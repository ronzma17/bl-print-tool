#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提单打印工具 - 第二张纸 (Pre-Arrival Notice)
不要直接双击, 请用 START_HERE.bat

做什么:
  把船公司发来的 Pre-Arrival Notice 每份取第 1 页, 拼成一个 PDF,
  顺序跟第一叠 (PRINT.pdf) 完全一致, 缺的位置插一张空白纸占位。
  出来的文件【单面打印】, 反面留白, 之后再塞回打印机印 GoFreight 的 Arrival Notice。

顺序从哪来:
  读第一叠生成的 _报告.csv, 照着里面的 HBL 顺序排。
  没有报告就退而求其次, 从 PDF 内容里找 HBL 自己排。

文件名要求:
  Apps Script 抓下来时已按提单号命名, 例如 CHN3966555.pdf

依赖: pip install pypdf
"""

import argparse
import csv
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.exit("缺少 pypdf,请先执行:  pip install pypdf")


# ============ 可调参数 ============
KEY_PATTERN = re.compile(r"[A-Z]{2,4}\d{6,}")        # MBL, 如 CHN3966555 / NGP4259647
HBL_PATTERN = re.compile(r"\b\d{4}[A-Z]\d{5,8}\b")   # HBL, 如 2609A336261
CONTAINER_PATTERN = re.compile(r"^[A-Z]{4}\d{7}$")   # 柜号 TIIU8576888, 长得像提单号, 要排除
TAKE_PAGES = [1]                                      # 每份只取第 1 页
OUT_NAME = "PRINT_SHEET2.pdf"
EXTRA_NAME = "_多出来的PREARRIVAL.pdf"
REPORT_NAME = "_报告_第二张纸.xlsx"
REPORT_CSV  = "_报告_第二张纸.csv"        # 没装 openpyxl 时的退路
# ==================================


def unzip_all(folder: Path):
    zips = list(folder.glob("*.zip"))
    if not zips:
        return
    n = 0
    for z in zips:
        try:
            with zipfile.ZipFile(z) as zf:
                for info in zf.infolist():
                    if info.is_dir():
                        continue
                    name = Path(info.filename).name
                    if not name.lower().endswith(".pdf"):
                        continue
                    t = folder / name
                    if t.exists():
                        continue
                    with zf.open(info) as s, open(t, "wb") as d:
                        d.write(s.read())
                    n += 1
        except Exception as e:
            print(f"  ! 解压失败 {z.name}: {e}")
    print(f"解压 {len(zips)} 个 zip,释放 {n} 个新 PDF")


ROW_PATTERN = re.compile(
    r"^([A-Z]{4}\d{7})\s+(\S+)\s+(\S+)\s+(.+?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s*$")
CNO_PATTERN = re.compile(r"^[A-Z]{4}\d{7}$")
TABLE_STOP = ("CHARGES", "PLEASE NOTE", "TOTAL CHARGES", "ABOVE PARTICULARS")


def writable_path(p: Path) -> Path:
    """Windows 上文件被 Excel / Adobe 开着时无法覆盖, 会抛 PermissionError。
       与其整个脚本崩掉, 不如换个带时间戳的名字存下来, 至少数据不丢。"""
    try:
        with open(p, "ab"):
            pass
        return p
    except PermissionError:
        alt = p.with_name(f"{p.stem}_{datetime.now():%H%M%S}{p.suffix}")
        print(f"  ! {p.name} 正被其他程序占用(多半是 Excel 或 Adobe 开着)")
        print(f"    改存为 {alt.name}")
        return alt
    except Exception:
        return p


def parse_containers(text):
    """从 Pre-Arrival 正文抠柜信息, 返回 [(柜号, 封条, 柜型, 件数, 重量, 材积), ...]

       有两种版式:
         单柜 -> 表格是一整行:  TIIU8566771 R9032062 40HC 1041CARTONS 27040 2401.00
         多柜 -> pypdf 按列抽:  两个柜号, 两个封条, 两个柜型 ... 竖着分组
       所以先试行式, 不行再按列拆。"""
    lines = text.split("\n")
    start = next((i for i, l in enumerate(lines)
                  if "CONTAINER" in l.upper() and "SEAL" in l.upper()), None)
    if start is None:
        return []

    block = []
    for l in lines[start + 1:]:
        u = l.strip().upper()
        if any(u.startswith(s) for s in TABLE_STOP):
            break
        if l.strip():
            block.append(l.strip())

    # 版式一: 一行一个柜
    rows = [ROW_PATTERN.match(l).groups() for l in block if ROW_PATTERN.match(l)]
    if rows:
        return [(a, b, c, d.strip(), e, f) for a, b, c, d, e, f in rows]

    # 版式二: 按列分组, 先数开头有几个柜号, 剩下的按组数平分
    tok = " ".join(block).split()
    n = 0
    while n < len(tok) and CNO_PATTERN.match(tok[n]):
        n += 1
    if n == 0:
        return []
    rest = tok[n:]
    if len(rest) == 6 * n:        # 件数和单位分开:  520 / CARTONS
        g = [rest[i * n:(i + 1) * n] for i in range(6)]
        seal, size, wt, ms = g[0], g[1], g[4], g[5]
        qty = [f"{g[2][i]}{g[3][i]}" for i in range(n)]
    elif len(rest) == 5 * n:      # 件数和单位黏一起: 520CARTONS
        g = [rest[i * n:(i + 1) * n] for i in range(5)]
        seal, size, qty, wt, ms = g[0], g[1], g[2], g[3], g[4]
    else:
        return []
    return [(tok[i], seal[i], size[i], qty[i], wt[i], ms[i]) for i in range(n)]


# (显示的表头, 内部取值的 key)。表头可以重复 —— 重量和 ETA 在前面放一份
# 方便核对, 后面原位置再放一份。
FIELDS = [
    ("第几张纸", "第几张纸"),
    ("HBL",     "HBL"),
    ("重量",     "重量"),        # C 列: 提到前面, 核对用
    ("ETA",     "ETA"),         # D 列: 提到前面, 核对用
    ("MBL",     "MBL"),
    ("文件",     "文件"),
    ("运输方式", "运输方式"),
    ("目的地",   "目的地"),
    ("柜数",     "柜数"),
    ("柜号",     "柜号"),
    ("封条",     "封条"),
    ("柜型",     "柜型"),
    ("件数",     "件数"),
    ("重量",     "重量"),        # 原本的位置
    ("材积",     "材积"),
    ("ETA",     "ETA"),         # 原本的位置
    ("状态",     "状态"),
]
HEADERS = [h for h, _ in FIELDS]
KEYS    = [k for _, k in FIELDS]


def write_report(out_dir, rows, extras, by_mbl, nokey, dupes):
    """优先写 xlsx(能上色), 没装 openpyxl 就退回 csv。
       Ramp 的票整行标黄, 缺 Pre-Arrival 的标红。"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  ! 没装 openpyxl, 报告改存成 csv(不带颜色)")
        print("    想要高亮请执行:  pip install openpyxl")
        return write_report_csv(out_dir, rows, extras, by_mbl, nokey, dupes)

    YELLOW = PatternFill("solid", fgColor="FFF2CC")     # Ramp
    RED    = PatternFill("solid", fgColor="FFC7CE")     # 缺件
    HEAD   = PatternFill("solid", fgColor="DDEBF7")

    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        ws.cell(1, c).font = Font(bold=True)
        ws.cell(1, c).fill = HEAD

    for r in rows:
        ws.append([r.get(k, "") for k in KEYS])
        i = ws.max_row
        fill = None
        if not r.get("文件"):
            fill = RED
        elif str(r.get("运输方式", "")).lower() == "ramp":
            fill = YELLOW
        if fill:
            for c in range(1, len(HEADERS) + 1):
                ws.cell(i, c).fill = fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{ws.max_row}"
    for c, (head, key) in enumerate(FIELDS, 1):
        longest = max([len(str(head))] + [len(str(r.get(key, ""))) for r in rows] or [0])
        ws.column_dimensions[get_column_letter(c)].width = min(max(longest + 2, 8), 40)

    # 第二个工作表放异常
    w2 = wb.create_sheet("异常")
    w2.append(["类型", "提单号", "文件"])
    for c in range(1, 4):
        w2.cell(1, c).font = Font(bold=True)
        w2.cell(1, c).fill = HEAD
    for m in sorted(extras):
        w2.append(["第一叠里没有的 Pre-Arrival", m, by_mbl[m].name])
    for n in nokey:
        w2.append(["认不出提单号", "", n])
    for n in dupes:
        w2.append(["重复的文件", "", n])
    for c, wdt in ((1, 30), (2, 16), (3, 40)):
        w2.column_dimensions[get_column_letter(c)].width = wdt

    path = writable_path(out_dir / REPORT_NAME)
    wb.save(path)

    ramp = sum(1 for r in rows if str(r.get("运输方式", "")).lower() == "ramp")
    print(f"  Ramp(转铁路) {ramp} 票, 已在报告里标黄")
    return path


def write_report_csv(out_dir, rows, extras, by_mbl, nokey, dupes):
    path = writable_path(out_dir / REPORT_CSV)
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        wtr = csv.writer(fh)
        wtr.writerow(HEADERS)
        for r in rows:
            wtr.writerow([r.get(k, "") for k in KEYS])
        fh.write("\n第一叠里没有的 Pre-Arrival\n")
        for m in sorted(extras):
            fh.write(f'"{m}","{by_mbl[m].name}"\n')
        if nokey:
            fh.write("\n认不出提单号的文件\n")
            for n in nokey:
                fh.write(f'"{n}"\n')
        if dupes:
            fh.write("\n重复的文件\n")
            for n in dupes:
                fh.write(f'"{n}"\n')
    return path


def parse_mode(text):
    """DEST.CARG MODE 的值。pypdf 抽出来时值和标签是分开的,
       所以直接找独立成行的 Port / Ramp / Door。
       注意不能用关键字搜全文 —— 条款里那句 "Last Free Day at Ramp"
       每一页都有, 会把所有票都误判成 Ramp。"""
    for l in text.split("\n"):
        v = l.strip()
        if v.lower() in ("port", "ramp", "door"):
            return v.capitalize()
    return ""


def parse_eta(text):
    """POD ETA。版面上它紧跟在 DEST.CARG MODE 的值后面,
       抽出来就是 Port/Ramp 那行的下一行。
       另一个日期(09/22/2026 07:21 PM)是通知的签发时间, 别搞混。"""
    lines = [l.strip() for l in text.split("\n")]
    bare = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    for i, l in enumerate(lines):
        if l.lower() in ("port", "ramp", "door"):
            for nxt in lines[i + 1:i + 4]:
                if bare.match(nxt):
                    return nxt
    for l in lines:                      # 兜底: 第一个不带时间的日期
        if bare.match(l):
            return l
    return ""


def final_dest(text):
    """最终目的地。DESTINATION 的标签和值在 PDF 里是分开的, 对不上,
       所以改成扫所有 "城市, ST" 格式的行:
       第一个通常是卸货港, 出现第二个不同的城市就是最终目的地(转铁路的情况)。
       收货人自己的地址(SAN MATEO)要排除掉。"""
    seen = []
    for l in text.split("\n"):
        v = l.strip()
        if v.upper().startswith(("PLACE OF ISSUE", "CLEARANCE POINT")):
            v = v.split(":", 1)[-1].strip()
        if not re.search(r",\s*[A-Z]{2}$", v):
            continue
        if "SAN MATEO" in v.upper() or len(v) > 40:
            continue
        if v not in seen:
            seen.append(v)
    if not seen:
        return ""
    for v in seen[1:]:
        if v != seen[0]:
            return v          # 跟卸货港不一样 -> 转运目的地
    return seen[0]


def join_col(rows, idx):
    """多个柜的同一项用逗号拼成一格:  40HC,40HC"""
    return ",".join(r[idx] for r in rows)


def key_from_text(txt):
    """从 PDF 正文里找提单号。
       坑: 柜号(TIIU8576888)也符合"字母+数字"的格式, 而且在正文里出现得比提单号早,
       直接通用匹配会把柜号当成提单号。所以先找带明确标签的。"""
    # 1) 最可靠: "The Chinese B/L Number as cross-reference is: CHN3966555"
    m = re.search(r"CROSS[- ]REFERENCE\s+IS[:\s]+([A-Z]{2,4}\d{6,})", txt)
    if m:
        return m.group(1)
    # 2) 次可靠: 清关用的 CMDUCHN3966555 这种 SCAC+提单号
    m = re.search(r"\b[A-Z]{4}([A-Z]{2,4}\d{6,})\b", txt)
    if m:
        return m.group(1)
    # 3) 兜底: 通用匹配, 但跳过柜号格式
    for m in KEY_PATTERN.finditer(txt):
        if not CONTAINER_PATTERN.match(m.group(0)):
            return m.group(0)
    return None


def first_page_text(path: Path):
    try:
        r = PdfReader(str(path))
        return len(r.pages), (r.pages[0].extract_text() or "").upper()
    except Exception as e:
        print(f"  ! 读不了 {path.name}: {e}")
        return 0, ""


def load_order(report: Path):
    """从第一叠的报告里读出 (HBL, MBL) 顺序。第一叠现在出的是 xlsx, 老的 csv 也还认。"""
    order = []
    with open(report, encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            mbl = (row.get("MBL") or "").strip()
            if not mbl:
                break
            order.append(((row.get("HBL") or "").strip(), mbl))
    return order


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="输入文件夹 4_DROP_PREARRIVAL_HERE")
    ap.add_argument("--out", required=True, help="输出文件夹 2_READY_TO_PRINT")
    ap.add_argument("--order-from", default=None, help="第一叠的 _报告.csv")
    ap.add_argument("--archive", default=None, help="归档文件夹 3_ARCHIVE")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    src = Path(args.folder)
    if not src.is_dir():
        sys.exit(f"找不到输入文件夹: {src}")
    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    unzip_all(src)

    pdfs = sorted(p for p in src.glob("*.pdf") if p.is_file())
    print(f"共 {len(pdfs)} 个 Pre-Arrival PDF\n")
    if not pdfs:
        print("输入文件夹是空的。先用 Apps Script 抓邮件附件,把 Drive 下载的 zip 放进来。")
        return

    # ---- 建立 MBL -> 文件 的索引 ----
    by_mbl, dupes, nokey = {}, [], []
    meta = {}
    for p in pdfs:
        npages, txt = first_page_text(p)
        m = KEY_PATTERN.search(p.stem.upper())
        mbl = m.group(0) if m else key_from_text(txt)   # 文件名没有就从正文找
        if not mbl:
            nokey.append(p.name)
            continue
        h = HBL_PATTERN.search(txt)
        meta[mbl] = (p, npages, h.group(0) if h else "", parse_containers(txt),
                     parse_mode(txt), final_dest(txt), parse_eta(txt))
        if mbl in by_mbl:
            dupes.append(p.name)
            continue
        by_mbl[mbl] = p

    # ---- 决定顺序 ----
    order, source = [], ""
    report = Path(args.order_from) if args.order_from else (out_dir / "_报告.csv")
    if report.exists():
        order = load_order(report)
        source = f"照第一叠的顺序 ({report.name})"
    else:
        order = sorted(((v[2], k) for k, v in meta.items()),
                       key=lambda x: (x[0] == "", x[0], x[1]))
        source = "第一叠的报告找不到,按 PDF 内容里的 HBL 自己排"
        print("!! 警告: " + source)
        print("!! 这一叠可能跟第一叠对不齐,叠之前务必逐张核对。\n")
    print(f"排序依据: {source}\n")

    # ---- 拼页 ----
    writer = None if args.dry_run else PdfWriter()
    rows, missing, used = [], 0, set()

    for i, (hbl, mbl) in enumerate(order, 1):
        p = by_mbl.get(mbl)
        if p is None:
            missing += 1
            if writer is not None:
                writer.add_blank_page(width=612, height=792)   # Letter 空白占位
            rows.append({"第几张纸": i, "HBL": hbl, "MBL": mbl, "文件": "",
                         "运输方式": "", "目的地": "", "ETA": "",
                         "柜数": 0, "柜号": "", "封条": "", "柜型": "",
                         "件数": "", "重量": "", "材积": "",
                         "状态": "缺 Pre-Arrival,已插空白纸占位"})
            continue
        used.add(mbl)
        _, npages, hbl_in_pdf, cons, mode, dest, eta = meta[mbl]
        take = [x for x in TAKE_PAGES if x <= npages]
        if writer is not None:
            r = PdfReader(str(p))
            for x in take:
                writer.add_page(r.pages[x - 1])
        note = ""
        if hbl and hbl_in_pdf and hbl != hbl_in_pdf:
            note = f"HBL 对不上 (PDF 里是 {hbl_in_pdf})"
        elif not take:
            note = "这个 PDF 是空的"
        if not cons:
            note = (note + " / " if note else "") + "柜信息读不出来"
        rows.append({"第几张纸": i, "HBL": hbl, "MBL": mbl, "文件": p.name,
                     "运输方式": mode, "目的地": dest, "ETA": eta,
                     "柜数": len(cons),
                     "柜号": join_col(cons, 0), "封条": join_col(cons, 1),
                     "柜型": join_col(cons, 2), "件数": join_col(cons, 3),
                     "重量": join_col(cons, 4), "材积": join_col(cons, 5),
                     "状态": note or "正常"})

    extras = [m for m in by_mbl if m not in used]

    # 一个都没对上时, 把两边的号打出来, 方便看是哪里不一致
    if order and missing == len(order):
        print("!! 一个都没对上, 两边的提单号完全不同。")
        print()
        print("   第一叠报告里的前 5 个 MBL:")
        for h, m in order[:5]:
            print(f"      {m}   (HBL {h})")
        print()
        print("   Pre-Arrival 文件夹里的前 5 个 MBL:")
        for m in sorted(by_mbl)[:5]:
            print(f"      {m}   <- {by_mbl[m].name}")
        print()
        print("   常见原因:")
        print("   1) 两边不是同一批货 (第一叠是别的船/航次)")
        print("   2) 2_READY_TO_PRINT 里的 _报告.csv 是上一批遗留的")
        print("      -> 先重跑第一叠, 或从 3_ARCHIVE 里找回对应那批的报告")
        print()

    print(f"第一叠共 {len(order)} 票")
    print(f"  正常     {len(order) - missing}")
    print(f"  缺失补白 {missing}")
    print(f"  多出来的 {len(extras)}")
    if nokey:
        print(f"  认不出提单号 {len(nokey)}")
    if dupes:
        print(f"  重复文件 {len(dupes)}")
    print()

    if writer is not None:
        with open(writable_path(out_dir / OUT_NAME), "wb") as fh:
            writer.write(fh)
        print(f"  {OUT_NAME}: {len(writer.pages)} 页 = {len(writer.pages)} 张纸 (单面打印)")

        if extras:
            w = PdfWriter()
            for m in sorted(extras):
                r = PdfReader(str(by_mbl[m]))
                for x in [y for y in TAKE_PAGES if y <= len(r.pages)]:
                    w.add_page(r.pages[x - 1])
            with open(writable_path(out_dir / EXTRA_NAME), "wb") as fh:
                w.write(fh)
            print(f"  {EXTRA_NAME}: {len(extras)} 票 (第一叠里没有的,不要混进去)")

    report_path = write_report(out_dir, rows, extras, by_mbl, nokey, dupes)

    print(f"\n报告: {report_path}")

    if args.dry_run:
        print("\n(检查模式: 没有生成 PDF, 什么文件都没动)")
    elif args.archive:
        stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
        dest = Path(args.archive) / (stamp + "_prearrival")
        dest.mkdir(parents=True, exist_ok=True)
        moved = 0
        for f in list(src.glob("*.zip")) + list(src.glob("*.pdf")):
            try:
                t = dest / f.name
                i = 1
                while t.exists():
                    t = dest / f"{f.stem}({i}){f.suffix}"
                    i += 1
                shutil.move(str(f), str(t))
                moved += 1
            except Exception as e:
                print(f"  ! 归档失败 {f.name}: {e}")
        print(f"已归档 {moved} 个原始文件 -> {dest.name}")

    print("\n打印设置: 【单面】打印, 不要勾双面。Shrink oversized pages, Letter 纸。")


if __name__ == "__main__":
    main()

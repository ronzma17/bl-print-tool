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
REPORT_NAME = "_报告_第二张纸.csv"
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
        meta[mbl] = (p, npages, h.group(0) if h else "")
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
                         "状态": "缺 Pre-Arrival,已插空白纸占位"})
            continue
        used.add(mbl)
        _, npages, hbl_in_pdf = meta[mbl]
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
        rows.append({"第几张纸": i, "HBL": hbl, "MBL": mbl,
                     "文件": p.name, "状态": note or "正常"})

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
        with open(out_dir / OUT_NAME, "wb") as fh:
            writer.write(fh)
        print(f"  {OUT_NAME}: {len(writer.pages)} 页 = {len(writer.pages)} 张纸 (单面打印)")

        if extras:
            w = PdfWriter()
            for m in sorted(extras):
                r = PdfReader(str(by_mbl[m]))
                for x in [y for y in TAKE_PAGES if y <= len(r.pages)]:
                    w.add_page(r.pages[x - 1])
            with open(out_dir / EXTRA_NAME, "wb") as fh:
                w.write(fh)
            print(f"  {EXTRA_NAME}: {len(extras)} 票 (第一叠里没有的,不要混进去)")

    report_path = out_dir / REPORT_NAME
    with open(report_path, "w", newline="", encoding="utf-8-sig") as fh:
        wtr = csv.DictWriter(fh, fieldnames=["第几张纸", "HBL", "MBL", "文件", "状态"])
        wtr.writeheader()
        wtr.writerows(rows)
        fh.write("\n第一叠里没有的 Pre-Arrival\n")
        for m in sorted(extras):
            fh.write(f'"{m}","{by_mbl[m].name}"\n')

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

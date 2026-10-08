#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提单打印工具 - 引擎
不要直接双击这个文件,请双击 START_HERE.bat

流程:
  解压 zip -> 按提单号配对 -> 合并成双面打印就绪的 PDF -> 原始文件归档
  下次再跑时, 上一批的成品 PDF 会自动挪进同一个归档文件夹, 不用手动删

每票 = [Transport B/L 全部页] + [Waybill 第1页] = 2 页 = 双面打印 1 张纸
        正面 B/L, 反面 Waybill

文件识别:
    短文件名  CHN1234567                -> Waybill (只取第1页)
    长文件名  CHN1122211 1111A1111111   -> Transport B/L
    同时读 PDF 内容交叉验证

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
KEY_PATTERN = re.compile(r"[A-Z]{2,4}\d{6,}")   # 提单号格式, 如 CHN3966555
WAYBILL_PAGES = [1]                             # Waybill 只打第几页
BL_HINT = "COMBINED TRANSPORT"
WB_HINT = "WAYBILL"
MARKER = "_batch.txt"                           # 记录本批归档到哪个文件夹
EXTRA_OUT = ("_需人工处理.pdf", "_报告.csv")
# ==================================


CONTAINER_PATTERN = re.compile(r"\b[A-Z]{4}\d{7}\b")  # 柜号, 别当成提单号
HBL_PATTERN = re.compile(r"\d{4}[A-Z]\d{5,8}")        # HBL, 如 2609A338346 / AIC2609A338346


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




def load_manifest(path: Path):
    """读船公司给的舱单(xls/xlsx/csv), 把里面所有 MBL 号抠出来。
       不认列名, 直接全表扫描找符合提单号格式的格子 —— 船公司的表头经常变,
       认列名很容易哪天就挂了, 扫格子反而稳。"""
    cells = []
    suf = path.suffix.lower()
    try:
        if suf == ".xls":
            import xlrd
            bk = xlrd.open_workbook(str(path))
            for sh in bk.sheets():
                for r in range(sh.nrows):
                    for c in range(sh.ncols):
                        cells.append(str(sh.cell_value(r, c)))
        elif suf in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook
            wb = load_workbook(str(path), read_only=True, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    cells.extend(str(v) for v in row if v is not None)
        else:                                   # csv / txt
            with open(path, encoding="utf-8-sig", errors="ignore") as fh:
                for line in fh:
                    cells.extend(line.replace(",", " ").split())
    except ImportError:
        print(f"  ! 读不了 {path.name}: 缺少 xlrd, 请执行  pip install xlrd")
        return set()
    except Exception as e:
        print(f"  ! 读不了 {path.name}: {e}")
        return set()

    keys = set()
    for v in cells:
        for m in KEY_PATTERN.finditer(v.upper()):
            k = m.group(0)
            if not CONTAINER_PATTERN.fullmatch(k):
                keys.add(k)
    return keys




def is_output_file(f: Path) -> bool:
    return f.is_file() and (f.name.startswith("PRINT") or f.name in EXTRA_OUT)


def archive_old_output(out_dir: Path, archive_root: Path):
    """把上一批留下的成品挪走, 跟那一批的原始文件归到同一个文件夹"""
    olds = [f for f in out_dir.iterdir() if is_output_file(f)]
    marker = out_dir / MARKER

    if not olds:
        if marker.exists():
            marker.unlink()
        return

    if archive_root is None:                      # 没配归档就只能删
        for f in olds:
            try:
                f.unlink()
            except Exception:
                pass
        print(f"清掉上一批的 {len(olds)} 个成品文件")
        return

    name = ""
    if marker.exists():
        try:
            name = marker.read_text(encoding="utf-8").strip()
        except Exception:
            name = ""
    if not name:                                  # 记号丢了就用文件修改时间兜底
        name = datetime.fromtimestamp(olds[0].stat().st_mtime).strftime("%Y-%m-%d_%H%M")

    dest = archive_root / name
    dest.mkdir(parents=True, exist_ok=True)
    moved = 0
    for f in olds:
        try:
            t = dest / f.name
            i = 1
            while t.exists():
                t = dest / f"{f.stem}({i}){f.suffix}"
                i += 1
            shutil.move(str(f), str(t))
            moved += 1
        except Exception as e:
            print(f"  ! 挪不动 {f.name}: {e}")
    print(f"上一批的 {moved} 个成品已归档 -> {name}")

    if marker.exists():
        marker.unlink()


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
                    target = folder / name
                    if target.exists():
                        continue
                    with zf.open(info) as s, open(target, "wb") as d:
                        d.write(s.read())
                    n += 1
        except Exception as e:
            print(f"  ! 解压失败 {z.name}: {e}")
    print(f"解压 {len(zips)} 个 zip,释放 {n} 个新 PDF")


def read_pdf(path: Path):
    try:
        r = PdfReader(str(path))
        return len(r.pages), (r.pages[0].extract_text() or "").upper()
    except Exception as e:
        print(f"  ! 读不了 {path.name}: {e}")
        return 0, ""


def kind_by_name(path: Path):
    """返回 (MBL号, 类型, HBL号)
       CHN3962684 2609A336281 -> ("CHN3962684", "bl", "2609A336281")
       CHN3962684             -> ("CHN3962684", "wb", "")
    """
    stem = path.stem.upper().strip()
    m = KEY_PATTERN.search(stem)
    if not m:
        return None, None, ""
    rest = (stem[:m.start()] + stem[m.end():]).strip(" _-")
    hbl = re.split(r"[\s_\-]+", rest)[0] if rest else ""
    return m.group(0), ("bl" if rest else "wb"), hbl


def kind_by_text(txt: str):
    if BL_HINT in txt:
        return "bl"
    if WB_HINT in txt:
        return "wb"
    return None


def archive_sources(src: Path, archive_root: Path) -> str:
    """把处理过的 zip 和 pdf 挪到归档文件夹, 返回文件夹名"""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    dest = archive_root / stamp
    dest.mkdir(parents=True, exist_ok=True)
    moved = 0
    picks = [f for f in src.iterdir()
             if f.is_file() and f.suffix.lower() in (".zip", ".pdf", ".xls", ".xlsx", ".xlsm")]
    for f in picks:
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
    print(f"已归档 {moved} 个原始文件 -> {stamp}")
    return stamp


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="输入文件夹 1_DROP_ZIPS_HERE")
    ap.add_argument("--out", default=None, help="输出文件夹 2_READY_TO_PRINT")
    ap.add_argument("--archive", default=None, help="归档文件夹 3_ARCHIVE")
    ap.add_argument("--chunk", type=int, default=0, help="每个输出文件放几票, 0=不拆")
    ap.add_argument("--manifest", default=None,
                    help="船公司舱单 xls/xlsx/csv, 跑完自动核对缺哪几票")
    ap.add_argument("--dry-run", action="store_true", help="只检查, 什么都不动")
    args = ap.parse_args()

    src = Path(args.folder)
    if not src.is_dir():
        sys.exit(f"找不到输入文件夹: {src}")
    out_dir = Path(args.out) if args.out else src / "_PRINT"
    out_dir.mkdir(parents=True, exist_ok=True)
    arc = Path(args.archive) if args.archive else None
    if arc:
        arc.mkdir(parents=True, exist_ok=True)

    if not args.dry_run:
        archive_old_output(out_dir, arc)

    unzip_all(src)

    pdfs = sorted(p for p in src.glob("*.pdf") if p.is_file())
    print(f"共 {len(pdfs)} 个 PDF\n")
    if not pdfs:
        print("输入文件夹里没有 PDF。请先把 zip 放进 1_DROP_ZIPS_HERE。")
        return

    groups, problems = {}, []
    for p in pdfs:
        key, k_name, hbl = kind_by_name(p)
        npages, txt = read_pdf(p)
        k_text = kind_by_text(txt)
        if not key:
            problems.append((p.name, "文件名里找不到提单号"))
            continue
        kind = k_name or k_text
        if kind is None:
            problems.append((p.name, "认不出是 B/L 还是 Waybill"))
            continue
        if k_text and k_name and k_text != k_name:
            problems.append((p.name, f"文件名像 {k_name} 但内容像 {k_text},已按文件名处理"))
        slot = groups.setdefault(key, {})
        if kind in slot:
            problems.append((p.name, f"{key} 的 {kind} 重复"))
            continue
        slot[kind] = (p, npages)
        if kind == "bl":
            slot["hbl"] = hbl

    matched, singles = [], []
    for key in groups:
        s = groups[key]
        if "bl" in s and "wb" in s:
            h = s.get("hbl", "")
            if not h:
                problems.append((s["bl"][0].name, f"{key}: B/L 文件名里没找到 HBL 号,排到最后"))
            matched.append((key, h, s["bl"], s["wb"]))
        else:
            f, _ = s.get("bl") or s["wb"]
            singles.append((key, f))
            problems.append((f.name, f"{key}: " + ("缺 Waybill" if "bl" in s else "缺 Transport B/L")))

    # 按 HBL 排序; 没有 HBL 的排最后, 再按 MBL 排
    matched.sort(key=lambda x: (x[1] == "", x[1], x[0]))
    singles.sort(key=lambda x: x[0])

    print(f"配对成功 {len(matched)} 票 | 落单 {len(singles)} | 问题 {len(problems)}\n")

    # ---- 跟船公司舱单核对 ----
    man_lines, man_info = [], {}
    mp = Path(args.manifest) if args.manifest else None
    if mp is None:                                   # 没指定就在输入文件夹里自己找
        found = sorted(f for f in src.iterdir()
                       if f.is_file() and f.suffix.lower() in (".xls", ".xlsx", ".xlsm"))
        if found:
            mp = found[0]
            print(f"发现舱单: {mp.name}\n")
    if mp:
        if not mp.exists():
            print(f"!! 找不到舱单: {mp}\n")
        else:
            want = load_manifest(mp)
            have = {k for k, _, _, _ in matched} | {k for k, _ in singles}
            lack = sorted(want - have)
            more = sorted(have - want)
            print(f"--- 舱单核对 ({mp.name}) ---")
            print(f"  舱单共      {len(want)} 票")
            print(f"  手上有      {len(have)} 票")
            print(f"  舱单有你没有 {len(lack)} 票")
            print(f"  你有舱单没有 {len(more)} 票")
            if lack:
                print("  缺这些:")
                for k in lack:
                    print(f"     {k}")
            if more:
                print("  多这些:")
                for k in more:
                    print(f"     {k}")
            print()
            man_lines.append(("舱单核对", mp.name))
            man_lines.append(("舱单票数", len(want)))
            man_lines.append(("手上票数", len(have)))
            for k in lack:
                man_lines.append(("舱单有你没有", k))
            for k in more:
                man_lines.append(("你有舱单没有", k))

    rows = []
    batches = ([matched] if args.chunk <= 0
               else [matched[i:i + args.chunk] for i in range(0, len(matched), args.chunk)])

    for bi, batch in enumerate(batches, 1):
        writer = None if args.dry_run else PdfWriter()
        sheet = 0
        for key, hbl, (bl, nbl), (wb, nwb) in batch:
            take = [i for i in WAYBILL_PAGES if i <= nwb]
            total = nbl + len(take)
            pad = 1 if total % 2 else 0
            if writer is not None:
                for pg in PdfReader(str(bl)).pages:
                    writer.add_page(pg)
                wbr = PdfReader(str(wb))
                for i in take:
                    writer.add_page(wbr.pages[i - 1])
                if pad:
                    last = writer.pages[-1]
                    writer.add_blank_page(width=last.mediabox.width,
                                          height=last.mediabox.height)
            start_sheet = sheet + 1
            sheet += (total + pad) // 2
            note = ""
            if nbl != 1:
                note = f"B/L 有 {nbl} 页, 正反面可能错位"
            if nwb < max(WAYBILL_PAGES):
                note = f"Waybill 只有 {nwb} 页"

            rows.append({
                "HBL": hbl, "MBL": key, "批次": bi, "第几张纸": start_sheet,
                "BL文件": bl.name, "BL页数": nbl,
                "Waybill文件": wb.name, "Waybill总页数": nwb,
                "合计页数": total + pad, "占纸张": (total + pad) // 2, "备注": note,
            })

        if writer is not None:
            name = "PRINT.pdf" if len(batches) == 1 else f"PRINT_{bi:02d}.pdf"
            with open(writable_path(out_dir / name), "wb") as fh:
                writer.write(fh)
            print(f"  {name}: {len(writer.pages)} 页 = {len(writer.pages)//2} 张纸")

    if singles and not args.dry_run:
        w = PdfWriter()
        for _, f in singles:
            for pg in PdfReader(str(f)).pages:
                w.add_page(pg)
            if len(w.pages) % 2:
                last = w.pages[-1]
                w.add_blank_page(width=last.mediabox.width, height=last.mediabox.height)
        with open(writable_path(out_dir / "_需人工处理.pdf"), "wb") as fh:
            w.write(fh)
        print(f"  _需人工处理.pdf: {len(singles)} 个落单文件")

    report_path = writable_path(out_dir / "_报告.csv")
    with open(report_path, "w", newline="", encoding="utf-8-sig") as fh:
        if rows:
            wtr = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            wtr.writeheader()
            wtr.writerows(rows)
        if man_lines:
            fh.write("\n舱单核对\n项目,内容\n")
            for x, y in man_lines:
                fh.write(f'"{x}","{y}"\n')
        fh.write("\n问题清单\n文件,原因\n")
        for n, why in problems:
            fh.write(f'"{n}","{why}"\n')

    print(f"\n报告: {report_path}")

    if args.dry_run:
        print("\n(检查模式: 没有生成 PDF, 什么文件都没动)")
    elif arc:
        stamp = archive_sources(src, arc)
        # 留个记号, 下次跑的时候把这批成品挪到同一个归档文件夹
        (out_dir / MARKER).write_text(stamp, encoding="utf-8")


if __name__ == "__main__":
    main()

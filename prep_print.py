#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
提单打印工具 - 引擎
不要直接双击这个文件,请双击 START_HERE.bat

流程:
  解压 zip -> 按提单号配对 -> 合并成双面打印就绪的 PDF -> 原始文件归档

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
    """靠文件名判断: 提单号后面还有别的东西 = B/L, 只有提单号 = Waybill"""
    stem = path.stem.upper().strip()
    m = KEY_PATTERN.search(stem)
    if not m:
        return None, None
    rest = (stem[:m.start()] + stem[m.end():]).strip(" _-")
    return m.group(0), ("bl" if rest else "wb")


def kind_by_text(txt: str):
    if BL_HINT in txt:
        return "bl"
    if WB_HINT in txt:
        return "wb"
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="输入文件夹 1_DROP_ZIPS_HERE")
    ap.add_argument("--out", default=None, help="输出文件夹 2_READY_TO_PRINT")
    ap.add_argument("--chunk", type=int, default=0, help="每个输出文件放几票, 0=不拆")
    ap.add_argument("--dry-run", action="store_true", help="只检查, 什么都不动")
    args = ap.parse_args()

    src = Path(args.folder)
    if not src.is_dir():
        sys.exit(f"找不到输入文件夹: {src}")
    out_dir = Path(args.out) if args.out else src / "_PRINT"
    out_dir.mkdir(parents=True, exist_ok=True)
    unzip_all(src)

    pdfs = sorted(p for p in src.glob("*.pdf") if p.is_file())
    print(f"共 {len(pdfs)} 个 PDF\n")
    if not pdfs:
        print("输入文件夹里没有 PDF。请先把 zip 放进 1_DROP_ZIPS_HERE。")
        return

    groups, problems = {}, []
    for p in pdfs:
        key, k_name = kind_by_name(p)
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

    matched, singles = [], []
    for key in groups:
        s = groups[key]
        if "bl" in s and "wb" in s:
            matched.append((key, "", s["bl"], s["wb"]))
        else:
            f, _ = s.get("bl") or s["wb"]
            singles.append((key, f))
            problems.append((f.name, f"{key}: " + ("缺 Waybill" if "bl" in s else "缺 Transport B/L")))

    matched.sort(key=lambda x: x[0])
    singles.sort(key=lambda x: x[0])

    print(f"配对成功 {len(matched)} 票 | 落单 {len(singles)} | 问题 {len(problems)}\n")

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
                "提单号": key, "批次": bi, "第几张纸": start_sheet,
                "BL文件": bl.name, "BL页数": nbl,
                "Waybill文件": wb.name, "Waybill总页数": nwb,
                "合计页数": total + pad, "占纸张": (total + pad) // 2, "备注": note,
            })

        if writer is not None:
            name = "PRINT.pdf" if len(batches) == 1 else f"PRINT_{bi:02d}.pdf"
            with open(out_dir / name, "wb") as fh:
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
        with open(out_dir / "_需人工处理.pdf", "wb") as fh:
            w.write(fh)
        print(f"  _需人工处理.pdf: {len(singles)} 个落单文件")

    report_path = (out_dir / "_报告.csv")
    with open(report_path, "w", newline="", encoding="utf-8-sig") as fh:
        if rows:
            wtr = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            wtr.writeheader()
            wtr.writerows(rows)
        fh.write("\n问题清单\n文件,原因\n")
        for n, why in problems:
            fh.write(f'"{n}","{why}"\n')

    print(f"\n报告: {report_path}")

    if args.dry_run:
        print("\n(检查模式: 没有生成 PDF, 什么文件都没动)")


if __name__ == "__main__":
    main()

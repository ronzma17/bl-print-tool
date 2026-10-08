#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把一堆 Combined Transport B/L + Waybill 的 PDF 按提单号配对,
合并成一个「双面打印就绪」的大 PDF。

顺序: 每票 = [Combined B/L 全部页] + [Waybill 全部页] (+ 空白页凑偶数)
这样用打印机的「双面 / 长边翻转」打出来, 每张纸正面 B/L 反面 Waybill。

用法:
    python merge_for_duplex_print.py "C:\\Users\\me\\Downloads\\0917"
    python merge_for_duplex_print.py "输入文件夹" --out "输出文件夹" --chunk 50

依赖:
    pip install pypdf
"""

import argparse
import csv
import re
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.exit("缺少 pypdf,请先执行:  pip install pypdf")


# ============ 可调参数 ============

# 提单号的格式。截图里是 CHN3962684 -> 3个大写字母 + 6位以上数字
KEY_PATTERN = re.compile(r"[A-Z]{2,4}\d{6,}")

# 正面(第一页)认哪些关键词
BL_KEYWORDS = ["COMBINED TRANSPORT BILL OF LADING", "BILL OF LADING"]

# 反面认哪些关键词
WB_KEYWORDS = ["WAYBILL", "NON NEGOTIABLE"]

# ==================================


def page1_text(path):
    """读第一页文字,失败返回空字符串(扫描件会是空的)"""
    try:
        reader = PdfReader(str(path))
        return (reader.pages[0].extract_text() or "").upper()
    except Exception as e:
        print(f"  ! 读不了 {path.name}: {e}")
        return ""


def classify(path, text):
    """判断这份 PDF 是正面(bl)还是反面(wb)"""
    if any(k in text for k in WB_KEYWORDS) and "COMBINED TRANSPORT" not in text:
        return "wb"
    if any(k in text for k in BL_KEYWORDS):
        return "bl"
    name = path.stem.upper()
    if "WAYBILL" in name:
        return "wb"
    if "BL" in name or "LADING" in name:
        return "bl"
    return "unknown"


def extract_key(path, text):
    """提单号:先从文件名找,找不到再从内容找"""
    m = KEY_PATTERN.search(path.stem.upper())
    if m:
        return m.group(0)
    m = KEY_PATTERN.search(text)
    return m.group(0) if m else None


def append_pdf(writer, path):
    reader = PdfReader(str(path))
    for page in reader.pages:
        writer.add_page(page)
    return len(reader.pages)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="放 PDF 的文件夹")
    ap.add_argument("--out", default=None)
    ap.add_argument("--chunk", type=int, default=0,
                    help="每个输出文件最多放几票(0=不拆分)")
    args = ap.parse_args()

    src = Path(args.folder)
    if not src.is_dir():
        sys.exit(f"找不到文件夹: {src}")
    out_dir = Path(args.out) if args.out else src / "_PRINT"
    out_dir.mkdir(exist_ok=True)

    pdfs = sorted(p for p in src.glob("*.pdf") if p.is_file())
    print(f"扫描到 {len(pdfs)} 个 PDF\n")

    groups, problems = {}, []
    for p in pdfs:
        text = page1_text(p)
        kind = classify(p, text)
        key = extract_key(p, text)
        if kind == "unknown":
            problems.append((p.name, "认不出是 B/L 还是 Waybill"))
            continue
        if not key:
            problems.append((p.name, "找不到提单号"))
            continue
        slot = groups.setdefault(key, {})
        if kind in slot:
            problems.append((p.name, f"提单号 {key} 的 {kind} 重复了"))
            continue
        slot[kind] = p

    matched, unmatched = [], []
    for key in sorted(groups):
        slot = groups[key]
        if "bl" in slot and "wb" in slot:
            matched.append((key, slot["bl"], slot["wb"]))
        else:
            have = "只有 B/L,缺 Waybill" if "bl" in slot else "只有 Waybill,缺 B/L"
            f = slot.get("bl") or slot.get("wb")
            unmatched.append((key, f))
            problems.append((f.name, f"{key}: {have}"))

    print(f"配对成功 {len(matched)} 票,有问题 {len(problems)} 个文件\n")

    rows = []
    batches = ([matched] if args.chunk <= 0
               else [matched[i:i + args.chunk] for i in range(0, len(matched), args.chunk)])

    for bi, batch in enumerate(batches, 1):
        writer = PdfWriter()
        for key, bl, wb in batch:
            start = len(writer.pages) + 1
            n_bl = append_pdf(writer, bl)
            n_wb = append_pdf(writer, wb)
            total = n_bl + n_wb
            padded = 0
            if total % 2:          # 奇数页 -> 补白页, 保证下一票从新的一张纸开始
                last = writer.pages[-1]
                writer.add_blank_page(width=last.mediabox.width,
                                      height=last.mediabox.height)
                padded, total = 1, total + 1
            rows.append({
                "提单号": key, "批次": bi, "起始页": start,
                "BL文件": bl.name, "BL页数": n_bl,
                "Waybill文件": wb.name, "Waybill页数": n_wb,
                "本票总页数": total, "补白页": padded,
                "占纸张数": total // 2,
            })
        suffix = "" if len(batches) == 1 else f"_{bi:02d}"
        target = out_dir / f"PRINT{suffix}.pdf"
        with open(target, "wb") as fh:
            writer.write(fh)
        print(f"  已生成 {target.name}  ({len(writer.pages)} 页 / {len(writer.pages)//2} 张纸)")

    if unmatched:
        writer = PdfWriter()
        for key, f in unmatched:
            append_pdf(writer, f)
            if len(writer.pages) % 2:
                last = writer.pages[-1]
                writer.add_blank_page(width=last.mediabox.width, height=last.mediabox.height)
        with open(out_dir / "_需人工处理.pdf", "wb") as fh:
            writer.write(fh)
        print(f"  已生成 _需人工处理.pdf ({len(unmatched)} 个落单文件)")

    with open(out_dir / "_报告.csv", "w", newline="", encoding="utf-8-sig") as fh:
        if rows:
            w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        fh.write("\n问题清单\n文件,原因\n")
        for name, why in problems:
            fh.write(f'"{name}","{why}"\n')

    print(f"\n报告: {out_dir / '_报告.csv'}")
    print("\n下一步: 打开 PRINT.pdf -> 打印 -> 勾选「双面打印」-> 翻转方式选「长边翻转」")


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

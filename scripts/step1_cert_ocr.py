# -*- coding: utf-8 -*-
"""
Step 1 — 编号 OCR 核查
======================
遍历目标文件夹下所有扫描版 PDF，OCR 提取证书编号，与文件名编号交叉比对，
输出 Excel 核查报告。

用法:
    python step1_cert_ocr.py [底稿文件夹] [输出Excel] [--type 证书类型]

参数:
    底稿文件夹   默认当前目录，自动递归所有子文件夹收集 .pdf
    输出Excel    默认 底稿文件夹/证书编号核查结果.xlsx
    --type      证书类型（realestate/patent/trademark/business_license/qualification）
                默认 realestate；未知类型会在报告中标记「未识别类型」

依赖: rapidocr-onnxruntime PyMuPDF pandas openpyxl numpy
"""
import os
import re
import sys
import argparse
from pathlib import Path
from collections import defaultdict

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import ocr_common as oc


def parse_args():
    p = argparse.ArgumentParser(description="扫描件编号 OCR 核查")
    p.add_argument("base_dir", nargs="?", default=".", help="底稿文件夹（默认当前目录）")
    p.add_argument("output", nargs="?", default=None, help="输出 Excel 路径")
    p.add_argument("--type", default="realestate",
                   help="证书类型: " + "/".join(oc.list_cert_types()))
    return p.parse_args()


def collect_pdfs(base: Path):
    """递归收集所有 .pdf 文件，返回 [(相对文件夹, 文件名, 完整路径)]。"""
    out = []
    for f in base.rglob("*"):
        if f.is_file() and f.suffix.lower() == ".pdf":
            rel = f.parent.relative_to(base) if f.parent != base else Path(".")
            out.append((str(rel), f.name, f))
    return out


def extract_cert_from_page_text(page_text, cert_type):
    """
    从单页 OCR 文本提取编号。
    返回 (main_num, all_candidates)。过滤"原证号"引用。
    """
    cfg = oc.get_cert_type(cert_type)
    if cfg is None or not page_text:
        return None, []

    lines = page_text.split("\n")
    all_candidates = []
    main_num = None

    for i, line in enumerate(lines):
        # 只处理含关键标识的行
        if cert_type == "realestate" and "不动产权第" not in line:
            continue
        window = " ".join(lines[max(0, i - 4): min(len(lines), i + 5)])

        # 过滤"原证号"引用（附记页里的旧证引用）
        is_original_ref = bool(re.search(r"(原证|原不动产权|原专利|原商标)", window))

        # 用该类型正则提取
        m_main, m_all = oc.extract_cert_number_from_text(window, cert_type)
        for n in m_all:
            if n not in all_candidates:
                all_candidates.append(n)

        if not is_original_ref and main_num is None and m_main:
            main_num = m_main

    return main_num, all_candidates


def main():
    args = parse_args()
    base = Path(args.base_dir)
    output = Path(args.output) if args.output else base / "证书编号核查结果.xlsx"
    cert_type = args.type

    cfg = oc.get_cert_type(cert_type)
    if cfg is None:
        print(f"[警告] 未知证书类型 '{cert_type}'，将尝试自动探测")
        known_types = oc.list_cert_types()
    else:
        known_types = [cert_type]

    print("=== 初始化 RapidOCR ===")
    engine = oc.get_ocr_engine()
    print("初始化完成\n")

    pdfs = collect_pdfs(base)
    print(f"共发现 {len(pdfs)} 个 PDF 文件\n")

    # 跨文件夹重复检测（按文件大小）
    size_map = defaultdict(list)
    for rel, name, fp in pdfs:
        size_map[fp.stat().st_size].append((rel, name, fp))
    dup_map = {}
    for sz, entries in size_map.items():
        if len(entries) >= 2:
            for rel, name, fp in entries:
                others = [e[1] for e in entries if e[2] != fp]
                dup_map[fp] = others

    results = []
    MAX_PAGES = 6
    consecutive_miss = 0  # 连续未提取计数，用于 OCR 异常区段检测

    for idx, (folder, filename, pdf_path) in enumerate(pdfs):
        print(f"[{idx+1}/{len(pdfs)}] {folder}/{filename}")

        row = {
            "序号": idx + 1,
            "文件夹": folder,
            "文件名": filename,
            "文件_编号": None,
            "OCR_编号": None,
            "OCR_所有候补编号": "",
            "OCR_所在页码": "",
            "是否一致": "",
            "匹配方式": "",
            "备注": "",
        }

        file_num = oc.extract_cert_number_from_filename(filename, cert_type) if cfg else None
        row["文件_编号"] = file_num

        if pdf_path in dup_map:
            row["备注"] = "跨文件夹重复: " + ", ".join(dup_map[pdf_path])

        try:
            doc = None
            best_main = None
            best_page = None
            all_candidates = []
            detected_type = None

            import fitz
            doc = fitz.open(str(pdf_path))
            for pi in range(min(doc.page_count, MAX_PAGES)):
                img, _ = oc.render_pdf_page(str(pdf_path), pi)
                if img is None:
                    continue
                page_text = oc.ocr_page_text(img, engine)
                if not page_text:
                    continue

                # 未指定类型时，自动探测
                if cfg is None:
                    for tk in known_types:
                        m_main, m_all = extract_cert_from_page_text(page_text, tk)
                        if m_main:
                            detected_type = tk
                            best_main = m_main
                            best_page = pi + 1
                            all_candidates = m_all
                            break
                else:
                    m_main, m_all = extract_cert_from_page_text(page_text, cert_type)
                    if m_main:
                        if best_main is None:
                            best_main = m_main
                            best_page = pi + 1
                        all_candidates.extend(m_all)

            if best_main:
                row["OCR_编号"] = best_main
                row["OCR_所有候补编号"] = ", ".join(dict.fromkeys(all_candidates))
                row["OCR_所在页码"] = str(best_page)
                if detected_type:
                    row["备注"] = (row["备注"] or "") + f"; 自动识别类型: {detected_type}"

                # 比对
                fn = oc.normalize_number(file_num)
                on = oc.normalize_number(best_main)
                if fn and on:
                    if fn == on:
                        row["是否一致"] = "一致"
                        row["匹配方式"] = "精确匹配"
                    elif on.startswith(fn) or fn.startswith(on):
                        row["是否一致"] = "不一致"
                        row["匹配方式"] = "前缀/后缀容错"
                        row["备注"] = (row["备注"] or "") + f"; 文件={fn} vs OCR={on}"
                    else:
                        row["是否一致"] = "不一致"
                        row["匹配方式"] = "未匹配"
                        row["备注"] = (row["备注"] or "") + f"; 文件={fn} vs OCR={on}"
                elif on and not fn:
                    row["是否一致"] = "仅OCR提取到"
                else:
                    row["是否一致"] = "不一致"
            else:
                row["是否一致"] = "OCR未提取到"
                consecutive_miss += 1
                if consecutive_miss >= 3:
                    row["备注"] = (row["备注"] or "") + "; OCR异常区段(连续≥3个未提取，建议整段人工复核)"
            if best_main:
                consecutive_miss = 0

        except Exception as e:
            row["备注"] = (row["备注"] or "") + "; 异常: " + str(e)[:100]
            row["是否一致"] = "处理失败"
        finally:
            if doc is not None:
                doc.close()

        results.append(row)

    df = pd.DataFrame(results)
    df.to_excel(str(output), index=False, engine="openpyxl")

    consistent = sum(1 for r in results if r["是否一致"] == "一致")
    inconsistent = sum(1 for r in results if r["是否一致"] == "不一致")
    other = len(results) - consistent - inconsistent

    print("\n=== 核查完成 ===")
    print(f"总数: {len(results)}  一致: {consistent}  不一致: {inconsistent}  其他: {other}")
    print(f"报告: {output}")


if __name__ == "__main__":
    main()

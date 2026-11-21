# -*- coding: utf-8 -*-
"""
Step 2 — 查册/清单 vs 底稿交叉比对
==================================
从查册（扫描版 PDF 或 Excel 清单）中提取所有证书编号，与底稿文件夹中的文件
逐一交叉核对，找出缺口（查册里有、底稿里缺的）。

用法:
    python step2_cross_check.py [查册文件或文件夹] [底稿文件夹] [输出Excel] [--type 证书类型]

参数:
    查册文件或文件夹   扫描版 PDF（单文件或文件夹）或 Excel 清单(.xlsx/.xls/.csv)
    底稿文件夹        存放底稿扫描件的文件夹（递归收集 .pdf）
    输出Excel         默认 底稿文件夹/查册交叉比对结果.xlsx
    --type            证书类型，默认 realestate

说明:
    - 查册为 Excel 时，需含编号列；脚本自动识别含"编号/证号/注册号/专利号"等关键词的列
    - 查册为扫描 PDF 时，先 OCR 提取所有证号，再比对
    - 自动识别"名下无不动产"等空查册
"""
import os
import re
import sys
import argparse
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
import ocr_common as oc


def parse_args():
    p = argparse.ArgumentParser(description="查册/清单 vs 底稿交叉比对")
    p.add_argument("checklist", help="查册文件/文件夹（PDF 或 Excel）")
    p.add_argument("base_dir", help="底稿文件夹")
    p.add_argument("output", nargs="?", default=None, help="输出 Excel 路径")
    p.add_argument("--type", default="realestate",
                   help="证书类型: " + "/".join(oc.list_cert_types()))
    return p.parse_args()


def collect_files(path: Path, exts):
    """收集文件：单文件返回 [path]，文件夹递归收集指定扩展名。"""
    if path.is_file():
        return [path]
    out = []
    for f in path.rglob("*"):
        if f.is_file() and f.suffix.lower() in exts:
            out.append(f)
    return out


def is_excel(path: Path):
    return path.suffix.lower() in (".xlsx", ".xls", ".csv")


def extract_numbers_from_excel(path: Path, cert_type):
    """从 Excel 清单读取编号列表。自动识别编号列。"""
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str)
    else:
        df = pd.read_excel(path, dtype=str)

    cfg = oc.get_cert_type(cert_type)
    label = cfg["label"] if cfg else ""

    # 找编号列：列名含编号/证号/注册号/专利号/信用代码等关键词
    num_col = None
    keywords = ["编号", "证号", "注册号", "专利号", "信用代码", "登记号", "证书号"]
    for col in df.columns:
        col_str = str(col)
        if any(k in col_str for k in keywords):
            num_col = col
            break

    if num_col is None and len(df.columns) >= 3:
        # 兜底：取第3列（常见结构：序号、权利人、编号）
        num_col = df.columns[2]

    if num_col is None:
        raise ValueError("无法识别 Excel 中的编号列，请确保含'编号/证号'等关键词列名")

    numbers = []
    for v in df[num_col].dropna():
        s = str(v).strip()
        if s and s not in ("nan", "None", ""):
            numbers.append(s)
    return numbers, str(num_col)


def extract_numbers_from_pdf(path: Path, cert_type, engine):
    """从扫描版 PDF 查册中 OCR 提取所有编号。返回 (numbers, is_empty)。"""
    cfg = oc.get_cert_type(cert_type)
    empty_marks = cfg["empty_mark"] if cfg else []
    import fitz

    numbers = []
    doc = fitz.open(str(path))
    all_text = ""
    for pi in range(doc.page_count):
        img, _ = oc.render_pdf_page(str(path), pi)
        if img is None:
            continue
        page_text = oc.ocr_page_text(img, engine)
        all_text += page_text + "\n"

        # 空查册检测
        if any(m in page_text for m in empty_marks):
            doc.close()
            return [], True

        # 提取编号
        m_main, m_all = oc.extract_cert_number_from_text(page_text, cert_type)
        for n in m_all:
            if n not in numbers:
                numbers.append(n)
    doc.close()

    return numbers, False


def collect_base_numbers(base_dir: Path, cert_type):
    """收集底稿文件夹中所有文件的编号（从文件名提取）。"""
    numbers = {}
    for f in base_dir.rglob("*"):
        if f.is_file() and f.suffix.lower() == ".pdf":
            num = oc.extract_cert_number_from_filename(f.name, cert_type)
            if num:
                norm = oc.normalize_number(num)
                numbers[norm] = f
    return numbers


def main():
    args = parse_args()
    checklist_path = Path(args.checklist)
    base_dir = Path(args.base_dir)
    output = Path(args.output) if args.output else base_dir / "查册交叉比对结果.xlsx"
    cert_type = args.type

    cfg = oc.get_cert_type(cert_type)
    if cfg is None:
        print(f"[警告] 未知证书类型 '{cert_type}'")
        cfg_label = cert_type
    else:
        cfg_label = cfg["label"]

    print("=== 初始化 RapidOCR ===")
    engine = oc.get_ocr_engine()
    print("初始化完成\n")

    # 1. 收集查册编号
    checklist_numbers = []  # (原始编号, 规范化编号, 来源)
    empty_marked = False

    files = collect_files(checklist_path, (".pdf", ".xlsx", ".xls", ".csv"))
    print(f"查册文件数: {len(files)}")

    for f in files:
        if is_excel(f):
            nums, col = extract_numbers_from_excel(f, cert_type)
            for n in nums:
                checklist_numbers.append((n, oc.normalize_number(n), f"{f.name}:{col}"))
            print(f"  [Excel] {f.name} 读取编号 {len(nums)} 条（列: {col}）")
        else:
            nums, is_empty = extract_numbers_from_pdf(f, cert_type, engine)
            if is_empty:
                empty_marked = True
                print(f"  [PDF] {f.name} 识别为空查册（名下无此证书）")
            else:
                for n in nums:
                    checklist_numbers.append((n, oc.normalize_number(n), f.name))
                print(f"  [PDF] {f.name} OCR 提取编号 {len(nums)} 条")

    # 2. 收集底稿编号
    base_numbers = collect_base_numbers(base_dir, cert_type)
    print(f"\n底稿文件夹编号数: {len(base_numbers)}")

    # 3. 交叉比对
    results = []
    for idx, (raw, norm, source) in enumerate(checklist_numbers):
        row = {
            "序号": idx + 1,
            "查册编号": raw,
            "来源": source,
            "底稿是否有": "",
            "匹配方式": "",
            "匹配文件": "",
        }
        if norm is None:
            row["底稿是否有"] = "编号无法识别"
            results.append(row)
            continue

        if norm in base_numbers:
            row["底稿是否有"] = "有"
            row["匹配方式"] = "精确匹配"
            row["匹配文件"] = str(base_numbers[norm])
        else:
            # 前缀/后缀容错
            matched = None
            matched_way = None
            for bnum, bpath in base_numbers.items():
                if bnum.startswith(norm) or norm.startswith(bnum):
                    matched = bpath
                    matched_way = "前缀/后缀容错"
                    break
            if matched:
                row["底稿是否有"] = "有(容错)"
                row["匹配方式"] = matched_way
                row["匹配文件"] = str(matched)
            else:
                row["底稿是否有"] = "缺失"
                row["匹配方式"] = "未匹配"

        results.append(row)

    # 4. 输出
    df = pd.DataFrame(results)
    df.to_excel(str(output), index=False, engine="openpyxl")

    n_missing = sum(1 for r in results if r["底稿是否有"] == "缺失")
    n_found = sum(1 for r in results if r["底稿是否有"] in ("有", "有(容错)"))

    print("\n=== 交叉比对完成 ===")
    print(f"查册编号总数: {len(results)}  底稿有: {n_found}  缺失: {n_missing}")
    if empty_marked:
        print("注意: 存在空查册（名下无此证书），已单独标记")
    print(f"报告: {output}")

    if n_missing:
        print("\n缺失编号:")
        for r in results:
            if r["底稿是否有"] == "缺失":
                print(f"  - {r['查册编号']} (来源: {r['来源']})")


if __name__ == "__main__":
    main()

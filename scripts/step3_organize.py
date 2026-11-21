# -*- coding: utf-8 -*-
"""
Step 3 — 清单匹配 + 按编号归档（复制不移动）
===========================================
按用户提供的清单，将源文件夹中的底稿文件（PDF/JPG/PNG）逐项匹配核对，
生成核对 Excel，并按编号统一命名复制到归档目录（源文件保持不动）。

用法（命令行）:
    python step3_organize.py [清单CSV/XLSX] [源文件夹] [输出Excel] [归档目录]

用法（作为库导入）:
    from step3_organize import check_and_organize
    result = check_and_organize(checklist, source_dirs, out_excel, out_dir)

清单格式（CSV/XLSX 需含表头）:
    序号,权利人,编号[,其他字段...]
"""
import os
import shutil
import sys
import argparse
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    os.system("pip install openpyxl -q")
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))


def build_source_index(source_dirs, key_extractor=None):
    """
    递归扫描源文件夹，建立 key -> (full_path, ext) 索引。

    参数:
        source_dirs: list of str — 源文件夹路径列表（递归子目录）
        key_extractor: callable(filename) -> str — 从文件名提取匹配 key。
                       默认取文件名（不含扩展名）。返回 None 则跳过。
    返回:
        dict {key: (full_path, real_ext)}

    规则:
        - 支持 .pdf / .jpg / .jpeg / .png
        - 同一 key 多格式时优先 pdf
    """
    index = {}

    def _scan_dir(d):
        if not os.path.isdir(d):
            return
        try:
            entries = os.listdir(d)
        except PermissionError:
            return
        for fname in entries:
            full = os.path.join(d, fname)
            if os.path.isdir(full):
                _scan_dir(full)
            elif os.path.isfile(full):
                ext = os.path.splitext(fname)[1].lower()
                if ext not in ('.pdf', '.jpg', '.jpeg', '.png'):
                    continue
                if key_extractor:
                    key = key_extractor(fname)
                    if key is None:
                        continue
                else:
                    key = os.path.splitext(fname)[0]
                existing = index.get(key)
                if existing is None:
                    index[key] = (full, ext)
                elif ext == '.pdf' and existing[1] in ('.jpg', '.jpeg', '.png'):
                    index[key] = (full, ext)  # pdf 优先

    for d in source_dirs:
        _scan_dir(d)
    return index


def check_and_organize(
    checklist,
    source_dirs,
    out_excel,
    out_dir,
    *,
    headers=None,
    sheet_name="底稿核对",
    key_extractor=None,
    overseas_checklist=None,
    overseas_out_dir=None,
    overseas_headers=None,
    overseas_sheet_name="境外底稿核对",
    extra_dir=None,
):
    """
    全流程：核对 + 生成 Excel + 按编号归档（复制，源文件不动）。

    参数:
        checklist: list of tuple — 每项 (序号, 权利人, 编号, *其他字段...)
        source_dirs: list of str — 源文件夹路径列表
        out_excel: str — 输出 Excel 路径
        out_dir: str — 归档输出目录
        headers: list of str — Excel 表头（不含自动追加的"底稿状态""底稿文件路径"两列）
        sheet_name: str — Sheet 名称
        key_extractor: callable(identifier_str) -> match_key — 清单编号转匹配 key
        overseas_checklist / overseas_out_dir / overseas_headers / overseas_sheet_name:
            境外清单及输出（可选）
        extra_dir: str — 多余文件存放目录（None 则不处理）
    返回:
        dict {total, found, missing, copied, extra_count, ...}
    """
    os.makedirs(out_dir, exist_ok=True)
    source_index = build_source_index(source_dirs, key_extractor=None)

    def _match_key(identifier):
        mk = identifier
        if key_extractor:
            m = key_extractor(identifier)
            if m is not None:
                mk = m
        return mk

    found, missing, used_keys = [], [], set()
    for item in checklist:
        match_key = _match_key(item[2])
        src_info = source_index.get(match_key)
        if src_info:
            found.append((*item, src_info[0], src_info[1]))
            used_keys.add(match_key)
        else:
            missing.append(item)

    # ---- 样式 ----
    hfont = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
    hfill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    tfont = Font(name='微软雅黑', bold=True, size=14)
    nfont = Font(name='微软雅黑', size=10)
    gfill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    rfill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    calign = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def _style(ws, r, ncols, font=nfont, fill=None):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font
            cell.alignment = calign
            cell.border = border
            if fill:
                cell.fill = fill

    def _make_sheet(wb, name, title, data_list, all_items, hdrs):
        ws = wb.create_sheet(name)
        full_headers = list(hdrs) + ["底稿状态", "底稿文件路径"]
        ncols = len(full_headers)
        last_col = openpyxl.utils.get_column_letter(ncols)

        ws.merge_cells(f'A1:{last_col}1')
        ws.cell(row=1, column=1, value=title)
        ws.cell(row=1, column=1).font = tfont
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')

        for i, h in enumerate(full_headers, 1):
            ws.cell(row=2, column=i, value=h)
        _style(ws, 2, ncols, font=hfont, fill=hfill)

        found_map = {item[0]: item[-2] for item in data_list}

        for idx, item in enumerate(all_items):
            row = idx + 3
            for ci, val in enumerate(item[:len(hdrs)]):
                ws.cell(row=row, column=ci + 1, value=val)
            src_path = found_map.get(item[0])
            if src_path:
                ws.cell(row=row, column=ncols - 1, value="有")
                ws.cell(row=row, column=ncols, value=src_path)
                _style(ws, row, ncols, fill=gfill)
            else:
                ws.cell(row=row, column=ncols - 1, value="无")
                _style(ws, row, ncols, fill=rfill)

        for i in range(len(hdrs)):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = \
                max(6, min(40, len(str(hdrs[i])) * 2 + 4))
        ws.column_dimensions[openpyxl.utils.get_column_letter(ncols - 1)].width = 10
        ws.column_dimensions[openpyxl.utils.get_column_letter(ncols)].width = 90

        sr = len(all_items) + 4
        found_n = len(data_list)
        miss_n = len(all_items) - found_n
        ws.merge_cells(f'A{sr}:{openpyxl.utils.get_column_letter(ncols - 2)}{sr}')
        ws.cell(row=sr, column=1, value=f"合计: {len(all_items)}条 | 有底稿: {found_n}条 | 缺失: {miss_n}条")
        ws.cell(row=sr, column=1).font = Font(name='微软雅黑', bold=True, size=11,
                                              color='C00000' if miss_n else '2F5496')
        ws.cell(row=sr, column=1).alignment = Alignment(horizontal='center', vertical='center')
        return found_n, miss_n

    # ---- 生成 Excel ----
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if headers is None:
        headers = ["序号", "权利人", "编号"]
    _make_sheet(wb, sheet_name, "底稿核对", found, checklist, headers)

    if missing:
        ws = wb.create_sheet("缺失清单")
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 80
        ws.cell(row=1, column=1, value="缺失底稿清单").font = tfont
        for i, m in enumerate(missing):
            ws.cell(row=i + 2, column=1, value=i + 1)
            ws.cell(row=i + 2, column=2, value=f"序号{m[0]}: {m[1]} | {m[2]}")
            _style(ws, i + 2, 2, fill=rfill)

    extra_keys = {k: v for k, v in source_index.items() if k not in used_keys}
    if extra_keys:
        ws = wb.create_sheet("多余文件")
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 110
        ws.cell(row=1, column=1, value="源文件夹中多余的文件（不在清单中）").font = \
            Font(name='微软雅黑', bold=True, size=11)
        for i, (k, (path, ext)) in enumerate(sorted(extra_keys.items())):
            ws.cell(row=i + 2, column=1, value=k)
            ws.cell(row=i + 2, column=2, value=path)

    # ---- 境外 ----
    ov_total = ov_found = ov_miss = 0
    if overseas_checklist:
        ov_found_items, ov_missing_items, ov_used = [], [], set()
        for item in overseas_checklist:
            match_key = _match_key(item[2])
            src_info = source_index.get(match_key)
            if src_info:
                ov_found_items.append((*item, src_info[0], src_info[1]))
                ov_used.add(match_key)
            else:
                ov_missing_items.append(item)
        ov_total, ov_found, ov_miss = len(overseas_checklist), len(ov_found_items), len(ov_missing_items)
        if overseas_headers is None:
            overseas_headers = list(headers)
        _make_sheet(wb, overseas_sheet_name, "境外底稿核对",
                    ov_found_items, overseas_checklist, overseas_headers)

    os.makedirs(os.path.dirname(out_excel), exist_ok=True)
    wb.save(out_excel)

    # ---- 归档（复制，源文件不动） ----
    copied = 0

    def _copy_items(items, target_dir, seq_width=3):
        nonlocal copied
        for item in items:
            seq, owner, identifier = item[0], item[1], item[2]
            safe_owner = str(owner).replace("/", "_").replace(" ", "")
            src_info = source_index.get(_match_key(identifier))
            if src_info:
                src_path, real_ext = src_info
                dst_name = f"{seq:0{seq_width}d}-{safe_owner}-{identifier}{real_ext}"
                shutil.copy2(src_path, os.path.join(target_dir, dst_name))
                copied += 1
            else:
                note_name = f"{seq:0{seq_width}d}-{safe_owner}-{identifier}_【缺失】.txt"
                with open(os.path.join(target_dir, note_name), 'w', encoding='utf-8') as nf:
                    nf.write(f"缺失底稿\n序号: {seq}\n权利人: {owner}\n编号: {identifier}\n")

    _copy_items(checklist, out_dir, 3)
    if overseas_checklist and overseas_out_dir:
        os.makedirs(overseas_out_dir, exist_ok=True)
        _copy_items(overseas_checklist, overseas_out_dir, 2)

    extra_copied = 0
    if extra_dir and extra_keys:
        os.makedirs(extra_dir, exist_ok=True)
        for idx, (key, (src_path, ext)) in enumerate(sorted(extra_keys.items()), 1):
            shutil.copy2(src_path, os.path.join(extra_dir, f"{idx:02d}-{key}{ext}"))
            extra_copied += 1

    return {
        "total": len(checklist),
        "found": len(found),
        "missing": len(missing),
        "missing_items": missing,
        "extra_count": len(extra_keys),
        "copied": copied,
        "overseas_total": ov_total,
        "overseas_found": ov_found,
        "overseas_missing": ov_miss,
        "extra_copied": extra_copied,
        "out_excel": out_excel,
        "out_dir": out_dir,
    }


def parse_args():
    p = argparse.ArgumentParser(description="清单匹配 + 按编号归档（复制不移动）")
    p.add_argument("checklist", help="清单文件（CSV/XLSX，含 序号/权利人/编号 列）")
    p.add_argument("source_dir", help="源文件夹")
    p.add_argument("output", nargs="?", default=None, help="输出 Excel 路径")
    p.add_argument("out_dir", nargs="?", default=None, help="归档目录")
    return p.parse_args()


def main():
    args = parse_args()
    cl_path = Path(args.checklist)
    source_dir = Path(args.source_dir)
    out_excel = args.output or str(source_dir / "底稿核对.xlsx")
    out_dir = args.out_dir or str(source_dir / "_归档")

    if cl_path.suffix.lower() == ".csv":
        df = pd.read_csv(cl_path, dtype=str)
    else:
        df = pd.read_excel(cl_path, dtype=str)

    headers = list(df.columns)
    checklist = [tuple(row) for row in df.itertuples(index=False, name=None)]

    result = check_and_organize(
        checklist, [str(source_dir)], out_excel, out_dir,
        headers=headers,
    )

    print("\n=== 归档完成 ===")
    print(f"总数: {result['total']}  有底稿: {result['found']}  缺失: {result['missing']}")
    print(f"归档复制: {result['copied']} 个文件（源文件未移动）")
    print(f"核对 Excel: {out_excel}")
    print(f"归档目录: {out_dir}")


if __name__ == "__main__":
    main()

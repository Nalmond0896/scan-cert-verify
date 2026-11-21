# -*- coding: utf-8 -*-
"""
scan-cert-verify 公共层
========================
RapidOCR 引擎单例 + 五类编号正则配置（可扩展）+ PDF 渲染 + 编号提取函数。

三个步骤脚本（step1/step2/step3）均从此模块导入公共能力。

编号类型扩展：
    向 CERT_TYPES 字典新增一项即可，同时更新 references/cert_regex_config.md。
    每一项包含：
        - label      类型中文名（用于报告展示）
        - file_re    从【文件名】提取编号的正则（含捕获组）
        - ocr_re     从【OCR 文本】提取编号的正则（含捕获组）
        - empty_mark 查册中表示"该主体名下无此类证书"的空结果标识（可多个）
"""
import os
import re
from pathlib import Path

import numpy as np
import fitz  # PyMuPDF

_ocr_engine = None


def get_ocr_engine():
    """RapidOCR 引擎单例，避免重复初始化（初始化约 2-3 秒）。"""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


# ============================================================
# 编号类型配置（默认五类，可扩展）
# ============================================================
CERT_TYPES = {
    "realestate": {
        "label": "不动产权证",
        # 文件名：不动产权第0372958号.pdf / 粤（2018）不动产权第0372959号.pdf
        "file_re": re.compile(r"不动产权第\s*(\d{5,7})\s*号"),
        # OCR 内页：不动产权第 0372958 号（可能带空格）
        "ocr_re": re.compile(r"不动产权第\s*(\d{6,7})"),
        "empty_mark": ["名下无不动产", "无不动产", "名下无房产"],
    },
    "patent": {
        "label": "专利证书",
        # 文件名：ZL202318765432.1.pdf / ZL 2023 0275890.6.pdf
        "file_re": re.compile(r"(ZL\s*[\d\s\.]+)"),
        # OCR 内页：专利号：ZL 2023 1 8765432.1（专利号后可能有空格、换行）
        "ocr_re": re.compile(r"专利号[：:]\s*(ZL[\d\s\.]*)"),
        "empty_mark": ["无专利", "名下无专利"],
    },
    "trademark": {
        "label": "商标注册证",
        # 文件名：第87654321号.pdf
        "file_re": re.compile(r"第\s*(\d{7,10})\s*号"),
        # OCR 内页：注册号：第 87654321 号
        "ocr_re": re.compile(r"注册号[：:]\s*第?\s*(\d{7,10})\s*号?"),
        "empty_mark": ["无商标", "名下无商标"],
    },
    "business_license": {
        "label": "营业执照",
        # 统一社会信用代码：91310000MA5XXXXXX7（18位字母数字）
        "file_re": re.compile(r"([0-9A-HJ-NP-RTUWXY]{18})"),
        "ocr_re": re.compile(r"统一社会信用代码[：:]\s*([0-9A-HJ-NP-RTUWXY]{18})"),
        "empty_mark": ["未查询到", "无登记信息"],
    },
    "qualification": {
        "label": "资质证书",
        # 资质证书编号格式各异，兜底匹配"编号：XXX"或"证书编号：XXX"
        "file_re": re.compile(r"编号[：:]?\s*([A-Za-z0-9\-\/]{6,})"),
        "ocr_re": re.compile(r"(?:证书)?编号[：:]\s*([A-Za-z0-9\-\/]{6,})"),
        "empty_mark": ["未查询到", "无相关资质"],
    },
}

# 统一社会信用代码校验位字符集（去掉了 I O S V Z 等易混淆字符）
# 仅用于 business_license 类型，其他类型不适用


def list_cert_types():
    """返回所有已注册的编号类型 key 列表。"""
    return list(CERT_TYPES.keys())


def get_cert_type(key):
    """按 key 取配置，未知 key 返回 None。"""
    return CERT_TYPES.get(key)


def detect_cert_type(text):
    """
    根据文本内容自动探测证书类型（返回 key 列表，可能为空）。
    用于 Step2 自动识别查册/清单里是什么类型的编号。
    """
    detected = []
    for key, cfg in CERT_TYPES.items():
        if cfg["ocr_re"].search(text):
            detected.append(key)
    return detected


# ============================================================
# PDF 渲染
# ============================================================
def render_pdf_page(pdf_path, page_index=0, scale=2):
    """
    渲染 PDF 指定页为 numpy 图像数组（RGB）。
    返回 (img_array, page)，page 供读取旋转信息等。
    """
    doc = fitz.open(str(pdf_path))
    if page_index >= doc.page_count:
        doc.close()
        return None, None
    page = doc[page_index]
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    # 统一转为 RGB（处理灰度/透明通道）
    if pix.n == 4:
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 4)
        img = img[:, :, :3]
    elif pix.n == 3:
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3)
    elif pix.n == 1:
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
        img = np.stack([img] * 3, axis=-1)
    else:
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    # 暂存 doc 引用不关闭，由调用方负责 close
    # 这里返回 (img, doc, page) 由调用方管理 doc 生命周期更安全
    return img, doc


# ============================================================
# OCR 文本提取
# ============================================================
def ocr_page_text(img_array, engine=None):
    """
    对单页图像做 OCR，返回纯文本（按坐标排序拼接）。
    """
    if engine is None:
        engine = get_ocr_engine()
    result, _ = engine(img_array)
    if not result:
        return ""
    # result: list of (box, text, conf)
    sorted_items = sorted(result, key=lambda x: (x[0][0][1], x[0][0][0]))
    return "\n".join(item[1] for item in sorted_items)


# ============================================================
# 编号提取
# ============================================================
def extract_cert_number_from_filename(filename, cert_type="realestate"):
    """从文件名提取编号（含捕获组），未匹配返回 None。"""
    cfg = get_cert_type(cert_type)
    if cfg is None:
        return None
    m = cfg["file_re"].search(filename)
    if not m:
        return None
    return m.group(1)


def extract_cert_number_from_text(text, cert_type="realestate"):
    """
    从 OCR 文本提取编号。返回 (main_num, all_candidates)。
    main_num 为最可能的编号，all_candidates 为该类型正则命中的所有候选。
    """
    cfg = get_cert_type(cert_type)
    if cfg is None or not text:
        return None, []
    matches = cfg["ocr_re"].findall(text)
    # 去重保序
    unique = list(dict.fromkeys(matches))
    main_num = unique[0] if unique else None
    return main_num, unique


def normalize_number(num_str):
    """规范化编号：去空格、去中文括号内可能的干扰，统一用于比对。"""
    if num_str is None:
        return None
    return re.sub(r"\s+", "", str(num_str)).strip()


def is_probably_year(num_str):
    """判断数字是否更像年份而非证书编号。"""
    try:
        n = int(re.sub(r"\D", "", str(num_str)))
        return 2000 <= n <= 2030
    except (ValueError, TypeError):
        return False

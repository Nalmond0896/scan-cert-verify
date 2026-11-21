# scan-cert-verify

投行尽调扫描件底稿通用核查工具 —— 用 RapidOCR 批量识别扫描版 PDF 中的编号，与清单/文件名交叉核对，输出 Excel 报告，并按编号归档。

## 能做什么

| 步骤 | 脚本 | 功能 |
|------|------|------|
| Step 1 | `scripts/step1_cert_ocr.py` | OCR 提取证书编号 → 与文件名比对 |
| Step 2 | `scripts/step2_cross_check.py` | 查册/清单 vs 底稿交叉比对，找缺口 |
| Step 3 | `scripts/step3_organize.py` | 清单匹配 + 按编号归档（复制不移动） |

三步可一条龙连跑，也可单独调用。

## 支持编号类型（默认五类，可扩展）

不动产证、专利证书、商标注册证、营业执照（统一社会信用代码）、资质证书。正则在 `scripts/ocr_common.py` 的 `CERT_TYPES` 字典中，加一项即可扩展新类型。

## 快速开始

```bash
pip install rapidocr-onnxruntime PyMuPDF pandas openpyxl numpy

# Step 1：编号 OCR 核查
python scripts/step1_cert_ocr.py [底稿文件夹] [输出Excel] --type realestate

# Step 2：查册交叉比对（查册支持扫描PDF / Excel清单）
python scripts/step2_cross_check.py [查册] [底稿文件夹] [输出Excel] --type realestate

# Step 3：清单归档
python scripts/step3_organize.py [清单.xlsx] [源文件夹] [输出Excel] [归档目录]
```

## 文档

- [SKILL.md](SKILL.md) —— 技能定义
- [使用说明.md](使用说明.md) —— 完整使用手册
- [references/cert_regex_config.md](references/cert_regex_config.md) —— 编号正则速查表

## 设计原则

**宁可多报，不可漏报。** 模糊匹配单独标注、OCR 异常区段整段提示、候补编号保留供人工复核。

## 环境

- Python 3.13 实测通过
- OCR 引擎：RapidOCR（不用 PaddleOCR/easyocr，避免 numpy ABI 冲突）

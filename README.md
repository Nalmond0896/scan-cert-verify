# scan-cert-verify

**扫描件底稿通用核查工具** —— 用 RapidOCR 批量识别扫描版 PDF 中的编号，与清单/文件名交叉核对，输出 Excel 报告，并按编号归档。

> 背景：尽调底稿（不动产权证、专利证书、商标注册证、营业执照、资质证书等）多为"扫描全能王"等 App 生成的扫描版 PDF，**无文字层**，传统 PDF 解析读不出内容，必须 OCR。

## 功能

| 步骤 | 脚本 | 功能 |
|------|------|------|
| Step 1 | `scripts/step1_cert_ocr.py` | OCR 提取证书编号 → 与文件名比对 |
| Step 2 | `scripts/step2_cross_check.py` | 查册/清单 vs 底稿交叉比对，找缺口 |
| Step 3 | `scripts/step3_organize.py` | 清单匹配 + 按编号归档（复制不移动） |

三步可一条龙连跑，也可单独调用任意一步。

## 支持编号类型（默认五类，可扩展）

| 类型 | 编号示例 |
|------|---------|
| 不动产权证 | `不动产权第 0372958 号` |
| 专利证书 | `专利号：ZL202318765432.1` |
| 商标注册证 | `第 87654321 号` |
| 营业执照 | `91310000MA5XXXXXX7`（统一社会信用代码） |
| 资质证书 | 视证书而定 |

新增类型：在 `scripts/ocr_common.py` 的 `CERT_TYPES` 字典加一项即可（正则配置集中管理）。

## 安装

```bash
pip install rapidocr-onnxruntime PyMuPDF pandas openpyxl numpy
```

> 使用 **RapidOCR**（PP-OCRv4 的 ONNX 移植版），零依赖冲突。不推荐 PaddleOCR/easyocr：PaddleOCR 在 Python 3.13 + numpy 2.x 下有 ABI 冲突，easyocr 模型下载慢易断。

## 快速开始

```bash
# Step 1：编号 OCR 核查（文件名 vs 证书内页编号）
python scripts/step1_cert_ocr.py [底稿文件夹] [输出Excel] --type realestate

# Step 2：查册交叉比对（查册支持扫描PDF / Excel清单）
python scripts/step2_cross_check.py [查册] [底稿文件夹] [输出Excel] --type realestate

# Step 3：清单归档（复制不移动源文件）
python scripts/step3_organize.py [清单.xlsx] [源文件夹] [输出Excel] [归档目录]
```

`--type` 可选：`realestate` / `patent` / `trademark` / `business_license` / `qualification`。

## 目录结构

```
scan-cert-verify/
├── SKILL.md                          # 技能定义
├── 使用说明.md                        # 完整使用手册
├── scripts/
│   ├── ocr_common.py                 # 公共层：RapidOCR 引擎 + 编号正则配置
│   ├── step1_cert_ocr.py             # Step 1 编号 OCR 核查
│   ├── step2_cross_check.py          # Step 2 查册交叉比对
│   └── step3_organize.py             # Step 3 清单归档
└── references/
    └── cert_regex_config.md          # 编号正则速查表
```

## 设计原则

**宁可多报，不可漏报。** 模糊匹配单独标注匹配方式、OCR 异常区段整段提示人工复核、候补编号保留供 Ctrl+F 验证。

## 环境

- Python 3.13 实测通过
- OCR 引擎：RapidOCR（CPU 模式，每份证书约 3-5 秒）

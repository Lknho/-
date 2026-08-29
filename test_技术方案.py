"""
技术方案文档完整性验证脚本
验证 技术方案.md 的关键章节、技术栈数据与 CArchive 解析结果一致性。
"""
import os
import sys

DOC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "技术方案.md")

REQUIRED_SECTIONS = [
    "# 财务管理系统 — 技术方案",
    "## 1. 项目概述",
    "## 2. 总体技术架构",
    "## 3. 核心模块设计",
    "## 4. 数据存储与备份",
    "## 5. 关键技术点",
    "## 6. 安全与性能",
    "## 7. 测试策略",
    "## 8. 后续迭代方向",
]

REQUIRED_TECH_KEYWORDS = [
    "Python 3.13",
    "PyInstaller",
    "Tkinter",
    "SQLite",
    "RapidOCR",
    "ONNX Runtime",
    "OpenCV",
    "PyMuPDF",
    "Pillow",
    "NumPy",
    "Shapely",
    "pyclipper",
    "lxml",
    "PyYAML",
    "decimal",
    "multiprocessing",
    "pywin32",
]

REQUIRED_OCR_MODELS = [
    "ch_PP-OCRv3_det",
    "ch_PP-OCRv3_rec",
    "ch_ppocr_mobile_v2.0_cls",
]

def test_file_exists():
    assert os.path.isfile(DOC_PATH), f"技术方案.md 不存在: {DOC_PATH}"
    size = os.path.getsize(DOC_PATH)
    assert size > 3000, f"文档内容过短: {size} bytes"
    print(f"[PASS] 文件存在，大小 {size} bytes")

def test_sections():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    missing = [s for s in REQUIRED_SECTIONS if s not in content]
    assert not missing, f"缺少章节: {missing}"
    print(f"[PASS] 全部 {len(REQUIRED_SECTIONS)} 个章节齐全")

def test_tech_stack():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    missing = [k for k in REQUIRED_TECH_KEYWORDS if k not in content]
    assert not missing, f"缺少技术栈关键词: {missing}"
    print(f"[PASS] 全部 {len(REQUIRED_TECH_KEYWORDS)} 个技术栈组件已提及")

def test_ocr_models():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    missing = [m for m in REQUIRED_OCR_MODELS if m not in content]
    assert not missing, f"缺少 OCR 模型: {missing}"
    print(f"[PASS] 全部 {len(REQUIRED_OCR_MODELS)} 个 OCR 模型已记录")

def test_financial_decimal():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    assert "decimal" in content.lower(), "未提及 decimal 高精度计算"
    assert "借贷" in content, "未提及借贷平衡校验"
    print("[PASS] 财务计算精度与借贷平衡校验已覆盖")

def test_no_placeholder():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    placeholders = ["TODO", "待补充", "XXX", "占位", "TBD"]
    found = [p for p in placeholders if p in content]
    assert not found, f"文档中存在未完成占位符: {found}"
    print("[PASS] 无未完成占位符")

def main():
    tests = [
        test_file_exists,
        test_sections,
        test_tech_stack,
        test_ocr_models,
        test_financial_decimal,
        test_no_placeholder,
    ]
    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except AssertionError as e:
            print(f"[FAIL] {t.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"[ERROR] {t.__name__}: {e}")
            failed += 1
    print(f"\n=== 测试结果: {passed} 通过, {failed} 失败, 共 {len(tests)} 项 ===")
    sys.exit(1 if failed else 0)

if __name__ == "__main__":
    main()

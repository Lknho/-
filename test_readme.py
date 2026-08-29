"""
README.md 完整性验证脚本
检查 GitHub 项目说明的关键章节与必要信息是否齐全。
"""
import os
import sys

DOC_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md")

REQUIRED_SECTIONS = [
    "# 财务管理系统",
    "## 功能特性",
    "## 技术栈",
    "## 快速开始",
    "## 项目结构",
    "## 开发指南",
    "## 测试",
    "## 打包发布",
    "## 贡献指南",
    "## 许可证",
]

REQUIRED_KEYWORDS = [
    "RapidOCR",
    "ONNX Runtime",
    "Tkinter",
    "SQLite",
    "OpenCV",
    "PyMuPDF",
    "decimal",
    "PyInstaller",
    "Python 3.13",
    "离线",
]

def test_file_exists():
    assert os.path.isfile(DOC_PATH), f"README.md 不存在: {DOC_PATH}"
    size = os.path.getsize(DOC_PATH)
    assert size > 2000, f"README 内容过短: {size} bytes"
    print(f"[PASS] 文件存在，大小 {size} bytes")

def test_sections():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    missing = [s for s in REQUIRED_SECTIONS if s not in content]
    assert not missing, f"缺少章节: {missing}"
    print(f"[PASS] 全部 {len(REQUIRED_SECTIONS)} 个章节齐全")

def test_keywords():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    missing = [k for k in REQUIRED_KEYWORDS if k not in content]
    assert not missing, f"缺少关键信息: {missing}"
    print(f"[PASS] 全部 {len(REQUIRED_KEYWORDS)} 个关键词已覆盖")

def test_about_description():
    """检查开头是否有一句话项目描述（GitHub About 用）"""
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        head = f.read(500)
    assert ">" in head, "开头缺少引用块形式的一句话描述"
    assert len(head.split(">")[1].strip()) > 20, "一句话描述过短"
    print("[PASS] 包含一句话项目描述（可用于 GitHub About）")

def test_no_placeholder():
    with open(DOC_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    placeholders = ["TODO", "待补充", "XXX", "TBD", "你的用户名"]
    found = [p for p in placeholders if p in content]
    # "你的用户名" 是 git clone 示例中的占位，允许存在但需提示
    hard_placeholders = [p for p in found if p != "你的用户名"]
    assert not hard_placeholders, f"存在未完成占位符: {hard_placeholders}"
    print("[PASS] 无未完成占位符（git clone 示例中的用户名占位属正常）")

def main():
    tests = [
        test_file_exists,
        test_sections,
        test_keywords,
        test_about_description,
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

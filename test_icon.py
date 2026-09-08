"""
图标资源验证脚本
验证 icon.jpg、icon.ico 文件完整性，以及桌面快捷方式配置。
"""
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def test_icon_jpg_exists():
    path = os.path.join(BASE_DIR, "icon.jpg")
    assert os.path.isfile(path), "icon.jpg 不存在"
    size = os.path.getsize(path)
    assert size > 10000, f"icon.jpg 过小: {size} bytes"
    print(f"[PASS] icon.jpg 存在 ({size} bytes)")

def test_icon_ico_exists():
    path = os.path.join(BASE_DIR, "icon.ico")
    assert os.path.isfile(path), "icon.ico 不存在"
    size = os.path.getsize(path)
    assert size > 10000, f"icon.ico 过小: {size} bytes"
    print(f"[PASS] icon.ico 存在 ({size} bytes)")

def test_icon_ico_valid():
    """验证 ICO 文件格式正确，包含多个尺寸"""
    path = os.path.join(BASE_DIR, "icon.ico")
    with open(path, "rb") as f:
        header = f.read(6)
    # ICO 文件头: reserved(2)=0, type(2)=1, count(2)
    reserved = int.from_bytes(header[0:2], "little")
    ico_type = int.from_bytes(header[2:4], "little")
    count = int.from_bytes(header[4:6], "little")
    assert reserved == 0, f"ICO reserved 字段错误: {reserved}"
    assert ico_type == 1, f"ICO type 字段错误: {ico_type} (应为1)"
    assert count >= 4, f"ICO 包含尺寸数过少: {count} (应>=4)"
    print(f"[PASS] icon.ico 格式有效，包含 {count} 个尺寸")

def test_icon_ico_with_pillow():
    """用 Pillow 验证 ICO 可正常读取"""
    try:
        from PIL import Image
    except ImportError:
        print("[SKIP] Pillow 未安装，跳过 Pillow 验证")
        return
    path = os.path.join(BASE_DIR, "icon.ico")
    img = Image.open(path)
    assert img.format == "ICO", f"格式错误: {img.format}"
    w, h = img.size
    assert w >= 16 and h >= 16, f"尺寸过小: {w}x{h}"
    assert w == h, f"图标应为正方形: {w}x{h}"
    # 验证可加载像素数据
    img.load()
    print(f"[PASS] Pillow 验证 icon.ico 可正常读取，主尺寸: {w}x{h}")

def test_desktop_shortcut():
    """验证桌面快捷方式存在且指向正确的 exe"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    lnk = os.path.join(desktop, "财务管理系统.lnk")
    # Windows 快捷方式验证
    if sys.platform == "win32":
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk)
            assert os.path.exists(shortcut.TargetPath), f"快捷方式目标不存在: {shortcut.TargetPath}"
            assert "财务管理系统.exe" in shortcut.TargetPath, "快捷方式目标不正确"
            assert "icon.ico" in shortcut.IconLocation, "快捷方式图标未设置"
            print(f"[PASS] 桌面快捷方式有效，目标: {shortcut.TargetPath}")
            print(f"       图标: {shortcut.IconLocation}")
        except ImportError:
            # win32com 不可用，只检查文件存在
            assert os.path.exists(lnk), "桌面快捷方式不存在"
            print(f"[PASS] 桌面快捷方式存在 (win32com 不可用，仅验证文件存在)")
    else:
        print("[SKIP] 非 Windows 平台，跳过快捷方式验证")

def test_exe_integrity():
    """验证 exe 文件存在且 PyInstaller overlay 完整"""
    exe_path = r"D:\财务\财务管理系统.exe"
    assert os.path.isfile(exe_path), "exe 文件不存在"
    size = os.path.getsize(exe_path)
    assert size > 100 * 1024 * 1024, f"exe 过小: {size} bytes (应>100MB)"
    # 验证 PyInstaller CArchive MAGIC
    with open(exe_path, "rb") as f:
        # 只读末尾部分搜索 MAGIC
        f.seek(max(0, size - 1024 * 1024))
        tail = f.read()
    magic = b"MEI\x0c\x0b\x0a\x0b\x0e"
    assert magic in tail, "未找到 PyInstaller CArchive MAGIC，overlay 可能损坏"
    print(f"[PASS] exe 完整性验证通过 ({size/1024/1024:.2f} MB, overlay 完整)")

def test_exe_icon_resource():
    """验证 exe 中包含可提取的图标资源"""
    if sys.platform != "win32":
        print("[SKIP] 非 Windows 平台，跳过 exe 图标验证")
        return
    exe_path = r"D:\财务\财务管理系统.exe"
    assert os.path.isfile(exe_path), "exe 文件不存在"
    import ctypes
    from ctypes import wintypes
    shell32 = ctypes.windll.shell32
    shell32.ExtractIconExW.argtypes = [
        wintypes.LPCWSTR, ctypes.c_int,
        ctypes.POINTER(wintypes.HICON),
        ctypes.POINTER(wintypes.HICON),
        wintypes.UINT
    ]
    shell32.ExtractIconExW.restype = wintypes.UINT
    user32 = ctypes.windll.user32
    user32.DestroyIcon.argtypes = [wintypes.HICON]
    n = shell32.ExtractIconExW(exe_path, -1, None, None, 0)
    assert n > 0, "exe 中没有图标资源"
    hicon = wintypes.HICON()
    extracted = shell32.ExtractIconExW(exe_path, 0, ctypes.byref(hicon), None, 1)
    assert extracted > 0 and hicon.value, "无法从 exe 提取图标"
    user32.DestroyIcon(hicon)
    print(f"[PASS] exe 图标资源验证通过 (可提取 {n} 个图标)")

def main():
    tests = [
        test_icon_jpg_exists,
        test_icon_ico_exists,
        test_icon_ico_valid,
        test_icon_ico_with_pillow,
        test_desktop_shortcut,
        test_exe_integrity,
        test_exe_icon_resource,
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

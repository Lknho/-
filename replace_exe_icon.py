"""
修改 PyInstaller onefile exe 的图标，保留 overlay 数据。
方法：
1. 解析 PE 头，找到 PE 数据末尾（overlay 开始位置）
2. 提取 overlay data
3. 对 PE 部分执行 UpdateResource 修改图标
4. 把 overlay data 附加回去
"""
import struct
import os
import ctypes
from ctypes import wintypes
import tempfile
import shutil

exe_path = r"D:\财务\财务管理系统.exe"
ico_path = r"D:\text\财务管理工具\icon.ico"

kernel32 = ctypes.windll.kernel32
kernel32.BeginUpdateResourceW.argtypes = [wintypes.LPCWSTR, wintypes.BOOL]
kernel32.BeginUpdateResourceW.restype = wintypes.HANDLE
kernel32.UpdateResourceW.argtypes = [wintypes.HANDLE, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.WORD, wintypes.LPCVOID, wintypes.DWORD]
kernel32.UpdateResourceW.restype = wintypes.BOOL
kernel32.EndUpdateResourceW.argtypes = [wintypes.HANDLE, wintypes.BOOL]
kernel32.EndUpdateResourceW.restype = wintypes.BOOL

RT_ICON_PTR = ctypes.cast(3, wintypes.LPCWSTR)
RT_GROUP_ICON_PTR = ctypes.cast(14, wintypes.LPCWSTR)
GROUP_ID_PTR = ctypes.cast(1, wintypes.LPCWSTR)

# ========== 1. 解析 PE 头，找到 overlay 开始位置 ==========
print("=== 解析 PE 头 ===")
with open(exe_path, "rb") as f:
    # DOS 头
    dos_header = f.read(64)
    e_lfanew = struct.unpack_from("<I", dos_header, 0x3C)[0]
    print(f"PE 头偏移: 0x{e_lfanew:X}")

    # PE 签名
    f.seek(e_lfanew)
    pe_sig = f.read(4)
    assert pe_sig == b"PE\x00\x00", "不是有效的 PE 文件"

    # COFF 头
    coff_header = f.read(20)
    num_sections = struct.unpack_from("<H", coff_header, 2)[0]
    size_opt_header = struct.unpack_from("<H", coff_header, 16)[0]
    print(f"节区数: {num_sections}, 可选头大小: {size_opt_header}")

    # 跳过可选头
    f.seek(e_lfanew + 24 + size_opt_header)

    # 读取节区表，计算 PE 数据末尾
    pe_end = 0
    for i in range(num_sections):
        section = f.read(40)
        name = section[0:8].rstrip(b"\x00").decode("ascii", errors="replace")
        virt_size = struct.unpack_from("<I", section, 8)[0]
        virt_addr = struct.unpack_from("<I", section, 12)[0]
        raw_size = struct.unpack_from("<I", section, 16)[0]
        raw_ptr = struct.unpack_from("<I", section, 20)[0]
        section_end = raw_ptr + raw_size
        if section_end > pe_end:
            pe_end = section_end
        print(f"  节区 {name}: raw_ptr=0x{raw_ptr:X} raw_size=0x{raw_size:X} end=0x{section_end:X}")

print(f"PE 数据末尾（overlay 开始）: 0x{pe_end:X} ({pe_end} bytes)")
file_size = os.path.getsize(exe_path)
overlay_size = file_size - pe_end
print(f"文件总大小: {file_size} bytes")
print(f"Overlay 大小: {overlay_size} bytes ({overlay_size/1024/1024:.2f} MB)")

# ========== 2. 提取 overlay data ==========
print("\n=== 提取 overlay ===")
with open(exe_path, "rb") as f:
    f.seek(pe_end)
    overlay_data = f.read()
print(f"Overlay 已提取: {len(overlay_data)} bytes")
# 验证 PyInstaller MAGIC (在 overlay 末尾附近)
# CArchive MAGIC: MEI\x0c\x0b\x0a\x0b\x0e
magic = b"MEI\x0c\x0b\x0a\x0b\x0e"
magic_pos = overlay_data.rfind(magic)
if magic_pos >= 0:
    print(f"PyInstaller CArchive MAGIC 找到于 overlay 偏移 0x{magic_pos:X}")
else:
    print("警告: 未找到 PyInstaller CArchive MAGIC")

# ========== 3. 写入临时 PE 文件（不含 overlay）==========
print("\n=== 创建临时 PE 文件 ===")
tmp_pe = tempfile.mktemp(suffix=".exe")
with open(exe_path, "rb") as f:
    pe_data = f.read(pe_end)
with open(tmp_pe, "wb") as f:
    f.write(pe_data)
print(f"临时 PE 文件: {tmp_pe} ({len(pe_data)} bytes)")

# ========== 4. 解析 ICO ==========
print("\n=== 解析 ICO ===")
with open(ico_path, "rb") as f:
    ico_data = f.read()
count = struct.unpack_from("<H", ico_data, 4)[0]
print(f"包含 {count} 个尺寸")

icons = []
offset = 6
for i in range(count):
    w, h, cc, res, planes, bit_count, bytes_in_res, image_offset = struct.unpack_from("<BBBBHHII", ico_data, offset)
    image_data = ico_data[image_offset:image_offset + bytes_in_res]
    icons.append(image_data)
    dw = 256 if w == 0 else w
    print(f"  [{i+1}] {dw}x{dw} {bit_count}bit {bytes_in_res}B")
    offset += 16

# 构建 GROUP_ICON
group_data = struct.pack("<HHH", 0, 1, count)
offset = 6
for i in range(count):
    w = ico_data[offset]
    h = ico_data[offset+1]
    cc = ico_data[offset+2]
    planes = struct.unpack_from("<H", ico_data, offset+4)[0]
    bit_count = struct.unpack_from("<H", ico_data, offset+6)[0]
    bytes_in_res = struct.unpack_from("<I", ico_data, offset+8)[0]
    group_data += struct.pack("<BBBBHHIH", w, h, cc, 0, planes, bit_count, bytes_in_res, i+1)
    offset += 16

# ========== 5. 修改临时 PE 的图标 ==========
print("\n=== 修改临时 PE 图标 ===")
h = kernel32.BeginUpdateResourceW(tmp_pe, False)
if not h:
    raise ctypes.WinError()
try:
    for i, data in enumerate(icons):
        icon_id_ptr = ctypes.cast(i+1, wintypes.LPCWSTR)
        buf = ctypes.create_string_buffer(data)
        ok = kernel32.UpdateResourceW(h, RT_ICON_PTR, icon_id_ptr, 0, buf, len(data))
        if not ok:
            raise RuntimeError(f"RT_ICON {i+1} 失败: {ctypes.GetLastError()}")
    gbuf = ctypes.create_string_buffer(group_data)
    ok = kernel32.UpdateResourceW(h, RT_GROUP_ICON_PTR, GROUP_ID_PTR, 0, gbuf, len(group_data))
    if not ok:
        raise RuntimeError(f"RT_GROUP_ICON 失败: {ctypes.GetLastError()}")
    ok = kernel32.EndUpdateResourceW(h, False)
    if not ok:
        raise RuntimeError(f"EndUpdateResource 失败: {ctypes.GetLastError()}")
    print("临时 PE 图标修改成功")
except Exception as e:
    kernel32.EndUpdateResourceW(h, True)
    os.remove(tmp_pe)
    raise

modified_pe_size = os.path.getsize(tmp_pe)
print(f"修改后 PE 大小: {modified_pe_size} bytes")

# ========== 6. 附加 overlay，替换原 exe ==========
print("\n=== 附加 overlay 并替换原 exe ===")
with open(tmp_pe, "rb") as f:
    modified_pe = f.read()

final_data = modified_pe + overlay_data
print(f"最终文件大小: {len(final_data)} bytes (PE {len(modified_pe)} + overlay {len(overlay_data)})")

# 先写入临时文件，再替换
tmp_final = tempfile.mktemp(suffix=".exe")
with open(tmp_final, "wb") as f:
    f.write(final_data)

# 替换原 exe
shutil.move(tmp_final, exe_path)
os.remove(tmp_pe)

print(f"\n=== 完成！ ===")
print(f"文件: {exe_path}")
print(f"大小: {os.path.getsize(exe_path)} bytes ({os.path.getsize(exe_path)/1024/1024:.2f} MB)")

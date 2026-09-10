"""
测试竖版图片缩放居中问题
模拟实际场景：竖版发票在110%缩放时
- 图片宽度476 < 窗口宽度1130（水平方向未超出）
- 图片高度674 > 窗口高度631（垂直方向超出）
- 预期：水平方向居中偏移，垂直方向使用滚动条
"""

def calculate_offset(dw, dh, frame_w, frame_h, actual_cw, actual_ch):
    """模拟新的偏移计算逻辑（分别处理水平和垂直方向）"""
    if dw <= actual_cw:
        offset_x = (actual_cw - dw) // 2
    else:
        offset_x = 0
    if dh <= actual_ch:
        offset_y = (actual_ch - dh) // 2
    else:
        offset_y = 0
    return (offset_x, offset_y)

def calculate_offset_old(dw, dh, frame_w, frame_h):
    """模拟旧的偏移计算逻辑（有bug的版本）"""
    if dw <= frame_w and dh <= frame_h:
        return ((frame_w - dw) // 2, (frame_h - dh) // 2)
    else:
        return (0, 0)

print("=" * 70)
print("测试1：竖版图片在110%缩放时（实际场景）")
print("=" * 70)
dw, dh = 476, 674
frame_w, frame_h = 1130, 631
actual_cw, actual_ch = 1115, 631

old_offset = calculate_offset_old(dw, dh, frame_w, frame_h)
new_offset = calculate_offset(dw, dh, frame_w, frame_h, actual_cw, actual_ch)

print(f"图片大小: {dw}x{dh}")
print(f"水平方向: dw={dw} <= actual_cw={actual_cw}? {dw <= actual_cw} (未超出，应居中)")
print(f"垂直方向: dh={dh} > actual_ch={actual_ch}? {dh > actual_ch} (超出，应滚动)")
print(f"旧逻辑偏移: {old_offset} (水平方向靠左！)")
print(f"新逻辑偏移: {new_offset} (水平方向居中！)")

expected_x = (actual_cw - dw) // 2
expected_y = 0
assert new_offset[0] == expected_x, f"水平偏移错误：期望{expected_x}，实际{new_offset[0]}"
assert new_offset[1] == expected_y, f"垂直偏移错误：期望{expected_y}，实际{new_offset[1]}"
print("测试1通过：竖版图片在110%缩放时水平方向居中，垂直方向使用滚动条")

print()
print("=" * 70)
print("测试2：横版图片在110%缩放时")
print("=" * 70)
dw, dh = 1200, 500
frame_w, frame_h = 1130, 631
actual_cw, actual_ch = 1130, 616

old_offset = calculate_offset_old(dw, dh, frame_w, frame_h)
new_offset = calculate_offset(dw, dh, frame_w, frame_h, actual_cw, actual_ch)

print(f"图片大小: {dw}x{dh}")
print(f"水平方向: dw={dw} > actual_cw={actual_cw}? {dw > actual_cw} (超出，应滚动)")
print(f"垂直方向: dh={dh} <= actual_ch={actual_ch}? {dh <= actual_ch} (未超出，应居中)")
print(f"旧逻辑偏移: {old_offset} (垂直方向靠上！)")
print(f"新逻辑偏移: {new_offset} (垂直方向居中！)")

expected_x = 0
expected_y = (actual_ch - dh) // 2
assert new_offset[0] == expected_x
assert new_offset[1] == expected_y
print("测试2通过：横版图片在110%缩放时水平方向使用滚动条，垂直方向居中")

print()
print("=" * 70)
print("测试3：图片完全未超出窗口时")
print("=" * 70)
dw, dh = 433, 613
frame_w, frame_h = 1130, 631
actual_cw, actual_ch = 1130, 631

old_offset = calculate_offset_old(dw, dh, frame_w, frame_h)
new_offset = calculate_offset(dw, dh, frame_w, frame_h, actual_cw, actual_ch)

print(f"图片大小: {dw}x{dh}")
print(f"旧逻辑偏移: {old_offset}")
print(f"新逻辑偏移: {new_offset}")

expected_x = (actual_cw - dw) // 2
expected_y = (actual_ch - dh) // 2
assert new_offset[0] == expected_x
assert new_offset[1] == expected_y
print("测试3通过：图片完全未超出窗口时水平和垂直方向都居中")

print()
print("=" * 70)
print("测试4：图片完全超出窗口时")
print("=" * 70)
dw, dh = 1500, 1000
frame_w, frame_h = 1130, 631
actual_cw, actual_ch = 1115, 616

old_offset = calculate_offset_old(dw, dh, frame_w, frame_h)
new_offset = calculate_offset(dw, dh, frame_w, frame_h, actual_cw, actual_ch)

print(f"图片大小: {dw}x{dh}")
print(f"旧逻辑偏移: {old_offset}")
print(f"新逻辑偏移: {new_offset}")

expected_x = 0
expected_y = 0
assert new_offset[0] == expected_x
assert new_offset[1] == expected_y
print("测试4通过：图片完全超出窗口时水平和垂直方向都使用滚动条")

print()
print("=" * 70)
print("所有测试通过！")
print("=" * 70)

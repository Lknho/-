"""
测试 split_invoice_image 单张发票自动裁剪功能
验证：一页一张的扫描PDF不再返回整页，而是自动检测票据区域并裁剪空白边距
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PIL import Image, ImageDraw
import importlib.util


def load_module():
    spec = importlib.util.spec_from_file_location(
        'finance_app',
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'finance_app.py')
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def make_test_image(w=1240, h=1754, inv_rect=(120, 200, 1120, 1400)):
    """创建模拟A4扫描图：白色背景+中间票据区域+文字线条"""
    img = Image.new('RGB', (w, h), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    l, t, r, b = inv_rect
    draw.rectangle([l, t, r, b], fill=(245, 245, 240), outline=(100, 100, 100), width=3)
    for y in range(t + 50, b - 50, 40):
        draw.line([(l + 30, y), (r - 30, y)], fill=(80, 80, 80), width=2)
    return img


def test_single_invoice_not_full_page():
    """单张发票时，返回的裁剪框不应是整页"""
    mod = load_module()
    img = make_test_image()
    w, h = img.size
    result, boxes = mod.split_invoice_image(img, return_boxes=True)
    assert len(result) == 1, f"应返回1张发票，实际{len(result)}"
    assert len(boxes) == 1, f"应返回1个框，实际{len(boxes)}"
    x1, y1, x2, y2 = boxes[0]
    is_full = (x1 == 0 and y1 == 0 and x2 == w and y2 == h)
    assert not is_full, "单张发票不应返回整页框"
    print(f"[PASS] 单张发票裁剪框: ({x1},{y1},{x2},{y2})，非整页")


def test_single_invoice_crop_accuracy():
    """单张发票时，裁剪框应接近票据区域（误差<=30px）"""
    mod = load_module()
    inv_rect = (120, 200, 1120, 1400)
    img = make_test_image(inv_rect=inv_rect)
    result, boxes = mod.split_invoice_image(img, return_boxes=True)
    x1, y1, x2, y2 = boxes[0]
    margin = 30
    assert abs(x1 - inv_rect[0]) <= margin, f"left误差过大: {x1} vs {inv_rect[0]}"
    assert abs(y1 - inv_rect[1]) <= margin, f"top误差过大: {y1} vs {inv_rect[1]}"
    assert abs(x2 - inv_rect[2]) <= margin, f"right误差过大: {x2} vs {inv_rect[2]}"
    assert abs(y2 - inv_rect[3]) <= margin, f"bottom误差过大: {y2} vs {inv_rect[3]}"
    print(f"[PASS] 裁剪框精度: 误差均<=30px")


def test_cropped_image_size():
    """裁剪后图像尺寸应与裁剪框一致"""
    mod = load_module()
    img = make_test_image()
    result, boxes = mod.split_invoice_image(img, return_boxes=True)
    x1, y1, x2, y2 = boxes[0]
    expected_w, expected_h = x2 - x1, y2 - y1
    actual_w, actual_h = result[0].size
    assert actual_w == expected_w, f"宽度不符: {actual_w} vs {expected_w}"
    assert actual_h == expected_h, f"高度不符: {actual_h} vs {expected_h}"
    print(f"[PASS] 裁剪后图像尺寸: {actual_w}x{actual_h}，与裁剪框一致")


def test_return_boxes_false():
    """return_boxes=False时只返回图像列表"""
    mod = load_module()
    img = make_test_image()
    result = mod.split_invoice_image(img, return_boxes=False)
    assert isinstance(result, list), "应返回列表"
    assert len(result) == 1, "应返回1张图像"
    assert isinstance(result[0], Image.Image), "元素应为PIL Image"
    print(f"[PASS] return_boxes=False返回图像列表")


if __name__ == '__main__':
    test_single_invoice_not_full_page()
    test_single_invoice_crop_accuracy()
    test_cropped_image_size()
    test_return_boxes_false()
    print("\n全部4项测试通过!")

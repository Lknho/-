"""测试一页一张PDF的发票区域检测（轮廓检测）"""
import sys
import os
sys.path.insert(0, r'D:\text\财务管理工具')

from finance_app import pdf_to_image, split_invoice_image

def test_single_invoice_contour_detection():
    """测试A4扫描件中单张发票的轮廓检测，裁剪框不应是整页"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份办公费.pdf'
    if not os.path.exists(pdf_path):
        print(f"跳过：测试PDF不存在 {pdf_path}")
        return True

    # 测试前3页
    all_pass = True
    for page_num in range(3):
        img = pdf_to_image(pdf_path, page_num=page_num, dpi=150)
        if img is None:
            print(f"第{page_num+1}页：转图像失败")
            continue

        orig_w, orig_h = img.size
        result, boxes = split_invoice_image(img, return_boxes=True)

        if not boxes:
            print(f"第{page_num+1}页：未检测到裁剪框")
            all_pass = False
            continue

        x1, y1, x2, y2 = boxes[0]
        crop_w = x2 - x1
        crop_h = y2 - y1
        crop_area_ratio = (crop_w * crop_h) / (orig_w * orig_h)

        print(f"第{page_num+1}页：原图{orig_w}x{orig_h}, 裁剪框({x1},{y1},{x2},{y2}), "
              f"尺寸{crop_w}x{crop_h}, 占比{crop_area_ratio:.1%}")

        # 裁剪框不应是整页（占比应小于95%）
        if crop_area_ratio >= 0.95:
            print(f"  失败：裁剪框占比{crop_area_ratio:.1%}，几乎是整页")
            all_pass = False
        # 裁剪框不应太小（占比应大于20%）
        elif crop_area_ratio < 0.20:
            print(f"  警告：裁剪框占比{crop_area_ratio:.1%}，可能过小")
        else:
            print(f"  通过：裁剪框占比{crop_area_ratio:.1%}，合理")

    return all_pass

def test_contour_helper():
    """测试_detect_invoice_contour辅助函数"""
    from finance_app import _detect_invoice_contour
    import cv2
    import numpy as np

    # 创建一个模拟A4扫描件：白色背景，中间一个蓝色矩形（模拟发票）
    w, h = 1240, 1754
    img = np.ones((h, w, 3), dtype=np.uint8) * 255
    # 发票区域：居中，占页面约70%宽，60%高
    inv_x1, inv_y1 = 180, 300
    inv_x2, inv_y2 = 1060, 1400
    cv2.rectangle(img, (inv_x1, inv_y1), (inv_x2, inv_y2), (200, 220, 240), -1)
    # 发票边框
    cv2.rectangle(img, (inv_x1, inv_y1), (inv_x2, inv_y2), (100, 100, 100), 3)
    # 发票内部一些文字线条
    for i in range(10):
        y = inv_y1 + 80 + i * 80
        cv2.line(img, (inv_x1 + 50, y), (inv_x2 - 50, y), (150, 150, 150), 2)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    result = _detect_invoice_contour(img, binary, w, h)
    if result is None:
        print("辅助函数测试：未检测到轮廓")
        return False

    x1, y1, x2, y2 = result
    print(f"辅助函数测试：检测到区域({x1},{y1},{x2},{y2}), 期望约({inv_x1},{inv_y1},{inv_x2},{inv_y2})")

    # 允许±30px误差
    tolerance = 30
    if (abs(x1 - inv_x1) <= tolerance and abs(y1 - inv_y1) <= tolerance and
        abs(x2 - inv_x2) <= tolerance and abs(y2 - inv_y2) <= tolerance):
        print("  通过：检测区域与预期一致")
        return True
    else:
        print(f"  失败：检测区域与预期偏差过大")
        return False

if __name__ == '__main__':
    print("=== 测试1：辅助函数轮廓检测 ===")
    r1 = test_contour_helper()
    print()
    print("=== 测试2：真实PDF单张发票检测 ===")
    r2 = test_single_invoice_contour_detection()
    print()
    if r1 and r2:
        print("全部测试通过！")
    else:
        print("部分测试失败！")
        sys.exit(1)

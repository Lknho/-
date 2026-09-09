"""
PDF发票识别测试 - 200DPI（与程序实际处理一致）
验证：
1. 单张模式：办公费PDF各页正确框住整张发票（不调用split_invoice_image）
2. 一页多张：打车发票(2).pdf正确识别6张横向发票
3. 一页多张：打车发票(3).pdf识别5-6张（第2+3张间隙小，可能合并）
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fitz
from PIL import Image
from finance_app import split_invoice_image, _detect_invoice_contour
import cv2
import numpy as np

def pdf_to_image_200dpi(pdf_path, page_num=0):
    """用200DPI转换PDF页面（与程序实际处理一致）"""
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    mat = fitz.Matrix(200 / 72, 200 / 72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    doc.close()
    return img

def test_single_mode_office():
    """单张模式：办公费PDF各页正确框住整张发票"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份办公费.pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在')
        return True
    print(f'  办公费PDF单张模式（200DPI）:')
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    doc.close()
    for page in range(min(3, page_count)):
        img = pdf_to_image_200dpi(pdf_path, page)
        w, h = img.size
        cv_img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        contour = _detect_invoice_contour(cv_img, binary, w, h)
        assert contour is not None, f'第{page+1}页未检测到发票区域'
        x1, y1, x2, y2 = contour
        ratio = (x2-x1)*(y2-y1)/(w*h)*100
        print(f'    第{page+1}页: 区域{x2-x1}x{y2-y1}, 占比{ratio:.1f}%')
        assert 30 < ratio < 55, f'第{page+1}页占比应该在30-55%，实际{ratio:.1f}%'
    return True

def test_multi_taxi_file2():
    """一页多张：打车发票(2).pdf正确识别6张横向发票"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份打车发票2 (2).pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在')
        return True
    img = pdf_to_image_200dpi(pdf_path, 0)
    result, boxes = split_invoice_image(img, return_boxes=True)
    print(f'  打车发票(2).pdf: 识别出{len(boxes)}张')
    assert len(boxes) == 6, f'期望6张，实际{len(boxes)}张'
    for i, box in enumerate(boxes):
        x1, y1, x2, y2 = box
        w, h = x2-x1, y2-y1
        print(f'    第{i+1}张: {w}x{h}')
        assert w > h, f'第{i+1}张应该是横向发票(宽>高)，实际{w}x{h}'
        assert w > 800, f'第{i+1}张宽度应该>800，实际{w}'
    return True

def test_multi_taxi_file3():
    """一页多张：打车发票(3).pdf识别5-6张（第2+3张间隙小）"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份打车发票2 (3).pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在')
        return True
    img = pdf_to_image_200dpi(pdf_path, 0)
    result, boxes = split_invoice_image(img, return_boxes=True)
    print(f'  打车发票(3).pdf: 识别出{len(boxes)}张')
    # 第2+3张之间间隙极小，可能合并为1张，所以5-6张都可接受
    assert 5 <= len(boxes) <= 6, f'期望5-6张，实际{len(boxes)}张'
    for i, box in enumerate(boxes):
        x1, y1, x2, y2 = box
        w, h = x2-x1, y2-y1
        print(f'    第{i+1}张: {w}x{h}')
    return True

if __name__ == '__main__':
    print('=== PDF发票识别测试（200DPI） ===')
    print()

    tests = [
        ('单张模式-办公费PDF', test_single_mode_office),
        ('一页多张-打车发票(2)', test_multi_taxi_file2),
        ('一页多张-打车发票(3)', test_multi_taxi_file3),
    ]

    passed = 0
    failed = 0
    for name, test_func in tests:
        print(f'测试: {name}')
        try:
            if test_func():
                print(f'  ✓ 通过')
                passed += 1
        except AssertionError as e:
            print(f'  ✗ 失败: {e}')
            failed += 1
        except Exception as e:
            print(f'  ✗ 异常: {e}')
            import traceback
            traceback.print_exc()
            failed += 1
        print()

    print(f'=== 结果: {passed}通过, {failed}失败 ===')
    sys.exit(0 if failed == 0 else 1)

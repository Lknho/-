"""
一页多张PDF发票识别测试 - 打车发票横向纵向排列场景
验证：
1. 一页6张横向打车发票纵向排列，能正确识别出6张
2. 每张发票尺寸合理（宽>高，横向发票）
3. 办公费PDF单张模式不受影响
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from finance_app import pdf_to_image, split_invoice_image

def test_taxi_invoices_file1():
    """文件1: 一页6张横向打车发票纵向排列"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份打车发票2 (2).pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在 {pdf_path}')
        return True
    img = pdf_to_image(pdf_path, page_num=0, dpi=150)
    result, boxes = split_invoice_image(img, return_boxes=True)
    print(f'  文件1识别出{len(boxes)}张发票')
    assert len(boxes) == 6, f'期望6张，实际{len(boxes)}张'
    for i, box in enumerate(boxes):
        x1, y1, x2, y2 = box
        w, h = x2 - x1, y2 - y1
        print(f'    第{i+1}张: {w}x{h}')
        assert w > h, f'第{i+1}张应该是横向发票(宽>高)，实际{w}x{h}'
        assert w > 500, f'第{i+1}张宽度应该>500，实际{w}'
        assert h > 200, f'第{i+1}张高度应该>200，实际{h}'
    return True

def test_taxi_invoices_file2():
    """文件2: 一页6张横向打车发票纵向排列（间隙更小）"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份打车发票2 (3).pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在 {pdf_path}')
        return True
    img = pdf_to_image(pdf_path, page_num=0, dpi=150)
    result, boxes = split_invoice_image(img, return_boxes=True)
    print(f'  文件2识别出{len(boxes)}张发票')
    assert len(boxes) == 6, f'期望6张，实际{len(boxes)}张'
    for i, box in enumerate(boxes):
        x1, y1, x2, y2 = box
        w, h = x2 - x1, y2 - y1
        print(f'    第{i+1}张: {w}x{h}')
        assert w > h, f'第{i+1}张应该是横向发票(宽>高)，实际{w}x{h}'
        assert w > 500, f'第{i+1}张宽度应该>500，实际{w}'
        assert h > 200, f'第{i+1}张高度应该>200，实际{h}'
    return True

def test_office_expense_single_mode():
    """办公费PDF单张模式不受影响"""
    pdf_path = r'C:\Users\waiter\Desktop\李堃煌\23\9月\汇总\9月份办公费.pdf'
    if not os.path.exists(pdf_path):
        print(f'  跳过: 文件不存在 {pdf_path}')
        return True
    print(f'  办公费PDF单张模式验证:')
    for page in range(3):
        img = pdf_to_image(pdf_path, page_num=page, dpi=150)
        result, boxes = split_invoice_image(img, return_boxes=True)
        w, h = img.size
        assert len(boxes) == 1, f'第{page+1}页期望1张，实际{len(boxes)}张'
        x1, y1, x2, y2 = boxes[0]
        ratio = (x2-x1)*(y2-y1)/(w*h)*100
        print(f'    第{page+1}页: 1张, 占比{ratio:.1f}%')
        assert 30 < ratio < 50, f'第{page+1}页裁剪框占比应该在30-50%，实际{ratio:.1f}%'
    return True

if __name__ == '__main__':
    print('=== 一页多张PDF发票识别测试 ===')
    print()

    tests = [
        ('打车发票文件1', test_taxi_invoices_file1),
        ('打车发票文件2', test_taxi_invoices_file2),
        ('办公费PDF单张模式', test_office_expense_single_mode),
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

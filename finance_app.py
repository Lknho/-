# -*- coding: utf-8 -*-
"""
财务管理软件
功能：人员管理、发票报销、工资结算、统计汇总
"""
import os
import sys
import time
import math

# 自动设置 Tcl/Tk 库路径（必须在 import tkinter 之前）
def _setup_tcl_tk():
    if getattr(sys, 'frozen', False):
        base = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    tcl_dir = os.path.join(base, 'tcl', 'tcl8.6')
    tk_dir = os.path.join(base, 'tcl', 'tk8.6')
    if os.path.isdir(tcl_dir):
        os.environ['TCL_LIBRARY'] = tcl_dir
    if os.path.isdir(tk_dir):
        os.environ['TK_LIBRARY'] = tk_dir

_setup_tcl_tk()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import os
import sys
import shutil
import zipfile
import re
from io import BytesIO
from datetime import datetime, timedelta
from PIL import Image, ImageTk, ImageDraw
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage


# ============================================================
# 路径与数据库
# ============================================================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DB_PATH = os.path.join(get_base_dir(), 'finance.db')
INVOICE_DIR = os.path.join(get_base_dir(), 'invoices')
os.makedirs(INVOICE_DIR, exist_ok=True)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        monthly_salary REAL NOT NULL DEFAULT 0,
        resigned INTEGER DEFAULT 0,
        insurance_paid INTEGER DEFAULT 0,
        insurance_amount REAL DEFAULT 0,
        remark TEXT,
        hire_date TEXT,
        resign_date TEXT
    )''')
    # 兼容旧数据库：添加新字段
    for col, ddl in [
        ('resigned', 'ALTER TABLE employees ADD COLUMN resigned INTEGER DEFAULT 0'),
        ('insurance_paid', 'ALTER TABLE employees ADD COLUMN insurance_paid INTEGER DEFAULT 0'),
        ('insurance_amount', 'ALTER TABLE employees ADD COLUMN insurance_amount REAL DEFAULT 0'),
        ('remark', 'ALTER TABLE employees ADD COLUMN remark TEXT'),
        ('hire_date', 'ALTER TABLE employees ADD COLUMN hire_date TEXT'),
        ('resign_date', 'ALTER TABLE employees ADD COLUMN resign_date TEXT'),
    ]:
        try:
            cols = [r[1] for r in c.execute("PRAGMA table_info(employees)").fetchall()]
            if col not in cols:
                c.execute(ddl)
        except Exception:
            pass
    c.execute('''CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_number TEXT,
        invoice_date TEXT,
        reimburser_id INTEGER,
        amount REAL NOT NULL DEFAULT 0,
        purpose TEXT,
        image_path TEXT,
        status INTEGER DEFAULT 0,
        type TEXT DEFAULT '发票',
        voucher_number TEXT,
        bank_account_id INTEGER,
        remark TEXT,
        seller TEXT,
        created_at TEXT,
        FOREIGN KEY (reimburser_id) REFERENCES employees(id),
        FOREIGN KEY (bank_account_id) REFERENCES bank_accounts(id)
    )''')
    # 兼容旧数据库：添加字段
    try:
        inv_cols = [r[1] for r in c.execute("PRAGMA table_info(invoices)").fetchall()]
        if 'type' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN type TEXT DEFAULT '发票'")
        if 'voucher_number' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN voucher_number TEXT")
        if 'bank_account_id' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN bank_account_id INTEGER")
        if 'remark' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN remark TEXT")
        if 'seller' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN seller TEXT")
        if 'batch_pending' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN batch_pending INTEGER DEFAULT 0")
        if 'printed' not in inv_cols:
            c.execute("ALTER TABLE invoices ADD COLUMN printed INTEGER DEFAULT 0")
    except Exception:
        pass
    c.execute('''CREATE TABLE IF NOT EXISTS salary_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        year INTEGER,
        month INTEGER,
        amount REAL NOT NULL DEFAULT 0,
        note TEXT,
        paid_at TEXT,
        bank_account_id INTEGER,
        FOREIGN KEY (employee_id) REFERENCES employees(id),
        FOREIGN KEY (bank_account_id) REFERENCES bank_accounts(id)
    )''')
    # 兼容旧数据库：salary_payments 添加 bank_account_id
    try:
        sal_cols = [r[1] for r in c.execute("PRAGMA table_info(salary_payments)").fetchall()]
        if 'bank_account_id' not in sal_cols:
            c.execute("ALTER TABLE salary_payments ADD COLUMN bank_account_id INTEGER")
    except Exception:
        pass

    c.execute('''CREATE TABLE IF NOT EXISTS bank_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        account_number TEXT,
        bank_name TEXT,
        remark TEXT,
        created_at TEXT
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS bank_transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        account_id INTEGER NOT NULL,
        direction TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        purpose TEXT,
        source TEXT,
        source_id INTEGER,
        transaction_date TEXT,
        created_at TEXT,
        FOREIGN KEY (account_id) REFERENCES bank_accounts(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS salary_adjustments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        effective_year INTEGER NOT NULL,
        effective_month INTEGER NOT NULL,
        new_salary REAL NOT NULL DEFAULT 0,
        reason TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees(id)
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS insurance_adjustments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        effective_year INTEGER NOT NULL,
        effective_month INTEGER NOT NULL,
        new_amount REAL NOT NULL DEFAULT 0,
        reason TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees(id)
    )''')
    # 五险一金细项调整表
    c.execute('''CREATE TABLE IF NOT EXISTS insurance_detail_adjustments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER NOT NULL,
        effective_year INTEGER NOT NULL,
        effective_month INTEGER NOT NULL,
        pension_company REAL DEFAULT 0,
        pension_personal REAL DEFAULT 0,
        medical_company REAL DEFAULT 0,
        medical_personal REAL DEFAULT 0,
        unemployment_company REAL DEFAULT 0,
        unemployment_personal REAL DEFAULT 0,
        injury_company REAL DEFAULT 0,
        injury_personal REAL DEFAULT 0,
        maternity_company REAL DEFAULT 0,
        maternity_personal REAL DEFAULT 0,
        housing_company REAL DEFAULT 0,
        housing_personal REAL DEFAULT 0,
        reason TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees(id)
    )''')

    # 兼容旧数据库：添加五险一金细项基础字段到employees
    detail_cols = [
        'pension_company', 'pension_personal',
        'medical_company', 'medical_personal',
        'unemployment_company', 'unemployment_personal',
        'injury_company', 'injury_personal',
        'maternity_company', 'maternity_personal',
        'housing_company', 'housing_personal',
    ]
    try:
        existing_cols = [r[1] for r in c.execute("PRAGMA table_info(employees)").fetchall()]
        for col in detail_cols:
            if col not in existing_cols:
                c.execute(f"ALTER TABLE employees ADD COLUMN {col} REAL DEFAULT 0")
    except Exception:
        pass

    c.execute('''CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )''')
    # 默认深色主题
    c.execute("INSERT OR IGNORE INTO settings(key, value) VALUES('app_theme', 'dark')")
    conn.commit()
    conn.close()


def get_setting(key, default=None):
    """读取设置项"""
    try:
        conn = get_db()
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        conn.close()
        return row['value'] if row else default
    except Exception:
        return default


def set_setting(key, value):
    """保存设置项"""
    try:
        conn = get_db()
        conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES(?,?)", (key, str(value)))
        conn.commit()
        conn.close()
    except Exception:
        pass


# ============================================================
# 日志管理器
# ============================================================
class LogManager:
    """运行日志管理器：每次启动生成一个.log文件，记录系统信息、操作步骤、错误、运行状态。
    支持日志轮转（保留N份，超出覆盖最旧）和自定义保存路径。"""

    _instance = None
    _log_file = None
    _start_time = None
    _operation_count = 0

    @classmethod
    def init(cls):
        """初始化日志，在程序启动时调用一次"""
        if cls._instance is not None:
            return
        cls._instance = cls()
        cls._start_time = datetime.now()
        cls._operation_count = 0

        # 读取日志配置
        keep_count = int(get_setting('log_keep_count', '10'))
        log_path = get_setting('log_path', '')
        if not log_path:
            log_path = os.path.join(get_base_dir(), 'log')
        os.makedirs(log_path, exist_ok=True)

        # 日志轮转：删除超过保留数量的旧日志
        try:
            log_files = sorted([
                os.path.join(log_path, f) for f in os.listdir(log_path)
                if f.endswith('.log') and os.path.isfile(os.path.join(log_path, f))
            ], key=lambda x: os.path.getmtime(x))
            while len(log_files) >= keep_count and log_files:
                oldest = log_files.pop(0)
                try:
                    os.remove(oldest)
                except Exception:
                    pass
        except Exception:
            pass

        # 创建本次运行的日志文件
        timestamp = cls._start_time.strftime('%Y%m%d_%H%M%S')
        cls._log_file = os.path.join(log_path, f'finance_{timestamp}.log')

        # 写入日志头
        cls._write_header()

    @classmethod
    def _write(cls, level, message):
        """写入一条日志"""
        if cls._log_file is None:
            return
        try:
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(cls._log_file, 'a', encoding='utf-8') as f:
                f.write(f'[{ts}] [{level}] {message}\n')
        except Exception:
            pass

    @classmethod
    def _write_header(cls):
        """写入系统信息头"""
        cls._write('INFO', '=' * 60)
        cls._write('INFO', '财务管理系统 - 运行日志')
        cls._write('INFO', '=' * 60)
        cls._write('INFO', f'启动时间: {cls._start_time.strftime("%Y-%m-%d %H:%M:%S")}')
        cls._write('INFO', f'程序版本: v1.0')
        cls._write('INFO', f'运行模式: {"打包EXE" if getattr(sys, "frozen", False) else "源码运行"}')
        cls._write('INFO', f'程序路径: {get_base_dir()}')
        cls._write('INFO', f'Python版本: {sys.version}')
        cls._write('INFO', f'操作系统: {os.name} / {sys.platform}')
        try:
            import platform
            cls._write('INFO', f'系统详情: {platform.system()} {platform.release()} {platform.version()}')
            cls._write('INFO', f'处理器: {platform.processor()}')
        except Exception:
            pass
        try:
            # 不创建临时Tk窗口（会破坏默认root），用已有的默认root或ctypes获取
            import tkinter as _tk
            if _tk._default_root is not None:
                sw = _tk._default_root.winfo_screenwidth()
                sh = _tk._default_root.winfo_screenheight()
                cls._write('INFO', f'屏幕分辨率: {sw}x{sh}')
            else:
                try:
                    import ctypes
                    user32 = ctypes.windll.user32
                    sw = user32.GetSystemMetrics(0)
                    sh = user32.GetSystemMetrics(1)
                    cls._write('INFO', f'屏幕分辨率: {sw}x{sh}')
                except Exception:
                    pass
        except Exception:
            pass
        try:
            conn = get_db()
            emp_count = conn.execute('SELECT COUNT(*) FROM employees').fetchone()[0]
            inv_count = conn.execute('SELECT COUNT(*) FROM invoices').fetchone()[0]
            sal_count = conn.execute('SELECT COUNT(*) FROM salary_records').fetchone()[0]
            conn.close()
            cls._write('INFO', f'数据统计: 人员{emp_count}人, 发票{inv_count}条, 工资{sal_count}条')
        except Exception:
            pass
        cls._write('INFO', '-' * 60)
        cls._write('INFO', '运行记录开始')
        cls._write('INFO', '-' * 60)

    @classmethod
    def info(cls, message):
        """记录普通信息"""
        cls._write('INFO', message)

    @classmethod
    def operation(cls, action, detail=''):
        """记录用户操作"""
        cls._operation_count += 1
        msg = f'[操作#{cls._operation_count}] {action}'
        if detail:
            msg += f' - {detail}'
        cls._write('OPERATION', msg)

    @classmethod
    def error(cls, message, exception=None):
        """记录错误"""
        cls._write('ERROR', message)
        if exception is not None:
            import traceback
            tb_str = ''.join(traceback.format_exception(type(exception), exception, exception.__traceback__))
            cls._write('ERROR', f'异常详情:\n{tb_str}')

    @classmethod
    def warning(cls, message):
        """记录警告"""
        cls._write('WARNING', message)

    @classmethod
    def status(cls, message):
        """记录运行状态"""
        cls._write('STATUS', message)

    @classmethod
    def close(cls):
        """程序退出时写入结束信息"""
        if cls._log_file is None:
            return
        end_time = datetime.now()
        duration = end_time - cls._start_time
        hours, remainder = divmod(int(duration.total_seconds()), 3600)
        minutes, seconds = divmod(remainder, 60)
        cls._write('INFO', '-' * 60)
        cls._write('INFO', f'运行结束: {end_time.strftime("%Y-%m-%d %H:%M:%S")}')
        cls._write('INFO', f'运行时长: {hours}小时{minutes}分{seconds}秒')
        cls._write('INFO', f'操作总数: {cls._operation_count}')
        cls._write('INFO', '=' * 60)
        cls._log_file = None
        cls._instance = None


def log_info(msg):
    LogManager.info(msg)

def log_op(action, detail=''):
    LogManager.operation(action, detail)

def log_err(msg, exc=None):
    LogManager.error(msg, exc)

def log_warn(msg):
    LogManager.warning(msg)

def log_status(msg):
    LogManager.status(msg)




def pdf_to_image(pdf_path, page_num=0, dpi=150):
    """将PDF指定页转为PIL Image，失败返回None"""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        if page_num >= len(doc):
            page_num = 0
        page = doc[page_num]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        doc.close()
        return img
    except Exception:
        return None


def get_invoice_image(image_path):
    """根据发票附件路径返回PIL Image（PDF自动转第一页）"""
    if not image_path:
        return None
    p = os.path.join(INVOICE_DIR, image_path)
    if not os.path.exists(p):
        return None
    if p.lower().endswith('.pdf'):
        return pdf_to_image(p)
    try:
        return Image.open(p)
    except Exception:
        return None


def _try_split(binary_crop, cv_img, top, left, h_gap_min, v_gap_min, min_col_ratio, min_row_h, min_col_w, min_content):
    """尝试用给定阈值分割，返回PIL Image列表或None。
    增强版：支持 2×2/2×3/多行多列、单行多列（横排）、多行单列（竖排）等布局，
    每行独立检测列分割点，不再要求所有行列数一致。"""
    import cv2
    import numpy as np
    ch, cw = binary_crop.shape

    # 水平投影分行
    row_proj = np.sum(binary_crop, axis=1) / 255
    row_is_blank = row_proj < cw * 0.002
    h_gaps = []
    in_gap = False
    gap_start = 0
    for y in range(ch):
        if row_is_blank[y] and not in_gap:
            in_gap = True
            gap_start = y
        elif not row_is_blank[y] and in_gap:
            in_gap = False
            if y - gap_start >= h_gap_min:
                h_gaps.append((gap_start + y) // 2)

    h_lines = [0] + h_gaps + [ch]
    rows = []
    for i in range(len(h_lines) - 1):
        y1, y2 = h_lines[i], h_lines[i + 1]
        if y2 - y1 >= min_row_h:
            rows.append((y1, y2))
    row_count = len(rows)
    if not (1 <= row_count <= 8):
        return None

    # 基于合并后行方向的智能合并：只合并横着的发票行（合并后行高<内容宽*0.6），不合并竖着的发票行
    if len(rows) >= 2:
        merged_direction = []
        for y1, y2 in rows:
            if merged_direction:
                prev_y1, prev_y2 = merged_direction[-1]
                gap = y1 - prev_y2
                prev_h = prev_y2 - prev_y1
                curr_h = y2 - y1
                # 计算合并后的行高和内容宽度
                merged_h = y2 - prev_y1
                merged_region = binary_crop[prev_y1:y2, :]
                merged_col_proj = np.sum(merged_region, axis=0) / 255
                merged_content_w = np.sum(merged_col_proj > merged_h * 0.005)
                # 只合并：间隙<80px 且 行高相似 且 合并后是横着的发票行（行高<内容宽*0.6）
                is_horizontal_after_merge = merged_h < merged_content_w * 0.6 if merged_content_w > 0 else False
                if gap < 80 and abs(prev_h - curr_h) < max(prev_h, curr_h) * 0.5 and is_horizontal_after_merge:
                    merged_direction[-1] = (prev_y1, y2)
                else:
                    merged_direction.append((y1, y2))
            else:
                merged_direction.append((y1, y2))
        rows = merged_direction
        row_count = len(rows)

    # 合并过矮或内容宽度过小的行：说明是发票内部表格行，合并到相邻行
    if len(rows) >= 3:
        row_h_pre = [y2 - y1 for y1, y2 in rows]
        median_h_pre = sorted(row_h_pre)[len(row_h_pre) // 2]
        merged_rows = []
        for y1, y2 in rows:
            rh_check = y2 - y1
            row_region_check = binary_crop[y1:y2, :]
            _, rw_check = row_region_check.shape
            col_proj_check = np.sum(row_region_check, axis=0) / 255
            content_w_check = np.sum(col_proj_check > rh_check * 0.005)
            # 行高<中位数*0.7 或 内容宽度<行宽*50%，都合并到相邻行
            is_short = rh_check < median_h_pre * 0.7
            is_narrow = content_w_check < rw_check * 0.50
            if (is_short or is_narrow) and merged_rows:
                prev_y1, prev_y2 = merged_rows[-1]
                merged_rows[-1] = (prev_y1, y2)
            else:
                merged_rows.append((y1, y2))
        rows = merged_rows
        row_count = len(rows)

    # 高行二次分割：如果某行高度明显大于中位数，说明可能合并了多张发票，用更小间隙再次分割
    if len(rows) >= 2:
        row_heights = [y2 - y1 for y1, y2 in rows]
        median_h = sorted(row_heights)[len(row_heights) // 2]
        new_rows = []
        for y1, y2 in rows:
            rh = y2 - y1
            if rh > median_h * 1.5 and rh > min_row_h * 1.5:
                row_region = binary_crop[y1:y2, :]
                sub_proj = np.sum(row_region, axis=1) / 255
                sub_blank = sub_proj < cw * 0.002
                sub_gaps = []
                in_gap = False
                gap_start = 0
                for y in range(rh):
                    if sub_blank[y] and not in_gap:
                        in_gap = True
                        gap_start = y
                    elif not sub_blank[y] and in_gap:
                        in_gap = False
                        gap_len = y - gap_start
                        # 降低最小间隙到3px，过滤边缘20px内的假间隙
                        if gap_len >= max(2, h_gap_min // 4) and gap_start > 15 and y < rh - 15:
                            sub_gaps.append((gap_start + y) // 2)
                if sub_gaps:
                    sub_lines = [0] + sub_gaps + [rh]
                    added = False
                    for si in range(len(sub_lines) - 1):
                        sy1, sy2 = sub_lines[si], sub_lines[si + 1]
                        if sy2 - sy1 >= min_row_h * 0.7:
                            new_rows.append((y1 + sy1, y1 + sy2))
                            added = True
                    if not added:
                        new_rows.append((y1, y2))
                else:
                    new_rows.append((y1, y2))
            else:
                new_rows.append((y1, y2))
        rows = new_rows
        row_count = len(rows)
        if not (1 <= row_count <= 8):
            return None

    # 高行二次分割后，检查是否还有异常偏高的行，强制上下平分
    if len(rows) >= 3:
        row_h_final = [y2 - y1 for y1, y2 in rows]
        median_final = sorted(row_h_final)[len(row_h_final) // 2]
        if median_final > 100:
            final_rows = []
            for y1, y2 in rows:
                rh = y2 - y1
                if rh > median_final * 1.3 and rh > 200:
                    mid_y = (y1 + y2) // 2
                    final_rows.append((y1, mid_y))
                    final_rows.append((mid_y, y2))
                else:
                    final_rows.append((y1, y2))
            rows = final_rows
            row_count = len(rows)

    # 每行内垂直投影分列（每行独立检测，允许行列数不一致）
    row_cols = []
    for y1, y2 in rows:
        row_region = binary_crop[y1:y2, :]
        rh, rw = row_region.shape
        # 横向发票判断：行高<行宽*0.65且内容集中，才整行作为1张
        if rh < rw * 0.65:
            col_proj_tmp = np.sum(row_region, axis=0) / 255
            content_width = np.sum(col_proj_tmp > rh * 0.005)
            if content_width > rw * 0.30:
                row_cols.append([(0, rw)])
                continue
        col_proj = np.sum(row_region, axis=0) / 255
        col_is_blank = col_proj < rh * 0.0015
        # 纵向发票行（行高>行宽*0.5）使用更小的分列间隙，检测紧密排列的发票
        is_vertical_row = rh > rw * 0.5
        effective_v_gap = 5 if is_vertical_row else v_gap_min
        v_gaps = []
        in_gap = False
        gap_start = 0
        for x in range(rw):
            if col_is_blank[x] and not in_gap:
                in_gap = True
                gap_start = x
            elif not col_is_blank[x] and in_gap:
                in_gap = False
                if x - gap_start >= effective_v_gap:
                    v_gaps.append((gap_start + x) // 2)
        v_lines = [0] + v_gaps + [rw]
        cols = []
        for i in range(len(v_lines) - 1):
            col_w = v_lines[i + 1] - v_lines[i]
            # 纵向发票行列宽要求更低（10%），其他行保持15%
            min_col_w_ratio = 0.10 if is_vertical_row else max(min_col_ratio, 0.15)
            if col_w >= rw * min_col_w_ratio:
                cols.append((v_lines[i], v_lines[i + 1]))
        # 纵向发票行：合并过窄的假列（宽度<中位数*0.6）
        if is_vertical_row and len(cols) >= 3:
            col_widths = [c2 - c1 for c1, c2 in cols]
            median_col_w = sorted(col_widths)[len(col_widths) // 2]
            if median_col_w > 100:
                merged_cols = []
                for c1, c2 in cols:
                    cw = c2 - c1
                    if cw < median_col_w * 0.6 and merged_cols:
                        # 合并到上一列
                        prev_c1, prev_c2 = merged_cols[-1]
                        merged_cols[-1] = (prev_c1, c2)
                    else:
                        merged_cols.append((c1, c2))
                cols = merged_cols
        if cols:
            # 宽高比检查：仅对横向发票行（行高<行宽*0.5）检查，如果多列且每列都是纵向说明是误分割
            # 纵向发票行（行高>行宽*0.5）不检查，因为每列确实是纵向的
            if len(cols) >= 2 and rh < rw * 0.5:
                row_h = y2 - y1
                all_vertical = all((c[1] - c[0]) < row_h for c in cols)
                if all_vertical:
                    cols = [(0, rw)]
            row_cols.append(cols)

    if not row_cols:
        return None

    # 行高一致性检查（放宽到3倍，允许尺寸差异略大的混排）
    if len(rows) >= 2:
        row_heights = [y2 - y1 for y1, y2 in rows]
        if max(row_heights) / min(row_heights) > 3.0:
            return None

    # 分割：每行按各自检测到的列分割
    result = []
    boxes = []
    for (y1, y2), cols in zip(rows, row_cols):
        row_region = binary_crop[y1:y2, :]
        rh, rw = row_region.shape
        for x1, x2 in cols:
            if x2 - x1 < min_col_w:
                continue
            if np.sum(row_region[:, x1:x2]) < min_content:
                continue
            ax1, ax2 = left + x1, left + x2
            ay1, ay2 = top + y1, top + y2
            crop_cv = cv_img[ay1:ay2, ax1:ax2]
            result.append(Image.fromarray(cv2.cvtColor(crop_cv, cv2.COLOR_BGR2RGB)))
            boxes.append((ax1, ay1, ax2, ay2))
    return (result, boxes) if len(result) >= 2 else None



def _detect_invoice_contour(cv_img, binary, w, h):
    """用基于内容密度的投影法精确找发票区域，返回(left, top, right, bottom)或None。
    适用于A4扫描件/电子发票中单张发票的精确定位，避免把整页A4当成发票。
    从页面边缘向内扫描，找到内容密度超过阈值的位置作为发票边界。"""
    try:
        import numpy as np
        # 水平投影：每行的内容像素数
        row_proj = np.sum(binary, axis=1) / 255
        # 垂直投影：每列的内容像素数
        col_proj = np.sum(binary, axis=0) / 255

        # 内容密度阈值：行/列的内容像素数占宽度/高度的比例
        # 用相对较高的阈值，排除边缘噪点和零星文字
        row_threshold = w * 0.015   # 行内容超过宽度1.5%才算有内容
        col_threshold = h * 0.015   # 列内容超过高度1.5%才算有内容

        # 从顶部向下扫描，找第一个内容密度超过阈值的行
        top = 0
        for y in range(h):
            if row_proj[y] >= row_threshold:
                top = y
                break

        # 从底部向上扫描
        bottom = h - 1
        for y in range(h - 1, -1, -1):
            if row_proj[y] >= row_threshold:
                bottom = y
                break

        # 从左向右扫描
        left = 0
        for x in range(w):
            if col_proj[x] >= col_threshold:
                left = x
                break

        # 从右向左扫描
        right = w - 1
        for x in range(w - 1, -1, -1):
            if col_proj[x] >= col_threshold:
                right = x
                break

        # 验证检测到的区域合理
        crop_w = right - left
        crop_h = bottom - top
        crop_area = crop_w * crop_h
        page_area = w * h

        # 区域太小（<10%页面）或太大（>98%页面）都不合理
        if crop_area < page_area * 0.10 or crop_area > page_area * 0.98:
            return None

        # 加安全边距
        margin = 15
        left = max(0, left - margin)
        top = max(0, top - margin)
        right = min(w, right + margin)
        bottom = min(h, bottom + margin)

        return (left, top, right, bottom)
    except Exception:
        return None



def split_invoice_image(pil_img, return_boxes=False):
    """将一页多张发票的图片分割成单张发票列表，返回PIL Image列表。
    多尺度投影法：尝试严格/中等/宽松三组阈值，取分割结果最合理的。
    单张发票返回原图。分割失败返回[原图]。
    return_boxes=True时返回(图片列表, 坐标列表)。"""
    try:
        import cv2
        import numpy as np
        # 横向页面（宽>高）：旋转90度变纵向处理，再旋转坐标回来
        rotated = False
        pw, ph = pil_img.size
        if pw > ph * 1.1:
            pil_img = pil_img.rotate(90, expand=True)
            rotated = True
        cv_img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
        h, w = gray.shape

        # Otsu二值化
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        binary = cv2.morphologyEx(binary, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))

        # 裁剪边距
        row_proj = np.sum(binary, axis=1) / 255
        col_proj = np.sum(binary, axis=0) / 255
        row_has = row_proj > w * 0.002
        col_has = col_proj > h * 0.002
        if not np.any(row_has) or not np.any(col_has):
            return ([pil_img], []) if return_boxes else [pil_img]
        top = max(0, int(np.argmax(row_has)) - 15)
        bottom = min(h, h - int(np.argmax(row_has[::-1])) + 15)
        left = max(0, int(np.argmax(col_has)) - 15)
        right = min(w, w - int(np.argmax(col_has[::-1])) + 15)
        binary_crop = binary[top:bottom, left:right]

        # 多尺度尝试：严格 -> 中等 -> 宽松 -> 极宽松
        # 不使用超宽松参数（v_gap_min<12），避免把单张发票内部文字间隙误判为发票间隙
        # 合并的发票由_secondary_split_wide二次分割处理
        params_list = [
            (50, 40, 0.15, 150, 80, 800),
            (40, 30, 0.12, 120, 70, 600),
            (30, 25, 0.10, 100, 60, 500),
            (15, 20, 0.08, 80, 50, 300),
            (10, 15, 0.06, 60, 40, 200),
        ]

        candidates = []
        for params in params_list:
            ret = _try_split(binary_crop, cv_img, top, left, *params)
            if ret:
                result, boxes = ret
                if 2 <= len(result) <= 16:
                    # 放宽"巨大框"判断：仅当某张占页面面积>45%，或宽>65%且高>65%才算误判
                    has_huge = any(
                        (im.size[0] * im.size[1]) > (w * h * 0.45) or
                        (im.size[0] > w * 0.65 and im.size[1] > h * 0.65)
                        for im in result
                    )
                    if not has_huge:
                        # 验证每块是"完整的发票"而非单张发票内部表格：
                        # 每块面积>页面4.5%、宽>10%页宽、高>10%页高
                        valid_sizes = all(
                            (im.size[0] * im.size[1]) > (w * h * 0.045) and
                            im.size[0] > w * 0.10 and im.size[1] > h * 0.10
                            for im in result
                        )
                        if valid_sizes:
                            candidates.append((result, boxes))

        if candidates:
            # 错误分割检测：如果2张发票高而窄、垂直位置一致、宽度之和接近整页，
            # 说明是把一张发票内部的表格竖线误判为分割线，应合并为1张走单张检测
            filtered_candidates = []
            for result, boxes in candidates:
                if len(boxes) == 2:
                    (x1, y1, x2, y2), (x3, y3, x4, y4) = boxes
                    h1 = y2 - y1
                    h2 = y4 - y3
                    w1 = x2 - x1
                    w2 = x4 - x3
                    # 两张都很高(>60%页高)、垂直位置接近、宽度之和>70%页宽
                    if (h1 > h * 0.60 and h2 > h * 0.60 and
                        abs(y1 - y3) < h * 0.10 and abs(y2 - y4) < h * 0.10 and
                        (w1 + w2) > w * 0.50):
                        # 这是错误分割，跳过此候选
                        continue
                filtered_candidates.append((result, boxes))
            candidates = filtered_candidates

        if candidates:
            common = {4, 5, 6, 8, 9, 12}
            common_candidates = [c for c in candidates if len(c[0]) in common]
            if common_candidates:
                best = max(common_candidates, key=lambda c: len(c[0]))
            else:
                best = max(candidates, key=lambda c: len(c[0]))
            # 宽发票二次分割：如果某张发票明显比中位数宽，尝试在其中找垂直缝隙再分割
            # 迭代二次分割：整页多张时需要多次分割（先分上下/左右，再细分）
            for _ in range(6):
                new_best = _secondary_split_wide(best[0], best[1], binary, w, h, tall_h_split=rotated)
                if len(new_best[0]) == len(best[0]):
                    break
                best = new_best
            # 迭代合并（一次可能只合并1对，需要多次）
            for _ in range(3):
                new_best = _merge_narrow_invoices(best[0], best[1], w, h)
                if len(new_best[0]) == len(best[0]):
                    break
                best = new_best
            # 连通区域分析核对：投影法结果很少（<=4张）时，用连通区域法核对
            if len(best[0]) <= 4:
                cc_ret = _connected_component_split(binary_crop, cv_img, top, left, w, h)
                if cc_ret and len(cc_ret[0]) > len(best[0]) and len(cc_ret[0]) <= 16:
                    # 连通区域法结果更多，用它替换投影法结果
                    best = cc_ret
            # 横向页面：旋转坐标回来
            if rotated and return_boxes:
                orig_w, orig_h = h, w
                rot_boxes = []
                for (x1, y1, x2, y2) in best[1]:
                    nx1 = orig_w - y2
                    ny1 = x1
                    nx2 = orig_w - y1
                    ny2 = x2
                    rot_boxes.append((nx1, ny1, nx2, ny2))
                return (best[0], rot_boxes)
            return best if return_boxes else best[0]

        # fallback：单张发票，用轮廓检测精确找发票区域，回退到投影法
        inv_result = _detect_invoice_contour(cv_img, binary, w, h)
        if inv_result is not None:
            left, top, right, bottom = inv_result
        cropped_img = pil_img.crop((left, top, right, bottom))
        if rotated and return_boxes:
            orig_w = h
            nx1, ny1 = orig_w - bottom, left
            nx2, ny2 = orig_w - top, right
            return ([cropped_img], [(nx1, ny1, nx2, ny2)])
        return ([cropped_img], [(left, top, right, bottom)]) if return_boxes else [cropped_img]
    except Exception as e:
        _log_ocr_error(f"发票图像分割失败: {e}")
        return ([pil_img], []) if return_boxes else [pil_img]


def _connected_component_split(binary_crop, cv_img, top, left, w, h):
    """连通区域分析法分割发票：检测所有连通区域，按y间隙分行、按x间隙分列，
    每组的外接矩形就是一张发票的边界框。用于与投影法核对，提高紧密排列发票的识别准确率。
    分割失败返回None。"""
    try:
        import cv2
        import numpy as np
        from PIL import Image

        ch, cw = binary_crop.shape
        if ch < 100 or cw < 100:
            return None

        # 形态学膨胀：填充发票内部小间隙
        kernel = np.ones((15, 15), np.uint8)
        binary_dilated = cv2.dilate(binary_crop, kernel, iterations=1)

        # 连通区域检测
        num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(binary_dilated, connectivity=8)

        # 收集连通区域（过滤极小噪声）
        components = []
        for i in range(1, num_labels):
            x, y, comp_w, comp_h, area = stats[i]
            if area >= 500:
                components.append((x, y, x + comp_w, y + comp_h))

        if len(components) < 2:
            return None

        # 第一步：按y中心间隙分行
        components.sort(key=lambda c: (c[1] + c[3]) / 2)
        y_centers = [(c[1] + c[3]) / 2 for c in components]
        # 检测y中心间隙：大于页高8%的间隙作为分行边界
        y_gaps = []
        for i in range(1, len(y_centers)):
            gap = y_centers[i] - y_centers[i - 1]
            if gap > ch * 0.05:
                y_gaps.append(i)
        # 分行
        rows = []
        prev_idx = 0
        for gap_idx in y_gaps:
            rows.append(components[prev_idx:gap_idx])
            prev_idx = gap_idx
        rows.append(components[prev_idx:])
        # 过滤空行和只有1个区域的行（可能是噪声）
        rows = [r for r in rows if len(r) >= 1]

        if len(rows) < 1:
            return None

        # 第二步：每行内用投影法分列（能检测更小的间隙）
        boxes = []
        for row in rows:
            if len(row) == 0:
                continue
            # 计算行的y范围
            row_y1 = min(c[1] for c in row)
            row_y2 = max(c[3] for c in row)
            row_h = row_y2 - row_y1
            # 提取行区域的原始二值图像（不使用膨胀后的，避免填充间隙）
            row_binary = binary_crop[row_y1:row_y2, :]
            # 水平投影（每列的白色像素数）
            col_proj = np.sum(row_binary, axis=0) / 255
            # 检测有内容的列（降低阈值，能检测更小的间隙）
            col_has = col_proj > row_h * 0.002
            # 找连续有内容的段
            cols = []
            in_col = False
            col_start = 0
            for x in range(cw):
                if col_has[x] and not in_col:
                    col_start = x
                    in_col = True
                elif not col_has[x] and in_col:
                    cols.append((col_start, x))
                    in_col = False
            if in_col:
                cols.append((col_start, cw))
            # 合并过窄的间隙（<8px的间隙合并）
            merged_cols = []
            for cx1, cx2 in cols:
                if merged_cols and cx1 - merged_cols[-1][1] < 8:
                    merged_cols[-1] = (merged_cols[-1][0], cx2)
                else:
                    merged_cols.append((cx1, cx2))
            # 过滤过窄的列
            valid_cols = [(cx1, cx2) for cx1, cx2 in merged_cols if (cx2 - cx1) >= cw * 0.05]
            # 计算行的内容范围（使用max(检测宽度, 行宽80%)，避免稀疏内容导致估算偏少）
            if valid_cols:
                content_x1 = min(cx1 for cx1, cx2 in valid_cols)
                content_x2 = max(cx2 for cx1, cx2 in valid_cols)
                detected_w = content_x2 - content_x1
            else:
                content_x1 = content_x2 = 0
                detected_w = 0
            content_w = detected_w if detected_w > cw * 0.3 else cw * 0.8
            # 检测行方向：竖着的发票行（行高>内容宽*0.5）还是横着的发票行
            is_vertical_row = row_h > (content_w * 0.5) if content_w > 0 else False
            # 高行二次分割：同时尝试水平平分和垂直平分，选择宽高比更接近典型出租车发票(1.8)的一种
            final_cols = []
            if row_h > 600 and content_w > 200:
                # 估算发票数量
                typical_w = 400
                typical_h = 350
                n_horizontal = max(2, round(content_w / typical_w))
                n_horizontal = min(n_horizontal, 4)
                n_vertical = max(2, round(row_h / typical_h))
                n_vertical = min(n_vertical, 4)
                # 计算两种平分方式的宽高比
                ratio_horizontal = (content_w / n_horizontal) / row_h
                ratio_vertical = content_w / (row_h / n_vertical)
                # 典型出租车发票宽高比约1.4，选择更接近的一种
                if abs(ratio_vertical - 1.8) < abs(ratio_horizontal - 1.8):
                    # 垂直平分（横躺着的发票行）
                    step_y = row_h / n_vertical
                    for i in range(n_vertical):
                        sy1 = int(row_y1 + i * step_y)
                        sy2 = int(row_y1 + (i + 1) * step_y)
                        boxes.append((left + max(0, content_x1 - 8), top + max(0, sy1 - 8), left + min(cw, content_x2 + 8), top + min(ch, sy2 + 8)))
                    continue
            # 基于典型发票宽度的智能强制平分
            if is_vertical_row and content_w > 200 and len(valid_cols) < 4:
                # 竖着的发票行：根据行高区分发票类型
                # 行高>650px：出租车发票（典型宽220px），否则：定额发票（典型宽500px）
                if row_h > 650:
                    typical_w = 400
                else:
                    typical_w = 500
                n_invoices = max(2, round(content_w / typical_w))
                n_invoices = min(n_invoices, 6)
                # 智能强制平分：在估算边界附近检测实际空白列
                step = content_w / n_invoices
                boundaries = [content_x1]
                for i in range(1, n_invoices):
                    est_boundary = content_x1 + i * step
                    # 在估算边界附近±60px检测空白列
                    search_start = max(0, int(est_boundary - 60))
                    search_end = min(cw, int(est_boundary + 60))
                    search_region = binary_crop[row_y1:row_y2, search_start:search_end]
                    col_proj_search = np.sum(search_region, axis=0) / 255
                    blank_cols = np.where(col_proj_search < row_h * 0.005)[0]
                    if len(blank_cols) > 0:
                        # 找到最接近估算边界的空白列
                        actual_boundary = search_start + blank_cols[np.argmin(np.abs(blank_cols + search_start - est_boundary))]
                    else:
                        actual_boundary = int(est_boundary)
                    boundaries.append(actual_boundary)
                boundaries.append(content_x2)
                for i in range(n_invoices):
                    sx1 = boundaries[i]
                    sx2 = boundaries[i + 1]
                    final_cols.append((sx1, sx2))
            elif not is_vertical_row and content_w > 300 and len(valid_cols) < 3:
                # 横着的发票行：典型发票宽度约400px，估算发票数量
                typical_w = 400
                n_invoices = max(2, round(content_w / typical_w))
                n_invoices = min(n_invoices, 4)
                # 智能强制平分：在估算边界附近检测实际空白列
                step = content_w / n_invoices
                boundaries = [content_x1]
                for i in range(1, n_invoices):
                    est_boundary = content_x1 + i * step
                    search_start = max(0, int(est_boundary - 60))
                    search_end = min(cw, int(est_boundary + 60))
                    search_region = binary_crop[row_y1:row_y2, search_start:search_end]
                    col_proj_search = np.sum(search_region, axis=0) / 255
                    blank_cols = np.where(col_proj_search < row_h * 0.005)[0]
                    if len(blank_cols) > 0:
                        actual_boundary = search_start + blank_cols[np.argmin(np.abs(blank_cols + search_start - est_boundary))]
                    else:
                        actual_boundary = int(est_boundary)
                    boundaries.append(actual_boundary)
                boundaries.append(content_x2)
                for i in range(n_invoices):
                    sx1 = boundaries[i]
                    sx2 = boundaries[i + 1]
                    final_cols.append((sx1, sx2))
            else:
                # 正常情况：使用检测到的列
                final_cols = valid_cols
            # 每列就是一张发票
            for cx1, cx2 in final_cols:
                # 扩展边界8px
                x1 = max(0, cx1 - 8)
                y1 = max(0, row_y1 - 8)
                x2 = min(cw, cx2 + 8)
                y2 = min(ch, row_y2 + 8)
                boxes.append((left + x1, top + y1, left + x2, top + y2))

        if len(boxes) < 2:
            return None
        if len(boxes) > 16:
            return None

        # 过滤太小的框
        valid_boxes = []
        for x1, y1, x2, y2 in boxes:
            if (x2 - x1) >= w * 0.06 and (y2 - y1) >= h * 0.06:
                if (x2 - x1) * (y2 - y1) >= w * h * 0.02:
                    valid_boxes.append((x1, y1, x2, y2))

        if len(valid_boxes) < 2:
            return None

        # 生成裁剪图片
        result = []
        final_boxes = []
        for x1, y1, x2, y2 in valid_boxes:
            crop_cv = cv_img[y1:y2, x1:x2]
            if crop_cv.size == 0:
                continue
            result.append(Image.fromarray(cv2.cvtColor(crop_cv, cv2.COLOR_BGR2RGB)))
            final_boxes.append((x1, y1, x2, y2))

        if len(result) < 2:
            return None

        return (result, final_boxes)
    except Exception as e:
        return None


def _find_inner_v_gap(region, min_gap=5, margin_ratio=0.10):
    """在区域内找垂直方向的内部空白缝隙，返回最靠近中心的(gap_start, gap_end)或None。
    优先选最靠近中心的缝隙（发票通常均匀排列），而非最长的缝隙。"""
    import numpy as np
    rh, rw = region.shape
    col_proj = np.sum(region, axis=0) / 255
    col_is_blank = col_proj < rh * 0.008
    margin = int(rw * margin_ratio)
    center = rw // 2
    best_gap = None
    best_dist = float('inf')
    in_gap = False
    gap_start = 0
    for x in range(rw):
        if col_is_blank[x] and not in_gap:
            in_gap = True
            gap_start = x
        elif not col_is_blank[x] and in_gap:
            in_gap = False
            gap_len = x - gap_start
            if gap_start >= margin and x <= rw - margin and gap_len >= min_gap:
                gap_center = (gap_start + x) // 2
                dist = abs(gap_center - center)
                if dist < best_dist:
                    best_dist = dist
                    best_gap = (gap_start, x)
    return best_gap


def _find_inner_h_gap(region, min_gap=5, margin_ratio=0.10):
    """在区域内找水平方向的内部空白缝隙，返回最靠近中心的(gap_start, gap_end)或None。
    优先选最靠近中心的缝隙，而非最长的缝隙。"""
    import numpy as np
    rh, rw = region.shape
    row_proj = np.sum(region, axis=1) / 255
    row_is_blank = row_proj < rw * 0.008
    margin = int(rh * margin_ratio)
    center = rh // 2
    best_gap = None
    best_dist = float('inf')
    in_gap = False
    gap_start = 0
    for y in range(rh):
        if row_is_blank[y] and not in_gap:
            in_gap = True
            gap_start = y
        elif not row_is_blank[y] and in_gap:
            in_gap = False
            gap_len = y - gap_start
            if gap_start >= margin and y <= rh - margin and gap_len >= min_gap:
                gap_center = (gap_start + y) // 2
                dist = abs(gap_center - center)
                if dist < best_dist:
                    best_dist = dist
                    best_gap = (gap_start, y)
    return best_gap


def _secondary_split_wide(images, boxes, binary, page_w, page_h, aggressive=True, tall_h_split=False):
    """宽发票二次分割：支持水平+垂直分割。
    对整页未分割的情况，先水平分割（上下），再迭代垂直分割（左右）。
    输入(图片列表, 坐标列表)，返回(图片列表, 坐标列表)。"""
    try:
        import numpy as np
        if len(boxes) < 1:
            return (images, boxes)
        rows = {}
        for i, (x1, y1, x2, y2) in enumerate(boxes):
            row_key = round(y1 / page_h, 1)
            if row_key not in rows:
                rows[row_key] = []
            rows[row_key].append(i)
        new_images = []
        new_boxes = []
        for i, (x1, y1, x2, y2) in enumerate(boxes):
            cur_w = x2 - x1
            cur_h = y2 - y1
            cur_area = cur_w * cur_h
            row_key = round(y1 / page_h, 1)
            row_indices = rows.get(row_key, [i])
            row_widths = [boxes[j][2] - boxes[j][0] for j in row_indices]
            min_w = min(row_widths) if row_widths else cur_w

            # 判断是否需要分割
            cond_v1 = len(row_indices) >= 2 and cur_w > min_w * 1.35 and cur_w > page_w * 0.25
            cond_full = len(boxes) == 1 and cur_area > page_w * page_h * 0.4
            all_widths = [b[2] - b[0] for b in boxes]
            global_median = sorted(all_widths)[len(all_widths) // 2] if len(all_widths) > 1 else cur_w
            cond_v2 = len(row_indices) == 1 and len(boxes) > 1 and cur_w > global_median * 1.8 and cur_w > page_w * 0.4
            # 宽扁框强制垂直分割（宽>页宽55% 且 高<页高60%），排除横向小发票（高<页高20%）
            cond_v3 = cur_w > page_w * 0.55 and cur_h < page_h * 0.60 and cur_w > page_w * 0.25 and cur_h > page_h * 0.35
            need_v_split = cond_v1 or cond_full or cond_v2 or cond_v3
            # 水平分割：1)整页单框 2)tall_h_split模式下竖长框（高>宽*1.3且同行<=2列）
            cond_tall_h = tall_h_split and cur_h > cur_w * 1.15 and len(row_indices) <= 2
            need_h_split = aggressive and cur_h > page_h * 0.30 and (cond_full or len(boxes) == 1 or cond_tall_h)

            region = binary[y1:y2, x1:x2]
            rh, rw = region.shape

            # 优先水平分割（整页多行情景），需上下两半都有足够内容
            if need_h_split:
                import numpy as np
                # 尝试多个分割位置：缝隙位置 -> 中间 -> 45%处（非激进模式仅用缝隙）
                h_candidates = []
                h_gap = _find_inner_h_gap(region, min_gap=10, margin_ratio=0.20)
                if h_gap:
                    h_candidates.append(y1 + (h_gap[0] + h_gap[1]) // 2)
                if aggressive:
                    h_candidates.append(y1 + int(cur_h * 0.50))
                    h_candidates.append(y1 + int(cur_h * 0.45))
                h_split_done = False
                for split_y in h_candidates:
                    top_h = split_y - y1
                    bot_h = y2 - split_y
                    if top_h < page_h * 0.12 or bot_h < page_h * 0.12:
                        continue
                    top_region = region[0:top_h, :]
                    bot_region = region[top_h:cur_h, :]
                    top_density = np.sum(top_region) / (top_h * rw * 255) if top_h > 0 else 0
                    bot_density = np.sum(bot_region) / (bot_h * rw * 255) if bot_h > 0 else 0
                    if top_density > 0.015 and bot_density > 0.015:
                        top_img = images[i].crop((0, 0, images[i].size[0], top_h))
                        bot_img = images[i].crop((0, top_h, images[i].size[0], images[i].size[1]))
                        new_images.append(top_img)
                        new_boxes.append((x1, y1, x2, split_y))
                        new_images.append(bot_img)
                        new_boxes.append((x1, split_y, x2, y2))
                        h_split_done = True
                        break
                if h_split_done:
                    continue

            # 垂直分割
            if need_v_split:
                v_gap = _find_inner_v_gap(region, min_gap=5, margin_ratio=0.10)
                split_x = None
                if v_gap:
                    gap_x = x1 + (v_gap[0] + v_gap[1]) // 2
                    if (gap_x - x1) > page_w * 0.10 and (x2 - gap_x) > page_w * 0.10:
                        split_x = gap_x
                if split_x is None and aggressive:
                    # 缝隙分割失败，回退到中间分割（仅激进模式）
                    split_x = (x1 + x2) // 2
                left_w = split_x - x1
                right_w = x2 - split_x
                if left_w > page_w * 0.10 and right_w > page_w * 0.10:
                    left_img = images[i].crop((0, 0, left_w, images[i].size[1]))
                    right_img = images[i].crop((left_w, 0, images[i].size[0], images[i].size[1]))
                    new_images.append(left_img)
                    new_boxes.append((x1, y1, split_x, y2))
                    new_images.append(right_img)
                    new_boxes.append((split_x, y1, x2, y2))
                    continue

            new_images.append(images[i])
            new_boxes.append((x1, y1, x2, y2))
        return (new_images, new_boxes)
    except Exception as e:
        _log_ocr_error(f"宽发票二次分割失败: {e}")
        return (images, boxes)


def _merge_narrow_invoices(images, boxes, page_w, page_h):
    """后处理：合并同行内明显偏窄的发票（通常是被折痕/分割线误分割的单张发票）。
    如果某行有发票宽度 < 同行中位数宽度*0.7，且相邻两张都偏窄，合并它们。"""
    try:
        import numpy as np
        if len(boxes) < 2:
            return (images, boxes)
        # 按行分组
        rows = {}
        for i, (x1, y1, x2, y2) in enumerate(boxes):
            row_key = round(y1 / page_h, 1)
            if row_key not in rows:
                rows[row_key] = []
            rows[row_key].append(i)
        # 找出需要合并的相邻窄发票
        merge_pairs = []
        for row_key, indices in rows.items():
            if len(indices) < 2:
                continue
            widths = [boxes[j][2] - boxes[j][0] for j in indices]
            # 用行内最大宽度作为参考（窄发票会拉低中位数，导致检测不到）
            max_w = max(widths)
            # 找出偏窄的发票索引（<最大宽度*0.7，更宽松）
            narrow = [indices[k] for k in range(len(indices)) if widths[k] < max_w * 0.70]
            # 合并相邻的窄发票
            for k in range(len(narrow) - 1):
                idx1, idx2 = narrow[k], narrow[k + 1]
                # 检查是否相邻（x坐标连续）
                if abs(boxes[idx1][2] - boxes[idx2][0]) < page_w * 0.02:
                    merge_pairs.append((idx1, idx2))
        if not merge_pairs:
            return (images, boxes)
        # 执行合并
        merged_indices = set()
        new_images = []
        new_boxes = []
        for i in range(len(boxes)):
            if i in merged_indices:
                continue
            # 检查是否是合并对的第一个
            pair = next((p for p in merge_pairs if p[0] == i), None)
            if pair:
                idx1, idx2 = pair
                x1 = min(boxes[idx1][0], boxes[idx2][0])
                y1 = min(boxes[idx1][1], boxes[idx2][1])
                x2 = max(boxes[idx1][2], boxes[idx2][2])
                y2 = max(boxes[idx1][3], boxes[idx2][3])
                # 合并图片
                merged_w = x2 - x1
                merged_h = y2 - y1
                from PIL import Image as PILImage
                merged_img = PILImage.new('RGB', (merged_w, merged_h), (255, 255, 255))
                # 粘贴左半
                left_x = boxes[idx1][0] - x1
                merged_img.paste(images[idx1], (left_x, boxes[idx1][1] - y1))
                # 粘贴右半
                right_x = boxes[idx2][0] - x1
                merged_img.paste(images[idx2], (right_x, boxes[idx2][1] - y1))
                new_images.append(merged_img)
                new_boxes.append((x1, y1, x2, y2))
                merged_indices.add(idx2)
            else:
                new_images.append(images[i])
                new_boxes.append(boxes[i])
        return (new_images, new_boxes)
    except Exception as e:
        _log_ocr_error(f"窄发票合并失败: {e}")
        return (images, boxes)


def annotate_invoice_boxes(pil_img):
    """在原图上用红框标注分割出的每张发票，返回标注后的PIL Image。"""
    try:
        import cv2
        import numpy as np
        result, boxes = split_invoice_image(pil_img, return_boxes=True)
        if not boxes or len(boxes) <= 1:
            return pil_img
        cv_img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        for i, (x1, y1, x2, y2) in enumerate(boxes):
            cv2.rectangle(cv_img, (x1, y1), (x2, y2), (0, 0, 255), 4)
            # 标注序号
            cv2.putText(cv_img, str(i + 1), (x1 + 8, y1 + 35),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.2, (0, 0, 255), 3)
        return Image.fromarray(cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB))
    except Exception as e:
        _log_ocr_error(f"标注发票框失败: {e}")
        return pil_img


def perspective_crop(orig_img, points):
    """将四边形4点透视校正为矩形并裁切，返回PIL Image。失败返回None。"""
    try:
        import cv2
        import numpy as np
        if len(points) != 4 or not all(len(p) == 2 for p in points):
            return None
        cv_img = cv2.cvtColor(np.array(orig_img), cv2.COLOR_RGB2BGR)
        src_pts = np.float32(points)
        w1 = np.linalg.norm(src_pts[1] - src_pts[0])
        w2 = np.linalg.norm(src_pts[2] - src_pts[3])
        h1 = np.linalg.norm(src_pts[3] - src_pts[0])
        h2 = np.linalg.norm(src_pts[2] - src_pts[1])
        out_w = max(50, int(max(w1, w2)))
        out_h = max(30, int(max(h1, h2)))
        dst_pts = np.float32([[0, 0], [out_w - 1, 0], [out_w - 1, out_h - 1], [0, out_h - 1]])
        M = cv2.getPerspectiveTransform(src_pts, dst_pts)
        warped = cv2.warpPerspective(cv_img, M, (out_w, out_h))
        return Image.fromarray(cv2.cvtColor(warped, cv2.COLOR_BGR2RGB))
    except Exception as e:
        _log_ocr_error(f"透视裁切失败: {e}")
        return None


def _ocr_image_return_boxes(pil_img):
    """对PIL图片做OCR，返回 (texts, line_boxes)。
    line_boxes[i] 是 texts[i] 的4点坐标（相对该图左上角）。失败返回 ([], [])。"""
    engine = get_ocr_engine()
    if engine is None:
        return [], []
    try:
        import tempfile
        import numpy as np
        tmp_dir = tempfile.gettempdir()
        tmp_path = os.path.join(tmp_dir, f"ocrb_{os.getpid()}_{int(time.time()*1000)}.jpg")
        pil_img.convert('RGB').save(tmp_path, 'JPEG', quality=95)
        result, elapse = engine(tmp_path)
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        if result:
            texts = [str(line[1]).strip() for line in result]
            boxes = [line[0] for line in result]
            return texts, boxes
        return [], []
    except Exception as e:
        _log_ocr_error(f"OCR返回框失败: {e}")
        return [], []


def _match_text_box(texts, line_boxes, needle):
    """在OCR文本行中查找包含字段文本的行，返回其4点坐标；找不到返回None"""
    if not needle or not texts:
        return None
    needle_s = str(needle).strip()
    if not needle_s:
        return None
    for i, t in enumerate(texts):
        ts = str(t).strip()
        if ts == needle_s or needle_s in ts or ts in needle_s:
            if i < len(line_boxes):
                return line_boxes[i]
    return None


def _find_amount_box(texts, line_boxes, amount):
    """在OCR结果中定位价税合计金额位置，返回4点框；找不到返回None。
    优先数字规范化精确匹配，再找'小写/价税合计'行，最后子串回退（容忍OCR把0识别成O等误差）。"""
    if not amount or not texts:
        return None
    target = str(amount).replace(',', '')
    # 1. 数字规范化精确匹配（容忍 ￥¥、空格、逗号）
    for _i, _t in enumerate(texts):
        if re.sub(r'[^0-9.]', '', str(_t)) == target:
            if _i < len(line_boxes):
                return line_boxes[_i]
    # 2. 优先"小写"/"价税合计"行内含金额数字（容忍OCR把0识别成O等误差）
    for _i, _t in enumerate(texts):
        _s = str(_t)
        if ('小写' in _s or '价税合计' in _s) and re.search(r'\d', _s):
            if _i < len(line_boxes):
                return line_boxes[_i]
    # 3. 子串匹配回退
    return _match_text_box(texts, line_boxes, str(amount))


def auto_field_boxes(orig_img, boxes):
    """对每张发票（轴对齐矩形boxes）自动OCR识别发票号/日期/金额，生成绿框（原图坐标4点）。
    返回 list，长度=len(boxes)，每项 {invoice_number: 4点或None, date: ..., amount: ...}"""
    import re
    results = []
    for box in boxes:
        fb = {'invoice_number': None, 'date': None, 'amount': None}
        try:
            x1, y1, x2, y2 = [int(v) for v in box]
            if x2 - x1 < 20 or y2 - y1 < 20:
                results.append(fb)
                continue
            sub = orig_img.crop((x1, y1, x2, y2))
            texts, line_boxes = _ocr_image_return_boxes(sub)
            if not texts:
                results.append(fb)
                continue
            # 提取字段文本
            inv_no = extract_invoice_number(texts)
            date_text = None
            amount_text = None
            today = datetime.now().date()
            full = '\n'.join(texts)
            # 日期：开票日期/日期标签优先，其次通用日期
            for pat in [r'开票日期[：:\s]*(\d{4})年(\d{1,2})月(\d{1,2})日',
                        r'开票日期[：:\s]*(\d{4})[-/](\d{1,2})[-/](\d{1,2})',
                        r'(\d{4})年(\d{1,2})月(\d{1,2})日',
                        r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})']:
                m = re.search(pat, full)
                if m:
                    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                    if 2000 <= y <= today.year and 1 <= mo <= 12 and 1 <= d <= 31:
                        date_text = f"{y:04d}-{mo:02d}-{d:02d}"
                        break
            # 金额：优先价税合计（小写）金额，如"（小写金额 62.70）""（小写）￥239.00"
            amount_text = None
            for pat in [r'小写[^0-9]{0,20}?([0-9][0-9,]*\.?\d{0,2})',
                        r'价税合计[^0-9¥￥]{0,30}?[¥￥]?\s*([0-9][0-9,]*\.?\d{0,2})',
                        r'[¥￥]\s*([0-9][0-9,]*\.?\d{0,2})',
                        r'([0-9][0-9,]*\.\d{2})\s*元',
                        r'([0-9][0-9,]*\.\d{2})']:
                m = re.search(pat, full)
                if m:
                    amount_text = m.group(1)
                    break
            # 位置匹配并转原图坐标（日期需按OCR原文格式多变形匹配）
            for key, val in (('invoice_number', inv_no), ('date', date_text), ('amount', amount_text)):
                if not val:
                    continue
                needles = [str(val)]
                if key == 'date' and '-' in str(val):
                    y, mo, d = str(val).split('-')
                    needles += [f"{y}年{int(mo)}月{int(d)}日", f"{y}年{mo}月{d}日",
                                f"{y}/{mo}/{d}", f"{y}-{mo}-{d}"]
                lb = None
                for nd in needles:
                    lb = _match_text_box(texts, line_boxes, nd)
                    if lb:
                        break
                if lb and len(lb) == 4:
                    pts = [(x1 + int(p[0]), y1 + int(p[1])) for p in lb]
                    # 转为4点顺序：左上、右上、右下、左下
                    fb[key] = pts
        except Exception as e:
            _log_ocr_error(f"自动字段识别失败: {e}")
        results.append(fb)
    return results


# OCR引擎单例（延迟加载，避免启动慢）
_ocr_engine = None

def get_ocr_engine():
    """获取OCR引擎单例，首次调用时加载模型"""
    global _ocr_engine
    if _ocr_engine is None:
        try:
            from rapidocr_onnxruntime import RapidOCR
            import rapidocr_onnxruntime.rapid_ocr_api as rapid_api
            import rapidocr_onnxruntime.utils as rapid_utils

            # 打包环境下手动修正 root_dir，指向 _MEIPASS/rapidocr_onnxruntime
            if getattr(sys, 'frozen', False):
                meipass = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
                pkg_dir = os.path.join(meipass, 'rapidocr_onnxruntime')
                if os.path.isdir(pkg_dir):
                    from pathlib import Path
                    rapid_api.root_dir = Path(pkg_dir)
                    rapid_utils.root_dir = Path(pkg_dir)
                    _log_ocr_error(f"OCR root_dir patched to: {pkg_dir}")
                    config_path = os.path.join(pkg_dir, 'config.yaml')
                    models_path = os.path.join(pkg_dir, 'models')
                    _log_ocr_error(f"config.yaml exists: {os.path.exists(config_path)}")
                    _log_ocr_error(f"models dir exists: {os.path.isdir(models_path)}")
                    if os.path.isdir(models_path):
                        _log_ocr_error(f"models: {os.listdir(models_path)}")

            _ocr_engine = RapidOCR()
            _log_ocr_error("OCR引擎加载成功")
        except Exception as e:
            print(f"OCR引擎加载失败: {e}")
            _log_ocr_error(f"OCR引擎加载失败: {e}")
            import traceback
            _log_ocr_error(traceback.format_exc())
            _ocr_engine = False  # 标记为加载失败，避免重复尝试
    return _ocr_engine if _ocr_engine is not False else None


def extract_invoice_number(text):
    """从OCR识别的文本中提取发票号码"""
    import re
    if not text:
        return None
    # 合并所有行（用空格连接，保留分行信息）
    if isinstance(text, list):
        lines = [str(t).strip() for t in text if t]
        full_text = ' '.join(lines)
    else:
        lines = [str(text).strip()]
        full_text = str(text)

    # 1. 优先匹配"发票号码"后面的数字（可能同行或下一行）
    for i, line in enumerate(lines):
        line_clean = line.replace(' ', '').replace('　', '')
        m = re.search(r'发票号码[:：]?(\d{8,20})', line_clean)
        if m:
            return m.group(1)
        if line_clean.endswith('发票号码') or line_clean.endswith('发票号码:') or line_clean.endswith('发票号码：'):
            if i + 1 < len(lines):
                next_clean = lines[i+1].replace(' ', '').replace('　', '')
                m = re.search(r'(\d{8,20})', next_clean)
                if m:
                    return m.group(1)

    # 1b. 出租车/客运发票：匹配"号码:"后面的数字（区别于"代码:"）
    for i, line in enumerate(lines):
        line_clean = line.replace(' ', '').replace('　', '')
        # "号码:00357432" 或 "号码：00357432"
        m = re.search(r'号码[:：](\d{6,20})', line_clean)
        if m:
            # 排除"发票号码"已匹配过的情况，以及"电话号码"等
            if '发票号码' not in line_clean and '电话' not in line_clean and '监督' not in line_clean:
                return m.group(1)
        # "号码"单独一行，下一行是数字
        if line_clean == '号码' or line_clean == '号码:':
            if i + 1 < len(lines):
                next_clean = lines[i+1].replace(' ', '').replace('　', '')
                m = re.search(r'^(\d{6,20})', next_clean)
                if m:
                    return m.group(1)

    # 2. 在全文中匹配"发票号码"关键词
    full_clean = full_text.replace(' ', '').replace('　', '')
    m = re.search(r'发票号码[:：]?(\d{8,20})', full_clean)
    if m:
        return m.group(1)

    # 3. 匹配"No."后面的数字
    m = re.search(r'No[.．:：\s]*(\d{8,20})', full_text, re.IGNORECASE)
    if m:
        return m.group(1)

    # 3b. 旧式磁介质火车票：车票号为"字母+6-8位数字"（如A053990），红色印在右上角
    is_magnetic_ticket = any(kw in full_text for kw in [
        '仅供报销使用', '报销凭证', '遗失不补', '退票改签', '检票口',
        '须交回车站', '开车', '二等座', '一等座', '软卧', '硬卧', '硬座', '无座'
    ])
    if is_magnetic_ticket:
        # 匹配独立的"1位大写字母+6-8位数字"（如A053990）
        letter_nums = re.findall(r'(?<![A-Za-z0-9])([A-Z]\d{6,8})(?![A-Za-z0-9])', full_text)
        if letter_nums:
            # 优先取较短的（车票号通常7位：1字母+6数字）
            letter_nums.sort(key=len)
            return letter_nums[0]
        # 从左侧竖排长编号中提取（如34084300370308A053990）
        m = re.search(r'(\d{10,20})([A-Z]\d{6,8})', full_text)
        if m:
            return m.group(2)

    # 3c. 全文容错匹配"发票号码"（竖排/旋转发票可能中间有其他字符）
    full_clean = full_text.replace(' ', '').replace('　', '').replace('\n', '')
    m = re.search(r'发票号码[^0-9]{0,5}(\d{8,20})', full_clean)
    if m:
        return m.group(1)

    # 4. 找所有8-20位连续数字，排除日期、发票代码、金额、印刷编号范围等
    all_nums = re.findall(r'\b(\d{8,20})\b', full_text)
    candidates = []
    for num in all_nums:
        # 排除日期格式（20240101等8位日期）
        if len(num) == 8 and num.startswith(('19', '20')):
            continue
        # 排除发票印刷编号范围（如 00300001-00600000 中的起始/结束编号）
        # 检查该数字前后是否有连字符和另一个数字
        idx = full_text.find(num)
        if idx >= 0:
            before = full_text[max(0, idx-3):idx]
            after = full_text[idx+len(num):idx+len(num)+3]
            if re.search(r'\d[-－~～]\d*$', before) or re.search(r'^[-－~～]\d', after):
                continue
            # 排除"印XXXX份"附近的编号（印刷信息）
            context = full_text[max(0, idx-30):idx+len(num)+10]
            if '印' in context and '份' in context:
                continue
        candidates.append(num)

    if candidates:
        # 优先返回8位的（最常见的发票号码长度）
        for num in candidates:
            if len(num) == 8:
                return num
        # 其次返回20位的（全电发票）
        for num in candidates:
            if len(num) == 20:
                return num
        # 最后返回第一个
        return candidates[0]

    return None


def ocr_invoice_number(image_path):
    """识别发票图片/PDF中的发票号码，返回(发票号, 全部文本)"""
    engine = get_ocr_engine()
    if engine is None:
        _log_ocr_error("OCR引擎加载失败")
        return None, None
    try:
        # PDF先转图片（提高DPI改善文字识别准确率）
        if image_path.lower().endswith('.pdf'):
            img = pdf_to_image(image_path, dpi=300)
            if img is None:
                _log_ocr_error(f"PDF转图片失败: {image_path}")
                return None, None
        else:
            try:
                img = Image.open(image_path)
                # 处理RGBA/P模式
                if img.mode != 'RGB':
                    img = img.convert('RGB')
            except Exception as e:
                _log_ocr_error(f"打开图片失败: {e}")
                return None, None

        # 压缩大图片（最大边不超过3000px，兼顾识别准确率和速度）
        max_side = 3000
        w, h = img.size
        if max(w, h) > max_side:
            ratio = max_side / max(w, h)
            img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

        # 保存到临时英文路径（避免中文路径导致OCR失败）
        import tempfile
        tmp_dir = tempfile.gettempdir()
        tmp_path = os.path.join(tmp_dir, f"ocr_{os.getpid()}_{int(time.time()*1000)}.jpg")
        img.save(tmp_path, 'JPEG', quality=95)
        path_to_ocr = tmp_path

        # 执行OCR
        result, elapse = engine(path_to_ocr)

        # 释放图片内存
        try:
            img.close()
        except Exception:
            pass
        img = None

        # 清理临时文件
        try:
            os.unlink(path_to_ocr)
        except Exception:
            pass

        if result:
            texts = [line[1] for line in result]
            _log_ocr_error(f"OCR成功，识别到{len(texts)}行文本: {texts[:5]}")
            inv_no = extract_invoice_number(texts)
            return inv_no, texts
        else:
            _log_ocr_error(f"OCR未识别到任何文本，耗时: {elapse}")
            return None, None
    except Exception as e:
        _log_ocr_error(f"OCR异常: {e}")
        import traceback
        _log_ocr_error(traceback.format_exc())
        # 释放图片内存
        try:
            if 'img' in locals():
                img.close()
        except Exception:
            pass
        return None, None


def _log_ocr_error(msg):
    """记录OCR错误到日志文件"""
    try:
        log_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(log_dir, 'ocr_log.txt')
        with open(log_path, 'a', encoding='utf-8') as f:
            from datetime import datetime
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


def _get_min_hire_year():
    try:
        conn = get_db()
        row = conn.execute("SELECT MIN(hire_date) as min_date FROM employees WHERE hire_date IS NOT NULL AND hire_date != ''").fetchone()
        conn.close()
        if row and row['min_date']:
            try:
                return int(str(row['min_date'])[:4])
            except (ValueError, IndexError):
                pass
    except Exception:
        pass
    return 2015

def get_monthly_salary(employee_id, year, month):
    """根据入职、离职、调薪记录计算指定年月的应发薪资"""
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id=?", (employee_id,)).fetchone()
    if not emp:
        conn.close()
        return 0

    # 检查入职时间
    if emp['hire_date']:
        try:
            hire_parts = emp['hire_date'].split('-')
            hire_y = int(hire_parts[0])
            hire_m = int(hire_parts[1])
            if (year, month) < (hire_y, hire_m):
                conn.close()
                return 0
        except Exception:
            pass

    # 检查离职时间
    if emp['resign_date']:
        try:
            resign_parts = emp['resign_date'].split('-')
            resign_y = int(resign_parts[0])
            resign_m = int(resign_parts[1])
            if (year, month) > (resign_y, resign_m):
                conn.close()
                return 0
        except Exception:
            pass

    # 基础薪资
    salary = float(emp['monthly_salary'])

    # 查找在指定年月之前或当月生效的最新调薪记录
    adj = conn.execute("""SELECT new_salary FROM salary_adjustments
                          WHERE employee_id=? AND (effective_year < ? OR (effective_year=? AND effective_month<=?))
                          ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1""",
                       (employee_id, year, year, month)).fetchone()
    if adj:
        salary = float(adj['new_salary'])

    conn.close()
    return salary

def get_monthly_insurance(employee_id, year, month):
    # 根据入职、离职、五险一金调整记录计算指定年月的五险一金金额
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id=?", (employee_id,)).fetchone()
    if not emp:
        conn.close()
        return 0
    # 检查入职时间
    if emp['hire_date']:
        try:
            hire_parts = emp['hire_date'].split('-')
            hire_y = int(hire_parts[0])
            hire_m = int(hire_parts[1])
            if (year, month) < (hire_y, hire_m):
                conn.close()
                return 0
        except Exception:
            pass
    # 检查离职时间
    if emp['resign_date']:
        try:
            resign_parts = emp['resign_date'].split('-')
            resign_y = int(resign_parts[0])
            resign_m = int(resign_parts[1])
            if (year, month) > (resign_y, resign_m):
                conn.close()
                return 0
        except Exception:
            pass
    # 基础金额
    amount = float(emp['insurance_amount'] or 0)
    # 查找在指定年月之前或当月生效的最新调整记录
    adj = conn.execute("SELECT new_amount FROM insurance_adjustments WHERE employee_id=? AND (effective_year < ? OR (effective_year=? AND effective_month<=?)) ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1",
                       (employee_id, year, year, month)).fetchone()
    if adj:
        amount = float(adj['new_amount'])
    conn.close()
    return amount

INSURANCE_ITEMS = [
    ('pension', '养老保险'),
    ('medical', '医疗保险'),
    ('unemployment', '失业保险'),
    ('injury', '工伤保险'),
    ('maternity', '生育保险'),
    ('housing', '住房公积金'),
]

INSURANCE_ITEMS_DICT = dict(INSURANCE_ITEMS)


def get_monthly_insurance_detail(employee_id, year, month):
    # 返回指定年月的五险一金细项 dict: {item: {company, personal, total}}
    conn = get_db()
    emp = conn.execute("SELECT * FROM employees WHERE id=?", (employee_id,)).fetchone()
    if not emp:
        conn.close()
        return {}
    # 检查入职离职
    if emp['hire_date']:
        try:
            hp = emp['hire_date'].split('-')
            if (year, month) < (int(hp[0]), int(hp[1])):
                conn.close()
                return {}
        except Exception:
            pass
    if emp['resign_date']:
        try:
            rp = emp['resign_date'].split('-')
            if (year, month) > (int(rp[0]), int(rp[1])):
                conn.close()
                return {}
        except Exception:
            pass
    # 基础值从employees取
    result = {}
    for key, name in INSURANCE_ITEMS:
        result[key] = {
            'name': name,
            'company': float(emp[f'{key}_company'] or 0),
            'personal': float(emp[f'{key}_personal'] or 0),
        }
    # 查找最新调整记录
    adj = conn.execute("SELECT * FROM insurance_detail_adjustments WHERE employee_id=? AND (effective_year < ? OR (effective_year=? AND effective_month<=?)) ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1",
                       (employee_id, year, year, month)).fetchone()
    if adj:
        for key, name in INSURANCE_ITEMS:
            result[key]['company'] = float(adj[f'{key}_company'] or 0)
            result[key]['personal'] = float(adj[f'{key}_personal'] or 0)
    conn.close()
    # 计算total
    for key in result:
        result[key]['total'] = result[key]['company'] + result[key]['personal']
    return result


def get_monthly_insurance_total(employee_id, year, month):
    # 返回该月五险一金总额（公司+个人）
    detail = get_monthly_insurance_detail(employee_id, year, month)
    if not detail:
        return 0
    return sum(d['total'] for d in detail.values())

def fmt_money(v):
    try:
        return f"¥{float(v):,.2f}"
    except Exception:
        return "¥0.00"


# Excel 样式
HEADER_FONT = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
HEADER_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
TITLE_FONT = Font(name='微软雅黑', bold=True, size=14)
TOTAL_FONT = Font(name='微软雅黑', bold=True, size=11)
TOTAL_FILL = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def _auto_width(ws, ncols, min_w=10, max_w=40):
    for c in range(1, ncols + 1):
        max_len = min_w
        for row in ws.iter_rows(min_col=c, max_col=c, values_only=True):
            v = row[0]
            if v is not None:
                length = sum(2 if ord(ch) > 127 else 1 for ch in str(v))
                max_len = max(max_len, length)
        ws.column_dimensions[get_column_letter(c)].width = min(max_len + 2, max_w)


def is_dark_theme():
    """判断当前是否为深色模式（独立函数，供控件使用）"""
    theme_setting = get_setting('app_theme', 'system')
    if theme_setting == 'dark':
        return True
    if theme_setting == 'light':
        return False
    # 跟随系统
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
        winreg.CloseKey(key)
        return value == 0
    except Exception:
        return False


class FlatCheckbutton(tk.Frame):
    """自定义复选框：深色模式下✓为白色，背景黑色"""
    def __init__(self, master=None, text='', variable=None, command=None, **kwargs):
        super().__init__(master)
        self.variable = variable
        self.command = command
        self._dark = is_dark_theme()

        if self._dark:
            self._bg = '#1C1C1E'
            self._btn_bg = '#000000'
            self._btn_fg = '#FFFFFF'
            self._border = '#FFFFFF'
            self._text_fg = '#F5F5F7'
            self._hover_bg = '#2C2C2E'
        else:
            self._bg = '#F0F1F2'
            self._btn_bg = '#FFFFFF'
            self._btn_fg = '#000000'
            self._border = '#D1D1D6'
            self._text_fg = '#1D1D1F'
            self._hover_bg = '#DCDDE0'

        self.configure(bg=self._bg)

        # 用Frame做边框容器，确保边框色可控
        self.border_frame = tk.Frame(self, bg=self._border, bd=0)
        self.border_frame.pack(side='left')
        self.btn = tk.Button(self.border_frame, text='', width=2, relief='flat', bd=0,
                            bg=self._btn_bg, fg=self._btn_fg,
                            activebackground=self._hover_bg, activeforeground=self._btn_fg,
                            highlightthickness=0, command=self._toggle,
                            font=('Segoe UI', 9, 'bold'), cursor='hand2',
                            padx=2, pady=0)
        self.btn.pack(padx=1, pady=1)

        self.label = tk.Label(self, text=text, bg=self._bg, fg=self._text_fg,
                             font=('Segoe UI', 9), cursor='hand2')
        self.label.pack(side='left', padx=(4, 0))
        self.label.bind('<Button-1>', lambda e: self._toggle())
        self.btn.bind('<Enter>', lambda e: self.btn.config(bg=self._hover_bg))
        self.btn.bind('<Leave>', lambda e: self.btn.config(bg=self._btn_bg))

        if self.variable:
            self.variable.trace_add('write', self._update)
        self._update()

    def _toggle(self):
        if self.variable:
            self.variable.set(not self.variable.get())
        if self.command:
            self.command()

    def _update(self, *args):
        if self.variable and self.variable.get():
            self.btn.config(text='✓')
        else:
            self.btn.config(text='')


# ============================================================
# 主应用
# ============================================================
class FinanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # 立即隐藏窗口，避免初始化期间显示白色
        self.withdraw()
        # 设置背景色，避免显示白色
        try:
            if self.get_current_theme() == 'dark':
                self.configure(bg='#000000')
            else:
                self.configure(bg='#E6E7E8')
        except Exception:
            self.configure(bg='#E6E7E8')
        # 关键：将默认root设为自己，确保所有StringVar绑定到app而非splash
        tk._default_root = self
        self.title("财务管理系统")
        self.geometry("1200x750")
        self.minsize(900, 600)

        # 设置窗口图标（使用exe内嵌图标，最可靠）
        try:
            import sys as _sys
            import os as _os
            if getattr(_sys, 'frozen', False):
                _icon_path = _sys.executable
            else:
                _icon_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), 'icon.ico')
            if _os.path.exists(_icon_path):
                self.iconbitmap(_icon_path)
        except Exception as _e:
            pass

        # 样式
        self.setup_styles()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=12, pady=(12, 6))

        self.emp_tab = EmployeeTab(self.notebook, self)
        self.inv_tab = InvoiceTab(self.notebook, self)
        self.batch_tab = BatchInvoiceTab(self.notebook, self)
        self.sal_tab = SalaryTab(self.notebook, self)
        self.stat_tab = StatsTab(self.notebook, self)
        self.bank_tab = BankTab(self.notebook, self)
        self.backup_tab = BackupTab(self.notebook, self)
        self.settings_tab = SettingsTab(self.notebook, self)

        self.notebook.add(self.emp_tab, text='人员管理')
        self.notebook.add(self.inv_tab, text='发票报销')
        self.notebook.add(self.batch_tab, text='批量添加发票')
        self.notebook.add(self.sal_tab, text='工资结算')
        self.notebook.add(self.stat_tab, text='统计汇总')
        self.notebook.add(self.bank_tab, text='银行账户')
        self.notebook.add(self.backup_tab, text='数据备份')
        self.notebook.add(self.settings_tab, text='设置')

        # 状态栏 - 现代扁平
        self.status_var = tk.StringVar(value="就绪")
        if self.get_current_theme() == 'dark':
            _status_bg = '#1C1C1E'
            _status_fg = '#86868B'
            _status_border = '#2C2C2E'
        else:
            _status_bg = '#DCDDE0'
            _status_fg = '#6E6E73'
            _status_border = '#D1D1D6'
        status_frame = tk.Frame(self, bg=_status_border, height=30)
        status_frame.pack(fill='x', side='bottom')
        status_inner = tk.Frame(status_frame, bg=_status_bg, height=28)
        status_inner.pack(fill='both', expand=True, pady=(1, 0))
        self._status_frame = status_inner
        status = tk.Label(status_inner, textvariable=self.status_var, anchor='w',
                          bg=_status_bg, fg=_status_fg, font=('微软雅黑', 9), padx=14)
        status.pack(fill='both', expand=True)
        self._status_label = status

        # 修复深色模式下Combobox下拉列表颜色
        self.after(100, self.fix_combobox_popup)
        # 标签页切换日志
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        try:
            log_info('主界面初始化完成，所有标签页已加载')
        except Exception:
            pass

    def _on_tab_changed(self, event):
        try:
            tab = event.widget.tab('current')['text']
            log_op('切换标签页', tab)
        except Exception:
            pass

    def apply_theme(self):
        """实时切换主题，无需重启"""
        # 1. 重新配置ttk样式
        self.setup_styles()

        # 2. 更新主窗口背景
        is_dark = self.get_current_theme() == 'dark'
        if is_dark:
            _bg = '#000000'
            _status_bg = '#1C1C1E'
            _status_fg = '#86868B'
            _status_border = '#2C2C2E'
            _canvas_bg = '#000000'
        else:
            _bg = '#E6E7E8'
            _status_bg = '#DCDDE0'
            _status_fg = '#6E6E73'
            _status_border = '#D1D1D6'
            _canvas_bg = '#E6E7E8'

        self.configure(bg=_bg)

        # 3. 更新状态栏
        try:
            self._status_frame.master.configure(bg=_status_border)
            self._status_frame.configure(bg=_status_bg)
            self._status_label.configure(bg=_status_bg, fg=_status_fg)
        except Exception:
            pass

        # 4. 递归更新所有Canvas背景（ScrollableTab等）
        def _update_canvases(widget):
            try:
                if isinstance(widget, tk.Canvas):
                    widget.configure(bg=_canvas_bg)
            except Exception:
                pass
            try:
                for child in widget.winfo_children():
                    _update_canvases(child)
            except Exception:
                pass
        _update_canvases(self)

        # 5. 更新FlatCheckbutton实例（整个过程包裹在try/except中，任何错误不影响主题切换）
        try:
            def _update_flat_checks(widget):
                try:
                    if widget.__class__.__name__ == 'FlatCheckbutton' and hasattr(widget, 'border_frame') and hasattr(widget, 'btn'):
                        widget._dark = is_dark
                        if is_dark:
                            widget._bg = '#1C1C1E'
                            widget._btn_bg = '#000000'
                            widget._btn_fg = '#FFFFFF'
                            widget._border = '#FFFFFF'
                            widget._text_fg = '#F5F5F7'
                            widget._hover_bg = '#2C2C2E'
                        else:
                            widget._bg = '#F0F1F2'
                            widget._btn_bg = '#FFFFFF'
                            widget._btn_fg = '#000000'
                            widget._border = '#D1D1D6'
                            widget._text_fg = '#1D1D1F'
                            widget._hover_bg = '#DCDDE0'
                        widget.configure(bg=widget._bg)
                        widget.border_frame.configure(bg=widget._border)
                        widget.btn.configure(bg=widget._btn_bg, fg=widget._btn_fg,
                                             activebackground=widget._hover_bg, activeforeground=widget._btn_fg)
                        if hasattr(widget, 'label'):
                            widget.label.configure(bg=widget._bg, fg=widget._text_fg)
                except Exception:
                    pass
                try:
                    for child in widget.winfo_children():
                        _update_flat_checks(child)
                except Exception:
                    pass
            _update_flat_checks(self)
        except Exception:
            pass

        # 6. 通用更新所有tk原生控件颜色（Listbox、Text、Label、Frame、Entry等）
        try:
            if is_dark:
                _widget_bg = '#1C1C1E'
                _widget_fg = '#F5F5F7'
                _select_bg = '#007AFF'
                _select_fg = '#FFFFFF'
                _insert_bg = '#FFFFFF'
                _dark_colors = {'#000000', '#1c1c1e', '#2c2c2e', '#0a0a0b', '#111827'}
            else:
                _widget_bg = '#FFFFFF'
                _widget_fg = '#1D1D1F'
                _select_bg = '#007AFF'
                _select_fg = '#FFFFFF'
                _insert_bg = '#000000'
                _dark_colors = {'#000000', '#1c1c1e', '#2c2c2e', '#0a0a0b', '#111827'}

            def _update_tk_widgets(widget):
                try:
                    wtype = widget.winfo_class()
                    if wtype == 'Listbox':
                        try:
                            widget.configure(bg=_widget_bg, fg=_widget_fg,
                                           selectbackground=_select_bg, selectforeground=_select_fg)
                        except Exception:
                            pass
                    elif wtype == 'Text':
                        try:
                            widget.configure(bg=_widget_bg, fg=_widget_fg,
                                           insertbackground=_insert_bg)
                        except Exception:
                            pass
                    elif wtype == 'Entry':
                        try:
                            widget.configure(bg=_widget_bg, fg=_widget_fg,
                                           insertbackground=_insert_bg)
                        except Exception:
                            pass
                    elif wtype == 'Label':
                        try:
                            cur_bg = widget.cget('bg').lower()
                            if cur_bg in _dark_colors:
                                widget.configure(bg=_widget_bg, fg=_widget_fg)
                        except Exception:
                            pass
                    elif wtype == 'Frame':
                        try:
                            cur_bg = widget.cget('bg').lower()
                            if cur_bg in _dark_colors:
                                widget.configure(bg=_widget_bg)
                        except Exception:
                            pass
                    elif wtype == 'Button':
                        try:
                            cur_bg = widget.cget('bg').lower()
                            if cur_bg in _dark_colors:
                                widget.configure(bg=_widget_bg, fg=_widget_fg,
                                               activebackground='#2C2C2E' if is_dark else '#DCDDE0')
                        except Exception:
                            pass
                    elif wtype == 'Canvas':
                        try:
                            widget.configure(bg='#000000' if is_dark else '#E6E7E8')
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    for child in widget.winfo_children():
                        _update_tk_widgets(child)
                except Exception:
                    pass
            _update_tk_widgets(self)
        except Exception:
            pass

        # 7. 重新修复Combobox下拉列表颜色
        self.after(100, self.fix_combobox_popup)

        self.set_status(f"主题已切换为{'深色' if is_dark else '浅色'}模式")

    def fix_combobox_popup(self):
        """遍历所有Combobox，绑定下拉列表弹出时的颜色修复"""
        if self.get_current_theme() != 'dark':
            return
        self._walk_fix_combo(self)

    def _walk_fix_combo(self, widget):
        """递归遍历子组件，找到Combobox并绑定postcommand"""
        try:
            if isinstance(widget, ttk.Combobox):
                widget.bind('<Button-1>', self._on_combo_pre_open, add='+')
                widget.bind('<KeyPress-Down>', self._on_combo_pre_open, add='+')
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._walk_fix_combo(child)
        except Exception:
            pass

    def _on_combo_pre_open(self, event):
        """Combobox即将弹出下拉列表时，延迟设置Listbox颜色"""
        combo = event.widget
        self.after(20, lambda c=combo: self._apply_combo_listbox_style(c))

    def _apply_combo_listbox_style(self, combo):
        """强制设置Combobox下拉列表Listbox的深色样式"""
        try:
            popup = combo.tk.call('ttk::combobox::PopdownWindow', combo)
            if popup:
                listbox = f"{popup}.f.l"
                combo.tk.call(listbox, 'configure', '-background', '#1C1C1E')
                combo.tk.call(listbox, 'configure', '-foreground', '#F5F5F7')
                combo.tk.call(listbox, 'configure', '-selectbackground', '#1C3A5C')
                combo.tk.call(listbox, 'configure', '-selectforeground', '#F5F5F7')
                combo.tk.call(listbox, 'configure', '-borderwidth', 0)
                frame = f"{popup}.f"
                combo.tk.call(frame, 'configure', '-background', '#2C2C2E')
        except Exception:
            pass

    def get_current_theme(self):
        """获取当前主题：light / dark"""
        theme_setting = get_setting('app_theme', 'dark')
        if theme_setting == 'light':
            return 'light'
        elif theme_setting == 'dark':
            return 'dark'
        else:
            # 跟随系统
            return self._detect_system_theme()

    def _detect_system_theme(self):
        """检测Windows系统主题（浅色/深色）"""
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
            value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
            winreg.CloseKey(key)
            return 'light' if value == 1 else 'dark'
        except Exception:
            return 'light'

    def setup_styles(self):
        """Fluent 2 风格样式配置（支持浅色/深色主题）"""
        theme = self.get_current_theme()
        is_dark = (theme == 'dark')

        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except Exception:
            pass

        # 颜色定义 - Apple风格
        if is_dark:
            C_PRIMARY = '#0A84FF'           # Apple深色模式蓝
            C_PRIMARY_HOVER = '#409CFF'     # hover提亮
            C_PRIMARY_PRESS = '#0066CC'     # 按下
            C_PRIMARY_LIGHT = '#1C3A5C'     # 选中行背景
            C_BG = '#000000'                # 纯黑背景
            C_CARD = '#1C1C1E'              # Apple深色分组背景
            C_TEXT = '#F5F5F7'              # Apple浅文字
            C_TEXT_SUB = '#86868B'          # Apple次文字
            C_BORDER = '#2C2C2E'            # 深色分隔线
            C_SUCCESS = '#30D158'           # Apple绿
            C_INPUT_BG = '#1C1C1E'          # 输入框背景
            C_HEADING_BG = '#2C2C2E'        # 表头背景
            C_ROW_ALT = '#1C1C1E'
            C_BTN_SEC_BG = '#2C2C2E'        # 次要按钮背景
            C_BTN_SEC_BORDER = '#3A3A3C'
            C_BTN_SEC_HOVER = '#3A3A3C'
            C_BTN_SEC_PRESS = '#1C1C1E'
        else:
            C_PRIMARY = '#007AFF'           # Apple系统蓝
            C_PRIMARY_HOVER = '#0051D5'     # hover加深
            C_PRIMARY_PRESS = '#003DA5'     # 按下
            C_PRIMARY_LIGHT = '#D6E4FF'     # 选中行浅蓝
            C_BG = '#E6E7E8'                # 柔和灰背景（不刺眼）
            C_CARD = '#F0F1F2'              # 浅灰卡片
            C_TEXT = '#1D1D1F'              # Apple黑
            C_TEXT_SUB = '#6E6E73'          # Apple灰
            C_BORDER = '#D1D1D6'            # 分隔线
            C_SUCCESS = '#34C759'           # Apple绿
            C_INPUT_BG = '#FFFFFF'          # 输入框背景保持白色
            C_HEADING_BG = '#DCDDE0'        # 表头背景
            C_ROW_ALT = '#ECEDEE'
            C_BTN_SEC_BG = '#DCDDE0'        # 次要按钮背景
            C_BTN_SEC_BORDER = '#D1D1D6'
            C_BTN_SEC_HOVER = '#D1D1D6'
            C_BTN_SEC_PRESS = '#C6C7C8'


        # 全局背景
        self.configure(bg=C_BG)

        # 字体 - Apple风格字重层级
        FONT_BASE = ('微软雅黑', 10)
        FONT_BOLD = ('微软雅黑', 10, 'bold')
        FONT_TITLE = ('微软雅黑', 13, 'bold')
        FONT_SECTION = ('微软雅黑', 11, 'bold')

        # TLabel
        style.configure('TLabel', font=FONT_BASE, foreground=C_TEXT, background=C_CARD)
        style.configure('Card.TLabel', font=FONT_BASE, foreground=C_TEXT, background=C_CARD)
        style.configure('Sub.TLabel', font=FONT_BASE, foreground=C_TEXT_SUB, background=C_CARD)
        style.configure('Title.TLabel', font=FONT_TITLE, foreground=C_TEXT, background=C_CARD)
        style.configure('Section.TLabel', font=FONT_SECTION, foreground=C_TEXT, background=C_CARD)
        style.configure('Success.TLabel', font=FONT_BOLD, foreground=C_SUCCESS, background=C_CARD)

        # TButton - Apple风格主按钮，饱满圆润
        style.configure('TButton', font=FONT_BOLD, foreground='white', background=C_PRIMARY,
                       borderwidth=0, focusthickness=0, padding=(24, 9), anchor='center', relief='flat')
        style.map('TButton',
                 background=[('active', C_PRIMARY_HOVER), ('pressed', C_PRIMARY_PRESS), ('disabled', C_PRIMARY_LIGHT)],
                 foreground=[('disabled', C_TEXT_SUB)],
                 relief=[('pressed', 'flat')])

        # 次要按钮 - Apple风格，浅灰背景无边框
        style.configure('Secondary.TButton', font=FONT_BASE, foreground=C_PRIMARY, background=C_BTN_SEC_BG,
                       borderwidth=0, focusthickness=0, padding=(20, 8), anchor='center', relief='flat')
        style.map('Secondary.TButton',
                 background=[('active', C_BTN_SEC_HOVER), ('pressed', C_BTN_SEC_PRESS)],
                 foreground=[('active', C_PRIMARY_HOVER)],
                 relief=[('pressed', 'flat')])

        # 强调按钮 - 主色填充，用于打印排版等关键操作
        style.configure('Accent.TButton', font=FONT_BOLD, foreground='white', background=C_PRIMARY,
                       borderwidth=0, focusthickness=0, padding=(24, 9), anchor='center', relief='flat')
        style.map('Accent.TButton',
                 background=[('active', C_PRIMARY_HOVER), ('pressed', C_PRIMARY_PRESS), ('disabled', C_PRIMARY_LIGHT)],
                 foreground=[('disabled', C_TEXT_SUB)],
                 relief=[('pressed', 'flat')])

        # TEntry - Apple极简输入框
        style.configure('TEntry', font=FONT_BASE, foreground=C_TEXT, fieldbackground=C_INPUT_BG,
                       bordercolor=C_BORDER, lightcolor=C_BORDER, darkcolor=C_BORDER, padding=8, relief='flat')
        style.map('TEntry', bordercolor=[('focus', C_PRIMARY)], lightcolor=[('focus', C_PRIMARY)])

        # TCombobox
        style.configure('TCombobox', font=FONT_BASE, foreground=C_TEXT, fieldbackground=C_INPUT_BG,
                       background=C_INPUT_BG, bordercolor=C_BORDER, arrowcolor=C_PRIMARY,
                       lightcolor=C_BORDER, darkcolor=C_BORDER, padding=7, relief='flat')
        style.map('TCombobox',
                 fieldbackground=[('readonly', C_INPUT_BG), ('focus', C_INPUT_BG), ('active', C_INPUT_BG)],
                 foreground=[('readonly', C_TEXT), ('focus', C_TEXT)],
                 background=[('readonly', C_INPUT_BG), ('active', C_INPUT_BG)],
                 bordercolor=[('focus', C_PRIMARY), ('active', C_PRIMARY)],
                 arrowcolor=[('active', C_PRIMARY_HOVER)])

        # TSpinbox
        style.configure('TSpinbox', font=FONT_BASE, foreground=C_TEXT, fieldbackground=C_INPUT_BG,
                       background=C_INPUT_BG, bordercolor=C_BORDER,
                       lightcolor=C_BORDER, darkcolor=C_BORDER, padding=7, relief='flat')
        style.map('TSpinbox',
                 fieldbackground=[('readonly', C_INPUT_BG), ('focus', C_INPUT_BG)],
                 foreground=[('readonly', C_TEXT)],
                 background=[('readonly', C_INPUT_BG)],
                 bordercolor=[('focus', C_PRIMARY)])

        # TCheckbutton
        style.configure('TCheckbutton', font=FONT_BASE, foreground=C_TEXT, background=C_CARD,
                       indicatorcolor=C_BTN_SEC_BORDER, indicatorbackground=C_INPUT_BG)
        style.map('TCheckbutton', background=[('active', C_CARD)],
                 indicatorcolor=[('selected', '#FFFFFF'), ('pressed', '#CCCCCC'),
                                ('active', C_PRIMARY_HOVER)],
                 indicatorbackground=[('selected', '#000000'), ('pressed', '#111111'),
                                     ('active', '#1A1A1A')],
                 foreground=[('active', C_TEXT), ('disabled', C_TEXT_SUB)])

        # TRadiobutton
        style.configure('TRadiobutton', font=FONT_BASE, foreground=C_TEXT, background=C_CARD)

        # TNotebook - Apple极简标签页
        style.configure('TNotebook', background=C_BG, borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure('TNotebook.Tab', font=FONT_BOLD, foreground=C_TEXT_SUB, background=C_BG,
                       padding=(28, 14), borderwidth=0)
        style.map('TNotebook.Tab',
                 background=[('selected', C_CARD), ('active', C_CARD)],
                 foreground=[('selected', C_PRIMARY), ('active', C_TEXT)])

        # TFrame
        style.configure('TFrame', background=C_CARD)
        style.configure('Card.TFrame', background=C_CARD)
        style.configure('BG.TFrame', background=C_BG)

        # TLabelframe - Apple风格，去掉重边框，用标题+间距
        style.configure('TLabelframe', background=C_CARD, bordercolor=C_BORDER,
                       borderwidth=1, relief='solid', padding=16)
        style.configure('TLabelframe.Label', font=FONT_SECTION, foreground=C_TEXT, background=C_CARD)

        # Treeview - Apple极简表格，无网格线
        style.configure('Treeview', font=FONT_BASE, foreground=C_TEXT, background=C_INPUT_BG,
                       fieldbackground=C_INPUT_BG, bordercolor=C_BORDER, rowheight=36, borderwidth=0)
        style.configure('Treeview.Heading', font=FONT_BOLD, foreground=C_TEXT_SUB, background=C_HEADING_BG,
                       bordercolor=C_BORDER, borderwidth=0, relief='flat', padding=(12, 12))
        style.map('Treeview', background=[('selected', C_PRIMARY_LIGHT)],
                 foreground=[('selected', C_TEXT)])
        style.map('Treeview.Heading', background=[('active', C_HEADING_BG)])

        # TScrollbar - Apple细滚动条
        style.configure('Vertical.TScrollbar', background=C_BORDER, troughcolor=C_CARD,
                       bordercolor=C_CARD, arrowcolor=C_TEXT_SUB, width=8, relief='flat')
        style.map('Vertical.TScrollbar', background=[('active', C_PRIMARY_LIGHT)])
        style.configure('Horizontal.TScrollbar', background=C_BORDER, troughcolor=C_CARD,
                       bordercolor=C_CARD, arrowcolor=C_TEXT_SUB, height=8, relief='flat')

        # TProgressbar
        style.configure('TProgressbar', troughcolor=C_BORDER, background=C_PRIMARY,
                       bordercolor=C_BORDER, thickness=4, relief='flat')

        # 设置选项菜单背景
        self.option_add('*TCombobox*Listbox.background', C_INPUT_BG)
        self.option_add('*TCombobox*Listbox.foreground', C_TEXT)
        self.option_add('*TCombobox*Listbox.selectBackground', C_PRIMARY_LIGHT)
        self.option_add('*TCombobox*Listbox.selectForeground', C_TEXT)
        self.option_add('*TCombobox*Listbox.font', FONT_BASE)

        # 设置所有tk组件默认背景
        self.option_add('*Frame.background', C_CARD)
        self.option_add('*LabelFrame.background', C_CARD)
        self.option_add('*Label.background', C_CARD)
        self.option_add('*Label.foreground', C_TEXT)
        self.option_add('*Button.background', C_PRIMARY)
        self.option_add('*Button.foreground', 'white')
        self.option_add('*Button.font', FONT_BASE)
        self.option_add('*Button.borderWidth', 0)
        self.option_add('*Button.padX', 16)
        self.option_add('*Button.padY', 6)
        self.option_add('*Entry.background', C_INPUT_BG)
        self.option_add('*Entry.foreground', C_TEXT)
        self.option_add('*Entry.font', FONT_BASE)
        self.option_add('*Text.background', C_INPUT_BG)
        self.option_add('*Text.foreground', C_TEXT)
        self.option_add('*Text.font', FONT_BASE)
        self.option_add('*Listbox.background', C_INPUT_BG)
        self.option_add('*Listbox.foreground', C_TEXT)
        self.option_add('*Listbox.font', FONT_BASE)
        self.option_add('*Checkbutton.background', C_CARD)
        self.option_add('*Checkbutton.foreground', C_TEXT)
        self.option_add('*Checkbutton.font', FONT_BASE)
        self.option_add('*Radiobutton.background', C_CARD)
        self.option_add('*Radiobutton.foreground', C_TEXT)
        self.option_add('*Radiobutton.font', FONT_BASE)
        self.option_add('*Scale.background', C_CARD)
        self.option_add('*Spinbox.background', C_INPUT_BG)
        self.option_add('*Spinbox.foreground', C_TEXT)
        self.option_add('*Spinbox.font', FONT_BASE)
        self.option_add('*Menu.background', C_CARD)
        self.option_add('*Menu.foreground', C_TEXT)
        self.option_add('*Menu.font', FONT_BASE)
        self.option_add('*Menubutton.background', C_CARD)
        self.option_add('*Menubutton.foreground', C_TEXT)
        self.option_add('*Message.background', C_CARD)
        self.option_add('*Message.foreground', C_TEXT)

        # 保存当前主题到实例变量
        self._current_theme = theme

    def refresh_all(self):
        self.emp_tab.refresh()
        self.inv_tab.refresh()
        self.sal_tab.refresh()
        self.stat_tab.refresh()
        self.bank_tab.refresh_accounts()
        if hasattr(self, 'backup_tab'):
            self.backup_tab.refresh_stats()

    def _refresh_other_tabs(self, current):
        """异步刷新除当前页外的其他标签页，避免UI卡顿"""
        try:
            if current != 'emp':
                self.emp_tab.refresh()
            if current != 'inv':
                self.inv_tab.refresh()
            if current != 'sal':
                self.sal_tab.refresh()
            if current != 'stat':
                self.stat_tab.refresh()
            if current != 'bank':
                self.bank_tab.refresh_accounts()
            if hasattr(self, 'backup_tab'):
                self.backup_tab.refresh_stats()
        except Exception:
            pass

    def set_status(self, msg):
        self.status_var.set(msg)
        try:
            log_status(msg)
        except Exception:
            pass
        self.after(3000, lambda: self.status_var.set("就绪"))


# ============================================================
# 可滚动Frame（所有标签页基类，确保窗口缩小时内容可滚动）
# ============================================================
class ScrollableTab(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        # 内容区最小宽度（子类可覆盖），小于此值出现水平滚动条
        self.min_content_w = 750
        # 根据主题设置Canvas背景色，消除白边
        _bg = '#000000' if is_dark_theme() else '#E6E7E8'
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=_bg, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hscroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.content = ttk.Frame(self.canvas)
        self.content.bind("<Configure>",
                          lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._canvas_window = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        # Canvas大小变化时，让content宽度至少填满Canvas；内容更宽时保持所需宽度，出现水平滚动条
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.scrollbar.set, xscrollcommand=self.hscroll.set)
        # pack顺序：先底部水平滚动条，再右侧垂直滚动条，最后canvas占剩余
        self.hscroll.pack(side="bottom", fill="x")
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        # 鼠标滚轮：普通滚轮垂直滚动，Shift+滚轮水平滚动
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _on_canvas_configure(self, event):
        """Canvas大小变化时：content宽度不小于min_content_w（空间不足时出现水平滚动条），
        空间充足时填满Canvas；高度至少填满可视区，避免底部露出Canvas背景"""
        self.canvas.itemconfig(self._canvas_window, width=max(event.width, self.min_content_w))
        need_h = self.content.winfo_reqheight()
        self.canvas.itemconfig(self._canvas_window, height=max(event.height, need_h))

    def _on_mousewheel(self, event):
        if event.state & 0x0001:  # Shift held → horizontal scroll
            self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ============================================================
# 人员管理
# ============================================================
class EmployeeTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.selected_id = None

        # 顶部表单
        form = ttk.LabelFrame(self.content, text="添加 / 修改人员", padding=10)
        form.pack(fill='x', padx=10, pady=8)

        ttk.Label(form, text="姓名:").grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.name_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.name_var, width=18).grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(form, text="月薪(元):").grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.salary_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.salary_var, width=14).grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(form, text="入职日期:").grid(row=0, column=4, sticky='e', padx=4, pady=4)
        self.hire_date_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.hire_date_var, width=12).grid(row=0, column=5, padx=4, pady=4)
        ttk.Label(form, text="(YYYY-MM-DD)", foreground='gray').grid(row=0, column=6, sticky='w', padx=2, pady=4)

        self.resigned_var = tk.BooleanVar(value=False)
        FlatCheckbutton(form, text="已离职", variable=self.resigned_var).grid(row=1, column=0, padx=8, pady=4)

        ttk.Label(form, text="离职日期:").grid(row=1, column=1, sticky='e', padx=4, pady=4)
        self.resign_date_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.resign_date_var, width=12).grid(row=1, column=2, padx=4, pady=4)
        ttk.Label(form, text="(YYYY-MM-DD)", foreground='gray').grid(row=1, column=3, sticky='w', padx=2, pady=4)

        self.insurance_var = tk.BooleanVar(value=False)
        FlatCheckbutton(form, text="五险一金已缴", variable=self.insurance_var).grid(row=1, column=4, padx=8, pady=4)

        ttk.Label(form, text="五险一金(个人):").grid(row=1, column=5, sticky='e', padx=4, pady=4)
        self.insurance_amount_var = tk.StringVar(value='0.00')
        ins_entry = ttk.Entry(form, textvariable=self.insurance_amount_var, width=12, state='readonly')
        ins_entry.grid(row=1, column=6, padx=4, pady=4)
        ttk.Button(form, text="设置细项", command=self.open_insurance_detail, width=8).grid(row=1, column=7, padx=2, pady=4)

        ttk.Label(form, text="备注:").grid(row=2, column=0, sticky='e', padx=4, pady=4)
        self.remark_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.remark_var, width=50).grid(row=2, column=1, columnspan=4,
                                                                      sticky='we', padx=4, pady=4)

        self.add_btn = ttk.Button(form, text="添加", command=self.add_or_update)
        self.add_btn.grid(row=2, column=5, padx=8, pady=4)
        self.cancel_btn = ttk.Button(form, text="取消修改", command=self.cancel_edit, state='disabled')
        self.cancel_btn.grid(row=2, column=6, padx=4, pady=4)

        # 表格
        tbl_frame = ttk.Frame(self.content)
        tbl_frame.pack(fill='both', expand=True, padx=10, pady=6)

        cols = ('id', 'name', 'monthly_salary', 'current_salary', 'hire_date', 'resign_date', 'insurance', 'remark')
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show='headings', selectmode='browse')
        self.tree.heading('id', text='编号')
        self.tree.heading('name', text='姓名')
        self.tree.heading('monthly_salary', text='基础月薪')
        self.tree.heading('current_salary', text='调薪后薪资')
        self.tree.heading('hire_date', text='入职日期')
        self.tree.heading('resign_date', text='离职日期')
        self.tree.heading('insurance', text='五险一金')
        self.tree.heading('remark', text='备注')
        self.tree.column('id', width=50, anchor='center')
        self.tree.column('name', width=110, anchor='center')
        self.tree.column('monthly_salary', width=90, anchor='e')
        self.tree.column('current_salary', width=90, anchor='e')
        self.tree.column('hire_date', width=85, anchor='center')
        self.tree.column('resign_date', width=85, anchor='center')
        self.tree.column('insurance', width=220, anchor='center')
        self.tree.column('remark', width=200, anchor='w')
        self.tree.pack(side='left', fill='both', expand=True)

        sb = ttk.Scrollbar(tbl_frame, orient='vertical', command=self.tree.yview)
        sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind('<Double-1>', self.on_double)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # 底部按钮
        btns = ttk.Frame(self.content)
        btns.pack(fill='x', padx=10, pady=6)
        ttk.Button(btns, text="删除选中", command=self.delete).pack(side='left', padx=4)
        ttk.Button(btns, text="调薪管理", command=self.open_salary_adjust).pack(side='left', padx=4)
        ttk.Button(btns, text="五险一金细项", command=self.open_insurance_detail).pack(side='left', padx=4)
        ttk.Button(btns, text="刷新", command=self.refresh).pack(side='left', padx=4)
        ttk.Label(btns, text="（双击记录可修改）", foreground='gray').pack(side='left', padx=10)

        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        conn = get_db()
        rows = conn.execute("SELECT * FROM employees ORDER BY id").fetchall()
        for r in rows:
            name = r['name']
            if r['resigned']:
                name += " (已离职)"
            # 五险一金：总金额/个人/公司（从细项调整表取最新）
            det_adj = conn.execute("SELECT * FROM insurance_detail_adjustments WHERE employee_id=? ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1", (r['id'],)).fetchone()
            if det_adj:
                det_pers = sum(float(det_adj[f'{k}_personal'] or 0) for k, _ in INSURANCE_ITEMS)
                det_comp = sum(float(det_adj[f'{k}_company'] or 0) for k, _ in INSURANCE_ITEMS)
                det_total = det_pers + det_comp
                ins_text = f"总{fmt_money(det_total)} 个{fmt_money(det_pers)} 公{fmt_money(det_comp)}"
            else:
                ins_amt = r['insurance_amount'] or 0
                if r['insurance_paid']:
                    ins_text = f"已缴 {fmt_money(ins_amt)}"
                else:
                    ins_text = f"未缴 {fmt_money(ins_amt)}" if ins_amt > 0 else "未缴"
            # 查询最新调薪后薪资
            adj = conn.execute(
                "SELECT new_salary FROM salary_adjustments WHERE employee_id=? "
                "ORDER BY effective_year DESC, effective_month DESC LIMIT 1",
                (r['id'],)).fetchone()
            if adj:
                cur_salary = fmt_money(adj['new_salary'])
            else:
                cur_salary = fmt_money(r['monthly_salary'])
            self.tree.insert('', 'end', values=(
                r['id'], name, fmt_money(r['monthly_salary']), cur_salary,
                r['hire_date'] or '', r['resign_date'] or '',
                ins_text, r['remark'] or ''))
        conn.close()

    def on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        self.selected_id = vals[0]

    def on_double(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        self.selected_id = vals[0]
        # 从数据库读取完整信息（因为表格中名字带了(已离职)后缀）
        conn = get_db()
        row = conn.execute("SELECT * FROM employees WHERE id=?", (self.selected_id,)).fetchone()
        conn.close()
        if not row:
            return
        self.name_var.set(row['name'])
        self.salary_var.set(str(row['monthly_salary'] or 0))
        self.hire_date_var.set(row['hire_date'] or '')
        self.resign_date_var.set(row['resign_date'] or '')
        self.resigned_var.set(bool(row['resigned']))
        self.insurance_var.set(bool(row['insurance_paid']))
        # 从五险一金细项表取最新的个人缴纳合计
        conn2 = get_db()
        det = conn2.execute("SELECT * FROM insurance_detail_adjustments WHERE employee_id=? ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1", (self.selected_id,)).fetchone()
        conn2.close()
        if det:
            pers_total = sum(float(det[f'{k}_personal'] or 0) for k, _ in INSURANCE_ITEMS)
            self.insurance_amount_var.set(f"{pers_total:.2f}")
        else:
            self.insurance_amount_var.set(str(row['insurance_amount'] or 0))
        self.remark_var.set(row['remark'] or '')
        self.add_btn.config(text="保存修改")
        self.cancel_btn.config(state='normal')

    def cancel_edit(self):
        self.selected_id = None
        self.name_var.set('')
        self.salary_var.set('')
        self.hire_date_var.set('')
        self.resign_date_var.set('')
        self.resigned_var.set(False)
        self.insurance_var.set(False)
        self.insurance_amount_var.set('0.00')
        self.remark_var.set('')
        self.add_btn.config(text="添加")
        self.cancel_btn.config(state='disabled')

    def add_or_update(self):
        name = self.name_var.get().strip()
        salary = self.salary_var.get().strip()
        hire_date = self.hire_date_var.get().strip()
        resign_date = self.resign_date_var.get().strip()
        resigned = 1 if self.resigned_var.get() else 0
        insurance_paid = 1 if self.insurance_var.get() else 0
        insurance_amount = self.insurance_amount_var.get().strip()
        remark = self.remark_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入姓名")
            return
        try:
            salary = float(salary) if salary else 0
        except ValueError:
            messagebox.showwarning("提示", "薪资必须是数字")
            return
        try:
            insurance_amount = float(insurance_amount) if insurance_amount else 0
        except ValueError:
            messagebox.showwarning("提示", "五险一金费用必须是数字")
            return
        # 校验日期格式
        for d_label, d_val in [("入职日期", hire_date), ("离职日期", resign_date)]:
            if d_val:
                try:
                    datetime.strptime(d_val, '%Y-%m-%d')
                except ValueError:
                    messagebox.showwarning("提示", f"{d_label}格式错误，请使用 YYYY-MM-DD 格式")
                    return
        conn = get_db()
        try:
            if self.selected_id:
                conn.execute("""UPDATE employees SET name=?, monthly_salary=?, hire_date=?, resign_date=?,
                             resigned=?, insurance_paid=?, insurance_amount=?, remark=? WHERE id=?""",
                             (name, salary, hire_date or None, resign_date or None,
                              resigned, insurance_paid, insurance_amount, remark, self.selected_id))
                self.app.set_status(f"已修改人员: {name}")
            else:
                conn.execute("""INSERT INTO employees(name, monthly_salary, hire_date, resign_date,
                             resigned, insurance_paid, insurance_amount, remark) VALUES(?,?,?,?,?,?,?,?)""",
                             (name, salary, hire_date or None, resign_date or None,
                              resigned, insurance_paid, insurance_amount, remark))
                self.app.set_status(f"已添加人员: {name}")
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "该姓名已存在")
            conn.close()
            return
        conn.close()
        self.cancel_edit()
        self.refresh()
        self.app.after(30, self.app._refresh_other_tabs, 'emp')

    def open_insurance_detail(self):
        # 五险一金细项管理（顶部总金额汇总，下面细项输入，支持添加/编辑/删除）
        if not self.selected_id:
            messagebox.showwarning("提示", "请先在列表中选择一个人员")
            return
        conn = get_db()
        emp = conn.execute("SELECT * FROM employees WHERE id=?", (self.selected_id,)).fetchone()
        conn.close()
        if not emp:
            return

        win = tk.Toplevel(self)
        win.title(f"五险一金 - {emp['name']}")
        win.geometry("860x640")
        win.minsize(700, 520)
        win.transient(self)

        editing_id = [None]

        # ===== 顶部：总金额汇总（自动从细项计算）=====
        total_frame = ttk.LabelFrame(win, text="五险一金总金额（自动汇总，不可手动输入）", padding=10)
        total_frame.pack(fill='x', padx=10, pady=6)

        ttk.Label(total_frame, text="公司缴纳合计:", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, padx=12, pady=4, sticky='e')
        sum_company_lbl = ttk.Label(total_frame, text="¥0.00", font=('微软雅黑', 11, 'bold'), foreground='#1565C0')
        sum_company_lbl.grid(row=0, column=1, padx=4, pady=4, sticky='w')

        ttk.Label(total_frame, text="个人缴纳合计:", font=('微软雅黑', 10, 'bold')).grid(row=0, column=2, padx=12, pady=4, sticky='e')
        sum_personal_lbl = ttk.Label(total_frame, text="¥0.00", font=('微软雅黑', 11, 'bold'), foreground='#C62828')
        sum_personal_lbl.grid(row=0, column=3, padx=4, pady=4, sticky='w')

        ttk.Label(total_frame, text="五险一金总额:", font=('微软雅黑', 10, 'bold')).grid(row=0, column=4, padx=12, pady=4, sticky='e')
        sum_total_lbl = ttk.Label(total_frame, text="¥0.00", font=('微软雅黑', 12, 'bold'), foreground='#2E7D32')
        sum_total_lbl.grid(row=0, column=5, padx=4, pady=4, sticky='w')

        ttk.Label(total_frame, text="（人员管理五险一金费用 = 个人缴纳合计）", foreground='gray', font=('微软雅黑', 8)).grid(row=1, column=0, columnspan=6, padx=4, pady=2)

        # ===== 细项输入表格 =====
        input_frame = ttk.LabelFrame(win, text="细项明细（填写各项公司/个人缴纳金额，元/月）", padding=8)
        input_frame.pack(fill='x', padx=10, pady=4)

        headers = ['项目', '公司缴纳', '个人缴纳', '小计']
        for ci, h in enumerate(headers):
            ttk.Label(input_frame, text=h, font=('微软雅黑', 9, 'bold')).grid(row=0, column=ci, padx=6, pady=4, sticky='n')

        ttk.Label(input_frame, text="生效:").grid(row=0, column=4, sticky='e', padx=4)
        det_year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Combobox(input_frame, textvariable=det_year_var, width=5,
                     values=[str(y) for y in range(_get_min_hire_year(), datetime.now().year + 1)],
                     state='readonly').grid(row=0, column=5, padx=2)
        ttk.Label(input_frame, text="年").grid(row=0, column=6, padx=1)
        det_month_var = tk.StringVar(value=str(datetime.now().month))
        ttk.Combobox(input_frame, textvariable=det_month_var, width=3,
                     values=[str(m) for m in range(1, 13)], state='readonly').grid(row=0, column=7, padx=2)
        ttk.Label(input_frame, text="月").grid(row=0, column=8, padx=1)

        entry_vars = {}
        subtotal_labels = {}
        for ri, (key, name) in enumerate(INSURANCE_ITEMS):
            ttk.Label(input_frame, text=name).grid(row=ri+1, column=0, padx=6, pady=2, sticky='w')
            comp_var = tk.StringVar(value='0')
            pers_var = tk.StringVar(value='0')
            ttk.Entry(input_frame, textvariable=comp_var, width=10).grid(row=ri+1, column=1, padx=4, pady=2)
            ttk.Entry(input_frame, textvariable=pers_var, width=10).grid(row=ri+1, column=2, padx=4, pady=2)
            sub_lbl = ttk.Label(input_frame, text="¥0.00")
            sub_lbl.grid(row=ri+1, column=3, padx=4, pady=2)
            entry_vars[key] = (comp_var, pers_var)
            subtotal_labels[key] = sub_lbl

        def recalc_all(*args):
            tc = tp = 0
            for key in entry_vars:
                try:
                    c = float(entry_vars[key][0].get() or 0)
                    p = float(entry_vars[key][1].get() or 0)
                except ValueError:
                    c = p = 0
                subtotal_labels[key].config(text=f"¥{c+p:,.2f}")
                tc += c
                tp += p
            sum_company_lbl.config(text=f"¥{tc:,.2f}")
            sum_personal_lbl.config(text=f"¥{tp:,.2f}")
            sum_total_lbl.config(text=f"¥{tc+tp:,.2f}")

        for key in entry_vars:
            entry_vars[key][0].trace_add('write', recalc_all)
            entry_vars[key][1].trace_add('write', recalc_all)

        ttk.Label(input_frame, text="原因:").grid(row=7, column=0, sticky='e', padx=4, pady=4)
        det_reason_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=det_reason_var, width=50).grid(row=7, column=1, columnspan=5, sticky='we', padx=4, pady=4)

        def clear_form():
            for key in entry_vars:
                entry_vars[key][0].set('0')
                entry_vars[key][1].set('0')
            det_reason_var.set('')
            det_year_var.set(str(datetime.now().year))
            det_month_var.set(str(datetime.now().month))
            editing_id[0] = None
            save_btn.config(text="添加记录")
            cancel_edit_btn.config(state='disabled')
            recalc_all()

        def save_detail():
            try:
                yr = int(det_year_var.get())
                mo = int(det_month_var.get())
            except ValueError:
                messagebox.showwarning("提示", "请选择生效年月")
                return
            values = {}
            tc = tp = 0
            for key in entry_vars:
                try:
                    cv = float(entry_vars[key][0].get() or 0)
                    pv = float(entry_vars[key][1].get() or 0)
                except ValueError:
                    messagebox.showwarning("提示", f"{INSURANCE_ITEMS_DICT[key]}金额必须是数字")
                    return
                values[f'{key}_company'] = cv
                values[f'{key}_personal'] = pv
                tc += cv
                tp += pv
            reason = det_reason_var.get().strip()
            conn = get_db()
            if editing_id[0]:
                set_clause = ','.join([f"{k}=?" for k in values.keys()])
                conn.execute(f"UPDATE insurance_detail_adjustments SET effective_year=?, effective_month=?, {set_clause}, reason=? WHERE id=?",
                             (yr, mo, *values.values(), reason, editing_id[0]))
                self.app.set_status(f"已修改五险一金：{yr}年{mo}月起 个人{fmt_money(tp)}")
            else:
                cols = ','.join(values.keys())
                placeholders = ','.join(['?'] * len(values))
                conn.execute(f"INSERT INTO insurance_detail_adjustments(employee_id, effective_year, effective_month, {cols}, reason, created_at) VALUES(?,?,?,{placeholders},?,?)",
                             (self.selected_id, yr, mo, *values.values(), reason, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                self.app.set_status(f"已添加五险一金：{yr}年{mo}月起 个人{fmt_money(tp)}")
            # 同步更新人员表的五险一金费用为个人总金额
            conn.execute("UPDATE employees SET insurance_amount=? WHERE id=?", (tp, self.selected_id))
            conn.commit()
            conn.close()
            clear_form()
            refresh_detail_list()
            self.app.after(30, self.refresh)

        save_btn = ttk.Button(input_frame, text="添加记录", command=save_detail)
        save_btn.grid(row=7, column=6, padx=4, pady=4, sticky='we')
        cancel_edit_btn = ttk.Button(input_frame, text="取消编辑", command=clear_form, state='disabled')
        cancel_edit_btn.grid(row=7, column=7, columnspan=2, padx=4, pady=4, sticky='we')

        # ===== 调整历史列表 =====
        list_frame = ttk.LabelFrame(win, text="历史记录（双击加载到上方编辑，右键可删除/编辑）", padding=6)
        list_frame.pack(fill='both', expand=True, padx=10, pady=6)

        det_cols = ('id', 'effective', 'company', 'personal', 'total', 'pension', 'medical', 'unemployment', 'injury', 'maternity', 'housing', 'reason')
        det_tree = ttk.Treeview(list_frame, columns=det_cols, show='headings', selectmode='browse')
        col_info = [('id', '编号', 45), ('effective', '生效年月', 80), ('company', '公司', 75),
                    ('personal', '个人', 75), ('total', '合计', 80), ('pension', '养老', 65),
                    ('medical', '医疗', 65), ('unemployment', '失业', 60), ('injury', '工伤', 55),
                    ('maternity', '生育', 55), ('housing', '公积金', 70), ('reason', '原因', 120)]
        for cid, text, w in col_info:
            det_tree.heading(cid, text=text)
            det_tree.column(cid, width=w, anchor='center')
        det_tree.pack(side='left', fill='both', expand=True)
        det_sb = ttk.Scrollbar(list_frame, orient='vertical', command=det_tree.yview)
        det_sb.pack(side='right', fill='y')
        det_tree.configure(yscrollcommand=det_sb.set)

        def refresh_detail_list():
            for i in det_tree.get_children():
                det_tree.delete(i)
            conn = get_db()
            rows = conn.execute("SELECT * FROM insurance_detail_adjustments WHERE employee_id=? ORDER BY effective_year DESC, effective_month DESC, id DESC", (self.selected_id,)).fetchall()
            conn.close()
            for r in rows:
                tc = sum(float(r[f'{k}_company'] or 0) for k, _ in INSURANCE_ITEMS)
                tp = sum(float(r[f'{k}_personal'] or 0) for k, _ in INSURANCE_ITEMS)
                det_tree.insert('', 'end', values=(
                    r['id'], f"{r['effective_year']}年{r['effective_month']}月",
                    f"¥{tc:.0f}", f"¥{tp:.0f}", f"¥{tc+tp:,.2f}",
                    f"¥{float(r['pension_company'] or 0)+float(r['pension_personal'] or 0):.0f}",
                    f"¥{float(r['medical_company'] or 0)+float(r['medical_personal'] or 0):.0f}",
                    f"¥{float(r['unemployment_company'] or 0)+float(r['unemployment_personal'] or 0):.0f}",
                    f"¥{float(r['injury_company'] or 0)+float(r['injury_personal'] or 0):.0f}",
                    f"¥{float(r['maternity_company'] or 0)+float(r['maternity_personal'] or 0):.0f}",
                    f"¥{float(r['housing_company'] or 0)+float(r['housing_personal'] or 0):.0f}",
                    r['reason'] or ''))

        def edit_detail(event=None):
            sel = det_tree.selection()
            if not sel:
                return
            det_id = det_tree.item(sel[0], 'values')[0]
            conn = get_db()
            row = conn.execute("SELECT * FROM insurance_detail_adjustments WHERE id=?", (det_id,)).fetchone()
            conn.close()
            if not row:
                return
            det_year_var.set(str(row['effective_year']))
            det_month_var.set(str(row['effective_month']))
            det_reason_var.set(row['reason'] or '')
            for key in entry_vars:
                entry_vars[key][0].set(str(float(row[f'{key}_company'] or 0)))
                entry_vars[key][1].set(str(float(row[f'{key}_personal'] or 0)))
            editing_id[0] = det_id
            save_btn.config(text="保存修改")
            cancel_edit_btn.config(state='normal')
            recalc_all()

        def delete_detail():
            sel = det_tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先在列表中选择要删除的记录")
                return
            det_id = det_tree.item(sel[0], 'values')[0]
            eff_text = det_tree.item(sel[0], 'values')[1]
            if not messagebox.askyesno("确认删除", f"确定删除 {eff_text} 的五险一金记录？\n\n删除后人员表的五险一金费用将自动重新计算。"):
                return
            try:
                conn = get_db()
                conn.execute("DELETE FROM insurance_detail_adjustments WHERE id=?", (det_id,))
                latest = conn.execute("SELECT * FROM insurance_detail_adjustments WHERE employee_id=? ORDER BY effective_year DESC, effective_month DESC, id DESC LIMIT 1", (self.selected_id,)).fetchone()
                if latest:
                    tp = sum(float(latest[f'{k}_personal'] or 0) for k, _ in INSURANCE_ITEMS)
                else:
                    tp = 0
                conn.execute("UPDATE employees SET insurance_amount=? WHERE id=?", (tp, self.selected_id))
                conn.commit()
                conn.close()
                if editing_id[0] == det_id:
                    clear_form()
                refresh_detail_list()
                self.app.after(30, self.refresh)
                self.app.set_status(f"已删除 {eff_text} 的五险一金记录")
            except Exception as e:
                messagebox.showerror("删除失败", str(e))

        det_tree.bind('<Double-1>', edit_detail)
        # 右键菜单
        ctx_menu = tk.Menu(win, tearoff=0)
        ctx_menu.add_command(label="编辑选中记录", command=lambda: edit_detail())
        ctx_menu.add_command(label="删除选中记录", command=delete_detail)
        def show_ctx(event):
            iid = det_tree.identify_row(event.y)
            if iid:
                det_tree.selection_set(iid)
                ctx_menu.tk_popup(event.x_root, event.y_root)
        det_tree.bind('<Button-3>', show_ctx)

        # 底部按钮（在所有函数定义之后创建，pack到最底部）
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', side='bottom', padx=10, pady=4)
        ttk.Button(btn_frame, text="删除选中记录", command=delete_detail).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="编辑选中记录", command=lambda: edit_detail()).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="清空表单", command=clear_form).pack(side='left', padx=4)

        refresh_detail_list()
        recalc_all()

    def open_insurance_adjust(self):
        # 打开五险一金调整管理窗口
        if not self.selected_id:
            messagebox.showwarning("提示", "请先在列表中选择一个人员")
            return
        conn = get_db()
        emp = conn.execute("SELECT * FROM employees WHERE id=?", (self.selected_id,)).fetchone()
        conn.close()
        if not emp:
            return

        win = tk.Toplevel(self)
        win.title(f"五险一金调整 - {emp['name']}")
        win.geometry("620x560")
        win.minsize(560, 480)
        win.transient(self)
        win.grab_set()

        info_frame = ttk.LabelFrame(win, text="当前五险一金信息", padding=8)
        info_frame.pack(fill='x', padx=10, pady=8)
        ttk.Label(info_frame, text=f"姓名：{emp['name']}").grid(row=0, column=0, padx=8, pady=4)
        base_amt = float(emp['insurance_amount'] or 0)
        ttk.Label(info_frame, text=f"基础金额：{fmt_money(base_amt)}").grid(row=0, column=1, padx=8, pady=4)
        paid_text = "已缴纳" if emp['insurance_paid'] else "未缴纳"
        ttk.Label(info_frame, text=f"状态：{paid_text}").grid(row=0, column=2, padx=8, pady=4)
        ttk.Label(info_frame, text=f"入职日期：{emp['hire_date'] or '未设置'}").grid(row=0, column=3, padx=8, pady=4)

        add_frame = ttk.LabelFrame(win, text="添加调整记录", padding=8)
        add_frame.pack(fill='x', padx=10, pady=4)

        ttk.Label(add_frame, text="生效年份:").grid(row=0, column=0, sticky='e', padx=4, pady=4)
        adj_year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Combobox(add_frame, textvariable=adj_year_var, width=6,
                     values=[str(y) for y in range(_get_min_hire_year(), datetime.now().year + 1)],
                     state='readonly').grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(add_frame, text="生效月份:").grid(row=0, column=2, sticky='e', padx=4, pady=4)
        adj_month_var = tk.StringVar(value=str(datetime.now().month))
        ttk.Combobox(add_frame, textvariable=adj_month_var, width=4,
                     values=[str(m) for m in range(1, 13)], state='readonly').grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(add_frame, text="新金额(元):").grid(row=0, column=4, sticky='e', padx=4, pady=4)
        adj_amount_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=adj_amount_var, width=12).grid(row=0, column=5, padx=4, pady=4)

        ttk.Label(add_frame, text="原因:").grid(row=1, column=0, sticky='e', padx=4, pady=4)
        adj_reason_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=adj_reason_var, width=40).grid(row=1, column=1, columnspan=4, sticky='we', padx=4, pady=4)

        def add_adjustment():
            try:
                year = int(adj_year_var.get())
                month = int(adj_month_var.get())
                new_amount = float(adj_amount_var.get()) if adj_amount_var.get() else 0
            except ValueError:
                messagebox.showwarning("提示", "请输入有效的年份、月份和金额")
                return
            if new_amount < 0:
                messagebox.showwarning("提示", "金额不能为负数")
                return
            reason = adj_reason_var.get().strip()
            conn = get_db()
            conn.execute("INSERT INTO insurance_adjustments(employee_id, effective_year, effective_month, new_amount, reason, created_at) VALUES(?,?,?,?,?,?)",
                         (self.selected_id, year, month, new_amount, reason, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()
            adj_amount_var.set('')
            adj_reason_var.set('')
            refresh_adj_list()
            self.app.set_status(f"已添加五险一金调整：{year}年{month}月起 {fmt_money(new_amount)}")

        ttk.Button(add_frame, text="添加调整", command=add_adjustment).grid(row=1, column=5, padx=8, pady=4)

        list_frame = ttk.LabelFrame(win, text="调整历史", padding=6)
        list_frame.pack(fill='both', expand=True, padx=10, pady=6)

        adj_cols = ('id', 'effective', 'new_amount', 'reason', 'created_at')
        adj_tree = ttk.Treeview(list_frame, columns=adj_cols, show='headings', selectmode='browse')
        for cid, text, w, anchor in [('id', '编号', 50, 'center'), ('effective', '生效年月', 100, 'center'), ('new_amount', '新金额', 100, 'e'), ('reason', '原因', 200, 'w'), ('created_at', '创建时间', 140, 'center')]:
            adj_tree.heading(cid, text=text)
            adj_tree.column(cid, width=w, anchor=anchor)
        adj_tree.pack(side='left', fill='both', expand=True)
        adj_sb = ttk.Scrollbar(list_frame, orient='vertical', command=adj_tree.yview)
        adj_sb.pack(side='right', fill='y')
        adj_tree.configure(yscrollcommand=adj_sb.set)
        adj_tree.bind('<Double-1>', edit_adjustment)

        def refresh_adj_list():
            for i in adj_tree.get_children():
                adj_tree.delete(i)
            conn = get_db()
            rows = conn.execute("SELECT * FROM insurance_adjustments WHERE employee_id=? ORDER BY effective_year DESC, effective_month DESC, id DESC", (self.selected_id,)).fetchall()
            conn.close()
            for r in rows:
                adj_tree.insert('', 'end', values=(r['id'], f"{r['effective_year']}年{r['effective_month']}月", fmt_money(r['new_amount']), r['reason'] or '', r['created_at'] or ''))

        def delete_adjustment():
            sel = adj_tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先选择要删除的调整记录")
                return
            adj_id = adj_tree.item(sel[0], 'values')[0]
            if not messagebox.askyesno("确认", "确定删除该调整记录？"):
                return
            conn = get_db()
            conn.execute("DELETE FROM insurance_adjustments WHERE id=?", (adj_id,))
            conn.commit()
            conn.close()
            refresh_adj_list()
            self.app.set_status("已删除五险一金调整记录")

        del_btn = ttk.Button(win, text="删除选中调整", command=delete_adjustment)
        del_btn.pack(side='bottom', pady=6)

        refresh_adj_list()

    def delete(self):
        if not self.selected_id:
            messagebox.showwarning("提示", "请先选择要删除的人员")
            return
        conn = get_db()
        row = conn.execute("SELECT name FROM employees WHERE id=?", (self.selected_id,)).fetchone()
        name = row['name'] if row else ''
        # 检查关联
        inv_cnt = conn.execute("SELECT COUNT(*) c FROM invoices WHERE reimburser_id=?",
                               (self.selected_id,)).fetchone()['c']
        sal_cnt = conn.execute("SELECT COUNT(*) c FROM salary_payments WHERE employee_id=?",
                               (self.selected_id,)).fetchone()['c']
        conn.close()
        msg = f"确定删除人员「{name}」？"
        if inv_cnt or sal_cnt:
            msg += f"\n\n该人员关联 {inv_cnt} 条发票记录（将保留，报销人显示为'已删除'）"
            msg += f"\n关联 {sal_cnt} 条工资发放记录（将一并删除）。"
        if messagebox.askyesno("确认删除", msg):
            conn = get_db()
            try:
                # 删除该人员所有关联记录
                conn.execute("DELETE FROM salary_payments WHERE employee_id=?", (self.selected_id,))
                conn.execute("DELETE FROM salary_adjustments WHERE employee_id=?", (self.selected_id,))
                conn.execute("DELETE FROM insurance_adjustments WHERE employee_id=?", (self.selected_id,))
                conn.execute("DELETE FROM insurance_detail_adjustments WHERE employee_id=?", (self.selected_id,))
                # 发票记录保留但解除外键关联
                conn.execute("UPDATE invoices SET reimburser_id=NULL WHERE reimburser_id=?",
                             (self.selected_id,))
                conn.execute("DELETE FROM employees WHERE id=?", (self.selected_id,))
                conn.commit()
                conn.close()
            except Exception as e:
                conn.close()
                messagebox.showerror("删除失败", str(e))
                return
            self.cancel_edit()
            self.refresh()  # 立即刷新当前页
            self.app.set_status(f"已删除人员: {name}")
            # 延迟异步刷新其他页，避免UI卡顿
            self.app.after(30, self.app._refresh_other_tabs, 'emp')

    def open_salary_adjust(self):
        """打开调薪管理窗口"""
        if not self.selected_id:
            messagebox.showwarning("提示", "请先在列表中选择一个人员")
            return
        conn = get_db()
        emp = conn.execute("SELECT * FROM employees WHERE id=?", (self.selected_id,)).fetchone()
        conn.close()
        if not emp:
            return

        win = tk.Toplevel(self)
        win.title(f"调薪管理 - {emp['name']}")
        win.geometry("620x560")
        win.minsize(560, 480)
        win.transient(self)

        # 当前薪资信息
        info_frame = ttk.LabelFrame(win, text="当前薪资信息", padding=8)
        info_frame.pack(fill='x', padx=10, pady=8)
        ttk.Label(info_frame, text=f"姓名：{emp['name']}").grid(row=0, column=0, padx=8, pady=4)
        ttk.Label(info_frame, text=f"基础月薪：{fmt_money(emp['monthly_salary'])}").grid(row=0, column=1, padx=8, pady=4)
        ttk.Label(info_frame, text=f"入职日期：{emp['hire_date'] or '未设置'}").grid(row=0, column=2, padx=8, pady=4)

        # 添加调薪记录表单
        add_frame = ttk.LabelFrame(win, text="添加调薪记录", padding=8)
        add_frame.pack(fill='x', padx=10, pady=4)

        ttk.Label(add_frame, text="生效年份:").grid(row=0, column=0, sticky='e', padx=4, pady=4)
        adj_year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Combobox(add_frame, textvariable=adj_year_var, width=6,
                     values=[str(y) for y in range(_get_min_hire_year(), datetime.now().year + 1)],
                     state='readonly').grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(add_frame, text="生效月份:").grid(row=0, column=2, sticky='e', padx=4, pady=4)
        adj_month_var = tk.StringVar(value=str(datetime.now().month))
        ttk.Combobox(add_frame, textvariable=adj_month_var, width=4,
                     values=[str(m) for m in range(1, 13)], state='readonly').grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(add_frame, text="新月薪(元):").grid(row=0, column=4, sticky='e', padx=4, pady=4)
        adj_salary_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=adj_salary_var, width=12).grid(row=0, column=5, padx=4, pady=4)

        ttk.Label(add_frame, text="原因:").grid(row=1, column=0, sticky='e', padx=4, pady=4)
        adj_reason_var = tk.StringVar()
        ttk.Entry(add_frame, textvariable=adj_reason_var, width=40).grid(row=1, column=1, columnspan=4, sticky='we', padx=4, pady=4)

        editing_adj_id = [None]

        def clear_adj_form():
            adj_salary_var.set('')
            adj_reason_var.set('')
            adj_year_var.set(str(datetime.now().year))
            adj_month_var.set(str(datetime.now().month))
            editing_adj_id[0] = None
            add_btn.config(text="添加调薪")
            cancel_btn.config(state='disabled')

        def save_adjustment():
            try:
                year = int(adj_year_var.get())
                month = int(adj_month_var.get())
                new_salary = float(adj_salary_var.get()) if adj_salary_var.get() else 0
            except ValueError:
                messagebox.showwarning("提示", "请输入有效的年份、月份和薪资")
                return
            if new_salary <= 0:
                messagebox.showwarning("提示", "薪资必须大于0")
                return
            reason = adj_reason_var.get().strip()
            conn = get_db()
            if editing_adj_id[0]:
                conn.execute("UPDATE salary_adjustments SET effective_year=?, effective_month=?, new_salary=?, reason=? WHERE id=?",
                             (year, month, new_salary, reason, editing_adj_id[0]))
                self.app.set_status(f"已修改调薪记录：{year}年{month}月起 {fmt_money(new_salary)}")
            else:
                conn.execute("INSERT INTO salary_adjustments(employee_id, effective_year, effective_month, new_salary, reason, created_at) VALUES(?,?,?,?,?,?)",
                             (self.selected_id, year, month, new_salary, reason, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                self.app.set_status(f"已添加调薪记录：{year}年{month}月起 {fmt_money(new_salary)}")
            conn.commit()
            conn.close()
            clear_adj_form()
            refresh_adj_list()
            self.app.after(30, self.refresh)

        def edit_adjustment(event=None):
            sel = adj_tree.selection()
            if not sel:
                return
            adj_id = adj_tree.item(sel[0], 'values')[0]
            conn = get_db()
            row = conn.execute("SELECT * FROM salary_adjustments WHERE id=?", (adj_id,)).fetchone()
            conn.close()
            if not row:
                return
            adj_year_var.set(str(row['effective_year']))
            adj_month_var.set(str(row['effective_month']))
            adj_salary_var.set(str(row['new_salary']))
            adj_reason_var.set(row['reason'] or '')
            editing_adj_id[0] = adj_id
            add_btn.config(text="保存修改")
            cancel_btn.config(state='normal')

        add_btn = ttk.Button(add_frame, text="添加调薪", command=save_adjustment)
        add_btn.grid(row=1, column=5, padx=4, pady=4)
        cancel_btn = ttk.Button(add_frame, text="取消编辑", command=clear_adj_form, state='disabled')
        cancel_btn.grid(row=1, column=4, padx=4, pady=4, sticky='e')

        # 底部按钮（先pack到bottom，确保不被列表挤出）
        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x', side='bottom', padx=10, pady=6)
        ttk.Button(btn_frame, text="删除选中调薪", command=lambda: delete_adjustment()).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="编辑选中调薪", command=lambda: edit_adjustment()).pack(side='left', padx=4)
        ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side='right', padx=4)

        # 调薪历史列表
        list_frame = ttk.LabelFrame(win, text="调薪历史（双击可编辑）", padding=6)
        list_frame.pack(fill='both', expand=True, padx=10, pady=6)

        adj_cols = ('id', 'effective', 'new_salary', 'reason', 'created_at')
        adj_tree = ttk.Treeview(list_frame, columns=adj_cols, show='headings', selectmode='browse')
        for cid, text, w, anchor in [('id', '编号', 50, 'center'),
                                     ('effective', '生效年月', 100, 'center'),
                                     ('new_salary', '新月薪', 100, 'e'),
                                     ('reason', '原因', 200, 'w'),
                                     ('created_at', '创建时间', 140, 'center')]:
            adj_tree.heading(cid, text=text)
            adj_tree.column(cid, width=w, anchor=anchor)
        adj_tree.pack(side='left', fill='both', expand=True)
        adj_sb = ttk.Scrollbar(list_frame, orient='vertical', command=adj_tree.yview)
        adj_sb.pack(side='right', fill='y')
        adj_tree.configure(yscrollcommand=adj_sb.set)
        adj_tree.bind('<Double-1>', edit_adjustment)

        def refresh_adj_list():
            for i in adj_tree.get_children():
                adj_tree.delete(i)
            conn = get_db()
            rows = conn.execute("""SELECT * FROM salary_adjustments WHERE employee_id=?
                                  ORDER BY effective_year DESC, effective_month DESC, id DESC""",
                               (self.selected_id,)).fetchall()
            conn.close()
            for r in rows:
                adj_tree.insert('', 'end', values=(
                    r['id'], f"{r['effective_year']}年{r['effective_month']:02d}月",
                    fmt_money(r['new_salary']), r['reason'] or '', r['created_at'] or ''))

        def delete_adjustment():
            sel = adj_tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先选择要删除的调薪记录")
                return
            adj_id = adj_tree.item(sel[0], 'values')[0]
            if not messagebox.askyesno("确认", "确定删除这条调薪记录吗？"):
                return
            conn = get_db()
            conn.execute("DELETE FROM salary_adjustments WHERE id=?", (adj_id,))
            conn.commit()
            conn.close()
            refresh_adj_list()
            self.app.set_status("已删除调薪记录")

        refresh_adj_list()


# ============================================================
# 发票报销
# ============================================================
class InvoiceTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.selected_id = None
        self.image_file = None  # 临时上传的图片路径
        self.min_content_w = 700  # 发票页表单最小宽度，小于此值出现水平滚动条

        # 重新排列pack顺序：先放右侧固定栏，最后放左侧expand的canvas，避免sash被挤出窗口
        self.canvas.pack_forget()
        self.scrollbar.pack_forget()
        self.hscroll.pack_forget()

        # 右侧发票预览列（直接pack在self上，不移动canvas，避免渲染问题）
        self.preview_panel = ttk.LabelFrame(self, text="发票预览（双击放大 / 拖动左边缘调整宽度）", padding=4,
                                            width=420)
        self.preview_panel.pack(side='right', fill='y', padx=(4, 0))
        self.preview_panel.pack_propagate(False)
        self._preview_photo = None
        self._preview_path = None
        self._preview_orig_img = None  # 缓存PIL原图
        self._preview_scale = 1.0

        # 顶部缩放工具栏
        toolbar = ttk.Frame(self.preview_panel)
        toolbar.pack(fill='x', pady=(0, 4))
        # 紧凑按钮样式
        _pb_style = 'PrevBar.TButton'
        ttk.Style().configure(_pb_style, padding=(4, 2), font=('', 8))
        for c in range(6):
            toolbar.columnconfigure(c, weight=1)
        ttk.Button(toolbar, text="−", width=2, style=_pb_style, command=self._preview_zoom_out).grid(row=0, column=0, padx=1, pady=1, sticky='we')
        _zoom_fg = '#F5F5F7' if is_dark_theme() else '#1D1D1F'
        _zoom_bg = '#1C1C1E' if is_dark_theme() else '#F0F1F2'
        self._zoom_label = tk.Label(toolbar, text="", width=4, anchor='center', font=('', 8), fg=_zoom_fg, bg=_zoom_bg)
        self._zoom_label.grid(row=0, column=1, padx=1, pady=1, sticky='we')
        ttk.Button(toolbar, text="+", width=2, style=_pb_style, command=self._preview_zoom_in).grid(row=0, column=2, padx=1, pady=1, sticky='we')
        ttk.Button(toolbar, text="适应", width=2, style=_pb_style, command=self._preview_fit).grid(row=0, column=3, padx=1, pady=1, sticky='we')
        ttk.Button(toolbar, text="1:1", width=2, style=_pb_style, command=self._preview_actual).grid(row=0, column=4, padx=1, pady=1, sticky='we')
        self._ocr_toggle_btn = ttk.Button(toolbar, text="OCR", width=3, style=_pb_style, command=self._toggle_preview_ocr)
        self._ocr_toggle_btn.grid(row=0, column=5, padx=1, pady=1, sticky='we')
        self._ocr_overlay_on = False
        self._ocr_results = []  # [(box, text, conf), ...] 原图坐标
        self._ocr_box_items = []  # [(canvas_item_id, result_index), ...]
        self._ocr_selected = set()  # 选中的结果索引
        self._ocr_drag = None  # 拖拽选择状态
        self._ocr_sel_rect = None  # 选择矩形canvas item

        # 预览Canvas + 双向滚动条
        prev_frame = ttk.Frame(self.preview_panel)
        prev_frame.pack(fill='both', expand=True)
        canvas_bg = '#141414' if is_dark_theme() else '#E6E7E8'
        self.preview_canvas = tk.Canvas(prev_frame, bg=canvas_bg, highlightthickness=0)
        self.preview_vscroll = ttk.Scrollbar(prev_frame, orient='vertical', command=self.preview_canvas.yview)
        self.preview_hscroll = ttk.Scrollbar(prev_frame, orient='horizontal', command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=self.preview_vscroll.set,
                                      xscrollcommand=self.preview_hscroll.set)
        self.preview_vscroll.pack(side='right', fill='y')
        self.preview_hscroll.pack(side='bottom', fill='x')
        self.preview_canvas.pack(side='left', fill='both', expand=True)
        self.preview_canvas.bind('<Double-1>', lambda e: self._preview_view_big())
        self.preview_canvas.bind('<Configure>', self._on_preview_resize)
        # Ctrl+滚轮缩放
        self.preview_canvas.bind('<Control-MouseWheel>', self._on_preview_wheel_zoom)
        # 点击OCR框复制文字 / 拖拽框选
        self.preview_canvas.bind('<Button-1>', self._on_preview_click_ocr)
        self.preview_canvas.bind('<B1-Motion>', self._on_preview_drag_ocr)
        self.preview_canvas.bind('<ButtonRelease-1>', self._on_preview_release_ocr)
        self.preview_canvas.bind('<Motion>', self._on_preview_ocr_hover)
        self.preview_canvas.bind('<Leave>', self._on_preview_ocr_leave)
        # 键盘滚动快捷键：方向键/WASD滚动，+/-缩放
        self.preview_canvas.bind('<Up>', lambda e: self._preview_key_scroll(0, -40))
        self.preview_canvas.bind('<Down>', lambda e: self._preview_key_scroll(0, 40))
        self.preview_canvas.bind('<Left>', lambda e: self._preview_key_scroll(-40, 0))
        self.preview_canvas.bind('<Right>', lambda e: self._preview_key_scroll(40, 0))
        self.preview_canvas.bind('<w>', lambda e: self._preview_key_scroll(0, -40))
        self.preview_canvas.bind('<s>', lambda e: self._preview_key_scroll(0, 40))
        self.preview_canvas.bind('<a>', lambda e: self._preview_key_scroll(-40, 0))
        self.preview_canvas.bind('<d>', lambda e: self._preview_key_scroll(40, 0))
        self.preview_canvas.bind('<W>', lambda e: self._preview_key_scroll(0, -40))
        self.preview_canvas.bind('<S>', lambda e: self._preview_key_scroll(0, 40))
        self.preview_canvas.bind('<A>', lambda e: self._preview_key_scroll(-40, 0))
        self.preview_canvas.bind('<D>', lambda e: self._preview_key_scroll(40, 0))
        self.preview_canvas.bind('<plus>', lambda e: self._preview_zoom_in())
        self.preview_canvas.bind('<minus>', lambda e: self._preview_zoom_out())
        self.preview_canvas.bind('<equal>', lambda e: self._preview_zoom_in())
        self.preview_canvas.focus_set()

        # OCR文字面板（可折叠，位于预览Canvas下方）
        self._ocr_panel_visible = False
        self.ocr_panel = ttk.Frame(self.preview_panel)
        ocr_header = ttk.Frame(self.ocr_panel)
        ocr_header.pack(fill='x')
        ttk.Button(ocr_header, text="隐藏", width=5, command=self._hide_ocr_panel).pack(side='right', padx=2)
        ttk.Label(ocr_header, text="在图片上拖拽框选文字自动复制，也可在下方选择复制", font=('', 8)).pack(side='left', padx=2)
        ocr_text_frame = ttk.Frame(self.ocr_panel)
        ocr_text_frame.pack(fill='both', expand=True)
        ocr_bg = '#1a1a1a' if is_dark_theme() else '#ffffff'
        ocr_fg = '#e0e0e0' if is_dark_theme() else '#222222'
        self.ocr_text = tk.Text(ocr_text_frame, height=8, wrap='word', font=('Consolas', 9),
                                bg=ocr_bg, fg=ocr_fg, insertbackground=ocr_fg,
                                selectbackground='#3B82F6', selectforeground='#ffffff',
                                relief='flat', padx=6, pady=4)
        ocr_text_scroll = ttk.Scrollbar(ocr_text_frame, orient='vertical', command=self.ocr_text.yview)
        self.ocr_text.configure(yscrollcommand=ocr_text_scroll.set)
        ocr_text_scroll.pack(side='right', fill='y')
        self.ocr_text.pack(side='left', fill='both', expand=True)
        self.ocr_text.bind('<Button-1>', self._on_ocr_text_click)
        self._ocr_line_ranges = []  # [(start_index, end_index, text), ...]

        # 可拖拽分隔条（加宽到12px，带握把纹理，hover高亮，双击重置）
        self._sash_bg_normal = '#3a3a3a' if is_dark_theme() else '#D0D0D0'
        self._sash_bg_hover = '#666666' if is_dark_theme() else '#999999'
        self._sash_grip = '#888888' if is_dark_theme() else '#AAAAAA'
        self.sash = tk.Canvas(self, width=12, cursor='sb_h_double_arrow', bg=self._sash_bg_normal,
                              highlightthickness=0, bd=0)
        self.sash.pack(side='right', fill='y')
        self._sash_drag = {'x': 0, 'w': 0, 'active': False}
        self.sash.bind('<Button-1>', self._on_sash_press)
        self.sash.bind('<B1-Motion>', self._on_sash_drag)
        self.sash.bind('<ButtonRelease-1>', self._on_sash_release)
        self.sash.bind('<Enter>', lambda e: self.sash.config(bg=self._sash_bg_hover))
        self.sash.bind('<Leave>', lambda e: self.sash.config(bg=self._sash_bg_normal))
        self.sash.bind('<Double-1>', self._on_sash_reset)
        self.sash.bind('<Configure>', self._draw_sash_grip)

        # 最后pack：滚动条最右，预览列在滚动条左，sash在预览列左，canvas占剩余
        self.scrollbar.pack(side='right', fill='y')
        self.hscroll.pack(side='bottom', fill='x')
        self.canvas.pack(side='left', fill='both', expand=True)

        # 窗口大小变化时自动收窄预览列，保证左侧至少750px
        self.bind('<Configure>', self._on_tab_resize)
        self._last_tab_w = 0

        # 顶部表单
        top_container = ttk.Frame(self.content)
        top_container.pack(fill='x', padx=10, pady=8)
        form = ttk.LabelFrame(top_container, text="添加 / 修改发票报销", padding=10)
        form.pack(fill='x')
        # 8列布局：4个标签列 + 4个输入列，按钮独立一行不与输入框抢宽度
        for col in range(8):
            if col in (0, 2, 4, 6):
                form.columnconfigure(col, weight=0, minsize=65)
            else:
                form.columnconfigure(col, weight=1)

        ttk.Label(form, text="发票号:").grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.inv_no_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.inv_no_var).grid(row=0, column=1, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="类型:").grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.inv_type_var = tk.StringVar(value='发票')
        ttk.Combobox(form, textvariable=self.inv_type_var, values=['发票', '支付记录'],
                     state='readonly').grid(row=0, column=3, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="报销人:").grid(row=0, column=4, sticky='e', padx=4, pady=4)
        self.reimburser_var = tk.StringVar()
        self.reimburser_cb = ttk.Combobox(form, textvariable=self.reimburser_var, state='readonly')
        self.reimburser_cb.grid(row=0, column=5, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="发票日期:").grid(row=0, column=6, sticky='e', padx=4, pady=4)
        self.inv_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(form, textvariable=self.inv_date_var).grid(row=0, column=7, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="金额(元):").grid(row=1, column=0, sticky='e', padx=4, pady=4)
        self.amount_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.amount_var).grid(row=1, column=1, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="用途:").grid(row=1, column=2, sticky='e', padx=4, pady=4)
        self.purpose_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.purpose_var).grid(row=1, column=3, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="凭证号:").grid(row=1, column=4, sticky='e', padx=4, pady=4)
        self.voucher_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.voucher_var).grid(row=1, column=5, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="银行账户:").grid(row=1, column=6, sticky='e', padx=4, pady=4)
        self.bank_var = tk.StringVar()
        self.bank_cb = ttk.Combobox(form, textvariable=self.bank_var, state='readonly')
        self.bank_cb.grid(row=1, column=7, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="附件:").grid(row=2, column=0, sticky='e', padx=4, pady=4)
        self.img_label_var = tk.StringVar(value="未选择")
        self._full_img_name = ""
        self.img_label = ttk.Label(form, textvariable=self.img_label_var, foreground='gray', anchor='w')
        self.img_label.grid(row=2, column=1, columnspan=5, sticky='we', padx=4, pady=4)
        self._attach_tooltip = None
        self.img_label.bind('<Enter>', lambda e: self._show_img_tooltip(e))
        self.img_label.bind('<Leave>', lambda e: self._hide_img_tooltip())

        ttk.Label(form, text="销售方:").grid(row=3, column=0, sticky='e', padx=4, pady=4)
        self.seller_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.seller_var).grid(row=3, column=1, columnspan=3, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="备注:").grid(row=3, column=4, sticky='e', padx=4, pady=4)
        self.remark_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.remark_var).grid(row=3, column=5, columnspan=3, sticky='we', padx=4, pady=4)

        ttk.Label(form, text="报销状态:").grid(row=2, column=6, sticky='e', padx=4, pady=4)
        self.status_var = tk.StringVar(value='未报销')
        self.status_cb = ttk.Combobox(form, textvariable=self.status_var,
                                       values=['未报销', '已报销'], state='readonly', width=10)
        self.status_cb.grid(row=2, column=7, sticky='we', padx=4, pady=4)

        # 按钮独立一行，左对齐均匀排列
        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=4, column=0, columnspan=8, sticky='we', padx=2, pady=(6, 2))
        for i in range(5):
            btn_frame.columnconfigure(i, weight=1)
        ttk.Button(btn_frame, text="选择图片/PDF", command=self.choose_image).grid(row=0, column=0, sticky='we', padx=2)
        ttk.Button(btn_frame, text="剪贴板粘贴", command=self.paste_from_clipboard, style='Secondary.TButton').grid(row=0, column=1, sticky='we', padx=2)
        self.ocr_btn = ttk.Button(btn_frame, text="识别发票号", command=self.ocr_invoice, style='Secondary.TButton')
        self.ocr_btn.grid(row=0, column=2, sticky='we', padx=2)
        self.add_btn = ttk.Button(btn_frame, text="添加", command=self.add_or_update)
        self.add_btn.grid(row=0, column=3, sticky='we', padx=2)
        self.cancel_btn = ttk.Button(btn_frame, text="取消修改", command=self.cancel_edit, state='disabled')
        self.cancel_btn.grid(row=0, column=4, sticky='we', padx=2)

        # 筛选（两行布局，确保窄窗口也能完整显示所有控件）
        filter_frame = ttk.Frame(self.content)
        filter_frame.pack(fill='x', padx=10, pady=4)
        # 第一行：状态、类型、年份、月份
        filter_row1 = ttk.Frame(filter_frame)
        filter_row1.pack(fill='x', pady=2)
        ttk.Label(filter_row1, text="状态:").pack(side='left', padx=4)
        self.filter_status = tk.StringVar(value='全部')
        ttk.Combobox(filter_row1, textvariable=self.filter_status, values=['全部', '未报销', '已报销', '待完善'],
                     width=8, state='readonly').pack(side='left', padx=4)
        ttk.Label(filter_row1, text="类型:").pack(side='left', padx=4)
        self.filter_type = tk.StringVar(value='全部')
        ttk.Combobox(filter_row1, textvariable=self.filter_type, values=['全部', '发票', '支付记录'],
                     width=8, state='readonly').pack(side='left', padx=4)
        ttk.Label(filter_row1, text="年份:").pack(side='left', padx=4)
        now = datetime.now()
        # 从人员最早入职年份开始，到当前年+5
        try:
            conn = get_db()
            row = conn.execute("SELECT MIN(hire_date) as min_date FROM employees WHERE hire_date IS NOT NULL AND hire_date != ''").fetchone()
            conn.close()
            min_year = 2015
            if row and row['min_date']:
                try:
                    min_year = int(str(row['min_date'])[:4])
                except (ValueError, IndexError):
                    pass
        except Exception:
            min_year = 2015
        self.filter_year = tk.StringVar(value=str(now.year))
        year_options = ['全部'] + [str(y) for y in range(min_year, now.year + 1)]
        ttk.Combobox(filter_row1, textvariable=self.filter_year, values=year_options,
                     width=6, state='readonly').pack(side='left', padx=4)
        ttk.Label(filter_row1, text="月份:").pack(side='left', padx=4)
        self.filter_month = tk.StringVar(value='全部')
        month_options = ['全部'] + [f"{m:02d}" for m in range(1, 13)]
        ttk.Combobox(filter_row1, textvariable=self.filter_month, values=month_options,
                     width=6, state='readonly').pack(side='left', padx=4)
        # 第二行：排版、报销人、发票号搜索、应用筛选按钮
        filter_row2 = ttk.Frame(filter_frame)
        filter_row2.pack(fill='x', pady=2)
        ttk.Label(filter_row2, text="排版:").pack(side='left', padx=4)
        self.filter_printed = tk.StringVar(value='全部')
        ttk.Combobox(filter_row2, textvariable=self.filter_printed, values=['全部', '未排版', '已排版'],
                     width=7, state='readonly').pack(side='left', padx=4)
        ttk.Label(filter_row2, text="报销人:").pack(side='left', padx=4)
        self.filter_person = tk.StringVar(value='全部')
        self.filter_person_cb = ttk.Combobox(filter_row2, textvariable=self.filter_person, width=12, state='readonly')
        self.filter_person_cb.pack(side='left', padx=4)
        ttk.Label(filter_row2, text="发票号:").pack(side='left', padx=4)
        self.filter_invoice_no = tk.StringVar()
        self.filter_invoice_no_entry = ttk.Entry(filter_row2, textvariable=self.filter_invoice_no, width=18)
        self.filter_invoice_no_entry.pack(side='left', padx=4)
        self.filter_invoice_no_entry.bind('<Return>', lambda e: self.refresh())
        ttk.Button(filter_row2, text="应用筛选", command=self.refresh).pack(side='left', padx=4)

        # 表格
        tbl_frame = ttk.Frame(self.content)
        tbl_frame.pack(fill='both', expand=True, padx=10, pady=6)

        cols = ('check', 'id', 'type', 'invoice_number', 'invoice_date', 'reimburser', 'amount', 'purpose', 'voucher', 'printed', 'status')
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show='tree headings', selectmode='browse')
        # #0列作为附件缩略图列（最左）
        self.tree.heading('#0', text='附件')
        self.tree.column('#0', width=60, anchor='center', stretch=False)
        self.checked_ids = set()
        headers = [('check', '选择', 45, 'center'),
                   ('id', '编号', 45, 'center'),
                   ('type', '类型', 55, 'center'),
                   ('invoice_number', '发票号', 110, 'center'),
                   ('invoice_date', '日期', 80, 'center'),
                   ('reimburser', '报销人', 80, 'center'),
                   ('amount', '金额', 80, 'e'),
                   ('purpose', '用途', 150, 'w'),
                   ('voucher', '凭证号', 95, 'center'),
                   ('printed', '排版', 50, 'center'),
                   ('status', '状态', 55, 'center')]
        for cid, text, w, anchor in headers:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor)
        # 缩略图缓存（防止被GC回收）
        self._thumb_cache = []
        self.tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(tbl_frame, orient='vertical', command=self.tree.yview)
        sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind('<Double-1>', self.on_double)
        self.tree.bind('<Button-1>', self.on_click)
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # 底部按钮 - 分两行，避免窗口模式下文字被截断
        btns = ttk.Frame(self.content)
        btns.pack(fill='x', padx=10, pady=6)
        # 第一行：9个操作按钮
        for col in range(9):
            btns.columnconfigure(col, weight=1)
        ttk.Button(btns, text="全选", command=self.select_all).grid(row=0, column=0, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="取消全选", command=self.deselect_all).grid(row=0, column=1, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="范围选择", command=self.range_select).grid(row=0, column=2, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="批量报销/取消", command=self.batch_set_status).grid(row=0, column=3, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="取消排版", command=self.cancel_printed).grid(row=0, column=4, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="批量删除", command=self.batch_delete).grid(row=0, column=5, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="查看图片", command=self.view_image).grid(row=0, column=6, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="删除", command=self.delete).grid(row=0, column=7, sticky='we', padx=2, pady=2)
        # 第9列：刷新
        ttk.Button(btns, text="刷新", command=self.refresh).grid(row=0, column=8, sticky='we', padx=2, pady=2)
        # 第二行：打印排版 + 导入导出 + 计数
        for col in range(9):
            btns.columnconfigure(col, weight=1)
        ttk.Button(btns, text="打印排版", command=self.print_layout, style='Accent.TButton').grid(row=1, column=0, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="导入Excel", command=self.import_excel).grid(row=1, column=1, sticky='we', padx=2, pady=2)
        ttk.Button(btns, text="导出Excel", command=self.export_excel).grid(row=1, column=2, sticky='we', padx=2, pady=2)
        ttk.Label(btns, text="勾选后一键指定报销人，不改变报销状态",
                  foreground='gray').grid(row=1, column=3, columnspan=3, sticky='we', padx=2, pady=2)
        self.checked_count_label = ttk.Label(btns, text="双击修改 · 已选0项", foreground='#3B82F6',
                                             anchor='e')
        self.checked_count_label.grid(row=1, column=6, columnspan=2, sticky='we', padx=2, pady=2)

        # 批量指定报销人按钮
        btns2 = ttk.Frame(self.content)
        btns2.pack(fill='x', padx=10, pady=(0, 6))
        ttk.Button(btns2, text="批量设置报销人", command=self.batch_set_reimburser, style='Accent.TButton')\
            .pack(side='left', padx=2)

        self._load_persons()
        self._load_bank_accounts()
        # 初始布局完成后调整一次预览列宽度，确保左侧内容不被挤压（延迟+重试等待窗口布局完成）
        self._preview_fit_retries = 0
        self.after(300, self._initial_preview_fit)

    def _initial_preview_fit(self):
        """初始布局后根据tab宽度收窄预览列（最多重试5次，等待窗口布局完成）"""
        try:
            w = self.winfo_width()
            if w < 100 and self._preview_fit_retries < 5:
                self._preview_fit_retries += 1
                self.after(200, self._initial_preview_fit)
                return
            if w > 100:
                fixed = self.min_content_w + 12 + 16 + 12
                max_preview = max(w - fixed, 280)
                if self.preview_panel.winfo_width() > max_preview:
                    self.preview_panel.config(width=max_preview)
                    self._on_preview_resize()
        except Exception:
            pass

    def _load_persons(self):
        conn = get_db()
        all_rows = conn.execute("SELECT name, resigned FROM employees ORDER BY name").fetchall()
        conn.close()
        # 新增/修改下拉只显示未离职人员
        active_names = [r['name'] for r in all_rows if not r['resigned']]
        self.reimburser_cb['values'] = active_names
        # 筛选下拉包含所有人员（含已离职，方便历史查询）
        all_names = [r['name'] for r in all_rows]
        self.filter_person_cb['values'] = ['全部'] + all_names

    def _load_bank_accounts(self):
        conn = get_db()
        rows = conn.execute("SELECT id, name FROM bank_accounts ORDER BY name").fetchall()
        conn.close()
        self.bank_cb['values'] = [r['name'] for r in rows]
        self._bank_map = {r['name']: r['id'] for r in rows}

    def refresh(self):
        self._load_persons()
        self._load_bank_accounts()
        # 配置待完善行样式
        if is_dark_theme():
            self.tree.tag_configure('pending', background='#3A2A00', foreground='#FFD60A')
            self.tree.tag_configure('checked', background='#2563EB', foreground='#FFFFFF')
        else:
            self.tree.tag_configure('pending', background='#FFF4CC', foreground='#8B6914')
            self.tree.tag_configure('checked', background='#93C5FD', foreground='#000000')
        for i in self.tree.get_children():
            self.tree.delete(i)
        conn = get_db()
        sql = """SELECT i.*, e.name as rname, e.resigned as resigned FROM invoices i
                 LEFT JOIN employees e ON i.reimburser_id=e.id WHERE 1=1"""
        params = []
        if self.filter_status.get() == '未报销':
            sql += " AND i.status=0 AND COALESCE(i.batch_pending,0)=0"
        elif self.filter_status.get() == '已报销':
            sql += " AND i.status=1 AND COALESCE(i.batch_pending,0)=0"
        elif self.filter_status.get() == '待完善':
            sql += " AND COALESCE(i.batch_pending,0)=1"
        if self.filter_type.get() and self.filter_type.get() != '全部':
            sql += " AND i.type=?"
            params.append(self.filter_type.get())
        if self.filter_year.get() and self.filter_year.get() != '全部':
            sql += " AND strftime('%Y', i.invoice_date)=?"
            params.append(self.filter_year.get())
        if self.filter_month.get() and self.filter_month.get() != '全部':
            sql += " AND strftime('%m', i.invoice_date)=?"
            params.append(self.filter_month.get())
        if self.filter_person.get() and self.filter_person.get() != '全部':
            sql += " AND e.name=?"
            params.append(self.filter_person.get())
        if self.filter_printed.get() == '未排版':
            sql += " AND COALESCE(i.printed,0)=0"
        elif self.filter_printed.get() == '已排版':
            sql += " AND i.printed=1"
        if hasattr(self, 'filter_invoice_no') and self.filter_invoice_no.get().strip():
            sql += " AND i.invoice_number LIKE ?"
            params.append(f"%{self.filter_invoice_no.get().strip()}%")
        sql += " ORDER BY i.invoice_date ASC, i.id ASC"
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        self._thumb_cache.clear()
        for seq, r in enumerate(rows, start=1):
            is_pending = (r['batch_pending'] == 1) if 'batch_pending' in r.keys() else False
            if is_pending:
                status = '待完善'
            else:
                status = '已报销' if r['status'] == 1 else '未报销'
            rname = r['rname'] or '(已删除)'
            if r['resigned']:
                rname += ' (已离职)'
            inv_type = r['type'] or '发票'
            printed_mark = '✓' if (r['printed'] == 1 if 'printed' in r.keys() else False) else ''
            # 生成附件缩略图（#0列）
            thumb_img = ''
            if r['image_path']:
                pil_img = get_invoice_image(r['image_path'])
                if pil_img:
                    try:
                        pil_img.thumbnail((45, 45), Image.LANCZOS)
                        photo = ImageTk.PhotoImage(pil_img)
                        self._thumb_cache.append(photo)
                        thumb_img = photo
                    except Exception:
                        pass
            tags = []
            if is_pending:
                tags.append('pending')
            is_checked = r['id'] in self.checked_ids
            if is_checked:
                tags.append('checked')
            check_mark = '✓' if is_checked else '☐'
            self.tree.insert('', 'end', iid=str(r['id']), image=thumb_img, values=(
                check_mark, seq, inv_type, r['invoice_number'] or '', r['invoice_date'] or '',
                rname, fmt_money(r['amount']),
                r['purpose'] or '', r['voucher_number'] or '', printed_mark, status), tags=tuple(tags))
        if hasattr(self, 'checked_count_label'):
            self.checked_count_label.config(text=f"双击修改 · 已选{len(self.checked_ids)}项")
        # 清除行选中状态，防止selected样式覆盖tag背景色
        try:
            for s in self.tree.selection():
                self.tree.selection_remove(s)
        except Exception:
            pass

    def on_select(self, event=None):
        sel = self.tree.selection()
        if sel:
            try:
                self.selected_id = int(sel[0])
            except (ValueError, TypeError):
                self.selected_id = None

    def on_click(self, event):
        """单击选择列（#1）切换勾选状态"""
        region = self.tree.identify('region', event.x, event.y)
        if region not in ('cell', 'tree'):
            return
        col = self.tree.identify_column(event.x)
        if col == '#1':  # #1是选择列（#0是附件缩略图列）
            row_id = self.tree.identify_row(event.y)
            if row_id:
                try:
                    inv_id = int(row_id)
                except (ValueError, TypeError):
                    inv_id = None
                if inv_id:
                    seq = self.tree.item(row_id, 'values')[1]
                    if inv_id in self.checked_ids:
                        self.checked_ids.remove(inv_id)
                        self.app.set_status(f"已取消勾选 编号{seq}")
                    else:
                        self.checked_ids.add(inv_id)
                        self.app.set_status(f"已勾选 编号{seq}")
                    self.refresh()
                return 'break'  # 阻止默认选中行为

    def on_double(self, event=None):
        """双击：附件列(#0)查看图片，其他列修改记录"""
        sel = self.tree.selection()
        if not sel:
            return
        try:
            self.selected_id = int(sel[0])
        except (ValueError, TypeError):
            self.selected_id = None
        col = self.tree.identify_column(event.x) if event else ''
        if col == '#0':
            # 附件列：查看图片
            self.view_image()
            return
        conn = get_db()
        row = conn.execute("""SELECT i.*, e.name as rname, b.name as bname FROM invoices i
                              LEFT JOIN employees e ON i.reimburser_id=e.id
                              LEFT JOIN bank_accounts b ON i.bank_account_id=b.id
                              WHERE i.id=?""",
                           (self.selected_id,)).fetchone()
        conn.close()
        if not row:
            return
        self.inv_no_var.set(row['invoice_number'] or '')
        self.inv_type_var.set(row['type'] or '发票')
        self.inv_date_var.set(row['invoice_date'] or '')
        self.reimburser_var.set(row['rname'] or '')
        self.amount_var.set(str(row['amount'] or 0))
        self.purpose_var.set(row['purpose'] or '')
        self.voucher_var.set(row['voucher_number'] or '')
        self.bank_var.set(row['bname'] or '')
        self.seller_var.set(row['seller'] or '')
        self.remark_var.set(row['remark'] or '')
        self.status_var.set('已报销' if row['status'] == 1 else '未报销')
        if row['image_path']:
            self.img_label_var.set(f"已有图片: {row['image_path']}")
        else:
            self.img_label_var.set("未选择")
        self.image_file = None  # 修改时默认不替换图片，除非用户重新选择
        self.add_btn.config(text="保存修改")
        self.cancel_btn.config(state='normal')
        # 右侧显示发票放大图，便于人工核对
        self._show_preview(row['image_path'] if row['image_path'] else None)

    def _resolve_attachment(self, img_path):
        """把发票附件路径解析为绝对路径（数据库存的是INVOICE_DIR下的文件名）"""
        if not img_path:
            return None
        if os.path.isabs(img_path):
            return img_path if os.path.exists(img_path) else None
        p = os.path.join(INVOICE_DIR, img_path)
        return p if os.path.exists(p) else None

    def _show_preview(self, img_path):
        """在右侧预览列显示发票图片（默认适应窗口，可缩放滚动）"""
        try:
            full = self._resolve_attachment(img_path)
            self._preview_path = full
            self._preview_orig_img = None
            # 切换发票时重置OCR状态
            self._ocr_results = []
            self._ocr_overlay_on = False
            self._ocr_selected = set()
            self._ocr_drag = None
            self._ocr_sel_rect = None
            if hasattr(self, '_ocr_toggle_btn'):
                self._ocr_toggle_btn.config(text="OCR", state='normal')
            if hasattr(self, 'ocr_text'):
                self.ocr_text.delete('1.0', 'end')
            if self._ocr_panel_visible:
                self._hide_ocr_panel()
            if not full:
                self._preview_photo = None
                self._draw_preview_placeholder("无发票图片")
                return
            img = None
            if str(full).lower().endswith('.pdf'):
                img = pdf_to_image(full, dpi=150)
            else:
                img = Image.open(full)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
            if img is None:
                self._draw_preview_placeholder("无法加载图片")
                return
            self._preview_orig_img = img
            self._preview_fit()
        except Exception as e:
            _log_ocr_error(f"预览发票失败: {e}")
            self._draw_preview_placeholder("预览失败")

    def _draw_preview_placeholder(self, text):
        """在预览Canvas上绘制占位提示文字"""
        try:
            self.preview_canvas.delete("all")
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            if cw <= 1:
                cw = 350
            if ch <= 1:
                ch = 400
            self.preview_canvas.create_text(cw // 2, ch // 2,
                                            text=text + "\n\n双击列表中的发票\n此处显示发票大图",
                                            fill='gray', justify='center', font=('', 10))
            self.preview_canvas.configure(scrollregion=(0, 0, cw, ch))
            self._zoom_label.config(text="")
        except Exception:
            pass

    def _render_preview(self):
        """按当前缩放比例渲染图片到预览Canvas，居中显示，放大后可滚动"""
        if self._preview_orig_img is None:
            return
        try:
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            if cw <= 10 or ch <= 10:
                self.after(80, self._render_preview)
                return
            ow, oh = self._preview_orig_img.size
            iw = max(1, int(ow * self._preview_scale))
            ih = max(1, int(oh * self._preview_scale))
            # 超大图限制渲染尺寸防止卡顿（超过4000px时降采样显示，但滚动范围仍按原图比例）
            disp_iw, disp_ih = iw, ih
            max_render = 4000
            if max(disp_iw, disp_ih) > max_render:
                r = max_render / max(disp_iw, disp_ih)
                disp_iw, disp_ih = int(disp_iw * r), int(disp_ih * r)
            img = self._preview_orig_img.resize((disp_iw, disp_ih), Image.LANCZOS)
            self._preview_photo = ImageTk.PhotoImage(img)
            self.preview_canvas.delete("all")
            x = max(0, (cw - iw) // 2)
            y = max(0, (ch - ih) // 2)
            self.preview_canvas.create_image(x, y, anchor='nw', image=self._preview_photo)
            self.preview_canvas.configure(scrollregion=(0, 0, max(iw, cw), max(ih, ch)))
            self._zoom_label.config(text=f"{int(self._preview_scale * 100)}%")
            # 叠加OCR文字框
            self._draw_ocr_overlay()
        except Exception as e:
            _log_ocr_error(f"渲染预览失败: {e}")

    def _preview_fit(self):
        """缩放图片适应当前Canvas窗口"""
        if self._preview_orig_img is None:
            return
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        if cw <= 10 or ch <= 10:
            self.after(100, self._preview_fit)
            return
        ow, oh = self._preview_orig_img.size
        self._preview_scale = min((cw - 20) / ow, (ch - 20) / oh)
        self._render_preview()

    def _preview_zoom_in(self):
        if self._preview_orig_img is None:
            return
        self._preview_scale = min(self._preview_scale * 1.25, 8.0)
        self._render_preview()

    def _preview_zoom_out(self):
        if self._preview_orig_img is None:
            return
        self._preview_scale = max(self._preview_scale / 1.25, 0.05)
        self._render_preview()

    def _preview_actual(self):
        """1:1原始大小"""
        if self._preview_orig_img is None:
            return
        self._preview_scale = 1.0
        self._render_preview()

    def _on_preview_wheel_zoom(self, event):
        """Ctrl+滚轮缩放"""
        if self._preview_orig_img is None:
            return
        if event.delta > 0:
            self._preview_zoom_in()
        else:
            self._preview_zoom_out()

    def _preview_key_scroll(self, dx, dy):
        """键盘滚动预览图片"""
        try:
            self.preview_canvas.xview_scroll(int(dx), 'units')
            self.preview_canvas.yview_scroll(int(dy), 'units')
        except Exception:
            pass

    def _on_preview_resize(self, event=None):
        """Canvas尺寸变化时重新渲染（保持当前缩放比例，居中）"""
        if self._preview_orig_img is not None:
            self._render_preview()
        else:
            self._draw_preview_placeholder("双击列表中的发票")

    # ========== 预览OCR文字识别与复制 ==========

    def _toggle_preview_ocr(self):
        """切换OCR识别：首次点击运行OCR并显示文字面板+框选，再次点击隐藏"""
        if self._preview_orig_img is None:
            messagebox.showinfo("提示", "请先在列表中选择一张发票")
            return
        if self._ocr_panel_visible:
            # 已显示 → 隐藏面板和框选
            self._hide_ocr_panel()
            self._ocr_overlay_on = False
            self._ocr_selected = set()
            self._ocr_toggle_btn.config(text="OCR")
            self.preview_canvas.config(cursor='')
            self._render_preview()
            return
        # 显示面板
        self._show_ocr_panel()
        if self._ocr_results:
            # 已有缓存结果，直接显示
            self._populate_ocr_text()
            self._ocr_overlay_on = True
            self._ocr_toggle_btn.config(text="OCR✓")
            self._render_preview()
        else:
            # 需要运行OCR
            self._ocr_toggle_btn.config(state='disabled', text="识别中")
            self.ocr_text.delete('1.0', 'end')
            self.ocr_text.insert('1.0', "正在识别，请稍候...")
            self.app.set_status("正在OCR识别发票文字...")
            import threading
            def worker():
                results = self._run_preview_ocr()
                self.after(0, lambda: self._on_preview_ocr_done(results))
            threading.Thread(target=worker, daemon=True).start()

    def _run_preview_ocr(self):
        """对当前预览图片运行OCR，返回 [(box, text, conf), ...]"""
        engine = get_ocr_engine()
        if engine is None:
            return []
        try:
            import tempfile
            img = self._preview_orig_img
            if img is None:
                return []
            # 压缩大图
            max_side = 3000
            w, h = img.size
            if max(w, h) > max_side:
                ratio = max_side / max(w, h)
                img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
            tmp_path = os.path.join(tempfile.gettempdir(),
                                    f"prev_ocr_{os.getpid()}_{int(time.time()*1000)}.jpg")
            img.save(tmp_path, 'JPEG', quality=95)
            result, elapse = engine(tmp_path)
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
            if result:
                return [(line[0], line[1], line[2]) for line in result]
            return []
        except Exception as e:
            _log_ocr_error(f"预览OCR失败: {e}")
            return []

    def _on_preview_ocr_done(self, results):
        """OCR完成回调"""
        self._ocr_results = results
        self._ocr_toggle_btn.config(state='normal', text="OCR✓")
        self._ocr_overlay_on = True
        if results:
            self.app.set_status(f"OCR完成，共{len(results)}行文字。在图片上拖拽框选要复制的内容，Ctrl+框选可追加")
            self._populate_ocr_text()
            self._render_preview()
        else:
            self.app.set_status("OCR未识别到文字")
            self.ocr_text.delete('1.0', 'end')
            self.ocr_text.insert('1.0', "未识别到文字")

    def _show_ocr_panel(self):
        """显示OCR文字面板"""
        self._ocr_panel_visible = True
        self.ocr_panel.pack(fill='x', pady=(4, 0))
        # 面板显示后重新适应图片（Canvas高度变小）
        self.after(100, self._render_preview)

    def _hide_ocr_panel(self):
        """隐藏OCR文字面板"""
        self._ocr_panel_visible = False
        self.ocr_panel.pack_forget()
        self.after(100, self._render_preview)

    def _populate_ocr_text(self):
        """把OCR结果填入文字面板，每行可点击复制"""
        self.ocr_text.delete('1.0', 'end')
        self._ocr_line_ranges = []
        for i, (box, text, conf) in enumerate(self._ocr_results):
            start = self.ocr_text.index('end-1c')
            self.ocr_text.insert('end', text + '\n')
            end = self.ocr_text.index('end-1c')
            self._ocr_line_ranges.append((start, end, text))
        self.ocr_text.config(state='normal')

    def _on_ocr_text_click(self, event):
        """点击文字面板某行，复制该行文字到剪贴板"""
        try:
            index = self.ocr_text.index(f"@{event.x},{event.y}")
            for start, end, text in self._ocr_line_ranges:
                if self.ocr_text.compare(index, '>=', start) and self.ocr_text.compare(index, '<=', end):
                    self.clipboard_clear()
                    self.clipboard_append(text)
                    self.app.set_status(f"已复制: {text}")
                    # 高亮该行
                    self.ocr_text.tag_remove('ocr_sel', '1.0', 'end')
                    self.ocr_text.tag_add('ocr_sel', start, end)
                    self.ocr_text.tag_config('ocr_sel', background='#3B82F6', foreground='#ffffff')
                    self.after(1500, lambda: self.ocr_text.tag_remove('ocr_sel', '1.0', 'end'))
                    return
        except Exception:
            pass

    def _copy_all_ocr_text(self):
        """复制全部OCR文字"""
        if not self._ocr_results:
            return
        all_text = '\n'.join(t for _, t, _ in self._ocr_results)
        self.clipboard_clear()
        self.clipboard_append(all_text)
        self.app.set_status(f"已复制全部{len(self._ocr_results)}行文字")

    def _on_preview_click_ocr(self, event):
        """点击预览Canvas上的OCR框，复制对应文字"""
        if not self._ocr_overlay_on or not self._ocr_results:
            return
        try:
            # Canvas坐标 → 图片坐标
            cx = self.preview_canvas.canvasx(event.x)
            cy = self.preview_canvas.canvasy(event.y)
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            ow, oh = self._preview_orig_img.size
            iw = max(1, int(ow * self._preview_scale))
            ih = max(1, int(oh * self._preview_scale))
            ox = max(0, (cw - iw) // 2)
            oy = max(0, (ch - ih) // 2)
            img_x = (cx - ox) / self._preview_scale
            img_y = (cy - oy) / self._preview_scale
            if img_x < 0 or img_y < 0 or img_x > ow or img_y > oh:
                return
            # 查找点击了哪个文字框
            for box, text, conf in self._ocr_results:
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                if min(xs) <= img_x <= max(xs) and min(ys) <= img_y <= max(ys):
                    self.clipboard_clear()
                    self.clipboard_append(text)
                    self.app.set_status(f"已复制: {text}")
                    return
        except Exception:
            pass

    def _draw_ocr_overlay(self):
        """在预览Canvas上绘制OCR文字框（选中的绿色高亮，未选中的蓝色）"""
        if not self._ocr_overlay_on or not self._ocr_results:
            return
        try:
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            ow, oh = self._preview_orig_img.size
            iw = max(1, int(ow * self._preview_scale))
            ih = max(1, int(oh * self._preview_scale))
            ox = max(0, (cw - iw) // 2)
            oy = max(0, (ch - ih) // 2)
            self._ocr_box_items = []
            self._ocr_img_offset = (ox, oy, ow, oh)
            for idx, (box, text, conf) in enumerate(self._ocr_results):
                pts = [(ox + p[0] * self._preview_scale, oy + p[1] * self._preview_scale) for p in box]
                flat = [c for pt in pts for c in pt]
                if idx in self._ocr_selected:
                    color = '#22C55E'  # 绿色选中
                    fill = '#22C55E'
                else:
                    color = '#3B82F6'  # 蓝色未选中
                    fill = '#3B82F6'
                item = self.preview_canvas.create_polygon(flat, outline=color, width=2,
                                                          fill=fill, stipple='gray12',
                                                          tags=('ocrbox',))
                self._ocr_box_items.append((item, idx))
        except Exception:
            pass

    def _canvas_to_img(self, cx, cy):
        """Canvas坐标转原图坐标"""
        ox, oy, ow, oh = self._ocr_img_offset
        ix = (cx - ox) / self._preview_scale
        iy = (cy - oy) / self._preview_scale
        return ix, iy

    def _box_at_img_point(self, ix, iy):
        """查找图片坐标点落在哪个OCR框内，返回索引或None"""
        for idx, (box, text, conf) in enumerate(self._ocr_results):
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            if min(xs) <= ix <= max(xs) and min(ys) <= iy <= max(ys):
                return idx
        return None

    def _boxes_in_rect(self, x1, y1, x2, y2):
        """查找与矩形相交的所有OCR框索引"""
        rx1, rx2 = min(x1, x2), max(x1, x2)
        ry1, ry2 = min(y1, y2), max(y1, y2)
        result = []
        for idx, (box, text, conf) in enumerate(self._ocr_results):
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            bx1, bx2 = min(xs), max(xs)
            by1, by2 = min(ys), max(ys)
            if bx2 >= rx1 and bx1 <= rx2 and by2 >= ry1 and by1 <= ry2:
                result.append(idx)
        return result

    def _on_preview_click_ocr(self, event):
        """鼠标按下：开始拖拽框选（OCR模式下）"""
        if not self._ocr_overlay_on or not self._ocr_results:
            return
        try:
            cx = self.preview_canvas.canvasx(event.x)
            cy = self.preview_canvas.canvasy(event.y)
            self._ocr_drag = {'start_x': cx, 'start_y': cy, 'cur_x': cx, 'cur_y': cy}
            # 绘制选择矩形
            if self._ocr_sel_rect:
                self.preview_canvas.delete(self._ocr_sel_rect)
            self._ocr_sel_rect = self.preview_canvas.create_rectangle(
                cx, cy, cx, cy, outline='#F59E0B', width=2, dash=(4, 3), tags=('ocrsel',))
        except Exception:
            pass

    def _on_preview_drag_ocr(self, event):
        """鼠标拖拽：更新选择矩形"""
        if not self._ocr_drag or not self._ocr_overlay_on:
            return
        try:
            cx = self.preview_canvas.canvasx(event.x)
            cy = self.preview_canvas.canvasy(event.y)
            self._ocr_drag['cur_x'] = cx
            self._ocr_drag['cur_y'] = cy
            if self._ocr_sel_rect:
                self.preview_canvas.coords(
                    self._ocr_sel_rect,
                    self._ocr_drag['start_x'], self._ocr_drag['start_y'], cx, cy)
        except Exception:
            pass

    def _on_preview_release_ocr(self, event):
        """鼠标释放：根据选择范围选中OCR框并复制文字"""
        if not self._ocr_drag or not self._ocr_overlay_on:
            return
        try:
            sx, sy = self._ocr_drag['start_x'], self._ocr_drag['start_y']
            cx = self.preview_canvas.canvasx(event.x)
            cy = self.preview_canvas.canvasy(event.y)
            dx = abs(cx - sx)
            dy = abs(cy - sy)
            # 清除选择矩形
            if self._ocr_sel_rect:
                self.preview_canvas.delete(self._ocr_sel_rect)
                self._ocr_sel_rect = None
            self._ocr_drag = None

            ctrl = (event.state & 0x0004) != 0  # Ctrl键
            if dx < 5 and dy < 5:
                # 单击：弹出文字选择器，可选择部分文字复制
                ix, iy = self._canvas_to_img(cx, cy)
                idx = self._box_at_img_point(ix, iy)
                if idx is not None:
                    _, text, _ = self._ocr_results[idx]
                    self._show_ocr_text_picker(text)
                elif not ctrl:
                    self._ocr_selected = set()
                    self._render_preview()
            else:
                # 拖拽框选：选中矩形内所有框并复制
                ix1, iy1 = self._canvas_to_img(sx, sy)
                ix2, iy2 = self._canvas_to_img(cx, cy)
                idxs = self._boxes_in_rect(ix1, iy1, ix2, iy2)
                if ctrl:
                    self._ocr_selected.update(idxs)
                else:
                    self._ocr_selected = set(idxs)
                # 重新绘制框
                self._render_preview()
                # 复制选中文字
                self._copy_selected_ocr()
        except Exception:
            pass

    def _copy_selected_ocr(self):
        """复制选中的OCR文字到剪贴板（按从上到下、从左到右排序）"""
        if not self._ocr_selected:
            self.app.set_status("未选中任何文字")
            return
        # 按框的顶部Y坐标排序，Y相近时按X排序
        selected = []
        for idx in self._ocr_selected:
            box, text, conf = self._ocr_results[idx]
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            selected.append((min(ys), min(xs), text))
        selected.sort(key=lambda t: (round(t[0] / 15), t[1]))
        text = '\n'.join(t[2] for t in selected)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.app.set_status(f"已复制{len(self._ocr_selected)}行（点击单个框可选择部分文字，Ctrl+框选追加）")

    def _show_ocr_text_picker(self, text):
        """弹出小窗口，让用户选择文字的一部分复制"""
        try:
            top = tk.Toplevel(self)
            top.title("选择文字")
            top.geometry("320x120")
            top.transient(self)
            # 在鼠标位置附近弹出
            try:
                mx = self.winfo_pointerx()
                my = self.winfo_pointery()
                top.update_idletasks()
                w = top.winfo_width()
                h = top.winfo_height()
                if w < 10:
                    w, h = 320, 120
                px = mx - w // 2
                py = my - h // 2
                # 确保不超出屏幕
                sw = top.winfo_screenwidth()
                sh = top.winfo_screenheight()
                px = max(0, min(px, sw - w))
                py = max(0, min(py, sh - h))
                top.geometry(f"+{px}+{py}")
            except Exception:
                pass
            top.grab_set()
            ttk.Label(top, text="选择要复制的文字（Ctrl+C复制）：", font=('', 9)).pack(anchor='w', padx=8, pady=(6, 2))
            entry = ttk.Entry(top, font=('Consolas', 11))
            entry.pack(fill='x', padx=8, pady=4)
            entry.insert(0, text)
            entry.select_range(0, 'end')
            entry.focus_set()
            def copy_and_close():
                sel = entry.selection_get() if entry.selection_present() else entry.get()
                self.clipboard_clear()
                self.clipboard_append(sel)
                self.app.set_status(f"已复制: {sel}")
                top.destroy()
            ttk.Button(top, text="复制选中内容", command=copy_and_close).pack(pady=4)
            entry.bind('<Return>', lambda e: copy_and_close())
            entry.bind('<Escape>', lambda e: top.destroy())
            top.bind('<Escape>', lambda e: top.destroy())
        except Exception:
            pass

    def _on_preview_ocr_hover(self, event):
        """鼠标悬停在OCR框上时显示手型光标"""
        if not self._ocr_overlay_on or not self._ocr_results:
            return
        try:
            cx = self.preview_canvas.canvasx(event.x)
            cy = self.preview_canvas.canvasy(event.y)
            ix, iy = self._canvas_to_img(cx, cy)
            if self._box_at_img_point(ix, iy) is not None:
                self.preview_canvas.config(cursor='crosshair')
            else:
                self.preview_canvas.config(cursor='tcross' if self._ocr_overlay_on else '')
        except Exception:
            pass

    def _on_preview_ocr_leave(self, event):
        """鼠标离开预览Canvas时恢复光标"""
        self.preview_canvas.config(cursor='')

    def _on_sash_press(self, event):
        self._sash_drag = {'x': event.x_root, 'w': self.preview_panel.winfo_width(), 'active': True}
        self.sash.config(bg=self._sash_bg_hover)

    def _draw_sash_grip(self, event=None):
        """在分隔条中间绘制握把点，提示可拖拽"""
        try:
            self.sash.delete("grip")
            w = self.sash.winfo_width()
            h = self.sash.winfo_height()
            if w < 4 or h < 4:
                return
            cx = w // 2
            # 中间一排圆点
            count = max(3, h // 40)
            spacing = h // (count + 1)
            for i in range(1, count + 1):
                y = spacing * i
                self.sash.create_oval(cx - 1, y - 1, cx + 1, y + 1,
                                      fill=self._sash_grip, outline='', tags="grip")
        except Exception:
            pass

    def _on_sash_drag(self, event):
        if not self._sash_drag.get('active'):
            return
        dx = event.x_root - self._sash_drag['x']
        new_w = self._sash_drag['w'] - dx  # 鼠标右移→预览列变窄
        try:
            total = self.winfo_width()
            max_w = max(total - self.min_content_w - 30, 280)
        except Exception:
            max_w = 1200
        new_w = max(260, min(new_w, max_w))
        self.preview_panel.config(width=new_w)
        self._on_preview_resize()

    def _on_sash_release(self, event):
        self._sash_drag['active'] = False
        self.sash.config(bg=self._sash_bg_normal)

    def _on_sash_reset(self, event):
        """双击分隔条重置预览列宽度"""
        self.preview_panel.config(width=420)
        self._on_preview_resize()

    def _on_tab_resize(self, event):
        """窗口/标签页大小变化时，自动收窄预览列以保证左侧内容完整显示"""
        if event.widget is not self:
            return
        if event.width == self._last_tab_w or event.width < 100:
            return
        self._last_tab_w = event.width
        # 左侧min_content_w + sash12 + 垂直滚动条16 + 边距
        fixed = self.min_content_w + 12 + 16 + 12
        max_preview = event.width - fixed
        if max_preview < 280:
            max_preview = 280
        cur = self.preview_panel.winfo_width()
        if cur > max_preview:
            self.preview_panel.config(width=max_preview)
            self._on_preview_resize()

    def _preview_view_big(self):
        """双击预览图，打开放大查看窗口（使用ImageViewer，支持ctrl+滚轮缩放、旋转、滚动）"""
        try:
            full = self._resolve_attachment(self._preview_path)
            if not full:
                messagebox.showinfo("提示", "没有可查看的发票附件")
                return
            # 构建简化的发票信息
            invoice_info = {
                'invoice_number': self.inv_no_var.get() or '无号',
                'invoice_date': self.inv_date_var.get() or '',
                'amount': self.amount_var.get() or 0,
                'purpose': self.purpose_var.get() or '',
            }
            # 使用ImageViewer类打开，支持ctrl+滚轮缩放、旋转、滚动等完整功能
            viewer = ImageViewer(self.app, str(full), invoice_info, invoice_id=self.selected_id)
        except Exception as e:
            _log_ocr_error(f"放大查看发票失败: {e}")
            messagebox.showerror("错误", f"打开发票大图失败: {e}")

    def _clear_preview(self):
        """清空右侧预览（面板始终显示，只清空图片并恢复提示文字）"""
        self._preview_photo = None
        self._preview_path = None
        self._preview_orig_img = None
        self._preview_scale = 1.0
        self._draw_preview_placeholder("双击列表中的发票")

    def cancel_edit(self):
        self.selected_id = None
        self.inv_no_var.set('')
        self.inv_type_var.set('发票')
        self.inv_date_var.set(datetime.now().strftime('%Y-%m-%d'))
        self.reimburser_var.set('')
        self.amount_var.set('')
        self.purpose_var.set('')
        self.voucher_var.set('')
        self.bank_var.set('')
        self.seller_var.set('')
        self.remark_var.set('')
        self.status_var.set('未报销')
        self.image_file = None
        self._full_img_name = ""
        self.img_label_var.set('未选择')
        self.add_btn.config(text="添加")
        self.cancel_btn.config(state='disabled')
        self._clear_preview()

    def paste_from_clipboard(self):
        """从剪贴板粘贴图片作为附件"""
        try:
            from PIL import ImageGrab
            img = ImageGrab.grabclipboard()
            if img is None:
                self.app.set_status("剪贴板中没有图片")
                return
            if isinstance(img, list):
                # 剪贴板可能是文件路径列表
                if img and os.path.isfile(img[0]):
                    self.image_file = img[0]
                    full_name = os.path.basename(img[0])
                    self._full_img_name = full_name
                    if len(full_name) > 20:
                        name, ext = os.path.splitext(full_name)
                        display = name[:15] + "..." + ext
                    else:
                        display = full_name
                    self.img_label_var.set(display)
                    self._reset_ocr_fields()
                    self.app.set_status("已从剪贴板粘贴文件")
                    return
                self.app.set_status("剪贴板内容不是图片")
                return
            # 是PIL Image，保存到invoices目录
            if img.mode != 'RGB':
                img = img.convert('RGB')
            fname = f"clipboard_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
            dst = os.path.join(INVOICE_DIR, fname)
            img.save(dst, 'JPEG', quality=92)
            self.image_file = dst
            self._full_img_name = fname
            self.img_label_var.set(fname)
            self._reset_ocr_fields()
            self.app.set_status("已从剪贴板粘贴图片")
        except Exception as e:
            _log_ocr_error(f"剪贴板粘贴失败: {e}")
            self.app.set_status(f"剪贴板粘贴失败: {e}")

    def _reset_ocr_fields(self):
        """重新选择附件后清空所有已识别字段"""
        self.inv_no_var.set('')
        self.inv_date_var.set(datetime.now().strftime('%Y-%m-%d'))
        self.amount_var.set('')
        self.seller_var.set('')
        self.remark_var.set('')
        self.inv_type_var.set('发票')
        self.app.set_status("已选择新附件，请点击「识别发票号」自动识别信息")

    def choose_image(self):
        path = filedialog.askopenfilename(title="选择发票图片或PDF",
                                          filetypes=[("发票文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.pdf"),
                                                     ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"),
                                                     ("PDF文件", "*.pdf"),
                                                     ("所有文件", "*.*")])
        if path:
            self.image_file = path
            full_name = os.path.basename(path)
            self._full_img_name = full_name
            # 截断显示：超过18个字符显示前15个+...+扩展名
            if len(full_name) > 20:
                name, ext = os.path.splitext(full_name)
                display = name[:15] + "..." + ext
            else:
                display = full_name
            self.img_label_var.set(display)
            # 重新上传附件后，清空所有已识别字段
            self._reset_ocr_fields()

    def _show_img_tooltip(self, event):
        """显示附件完整文件名tooltip"""
        if not self._full_img_name:
            return
        x = event.widget.winfo_rootx() + 20
        y = event.widget.winfo_rooty() + event.widget.winfo_height() + 5
        self._attach_tooltip = tk.Toplevel()
        self._attach_tooltip.wm_overrideredirect(True)
        self._attach_tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self._attach_tooltip, text=self._full_img_name,
                        background="#FFFFE0", foreground="#000000",
                        borderwidth=1, relief='solid', font=('微软雅黑', 9),
                        wraplength=400, justify='left')
        label.pack(padx=5, pady=3)

    def _hide_img_tooltip(self):
        """隐藏tooltip"""
        if self._attach_tooltip:
            self._attach_tooltip.destroy()
            self._attach_tooltip = None

    def ocr_invoice(self):
        """手动识别发票号码"""
        ocr_path = self.image_file
        if not ocr_path and self.selected_id:
            # 编辑模式：用数据库中已有的附件
            conn = get_db()
            row = conn.execute("SELECT image_path FROM invoices WHERE id=?", (self.selected_id,)).fetchone()
            conn.close()
            if row and row['image_path']:
                ocr_path = os.path.join(INVOICE_DIR, row['image_path'])
                if not os.path.exists(ocr_path):
                    ocr_path = None
        if not ocr_path:
            messagebox.showwarning("提示", "请先选择发票附件")
            return
        self.ocr_btn.config(state='disabled', text="识别中...")
        self.app.set_status("正在识别发票号码...")
        import threading
        def ocr_worker():
            inv_no, texts = ocr_invoice_number(ocr_path)
            self.after(0, lambda: self._on_ocr_done(inv_no, texts))
        threading.Thread(target=ocr_worker, daemon=True).start()

    def _on_ocr_done(self, inv_no, texts=None):
        """OCR识别完成后的回调，自动填充所有可识别字段，不弹窗"""
        self.ocr_btn.config(state='normal', text="识别发票号")

        if not texts:
            self.app.set_status("识别完成，未提取到有效信息，请手动填写")
            return

        # 自动判断类型
        auto_type = self._detect_document_type(texts)
        if auto_type:
            self.inv_type_var.set(auto_type)

        # 提取并填充各字段（识别到就覆盖，因为用户主动点了识别按钮）
        extracted = self._extract_invoice_fields(texts)

        # 发票号/交易单号：优先用ocr识别的发票号，其次用字段提取的交易单号
        if not inv_no and extracted.get('invoice_number'):
            inv_no = extracted['invoice_number']
        if inv_no:
            self.inv_no_var.set(inv_no)

        # 日期（识别到就覆盖，默认值是今天会挡住）
        if extracted.get('date'):
            self.inv_date_var.set(extracted['date'])

        # 金额
        if extracted.get('amount'):
            self.amount_var.set(str(extracted['amount']))

        # 销售方
        if extracted.get('seller'):
            self.seller_var.set(extracted['seller'])

        # 备注
        if extracted.get('remark'):
            self.remark_var.set(extracted['remark'])

        # 状态栏提示（不弹窗）
        parts = []
        if inv_no:
            parts.append(f"单号 {inv_no}")
        if auto_type:
            parts.append(f"类型 {auto_type}")
        if extracted.get('date'):
            parts.append(f"日期 {extracted['date']}")
        if extracted.get('amount'):
            parts.append(f"金额 {extracted['amount']}")
        if extracted.get('seller'):
            parts.append(f"销售方 {extracted['seller'][:15]}")
        msg = "识别完成，已自动填充：" + "、".join(parts) if parts else "识别完成，请手动补充信息"
        self.app.set_status(msg)

    def _extract_invoice_fields(self, texts):
        """从OCR文本中提取日期、金额、销售方、备注、发票号（支持发票、微信/支付宝支付记录）"""
        import re
        result = {}
        if isinstance(texts, list):
            lines = [str(t).strip() for t in texts if t]
            full_text = '\n'.join(lines)
        else:
            full_text = str(texts)
            lines = full_text.split('\n')

        # 判断是否为支付记录（微信/支付宝）
        # 发票强特征词优先：只要含发票号码/代码/价税合计/开票日期等，一律按发票处理，
        # 避免"收款方名称"等发票字段误触支付判定导致金额错乱
        _invoice_strong = any(kw in full_text for kw in
                             ['发票号码', '发票代码', '价税合计', '开票日期',
                              '发票专用章', '发票联', '通用机打发票', '电子发票',
                              '增值税', '机动车销售统一发票', '旅客运输服务'])
        is_payment = (not _invoice_strong) and any(kw in full_text for kw in
                        ['微信支付', '支付宝', '转账时间', '支付时间', '付款方留言',
                         '收款方备注', '转账单号', '交易成功', '支付成功',
                         '扫二维码付款', '账单详情', '商品说明', '缴费成功',
                         '缴费状态', '缴费单位', '缴费编号', '商户全称',
                         '交易详情', '订单详情', '支付详情', '到账', '实付金额',
                         '交易方式', '商家订单号', '商户单号'])
        result['is_payment'] = is_payment

        # ========== 1. 提取日期 ==========
        date = None
        # 优先匹配发票的"开票日期"（支持冒号、空格、年月日间空格分隔）
        for pat in [r'开票日期[：:\s]+(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日',
                    r'开票日期[：:\s]+(\d{4})[-/](\d{1,2})[-/](\d{1,2})',
                    r'开\s*票\s*日\s*期[：:\s]*(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日',
                    r'开\s*票\s*日\s*期[：:\s]*(\d{4})[-/](\d{1,2})[-/](\d{1,2})']:
            m = re.search(pat, full_text)
            if m:
                date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                break
        # 支付记录：转账时间/支付时间
        if not date:
            for prefix in ['转账时间', '支付时间', '交易时间', '付款时间']:
                m = re.search(rf'{prefix}[：:\s]*(\d{{4}})年\s*(\d{{1,2}})月\s*(\d{{1,2}})日', full_text)
                if m:
                    date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                    break
                m = re.search(rf'{prefix}[：:\s]*(\d{{4}})[./-](\d{{1,2}})[./-](\d{{1,2}})', full_text)
                if m:
                    date = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                    break
        # 通用日期格式（取第一个合理的日期，排除未来日期）
        if not date:
            today = datetime.now().date()
            for pat in [r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日',
                        r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})']:
                matches = re.findall(pat, full_text)
                for m in matches:
                    try:
                        y, mo, d = int(m[0]), int(m[1]), int(m[2])
                        if 2000 <= y <= today.year and 1 <= mo <= 12 and 1 <= d <= 31:
                            candidate = f"{y:04d}-{mo:02d}-{d:02d}"
                            # 排除未来日期
                            if datetime.strptime(candidate, '%Y-%m-%d').date() <= today:
                                date = candidate
                                break
                    except (ValueError, IndexError):
                        continue
                if date:
                    break
        if date:
            result['date'] = date

        # ========== 2. 提取金额 ==========
        amount = None
        # 剔除日期/时间串，避免 "2024-10-06" "-05-14" "-14:00" 等干扰金额提取
        clean = re.sub(r'(?<!\d)(?:19|20)\d{2}[./-]\d{1,2}[./-]\d{1,2}', ' ', full_text)  # 2024-10-06 / 2024.04.18（容忍时间粘连如 2024-09-1023:02）
        clean = re.sub(r'(?<!\d)(?:19|20)\d{2}[-/]\d{1,2}', ' ', clean)                    # 2024-10（19/20 前缀确保不误删金额如 1020.00）
        clean = re.sub(r'(?<!\d)\d{1,2}[-/]\d{1,2}(?!\d)', ' ', clean)                      # 05-14 月-日
        clean = re.sub(r'\d{1,2}:\d{1,2}(:\d{1,2})?', ' ', clean)                           # 14:00 时间
        # 发票：价税合计
        m = re.search(r'(?:价税合计|小写)[^0-9￥¥]*[￥¥]?\s*([\d,]+\.?\d*)', full_text)
        if m:
            amount = m.group(1).replace(',', '')
        # 铁路/动车/飞机票：票价 ¥X.XX
        if not amount:
            m = re.search(r'票价[：:\s]*[￥¥]?\s*([\d,]+\.?\d*)', full_text)
            if m:
                amount = m.group(1).replace(',', '')
        # 出租车/客运发票：金额 X.XX元
        if not amount:
            m = re.search(r'金额[：:\s]*([\d,]+\.?\d*)\s*元', full_text)
            if m:
                amount = m.group(1).replace(',', '')
        # 出租车发票：金额 X-XX元（OCR常把小数点误识别成负号，如 7-90元 = 7.90元）
        if not amount:
            m = re.search(r'金额[：:\s]*(\d+)-(\d{1,2})\s*元', full_text)
            if m:
                amount = f"{int(m.group(1))}.{m.group(2)}"
        # 出租车竖排卷式发票：X-XX元（"金额"标签在数值之后，无前缀直接匹配）
        if not amount:
            m = re.search(r'(?<!\d)(\d{1,4})-(\d{1,2})\s*元', full_text)
            if m:
                amount = f"{int(m.group(1))}.{m.group(2)}"
        # 高速通行费发票：金额¥:XX 或 金额Y:XX（OCR可能把¥识别成Y）
        if not amount:
            m = re.search(r'金额[¥￥Y][：:\s]*([\d,]+\.?\d*)', full_text)
            if m:
                amount = m.group(1).replace(',', '')
        # 出租车发票：金额后面跟数字（无"元"字时）
        if not amount:
            for i, line in enumerate(lines):
                if '金额' in line and '价税' not in line:
                    m = re.search(r'金额[：:\s¥￥Y]*([\d,]+\.?\d*)', line)
                    if m:
                        val = m.group(1).replace(',', '')
                        try:
                            if 0.01 <= float(val) <= 99999:
                                amount = val
                                break
                        except ValueError:
                            pass
        # 发票通用回退：找所有¥金额，取合理范围内的最大值
        if not amount and not is_payment:
            yuan_amounts = re.findall(r'[￥¥]\s*([\d,]+\.?\d*)', full_text)
            valid_amounts = []
            for a in yuan_amounts:
                try:
                    fval = float(a.replace(',', ''))
                    if 0.01 <= fval <= 99999:
                        valid_amounts.append(fval)
                except ValueError:
                    pass
            if valid_amounts:
                amount = f"{max(valid_amounts):.2f}"
        elif not amount and is_payment:
            # 支付记录：优先取带负号的金额（-0.10, -142.00）
            decimal_amounts = []  # 带小数点的金额（优先，保留正负号）
            int_amounts = []      # 整数金额（保留正负号）
            for m in re.finditer(r'([-－])\s*(\d+\.?\d*)', clean):
                sign = m.group(1)
                val = m.group(2)
                try:
                    fval = float(val.replace(',', ''))
                    if sign in '-－':
                        fval = -fval
                    if -999999 <= fval <= 999999:
                        if '.' in val:
                            decimal_amounts.append(fval)
                        else:
                            int_amounts.append(fval)
                except ValueError:
                    pass
            # 优先用带小数点的金额（支付记录金额通常带小数）
            if decimal_amounts:
                amount = f"{max(decimal_amounts):.2f}"
            elif int_amounts:
                amount = f"{max(int_amounts):.2f}"
            # ￥金额（无负号时取绝对值最大的合理金额）
            if not amount:
                amounts = re.findall(r'[￥¥]\s*([\d,]+\.?\d*)', full_text)
                if amounts:
                    try:
                        amount = max(float(a.replace(',', '')) for a in amounts)
                        amount = f"{amount:.2f}"
                    except ValueError:
                        pass
            # 纯数字金额（无符号无￥，如缴费详情里的 500.00）
            if not amount:
                nums = re.findall(r'(?<![\d.])(\d{1,7}\.\d{1,2})(?![\d.])', clean)
                if nums:
                    try:
                        amount = f"{max(float(a) for a in nums):.2f}"
                    except ValueError:
                        pass
        # ===== 通用金额回退（发票与支付记录均适用，兜底） =====
        if not amount:
            # X.XX元 形式（如 共180.00元 / 支付180.00元）
            m = re.search(r'([\d,]+\.\d{1,2})\s*元', full_text)
            if m:
                try:
                    fval = float(m.group(1).replace(',', ''))
                    if 0.01 <= fval <= 999999:
                        amount = f"{fval:.2f}"
                except ValueError:
                    pass
        if not amount:
            # 纯 X.XX 数字（银行转账截图等），排除日期时间
            nums = re.findall(r'(?<![\d.])(\d{1,7}\.\d{1,2})(?![\d.])', clean)
            if nums:
                try:
                    amount = f"{max(float(a) for a in nums):.2f}"
                except ValueError:
                    pass
        if amount:
            try:
                result['amount'] = f"{float(amount):.2f}"
            except ValueError:
                pass

        # ========== 3. 提取销售方/收款方 ==========
        seller = None
        if is_payment:
            # 微信：扫二维码付款-给XXX / 转账给XXX
            m = re.search(r'(?:给|转账给|收款方)\s*([^\n\d\-]{2,30})', full_text)
            if m:
                seller = m.group(1).strip()
            # 支付宝：对方账户 XXX / 收款方名称在金额上方
            if not seller:
                m = re.search(r'对方账户[：:\s]*([^\n\d\*]{2,20})', full_text)
                if m:
                    seller = m.group(1).strip()
            # 支付宝/微信：找金额行（-X.XX）上方最近的非空文字作为商家/收款方
            if not seller:
                for i, line in enumerate(lines):
                    if re.match(r'^[-－]\s*\d+\.?\d*$', line.strip()):
                        # 金额行，往上找商家名（排除"交易成功"等状态文字）
                        for j in range(i-1, max(i-5, -1), -1):
                            txt = lines[j].strip()
                            if txt and len(txt) > 1 and len(txt) <= 30 \
                                    and not re.match(r'^[\d\s\-.,:：]+$', txt) \
                                    and not any(kw in txt for kw in ['交易成功', '支付成功', '转账成功', '账单详情', '全部账单']):
                                seller = txt
                                break
                        break
        else:
            # 铁路/动车/飞机票：销售方为铁路/航空公司
            is_railway = any(kw in full_text for kw in [
                '铁路', '电子客票', '票价', '次列车', '开车', '座位号',
                '一等座', '二等座', '软卧', '硬卧', '硬座', '无座',
                '仅供报销使用', '报销凭证', '遗失不补', '退票改签', '检票口', '须交回车站'
            ])
            if is_railway:
                # 尝试提取出发站作为销售方参考，否则用中国铁路
                station_match = re.search(r'([\u4e00-\u9fa5]{2,8}站)\s*$', full_text, re.MULTILINE)
                if station_match:
                    seller = f"中国铁路（{station_match.group(1)}出发）"
                else:
                    seller = "中国铁路"
            # 高速通行费发票：销售方在标题中（XX高速公路有限责任公司通用机打发票）
            if not seller:
                is_highway = any(kw in full_text for kw in ['高速公路', '通行费', '入口:', '出口:', '车型:'])
                if is_highway:
                    m = re.search(r'([\u4e00-\u9fa5]{2,20}(?:高速公路|公路|交通|投资)[\u4e00-\u9fa5]{0,10}(?:有限责任公司|有限公司|公司))', full_text)
                    if m:
                        seller = m.group(1).strip()
                    else:
                        # 从标题行提取：去掉"通用机打发票"等后缀
                        for line in lines:
                            if '通用机打发票' in line or ('发票' in line and ('高速' in line or '公路' in line)):
                                seller = re.sub(r'通用机打发票|发票联|发票$', '', line).strip()
                                if seller and len(seller) > 2:
                                    break
            # 发票：电子发票布局是购买方在左、销售方在右
            if not seller:
                names = re.findall(r'名称[：:]\s*(.+)', full_text)
                has_buyer = '购买方' in full_text
                has_seller = '销售方' in full_text
                if has_buyer and has_seller and len(names) >= 2:
                    seller = names[1].strip()
                elif has_seller and len(names) >= 1:
                    seller_section = False
                    for line in lines:
                        if '销售方' in line:
                            seller_section = True
                            continue
                        if seller_section and '名称' in line:
                            m = re.search(r'名称[：:]\s*(.+)', line)
                            if m:
                                seller = m.group(1).strip()
                                break
                        if '购买方' in line and seller_section:
                            break
                    if not seller and names:
                        seller = names[-1].strip()
                elif len(names) >= 2:
                    seller = names[1].strip()
                elif len(names) == 1:
                    seller = names[0].strip()

        if seller:
            seller = re.sub(r'\s+', ' ', seller).strip()
            # 清理末尾的标点符号和特殊符号（>、<、★、☆等）
            seller = seller.rstrip('，。、；：>><<★☆\t ')
            # 清理末尾的箭头符号（如 "堃煌>"）
            seller = re.sub(r'[>＞]+$', '', seller).strip()
            if len(seller) > 50:
                seller = seller[:50]
            result['seller'] = seller

        # ========== 4. 提取备注/附言/留言 ==========
        remark = None
        # 排除列表：明显不是备注的内容
        exclude_patterns = ['二维码收款', '扫二维码付款', '转账给', '收款方', '付款方',
                           '交易成功', '支付成功', '转账时间', '支付时间', '交易时间',
                           '付款时间', '商品说明', '订单号', '交易单号', '转账单号',
                           '商户单号', '发票号码', '发票号', '金额', '价税合计',
                           '添加', '请选择', '更多', '查看', '申请', '对此订单',
                           '结算单', '流水号', '零售结算', '订单描述',
                           '开票人', '复核', '收款人', '开票员', '机器编号', '校验码']

        def is_valid_remark(val):
            """校验备注内容是否合理"""
            if not val or len(val) < 1:
                return False
            if len(val) > 100:
                return False
            # 排除纯数字、纯符号
            if re.match(r'^[\d\s\-.,￥¥:：>><]+$', val):
                return False
            # 排除明显是交易类型/标题/占位符/订单号的内容
            for ex in exclude_patterns:
                if ex in val:
                    return False
            # 排除看起来像订单号的（字母+数字组合，长度>10）
            if re.search(r'[A-Za-z]{2,}\d{8,}', val) and len(val) > 10:
                return False
            return True

        # 第一优先级：明确的用户留言字段（微信/支付宝用户主动填写）
        high_priority_keywords = ['付款方留言', '收款理由', '付款留言', '留言', '转账说明', '附言']
        for i, line in enumerate(lines):
            for kw in high_priority_keywords:
                if kw in line:
                    # 当前行匹配（冒号或空格分隔）
                    m = re.search(rf'{kw}[：:\s]*(.+)', line)
                    if m:
                        val = m.group(1).strip()
                        if is_valid_remark(val):
                            remark = val
                            break
                    # 当前行只有关键词，检查下一行
                    if not remark and i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if is_valid_remark(next_line):
                            remark = next_line
                            break
            if remark:
                break

        # 第二优先级：用户填写的备注/摘要/用途（排除商家自动生成的商品说明）
        if not remark:
            low_priority_keywords = ['收款方备注', '备注', '摘要', '用途']
            for i, line in enumerate(lines):
                for kw in low_priority_keywords:
                    if kw in line:
                        m = re.search(rf'{kw}[：:\s]*(.+)', line)
                        if m:
                            val = m.group(1).strip()
                            if is_valid_remark(val):
                                remark = val
                                break
                        if not remark and i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            if is_valid_remark(next_line):
                                remark = next_line
                                break
                if remark:
                    break

        if remark:
            remark = re.sub(r'\s+', ' ', remark).strip()
            remark = remark.rstrip('，。、；：')
            result['remark'] = remark

        # ========== 5. 提取发票号/交易单号 ==========
        inv_no = None
        if is_payment:
            # 支付记录：只有明确出现"发票号码/发票号"才提取，交易单号/订单号/商户单号不作为发票号
            for kw in ['发票号码', '发票号']:
                m = re.search(rf'{kw}[：:]\s*([A-Za-z0-9]+)', full_text)
                if m:
                    inv_no = m.group(1)
                    break
        else:
            for kw in ['发票号码', '发票号', '号码', '转账单号', '交易单号', '订单号', '商户单号']:
                m = re.search(rf'{kw}[：:\s]*([A-Za-z0-9]+)', full_text)
                if m:
                    inv_no = m.group(1)
                    break
            if not inv_no:
                # 找8-28位连续数字（排除日期、排除印刷范围号如 00000001-00200000）
                all_nums = re.findall(r'\b(\d{8,28})\b', full_text)
                # 收集被短横线连接的印刷范围号
                range_nums = set()
                for m in re.finditer(r'(\d{6,})-(\d{6,})', full_text):
                    range_nums.add(m.group(1))
                    range_nums.add(m.group(2))
                for n in all_nums:
                    if n in range_nums:
                        continue
                    if not (len(n) == 8 and n.startswith(('19', '20'))):
                        inv_no = n
                        break
        if inv_no:
            result['invoice_number'] = inv_no

        return result

    def _parse_payment_list(self, texts):
        '''从支付记录列表截图（微信账单/支付宝账单/电信充值记录/生活缴费等）的OCR文本中提取多条支付记录。
        返回列表，每条包含 {merchant, amount, date, remark, is_payment: True}
        '''
        import re
        if not texts:
            return []
        if isinstance(texts, list):
            lines = [str(t).strip() for t in texts if t and str(t).strip()]
        else:
            lines = str(texts).split('\n')
            lines = [l.strip() for l in lines if l.strip()]
        full_text = '\n'.join(lines)

        # ========== 第一步：找所有月份标题（YYYY年MM月），用于补全无年份日期 ==========
        month_headers = []
        for m in re.finditer(r'(\d{4})年\s*(\d{1,2})月', full_text):
            try:
                y, mo = int(m.group(1)), int(m.group(2))
                if 2000 <= y <= 2100 and 1 <= mo <= 12:
                    month_headers.append({'year': y, 'month': mo, 'pos': m.start()})
            except (ValueError, IndexError):
                continue

        def get_year_for_pos(pos):
            '''根据位置查找最近的月份标题，返回年份'''
            best = None
            for mh in month_headers:
                if mh['pos'] <= pos:
                    best = mh
                else:
                    break
            return best['year'] if best else datetime.now().year

        # ========== 第二步：找所有日期（带年份和不带年份） ==========
        all_dates = []

        # 带年份的完整日期：YYYY年M月D日 或 YYYY-MM-DD
        for pat in [
            r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日\s*(\d{1,2}:\d{2})?',
            r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s*(\d{1,2}:\d{2})?',
        ]:
            for m in re.finditer(pat, full_text):
                try:
                    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                    if 2000 <= y <= 2100 and 1 <= mo <= 12 and 1 <= d <= 31:
                        all_dates.append({
                            'date': f"{y:04d}-{mo:02d}-{d:02d}",
                            'pos': m.start(),
                        })
                except (ValueError, IndexError):
                    continue

        # 不带年份的日期：M月D日（生活缴费格式，如 01月19日16:30）
        for m in re.finditer(r'(?<!\d)(\d{1,2})月\s*(\d{1,2})日\s*(\d{1,2}:\d{2})?', full_text):
            try:
                mo, d = int(m.group(1)), int(m.group(2))
                if 1 <= mo <= 12 and 1 <= d <= 31:
                    y = get_year_for_pos(m.start())
                    all_dates.append({
                        'date': f"{y:04d}-{mo:02d}-{d:02d}",
                        'pos': m.start(),
                    })
            except (ValueError, IndexError):
                continue

        if len(all_dates) < 2:
            return []

        all_dates.sort(key=lambda x: x['pos'])

        # ========== 第三步：金额模式 ==========
        amount_patterns = [
            r'([-－])\s*(\d+\.?\d*)',
            r'[￥¥]\s*(\d+\.?\d*)',
            r'(\d+\.?\d*)\s*元',
        ]

        # ========== 第四步：生活缴费类别关键词 ==========
        utility_categories = ['电费', '水费', '燃气费', '煤气费', '宽带费', '物业费',
                              '取暖费', '有线电视费', '通讯费', '手机费', '固话费']

        records = []
        seen = set()

        for i, dt in enumerate(all_dates):
            search_start = max(0, dt['pos'] - 500)
            search_end = all_dates[i + 1]['pos'] if i + 1 < len(all_dates) else len(full_text)
            segment = full_text[search_start:search_end]

            # 找金额
            amount = None
            for pat in amount_patterns:
                for m in re.finditer(pat, segment):
                    try:
                        if m.group(1) in '-－':
                            val = -float(m.group(2).replace(',', ''))
                        else:
                            val = float(m.group(1).replace(',', ''))
                        if -999999 <= val <= 999999 and val != 0:
                            # 过滤网络速度（如 9.00 K/s）
                            amt_text = m.group(0)
                            seg_after = segment[m.end():m.end()+15] if m.end() < len(segment) else ''
                            if re.search(r'[KMGkmg]/?s', seg_after) or 'K/s' in amt_text or 'M/s' in amt_text:
                                continue
                            # 过滤明显的网速数值（小于10且附近有K/s/M/s）
                            if abs(val) < 10 and re.search(r'[KMGkmg]/?s', segment[max(0,m.start()-20):m.end()+20]):
                                continue
                            if amount is None or abs(val) > abs(amount):
                                amount = val
                    except (ValueError, IndexError):
                        continue
                if amount is not None:
                    break

            if amount is None:
                continue

            # 找商家/类别
            merchant = None
            # 优先找生活缴费类别
            for cat in utility_categories:
                if cat in segment:
                    merchant = cat
                    break
            # 再找金额行上方的商家名
            if not merchant:
                seg_lines = segment.split('\n')
                amount_abs = abs(amount)
                amount_str = f"{amount_abs:.2f}"
                amount_int = str(int(amount_abs))
                for j, line in enumerate(seg_lines):
                    if amount_str in line or amount_int in line or re.search(rf'[-－￥¥]?\s*{amount_int}', line):
                        for k in range(max(0, j - 5), j):
                            txt = seg_lines[k].strip()
                            if (txt and 2 <= len(txt) <= 40
                                    and not re.match(r'^[\d\s\-.,:：年月日时分￥¥元]+$', txt)
                                    and not any(kw in txt for kw in [
                                        '交易成功', '支付成功', '转账成功', '账单详情', '全部账单',
                                        '当前状态', '支付方式', '付款方式', '未知支付方式',
                                        '查看更多', '仅支持查找', '找不到想要', '下载账单',
                                        '充值记录', '查询号码', '安装地址', '余额查询',
                                        '生活缴费', '微信支付', '支付宝',
                                        # 新增：常见UI按钮和状态文字
                                        '查看详情', '赚奖励', '取消', '确认', '删除', '编辑',
                                        'K/s', 'M/s', 'KB/s', 'MB/s', '网速', '网络',
                                        '服务', '客服', '帮助', '设置', '更多', '返回',
                                        '未知商家', '暂无数据', '加载中', '正在加载',
                                        '（84）', '（8', '）',
                                        '已完成', '已支付', '已退款', '退款中', '待支付',
                                        '收入', '支出', '余额', '零钱', '银行卡',
                                    ])):
                                merchant = txt
                                break
                        break
            if not merchant:
                candidates = re.findall(r'[\u4e00-\u9fa5（）()]{2,30}', segment)
                for c in candidates:
                    if not any(kw in c for kw in ['年月', '日时分', '支付方式', '充值记录', '查询', '账单', '生活缴费']):
                        merchant = c
                        break

            # 备注
            remark = None
            for kw in ['留言', '备注', '商品说明', '缴费项目', '收费项目', '商品']:
                m = re.search(rf'{kw}[：:\s]*([^\n]{{2,30}})', segment)
                if m:
                    remark = m.group(1).strip()
                    break

            # 商家名后处理：过滤UI元素
            ui_merchants = {'查看详情', '赚奖励', '取消', '确认', '删除', '编辑', 'K/s', 'M/s',
                           '服务', '客服', '帮助', '设置', '更多', '返回', '未知商家', '暂无数据',
                           '已完成', '已支付', '已退款', '退款中', '待支付', '收入', '支出',
                           '余额', '零钱', '银行卡', '安装地址', '查询号码'}
            # 商家名统一
            if merchant == '电信' or merchant == '中国电信':
                merchant = '中国电信'
            if merchant in ui_merchants or (merchant and len(merchant) <= 1):
                # 尝试从生活缴费类别中找
                for cat in utility_categories:
                    if cat in segment:
                        merchant = cat
                        break
                else:
                    # 尝试从关键词推断
                    if '电信' in segment or '宽带' in segment:
                        merchant = '中国电信'
                    elif '翼支付' in segment:
                        merchant = '翼支付'
                    else:
                        merchant = '未识别'
            rec = {
                'merchant': merchant or '未识别',
                'amount': f"{abs(amount):.2f}",
                'date': dt['date'],
                'remark': remark or '',
                'is_payment': True,
                'invoice_number': '',
            }
            key = (rec['date'], rec['amount'], rec['merchant'])
            if key not in seen:
                seen.add(key)
                records.append(rec)

        return records


    def _detect_document_type(self, texts):
        """根据OCR文本判断单据类型，返回'发票'、'支付记录'或None"""
        if not texts:
            return None
        if isinstance(texts, list):
            full_text = ' '.join(str(t) for t in texts)
        else:
            full_text = str(texts)

        # 发票关键词
        invoice_keywords = ['发票号码', '发票代码', '纳税人识别号', '价税合计',
                           '开票日期', '购买方', '销售方', '发票专用章', '税额',
                           '税率', '校验码', '机器编号']
        # 支付记录关键词（微信/支付宝/银行转账等）
        payment_keywords = ['支付', '转账', '付款', '收款', '交易', '流水',
                           '商户', '订单号', '交易号', '支付宝', '微信',
                           '财付通', '银行转账', '付款方', '收款方', '交易时间',
                           '转账时间', '支付时间', '付款方留言', '收款方备注',
                           '转账单号', '交易成功', '支付成功', '扫二维码付款',
                           '账单详情', '商品说明', '全部账单', '当前状态',
                           '支付方式', '付款方式']

        has_invoice = any(kw in full_text for kw in invoice_keywords)
        has_payment = any(kw in full_text for kw in payment_keywords)

        if has_invoice:
            return '发票'
        if has_payment:
            return '支付记录'
        return None

    def add_or_update(self):
        inv_no = self.inv_no_var.get().strip()
        inv_type = self.inv_type_var.get().strip() or '发票'
        inv_date = self.inv_date_var.get().strip()
        reimburser = self.reimburser_var.get().strip()
        amount = self.amount_var.get().strip()
        purpose = self.purpose_var.get().strip()
        voucher = self.voucher_var.get().strip()
        bank_name = self.bank_var.get().strip()
        seller = self.seller_var.get().strip()
        remark = self.remark_var.get().strip()
        bank_id = self._bank_map.get(bank_name) if bank_name else None

        if not reimburser:
            messagebox.showwarning("提示", "请选择报销人")
            return
        try:
            amount = float(amount) if amount else 0
        except ValueError:
            messagebox.showwarning("提示", "金额必须是数字")
            return

        conn = get_db()
        # 发票号重复校验（非空时检查，修改时排除自身）
        if inv_no:
            sql = "SELECT id FROM invoices WHERE invoice_number=?"
            params = [inv_no]
            if self.selected_id:
                sql += " AND id != ?"
                params.append(self.selected_id)
            exist = conn.execute(sql, params).fetchone()
            if exist:
                conn.close()
                messagebox.showerror("发票号重复",
                                     f"发票号「{inv_no}」已存在！\n\n"
                                     f"该发票已录入，请勿重复报销。")
                return

        emp = conn.execute("SELECT id FROM employees WHERE name=?", (reimburser,)).fetchone()
        if not emp:
            messagebox.showerror("错误", "报销人不存在，请先在人员管理中添加")
            conn.close()
            return
        emp_id = emp['id']

        if self.selected_id:
            # 修改模式
            img_path = None
            if self.image_file:
                # 用户选择了新图片，替换
                try:
                    # 先删除旧图片
                    old = conn.execute("SELECT image_path FROM invoices WHERE id=?",
                                       (self.selected_id,)).fetchone()
                    if old and old['image_path']:
                        oldp = os.path.join(INVOICE_DIR, old['image_path'])
                        if os.path.exists(oldp):
                            try:
                                os.remove(oldp)
                            except Exception:
                                pass
                    ext = os.path.splitext(self.image_file)[1].lower()
                    fname = f"inv_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{ext}"
                    dst = os.path.join(INVOICE_DIR, fname)
                    shutil.copy2(self.image_file, dst)
                    img_path = fname
                except Exception as e:
                    messagebox.showerror("错误", f"图片复制失败: {e}")
                    conn.close()
                    return
                new_status = 1 if self.status_var.get() == '已报销' else 0
                conn.execute("""UPDATE invoices SET invoice_number=?, type=?, invoice_date=?, reimburser_id=?,
                             amount=?, purpose=?, voucher_number=?, bank_account_id=?, image_path=?, seller=?, remark=?, status=?, batch_pending=0 WHERE id=?""",
                             (inv_no, inv_type, inv_date, emp_id, amount, purpose, voucher, bank_id, img_path, seller, remark, new_status, self.selected_id))
            else:
                new_status = 1 if self.status_var.get() == '已报销' else 0
                conn.execute("""UPDATE invoices SET invoice_number=?, type=?, invoice_date=?, reimburser_id=?,
                             amount=?, purpose=?, voucher_number=?, bank_account_id=?, seller=?, remark=?, status=?, batch_pending=0 WHERE id=?""",
                             (inv_no, inv_type, inv_date, emp_id, amount, purpose, voucher, bank_id, seller, remark, new_status, self.selected_id))
            # 处理报销状态变化对应的银行记录
            old_row = conn.execute("SELECT status, bank_account_id, amount, purpose, invoice_number, invoice_date FROM invoices WHERE id=?", (self.selected_id,)).fetchone()
            # Note: old_row is fetched AFTER update, so status is already new. We need to track old status.
            # Simpler: delete existing bank tx for this invoice, then if new_status=1 and has bank, create new one
            conn.execute("DELETE FROM bank_transactions WHERE source='发票报销' AND source_id=?", (self.selected_id,))
            if new_status == 1 and bank_id:
                purpose = f"发票报销：{purpose or ''}（发票号：{inv_no or '无'}）"
                conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                              source, source_id, transaction_date, created_at) VALUES(?,?,?,?,?,?,?,?)""",
                             (bank_id, '出项', amount, purpose, '发票报销',
                              self.selected_id, inv_date or datetime.now().strftime('%Y-%m-%d'),
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()
            self.app.set_status(f"已修改发票: {inv_no or '(无号)'}")
        else:
            # 添加模式
            img_path = None
            if self.image_file:
                try:
                    ext = os.path.splitext(self.image_file)[1].lower()
                    fname = f"inv_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{ext}"
                    dst = os.path.join(INVOICE_DIR, fname)
                    shutil.copy2(self.image_file, dst)
                    img_path = fname
                except Exception as e:
                    messagebox.showerror("错误", f"图片复制失败: {e}")
                    conn.close()
                    return
            conn.execute("""INSERT INTO invoices(invoice_number, type, invoice_date, reimburser_id, amount,
                          purpose, voucher_number, bank_account_id, image_path, status, created_at, seller, remark)
                          VALUES(?,?,?,?,?,?,?,?,?,?,0,?,?)""",
                         (inv_no, inv_type, inv_date, emp_id, amount, purpose, voucher, bank_id, img_path,
                          datetime.now().strftime('%Y-%m-%d %H:%M:%S'), seller, remark))
            conn.commit()
            conn.close()
            self.app.set_status(f"已添加{inv_type}: {inv_no or '(无号)'} 金额 {fmt_money(amount)}")

        self.cancel_edit()
        # 添加/修改后重置年份筛选为全部，确保新记录可见
        self.filter_year.set('全部')
        self.filter_month.set('全部')
        self.refresh()  # 立即刷新当前页
        # 延迟异步刷新其他页
        self.app.after(30, self.app._refresh_other_tabs, 'inv')

    def set_status(self, status):
        if not self.selected_id:
            messagebox.showwarning("提示", "请先选择发票")
            return
        conn = get_db()
        row = conn.execute("SELECT * FROM invoices WHERE id=?", (self.selected_id,)).fetchone()
        if not row:
            conn.close()
            return
        conn.execute("UPDATE invoices SET status=? WHERE id=?", (status, self.selected_id))
        if status == 1 and row['bank_account_id']:
            # 标记已报销：生成银行出项记录（如果不存在）
            exist = conn.execute("""SELECT id FROM bank_transactions WHERE source='发票报销'
                                  AND source_id=?""", (self.selected_id,)).fetchone()
            if not exist:
                purpose = f"发票报销：{row['purpose'] or ''}（发票号：{row['invoice_number'] or '无'}）"
                conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                              source, source_id, transaction_date, created_at) VALUES(?,?,?,?,?,?,?,?)""",
                             (row['bank_account_id'], '出项', row['amount'], purpose, '发票报销',
                              self.selected_id, row['invoice_date'] or datetime.now().strftime('%Y-%m-%d'),
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        elif status == 0:
            # 标记未报销：删除对应的银行出项记录
            conn.execute("DELETE FROM bank_transactions WHERE source='发票报销' AND source_id=?",
                         (self.selected_id,))
        conn.commit()
        conn.close()
        self.refresh()
        self.app.stat_tab.refresh()
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()
        self.app.set_status("已更新报销状态")

    def delete(self):
        if not self.selected_id:
            messagebox.showwarning("提示", "请先选择要删除的发票")
            return
        if not messagebox.askyesno("确认", "确定删除该发票记录？"):
            return
        conn = get_db()
        row = conn.execute("SELECT image_path FROM invoices WHERE id=?", (self.selected_id,)).fetchone()
        img = row['image_path'] if row else None
        conn.execute("DELETE FROM bank_transactions WHERE source='发票报销' AND source_id=?", (self.selected_id,))
        conn.execute("DELETE FROM invoices WHERE id=?", (self.selected_id,))
        conn.commit()
        conn.close()
        if img:
            p = os.path.join(INVOICE_DIR, img)
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
        self.selected_id = None
        self.refresh()  # 立即刷新当前页
        self.app.set_status("已删除发票")
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()
        # 延迟异步刷新其他页
        self.app.after(30, self.app._refresh_other_tabs, 'inv')

    def select_all(self):
        """全选当前筛选结果"""
        for item in self.tree.get_children():
            try:
                inv_id = int(item)
            except (ValueError, TypeError):
                inv_id = None
            if inv_id:
                self.checked_ids.add(inv_id)
        self.refresh()

    def deselect_all(self):
        """取消全选"""
        self.checked_ids.clear()
        self.refresh()

    def range_select(self):
        """按编号批量勾选，支持连续(3-9)、断开(3,6,9)、混合(3-6,9)"""
        from tkinter import simpledialog
        raw = simpledialog.askstring("范围选择",
            "输入编号（支持：3-9 连续、3,6,9 断开、3-6,9 混合）：",
            parent=self)
        if not raw:
            return
        # 解析编号集合
        target_ids = set()
        for part in raw.replace('，', ',').replace('—', '-').replace('～', '-').split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                ab = part.split('-', 1)
                try:
                    a, b = int(ab[0].strip()), int(ab[1].strip())
                    if a > b:
                        a, b = b, a
                    for n in range(a, b + 1):
                        target_ids.add(n)
                except (ValueError, TypeError):
                    continue
            else:
                try:
                    target_ids.add(int(part))
                except (ValueError, TypeError):
                    continue
        if not target_ids:
            messagebox.showwarning("提示", "未解析到有效编号")
            return
        count = 0
        for item in self.tree.get_children():
            vals = self.tree.item(item, 'values')
            if len(vals) > 1:
                try:
                    seq = int(vals[1])
                    if seq in target_ids:
                        inv_id = int(item)
                        self.checked_ids.add(inv_id)
                        count += 1
                except (ValueError, TypeError):
                    continue
        self.refresh()
        self.app.set_status(f"范围选择已勾选 {count} 条（输入：{raw}）")

    def batch_set_status(self):
        """批量切换报销状态：未报销→已报销，已报销→未报销"""
        if not self.checked_ids:
            messagebox.showwarning("提示", "请先勾选要操作的发票")
            return
        conn = get_db()
        # 统计当前状态
        paid_count = 0
        unpaid_count = 0
        for inv_id in list(self.checked_ids):
            row = conn.execute("SELECT status FROM invoices WHERE id=?", (inv_id,)).fetchone()
            if row:
                if row['status'] == 1:
                    paid_count += 1
                else:
                    unpaid_count += 1
        conn.close()
        # 决定操作方向
        if paid_count > 0 and unpaid_count == 0:
            action = 'unreimburse'
            msg = f"选中的 {paid_count} 张发票当前都是已报销，确定要标记为未报销吗？"
        elif unpaid_count > 0 and paid_count == 0:
            action = 'reimburse'
            msg = f"确定将选中的 {unpaid_count} 张发票标记为已报销？"
        else:
            action = 'reimburse'
            msg = f"选中的 {len(self.checked_ids)} 张发票中，{unpaid_count} 张未报销、{paid_count} 张已报销。\n确定将未报销的标记为已报销？（已报销的保持不变）"
        if not messagebox.askyesno("确认", msg):
            return
        conn = get_db()
        count = 0
        for inv_id in list(self.checked_ids):
            row = conn.execute("SELECT * FROM invoices WHERE id=?", (inv_id,)).fetchone()
            if not row:
                continue
            if action == 'reimburse':
                if row['status'] == 1:
                    continue  # 已经是已报销，跳过
                conn.execute("UPDATE invoices SET status=1 WHERE id=?", (inv_id,))
                if row['bank_account_id']:
                    exist = conn.execute("SELECT id FROM bank_transactions WHERE source='发票报销' AND source_id=?", (inv_id,)).fetchone()
                    if not exist:
                        purpose = f"发票报销：{row['purpose'] or ''}（发票号：{row['invoice_number'] or '无'}）"
                        conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                                      source, source_id, transaction_date, created_at) VALUES(?,?,?,?,?,?,?,?)""",
                                     (row['bank_account_id'], '出项', row['amount'], purpose, '发票报销',
                                      inv_id, row['invoice_date'] or datetime.now().strftime('%Y-%m-%d'),
                                      datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                count += 1
            else:  # unreimburse
                if row['status'] != 1:
                    continue  # 不是已报销，跳过
                conn.execute("UPDATE invoices SET status=0 WHERE id=?", (inv_id,))
                # 删除对应的银行出项记录
                conn.execute("DELETE FROM bank_transactions WHERE source='发票报销' AND source_id=?", (inv_id,))
                count += 1
        conn.commit()
        conn.close()
        self.checked_ids.clear()
        self.refresh()
        self.app.stat_tab.refresh()
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()
        if action == 'reimburse':
            self.app.set_status(f"已批量报销 {count} 张发票")
        else:
            self.app.set_status(f"已批量取消报销 {count} 张发票")

    def batch_set_reimburser(self):
        """批量指定报销人（含待完善发票）。指定报销人后发票从待完善自动变为未报销。"""
        import tkinter as tk
        if not self.checked_ids:
            messagebox.showwarning("提示", "请先勾选要设置报销人的发票")
            return
        conn = get_db()
        persons = conn.execute("SELECT id, name, resigned FROM employees ORDER BY name").fetchall()
        conn.close()
        if not persons:
            messagebox.showwarning("提示", "请先在人员管理中添加人员")
            return
        win = tk.Toplevel(self)
        win.title("批量设置报销人")
        win.geometry("400x200")
        bg = '#1a1a1a' if is_dark_theme() else '#F5F7FA'
        win.configure(bg=bg)
        win.transient(self)
        win.grab_set()
        frame = ttk.Frame(win, padding=16)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text=f"为选中的 {len(self.checked_ids)} 张发票设置报销人：",
                  font=('微软雅黑', 11, 'bold')).pack(anchor='w', pady=(0, 10))
        var = tk.StringVar()
        active = [r['name'] for r in persons if not r['resigned']]
        if not active:
            active = [r['name'] for r in persons]
        cb = ttk.Combobox(frame, textvariable=var, values=active, state='readonly')
        cb.pack(fill='x', pady=(0, 14))
        if active:
            var.set(active[0])
        tip = ttk.Label(frame, text="仅设置报销人，不改变发票的报销状态", foreground='gray')
        tip.pack(anchor='w', pady=(0, 10))
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')
        ttk.Button(btn_frame, text="取消", command=win.destroy).pack(side='right', padx=4)
        def do_set():
            name = var.get().strip()
            if not name:
                messagebox.showwarning("提示", "请选择报销人")
                return
            emp_id = next((r['id'] for r in persons if r['name'] == name), None)
            if emp_id is None:
                return
            conn = get_db()
            ph = ','.join('?' * len(self.checked_ids))
            # 仅更新报销人，保留原有的报销状态（已报销/未报销/待完善）与 batch_pending 标记
            conn.execute(
                "UPDATE invoices SET reimburser_id=? WHERE id IN (%s)" % ph,
                tuple([emp_id] + list(self.checked_ids)))
            conn.commit()
            conn.close()
            count = len(self.checked_ids)
            self.checked_ids.clear()
            win.destroy()
            self.refresh()
            self.app.after(30, self.app._refresh_other_tabs, 'inv')
            self.app.set_status(f"已为 {count} 张发票设置报销人: {name}")
        ttk.Button(btn_frame, text="确定设置", command=do_set, style='Accent.TButton').pack(side='right', padx=4)

    def cancel_printed(self):
        """取消勾选发票的排版标记"""
        if not self.checked_ids:
            messagebox.showwarning("提示", "请先勾选要取消排版标记的发票")
            return
        if not messagebox.askyesno("确认", f"确定要取消 {len(self.checked_ids)} 张发票的排版标记吗？"):
            return
        conn = get_db()
        placeholders = ','.join('?' * len(self.checked_ids))
        conn.execute(f"UPDATE invoices SET printed=0 WHERE id IN ({placeholders})", tuple(self.checked_ids))
        conn.commit()
        conn.close()
        self.app.set_status(f"已取消 {len(self.checked_ids)} 张发票的排版标记")
        self.refresh()

    def batch_delete(self):
        """批量删除"""
        if not self.checked_ids:
            messagebox.showwarning("提示", "请先勾选要删除的发票")
            return
        if not messagebox.askyesno("确认", f"确定删除选中的 {len(self.checked_ids)} 张发票？此操作不可恢复！"):
            return
        conn = get_db()
        count = 0
        for inv_id in list(self.checked_ids):
            row = conn.execute("SELECT image_path FROM invoices WHERE id=?", (inv_id,)).fetchone()
            img = row['image_path'] if row else None
            conn.execute("DELETE FROM bank_transactions WHERE source='发票报销' AND source_id=?", (inv_id,))
            conn.execute("DELETE FROM invoices WHERE id=?", (inv_id,))
            if img:
                p = os.path.join(INVOICE_DIR, img)
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            count += 1
        conn.commit()
        conn.close()
        self.checked_ids.clear()
        self.refresh()
        self.app.stat_tab.refresh()
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()
        self.app.set_status(f"已批量删除 {count} 张发票")

    def view_image(self):
        if not self.selected_id:
            messagebox.showwarning("提示", "请先选择发票")
            return
        conn = get_db()
        row = conn.execute("SELECT id, image_path, invoice_number, amount, purpose, invoice_date FROM invoices WHERE id=?",
                           (self.selected_id,)).fetchone()
        conn.close()
        if not row or not row['image_path']:
            messagebox.showinfo("提示", "该发票没有上传图片/PDF")
            return
        p = os.path.join(INVOICE_DIR, row['image_path'])
        if not os.path.exists(p):
            messagebox.showerror("错误", f"文件不存在: {p}")
            return
        # PDF和图片都在程序内查看（PDF自动转图片）
        ImageViewer(self, p, row, invoice_id=row['id'])

    def export_excel(self):
        """导出发票报销明细到Excel（含发票图片）"""
        conn = get_db()
        sql = """SELECT i.*, e.name as rname FROM invoices i
                 LEFT JOIN employees e ON i.reimburser_id=e.id WHERE 1=1"""
        params = []
        if self.filter_status.get() == '未报销':
            sql += " AND i.status=0"
        elif self.filter_status.get() == '已报销':
            sql += " AND i.status=1"
        if self.filter_person.get() and self.filter_person.get() != '全部':
            sql += " AND e.name=?"
            params.append(self.filter_person.get())
        sql += " ORDER BY i.invoice_date DESC, i.id DESC"
        rows = conn.execute(sql, params).fetchall()
        conn.close()

        if not rows:
            messagebox.showinfo("提示", "没有可导出的数据")
            return

        default_name = f"发票报销明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出Excel", defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "发票报销明细"

        NCOLS = 13  # 编号,类型,发票号,日期,报销人,金额,用途,销售方,备注,凭证号,状态,创建时间,发票图片
        # 标题
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
        ws['A1'] = "发票报销明细表"
        ws['A1'].font = TITLE_FONT
        ws['A1'].alignment = CENTER
        ws.row_dimensions[1].height = 28

        filter_info = f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}    状态筛选：{self.filter_status.get()}    报销人：{self.filter_person.get()}"
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
        ws['A2'] = filter_info
        ws['A2'].font = Font(name='微软雅黑', size=9, color='666666')
        ws['A2'].alignment = LEFT

        headers = ['编号', '类型', '发票号', '发票日期', '报销人', '金额(元)', '用途', '销售方', '备注', '凭证号', '状态', '创建时间', '发票图片']
        for c, h in enumerate(headers, 1):
            ws.cell(row=3, column=c, value=h)
        _style_header(ws, 3, NCOLS)

        IMG_W = 140  # 图片显示宽度（像素）
        ROW_H = 110
        total_amount = 0
        unpaid_amount = 0
        img_count = 0
        _img_refs = []  # 保持图片对象引用，防止被GC回收

        for idx, r in enumerate(rows, 4):
            status = '已报销' if r['status'] == 1 else '未报销'
            ws.cell(row=idx, column=1, value=r['id']).alignment = CENTER
            ws.cell(row=idx, column=2, value=r['type'] or '发票').alignment = CENTER
            ws.cell(row=idx, column=3, value=r['invoice_number'] or '').alignment = CENTER
            ws.cell(row=idx, column=4, value=r['invoice_date'] or '').alignment = CENTER
            ws.cell(row=idx, column=5, value=r['rname'] or '(已删除)').alignment = CENTER
            cell_amt = ws.cell(row=idx, column=6, value=float(r['amount']))
            cell_amt.number_format = '#,##0.00'
            cell_amt.alignment = RIGHT
            ws.cell(row=idx, column=7, value=r['purpose'] or '').alignment = LEFT
            ws.cell(row=idx, column=8, value=r['seller'] or '').alignment = LEFT
            ws.cell(row=idx, column=9, value=r['remark'] or '').alignment = LEFT
            ws.cell(row=idx, column=10, value=r['voucher_number'] or '').alignment = CENTER
            ws.cell(row=idx, column=11, value=status).alignment = CENTER
            ws.cell(row=idx, column=12, value=r['created_at'] or '').alignment = CENTER
            for c in range(1, NCOLS + 1):
                ws.cell(row=idx, column=c).border = THIN_BORDER

            # 插入发票图片（PDF自动转第一页图片）
            if r['image_path']:
                img_path = os.path.join(INVOICE_DIR, r['image_path'])
                if os.path.exists(img_path):
                    try:
                        if img_path.lower().endswith('.pdf'):
                            pil_img = pdf_to_image(img_path)
                        else:
                            pil_img = Image.open(img_path)
                        if pil_img:
                            w, h = pil_img.size
                            ratio = IMG_W / w
                            new_h = int(h * ratio)
                            resized = pil_img.resize((IMG_W, new_h), Image.LANCZOS)
                            if resized.mode not in ('RGB', 'RGBA'):
                                resized = resized.convert('RGB')
                            buf = BytesIO()
                            resized.save(buf, format='PNG')
                            buf.seek(0)
                            xl_img = XLImage(buf)
                            xl_img.width = IMG_W
                            xl_img.height = new_h
                            _img_refs.append(buf)
                            cell_ref = f"{get_column_letter(NCOLS)}{idx}"
                            ws.add_image(xl_img, cell_ref)
                            ws.row_dimensions[idx].height = max(ROW_H, new_h * 0.8)
                            img_count += 1
                            ws.cell(row=idx, column=NCOLS, value='').border = THIN_BORDER
                        else:
                            ws.cell(row=idx, column=NCOLS, value='图片加载失败').alignment = CENTER
                    except Exception as ex:
                        ws.cell(row=idx, column=NCOLS, value='图片加载失败').alignment = CENTER
                else:
                    ws.cell(row=idx, column=NCOLS, value='文件缺失').alignment = CENTER
            else:
                ws.cell(row=idx, column=NCOLS, value='无').alignment = CENTER

            total_amount += float(r['amount'])
            if r['status'] == 0:
                unpaid_amount += float(r['amount'])

        # 合计行
        total_row = len(rows) + 4
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=5)
        ws.cell(row=total_row, column=1, value=f"合计（共{len(rows)}条，含附件{img_count}个）").font = TOTAL_FONT
        ws.cell(row=total_row, column=1).alignment = RIGHT
        ws.cell(row=total_row, column=1).fill = TOTAL_FILL
        cell_total = ws.cell(row=total_row, column=6, value=total_amount)
        cell_total.number_format = '#,##0.00'
        cell_total.font = TOTAL_FONT
        cell_total.fill = TOTAL_FILL
        cell_total.alignment = RIGHT
        ws.merge_cells(start_row=total_row, start_column=7, end_row=total_row, end_column=NCOLS)
        ws.cell(row=total_row, column=7, value=f"其中未报销：¥{unpaid_amount:,.2f}").font = TOTAL_FONT
        ws.cell(row=total_row, column=7).fill = TOTAL_FILL
        ws.cell(row=total_row, column=7).alignment = LEFT
        for c in range(1, NCOLS + 1):
            ws.cell(row=total_row, column=c).border = THIN_BORDER

        # 列宽
        widths = [7, 7, 15, 11, 9, 11, 20, 25, 20, 13, 8, 17, 22]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A4'

        try:
            wb.save(path)
            self.app.set_status(f"已导出Excel: {os.path.basename(path)}")
            messagebox.showinfo("导出成功", f"已导出 {len(rows)} 条发票记录（含 {img_count} 张图片）到：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def import_excel(self):
        """从Excel导入发票报销数据"""
        path = filedialog.askopenfilename(title="选择要导入的Excel文件",
                                          filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取Excel文件：{e}")
            return

        # 自动识别表头行（查找包含"发票号"或"报销人"的行）
        header_row = None
        col_map = {}
        for row_idx in range(1, min(ws.max_row, 10) + 1):
            row_data = [str(ws.cell(row=row_idx, column=c).value or '').strip() for c in range(1, ws.max_column + 1)]
            if '发票号' in row_data or '报销人' in row_data or '金额' in row_data:
                header_row = row_idx
                for c_idx, val in enumerate(row_data, 1):
                    col_map[val] = c_idx
                break

        if header_row is None:
            messagebox.showerror("导入失败", "未识别到表头行，请确保Excel包含「发票号」「报销人」「金额」等列名。")
            return

        # 列名映射（兼容多种写法）
        def get_col(*names):
            for n in names:
                if n in col_map:
                    return col_map[n]
            return None

        col_no = get_col('发票号', '发票号码', '编号')
        col_type = get_col('类型', '票据类型')
        col_date = get_col('发票日期', '日期', '开票日期')
        col_person = get_col('报销人', '报销人员', '员工')
        col_amount = get_col('金额', '金额(元)', '报销金额')
        col_purpose = get_col('用途', '摘要', '事由', '费用项目')
        col_seller = get_col('销售方', '销售方名称', '商户', '收款方')
        col_remark = get_col('备注', '附言', '说明')
        col_voucher = get_col('凭证号', '凭证编号')
        col_status = get_col('状态', '报销状态')

        if col_amount is None or col_person is None:
            messagebox.showerror("导入失败", "Excel必须包含「报销人」和「金额」列。")
            return

        conn = get_db()
        success = 0
        failed = 0
        errors = []

        for row_idx in range(header_row + 1, ws.max_row + 1):
            def cell_val(col):
                if col is None:
                    return ''
                v = ws.cell(row=row_idx, column=col).value
                return str(v).strip() if v is not None else ''

            person_name = cell_val(col_person)
            amount_str = cell_val(col_amount)

            # 跳过空行
            if not person_name and not amount_str:
                continue

            try:
                amount = float(amount_str.replace(',', '').replace('¥', '')) if amount_str else 0
            except ValueError:
                failed += 1
                errors.append(f"第{row_idx}行：金额格式错误「{amount_str}」")
                continue

            if amount <= 0:
                failed += 1
                errors.append(f"第{row_idx}行：金额必须大于0")
                continue

            # 查找报销人
            emp = conn.execute("SELECT id FROM employees WHERE name=?", (person_name,)).fetchone()
            if not emp:
                failed += 1
                errors.append(f"第{row_idx}行：报销人「{person_name}」不存在，请先在人员管理中添加")
                continue
            emp_id = emp['id']

            inv_no = cell_val(col_no)
            inv_type = cell_val(col_type) or '发票'
            if inv_type not in ('发票', '支付记录'):
                inv_type = '发票'
            inv_date = cell_val(col_date) or datetime.now().strftime('%Y-%m-%d')
            purpose = cell_val(col_purpose)
            seller = cell_val(col_seller)
            remark = cell_val(col_remark)
            voucher = cell_val(col_voucher)
            status_str = cell_val(col_status)
            status = 1 if ('已报销' in status_str or status_str == '1') else 0

            # 发票号重复校验
            if inv_no:
                exist = conn.execute("SELECT id FROM invoices WHERE invoice_number=?", (inv_no,)).fetchone()
                if exist:
                    failed += 1
                    errors.append(f"第{row_idx}行：发票号「{inv_no}」已存在，跳过")
                    continue

            try:
                conn.execute("""INSERT INTO invoices(invoice_number, type, invoice_date, reimburser_id,
                              amount, purpose, seller, remark, voucher_number, status, created_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                             (inv_no, inv_type, inv_date, emp_id, amount, purpose, seller, remark, voucher, status,
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                success += 1
            except Exception as e:
                failed += 1
                errors.append(f"第{row_idx}行：导入失败 - {e}")

        conn.commit()
        conn.close()

        self.refresh()
        self.app.stat_tab.refresh()

        msg = f"导入完成！\n\n成功：{success} 条\n失败：{failed} 条"
        if errors:
            msg += "\n\n失败详情：\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                msg += f"\n...（共{len(errors)}条错误）"
        messagebox.showinfo("导入结果", msg)
        self.app.set_status(f"发票导入完成：成功{success}条，失败{failed}条")

    def print_layout(self):
        """勾选的发票进行打印排版"""
        if not self.checked_ids:
            messagebox.showwarning("提示", "请先勾选要打印排版的发票（点击选择列的✓）")
            return
        conn = get_db()
        placeholders = ','.join('?' * len(self.checked_ids))
        all_rows = conn.execute(
            "SELECT id, invoice_number, invoice_date, amount, purpose, image_path, printed "
            "FROM invoices WHERE id IN (%s)" % placeholders, tuple(self.checked_ids)).fetchall()
        conn.close()
        rows = [r for r in all_rows if r['image_path']]
        if not rows:
            messagebox.showinfo("提示", "勾选的发票中没有带附件的，无法排版打印")
            return
        printed_count = sum(1 for r in rows if r['printed'] == 1)
        if printed_count > 0:
            if not messagebox.askyesno("提示",
                f"勾选的 {len(rows)} 张发票中有 {printed_count} 张已经排过版了，是否继续排版？"):
                return
        skipped = len(all_rows) - len(rows)
        if skipped > 0:
            self.app.set_status(f"打印排版：{skipped} 张勾选发票无附件已跳过，排版 {len(rows)} 张")
        self.app.set_status(f"正在打开打印排版编辑器（{len(rows)} 张发票）...")
        try:
            PrintLayoutEditor(self, self.app, [dict(r) for r in rows])
        except Exception as e:
            _log_ocr_error(f"打开打印排版编辑器失败: {e}")
            messagebox.showerror("错误", f"打开打印排版编辑器失败: {e}")


# ============================================================
# 打印排版编辑器（发票自动排版 + 人工检查 + 导出图片）
# ============================================================
class PrintLayoutEditor(tk.Toplevel):
    """打印排版编辑器：
    1. 自动排版：大发票一页2张、中等发票一页6张、小发票自由排列填充空白
    2. 人工检查：点击选中、拖拽移动、边中点单边拉伸、四角双向缩放、旋转90°、移除
    3. 纸张：A4 / A3 / 自定义(cm)，自动分页 + 手动添加页面
    4. 红线标出打印边界，发票不超出纸张范围
    5. 导出：渲染为图片（PNG），询问保存文件夹
    """
    PAPER_SIZES = {'A4': (210, 297), 'A3': (297, 420)}

    def __init__(self, parent, app, invoice_rows):
        super().__init__(parent)
        self.parent = parent
        self.app = app
        self.title("打印排版 - 人工检查")
        self.configure(bg='#1a1a1e')
        self.geometry('1100x760')
        self.minsize(820, 580)
        self.transient(parent)
        self.invoice_rows = invoice_rows
        self.items = []      # [{path, num, aspect, ow, oh}]
        self.pages = []      # [[ {idx,x,y,w,h,rotation}, ... ], ...]
        self.paper_w, self.paper_h = 210, 297
        self.dpi = 300
        self.current_page = 0
        self.selected = None  # 当前页选中的layout项索引
        self._drag = None     # 拖拽起点
        self._drag_mode = 'move'  # move/corner/edge
        self._handle_idx = 0
        self._drag_orig = None    # 拖拽起始时选中的布局状态（用于等比缩放）
        self._thumbs = []    # 保持PhotoImage引用

        # 加载发票图片
        self._load_items()
        if not self.items:
            messagebox.showwarning("提示", "所选发票附件无法读取图片")
            self.destroy()
            return
        # 自动排版
        self._auto_layout()
        # 构建界面
        self._build_ui()
        # 绘制
        self.after(80, self._draw)

    # ---------- 数据加载 ----------
    def _load_items(self):
        for r in self.invoice_rows:
            img = get_invoice_image(r['image_path'])
            if not img:
                continue
            if img.mode != 'RGB':
                img = img.convert('RGB')
            w, h = img.size
            self.items.append({
                'id': r['id'],
                'path': os.path.join(INVOICE_DIR, r['image_path']),
                'num': r['invoice_number'] or '',
                'aspect': (w / h) if h > 0 else 1.0,
                'ow': w, 'oh': h,
            })

    # ---------- 自动排版 ----------
    def _auto_layout(self):
        # 用"短边+宽高比"分类（更抗扫描分辨率影响）：
        # - 大票：横版(宽>=0.55高)且物理尺寸大(短边>=1100px)，如增值税电子发票/横版普通发票 → 一页2张横放清晰
        # - 竖版中票：竖版(宽<0.55高)且较大(短边>=900px)，如福建增值税普通发票(竖版) → 一页2张竖放，保持可读尺寸
        # - 中票：其余较宽发票（定额发票/支付记录/出租票等）→ 一页6张(2行×3列)紧凑
        # - 小票：极小的窄长票 → 独立排页
        n = len(self.items)
        if n == 0:
            self.pages = [[]]
            return
        shorts = [min(self.items[i]['ow'], self.items[i]['oh']) for i in range(n)]
        aspects = [self.items[i]['aspect'] for i in range(n)]  # ow/oh
        large_g = [i for i in range(n) if aspects[i] >= 0.55 and shorts[i] >= 800]
        large_set = set(large_g)
        rest1 = [i for i in range(n) if i not in large_set]
        # 竖版中票（增值税普通发票等竖版较大发票）：一页2张竖放
        mid_large_g = [i for i in rest1 if aspects[i] < 0.55 and shorts[i] >= 900]
        ml_set = set(mid_large_g)
        rest = [i for i in range(n) if i not in large_set and i not in ml_set]
        if rest:
            rshorts = [shorts[i] for i in rest]
            span_r = max(rshorts) / max(1, min(rshorts))
            if len(rest) > 1 and span_r >= 1.5:
                # 剩余仍有显著尺寸差异：按相对比例分中/小
                mx = max(rshorts)
                middle_g = [i for i in rest if shorts[i] >= mx * 0.4]
                small_g = [i for i in rest if shorts[i] < mx * 0.4]
            else:
                # 尺寸接近：按绝对值分档
                med_short = sorted(rshorts)[len(rshorts) // 2]
                if med_short >= 900:
                    middle_g, small_g = rest, []
                elif med_short >= 450:
                    middle_g = [i for i in rest if shorts[i] >= 450]
                    small_g = [i for i in rest if shorts[i] < 450]
                else:
                    middle_g, small_g = [], rest
        else:
            middle_g, small_g = [], []
        large_g.sort(key=lambda i: shorts[i], reverse=True)
        mid_large_g.sort(key=lambda i: shorts[i], reverse=True)
        middle_g.sort(key=lambda i: shorts[i], reverse=True)
        small_g.sort(key=lambda i: shorts[i], reverse=True)
        h_ratio = self.paper_h / self.paper_w
        h_limit = h_ratio * 0.94
        self.pages = []
        # 阶段1：大发票每页2张（上下排列，占满页宽）
        if large_g:
            self.pages.extend(self._layout_fixed(large_g, 2, 0.47, h_limit, h_ratio))
        # 阶段1b：竖版中票每页2张竖放（保持可读尺寸）
        if mid_large_g:
            self.pages.extend(self._layout_medium2(mid_large_g, h_ratio))
        # 阶段2：中等发票每页6张（2行×3列填满整页）
        if middle_g:            self.pages.extend(self._layout_fixed(middle_g, 6, 0.30, h_limit, h_ratio))
        # 阶段3：小发票独立排页（每张约半页高≈148mm，保持可读尺寸，放不下自动分页）
        if small_g:
            self.pages.extend(self._layout_small(small_g, h_ratio))
        # 清理空页
        self.pages = [pg for pg in self.pages if pg]
        if not self.pages:
            self.pages = [[]]

    def _layout_medium2(self, indices, page_h_ratio):
        """竖版中等发票（如福建增值税普通发票竖版）：一页2张上下竖放，每张约半页高，保持可读尺寸不缩小。
        竖放不旋转（竖版票竖放更自然），下方留少量空白可供手动填充小票。"""
        pages = []
        GAP = 0.014
        half_h = (page_h_ratio - 3 * GAP) / 2.0
        for start in range(0, len(indices), 2):
            group = indices[start:start + 2]
            page = []
            y = GAP
            for idx in group:
                item = self.items[idx]
                aspect = item['aspect']  # 竖版 ow/oh < 0.55
                h = half_h
                w = h * aspect
                if w > 0.94:
                    w = 0.94
                    h = w / aspect
                page.append({'idx': idx, 'x': (1.0 - w) / 2, 'y': y,
                             'w': w, 'h': h, 'rotation': 0})
                y += h + GAP
            pages.append(page)
        return pages

    def _layout_fixed(self, indices, per_page, tw, h_limit, page_h_ratio):
        """固定张数自动排版：
        - 大票(per_page=2)：每页2张，上下排列，每张占满页宽（约半页高），清晰不浪费
        - 中票(per_page=6)：每页6张，2行×3列，每张约半页高，填满整页
        自动分页。"""
        pages = []
        GAP = 0.014
        if per_page == 2:
            # ---- 大票：每页2张，上下排列 ----
            half_h = (page_h_ratio - 3 * GAP) / 2.0
            for start in range(0, len(indices), 2):
                group = indices[start:start + 2]
                page = []
                y = GAP
                for idx in group:
                    item = self.items[idx]
                    rotation = 0
                    aspect = item['aspect']
                    if aspect < 0.7:
                        rotation = 90
                        aspect = 1.0 / aspect
                    w = 0.94
                    h = w / aspect
                    if h > half_h:
                        h = half_h
                        w = h * aspect
                    page.append({'idx': idx, 'x': (1.0 - w) / 2, 'y': y,
                                 'w': w, 'h': h, 'rotation': rotation})
                    y += h + GAP
                pages.append(page)
            return pages
        # ---- 中票：每页6张，2行×3列竖排（不旋转），每张约半页高，填满整页 ----
        # 数量不足6张时动态调整：<=2张每页2张上下排列，3-4张每行2张，5-6张2行×3列
        for start in range(0, len(indices), 6):
            group = indices[start:start + 6]
            cnt = len(group)
            # 根据数量决定行列数
            if cnt <= 2:
                cols, rows = 1, 2
            elif cnt <= 4:
                cols, rows = 2, 2
            else:
                cols, rows = 3, 2
            maxw = (1.0 - (cols + 1) * GAP) / cols
            rowh = (page_h_ratio - (rows + 1) * GAP) / rows
            info = []
            for idx in group:
                item = self.items[idx]
                rotation = 0
                aspect = item['aspect']
                if aspect < 0.35:
                    rotation = 90
                    aspect = 1.0 / aspect
                h = rowh
                w = h * aspect
                if w > maxw:
                    w = maxw
                    h = w / aspect
                info.append((idx, w, h, rotation))
            page = []
            for row in range(rows):
                y = GAP + row * (rowh + GAP)
                for col in range(cols):
                    n = row * cols + col
                    if n >= len(info):
                        break
                    idx, w, h, rotation = info[n]
                    x = GAP + col * (maxw + GAP) + (maxw - w) / 2
                    page.append({'idx': idx, 'x': x, 'y': y,
                                 'w': w, 'h': h, 'rotation': rotation})
            pages.append(page)
        return pages

    def _layout_small(self, indices, page_h_ratio):
        """小发票独立排页：每张高度≈A4长度一半（约148mm），宽度等比例缩放。
        每行最多放3张，放不下换行，一页最多2行，满2行自动开新页（不缩小）。
        发票之间保留明显空隙（GAP），便于裁剪。"""
        pages = []
        GAP = 0.03
        half_h = (page_h_ratio - 3 * GAP) / 2.0
        info = []
        for idx in indices:
            item = self.items[idx]
            rotation = 0
            aspect = item['aspect']
            # 极窄长条(宽<0.2高)才旋转横放，避免竖版窄长票被旋转成占页宽的大块浪费版面
            if aspect < 0.20:
                rotation = 90
                aspect = 1.0 / aspect
            h = half_h
            w = h * aspect
            info.append({'idx': idx, 'w': w, 'h': h, 'rotation': rotation})
        page = []
        cur_row = []
        x = GAP
        y = GAP
        row_no = 1
        for it in info:
            if x + it['w'] > 1.0 - GAP and cur_row:
                # 换行（当前行放不下）
                page.extend(cur_row)
                cur_row = []
                row_no += 1
                if row_no > 2:
                    # 一页最多2行，开新页
                    pages.append(page)
                    page = []
                    row_no = 1
                x = GAP
                y = GAP + (row_no - 1) * (half_h + GAP)
            if x + it['w'] > 1.0 - GAP:
                # 单张也放不下（极宽）：等比缩小到页宽（保持宽高比）
                scale = (1.0 - 2 * GAP) / it['w']
                it['w'] *= scale
                it['h'] *= scale
            it['x'] = x
            it['y'] = y
            cur_row.append(it)
            x += it['w'] + GAP
        page.extend(cur_row)
        pages.append(page)
        return pages

    def _shelf_layout(self, indices, tw, h_limit, page_h_ratio):
        """货架算法：统一宽度，逐行放置，放不下换页。返回页面列表。
        很窄长的发票（宽高比<0.6）自动旋转90°横放，节省版面。"""
        pages = []
        cur = []
        x = 0.012
        y = 0.012
        row_h = 0
        GAP = 0.014
        for idx in indices:
            item = self.items[idx]
            rotation = 0
            aspect = item['aspect']
            if aspect < 0.6:
                rotation = 90
                aspect = 1.0 / aspect
            w = tw
            h = w / aspect
            if h > h_limit:
                h = h_limit
                w = h * aspect
            if x + w > 1.0 - GAP:
                x = GAP
                y += row_h + GAP
                row_h = 0
            if y + h > page_h_ratio - GAP:
                pages.append(cur)
                cur = []
                x = GAP
                y = GAP
                row_h = 0
            cur.append({'idx': idx, 'x': x, 'y': y, 'w': w, 'h': h, 'rotation': rotation})
            x += w + GAP
            row_h = max(row_h, h)
        if cur:
            pages.append(cur)
        return pages

    def _fill_blank(self, page, small_indices, page_h_ratio):
        """网格法在页面空白区域填充小发票，保持宽高比不变形。
        占用标记用ceil取整确保不重叠，放置用浮点坐标保持比例。"""
        GRID_COLS, GRID_ROWS = 20, 28
        occupied = [[False] * GRID_COLS for _ in range(GRID_ROWS)]
        for lo in page:
            x0 = max(0, min(GRID_COLS - 1, int(lo['x'] * GRID_COLS)))
            x1 = max(0, min(GRID_COLS, int(math.ceil((lo['x'] + lo['w']) * GRID_COLS))))
            y_top = lo['y'] / page_h_ratio
            y_bot = (lo['y'] + lo['h']) / page_h_ratio
            y0 = max(0, min(GRID_ROWS - 1, int(y_top * GRID_ROWS)))
            y1 = max(0, min(GRID_ROWS, int(math.ceil(y_bot * GRID_ROWS))))
            for gy in range(y0, y1):
                for gx in range(x0, x1):
                    occupied[gy][gx] = True
        filled = []
        for idx in small_indices:
            item = self.items[idx]
            rotation = 0
            aspect = item['aspect']
            if aspect < 0.6:
                rotation = 90
                aspect = 1.0 / aspect
            w = 0.20
            h = w / aspect
            if h > page_h_ratio - 0.02:
                h = page_h_ratio - 0.02
                w = h * aspect
            gw = max(1, int(math.ceil(w * GRID_COLS)))
            gh = max(1, int(math.ceil((h / page_h_ratio) * GRID_ROWS)))
            placed = False
            for gy in range(GRID_ROWS - gh + 1):
                for gx in range(GRID_COLS - gw + 1):
                    if all(not occupied[gy + dy][gx + dx] for dy in range(gh) for dx in range(gw)):
                        x = gx / GRID_COLS + 0.004
                        y = (gy / GRID_ROWS) * page_h_ratio + 0.004
                        lo = {'idx': idx, 'x': x, 'y': y, 'w': w, 'h': h, 'rotation': rotation}
                        page.append(lo)
                        # 标记占用（ceil覆盖实际区域）
                        x0m = int(x * GRID_COLS)
                        x1m = int(math.ceil((x + w) * GRID_COLS))
                        y0m = int(y / page_h_ratio * GRID_ROWS)
                        y1m = int(math.ceil((y + h) / page_h_ratio * GRID_ROWS))
                        for dy in range(y0m, min(y1m, GRID_ROWS)):
                            for dx in range(x0m, min(x1m, GRID_COLS)):
                                occupied[dy][dx] = True
                        placed = True
                        break
                if placed:
                    break
            if placed:
                filled.append(idx)
        return filled

    # ---------- UI ----------
    def _build_ui(self):
        # 顶部工具栏
        top = ttk.Frame(self, padding=8)
        top.pack(fill='x', side='top')
        ttk.Label(top, text="纸张:").pack(side='left', padx=4)
        self.paper_var = tk.StringVar(value='A4')
        paper_cb = ttk.Combobox(top, textvariable=self.paper_var, values=['A4', 'A3', '自定义'],
                                width=6, state='readonly')
        paper_cb.pack(side='left', padx=4)
        paper_cb.bind('<<ComboboxSelected>>', self._on_paper_change)
        self.cust_w_var = tk.StringVar(value='21')
        self.cust_h_var = tk.StringVar(value='29.7')
        self.cust_w_entry = ttk.Entry(top, textvariable=self.cust_w_var, width=5)
        self.cust_h_entry = ttk.Entry(top, textvariable=self.cust_h_var, width=5)
        ttk.Label(top, text=" 宽(cm):").pack(side='left', padx=(6, 0))
        self.cust_w_entry.pack(side='left', padx=2)
        ttk.Label(top, text="高(cm):").pack(side='left', padx=2)
        self.cust_h_entry.pack(side='left', padx=2)
        ttk.Button(top, text="应用纸张", command=self._apply_paper).pack(side='left', padx=6)
        ttk.Separator(top, orient='vertical').pack(side='left', fill='y', padx=8)
        ttk.Button(top, text="↻ 旋转选中90°", command=self._rotate_selected).pack(side='left', padx=4)
        ttk.Button(top, text="移除选中", command=self._remove_selected).pack(side='left', padx=4)
        ttk.Button(top, text="重新自动排版", command=self._reauto_layout).pack(side='left', padx=4)
        ttk.Button(top, text="＋ 添加页面", command=self._add_page).pack(side='left', padx=4)
        ttk.Separator(top, orient='vertical').pack(side='left', fill='y', padx=8)
        self.export_btn = ttk.Button(top, text="导出图片...", command=self._export,
                                     style='Accent.TButton')
        self.export_btn.pack(side='right', padx=4)
        ttk.Label(top, text=f"共 {len(self.invoice_rows)} 张发票", foreground='#888').pack(side='right', padx=8)

        # 提示栏
        tip = ttk.Label(self,
                        text="点击选中发票 → 拖拽移动；选中后拖动边框边中点=单边拉伸，拖动四角=双向缩放；可旋转/移除；红线内为打印范围。",
                        foreground='#999', padding=(12, 0))
        tip.pack(fill='x')

        # 画布
        self.canvas = tk.Canvas(self, bg='#2a2a2f', highlightthickness=0)
        self.canvas.pack(fill='both', expand=True, padx=8, pady=4)
        self.canvas.bind('<Button-1>', self._on_click)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)

        # 底栏（页码）
        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill='x', side='bottom')
        ttk.Button(bottom, text="◀ 上一页", command=lambda: self._goto_page(self.current_page - 1)).pack(side='left', padx=4)
        self.page_label = ttk.Label(bottom, text="", width=24, anchor='center')
        self.page_label.pack(side='left', padx=4)
        ttk.Button(bottom, text="下一页 ▶", command=lambda: self._goto_page(self.current_page + 1)).pack(side='left', padx=4)
        ttk.Separator(bottom, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Button(bottom, text="⬅ 移到上一页", command=lambda: self._move_to_page(-1)).pack(side='left', padx=4)
        ttk.Button(bottom, text="移到下一页 ➡", command=lambda: self._move_to_page(1)).pack(side='left', padx=4)
        ttk.Button(bottom, text="📄 移动到指定页...", command=self._move_to_page_picker).pack(side='left', padx=4)
        self.sel_label = ttk.Label(bottom, text="", foreground='#3B82F6')
        self.sel_label.pack(side='right', padx=4)

        self.canvas.bind('<Configure>', lambda e: self._draw())

    # ---------- 交互 ----------
    def _on_click(self, event):
        ox, oy, pw, ph = self._get_page_rect()
        if pw <= 0:
            return
        page = self.pages[self.current_page]
        hr = self.paper_h / self.paper_w
        # 1) 若已有选中发票，先检测其控制点
        if self.selected is not None and self.selected < len(page):
            lo = page[self.selected]
            handles = self._get_handles(lo, ox, oy, pw, ph)
            for hi, (hx, hy) in enumerate(handles):
                if abs(event.x - hx) <= 9 and abs(event.y - hy) <= 9:
                    self._drag_mode = 'corner' if hi < 4 else 'edge'
                    self._handle_idx = hi
                    self._drag = (event.x, event.y)
                    self._drag_orig = dict(lo)
                    return
        # 2) 检测发票内部（从后往前，优先上层）
        self.selected = None
        for li in range(len(page) - 1, -1, -1):
            lo = page[li]
            px = ox + lo['x'] * pw
            py = oy + lo['y'] * pw
            if px <= event.x <= px + lo['w'] * pw and py <= event.y <= py + lo['h'] * pw:
                self.selected = li
                self._drag_mode = 'move'
                self._drag = (event.x, event.y)
                self._drag_orig = dict(lo)
                break
        if self.selected is None:
            self._drag = None
        self._update_sel_label()
        self._draw()

    def _on_drag(self, event):
        if self.selected is None or self._drag is None:
            return
        page = self.pages[self.current_page]
        if self.selected >= len(page):
            return
        lo = page[self.selected]
        ox, oy, pw, ph = self._get_page_rect()
        if pw <= 0:
            return
        hr = self.paper_h / self.paper_w
        dx = (event.x - self._drag[0]) / pw
        dy = (event.y - self._drag[1]) / pw
        mode = self._drag_mode
        if mode == 'move':
            lo['x'] = max(0.0, min(1.0 - lo['w'], lo['x'] + dx))
            lo['y'] = max(0.0, min(hr - lo['h'], lo['y'] + dy))
        elif mode == 'edge':
            self._resize_edge(lo, dx, dy, hr)
        elif mode == 'corner':
            self._resize_corner(lo, event, ox, oy, pw, ph, hr)
        self._drag = (event.x, event.y)
        self._draw()

    def _on_release(self, event):
        self._drag = None
        self._drag_mode = 'move'

    def _get_handles(self, lo, ox, oy, pw, ph):
        """返回选中发票的8个控制点屏幕坐标：前4个为四角，后4个为四边中点"""
        px = ox + lo['x'] * pw
        py = oy + lo['y'] * pw
        w2 = lo['w'] * pw
        h2 = lo['h'] * pw
        x1, y1 = px + w2, py + h2
        return [(px, py), (x1, py), (px, y1), (x1, y1),     # 左上 右上 左下 右下
                (px + w2 / 2, py), (px, py + h2 / 2),       # 上中 左中
                (x1, py + h2 / 2), (px + w2 / 2, y1)]       # 右中 下中

    def _resize_edge(self, lo, dx, dy, hr):
        """边中点单边拉伸：对边固定，只移动该边"""
        x0, y0 = lo['x'], lo['y']
        x1, y1 = lo['x'] + lo['w'], lo['y'] + lo['h']
        idx = self._handle_idx
        if idx == 4:    # 上中点 → 调整 y0
            y0 = max(0.0, min(y0 + dy, y1 - 0.02))
        elif idx == 5:  # 左中点 → 调整 x0
            x0 = max(0.0, min(x0 + dx, x1 - 0.02))
        elif idx == 6:  # 右中点 → 调整 x1
            x1 = min(1.0, max(x1 + dx, x0 + 0.02))
        elif idx == 7:  # 下中点 → 调整 y1
            y1 = min(hr, max(y1 + dy, y0 + 0.02))
        lo['x'], lo['y'] = x0, y0
        lo['w'], lo['h'] = x1 - x0, y1 - y0

    def _resize_corner(self, lo, event, ox, oy, pw, ph, hr):
        """四角等比缩放：对角固定，保持宽高比不变（不会变形）"""
        orig = self._drag_orig
        ow = orig['w']
        oh = orig['h']
        aspect = ow / oh if oh > 0 else 1.0
        ax0 = orig['x']
        ay0 = orig['y']
        idx = self._handle_idx
        # 锚点 = 拖拽角的对角
        if idx == 0:      # 左上 → 锚右下
            ax, ay = ax0 + ow, ay0 + oh
        elif idx == 1:    # 右上 → 锚左下
            ax, ay = ax0, ay0 + oh
        elif idx == 2:    # 左下 → 锚右上
            ax, ay = ax0 + ow, ay0
        else:             # 右下 → 锚左上
            ax, ay = ax0, ay0
        # 鼠标相对锚点的距离（页宽单位）
        dw = (event.x - (ox + ax * pw)) / pw
        dh = (event.y - (oy + ay * pw)) / pw
        # 保持宽高比：取两个距离中较小者折算
        nw = max(0.02, min(abs(dw), abs(dh) * aspect))
        # 限制尺寸不超出页面（按锚点到边缘的可用空间）
        if idx in (0, 2):
            avail_w = ax
        else:
            avail_w = 1.0 - ax
        if idx in (0, 1):
            avail_h = ay
        else:
            avail_h = hr - ay
        nw = min(nw, avail_w, avail_h * aspect)
        nh = nw / aspect
        # 新左上角（锚点固定）
        if idx in (0, 2):
            x = ax - nw
        else:
            x = ax
        if idx in (0, 1):
            y = ay - nh
        else:
            y = ay
        lo['x'], lo['y'] = x, y
        lo['w'], lo['h'] = nw, nh

    def _rotate_selected(self):
        if self.selected is None:
            messagebox.showinfo("提示", "请先在画布上点击选中一张发票")
            return
        page = self.pages[self.current_page]
        lo = page[self.selected]
        lo['rotation'] = (lo['rotation'] + 90) % 360
        lo['w'], lo['h'] = lo['h'], lo['w']
        self._draw()

    def _remove_selected(self):
        if self.selected is None:
            messagebox.showinfo("提示", "请先在画布上点击选中一张发票")
            return
        page = self.pages[self.current_page]
        del page[self.selected]
        self.selected = None
        if not page:
            self.pages.pop(self.current_page)
            if self.current_page >= len(self.pages):
                self.current_page = len(self.pages) - 1
            if not self.pages:
                self.pages = [[]]
                self.current_page = 0
        self._update_sel_label()
        self._draw()

    def _add_page(self):
        """手动添加一页空白页"""
        self.pages.append([])
        self.current_page = len(self.pages) - 1
        self.selected = None
        self._update_sel_label()
        self._draw()

    def _goto_page(self, p):
        if 0 <= p < len(self.pages):
            self.current_page = p
            self.selected = None
            self._update_sel_label()
            self._draw()

    def _move_to_page(self, delta):
        """把选中的发票移动到上一页/下一页（填补空白，避免浪费页面）"""
        if self.selected is None:
            messagebox.showinfo("提示", "请先在画布上点击选中一张发票")
            return
        target = self.current_page + delta
        if target < 0:
            messagebox.showinfo("提示", "已经是第一页，无法上移")
            return
        self._move_selected_to(target)

    def _move_selected_to(self, target):
        """把选中的发票移动到指定页面（target == len(pages) 表示新建空白页）"""
        if self.selected is None:
            messagebox.showinfo("提示", "请先在画布上点击选中一张发票")
            return
        page = self.pages[self.current_page]
        lo = page.pop(self.selected)
        self.selected = None
        if target >= len(self.pages):
            self.pages.append([])
        target_page = self.pages[target]
        # 放置到目标页的空白位置：优先放在已有发票下方的空白行，否则居中放页面底部空位
        hr = self.paper_h / self.paper_w
        lo = dict(lo)
        lo['x'] = max(0.012, min(1.0 - lo['w'] - 0.012, (1.0 - lo['w']) / 2.0))
        lo['y'] = self._find_free_y(target_page, lo['h'], hr)
        target_page.append(lo)
        # 当前页空了则移除
        if not page:
            self.pages.pop(self.current_page)
        self.current_page = min(target, len(self.pages) - 1)
        self._update_sel_label()
        self._draw()

    def _move_to_page_picker(self):
        """弹窗选择目标页（任意已有页或新建页），把选中发票移动过去"""
        if self.selected is None:
            messagebox.showinfo("提示", "请先在画布上点击选中一张发票")
            return
        import tkinter as tk
        win = tk.Toplevel(self)
        win.title("移动到指定页面")
        win.geometry("360x320")
        bg = '#1a1a1a' if is_dark_theme() else '#F5F7FA'
        win.configure(bg=bg)
        win.transient(self)
        win.grab_set()
        frame = ttk.Frame(win, padding=14)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="选择目标页面（可移动到任意已有页或新建页）：",
                  font=('微软雅黑', 10, 'bold')).pack(anchor='w', pady=(0, 10))
        lb = tk.Listbox(frame, height=min(len(self.pages) + 1, 12))
        for i, pg in enumerate(self.pages):
            lb.insert('end', f"第 {i+1} 页（{len(pg)} 张发票）")
        lb.insert('end', "＋ 新建空白页")
        lb.pack(fill='both', expand=True, pady=(0, 10))
        if self.current_page < len(self.pages):
            lb.selection_set(self.current_page)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')
        def do():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("提示", "请选择目标页面")
                return
            t = sel[0]
            win.destroy()
            self._move_selected_to(t)
        ttk.Button(btn_frame, text="取消", command=win.destroy).pack(side='right', padx=4)
        ttk.Button(btn_frame, text="确定移动", command=do, style='Accent.TButton').pack(side='right', padx=4)

    def _find_free_y(self, page, h, hr):
        """在页面中寻找可放置发票的空白 y 位置（自上而下找首个不与已有发票重叠且不越界的 y）"""
        if not page:
            return 0.02
        # 取所有已有发票的下边界，从最大下边界下方放置（简单且直观）
        max_bottom = 0.02
        for lo in page:
            bottom = lo['y'] + lo['h']
            if bottom > max_bottom:
                max_bottom = bottom
        y = max_bottom + 0.02
        if y + h > 1.0:
            # 底部放不下，尝试从顶部重新找空隙
            y = 0.02
        return y

    def _reauto_layout(self):
        self._auto_layout()
        self.current_page = 0
        self.selected = None
        self._update_sel_label()
        self._draw()

    def _on_paper_change(self, event):
        if self.paper_var.get() == '自定义':
            self.cust_w_entry.config(state='normal')
            self.cust_h_entry.config(state='normal')
        else:
            self.cust_w_entry.config(state='disabled')
            self.cust_h_entry.config(state='disabled')
            wmm, hmm = self.PAPER_SIZES[self.paper_var.get()]
            self.paper_w, self.paper_h = wmm, hmm
            self._reauto_layout()

    def _apply_paper(self):
        if self.paper_var.get() == '自定义':
            try:
                wcm = float(self.cust_w_var.get())
                hcm = float(self.cust_h_var.get())
                if wcm <= 0 or hcm <= 0:
                    raise ValueError
                self.paper_w, self.paper_h = wcm * 10, hcm * 10
            except ValueError:
                messagebox.showwarning("提示", "请输入正确的自定义尺寸（cm，正数）")
                return
            self._reauto_layout()
        else:
            wmm, hmm = self.PAPER_SIZES[self.paper_var.get()]
            self.paper_w, self.paper_h = wmm, hmm
            self._reauto_layout()

    # ---------- 绘制 ----------
    def _get_page_rect(self):
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw < 10 or ch < 10:
            return (0, 0, 0, 0)
        page_h_ratio = self.paper_h / self.paper_w
        pw = cw - 50
        ph = pw * page_h_ratio
        if ph > ch - 50:
            ph = ch - 50
            pw = ph / page_h_ratio
        ox = (cw - pw) / 2
        oy = (ch - ph) / 2
        return (ox, oy, pw, ph)

    def _draw(self):
        c = self.canvas
        c.delete('all')
        cw = c.winfo_width()
        ch = c.winfo_height()
        if cw < 10 or ch < 10:
            self.after(60, self._draw)
            return
        ox, oy, pw, ph = self._get_page_rect()
        if pw <= 0:
            return
        hr = self.paper_h / self.paper_w
        # 页面白色矩形 + 阴影 + 红色打印边界
        c.create_rectangle(ox + 5, oy + 5, ox + pw + 5, oy + ph + 5, fill='#000000', outline='')
        c.create_rectangle(ox, oy, ox + pw, oy + ph, fill='#ffffff', outline='#FF3B30', width=2)
        # 纸张尺寸标注
        c.create_text(ox + pw / 2, oy - 10, text=f"{self.paper_w}×{self.paper_h} mm（红线内为打印范围）",
                      fill='#FF6B6B', font=('微软雅黑', 9))
        page = self.pages[self.current_page]
        for li, lo in enumerate(page):
            px = ox + lo['x'] * pw
            py = oy + lo['y'] * pw
            w2 = lo['w'] * pw
            h2 = lo['h'] * pw
            photo = self._get_display_photo(lo, w2, h2)
            if photo:
                c.create_image(px, py, anchor='nw', image=photo)
            is_sel = (li == self.selected)
            color = '#3B82F6' if is_sel else '#666666'
            width = 2 if is_sel else 1
            c.create_rectangle(px, py, px + w2, py + h2, outline=color, width=width)
            c.create_text(px + 3, py + 3, text=str(li + 1), anchor='nw',
                          fill='#FFFFFF', font=('Arial', 8, 'bold'))
            if is_sel:
                # 绘制8个控制点：4角(方块) + 4边中点(圆点)
                handles = self._get_handles(lo, ox, oy, pw, ph)
                for hi, (hx, hy) in enumerate(handles):
                    r = 5 if hi < 4 else 4
                    if hi < 4:
                        c.create_rectangle(hx - r, hy - r, hx + r, hy + r,
                                           fill='#FFD60A', outline='#3B82F6', width=1)
                    else:
                        c.create_oval(hx - r, hy - r, hx + r, hy + r,
                                      fill='#FFD60A', outline='#3B82F6', width=1)
        # 页码
        self.page_label.config(text=f"第 {self.current_page + 1} / {len(self.pages)} 页")

    def _get_display_photo(self, lo, w2, h2):
        """获取旋转+缩放后的PhotoImage（带缓存，保持宽高比）"""
        item = self.items[lo['idx']]
        key = (lo['idx'], lo['rotation'], int(w2), int(h2))
        if item.get('_cache_key') == key and item.get('_cache_photo') is not None:
            return item['_cache_photo']
        img = get_invoice_image(item['path'])
        if img is None:
            return None
        if img.mode != 'RGB':
            img = img.convert('RGB')
        if lo['rotation']:
            img = img.rotate(lo['rotation'], expand=True)
        wpx = max(1, int(w2))
        hpx = max(1, int(h2))
        img = img.resize((wpx, hpx), Image.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        item['_cache_photo'] = photo
        item['_cache_key'] = key
        self._thumbs.append(photo)
        return photo

    # ---------- 导出 ----------
    def _export(self):
        if not self.pages or all(len(p) == 0 for p in self.pages):
            messagebox.showinfo("提示", "没有可导出的排版内容")
            return
        folder = filedialog.askdirectory(title="选择保存排版图片的文件夹")
        if not folder:
            return
        pw_px = int(round(self.paper_w / 25.4 * self.dpi))
        ph_px = int(round(self.paper_h / 25.4 * self.dpi))
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        saved = []
        total_pages = len(self.pages)
        try:
            for pi, page in enumerate(self.pages):
                if not page:
                    continue
                canvas_img = Image.new('RGB', (pw_px, ph_px), 'white')
                for lo in page:
                    item = self.items[lo['idx']]
                    img = get_invoice_image(item['path'])
                    if img is None:
                        continue
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    if lo['rotation']:
                        img = img.rotate(lo['rotation'], expand=True)
                    x = int(lo['x'] * pw_px)
                    y = int(lo['y'] * pw_px)
                    wpx = max(1, int(lo['w'] * pw_px))
                    hpx = max(1, int(lo['h'] * pw_px))
                    img = img.resize((wpx, hpx), Image.LANCZOS)
                    canvas_img.paste(img, (x, y))
                if total_pages > 1:
                    fname = f"打印排版_{ts}_第{pi+1}页.png"
                else:
                    fname = f"打印排版_{ts}.png"
                out_path = os.path.join(folder, fname)
                canvas_img.save(out_path, 'PNG')
                saved.append(out_path)
        except Exception as e:
            _log_ocr_error(f"导出打印排版失败: {e}")
            messagebox.showerror("错误", f"导出排版图片失败: {e}")
            return
        messagebox.showinfo("导出完成",
                            f"已导出 {len(saved)} 张排版图片到：\n{folder}\n\n"
                            f"共 {len(saved)} 页，分辨率 {self.dpi} DPI")
        self.app.set_status(f"打印排版已导出 {len(saved)} 张图片到 {folder}")
        # 标记已排版的发票
        try:
            laid_out_ids = set()
            for page in self.pages:
                for lo in page:
                    item = self.items[lo['idx']]
                    if 'id' in item:
                        laid_out_ids.add(item['id'])
            if laid_out_ids:
                conn = get_db()
                for inv_id in laid_out_ids:
                    conn.execute("UPDATE invoices SET printed=1 WHERE id=?", (inv_id,))
                conn.commit()
                conn.close()
                if hasattr(self, 'parent') and hasattr(self.parent, 'refresh'):
                    self.parent.refresh()
        except Exception as e:
            _log_ocr_error(f"标记已排版失败: {e}")

    def _update_sel_label(self):
        if self.selected is not None:
            try:
                page = self.pages[self.current_page]
                lo = page[self.selected]
                item = self.items[lo['idx']]
                self.sel_label.config(text=f"选中：{item['num'] or '无号'}（旋转 {lo['rotation']}°）")
            except Exception:
                self.sel_label.config(text="")
        else:
            self.sel_label.config(text="")
class ImageViewer(tk.Toplevel):
    def __init__(self, parent, img_path, invoice_info, invoice_id=None):
        super().__init__(parent)
        self.title(f"发票预览 - {invoice_info['invoice_number'] or '无号'}")
        _viewer_bg = '#1C1C1E' if is_dark_theme() else '#E6E7E8'
        _viewer_canvas_bg = '#000000' if is_dark_theme() else '#FFFFFF'
        self.configure(bg=_viewer_bg)
        self._img_path = img_path
        self._rotation = 0
        self._invoice_id = invoice_id
        self.minsize(500, 700)
        # 根据图片比例设置初始窗口大小（信息栏约120px + 控制栏约50px + 边距约30px）
        try:
            _tmp_img = Image.open(img_path)
            _iw, _ih = _tmp_img.size
            _tmp_img.close()
            _ui_h = 400  # 信息栏+控制栏(两行)+边距总高度（确保底部按钮完整显示）
            _ui_w = 40   # 左右边距+滚动条
            _max_w = int(self.winfo_screenwidth() * 0.75)
            _max_h = int(self.winfo_screenheight() * 0.85)
            # 计算图片在可用区域内的最大缩放
            _avail_w = max(100, _max_w - _ui_w)
            _avail_h = max(100, _max_h - _ui_h)
            _scale = min(_avail_w / _iw, _avail_h / _ih, 1.5)
            _init_w = max(500, int(_iw * _scale) + _ui_w)
            _init_h = max(700, int(_ih * _scale) + _ui_h)
            self.geometry(f"{_init_w}x{_init_h}")
            # 设置窗口初始位置：居中偏上，避免被任务栏遮挡
            _screen_w = self.winfo_screenwidth()
            _screen_h = self.winfo_screenheight()
            _x = max(0, (_screen_w - _init_w) // 2)
            _y = max(0, (_screen_h - _init_h) // 4)
            self.geometry(f"{_init_w}x{_init_h}+{_x}+{_y}")
        except Exception:
            self.geometry("700x900")
            _screen_w = self.winfo_screenwidth()
            _screen_h = self.winfo_screenheight()
            _x = max(0, (_screen_w - 700) // 2)
            _y = max(0, (_screen_h - 800) // 4)
            self.geometry(f"700x800+{_x}+{_y}")

        # 信息栏（顶部）
        info = ttk.Frame(self, padding=8)
        info.pack(fill='x', side='top')
        ttk.Label(info, text=f"发票号: {invoice_info['invoice_number'] or '-'}",
                  font=('微软雅黑', 10, 'bold')).pack(anchor='w')
        ttk.Label(info, text=f"日期: {invoice_info['invoice_date'] or '-'}").pack(anchor='w')
        ttk.Label(info, text=f"金额: {fmt_money(invoice_info['amount'])}").pack(anchor='w')
        ttk.Label(info, text=f"用途: {invoice_info['purpose'] or '-'}").pack(anchor='w')

        # 控制栏（底部，两行布局，确保窄窗口也能完整显示所有按钮）
        ctrl = ttk.Frame(self, padding=6)
        ctrl.pack(fill='x', side='bottom')
        # 第一行：旋转和缩放按钮
        ctrl_row1 = ttk.Frame(ctrl)
        ctrl_row1.pack(fill='x', pady=2)
        ttk.Button(ctrl_row1, text="↺ 左旋90°", command=lambda: self._rotate(-90)).pack(side='left', padx=4)
        ttk.Button(ctrl_row1, text="↻ 右旋90°", command=lambda: self._rotate(90)).pack(side='left', padx=4)
        ttk.Separator(ctrl_row1, orient='vertical').pack(side='left', fill='y', padx=4)
        ttk.Button(ctrl_row1, text="放大", command=lambda: self._zoom(1.2)).pack(side='left', padx=4)
        ttk.Button(ctrl_row1, text="缩小", command=lambda: self._zoom(0.8)).pack(side='left', padx=4)
        ttk.Button(ctrl_row1, text="适应窗口", command=lambda: self._fit()).pack(side='left', padx=4)
        # 第二行：关闭按钮（右对齐）
        ctrl_row2 = ttk.Frame(ctrl)
        ctrl_row2.pack(fill='x', pady=2)
        ttk.Button(ctrl_row2, text="关闭并保存", command=self._close_and_save).pack(side='right', padx=4)

        # 图片区域（中间，带水平+垂直滚动条）
        img_frame = ttk.Frame(self)
        img_frame.pack(fill='both', expand=True, padx=10, pady=8)
        # 先pack水平滚动条（底部），再pack canvas区域，确保布局正确
        sb_x = ttk.Scrollbar(img_frame, orient='horizontal', command=self.canvas.xview) if hasattr(self, 'canvas') else None
        canvas_frame = ttk.Frame(img_frame)
        canvas_frame.pack(side='top', fill='both', expand=True)
        self.canvas = tk.Canvas(canvas_frame, bg=_viewer_canvas_bg, highlightthickness=1,
                                highlightbackground='#333' if is_dark_theme() else '#ccc')
        self.canvas.pack(side='left', fill='both', expand=True)
        sb_y = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.canvas.yview)
        sb_y.pack(side='right', fill='y')
        sb_x = ttk.Scrollbar(img_frame, orient='horizontal', command=self.canvas.xview)
        sb_x.pack(side='bottom', fill='x')
        self.canvas.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        # 鼠标滚轮快捷键（默认Alt+滚轮水平滚动，因为Shift+滚轮会被Windows鼠标驱动拦截）
        _zoom_mod = get_setting('shortcut_zoom', 'Ctrl')
        _hscroll_mod = get_setting('shortcut_hscroll', 'Alt')
        _mod_state = {'Ctrl': 0x0004, 'Shift': 0x0001, 'Alt': 0x00020000}
        _zoom_state = _mod_state.get(_zoom_mod, 0x0004)
        _hscroll_state = _mod_state.get(_hscroll_mod, 0x00020000)

        def _on_wheel(event):
            st = event.state
            has_zoom = (_zoom_mod != '无') and bool(st & _zoom_state)
            has_hscroll = (_hscroll_mod != '无') and bool(st & _hscroll_state)
            if has_zoom:
                if event.delta > 0:
                    self._zoom(1.15)
                else:
                    self._zoom(0.87)
            elif has_hscroll:
                if event.delta > 0:
                    self.canvas.xview_scroll(-10, 'units')
                else:
                    self.canvas.xview_scroll(10, 'units')
            else:
                if event.delta > 0:
                    self.canvas.yview_scroll(-5, 'units')
                else:
                    self.canvas.yview_scroll(5, 'units')
            return 'break'

        # bind_all 确保全窗口任何控件都能触发滚轮事件
        self.bind_all('<MouseWheel>', _on_wheel)

        # 左键按住拖动平移图片（增量方式，最直接可靠）
        self._pan_dx = 0
        self._pan_dy = 0

        def _on_pan_start(event):
            self._pan_dx = event.x
            self._pan_dy = event.y
            self.canvas.config(cursor='fleur')

        def _on_pan_move(event):
            dx = event.x - self._pan_dx
            dy = event.y - self._pan_dy
            self._pan_dx = event.x
            self._pan_dy = event.y
            # 用scroll直接平移，步长=像素差/10，效果明显
            if dx != 0:
                self.canvas.xview_scroll(-dx, 'units')
            if dy != 0:
                self.canvas.yview_scroll(-dy, 'units')

        def _on_pan_end(event):
            self.canvas.config(cursor='')

        # 绑定到canvas和canvas_frame，确保点击图片任何位置都能触发
        self.canvas.bind('<ButtonPress-1>', _on_pan_start)
        self.canvas.bind('<B1-Motion>', _on_pan_move)
        self.canvas.bind('<ButtonRelease-1>', _on_pan_end)
        canvas_frame.bind('<ButtonPress-1>', _on_pan_start)
        canvas_frame.bind('<B1-Motion>', _on_pan_move)
        canvas_frame.bind('<ButtonRelease-1>', _on_pan_end)

        # 键盘方向键滚动（100%可靠，核对发票效率高）
        def _scroll_left(event):
            self.canvas.xview_scroll(-3, 'units')
            return 'break'
        def _scroll_right(event):
            self.canvas.xview_scroll(3, 'units')
            return 'break'
        def _scroll_up(event):
            self.canvas.yview_scroll(-3, 'units')
            return 'break'
        def _scroll_down(event):
            self.canvas.yview_scroll(3, 'units')
            return 'break'
        # PageUp/PageDown 大幅度滚动
        def _page_up(event):
            self.canvas.yview_scroll(-1, 'pages')
            return 'break'
        def _page_down(event):
            self.canvas.yview_scroll(1, 'pages')
            return 'break'
        self.bind('<Left>', _scroll_left)
        self.bind('<Right>', _scroll_right)
        self.bind('<Up>', _scroll_up)
        self.bind('<Down>', _scroll_down)
        self.bind('<Prior>', _page_up)
        self.bind('<Next>', _page_down)
        # 让canvas获取焦点，确保方向键生效
        self.canvas.focus_set()

        # 提示标签
        _tip = "Ctrl+滚轮:放大  滚轮:上下移  ←→↑↓方向键:滚动  PageUp/Down:翻页  左键拖动:平移（需先放大）"
        ttk.Label(img_frame, text=_tip,
                 font=('微软雅黑', 8), foreground='#888').pack(side='bottom', pady=(2, 0))

        self._img = None
        self._photo = None
        self.zoom = -1  # 默认适应窗口
        try:
            if img_path.lower().endswith('.pdf'):
                self._img = pdf_to_image(img_path)
            else:
                self._img = Image.open(img_path)
            if self._img:
                self._orig = self._img.copy()
                sw = self.winfo_screenwidth()
                sh = self.winfo_screenheight()
                iw, ih = self._orig.size
                max_w = int(sw * 0.8)
                max_h = int(sh * 0.8)
                avail_h = max_h - 220
                avail_w = max_w - 40
                ratio = min(avail_w / iw, avail_h / ih, 1.0)
                if ratio <= 0:
                    ratio = 0.5
                win_w = max(420, int(iw * ratio) + 40)
                win_h = max(520, int(ih * ratio) + 220)
                win_w = min(win_w, max_w)
                win_h = min(win_h, max_h)
                self.geometry(f"{win_w}x{win_h}")
        except Exception as e:
            ttk.Label(self, text=f"无法加载图片: {e}", foreground='red').pack(pady=20)

        self.bind('<Configure>', lambda e: self._render())
        self.after(200, self._render)

    def _rotate(self, angle):
        if not self._orig:
            return
        self._rotation = (self._rotation + angle) % 360
        self._orig = self._orig.rotate(angle, expand=True)
        self.zoom = -1
        self._render()

    def _close_and_save(self):
        if self._rotation != 0 and self._orig:
            try:
                if self._img_path.lower().endswith('.pdf'):
                    # PDF旋转后保存为同名jpg
                    base = os.path.splitext(os.path.basename(self._img_path))[0]
                    new_name = f"{base}_rotated.jpg"
                    new_path = os.path.join(INVOICE_DIR, new_name)
                    self._orig.save(new_path, 'JPEG', quality=92)
                    # 更新数据库image_path
                    if self._invoice_id:
                        conn = get_db()
                        try:
                            conn.execute("UPDATE invoices SET image_path=? WHERE id=?",
                                         (new_name, self._invoice_id))
                            conn.commit()
                        finally:
                            conn.close()
                else:
                    ext = os.path.splitext(self._img_path)[1].lower()
                    if ext in ('.jpg', '.jpeg'):
                        self._orig.save(self._img_path, 'JPEG', quality=92)
                    elif ext == '.png':
                        self._orig.save(self._img_path, 'PNG')
                    else:
                        self._orig.save(self._img_path)
            except Exception as e:
                _log_ocr_error(f"保存旋转图片失败: {e}")
        self.destroy()

    def _zoom(self, factor):
        if self.zoom == -1:
            self.zoom = 1.0
        self.zoom *= factor
        self.zoom = max(0.1, min(self.zoom, 10))
        self._render()

    def _fit(self):
        self.zoom = -1
        self._render()

    def _orig_size(self):
        self.zoom = 1.0
        self._render()

    def _render(self):
        if not self._img:
            return
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        # 窗口未完全显示时用默认值
        if cw < 10:
            cw = 800
        if ch < 10:
            ch = 600
        iw, ih = self._orig.size
        if self.zoom == -1:
            # 适应窗口：保持zoom=-1，每次都重新计算
            ratio = min(cw / iw, ch / ih)
            if ratio <= 0:
                ratio = 1.0
            nw = max(1, int(iw * ratio))
            nh = max(1, int(ih * ratio))
        else:
            nw = max(1, int(iw * self.zoom))
            nh = max(1, int(ih * self.zoom))
        resized = self._orig.resize((nw, nh), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(resized)
        # 释放resized图片内存（PhotoImage已复制数据）
        try:
            resized.close()
        except Exception:
            pass
        self.canvas.delete('all')
        # 图片大于画布时左上角对齐，支持滚动；小于时居中（始终居中显示）
        if nw > cw or nh > ch:
            self.canvas.create_image(0, 0, image=self._photo, anchor='nw')
            self.canvas.configure(scrollregion=(0, 0, nw, nh))
        else:
            # 居中显示，确保左右上下留白均匀
            _cx = max(0, (cw - nw) // 2)
            _cy = max(0, (ch - nh) // 2)
            self.canvas.create_image(_cx + nw // 2, _cy + nh // 2, image=self._photo, anchor='center')
            self.canvas.configure(scrollregion=(0, 0, cw, ch))


# ============================================================
# 工资结算
# ============================================================
class SalaryTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.selected_payment_id = None

        # 顶部：年月选择 + 发放表单
        top = ttk.Frame(self.content)
        top.pack(fill='x', padx=10, pady=8)

        ttk.Label(top, text="年份:").pack(side='left', padx=4)
        now = datetime.now()
        self.year_var = tk.StringVar(value=str(now.year))
        min_y = _get_min_hire_year()
        years = [str(y) for y in range(min_y, now.year + 1)]
        ttk.Combobox(top, textvariable=self.year_var, values=years, width=6,
                     state='readonly').pack(side='left', padx=4)
        ttk.Label(top, text="月份:").pack(side='left', padx=4)
        self.month_var = tk.StringVar(value=str(now.month))
        ttk.Combobox(top, textvariable=self.month_var, values=[str(m) for m in range(1, 13)],
                     width=4, state='readonly').pack(side='left', padx=4)
        ttk.Button(top, text="查询", command=self.refresh).pack(side='left', padx=8)
        ttk.Button(top, text="导入Excel", command=self.import_excel).pack(side='right', padx=4)
        ttk.Button(top, text="导出Excel", command=self.export_excel).pack(side='right', padx=4)

        # 发放表单
        pay_form = ttk.LabelFrame(self.content, text="发放 / 修改工资", padding=8)
        pay_form.pack(fill='x', padx=10, pady=4)

        ttk.Label(pay_form, text="员工:").grid(row=0, column=0, sticky='e', padx=4, pady=3)
        self.pay_emp_var = tk.StringVar()
        self.pay_emp_cb = ttk.Combobox(pay_form, textvariable=self.pay_emp_var, width=14, state='readonly')
        self.pay_emp_cb.grid(row=0, column=1, padx=4, pady=3)

        ttk.Label(pay_form, text="发放金额:").grid(row=0, column=2, sticky='e', padx=4, pady=3)
        self.pay_amount_var = tk.StringVar()
        ttk.Entry(pay_form, textvariable=self.pay_amount_var, width=12).grid(row=0, column=3, padx=4, pady=3)

        ttk.Label(pay_form, text="备注:").grid(row=0, column=4, sticky='e', padx=4, pady=3)
        self.pay_note_var = tk.StringVar()
        ttk.Entry(pay_form, textvariable=self.pay_note_var, width=18).grid(row=0, column=5, padx=4, pady=3)

        ttk.Label(pay_form, text="银行账户:").grid(row=0, column=6, sticky='e', padx=4, pady=3)
        self.pay_bank_var = tk.StringVar()
        self.pay_bank_cb = ttk.Combobox(pay_form, textvariable=self.pay_bank_var, width=12, state='readonly')
        self.pay_bank_cb.grid(row=0, column=7, padx=4, pady=3)

        self.pay_btn = ttk.Button(pay_form, text="确认发放", command=self.pay_or_update)
        self.pay_btn.grid(row=0, column=8, padx=8, pady=3)
        self.pay_cancel_btn = ttk.Button(pay_form, text="取消修改", command=self.cancel_edit, state='disabled')
        self.pay_cancel_btn.grid(row=0, column=9, padx=4, pady=3)

        # 当月结算表格
        tbl_frame = ttk.LabelFrame(self.content, text="当月薪资结算", padding=6)
        tbl_frame.pack(fill='both', expand=True, padx=10, pady=6)
        tbl_wrap = ttk.Frame(tbl_frame)
        tbl_wrap.pack(fill='both', expand=True)

        cols = ('id', 'name', 'monthly_salary', 'paid', 'unpaid')
        self.tree = ttk.Treeview(tbl_wrap, columns=cols, show='headings', selectmode='browse')
        for cid, text, w, anchor in [('id', '编号', 60, 'center'),
                                     ('name', '姓名', 150, 'center'),
                                     ('monthly_salary', '应发薪资', 140, 'e'),
                                     ('paid', '已结', 140, 'e'),
                                     ('unpaid', '未结', 140, 'e')]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor)
        self.tree.pack(side='left', fill='both', expand=True)
        tbl_sb = ttk.Scrollbar(tbl_wrap, orient='vertical', command=self.tree.yview)
        tbl_sb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=tbl_sb.set)

        # 发放记录
        rec_frame = ttk.LabelFrame(self.content, text="发放记录（当月）", padding=6)
        rec_frame.pack(fill='both', expand=True, padx=10, pady=6)
        rec_wrap = ttk.Frame(rec_frame)
        rec_wrap.pack(fill='both', expand=True)

        cols2 = ('id', 'name', 'year_month', 'amount', 'note', 'paid_at')
        self.rec_tree = ttk.Treeview(rec_wrap, columns=cols2, show='headings', selectmode='browse')
        for cid, text, w, anchor in [('id', '编号', 60, 'center'),
                                     ('name', '员工', 120, 'center'),
                                     ('year_month', '年月', 80, 'center'),
                                     ('amount', '金额', 120, 'e'),
                                     ('note', '备注', 220, 'w'),
                                     ('paid_at', '发放时间', 150, 'center')]:
            self.rec_tree.heading(cid, text=text)
            self.rec_tree.column(cid, width=w, anchor=anchor)
        self.rec_tree.pack(side='left', fill='both', expand=True)
        rec_sb = ttk.Scrollbar(rec_wrap, orient='vertical', command=self.rec_tree.yview)
        rec_sb.pack(side='right', fill='y')
        self.rec_tree.configure(yscrollcommand=rec_sb.set)
        self.rec_tree.bind('<Double-1>', self.on_rec_double)

        # 底部
        btns = ttk.Frame(self.content)
        btns.pack(fill='x', padx=10, pady=4)
        ttk.Button(btns, text="删除选中发放记录", command=self.delete_payment).pack(side='left', padx=4)
        ttk.Button(btns, text="刷新", command=self.refresh).pack(side='left', padx=4)
        ttk.Label(btns, text="（双击发放记录可修改）", foreground='gray').pack(side='left', padx=10)

        self._load_emps()

    def _load_emps(self):
        conn = get_db()
        rows = conn.execute("SELECT name FROM employees WHERE resigned=0 ORDER BY name").fetchall()
        conn.close()
        self.pay_emp_cb['values'] = [r['name'] for r in rows]

    def _load_bank_accounts(self):
        conn = get_db()
        rows = conn.execute("SELECT id, name FROM bank_accounts ORDER BY name").fetchall()
        conn.close()
        self.pay_bank_cb['values'] = [r['name'] for r in rows]
        self._bank_map = {r['name']: r['id'] for r in rows}

    def refresh(self):
        self._load_emps()
        self._load_bank_accounts()
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            return

        # 结算表
        for i in self.tree.get_children():
            self.tree.delete(i)
        conn = get_db()
        emps = conn.execute("SELECT * FROM employees ORDER BY id").fetchall()
        for e in emps:
            paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                   WHERE employee_id=? AND year=? AND month=?""",
                                (e['id'], year, month)).fetchone()['s']
            # 根据入职、离职、调薪记录计算当月应发薪资
            total = get_monthly_salary(e['id'], year, month)
            unpaid = total - paid
            name = e['name']
            if e['resigned']:
                name += ' (已离职)'
            self.tree.insert('', 'end', values=(e['id'], name, fmt_money(total),
                                                fmt_money(paid), fmt_money(unpaid)))
        # 发放记录
        for i in self.rec_tree.get_children():
            self.rec_tree.delete(i)
        rows = conn.execute("""SELECT p.*, e.name, e.resigned FROM salary_payments p
                               LEFT JOIN employees e ON p.employee_id=e.id
                               WHERE p.year=? AND p.month=? ORDER BY p.id DESC""",
                            (year, month)).fetchall()
        conn.close()
        for r in rows:
            name = r['name'] or '(已删除)'
            if r['resigned']:
                name += ' (已离职)'
            self.rec_tree.insert('', 'end', values=(
                r['id'], name, f"{r['year']}-{r['month']:02d}",
                fmt_money(r['amount']), r['note'] or '', r['paid_at'] or ''))

    def on_rec_double(self, event=None):
        """双击发放记录加载到表单修改"""
        sel = self.rec_tree.selection()
        if not sel:
            return
        pid = self.rec_tree.item(sel[0], 'values')[0]
        conn = get_db()
        row = conn.execute("""SELECT p.*, e.name, b.name as bname FROM salary_payments p
                              LEFT JOIN employees e ON p.employee_id=e.id
                              LEFT JOIN bank_accounts b ON p.bank_account_id=b.id
                              WHERE p.id=?""",
                           (pid,)).fetchone()
        conn.close()
        if not row:
            return
        self.selected_payment_id = pid
        self.pay_emp_var.set(row['name'] or '')
        self.pay_amount_var.set(str(row['amount'] or 0))
        self.pay_note_var.set(row['note'] or '')
        self.pay_bank_var.set(row['bname'] or '')
        self.year_var.set(str(row['year']))
        self.month_var.set(str(row['month']))
        self.pay_btn.config(text="保存修改")
        self.pay_cancel_btn.config(state='normal')

    def cancel_edit(self):
        self.selected_payment_id = None
        self.pay_emp_var.set('')
        self.pay_amount_var.set('')
        self.pay_note_var.set('')
        self.pay_bank_var.set('')
        self.pay_btn.config(text="确认发放")
        self.pay_cancel_btn.config(state='disabled')

    def pay_or_update(self):
        emp_name = self.pay_emp_var.get().strip()
        amount = self.pay_amount_var.get().strip()
        note = self.pay_note_var.get().strip()
        bank_name = self.pay_bank_var.get().strip()
        bank_id = self._bank_map.get(bank_name) if bank_name else None
        if not emp_name:
            messagebox.showwarning("提示", "请选择员工")
            return
        try:
            amount = float(amount) if amount else 0
        except ValueError:
            messagebox.showwarning("提示", "金额必须是数字")
            return
        if amount <= 0:
            messagebox.showwarning("提示", "发放金额必须大于0")
            return
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            return
        conn = get_db()
        emp = conn.execute("SELECT id FROM employees WHERE name=?", (emp_name,)).fetchone()
        if not emp:
            messagebox.showerror("错误", "员工不存在")
            conn.close()
            return

        if self.selected_payment_id:
            # 修改模式
            conn.execute("""UPDATE salary_payments SET employee_id=?, year=?, month=?, amount=?, note=?,
                         bank_account_id=? WHERE id=?""",
                         (emp['id'], year, month, amount, note, bank_id, self.selected_payment_id))
            # 更新银行出项记录：先删旧的，再根据新账户生成
            conn.execute("DELETE FROM bank_transactions WHERE source='工资发放' AND source_id=?",
                         (self.selected_payment_id,))
            if bank_id:
                purpose = f"工资发放：{emp_name} {year}年{month}月（{note or ''}）"
                conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                              source, source_id, transaction_date, created_at) VALUES(?,?,?,?,?,?,?,?)""",
                             (bank_id, '出项', amount, purpose, '工资发放', self.selected_payment_id,
                              datetime.now().strftime('%Y-%m-%d'),
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()
            self.app.set_status(f"已修改发放记录: {emp_name} {fmt_money(amount)}")
        else:
            # 发放模式
            conn.execute("""INSERT INTO salary_payments(employee_id, year, month, amount, note, paid_at, bank_account_id)
                      VALUES(?,?,?,?,?,?,?)""",
                         (emp['id'], year, month, amount, note,
                          datetime.now().strftime('%Y-%m-%d %H:%M:%S'), bank_id))
            new_id = conn.execute("SELECT last_insert_rowid() id").fetchone()['id']
            if bank_id:
                purpose = f"工资发放：{emp_name} {year}年{month}月（{note or ''}）"
                conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                              source, source_id, transaction_date, created_at) VALUES(?,?,?,?,?,?,?,?)""",
                             (bank_id, '出项', amount, purpose, '工资发放', new_id,
                              datetime.now().strftime('%Y-%m-%d'),
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()
            self.app.set_status(f"已为 {emp_name} 发放工资 {fmt_money(amount)}")

        self.cancel_edit()
        self.refresh()
        self.app.stat_tab.refresh()
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()

    def delete_payment(self):
        sel = self.rec_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择发放记录")
            return
        pid = self.rec_tree.item(sel[0], 'values')[0]
        if not messagebox.askyesno("确认", "确定删除该发放记录？"):
            return
        conn = get_db()
        conn.execute("DELETE FROM bank_transactions WHERE source='工资发放' AND source_id=?", (pid,))
        conn.execute("DELETE FROM salary_payments WHERE id=?", (pid,))
        conn.commit()
        conn.close()
        self.refresh()
        self.app.stat_tab.refresh()
        self.app.bank_tab.refresh_accounts()
        if self.app.bank_tab.selected_account_id:
            self.app.bank_tab.refresh_transactions()
            self.app.bank_tab.refresh_stats()
        self.app.stat_tab.refresh()
        self.app.set_status("已删除发放记录")

    def export_excel(self):
        """导出当月工资结算到Excel"""
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
        except ValueError:
            messagebox.showwarning("提示", "请先选择年份和月份")
            return

        conn = get_db()
        emps = conn.execute("SELECT * FROM employees ORDER BY id").fetchall()
        payments = conn.execute("""SELECT p.*, e.name FROM salary_payments p
                                   LEFT JOIN employees e ON p.employee_id=e.id
                                   WHERE p.year=? AND p.month=? ORDER BY p.id DESC""",
                                (year, month)).fetchall()
        conn.close()

        default_name = f"工资结算_{year}年{month:02d}月_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出Excel", defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return

        wb = Workbook()

        # ===== Sheet1: 当月薪资结算 =====
        ws1 = wb.active
        ws1.title = "当月薪资结算"

        ws1.merge_cells('A1:E1')
        ws1['A1'] = f"{year}年{month:02d}月 薪资结算表"
        ws1['A1'].font = TITLE_FONT
        ws1['A1'].alignment = CENTER
        ws1.row_dimensions[1].height = 28

        ws1.merge_cells('A2:E2')
        ws1['A2'] = f"导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws1['A2'].font = Font(name='微软雅黑', size=9, color='666666')
        ws1['A2'].alignment = LEFT

        headers1 = ['编号', '姓名', '应发薪资(元)', '已结(元)', '未结(元)']
        for c, h in enumerate(headers1, 1):
            ws1.cell(row=3, column=c, value=h)
        _style_header(ws1, 3, len(headers1))

        total_payable = 0
        total_paid = 0
        total_unpaid = 0
        for idx, e in enumerate(emps, 4):
            conn = get_db()
            paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                   WHERE employee_id=? AND year=? AND month=?""",
                                (e['id'], year, month)).fetchone()['s']
            conn.close()
            payable = get_monthly_salary(e['id'], year, month)
            unpaid = payable - paid
            ws1.cell(row=idx, column=1, value=e['id']).alignment = CENTER
            ws1.cell(row=idx, column=2, value=e['name']).alignment = CENTER
            for col, val in [(3, payable), (4, paid), (5, unpaid)]:
                cell = ws1.cell(row=idx, column=col, value=val)
                cell.number_format = '#,##0.00'
                cell.alignment = RIGHT
            for c in range(1, 6):
                ws1.cell(row=idx, column=c).border = THIN_BORDER
            total_payable += payable
            total_paid += paid
            total_unpaid += unpaid

        # 合计行
        tr = len(emps) + 4
        ws1.merge_cells(f'A{tr}:B{tr}')
        ws1.cell(row=tr, column=1, value=f"合计（共{len(emps)}人）").font = TOTAL_FONT
        ws1.cell(row=tr, column=1).alignment = RIGHT
        ws1.cell(row=tr, column=1).fill = TOTAL_FILL
        for col, val in [(3, total_payable), (4, total_paid), (5, total_unpaid)]:
            cell = ws1.cell(row=tr, column=col, value=val)
            cell.number_format = '#,##0.00'
            cell.font = TOTAL_FONT
            cell.fill = TOTAL_FILL
            cell.alignment = RIGHT
        for c in range(1, 6):
            ws1.cell(row=tr, column=c).border = THIN_BORDER

        _auto_width(ws1, len(headers1))
        ws1.freeze_panes = 'A4'

        # ===== Sheet2: 发放记录 =====
        ws2 = wb.create_sheet("发放记录")
        ws2.merge_cells('A1:F1')
        ws2['A1'] = f"{year}年{month:02d}月 工资发放记录"
        ws2['A1'].font = TITLE_FONT
        ws2['A1'].alignment = CENTER
        ws2.row_dimensions[1].height = 28

        headers2 = ['编号', '员工', '年月', '金额(元)', '备注', '发放时间']
        for c, h in enumerate(headers2, 1):
            ws2.cell(row=2, column=c, value=h)
        _style_header(ws2, 2, len(headers2))

        pay_total = 0
        for idx, p in enumerate(payments, 3):
            ws2.cell(row=idx, column=1, value=p['id']).alignment = CENTER
            ws2.cell(row=idx, column=2, value=p['name'] or '(已删除)').alignment = CENTER
            ws2.cell(row=idx, column=3, value=f"{p['year']}-{p['month']:02d}").alignment = CENTER
            cell = ws2.cell(row=idx, column=4, value=float(p['amount']))
            cell.number_format = '#,##0.00'
            cell.alignment = RIGHT
            ws2.cell(row=idx, column=5, value=p['note'] or '').alignment = LEFT
            ws2.cell(row=idx, column=6, value=p['paid_at'] or '').alignment = CENTER
            for c in range(1, 7):
                ws2.cell(row=idx, column=c).border = THIN_BORDER
            pay_total += float(p['amount'])

        if payments:
            tr2 = len(payments) + 3
            ws2.merge_cells(f'A{tr2}:C{tr2}')
            ws2.cell(row=tr2, column=1, value=f"合计（共{len(payments)}笔）").font = TOTAL_FONT
            ws2.cell(row=tr2, column=1).alignment = RIGHT
            ws2.cell(row=tr2, column=1).fill = TOTAL_FILL
            cell = ws2.cell(row=tr2, column=4, value=pay_total)
            cell.number_format = '#,##0.00'
            cell.font = TOTAL_FONT
            cell.fill = TOTAL_FILL
            cell.alignment = RIGHT
            ws2.merge_cells(f'E{tr2}:F{tr2}')
            for c in range(1, 7):
                ws2.cell(row=tr2, column=c).border = THIN_BORDER

        _auto_width(ws2, len(headers2))
        ws2.freeze_panes = 'A3'

        try:
            wb.save(path)
            self.app.set_status(f"已导出Excel: {os.path.basename(path)}")
            messagebox.showinfo("导出成功", f"已导出工资结算表到：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def import_excel(self):
        """从Excel导入工资发放记录"""
        path = filedialog.askopenfilename(title="选择要导入的Excel文件",
                                          filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取Excel文件：{e}")
            return

        # 自动识别表头行
        header_row = None
        col_map = {}
        for row_idx in range(1, min(ws.max_row, 10) + 1):
            row_data = [str(ws.cell(row=row_idx, column=c).value or '').strip() for c in range(1, ws.max_column + 1)]
            if '员工' in row_data or '姓名' in row_data or '金额' in row_data:
                header_row = row_idx
                for c_idx, val in enumerate(row_data, 1):
                    col_map[val] = c_idx
                break

        if header_row is None:
            messagebox.showerror("导入失败", "未识别到表头行，请确保Excel包含「员工」「金额」等列名。")
            return

        def get_col(*names):
            for n in names:
                if n in col_map:
                    return col_map[n]
            return None

        col_emp = get_col('员工', '姓名', '员工姓名', '人员')
        col_amount = get_col('金额', '金额(元)', '发放金额', '实发金额')
        col_ym = get_col('年月', '月份', '年-月', '发放月份')
        col_year = get_col('年', '年份', '年度')
        col_month = get_col('月', '月份', '月度')
        col_note = get_col('备注', '说明', '摘要')

        if col_emp is None or col_amount is None:
            messagebox.showerror("导入失败", "Excel必须包含「员工」和「金额」列。")
            return

        conn = get_db()
        success = 0
        failed = 0
        errors = []
        now_year = datetime.now().year
        now_month = datetime.now().month

        for row_idx in range(header_row + 1, ws.max_row + 1):
            def cell_val(col):
                if col is None:
                    return ''
                v = ws.cell(row=row_idx, column=col).value
                return str(v).strip() if v is not None else ''

            emp_name = cell_val(col_emp)
            amount_str = cell_val(col_amount)

            if not emp_name and not amount_str:
                continue

            try:
                amount = float(amount_str.replace(',', '').replace('¥', '')) if amount_str else 0
            except ValueError:
                failed += 1
                errors.append(f"第{row_idx}行：金额格式错误「{amount_str}」")
                continue

            if amount <= 0:
                failed += 1
                errors.append(f"第{row_idx}行：金额必须大于0")
                continue

            # 解析年月
            year = now_year
            month = now_month
            if col_ym:
                ym_str = cell_val(col_ym)
                if ym_str:
                    import re
                    m = re.search(r'(\d{4})[-/年](\d{1,2})', ym_str)
                    if m:
                        year = int(m.group(1))
                        month = int(m.group(2))
                    else:
                        m2 = re.search(r'(\d{1,2})月', ym_str)
                        if m2:
                            month = int(m2.group(1))
            if col_year:
                y_str = cell_val(col_year)
                if y_str:
                    try:
                        year = int(y_str)
                    except ValueError:
                        pass
            if col_month:
                m_str = cell_val(col_month)
                if m_str:
                    try:
                        month = int(m_str)
                    except ValueError:
                        pass

            if month < 1 or month > 12:
                month = now_month

            # 查找员工
            emp = conn.execute("SELECT id FROM employees WHERE name=?", (emp_name,)).fetchone()
            if not emp:
                failed += 1
                errors.append(f"第{row_idx}行：员工「{emp_name}」不存在，请先在人员管理中添加")
                continue
            emp_id = emp['id']

            note = cell_val(col_note)

            try:
                conn.execute("""INSERT INTO salary_payments(employee_id, year, month, amount, note, paid_at)
                              VALUES(?,?,?,?,?,?)""",
                             (emp_id, year, month, amount, note,
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                success += 1
            except Exception as e:
                failed += 1
                errors.append(f"第{row_idx}行：导入失败 - {e}")

        conn.commit()
        conn.close()

        self.refresh()
        self.app.stat_tab.refresh()

        msg = f"导入完成！\n\n成功：{success} 条\n失败：{failed} 条"
        if errors:
            msg += "\n\n失败详情：\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                msg += f"\n...（共{len(errors)}条错误）"
        messagebox.showinfo("导入结果", msg)
        self.app.set_status(f"工资导入完成：成功{success}条，失败{failed}条")


# ============================================================
# 统计汇总
# ============================================================
class StatsTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app

        # 选择区
        top = ttk.Frame(self.content)
        top.pack(fill='x', padx=10, pady=8)
        ttk.Label(top, text="统计年份:").pack(side='left', padx=4)
        now = datetime.now()
        self.year_var = tk.StringVar(value=str(now.year))
        min_y = _get_min_hire_year()
        years = ['全部'] + [str(y) for y in range(min_y, now.year + 1)]
        ttk.Combobox(top, textvariable=self.year_var, values=years, width=6,
                     state='readonly').pack(side='left', padx=4)
        ttk.Button(top, text="刷新统计", command=self.refresh).pack(side='left', padx=8)
        ttk.Button(top, text="导入Excel", command=self.import_excel).pack(side='right', padx=4)
        ttk.Button(top, text="导出Excel", command=self.export_excel).pack(side='right', padx=4)

        # 用 Notebook 分子标签
        self.sub_nb = ttk.Notebook(self.content)
        self.sub_nb.pack(fill='both', expand=True, padx=10, pady=6)

        # ---- 发票统计 ----
        inv_frame = ttk.Frame(self.sub_nb)
        self.sub_nb.add(inv_frame, text='  发票报销统计  ')

        # 按人汇总
        ttk.Label(inv_frame, text="按人员汇总（含发票+支付记录）", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(6, 2))
        cols = ('name', 'inv_amt', 'pay_amt', 'total', 'pending', 'unpaid', 'paid')
        self.inv_person_tree = self._make_tree(inv_frame, cols,
            [('name', '姓名', 120, 'center'),
             ('inv_amt', '发票金额', 110, 'e'),
             ('pay_amt', '支付记录', 110, 'e'),
             ('total', '报销总额', 110, 'e'),
             ('pending', '待完善', 110, 'e'),
             ('unpaid', '未报销', 110, 'e'),
             ('paid', '已报销', 110, 'e')])

        # 按月汇总
        ttk.Label(inv_frame, text="按月汇总（含发票+支付记录）", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(10, 2))
        cols2 = ('month', 'inv_amt', 'pay_amt', 'total', 'pending', 'unpaid', 'paid')
        self.inv_month_tree = self._make_tree(inv_frame, cols2,
            [('month', '月份', 90, 'center'),
             ('inv_amt', '发票金额', 110, 'e'),
             ('pay_amt', '支付记录', 110, 'e'),
             ('total', '报销总额', 110, 'e'),
             ('pending', '待完善', 110, 'e'),
             ('unpaid', '未报销', 110, 'e'),
             ('paid', '已报销', 110, 'e')])

        # ---- 工资统计 ----
        sal_frame = ttk.Frame(self.sub_nb)
        self.sub_nb.add(sal_frame, text='  工资结算统计  ')

        ttk.Label(sal_frame, text="按人员年度汇总", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(6, 2))
        cols3 = ('name', 'total_salary', 'paid', 'unpaid')
        self.sal_person_tree = self._make_tree(sal_frame, cols3,
            [('name', '姓名', 150, 'center'), ('total_salary', '年度应发', 140, 'e'),
             ('paid', '已结', 140, 'e'), ('unpaid', '未结', 140, 'e')])

        ttk.Label(sal_frame, text="按月度汇总", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(10, 2))
        cols4 = ('month', 'total_salary', 'paid', 'unpaid')
        self.sal_month_tree = self._make_tree(sal_frame, cols4,
            [('month', '月份', 120, 'center'), ('total_salary', '应发总额', 140, 'e'),
             ('paid', '已结', 140, 'e'), ('unpaid', '未结', 140, 'e')])

        # ---- 五险一金统计 ----
        ins_frame = ttk.Frame(self.sub_nb)
        self.sub_nb.add(ins_frame, text='  五险一金统计  ')

        ttk.Label(ins_frame, text="按人员汇总（公司/个人/合计）", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(6, 2))
        cols5 = ('name', 'company', 'personal', 'total')
        self.ins_person_tree = self._make_tree(ins_frame, cols5,
            [('name', '姓名', 150, 'center'),
             ('company', '公司缴纳', 130, 'e'),
             ('personal', '个人缴纳', 130, 'e'),
             ('total', '合计', 130, 'e')])

        ttk.Label(ins_frame, text="按月度汇总（公司/个人/合计）", font=('微软雅黑', 10, 'bold')).pack(anchor='w', padx=6, pady=(10, 2))
        cols6 = ('month', 'company', 'personal', 'total')
        self.ins_month_tree = self._make_tree(ins_frame, cols6,
            [('month', '月份', 120, 'center'),
             ('company', '公司缴纳', 130, 'e'),
             ('personal', '个人缴纳', 130, 'e'),
             ('total', '合计', 130, 'e')])

    def _make_tree(self, parent, cols, headers):
        f = ttk.Frame(parent)
        f.pack(fill='both', expand=True, padx=6, pady=2)
        tree = ttk.Treeview(f, columns=cols, show='headings', height=8)
        for cid, text, w, anchor in headers:
            tree.heading(cid, text=text)
            tree.column(cid, width=w, anchor=anchor)
        tree.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(f, orient='vertical', command=tree.yview)
        sb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=sb.set)
        return tree

    def refresh(self):
        year_sel = self.year_var.get()
        is_all = (year_sel == '全部')
        conn = get_db()

        # ========== 发票统计：Python端聚合，兼容各种日期格式和无报销人 ==========
        import re as _re
        def _parse_year_month(date_str):
            if not date_str:
                return None, None, None
            s = str(date_str).strip()
            m = _re.match(r'(\d{4})[-/.年]?(\d{1,2})', s)
            if m:
                y = int(m.group(1))
                mo = int(m.group(2))
                if 1 <= mo <= 12:
                    return y, mo, f"{y}-{mo:02d}"
            return None, None, None

        emp_map = {}
        for e in conn.execute("SELECT id, name, resigned FROM employees ORDER BY name").fetchall():
            emp_map[e['id']] = {'name': e['name'], 'resigned': e['resigned']}

        all_inv = conn.execute("""SELECT i.*, e.name as rname FROM invoices i
                                   LEFT JOIN employees e ON i.reimburser_id=e.id""").fetchall()

        person_agg = {}
        for emp_id, info in emp_map.items():
            person_agg[emp_id] = {
                'name': info['name'], 'resigned': info['resigned'],
                'inv_amt': 0, 'pay_amt': 0, 'total': 0,
                'pending': 0, 'unpaid': 0, 'paid': 0
            }
        person_agg['none'] = {
            'name': '未指定报销人', 'resigned': 0,
            'inv_amt': 0, 'pay_amt': 0, 'total': 0,
            'pending': 0, 'unpaid': 0, 'paid': 0
        }

        month_agg = {}

        for inv in all_inv:
            y, mo, ym = _parse_year_month(inv['invoice_date'])
            if not is_all and y != int(year_sel):
                continue
            amt = float(inv['amount'] or 0)
            is_pay = (inv['type'] == '支付记录')
            is_pending = (inv['batch_pending'] == 1)
            is_paid = (inv['status'] == 1)

            key = inv['reimburser_id'] if inv['reimburser_id'] in person_agg else 'none'
            pa = person_agg[key]
            pa['total'] += amt
            if is_pay:
                pa['pay_amt'] += amt
            else:
                pa['inv_amt'] += amt
            if is_pending:
                pa['pending'] += amt
            if is_paid:
                pa['paid'] += amt
            else:
                pa['unpaid'] += amt

            if ym:
                if ym not in month_agg:
                    month_agg[ym] = {
                        'inv_amt': 0, 'pay_amt': 0, 'total': 0,
                        'pending': 0, 'unpaid': 0, 'paid': 0
                    }
                ma = month_agg[ym]
                ma['total'] += amt
                if is_pay:
                    ma['pay_amt'] += amt
                else:
                    ma['inv_amt'] += amt
                if is_pending:
                    ma['pending'] += amt
                if is_paid:
                    ma['paid'] += amt
                else:
                    ma['unpaid'] += amt

        for i in self.inv_person_tree.get_children():
            self.inv_person_tree.delete(i)
        for key in sorted(person_agg.keys(), key=lambda k: (k == 'none', person_agg[k]['name'])):
            pa = person_agg[key]
            if pa['total'] == 0 and key != 'none':
                continue
            name = pa['name']
            if pa['resigned']:
                name += ' (已离职)'
            self.inv_person_tree.insert('', 'end', values=(
                name, fmt_money(pa['inv_amt']), fmt_money(pa['pay_amt']),
                fmt_money(pa['total']), fmt_money(pa['pending']),
                fmt_money(pa['unpaid']), fmt_money(pa['paid'])))

        for i in self.inv_month_tree.get_children():
            self.inv_month_tree.delete(i)
        for ym in sorted(month_agg.keys()):
            ma = month_agg[ym]
            self.inv_month_tree.insert('', 'end', values=(
                ym, fmt_money(ma['inv_amt']), fmt_money(ma['pay_amt']),
                fmt_money(ma['total']), fmt_money(ma['pending']),
                fmt_money(ma['unpaid']), fmt_money(ma['paid'])))

        # === 工资按人年度 ===
        for i in self.sal_person_tree.get_children():
            self.sal_person_tree.delete(i)
        emps = conn.execute("SELECT * FROM employees ORDER BY name").fetchall()
        if is_all:
            # 全部年份：确定年份范围
            yr_min = conn.execute("SELECT MIN(year) FROM salary_payments").fetchone()[0]
            yr_max = conn.execute("SELECT MAX(year) FROM salary_payments").fetchone()[0]
            if not yr_min or not yr_max:
                yr_min = yr_max = datetime.now().year
            years_range = range(int(yr_min), int(yr_max) + 1)
            for e in emps:
                total = sum(get_monthly_salary(e['id'], y, m)
                            for y in years_range for m in range(1, 13))
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE employee_id=?""", (e['id'],)).fetchone()['s']
                unpaid = total - paid
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.sal_person_tree.insert('', 'end', values=(
                    name, fmt_money(total), fmt_money(paid), fmt_money(unpaid)))
        else:
            year = int(year_sel)
            for e in emps:
                total = sum(get_monthly_salary(e['id'], year, m) for m in range(1, 13))
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE employee_id=? AND year=?""",
                                    (e['id'], year)).fetchone()['s']
                unpaid = total - paid
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.sal_person_tree.insert('', 'end', values=(
                    name, fmt_money(total), fmt_money(paid), fmt_money(unpaid)))

        # === 工资按月 ===
        for i in self.sal_month_tree.get_children():
            self.sal_month_tree.delete(i)
        if is_all:
            # 全部年份：按年-月分组
            yr_min = conn.execute("SELECT MIN(year) FROM salary_payments").fetchone()[0]
            yr_max = conn.execute("SELECT MAX(year) FROM salary_payments").fetchone()[0]
            if not yr_min or not yr_max:
                yr_min = yr_max = datetime.now().year
            years_range = range(int(yr_min), int(yr_max) + 1)
            for y in years_range:
                for m in range(1, 13):
                    monthly_total = sum(get_monthly_salary(e['id'], y, m) for e in emps)
                    if monthly_total == 0:
                        continue
                    paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                           WHERE year=? AND month=?""", (y, m)).fetchone()['s']
                    unpaid = monthly_total - paid
                    self.sal_month_tree.insert('', 'end', values=(
                        f"{y}-{m:02d}", fmt_money(monthly_total), fmt_money(paid), fmt_money(unpaid)))
        else:
            year = int(year_sel)
            for m in range(1, 13):
                monthly_total = sum(get_monthly_salary(e['id'], year, m) for e in emps)
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE year=? AND month=?""", (year, m)).fetchone()['s']
                unpaid = monthly_total - paid
                self.sal_month_tree.insert('', 'end', values=(
                    f"{year}-{m:02d}", fmt_money(monthly_total), fmt_money(paid), fmt_money(unpaid)))

        # ========== 五险一金统计 ==========
        for i in self.ins_person_tree.get_children():
            self.ins_person_tree.delete(i)
        for i in self.ins_month_tree.get_children():
            self.ins_month_tree.delete(i)

        emps_ins = conn.execute("SELECT * FROM employees ORDER BY name").fetchall()

        if is_all:
            # 全部年份：从最早入职到当前年
            yr_min_ins = _get_min_hire_year()
            yr_max_ins = datetime.now().year
            years_range_ins = range(yr_min_ins, yr_max_ins + 1)
            for e in emps_ins:
                comp_total = pers_total = 0
                for y in years_range_ins:
                    for m in range(1, 13):
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            comp_total += sum(d['company'] for d in det.values())
                            pers_total += sum(d['personal'] for d in det.values())
                if comp_total == 0 and pers_total == 0:
                    continue
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.ins_person_tree.insert('', 'end', values=(
                    name, fmt_money(comp_total), fmt_money(pers_total),
                    fmt_money(comp_total + pers_total)))
            # 按月
            for y in years_range_ins:
                for m in range(1, 13):
                    m_comp = m_pers = 0
                    for e in emps_ins:
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            m_comp += sum(d['company'] for d in det.values())
                            m_pers += sum(d['personal'] for d in det.values())
                    if m_comp == 0 and m_pers == 0:
                        continue
                    self.ins_month_tree.insert('', 'end', values=(
                        f"{y}-{m:02d}", fmt_money(m_comp), fmt_money(m_pers),
                        fmt_money(m_comp + m_pers)))
        else:
            year_ins = int(year_sel)
            for e in emps_ins:
                comp_total = pers_total = 0
                for m in range(1, 13):
                    det = get_monthly_insurance_detail(e['id'], year_ins, m)
                    if det:
                        comp_total += sum(d['company'] for d in det.values())
                        pers_total += sum(d['personal'] for d in det.values())
                if comp_total == 0 and pers_total == 0:
                    continue
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.ins_person_tree.insert('', 'end', values=(
                    name, fmt_money(comp_total), fmt_money(pers_total),
                    fmt_money(comp_total + pers_total)))
            # 按月
            for m in range(1, 13):
                m_comp = m_pers = 0
                for e in emps_ins:
                    det = get_monthly_insurance_detail(e['id'], year_ins, m)
                    if det:
                        m_comp += sum(d['company'] for d in det.values())
                        m_pers += sum(d['personal'] for d in det.values())
                if m_comp == 0 and m_pers == 0:
                    continue
                self.ins_month_tree.insert('', 'end', values=(
                    f"{year_ins}-{m:02d}", fmt_money(m_comp), fmt_money(m_pers),
                    fmt_money(m_comp + m_pers)))

        conn.close()

    def export_excel(self):
        """导出统计汇总到Excel（发票统计+工资统计，共4个Sheet）"""
        year_sel = self.year_var.get()
        is_all = (year_sel == '全部')
        if not is_all:
            try:
                year = int(year_sel)
            except ValueError:
                messagebox.showwarning("提示", "请先选择统计年份")
                return

        conn = get_db()

        # 发票按人
        if is_all:
            inv_person = conn.execute("""SELECT e.name,
                COALESCE(SUM(CASE WHEN i.type='支付记录' THEN i.amount ELSE 0 END),0) pay_amt,
                COALESCE(SUM(CASE WHEN COALESCE(i.type,'发票')='发票' THEN i.amount ELSE 0 END),0) inv_amt,
                COALESCE(SUM(i.amount),0) total,
                COALESCE(SUM(CASE WHEN COALESCE(i.batch_pending,0)=1 THEN i.amount ELSE 0 END),0) pending,
                COALESCE(SUM(CASE WHEN COALESCE(i.status,0)=0 THEN i.amount ELSE 0 END),0) unpaid,
                COALESCE(SUM(CASE WHEN i.status=1 THEN i.amount ELSE 0 END),0) paid
                FROM employees e LEFT JOIN invoices i ON i.reimburser_id=e.id
                GROUP BY e.id ORDER BY e.name""").fetchall()
        else:
            inv_person = conn.execute("""SELECT e.name,
                COALESCE(SUM(CASE WHEN i.type='支付记录' THEN i.amount ELSE 0 END),0) pay_amt,
                COALESCE(SUM(CASE WHEN COALESCE(i.type,'发票')='发票' THEN i.amount ELSE 0 END),0) inv_amt,
                COALESCE(SUM(i.amount),0) total,
                COALESCE(SUM(CASE WHEN COALESCE(i.batch_pending,0)=1 THEN i.amount ELSE 0 END),0) pending,
                COALESCE(SUM(CASE WHEN COALESCE(i.status,0)=0 THEN i.amount ELSE 0 END),0) unpaid,
                COALESCE(SUM(CASE WHEN i.status=1 THEN i.amount ELSE 0 END),0) paid
                FROM employees e LEFT JOIN invoices i ON i.reimburser_id=e.id
                AND strftime('%Y', i.invoice_date)=?
                GROUP BY e.id ORDER BY e.name""", (str(year),)).fetchall()

        # 发票按月
        if is_all:
            inv_month = conn.execute("""SELECT strftime('%Y-%m', invoice_date) ym,
                COALESCE(SUM(CASE WHEN type='支付记录' THEN amount ELSE 0 END),0) pay_amt,
                COALESCE(SUM(CASE WHEN COALESCE(type,'发票')='发票' THEN amount ELSE 0 END),0) inv_amt,
                COALESCE(SUM(amount),0) total,
                COALESCE(SUM(CASE WHEN COALESCE(batch_pending,0)=1 THEN amount ELSE 0 END),0) pending,
                COALESCE(SUM(CASE WHEN COALESCE(status,0)=0 THEN amount ELSE 0 END),0) unpaid,
                COALESCE(SUM(CASE WHEN status=1 THEN amount ELSE 0 END),0) paid
                FROM invoices GROUP BY ym ORDER BY ym""").fetchall()
        else:
            inv_month = conn.execute("""SELECT strftime('%Y-%m', invoice_date) ym,
                COALESCE(SUM(CASE WHEN type='支付记录' THEN amount ELSE 0 END),0) pay_amt,
                COALESCE(SUM(CASE WHEN COALESCE(type,'发票')='发票' THEN amount ELSE 0 END),0) inv_amt,
                COALESCE(SUM(amount),0) total,
                COALESCE(SUM(CASE WHEN COALESCE(batch_pending,0)=1 THEN amount ELSE 0 END),0) pending,
                COALESCE(SUM(CASE WHEN COALESCE(status,0)=0 THEN amount ELSE 0 END),0) unpaid,
                COALESCE(SUM(CASE WHEN status=1 THEN amount ELSE 0 END),0) paid
                FROM invoices WHERE strftime('%Y', invoice_date)=?
                GROUP BY ym ORDER BY ym""", (str(year),)).fetchall()

        # 工资按人年度
        emps = conn.execute("SELECT * FROM employees ORDER BY name").fetchall()
        sal_person = []
        if is_all:
            yr_min = conn.execute("SELECT MIN(year) FROM salary_payments").fetchone()[0]
            yr_max = conn.execute("SELECT MAX(year) FROM salary_payments").fetchone()[0]
            if not yr_min or not yr_max:
                yr_min = yr_max = datetime.now().year
            years_range = range(int(yr_min), int(yr_max) + 1)
            for e in emps:
                total = sum(get_monthly_salary(e['id'], y, m)
                            for y in years_range for m in range(1, 13))
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE employee_id=?""", (e['id'],)).fetchone()['s']
                sal_person.append({'name': e['name'], 'total': total, 'paid': paid, 'unpaid': total - paid})
        else:
            for e in emps:
                total = sum(get_monthly_salary(e['id'], year, m) for m in range(1, 13))
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE employee_id=? AND year=?""",
                                    (e['id'], year)).fetchone()['s']
                sal_person.append({'name': e['name'], 'total': total, 'paid': paid, 'unpaid': total - paid})

        # 工资按月
        sal_month = []
        if is_all:
            yr_min = conn.execute("SELECT MIN(year) FROM salary_payments").fetchone()[0]
            yr_max = conn.execute("SELECT MAX(year) FROM salary_payments").fetchone()[0]
            if not yr_min or not yr_max:
                yr_min = yr_max = datetime.now().year
            years_range = range(int(yr_min), int(yr_max) + 1)
            for y in years_range:
                for m in range(1, 13):
                    monthly_total = sum(get_monthly_salary(e['id'], y, m) for e in emps)
                    if monthly_total == 0:
                        continue
                    paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                           WHERE year=? AND month=?""", (y, m)).fetchone()['s']
                    sal_month.append({'month': f"{y}-{m:02d}", 'total': monthly_total,
                                      'paid': paid, 'unpaid': monthly_total - paid})
        else:
            for m in range(1, 13):
                monthly_total = sum(get_monthly_salary(e['id'], year, m) for e in emps)
                paid = conn.execute("""SELECT COALESCE(SUM(amount),0) s FROM salary_payments
                                       WHERE year=? AND month=?""", (year, m)).fetchone()['s']
                sal_month.append({'month': f"{year}-{m:02d}", 'total': monthly_total,
                                  'paid': paid, 'unpaid': monthly_total - paid})

        # 五险一金统计数据
        ins_person = []
        ins_month = []
        if is_all:
            yr_min_ins = _get_min_hire_year()
            yr_max_ins = datetime.now().year
            years_range_ins = range(yr_min_ins, yr_max_ins + 1)
            for e in emps:
                comp_t = pers_t = 0
                for y in years_range_ins:
                    for m in range(1, 13):
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            comp_t += sum(d['company'] for d in det.values())
                            pers_t += sum(d['personal'] for d in det.values())
                if comp_t > 0 or pers_t > 0:
                    ins_person.append({'name': e['name'], 'company': comp_t,
                                       'personal': pers_t, 'total': comp_t + pers_t})
            for y in years_range_ins:
                for m in range(1, 13):
                    mc = mp = 0
                    for e in emps:
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            mc += sum(d['company'] for d in det.values())
                            mp += sum(d['personal'] for d in det.values())
                    if mc > 0 or mp > 0:
                        ins_month.append({'month': f"{y}-{m:02d}", 'company': mc,
                                          'personal': mp, 'total': mc + mp})
        else:
            for e in emps:
                comp_t = pers_t = 0
                for m in range(1, 13):
                    det = get_monthly_insurance_detail(e['id'], year, m)
                    if det:
                        comp_t += sum(d['company'] for d in det.values())
                        pers_t += sum(d['personal'] for d in det.values())
                if comp_t > 0 or pers_t > 0:
                    ins_person.append({'name': e['name'], 'company': comp_t,
                                       'personal': pers_t, 'total': comp_t + pers_t})
            for m in range(1, 13):
                mc = mp = 0
                for e in emps:
                    det = get_monthly_insurance_detail(e['id'], year, m)
                    if det:
                        mc += sum(d['company'] for d in det.values())
                        mp += sum(d['personal'] for d in det.values())
                if mc > 0 or mp > 0:
                    ins_month.append({'month': f"{year}-{m:02d}", 'company': mc,
                                      'personal': mp, 'total': mc + mp})
        conn.close()

        year_label = "全部" if is_all else f"{year}年度"
        default_name = f"统计汇总_{year_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出统计Excel", defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return

        wb = Workbook()

        def _write_sheet(ws, title, headers, data_rows, total_cols=None):
            ncols = len(headers)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
            ws['A1'] = title
            ws['A1'].font = TITLE_FONT
            ws['A1'].alignment = CENTER
            ws.row_dimensions[1].height = 26
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
            ws['A2'] = f"统计年份：{year}年    导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            ws['A2'].font = Font(name='微软雅黑', size=9, color='666666')
            ws['A2'].alignment = LEFT
            for c, h in enumerate(headers, 1):
                ws.cell(row=3, column=c, value=h)
            _style_header(ws, 3, ncols)
            for idx, row in enumerate(data_rows, 4):
                for c, val in enumerate(row, 1):
                    cell = ws.cell(row=idx, column=c, value=val)
                    cell.border = THIN_BORDER
                    if c == 1:
                        cell.alignment = CENTER
                    elif isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00'
                        cell.alignment = RIGHT
                    else:
                        cell.alignment = CENTER
            if total_cols:
                tr = len(data_rows) + 4
                ws.cell(row=tr, column=1, value="合计").font = TOTAL_FONT
                ws.cell(row=tr, column=1).fill = TOTAL_FILL
                ws.cell(row=tr, column=1).alignment = CENTER
                ws.cell(row=tr, column=1).border = THIN_BORDER
                for ci in total_cols:
                    s = sum(float(r[ci - 1]) for r in data_rows if isinstance(r[ci - 1], (int, float)))
                    cell = ws.cell(row=tr, column=ci, value=s)
                    cell.number_format = '#,##0.00'
                    cell.font = TOTAL_FONT
                    cell.fill = TOTAL_FILL
                    cell.alignment = RIGHT
                    cell.border = THIN_BORDER
                for c in range(1, ncols + 1):
                    ws.cell(row=tr, column=c).border = THIN_BORDER
                    if c not in total_cols and c != 1:
                        ws.cell(row=tr, column=c).fill = TOTAL_FILL
            _auto_width(ws, ncols)
            ws.freeze_panes = 'A4'

        # Sheet1: 发票-按人员
        ws1 = wb.active
        ws1.title = "发票-按人员"
        rows1 = [(r['name'], float(r['inv_amt']), float(r['pay_amt']),
                  float(r['total']), float(r['pending']), float(r['unpaid']), float(r['paid'])) for r in inv_person]
        _write_sheet(ws1, f"{year_label} 发票报销统计（按人员，含发票+支付记录）",
                     ['姓名', '发票金额', '支付记录', '报销总额', '待完善', '未报销', '已报销'],
                     rows1, total_cols=[2, 3, 4, 5, 6, 7])

        # Sheet2: 发票-按月
        ws2 = wb.create_sheet("发票-按月")
        rows2 = [(r['ym'] or '-', float(r['inv_amt']), float(r['pay_amt']),
                  float(r['total']), float(r['pending']), float(r['unpaid']), float(r['paid'])) for r in inv_month]
        _write_sheet(ws2, f"{year_label} 发票报销统计（按月，含发票+支付记录）",
                     ['月份', '发票金额', '支付记录', '报销总额', '待完善', '未报销', '已报销'],
                     rows2, total_cols=[2, 3, 4, 5, 6, 7])

        # Sheet3: 工资-按人员年度
        ws3 = wb.create_sheet("工资-按人员年度")
        rows3 = [(r['name'], float(r['total']), float(r['paid']), float(r['unpaid'])) for r in sal_person]
        _write_sheet(ws3, f"{year_label} 工资结算统计（按人员）",
                     ['姓名', '年度应发', '已结', '未结'], rows3, total_cols=[2, 3, 4])

        # Sheet4: 工资-按月度
        ws4 = wb.create_sheet("工资-按月度")
        rows4 = [(r['month'], float(r['total']), float(r['paid']), float(r['unpaid'])) for r in sal_month]
        _write_sheet(ws4, f"{year_label} 工资结算统计（按月度）",
                     ['月份', '应发总额', '已结', '未结'], rows4, total_cols=[2, 3, 4])

        # Sheet5: 五险一金-按人员
        ws5 = wb.create_sheet("五险一金-按人员")
        rows5 = [(r['name'], float(r['company']), float(r['personal']), float(r['total'])) for r in ins_person]
        _write_sheet(ws5, f"{year_label} 五险一金统计（按人员，公司+个人）",
                     ['姓名', '公司缴纳', '个人缴纳', '合计'], rows5, total_cols=[2, 3, 4])

        # Sheet6: 五险一金-按月度
        ws6 = wb.create_sheet("五险一金-按月度")
        rows6 = [(r['month'], float(r['company']), float(r['personal']), float(r['total'])) for r in ins_month]
        _write_sheet(ws6, f"{year_label} 五险一金统计（按月度，公司+个人）",
                     ['月份', '公司缴纳', '个人缴纳', '合计'], rows6, total_cols=[2, 3, 4])

        try:
            wb.save(path)
            self.app.set_status(f"已导出统计Excel: {os.path.basename(path)}")
            messagebox.showinfo("导出成功", f"已导出 {year_label} 统计汇总到：\n{path}\n\n包含6个工作表：发票按人员、发票按月、工资按人员、工资按月度、五险一金按人员、五险一金按月度")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def import_excel(self):
        """统计页导入提示：统计数据为自动计算结果，请到对应页面导入原始数据"""
        messagebox.showinfo("导入说明",
            "统计汇总页面的数据是根据发票报销和工资发放记录自动计算生成的，\n"
            "无法直接导入统计数据。\n\n"
            "如需批量导入数据，请切换到以下页面操作：\n"
            "• 发票报销 → 导入Excel（导入发票报销明细）\n"
            "• 工资结算 → 导入Excel（导入工资发放记录）\n"
            "• 银行账户 → 导入Excel（导入银行交易明细）\n\n"
            "导入原始数据后，统计汇总页面会自动更新。")


# ============================================================
# 银行账户管理
# ============================================================
class BankTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.selected_account_id = None
        self.editing_account_id = None
        self.editing_tx_id = None

        # 顶部：账户管理
        top = ttk.LabelFrame(self.content, text="账户管理")
        top.pack(fill='x', padx=10, pady=6)

        form = ttk.Frame(top)
        form.pack(fill='x', padx=8, pady=6)
        ttk.Label(form, text="账户名称:").grid(row=0, column=0, sticky='e', padx=4, pady=4)
        self.acc_name_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.acc_name_var, width=18).grid(row=0, column=1, padx=4, pady=4)
        ttk.Label(form, text="账号:").grid(row=0, column=2, sticky='e', padx=4, pady=4)
        self.acc_num_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.acc_num_var, width=20).grid(row=0, column=3, padx=4, pady=4)
        ttk.Label(form, text="开户行:").grid(row=0, column=4, sticky='e', padx=4, pady=4)
        self.acc_bank_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.acc_bank_var, width=16).grid(row=0, column=5, padx=4, pady=4)
        ttk.Label(form, text="备注:").grid(row=1, column=0, sticky='e', padx=4, pady=4)
        self.acc_remark_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.acc_remark_var, width=50).grid(row=1, column=1, columnspan=4, sticky='we', padx=4, pady=4)
        self.acc_add_btn = ttk.Button(form, text="添加账户", command=self.add_or_update_account)
        self.acc_add_btn.grid(row=1, column=5, padx=4, pady=4)
        self.acc_cancel_btn = ttk.Button(form, text="取消", command=self.cancel_account_edit, state='disabled')
        self.acc_cancel_btn.grid(row=0, column=6, padx=4, pady=4)

        # 账户列表
        acc_frame = ttk.LabelFrame(self.content, text="账户列表（双击可修改）")
        acc_frame.pack(fill='x', padx=10, pady=4)
        acc_tbl = ttk.Frame(acc_frame)
        acc_tbl.pack(fill='x', padx=4, pady=4)
        acc_cols = ('id', 'name', 'account_number', 'bank_name', 'balance', 'remark')
        self.acc_tree = ttk.Treeview(acc_tbl, columns=acc_cols, show='headings', height=4)
        acc_headers = [('id', '编号', 50, 'center'), ('name', '账户名称', 150, 'w'),
                       ('account_number', '账号', 180, 'center'), ('bank_name', '开户行', 150, 'w'),
                       ('balance', '当前余额', 120, 'e'), ('remark', '备注', 200, 'w')]
        for cid, text, w, anchor in acc_headers:
            self.acc_tree.heading(cid, text=text)
            self.acc_tree.column(cid, width=w, anchor=anchor)
        self.acc_tree.pack(side='left', fill='x', expand=True)
        acc_sb = ttk.Scrollbar(acc_tbl, orient='vertical', command=self.acc_tree.yview)
        acc_sb.pack(side='right', fill='y')
        self.acc_tree.configure(yscrollcommand=acc_sb.set)
        self.acc_tree.bind('<<TreeviewSelect>>', self.on_account_select)
        self.acc_tree.bind('<Double-1>', self.on_account_double)

        acc_btns = ttk.Frame(acc_frame)
        acc_btns.pack(fill='x', padx=4, pady=2)
        ttk.Button(acc_btns, text="删除选中账户", command=self.delete_account).pack(side='left', padx=4)
        ttk.Button(acc_btns, text="刷新", command=self.refresh_accounts).pack(side='left', padx=4)

        # 交易记录
        tx_frame = ttk.LabelFrame(self.content, text="交易记录")
        tx_frame.pack(fill='both', expand=True, padx=10, pady=4)

        # 交易表单
        tx_form = ttk.Frame(tx_frame)
        tx_form.pack(fill='x', padx=8, pady=4)
        ttk.Label(tx_form, text="方向:").grid(row=0, column=0, sticky='e', padx=4, pady=3)
        self.tx_dir_var = tk.StringVar(value='出项')
        ttk.Combobox(tx_form, textvariable=self.tx_dir_var, values=['进项', '出项'],
                     width=8, state='readonly').grid(row=0, column=1, padx=4, pady=3)
        ttk.Label(tx_form, text="金额:").grid(row=0, column=2, sticky='e', padx=4, pady=3)
        self.tx_amount_var = tk.StringVar()
        ttk.Entry(tx_form, textvariable=self.tx_amount_var, width=12).grid(row=0, column=3, padx=4, pady=3)
        ttk.Label(tx_form, text="日期:").grid(row=0, column=4, sticky='e', padx=4, pady=3)
        self.tx_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(tx_form, textvariable=self.tx_date_var, width=12).grid(row=0, column=5, padx=4, pady=3)
        ttk.Label(tx_form, text="用途:").grid(row=0, column=6, sticky='e', padx=4, pady=3)
        self.tx_purpose_var = tk.StringVar()
        ttk.Entry(tx_form, textvariable=self.tx_purpose_var, width=24).grid(row=0, column=7, padx=4, pady=3)
        self.tx_add_btn = ttk.Button(tx_form, text="添加交易", command=self.add_or_update_tx)
        self.tx_add_btn.grid(row=0, column=8, padx=4, pady=3)
        self.tx_cancel_btn = ttk.Button(tx_form, text="取消", command=self.cancel_tx_edit, state='disabled')
        self.tx_cancel_btn.grid(row=0, column=9, padx=4, pady=3)

        # 交易列表
        tx_tbl = ttk.Frame(tx_frame)
        tx_tbl.pack(fill='both', expand=True, padx=4, pady=4)
        tx_cols = ('id', 'transaction_date', 'direction', 'amount', 'purpose', 'source', 'balance_after')
        self.tx_tree = ttk.Treeview(tx_tbl, columns=tx_cols, show='headings', height=8)
        tx_headers = [('id', '编号', 50, 'center'), ('transaction_date', '日期', 100, 'center'),
                      ('direction', '方向', 70, 'center'), ('amount', '金额', 120, 'e'),
                      ('purpose', '用途/摘要', 250, 'w'), ('source', '来源', 100, 'center'),
                      ('balance_after', '交易后余额', 120, 'e')]
        for cid, text, w, anchor in tx_headers:
            self.tx_tree.heading(cid, text=text)
            self.tx_tree.column(cid, width=w, anchor=anchor)
        self.tx_tree.pack(side='left', fill='both', expand=True)
        tx_sb = ttk.Scrollbar(tx_tbl, orient='vertical', command=self.tx_tree.yview)
        tx_sb.pack(side='right', fill='y')
        self.tx_tree.configure(yscrollcommand=tx_sb.set)
        self.tx_tree.bind('<Double-1>', self.on_tx_double)

        tx_btns = ttk.Frame(tx_frame)
        tx_btns.pack(fill='x', padx=4, pady=2)
        ttk.Button(tx_btns, text="删除选中交易", command=self.delete_tx).pack(side='left', padx=4)
        ttk.Button(tx_btns, text="导入Excel", command=self.import_excel).pack(side='right', padx=4)
        ttk.Button(tx_btns, text="导出Excel", command=self.export_excel).pack(side='right', padx=4)

        # 统计汇总
        stat_frame = ttk.LabelFrame(self.content, text="账户统计（逐月汇总 / 逐年累计）")
        stat_frame.pack(fill='x', padx=10, pady=4)
        stat_wrap = ttk.Frame(stat_frame)
        stat_wrap.pack(fill='x', padx=4, pady=4)
        stat_cols = ('period', 'income', 'expense', 'net')
        self.stat_tree = ttk.Treeview(stat_wrap, columns=stat_cols, show='headings', height=5)
        stat_headers = [('period', '期间', 120, 'center'), ('income', '进项合计', 150, 'e'),
                        ('expense', '出项合计', 150, 'e'), ('net', '净额', 150, 'e')]
        for cid, text, w, anchor in stat_headers:
            self.stat_tree.heading(cid, text=text)
            self.stat_tree.column(cid, width=w, anchor=anchor)
        self.stat_tree.pack(side='left', fill='x', expand=True)
        stat_sb = ttk.Scrollbar(stat_wrap, orient='vertical', command=self.stat_tree.yview)
        stat_sb.pack(side='right', fill='y')
        self.stat_tree.configure(yscrollcommand=stat_sb.set)

        self.refresh_accounts()

    def _get_balance(self, account_id):
        """计算账户当前余额"""
        conn = get_db()
        row = conn.execute("""SELECT
            COALESCE(SUM(CASE WHEN direction='进项' THEN amount ELSE 0 END),0) as income,
            COALESCE(SUM(CASE WHEN direction='出项' THEN amount ELSE 0 END),0) as expense
            FROM bank_transactions WHERE account_id=?""", (account_id,)).fetchone()
        conn.close()
        return (row['income'] or 0) - (row['expense'] or 0)

    def refresh_accounts(self):
        for i in self.acc_tree.get_children():
            self.acc_tree.delete(i)
        conn = get_db()
        rows = conn.execute("SELECT * FROM bank_accounts ORDER BY id").fetchall()
        conn.close()
        for r in rows:
            bal = self._get_balance(r['id'])
            self.acc_tree.insert('', 'end', values=(
                r['id'], r['name'], r['account_number'] or '', r['bank_name'] or '',
                fmt_money(bal), r['remark'] or ''))

    def on_account_select(self, event=None):
        sel = self.acc_tree.selection()
        if sel:
            self.selected_account_id = self.acc_tree.item(sel[0], 'values')[0]
            self.refresh_transactions()
            self.refresh_stats()

    def on_account_double(self, event=None):
        sel = self.acc_tree.selection()
        if not sel:
            return
        self.editing_account_id = self.acc_tree.item(sel[0], 'values')[0]
        conn = get_db()
        row = conn.execute("SELECT * FROM bank_accounts WHERE id=?", (self.editing_account_id,)).fetchone()
        conn.close()
        if not row:
            return
        self.acc_name_var.set(row['name'] or '')
        self.acc_num_var.set(row['account_number'] or '')
        self.acc_bank_var.set(row['bank_name'] or '')
        self.acc_remark_var.set(row['remark'] or '')
        self.acc_add_btn.config(text="保存修改")
        self.acc_cancel_btn.config(state='normal')

    def cancel_account_edit(self):
        self.editing_account_id = None
        self.acc_name_var.set('')
        self.acc_num_var.set('')
        self.acc_bank_var.set('')
        self.acc_remark_var.set('')
        self.acc_add_btn.config(text="添加账户")
        self.acc_cancel_btn.config(state='disabled')

    def add_or_update_account(self):
        name = self.acc_name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入账户名称")
            return
        acc_num = self.acc_num_var.get().strip()
        bank = self.acc_bank_var.get().strip()
        remark = self.acc_remark_var.get().strip()
        conn = get_db()
        if self.editing_account_id:
            conn.execute("""UPDATE bank_accounts SET name=?, account_number=?, bank_name=?, remark=?
                         WHERE id=?""", (name, acc_num, bank, remark, self.editing_account_id))
            self.app.set_status(f"已修改账户: {name}")
        else:
            conn.execute("""INSERT INTO bank_accounts(name, account_number, bank_name, remark, created_at)
                          VALUES(?,?,?,?,?)""", (name, acc_num, bank, remark,
                          datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.app.set_status(f"已添加账户: {name}")
        conn.commit()
        conn.close()
        self.cancel_account_edit()
        self.refresh_accounts()

    def delete_account(self):
        sel = self.acc_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的账户")
            return
        acc_id = self.acc_tree.item(sel[0], 'values')[0]
        acc_name = self.acc_tree.item(sel[0], 'values')[1]
        conn = get_db()
        tx_cnt = conn.execute("SELECT COUNT(*) c FROM bank_transactions WHERE account_id=?", (acc_id,)).fetchone()['c']
        conn.close()
        msg = f"确定删除账户「{acc_name}」？"
        if tx_cnt:
            msg += f"\n\n该账户有 {tx_cnt} 条交易记录，删除后交易记录也将一并删除。"
        if not messagebox.askyesno("确认删除", msg):
            return
        conn = get_db()
        conn.execute("DELETE FROM bank_transactions WHERE account_id=?", (acc_id,))
        conn.execute("DELETE FROM bank_accounts WHERE id=?", (acc_id,))
        conn.commit()
        conn.close()
        self.selected_account_id = None
        self.refresh_accounts()
        self.refresh_transactions()
        self.refresh_stats()
        self.app.set_status(f"已删除账户: {acc_name}")

    def refresh_transactions(self):
        for i in self.tx_tree.get_children():
            self.tx_tree.delete(i)
        if not self.selected_account_id:
            return
        conn = get_db()
        rows = conn.execute("""SELECT * FROM bank_transactions WHERE account_id=?
                             ORDER BY transaction_date DESC, id DESC""", (self.selected_account_id,)).fetchall()
        conn.close()
        # 计算每笔交易后的余额（按时间正序累计，再倒序显示）
        balance = 0
        balances = {}
        for r in sorted(rows, key=lambda x: (x['transaction_date'] or '', x['id'])):
            if r['direction'] == '进项':
                balance += r['amount']
            else:
                balance -= r['amount']
            balances[r['id']] = balance
        for r in rows:
            source_text = r['source'] or '手动'
            self.tx_tree.insert('', 'end', values=(
                r['id'], r['transaction_date'] or '', r['direction'],
                fmt_money(r['amount']), r['purpose'] or '', source_text,
                fmt_money(balances.get(r['id'], 0))))

    def on_tx_double(self, event=None):
        sel = self.tx_tree.selection()
        if not sel:
            return
        self.editing_tx_id = self.tx_tree.item(sel[0], 'values')[0]
        conn = get_db()
        row = conn.execute("SELECT * FROM bank_transactions WHERE id=?", (self.editing_tx_id,)).fetchone()
        conn.close()
        if not row:
            return
        self.tx_dir_var.set(row['direction'])
        self.tx_amount_var.set(str(row['amount'] or 0))
        self.tx_date_var.set(row['transaction_date'] or '')
        self.tx_purpose_var.set(row['purpose'] or '')
        self.tx_add_btn.config(text="保存修改")
        self.tx_cancel_btn.config(state='normal')

    def cancel_tx_edit(self):
        self.editing_tx_id = None
        self.tx_dir_var.set('出项')
        self.tx_amount_var.set('')
        self.tx_date_var.set(datetime.now().strftime('%Y-%m-%d'))
        self.tx_purpose_var.set('')
        self.tx_add_btn.config(text="添加交易")
        self.tx_cancel_btn.config(state='disabled')

    def add_or_update_tx(self):
        if not self.selected_account_id:
            messagebox.showwarning("提示", "请先选择一个银行账户")
            return
        direction = self.tx_dir_var.get()
        amount = self.tx_amount_var.get().strip()
        tx_date = self.tx_date_var.get().strip()
        purpose = self.tx_purpose_var.get().strip()
        try:
            amount = float(amount) if amount else 0
        except ValueError:
            messagebox.showwarning("提示", "金额必须是数字")
            return
        if amount <= 0:
            messagebox.showwarning("提示", "金额必须大于0")
            return
        conn = get_db()
        if self.editing_tx_id:
            conn.execute("""UPDATE bank_transactions SET direction=?, amount=?, transaction_date=?,
                         purpose=? WHERE id=?""", (direction, amount, tx_date, purpose, self.editing_tx_id))
            self.app.set_status("已修改交易记录")
        else:
            conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                          source, transaction_date, created_at) VALUES(?,?,?,?,?,?,?)""",
                         (self.selected_account_id, direction, amount, purpose, '手动', tx_date,
                          datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.app.set_status(f"已添加{direction}交易: {fmt_money(amount)}")
        conn.commit()
        conn.close()
        self.cancel_tx_edit()
        self.refresh_transactions()
        self.refresh_accounts()
        self.refresh_stats()

    def delete_tx(self):
        sel = self.tx_tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择要删除的交易")
            return
        tx_id = self.tx_tree.item(sel[0], 'values')[0]
        if not messagebox.askyesno("确认", "确定删除该交易记录？"):
            return
        conn = get_db()
        conn.execute("DELETE FROM bank_transactions WHERE id=?", (tx_id,))
        conn.commit()
        conn.close()
        self.refresh_transactions()
        self.refresh_accounts()
        self.refresh_stats()
        self.app.set_status("已删除交易记录")

    def refresh_stats(self):
        for i in self.stat_tree.get_children():
            self.stat_tree.delete(i)
        if not self.selected_account_id:
            return
        conn = get_db()
        # 逐月汇总
        rows = conn.execute("""SELECT substr(transaction_date,1,7) as ym,
            COALESCE(SUM(CASE WHEN direction='进项' THEN amount ELSE 0 END),0) as income,
            COALESCE(SUM(CASE WHEN direction='出项' THEN amount ELSE 0 END),0) as expense
            FROM bank_transactions WHERE account_id=? AND transaction_date IS NOT NULL
            GROUP BY ym ORDER BY ym DESC""", (self.selected_account_id,)).fetchall()
        for r in rows:
            net = (r['income'] or 0) - (r['expense'] or 0)
            self.stat_tree.insert('', 'end', values=(
                f"{r['ym']} 月", fmt_money(r['income']), fmt_money(r['expense']), fmt_money(net)))
        # 逐年累计
        year_rows = conn.execute("""SELECT substr(transaction_date,1,4) as yr,
            COALESCE(SUM(CASE WHEN direction='进项' THEN amount ELSE 0 END),0) as income,
            COALESCE(SUM(CASE WHEN direction='出项' THEN amount ELSE 0 END),0) as expense
            FROM bank_transactions WHERE account_id=? AND transaction_date IS NOT NULL
            GROUP BY yr ORDER BY yr DESC""", (self.selected_account_id,)).fetchall()
        for r in year_rows:
            net = (r['income'] or 0) - (r['expense'] or 0)
            self.stat_tree.insert('', 'end', values=(
                f"{r['yr']} 年累计", fmt_money(r['income']), fmt_money(r['expense']), fmt_money(net)))
        # ========== 五险一金统计 ==========
        for i in self.ins_person_tree.get_children():
            self.ins_person_tree.delete(i)
        for i in self.ins_month_tree.get_children():
            self.ins_month_tree.delete(i)

        emps_ins = conn.execute("SELECT * FROM employees ORDER BY name").fetchall()

        if is_all:
            # 全部年份：从最早入职到当前年
            yr_min_ins = _get_min_hire_year()
            yr_max_ins = datetime.now().year
            years_range_ins = range(yr_min_ins, yr_max_ins + 1)
            for e in emps_ins:
                comp_total = pers_total = 0
                for y in years_range_ins:
                    for m in range(1, 13):
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            comp_total += sum(d['company'] for d in det.values())
                            pers_total += sum(d['personal'] for d in det.values())
                if comp_total == 0 and pers_total == 0:
                    continue
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.ins_person_tree.insert('', 'end', values=(
                    name, fmt_money(comp_total), fmt_money(pers_total),
                    fmt_money(comp_total + pers_total)))
            # 按月
            for y in years_range_ins:
                for m in range(1, 13):
                    m_comp = m_pers = 0
                    for e in emps_ins:
                        det = get_monthly_insurance_detail(e['id'], y, m)
                        if det:
                            m_comp += sum(d['company'] for d in det.values())
                            m_pers += sum(d['personal'] for d in det.values())
                    if m_comp == 0 and m_pers == 0:
                        continue
                    self.ins_month_tree.insert('', 'end', values=(
                        f"{y}-{m:02d}", fmt_money(m_comp), fmt_money(m_pers),
                        fmt_money(m_comp + m_pers)))
        else:
            year_ins = int(year_sel)
            for e in emps_ins:
                comp_total = pers_total = 0
                for m in range(1, 13):
                    det = get_monthly_insurance_detail(e['id'], year_ins, m)
                    if det:
                        comp_total += sum(d['company'] for d in det.values())
                        pers_total += sum(d['personal'] for d in det.values())
                if comp_total == 0 and pers_total == 0:
                    continue
                name = e['name']
                if e['resigned']:
                    name += ' (已离职)'
                self.ins_person_tree.insert('', 'end', values=(
                    name, fmt_money(comp_total), fmt_money(pers_total),
                    fmt_money(comp_total + pers_total)))
            # 按月
            for m in range(1, 13):
                m_comp = m_pers = 0
                for e in emps_ins:
                    det = get_monthly_insurance_detail(e['id'], year_ins, m)
                    if det:
                        m_comp += sum(d['company'] for d in det.values())
                        m_pers += sum(d['personal'] for d in det.values())
                if m_comp == 0 and m_pers == 0:
                    continue
                self.ins_month_tree.insert('', 'end', values=(
                    f"{year_ins}-{m:02d}", fmt_money(m_comp), fmt_money(m_pers),
                    fmt_money(m_comp + m_pers)))

        conn.close()

    def export_excel(self):
        if not self.selected_account_id:
            messagebox.showwarning("提示", "请先选择一个银行账户")
            return
        conn = get_db()
        acc = conn.execute("SELECT * FROM bank_accounts WHERE id=?", (self.selected_account_id,)).fetchone()
        rows = conn.execute("""SELECT * FROM bank_transactions WHERE account_id=?
                             ORDER BY transaction_date DESC, id DESC""", (self.selected_account_id,)).fetchall()
        conn.close()
        if not rows:
            messagebox.showinfo("提示", "该账户没有交易记录")
            return
        default_name = f"银行账户_{acc['name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = filedialog.asksaveasfilename(title="导出Excel", defaultextension=".xlsx",
                                            initialfile=default_name, filetypes=[("Excel文件", "*.xlsx")])
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "交易明细"
        NCOLS = 7
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
        ws['A1'] = f"{acc['name']} - 银行账户交易明细"
        ws['A1'].font = TITLE_FONT
        ws['A1'].alignment = CENTER
        ws.row_dimensions[1].height = 28
        info = f"账号：{acc['account_number'] or ''}    开户行：{acc['bank_name'] or ''}    导出时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
        ws['A2'] = info
        ws['A2'].font = Font(name='微软雅黑', size=9, color='666666')
        headers = ['编号', '日期', '方向', '金额(元)', '用途/摘要', '来源', '交易后余额']
        for c, h in enumerate(headers, 1):
            ws.cell(row=3, column=c, value=h)
        _style_header(ws, 3, NCOLS)
        balance = 0
        balances = {}
        for r in sorted(rows, key=lambda x: (x['transaction_date'] or '', x['id'])):
            if r['direction'] == '进项':
                balance += r['amount']
            else:
                balance -= r['amount']
            balances[r['id']] = balance
        total_in = 0
        total_out = 0
        for idx, r in enumerate(rows, 4):
            ws.cell(row=idx, column=1, value=r['id']).alignment = CENTER
            ws.cell(row=idx, column=2, value=r['transaction_date'] or '').alignment = CENTER
            ws.cell(row=idx, column=3, value=r['direction']).alignment = CENTER
            cell_amt = ws.cell(row=idx, column=4, value=float(r['amount']))
            cell_amt.number_format = '#,##0.00'
            cell_amt.alignment = RIGHT
            ws.cell(row=idx, column=5, value=r['purpose'] or '').alignment = LEFT
            ws.cell(row=idx, column=6, value=r['source'] or '手动').alignment = CENTER
            cell_bal = ws.cell(row=idx, column=7, value=float(balances.get(r['id'], 0)))
            cell_bal.number_format = '#,##0.00'
            cell_bal.alignment = RIGHT
            for c in range(1, NCOLS + 1):
                ws.cell(row=idx, column=c).border = THIN_BORDER
            if r['direction'] == '进项':
                total_in += r['amount']
            else:
                total_out += r['amount']
        total_row = len(rows) + 4
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
        ws.cell(row=total_row, column=1, value=f"合计（共{len(rows)}笔）").font = TOTAL_FONT
        ws.cell(row=total_row, column=1).alignment = RIGHT
        ws.cell(row=total_row, column=1).fill = TOTAL_FILL
        ws.cell(row=total_row, column=3, value="进项").font = TOTAL_FONT
        ws.cell(row=total_row, column=3).fill = TOTAL_FILL
        ws.cell(row=total_row, column=3).alignment = CENTER
        cell_in = ws.cell(row=total_row, column=4, value=total_in)
        cell_in.number_format = '#,##0.00'
        cell_in.font = TOTAL_FONT
        cell_in.fill = TOTAL_FILL
        cell_in.alignment = RIGHT
        ws.cell(row=total_row, column=5, value="出项").font = TOTAL_FONT
        ws.cell(row=total_row, column=5).fill = TOTAL_FILL
        ws.cell(row=total_row, column=5).alignment = CENTER
        cell_out = ws.cell(row=total_row, column=6, value=total_out)
        cell_out.number_format = '#,##0.00'
        cell_out.font = TOTAL_FONT
        cell_out.fill = TOTAL_FILL
        cell_out.alignment = RIGHT
        cell_net = ws.cell(row=total_row, column=7, value=total_in - total_out)
        cell_net.number_format = '#,##0.00'
        cell_net.font = TOTAL_FONT
        cell_net.fill = TOTAL_FILL
        cell_net.alignment = RIGHT
        for c in range(1, NCOLS + 1):
            ws.cell(row=total_row, column=c).border = THIN_BORDER
        widths = [8, 12, 8, 14, 30, 10, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A4'
        try:
            wb.save(path)
            self.app.set_status(f"已导出Excel: {os.path.basename(path)}")
            messagebox.showinfo("导出成功", f"已导出 {len(rows)} 条交易记录到：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def import_excel(self):
        """从Excel导入银行交易记录"""
        if not self.selected_account_id:
            messagebox.showwarning("提示", "请先在上方选择一个银行账户")
            return
        path = filedialog.askopenfilename(title="选择要导入的Excel文件",
                                          filetypes=[("Excel文件", "*.xlsx *.xls")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取Excel文件：{e}")
            return

        # 自动识别表头行
        header_row = None
        col_map = {}
        for row_idx in range(1, min(ws.max_row, 10) + 1):
            row_data = [str(ws.cell(row=row_idx, column=c).value or '').strip() for c in range(1, ws.max_column + 1)]
            if '方向' in row_data or '金额' in row_data or '日期' in row_data:
                header_row = row_idx
                for c_idx, val in enumerate(row_data, 1):
                    col_map[val] = c_idx
                break

        if header_row is None:
            messagebox.showerror("导入失败", "未识别到表头行，请确保Excel包含「日期」「方向」「金额」等列名。")
            return

        def get_col(*names):
            for n in names:
                if n in col_map:
                    return col_map[n]
            return None

        col_date = get_col('日期', '交易日期', '记账日期')
        col_dir = get_col('方向', '收支', '类型', '借贷')
        col_amount = get_col('金额', '金额(元)', '交易金额', '发生额')
        col_income = get_col('进项', '收入', '贷方', '转入')
        col_expense = get_col('出项', '支出', '借方', '转出')
        col_purpose = get_col('用途', '摘要', '备注', '说明', '交易用途')

        if col_amount is None and col_income is None and col_expense is None:
            messagebox.showerror("导入失败", "Excel必须包含「金额」列，或「进项」「出项」分列。")
            return

        conn = get_db()
        success = 0
        failed = 0
        errors = []

        for row_idx in range(header_row + 1, ws.max_row + 1):
            def cell_val(col):
                if col is None:
                    return ''
                v = ws.cell(row=row_idx, column=col).value
                return str(v).strip() if v is not None else ''

            # 解析金额和方向
            direction = '出项'
            amount = 0

            if col_amount:
                amt_str = cell_val(col_amount)
                dir_str = cell_val(col_dir)
                try:
                    amount = float(amt_str.replace(',', '').replace('¥', '')) if amt_str else 0
                except ValueError:
                    amount = 0
                if '进' in dir_str or '收' in dir_str or '贷' in dir_str or '转入' in dir_str:
                    direction = '进项'
                elif '出' in dir_str or '支' in dir_str or '借' in dir_str or '转出' in dir_str:
                    direction = '出项'
            else:
                # 进项/出项分列
                in_str = cell_val(col_income) if col_income else ''
                out_str = cell_val(col_expense) if col_expense else ''
                try:
                    in_amt = float(in_str.replace(',', '').replace('¥', '')) if in_str else 0
                except ValueError:
                    in_amt = 0
                try:
                    out_amt = float(out_str.replace(',', '').replace('¥', '')) if out_str else 0
                except ValueError:
                    out_amt = 0
                if in_amt > 0:
                    direction = '进项'
                    amount = in_amt
                elif out_amt > 0:
                    direction = '出项'
                    amount = out_amt

            if amount <= 0:
                # 跳过空行或金额为0的行
                if not cell_val(col_date) and not cell_val(col_purpose):
                    continue
                failed += 1
                errors.append(f"第{row_idx}行：金额无效或为0")
                continue

            tx_date = cell_val(col_date) or datetime.now().strftime('%Y-%m-%d')
            purpose = cell_val(col_purpose)

            try:
                conn.execute("""INSERT INTO bank_transactions(account_id, direction, amount, purpose,
                              source, transaction_date, created_at) VALUES(?,?,?,?,?,?,?)""",
                             (self.selected_account_id, direction, amount, purpose, '手动导入', tx_date,
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                success += 1
            except Exception as e:
                failed += 1
                errors.append(f"第{row_idx}行：导入失败 - {e}")

        conn.commit()
        conn.close()

        self.refresh_transactions()
        self.refresh_accounts()
        self.refresh_stats()

        msg = f"导入完成！\n\n成功：{success} 条\n失败：{failed} 条"
        if errors:
            msg += "\n\n失败详情：\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                msg += f"\n...（共{len(errors)}条错误）"
        messagebox.showinfo("导入结果", msg)
        self.app.set_status(f"银行交易导入完成：成功{success}条，失败{failed}条")


# ============================================================
# 交互式红框标注编辑器
# ============================================================
class FlowFrame(ttk.Frame):
    """自动换行的流式容器：子控件放不下时自动换到下一行，保证所有控件完整可见。"""
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind('<Configure>', lambda e: self._relayout())
        self._last_w = -1
        self._last_h = -1

    def _relayout(self, force=False):
        w = self.winfo_width()
        if w <= 10:
            return
        avail = w - 6
        x, y = 3, 3
        max_h = 0
        for c in self.winfo_children():
            cw = c.winfo_reqwidth() + 10
            ch = c.winfo_reqheight()
            if x + cw > avail and x > 3:
                x = 3
                y += max_h + 4
                max_h = 0
            c.place(x=x, y=y)
            x += cw
            max_h = max(max_h, ch)
        total_h = y + max_h + 6
        if self._last_w != w or self._last_h != total_h or force:
            self._last_w = w
            self._last_h = total_h
            try:
                # place 布局的子控件不贡献请求高度，必须显式设置高度避免塌陷
                self.configure(height=total_h)
            except Exception:
                pass


class AnnotationEditor(tk.Toplevel):
    """人工确认编辑器：四边形红框=发票裁切范围，四边形绿框=字段强化识别（发票号/日期/金额）。"""
    VERTEX_SIZE = 10
    FIELD_TYPES = [
        ('invoice_number', '发票号', '#34C759'),
        ('date', '日期', '#007AFF'),
        ('amount', '金额', '#FF9500'),
    ]

    def __init__(self, parent, annotations, on_confirm):
        super().__init__(parent)
        self.title("人工确认裁切范围 - 红框发票/绿框字段强化")
        self.geometry("1150x820")
        self.minsize(1080, 680)
        _editor_bg = '#1C1C1E' if is_dark_theme() else '#E6E7E8'
        self.configure(bg=_editor_bg)
        self.annotations = annotations
        self.on_confirm = on_confirm
        self.current_idx = 0
        self.selected_box = -1
        self.selected_vertex = -1  # 0-3, -1表示移动整个框
        self.selected_edge = -1    # 0-3 边中点拉伸（0上/1右/2下/3左），-1无
        self.selected_field = None  # (box_idx, field_name) 或 None
        self.selected_field_edge = -1  # 绿框边中点拉伸（0上/1右/2下/3左），-1无
        self.drag_start = None
        self.drag_box_start = None
        self._photo = None
        self._img_offset = (0, 0)
        self._scale = 1.0
        # 缩放功能变量
        self._user_zoom = 1.0
        self._base_scale = 1.0
        self._img_size = (0, 0)
        # 缩放画质控制：缩放过程中用NEAREST（快），停止后用LANCZOS（清晰）
        self._zoom_quality = Image.LANCZOS
        self._zoom_stop_after_id = None
        self._last_redraw_time = 0

        # 将矩形框转换为四边形4顶点，并初始化每个发票的字段框
        for ann in self.annotations:
            new_boxes = []
            for box in ann['boxes']:
                try:
                    if len(box) == 4 and all(isinstance(p, (tuple, list)) and len(p) == 2 for p in box):
                        # 已经是4点四边形格式
                        new_boxes.append([tuple(p) for p in box])
                    else:
                        # 矩形(x1,y1,x2,y2)转4点
                        x1, y1, x2, y2 = box
                        new_boxes.append([(float(x1), float(y1)), (float(x2), float(y1)),
                                          (float(x2), float(y2)), (float(x1), float(y2))])
                except Exception as e:
                    _log_ocr_error(f"框格式转换失败: {box}, 错误: {e}")
                    # 兜底：用全图
                    iw, ih = ann.get('img_size', (100, 100))
                    new_boxes.append([(0, 0), (float(iw), 0), (float(iw), float(ih)), (0, float(ih))])
            ann['boxes'] = new_boxes
            # 每个发票对应一组字段框 {field_name: 4points or None}
            try:
                if 'field_boxes' not in ann or not isinstance(ann.get('field_boxes'), list) or len(ann['field_boxes']) != len(ann['boxes']):
                    ann['field_boxes'] = [{'invoice_number': None, 'date': None, 'amount': None}
                                           for _ in ann['boxes']]
                else:
                    # 确保每个字段框字典有完整的key
                    for fb in ann['field_boxes']:
                        for k in ('invoice_number', 'date', 'amount'):
                            if k not in fb:
                                fb[k] = None
            except Exception as e:
                _log_ocr_error(f"字段框初始化失败: {e}")
                ann['field_boxes'] = [{'invoice_number': None, 'date': None, 'amount': None}
                                       for _ in ann['boxes']]

        # 顶部工具栏 - 流式自动换行，任何窗口宽度下按钮都完整可见
        self.toolbar = FlowFrame(self)
        self.toolbar.pack(fill='x', padx=10, pady=(8, 4))
        self.prev_btn = ttk.Button(self.toolbar, text="◀ 上一页", command=self.prev_page)
        self.page_label = ttk.Label(self.toolbar, text="", font=('微软雅黑', 11, 'bold'))
        self.next_btn = ttk.Button(self.toolbar, text="下一页 ▶", command=self.next_page)
        ttk.Button(self.toolbar, text="添加发票框", command=self.add_box)
        ttk.Button(self.toolbar, text="删除选中", command=self.delete_selected)
        ttk.Button(self.toolbar, text="重置红框", command=self.reset_boxes)
        self.auto_rec_btn = ttk.Button(self.toolbar, text="自动识别字段",
                                       command=self.auto_recognize_fields, style='Accent.TButton')

        ttk.Label(self.toolbar, text="  字段强化:", foreground='#A0A0A0')
        self.field_buttons = {}
        for field_name, label, color in self.FIELD_TYPES:
            btn = ttk.Button(self.toolbar, text=f"🟢{label}", command=lambda f=field_name: self.add_field_box(f))
            self.field_buttons[field_name] = btn
        ttk.Button(self.toolbar, text="删除选中字段框", command=self.delete_invoice_fields)
        ttk.Button(self.toolbar, text="清除本页框", command=self.clear_field_boxes)
        ttk.Label(self.toolbar, text="💡 顶点=斜调 边中点=整条边 框内=移动",
                 foreground='#FFD60A', font=('微软雅黑', 9))
        try:
            self.toolbar._relayout(True)
        except Exception:
            pass

        # Canvas区域 - 使用Grid布局支持滚动条
        _canvas_bg = '#000000' if is_dark_theme() else '#FFFFFF'
        self.canvas_frame = ttk.Frame(self)
        self.canvas_frame.pack(fill='both', expand=True, padx=10, pady=4)
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(self.canvas_frame, bg=_canvas_bg, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky='nsew')
        # 垂直滚动条（右边，红色框位置）
        self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient='vertical', command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky='ns')
        # 水平滚动条（下边，绿色框位置）
        self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient='horizontal', command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky='ew')
        self.canvas.config(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        # 初始隐藏滚动条
        self.v_scrollbar.grid_remove()
        self.h_scrollbar.grid_remove()
        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.canvas.bind('<Configure>', lambda e: self.redraw())
        # Ctrl+鼠标滚轮缩放
        self.canvas.bind('<Control-MouseWheel>', self.on_ctrl_mouse_wheel)
        self.bind('<Control-MouseWheel>', self.on_ctrl_mouse_wheel)

        # 底部按钮
        bottom = ttk.Frame(self)
        bottom.pack(fill='x', padx=10, pady=10)
        self.box_count_label = ttk.Label(bottom, text="")
        self.box_count_label.pack(side='left', padx=10)
        ttk.Button(bottom, text="🔍 重置缩放", command=self.reset_zoom).pack(side='left', padx=6)
        ttk.Button(bottom, text="取消", command=self.destroy).pack(side='right', padx=6)
        ttk.Button(bottom, text="确认全部并导入", command=self.confirm_all).pack(side='right', padx=6)
        ttk.Button(bottom, text="确认本页并下一页", command=self.confirm_current_and_next).pack(side='right', padx=6)

        self.load_page(0)
        self.grab_set()

    def load_page(self, idx):
        self.current_idx = idx
        ann = self.annotations[idx]
        self.page_label.config(text=f"第 {idx+1}/{len(self.annotations)} 页 - {ann['source_file']} 第{ann['page_num']}页")
        self.toolbar._relayout(True)
        self.selected_box = -1
        self.selected_vertex = -1
        self.selected_edge = -1
        self.selected_field = None
        self.selected_field_edge = -1
        # 记录每个字段框的所有Canvas元素ID，便于直接删除
        self._field_element_ids = {}
        # 翻页时重置缩放
        self._user_zoom = 1.0
        self.redraw()

    def _to_screen(self, x, y):
        return (self._img_offset[0] + x * self._scale, self._img_offset[1] + y * self._scale)

    def _to_img(self, sx, sy):
        # 考虑滚动位置：使用canvasx/canvasy将屏幕坐标转换为Canvas坐标
        try:
            cx = self.canvas.canvasx(sx)
            cy = self.canvas.canvasy(sy)
        except Exception:
            cx, cy = sx, sy
        return ((cx - self._img_offset[0]) / self._scale, (cy - self._img_offset[1]) / self._scale)

    def redraw(self):
        try:
            self.canvas.delete('all')
            ann = self.annotations[self.current_idx]
            try:
                orig_img = Image.open(ann['orig_path'])
            except Exception:
                return
            # 使用canvas_frame的大小计算_base_scale，避免滚动条显示/隐藏时大小变化
            frame_w = self.canvas_frame.winfo_width() or 1000
            frame_h = self.canvas_frame.winfo_height() or 600
            if frame_w < 50: frame_w = 1000
            if frame_h < 50: frame_h = 600
            iw, ih = orig_img.size
            self._img_size = (iw, ih)
            # 预留18像素滚动条空间，确保100%时图片不会因为滚动条出现而大小突变
            SCROLLBAR_SPACE = 18
            calc_cw = max(1, frame_w - SCROLLBAR_SPACE)
            calc_ch = max(1, frame_h - SCROLLBAR_SPACE)
            self._base_scale = min(calc_cw / iw, calc_ch / ih, 1.0)
            # 总缩放比例 = 基础缩放 × 用户缩放
            total_scale = self._base_scale * self._user_zoom
            self._scale = total_scale
            dw, dh = int(iw * total_scale), int(ih * total_scale)
            
            # 先根据frame大小判断是否需要滚动条，更新布局后获取Canvas实际大小
            need_h = dw > frame_w
            need_v = dh > frame_h
            if need_h:
                self.h_scrollbar.grid()
            else:
                self.h_scrollbar.grid_remove()
            if need_v:
                self.v_scrollbar.grid()
            else:
                self.v_scrollbar.grid_remove()
            self.canvas.update_idletasks()
            self.canvas_frame.update_idletasks()
            actual_cw = self.canvas.winfo_width() or frame_w
            actual_ch = self.canvas.winfo_height() or frame_h
            if actual_cw < 50: actual_cw = frame_w
            if actual_ch < 50: actual_ch = frame_h
            
            # 关键：分别处理水平和垂直方向
            # 未超出的方向居中，超出的方向使用滚动条（偏移0）
            if dw <= actual_cw:
                offset_x = (actual_cw - dw) // 2
                self.h_scrollbar.grid_remove()
            else:
                offset_x = 0
                self.h_scrollbar.grid()
            if dh <= actual_ch:
                offset_y = (actual_ch - dh) // 2
                self.v_scrollbar.grid_remove()
            else:
                offset_y = 0
                self.v_scrollbar.grid()
            self._img_offset = (offset_x, offset_y)
            
            # 设置scrollregion
            scroll_w = max(dw, actual_cw)
            scroll_h = max(dh, actual_ch)
            self.canvas.config(scrollregion=(0, 0, scroll_w, scroll_h))
            
            # 如果图片完全未超出窗口，重置视图位置
            if dw <= actual_cw and dh <= actual_ch:
                try:
                    self.canvas.xview_moveto(0)
                    self.canvas.yview_moveto(0)
                except Exception:
                    pass
            
            resized = orig_img.resize((dw, dh), self._zoom_quality)
            self._photo = ImageTk.PhotoImage(resized)
            ox, oy = self._img_offset
            self.canvas.create_image(ox, oy, anchor='nw', image=self._photo)

            for i, points in enumerate(ann['boxes']):
                screen_pts = []
                for px, py in points:
                    sx, sy = self._to_screen(px, py)
                    screen_pts.extend([sx, sy])
                is_sel = (i == self.selected_box and self.selected_field is None)
                color = '#FF3B30' if is_sel else '#FF6B6B'
                width = 3 if is_sel else 2
                dash = None if is_sel else (4, 2)
                self.canvas.create_polygon(screen_pts, outline=color, width=width, fill='', dash=dash)
                # 序号
                cx = sum(p[0] for p in points) / 4
                cy = sum(p[1] for p in points) / 4
                scx, scy = self._to_screen(cx, cy)
                self.canvas.create_text(scx, scy - 10, text=str(i + 1), fill=color, font=('Arial', 14, 'bold'))
                # 红框顶点（仅当红框被选中且没有选中绿框时）
                if is_sel:
                    for vi, (px, py) in enumerate(points):
                        sx, sy = self._to_screen(px, py)
                        self.canvas.create_oval(sx - self.VERTEX_SIZE // 2, sy - self.VERTEX_SIZE // 2,
                                                sx + self.VERTEX_SIZE // 2, sy + self.VERTEX_SIZE // 2,
                                                fill='#FFD60A', outline='#FF3B30', width=2)
                    # 红框4条边的中点（可整条边拉伸）
                    edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
                    for ei, (a, b) in enumerate(edge_ends):
                        mx = (points[a][0] + points[b][0]) / 2
                        my = (points[a][1] + points[b][1]) / 2
                        sx, sy = self._to_screen(mx, my)
                        fill_c = '#FFFFFF' if ei == self.selected_edge else '#FF9500'
                        self.canvas.create_rectangle(sx - self.VERTEX_SIZE // 2, sy - self.VERTEX_SIZE // 2,
                                                     sx + self.VERTEX_SIZE // 2, sy + self.VERTEX_SIZE // 2,
                                                     fill=fill_c, outline='#FF3B30', width=2)

            # 绘制绿框（字段强化框），在红框之上
            field_colors = {fn: color for fn, label, color in self.FIELD_TYPES}
            field_labels = {fn: lb for fn, lb, _ in self.FIELD_TYPES}
            # 清空旧的元素ID记录
            self._field_element_ids = {}
            for box_idx in range(len(ann['boxes'])):
                fb = ann['field_boxes'][box_idx]
                for field_name, fpoints in fb.items():
                    if fpoints is None:
                        continue
                    # 给该字段框的所有Canvas元素添加统一标签，便于单独删除
                    _field_tag = f"field_{box_idx}_{field_name}"
                    _elem_ids = []
                    screen_pts = []
                    for px, py in fpoints:
                        sx, sy = self._to_screen(px, py)
                        screen_pts.extend([sx, sy])
                    is_sel = (self.selected_field == (box_idx, field_name))
                    fcolor = field_colors.get(field_name, '#34C759')
                    width = 3 if is_sel else 2
                    dash = None if is_sel else (5, 3)
                    _id = self.canvas.create_polygon(screen_pts, outline=fcolor, width=width, fill='', dash=dash, tags=(_field_tag,))
                    _elem_ids.append(_id)
                    # 字段标签
                    cx = sum(p[0] for p in fpoints) / 4
                    cy = sum(p[1] for p in fpoints) / 4
                    scx, scy = self._to_screen(cx, cy)
                    _id = self.canvas.create_text(scx, scy, text=field_labels.get(field_name, field_name),
                                           fill=fcolor, font=('微软雅黑', 9, 'bold'), tags=(_field_tag,))
                    _elem_ids.append(_id)
                    # 绿框顶点
                    if is_sel:
                        for vi, (px, py) in enumerate(fpoints):
                            sx, sy = self._to_screen(px, py)
                            _id = self.canvas.create_oval(sx - self.VERTEX_SIZE // 2, sy - self.VERTEX_SIZE // 2,
                                                    sx + self.VERTEX_SIZE // 2, sy + self.VERTEX_SIZE // 2,
                                                    fill='#FFD60A', outline=fcolor, width=2, tags=(_field_tag,))
                            _elem_ids.append(_id)
                        # 绿框4条边的中点（可整条边拉伸，与红框一致）
                        edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
                        for ei, (a, b) in enumerate(edge_ends):
                            mx = (fpoints[a][0] + fpoints[b][0]) / 2
                            my = (fpoints[a][1] + fpoints[b][1]) / 2
                            sx, sy = self._to_screen(mx, my)
                            edge_sel = (self.selected_field == (box_idx, field_name) and
                                        self.selected_field_edge == ei)
                            fill_c = '#FFFFFF' if edge_sel else '#FF9500'
                            _id = self.canvas.create_rectangle(sx - self.VERTEX_SIZE // 2, sy - self.VERTEX_SIZE // 2,
                                                         sx + self.VERTEX_SIZE // 2, sy + self.VERTEX_SIZE // 2,
                                                         fill=fill_c, outline=fcolor, width=2, tags=(_field_tag,))
                            _elem_ids.append(_id)

                    # 保存该字段框的所有元素ID
                    self._field_element_ids[(box_idx, field_name)] = _elem_ids

            # 更新底部统计
            field_count = sum(1 for fb in ann.get('field_boxes', []) for v in fb.values() if v is not None)
            self.box_count_label.config(text=f"当前页 {len(ann['boxes'])} 张发票，{field_count} 个字段强化框")
            # 释放图片内存（PhotoImage已复制数据）
            try:
                resized.close()
            except Exception:
                pass
            try:
                orig_img.close()
            except Exception:
                pass

        except Exception as e:
            _log_ocr_error(f"人工确认编辑器redraw失败: {e}")

    def _point_in_polygon(self, px, py, points):
        """判断点是否在多边形内（射线法）"""
        n = len(points)
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = points[i]
            xj, yj = points[j]
            if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    def on_ctrl_mouse_wheel(self, event, delta=None):
        """Ctrl+鼠标滚轮缩放，以鼠标位置为中心"""
        try:
            import time
            wheel_delta = delta if delta is not None else event.delta
            if wheel_delta == 0:
                return
            # 节流：限制最大重绘频率为60fps（16ms）
            now = time.time()
            if now - self._last_redraw_time < 0.016:
                # 频率限制内，保存参数稍后执行
                self._pending_wheel_event = (event, delta)
                if self._zoom_stop_after_id is None:
                    self._zoom_stop_after_id = self.after(16, self._flush_pending_wheel)
                return
            self._last_redraw_time = now
            # 缩放过程中切换到快速画质NEAREST
            self._zoom_quality = Image.NEAREST
            # 取消之前的停止后高画质重绘
            if self._zoom_stop_after_id is not None:
                try:
                    self.after_cancel(self._zoom_stop_after_id)
                except Exception:
                    pass
            # 200ms后没有滚轮事件，切换回高画质LANCZOS重新渲染
            self._zoom_stop_after_id = self.after(200, self._zoom_stop_high_quality)
            old_zoom = self._user_zoom
            if wheel_delta > 0:
                new_zoom = min(old_zoom * 1.1, 5.0)
            else:
                new_zoom = max(old_zoom / 1.1, 0.2)
            if abs(new_zoom - old_zoom) < 0.001:
                return
            old_total_scale = self._base_scale * old_zoom
            new_total_scale = self._base_scale * new_zoom
            # 鼠标对应的图片坐标（考虑滚动位置和_img_offset）
            try:
                scroll_x = self.canvas.canvasx(event.x)
                scroll_y = self.canvas.canvasy(event.y)
            except Exception:
                scroll_x, scroll_y = event.x, event.y
            img_x = (scroll_x - self._img_offset[0]) / old_total_scale
            img_y = (scroll_y - self._img_offset[1]) / old_total_scale
            # 缩放后新的图片尺寸
            frame_w = self.canvas_frame.winfo_width() or 1000
            frame_h = self.canvas_frame.winfo_height() or 600
            iw, ih = self._img_size
            new_dw, new_dh = int(iw * new_total_scale), int(ih * new_total_scale)
            # 计算新的滚动位置，使鼠标指向的位置保持不变
            new_scroll_x = max(0.0, min(1.0, (img_x * new_total_scale - event.x) / max(1, new_dw)))
            new_scroll_y = max(0.0, min(1.0, (img_y * new_total_scale - event.y) / max(1, new_dh)))
            self._user_zoom = new_zoom
            self.redraw()
            self.canvas.update_idletasks()
            # 设置滚动位置（以鼠标为中心）
            try:
                actual_cw = self.canvas.winfo_width() or frame_w
                actual_ch = self.canvas.winfo_height() or frame_h
                if new_dw > actual_cw:
                    self.canvas.xview_moveto(new_scroll_x)
                if new_dh > actual_ch:
                    self.canvas.yview_moveto(new_scroll_y)
            except Exception as e:
                pass
            # 显示缩放比例
            zoom_percent = int(new_zoom * 100)
            if hasattr(self, 'box_count_label'):
                current_text = self.box_count_label.cget('text')
                if '缩放' not in current_text:
                    self.box_count_label.config(text=current_text + f" | 缩放: {zoom_percent}%")
        except Exception as e:
            pass

    def _flush_pending_wheel(self):
        """执行待处理的滚轮事件（节流补充）"""
        try:
            self._zoom_stop_after_id = None
            if hasattr(self, '_pending_wheel_event') and self._pending_wheel_event is not None:
                event, delta = self._pending_wheel_event
                self._pending_wheel_event = None
                self.on_ctrl_mouse_wheel(event, delta)
        except Exception:
            pass

    def _zoom_stop_high_quality(self):
        """缩放停止后切换回高画质LANCZOS重新渲染"""
        try:
            self._zoom_stop_after_id = None
            self._zoom_quality = Image.LANCZOS
            self.redraw()
        except Exception:
            pass

    def reset_zoom(self):
        """重置缩放到原始大小（100%）"""
        try:
            self._user_zoom = 1.0
            self.redraw()
            try:
                self.canvas.xview_moveto(0)
                self.canvas.yview_moveto(0)
            except Exception:
                pass
            if hasattr(self, 'box_count_label'):
                current_text = self.box_count_label.cget('text')
                if '缩放' in current_text:
                    idx = current_text.find(' | 缩放')
                    self.box_count_label.config(text=current_text[:idx])
        except Exception as e:
            pass

    def on_mouse_down(self, event):
        ann = self.annotations[self.current_idx]
        ix, iy = self._to_img(event.x, event.y)

        # 1. 检查是否点击选中绿框的顶点
        if self.selected_field is not None:
            box_idx, field_name = self.selected_field
            fpoints = ann['field_boxes'][box_idx][field_name]
            if fpoints:
                for vi, (px, py) in enumerate(fpoints):
                    sx, sy = self._to_screen(px, py)
                    if abs(event.x - sx) <= self.VERTEX_SIZE and abs(event.y - sy) <= self.VERTEX_SIZE:
                        self.selected_vertex = vi
                        self.selected_edge = -1
                        self.selected_field_edge = -1
                        self.drag_start = (event.x, event.y)
                        self.drag_box_start = [p[:] for p in fpoints]
                        return
                # 1b. 绿框边中点（整条边拉伸，与红框一致）
                edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
                for ei, (a, b) in enumerate(edge_ends):
                    mx = (fpoints[a][0] + fpoints[b][0]) / 2
                    my = (fpoints[a][1] + fpoints[b][1]) / 2
                    sx, sy = self._to_screen(mx, my)
                    if abs(event.x - sx) <= self.VERTEX_SIZE + 2 and abs(event.y - sy) <= self.VERTEX_SIZE + 2:
                        self.selected_field_edge = ei
                        self.selected_vertex = -1
                        self.selected_edge = -1
                        self.drag_start = (event.x, event.y)
                        self.drag_box_start = [p[:] for p in fpoints]
                        self.redraw()
                        return

        # 2. 检查是否点击选中红框的顶点（仅当没有选中绿框时）
        if self.selected_field is None and self.selected_box >= 0:
            points = ann['boxes'][self.selected_box]
            for vi, (px, py) in enumerate(points):
                sx, sy = self._to_screen(px, py)
                if abs(event.x - sx) <= self.VERTEX_SIZE and abs(event.y - sy) <= self.VERTEX_SIZE:
                    self.selected_vertex = vi
                    self.selected_edge = -1
                    self.drag_start = (event.x, event.y)
                    self.drag_box_start = [p[:] for p in points]
                    return

            # 2b. 检查是否点击选中红框的边中点（整条边拉伸）
            edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
            for ei, (a, b) in enumerate(edge_ends):
                mx = (points[a][0] + points[b][0]) / 2
                my = (points[a][1] + points[b][1]) / 2
                sx, sy = self._to_screen(mx, my)
                if abs(event.x - sx) <= self.VERTEX_SIZE + 2 and abs(event.y - sy) <= self.VERTEX_SIZE + 2:
                    self.selected_edge = ei
                    self.selected_vertex = -1
                    self.drag_start = (event.x, event.y)
                    self.drag_box_start = [p[:] for p in points]
                    self.redraw()
                    return

        # 3. 检查是否点击绿框内部（从后往前，后绘制的在上层）
        for box_idx in range(len(ann['boxes']) - 1, -1, -1):
            fb = ann['field_boxes'][box_idx]
            for field_name in ('amount', 'date', 'invoice_number'):
                fpoints = fb.get(field_name)
                if fpoints and self._point_in_polygon(ix, iy, fpoints):
                    self.selected_field = (box_idx, field_name)
                    self.selected_box = box_idx
                    self.selected_vertex = -1
                    self.selected_edge = -1
                    self.selected_field_edge = -1
                    self.drag_start = (event.x, event.y)
                    self.drag_box_start = [p[:] for p in fpoints]
                    self.redraw()
                    return

        # 4. 检查是否点击红框内部（从后往前）
        for i in range(len(ann['boxes']) - 1, -1, -1):
            if self._point_in_polygon(ix, iy, ann['boxes'][i]):
                self.selected_box = i
                self.selected_field = None
                self.selected_vertex = -1
                self.selected_edge = -1
                self.selected_field_edge = -1
                self.drag_start = (event.x, event.y)
                self.drag_box_start = [p[:] for p in ann['boxes'][i]]
                self.redraw()
                return

        # 5. 点击空白：取消选中
        self.selected_box = -1
        self.selected_field = None
        self.selected_vertex = -1
        self.selected_edge = -1
        self.selected_field_edge = -1
        self.redraw()

    def on_mouse_drag(self, event):
        try:
            if self.drag_start is None:
                return
            ann = self.annotations[self.current_idx]
            dx = (event.x - self.drag_start[0]) / self._scale
            dy = (event.y - self.drag_start[1]) / self._scale
            iw, ih = ann['img_size']

            # 判断拖拽的是绿框还是红框
            if self.selected_field is not None:
                box_idx, field_name = self.selected_field
                if self.selected_vertex >= 0:
                    # 移动绿框单个顶点
                    points = [p[:] for p in self.drag_box_start]
                    nx = max(0, min(iw, points[self.selected_vertex][0] + dx))
                    ny = max(0, min(ih, points[self.selected_vertex][1] + dy))
                    points[self.selected_vertex] = (nx, ny)
                    ann['field_boxes'][box_idx][field_name] = points
                elif self.selected_field_edge >= 0:
                    # 绿框边中点拉伸：整条边往外/往里拉（该边两个端点同步移动）
                    points = [p[:] for p in self.drag_box_start]
                    edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
                    i1, i2 = edge_ends[self.selected_field_edge]
                    if self.selected_field_edge in (0, 2):
                        # 上/下边：只改变 y（垂直拉伸）
                        for vi in (i1, i2):
                            ny = max(0, min(ih, points[vi][1] + dy))
                            points[vi] = (points[vi][0], ny)
                    else:
                        # 左/右边：只改变 x（水平拉伸）
                        for vi in (i1, i2):
                            nx = max(0, min(iw, points[vi][0] + dx))
                            points[vi] = (nx, points[vi][1])
                    ann['field_boxes'][box_idx][field_name] = points
                else:
                    # 移动整个绿框
                    points = []
                    for px, py in self.drag_box_start:
                        nx = max(0, min(iw, px + dx))
                        ny = max(0, min(ih, py + dy))
                        points.append((nx, ny))
                    ann['field_boxes'][box_idx][field_name] = points
            elif self.selected_box >= 0:
                if self.selected_vertex >= 0:
                    # 移动红框单个顶点
                    points = [p[:] for p in self.drag_box_start]
                    nx = max(0, min(iw, points[self.selected_vertex][0] + dx))
                    ny = max(0, min(ih, points[self.selected_vertex][1] + dy))
                    points[self.selected_vertex] = (nx, ny)
                    ann['boxes'][self.selected_box] = points
                elif self.selected_edge >= 0:
                    # 边中点拖拽：整条边往外/往里拉伸（该边两个端点同步移动）
                    points = [p[:] for p in self.drag_box_start]
                    edge_ends = [(0, 1), (1, 2), (2, 3), (3, 0)]
                    i1, i2 = edge_ends[self.selected_edge]
                    if self.selected_edge in (0, 2):
                        # 上/下边：只改变 y（垂直拉伸）
                        for vi in (i1, i2):
                            ny = max(0, min(ih, points[vi][1] + dy))
                            points[vi] = (points[vi][0], ny)
                    else:
                        # 左/右边：只改变 x（水平拉伸）
                        for vi in (i1, i2):
                            nx = max(0, min(iw, points[vi][0] + dx))
                            points[vi] = (nx, points[vi][1])
                    ann['boxes'][self.selected_box] = points
                else:
                    # 移动整个红框
                    points = []
                    for px, py in self.drag_box_start:
                        nx = max(0, min(iw, px + dx))
                        ny = max(0, min(ih, py + dy))
                        points.append((nx, ny))
                    ann['boxes'][self.selected_box] = points
            self.redraw()

        except Exception as e:
            _log_ocr_error(f"人工确认编辑器on_mouse_drag失败: {e}")

    def on_mouse_up(self, event):
        self.drag_start = None
        self.drag_box_start = None
        self.selected_edge = -1
        self.selected_field_edge = -1

    def add_box(self):
        ann = self.annotations[self.current_idx]
        iw, ih = ann['img_size']
        w, h = iw // 4, ih // 4
        x1 = (iw - w) // 2
        y1 = (ih - h) // 2
        ann['boxes'].append([(x1, y1), (x1 + w, y1), (x1 + w, y1 + h), (x1, y1 + h)])
        ann['field_boxes'].append({'invoice_number': None, 'date': None, 'amount': None})
        self.selected_box = len(ann['boxes']) - 1
        self.selected_field = None
        self.selected_vertex = -1
        self.redraw()

    def delete_selected(self):
        if self.selected_field is not None:
            # 删除选中的绿框
            box_idx, field_name = self.selected_field
            self.annotations[self.current_idx]['field_boxes'][box_idx][field_name] = None
            self.selected_field = None
            self.redraw()
            return
        if self.selected_box < 0:
            return
        ann = self.annotations[self.current_idx]
        if len(ann['boxes']) <= 1:
            return
        del ann['boxes'][self.selected_box]
        del ann['field_boxes'][self.selected_box]
        self.selected_box = -1
        self.selected_vertex = -1
        self.redraw()

    def add_field_box(self, field_name):
        """为当前选中的红框（发票）添加一个字段强化绿框"""
        try:
            ann = self.annotations[self.current_idx]
            # 如果没有选中红框，自动选中第一个
            if self.selected_box < 0 or self.selected_box >= len(ann['boxes']):
                if len(ann['boxes']) > 0:
                    self.selected_box = 0
                    self.selected_field = None
                    self.redraw()
                else:
                    messagebox.showinfo("提示", "没有可选择的发票框")
                    return
            box_idx = self.selected_box
            red_box = ann['boxes'][box_idx]
            # 确保field_boxes存在
            if 'field_boxes' not in ann or len(ann['field_boxes']) <= box_idx:
                ann['field_boxes'] = [{'invoice_number': None, 'date': None, 'amount': None}
                                       for _ in ann['boxes']]
            # 默认绿框放在红框中心偏上区域，大小更明显
            cx = sum(p[0] for p in red_box) / 4
            cy = sum(p[1] for p in red_box) / 4
            xs = [p[0] for p in red_box]
            ys = [p[1] for p in red_box]
            bw = max(xs) - min(xs)
            bh = max(ys) - min(ys)
            # 绿框尺寸：宽60%，高20%，更明显
            fw, fh = bw * 0.6, bh * 0.20
            x1 = max(0, cx - fw / 2)
            y1 = max(0, cy - fh / 2)
            iw, ih = ann['img_size']
            x2 = min(iw, x1 + fw)
            y2 = min(ih, y1 + fh)
            ann['field_boxes'][box_idx][field_name] = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
            self.selected_field = (box_idx, field_name)
            self.redraw()
        except Exception as e:
            _log_ocr_error(f"添加字段框失败: {e}")
            messagebox.showerror("错误", f"添加字段框失败: {e}")

    def clear_field_boxes(self):
        """清除当前页所有识别框（绿框）"""
        ann = self.annotations[self.current_idx]
        for fb in ann.get('field_boxes', []):
            for k in fb:
                fb[k] = None
        self.selected_field = None
        self.redraw()

    def delete_invoice_fields(self):
        """删除当前选中的那个字段识别框（绿框），只删选中的，不删其他字段框"""
        ann = self.annotations[self.current_idx]
        # 必须先选中一个字段框（绿框）
        if self.selected_field is None:
            messagebox.showinfo("提示", "请先点击选中要删除的字段框（绿框），再点击此按钮")
            return
        box_idx, field_name = self.selected_field
        if 'field_boxes' not in ann or len(ann['field_boxes']) <= box_idx:
            return
        # 只删除选中的那个字段框的数据（和clear_field_boxes一样的逻辑）
        ann['field_boxes'][box_idx][field_name] = None
        self.selected_field = None
        # 和clear_field_boxes完全一样：直接调用redraw()，不做任何额外操作
        self.redraw()

    def auto_recognize_fields(self):
        """对当前页所有发票自动OCR识别发票号/日期/金额并生成绿框（后台线程，避免卡顿）"""
        ann = self.annotations[self.current_idx]
        boxes = ann.get('boxes', [])
        if not boxes:
            messagebox.showinfo("提示", "当前页没有发票框")
            return
        # 防止重复触发
        if getattr(self, '_auto_recognizing', False):
            return
        self._auto_recognizing = True
        if hasattr(self, 'auto_rec_btn'):
            try:
                self.auto_rec_btn.config(state='disabled')
            except Exception:
                pass
        self._set_auto_status("正在自动识别字段...")
        import threading
        orig_path = ann.get('orig_path')
        page_idx = self.current_idx

        def worker():
            try:
                import cv2
                import numpy as np
                from PIL import Image as _PIL
                orig = _PIL.open(orig_path)
                if orig.mode != 'RGB':
                    orig = orig.convert('RGB')
                new_fbs = []
                for box in boxes:
                    fb = {'invoice_number': None, 'date': None, 'amount': None}
                    try:
                        # 取轴对齐包围盒裁切（红框多为轴对齐；斜框近似处理）
                        xs = [p[0] for p in box]
                        ys = [p[1] for p in box]
                        bx1, by1 = int(min(xs)), int(min(ys))
                        bx2, by2 = int(max(xs)), int(max(ys))
                        if bx2 - bx1 < 20 or by2 - by1 < 20:
                            new_fbs.append(fb)
                            continue
                        sub = orig.crop((bx1, by1, bx2, by2))
                        texts, line_boxes = _ocr_image_return_boxes(sub)
                        if texts:
                            inv_no = extract_invoice_number(texts)
                            extracted = {}
                            try:
                                extracted = self.master.app.inv_tab._extract_invoice_fields(texts)
                            except Exception:
                                pass
                            date_text = extracted.get('date')
                            amount_text = extracted.get('amount')
                            for key, val in (('invoice_number', inv_no),
                                             ('date', date_text), ('amount', amount_text)):
                                if not val:
                                    continue
                                needles = [str(val)]
                                if key == 'date' and '-' in str(val):
                                    y, mo, d = str(val).split('-')
                                    needles += [f"{y}年{int(mo)}月{int(d)}日",
                                                f"{y}年{mo}月{d}日",
                                                f"{y}/{mo}/{d}", f"{y}-{mo}-{d}"]
                                if key == 'amount':
                                    needles += [str(val) + '元', str(val)]
                                lb = None
                                if key == 'amount':
                                    lb = _find_amount_box(texts, line_boxes, val)
                                if not lb:
                                    for nd in needles:
                                        lb = _match_text_box(texts, line_boxes, nd)
                                        if lb:
                                            break
                                if lb and len(lb) == 4:
                                    pts = [(bx1 + int(p[0]), by1 + int(p[1])) for p in lb]
                                    fb[key] = pts
                    except Exception:
                        pass
                    new_fbs.append(fb)
                self.after(0, lambda nf=new_fbs: self._apply_auto_fields(nf, page_idx))
            except Exception as e:
                _log_ocr_error(f"自动识别字段线程失败: {e}")
                self.after(0, lambda: (self._set_auto_status("自动识别失败"),
                                       self._restore_auto_btn()))
            finally:
                self._auto_recognizing = False

        threading.Thread(target=worker, daemon=True).start()

    def _apply_auto_fields(self, new_fbs, page_idx):
        """应用自动识别的绿框结果（主线程）"""
        try:
            if page_idx < 0 or page_idx >= len(self.annotations):
                return
            ann = self.annotations[page_idx]
            if 'field_boxes' not in ann or len(ann['field_boxes']) != len(new_fbs):
                ann['field_boxes'] = [{'invoice_number': None, 'date': None, 'amount': None}
                                      for _ in ann['boxes']]
            for i, fb in enumerate(new_fbs):
                if i < len(ann['field_boxes']):
                    ann['field_boxes'][i] = fb
            if page_idx == self.current_idx:
                self.redraw()
            self._set_auto_status("自动识别完成（绿框=识别位置，可调整或删除）")
            self._restore_auto_btn()
        except Exception as e:
            _log_ocr_error(f"应用自动识别结果失败: {e}")
            self._set_auto_status("自动识别失败")
            self._restore_auto_btn()

    def _restore_auto_btn(self):
        try:
            if hasattr(self, 'auto_rec_btn'):
                self.auto_rec_btn.config(state='normal')
        except Exception:
            pass

    def _set_auto_status(self, text):
        """在页面标签处显示自动识别状态（替换式，不累积）"""
        try:
            if hasattr(self, 'page_label'):
                cur = self.page_label.cget('text')
                if '第' in cur:
                    base = cur.split('|')[0].rstrip()
                    self.page_label.config(text=f"{base}  |  {text}")
                if hasattr(self, 'toolbar'):
                    self.toolbar._relayout(True)
        except Exception:
            pass

    def reset_boxes(self):
        ann = self.annotations[self.current_idx]
        try:
            img = Image.open(ann['orig_path'])
            _, boxes = split_invoice_image(img, return_boxes=True)
            new_boxes = []
            for box in (boxes if boxes else [(0, 0, ann['img_size'][0], ann['img_size'][1])]):
                x1, y1, x2, y2 = box
                new_boxes.append([(x1, y1), (x2, y1), (x2, y2), (x1, y2)])
            ann['boxes'] = new_boxes
            ann['field_boxes'] = [{'invoice_number': None, 'date': None, 'amount': None}
                                   for _ in new_boxes]
            self.selected_box = -1
            self.selected_field = None
            self.selected_vertex = -1
            self.selected_edge = -1
            self.selected_field_edge = -1
            self.redraw()
        except Exception as e:
            messagebox.showerror("错误", f"重新分割失败: {e}")

    def prev_page(self):
        if self.current_idx > 0:
            self.load_page(self.current_idx - 1)

    def next_page(self):
        if self.current_idx < len(self.annotations) - 1:
            self.load_page(self.current_idx + 1)

    def confirm_current_and_next(self):
        if self.current_idx < len(self.annotations) - 1:
            self.load_page(self.current_idx + 1)
        else:
            self.confirm_all()

    def confirm_all(self):
        total = sum(len(a['boxes']) for a in self.annotations)
        field_total = sum(1 for a in self.annotations for fb in a.get('field_boxes', [])
                         for v in fb.values() if v is not None)
        msg = f"共 {len(self.annotations)} 页，{total} 张发票"
        if field_total > 0:
            msg += f"，{field_total} 个字段强化框"
        msg += "，确认导入到待完善列表？"
        if not messagebox.askyesno("确认", msg):
            return
        self.destroy()
        self.on_confirm(self.annotations)


# ============================================================
# 批量添加发票
# ============================================================
class PaymentListEditor(tk.Toplevel):
    '''支付记录列表可视化确认编辑器：原图红框标注每条记录范围，
    可拖拽调整红框，日期/金额彩色识别框可调节，框内OCR，长图可滚动，支持学习。'''
    VERTEX_SIZE = 9
    EDGE_SIZE = 7
    FIELD_TYPES = [
        ('date', '日期框', '#007AFF'),
        ('amount', '金额框', '#FF9500'),
    ]
    # 学习数据文件
    LEARN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'paylist_learn.json')

    def __init__(self, parent, image_path, records, on_confirm, recognition_method='date'):
        super().__init__(parent)
        self.result = None  # 最开始就设为None，确保关闭窗口不会误导入
        self._confirmed = False  # 只有点击确认导入才设为True
        self.recognition_method = recognition_method  # 'date'(日期识别) 或 'monthly'(月度识别)
        self._ui_ready = False  # UI创建完成后才设为True，避免_regenerate_boxes访问未创建的控件
        self.title(f"确认支付记录 - {os.path.basename(image_path)}")
        self.geometry("1250x820")
        self.minsize(950, 620)
        self.image_path = image_path
        self.on_confirm = on_confirm
        self.selected_box = -1
        self.selected_field = None  # (box_idx, field_name)
        self.drag_mode = None  # 'move_box', 'vertex_box', 'edge_box', 'move_field', 'vertex_field', 'edge_field'
        self.drag_extra = None
        self.drag_start = None
        self.drag_box_start = None
        self._photo = None
        self._img_w = 0
        self._img_h = 0
        self._scale = 1.0
        self._learn_data = self._load_learn_data()

        # 加载原图
        try:
            self.original_img = Image.open(image_path)
            self._img_w, self._img_h = self.original_img.size
        except Exception:
            messagebox.showerror("错误", "无法加载图片")
            self.destroy()
            return

        # 初始化框数据：根据选择的识别方案生成框
        self.boxes = []
        self._regenerate_boxes(records)

        # === 顶部工具栏（FlowFrame自动换行）===
        toolbar = FlowFrame(self, padding=6)
        toolbar.pack(fill='x', side='top')
        ttk.Label(toolbar, text=f"共 {len(self.boxes)} 条", font=('微软雅黑', 10, 'bold')).pack(side='left', padx=6)
        # 识别方案切换
        ttk.Label(toolbar, text="识别方案:", foreground='#FF9500').pack(side='left', padx=(8,2))
        self.method_var = tk.StringVar(value='日期识别' if self.recognition_method == 'date' else '月度识别')
        method_cb = ttk.Combobox(toolbar, textvariable=self.method_var, values=['日期识别', '月度识别'],
                                  width=8, state='readonly')
        method_cb.pack(side='left', padx=2)
        method_cb.bind('<<ComboboxSelected>>', lambda e: self._on_method_change())
        ttk.Button(toolbar, text="添加框", command=self.add_box).pack(side='left', padx=2)
        ttk.Button(toolbar, text="删除选中", command=self.delete_box).pack(side='left', padx=2)
        ttk.Button(toolbar, text="全选", command=lambda: self._set_all(True)).pack(side='left', padx=2)
        ttk.Button(toolbar, text="全不选", command=lambda: self._set_all(False)).pack(side='left', padx=2)
        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Label(toolbar, text="字段框:", foreground='#007AFF').pack(side='left')
        self.field_buttons = {}
        for fname, label, color in self.FIELD_TYPES:
            btn = ttk.Button(toolbar, text=label, command=lambda f=fname: self.add_field_box(f))
            btn.pack(side='left', padx=2)
            self.field_buttons[fname] = btn
        ttk.Button(toolbar, text="删除字段框", command=self.delete_field_box).pack(side='left', padx=2)
        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Label(toolbar, text="OCR:", foreground='#FF9500').pack(side='left')
        ttk.Button(toolbar, text="识别日期框", command=lambda: self.ocr_field('date')).pack(side='left', padx=2)
        ttk.Button(toolbar, text="识别金额框", command=lambda: self.ocr_field('amount')).pack(side='left', padx=2)
        ttk.Button(toolbar, text="识别当前红框", command=lambda: self.ocr_box_all()).pack(side='left', padx=2)
        ttk.Button(toolbar, text="识别全部红框", command=self.ocr_all_boxes).pack(side='left', padx=2)
        ttk.Button(toolbar, text="学习当前框位置", command=self.learn_positions).pack(side='left', padx=2)
        ttk.Label(toolbar, text="💡顶点=缩放 边=拉伸 框内=移动", foreground='#888', font=('微软雅黑', 9)).pack(side='left', padx=8)
        try:
            toolbar._relayout(True)
        except Exception:
            pass

        # === 主区域 ===
        main = ttk.Frame(self)
        main.pack(fill='both', expand=True, padx=6, pady=4)

        canvas_frame = ttk.Frame(main)
        canvas_frame.pack(side='left', fill='both', expand=True)
        _canvas_bg2 = '#1C1C1E' if is_dark_theme() else '#FFFFFF'
        self.canvas = tk.Canvas(canvas_frame, bg=_canvas_bg2, highlightthickness=0)
        self.canvas.pack(side='left', fill='both', expand=True)
        vsb = ttk.Scrollbar(canvas_frame, orient='vertical', command=self.canvas.yview)
        vsb.pack(side='right', fill='y')
        self.canvas.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(main, orient='horizontal', command=self.canvas.xview)
        hsb.pack(side='bottom', fill='x', after=canvas_frame)
        self.canvas.configure(xscrollcommand=hsb.set)

        self.canvas.bind('<Button-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.canvas.bind('<Configure>', lambda e: self.redraw())
        self.canvas.bind('<MouseWheel>', lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        # 右侧信息面板
        info_panel = ttk.Frame(main, width=300)
        info_panel.pack(side='right', fill='y', padx=(6, 0))
        ttk.Label(info_panel, text="选中记录信息", font=('微软雅黑', 10, 'bold')).pack(pady=(8, 4))
        ttk.Separator(info_panel).pack(fill='x', pady=4)
        self.info_frame = ttk.Frame(info_panel)
        self.info_frame.pack(fill='x', padx=8, pady=4)

        self.include_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.info_frame, text="包含此记录", variable=self.include_var,
                        command=self._on_include_change).pack(anchor='w', pady=2)

        for label, varname in [("日期:", 'date_var'), ("金额:", 'amount_var'), ("类型/商家:", 'merchant_var'), ("备注:", 'remark_var')]:
            ttk.Label(self.info_frame, text=label).pack(anchor='w', pady=(6, 0))
            setattr(self, varname, tk.StringVar())
            ttk.Entry(self.info_frame, textvariable=getattr(self, varname), width=26).pack(fill='x', pady=2)

        self.date_var.trace_add('write', lambda *a: self._on_info_change('date'))
        self.amount_var.trace_add('write', lambda *a: self._on_info_change('amount'))
        self.merchant_var.trace_add('write', lambda *a: self._on_info_change('merchant'))
        self.remark_var.trace_add('write', lambda *a: self._on_info_change('remark'))

        ttk.Separator(info_panel).pack(fill='x', pady=8)
        self.box_info_label = ttk.Label(info_panel, text="", foreground='#888', font=('微软雅黑', 9), justify='left')
        self.box_info_label.pack(pady=4, anchor='w', padx=8)

        # === 底部 ===
        bottom = ttk.Frame(self, padding=8)
        bottom.pack(fill='x', side='bottom')
        # 左下角状态栏
        self.status_label = ttk.Label(bottom, text="", foreground='#888', font=('微软雅黑', 8))
        self.status_label.pack(side='left', padx=4)
        self.count_label = ttk.Label(bottom, text="", foreground='#007AFF')
        self.count_label.pack(side='left', padx=16)
        ttk.Button(bottom, text="取消", command=self._cancel).pack(side='right', padx=6)
        ttk.Button(bottom, text="确认导入", command=self._confirm).pack(side='right', padx=6)

        self._update_info_panel()
        self._update_count()
        self.redraw()
        self._ui_ready = True  # UI创建完成，后续_regenerate_boxes可正常更新
        # 确保窗口在最前并获得焦点
        top = parent.winfo_toplevel() if hasattr(parent, 'winfo_toplevel') else parent
        self.transient(top)
        self.attributes('-topmost', True)
        self.update_idletasks()
        self.result = None  # 存储确认结果，只有用户点击确认后才设置
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except Exception:
            pass
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Ctrl+A 全选红框（标记为包含）
        self.bind('<Control-a>', self._select_all_boxes)
        self.bind('<Control-A>', self._select_all_boxes)

    def _load_learn_data(self):
        try:
            if os.path.exists(self.LEARN_FILE):
                with open(self.LEARN_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_learn_data(self):
        try:
            os.makedirs(os.path.dirname(self.LEARN_FILE), exist_ok=True)
            with open(self.LEARN_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._learn_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _apply_learned_fields(self, x1, y1, x2, y2):
        """根据学习数据生成字段框的初始位置。"""
        field_boxes = {}
        bw = x2 - x1
        bh = y2 - y1
        for fname, _, _ in self.FIELD_TYPES:
            rel = self._learn_data.get(fname)
            if rel and bw > 0 and bh > 0:
                field_boxes[fname] = {
                    'x1': x1 + bw * rel['x1'], 'y1': y1 + bh * rel['y1'],
                    'x2': x1 + bw * rel['x2'], 'y2': y1 + bh * rel['y2'],
                }
            else:
                field_boxes[fname] = None
        return field_boxes

    def learn_positions(self):
        '''学习当前选中框的字段框相对位置（相对于红框）'''
        if self.selected_box < 0:
            messagebox.showinfo("提示", "请先选中一个红框")
            return
        b = self.boxes[self.selected_box]
        bw = b['x2'] - b['x1']
        bh = b['y2'] - b['y1']
        if bw <= 0 or bh <= 0:
            return
        learned = 0
        for fname, _, _ in self.FIELD_TYPES:
            fb = b['field_boxes'].get(fname)
            if fb:
                self._learn_data[fname] = {
                    'x1': (fb['x1'] - b['x1']) / bw,
                    'y1': (fb['y1'] - b['y1']) / bh,
                    'x2': (fb['x2'] - b['x1']) / bw,
                    'y2': (fb['y2'] - b['y1']) / bh,
                }
                learned += 1
        self._save_learn_data()
        self._set_status(f"学习完成: 已学习 {learned} 个字段框的相对位置")

    def _get_default_box(self, idx, total, start_y=0, record_boundaries=None, date_positions=None):
        """生成默认红框位置。
        支持两种方案：date_positions（基于日期位置）或 record_boundaries（基于文本聚类）。
        否则回退到均匀分配。"""
        h = self._img_h
        w = self._img_w
        if date_positions and idx < len(date_positions):
            # 日期识别：基于记录顶部精确框选，每条记录不重叠
            dp = date_positions[idx]
            if isinstance(dp, (list, tuple)) and len(dp) >= 3:
                record_top, date_y, month_y = dp[0], dp[1], dp[2]
            elif isinstance(dp, (list, tuple)) and len(dp) == 2:
                record_top, date_y, month_y = max(0, dp[0] - 60), dp[0], dp[1]
            else:
                record_top, date_y, month_y = max(0, dp - 60), dp, 0
            # 上边界：优先用记录顶部，每月第一条用月份标题
            if month_y > 0 and (idx == 0 or month_y != (date_positions[idx-1][2] if isinstance(date_positions[idx-1], (list,tuple)) and len(date_positions[idx-1]) >= 3 else 0)):
                y1 = max(0, min(record_top, month_y))
            else:
                y1 = max(0, record_top)
            # 下边界：下一条记录的顶部
            if idx + 1 < len(date_positions):
                next_dp = date_positions[idx + 1]
                if isinstance(next_dp, (list, tuple)) and len(next_dp) >= 3:
                    y2 = min(h, next_dp[0] - 8)
                elif isinstance(next_dp, (list, tuple)):
                    y2 = min(h, next_dp[0] - 8)
                else:
                    y2 = min(h, next_dp - 8)
            else:
                y2 = min(h, date_y + 200)
            if y2 - y1 < 80:
                y2 = min(h, y1 + 120)
            return 20, y1, w - 20, y2
        if record_boundaries and idx < len(record_boundaries):
            # 方案二：基于实际记录边界精确框选
            y1, y2 = record_boundaries[idx]
            y1 = max(0, y1)
            y2 = min(h, y2)
            if y2 - y1 < 100:
                y2 = min(h, y1 + 150)
            return 20, y1, w - 20, y2
        # 回退到均匀分配
        available_h = h - start_y
        slot_h = available_h / max(total, 1)
        box_h = max(200, min(450, int(slot_h * 0.95)))
        y1 = int(start_y + idx * slot_h + slot_h * 0.025)
        y2 = min(h, y1 + box_h)
        y1 = max(start_y, y1)
        return 20, y1, w - 20, y2

    def _find_record_boundaries(self):
        """OCR整张图，基于文本行聚类+月份标题检测来推断每条记录的边界。
        优化：阈值=中位行高x1.8（经水电费截图分析，1.5-2.0最佳），
        月份标题作为强制分隔点，过滤顶部状态栏区域。
        返回 [(y1, y2), ...] 列表。"""
        try:
            engine = get_ocr_engine()
            if engine is None:
                return []
            import tempfile, re
            tmp = os.path.join(tempfile.gettempdir(), f"find_records_{int(time.time()*1000)}.jpg")
            if self.original_img.mode != 'RGB':
                self.original_img.convert('RGB').save(tmp, 'JPEG', quality=85)
            else:
                self.original_img.save(tmp, 'JPEG', quality=85)
            result, _ = engine(tmp)
            try:
                os.unlink(tmp)
            except:
                pass
            if not result:
                return []
            # 收集所有文本行
            lines = []
            month_titles = []  # 月份标题位置
            month_pattern = r'\d{4}年\s*\d{1,2}月'
            for line in result:
                if not line or len(line) < 2:
                    continue
                bbox = line[0]
                text = str(line[1])
                if bbox and len(bbox) >= 2:
                    ytop = int(min(p[1] for p in bbox))
                    ybot = int(max(p[1] for p in bbox))
                    lines.append((ytop, ybot, text))
                    if re.search(month_pattern, text):
                        month_titles.append(ytop)
            if not lines:
                return []
            lines.sort(key=lambda x: x[0])
            # 过滤顶部状态栏（y < 250的区域通常是手机状态栏/标题栏）
            content_start = 250
            # 如果有月份标题，从第一个月份标题前50px开始
            if month_titles:
                content_start = max(0, min(month_titles) - 50)
            lines = [l for l in lines if l[0] >= content_start]
            if not lines:
                return []
            # 自适应聚类阈值：中位行高 x 1.8（经分析1.5-2.0最佳，最小35px）
            heights = [l[1] - l[0] for l in lines if l[1] - l[0] > 0]
            heights.sort()
            median_h = heights[len(heights) // 2] if heights else 35
            gap_threshold = max(35, int(median_h * 1.8))
            # 月份标题作为强制分隔点
            def is_near_month_title(y):
                for my in month_titles:
                    if abs(y - my) < 30:
                        return True
                return False
            clusters = []
            current_cluster = [lines[0]]
            for i in range(1, len(lines)):
                prev_bot = current_cluster[-1][1]
                curr_top = lines[i][0]
                gap = curr_top - prev_bot
                # 间距超过阈值，或遇到月份标题，新记录开始
                if gap > gap_threshold or is_near_month_title(curr_top):
                    clusters.append(current_cluster)
                    current_cluster = [lines[i]]
                else:
                    current_cluster.append(lines[i])
            clusters.append(current_cluster)
            # 计算每个聚类的边界
            boundaries = []
            for cluster in clusters:
                y1 = max(0, cluster[0][0] - 12)
                y2 = min(self._img_h, cluster[-1][1] + 12)
                if y2 - y1 > 60:  # 过滤太小的噪声聚类
                    boundaries.append((y1, y2))
            return boundaries
        except Exception as e:
            _log_ocr_error(f"查找记录边界失败: {e}")
            return []

    def _find_all_date_positions(self):
        """日期识别：基于日期位置+记录间距推算每条支付记录的顶部。
        日期通常在记录的60%位置，用相邻日期的间距推算记录高度，
        往上偏移记录高度的45%即为记录顶部。
        返回(记录顶部y, 日期y, 月份标题y)列表。"""
        try:
            engine = get_ocr_engine()
            if engine is None:
                return []
            import tempfile, re
            tmp = os.path.join(tempfile.gettempdir(), f"find_dates_{int(time.time()*1000)}.jpg")
            if self.original_img.mode != 'RGB':
                self.original_img.convert('RGB').save(tmp, 'JPEG', quality=85)
            else:
                self.original_img.save(tmp, 'JPEG', quality=85)
            result, _ = engine(tmp)
            try:
                os.unlink(tmp)
            except:
                pass
            if not result:
                return []
            all_lines = []
            for line in result:
                if not line or len(line) < 2:
                    continue
                bbox = line[0]
                text = str(line[1])
                if bbox and len(bbox) >= 2:
                    ytop = int(min(p[1] for p in bbox))
                    ybot = int(max(p[1] for p in bbox))
                    all_lines.append((ytop, ybot, text))
            all_lines.sort(key=lambda x: x[0])
            if not all_lines:
                return []
            # UI关键词（注意：金额行如"-99.80"不算UI）
            ui_keywords = ['搜索', '查找', '全部账单', '账单详情', '下载账单',
                          '查看更多', '找不到想要', '仅支持查找', '筛选', '排序',
                          '取消', '确认', '删除', '编辑', '返回', '更多',
                          '微信支付', '支付宝', '生活缴费', '充值记录',
                          '交易成功', '支付成功', '转账成功',
                          'K/s', 'M/s', 'KB/s', 'MB/s', '网速',
                          '当前状态', '支付方式', '付款方式', '未知支付方式',
                          '交易发生在', '曾经注销', '查找本人银行卡']
            amount_pattern = r'[￥¥]\s*\d+\.?\d*|\d+\.?\d*\s*元|[-－]\s*\d+\.?\d*'
            date_patterns = [
                r'\d{4}[-/]\d{1,2}[-/]\d{1,2}',
                r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日',
                r'(?<!\d)\d{1,2}月\s*\d{1,2}日',
            ]
            month_pattern = r'(\d{4})年\s*(\d{1,2})月'
            def is_ui_line(text):
                # 先排除金额行（金额不是UI）
                if re.search(amount_pattern, text):
                    return False
                for kw in ui_keywords:
                    if kw in text:
                        return True
                # 纯数字/符号行（状态栏时间、网速等），但排除金额
                cleaned = re.sub(amount_pattern, '', text)
                if re.match(r'^[\d\s:：./%\-－]+$', cleaned) and cleaned.strip():
                    return True
                return False
            # 找月份标题
            month_titles = []
            for ytop, ybot, text in all_lines:
                m = re.search(month_pattern, text)
                if m:
                    month_titles.append((ytop, int(m.group(1)), int(m.group(2))))
            # 找所有日期行（附近必须有金额，且不是UI区域）
            date_lines = []
            for i, (ytop, ybot, text) in enumerate(all_lines):
                for pat in date_patterns:
                    if re.search(pat, text):
                        # 排除UI消息中的日期（如"仅支持查找2025年1月1日以后的账单"）
                        if is_ui_line(text):
                            break
                        has_amount = False
                        is_ui_area = False
                        # 用像素距离检查（120px内），不用行数，避免远处的UI按钮误判
                        for j in range(max(0, i-15), min(len(all_lines), i+16)):
                            if abs(all_lines[j][0] - ytop) > 120:
                                continue
                            if re.search(amount_pattern, all_lines[j][2]):
                                has_amount = True
                            if is_ui_line(all_lines[j][2]):
                                is_ui_area = True
                        if has_amount and not is_ui_area:
                            date_lines.append((ytop, ybot, text, i))
                        break
            date_lines.sort(key=lambda x: x[0])
            filtered_dates = []
            for dl in date_lines:
                if not filtered_dates or dl[0] - filtered_dates[-1][0] > 30:
                    filtered_dates.append(dl)
            if not filtered_dates:
                return []
            # 计算平均记录高度（相邻日期的间距中位数）
            date_ys = [dl[0] for dl in filtered_dates]
            gaps = []
            for k in range(1, len(date_ys)):
                g = date_ys[k] - date_ys[k-1]
                if 50 < g < 500:
                    gaps.append(g)
            gaps.sort()
            if gaps:
                avg_record_h = gaps[len(gaps)//2]
            else:
                avg_record_h = 200
            # 日期在记录中的位置约55-60%，记录顶部 = 日期y - 记录高度*0.45
            top_offset = max(60, int(avg_record_h * 0.45))
            positions = []
            for idx, (date_ytop, date_ybot, date_text, date_idx) in enumerate(filtered_dates):
                month_y = 0
                for my, yr, mo in month_titles:
                    if my <= date_ytop:
                        month_y = my
                    else:
                        break
                record_top = date_ytop - top_offset
                # 往上扫描，遇到UI行（搜索栏）则记录顶部在UI行下方
                for j in range(date_idx - 1, max(0, date_idx - 20), -1):
                    if is_ui_line(all_lines[j][2]):
                        ui_bottom = all_lines[j][1]
                        if ui_bottom > record_top:
                            record_top = ui_bottom + 10
                        break
                # 每月第一条记录包含月份标题
                if month_y > 0:
                    prev_month_y = 0
                    if idx > 0:
                        for my, yr, mo in month_titles:
                            if my <= filtered_dates[idx-1][0]:
                                prev_month_y = my
                            else:
                                break
                    if idx == 0 or prev_month_y != month_y:
                        record_top = min(record_top, max(0, month_y - 10))
                record_top = max(0, record_top)
                positions.append((record_top, date_ytop, max(0, month_y - 15)))
            return positions
        except Exception as e:
            _log_ocr_error(f"查找日期位置失败: {e}")
            return []

    def _regenerate_boxes(self, records=None):
        """根据当前选择的识别方案重新生成所有红框。"""
        if records is None:
            # 从现有框中恢复记录数据
            records = [{'date': b.get('date',''), 'amount': b.get('amount',''),
                        'merchant': b.get('merchant',''), 'remark': b.get('remark','')}
                       for b in self.boxes]
        total = len(records)
        self.boxes = []
        if self.recognition_method == 'monthly':
            # 月度识别：按月份分组，汇总每月金额，每月一个红框
            month_data = self._group_records_by_month(records)
            month_boundaries = self._find_month_boundaries()
            first_y = month_boundaries[0][1] if month_boundaries else self._find_first_date_y()
            for i, (month_key, mrec) in enumerate(month_data):
                x1, y1, x2, y2 = self._get_month_box(i, len(month_data), start_y=first_y,
                                                      month_boundaries=month_boundaries, month_key=month_key)
                field_boxes = self._apply_learned_fields(x1, y1, x2, y2)
                self.boxes.append({
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                    'date': mrec.get('date',''), 'amount': mrec.get('amount',''),
                    'merchant': mrec.get('merchant',''), 'remark': mrec.get('remark',''),
                    'include': True, 'field_boxes': field_boxes,
                })
        else:
            # 日期识别：按日期位置逐条框选，每条记录从一个日期到下一个日期，不重叠
            date_positions = self._find_all_date_positions()
            if date_positions:
                first_dp = date_positions[0]
                if isinstance(first_dp, (list, tuple)) and len(first_dp) >= 3:
                    first_y = min(first_dp[0], first_dp[2]) if first_dp[2] > 0 else first_dp[0]
                elif isinstance(first_dp, (list, tuple)):
                    first_y = first_dp[0]
                else:
                    first_y = first_dp - 60
            else:
                first_y = self._find_first_date_y()
            for i, rec in enumerate(records):
                x1, y1, x2, y2 = self._get_default_box(i, total, start_y=first_y, date_positions=date_positions)
                field_boxes = self._apply_learned_fields(x1, y1, x2, y2)
                self.boxes.append({
                    'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                    'date': rec.get('date',''), 'amount': rec.get('amount',''),
                    'merchant': rec.get('merchant',''), 'remark': rec.get('remark',''),
                    'include': True, 'field_boxes': field_boxes,
                })
        self.selected_box = -1
        self.selected_field = None
        if self._ui_ready:
            self._update_info_panel()
            self._update_count()
            self.redraw()

    def _on_method_change(self):
        """下拉框切换识别方案。"""
        method = 'monthly' if self.method_var.get() == '月度识别' else 'date'
        self._switch_recognition_method(method)

    def _switch_recognition_method(self, method):
        """切换识别方案并重新生成框。"""
        if method == self.recognition_method:
            return
        self.recognition_method = method
        method_name = '月度识别' if method == 'monthly' else '日期识别'
        self._set_status(f"正在使用【{method_name}】方案重新识别分隔框...")
        self.update_idletasks()
        self._regenerate_boxes()
        self._set_status(f"已切换为【{method_name}】方案，识别到 {len(self.boxes)} 条记录")

    def _find_first_date_y(self):
        """OCR整张图，找到第一个月份标题或日期的y坐标，作为红框起始位置。
        优先找月份标题（如2025年10月），从月份标题开始框，确保人工审核能看到月份。"""
        try:
            engine = get_ocr_engine()
            if engine is None:
                return 0
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), f"find_date_{int(time.time()*1000)}.jpg")
            if self.original_img.mode != 'RGB':
                self.original_img.convert('RGB').save(tmp, 'JPEG', quality=85)
            else:
                self.original_img.save(tmp, 'JPEG', quality=85)
            result, _ = engine(tmp)
            try:
                os.unlink(tmp)
            except:
                pass
            if not result:
                return 0
            import re
            # 优先匹配月份标题（如"2025年10月"），从月份标题开始框
            month_pattern = r'\d{4}年\s*\d{1,2}月'
            for line in result:
                if not line or len(line) < 2:
                    continue
                text = str(line[1])
                if re.search(month_pattern, text):
                    bbox = line[0]
                    if bbox and len(bbox) >= 2:
                        y = min(p[1] for p in bbox)
                        return max(0, int(y) - 20)  # 往上留20px余量
            # 没找到月份标题，再找第一个日期
            date_patterns = [
                r'\d{4}[-/]\d{1,2}[-/]\d{1,2}',
                r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日',
                r'\d{1,2}月\s*\d{1,2}日',
            ]
            for line in result:
                if not line or len(line) < 2:
                    continue
                text = str(line[1])
                for pat in date_patterns:
                    if re.search(pat, text):
                        bbox = line[0]
                        if bbox and len(bbox) >= 2:
                            y = min(p[1] for p in bbox)
                            return max(0, int(y) - 30)
                        return 0
            return 0
        except Exception as e:
            _log_ocr_error(f"查找第一个日期位置失败: {e}")
            return 0

    def _find_month_boundaries(self):
        """OCR整张图，找到所有月份标题的y坐标，用于按月份划分红框。
        返回 [(month_text, y_top, y_bottom), ...] 列表。"""
        try:
            engine = get_ocr_engine()
            if engine is None:
                return []
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), f"find_months_{int(time.time()*1000)}.jpg")
            if self.original_img.mode != 'RGB':
                self.original_img.convert('RGB').save(tmp, 'JPEG', quality=85)
            else:
                self.original_img.save(tmp, 'JPEG', quality=85)
            result, _ = engine(tmp)
            try:
                os.unlink(tmp)
            except:
                pass
            if not result:
                return []
            import re
            month_pattern = r'(\d{4}年\s*\d{1,2}月)'
            months = []
            for line in result:
                if not line or len(line) < 2:
                    continue
                text = str(line[1])
                m = re.search(month_pattern, text)
                if m:
                    bbox = line[0]
                    if bbox and len(bbox) >= 2:
                        y = min(p[1] for p in bbox)
                        months.append((m.group(1), max(0, int(y) - 20)))
            # 计算每个月的结束位置（下一个月的开始位置）
            boundaries = []
            for i, (mtext, ytop) in enumerate(months):
                ybottom = months[i+1][1] if i + 1 < len(months) else self._img_h
                boundaries.append((mtext, ytop, ybottom))
            return boundaries
        except Exception as e:
            _log_ocr_error(f"查找月份边界失败: {e}")
            return []

    def _group_records_by_month(self, records):
        from collections import OrderedDict
        months = OrderedDict()
        for rec in records:
            date_str = rec.get('date', '')
            if len(date_str) >= 7:
                month_key = date_str[:7]
            else:
                month_key = '未知'
            if month_key not in months:
                months[month_key] = {'date': month_key + '-01', 'amount': 0.0,
                                      'merchant': '', 'remark': '', 'count': 0}
            try:
                amt = float(rec.get('amount', 0))
            except (ValueError, TypeError):
                amt = 0
            months[month_key]['amount'] += amt
            months[month_key]['count'] += 1
            if not months[month_key]['merchant'] and rec.get('merchant'):
                months[month_key]['merchant'] = rec.get('merchant')
        result = []
        for mk, m in months.items():
            remark = f"共{m['count']}笔，合计"
            result.append((mk, {
                'date': m['date'],
                'amount': m['amount'],
                'merchant': m['merchant'],
                'remark': remark,
            }))
        return result

    def _get_month_box(self, idx, total, start_y=0, month_boundaries=None, month_key=None):
        h = self._img_h
        w = self._img_w
        if month_boundaries and idx < len(month_boundaries):
            mtext, ytop, ybottom = month_boundaries[idx]
            y1 = max(0, int(ytop))
            y2 = min(h, int(ybottom))
            if y2 - y1 < 100:
                y2 = min(h, y1 + 200)
            return 20, y1, w - 20, y2
        available_h = h - start_y
        slot_h = available_h / max(total, 1)
        box_h = max(250, min(600, int(slot_h * 0.95)))
        y1 = int(start_y + idx * slot_h)
        y2 = min(h, y1 + box_h)
        return 20, y1, w - 20, y2

    def _set_all(self, val):
        for b in self.boxes:
            b['include'] = val
        self._update_count()
        self.redraw()

    def _select_all_boxes(self, event=None):
        """Ctrl+A: 全选所有红框（标记为包含）"""
        for b in self.boxes:
            b['include'] = True
        self._update_count()
        self.redraw()
        self._set_status(f"已全选 {len(self.boxes)} 个红框")
        return 'break'  # 阻止默认行为

    def _set_status(self, msg):
        """设置左下角状态栏文字"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=msg)

    def _update_count(self):
        if not hasattr(self, 'count_label'):
            return
        count = sum(1 for b in self.boxes if b['include'])
        self.count_label.config(text=f"已选 {count}/{len(self.boxes)} 条")

    def _on_include_change(self):
        if 0 <= self.selected_box < len(self.boxes):
            self.boxes[self.selected_box]['include'] = self.include_var.get()
            self._update_count()
            self.redraw()

    def _on_info_change(self, field):
        if 0 <= self.selected_box < len(self.boxes):
            self.boxes[self.selected_box][field] = getattr(self, field + '_var').get()

    def _update_info_panel(self):
        # UI变量可能尚未创建（_regenerate_boxes在__init__早期调用）
        if not hasattr(self, 'date_var') or not hasattr(self, 'box_info_label'):
            return
        if 0 <= self.selected_box < len(self.boxes):
            b = self.boxes[self.selected_box]
            self.include_var.set(b['include'])
            self.date_var.set(b['date'])
            self.amount_var.set(b['amount'])
            self.merchant_var.set(b['merchant'])
            self.remark_var.set(b['remark'])
            field_info = []
            for fname, label, color in self.FIELD_TYPES:
                fb = b['field_boxes'].get(fname)
                field_info.append(f"{label}: {'已设置' if fb else '未设置'}")
            self.box_info_label.config(text=f"框 #{self.selected_box+1}\n位置: ({int(b['x1'])},{int(b['y1'])})-({int(b['x2'])},{int(b['y2'])})\n尺寸: {int(b['x2']-b['x1'])}x{int(b['y2']-b['y1'])}\n" + "\n".join(field_info))
        else:
            for v in ['date_var', 'amount_var', 'merchant_var', 'remark_var']:
                getattr(self, v).set('')
            self.box_info_label.config(text="未选中框\n点击红框选中")
        self._update_count()

    def add_box(self):
        # 新框生成在当前可见区域中间，而不是固定在顶部
        try:
            top_frac, bottom_frac = self.canvas.yview()
            disp_h = self._img_h * self._scale
            visible_top = top_frac * disp_h
            visible_bottom = bottom_frac * disp_h
            visible_middle = (visible_top + visible_bottom) / 2
            y = int(visible_middle / self._scale) - 90  # 框高180，往上偏移一半
            y = max(0, min(self._img_h - 200, y))
        except Exception:
            y = 50
        field_boxes = {fname: None for fname, _, _ in self.FIELD_TYPES}
        self.boxes.append({
            'x1': 50, 'y1': y, 'x2': self._img_w - 50, 'y2': y + 180,
            'date': '', 'amount': '', 'merchant': '', 'remark': '', 'include': True,
            'field_boxes': field_boxes,
        })
        self.selected_box = len(self.boxes) - 1
        self.selected_field = None
        self._update_info_panel()
        self.redraw()
        # 滚动到新框位置，确保可见
        try:
            box_center = (y + 90) * self._scale
            scrollregion_h = self._img_h * self._scale
            target_frac = max(0, min(0.9, (box_center - 100) / scrollregion_h))
            self.canvas.yview_moveto(target_frac)
        except Exception:
            pass

    def delete_box(self):
        if 0 <= self.selected_box < len(self.boxes):
            del self.boxes[self.selected_box]
            self.selected_box = -1
            self.selected_field = None
            self._update_info_panel()
            self.redraw()

    def add_field_box(self, field_name):
        if self.selected_box < 0:
            messagebox.showinfo("提示", "请先选中一个红框")
            return
        b = self.boxes[self.selected_box]
        bw = b['x2'] - b['x1']
        bh = b['y2'] - b['y1']
        # 默认放在红框上部1/3区域
        fb = {
            'x1': b['x1'] + bw * 0.05,
            'y1': b['y1'] + bh * (0.1 if field_name == 'date' else 0.4),
            'x2': b['x1'] + bw * 0.6,
            'y2': b['y1'] + bh * (0.25 if field_name == 'date' else 0.55),
        }
        b['field_boxes'][field_name] = fb
        self.selected_field = (self.selected_box, field_name)
        self._update_info_panel()
        self.redraw()

    def delete_field_box(self):
        if self.selected_field:
            bi, fn = self.selected_field
            self.boxes[bi]['field_boxes'][fn] = None
            self.selected_field = None
            self._update_info_panel()
            self.redraw()

    def ocr_field(self, field_name):
        if self.selected_box < 0:
            self._set_status("请先选中一个红框")
            return
        b = self.boxes[self.selected_box]
        fb = b['field_boxes'].get(field_name)
        if not fb:
            field_label = '日期' if field_name == 'date' else '金额'
            self._set_status(f"请先添加{field_label}字段框")
            return
        x1, y1, x2, y2 = int(fb['x1']), int(fb['y1']), int(fb['x2']), int(fb['y2'])
        if x2 - x1 < 10 or y2 - y1 < 10:
            self._set_status("字段框太小，无法识别")
            return
        self._set_status("正在识别...")
        self.update_idletasks()
        text = self._ocr_region(x1, y1, x2, y2)
        if not text:
            self._set_status("未识别到文字")
            return
        if field_name == 'date':
            date = self._extract_date(text)
            if date:
                b['date'] = date
                self._set_status(f"识别成功: 日期={date}")
            else:
                self._set_status("未提取到日期")
        elif field_name == 'amount':
            # 金额框支持多笔金额求和
            total, count = self._extract_all_amounts(text)
            if count > 0:
                b['amount'] = f"{total:.2f}"
                self._set_status(f"识别成功: 共{count}笔金额，合计={total:.2f}")
            else:
                self._set_status("未提取到金额")
        self._update_info_panel()
        self.redraw()

    def ocr_box_all(self):
        if self.selected_box < 0:
            messagebox.showinfo("提示", "请先选中一个红框")
            return
        b = self.boxes[self.selected_box]
        x1, y1, x2, y2 = int(b['x1']), int(b['y1']), int(b['x2']), int(b['y2'])
        full_result = self._ocr_region(x1, y1, x2, y2, return_full=True)
        if not full_result:
            messagebox.showinfo("提示", "未识别到文字")
            return
        text = '\n'.join([item['text'] for item in full_result])
        date = self._extract_date(text)
        total, count = self._extract_all_amounts(text)
        merchant = self._extract_merchant(text)
        if date:
            b['date'] = date
        if count > 0:
            b['amount'] = f'{total:.2f}'
        if merchant:
            b['merchant'] = merchant

        # 在识别到的日期/金额文本位置自动生成字段框
        import re
        date_found = False
        amount_found = False
        for item in full_result:
            if not item.get('box'):
                continue
            itext = item['text']
            box = item['box']
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            bx1, by1, bx2, by2 = min(xs), min(ys), max(xs), max(ys)
            # 限制在红框范围内
            bx1 = max(b['x1'], bx1)
            by1 = max(b['y1'], by1)
            bx2 = min(b['x2'], bx2)
            by2 = min(b['y2'], by2)
            # 检测日期
            if not date_found and (re.search(r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日', itext)
                                    or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', itext)
                                    or re.search(r'\d{1,2}月\s*\d{1,2}日', itext)):
                b['field_boxes']['date'] = {'x1': bx1 - 5, 'y1': by1 - 5, 'x2': bx2 + 5, 'y2': by2 + 5}
                date_found = True
            # 检测金额：扩展金额框覆盖所有金额（用并集）
            if re.search(r'[￥¥]\s*\d+', itext) or re.search(r'(?<!\d)[-－]\s*\d+\.\d+', itext) or re.search(r'\d+\.\d+\s*元', itext):
                existing = b['field_boxes'].get('amount')
                if existing:
                    # 扩展已有金额框，包含新的金额位置
                    existing['x1'] = min(existing['x1'], bx1 - 5)
                    existing['y1'] = min(existing['y1'], by1 - 5)
                    existing['x2'] = max(existing['x2'], bx2 + 5)
                    existing['y2'] = max(existing['y2'], by2 + 5)
                else:
                    b['field_boxes']['amount'] = {'x1': bx1 - 5, 'y1': by1 - 5, 'x2': bx2 + 5, 'y2': by2 + 5}
                amount_found = True

        self._update_info_panel()
        self.redraw()
        field_msg = []
        if date_found: field_msg.append("日期框")
        if amount_found: field_msg.append("金额框")
        status = f"识别完成: 日期={b['date'] or '无'} 金额={b['amount'] or '无'}"
        if field_msg:
            status += f" 自动生成{','.join(field_msg)}"
        self._set_status(status)

    def ocr_all_boxes(self):
        """识别所有红框（后台线程，避免UI卡死）"""
        import threading
        total = len(self.boxes)
        if total == 0:
            self._set_status("没有红框可识别")
            return
        self._set_status(f"正在识别全部 {total} 个红框...")
        self.update_idletasks()

        def worker():
            recognized = 0
            date_fields = 0
            amount_fields = 0
            import re
            for i, b in enumerate(self.boxes):
                try:
                    x1, y1, x2, y2 = int(b['x1']), int(b['y1']), int(b['x2']), int(b['y2'])
                    if x2 - x1 < 20 or y2 - y1 < 20:
                        continue
                    full_result = self._ocr_region(x1, y1, x2, y2, return_full=True)
                    if not full_result:
                        continue
                    text = '\n'.join([item['text'] for item in full_result])
                    date = self._extract_date(text)
                    total, count = self._extract_all_amounts(text)
                    merchant = self._extract_merchant(text)
                    if date: b['date'] = date
                    if count > 0: b['amount'] = f"{total:.2f}"
                    if merchant: b['merchant'] = merchant
                    recognized += 1
                    # 自动生成字段框
                    for item in full_result:
                        if not item.get('box'): continue
                        itext = item['text']
                        box = item['box']
                        xs = [p[0] for p in box]
                        ys = [p[1] for p in box]
                        bx1, by1, bx2, by2 = min(xs), min(ys), max(xs), max(ys)
                        bx1 = max(b['x1'], bx1); by1 = max(b['y1'], by1)
                        bx2 = min(b['x2'], bx2); by2 = min(b['y2'], by2)
                        if not b['field_boxes'].get('date') and (re.search(r'\d{4}年\s*\d{1,2}月\s*\d{1,2}日', itext)
                                or re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', itext)
                                or re.search(r'\d{1,2}月\s*\d{1,2}日', itext)):
                            b['field_boxes']['date'] = {'x1': bx1-5, 'y1': by1-5, 'x2': bx2+5, 'y2': by2+5}
                            date_fields += 1
                        if re.search(r'[￥¥]\s*\d+', itext) or re.search(r'(?<!\d)[-－]\s*\d+\.\d+', itext) or re.search(r'\d+\.\d+\s*元', itext):
                            existing = b['field_boxes'].get('amount')
                            if existing:
                                existing['x1'] = min(existing['x1'], bx1-5)
                                existing['y1'] = min(existing['y1'], by1-5)
                                existing['x2'] = max(existing['x2'], bx2+5)
                                existing['y2'] = max(existing['y2'], by2+5)
                            else:
                                b['field_boxes']['amount'] = {'x1': bx1-5, 'y1': by1-5, 'x2': bx2+5, 'y2': by2+5}
                                amount_fields += 1
                except Exception:
                    continue
            # 更新UI（在主线程）
            self.after(0, lambda: self._update_info_panel())
            self.after(0, lambda: self.redraw())
            self.after(0, lambda: self._set_status(
                f"识别完成: 共{recognized}/{total}条记录, 日期框{date_fields}个, 金额框{amount_fields}个"))

        threading.Thread(target=worker, daemon=True).start()

    def _ocr_region(self, x1, y1, x2, y2, return_full=False):
        try:
            cropped = self.original_img.crop((x1, y1, x2, y2))
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), f"ocr_{int(time.time()*1000)}.jpg")
            if cropped.mode != 'RGB':
                cropped = cropped.convert('RGB')
            cropped.save(tmp, 'JPEG', quality=95)
            engine = get_ocr_engine()
            if engine is None:
                return '' if not return_full else []
            result, _ = engine(tmp)
            try:
                os.unlink(tmp)
            except Exception:
                pass
            if return_full:
                # 返回完整结果，坐标转换为原图绝对坐标
                full = []
                for line in result:
                    if line and len(line) > 1:
                        box = line[0]
                        text = line[1]
                        # 转换坐标：加上裁切偏移
                        try:
                            if isinstance(box[0], (list, tuple)):
                                abs_box = [(p[0] + x1, p[1] + y1) for p in box]
                            else:
                                abs_box = [(box[i] + x1, box[i+1] + y1) for i in range(0, len(box), 2)]
                            full.append({'box': abs_box, 'text': text})
                        except Exception:
                            full.append({'box': None, 'text': text})
                return full
            texts = [line[1] for line in result if line and len(line) > 1]
            return '\n'.join(texts)
        except Exception as e:
            messagebox.showerror("错误", f"OCR失败: {e}")
            return '' if not return_full else []

    def _extract_date(self, text):
        import re
        m = re.search(r'(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
        if m:
            return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', text)
        if m:
            return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        m = re.search(r'(\d{1,2})月\s*(\d{1,2})日', text)
        if m:
            y = datetime.now().year
            return f"{y:04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
        return ''

    def _extract_amount(self, text):
        import re
        # 注意：[-－]模式用(?<!\d)排除日期分隔符（如2025-10-12中的-10）
        for pat in [r'[￥¥]\s*(\d+\.?\d*)', r'(?<!\d)[-－]\s*(\d+\.?\d*)', r'(\d+\.?\d*)\s*元']:
            m = re.search(pat, text)
            if m:
                try:
                    val = float(m.group(1))
                    if val > 0:
                        return f"{val:.2f}"
                except ValueError:
                    continue
        return ''

    def _extract_all_amounts(self, text):
        """提取文本中所有金额并求和，返回 (总额, 笔数)。
        排除日期分隔符（如2025-10-12中的-10）。"""
        import re
        amounts = []
        # 匹配 ￥XX.XX, ¥XX.XX, -XX.XX(非日期), XX.XX元
        for pat in [r'[￥¥]\s*(\d+\.?\d*)', r'(?<!\d)[-－]\s*(\d+\.?\d*)', r'(\d+\.?\d*)\s*元']:
            for m in re.finditer(pat, text):
                try:
                    val = float(m.group(1))
                    if val > 0:
                        amounts.append(val)
                except ValueError:
                    continue
        if not amounts:
            return 0.0, 0
        total = sum(amounts)
        return total, len(amounts)

    def _extract_merchant(self, text):
        import re
        for line in text.split('\n'):
            line = line.strip()
            if 2 <= len(line) <= 15 and re.match(r'^[\u4e00-\u9fa5（）()]+$', line):
                if not any(kw in line for kw in ['年月', '日时分', '支付', '账单', '记录', '缴费', '成功']):
                    return line
        return ''

    # === Canvas绘制和交互 ===
    def redraw(self):
        self.canvas.delete('all')
        cw = self.canvas.winfo_width()
        if cw < 10:
            cw = 800
        self._scale = min(1.0, cw / self._img_w)
        disp_w = int(self._img_w * self._scale)
        disp_h = int(self._img_h * self._scale)

        disp_img = self.original_img.copy()
        disp_img.thumbnail((disp_w, disp_h), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(disp_img)
        self.canvas.create_image(0, 0, anchor='nw', image=self._photo)
        # 释放disp_img内存（PhotoImage已复制数据）
        try:
            disp_img.close()
        except Exception:
            pass
        self.canvas.configure(scrollregion=(0, 0, disp_w, disp_h))

        for i, b in enumerate(self.boxes):
            x1 = b['x1'] * self._scale
            y1 = b['y1'] * self._scale
            x2 = b['x2'] * self._scale
            y2 = b['y2'] * self._scale
            is_sel = (i == self.selected_box)
            color = '#FF3B30' if b['include'] else '#666666'
            width = 3 if is_sel else 2
            dash = None if is_sel else (5, 3)
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=width, dash=dash)
            label = f"#{i+1}"
            if b['date']:
                label += f" {b['date']}"
            if b['amount']:
                label += f" ¥{b['amount']}"
            self.canvas.create_text(x1 + 4, y1 + 4, anchor='nw', text=label, fill=color, font=('微软雅黑', 9, 'bold'))

            # 字段框
            for fname, flabel, fcolor in self.FIELD_TYPES:
                fb = b['field_boxes'].get(fname)
                if fb:
                    fx1 = fb['x1'] * self._scale
                    fy1 = fb['y1'] * self._scale
                    fx2 = fb['x2'] * self._scale
                    fy2 = fb['y2'] * self._scale
                    is_field_sel = (self.selected_field == (i, fname))
                    fw = 2 if is_field_sel else 1
                    self.canvas.create_rectangle(fx1, fy1, fx2, fy2, outline=fcolor, width=fw, dash=(3, 2) if not is_field_sel else None)
                    self.canvas.create_text(fx1 + 2, fy1 + 2, anchor='nw', text=flabel, fill=fcolor, font=('微软雅黑', 8))
                    if is_field_sel:
                        vs = self.VERTEX_SIZE
                        es = self.EDGE_SIZE
                        # 4个顶点（黄色，更明显）
                        for (vx, vy) in [(fx1, fy1), (fx2, fy1), (fx2, fy2), (fx1, fy2)]:
                            self.canvas.create_rectangle(vx - vs, vy - vs, vx + vs, vy + vs, fill='#FFD60A', outline='#000', width=1)
                        # 4个边中点（蓝色）
                        fcx, fcy = (fx1 + fx2) / 2, (fy1 + fy2) / 2
                        for (ex, ey) in [(fcx, fy1), (fx2, fcy), (fcx, fy2), (fx1, fcy)]:
                            self.canvas.create_rectangle(ex - es, ey - es, ex + es, ey + es, fill='#007AFF', outline='#fff', width=1)

            # 选中红框的顶点和边
            if is_sel:
                vs = self.VERTEX_SIZE
                es = self.EDGE_SIZE
                for (vx, vy) in [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]:
                    self.canvas.create_rectangle(vx - vs, vy - vs, vx + vs, vy + vs, fill='#FFD60A', outline='#000', width=1)
                cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
                for (ex, ey) in [(cx, y1), (x2, cy), (cx, y2), (x1, cy)]:
                    self.canvas.create_rectangle(ex - es, ey - es, ex + es, ey + es, fill='#007AFF', outline='#fff', width=1)

    def _get_canvas_coords(self, event):
        '''获取考虑滚动偏移的Canvas坐标'''
        return self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

    def _hit_test(self, x, y):
        vs = self.VERTEX_SIZE
        es = self.EDGE_SIZE
        # 1. 测试选中字段框的顶点和边中点
        if self.selected_field:
            bi, fn = self.selected_field
            b = self.boxes[bi]
            fb = b['field_boxes'].get(fn)
            if fb:
                fx1, fy1, fx2, fy2 = fb['x1']*self._scale, fb['y1']*self._scale, fb['x2']*self._scale, fb['y2']*self._scale
                vertices = [(fx1, fy1, 0), (fx2, fy1, 1), (fx2, fy2, 2), (fx1, fy2, 3)]
                for vx, vy, vi in vertices:
                    if abs(x - vx) <= vs and abs(y - vy) <= vs:
                        return ('vertex_field', bi, (fn, vi))
                # 边中点
                fcx, fcy = (fx1 + fx2) / 2, (fy1 + fy2) / 2
                edges = [(fcx, fy1, 0), (fx2, fcy, 1), (fcx, fy2, 2), (fx1, fcy, 3)]
                for ex, ey, ei in edges:
                    if abs(x - ex) <= es and abs(y - ey) <= es:
                        return ('edge_field', bi, (fn, ei))
                if fx1 <= x <= fx2 and fy1 <= y <= fy2:
                    return ('move_field', bi, fn)
        # 2. 测试选中红框的顶点和边
        if 0 <= self.selected_box < len(self.boxes):
            b = self.boxes[self.selected_box]
            x1, y1, x2, y2 = b['x1']*self._scale, b['y1']*self._scale, b['x2']*self._scale, b['y2']*self._scale
            vertices = [(x1, y1, 0), (x2, y1, 1), (x2, y2, 2), (x1, y2, 3)]
            for vx, vy, vi in vertices:
                if abs(x - vx) <= vs and abs(y - vy) <= vs:
                    return ('vertex_box', self.selected_box, vi)
            cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
            edges = [(cx, y1, 0), (x2, cy, 1), (cx, y2, 2), (x1, cy, 3)]
            for ex, ey, ei in edges:
                if abs(x - ex) <= es and abs(y - ey) <= es:
                    return ('edge_box', self.selected_box, ei)
        # 3. 测试字段框（未选中的）
        for i, b in enumerate(self.boxes):
            for fname, _, _ in self.FIELD_TYPES:
                fb = b['field_boxes'].get(fname)
                if fb:
                    fx1, fy1, fx2, fy2 = fb['x1']*self._scale, fb['y1']*self._scale, fb['x2']*self._scale, fb['y2']*self._scale
                    if fx1 <= x <= fx2 and fy1 <= y <= fy2:
                        return ('select_field', i, fname)
        # 4. 测试红框内部（从后往前）
        for i in range(len(self.boxes) - 1, -1, -1):
            b = self.boxes[i]
            x1, y1, x2, y2 = b['x1']*self._scale, b['y1']*self._scale, b['x2']*self._scale, b['y2']*self._scale
            if x1 <= x <= x2 and y1 <= y <= y2:
                return ('move_box', i, None)
        return (None, -1, None)

    def on_ctrl_mouse_wheel(self, event, delta=None):
        """Ctrl+鼠标滚轮缩放，以鼠标位置为中心"""
        try:
            import time
            wheel_delta = delta if delta is not None else event.delta
            if wheel_delta == 0:
                return
            # 节流：限制最大重绘频率为60fps（16ms）
            now = time.time()
            if now - self._last_redraw_time < 0.016:
                # 频率限制内，保存参数稍后执行
                self._pending_wheel_event = (event, delta)
                if self._zoom_stop_after_id is None:
                    self._zoom_stop_after_id = self.after(16, self._flush_pending_wheel)
                return
            self._last_redraw_time = now
            # 缩放过程中切换到快速画质NEAREST
            self._zoom_quality = Image.NEAREST
            # 取消之前的停止后高画质重绘
            if self._zoom_stop_after_id is not None:
                try:
                    self.after_cancel(self._zoom_stop_after_id)
                except Exception:
                    pass
            # 200ms后没有滚轮事件，切换回高画质LANCZOS重新渲染
            self._zoom_stop_after_id = self.after(200, self._zoom_stop_high_quality)
            old_zoom = self._user_zoom
            if wheel_delta > 0:
                new_zoom = min(old_zoom * 1.1, 5.0)
            else:
                new_zoom = max(old_zoom / 1.1, 0.2)
            if abs(new_zoom - old_zoom) < 0.001:
                return
            old_total_scale = self._base_scale * old_zoom
            new_total_scale = self._base_scale * new_zoom
            # 鼠标对应的图片坐标（考虑滚动位置和_img_offset）
            try:
                scroll_x = self.canvas.canvasx(event.x)
                scroll_y = self.canvas.canvasy(event.y)
            except Exception:
                scroll_x, scroll_y = event.x, event.y
            img_x = (scroll_x - self._img_offset[0]) / old_total_scale
            img_y = (scroll_y - self._img_offset[1]) / old_total_scale
            # 缩放后新的图片尺寸
            frame_w = self.canvas_frame.winfo_width() or 1000
            frame_h = self.canvas_frame.winfo_height() or 600
            iw, ih = self._img_size
            new_dw, new_dh = int(iw * new_total_scale), int(ih * new_total_scale)
            # 计算新的滚动位置，使鼠标指向的位置保持不变
            new_scroll_x = max(0.0, min(1.0, (img_x * new_total_scale - event.x) / max(1, new_dw)))
            new_scroll_y = max(0.0, min(1.0, (img_y * new_total_scale - event.y) / max(1, new_dh)))
            self._user_zoom = new_zoom
            self.redraw()
            self.canvas.update_idletasks()
            # 设置滚动位置（以鼠标为中心）
            try:
                actual_cw = self.canvas.winfo_width() or frame_w
                actual_ch = self.canvas.winfo_height() or frame_h
                if new_dw > actual_cw:
                    self.canvas.xview_moveto(new_scroll_x)
                if new_dh > actual_ch:
                    self.canvas.yview_moveto(new_scroll_y)
            except Exception as e:
                pass
            # 显示缩放比例
            zoom_percent = int(new_zoom * 100)
            if hasattr(self, 'box_count_label'):
                current_text = self.box_count_label.cget('text')
                if '缩放' not in current_text:
                    self.box_count_label.config(text=current_text + f" | 缩放: {zoom_percent}%")
        except Exception as e:
            pass

    def _flush_pending_wheel(self):
        """执行待处理的滚轮事件（节流补充）"""
        try:
            self._zoom_stop_after_id = None
            if hasattr(self, '_pending_wheel_event') and self._pending_wheel_event is not None:
                event, delta = self._pending_wheel_event
                self._pending_wheel_event = None
                self.on_ctrl_mouse_wheel(event, delta)
        except Exception:
            pass

    def _zoom_stop_high_quality(self):
        """缩放停止后切换回高画质LANCZOS重新渲染"""
        try:
            self._zoom_stop_after_id = None
            self._zoom_quality = Image.LANCZOS
            self.redraw()
        except Exception:
            pass

    def reset_zoom(self):
        """重置缩放到原始大小（100%）"""
        try:
            self._user_zoom = 1.0
            self.redraw()
            try:
                self.canvas.xview_moveto(0)
                self.canvas.yview_moveto(0)
            except Exception:
                pass
            if hasattr(self, 'box_count_label'):
                current_text = self.box_count_label.cget('text')
                if '缩放' in current_text:
                    idx = current_text.find(' | 缩放')
                    self.box_count_label.config(text=current_text[:idx])
        except Exception as e:
            pass

    def on_mouse_down(self, event):
        x, y = self._get_canvas_coords(event)
        mode, box_idx, extra = self._hit_test(x, y)
        if mode is None:
            self.selected_box = -1
            self.selected_field = None
            self._update_info_panel()
            self.redraw()
            return
        if mode == 'select_field':
            self.selected_box = box_idx
            self.selected_field = (box_idx, extra)
            self._update_info_panel()
            self.redraw()
            return
        if mode in ('move_field', 'vertex_field', 'edge_field'):
            self.selected_box = box_idx
            if mode == 'move_field':
                self.selected_field = (box_idx, extra)
                self.drag_mode = 'move_field'
                self.drag_extra = extra
            elif mode == 'vertex_field':
                fn, vi = extra
                self.selected_field = (box_idx, fn)
                self.drag_mode = 'vertex_field'
                self.drag_extra = (fn, vi)
            elif mode == 'edge_field':
                fn, ei = extra
                self.selected_field = (box_idx, fn)
                self.drag_mode = 'edge_field'
                self.drag_extra = (fn, ei)
            b = self.boxes[box_idx]
            fb = b['field_boxes'][self.selected_field[1]]
            self.drag_start = (x, y)
            self.drag_box_start = (fb['x1'], fb['y1'], fb['x2'], fb['y2'])
            self._update_info_panel()
            self.redraw()
            return
        # 红框操作
        self.selected_box = box_idx
        self.selected_field = None
        self.drag_mode = mode
        self.drag_extra = extra
        self.drag_start = (x, y)
        b = self.boxes[box_idx]
        self.drag_box_start = (b['x1'], b['y1'], b['x2'], b['y2'])
        self._update_info_panel()
        self.redraw()

    def on_mouse_drag(self, event):
        if self.drag_mode is None:
            return
        x, y = self._get_canvas_coords(event)
        dx = (x - self.drag_start[0]) / self._scale
        dy = (y - self.drag_start[1]) / self._scale

        if self.drag_mode in ('move_field', 'vertex_field'):
            bi, fn = self.selected_field
            fb = self.boxes[bi]['field_boxes'][fn]
            ox1, oy1, ox2, oy2 = self.drag_box_start
            if self.drag_mode == 'move_field':
                w = ox2 - ox1
                h = oy2 - oy1
                b = self.boxes[bi]
                nx1 = max(b['x1'], min(b['x2'] - w, ox1 + dx))
                ny1 = max(b['y1'], min(b['y2'] - h, oy1 + dy))
                fb['x1'], fb['y1'], fb['x2'], fb['y2'] = nx1, ny1, nx1 + w, ny1 + h
            elif self.drag_mode == 'vertex_field':
                _, vi = self.drag_extra
                b = self.boxes[bi]  # 红框边界作为约束
                rx1, ry1, rx2, ry2 = b['x1'], b['y1'], b['x2'], b['y2']
                nx1, ny1, nx2, ny2 = ox1, oy1, ox2, oy2
                if vi == 0:
                    nx1 = max(rx1, min(rx2 - 5, ox1 + dx)); ny1 = max(ry1, min(ry2 - 5, oy1 + dy))
                elif vi == 1:
                    nx2 = min(rx2, max(rx1 + 5, ox2 + dx)); ny1 = max(ry1, min(ry2 - 5, oy1 + dy))
                elif vi == 2:
                    nx2 = min(rx2, max(rx1 + 5, ox2 + dx)); ny2 = min(ry2, max(ry1 + 5, oy2 + dy))
                elif vi == 3:
                    nx1 = max(rx1, min(rx2 - 5, ox1 + dx)); ny2 = min(ry2, max(ry1 + 5, oy2 + dy))
                fb['x1'], fb['y1'], fb['x2'], fb['y2'] = nx1, ny1, nx2, ny2
            elif self.drag_mode == 'edge_field':
                _, ei = self.drag_extra
                b = self.boxes[bi]
                rx1, ry1, rx2, ry2 = b['x1'], b['y1'], b['x2'], b['y2']
                nx1, ny1, nx2, ny2 = ox1, oy1, ox2, oy2
                if ei == 0: ny1 = max(ry1, min(ry2 - 5, oy1 + dy))  # 上边
                elif ei == 1: nx2 = min(rx2, max(rx1 + 5, ox2 + dx))  # 右边
                elif ei == 2: ny2 = min(ry2, max(ry1 + 5, oy2 + dy))  # 下边
                elif ei == 3: nx1 = max(rx1, min(rx2 - 5, ox1 + dx))  # 左边
                fb['x1'], fb['y1'], fb['x2'], fb['y2'] = nx1, ny1, nx2, ny2
            self.redraw()
            return

        b = self.boxes[self.selected_box]
        ox1, oy1, ox2, oy2 = self.drag_box_start
        if self.drag_mode == 'move_box':
            w = ox2 - ox1
            h = oy2 - oy1
            nx1 = max(0, min(self._img_w - w, ox1 + dx))
            ny1 = max(0, min(self._img_h - h, oy1 + dy))
            b['x1'], b['y1'], b['x2'], b['y2'] = nx1, ny1, nx1 + w, ny1 + h
        elif self.drag_mode == 'vertex_box':
            vi = self.drag_extra
            nx1, ny1, nx2, ny2 = ox1, oy1, ox2, oy2
            if vi == 0:
                nx1 = max(0, min(ox2 - 10, ox1 + dx)); ny1 = max(0, min(oy2 - 10, oy1 + dy))
            elif vi == 1:
                nx2 = min(self._img_w, max(ox1 + 10, ox2 + dx)); ny1 = max(0, min(oy2 - 10, oy1 + dy))
            elif vi == 2:
                nx2 = min(self._img_w, max(ox1 + 10, ox2 + dx)); ny2 = min(self._img_h, max(oy1 + 10, oy2 + dy))
            elif vi == 3:
                nx1 = max(0, min(ox2 - 10, ox1 + dx)); ny2 = min(self._img_h, max(oy1 + 10, oy2 + dy))
            b['x1'], b['y1'], b['x2'], b['y2'] = nx1, ny1, nx2, ny2
        elif self.drag_mode == 'edge_box':
            ei = self.drag_extra
            nx1, ny1, nx2, ny2 = ox1, oy1, ox2, oy2
            if ei == 0: ny1 = max(0, min(oy2 - 10, oy1 + dy))
            elif ei == 1: nx2 = min(self._img_w, max(ox1 + 10, ox2 + dx))
            elif ei == 2: ny2 = min(self._img_h, max(oy1 + 10, oy2 + dy))
            elif ei == 3: nx1 = max(0, min(ox2 - 10, ox1 + dx))
            b['x1'], b['y1'], b['x2'], b['y2'] = nx1, ny1, nx2, ny2
        self._update_info_panel()
        self.redraw()

    def on_mouse_up(self, event):
        self.drag_mode = None
        self.drag_extra = None
        self.drag_start = None
        self.drag_box_start = None

    def _on_close(self):
        """窗口关闭按钮：等同于取消，不导入任何数据"""
        self._confirmed = False
        self.result = None
        self.destroy()

    def _cancel(self):
        self._confirmed = False
        self.result = None
        self.destroy()

    def _confirm(self):
        result = []
        for b in self.boxes:
            if b['include']:
                result.append({
                    'date': b['date'], 'amount': b['amount'],
                    'merchant': b['merchant'], 'remark': b['remark'],
                    'x1': b['x1'], 'y1': b['y1'], 'x2': b['x2'], 'y2': b['y2'],
                })
        self._confirmed = True
        self.result = result
        self.destroy()


class BatchInvoiceTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.selected_files = []
        self._processing = False
        self._editor_opening = False

        # 标题
        title = ttk.Label(self.content, text="批量添加发票", font=('微软雅黑', 14, 'bold'))
        title.pack(pady=(16, 8))

        # 说明
        desc = ttk.Label(self.content, text="两种模式：「一页一张」每页/每张图当一张发票；「一页多张」自动分割扫描页中排列的多张发票。两种模式均会弹出四边形红框编辑器，拖拽黄色顶点可斜向调整裁切范围，确认后透视校正并导入「发票报销」页的「待完善」状态，补充报销人后保存即整合。",
                         foreground='gray', wraplength=900, justify='left')
        desc.pack(fill='x', padx=20, pady=(0, 12))

        # 模式选择
        mode_frame = ttk.LabelFrame(self.content, text="识别模式", padding=10)
        mode_frame.pack(fill='x', padx=20, pady=4)
        self.mode_var = tk.StringVar(value='single')
        ttk.Radiobutton(mode_frame, text="一页一张（PDF每页一张 / 单张图片）", variable=self.mode_var, value='single').pack(side='left', padx=15)
        ttk.Radiobutton(mode_frame, text="一页多张（自动分割扫描页中的多张发票）", variable=self.mode_var, value='multi').pack(side='left', padx=15)

        # 操作区
        form = ttk.LabelFrame(self.content, text="批量设置", padding=12)
        form.pack(fill='x', padx=20, pady=8)

        ttk.Label(form, text="默认日期（识别不到时使用）:").grid(row=0, column=0, sticky='e', padx=4, pady=6)
        self.default_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(form, textvariable=self.default_date_var, width=14).grid(row=0, column=1, sticky='w', padx=4, pady=6)

        ttk.Label(form, text="已选文件:").grid(row=0, column=2, sticky='e', padx=(20, 4), pady=6)
        self.file_count_var = tk.StringVar(value="0 个")
        ttk.Label(form, textvariable=self.file_count_var, foreground='#007AFF').grid(row=0, column=3, sticky='w', padx=4, pady=6)

        ttk.Button(form, text="选择文件", command=self.choose_files).grid(row=0, column=4, padx=8, pady=6)
        self.start_btn = ttk.Button(form, text="开始识别", command=self.start_batch)
        self.start_btn.grid(row=0, column=5, padx=8, pady=6)
        self.paylist_btn = ttk.Button(form, text="识别支付记录列表", command=self.recognize_payment_list)
        self.paylist_btn.grid(row=0, column=6, padx=8, pady=6)

        # 文件列表
        list_frame = ttk.LabelFrame(self.content, text="待处理文件", padding=8)
        list_frame.pack(fill='both', expand=True, padx=20, pady=8)
        list_wrap = ttk.Frame(list_frame)
        list_wrap.pack(fill='both', expand=True)
        self.file_list = tk.Listbox(list_wrap, height=10, activestyle='dotbox')
        self.file_list.pack(side='left', fill='both', expand=True)
        list_sb = ttk.Scrollbar(list_wrap, orient='vertical', command=self.file_list.yview)
        list_sb.pack(side='right', fill='y')
        self.file_list.configure(yscrollcommand=list_sb.set)

        # 进度
        prog_frame = ttk.Frame(self.content)
        prog_frame.pack(fill='x', padx=20, pady=8)
        self.progress_var = tk.StringVar(value="就绪")
        ttk.Label(prog_frame, textvariable=self.progress_var).pack(anchor='w', pady=(0, 4))
        self.progress_bar = ttk.Progressbar(prog_frame, mode='determinate')
        self.progress_bar.pack(fill='x')

        # 结果统计
        self.result_var = tk.StringVar(value="")
        ttk.Label(self.content, textvariable=self.result_var, foreground='#34C759',
                  font=('微软雅黑', 10)).pack(pady=(8, 4))

        # 详细分割结果
        detail_frame = ttk.LabelFrame(self.content, text="分割详情（每个文件的页数 / 分割出的发票数）", padding=6)
        detail_frame.pack(fill='both', expand=True, padx=20, pady=(0, 12))
        self.detail_text = tk.Text(detail_frame, height=6, wrap='word',
                                   font=('微软雅黑', 9), relief='flat')
        self.detail_text.pack(side='left', fill='both', expand=True)
        detail_sb = ttk.Scrollbar(detail_frame, orient='vertical', command=self.detail_text.yview)
        detail_sb.pack(side='right', fill='y')
        self.detail_text.configure(yscrollcommand=detail_sb.set, state='disabled')

    def choose_files(self):
        paths = filedialog.askopenfilenames(title="选择发票文件（可多选）",
                                            filetypes=[("发票文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.pdf"),
                                                       ("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"),
                                                       ("PDF文件", "*.pdf"),
                                                       ("所有文件", "*.*")])
        if paths:
            self.selected_files = list(paths)
            self.file_list.delete(0, 'end')
            for p in self.selected_files:
                self.file_list.insert('end', os.path.basename(p))
            self.file_count_var.set(f"{len(self.selected_files)} 个")
            self.app.set_status(f"已选择 {len(self.selected_files)} 个文件")

    def start_batch(self):
        if self._processing or getattr(self, '_editor_opening', False):
            return
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择文件")
            return
        default_date = self.default_date_var.get().strip()
        if default_date:
            try:
                datetime.strptime(default_date, '%Y-%m-%d')
            except ValueError:
                messagebox.showwarning("提示", "默认日期格式错误，请使用 YYYY-MM-DD")
                return

        self._processing = True
        self.start_btn.config(state='disabled', text="识别中...")
        self.progress_bar['value'] = 0
        self.progress_var.set("正在准备...")
        self.result_var.set("")

        import threading
        def worker():
            results = {'success': 0, 'failed': 0, 'total_pages': 0, 'total_split': 0,
                       'details': [], 'annotate_paths': [], 'pending_annotations': []}
            total = len(self.selected_files)
            for idx, fpath in enumerate(self.selected_files):
                self.after(0, lambda i=idx, t=total, n=os.path.basename(fpath):
                           self._update_progress(i, t, f"处理中: {n}"))
                try:
                    info = self._process_file(fpath, default_date)
                    results['success'] += info['saved']
                    results['total_pages'] += info['pages']
                    results['total_split'] += info['split']
                    results['annotate_paths'].extend(info.get('annotate_paths', []))
                    results['pending_annotations'].extend(info.get('pending_annotations', []))
                    results['details'].append({
                        'name': os.path.basename(fpath),
                        'pages': info['pages'],
                        'split': info['split'],
                        'saved': info['saved'],
                        'page_details': info.get('page_details', []),
                        'error': None
                    })
                except Exception as e:
                    results['failed'] += 1
                    _log_ocr_error(f"批量处理失败 {fpath}: {e}")
                    results['details'].append({
                        'name': os.path.basename(fpath),
                        'pages': 0, 'split': 0, 'saved': 0,
                        'error': str(e)
                    })
            self.after(0, lambda: self._on_batch_done(results))

        threading.Thread(target=worker, daemon=True).start()

    def recognize_payment_list(self):
        '''识别支付记录列表截图（微信/支付宝/电信账单等，一张图含多条记录），
        解析后逐条添加到发票报销待完善列表'''
        if self._processing:
            return
        if not self.selected_files:
            messagebox.showwarning("提示", "请先选择支付记录列表截图文件")
            return
        self._processing = True
        self._paylist_queue = []  # 收集所有识别结果，逐个确认
        self._paylist_imported = 0  # 实际导入的记录数
        self.paylist_btn.config(state='disabled', text="识别中...")
        self.progress_bar['value'] = 0
        self.progress_var.set("正在识别支付记录列表...")
        self.result_var.set("")

        import threading
        def worker():
            engine = get_ocr_engine()
            if engine is None:
                self.after(0, lambda: messagebox.showerror("错误", "OCR引擎初始化失败"))
                self.after(0, lambda: self._paylist_done(0, 0, 0))
                return
            total = len(self.selected_files)
            total_recognized = 0
            total_files = 0
            failed = 0
            details = []
            for idx, fpath in enumerate(self.selected_files):
                fname = os.path.basename(fpath)
                self.after(0, lambda i=idx, t=total, n=fname:
                           self._update_progress(i, t, f"识别中: {n}"))
                try:
                    if fpath.lower().endswith('.pdf'):
                        img = pdf_to_image(fpath)
                    else:
                        img = Image.open(fpath)
                    if img is None:
                        failed += 1
                        details.append(f"[跳过] {fname}: 无法加载")
                        continue
                    max_side = 8000
                    w, h = img.size
                    if max(w, h) > max_side:
                        ratio = max_side / max(w, h)
                        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
                    import tempfile
                    # 保存到持久临时目录供编辑器使用
                    editor_tmp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'temp_editor')
                    os.makedirs(editor_tmp_dir, exist_ok=True)
                    tmp_path = os.path.join(editor_tmp_dir,
                                            f"paylist_{os.getpid()}_{int(time.time()*1000)}.jpg")
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    img.save(tmp_path, 'JPEG', quality=92)
                    result, _ = engine(tmp_path)
                    # 不删除tmp_path，供编辑器使用
                    if not result:
                        failed += 1
                        details.append(f"[无文字] {fname}")
                        continue
                    texts = [line[1] for line in result if line and len(line) > 1]
                    records = self.app.inv_tab._parse_payment_list(texts)
                    if not records:
                        details.append(f"[未识别到记录] {fname}")
                        continue
                    # 收集结果，等全部识别完后逐个弹出确认编辑器
                    self._paylist_queue.append((records, fname, tmp_path))
                    total_recognized += len(records)
                    total_files += 1
                    details.append(f"{fname}: 识别到 {len(records)} 条，待人工确认")
                except Exception as e:
                    failed += 1
                    _log_ocr_error(f"支付列表识别失败 {fname}: {e}")
                    details.append(f"[失败] {fname}: {e}")
            # 全部识别完后，在主线程逐个处理确认编辑器
            self.after(0, lambda: self._process_paylist_queue(total_files, total_recognized, failed, details))

        threading.Thread(target=worker, daemon=True).start()

    def _crop_record_image(self, img, ocr_lines, rec, fname):
        """根据记录的日期精确定位，在日期行附近小范围内找金额和商家，
        计算裁切区域并保存为单独的图片附件。返回保存的文件路径或空字符串。"""
        import re
        try:
            date_str = rec.get('date', '')  # YYYY-MM-DD
            amount_str = rec.get('amount', '')
            merchant = rec.get('merchant', '')

            date_parts = date_str.split('-') if date_str else []
            if len(date_parts) != 3:
                return ''
            y, m, d = date_parts
            mi, di = int(m), int(d)

            # 构建精确的日期匹配模式
            date_patterns = [
                rf'{mi}月\s*{di}日',
                rf'{m}月\s*{d}日',
                rf'{y}年\s*{mi}月\s*{di}日',
                rf'{y}-{m}-{d}',
            ]

            # 第一步：精确定位日期行
            date_box = None
            date_y = None
            for line in ocr_lines:
                if not line or len(line) < 2:
                    continue
                text = str(line[1]).strip()
                if not text:
                    continue
                for pat in date_patterns:
                    if re.search(pat, text):
                        date_box = line[0]
                        # 计算日期行的中心y坐标
                        if isinstance(date_box[0], (list, tuple)):
                            date_y = sum(p[1] for p in date_box) / 4
                        else:
                            date_y = sum(date_box[1::2]) / 4
                        break
                if date_box:
                    break

            if date_y is None:
                return ''

            # 第二步：在日期行上下小范围内（±200像素）找金额和商家
            search_range = 250
            amount_val = float(amount_str) if amount_str else 0
            matched_boxes = [date_box]

            for line in ocr_lines:
                if not line or len(line) < 2:
                    continue
                box = line[0]
                text = str(line[1]).strip()
                if not text:
                    continue
                # 计算该行中心y坐标
                if isinstance(box[0], (list, tuple)):
                    line_y = sum(p[1] for p in box) / 4
                else:
                    line_y = sum(box[1::2]) / 4
                # 只考虑日期行附近的行
                if abs(line_y - date_y) > search_range:
                    continue
                # 匹配金额（带￥或-前缀更精确）
                amount_matched = False
                for prefix in ['￥', '¥', '-', '－']:
                    if f"{prefix}{amount_val:.2f}" in text or f"{prefix} {amount_val:.2f}" in text:
                        amount_matched = True
                        break
                if not amount_matched and f"{amount_val:.0f}" in text and len(text) <= 15:
                    # 整数金额且行不长
                    amount_matched = True
                if amount_matched:
                    matched_boxes.append(box)
                    continue
                # 匹配商家（短文本）
                if merchant and merchant in text and len(text) <= 10:
                    matched_boxes.append(box)

            # 计算所有匹配框的外接矩形
            all_x = []
            all_y = []
            for box in matched_boxes:
                try:
                    if isinstance(box[0], (list, tuple)):
                        xs = [p[0] for p in box]
                        ys = [p[1] for p in box]
                    else:
                        xs = box[0::2]
                        ys = box[1::2]
                    all_x.extend(xs)
                    all_y.extend(ys)
                except Exception:
                    continue

            if not all_x or not all_y:
                return ''

            img_w, img_h = img.size
            pad_x = 25
            pad_y = 15
            x1 = max(0, int(min(all_x)) - pad_x)
            y1 = max(0, int(min(all_y)) - pad_y)
            x2 = min(img_w, int(max(all_x)) + pad_x)
            y2 = min(img_h, int(max(all_y)) + pad_y)

            # 确保区域有合理大小
            if x2 - x1 < 80 or y2 - y1 < 40:
                return ''
            # 限制最大高度（避免裁太大）
            if y2 - y1 > 500:
                y2 = y1 + 500

            # 裁切并保存
            cropped = img.crop((x1, y1, x2, y2))
            timestamp = int(time.time() * 1000)
            base = os.path.splitext(os.path.basename(fname))[0][:15]
            save_name = f"paylist_{base}_{date_str}_{timestamp}.jpg"
            save_path = os.path.join(INVOICE_DIR, save_name)
            os.makedirs(INVOICE_DIR, exist_ok=True)
            if cropped.mode != 'RGB':
                cropped = cropped.convert('RGB')
            cropped.save(save_path, 'JPEG', quality=90)
            return save_path
        except Exception as e:
            _log_ocr_error(f"裁切记录图片失败: {e}")
            return ''


    def _show_paylist_confirm(self, records, fname, image_path=None):
        """显示支付记录可视化确认编辑器，用户调整红框和OCR后导入选中的记录。
        只有用户在编辑器中点击'确认导入'后，记录才会被添加到待完善列表。"""
        if image_path is None or not os.path.exists(image_path):
            messagebox.showerror("错误", f"无法获取原图路径: {image_path}\n请重新选择文件识别")
            return

        self.app.set_status(f"支付记录人工确认中: {fname} (共{len(records)}条，确认后才导入)")
        try:
            editor = PaymentListEditor(self, image_path, records, None)
            editor.wait_window()  # 阻塞直到用户关闭窗口
        except Exception as e:
            _log_ocr_error(f"打开支付记录编辑器失败: {e}")
            messagebox.showerror("错误", f"打开确认编辑器失败: {e}\n请尝试重新识别。")
            return

        # 只有用户点击"确认导入"后，_confirmed才为True
        confirmed = getattr(editor, '_confirmed', False)
        result = getattr(editor, 'result', None)
        if not confirmed or result is None:
            self.app.set_status(f"支付记录确认已取消（{fname}）")
            return
        if not result:
            self.app.set_status(f"没有选中任何记录（{fname}）")
            messagebox.showinfo("提示", "没有选中任何记录，未导入")
            return

        # 用户已确认，裁切每个框并保存为附件，然后导入
        conn = get_db()
        saved = 0
        try:
            for rec in result:
                img_filename = ''
                try:
                    x1, y1, x2, y2 = int(rec['x1']), int(rec['y1']), int(rec['x2']), int(rec['y2'])
                    if x2 - x1 > 20 and y2 - y1 > 20:
                        cropped = Image.open(image_path).crop((x1, y1, x2, y2))
                        if cropped.mode != 'RGB':
                            cropped = cropped.convert('RGB')
                        timestamp = int(time.time() * 1000)
                        base = os.path.splitext(os.path.basename(fname))[0][:15]
                        save_name = f"paylist_{base}_{rec.get('date','')}_{timestamp}.jpg"
                        save_path = os.path.join(INVOICE_DIR, save_name)
                        os.makedirs(INVOICE_DIR, exist_ok=True)
                        cropped.save(save_path, 'JPEG', quality=90)
                        img_filename = save_name
                except Exception as e:
                    _log_ocr_error(f"裁切记录图片失败: {e}")
                conn.execute(
                    "INSERT INTO invoices(invoice_number, type, invoice_date, "
                    "reimburser_id, amount, purpose, voucher_number, bank_account_id, "
                    "image_path, status, created_at, seller, remark, batch_pending) "
                    "VALUES(?,?,?,?,?,?,?,?,?,?,0,?,?,1)",
                    ('', '支付记录', rec.get('date', ''), None,
                     float(rec.get('amount', 0) or 0),
                     rec.get('remark', '') or '',
                     None, None, img_filename,
                     datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                     rec.get('merchant', ''), rec.get('remark', '')))
                saved += 1
            conn.commit()
        finally:
            conn.close()
        self.app.set_status(f"支付记录导入完成：{saved} 条（{fname}）")
        if hasattr(self.app, 'inv_tab') and hasattr(self.app.inv_tab, 'refresh'):
            self.app.inv_tab.refresh()
        # 累计导入数
        self._paylist_imported += saved
        # 不单独弹messagebox，由_paylist_done统一提示

    def _process_paylist_queue(self, total_files, total_recognized, failed, details):
        """逐个弹出确认编辑器，全部处理完后调用_paylist_done。
        只有用户点击确认导入后，记录才会被插入到待完善列表。"""
        if not self._paylist_queue:
            self._paylist_done(total_files, total_recognized, failed, details)
            return
        records, fname, tmp_path = self._paylist_queue.pop(0)
        # 显示确认编辑器（内部wait_window阻塞，确认后才插入）
        self._show_paylist_confirm(records, fname, tmp_path)
        # 处理下一个
        self.after(100, lambda: self._process_paylist_queue(total_files, total_recognized, failed, details))

    def _paylist_done(self, files, records, failed, details=None):
        self._processing = False
        self.paylist_btn.config(state='normal', text="识别支付记录列表")
        self.progress_var.set(f"完成：{files} 个文件，识别 {records} 条，失败 {failed} 个")
        self.result_var.set(f"【支付记录列表】共识别 {records} 条记录，请在弹出的确认窗口中勾选后导入")
        self.app.set_status(f"支付记录列表识别完成：{records} 条待人工确认")
        if hasattr(self.app, 'inv_tab') and hasattr(self.app.inv_tab, 'refresh'):
            self.app.inv_tab.refresh()
        if details:
            self.detail_text.configure(state='normal')
            self.detail_text.delete('1.0', 'end')
            self.detail_text.insert('end', "支付记录列表识别详情\n")
            self.detail_text.insert('end', "=" * 60 + "\n")
            for d in details:
                self.detail_text.insert('end', d + "\n")
            self.detail_text.configure(state='disabled')
        if records > 0:
            if self._paylist_imported > 0:
                messagebox.showinfo("导入完成", f"共识别 {records} 条支付记录\n已导入 {self._paylist_imported} 条到待完善列表\n请前往发票报销页补充报销人信息")
            else:
                messagebox.showinfo("已取消", f"共识别 {records} 条支付记录\n已全部取消，未导入任何记录")

    def _update_progress(self, idx, total, msg):
        self.progress_bar['maximum'] = total
        self.progress_bar['value'] = idx + 1
        self.progress_var.set(msg)

    def _process_file(self, fpath, default_date):
        """处理单个文件。两种模式都保存原图和框坐标到pending_annotations，等待人工确认。"""
        ext = os.path.splitext(fpath)[1].lower()
        saved = 0
        pages = 0
        split_count = 0
        page_details = []
        annotate_paths = []
        pending_annotations = []
        use_split = (self.mode_var.get() == 'multi')

        if ext == '.pdf':
            try:
                import fitz
                doc = fitz.open(fpath)
                page_count = len(doc)
                pages = page_count
                for page_num in range(page_count):
                    try:
                        page = doc[page_num]
                        mat = fitz.Matrix(200 / 72, 200 / 72)
                        pix = page.get_pixmap(matrix=mat)
                        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        if use_split:
                            sub_images, boxes = split_invoice_image(img, return_boxes=True)
                            page_split = len(sub_images)
                            split_count += page_split
                            page_details.append(f"第{page_num+1}页→{page_split}张（待确认）")
                            # 释放sub_images内存（不需要使用，只需要boxes）
                            for _si in sub_images:
                                try:
                                    _si.close()
                                except Exception:
                                    pass
                            sub_images = None
                        else:
                            page_split = 1
                            split_count += page_split
                            # 单张模式：直接用_detect_invoice_contour检测整张发票区域，不调用split_invoice_image（会内部分割）
                            import cv2 as _cv2
                            import numpy as _np
                            _cv_img = _cv2.cvtColor(_np.array(img), _cv2.COLOR_RGB2BGR)
                            _gray = _cv2.cvtColor(_cv_img, _cv2.COLOR_BGR2GRAY)
                            _, _binary = _cv2.threshold(_gray, 0, 255, _cv2.THRESH_BINARY_INV + _cv2.THRESH_OTSU)
                            _contour = _detect_invoice_contour(_cv_img, _binary, img.size[0], img.size[1])
                            if _contour:
                                boxes = [_contour]
                            else:
                                boxes = [(0, 0, img.size[0], img.size[1])]
                            page_details.append(f"第{page_num+1}页→1张（待确认）")
                            # 释放OpenCV图像内存
                            del _cv_img, _gray, _binary, _contour
                        # 保存原图到临时位置
                        orig_name = f"pending_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_p{page_num+1}.jpg"
                        orig_path = os.path.join(INVOICE_DIR, orig_name)
                        img.save(orig_path, 'JPEG', quality=92)
                        fboxes = None
                        pending_annotations.append({
                            'orig_path': orig_path,
                            'orig_name': orig_name,
                            'boxes': boxes,
                            'field_boxes': fboxes,
                            'source_file': os.path.basename(fpath),
                            'page_num': page_num + 1,
                            'default_date': default_date,
                            'img_size': img.size,
                        })
                        # 释放图片和pixmap内存
                        img.close()
                        pix = None
                        img = None
                    except Exception as pe:
                        _log_ocr_error(f"PDF第{page_num+1}页处理失败 {fpath}: {pe}")
                        page_details.append(f"第{page_num+1}页→处理失败")
                doc.close()
            except Exception as e:
                _log_ocr_error(f"PDF打开失败 {fpath}: {e}")
                raise
        else:
            pages = 1
            try:
                pil_img = Image.open(fpath)
                if pil_img.mode != 'RGB':
                    pil_img = pil_img.convert('RGB')
                if use_split:
                    sub_images, boxes = split_invoice_image(pil_img, return_boxes=True)
                    split_count = len(sub_images)
                    page_details.append(f"第1页→{split_count}张（待确认）")
                    orig_name = f"pending_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
                    orig_path = os.path.join(INVOICE_DIR, orig_name)
                    pil_img.save(orig_path, 'JPEG', quality=92)
                    fboxes = None
                    pending_annotations.append({
                        'orig_path': orig_path,
                        'orig_name': orig_name,
                        'boxes': boxes,
                        'field_boxes': fboxes,
                        'source_file': os.path.basename(fpath),
                        'page_num': 1,
                        'default_date': default_date,
                        'img_size': pil_img.size,
                    })
                    # 释放sub_images内存
                    for _si in sub_images:
                        try:
                            _si.close()
                        except Exception:
                            pass
                    sub_images = None
                else:
                    # 一页一张：直接用_detect_invoice_contour检测整张发票区域，不调用split_invoice_image（会内部分割）
                    split_count = 1
                    import cv2 as _cv2
                    import numpy as _np
                    _cv_img = _cv2.cvtColor(_np.array(pil_img), _cv2.COLOR_RGB2BGR)
                    _gray = _cv2.cvtColor(_cv_img, _cv2.COLOR_BGR2GRAY)
                    _, _binary = _cv2.threshold(_gray, 0, 255, _cv2.THRESH_BINARY_INV + _cv2.THRESH_OTSU)
                    _contour = _detect_invoice_contour(_cv_img, _binary, pil_img.size[0], pil_img.size[1])
                    if _contour:
                        boxes = [_contour]
                    else:
                        iw, ih = pil_img.size
                        boxes = [(0, 0, iw, ih)]
                    page_details.append(f"第1页→1张（待确认）")
                    orig_name = f"pending_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
                    orig_path = os.path.join(INVOICE_DIR, orig_name)
                    pil_img.save(orig_path, 'JPEG', quality=92)
                    fboxes = None
                    pending_annotations.append({
                        'orig_path': orig_path,
                        'orig_name': orig_name,
                        'boxes': boxes,
                        'field_boxes': fboxes,
                        'source_file': os.path.basename(fpath),
                        'page_num': 1,
                        'default_date': default_date,
                        'img_size': pil_img.size,
                    })
                    # 释放OpenCV图像内存
                    del _cv_img, _gray, _binary, _contour
                # 释放图片内存
                pil_img.close()
                pil_img = None
            except Exception as e:
                _log_ocr_error(f"图片处理失败 {fpath}: {e}")
                raise

        return {'pages': pages, 'split': split_count, 'saved': saved,
                'page_details': page_details, 'annotate_paths': annotate_paths,
                'pending_annotations': pending_annotations}

    def _ocr_and_save(self, img_path, img_filename, default_date, overrides=None):
        """OCR识别图片并保存到数据库（batch_pending=1）。overrides为字段强化框的OCR结果，优先使用。"""
        inv_no, texts = ocr_invoice_number(img_path)

        # 复用发票报销页的字段提取逻辑
        extracted = {}
        auto_type = None
        try:
            auto_type = self.app.inv_tab._detect_document_type(texts)
            extracted = self.app.inv_tab._extract_invoice_fields(texts)
        except Exception:
            pass

        inv_type = auto_type or extracted.get('type') or '发票'
        inv_date = extracted.get('date') or default_date
        amount = extracted.get('amount') or 0
        seller = extracted.get('seller') or ''
        remark = extracted.get('remark') or ''
        if not inv_no and extracted.get('invoice_number'):
            inv_no = extracted['invoice_number']

        # 应用字段强化框的OCR覆盖（绿框识别结果优先）
        if overrides:
            if overrides.get('invoice_number'):
                inv_no = overrides['invoice_number']
            if overrides.get('date'):
                inv_date = overrides['date']
            if overrides.get('amount') is not None and overrides['amount'] != 0:
                amount = overrides['amount']
            if overrides.get('seller'):
                seller = overrides['seller']
            if overrides.get('remark'):
                remark = overrides['remark']

        # 保存到数据库，reimburser_id=NULL（待用户补充），batch_pending=1
        conn = get_db()
        try:
            conn.execute("""INSERT INTO invoices(invoice_number, type, invoice_date, reimburser_id, amount,
                          purpose, voucher_number, bank_account_id, image_path, status, created_at, seller, remark, batch_pending)
                          VALUES(?,?,?,?,?,?,?,?,?,?,0,?,?,1)""",
                         (inv_no or '', inv_type, inv_date, None, amount,
                          '', None, None, img_filename,
                          datetime.now().strftime('%Y-%m-%d %H:%M:%S'), seller, remark))
            conn.commit()
        finally:
            conn.close()

    def _on_batch_done(self, results):
        self._processing = False
        self.start_btn.config(state='normal', text="开始识别")
        pending = results.get('pending_annotations', [])
        total_pending = sum(len(a['boxes']) for a in pending)
        self.progress_var.set(f"完成：{results['total_pages']} 页，{total_pending} 张待确认，失败 {results['failed']} 个文件")
        mode_text = "一页多张模式" if self.mode_var.get() == 'multi' else "一页一张模式"
        self.result_var.set(f"【{mode_text}】共 {results['total_pages']} 页，{total_pending} 张发票待人工确认裁切范围")
        self.app.set_status(f"批量识别完成：{total_pending} 张发票待人工确认")

        # 显示详细分割结果
        self.detail_text.configure(state='normal')
        self.detail_text.delete('1.0', 'end')
        self.detail_text.insert('end', f"识别模式：{mode_text}\n")
        self.detail_text.insert('end', f"总计：{len(results['details'])} 个文件，{results['total_pages']} 页，{total_pending} 张发票待确认\n")
        self.detail_text.insert('end', "=" * 60 + "\n")
        for d in results['details']:
            if d['error']:
                self.detail_text.insert('end', f"[失败] {d['name']}\n  错误：{d['error']}\n")
            else:
                self.detail_text.insert('end',
                    f"{d['name']}\n  {d['pages']} 页 → {d['split']} 张发票（待确认）\n")
                if d.get('page_details'):
                    self.detail_text.insert('end', f"  详情: {', '.join(d['page_details'])}\n")
        self.detail_text.insert('end', "=" * 60 + "\n")
        if pending:
            self.detail_text.insert('end', f"⚠ 共 {len(pending)} 页 {total_pending} 张发票待人工确认裁切范围\n")
            self.detail_text.insert('end', "请在弹出的编辑器中拖拽黄色顶点调整四边形红框，确认后透视校正裁切并导入待完善列表。\n")
        self.detail_text.configure(state='disabled')

        # 保存标注图路径，提供查看按钮
        self._last_annotate_paths = results.get('annotate_paths', [])
        if getattr(self, 'view_annotate_btn', None) is not None:
            try:
                self.view_annotate_btn.destroy()
            except Exception:
                pass
        if self._last_annotate_paths:
            self.view_annotate_btn = ttk.Button(self.content, text=f"查看分割标注图（{len(self._last_annotate_paths)}张）",
                                                command=self._view_annotate_images)
            self.view_annotate_btn.pack(pady=4)
        else:
            self.view_annotate_btn = None

        # 打开人工确认编辑器（一页一张和一页多张都走此流程）
        if pending:
            def _open_editor():
                try:
                    editor = AnnotationEditor(self, pending, self._confirm_import)
                    editor.transient(self.winfo_toplevel())
                    editor.grab_set()
                    editor.lift()
                except Exception as e:
                    _log_ocr_error(f"打开人工确认编辑器失败: {e}")
                    messagebox.showerror("错误", f"打开确认编辑器失败: {e}\n请尝试重新识别。")
                finally:
                    self._editor_opening = False
            self._editor_opening = True
            self.after(150, _open_editor)
        else:
            error_detail = "\n".join([f"  {d['name']}: {d.get('error','未知错误')}"
                                       for d in results['details'] if d.get('error')])
            messagebox.showwarning("提示",
                f"没有可确认的发票。\n成功处理 {results['success']} 个，失败 {results['failed']} 个。\n"
                f"失败详情：\n{error_detail}\n\n请检查文件是否为有效的图片或PDF。")

    def _confirm_import(self, annotations):
        """根据人工确认后的红框重新裁切、OCR并导入待完善列表。后台线程运行，避免UI卡死。"""
        total = sum(len(a['boxes']) for a in annotations)
        if total == 0:
            messagebox.showinfo("提示", "没有可导入的发票")
            return

        # 创建进度窗口
        progress_win = tk.Toplevel(self)
        progress_win.title("正在导入")
        progress_win.geometry("360x120")
        _prog_bg = '#1C1C1E' if is_dark_theme() else '#E6E7E8'
        _prog_fg = '#F5F5F7' if is_dark_theme() else '#1D1D1F'
        progress_win.configure(bg=_prog_bg)
        progress_win.transient(self)
        progress_win.grab_set()
        tk.Label(progress_win, text="正在裁切并识别发票，请稍候...", bg=_prog_bg, fg=_prog_fg,
                font=('微软雅黑', 10), pady=10).pack()
        progress_bar = ttk.Progressbar(progress_win, maximum=total, length=300)
        progress_bar.pack(pady=8)
        count_label = tk.Label(progress_win, text=f"0 / {total}", bg=_prog_bg, fg=_prog_fg,
                              font=('微软雅黑', 10, 'bold'))
        count_label.pack(pady=(0, 8))
        progress_win.update()

        def worker():
            import cv2
            import numpy as np
            imported = 0
            for ann in annotations:
                try:
                    orig_img = Image.open(ann['orig_path'])
                    default_date = ann.get('default_date', '')
                    for bi, points in enumerate(ann['boxes']):
                        try:
                            # 1. 透视裁切发票（红框）
                            crop = perspective_crop(orig_img, points)
                            if crop is None:
                                continue
                            fname = f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_p{ann['page_num']}_{bi+1}.jpg"
                            dst = os.path.join(INVOICE_DIR, fname)
                            crop.save(dst, 'JPEG', quality=92)

                            # 2. 字段强化OCR（绿框）：裁切每个字段区域，单独OCR后覆盖整体识别结果
                            overrides = {}
                            field_boxes_list = ann.get('field_boxes', [])
                            if bi < len(field_boxes_list):
                                fb = field_boxes_list[bi]
                                for field_name, fpoints in fb.items():
                                    if fpoints is None:
                                        continue
                                    field_crop = perspective_crop(orig_img, fpoints)
                                    if field_crop is None:
                                        continue
                                    # 临时保存字段图片并OCR
                                    tmp_name = f"_field_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.jpg"
                                    tmp_path = os.path.join(INVOICE_DIR, tmp_name)
                                    try:
                                        field_crop.save(tmp_path, 'JPEG', quality=95)
                                        f_inv_no, f_texts = ocr_invoice_number(tmp_path)
                                        f_extracted = {}
                                        try:
                                            f_extracted = self.app.inv_tab._extract_invoice_fields(f_texts)
                                        except Exception:
                                            pass
                                        if field_name == 'invoice_number':
                                            if f_inv_no:
                                                overrides['invoice_number'] = f_inv_no
                                            elif f_extracted.get('invoice_number'):
                                                overrides['invoice_number'] = f_extracted['invoice_number']
                                        elif field_name == 'date':
                                            if f_extracted.get('date'):
                                                overrides['date'] = f_extracted['date']
                                        elif field_name == 'amount':
                                            if f_extracted.get('amount'):
                                                overrides['amount'] = f_extracted['amount']
                                    except Exception as e:
                                        _log_ocr_error(f"字段{field_name}强化OCR失败: {e}")
                                    finally:
                                        try:
                                            os.remove(tmp_path)
                                        except Exception:
                                            pass

                            # 3. 整体OCR + 字段覆盖 → 保存
                            self._ocr_and_save(dst, fname, default_date, overrides=overrides)
                            imported += 1
                        except Exception as e:
                            _log_ocr_error(f"裁切第{bi+1}张失败: {e}")
                        # 更新进度
                        self.after(0, lambda i=imported: (
                            progress_bar.configure(value=i),
                            count_label.config(text=f"{i} / {total}")
                        ))
                    try:
                        os.remove(ann['orig_path'])
                    except Exception:
                        pass
                except Exception as e:
                    _log_ocr_error(f"确认导入失败 {ann.get('source_file')}: {e}")

            def done():
                try:
                    progress_win.destroy()
                except Exception:
                    pass
                self.app.set_status(f"已导入 {imported} 张确认后的发票到待完善列表")
                messagebox.showinfo("导入完成", f"已成功导入 {imported} 张发票到待完善列表。\n请前往「发票报销」页面补充报销人等信息。")
                try:
                    self.app.inv_tab.filter_status.set('待完善')
                    self.app.inv_tab.filter_year.set('全部')
                    self.app.inv_tab.filter_month.set('全部')
                    self.app.inv_tab.filter_person.set('全部')
                    self.app.inv_tab.refresh()
                    self.app.notebook.select(self.app.inv_tab)
                except Exception:
                    pass
            self.after(0, done)

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def _view_annotate_images(self):
        """查看本次批量识别的分割标注图"""
        if not getattr(self, '_last_annotate_paths', []):
            messagebox.showinfo("提示", "没有分割标注图")
            return
        win = tk.Toplevel(self)
        win.title("分割标注图预览")
        win.geometry("900x700")
        _preview_bg = '#1C1C1E' if is_dark_theme() else '#E6E7E8'
        _preview_fg = '#F5F5F7' if is_dark_theme() else '#1D1D1F'
        canvas = tk.Canvas(win, bg=_preview_bg)
        sb = ttk.Scrollbar(win, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor='nw')
        frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        thumbs = []
        for i, path in enumerate(self._last_annotate_paths):
            if not os.path.exists(path):
                continue
            try:
                img = Image.open(path)
                img.thumbnail((400, 500), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                thumbs.append(photo)
                col = i % 2
                row = i // 2
                lbl = tk.Label(frame, image=photo, bg=_preview_bg, cursor='hand2')
                lbl.grid(row=row, column=col, padx=10, pady=10)
                lbl.bind('<Button-1>', lambda e, p=path: self._open_full_image(p))
                tk.Label(frame, text=os.path.basename(path), fg=_preview_fg, bg=_preview_bg,
                         wraplength=400).grid(row=row + 1, column=col, pady=(0, 10))
            except Exception:
                pass
        win.configure(bg=_preview_bg)

    def _open_full_image(self, path):
        """打开完整标注图查看"""
        try:
            img = Image.open(path)
            viewer = tk.Toplevel(self)
            viewer.title(os.path.basename(path))
            canvas = tk.Canvas(viewer, bg='#000')
            canvas.pack(fill='both', expand=True)
            # 自适应窗口
            def fit_img(event=None):
                w, h = viewer.winfo_width(), viewer.winfo_height()
                if w < 10 or h < 10:
                    return
                ratio = min(w / img.width, h / img.height, 1.0)
                nw, nh = int(img.width * ratio), int(img.height * ratio)
                resized = img.resize((nw, nh), Image.LANCZOS)
                photo = ImageTk.PhotoImage(resized)
                canvas.delete('all')
                canvas.create_image(w // 2, h // 2, image=photo)
                canvas._photo = photo
            viewer.bind('<Configure>', fit_img)
            viewer.geometry(f"{min(img.width, 1000)}x{min(img.height, 800)}")
        except Exception as e:
            messagebox.showerror("错误", f"打开图片失败: {e}")



# ============================================================
# 设置
# ============================================================
class SettingsTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app

        title = ttk.Label(self.content, text="系统设置", font=('微软雅黑', 14, 'bold'))
        title.pack(pady=(16, 8))

        # ---- 界面设置 ----
        ui_frame = ttk.LabelFrame(self.content, text="界面设置", padding=12)
        ui_frame.pack(fill='x', padx=30, pady=8)

        ttk.Label(ui_frame, text="主题模式:").grid(row=0, column=0, sticky='e', padx=4, pady=6)
        _saved_theme = get_setting('app_theme', 'system')
        _theme_display = {'light': '白天', 'dark': '黑暗', 'system': '跟随系统'}.get(_saved_theme, '跟随系统')
        self.theme_var = tk.StringVar(value=_theme_display)
        theme_options = ['白天', '黑暗', '跟随系统']
        self.theme_cb = ttk.Combobox(ui_frame, textvariable=self.theme_var,
                                     values=theme_options, width=12, state='readonly')
        self.theme_cb.grid(row=0, column=1, sticky='w', padx=4, pady=6)

        ttk.Button(ui_frame, text="保存主题设置", command=self.save_theme_setting).grid(
            row=0, column=2, sticky='w', padx=12, pady=6)

        ttk.Label(ui_frame, text="保存后立即生效，无需重启",
                 font=('微软雅黑', 8), foreground='#888').grid(
            row=0, column=3, sticky='w', padx=4, pady=6)

        # ---- 日志设置 ----
        log_frame = ttk.LabelFrame(self.content, text="日志设置", padding=12)
        log_frame.pack(fill='x', padx=30, pady=8)

        ttk.Label(log_frame, text="日志保留份数:").grid(row=0, column=0, sticky='e', padx=4, pady=6)
        _log_keep = get_setting('log_keep_count', '10')
        self.log_keep_var = tk.StringVar(value=_log_keep)
        ttk.Spinbox(log_frame, from_=1, to=100, width=6, textvariable=self.log_keep_var).grid(
            row=0, column=1, sticky='w', padx=4, pady=6)
        ttk.Label(log_frame, text="份（超出自动覆盖最旧日志）",
                 font=('微软雅黑', 8), foreground='#888').grid(
            row=0, column=2, sticky='w', padx=4, pady=6)

        ttk.Label(log_frame, text="日志保存路径:").grid(row=1, column=0, sticky='e', padx=4, pady=6)
        _log_path = get_setting('log_path', '')
        if not _log_path:
            _log_path = os.path.join(get_base_dir(), 'log')
        self.log_path_var = tk.StringVar(value=_log_path)
        ttk.Entry(log_frame, textvariable=self.log_path_var, width=50).grid(
            row=1, column=1, columnspan=2, sticky='we', padx=4, pady=6)
        ttk.Button(log_frame, text="浏览", command=self.choose_log_path).grid(
            row=1, column=3, sticky='w', padx=4, pady=6)

        ttk.Button(log_frame, text="保存日志设置", command=self.save_log_settings).grid(
            row=2, column=1, sticky='w', padx=4, pady=8)
        ttk.Button(log_frame, text="打开日志文件夹", command=self.open_log_folder).grid(
            row=2, column=2, sticky='w', padx=4, pady=8)


        # ---- 关于 ----
        about_frame = ttk.LabelFrame(self.content, text="关于", padding=12)
        about_frame.pack(fill='x', padx=30, pady=8)
        ttk.Label(about_frame, text="财务管理系统", font=('微软雅黑', 12, 'bold')).pack(anchor='w', pady=2)
        ttk.Label(about_frame, text="功能：人员管理、发票报销、工资结算、银行账户、统计汇总、数据备份",
                 font=('微软雅黑', 9), foreground='#666').pack(anchor='w', pady=1)

    def save_theme_setting(self):
        theme_map = {'白天': 'light', '黑暗': 'dark', '跟随系统': 'system'}
        val = theme_map.get(self.theme_var.get(), 'system')
        conn = get_db()
        conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES('app_theme', ?)", (val,))
        conn.commit()
        conn.close()
        # 实时切换主题，无需重启
        self.app.apply_theme()
        messagebox.showinfo("保存成功", f"主题已切换为：{self.theme_var.get()}")

    def choose_log_path(self):
        path = filedialog.askdirectory(title="选择日志保存文件夹", initialdir=self.log_path_var.get())
        if path:
            self.log_path_var.set(path)

    def save_log_settings(self):
        try:
            keep = int(self.log_keep_var.get())
            if keep < 1:
                keep = 1
        except (ValueError, TypeError):
            keep = 10
        path = self.log_path_var.get().strip()
        if not path:
            path = os.path.join(get_base_dir(), 'log')
        try:
            os.makedirs(path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建日志文件夹：{e}")
            return
        set_setting('log_keep_count', str(keep))
        set_setting('log_path', path)
        log_op('保存日志设置', f'保留{keep}份, 路径={path}')
        self.app.set_status(f"日志设置已保存：保留{keep}份，路径={path}")
        messagebox.showinfo("保存成功", f"日志设置已保存\n保留份数：{keep}\n保存路径：{path}")

    def open_log_folder(self):
        path = self.log_path_var.get().strip()
        if not path:
            path = os.path.join(get_base_dir(), 'log')
        try:
            os.makedirs(path, exist_ok=True)
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
            log_op('打开日志文件夹', path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：{e}")





# ============================================================
# 数据备份与恢复
# ============================================================
class BackupTab(ScrollableTab):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app

        # 标题
        title = ttk.Label(self.content, text="数据备份与恢复", font=('微软雅黑', 14, 'bold'))
        title.pack(pady=(16, 8))

        # 数据统计卡片
        stat_frame = ttk.LabelFrame(self.content, text="当前数据统计", padding=16)
        stat_frame.pack(fill='x', padx=30, pady=8)

        self.stats_vars = {}
        stats = [
            ('emp_count', '人员数量', '人'),
            ('inv_count', '发票记录', '条'),
            ('inv_unpaid', '未报销金额', '元'),
            ('sal_count', '工资发放记录', '条'),
            ('img_count', '发票图片', '张'),
            ('db_size', '数据库大小', 'KB'),
        ]
        for i, (key, label, unit) in enumerate(stats):
            row = i // 3
            col = i % 3
            ttk.Label(stat_frame, text=label + "：", font=('微软雅黑', 10)).grid(
                row=row, column=col * 2, sticky='e', padx=(20, 4), pady=6)
            var = tk.StringVar(value="-")
            self.stats_vars[key] = var
            ttk.Label(stat_frame, textvariable=var, font=('微软雅黑', 10, 'bold'),
                      foreground='#217346').grid(row=row, column=col * 2 + 1, sticky='w', pady=6)

        # 操作按钮
        btn_frame = ttk.Frame(self.content)
        btn_frame.pack(pady=20)

        self.backup_btn = ttk.Button(btn_frame, text="📦 备份数据", command=self.do_backup)
        self.backup_btn.pack(side='left', padx=12, ipadx=20, ipady=8)

        self.restore_btn = ttk.Button(btn_frame, text="📂 从备份恢复", command=self.do_restore)
        self.restore_btn.pack(side='left', padx=12, ipadx=20, ipady=8)

        self.refresh_btn = ttk.Button(btn_frame, text="🔄 刷新统计", command=self.refresh_stats)
        self.refresh_btn.pack(side='left', padx=12, ipadx=20, ipady=8)

        # 自动备份设置
        auto_frame = ttk.LabelFrame(self.content, text="自动备份设置", padding=12)
        auto_frame.pack(fill='x', padx=30, pady=8)

        self.auto_enabled_var = tk.BooleanVar(value=False)
        FlatCheckbutton(auto_frame, text="启用自动备份", variable=self.auto_enabled_var,
                        command=self.on_auto_toggle).grid(row=0, column=0, columnspan=2, sticky='w', padx=4, pady=4)

        ttk.Label(auto_frame, text="备份间隔:").grid(row=1, column=0, sticky='e', padx=4, pady=4)
        self.interval_var = tk.StringVar(value='每1小时')
        interval_options = ['每30分钟', '每1小时', '每2小时', '每4小时', '每8小时', '每天', '每周']
        self.interval_cb = ttk.Combobox(auto_frame, textvariable=self.interval_var,
                                        values=interval_options, width=12, state='readonly')
        self.interval_cb.grid(row=1, column=1, sticky='w', padx=4, pady=4)

        ttk.Label(auto_frame, text="保留备份数:").grid(row=1, column=2, sticky='e', padx=4, pady=4)
        self.keep_count_var = tk.StringVar(value='10')
        ttk.Spinbox(auto_frame, from_=1, to=100, width=6, textvariable=self.keep_count_var).grid(
            row=1, column=3, sticky='w', padx=4, pady=4)

        ttk.Label(auto_frame, text="备份路径:").grid(row=2, column=0, sticky='e', padx=4, pady=4)
        self.backup_path_var = tk.StringVar()
        ttk.Entry(auto_frame, textvariable=self.backup_path_var, width=40).grid(
            row=2, column=1, columnspan=2, sticky='we', padx=4, pady=4)
        ttk.Button(auto_frame, text="浏览", command=self.choose_backup_path).grid(
            row=2, column=3, sticky='w', padx=4, pady=4)

        ttk.Button(auto_frame, text="保存设置", command=self.save_auto_settings).grid(
            row=3, column=0, sticky='w', padx=4, pady=6)

        self.auto_status_var = tk.StringVar(value="自动备份未启用")
        ttk.Label(auto_frame, textvariable=self.auto_status_var, foreground='#217346',
                 font=('微软雅黑', 9)).grid(row=3, column=1, columnspan=3, sticky='w', padx=4, pady=6)

        # 定时器相关
        self._auto_timer = None
        self._next_backup_time = None

        # 说明
        info = ttk.LabelFrame(self.content, text="说明", padding=12)
        info.pack(fill='x', padx=30, pady=8)
        tips = [
            "• 备份会将数据库（人员、发票、工资记录、银行账户）和发票图片打包为一个 .zip 文件",
            "• 建议定期备份，备份文件可保存到U盘、网盘等安全位置",
            "• 恢复数据会覆盖当前所有数据，请谨慎操作，恢复前建议先备份当前数据",
            "• 数据文件（finance.db）和发票图片（invoices文件夹）位于程序同目录",
            "• 自动备份启用后，程序运行期间按设定间隔自动备份，超出保留数量自动删除最早的备份",
        ]
        for tip in tips:
            ttk.Label(info, text=tip, font=('微软雅黑', 9), foreground='#555').pack(anchor='w', pady=2)

        self.load_auto_settings()
        self.refresh_stats()

    def refresh_stats(self):
        """刷新数据统计"""
        try:
            conn = get_db()
            emp_count = conn.execute("SELECT COUNT(*) c FROM employees").fetchone()['c']
            inv_count = conn.execute("SELECT COUNT(*) c FROM invoices").fetchone()['c']
            inv_unpaid = conn.execute("SELECT COALESCE(SUM(amount),0) s FROM invoices WHERE status=0").fetchone()['s']
            sal_count = conn.execute("SELECT COUNT(*) c FROM salary_payments").fetchone()['c']
            conn.close()
        except Exception:
            emp_count = inv_count = sal_count = 0
            inv_unpaid = 0

        # 发票文件数量（图片+PDF）
        img_count = 0
        if os.path.isdir(INVOICE_DIR):
            img_count = len([f for f in os.listdir(INVOICE_DIR)
                             if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.pdf'))])

        # 数据库大小
        db_size = 0
        if os.path.exists(DB_PATH):
            db_size = os.path.getsize(DB_PATH) / 1024

        self.stats_vars['emp_count'].set(f"{emp_count} 人")
        self.stats_vars['inv_count'].set(f"{inv_count} 条")
        self.stats_vars['inv_unpaid'].set(f"¥{inv_unpaid:,.2f}")
        self.stats_vars['sal_count'].set(f"{sal_count} 条")
        self.stats_vars['img_count'].set(f"{img_count} 张")
        self.stats_vars['db_size'].set(f"{db_size:.1f} KB")

    def _backup_to_file(self, path):
        """核心备份逻辑：将数据打包到指定zip路径"""
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # 添加数据库
            if os.path.exists(DB_PATH):
                zf.write(DB_PATH, 'finance.db')
            # 添加发票图片
            if os.path.isdir(INVOICE_DIR):
                for fname in os.listdir(INVOICE_DIR):
                    fpath = os.path.join(INVOICE_DIR, fname)
                    if os.path.isfile(fpath):
                        zf.write(fpath, f'invoices/{fname}')
            # 添加清单
            manifest = (f"财务管理系统数据备份\n"
                        f"备份时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"版本：1.0\n")
            zf.writestr('backup_manifest.txt', manifest)
        return True

    def do_backup(self):
        """备份数据到zip"""
        if not os.path.exists(DB_PATH):
            messagebox.showwarning("提示", "数据库文件不存在，无需备份")
            return

        default_name = f"财务数据备份_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        path = filedialog.asksaveasfilename(
            title="备份数据到", defaultextension=".zip",
            initialfile=default_name,
            filetypes=[("备份文件", "*.zip")])
        if not path:
            return

        try:
            self._backup_to_file(path)
            size_mb = os.path.getsize(path) / (1024 * 1024)
            self.app.set_status(f"备份成功: {os.path.basename(path)}")
            messagebox.showinfo("备份成功",
                                f"数据已备份到：\n{path}\n\n文件大小：{size_mb:.2f} MB")
        except Exception as e:
            messagebox.showerror("备份失败", str(e))

    def do_restore(self):
        """从备份zip恢复数据"""
        path = filedialog.askopenfilename(
            title="选择备份文件", filetypes=[("备份文件", "*.zip")])
        if not path:
            return

        # 验证备份文件
        try:
            with zipfile.ZipFile(path, 'r') as zf:
                names = zf.namelist()
                if 'finance.db' not in names:
                    messagebox.showerror("恢复失败", "该备份文件中未找到 finance.db，不是有效的数据备份")
                    return
        except zipfile.BadZipFile:
            messagebox.showerror("恢复失败", "文件不是有效的zip备份文件")
            return
        except Exception as e:
            messagebox.showerror("恢复失败", str(e))
            return

        # 确认
        if not messagebox.askyesno("确认恢复",
                                   "恢复数据将覆盖当前所有数据（人员、发票、工资记录、图片）！\n\n"
                                   "建议先备份当前数据。\n\n确定要继续恢复吗？"):
            return

        try:
            # 重置各页面选中状态
            self.app.emp_tab.selected_id = None
            self.app.inv_tab.selected_id = None
            self.app.sal_tab.selected_payment_id = None
            self.app.bank_tab.selected_account_id = None
            self.app.bank_tab.editing_account_id = None
            self.app.bank_tab.editing_tx_id = None

            # 先备份当前数据为 .bak
            if os.path.exists(DB_PATH):
                bak_path = DB_PATH + '.bak'
                shutil.copy2(DB_PATH, bak_path)

            # 解压
            extract_dir = get_base_dir()
            with zipfile.ZipFile(path, 'r') as zf:
                for name in zf.namelist():
                    if name == 'backup_manifest.txt':
                        continue
                    # 安全检查：防止路径穿越
                    if name.startswith('..') or name.startswith('/') or ':' in name:
                        continue
                    out_path = os.path.join(extract_dir, name)
                    if name.endswith('/'):
                        os.makedirs(out_path, exist_ok=True)
                        continue
                    os.makedirs(os.path.dirname(out_path), exist_ok=True)
                    with zf.open(name) as src, open(out_path, 'wb') as dst:
                        shutil.copyfileobj(src, dst)

            # 重新初始化数据库（确保表结构完整）
            init_db()

            # 刷新所有页面（带异常保护）
            try:
                self.app.emp_tab.refresh()
            except Exception:
                pass
            try:
                self.app.inv_tab.refresh()
            except Exception:
                pass
            try:
                self.app.sal_tab.refresh()
            except Exception:
                pass
            try:
                self.app.stat_tab.refresh()
            except Exception:
                pass
            try:
                self.app.bank_tab.refresh_accounts()
            except Exception:
                pass
            try:
                self.refresh_stats()
            except Exception:
                pass

            self.app.set_status("数据恢复成功")
            messagebox.showinfo("恢复成功", "数据已从备份恢复完成。\n\n（原数据库已备份为 finance.db.bak）")
        except Exception as e:
            messagebox.showerror("恢复失败", f"{str(e)}\n\n原数据已保留为 finance.db.bak，可手动恢复。")

    # ===== 自动备份相关方法 =====
    def load_auto_settings(self):
        """加载自动备份设置"""
        enabled = get_setting('auto_backup_enabled', '0') == '1'
        interval = get_setting('auto_backup_interval', '每1小时')
        keep_count = get_setting('auto_backup_keep_count', '10')
        backup_path = get_setting('auto_backup_path', '')
        self.auto_enabled_var.set(enabled)
        self.interval_var.set(interval)
        self.keep_count_var.set(keep_count)
        if not backup_path:
            # 默认备份路径为程序同目录下的 backup 文件夹
            backup_path = os.path.join(get_base_dir(), 'backup')
        self.backup_path_var.set(backup_path)
        self._update_auto_status()
        if enabled:
            self._schedule_next_backup()

    def save_theme_setting(self):
        """保存主题设置"""
        theme_display = self.theme_var.get()
        if theme_display == '白天':
            theme_value = 'light'
        elif theme_display == '黑暗':
            theme_value = 'dark'
        else:
            theme_value = 'system'
        set_setting('app_theme', theme_value)
        # 实时切换主题，无需重启
        if hasattr(self.app, 'apply_theme'):
            self.app.apply_theme()
        messagebox.showinfo("主题设置", f"已切换为「{theme_display}」模式。")

    def save_auto_settings(self):
        """保存自动备份设置"""
        try:
            keep_count = int(self.keep_count_var.get())
            if keep_count < 1:
                keep_count = 1
        except ValueError:
            keep_count = 10
            self.keep_count_var.set('10')

        backup_path = self.backup_path_var.get().strip()
        if not backup_path:
            backup_path = os.path.join(get_base_dir(), 'backup')
            self.backup_path_var.set(backup_path)

        # 确保备份目录存在
        try:
            os.makedirs(backup_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建备份目录：{e}")
            return

        set_setting('auto_backup_enabled', '1' if self.auto_enabled_var.get() else '0')
        set_setting('auto_backup_interval', self.interval_var.get())
        set_setting('auto_backup_keep_count', str(keep_count))
        set_setting('auto_backup_path', backup_path)

        self._update_auto_status()
        # 重新调度
        self._cancel_timer()
        if self.auto_enabled_var.get():
            self._schedule_next_backup()

        messagebox.showinfo("保存成功", "自动备份设置已保存。")

    def choose_backup_path(self):
        """选择备份目录"""
        path = filedialog.askdirectory(title="选择备份文件保存目录")
        if path:
            self.backup_path_var.set(path)

    def on_auto_toggle(self):
        """启用/禁用自动备份切换"""
        self._update_auto_status()

    def _get_interval_seconds(self):
        """获取备份间隔秒数"""
        mapping = {
            '每30分钟': 1800,
            '每1小时': 3600,
            '每2小时': 7200,
            '每4小时': 14400,
            '每8小时': 28800,
            '每天': 86400,
            '每周': 604800,
        }
        return mapping.get(self.interval_var.get(), 3600)

    def _update_auto_status(self):
        """更新自动备份状态显示"""
        if self.auto_enabled_var.get():
            interval = self.interval_var.get()
            keep = self.keep_count_var.get()
            path = self.backup_path_var.get()
            if self._next_backup_time:
                next_str = self._next_backup_time.strftime('%Y-%m-%d %H:%M:%S')
                self.auto_status_var.set(f"自动备份已启用 | 间隔：{interval} | 保留：{keep}份 | 下次：{next_str}")
            else:
                self.auto_status_var.set(f"自动备份已启用 | 间隔：{interval} | 保留：{keep}份 | 路径：{path}")
        else:
            self.auto_status_var.set("自动备份未启用")

    def _cancel_timer(self):
        """取消定时器"""
        if self._auto_timer:
            try:
                self.app.after_cancel(self._auto_timer)
            except Exception:
                pass
            self._auto_timer = None

    def _schedule_next_backup(self):
        """调度下次自动备份"""
        self._cancel_timer()
        if not self.auto_enabled_var.get():
            return
        interval_sec = self._get_interval_seconds()
        self._next_backup_time = datetime.now() + timedelta(seconds=interval_sec)
        self._update_auto_status()
        # tkinter after 用毫秒
        self._auto_timer = self.app.after(interval_sec * 1000, self._do_auto_backup)

    def _do_auto_backup(self):
        """执行自动备份"""
        if not self.auto_enabled_var.get():
            return
        try:
            backup_path = self.backup_path_var.get().strip()
            if not backup_path:
                backup_path = os.path.join(get_base_dir(), 'backup')
            os.makedirs(backup_path, exist_ok=True)

            filename = f"自动备份_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            full_path = os.path.join(backup_path, filename)

            if os.path.exists(DB_PATH):
                self._backup_to_file(full_path)
                self.app.set_status(f"自动备份完成: {filename}")
                # 轮转备份
                self._rotate_backups(backup_path)
            else:
                self.app.set_status("自动备份跳过：数据库文件不存在")
        except Exception as e:
            self.app.set_status(f"自动备份失败: {e}")

        # 调度下次
        self._schedule_next_backup()

    def _rotate_backups(self, backup_path):
        """轮转备份文件，超出保留数量删除最早的"""
        try:
            keep_count = int(self.keep_count_var.get())
        except ValueError:
            keep_count = 10

        # 列出所有自动备份zip文件，按修改时间排序
        backups = []
        if os.path.isdir(backup_path):
            for f in os.listdir(backup_path):
                if f.startswith('自动备份_') and f.endswith('.zip'):
                    full = os.path.join(backup_path, f)
                    if os.path.isfile(full):
                        backups.append((os.path.getmtime(full), full))

        # 按时间升序（最早的在前）
        backups.sort(key=lambda x: x[0])

        # 删除超出数量的最早备份
        while len(backups) > keep_count:
            _, old_path = backups.pop(0)
            try:
                os.remove(old_path)
            except Exception:
                pass

    def stop_auto_backup(self):
        """停止自动备份（程序退出时调用）"""
        self._cancel_timer()


# ============================================================
# 入口
# ============================================================
if __name__ == '__main__':
    # 崩溃日志钩子（便于排查打包后问题）
    def _excepthook(tp, val, tb):
        import traceback
        try:
            with open(os.path.join(get_base_dir(), 'crash_log.txt'), 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now()}]\n")
                f.write(''.join(traceback.format_exception(tp, val, tb)))
                f.write("\n")
        except Exception:
            pass
        # 同时写入运行日志
        try:
            LogManager.error('程序崩溃', val)
        except Exception:
            pass
    sys.excepthook = _excepthook

    # Tkinter回调异常处理（捕获after回调和事件绑定中的异常，避免闪退）
    _orig_report_callback_exception = tk.Tk.report_callback_exception
    def _report_callback_exception(self, exc, val, tb):
        import traceback
        try:
            with open(os.path.join(get_base_dir(), 'crash_log.txt'), 'a', encoding='utf-8') as f:
                f.write(f"[{datetime.now()}] [Tkinter回调异常]\n")
                f.write(''.join(traceback.format_exception(exc, val, tb)))
                f.write("\n")
        except Exception:
            pass
        try:
            LogManager.error('Tkinter回调异常', val)
        except Exception:
            pass
        # 不重新抛出异常，避免程序闪退
    tk.Tk.report_callback_exception = _report_callback_exception

    # === 先初始化数据库（读取主题设置需要） ===
    init_db()

    # 根据主题设置确定启动界面颜色
    _splash_is_dark = is_dark_theme()
    if _splash_is_dark:
        _sp_bg = '#0A0A0B'
        _sp_title_fg = '#FFFFFF'
        _sp_sub_fg = '#6E6E73'
        _sp_trough = '#1C1C1E'
        _sp_accent = '#3B82F6'
        _sp_version_fg = '#48484A'
    else:
        _sp_bg = '#E6E7E8'
        _sp_title_fg = '#1D1D1F'
        _sp_sub_fg = '#6E6E73'
        _sp_trough = '#D1D1D6'
        _sp_accent = '#007AFF'
        _sp_version_fg = '#A1A1A6'

    # === 启动界面 ===
    splash = tk.Tk()
    splash.withdraw()  # 立即隐藏窗口，这是避免闪白的关键
    splash.overrideredirect(True)
    splash_width = 460
    splash_height = 280
    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = (screen_w - splash_width) // 2
    y = (screen_h - splash_height) // 2
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
    splash.configure(bg=_sp_bg)
    splash.attributes('-topmost', True)
    splash.attributes('-alpha', 0.0)  # 完全透明，deiconify后不会闪白

    # 启动界面内容
    splash_frame = tk.Frame(splash, bg=_sp_bg, width=splash_width, height=splash_height)
    splash_frame.pack(fill='both', expand=True)
    splash_frame.pack_propagate(False)

    tk.Label(splash_frame, text="财务管理系统", font=('微软雅黑', 22, 'bold'),
             bg=_sp_bg, fg=_sp_title_fg).pack(pady=(50, 4))
    tk.Label(splash_frame, text="Finance Management System", font=('微软雅黑', 9),
             bg=_sp_bg, fg=_sp_sub_fg).pack(pady=(0, 30))

    progress_frame = tk.Frame(splash_frame, bg=_sp_bg)
    progress_frame.pack(fill='x', padx=50)
    progress_style = ttk.Style()
    progress_style.theme_use('clam')
    progress_style.configure('Splash.Horizontal.TProgressbar',
                              troughcolor=_sp_trough, background=_sp_accent,
                              bordercolor=_sp_trough, lightcolor=_sp_accent,
                              darkcolor=_sp_accent, thickness=6)
    progress = ttk.Progressbar(progress_frame, style='Splash.Horizontal.TProgressbar',
                                 mode='determinate', maximum=100, length=360)
    progress.pack(fill='x')

    status_label = tk.Label(splash_frame, text="正在启动...", font=('微软雅黑', 9),
                            bg=_sp_bg, fg=_sp_sub_fg)
    status_label.pack(pady=(12, 0))

    tk.Label(splash_frame, text="v1.0", font=('微软雅黑', 8),
             bg=_sp_bg, fg=_sp_version_fg).pack(side='bottom', pady=10)

    # 所有控件创建完成后，先deiconify（此时已完全透明，不会闪白），再淡入
    splash.update_idletasks()
    splash.update()
    splash.deiconify()  # 显示窗口，此时-alpha=0.0完全透明
    splash.update_idletasks()
    splash.update()
    # 淡入效果：从0.1逐步增加到1.0
    for _alpha in [i / 10.0 for i in range(1, 11)]:
        splash.attributes('-alpha', _alpha)
        splash.update_idletasks()
        splash.update()

    def update_splash(val, text):
        progress['value'] = val
        status_label.config(text=text)
        splash.update_idletasks()
        splash.update()

    # 初始化步骤（手动刷新，不用mainloop避免事件循环冲突）
    update_splash(10, "正在初始化...")

    update_splash(15, "正在初始化日志系统...")
    try:
        LogManager.init()
        log_info('程序启动，日志系统初始化完成')
    except Exception as _le:
        print(f'日志初始化失败: {_le}')

    update_splash(35, "正在加载界面组件...")
    app = FinanceApp()
    app.withdraw()
    try:
        app.attributes('-alpha', 0.0)  # 完全透明，避免创建时闪白色
    except Exception:
        pass

    update_splash(65, "正在加载业务数据...")
    app.update_idletasks()

    update_splash(90, "正在准备运行环境...")
    # 短暂停留让用户看到完成状态
    for _ in range(15):
        splash.update()
        splash.after(30)
        splash.update()

    update_splash(100, "启动完成")
    splash.update()
    splash.after(200)
    splash.update()

    # 确保启动界面完全关闭后再显示主窗口
    try:
        splash.update_idletasks()
        splash.update()
        splash.destroy()
    except Exception:
        pass
    # 等待一下，确保启动界面完全销毁
    try:
        app.update_idletasks()
        app.update()
    except Exception:
        pass
    # 恢复默认root窗口（splash作为第一个Tk曾是默认root，销毁后需指定app为默认root）
    try:
        tk._default_root = app
    except Exception:
        pass

    # 确保主窗口是完全透明的，然后再显示
    try:
        app.attributes('-alpha', 0.0)
    except Exception:
        pass
    app.deiconify()
    app.title("财务管理系统")
    app.geometry("1280x820")
    app.minsize(1150, 760)
    app.lift()
    app.focus_force()
    # 主窗口淡入显示，避免闪白色
    app.update_idletasks()
    app.update()
    for _alpha in [i / 10.0 for i in range(1, 11)]:
        try:
            app.attributes('-alpha', _alpha)
            app.update_idletasks()
            app.update()
        except Exception:
            break
    try:
        app.mainloop()
    finally:
        try:
            LogManager.close()
        except Exception:
            pass

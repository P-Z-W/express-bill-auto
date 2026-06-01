# -*- coding: utf-8 -*-
"""
公共工具函数库 V2.4
==================================================
【架构说明】
  统一管理所有模块共用的工具函数和样式常量
  各业务模块从此文件导入，消除重复代码

【包含内容】
  1. ensure_folder()     统一文件夹创建
  2. read_excel()        统一Excel读取（带异常捕获）
  3. clean_dataframe()   统一数据清洗
  4. get_last_month_str() 获取上月字符串（YYYYMM格式）
  5. openpyxl 样式常量   统一边框/字体/对齐样式

【使用方式】
  from utils import ensure_folder, read_excel, clean_dataframe
  from utils import FULL_BORDER, HEADER_FONT, CENTER_ALIGN
==================================================
"""
import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# ====================== 1. 文件夹管理 ======================
def ensure_folder(folder_path):
    """
    确保文件夹存在，不存在则自动创建
    统一替代各模块的 ensure_folder / ensure_output_dir / init_folder
    """
    os.makedirs(folder_path, exist_ok=True)


# ====================== 2. Excel 读取 ======================
def read_excel(file_path, nrows=None):
    """
    统一读取 Excel 文件，带完整异常捕获
    nrows=None 读全部，nrows=0 只读表头（用于快速识别列名）
    返回 DataFrame 或 None（文件不存在/读取失败）
    """
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在：{file_path}")
        return None
    try:
        df = pd.read_excel(file_path, engine="openpyxl", nrows=nrows)
        return df
    except Exception as e:
        print(f"❌ 读取失败：{file_path}，错误：{str(e)}")
        return None


# ====================== 3. 数据清洗 ======================
def clean_dataframe(df):
    """
    通用数据清洗：去全空行、去重、重置索引
    统一替代各模块的 clean_dataframe
    """
    if df is None or df.empty:
        return None
    df = df.dropna(how="all")
    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    return df


# ====================== 4. 日期工具 ======================
def get_last_month_str():
    """
    获取上个月的字符串，格式 YYYYMM
    例：6月运行 → 返回 "202605"
    统一替代 order_db.py 里自己计算的版本
    直接从 settings 读取，保证和其他模块月份一致
    """
    from config import settings
    # PROCESS_MONTH 格式是 "2026-05"，去掉横线得到 "202605"
    return settings.PROCESS_MONTH.replace("-", "")


# ====================== 5. openpyxl 样式常量 ======================
# 统一替代各模块重复定义的样式，保持全项目Excel风格一致

# 细边框
_THIN_SIDE  = Side(style="thin", color="000000")
FULL_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE,
    top=_THIN_SIDE,  bottom=_THIN_SIDE
)

# 表头样式
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

# 普通加粗
BOLD_FONT   = Font(bold=True, size=11)

# 对齐方式
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center")
RIGHT_ALIGN  = Alignment(horizontal="right",  vertical="center")
WRAP_CENTER_ALIGN = Alignment(wrap_text=True, horizontal="center", vertical="center")
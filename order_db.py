# -*- coding: utf-8 -*-
"""
数据库订单下载模块（独立可运行）
==================================================
功能：
  1. 从阿里云 MySQL 数据库下载订单数据
  2. 读取 config/SQL-config.txt 中的查询语句
  3. 导出 Excel 到 output 文件夹
  4. 自动美化表格（高速版，不循环全表，只做关键美化）
  5. 独立运行 / 被 main.py 调用均可

文件结构：
  1. 配置区       → 数据库、路径固定配置
  2. 工具函数区   → 文件夹创建、时间计算
  3. 数据库操作区 → 连接、读取SQL、查询
  4. 导出美化区   → 高速导出 + 轻量美化
  5. 主流程区     → 对外提供 run_download_orders() 接口
==================================================
"""
import pymysql
import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ========================== 【配置区】固定配置 ==========================
# 阿里云数据库连接信息
DB_CONFIG = {
    "host": "yiboerp.rwlb.rds.aliyuncs.com",
    "port": 3306,
    "user": "only_read",
    "password": "1oEb4y(lBb09F",
    "database": "yibo",
    "charset": "utf8mb4"
}

# 路径规范（项目统一结构）
CONFIG_FOLDER = "config"          # 配置文件目录
OUTPUT_FOLDER = "output"          # 输出文件目录
SQL_FILE_PATH = os.path.join(CONFIG_FOLDER, "SQL-config.txt")  # SQL语句路径

# ========================== 【工具函数区】 ==========================
def ensure_folder(folder_path):
    """
    确保文件夹存在，不存在则自动创建
    :param folder_path: 文件夹路径
    """
    os.makedirs(folder_path, exist_ok=True)

def get_last_month_str():
    """
    获取上个月日期字符串，格式：YYYYMM
    例：202604
    """
    today = datetime.now()
    first_day_this_month = datetime(today.year, today.month, 1)
    last_day_last_month = first_day_this_month - timedelta(days=1)
    return last_day_last_month.strftime("%Y%m")

# ========================== 【数据库操作区】 ==========================
def connect_db():
    """
    创建并返回数据库连接
    :return: 连接对象 / None
    """
    try:
        conn = pymysql.connect(**DB_CONFIG)
        print("✅ 数据库连接成功")
        return conn
    except Exception as e:
        print(f"❌ 数据库连接失败：{str(e)}")
        return None

def load_sql_from_config():
    """
    从 config/SQL-config.txt 读取SQL语句
    :return: SQL字符串 / None
    """
    try:
        with open(SQL_FILE_PATH, "r", encoding="utf-8") as f:
            sql = f.read().strip()
        print("✅ 成功读取外部SQL配置文件")
        return sql
    except Exception as e:
        print(f"❌ 读取SQL文件失败：{str(e)}")
        return None

def query_database(conn, sql):
    """
    执行SQL查询，返回DataFrame
    :param conn: 数据库连接
    :param sql: 查询语句
    :return: 查询结果DataFrame
    """
    df = pd.read_sql(sql, conn)
    print(f"📊 查询完成，获取数据行数：{len(df)}")
    return df

# ========================== 【导出美化区】（高速美化 ✅） ==========================
def export_and_pretty(df):
    """
    导出Excel + 高速美化（只美化关键部分，不循环全表，速度极快）
    美化内容：
      1. 表头加粗、居中
      2. 全表边框
      3. 日期列格式化
      4. 运单号设为文本格式
      5. 自动适配列宽（仅前100行估算，大幅提速）
    """
    # 确保输出目录存在
    ensure_folder(OUTPUT_FOLDER)

    # 生成文件名
    month_str = get_last_month_str()
    filename = f"毅播快递数据_{month_str}.xlsx"
    save_path = os.path.join(OUTPUT_FOLDER, filename)

    # 1. 快速导出Excel
    with pd.ExcelWriter(save_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False)

    # 2. 打开文件进行轻量美化
    wb = load_workbook(save_path)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column

    # 样式定义（只创建一次）
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    # -------------------- 高速全表边框（不逐行循环） --------------------
    last_col = chr(ord('A') + max_col - 1)
    data_range = f"A1:{last_col}{max_row}"
    for row in ws[data_range]:
        for cell in row:
            cell.border = border

    # -------------------- 表头美化 --------------------
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = center_align

    # -------------------- 日期列格式化（解决 ##### 显示问题） --------------------
    if max_col >= 2:
        ws.column_dimensions['B'].width = 20
        for cell in ws['B']:
            cell.number_format = "yyyy-mm-dd hh:mm:ss"

    # -------------------- 运单号设为文本格式 --------------------
    for cell in ws['A']:
        cell.number_format = "@"

    # -------------------- 自动列宽（前100行估算，超快） --------------------
    for col_idx in range(1, max_col + 1):
        col_letter = chr(ord('A') + col_idx - 1)
        max_len = 0
        # 只遍历前100行，避免十万行循环
        for cell in ws[col_letter][:100]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 25)

    # 保存文件
    wb.save(save_path)
    print(f"💾 导出并美化完成：{filename}")
    return save_path

# ========================== 【主流程区】对外接口 ==========================
def run_download_orders():
    """
    模块主流程：
    连接数据库 → 读取SQL → 查询数据 → 导出并美化 → 返回数据
    可供 main.py 调用
    """
    conn = None
    try:
        # 1. 连接数据库
        conn = connect_db()
        if not conn:
            return None

        # 2. 读取SQL语句
        sql = load_sql_from_config()
        if not sql:
            return None

        # 3. 执行查询
        df = query_database(conn, sql)
        if df.empty:
            print("⚠️ 查询结果为空，无需导出")
            return None

        # 4. 导出并美化
        export_and_pretty(df)
        return df

    except Exception as e:
        print(f"❌ 流程执行失败：{str(e)}")
        return None

    finally:
        # 无论是否异常，确保关闭连接
        if conn:
            conn.close()
            print("🔌 数据库连接已关闭")

# ========================== 独立运行入口 ==========================
if __name__ == "__main__":
    run_download_orders()
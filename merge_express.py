# -*- coding: utf-8 -*-
"""
快递单合并模块（V1.4 稳定版）
==================================================
【框架说明】
  功能：读取data目录下申通、中通原始账单
       → 清洗格式 → 统一字段 → 合并成一张总表
       → 输出到 output/清洗合并总账单.xlsx
  结构：工具函数 → 数据清洗 → 分快递处理 → 合并 → 导出美化
【可独立运行】
==================================================
"""
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 【仅修改这一行】
from config import settings

# ====================== 路径配置（从settings读取）======================
DATA_FOLDER    = settings.DATA_FOLDER
OUTPUT_FOLDER  = settings.OUTPUT_FOLDER
OUTPUT_FILE    = os.path.join(OUTPUT_FOLDER, settings.EXPRESS_OUTPUT_FILE)

# ====================== 工具函数 ======================
def ensure_output_dir():
    """确保输出文件夹存在"""
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

def read_excel_file(file_path):
    """读取Excel文件，异常捕获"""
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在：{file_path}")
        return None
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        return df
    except Exception as e:
        print(f"❌ 读取失败：{file_path}, 错误：{str(e)}")
        return None

# ====================== 数据清洗 ======================
def clean_dataframe(df):
    """通用清洗：去空行、去重、重置索引"""
    if df is None or df.empty:
        return None
    df = df.dropna(how="all")
    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    return df

# ====================== 申通账单处理 ======================
def process_shentong(df):
    """申通格式标准化"""
    df = clean_dataframe(df)
    if df is None:
        return None

    use_cols = ["业务时间", "运单号", "订单目的省份", "订单目的城市", "结算重量"]
    df = df[use_cols]

    df.rename(columns={
        "订单目的省份": "目的省份",
        "订单目的城市": "目的城市"
    }, inplace=True)

    df["运单号"] = df["运单号"].astype(str).str.strip()
    df["快递类型"] = "申通"
    df["所属团队"] = ""
    return df

# ====================== 中通账单处理 ======================
def process_zhongtong(df):
    """中通格式标准化"""
    df = clean_dataframe(df)
    if df is None:
        return None

    use_cols = ["扫描时间", "运单号", "目的地", "目的地", "结算重量"]
    df = df[use_cols]
    df.columns = ["业务时间", "运单号", "目的省份", "目的城市", "结算重量"]

    df["运单号"] = df["运单号"].astype(str).str.strip()
    df["快递类型"] = "中通"
    df["所属团队"] = ""
    return df

# ====================== 合并两张账单 ======================
def merge_express_bills(df_st, df_zt):
    """合并申通、中通数据，统一字段顺序"""
    final_columns = ["业务时间", "运单号", "目的省份", "目的城市", "结算重量", "快递类型", "所属团队"]
    merged = pd.concat([df_st, df_zt], ignore_index=True)
    merged = merged[final_columns]
    merged = clean_dataframe(merged)
    merged = merged.dropna(subset=["运单号"])
    return merged

# ====================== 导出并美化 ======================
def save_and_style(df):
    """导出Excel并添加表头、边框、居中样式"""
    ensure_output_dir()
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 表头加粗居中
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # 数据行加边框
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border

    wb.save(OUTPUT_FILE)
    print(f"✅ 合并完成，文件已生成：{OUTPUT_FILE}")

# ====================== 主流程 ======================
def run_merge_process():
    """快递合并主流程"""
    print("🚀 开始合并快递账单...")

    file_st = os.path.join(DATA_FOLDER, settings.EXPRESS_INPUT_ST)
    file_zt = os.path.join(DATA_FOLDER, settings.EXPRESS_INPUT_ZT)

    df_st = read_excel_file(file_st)
    df_zt = read_excel_file(file_zt)

    df_st_clean = process_shentong(df_st)
    df_zt_clean = process_zhongtong(df_zt)

    df_all = merge_express_bills(df_st_clean, df_zt_clean)
    save_and_style(df_all)
    return df_all

# 独立运行入口
if __name__ == "__main__":
    run_merge_process()
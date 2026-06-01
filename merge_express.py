# -*- coding: utf-8 -*-
"""
快递单合并模块（V2.1 稳定版）
==================================================
【框架说明】
  功能：读取 data/YYYY-MM/ 目录下申通、中通原始账单
       → 清洗格式 → 统一字段 → 合并成一张总表
       → 输出到 output/YYYY-MM/清洗合并总账单.xlsx
  结构：工具函数 → 数据清洗 → 分快递处理 → 合并 → 导出美化

【V2.1 新增】
  - 自动扫描 data/YYYY-MM/ 目录，按列名特征识别申通/中通
  - 不再依赖固定文件名，文件叫任何名字都能自动识别
  - 识别逻辑：申通含"业务时间"列，中通含"扫描时间"列
  - 支持任意一方为空（只有申通或只有中通也能跑）

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
DATA_FOLDER   = settings.DATA_FOLDER
OUTPUT_FOLDER = settings.OUTPUT_FOLDER
OUTPUT_FILE   = os.path.join(OUTPUT_FOLDER, settings.EXPRESS_OUTPUT_FILE)

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

# ====================== V2.1新增：自动扫描识别文件 ======================
def auto_find_express_files(data_folder):
    """
    自动扫描 data/YYYY-MM/ 目录，按列名特征识别申通/中通文件
    业务逻辑：
      - 申通账单含列"业务时间" → 识别为申通
      - 中通账单含列"扫描时间" → 识别为中通
    无论文件名叫什么都能自动匹配，无需每月重命名
    返回：(申通文件路径 或 None, 中通文件路径 或 None)
    """
    if not os.path.exists(data_folder):
        print(f"❌ 数据目录不存在：{data_folder}")
        return None, None

    # 扫描目录下所有 xlsx 文件，排除 Excel 临时文件（~开头）
    xlsx_files = [
        os.path.join(data_folder, f)
        for f in os.listdir(data_folder)
        if f.endswith(".xlsx") and not f.startswith("~")
    ]

    if not xlsx_files:
        print(f"❌ {data_folder} 目录下没有找到任何xlsx文件")
        return None, None

    file_st, file_zt = None, None

    for fpath in xlsx_files:
        try:
            # 只读表头，速度快，不加载全部数据
            df_head = pd.read_excel(fpath, engine="openpyxl", nrows=0)
            cols = df_head.columns.tolist()

            if "业务时间" in cols and file_st is None:
                file_st = fpath
                print(f"✅ 识别到申通账单：{os.path.basename(fpath)}")
            elif "扫描时间" in cols and file_zt is None:
                file_zt = fpath
                print(f"✅ 识别到中通账单：{os.path.basename(fpath)}")

        except Exception as e:
            print(f"⚠️  读取文件头失败，跳过：{os.path.basename(fpath)}，错误：{str(e)}")
            continue

    if file_st is None:
        print(f"❌ 未识别到申通账单（需含'业务时间'列），请检查 {data_folder} 目录")
    if file_zt is None:
        print(f"❌ 未识别到中通账单（需含'扫描时间'列），请检查 {data_folder} 目录")

    return file_st, file_zt

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

    df["运单号"]   = df["运单号"].astype(str).str.strip()
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

    df["运单号"]   = df["运单号"].astype(str).str.strip()
    df["快递类型"] = "中通"
    df["所属团队"] = ""
    return df

# ====================== 合并两张账单 ======================
def merge_express_bills(df_st, df_zt):
    """
    合并申通、中通数据，统一字段顺序
    支持任意一方为空：只有申通或只有中通也能正常运行
    """
    final_columns = ["业务时间", "运单号", "目的省份", "目的城市", "结算重量", "快递类型", "所属团队"]

    dfs = [d for d in [df_st, df_zt] if d is not None]
    if not dfs:
        raise ValueError("申通和中通账单均为空，无法合并")

    merged = pd.concat(dfs, ignore_index=True)
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

    thin   = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 表头加粗居中
    for cell in ws[1]:
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border

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

    # V2.1：自动扫描识别文件，无需固定文件名
    file_st, file_zt = auto_find_express_files(DATA_FOLDER)

    df_st = read_excel_file(file_st) if file_st else None
    df_zt = read_excel_file(file_zt) if file_zt else None

    df_st_clean = process_shentong(df_st)  if df_st is not None else None
    df_zt_clean = process_zhongtong(df_zt) if df_zt is not None else None

    df_all = merge_express_bills(df_st_clean, df_zt_clean)
    save_and_style(df_all)
    return df_all

# 独立运行入口
if __name__ == "__main__":
    run_merge_process()
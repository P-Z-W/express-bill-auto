# -*- coding: utf-8 -*-
"""
快递单合并模块：读取、清洗、合并、美化格式（速度优化版）
"""
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ====================== 自动创建output文件夹 ======================
def ensure_output_dir():
    if not os.path.exists("output"):
        os.makedirs("output")

# ====================== 读取Excel（提速：低内存模式） ======================
def read_excel(file_path):
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在：{file_path}")
        return None

    try:
        # 低内存加载，大数据提速明显
        df = pd.read_excel(file_path, engine="openpyxl", dtype_backend="numpy_nullable")
        print(f"✅ 读取成功：{os.path.basename(file_path)} | 行数：{len(df)}")
        return df
    except Exception as e:
        print(f"❌ 读取失败：{str(e)}")
        return None

# ====================== 深度清洗（提速：向量化，不循环） ======================
def deep_clean_df(df):
    if df is None or df.empty:
        return None

    df = df.dropna(how="all")
    df = df.drop_duplicates()
    df = df[df.count(axis=1) > 1]
    df = df.reset_index(drop=True)
    return df

# ====================== 处理申通（提速：减少拷贝） ======================
def handle_shentong_df(df_st):
    df_clean = deep_clean_df(df_st)
    use_cols = ["业务时间", "运单号", "订单目的省份", "订单目的城市", "结算重量"]
    df_use = df_clean[use_cols]

    df_use.rename(columns={
        "订单目的省份": "目的省份",
        "订单目的城市": "目的城市"
    }, inplace=True)

    df_use["运单号"] = df_use["运单号"].astype("string").str.strip()
    df_use["快递"] = "申通"
    df_use["团队"] = ""
    return df_use

# ====================== 处理中通（提速：减少拷贝） ======================
def handle_zhongtong_df(df_zt):
    df_clean = deep_clean_df(df_zt)
    use_cols = ["扫描时间", "运单号", "目的地", "目的地市", "结算重量"]
    df_use = df_clean[use_cols]

    df_use.rename(columns={
        "扫描时间": "业务时间",
        "目的地": "目的省份",
        "目的地市": "目的城市"
    }, inplace=True)

    df_use["运单号"] = df_use["运单号"].astype("string").str.strip()
    df_use["快递"] = "中通"
    df_use["团队"] = ""
    return df_use

# ====================== 合并表格（提速：最小化处理） ======================
def merge_two_df(df_st_final, df_zt_final):
    final_cols = ["业务时间", "运单号", "目的省份", "目的城市", "结算重量", "快递", "团队"]
    df_total = pd.concat([df_st_final, df_zt_final], ignore_index=True)
    df_total = df_total[final_cols]
    df_total = deep_clean_df(df_total)
    df_total = df_total.dropna(subset=["运单号"])
    return df_total

# ====================== Excel美化（提速：只循环有效区域） ======================
def save_and_style_excel(df_total, output_path):
    # 快速保存
    with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
        df_total.to_excel(writer, index=False)

    wb = load_workbook(output_path)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column

    # 样式只创建一次
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    # 表头一次性设置
    for cell in ws[1]:
        cell.font = bold_font
        cell.alignment = center_align
        cell.border = border

    # 批量边框（提速：不判断空，直接整列设置）
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = border

    # 运单号文本格式
    for cell in ws["B"]:
        cell.number_format = "@"

    # 结算重量数字格式
    for cell in ws["E"]:
        cell.number_format = "0.00"

    # 列宽
    widths = {"A": 22, "B": 25, "C": 15, "D": 15, "E": 12, "F": 10, "G": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    wb.save(output_path)
    print(f"💾 美化完成：{output_path}")

# ====================== 主流程 ======================
def run_merge_process():
    print("=" * 60)
    print("        中通+申通快递合并模块（提速优化版）")
    print("=" * 60)

    ensure_output_dir()

    path_st = "data/2026年4月钟村毅播云仓对帐单.xlsx"
    path_zt = "data/咸辣甜服饰-4月份.xlsx"

    df_st_raw = read_excel(path_st)
    df_zt_raw = read_excel(path_zt)

    df_st_done = handle_shentong_df(df_st_raw)
    df_zt_done = handle_zhongtong_df(df_zt_raw)

    df_all = merge_two_df(df_st_done, df_zt_done)
    save_and_style_excel(df_all, "output/清洗合并总账单.xlsx")

    print(f"\n🎉 合并完成！总数据：{len(df_all)} 条")
    return df_all
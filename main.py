# -*- coding: utf-8 -*-
"""
快递对账自动化系统
模块化结构：
1. 读取Excel
2. 数据清洗
3. 主流程
"""
import pandas as pd
import os

# ==========================================
# 函数1：读取 Excel 文件（单独模块）
# ==========================================
def read_excel(file_path):
    """
    功能：仅读取Excel，不做任何清洗
    返回：原始数据 DataFrame
    """
    if not os.path.exists(file_path):
        print(f"❌ 文件不存在：{file_path}")
        return None

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        print("✅ 读取成功")
        print(f"📊 原始数据条数：{len(df)}")
        return df
    except Exception as e:
        print(f"❌ 读取失败：{e}")
        return None

# ==========================================
# 函数2：数据清洗（单独模块）
# ==========================================
def clean_data(df):
    """
    功能：仅清洗数据，不负责读取
    返回：清洗后的数据
    """
    if df is None or df.empty:
        return None

    print("\n🧹 开始数据清洗...")

    # 去完全重复行
    df = df.drop_duplicates()

    # 去全空行
    df = df.dropna(how="all")

    # 重置索引
    df = df.reset_index(drop=True)

    print(f"✅ 清洗完成 → 剩余 {len(df)} 条")
    return df

# ==========================================
# 主程序：流程控制
# ==========================================
def main():
    print("=" * 50)
    print("   快递对账自动化系统 V1.2")
    print("=" * 50)

    # 1. 读取
    file_path = "data/2026年4月钟村毅播云仓对帐单.xlsx"
    df_original = read_excel(file_path)

    # 2. 清洗
    df_clean = clean_data(df_original)

    # 3. 完成
    if df_clean is not None:
        print("\n🎉 流程完成：读取 → 清洗 全部成功！")

if __name__ == "__main__":
    main()
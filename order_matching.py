# -*- coding: utf-8 -*-
"""
快递订单对账匹配模块（V1.4）
==================================================
【框架说明】
  功能：运单号匹配 → 标注是否匹配 → 回填团队名称 → 导出结果
  依赖：先运行 merge_express.py 生成清洗合并总账单
        先运行 order_db.py 生成毅播快递数据_YYYYMM.xlsx
  输出：output/最终对账结果.xlsx
  可独立运行
==================================================
"""
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# 导入项目统一配置
from config import settings

# ====================== 路径配置（从settings读取）======================
OUTPUT_FOLDER = settings.OUTPUT_FOLDER
EXPRESS_FILE = os.path.join(OUTPUT_FOLDER, settings.EXPRESS_OUTPUT_FILE)
RESULT_FILE = os.path.join(OUTPUT_FOLDER, settings.RESULT_FILE)


# ====================== 工具函数 ======================
def ensure_folder(folder_path):
    """确保输出文件夹存在"""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def load_source_data():
    """读取两份数据源，并统一运单号格式"""
    if not os.path.exists(EXPRESS_FILE):
        print(f"❌ 未找到清洗合并总账单，请先运行 merge_express.py")
        return None, None
    df_express = pd.read_excel(EXPRESS_FILE, engine="openpyxl")
    print(f"✅ 成功读取清洗合并总账单，共 {len(df_express)} 条快递记录")

    order_files = [f for f in os.listdir(OUTPUT_FOLDER)
                   if f.startswith(settings.ORDER_FILE_PREFIX)]
    if not order_files:
        print(f"❌ 未找到订单数据文件，请先运行 order_db.py")
        return None, None

    latest_order_file = order_files[-1]
    order_file_path = os.path.join(OUTPUT_FOLDER, latest_order_file)
    df_order = pd.read_excel(order_file_path, engine="openpyxl")
    print(f"✅ 成功读取订单数据文件：{latest_order_file}，共 {len(df_order)} 条订单记录")

    return df_express, df_order


# ====================== 核心匹配逻辑 ======================
def match_team_by_waybill():
    df_express, df_order = load_source_data()
    if df_express is None or df_order is None:
        return None

    # 关键修复：处理快递账单的运单号，去除小数后缀
    df_express["运单号"] = df_express["运单号"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df_order["运单号"] = df_order["运单号"].astype(str).str.strip()

    # 构建映射关系
    waybill_team_map = df_order[["运单号", "所属团队"]].drop_duplicates("运单号").set_index("运单号")[
        "所属团队"].to_dict()

    # 回填团队名称
    df_express["所属团队"] = df_express["运单号"].map(waybill_team_map)
    df_express["所属团队"] = df_express["所属团队"].fillna("未匹配")

    matched_count = len(df_express[df_express["所属团队"] != "未匹配"])
    print(f"✅ 匹配完成：共 {len(df_express)} 条，成功匹配 {matched_count} 条，未匹配 {len(df_express) - matched_count} 条")

    return df_express


# ====================== 导出与美化（全固定列宽版）======================
def export_styled_result(df_result):
    ensure_folder(OUTPUT_FOLDER)
    df_result.to_excel(RESULT_FILE, index=False, engine="openpyxl")

    wb = load_workbook(RESULT_FILE)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column

    # 样式定义
    thin_border = Side(style="thin")
    border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    # 1. 全表添加边框（高效写法）
    last_col = chr(ord('A') + max_col - 1)
    for row in ws[f"A1:{last_col}{max_row}"]:
        for cell in row:
            cell.border = border

    # 2. 表头样式
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = center_align

    # 3. 【全固定列宽】按你表格字段设置的合理宽度
    # A列：业务时间
    ws.column_dimensions['A'].width = 20
    # B列：运单号
    ws.column_dimensions['B'].width = 18
    # C列：目的省份
    ws.column_dimensions['C'].width = 12
    # D列：目的城市
    ws.column_dimensions['D'].width = 12
    # E列：结算重量
    ws.column_dimensions['E'].width = 10
    # F列：快递类型
    ws.column_dimensions['F'].width = 12
    # G列：所属团队
    ws.column_dimensions['G'].width = 16

    wb.save(RESULT_FILE)
    print(f"✅ 最终对账结果已保存：{RESULT_FILE}")


# ====================== 主流程 ======================
def run_reconciliation():
    print("=" * 60)
    print("📌 启动快递订单对账匹配程序（V1.4）")
    print("=" * 60)

    matched_data = match_team_by_waybill()
    if matched_data is not None:
        export_styled_result(matched_data)
        print("\n🎉 所有对账匹配流程执行完毕！")


if __name__ == "__main__":
    run_reconciliation()
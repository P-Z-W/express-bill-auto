# -*- coding: utf-8 -*-
"""
客户对账账单拆分 V2.1 稳定版
==================================================
【核心功能】
1. 按团队拆分账单，自动生成标准化Excel格式
2. 运费核算：全国均重/单票/新西1-3kg/新西1kg内加收费用计算
3. 专项统计：新西1kg订单数、1%占比判断、应结算单数核算
4. 文件命名：{合计金额}-{团队名}_{业务年月}_快递加收费.xlsx
5. 0费用过滤：普通团队合计金额为0不生成文件

【V2.0版本更新】（完整保留）
1. 表格样式优化：加收费汇总标题横跨L→S全列，排版更规整
2. 千耀传媒专项优化：双条件筛选+快递类型列仅全国均重显示
3. 专项统计表下移适配+完整闭合边框
4. 千耀传媒金额0强制输出，其他团队保持原过滤规则

【V2.1 变更】
- SOURCE_FILE / SAVE_DIR 全部走 settings，支持月份子目录
- output/YYYY-MM/最终对账结果.xlsx → 读取源文件
- output/YYYY-MM/客户账单/ → 输出拆分账单
- 其余所有逻辑、样式、计算完全保留原V2.0框架
==================================================
"""
import os
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from config import settings

# ===================== 全局配置（V2.1：路径走settings）=====================
# V2.1 核心改动：SOURCE_FILE 和 SAVE_DIR 改为从 settings 读取，支持月份子目录
SOURCE_FILE = os.path.join(settings.OUTPUT_FOLDER, settings.RESULT_FILE)          # output/YYYY-MM/最终对账结果.xlsx
CONFIG_FILE = os.path.join(settings.CONFIG_FOLDER, "price_config.xlsx")           # config/price_config.xlsx
SAVE_DIR    = os.path.join(settings.OUTPUT_FOLDER, "客户账单")                    # output/YYYY-MM/客户账单/

# 以下配置保持原框架不变
GROUP_COL      = "所属团队"
FILTER_INVALID = True
TEST_TEAMS     = []          # 白名单：空=全量运行
SPECIAL_TEAM   = "千耀传媒"  # 特殊团队标记

# ===================== 全局样式（无修改，保留原框架）=====================
THIN_SIDE      = Side(style="thin", color="000000")
FULL_BORDER    = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_FONT    = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL    = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
BOLD_FONT      = Font(bold=True, size=11)
RED_BOLD_FONT  = Font(bold=True, color="FF0000")
CENTER_ALIGN      = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN        = Alignment(horizontal="left",   vertical="center")
RIGHT_ALIGN       = Alignment(horizontal="right",  vertical="center")
WRAP_CENTER_ALIGN = Alignment(wrap_text=True, horizontal="center", vertical="center")

COL_WIDTH = {
    'A': 20, 'B': 18, 'C': 14, 'D': 12, 'E': 8,
    'F': 12, 'G': 12, 'H': 13, 'I': 13,
    'L': 8,  'M': 18, 'N': 12, 'O': 14,
    'P': 14, 'Q': 12, 'R': 15, 'S': 15
}
ROW_HEIGHT  = 24
TEAM_CONFIG = {}

# ===================== 工具函数（无修改，保留原框架）=====================
def init_folder():
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

def safe_file_name(raw_name: str) -> str:
    illegal_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    name = str(raw_name)
    for char in illegal_chars:
        name = name.replace(char, "_")
    return name

def load_team_config():
    global TEAM_CONFIG
    if not os.path.exists(CONFIG_FILE):
        print(f"【警告】配置文件不存在 {CONFIG_FILE}")
        return
    try:
        df = pd.read_excel(CONFIG_FILE, engine="openpyxl", sheet_name="客户快递加收单价信息记录")
        for _, row in df.iterrows():
            team = str(row.iloc[3]).strip()
            if not team or team == "nan":
                continue
            st_avg = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
            zt_avg = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0
            st_fee = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0
            zt_fee = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
            TEAM_CONFIG[team] = (st_avg, zt_avg, st_fee, zt_fee)
        print(f"【配置】加载 {len(TEAM_CONFIG)} 个团队规则")
    except Exception as e:
        print(f"【错误】读取配置失败: {str(e)}")

def get_team_rule(team_name: str) -> tuple:
    return TEAM_CONFIG.get(team_name, (0.0, 0.0, 0.0, 0.0))

def set_worksheet_style(ws, max_row: int, max_col: int):
    for col, width in COL_WIDTH.items():
        ws.column_dimensions[col].width = width
    for row in range(1, max_row + 1):
        ws.row_dimensions[row].height = ROW_HEIGHT
    for row in range(1, max_row + 1):
        ws[f"B{row}"].number_format = "@"
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border    = FULL_BORDER
            cell.alignment = CENTER_ALIGN
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

# ===================== 核心汇总统计（完整保留V2.0逻辑）=====================
def add_summary_area(ws, team_df: pd.DataFrame, team_name: str) -> float:
    st_avg_all, zt_avg_all, st_fee_all, zt_fee_all = get_team_rule(team_name)
    print(f"\n【调试】当前团队:{team_name} 申通:{st_avg_all}/{st_fee_all} 中通:{zt_avg_all}/{zt_fee_all}")

    col_start = 12
    col_end   = 19  # 标题横跨L→S全列

    # 基础统计
    total_order  = len(team_df)
    xixi_order   = len(team_df[(team_df["结算重量"] < 1) & (team_df["目的省份"].str.contains("新疆|西藏", na=False))])
    threshold    = total_order * 0.01
    is_over_flag = "是" if xixi_order >= threshold else "否"
    settle_order = math.ceil(xixi_order - threshold) if is_over_flag == "是" else 0
    r11_value    = settle_order * 10

    # 汇总大标题（横跨L→S）
    ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
    title_cell            = ws.cell(row=1, column=col_start)
    title_cell.value      = "加收费汇总"
    title_cell.font       = HEADER_FONT
    title_cell.fill       = HEADER_FILL
    title_cell.alignment  = CENTER_ALIGN
    title_cell.border     = FULL_BORDER

    # 表头分支
    if team_name == SPECIAL_TEAM:
        summary_headers = ["序号","实际计算方式","快递类型","发货单量","结算重量","平均重量","超出重量","应付金额"]
        calc_list = [
            ("全国均重","申通"),
            ("全国均重","中通"),
            ("单票","申通"),
            ("新西1-3公斤","申通"),
            ("新西1kg内（包1%）","中通"),
            ("合计","")
        ]
    else:
        summary_headers = ["序号","实际计算方式","发货单量","结算重量","平均重量","超出重量","应付金额"]
        calc_list = [
            ("全国均重",""),
            ("单票",""),
            ("新西1-3公斤",""),
            ("新西1kg内（包1%）",""),
            ("合计","")
        ]

    # 写入表头
    for idx, text in enumerate(summary_headers):
        cell           = ws.cell(row=2, column=col_start + idx)
        cell.value     = text
        cell.border    = FULL_BORDER
        cell.alignment = CENTER_ALIGN
        cell.font      = BOLD_FONT

    summary_rows  = []
    total_fee_all = 0.0

    # ========== 【千耀传媒】核心逻辑（完整保留V2.0：双条件筛选）==========
    if team_name == SPECIAL_TEAM:
        for idx, (calc_name, express_type) in enumerate(calc_list):
            seq          = idx + 1
            order_cnt    = 0
            total_weight = 0.0
            avg_weight   = 0.0
            exceed_weight = 0.0
            row_fee      = 0.0

            # 双条件筛选：实际计算方式 + 快递类型
            if calc_name != "合计":
                filter_df = team_df[(team_df["实际计算方式"] == calc_name) & (team_df["快递类型"] == express_type)]
                order_cnt = len(filter_df)
                if order_cnt > 0:
                    total_weight = round(filter_df["结算重量"].sum(), 2)
                    avg_weight   = round(total_weight / order_cnt, 2)

            # 按快递类型匹配配置
            if express_type == "申通":
                use_avg, use_fee = st_avg_all, st_fee_all
            elif express_type == "中通":
                use_avg, use_fee = zt_avg_all, zt_fee_all
            else:
                use_avg, use_fee = st_avg_all, st_fee_all

            # 超出重量计算（负数自动置0）
            if calc_name == "全国均重":
                exceed_weight   = max(round(avg_weight - use_avg, 2), 0.0)
                single_fee      = round((exceed_weight / 0.1) * use_fee, 2)
                row_fee         = round(single_fee * order_cnt, 2)
                total_fee_all  += row_fee
            elif calc_name in ("单票","新西1-3公斤"):
                fee_sum = 0.0
                for _, row in filter_df.iterrows():
                    w        = float(row["结算重量"]) if pd.notna(row["结算重量"]) else 0.0
                    diff     = max(w - use_avg, 0.0)
                    fee_sum += round((diff / 0.1) * use_fee, 2)
                row_fee        = round(fee_sum, 2)
                total_fee_all += row_fee
            elif calc_name == "新西1kg内（包1%）":
                row_fee        = r11_value
                total_fee_all += row_fee

            if calc_name == "新西1kg内（包1%）":
                row_data = [seq, calc_name, express_type, "", "", "", "", row_fee]
            elif calc_name == "合计":
                row_data = [seq, calc_name, "", total_order, "", "", "", "=SUM(S3:S7)"]
            else:
                row_data = [seq, calc_name, express_type, order_cnt, total_weight, avg_weight, exceed_weight, row_fee]

            if calc_name in ("单票","新西1-3公斤","新西1kg内（包1%）","合计"):
                row_data[4] = row_data[5] = row_data[6] = ""

            summary_rows.append(row_data)

    else:
        # 普通团队：原逻辑完全不变
        for idx, (calc_name, _) in enumerate(calc_list):
            seq           = idx + 1
            order_cnt     = 0
            total_weight  = 0.0
            avg_weight    = 0.0
            exceed_weight = 0.0
            row_fee       = 0.0

            if calc_name != "合计":
                filter_df = team_df[team_df["实际计算方式"] == calc_name]
                order_cnt = len(filter_df)
                if order_cnt > 0:
                    total_weight = round(filter_df["结算重量"].sum(), 2)
                    avg_weight   = round(total_weight / order_cnt, 2)

            if calc_name == "全国均重":
                exceed_weight  = max(round(avg_weight - st_avg_all, 2), 0.0)
                single_fee     = round((exceed_weight / 0.1) * st_fee_all, 2)
                row_fee        = round(single_fee * order_cnt, 2)
                total_fee_all += row_fee
            elif calc_name in ("单票","新西1-3公斤"):
                fee_sum = 0.0
                for _, row in filter_df.iterrows():
                    w        = float(row["结算重量"]) if pd.notna(row["结算重量"]) else 0.0
                    diff     = max(w - st_avg_all, 0.0)
                    fee_sum += round((diff / 0.1) * st_fee_all, 2)
                row_fee        = round(fee_sum, 2)
                total_fee_all += row_fee
            elif calc_name == "新西1kg内（包1%）":
                row_fee        = r11_value
                total_fee_all += row_fee

            if calc_name == "新西1kg内（包1%）":
                row_data = [seq, calc_name, "", "", "", "", row_fee]
            elif calc_name == "合计":
                row_data = [seq, calc_name, total_order, "", "", "", "=SUM(R3:R6)"]
            else:
                row_data = [seq, calc_name, order_cnt, total_weight, avg_weight, exceed_weight, row_fee]

            if calc_name in ("单票","新西1-3公斤","新西1kg内（包1%）","合计"):
                row_data[3] = row_data[4] = row_data[5] = ""

            summary_rows.append(row_data)

    total_fee_all = round(total_fee_all, 2)

    # 写入汇总表内容
    for row_idx, row_data in enumerate(summary_rows, start=3):
        for col_idx, val in enumerate(row_data):
            cell        = ws.cell(row=row_idx, column=col_start + col_idx)
            cell.value  = val
            cell.border = FULL_BORDER
            if col_idx == 0:
                cell.alignment = CENTER_ALIGN
            elif col_idx == 1:
                cell.alignment = LEFT_ALIGN
            else:
                cell.alignment = RIGHT_ALIGN
            if row_data[1] == "合计":
                if col_idx == 1:
                    cell.font      = BOLD_FONT
                    cell.alignment = CENTER_ALIGN
                if (team_name == SPECIAL_TEAM and col_idx == 7) or (team_name != SPECIAL_TEAM and col_idx == 6):
                    cell.font = RED_BOLD_FONT
        ws.row_dimensions[row_idx].height = ROW_HEIGHT

    # ========== 专项统计表（完整保留V2.0：下移+完整边框）==========
    if team_name == SPECIAL_TEAM:
        stat_row_t1, stat_row_t2, stat_row_data = 10, 11, 12
        stat_headers = [
            "序号","实际计算方式","快递类型","总发货单量","新西1kg内订单数",
            "是否超1%","新西1kg内应结算单数","新西1kg内应付金额"
        ]
    else:
        stat_row_t1, stat_row_t2, stat_row_data = 9, 10, 11
        stat_headers = [
            "序号","实际计算方式","总发货单量","新西1kg内订单数",
            "是否超1%","新西1kg内应结算单数","新西1kg内应付金额"
        ]

    for idx in range(len(stat_headers)):
        col = col_start + idx
        ws.merge_cells(start_row=stat_row_t1, start_column=col, end_row=stat_row_t2, end_column=col)
        cell           = ws.cell(row=stat_row_t1, column=col)
        cell.value     = stat_headers[idx]
        cell.alignment = WRAP_CENTER_ALIGN
        cell.font      = BOLD_FONT
        cell.border    = FULL_BORDER

    if team_name == SPECIAL_TEAM:
        cell           = ws.cell(row=stat_row_data, column=col_start+2, value="中通")
        cell.alignment = CENTER_ALIGN
        cell.border    = FULL_BORDER

    stat_data = ["1","新西1kg内（包1%）", total_order, xixi_order, is_over_flag, settle_order, r11_value]
    if team_name == SPECIAL_TEAM:
        stat_data.insert(2, "")

    for idx, val in enumerate(stat_data):
        cell        = ws.cell(row=stat_row_data, column=col_start + idx)
        cell.value  = val
        cell.border = FULL_BORDER
        if idx == 0:
            cell.alignment = CENTER_ALIGN
        elif idx == 1:
            cell.alignment = LEFT_ALIGN
        else:
            cell.alignment = RIGHT_ALIGN

    # 专项表外框闭合
    for row in range(stat_row_t1, stat_row_data + 1):
        for col in range(col_start, col_start + len(stat_headers)):
            ws.cell(row=row, column=col).border = FULL_BORDER

    ws.row_dimensions[stat_row_t1].height  = ROW_HEIGHT
    ws.row_dimensions[stat_row_t2].height  = ROW_HEIGHT
    ws.row_dimensions[stat_row_data].height = ROW_HEIGHT

    print("✅ 汇总表、专项统计表写入完成")
    return total_fee_all

# ===================== 数据读取与文件拆分（无修改）=====================
def load_source_data() -> pd.DataFrame:
    if not os.path.exists(SOURCE_FILE):
        raise FileNotFoundError(f"原始数据文件不存在 {SOURCE_FILE}")
    df = pd.read_excel(SOURCE_FILE, engine="openpyxl")
    print(f"【数据】原始条数：{len(df)}")
    if FILTER_INVALID:
        df = df[(df[GROUP_COL].notna()) & (df[GROUP_COL] != "未匹配")]
        print(f"【数据】有效条数：{len(df)}")
    df["运单号"] = df["运单号"].astype(str)
    return df

def split_by_group(df: pd.DataFrame):
    group_list = df.groupby(GROUP_COL)
    print(f"\n【分组】共识别 {len(group_list)} 个团队")
    print("===== 开始拆分账单 =====")

    for team_name, team_data in group_list:
        if TEST_TEAMS and team_name not in TEST_TEAMS:
            print(f"【跳过】{team_name}（白名单外团队）")
            continue

        safe_team_name = safe_file_name(team_name)
        temp_path      = os.path.join(SAVE_DIR, f"temp_{safe_team_name}.xlsx")
        team_data.to_excel(temp_path, index=False, engine="openpyxl")

        wb = load_workbook(temp_path)
        ws = wb.active
        set_worksheet_style(ws, ws.max_row, ws.max_column)
        total_fee = add_summary_area(ws, team_data, team_name)
        wb.save(temp_path)
        wb.close()

        # 提取业务年月
        try:
            team_data["业务时间"] = pd.to_datetime(team_data["业务时间"])
            if len(team_data) >= 2:
                target_time = team_data.iloc[1]["业务时间"]
            else:
                target_time = team_data.iloc[0]["业务时间"]
            year_month = target_time.strftime("%Y-%m")
        except Exception as e:
            print(f"【警告】{team_name} 业务年月异常：{str(e)}，使用默认值0000-00")
            year_month = "0000-00"

        # 千耀传媒金额0不跳过
        if abs(total_fee) < 1e-6 and team_name != SPECIAL_TEAM:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            print(f"【跳过】{team_name}：合计应付为0，不生成文件")
            continue

        # 生成最终文件
        fee_str      = f"{total_fee:.2f}"
        new_file_name = f"{fee_str}-{safe_team_name}_{year_month}_快递加收费.xlsx"
        new_path     = os.path.join(SAVE_DIR, new_file_name)
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(temp_path, new_path)

        print(f"✅ 生成完成：{new_file_name} | 单据数：{len(team_data)} | 合计金额：{total_fee}")

# ===================== 程序入口 =====================
def main():
    print("=" * 65)
    print("        客户对账账单拆分程序 V2.1 稳定版")
    print("=" * 65)
    try:
        init_folder()
        load_team_config()
        source_df = load_source_data()
        split_by_group(source_df)
        print("\n" + "=" * 65)
        print("🎉 全部任务执行完毕，可查看输出目录")
        print("=" * 65)
    except Exception as e:
        print(f"\n❌ 运行异常：{str(e)}")

if __name__ == "__main__":
    print("🚀 启动快递账单拆分程序（V2.1）")
    main()
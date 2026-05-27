# -*- coding: utf-8 -*-
"""
客户对账账单拆分 V1.8 完整版
核心功能：
1. 加载团队运费配置，按团队拆分原始对账数据，自动生成标准化Excel
2. 运费核算：全国均重 / 单票 / 新西1-3kg 加收费用计算
4. 专项统计：新西1kg订单数、占比1%判断、应结算单数核算
5. 文件命名：{合计金额}-{团队名}_{业务年月}_快递加收费.xlsx
   - 合计金额：代码内计算（规避Excel公式读取异常）
   - 业务年月：取自【业务时间】列第二行数据，格式 YYYY-MM
6. 新增过滤：合计应付金额为0的团队不生成文件，打印日志提示
切换说明：TEST_TEAMS 为空=全量运行；填入团队名=仅运行指定团队
"""
import os
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ===================== 全局配置 =====================
SOURCE_FILE = "output/最终对账结果.xlsx"
CONFIG_FILE = "config/price_config.xlsx"
SAVE_DIR = "output/客户账单"
GROUP_COL = "所属团队"
BILL_SUFFIX = "2026-05对账单"
FILTER_INVALID = True

# 白名单：空列表=运行全部团队；填写团队名=仅运行指定团队
TEST_TEAMS = []

# ===================== 全局统一样式 =====================
THIN_SIDE = Side(style="thin", color="000000")
FULL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
BOLD_FONT = Font(bold=True, size=11)
RED_BOLD_FONT = Font(bold=True, color="FF0000", size=11)

CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")
RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")
WRAP_CENTER_ALIGN = Alignment(wrap_text=True, horizontal="center", vertical="center")

COL_WIDTH = {
    'A': 20, 'B': 18, 'C': 14, 'D': 12, 'E': 8,
    'F': 12, 'G': 12, 'H': 13, 'I': 13,
    'L': 8, 'M': 18, 'N': 12, 'O': 14,
    'P': 14, 'Q': 12, 'R': 15
}
ROW_HEIGHT = 24
TEAM_CONFIG = {}

# ===================== 工具函数 =====================
def init_folder():
    """创建账单保存目录"""
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)
        print(f"【目录】创建 {SAVE_DIR}")


def safe_file_name(raw_name: str) -> str:
    """过滤文件名非法字符，避免保存失败"""
    illegal_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    name = str(raw_name)
    for char in illegal_chars:
        name = name.replace(char, "_")
    return name


def load_team_config():
    """加载团队运费规则并缓存"""
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
            st_fee = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0
            zt_avg = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0
            zt_fee = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
            TEAM_CONFIG[team] = (st_avg, zt_avg, st_fee, zt_fee)
        print(f"【配置】加载 {len(TEAM_CONFIG)} 个团队规则")
    except Exception as e:
        print(f"【错误】读取配置失败: {str(e)}")


def get_team_rule(team_name: str) -> tuple:
    """获取指定团队运费规则，无配置返回默认0值"""
    return TEAM_CONFIG.get(team_name, (0.0, 0.0, 0.0, 0.0))


def set_worksheet_style(ws, max_row: int, max_col: int):
    """统一工作表样式：列宽、行高、边框、对齐、运单号文本格式"""
    # 设置列宽
    for col, width in COL_WIDTH.items():
        ws.column_dimensions[col].width = width
    # 设置行高
    for row in range(1, max_row + 1):
        ws.row_dimensions[row].height = ROW_HEIGHT
    # 运单号设为文本格式，避免科学计数
    for row in range(1, max_row + 1):
        ws[f"B{row}"].number_format = "@"
    # 全局边框+居中对齐
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border = FULL_BORDER
            cell.alignment = CENTER_ALIGN
    # 表头样式
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

# ===================== 核心汇总统计 =====================
def add_summary_area(ws, team_df: pd.DataFrame, team_name: str) -> float:
    """写入右侧汇总表+新西1kg专项统计表，返回合计应付金额"""
    st_avg, _, st_fee, _ = get_team_rule(team_name)
    print(f"\n【调试】当前团队:{team_name} 申通拉均:{st_avg} 加收单价:{st_fee}")

    col_start = 12
    col_end = 18

    # 基础统计数据
    total_order = len(team_df)
    xixi_order = len(team_df[(team_df["结算重量"] < 1) & (team_df["目的省份"].str.contains("新疆|西藏", na=False))])
    threshold = total_order * 0.01
    is_over_flag = "是" if xixi_order >= threshold else "否"
    settle_order = math.ceil(xixi_order - threshold) if is_over_flag == "是" else 0
    r11_value = settle_order * 10  # R11 金额计算

    # 汇总区域大标题
    ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
    title_cell = ws.cell(row=1, column=col_start)
    title_cell.value = "加收费汇总"
    title_cell.font = HEADER_FONT
    title_cell.fill = HEADER_FILL
    title_cell.alignment = CENTER_ALIGN
    title_cell.border = FULL_BORDER

    # 汇总表表头
    summary_headers = ["序号", "实际计算方式", "发货单量", "结算重量", "平均重量", "超出重量", "应付金额"]
    for idx, text in enumerate(summary_headers):
        cell = ws.cell(row=2, column=col_start + idx)
        cell.value = text
        cell.border = FULL_BORDER
        cell.alignment = CENTER_ALIGN
        cell.font = BOLD_FONT

    # 分类计算运费
    calc_types = ["全国均重", "单票", "新西1-3公斤", "新西1kg内（包1%）", "合计"]
    summary_rows = []
    for idx, calc_name in enumerate(calc_types):
        if calc_name != "合计":
            filter_df = team_df[team_df["实际计算方式"] == calc_name]
            order_cnt = len(filter_df)
            total_weight = round(filter_df["结算重量"].sum(), 2)
            avg_weight = round(total_weight / order_cnt, 2) if order_cnt > 0 else 0.0
        else:
            order_cnt = 0
            total_weight = 0.0
            avg_weight = 0.0

        seq = idx + 1
        exceed_weight = ""
        total_fee = ""

        # 全国均重费用
        if calc_name == "全国均重":
            exceed_weight = round(avg_weight - st_avg, 2)
            exceed_weight = 0.0 if exceed_weight < 0 else exceed_weight
            single_fee = round((exceed_weight / 0.1) * st_fee, 2)
            total_fee = round(single_fee * order_cnt, 2)
            print(f"【调试】全国均重 平均重量:{avg_weight} 超出:{exceed_weight} 总费用:{total_fee}")
        # 单票 / 新西1-3公斤费用
        elif calc_name in ("单票", "新西1-3公斤"):
            fee_sum = 0.0
            for _, row in filter_df.iterrows():
                w = float(row["结算重量"]) if pd.notna(row["结算重量"]) else 0.0
                e = w - st_avg
                e = 0.0 if e < 0 else e
                fee_sum += round((e / 0.1) * st_fee, 2)
            total_fee = round(fee_sum, 2)
            print(f"【调试】{calc_name} 总费用:{total_fee}")

        # 组装行数据
        if calc_name == "新西1kg内（包1%）":
            row_data = [seq, calc_name, "", "", "", "", r11_value]
        elif calc_name == "合计":
            row_data = [seq, calc_name, total_order, "", "", "", "=SUM(R3:R6)"]
        else:
            row_data = [seq, calc_name, order_cnt, total_weight, avg_weight, exceed_weight, total_fee]

        # 清空指定字段
        if calc_name in ("单票", "新西1-3公斤", "新西1kg内（包1%）", "合计"):
            row_data[3] = row_data[4] = row_data[5] = ""

        summary_rows.append(row_data)

    # 代码内计算总金额（用于文件名 & 过滤判断，规避浮点误差）
    total_fee_all = 0.0
    for i in range(4):
        fee = summary_rows[i][6]
        total_fee_all += fee if isinstance(fee, (int, float)) else 0.0
    total_fee_all = round(total_fee_all, 2)
    print(f"【调试】合计应付金额：{total_fee_all}")

    # 写入汇总表数据与样式
    for row_idx, row_data in enumerate(summary_rows, start=3):
        for col_idx, val in enumerate(row_data):
            cell = ws.cell(row=row_idx, column=col_start + col_idx)
            cell.value = val
            cell.border = FULL_BORDER

            # 单元格对齐
            if col_idx == 0:
                cell.alignment = CENTER_ALIGN
            elif col_idx == 1:
                cell.alignment = LEFT_ALIGN
            else:
                cell.alignment = RIGHT_ALIGN

            # 合计行样式
            if row_data[1] == "合计":
                if col_idx == 1:
                    cell.font = BOLD_FONT
                    cell.alignment = CENTER_ALIGN
                if col_idx == 6:
                    cell.font = RED_BOLD_FONT
        ws.row_dimensions[row_idx].height = ROW_HEIGHT

    # 新西1kg专项统计表
    stat_row_t1, stat_row_t2, stat_row_data = 9, 10, 11
    stat_headers = [
        "序号", "实际计算方式", "总发货单量", "新西1kg内订单数",
        "是否超1%", "新西1kg内应结算单数", "新西1kg内（包1%）应付金额"
    ]
    stat_data = ["1", "新西1kg内（包1%）", total_order, xixi_order, is_over_flag, settle_order, r11_value]

    # 标题列合并
    for idx in range(len(stat_headers)):
        col = col_start + idx
        ws.merge_cells(start_row=stat_row_t1, start_column=col, end_row=stat_row_t2, end_column=col)
    # 写入表头
    for idx, text in enumerate(stat_headers):
        cell = ws.cell(row=stat_row_t1, column=col_start + idx)
        cell.value = text
        cell.alignment = WRAP_CENTER_ALIGN
        cell.font = BOLD_FONT
    # 写入数据
    for idx, val in enumerate(stat_data):
        cell = ws.cell(row=stat_row_data, column=col_start + idx)
        cell.value = val
        if idx == 0:
            cell.alignment = CENTER_ALIGN
        elif idx == 1:
            cell.alignment = LEFT_ALIGN
        else:
            cell.alignment = RIGHT_ALIGN

    # 行高与边框
    ws.row_dimensions[stat_row_t1].height = ROW_HEIGHT
    ws.row_dimensions[stat_row_t2].height = ROW_HEIGHT
    ws.row_dimensions[stat_row_data].height = ROW_HEIGHT
    for row in range(stat_row_t1, stat_row_data + 1):
        for col in range(col_start, col_end + 1):
            ws.cell(row=row, column=col).border = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    print("✅ 汇总表、专项统计表写入完成")
    return total_fee_all

# ===================== 数据读取与文件拆分 =====================
def load_source_data() -> pd.DataFrame:
    """读取原始对账数据，过滤无效数据，运单号转为文本"""
    if not os.path.exists(SOURCE_FILE):
        raise FileNotFoundError(f"原始数据文件不存在 {SOURCE_FILE}")
    df = pd.read_excel(SOURCE_FILE, engine="openpyxl")
    print(f"【数据】原始条数：{len(df)}")

    # 过滤无效团队数据
    if FILTER_INVALID:
        df = df[(df[GROUP_COL].notna()) & (df[GROUP_COL] != "未匹配")]
        print(f"【数据】有效条数：{len(df)}")

    df["运单号"] = df["运单号"].astype(str)
    print("【格式】运单号已转为文本")
    return df


def split_by_group(df: pd.DataFrame):
    """按团队拆分数据，生成Excel并按规则命名，过滤费用为0的团队"""
    group_list = df.groupby(GROUP_COL)
    print(f"\n【分组】共识别 {len(group_list)} 个团队")
    print("===== 开始拆分账单 =====")

    for team_name, team_data in group_list:
        # 白名单过滤：空列表则全部运行
        if TEST_TEAMS and team_name not in TEST_TEAMS:
            print(f"【跳过】{team_name}（白名单外团队）")
            continue

        safe_team_name = safe_file_name(team_name)
        # 临时文件中转，避免文件占用
        temp_path = os.path.join(SAVE_DIR, f"temp_{safe_team_name}.xlsx")
        team_data.to_excel(temp_path, index=False, engine="openpyxl")

        # 写入样式、汇总数据，获取合计应付金额
        wb = load_workbook(temp_path)
        ws = wb.active
        set_worksheet_style(ws, ws.max_row, ws.max_column)
        total_fee = add_summary_area(ws, team_data, team_name)
        wb.save(temp_path)
        wb.close()

        # 提取业务年月（取第二行数据，兼容行数不足场景）
        try:
            team_data["业务时间"] = pd.to_datetime(team_data["业务时间"])
            if len(team_data) >= 2:
                target_time = team_data.iloc[1]["业务时间"]
            else:
                target_time = team_data.iloc[0]["业务时间"]
            year_month = target_time.strftime("%Y-%m")
        except Exception as e:
            print(f"【警告】{team_name} 业务年月提取异常：{str(e)}，使用默认值0000-00")
            year_month = "0000-00"

        # 费用为0过滤（兼容浮点误差）
        if abs(total_fee) < 1e-6:
            # 删除临时文件，避免残留
            if os.path.exists(temp_path):
                os.remove(temp_path)
            print(f"【跳过】{team_name}：合计应付金额为{total_fee:.2f}，无费用账单不生成文件")
            continue

        # 拼接最终文件名
        fee_str = f"{total_fee:.2f}"
        new_file_name = f"{fee_str}-{safe_team_name}_{year_month}_快递加收费.xlsx"
        new_path = os.path.join(SAVE_DIR, new_file_name)

        # 覆盖同名文件
        if os.path.exists(new_path):
            os.remove(new_path)
        os.rename(temp_path, new_path)

        print(f"✅ 生成完成：{new_file_name} | 单据数：{len(team_data)} | 合计金额：{total_fee}")

# ===================== 程序入口 =====================
def main():
    print("=" * 65)
    print("        客户对账账单拆分程序 V1.8 完整版")
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
    main()
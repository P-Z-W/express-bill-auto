# -*- coding: utf-8 -*-
"""
快递订单对账匹配模块（V1.7 最终功能版）
==================================================
【功能说明】
  1. 运单号匹配 → 回填所属团队（支持批量匹配，自动清洗运单号格式）
  2. 智能判断计费模式 → 按结算重量+目的省份自动生成「实际计算方式」
  3. 精准计算应付金额 → 读取config/price_config.xlsx，按省份/重量/快递类型差异化计算
  4. 格式化导出结果 → 带边框、表头样式、固定列宽的Excel，便于人工核对
  5. 完善容错机制 → 缺失文件/配置时给出明确报错，避免程序异常崩溃

【计费规则详情】（基于price_config.xlsx字段设计）
  1. 全国均重模式
     - 触发条件：结算重量≤1kg OR 重量1-3kg但省份非新疆/西藏
     - 计算逻辑：金额留空（需人工按均重规则补充）
     - 适配场景：普通省份小重量订单

  2. 单票模式（重量≥3kg）
     - 触发条件：结算重量≥3kg（所有省份通用）
     - 计算公式：应付金额 = 结算重量 × 对应省份续重单价 + (对应省份超3kg面单费 - 对应快递充单价格)
     - 字段来源：
       → 续重单价：申通/中通报价sheet的E列（iloc[4]）
       → 超3kg面单费：申通报价sheet的D列（iloc[3]）
       → 充单价格：客户快递加收单价信息记录sheet的L列（iloc[11]）

  3. 新西1-3公斤模式（新疆/西藏专属）
     - 触发条件：目的省份为新疆/西藏 AND 结算重量1kg＜重量＜3kg
     - 计算公式：应付金额 = 对应省份3kg内面单费 + 结算重量 × 对应省份续重单价 - 对应快递充单价格

【智能省份匹配说明】
  不再固定截取前2个字，采用开头模糊匹配：
  例：黑龙江省 → 匹配报价表「黑龙江」
      新疆维吾尔自治区 → 匹配报价表「新疆」
      西藏自治区 → 匹配报价表「西藏」
  无需手动改Excel省份名称，自动兼容带「省/自治区」后缀的订单省份

【price_config.xlsx字段约定】
  申通报价                 A列=省份  B列=3kg内面单费  D列=超3kg面单费  E列=续重单价
  中通报价                 同上结构规则
  客户快递加收单价信息记录  K列=快递类型  L列=充单价格

【维护备注】
  1. 智能省份匹配，兼容省、自治区后缀，不用改配置表
  2. 新西固定取B列3.2，单票取D列，续重统一E列
  3. 兼容Python3.14 + 新版pandas，修复fillna语法
  4. 所有报错中文提示，日志清晰，方便排查
==================================================
"""
import pandas as pd
import os
from sys import exit
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# 导入项目统一配置（从settings.py读取路径，避免硬编码）
from config import settings

# ====================== 路径配置（统一从settings读取，避免硬编码）======================
# 输出文件夹路径（自动创建，无需手动新建）
OUTPUT_FOLDER = settings.OUTPUT_FOLDER
# 清洗合并总账单路径（merge_express.py生成）
EXPRESS_FILE = os.path.join(OUTPUT_FOLDER, settings.EXPRESS_OUTPUT_FILE)
# 最终对账结果输出路径（带格式的Excel）
RESULT_FILE = os.path.join(OUTPUT_FOLDER, settings.RESULT_FILE)
# 报价表配置路径
PRICE_CONFIG_PATH = "config/price_config.xlsx"


# ====================== 工具基础函数（通用工具，后续可复用）=====================
def ensure_folder(folder_path):
    """
    确保目标文件夹存在，不存在则自动创建
    :param folder_path: 文件夹路径
    """
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"📂 自动创建输出文件夹：{folder_path}")


# ====================== 【1】计费模式判断函数 ======================
def get_calc_type(dest_prov, weight):
    """
    根据目的省份+结算重量自动判断计费模式
    单票：重量≥3
    新西1-3公斤：新疆/西藏 且 1<重量<3
    其余全部为全国均重
    """
    if weight >= 3:
        return "单票"
    # 智能匹配省份开头，不再截取固定字数
    if dest_prov.startswith("新疆") or dest_prov.startswith("西藏"):
        return "新西1-3公斤"
    return "全国均重"


# ====================== 【2】报价配置加载函数 ======================
def load_price_config():
    """
    读取申通、中通报价 + 充单价格
    拆分存储：
    - 面单费_3kg内 → B列 新西用
    - 面单费_超3kg → D列 单票用
    - 续重单价 → E列 通用
    """
    price_dict = {}

    # ---------------------- 读取申通报价 ----------------------
    try:
        df_st = pd.read_excel(PRICE_CONFIG_PATH, sheet_name="申通报价")
    except Exception as e:
        print(f"❌ 读取申通报价失败：{str(e)}")
        exit(1)

    st_map = {}
    for idx, row in df_st.iterrows():
        if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
            continue
        prov_name = str(row.iloc[0]).strip()
        try:
            st_map[prov_name] = {
                "面单费_3kg内": float(row.iloc[1]),  # B列 3kg内面单
                "面单费_超3kg": float(row.iloc[3]),  # D列 超3kg面单
                "续重单价": float(row.iloc[4])       # E列 续重
            }
        except Exception as e:
            print(f"❌ 申通报价第{idx+2}行数据异常：{str(e)}")
            print(f"👉 请检查该行B/D/E列是否为数字（无文字/空值/特殊符号）")
            exit(1)
    price_dict["申通"] = st_map

    # ---------------------- 读取中通报价 ----------------------
    try:
        df_zt = pd.read_excel(PRICE_CONFIG_PATH, sheet_name="中通报价")
    except Exception as e:
        print(f"❌ 读取中通报价失败：{str(e)}")
        exit(1)

    zt_map = {}
    for idx, row in df_zt.iterrows():
        if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
            continue
        prov_name = str(row.iloc[0]).strip()
        try:
            zt_map[prov_name] = {
                "面单费_3kg内": float(row.iloc[1]),
                "面单费_超3kg": float(row.iloc[3]),
                "续重单价": float(row.iloc[4])
            }
        except Exception as e:
            print(f"❌ 中通报价第{idx+2}行数据异常：{str(e)}")
            exit(1)
    price_dict["中通"] = zt_map

    # ---------------------- 读取充单价格 ----------------------
    try:
        df_charge = pd.read_excel(PRICE_CONFIG_PATH, sheet_name="客户快递加收单价信息记录")
    except Exception as e:
        print(f"❌ 读取充单价格失败：{str(e)}")
        exit(1)

    charge_map = {}
    for idx, row in df_charge.iterrows():
        if pd.isna(row.iloc[10]) or pd.isna(row.iloc[11]):
            continue
        type_name = str(row.iloc[10]).strip()
        if type_name in ["申通", "中通"]:
            try:
                charge_map[type_name] = float(row.iloc[11])
            except Exception as e:
                print(f"❌ 充单价格第{idx+2}行数据异常：{str(e)}")
                exit(1)

    # 强制校验必须有申通中通
    if "申通" not in charge_map or "中通" not in charge_map:
        print("=" * 60)
        print("❌ 致命错误：充单价格配置不全")
        print(f"👉 当前已读取到的充单价格：{charge_map}")
        print("👉 请补充：客户快递加收单价信息记录sheet中，K列填'申通'/'中通'，L列填对应充单价格")
        print("=" * 60)
        exit(1)

    price_dict["充单价"] = charge_map
    print(f"✅ 报价表加载完成：")
    print(f"   - 申通覆盖省份数：{len(st_map)} 个")
    print(f"   - 中通覆盖省份数：{len(zt_map)} 个")
    print(f"   - 充单价格：{charge_map}")
    return price_dict


# ====================== 【3】应付金额计算函数 ======================
def calculate_single_fee(row, price_dict):
    """
    按计费模式，计算单条订单的应付金额
    智能省份匹配：遍历报价省份，用开头匹配
    """
    calc_method = row["实际计算方式"]
    express_type = str(row["快递类型"]).strip()
    dest_prov = str(row["目的省份"]).strip()
    try:
        weight = float(row["结算重量"])
    except Exception as e:
        print(f"⚠️  运单号{row['运单号']}的结算重量异常：{str(e)}，金额置0")
        return round(0.00, 2)

    # 全国均重模式：金额留空
    if calc_method == "全国均重":
        return ""

    # 非申通/中通：金额置0
    if express_type not in ["申通", "中通"]:
        print(f"⚠️  运单号{row['运单号']}：未知快递类型'{express_type}'，金额置0")
        return round(0.00, 2)

    charge_price = price_dict["充单价"][express_type]
    express_prov_map = price_dict[express_type]

    # 智能开头匹配省份
    match_prov = None
    for p_key in express_prov_map:
        if dest_prov.startswith(p_key):
            match_prov = p_key
            break

    if not match_prov:
        print(f"⚠️  运单号{row['运单号']}：{express_type}无'{dest_prov}'省份的报价，金额置0")
        return round(0.00, 2)

    prov_data = express_prov_map[match_prov]

    # 单票模式
    if calc_method == "单票":
        fee = weight * prov_data["续重单价"] + (prov_data["面单费_超3kg"] - charge_price)
    # 新西1-3公斤模式
    elif calc_method == "新西1-3公斤":
        fee = prov_data["面单费_3kg内"] + weight * prov_data["续重单价"] - charge_price
    else:
        print(f"⚠️  运单号{row['运单号']}：未知计费模式'{calc_method}'，金额置0")
        return round(0.00, 2)

    return round(fee, 2)


# ====================== 【4】源数据读取函数 ======================
def load_source_data():
    """
    读取merge_express.py生成的清洗合并账单，和order_db.py生成的订单匹配文件
    """
    if not os.path.exists(EXPRESS_FILE):
        print("=" * 60)
        print("❌ 未找到清洗合并总账单")
        print(f"👉 请先运行 merge_express.py，生成文件：{EXPRESS_FILE}")
        print("👉 该文件需包含列：运单号、目的省份、结算重量、快递类型")
        print("=" * 60)
        exit(1)
    try:
        df_express = pd.read_excel(EXPRESS_FILE, engine="openpyxl")
    except Exception as e:
        print(f"❌ 读取清洗合并账单失败：{str(e)}")
        exit(1)
    # 校验关键列
    required_cols_express = ["运单号", "目的省份", "结算重量", "快递类型"]
    missing_cols = [col for col in required_cols_express if col not in df_express.columns]
    if missing_cols:
        print(f"❌ 清洗合并账单缺少关键列：{missing_cols}")
        exit(1)
    print(f"✅ 成功读取清洗合并总账单：共 {len(df_express)} 条订单记录")

    # 查找订单文件
    order_files = [f for f in os.listdir(OUTPUT_FOLDER)
                   if f.startswith(settings.ORDER_FILE_PREFIX)]
    if not order_files:
        print("=" * 60)
        print("❌ 未找到订单匹配文件")
        print(f"👉 请先运行 order_db.py，生成前缀为'{settings.ORDER_FILE_PREFIX}'的文件")
        print("=" * 60)
        exit(1)
    # 取最新文件
    latest_order_file = sorted(order_files)[-1]
    order_file_path = os.path.join(OUTPUT_FOLDER, latest_order_file)
    try:
        df_order = pd.read_excel(order_file_path, engine="openpyxl")
    except Exception as e:
        print(f"❌ 读取订单匹配文件'{latest_order_file}'失败：{str(e)}")
        exit(1)
    # 校验关键列
    required_cols_order = ["运单号", "所属团队"]
    missing_cols_order = [col for col in required_cols_order if col not in df_order.columns]
    if missing_cols_order:
        print(f"❌ 订单匹配文件缺少关键列：{missing_cols_order}")
        exit(1)
    print(f"✅ 成功读取订单匹配文件：{latest_order_file}，共 {len(df_order)} 条映射记录")

    return df_express, df_order


# ====================== 【5】运单号匹配函数 ======================
def match_team_by_waybill():
    """
    核心业务：加载 → 清洗运单号 → 匹配团队 → 生成计算方式 → 算金额
    """
    df_express, df_order = load_source_data()

    # 运单号清洗
    df_express["运单号"] = df_express["运单号"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df_order["运单号"] = df_order["运单号"].astype(str).str.strip()
    print("✅ 运单号清洗完成：已去除末尾'.0'，统一格式")

    # 构建映射字典
    waybill_team_map = df_order[["运单号", "所属团队"]].drop_duplicates(subset="运单号")\
        .set_index("运单号")["所属团队"].to_dict()
    print(f"✅ 构建运单号-团队映射：共 {len(waybill_team_map)} 个唯一运单号")

    # 回填所属团队 兼容新版pandas
    df_express["所属团队"] = df_express["运单号"].map(waybill_team_map)
    df_express["所属团队"] = df_express["所属团队"].fillna("未匹配")

    # 新增空列
    df_express["实际计算方式"] = ""
    df_express["单票应付金额"] = ""

    # 批量判断计费模式
    print("🔄 正在批量判断计费模式...")
    df_express["实际计算方式"] = df_express.apply(
        lambda row: get_calc_type(row["目的省份"], row["结算重量"]),
        axis=1
    )
    method_distribution = df_express["实际计算方式"].value_counts()
    print(f"✅ 计费模式判断完成，分布如下：")
    for method, count in method_distribution.items():
        print(f"   - {method}：{count} 条（占比：{count/len(df_express)*100:.1f}%）")

    # 批量计算金额
    print("🔄 正在加载报价表并计算应付金额...")
    price_dict = load_price_config()
    df_express["单票应付金额"] = df_express.apply(
        lambda row: calculate_single_fee(row, price_dict),
        axis=1
    )

    # 统计匹配结果
    matched_count = len(df_express[df_express["所属团队"] != "未匹配"])
    unmatched_count = len(df_express) - matched_count
    print(f"✅ 运单号匹配完成：")
    print(f"   - 总订单数：{len(df_express)} 条")
    print(f"   - 已匹配团队：{matched_count} 条（{matched_count/len(df_express)*100:.1f}%）")
    print(f"   - 未匹配团队：{unmatched_count} 条（{unmatched_count/len(df_express)*100:.1f}%）")

    return df_express


# ====================== 【6】结果导出函数 ======================
def export_styled_result(df_result):
    """
    导出带边框、表头蓝色背景、居中对齐、固定列宽Excel
    """
    ensure_folder(OUTPUT_FOLDER)

    # 导出基础Excel
    try:
        df_result.to_excel(RESULT_FILE, index=False, engine="openpyxl")
    except Exception as e:
        print(f"❌ 导出Excel失败：{str(e)}")
        exit(1)

    # 加载样式
    try:
        wb = load_workbook(RESULT_FILE)
        ws = wb.active
        max_row = ws.max_row
        max_col = ws.max_column
    except Exception as e:
        print(f"❌ 加载Excel样式失败：{str(e)}")
        exit(1)

    # 样式定义
    thin_border = Side(style="thin", color="000000")
    border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")

    # 全表边框+居中
    last_col_letter = chr(ord('A') + max_col - 1)
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            cell.border = border
            cell.alignment = center_align

    # 表头样式
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    # 固定列宽
    column_widths = {
        'A': 20,   # 运单号
        'B': 18,   # 快递类型
        'C': 12,   # 目的省份
        'D': 12,   # 结算重量
        'E': 10,   # 其他字段
        'F': 16,   # 所属团队
        'G': 16,   # 实际计算方式
        'H': 16,   # 单票应付金额
        'I': 28    # 备注
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # 保存
    try:
        wb.save(RESULT_FILE)
    except Exception as e:
        print(f"❌ 保存样式Excel失败：{str(e)}")
        exit(1)

    print(f"✅ 最终对账结果已导出：{RESULT_FILE}")
    print(f"   - 包含字段：{list(df_result.columns)}")
    print(f"   - 样式：全表边框+表头蓝色背景+居中对齐+固定列宽")


# ====================== 【7】主函数 ======================
def run_reconciliation():
    """
    程序主入口：打印日志 → 执行流程 → 导出结果
    """
    print("=" * 80)
    print("📦 快递订单对账匹配程序（V1.7 最终功能版）")
    print(f"⏰ 启动时间：{pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📂 输出路径：{OUTPUT_FOLDER}")
    print("=" * 80)

    try:
        matched_data = match_team_by_waybill()
    except Exception as e:
        print(f"\n❌ 程序执行失败：{str(e)}")
        exit(1)

    if matched_data is not None and len(matched_data) > 0:
        export_styled_result(matched_data)
        print("\n🎉 全部流程执行完毕！可打开以下文件进行对账：")
        print(f"   → {RESULT_FILE}")
    else:
        print("\n⚠️  无有效对账数据可导出（可能源文件为空）")

    print("\n" + "=" * 80)


# ====================== 程序启动 ======================
if __name__ == "__main__":
    run_reconciliation()
# -*- coding: utf-8 -*-
"""
快递账单对账总控程序 V2.1
==================================================
【执行流程】
  运行即启动 → 自动建目录 → 弹窗引导放文件
  → ① 下载数据库订单  → output/YYYY-MM/毅播快递数据_YYYYMM.xlsx
  → ② 清洗合并快递账单 → output/YYYY-MM/清洗合并总账单.xlsx
  → ③ 运单号匹配对账   → output/YYYY-MM/最终对账结果.xlsx
  → ④ 按团队拆分账单   → output/YYYY-MM/客户账单/各团队Excel
  → ⑤ 生成README.md   → 控制台显示全文，可一键复制提交Git

【V2.1 新增】
  1. 自动识别上月月份，无需手动改任何配置
  2. 启动直接显示月份信息，运行即执行，无需确认
  3. 自动创建 data/YYYY-MM/ 和 output/YYYY-MM/ 目录
  4. data 目录为空时自动弹出文件夹窗口，放好文件按回车继续
  5. 保留V2.0：完成后自动生成README.md + 控制台一键复制Git

【可独立运行】
==================================================
"""
import sys
import os

# 将当前目录加入 Python 路径，确保各模块可正常导入
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def print_separator(title):
    """打印流程分隔符，便于控制台区分各步骤"""
    print("\n" + "=" * 70)
    print(f"📌 {title}")
    print("=" * 70)



def ensure_dirs(data_folder, output_folder, process_month):
    """
    自动创建月份子目录 + 引导用户放入账单文件
    业务逻辑：
      1. 自动创建 data/YYYY-MM/ 和 output/YYYY-MM/
      2. 检测 data/YYYY-MM/ 是否已有 xlsx 文件
         - 有文件 → 直接通过，不打扰用户
         - 没文件 → 自动弹出文件夹窗口，等用户放好后按回车继续
      3. 用户按回车后二次校验，防止忘记放文件就继续
    返回 True = 目录就绪且有文件，False = 文件仍为空则退出
    """
    os.makedirs(data_folder,  exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    def has_excel_files():
        """检测 data 目录是否有 xlsx 文件（排除 Excel 临时文件）"""
        return any(
            f.endswith(".xlsx") and not f.startswith("~")
            for f in os.listdir(data_folder)
        )

    # 已有文件，直接通过，不打扰
    if has_excel_files():
        return True

    # 没有文件：自动弹出文件夹窗口，引导用户放入账单
    abs_path = os.path.abspath(data_folder)
    print(f"\n📂 data/{process_month}/ 目录为空，已自动打开文件夹：")
    print(f"   {abs_path}")
    print(f"   请将申通、中通原始账单 Excel 放入该文件夹，放好后按回车继续...")

    # 自动弹出 Windows 文件夹窗口
    os.startfile(abs_path)

    # 等待用户放好文件后按回车
    input()

    # 二次校验：防止用户忘记放文件直接回车
    if not has_excel_files():
        print(f"\n❌ data/{process_month}/ 仍为空，请放入账单文件后重新运行。")
        return False

    print(f"✅ 检测到账单文件，继续运行...")
    return True


def create_readme_file(process_month):
    """
    生成 README.md 并返回内容
    业务逻辑：每次运行完自动更新项目说明文档
             控制台显示全文，直接框选 Ctrl+C 即可复制提交 Git
    """
    readme_text = f"""# 快递自动对账·账单系统 V2.1

## 程序说明
全自动完成：订单下载 → 快递账单合并 → 运单匹配对账 → 按团队拆分账单
V2.1 新增按月归档，各月数据完全隔离，自动识别处理月份无需手动修改配置

## 版本更新 V2.1
1. 按月归档：data/YYYY-MM/ 和 output/YYYY-MM/ 各月独立，互不覆盖
2. 自动识别上月：5月运行处理4月数据，6月运行处理5月数据
3. 启动月份确认：防止误操作处理错月份
4. 防覆盖保护：当月已有结果时提示用户确认，默认拒绝覆盖
5. data目录空文件检测：自动提示用户放入账单文件

## 执行流程
1. 从MySQL自动下载上月订单数据
2. 清洗合并 data/{process_month}/ 申通、中通原始账单
3. 运单号自动匹配回填所属团队、智能判断计费模式
4. 按团队自动拆分生成独立标准化Excel账单

## 目录结构
- config/          配置文件、price_config.xlsx价格配置表
- data/YYYY-MM/    存放当月申通、中通原始账单
- output/YYYY-MM/  所有输出文件 + 拆分后客户账单

## 每月操作步骤
1. 直接运行：python main.py
2. 程序自动弹出 data/YYYY-MM/ 文件夹，将申通、中通原始账单放入即可
3. 文件放好后按回车，程序自动识别并继续运行
4. 前往 output/YYYY-MM/ 查看结果

## 特性
- 全流程全自动无人工干预
- 申通/中通差异化计费规则
- 自动创建目录、Excel自动样式美化
- 全局异常捕获，报错清晰可排查
- 按月归档，历史数据永久保留
"""
    # 写入本地 README.md
    with open("README.md", "w", encoding="utf-8") as f:
        f.write(readme_text)
    return readme_text


def main():
    # -------------------- 读取本次处理月份 --------------------
    # 从 settings 读取自动计算好的上月月份，无需手动修改
    from config import settings
    process_month = settings.PROCESS_MONTH
    data_folder   = settings.DATA_FOLDER
    output_folder = settings.OUTPUT_FOLDER

    print(f"\n🚀 快递账单对账全流程程序（V2.1）启动")
    print("\n" + "=" * 70)
    print(f"  🗓  处理月份：{process_month}")
    print(f"  📂  数据目录：data/{process_month}/")
    print(f"  📂  输出目录：output/{process_month}/")
    print("=" * 70)

    # -------------------- Step 0：自动建目录 + 引导放文件 --------------------
    print(f"\n📁 正在初始化 {process_month} 月份目录...")
    if not ensure_dirs(data_folder, output_folder, process_month):
        sys.exit(1)
    print(f"   ✅ data/{process_month}/   就绪")
    print(f"   ✅ output/{process_month}/ 就绪")

    # -------------------- 主流程：四步串联 --------------------
    try:
        # 延迟导入模块，避免启动时因配置未初始化报错
        from modules.express import order_db
        from modules.express import merge_express
        from modules.express import order_matching
        from modules.express import split_bill_by_team

        # ---------- 第一步：从数据库下载订单数据 ----------
        print_separator(f"第一步：从数据库下载 {process_month} 订单数据")
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            print("❌ 订单数据下载失败或为空，终止流程")
            sys.exit(1)

        # ---------- 第二步：清洗合并申通/中通快递账单 ----------
        print_separator(f"第二步：清洗合并 {process_month} 快递账单")
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            print("❌ 快递账单合并失败或为空，终止流程")
            sys.exit(1)

        # ---------- 第三步：运单号匹配对账+计费 ----------
        print_separator(f"第三步：运单号匹配对账（{process_month}）")
        order_matching.run_reconciliation()

        # ---------- 第四步：按团队拆分客户账单 ----------
        print_separator(f"第四步：按团队拆分客户账单（{process_month}）")
        split_bill_by_team.main()

        # ---------- 第五步：生成 README.md ----------
        # 业务逻辑：每次跑完自动更新文档，控制台显示全文便于复制提交 Git
        print_separator("第五步：生成项目说明 README.md")
        readme_content = create_readme_file(process_month)
        print("✅ 已生成本地 README.md 文件")
        print("📋 下方为README全文，直接框选 Ctrl+C 即可复制提交Git：")
        print("=" * 80)
        print(readme_content)
        print("=" * 80)

        # -------------------- 完成汇总 --------------------
        print_separator(f"✅ {process_month} 全部流程执行完成 🎉")
        print("📦 生成文件清单：")
        print(f"   - 订单数据：output/{process_month}/毅播快递数据_YYYYMM.xlsx")
        print(f"   - 合并账单：output/{process_month}/清洗合并总账单.xlsx")
        print(f"   - 对账结果：output/{process_month}/最终对账结果.xlsx")
        print(f"   - 拆分账单：output/{process_month}/客户账单/ 各团队独立Excel")
        print(f"\n🎉 全部任务执行完毕，请前往 output/{process_month}/ 查看结果")

    except Exception as e:
        # 全局异常捕获：打印详细报错，不再闪退
        print(f"\n❌ 程序运行异常：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
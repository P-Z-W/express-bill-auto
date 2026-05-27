# -*- coding: utf-8 -*-
"""
快递账单对账总控程序 V2.0
==================================================
【执行流程】
1. 下载数据库订单数据 → output/毅播快递数据_YYYYMM.xlsx
2. 清洗合并申通/中通账单 → output/清洗合并总账单.xlsx
3. 运单号匹配对账 → output/最终对账结果.xlsx
4. 按团队拆分账单 → output/客户账单 分团队自动生成Excel
【版本更新 V2.0】
1. 千耀传媒：双条件独立计算、快递类型列仅全国均重显示
2. 加收费汇总标题横跨L→S整列，表格排版优化
3. 专项统计表完整闭合边框、下移不重叠
4. 普通团队0金额自动过滤，千耀金额0强制输出
5. 程序全局异常捕获，解决直接闪退问题
6. 运行完成自动生成README.md，控制台可一键复制提交Git
【可独立运行】
==================================================
"""
import sys
import os

# 将当前目录加入Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def print_separator(title):
    """打印流程分隔符"""
    print("\n" + "="*70)
    print(f"📌 {title}")
    print("="*70)

def create_readme_file():
    """生成README.md 并返回内容，支持控制台一键复制"""
    readme_text = """# 快递自动对账&账单系统 V2.0

## 程序说明
全自动完成：订单下载 → 快递账单合并 → 运单匹配对账 → 按团队拆分账单
V2.0 优化千耀传媒专属计算逻辑与Excel表格排版样式

## 版本更新 V2.0
1. 千耀新增快递类型列，全国均重拆分申通/中通独立核算
2. 按【所属团队+快递类型】双条件匹配配置拉均、加收单价
3. 超出重量自动负数置0，费用计算精准无误
4. 加收费汇总标题横跨L→S整列，排版规整美观
5. 专项统计表下移+完整闭合边框，无缺线无重叠
6. 普通团队自动过滤0金额，千耀传媒金额0强制保留输出

## 执行流程
1. 从MySQL自动下载上月订单数据
2. 清洗合并data目录申通、中通原始账单
3. 运单号自动匹配回填所属团队、智能判断计费模式
4. 按团队自动拆分生成独立标准化Excel账单

## 目录结构
- config/        配置文件、price_config.xlsx价格配置表
- data/          存放申通、中通原始账单
- output/        所有输出文件 + 拆分后客户账单

## 运行方式
直接运行：
python main.py

## 特性
- 全流程全自动无人工干预
- 申通/中通差异化计费规则
- 自动创建目录、Excel自动样式美化
- 全局异常捕获，报错清晰可排查
"""
    # 写入本地README.md
    with open("README.md", "w", encoding="utf-8") as f:
        f.write(readme_text)
    return readme_text

def main():
    try:
        # 延迟导入模块，避免启动闪退 + 异常可捕获
        import order_db
        import merge_express
        import order_matching
        import split_bill_by_team

        # 第一步：下载数据库订单数据
        print_separator("第一步：从数据库下载订单数据")
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            print("❌ 订单数据下载失败或为空，终止流程")
            return

        # 第二步：清洗合并申通/中通快递账单
        print_separator("第二步：清洗合并申通/中通快递账单")
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            print("❌ 快递账单合并失败或为空，终止流程")
            return

        # 第三步：运单号匹配对账
        print_separator("第三步：运单号匹配并生成最终对账结果")
        order_matching.run_reconciliation()

        # 第四步：按团队拆分账单
        print_separator("第四步：按团队拆分客户账单")
        split_bill_by_team.main()

        # 生成README.md + 控制台一键复制
        print_separator("生成项目说明 README.md")
        readme_content = create_readme_file()
        print("✅ 已生成本地 README.md 文件")
        print("📋 下方为README全文，直接框选 Ctrl+C 即可复制提交Git：")
        print("="*80)
        print(readme_content)
        print("="*80)

        # 流程完成汇总
        print_separator("所有流程执行完成 ✨")
        print("📁 生成文件清单：")
        print("   - 订单数据：output/毅播快递数据_对应月份.xlsx")
        print("   - 合并账单：output/清洗合并总账单.xlsx")
        print("   - 对账结果：output/最终对账结果.xlsx")
        print("   - 拆分账单：output/客户账单/ 各团队独立Excel")
        print("\n🎉 全部任务执行完毕！")

    except Exception as e:
        # 全局异常捕获，打印详细错误，不再闪退
        print(f"\n❌ 程序运行异常：{str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    print("🚀 启动快递账单对账全流程程序（V2.0）")
    main()
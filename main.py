# -*- coding: utf-8 -*-
"""
快递账单对账总控程序（V1.4）
==================================================
【执行流程】
1. 下载数据库订单数据 → output/毅播快递数据_YYYYMM.xlsx
2. 清洗合并申通/中通账单 → output/清洗合并总账单.xlsx
3. 运单号匹配对账 → output/最终对账结果.xlsx
【可独立运行】
==================================================
"""
import sys
import os

# 将当前目录加入Python路径（确保能导入同级模块）
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入各功能模块
import merge_express
import order_db
import order_matching

def print_separator(title):
    """打印流程分隔符"""
    print("\n" + "="*70)
    print(f"📌 {title}")
    print("="*70)

def main():
    """主控制流程"""
    try:
        # 第一步：下载数据库订单数据
        print_separator("第一步：从数据库下载订单数据")
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            print("❌ 订单数据下载失败或为空，终止流程")
            return

        # 第二步：清洗合并快递账单
        print_separator("第二步：清洗并合并申通/中通快递账单")
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            print("❌ 快递账单合并失败或为空，终止流程")
            return

        # 第三步：运单号匹配对账
        print_separator("第三步：运单号匹配并生成最终对账结果")
        order_matching.run_reconciliation()

        # 流程完成
        print_separator("所有流程执行完成 ✨")
        print("📁 生成文件清单：")
        print(f"   - 订单数据：{order_db.OUTPUT_FOLDER}/{order_db.ORDER_FILE_PREFIX}{order_db.get_last_month_str()}.xlsx")
        print(f"   - 合并账单：{merge_express.OUTPUT_FILE}")
        print(f"   - 对账结果：{order_matching.RESULT_FILE}")
        print("\n🎉 全部任务执行完毕！")

    except Exception as e:
        print(f"\n❌ 程序执行出错：{str(e)}")
        # 异常时终止程序并返回错误码
        sys.exit(1)

if __name__ == "__main__":
    print("🚀 启动快递账单对账全流程程序（V1.4）")
    main()
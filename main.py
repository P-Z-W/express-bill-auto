# -*- coding: utf-8 -*-
"""
项目总入口：调用所有模块
1. 合并快递单
2. 下载订单（后续）
3. 订单匹配（后续）
4. 按团队结算（后续）
"""
from merge_express import run_merge_process

if __name__ == "__main__":
    print("🚀 项目启动：快递对账自动化系统")
    print("=" * 60)

    # ========== 1. 执行快递单合并 ==========
    df_express_total = run_merge_process()

    # ========== 后续：下载订单、匹配、对账、算钱 在这里调用 ==========
    # from order_db import download_orders
    # from order_matching import run_reconciliation

    print("\n✅ 全部流程执行完毕！")
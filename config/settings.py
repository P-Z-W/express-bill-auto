# -*- coding: utf-8 -*-
"""
项目统一配置文件（V1.4 标准版）
==================================================
【框架说明】
  本文件仅存放项目所有固定配置，不包含任何业务逻辑。
  所有路径、数据库、文件名统一在此管理，后续只需修改此处。
【可配置项】
  1. 文件夹路径
  2. 快递账单输入/输出名称
  3. 数据库连接信息
  4. SQL配置文件路径
  5. 订单文件前缀
==================================================
"""
import os

# ==================== 文件夹路径 ====================
DATA_FOLDER    = "data"
OUTPUT_FOLDER  = "output"
CONFIG_FOLDER  = "config"

# ==================== 快递合并模块 ====================
EXPRESS_INPUT_ST      = "2026年4月钟村毅播云仓对帐单.xlsx"
EXPRESS_INPUT_ZT      = "咸辣甜服饰-4月份.xlsx"
EXPRESS_OUTPUT_FILE   = "清洗合并总账单.xlsx"

# ==================== 数据库订单模块 ====================
DB_CONFIG = {
    "host": "yiboerp.rwlb.rds.aliyuncs.com",
    "port": 3306,
    "user": "only_read",
    "password": "1oEb4y(lBb09F",
    "database": "yibo",
    "charset": "utf8mb4"
}
SQL_FILE_PATH         = os.path.join(CONFIG_FOLDER, "SQL-config.txt")
ORDER_FILE_PREFIX     = "毅播快递数据_"

# ==================== 对账匹配模块 ====================
RESULT_FILE           = "最终对账结果.xlsx"
# -*- coding: utf-8 -*-
"""
项目统一配置文件 V2.1
==================================================
【架构说明】
  本文件仅存放项目所有固定配置，不包含任何业务逻辑。
  所有路径、数据库、文件名统一在此管理，后续只需改这里。

【V2.1 新增】
  - 自动计算"上个月"月份，无需手动修改任何配置
  - data/ 和 output/ 改为按月份子目录存放
  - 5月运行 → 自动处理 2026-04 数据
  - 6月运行 → 自动处理 2026-05 数据，与4月完全隔离
  - 数据库密码从 .env 文件读取，不提交到Git

【.env 文件说明】
  项目根目录新建 .env 文件，内容如下：
    DB_HOST=你的数据库地址
    DB_PORT=3306
    DB_USER=账号
    DB_PASSWORD=密码
    DB_NAME=数据库名
    DB_CHARSET=utf8mb4
  .env 已加入 .gitignore，永远不会提交到Git

【可配置项】
  1. 文件夹路径（含月份子目录）
  2. 快递账单输入/输出文件名
  3. 数据库连接信息（从.env读取）
  4. SQL配置文件路径
  5. 订单文件前缀
  6. SQL动态日期扩展天数
==================================================
"""
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

# ==================== 加载 .env 文件 ====================
# 从项目根目录读取 .env，把变量载入环境
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), ".env"))

# ==================== 自动计算上个月 ====================
# 业务逻辑：每月运行时处理的是"上个月"的数据
# 例：5月运行 → PROCESS_MONTH = "2026-04"
#     6月运行 → PROCESS_MONTH = "2026-05"
_today          = datetime.now()
_first_of_month = _today.replace(day=1)               # 本月1号
_last_month_day = _first_of_month - timedelta(days=1) # 上月最后一天
PROCESS_MONTH   = _last_month_day.strftime("%Y-%m")   # → "2026-05"

# ==================== 文件夹路径（按月归档）====================
# 每个月的数据完全隔离，互不覆盖
CONFIG_FOLDER  = "config"
DATA_FOLDER    = os.path.join("data",   PROCESS_MONTH)  # data/2026-05
OUTPUT_FOLDER  = os.path.join("output", PROCESS_MONTH)  # output/2026-05

# ==================== 快递合并模块 ====================
EXPRESS_INPUT_ST    = "申通账单.xlsx"       # 申通原始账单文件名
EXPRESS_INPUT_ZT    = "中通账单.xlsx"       # 中通原始账单文件名
EXPRESS_OUTPUT_FILE = "清洗合并总账单.xlsx" # 合并后输出文件名

# ==================== 数据库订单模块 ====================
# 从 .env 文件读取，密码不出现在代码里
DB_CONFIG = {
    "host":     os.getenv("DB_HOST",     ""),
    "port":     int(os.getenv("DB_PORT", "3306")),
    "user":     os.getenv("DB_USER",     ""),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME",     ""),
    "charset":  os.getenv("DB_CHARSET",  "utf8mb4"),
}
SQL_FILE_PATH     = os.path.join(CONFIG_FOLDER, "SQL-config.txt")
ORDER_FILE_PREFIX = "毅播快递数据_"  # 订单导出文件名前缀，后接 YYYYMM

# ==================== SQL 动态日期范围 ====================
# 业务逻辑：账单数据包含上月底和下月初的运单，需要多捞前后几天
# 可配置：调整天数覆盖业务边界，不用改任何其他代码
SQL_EXTEND_DAYS_BEFORE = 15  # 往前多捞天数：处理月1号往前推15天
SQL_EXTEND_DAYS_AFTER  = 5   # 往后多捞天数：处理月最后一天往后推5天

# 自动计算 START_DATE 和 END_DATE
# 例：处理 2026-05
#   → START_DATE = 2026-04-16 00:00:00（5月1号往前15天）
#   → END_DATE   = 2026-06-05 23:59:59（5月31号往后5天）
_process_first_day = _last_month_day.replace(day=1)      # 处理月1号
_process_last_day  = _first_of_month - timedelta(days=1) # 处理月最后一天

SQL_START_DATE = (_process_first_day - timedelta(days=SQL_EXTEND_DAYS_BEFORE))\
    .strftime("%Y-%m-%d 00:00:00")
SQL_END_DATE   = (_process_last_day  + timedelta(days=SQL_EXTEND_DAYS_AFTER))\
    .strftime("%Y-%m-%d 23:59:59")

# ==================== 对账匹配模块 ====================
RESULT_FILE = "最终对账结果.xlsx"  # 运单匹配+计费后的输出文件名
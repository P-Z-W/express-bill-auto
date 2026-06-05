---
name: project-overview
description: 毅播快递对账系统的整体架构和业务流程说明
metadata:
  type: project
---

# 毅播快递对账系统（express-bill-auto）

**Why:** 该系统为"毅播"公司自动化完成每月快递费用对账，取代手工核对申通/中通账单与订单数据的繁琐工作。

**How to apply:** 理解任何功能需求时，都要结合4步核心流程和Web管理界面的定位来判断影响范围。

## 技术栈
- 后端：Python + Flask（V3.1）
- 数据处理：pandas + openpyxl
- 数据库：阿里云 MySQL（pymysql）
- 前端：多页面HTML模板（templates/）

## 4步核心流程

1. **order_db.py** — 连接 MySQL，用 config/SQL-config.txt 模板（含{START_DATE}/{END_DATE}占位符）查询上月订单，导出到 `output/YYYY-MM/毅播快递数据_YYYYMM.xlsx`

2. **merge_express.py** — 扫描 `data/YYYY-MM/` 目录，通过特征列自动识别快递类型（申通="业务时间"列，中通="扫描时间"列），清洗合并为统一格式，输出 `清洗合并总账单.xlsx`

3. **order_matching.py** — 将快递账单运单号与订单数据匹配回填所属团队，按【结算重量+目的省份】判断计费模式（全国均重/单票/新西1-3公斤），从 price_config.xlsx 读取报价计算应付金额，输出 `最终对账结果.xlsx`

4. **split_bill_by_team.py** — 按团队拆分，含汇总统计区域和新西1%特殊规则，文件命名为 `{金额}-{团队名}_{月份}_快递加收费.xlsx`

## 目录结构
- `config/price_config.xlsx` — 申通/中通报价 + 客户加收单价（3个sheet）
- `config/express_config.json` — 快递公司配置（identify_column、enabled）
- `config/SQL-config.txt` — SQL模板
- `data/YYYY-MM/` — 原始快递账单（每月）
- `output/YYYY-MM/` — 输出结果（每月归档）

## Web界面（app.py）
- `/` 首页看板：团队费用、申通/中通占比、月度趋势、环比
- `/run` 运行页：上传账单、触发运行、SSE实时日志
- `/history` 历史记录
- `/stats` 统计报表 + 异常运单分析
- `/config` 系统配置（报价Web化管理）

## 特殊业务规则
- "千耀传媒"团队有专项统计逻辑（按申通/中通分开计算）
- 新疆/西藏1-3kg走"新西1-3公斤"计费模式
- 新西1kg内订单数超总量1%才收费（r11_value = 超出单数×10）
- SQL日期范围：处理月前扩15天、后扩5天（可配置）

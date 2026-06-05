---
name: project-overview
description: yibo-backoffice 快递对账系统整体架构、业务流程、V4.1 UI 特性说明
metadata:
  type: project
---

# 毅播云仓内部管理平台 (yibo-backoffice)

**Why:** 自动化完成每月快递费用对账，取代手工核对申通/中通账单与订单数据的繁琐工作，支持局域网 Web 操作。

**How to apply:** 理解任何功能需求时结合4步核心流程、多模块架构和 Web 管理界面的定位来判断影响范围。

## 当前版本：V4.1（2026-06-05）

## 技术栈
- 后端：Python + Flask，Blueprint 多模块架构
- 数据处理：pandas + openpyxl
- 数据库：阿里云 MySQL（pymysql）
- 前端：Jinja2 模板 + 原生 JS + Chart.js，SSE 实时日志推送

## 多模块架构（V4.0+）

```
core/
  cache.py        — Excel/接口结果三层缓存
  task_runner.py  — 任务锁、LogCapture 线程隔离、SSE 队列
  utils.py        — openpyxl 样式常量、公共工具

modules/
  express/        — 快递对账（当前唯一完整模块）
    routes.py     — express_bp，全部路由（无 url_prefix，保持原 URL 不变）
    order_db.py   — 数据库订单下载
    merge_express.py     — 账单合并
    order_matching.py    — 运单对账匹配
    split_bill_by_team.py — 按团队拆分账单
  storage/routes.py  — 仓储费（占位，/storage 前缀）
  finance/routes.py  — 财务汇总（占位，/finance 前缀）
```

## 4步核心流程

1. **order_db.py** — 连接 MySQL，用 `config/SQL-config.txt` 模板查询上月订单，导出到 `output/YYYY-MM/毅播快递数据_YYYYMM.xlsx`

2. **merge_express.py** — 扫描 `data/YYYY-MM/`，通过特征列自动识别快递类型（申通="业务时间"，中通="扫描时间"），清洗合并为统一格式

3. **order_matching.py** — 运单号匹配回填团队，按【结算重量+目的省份】判断计费模式，从 `price_config.xlsx` 读取报价计算应付金额

4. **split_bill_by_team.py** — 按团队拆分，含汇总统计和新西1%特殊规则，文件命名 `{金额}-{团队名}_{月份}_快递加收费.xlsx`

## Web 路由（express_bp，无前缀）

| 路径 | 说明 |
|---|---|
| `/` | 看板：费用趋势、团队 TOP5、快递占比、环比 |
| `/run` | 运行：上传账单、触发对账、SSE 实时进度 |
| `/history` | 历史：运行次数、结果、耗时、下载 |
| `/stats` | 统计：团队汇总、异常运单分析、对账预览 |
| `/config` | 配置：报价管理、快递管理、运行参数 |

SSE 信号：`__STEP__0–4`（进度推进）、`__STEP_FAIL__N`（步骤失败）、`__DONE__`、`__PING__`

## V4.1 UI 特性
- 220px 固定左侧边栏，SVG 矢量导航图标
- 运行页步骤时间线（CSS `::before` 圆形指示器，active/done/error 状态）
- 骨架屏加载动画（`.sk` shimmer，历史/统计/配置页）
- Stat 卡片顶部 2px accent 色条区分类型

## 特殊业务规则
- 新疆/西藏 1–3kg → 新西1-3公斤计费
- 新西 1kg 内订单超总量 1% 才收费（r11_value = 超出单数×10）
- SQL 日期：处理月前扩15天、后扩5天（`settings_override.json` 可覆盖）

## 配置文件
- `config/settings.py` — 全局配置，月份自动取上月
- `config/price_config.xlsx` — 申通/中通报价 + 客户加收单价（3个 sheet）
- `config/express_config.json` — 快递公司开关配置
- `config/SQL-config.txt` — SQL 模板（`.gitignore` 排除，含数据库密码）

## Git
远程：`https://github.com/P-Z-W/yibo-backoffice.git`，主分支 `main`，最新 tag `v4.1`

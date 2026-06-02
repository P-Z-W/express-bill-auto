# -*- coding: utf-8 -*-
"""
毅播快递对账系统 - Web服务主程序 V2.6
==================================================
【说明】
  Flask轻量Web服务，提供局域网浏览器操作界面
  五个业务模块完全不改动，app.py只负责：
    1. 接收上传的申通/中通账单，存入 data/YYYY-MM/
    2. 按顺序调用四个业务模块，实时推送日志到浏览器
    3. 运行完成后打包 output/YYYY-MM/ 提供下载

【V2.5 新增】
  运行日志自动保存到 output/YYYY-MM/run.log
  追加模式：多次运行不覆盖，每次运行用分隔线区分

【V2.6 新增】
  历史记录查看：页面底部显示所有月份处理记录
  每个月份独立下载按钮，直接下载对应月份zip
==================================================
"""
import os
import sys
import io
import queue
import threading
import zipfile
import json
from datetime import datetime, timedelta
from flask import Flask, render_template, request, Response, jsonify, send_file

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import settings

app = Flask(__name__)

# ====================== 全局状态 ======================
task_lock   = threading.Lock()
is_running  = False
log_queue   = queue.Queue()
task_result = {"success": False, "output_folder": ""}


# ====================== 工具函数 ======================
def get_process_month():
    """获取本次处理月份（上个月）"""
    today          = datetime.now()
    first_of_month = today.replace(day=1)
    last_month     = first_of_month - timedelta(days=1)
    return last_month.strftime("%Y-%m")


def push_log(msg):
    """推送日志到浏览器队列，不调用print避免死循环"""
    log_queue.put(msg)


# ====================== 日志文件工具 ======================
def get_log_path(output_folder):
    """获取日志文件路径：output/YYYY-MM/run.log"""
    return os.path.join(output_folder, "run.log")


def get_run_count(log_path):
    """统计 run.log 里已有几次运行记录"""
    if not os.path.exists(log_path):
        return 1
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()
        return content.count("【第") + 1
    except Exception:
        return 1


def write_log_header(log_file, process_month, run_count):
    """写入每次运行的分隔头"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file.write("\n" + "=" * 60 + "\n")
    log_file.write(f"【第{run_count}次运行】{now}\n")
    log_file.write(f"处理月份：{process_month}\n")
    log_file.write("=" * 60 + "\n")
    log_file.flush()


def write_log_footer(log_file, success, start_time):
    """写入运行结果和耗时"""
    end_time = datetime.now()
    elapsed  = end_time - start_time
    minutes  = int(elapsed.total_seconds() // 60)
    seconds  = int(elapsed.total_seconds() % 60)
    result   = "成功 ✅" if success else "失败 ❌"
    log_file.write("\n" + "-" * 60 + "\n")
    log_file.write(f"运行结果：{result}\n")
    log_file.write(f"结束时间：{end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
    log_file.write(f"耗时：{minutes}分{seconds}秒\n")
    log_file.write("-" * 60 + "\n")
    log_file.flush()


# ====================== LogCapture ======================
class LogCapture(io.StringIO):
    """
    劫持标准输出，捕获业务模块所有 print() 推入队列
    同时追加写入 run.log 日志文件
    """
    def __init__(self, real_stdout, log_file=None):
        super().__init__()
        self._real_stdout = real_stdout
        self._log_file    = log_file

    def write(self, msg):
        if msg.strip():
            log_queue.put(msg.strip())
            self._real_stdout.write(msg + "\n")
            if self._log_file:
                self._log_file.write(msg.strip() + "\n")
                self._log_file.flush()
        return len(msg)

    def flush(self):
        pass


# ====================== 业务任务线程 ======================
def run_task(data_folder, output_folder, process_month):
    """
    独立线程里按顺序执行四个业务模块
    所有日志实时推送到浏览器，同时追加写入 run.log
    """
    global is_running, task_result

    start_time = datetime.now()
    success    = False

    os.makedirs(output_folder, exist_ok=True)

    log_path  = get_log_path(output_folder)
    run_count = get_run_count(log_path)
    log_file  = open(log_path, "a", encoding="utf-8")
    write_log_header(log_file, process_month, run_count)

    real_stdout = sys.stdout
    sys.stdout  = LogCapture(real_stdout, log_file)

    try:
        push_log(f"🚀 开始处理 {process_month} 月份数据")
        push_log(f"📂 数据目录：{data_folder}")
        push_log(f"📂 输出目录：{output_folder}")
        push_log("=" * 60)
        push_log("__STEP__0")  # 通知前端：开始运行

        # ---------- 第一步：数据库下载订单 ----------
        push_log(f"\n📌 第一步：从数据库下载 {process_month} 订单数据")
        import order_db
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            push_log("❌ 订单数据下载失败或为空，终止流程")
            push_log("__STEP_FAIL__1")
            task_result = {"success": False, "output_folder": output_folder}
            return
        push_log("__STEP__1")  # 通知前端：第一步完成 25%

        # ---------- 第二步：合并快递账单 ----------
        push_log(f"\n📌 第二步：清洗合并 {process_month} 快递账单")
        import merge_express
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            push_log("❌ 快递账单合并失败或为空，终止流程")
            push_log("__STEP_FAIL__2")
            task_result = {"success": False, "output_folder": output_folder}
            return
        push_log("__STEP__2")  # 通知前端：第二步完成 50%

        # ---------- 第三步：运单匹配对账 ----------
        push_log(f"\n📌 第三步：运单号匹配对账（{process_month}）")
        import order_matching
        order_matching.run_reconciliation()
        push_log("__STEP__3")  # 通知前端：第三步完成 75%

        # ---------- 第四步：按团队拆分账单 ----------
        push_log(f"\n📌 第四步：按团队拆分客户账单（{process_month}）")
        import split_bill_by_team
        split_bill_by_team.main()
        push_log("__STEP__4")  # 通知前端：第四步完成 100%

        push_log("\n" + "=" * 60)
        push_log(f"🎉 {process_month} 全部流程执行完成！")
        push_log("📦 可点击下方按钮下载结果文件")

        success     = True
        task_result = {"success": True, "output_folder": output_folder}

    except Exception as e:
        import traceback
        push_log(f"\n❌ 程序运行异常：{str(e)}")
        push_log(traceback.format_exc())
        task_result = {"success": False, "output_folder": output_folder}

    finally:
        write_log_footer(log_file, success, start_time)
        log_file.close()
        sys.stdout = real_stdout
        is_running = False
        push_log("__DONE__")


# ====================== 路由：首页 ======================
@app.route("/")
def index():
    process_month = get_process_month()
    return render_template("index.html", process_month=process_month)


# ====================== 路由：上传账单文件 ======================
@app.route("/upload", methods=["POST"])
def upload():
    """接收上传文件存入 data/YYYY-MM/"""
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "没有收到文件"}), 400

    file          = request.files["file"]
    process_month = get_process_month()
    data_folder   = os.path.join("data", process_month)
    os.makedirs(data_folder, exist_ok=True)

    save_path = os.path.join(data_folder, file.filename)
    file.save(save_path)

    return jsonify({
        "ok":  True,
        "msg": f"✅ {file.filename} 已上传到 data/{process_month}/"
    })


# ====================== 路由：开始运行 ======================
@app.route("/run", methods=["POST"])
def run():
    """触发业务流程，并发保护：同一时间只允许一个任务"""
    global is_running

    with task_lock:
        if is_running:
            return jsonify({
                "ok":  False,
                "msg": "⚠️ 当前有任务正在处理，请等待完成后再试"
            }), 429
        is_running = True

    while not log_queue.empty():
        try:
            log_queue.get_nowait()
        except queue.Empty:
            break

    process_month = get_process_month()
    data_folder   = os.path.join("data",   process_month)
    output_folder = os.path.join("output", process_month)

    xlsx_files = [
        f for f in os.listdir(data_folder)
        if f.endswith(".xlsx") and not f.startswith("~")
    ] if os.path.exists(data_folder) else []

    if not xlsx_files:
        is_running = False
        return jsonify({
            "ok":  False,
            "msg": f"❌ data/{process_month}/ 目录为空，请先上传账单文件"
        }), 400

    t = threading.Thread(
        target=run_task,
        args=(data_folder, output_folder, process_month),
        daemon=True
    )
    t.start()

    return jsonify({"ok": True, "msg": "任务已启动"})


# ====================== 路由：实时日志（SSE）======================
@app.route("/logs")
def logs():
    """Server-Sent Events：实时把日志队列内容推送给浏览器"""
    def generate():
        while True:
            try:
                msg = log_queue.get(timeout=30)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg == "__DONE__":
                    break
            except queue.Empty:
                yield f"data: {json.dumps('__PING__', ensure_ascii=False)}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control":               "no-cache",
            "X-Accel-Buffering":           "no",
            "Access-Control-Allow-Origin": "*"
        }
    )


# ====================== 路由：查询任务状态 ======================
@app.route("/status")
def status():
    return jsonify({
        "running": is_running,
        "result":  task_result
    })


# ====================== 路由：下载当月结果zip ======================
@app.route("/download")
def download():
    """下载当月结果zip"""
    process_month = get_process_month()
    return download_month(process_month)


# ====================== 路由：下载指定月份zip ======================
@app.route("/download/<month>")
def download_month_route(month):
    """下载指定月份结果zip，供历史记录页面使用"""
    return download_month(month)


def download_month(month):
    """打包指定月份 output/YYYY-MM/ 为zip提供下载"""
    output_folder = os.path.join("output", month)

    if not os.path.exists(output_folder):
        return "结果文件不存在", 404

    zip_path = os.path.join("output", f"{month}_结果.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname   = os.path.relpath(file_path, "output")
                zf.write(file_path, arcname)

    return send_file(
        os.path.abspath(zip_path),
        as_attachment=True,
        download_name=f"{month}_对账结果.zip"
    )


# ====================== 路由：历史记录 ======================
@app.route("/history")
def history():
    """
    扫描 output/ 目录，返回所有月份的处理记录
    读取每个月的 run.log 最后一次运行结果
    按月份倒序排列
    """
    output_base = "output"
    records     = []

    if not os.path.exists(output_base):
        return jsonify([])

    for folder in sorted(os.listdir(output_base), reverse=True):
        folder_path = os.path.join(output_base, folder)
        if not os.path.isdir(folder_path):
            continue
        if len(folder) != 7 or folder[4] != "-":
            continue

        log_path = os.path.join(folder_path, "run.log")
        record   = {
            "month":         folder,
            "has_log":       os.path.exists(log_path),
            "run_count":     0,
            "last_result":   "无记录",
            "last_time":     "",
            "last_duration": ""
        }

        if os.path.exists(log_path):
            try:
                with open(log_path, "r", encoding="utf-8") as f:
                    content_log = f.read()

                record["run_count"] = content_log.count("【第")

                sections = content_log.split("【第")
                if len(sections) > 1:
                    last_section = sections[-1]

                    for line in last_section.split("\n"):
                        if "】" in line:
                            record["last_time"] = line.split("】")[-1].strip()
                            break

                    if "成功 ✅" in last_section:
                        record["last_result"] = "成功"
                    elif "失败 ❌" in last_section:
                        record["last_result"] = "失败"

                    for line in last_section.split("\n"):
                        if line.startswith("耗时："):
                            record["last_duration"] = line.replace("耗时：", "").strip()
                            break

            except Exception:
                record["last_result"] = "日志读取失败"

        records.append(record)

    return jsonify(records)




# ====================== 路由：统计数据 ======================
@app.route("/stats/<month>")
def stats(month):
    """
    读取指定月份 output/YYYY-MM/最终对账结果.xlsx
    按团队分组求和，返回统计数据供图表渲染
    同时扫描所有月份，返回趋势数据
    """
    import pandas as pd

    output_base = "output"
    result      = {
        "month":        month,
        "team_stats":   [],   # 本月各团队费用
        "monthly_trend": []   # 各月总费用趋势
    }

    # ---- 本月各团队费用 ----
    result_file = os.path.join(output_base, month, "最终对账结果.xlsx")
    if os.path.exists(result_file):
        try:
            df = pd.read_excel(result_file, engine="openpyxl")
            if "所属团队" in df.columns and "单票应付金额" in df.columns:
                df["单票应付金额"] = pd.to_numeric(df["单票应付金额"], errors="coerce").fillna(0)
                team_group = df.groupby("所属团队")["单票应付金额"].sum().reset_index()
                team_group = team_group[team_group["单票应付金额"] > 0]
                team_group = team_group.sort_values("单票应付金额", ascending=False)
                result["team_stats"] = [
                    {"team": row["所属团队"], "amount": round(float(row["单票应付金额"]), 2)}
                    for _, row in team_group.iterrows()
                ]
        except Exception as e:
            print(f"读取统计数据失败：{str(e)}")

    # ---- 各月总费用趋势 ----
    monthly = []
    if os.path.exists(output_base):
        for folder in sorted(os.listdir(output_base)):
            folder_path = os.path.join(output_base, folder)
            if not os.path.isdir(folder_path):
                continue
            if len(folder) != 7 or folder[4] != "-":
                continue
            f = os.path.join(folder_path, "最终对账结果.xlsx")
            if not os.path.exists(f):
                continue
            try:
                df_m = pd.read_excel(f, engine="openpyxl")
                if "单票应付金额" in df_m.columns:
                    df_m["单票应付金额"] = pd.to_numeric(df_m["单票应付金额"], errors="coerce").fillna(0)
                    total = round(float(df_m["单票应付金额"].sum()), 2)
                    monthly.append({"month": folder, "total": total})
            except Exception:
                continue
    result["monthly_trend"] = monthly

    return jsonify(result)



# ====================== 路由：读取价格配置 ======================
@app.route("/price/get")
def price_get():
    """
    读取 config/price_config.xlsx 三个Sheet
    返回申通报价、中通报价、充单价格供前端渲染表格
    """
    import pandas as pd

    price_file = os.path.join("config", "price_config.xlsx")
    if not os.path.exists(price_file):
        return jsonify({"ok": False, "msg": "price_config.xlsx 不存在"}), 404

    result = {"ok": True, "shentong": [], "zhongtong": [], "charge": []}

    try:
        # 读取申通报价
        df_st = pd.read_excel(price_file, sheet_name="申通报价", engine="openpyxl")
        for _, row in df_st.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
                continue
            result["shentong"].append({
                "province":   str(row.iloc[0]).strip(),
                "fee_3kg":    float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                "fee_over3kg": float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                "unit_price": float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
            })

        # 读取中通报价
        df_zt = pd.read_excel(price_file, sheet_name="中通报价", engine="openpyxl")
        for _, row in df_zt.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
                continue
            result["zhongtong"].append({
                "province":    str(row.iloc[0]).strip(),
                "fee_3kg":     float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                "fee_over3kg": float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                "unit_price":  float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
            })

        # 读取充单价格
        df_charge = pd.read_excel(price_file, sheet_name="客户快递加收单价信息记录", engine="openpyxl")
        for _, row in df_charge.iterrows():
            if pd.isna(row.iloc[10]) or pd.isna(row.iloc[11]):
                continue
            type_name = str(row.iloc[10]).strip()
            if type_name in ["申通", "中通"]:
                result["charge"].append({
                    "type":   type_name,
                    "price":  float(row.iloc[11])
                })

    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500

    return jsonify(result)


# ====================== 路由：保存价格配置 ======================
@app.route("/price/save", methods=["POST"])
def price_save():
    """
    接收前端修改后的报价数据，更新 price_config.xlsx 对应Sheet
    只更新B/D/E列（3kg内面单、超3kg面单、续重单价）和充单价格
    不改变Excel的其他格式和内容
    """
    import pandas as pd
    from openpyxl import load_workbook

    data       = request.json
    price_file = os.path.join("config", "price_config.xlsx")

    if not os.path.exists(price_file):
        return jsonify({"ok": False, "msg": "price_config.xlsx 不存在"}), 404

    try:
        wb = load_workbook(price_file)

        # 更新申通报价
        if "shentong" in data and "申通报价" in wb.sheetnames:
            ws = wb["申通报价"]
            province_map = {str(row[0].value).strip(): row for row in ws.iter_rows(min_row=2) if row[0].value}
            for item in data["shentong"]:
                prov = item["province"]
                if prov in province_map:
                    row = province_map[prov]
                    row[1].value = item["fee_3kg"]      # B列
                    row[3].value = item["fee_over3kg"]  # D列
                    row[4].value = item["unit_price"]   # E列

        # 更新中通报价
        if "zhongtong" in data and "中通报价" in wb.sheetnames:
            ws = wb["中通报价"]
            province_map = {str(row[0].value).strip(): row for row in ws.iter_rows(min_row=2) if row[0].value}
            for item in data["zhongtong"]:
                prov = item["province"]
                if prov in province_map:
                    row = province_map[prov]
                    row[1].value = item["fee_3kg"]
                    row[3].value = item["fee_over3kg"]
                    row[4].value = item["unit_price"]

        # 更新充单价格
        if "charge" in data and "客户快递加收单价信息记录" in wb.sheetnames:
            ws = wb["客户快递加收单价信息记录"]
            for row in ws.iter_rows(min_row=2):
                if row[10].value and str(row[10].value).strip() in ["申通", "中通"]:
                    type_name = str(row[10].value).strip()
                    for item in data["charge"]:
                        if item["type"] == type_name:
                            row[11].value = item["price"]

        wb.save(price_file)
        return jsonify({"ok": True, "msg": "价格配置已保存"})

    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500

# ====================== 启动 ======================
if __name__ == "__main__":
    print("=" * 60)
    print("  YiBo Express Bill System - Web Service")
    print("  LAN Access:   http://[LAN IP]:5000")
    print("  Local Access: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
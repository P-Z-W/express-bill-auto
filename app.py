# -*- coding: utf-8 -*-
"""
毅播快递对账系统 - Web服务主程序 V3.0.1
==================================================
【V3.0.1 性能优化】
  1. Excel 读取缓存（_excel_cache）
     - 相同文件、相同 mtime → 直接返回内存 DataFrame，跳过磁盘IO
     - 文件更新后自动失效，保证数据准确
  2. /stats/<month> 结果缓存（_stats_cache）
     - 计算结果按月缓存，文件未变动时直接返回
     - monthly_trend 同理，只在有新月份文件时重新扫描
  3. /history/api 结果缓存（_history_cache）
     - 按 output/ 目录 mtime 变化自动失效
  4. 前端 dashboard 并行请求（Promise.all）
     - /history/api + /stats/ + /unmatched/ 三个请求并发发出
     - 总耗时 ≈ 最慢那个，而不是三者之和
  5. /stats/<month> 中 monthly_trend 仅读 "单票应付金额" 一列
     - 用 usecols 参数大幅减少 Excel 读取数据量

【V3.0 BugFix（保留）】
  - 修复 /history 路由注册两次导致的路由冲突
  - 拆分 /history（页面）和 /history/api（JSON接口）

【路由说明】
  /          → 首页看板（HTML）
  /run       → 运行操作页（HTML）
  /history   → 历史记录页（HTML）
  /stats     → 统计报表页（HTML）
  /config    → 系统配置页（HTML）

【API接口】
  /upload              POST  上传账单文件
  /run                 POST  触发运行
  /logs                GET   SSE实时日志
  /status              GET   任务状态
  /download            GET   下载当月zip
  /download/<month>    GET   下载指定月份zip
  /history/api         GET   历史记录列表（JSON，带缓存）
  /stats/<month>       GET   统计数据（JSON，带缓存）
  /unmatched/<month>   GET   未匹配分析（JSON，带缓存）
  /preview/<month>     GET   对账结果预览（JSON）
  /price/get           GET   读取价格配置
  /price/save          POST  保存价格配置
  /express/config      GET   读取快递配置
  /express/save        POST  保存快递配置
==================================================
"""
import os
import sys
import io
import json
import queue
import threading
import zipfile
from datetime import datetime, timedelta
from flask import Flask, render_template, request, Response, jsonify, send_file

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from config import settings

app = Flask(__name__)

# ====================== 全局状态 ======================
task_lock   = threading.Lock()
is_running  = False
log_queue   = queue.Queue()
task_result = {"success": False, "output_folder": "", "elapsed": ""}

EXPRESS_CONFIG_PATH = os.path.join("config", "express_config.json")

# ====================== 性能缓存 ======================
# Excel 文件缓存：key=文件路径, value={"mtime": float, "df": DataFrame}
# 同一文件只要 mtime 没变，直接返回内存中的 DataFrame，无需重复读磁盘
_excel_cache   = {}
_cache_lock    = threading.Lock()

# 接口结果缓存
# _stats_cache:   key=month, value={"file_mtime": float, "result": dict}
# _history_cache: {"dir_mtime": float, "result": list}
# _unmatched_cache: key=month, value={"file_mtime": float, "result": dict}
_stats_cache     = {}
_history_cache   = {"dir_mtime": -1, "result": None}
_unmatched_cache = {}


def _get_mtime(path):
    """安全获取文件/目录的修改时间，不存在返回 -1"""
    try:
        return os.path.getmtime(path)
    except OSError:
        return -1


def _read_excel_cached(file_path, usecols=None):
    """
    带缓存的 Excel 读取
    - 首次读取后缓存 DataFrame 到内存
    - 再次请求时检查 mtime：未变化直接返回缓存，避免重复磁盘IO
    - usecols 指定只读哪些列，减少大文件读取量
    """
    import pandas as pd
    if not os.path.exists(file_path):
        return None

    mtime = _get_mtime(file_path)
    cache_key = f"{file_path}|{str(usecols)}"

    with _cache_lock:
        cached = _excel_cache.get(cache_key)
        if cached and cached["mtime"] == mtime:
            return cached["df"]

    # 缓存未命中或文件已更新，重新读取
    try:
        df = pd.read_excel(file_path, engine="openpyxl", usecols=usecols)
        with _cache_lock:
            _excel_cache[cache_key] = {"mtime": mtime, "df": df}
        return df
    except Exception as e:
        print(f"读取Excel失败：{file_path}，{str(e)}")
        return None


def invalidate_cache(month=None):
    """
    任务完成后主动清除相关缓存，确保下次请求拿到最新数据
    month=None 时清除全部缓存
    """
    global _history_cache
    with _cache_lock:
        if month:
            # 清除指定月份的文件缓存和结果缓存
            keys_to_del = [k for k in _excel_cache if f"output/{month}" in k or f"output\\{month}" in k]
            for k in keys_to_del:
                del _excel_cache[k]
            _stats_cache.pop(month, None)
            _unmatched_cache.pop(month, None)
        else:
            _excel_cache.clear()
            _stats_cache.clear()
            _unmatched_cache.clear()
        _history_cache = {"dir_mtime": -1, "result": None}


# ====================== 工具函数 ======================
def get_process_month():
    today          = datetime.now()
    first_of_month = today.replace(day=1)
    last_month     = first_of_month - timedelta(days=1)
    return last_month.strftime("%Y-%m")


def load_express_config():
    """读取快递公司配置，文件不存在时返回默认值"""
    try:
        with open(EXPRESS_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"express_list": [
            {"name": "申通", "identify_column": "业务时间", "enabled": True},
            {"name": "中通", "identify_column": "扫描时间", "enabled": True}
        ]}


def push_log(msg):
    log_queue.put(msg)


# ====================== 日志工具 ======================
def get_log_path(output_folder):
    return os.path.join(output_folder, "run.log")


def get_run_count(log_path):
    if not os.path.exists(log_path):
        return 1
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            return f.read().count("【第") + 1
    except Exception:
        return 1


def write_log_header(log_file, process_month, run_count):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file.write("\n" + "=" * 60 + "\n")
    log_file.write(f"【第{run_count}次运行】{now}\n")
    log_file.write(f"处理月份：{process_month}\n")
    log_file.write("=" * 60 + "\n")
    log_file.flush()


def write_log_footer(log_file, success, start_time):
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
        push_log("=" * 60)
        push_log("__STEP__0")

        push_log(f"\n📌 第一步：从数据库下载 {process_month} 订单数据")
        import order_db
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            push_log("❌ 订单数据下载失败或为空，终止流程")
            push_log("__STEP_FAIL__1")
            task_result = {"success": False, "output_folder": output_folder, "elapsed": ""}
            return
        push_log("__STEP__1")

        push_log(f"\n📌 第二步：清洗合并 {process_month} 快递账单")
        import merge_express
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            push_log("❌ 快递账单合并失败或为空，终止流程")
            push_log("__STEP_FAIL__2")
            task_result = {"success": False, "output_folder": output_folder, "elapsed": ""}
            return
        push_log("__STEP__2")

        push_log(f"\n📌 第三步：运单号匹配对账（{process_month}）")
        import order_matching
        order_matching.run_reconciliation()
        push_log("__STEP__3")

        push_log(f"\n📌 第四步：按团队拆分客户账单（{process_month}）")
        import split_bill_by_team
        split_bill_by_team.main()
        push_log("__STEP__4")

        push_log("\n" + "=" * 60)
        push_log(f"🎉 {process_month} 全部流程执行完成！")

        elapsed     = datetime.now() - start_time
        minutes     = int(elapsed.total_seconds() // 60)
        seconds     = int(elapsed.total_seconds() % 60)
        elapsed_str = f"{minutes}分{seconds}秒"
        push_log(f"⏱ 总耗时：{elapsed_str}")

        success     = True
        task_result = {"success": True, "output_folder": output_folder, "elapsed": elapsed_str}

    except Exception as e:
        import traceback
        push_log(f"\n❌ 程序运行异常：{str(e)}")
        push_log(traceback.format_exc())
        elapsed     = datetime.now() - start_time
        elapsed_str = f"{int(elapsed.total_seconds()//60)}分{int(elapsed.total_seconds()%60)}秒"
        task_result = {"success": False, "output_folder": output_folder, "elapsed": elapsed_str}

    finally:
        write_log_footer(log_file, success, start_time)
        log_file.close()
        sys.stdout = real_stdout
        is_running = False
        # 任务完成后清除该月缓存，下次请求拿到最新数据
        invalidate_cache(process_month)
        push_log("__DONE__")


# ====================== 页面路由（HTML）======================
@app.route("/")
def dashboard():
    return render_template("dashboard.html",
        process_month=get_process_month(),
        active_page="dashboard")


@app.route("/run")
def run_page():
    config = load_express_config()
    return render_template("run.html",
        process_month=get_process_month(),
        active_page="run",
        express_list=config.get("express_list", []))


@app.route("/history")
def history_page():
    return render_template("history.html",
        process_month=get_process_month(),
        active_page="history")


@app.route("/stats")
def stats_page():
    return render_template("stats.html",
        process_month=get_process_month(),
        active_page="stats")


@app.route("/config")
def config_page():
    return render_template("config.html",
        process_month=get_process_month(),
        active_page="config")


# ====================== API：上传文件 ======================
@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "没有收到文件"}), 400
    file          = request.files["file"]
    process_month = get_process_month()
    data_folder   = os.path.join("data", process_month)
    os.makedirs(data_folder, exist_ok=True)
    save_path = os.path.join(data_folder, file.filename)
    file.save(save_path)
    return jsonify({"ok": True, "msg": f"✅ {file.filename} 已上传到 data/{process_month}/"})


# ====================== API：触发运行 ======================
@app.route("/run", methods=["POST"])
def run():
    global is_running
    with task_lock:
        if is_running:
            return jsonify({"ok": False, "msg": "⚠️ 当前有任务正在处理，请等待完成后再试"}), 429
        is_running = True

    while not log_queue.empty():
        try: log_queue.get_nowait()
        except queue.Empty: break

    process_month = get_process_month()
    data_folder   = os.path.join("data",   process_month)
    output_folder = os.path.join("output", process_month)

    xlsx_files = [
        f for f in os.listdir(data_folder)
        if f.endswith(".xlsx") and not f.startswith("~")
    ] if os.path.exists(data_folder) else []

    if not xlsx_files:
        is_running = False
        return jsonify({"ok": False, "msg": f"❌ data/{process_month}/ 目录为空，请先上传账单文件"}), 400

    t = threading.Thread(
        target=run_task,
        args=(data_folder, output_folder, process_month),
        daemon=True
    )
    t.start()
    return jsonify({"ok": True, "msg": "任务已启动"})


# ====================== API：SSE实时日志 ======================
@app.route("/logs")
def logs():
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
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Access-Control-Allow-Origin": "*"
        }
    )


# ====================== API：任务状态 ======================
@app.route("/status")
def status():
    return jsonify({"running": is_running, "result": task_result})


# ====================== API：下载zip ======================
@app.route("/download")
def download():
    return download_month(get_process_month())


@app.route("/download/<month>")
def download_month_route(month):
    return download_month(month)


def download_month(month):
    output_folder = os.path.join("output", month)
    if not os.path.exists(output_folder):
        return "结果文件不存在", 404
    zip_path = os.path.join("output", f"{month}_结果.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                file_path = os.path.join(root, file)
                zf.write(file_path, os.path.relpath(file_path, "output"))
    return send_file(
        os.path.abspath(zip_path),
        as_attachment=True,
        download_name=f"{month}_对账结果.zip"
    )


# ====================== API：历史记录（JSON，带缓存）======================
@app.route("/history/api")
def history_api():
    """
    历史记录 JSON 接口
    缓存策略：监听 output/ 目录 mtime，目录结构未变化时直接返回缓存
    """
    global _history_cache
    output_base = "output"
    dir_mtime   = _get_mtime(output_base)

    # 命中缓存
    with _cache_lock:
        if _history_cache["result"] is not None and _history_cache["dir_mtime"] == dir_mtime:
            return jsonify(_history_cache["result"])

    if not os.path.exists(output_base):
        return jsonify([])

    records = []
    for folder in sorted(os.listdir(output_base), reverse=True):
        folder_path = os.path.join(output_base, folder)
        if not os.path.isdir(folder_path) or len(folder) != 7 or folder[4] != "-":
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

    with _cache_lock:
        _history_cache = {"dir_mtime": dir_mtime, "result": records}

    return jsonify(records)


# ====================== API：统计数据（带缓存）======================
@app.route("/stats/<month>")
def stats(month):
    import pandas as pd
    output_base = "output"
    result_file = os.path.join(output_base, month, "最终对账结果.xlsx")
    file_mtime  = _get_mtime(result_file)

    # 命中缓存：文件未变动直接返回
    with _cache_lock:
        cached = _stats_cache.get(month)
        if cached and cached["file_mtime"] == file_mtime and file_mtime != -1:
            return jsonify(cached["result"])

    result = {
        "month":         month,
        "team_stats":    [],
        "monthly_trend": [],
        "team_summary":  []
    }

    # 读取当月统计（使用缓存读取）
    if os.path.exists(result_file):
        try:
            df = _read_excel_cached(result_file)
            if df is not None:
                df = df[df["所属团队"] != "未匹配"] if "所属团队" in df.columns else df
                if "所属团队" in df.columns and "单票应付金额" in df.columns:
                    df["单票应付金额"] = pd.to_numeric(df["单票应付金额"], errors="coerce").fillna(0)
                    team_group = df.groupby("所属团队").agg(
                        amount=("单票应付金额", "sum"),
                        count=("运单号", "count")
                    ).reset_index()
                    team_group = team_group[team_group["count"] > 0].sort_values("amount", ascending=False)
                    result["team_stats"] = [
                        {"team": row["所属团队"], "amount": round(float(row["amount"]), 2)}
                        for _, row in team_group.iterrows()
                    ]

                    summary = []
                    total_single = total_average = total_amount = 0
                    for team, group in df.groupby("所属团队"):
                        single_amount = round(float(
                            group[group["实际计算方式"] == "单票"]["单票应付金额"].sum()
                        ), 2) if "实际计算方式" in group.columns else 0
                        average_count = len(
                            group[group["实际计算方式"] == "全国均重"]
                        ) if "实际计算方式" in group.columns else 0
                        team_amount   = round(float(group["单票应付金额"].sum()), 2)
                        total_single += single_amount
                        total_average += average_count
                        total_amount  += team_amount
                        summary.append({
                            "team":          team,
                            "single_amount": single_amount,
                            "average_count": average_count,
                            "total_amount":  team_amount
                        })
                    summary.sort(key=lambda x: x["total_amount"], reverse=True)
                    summary.append({
                        "team":          "合计",
                        "single_amount": round(total_single, 2),
                        "average_count": total_average,
                        "total_amount":  round(total_amount, 2)
                    })
                    result["team_summary"] = summary
        except Exception as e:
            print(f"读取统计数据失败：{str(e)}")

    # 汇总各月趋势
    # 优化：只读"单票应付金额"一列，减少大文件读取量
    monthly = []
    if os.path.exists(output_base):
        for folder in sorted(os.listdir(output_base)):
            folder_path = os.path.join(output_base, folder)
            if not os.path.isdir(folder_path) or len(folder) != 7 or folder[4] != "-":
                continue
            f = os.path.join(folder_path, "最终对账结果.xlsx")
            if not os.path.exists(f):
                continue
            try:
                # usecols 只读需要的列，速度比读全表快数倍
                df_m = _read_excel_cached(f, usecols=["单票应付金额"])
                if df_m is not None and "单票应付金额" in df_m.columns:
                    df_m["单票应付金额"] = pd.to_numeric(
                        df_m["单票应付金额"], errors="coerce"
                    ).fillna(0)
                    monthly.append({
                        "month": folder,
                        "total": round(float(df_m["单票应付金额"].sum()), 2)
                    })
            except Exception:
                continue
    result["monthly_trend"] = monthly

    # 写入缓存
    with _cache_lock:
        _stats_cache[month] = {"file_mtime": file_mtime, "result": result}

    return jsonify(result)


# ====================== API：未匹配分析（带缓存）======================
@app.route("/unmatched/<month>")
def unmatched(month):
    import pandas as pd
    result_file = os.path.join("output", month, "最终对账结果.xlsx")
    if not os.path.exists(result_file):
        return jsonify({"ok": False, "msg": f"{month} 对账结果文件不存在"}), 404

    file_mtime = _get_mtime(result_file)

    # 命中缓存
    with _cache_lock:
        cached = _unmatched_cache.get(month)
        if cached and cached["file_mtime"] == file_mtime:
            return jsonify(cached["result"])

    df = _read_excel_cached(result_file)
    if df is None:
        return jsonify({"ok": False, "msg": "读取文件失败"}), 500

    total           = len(df)
    df_unmatched    = df[df["所属团队"] == "未匹配"] if "所属团队" in df.columns else df
    unmatched_count = len(df_unmatched)
    matched_count   = total - unmatched_count
    ratio           = round(unmatched_count / total * 100, 1) if total > 0 else 0
    by_express      = {}
    if "快递类型" in df_unmatched.columns:
        for t, g in df_unmatched.groupby("快递类型"):
            by_express[str(t)] = len(g)
    samples = (
        df_unmatched["运单号"].astype(str).head(5).tolist()
        if "运单号" in df_unmatched.columns else []
    )
    result = {
        "ok":         True,
        "month":      month,
        "total":      total,
        "matched":    matched_count,
        "unmatched":  unmatched_count,
        "ratio":      ratio,
        "by_express": by_express,
        "samples":    samples
    }

    with _cache_lock:
        _unmatched_cache[month] = {"file_mtime": file_mtime, "result": result}

    return jsonify(result)


# ====================== API：对账结果预览 ======================
@app.route("/preview/<month>")
def preview(month):
    import pandas as pd
    result_file = os.path.join("output", month, "最终对账结果.xlsx")
    if not os.path.exists(result_file):
        return jsonify({"ok": False, "msg": f"{month} 对账结果文件不存在"}), 404

    # 预览接口复用缓存的完整 DataFrame，避免重复读磁盘
    df = _read_excel_cached(result_file)
    if df is None:
        return jsonify({"ok": False, "msg": "读取文件失败"}), 500

    df      = df.fillna("")
    page    = int(request.args.get("page", 1))
    size    = int(request.args.get("size", 100))
    filter_ = request.args.get("filter", "all")
    keyword = request.args.get("keyword", "").strip()

    total           = len(df)
    matched_count   = len(df[df["所属团队"] != "未匹配"]) if "所属团队" in df.columns else 0
    unmatched_count = total - matched_count

    if "所属团队" in df.columns and "实际计算方式" in df.columns:
        if filter_ == "matched":     df = df[df["所属团队"] != "未匹配"]
        elif filter_ == "unmatched": df = df[df["所属团队"] == "未匹配"]
        elif filter_ == "single":    df = df[df["实际计算方式"] == "单票"]
        elif filter_ == "average":   df = df[df["实际计算方式"] == "全国均重"]

    if keyword:
        mask = pd.Series([False] * len(df), index=df.index)
        for col in ["运单号", "所属团队"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.contains(keyword, na=False)
        df = df[mask]

    filtered_count = len(df)
    page_df        = df.iloc[(page - 1) * size: page * size]
    show_cols      = [
        c for c in ["运单号", "所属团队", "目的省份", "结算重量", "快递类型", "实际计算方式", "单票应付金额"]
        if c in page_df.columns
    ]
    return jsonify({
        "ok":          True,
        "month":       month,
        "total":       total,
        "matched":     matched_count,
        "unmatched":   unmatched_count,
        "filtered":    filtered_count,
        "page":        page,
        "size":        size,
        "total_pages": max(1, -(-filtered_count // size)),
        "rows":        page_df[show_cols].to_dict(orient="records")
    })


# ====================== API：价格配置 ======================
@app.route("/price/get")
def price_get():
    import pandas as pd
    price_file = os.path.join("config", "price_config.xlsx")
    if not os.path.exists(price_file):
        return jsonify({"ok": False, "msg": "price_config.xlsx 不存在"}), 404
    result = {"ok": True, "shentong": [], "zhongtong": [], "charge": []}
    try:
        for sheet, key in [("申通报价", "shentong"), ("中通报价", "zhongtong")]:
            df = pd.read_excel(price_file, sheet_name=sheet)
            for _, row in df.iterrows():
                if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "":
                    continue
                result[key].append({
                    "province":    str(row.iloc[0]).strip(),
                    "fee_3kg":     float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                    "fee_over3kg": float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                    "unit_price":  float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
                })
        df_charge = pd.read_excel(price_file, sheet_name="客户快递加收单价信息记录")
        for _, row in df_charge.iterrows():
            if pd.isna(row.iloc[10]) or pd.isna(row.iloc[11]):
                continue
            type_name = str(row.iloc[10]).strip()
            if type_name in ["申通", "中通"]:
                result["charge"].append({"type": type_name, "price": float(row.iloc[11])})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500
    return jsonify(result)


@app.route("/price/save", methods=["POST"])
def price_save():
    from openpyxl import load_workbook
    data       = request.json
    price_file = os.path.join("config", "price_config.xlsx")
    if not os.path.exists(price_file):
        return jsonify({"ok": False, "msg": "price_config.xlsx 不存在"}), 404
    try:
        wb = load_workbook(price_file)
        for sheet_key, sheet_name in [("shentong", "申通报价"), ("zhongtong", "中通报价")]:
            if sheet_key in data and sheet_name in wb.sheetnames:
                ws   = wb[sheet_name]
                pmap = {
                    str(row[0].value).strip(): row
                    for row in ws.iter_rows(min_row=2) if row[0].value
                }
                for item in data[sheet_key]:
                    if item["province"] in pmap:
                        row          = pmap[item["province"]]
                        row[1].value = item["fee_3kg"]
                        row[3].value = item["fee_over3kg"]
                        row[4].value = item["unit_price"]
        if "charge" in data and "客户快递加收单价信息记录" in wb.sheetnames:
            ws = wb["客户快递加收单价信息记录"]
            for row in ws.iter_rows(min_row=2):
                if row[10].value and str(row[10].value).strip() in ["申通", "中通"]:
                    for item in data["charge"]:
                        if item["type"] == str(row[10].value).strip():
                            row[11].value = item["price"]
        wb.save(price_file)
        return jsonify({"ok": True, "msg": "价格配置已保存"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


# ====================== API：快递配置 ======================
@app.route("/express/config")
def express_config():
    config = load_express_config()
    return jsonify({"ok": True, **config})


@app.route("/express/save", methods=["POST"])
def express_save():
    data = request.json
    try:
        with open(EXPRESS_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return jsonify({"ok": True, "msg": "快递配置已保存"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


# ====================== 启动 ======================
if __name__ == "__main__":
    print("=" * 60)
    print("  YiBo Express Bill System V3.0.1")
    print("  LAN Access:   http://[LAN IP]:5000")
    print("  Local Access: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
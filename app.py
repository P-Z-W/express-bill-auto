# -*- coding: utf-8 -*-
"""
毅播快递对账系统 - Web服务主程序
==================================================
【说明】
  Flask轻量Web服务，提供局域网浏览器操作界面
  五个业务模块完全不改动，app.py只负责：
    1. 接收上传的申通/中通账单，存入 data/YYYY-MM/
    2. 按顺序调用四个业务模块，实时推送日志到浏览器
    3. 运行完成后打包 output/YYYY-MM/ 提供下载

【启动方式】
  python app.py
  浏览器访问 http://局域网IP:5000

【并发保护】
  同一时间只允许一个任务运行，其他人点击会收到提示

【V2.1修复】
  修复 push_log 和 LogCapture 互相触发导致的日志死循环问题
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
    """
    推送日志到队列
    【修复】去掉 print()，断开和 LogCapture 的循环触发链
    只往队列写，不再调用print，彻底避免死循环
    """
    log_queue.put(msg)


class LogCapture(io.StringIO):
    """
    劫持标准输出，捕获业务模块所有 print() 推入队列
    【修复】直接写队列，不调用 push_log，彻底断开循环链
    保留 _real_stdout 把内容同步打印到服务器控制台方便调试
    """
    def __init__(self, real_stdout):
        super().__init__()
        self._real_stdout = real_stdout

    def write(self, msg):
        if msg.strip():
            log_queue.put(msg.strip())            # 直接写队列，不经过push_log
            self._real_stdout.write(msg + "\n")   # 同步打印到服务器控制台
        return len(msg)

    def flush(self):
        pass  # 防止flush报错


# ====================== 业务任务线程 ======================
def run_task(data_folder, output_folder, process_month):
    """
    独立线程里按顺序执行四个业务模块
    LogCapture 劫持标准输出，所有模块的 print 实时推送到浏览器
    """
    global is_running, task_result

    real_stdout = sys.stdout           # 保存真实stdout
    sys.stdout  = LogCapture(real_stdout)  # 替换为劫持版本

    try:
        push_log(f"🚀 开始处理 {process_month} 月份数据")
        push_log(f"📂 数据目录：{data_folder}")
        push_log(f"📂 输出目录：{output_folder}")
        push_log("=" * 60)

        os.makedirs(output_folder, exist_ok=True)

        # ---------- 第一步：数据库下载订单 ----------
        push_log(f"\n📌 第一步：从数据库下载 {process_month} 订单数据")
        import order_db
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            push_log("❌ 订单数据下载失败或为空，终止流程")
            task_result = {"success": False, "output_folder": output_folder}
            return

        # ---------- 第二步：合并快递账单 ----------
        push_log(f"\n📌 第二步：清洗合并 {process_month} 快递账单")
        import merge_express
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            push_log("❌ 快递账单合并失败或为空，终止流程")
            task_result = {"success": False, "output_folder": output_folder}
            return

        # ---------- 第三步：运单匹配对账 ----------
        push_log(f"\n📌 第三步：运单号匹配对账（{process_month}）")
        import order_matching
        order_matching.run_reconciliation()

        # ---------- 第四步：按团队拆分账单 ----------
        push_log(f"\n📌 第四步：按团队拆分客户账单（{process_month}）")
        import split_bill_by_team
        split_bill_by_team.main()

        push_log("\n" + "=" * 60)
        push_log(f"🎉 {process_month} 全部流程执行完成！")
        push_log("📦 可点击下方按钮下载结果文件")

        task_result = {"success": True, "output_folder": output_folder}

    except Exception as e:
        import traceback
        push_log(f"\n❌ 程序运行异常：{str(e)}")
        push_log(traceback.format_exc())
        task_result = {"success": False, "output_folder": output_folder}

    finally:
        sys.stdout = real_stdout   # 恢复真实stdout
        is_running = False
        push_log("__DONE__")       # 通知前端任务结束


# ====================== 路由：首页 ======================
@app.route("/")
def index():
    process_month = get_process_month()
    return render_template("index.html", process_month=process_month)


# ====================== 路由：上传账单文件 ======================
@app.route("/upload", methods=["POST"])
def upload():
    """
    接收上传文件存入 data/YYYY-MM/
    文件保留原始文件名，merge_express.py 自动扫描识别申通/中通
    """
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

    # 清空日志队列
    while not log_queue.empty():
        try:
            log_queue.get_nowait()
        except queue.Empty:
            break

    process_month = get_process_month()
    data_folder   = os.path.join("data",   process_month)
    output_folder = os.path.join("output", process_month)

    # 检查是否有账单文件
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

    # 在独立线程里跑业务，不阻塞 Web 服务
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
    """
    Server-Sent Events：实时把日志队列内容推送给浏览器
    前端 EventSource 监听，收到日志实时显示
    """
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


# ====================== 路由：下载结果zip ======================
@app.route("/download")
def download():
    """
    把 output/YYYY-MM/ 打包成 zip 提供浏览器下载
    本地 output 文件永久保留作为备份
    """
    process_month = get_process_month()
    output_folder = os.path.join("output", process_month)

    if not os.path.exists(output_folder):
        return "结果文件不存在，请先运行任务", 404

    zip_path = os.path.join("output", f"{process_month}_结果.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname   = os.path.relpath(file_path, "output")
                zf.write(file_path, arcname)

    return send_file(
        os.path.abspath(zip_path),
        as_attachment=True,
        download_name=f"{process_month}_对账结果.zip"
    )


# ====================== 启动 ======================
if __name__ == "__main__":
    print("=" * 60)
    print("  毅播快递对账系统 Web服务启动")
    print("  局域网访问地址：http://[本机IP]:5000")
    print("  本机访问地址：  http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
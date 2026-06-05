# -*- coding: utf-8 -*-
import os
import sys
import io
import queue
import threading
from datetime import datetime

from core.cache import invalidate_cache

task_lock   = threading.Lock()
is_running  = False
log_queue   = queue.Queue()
task_result = {"success": False, "output_folder": "", "elapsed": ""}


def push_log(msg):
    log_queue.put(msg)


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


class LogCapture(io.StringIO):
    def __init__(self, real_stdout, log_file=None, task_thread_id=None):
        super().__init__()
        self._real_stdout    = real_stdout
        self._log_file       = log_file
        self._task_thread_id = task_thread_id

    def write(self, msg):
        is_task = (
            self._task_thread_id is None or
            threading.current_thread().ident == self._task_thread_id
        )
        if msg.strip() and is_task:
            log_queue.put(msg.strip())
            if self._log_file:
                self._log_file.write(msg.strip() + "\n")
                self._log_file.flush()
        self._real_stdout.write(msg)
        return len(msg)

    def flush(self):
        pass


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
    sys.stdout  = LogCapture(real_stdout, log_file, task_thread_id=threading.current_thread().ident)

    try:
        push_log(f"🚀 开始处理 {process_month} 月份数据")
        push_log("=" * 60)
        push_log("__STEP__0")

        push_log(f"\n📌 第一步：从数据库下载 {process_month} 订单数据")
        from modules.express import order_db
        order_df = order_db.run_download_orders()
        if order_df is None or order_df.empty:
            push_log("❌ 订单数据下载失败或为空，终止流程")
            push_log("__STEP_FAIL__1")
            task_result = {"success": False, "output_folder": output_folder, "elapsed": ""}
            return
        push_log("__STEP__1")

        push_log(f"\n📌 第二步：清洗合并 {process_month} 快递账单")
        from modules.express import merge_express
        express_df = merge_express.run_merge_process()
        if express_df is None or express_df.empty:
            push_log("❌ 快递账单合并失败或为空，终止流程")
            push_log("__STEP_FAIL__2")
            task_result = {"success": False, "output_folder": output_folder, "elapsed": ""}
            return
        push_log("__STEP__2")

        push_log(f"\n📌 第三步：运单号匹配对账（{process_month}）")
        from modules.express import order_matching
        order_matching.run_reconciliation()
        push_log("__STEP__3")

        push_log(f"\n📌 第四步：按团队拆分客户账单（{process_month}）")
        from modules.express import split_bill_by_team
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
        invalidate_cache(process_month)
        push_log("__DONE__")
        is_running = False

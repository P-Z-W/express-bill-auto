# -*- coding: utf-8 -*-
import os
import threading

_excel_cache     = {}
_cache_lock      = threading.Lock()
_stats_cache     = {}
_history_cache   = {"dir_mtime": -1, "result": None}
_unmatched_cache = {}


def _get_mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return -1


def _read_excel_cached(file_path, usecols=None):
    import pandas as pd
    if not os.path.exists(file_path):
        return None
    mtime     = _get_mtime(file_path)
    cache_key = f"{file_path}|{str(usecols)}"
    with _cache_lock:
        cached = _excel_cache.get(cache_key)
        if cached and cached["mtime"] == mtime:
            return cached["df"]
    try:
        df = pd.read_excel(file_path, engine="openpyxl", usecols=usecols)
        with _cache_lock:
            _excel_cache[cache_key] = {"mtime": mtime, "df": df}
        return df
    except Exception as e:
        print(f"读取Excel失败：{file_path}，{str(e)}")
        return None


def invalidate_cache(month=None):
    global _history_cache
    with _cache_lock:
        if month:
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

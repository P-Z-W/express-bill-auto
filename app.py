# -*- coding: utf-8 -*-
"""
毅播云仓内部管理平台 V4.0
==================================================
多模块架构入口：注册各业务 Blueprint，启动 Flask 服务

模块说明：
  express  → 快递对账（/run /history /stats /config 及全部 API）
  storage  → 仓储费（占位，待开发）
  finance  → 财务汇总（占位，待开发）
==================================================
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask
from modules.express.routes import express_bp
from modules.storage.routes import storage_bp
from modules.finance.routes import finance_bp

app = Flask(__name__)
app.register_blueprint(express_bp)
app.register_blueprint(storage_bp, url_prefix="/storage")
app.register_blueprint(finance_bp, url_prefix="/finance")

if __name__ == "__main__":
    print("=" * 60)
    print("  YiBo BackOffice Platform V4.0")
    print("  LAN Access:   http://[LAN IP]:5000")
    print("  Local Access: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)

# -*- coding: utf-8 -*-
from flask import Blueprint, jsonify

finance_bp = Blueprint("finance", __name__)


@finance_bp.route("/")
def index():
    return jsonify({"module": "finance", "status": "coming_soon"})

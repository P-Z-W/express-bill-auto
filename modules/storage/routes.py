# -*- coding: utf-8 -*-
from flask import Blueprint, jsonify

storage_bp = Blueprint("storage", __name__)


@storage_bp.route("/")
def index():
    return jsonify({"module": "storage", "status": "coming_soon"})

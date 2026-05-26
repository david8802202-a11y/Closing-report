"""
PDF → Excel 結案表填充器
功能:上傳銀河專案 PDF,自動填入結案表模板對應位置
"""
import base64

# ==================== 內建模板資料 ====================
"""內建模板資料(base64 編碼)"""

CLOSURE_TEMPLATE_B64 = """UEsEMwYAAAAAFwAXACIGAABRegEAAAA=""" 
MONTHLY_TEMPLATE_B64 = """UEsFBYAAAAAFwAXACIGAABRegEAAAA=""" 

# ==================== 內建模板載入函式 ====================

def get_builtin_template(report_type):
    """取得內建模板的 bytes 資料"""
    if report_type == "月報表":
        return base64.b64decode(MONTHLY_TEMPLATE_B64)
    elif report_type == "結案表":
        return base64.b64decode(CLOSURE_TEMPLATE_B64)
    raise ValueError(f"未知的報表類型: {report_type}")


import streamlit as st
import pdfplumber
import re
import unicodedata
import subprocess
import tempfile
import os
import shutil
from pathlib import Path
from io import BytesIO
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="PDF → 報表填充器", page_icon="📊", layout="wide")

# ==================== PDF 抽取邏輯 ====================

def normalize(s):
    """去空白、Unicode 正規化(把 ⽂→文 等異體字統一)"""
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize('NFKC', s)
    return re.sub(r'\s+', '', s)


def is_main_data_row(row):
    """判斷是否為主表資料列"""
    if not row or len(row) < 13:
        return False
    站版 = normalize(row[0])
    if not 站版 or 站版 == "合計" or "站版" in 站版:
        return False
    try:
        int(row[4])
        return True
    except (ValueError, TypeError):
        return False


# 主軸縮寫對照(PDF 中主軸欄位常被切成單字)
主軸縮寫對照 = {
    "自": "自品置入",
    "需": "需求置入",
    "圖": "圖書/閱讀",
    "會": "會員/品牌",
    "百": "百貨",
    "競": "競品置入",
    "全": "全域置入",
    "社": "社團置入",
    "消": "消費者需求置入",
    "情": "情境話題",
    "產": "產品話題",
    "話": "話題",
}

# 站版前綴關鍵字 — 用來辨識一行是否以「站版」開頭
站版前綴 = ["Threads", "PTT", "Facebook", "Dcard", "Mobile01", "eyny", "Plurk",
             "BabyHome", "FG", "伊莉", "狄卡", "批踢踢", "FashionGuide"]

# 行尾「主軸 + 日期前半 + 9 個數字」的通用特徵
ROW_PATTERN = re.compile(
    r'([^\s\d][^\s]*)\s+'                                          # 主軸縮寫(非數字、非空白開頭)
    r'\d{4}-\s+'                                                   # 日期 yyyy-
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+'                            # 發文 回應 聲量 正向
    r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)'                       # 正面 負面 議題 產品 複合
    r'\s*$', re.MULTILINE
)


def get_main_category_from_pdf(pdf_path):
    """從 PDF「操作主軸」區塊抽出實際主軸全名清單"""
    with

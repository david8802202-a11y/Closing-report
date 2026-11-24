"""共用 Word 處理工具 — 給 closure_report.py 和 monthly_report.py 用

內容:
- XML 控制字元清理(_sanitize_data/_sanitize_str)
- 段落操作(_replace_paragraph_text、_insert_paragraph_after、_remove_paragraph)
- 儲存格操作(_set_cell_text)
- 日期工具(_minus_months、_count_business_days、_TW_HOLIDAYS)
"""
import re
from copy import deepcopy
from datetime import date, timedelta

from docx.oxml.ns import qn


# ==================== XML 字元清理 ====================
# XML 1.0 禁用字元:
# - 控制字元 0x00~0x08、0x0B、0x0C、0x0E~0x1F、0x7F
# - Unicode lone surrogates U+D800~U+DFFF
# - 非字元 code point U+FFFE、U+FFFF
# 合法保留:\t (0x09)、\n (0x0A)、\r (0x0D)
_INVALID_XML_CHARS_RE = re.compile(
    '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F'
    '\uD800-\uDFFF'
    '￾￿]'
)


def _sanitize_str(s):
    """從字串裡移除 XML 不接受的字元;非字串原樣回傳。
    多重保險:bytes→decode、regex 去控制字元、encode/decode 清理 lone surrogate。
    """
    if s is None:
        return s
    if isinstance(s, bytes):
        s = s.decode('utf-8', errors='replace')
    elif not isinstance(s, str):
        return s
    s = _INVALID_XML_CHARS_RE.sub('', s)
    s = s.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
    return s


def _sanitize_data(obj):
    """遞迴清理 dict/list/str 裡的 XML 非法字元;bytes/數字/None/bool 原樣保留"""
    if isinstance(obj, str):
        return _sanitize_str(obj)
    if isinstance(obj, dict):
        return {k: _sanitize_data(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize_data(v) for v in obj]
    if isinstance(obj, tuple):
        return tuple(_sanitize_data(v) for v in obj)
    return obj


# ==================== 段落操作 ====================
def _replace_paragraph_text(paragraph, new_text: str):
    """替換段落文字,保留第一個 run 的格式;其他 run 會被移除"""
    if not paragraph.runs:
        paragraph.text = new_text
        return
    first_run = paragraph.runs[0]
    first_run.text = new_text
    for run in paragraph.runs[1:]:
        run._r.getparent().remove(run._r)


def _replace_paragraph_text_or_add(paragraph, text: str):
    """替換段落文字(保留格式),若無 run 則新增一個"""
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = text
        for run in paragraph.runs[1:]:
            run._r.getparent().remove(run._r)
    else:
        paragraph.add_run(text)


def _insert_paragraph_after(paragraph, text: str = "", style=None):
    """在指定段落後插入新段落,複製原段落樣式;回傳新 Paragraph"""
    new_p = deepcopy(paragraph._p)
    for child in list(new_p):
        if child.tag == qn('w:r'):
            new_p.remove(child)
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        _replace_paragraph_text_or_add(new_para, text)
    return new_para


def _remove_paragraph(paragraph):
    """從文件中移除指定段落"""
    p = paragraph._p
    p.getparent().remove(p)


# ==================== 儲存格操作 ====================
def _set_cell_text(cell, text: str, bold: bool = False):
    """設定表格儲存格文字,保留原格式"""
    for para in cell.paragraphs:
        if para.runs:
            first_run = para.runs[0]
            first_run.text = text
            if bold:
                first_run.bold = True
            for run in para.runs[1:]:
                run._r.getparent().remove(run._r)
            for extra in cell.paragraphs[1:]:
                _remove_paragraph(extra)
            return
    cell.text = text


# ==================== 日期工具 ====================
def _minus_months(d, months: int):
    """把 date d 往前推 months 個月(日期保留;該月天數不足就取月底)"""
    from calendar import monthrange
    y = d.year
    m = d.month - months
    while m <= 0:
        m += 12
        y -= 1
    max_day = monthrange(y, m)[1]
    return date(y, m, min(d.day, max_day))


# 台灣「見紅就放」假日表(2024~2027),含國定假日、補假、彈性放假
_TW_HOLIDAYS = {
    # 2024
    date(2024, 1, 1),
    date(2024, 2, 8), date(2024, 2, 9),
    date(2024, 2, 10), date(2024, 2, 11), date(2024, 2, 12), date(2024, 2, 13), date(2024, 2, 14),
    date(2024, 2, 28),
    date(2024, 4, 4), date(2024, 4, 5),
    date(2024, 5, 1),
    date(2024, 6, 10),
    date(2024, 9, 17),
    date(2024, 10, 10),
    # 2025
    date(2025, 1, 1),
    date(2025, 1, 27), date(2025, 1, 28), date(2025, 1, 29), date(2025, 1, 30), date(2025, 1, 31),
    date(2025, 2, 28),
    date(2025, 4, 3), date(2025, 4, 4),
    date(2025, 5, 1),
    date(2025, 5, 30), date(2025, 5, 31),
    date(2025, 9, 29),
    date(2025, 10, 6),
    date(2025, 10, 10),
    # 2026
    date(2026, 1, 1),
    date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18), date(2026, 2, 19), date(2026, 2, 20),
    date(2026, 2, 27),
    date(2026, 4, 3), date(2026, 4, 6),
    date(2026, 5, 1),
    date(2026, 6, 19),
    date(2026, 9, 25),
    date(2026, 9, 28),
    date(2026, 10, 9), date(2026, 10, 12),
    # 2027
    date(2027, 1, 1),
    date(2027, 2, 5), date(2027, 2, 8), date(2027, 2, 9), date(2027, 2, 10), date(2027, 2, 11), date(2027, 2, 12),
    date(2027, 3, 1),
    date(2027, 4, 5),
    date(2027, 5, 3),
    date(2027, 6, 9),
    date(2027, 9, 14),
    date(2027, 10, 11),
}


def _count_business_days(start, end) -> int:
    """算工作天數(扣週末 + 台灣國定假日「見紅就放」,含起訖兩天)"""
    if end < start:
        return 0
    days = 0
    current = start
    while current <= end:
        if current.weekday() < 5 and current not in _TW_HOLIDAYS:
            days += 1
        current += timedelta(days=1)
    return days

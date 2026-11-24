"""內文指名度查詢:從「專案發文總覽 PDF」抽文章、關鍵字搜尋、UI 渲染"""
import re
import unicodedata
from datetime import datetime

import pdfplumber


# ==================== PDF 抽取 ====================
def extract_posts_with_content(pdf_path):
    """從「專案發文總覽 PDF」抽出每篇文章的 title.標題 + cnt.內文

    結構特徵:
    - 每篇文章都有「目標站版」「網址紀錄」「cnt.內文」這 3 個必有欄位
    - 「title.標題」可能在「主文」會出現,但「推文」有時會省略

    用「目標站版」當分割錨點(每篇必有),確保抓到所有文章。
    """
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for p in pdf.pages:
            full_text += unicodedata.normalize('NFKC', p.extract_text() or '') + "\n"

    station_matches = list(re.finditer(r'目標站版\s+(.+?)(?=\n)', full_text))

    posts = []
    最近的標題 = ""

    for i, m in enumerate(station_matches):
        start = m.start()
        next_start = station_matches[i + 1].start() if i + 1 < len(station_matches) else len(full_text)
        block = full_text[start:next_start]

        title_match = re.search(r'title\.標題\s+(.+?)(?=\n)', block)
        if title_match:
            title = title_match.group(1).strip().replace('\n', ' ')
            最近的標題 = title
        else:
            title = f"#? RE: {最近的標題}" if 最近的標題 else "(無標題)"

        content_match = re.search(r'cnt\.內文(.*?)(?:\n\s*Time\.發文時間|\n\s*截圖|\Z)', block, re.DOTALL)
        content = content_match.group(1).strip() if content_match else ""

        posts.append({"title": title, "content": content})

    return posts


# ==================== 標題正規化 ====================
def normalize_thread_title(title):
    """把標題去掉 #N RE: 前綴 + 過濾 PDF 圖示字符,得到主討論串名稱"""
    if not title:
        return ""
    cleaned = title
    cleaned = ''.join(ch for ch in cleaned if not (0xE000 <= ord(ch) <= 0xF8FF))
    cleaned = cleaned.strip()
    cleaned = re.sub(r'^#\S*\s*RE:\s*', '', cleaned)
    cleaned = re.sub(r'^RE:\s*', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()


# ==================== 搜尋 ====================
def search_posts(posts, keyword, search_in="content"):
    """依關鍵字搜尋文章

    Args:
        posts: extract_posts_with_content() 的回傳
        keyword: 關鍵字,支援:
            - 單一詞: "Friday"
            - 複合 AND: "Friday+稻草人"
            - 複合 OR:  "Friday,稻草人" 或 "Friday|稻草人"
        search_in:
            - "content": 內文 — 每篇獨立計算
            - "title": 標題 — 同一主討論串視為同一篇

    Returns: list of dict, 每筆含 {title, content, snippet_source, thread_size}
    """
    if not keyword.strip():
        return []

    kw = keyword.strip()
    if '+' in kw:
        parts = [k.strip().lower() for k in kw.split('+') if k.strip()]
        mode = "AND"
    elif ',' in kw or '|' in kw:
        parts = [k.strip().lower() for k in re.split(r'[,|]', kw) if k.strip()]
        mode = "OR"
    else:
        parts = [kw.lower()]
        mode = "AND"

    def match(text):
        t = text.lower()
        return all(k in t for k in parts) if mode == "AND" else any(k in t for k in parts)

    if search_in == "title":
        from collections import OrderedDict
        groups = OrderedDict()
        for p in posts:
            key = normalize_thread_title(p["title"])
            if key not in groups:
                groups[key] = []
            groups[key].append(p)

        matched = []
        for thread_title, group_posts in groups.items():
            if match(thread_title):
                rep = group_posts[0]
                matched.append({
                    "title": thread_title,
                    "content": rep["content"],
                    "snippet_source": thread_title,
                    "thread_size": len(group_posts),
                })
        return matched
    else:
        matched = []
        for p in posts:
            if match(p["content"]):
                matched.append({
                    "title": p["title"],
                    "content": p["content"],
                    "snippet_source": p["content"],
                    "thread_size": 1,
                })
        return matched


# ==================== Streamlit UI ====================
def render_ui(pdf_path):
    """渲染內文指名度查詢功能的完整 UI(從 app.py 呼叫)"""
    import streamlit as st

    with st.spinner("🔍 解析 PDF 文章..."):
        posts = extract_posts_with_content(pdf_path)

    if not posts:
        st.error("❌ 無法從 PDF 抽出文章。請確認上傳的是『專案發文總覽』PDF(內含 cnt.內文 欄位)")
        return

    st.success(f"✅ 解析成功!共抓到 **{len(posts)}** 篇文章")

    total_content_chars = sum(len(p["content"]) for p in posts)
    total_title_chars = sum(len(p["title"]) for p in posts)
    c1, c2, c3 = st.columns(3)
    c1.metric("文章篇數", len(posts))
    c2.metric("內文總字數", f"{total_content_chars:,}")
    c3.metric("標題總字數", f"{total_title_chars:,}")

    st.divider()
    st.subheader("🔎 關鍵字搜尋")
    col_kw, col_scope = st.columns([3, 1])
    with col_kw:
        keyword = st.text_input(
            "輸入關鍵字",
            placeholder="例如:Friday、Friday+稻草人(AND)、Friday,稻草人(OR)",
            key="search_keyword",
        )
    with col_scope:
        search_in = st.radio("搜尋範圍", ["內文", "標題"], key="search_scope")
    search_in_key = "content" if search_in == "內文" else "title"

    st.caption("💡 語法說明:`+` = AND(都要含)、`,` 或 `|` = OR(任一含)、不分大小寫")

    if not keyword.strip():
        st.info("👆 請輸入關鍵字後查詢")
        return

    matched = search_posts(posts, keyword, search_in=search_in_key)

    if search_in_key == "title":
        from collections import OrderedDict
        thread_set = OrderedDict()
        for p in posts:
            key = normalize_thread_title(p["title"])
            thread_set[key] = True
        total_base = len(thread_set)
        base_label = "總討論串數"
    else:
        total_base = len(posts)
        base_label = "總篇數"

    st.divider()
    st.markdown("### 📊 搜尋結果")
    col_r1, col_r2, col_r3 = st.columns(3)
    unit = "串" if search_in_key == "title" else "篇"
    col_r1.metric(f"在「{search_in}」中命中", f"{len(matched)} {unit}")
    col_r2.metric(base_label, f"{total_base} {unit}")
    pct = (len(matched) / total_base * 100) if total_base else 0
    col_r3.metric("命中比例", f"{pct:.1f}%")

    if search_in_key == "title":
        st.caption("💡 標題模式:同一主討論串(如「#1 RE: XXX」「#2 RE: XXX」)合併為 1 串計算")
    else:
        st.caption("💡 內文模式:主文與每篇推文/回應各算 1 篇")

    if not matched:
        st.info("沒有任何文章命中關鍵字")
        return

    label_word = "討論串" if search_in_key == "title" else "文章"
    st.subheader(f"📋 命中{label_word}清單({len(matched)} {unit})")
    import pandas as pd
    rows = []
    for i, p in enumerate(matched, 1):
        target = p["snippet_source"]
        first_kw = re.split(r'[+,|]', keyword.strip())[0].strip()
        idx = target.lower().find(first_kw.lower())
        if idx >= 0:
            ctx_start = max(0, idx - 20)
            ctx_end = min(len(target), idx + len(first_kw) + 40)
            snippet = ("..." if ctx_start > 0 else "") + target[ctx_start:ctx_end].replace('\n', ' ') + ("..." if ctx_end < len(target) else "")
        else:
            snippet = target[:60].replace('\n', ' ') + ("..." if len(target) > 60 else "")

        row = {
            "#": i,
            "標題": p["title"][:50],
            f"{search_in}命中片段": snippet,
        }
        if search_in_key == "title" and p.get("thread_size", 1) > 1:
            row["該串文章數"] = p["thread_size"]
        rows.append(row)
    df_results = pd.DataFrame(rows)
    st.dataframe(df_results, use_container_width=True, hide_index=True)

    csv_data = df_results.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        "📥 下載命中清單(CSV)",
        data=csv_data,
        file_name=f"內文指名度_{keyword.replace('+','_').replace(',','_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
    )

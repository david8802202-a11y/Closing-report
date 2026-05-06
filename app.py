import streamlit as st
import pandas as pd
import io
import pdfplumber

st.set_page_config(page_title="口碑結案自動生成器", page_icon="📊")

st.title("📊 口碑結案報告自動生成器")
st.markdown("上傳 Beauty2 匯出的 PDF 與結案表格模板，自動將數據填入模板中。")

st.write("---")

# 1. 檔案上傳區
col1, col2 = st.columns(2)
with col1:
    uploaded_pdf = st.file_uploader("1. 上傳網頁 PDF (Beauty2)", type=["pdf"])
with col2:
    uploaded_template = st.file_uploader("2. 上傳結案模板 (.xlsx)", type=["xlsx", "xls"])

if st.button("🚀 開始分析與生成", type="primary"):
    if uploaded_pdf is None or uploaded_template is None:
        st.error("請確認 PDF 和 Excel 模板都已上傳！")
    else:
        with st.spinner("資料處理中，請稍候..."):
            try:
                # ==========================================
                # 步驟 A: 解析 PDF (這裡先做簡單的示範讀取)
                # ==========================================
                st.info("正在解析 PDF 內容...")
                pdf_text_summary = []
                with pdfplumber.open(uploaded_pdf) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            pdf_text_summary.append(f"第 {i+1} 頁讀取成功。")
                
                # 這裡未來會放入將 text 轉換為結構化數據的邏輯
                
                # ==========================================
                # 步驟 B: 處理 Excel 模板
                # ==========================================
                st.info("正在讀取結案表格模板...")
                # 讀取模板以確認格式
                df_template = pd.read_excel(uploaded_template, sheet_name=None)
                sheet_names = list(df_template.keys())
                
                # 這裡未來會放入將 PDF 數據填入 Excel 的邏輯
                
                # ==========================================
                # 步驟 C: 準備下載檔案 (目前先回傳原模板示範)
                # ==========================================
                st.success("✅ 處理完成！")
                st.write(f"已識別模板分頁：{', '.join(sheet_names)}")
                
            # 取得原檔案的名稱與格式，避免副檔名衝突
                original_filename = uploaded_template.name

            # 將處理好的檔案轉為可下載格式
                output = io.BytesIO()
                uploaded_template.seek(0)
                output.write(uploaded_template.read())

# 讓下載出來的檔案，動態跟隨原本的副檔名
                st.download_button(
                    label="📥 下載完成的結案報告",
                    data=output.getvalue(),
                    file_name=f"自動產出_{original_filename}", 
                    mime="application/octet-stream"
                )
            except Exception as e:
                st.error(f"處理過程中發生錯誤：{e}")

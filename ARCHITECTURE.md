# 檔案架構說明(重構後)

## 檔案清單與職責

```
app.py                36 KB  主入口 + 共用 PDF 解析 + Streamlit UI 路由
closure_report.py     2.5 MB 結案:Excel + Word + 結案指令詞 + 預覽
monthly_report.py     2.4 MB 月報:Excel + Word + 簡報指令詞 + 預覽
indexing_search.py    12 KB  內文指名度查詢(獨立)
docx_utils.py         8 KB   共用 Word 工具(清字元、插段落、填儲存格、工作天數)
requirements.txt
README.md
```

## 誰負責什麼

| 要改的東西 | 改哪個檔 |
|-----------|---------|
| PDF 主表抽取邏輯、UI 版面、新增功能分支 | `app.py` |
| 結案 Excel 填法、結案 KPI 計算、結案 Word 範本、結案指令詞 | `closure_report.py` |
| 月報 Excel 填法、月報 Word 範本、月報簡報指令詞 | `monthly_report.py` |
| 內文指名度搜尋演算法、關鍵字 UI | `indexing_search.py` |
| XML 字元清理、段落操作、工作天數算法、台灣假日 | `docx_utils.py` |

## 相依關係

```
app.py
├── closure_report
│   └── docx_utils
├── monthly_report
│   └── docx_utils
└── indexing_search
```

`docx_utils` 不相依任何 app 模組,可獨立測試。

## Word 範本在哪

**內嵌在對應模組**:
- 結案 Excel 範本 → `closure_report.py` 內 `CLOSURE_TEMPLATE_B64`
- 結案 Word 範本 → `closure_report.py` 內 `WORD_TEMPLATE_B64`
- 結案指令詞範本 → `closure_report.py` 內 `PROMPT_TEMPLATE_B64`
- 月報 Excel 範本 → `monthly_report.py` 內 `MONTHLY_TEMPLATE_B64`
- 月報 Word 範本 → `monthly_report.py` 內 `MONTHLY_WORD_TEMPLATE_B64`
- 月報簡報指令詞範本 → `monthly_report.py` 內 `MONTHLY_PROMPT_TEMPLATE_B64`

## 之後要新增第 3、第 4 種報表

照 `monthly_report.py` 的模板複製一個新模組(例如 `season_report.py`),然後:
1. 把新的範本 b64 字串塞進去
2. 寫 `fill_template()`、`build_word_report()`、`build_prompt_text()`、`render_form()`、`_render_preview()`
3. 在 `app.py` 的 radio 加一個選項,引用新模組
4. 完成

跟現有的結案/月報完全不會互相干擾。

## 已經可以刪掉的舊檔

- `app_g.py`(舊的簡化版備份,不再使用)
- `word_report.py`(內容已搬進 `closure_report.py`)

## 部署步驟

把下列 5 個檔案覆蓋/新增到 repo 根目錄,刪掉 `app_g.py` 和 `word_report.py`:

```
app.py
closure_report.py
monthly_report.py
indexing_search.py
docx_utils.py
```

commit → push → Streamlit Cloud 右上角 ⋮ → **Reboot app** 強制重新載入所有 module

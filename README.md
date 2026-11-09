# WORD清洗MD (word-cleaner-md)

將 Word 文件（`.docx`）進行結構化解析、文本清洗並轉換為標準 Markdown 格式的自動化工具。

---

## 專案特色與目標

- **Word 結構解析**：自動識別標題層級、段落、列表、表格與圖檔引用。
- **樣式清洗與標準化**：去除冗餘樣式、修復非標準換行、統一標點符號與編碼。
- **乾淨 Markdown 輸出**：產出語義清晰、相容性高的 GitHub Flavored Markdown (GFM)。

---

## 目錄結構規劃（預期）

```text
WORD清洗MD/
├── .gitignore          # Git 忽略設定
├── README.md           # 專案說明文件
├── requirements.txt    # 相依套件清單
├── input/              # 待處理 Word 來源檔（預設已忽略）
├── output/             # 清洗完成之 Markdown 輸出檔（預設已忽略）
└── src/                # 核心原始碼
```

---

## 快速開始

### 1. 環境需求
- Python 3.9+
- Git

### 2. 安裝虛擬環境與相依套件（規劃中）
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# 安裝依賴
pip install -r requirements.txt
```

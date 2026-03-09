
# 📦 ERP 訂購系統自動化工具

本專案是一個使用 **Python + Selenium** 開發的 ERP / B2B 訂購系統自動化工具。

透過讀取 Excel 訂單資料，自動登入訂購平台，搜尋商品並加入購物車，
可大幅減少人工逐筆下單的時間，提升企業採購與訂單處理效率。

---

# 🚀 專案功能

本系統提供以下自動化功能：

- 🔐 自動登入 ERP / B2B 訂購系統
- 📊 從 Excel 批次讀取訂單資料
- 🔎 自動搜尋商品
- 🛒 自動加入購物車
- ❌ 記錄失敗訂單並輸出 Excel
- ⚡ 完整自動化下單流程

---

# 🧠 技術亮點

- Python 自動化腳本開發
- Selenium 網頁自動化
- Excel 資料處理 (Pandas)
- Web Element 定位與操作
- Alert 視窗處理
- 批次資料處理流程設計

---

# 🏗 系統流程

Excel 訂單資料  
      │  
      ▼  
讀取 Excel (Pandas)  
      │  
      ▼  
登入 ERP 系統  
      │  
      ▼  
搜尋商品  
      │  
      ▼  
加入購物車  
      │  
      ▼  
輸出失敗訂單  

---

# 🛠 技術架構

| 技術 | 說明 |
|-----|-----|
| Python | 自動化程式開發 |
| Selenium | 網頁操作自動化 |
| Pandas | Excel 資料處理 |
| ChromeDriver | 瀏覽器控制 |

---

# 📂 專案結構

crerp
│
├── 文具.py           # 主程式（自動化下單）
├── product.xlsx      # 訂單資料來源
├── 錯誤清單.xlsx      # 失敗訂單輸出
└── README.md

---

# 📄 程式說明

## 文具.py

主要功能：

- 初始化 Selenium 瀏覽器
- 自動登入訂購系統
- 進入訂購頁面
- 搜尋商品
- 加入購物車
- 從 Excel 批次處理訂單
- 匯出失敗訂單

---

# ⚙️ 安裝方式

git clone https://github.com/leolin0330/crerp.git

安裝套件

pip install selenium pandas webdriver-manager

---

# ▶️ 執行方式

python 文具.py

---

# 📈 專案用途

- ERP / B2B 系統自動化下單
- 批次訂單處理
- 企業採購流程自動化
- 減少人工操作時間

---

# 👨‍💻 作者

Timmy Lin

ERP 系統工程師

技能：

- ERP 系統維護
- Python 自動化
- SQL / 資料分析
- AI 應用開發

---

# ⭐ GitHub

https://github.com/leolin0330/crerp

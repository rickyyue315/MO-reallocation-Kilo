# 庫存調貨建議系統(澳門優先版) v1.0

## 系統概述

庫存調貨建議系統(澳門優先版)v1.0是一個基於Streamlit的Web應用程序，旨在根據庫存、銷量和安全庫存數據，自動生成跨店鋪的商品調貨建議，以優化庫存分布，滿足銷售需求。

本版本支持**三模式系統**：A(保守轉貨)/B(加強轉貨)/C(全量轉貨)，並特別針對澳門地區的庫存管理需求進行了優化。

## 功能特點

- **三模式系統**：支持A(保守轉貨)/B(加強轉貨)/C(全量轉貨)的轉貨模式
- **智能調貨算法**：基於優先級的轉出/接收規則，實現庫存優化分配
- **數據預處理與驗證**：自動讀取Excel文件，處理數據類型轉換、缺失值和異常值
- **質量檢查**：確保調貨建議的準確性和合理性
- **可視化結果**：提供調貨建議詳情和統計圖表
- **Excel導出**：生成包含調貨建議和統計摘要的Excel文件
- **模擬數據生成**：內置測試數據生成器，方便系統測試

## 系統要求

- Python 3.8+
- 依賴包：pandas, openpyxl, streamlit, numpy, xlsxwriter, matplotlib, seaborn

## 快速開始

### 1. 克隆或下載項目

```bash
git clone <repository-url>
cd inventory-transfer-system
```

### 2. 安裝依賴

#### 方法一：使用自動安裝腳本（推薦）

```bash
# Windows
python install_dependencies.py

# Linux/macOS
python3 install_dependencies.py
```

#### 方法二：手動安裝

```bash
pip install pandas openpyxl streamlit numpy xlsxwriter matplotlib seaborn
```

### 3. 運行系統

#### 方法一：使用運行腳本（推薦）

```bash
# Windows
run.bat

# Linux/macOS
chmod +x run.sh
./run.sh
```

#### 方法二：直接運行

```bash
streamlit run app.py
```

### 4. 訪問系統

打開瀏覽器，訪問 `http://localhost:8501`

## 使用說明

### 1. 上傳數據文件

- 點擊"瀏覽文件"按鈕上傳Excel文件（.xlsx格式）
- 確保文件包含以下必需欄位：
  - Article（商品編號，12位文本格式）
  - Article Description（商品描述）或 Article Long Text (60 Chars)
  - OM（OM編號）
  - RP Type（店鋪類型：ND或RF）
  - Site（店鋪代碼）
  - MOQ（最低派貨數量）
  - SaSa Net Stock（淨庫存）
  - Pending Received（待收貨數量）
  - Safety Stock（安全庫存）
  - Last Month Sold Qty（上月銷量）
  - MTD Sold Qty（本月銷量）

### 2. 選擇模式

選擇轉貨模式：

**轉貨模式：**
- **A模式(保守轉貨)**：轉出後剩餘庫存不低於安全庫存，轉出類型為RF過剩轉出
- **B模式(加強轉貨)**：轉出後剩餘庫存可能低於安全庫存，轉出類型為RF加強轉出
- **C模式(全量轉貨)**：忽視A模式及B模式的限制，ND Shop可以轉去ND Shop，需要限制同一個OM組別及同一個Article，轉出店舖的銷售量必須為同組最少，接收店舖的銷售量必須為同組最多，轉出店舖的銷售量如果為0件，轉出數量可全數轉出

### 3. 數據預處理

- 系統自動讀取並驗證數據
- 執行數據類型轉換、缺失值處理和異常值校正
- 顯示處理結果和數據統計

### 4. 生成調貨建議

- 點擊"生成調貨建議"按鈕
- 系統基於選擇的模式生成調貨建議
- 執行質量檢查，確保建議的準確性

### 5. 查看結果

- 查看調貨建議統計和詳情
- 瀏覽統計圖表，了解轉出類型和接收優先級分布

### 6. 下載結果

- 點擊"下載調貨建議Excel文件"按鈕
- 獲取包含以下工作表的Excel文件：
  - 調貨建議（Transfer Recommendations）
  - 統計摘要（Summary Dashboard）

## 實際Excel欄位結構（Transfer Recommendations）

1. Article (A欄, 12字符)
2. Product Desc (B欄, 25字符)
3. Transfer OM (C欄, 12字符)
4. Transfer Site (D欄, 12字符)
5. Receive OM (E欄, 12字符)
6. Receive Site (F欄, 12字符)
7. Transfer Qty (G欄, 10字符)
8. Transfer Site Original Stock (H欄, 18字符)
9. Transfer Site After Transfer Stock (I欄, 20字符)
10. Transfer Site Safety Stock (J欄, 18字符)
11. Transfer Site MOQ (K欄, 12字符)
12. Transfer Site Last Month Sold Qty (M欄, 18字符)
13. Transfer Site MTD Sold Qty (N欄, 15字符)
14. Receive Site Last Month Sold Qty (O欄, 18字符)
15. Receive Site MTD Sold Qty (P欄, 15字符)
16. Receive Original Stock (Q欄, 15字符)
17. Remark (L欄, 25字符)
18. Notes (R欄, 200字符) ← 寬度優化

## 業務邏輯詳解

### 有效銷量計算

為了簡化銷量比較，系統定義"有效銷量"欄位：
- 使用"上月銷量" 及 "本月銷量"

### A模式(保守轉貨) - 轉出規則

**優先級1：ND類型轉出**
- 條件：RP Type為"ND"
- 可轉數量：全部SaSa Net Stock
- 轉出類型：ND轉出

**優先級2：RF類型過剩轉出**
- 條件1：RP Type為"RF"
- 條件2：庫存充足（SaSa Net Stock + Pending Received > Safety Stock）
- 條件3：該店鋪的有效銷量不是此Article中的最高值
- 條件4：轉出後剩餘庫存不低於安全庫存
- 基礎可轉出 = (庫存+在途) - 安全庫存
- 上限控制 = (庫存+在途) × 40%，但最少出貨2件
- 實際轉出 = min(基礎可轉出, max(上限控制, 2))
- 轉出類型：RF過剩轉出

### B模式(加強轉貨) - 轉出規則

**優先級1：ND類型轉出**
- 條件：RP Type為"ND"
- 可轉數量：全部SaSa Net Stock
- 轉出類型：ND轉出

**優先級2：RF類型轉出**
- 條件1：RP Type為"RF"
- 條件2：(庫存+在途) > (MOQ數量)
- 條件3：該店鋪的有效銷量不是此Article中的最高值
- 基礎可轉出 = (庫存+在途) – (MOQ數量)
- 上限控制 = (庫存+在途) × 80%，但最少出貨2件
- 實際轉出 = min(基礎可轉出, max(上限控制, 2))
- 轉出類型判斷：
  - 如果轉出後剩餘庫存 ≥ 安全庫存：RF過剩轉出
  - 如果轉出後剩餘庫存 < 安全庫存：RF加強轉出

### C模式(全量轉貨) - 轉出規則

**優先級1：ND類型轉出**
- 條件1：RP Type為"ND"
- 條件2：轉出店舖的銷售量為同OM同Article中最少
- 可轉數量：全部SaSa Net Stock
- 轉出類型：ND轉出
- 限制：只能轉給同OM的ND店鋪

**優先級2：RF類型轉出**
- 條件1：RP Type為"RF"
- 條件2：轉出店舖的銷售量為同OM同Article中最少
- 條件3：轉出店舖的銷售量為0件
- 可轉數量：全部SaSa Net Stock
- 轉出類型：C模式全量轉出
- 限制：只能轉給同OM的RF店鋪

### 接收規則（三種模式通用）

**優先級1：緊急缺貨補貨**
- 條件1：RP Type為"RF"
- 條件2：完全無庫存+在途（SaSa Net Stock + Pending Received = 0）
- 條件3：曾有銷售記錄（Effective Sold Qty > 0）
- 需求數量：Safety Stock

**優先級2：潛在缺貨補貨**
- 條件1：RP Type為"RF"
- 條件2：庫存不足（SaSa Net Stock + Pending Received < Safety Stock）
- 條件3：該店鋪的有效銷量是此Article中的最高值
- 需求數量：Safety Stock - (SaSa Net Stock + Pending Received)

### 匹配算法

1. 按優先級順序進行匹配：
   - ND轉出 → 緊急缺貨
   - ND轉出 → 潛在缺貨
   - RF過剩轉出 → 緊急缺貨
   - RF過剩轉出 → 潛在缺貨
   - RF加強轉出 → 緊急缺貨
   - RF加強轉出 → 潛在缺貨
   - C模式全量轉出 → 緊急缺貨
   - C模式全量轉出 → 潛在缺貨
2. 每次匹配的轉移數量為min(轉出方的可轉數量, 接收方的需求數量)
3. 更新雙方的剩餘可轉數量和需求數量
4. 調貨數量優化：如果只有1件，嘗試調高到2件（前提是不影響轉出RF店鋪的庫存）
5. HA店舖,H店舖,HC店舖可以出貨去HD店舖, 但HD店舖絕對不能出貨去HA店舖,H店舖,HC店舖
6. HA店舖,H店舖,HC店舖轉去HA店舖,H店舖,HC店舖需限制同組OM, 但轉去HD店舖不受OM限制

## 質量檢查

系統自動執行以下質量檢查：

- [ ] 轉出與接收的Article必須完全一致
- [ ] Transfer Qty必須為正整數
- [ ] Transfer Qty不得超過轉出店鋪的原始SaSa Net Stock
- [ ] Transfer Site和Receive Site不能相同
- [ ] 最終輸出的Article欄位必須是12位文本格式

## 故障排除

### 常見問題

1. **文件上傳失敗**
   - 確保文件格式為.xlsx
   - 檢查文件是否損壞
   - 確認文件大小在允許範圍內

2. **數據處理錯誤**
   - 檢查文件是否包含所有必需欄位
   - 確認Article欄位為12位文本格式
   - 檢查數據中是否有特殊字符

3. **沒有生成調貨建議**
   - 確認數據中同時包含ND和RF類型的店鋪
   - 檢查庫存和銷量數據是否合理
   - 確認安全庫存和MOQ設置是否正確

4. **模式選擇錯誤**
   - 確認選擇了正確的轉貨模式
   - 系統支持A模式(保守轉貨)、B模式(加強轉貨)和C模式(全量轉貨)

5. **依賴安裝失敗**
   - 嘗試使用專用安裝腳本：`python install_dependencies.py`
   - 手動安裝核心依賴：`pip install pandas openpyxl streamlit numpy xlsxwriter matplotlib seaborn`
   - 升級pip：`python -m pip install --upgrade pip`
   - 使用國內鏡像源：`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple/`

### 日誌查看

應用程序運行時會生成日誌，可在控制台查看詳細錯誤信息。

## 技術架構

```
inventory_transfer_system/
├── app.py                 # Streamlit主應用 v1.9
├── data_processor.py      # 數據預處理模組 v1.8.1
├── business_logic.py      # 業務邏輯模組 v1.9
├── excel_generator.py     # Excel輸出模組 v1.8
├── requirements.txt       # 依賴包列表
├── install_dependencies.py # 依賴安裝腳本
├── README.md             # 使用說明
├── VERSION.md            # 版本更新記錄
├── run.bat               # Windows運行腳本
├── run.sh                # Linux/macOS運行腳本
├── test_system.py        # 系統測試腳本
├── test_chart_v1.8.py    # 圖表測試腳本
└── test_real_data.py     # 真實數據測試腳本
```

## 測試

### 運行系統測試

```bash
# 運行完整系統測試
python test_system.py
```

### 測試覆蓋範圍

- 數據處理模塊測試
- 業務邏輯模塊測試
- Excel生成模塊測試
- 集成測試
- 邊界情況測試

## 版本歷史

詳細的版本更新記錄請參見 [VERSION.md](VERSION.md)

## 貢獻

歡迎提交問題報告和功能請求。

## 許可證

本項目採用 MIT 許可證。

## 聯繫方式

如有問題或建議，請通過以下方式聯繫：
- 郵箱：[您的郵箱]
- 問題追踪：[GitHub Issues鏈接]

---

**庫存調貨建議系統(澳門優先版) v1.0** - 優化庫存管理，提升銷售效率
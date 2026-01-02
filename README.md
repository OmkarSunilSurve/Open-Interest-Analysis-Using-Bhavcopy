# 📊 Derivative Positioning & Open Interest Analysis (Excel VBA)

A complete **Excel VBA–based analytics framework** for analyzing **derivatives market positioning** using **price action and Open Interest (OI)** from NSE FO Bhavcopy data.

This project automates:

* Bhavcopy download & cleaning
* Futures & underlying aggregation
* Open Interest buildup classification
* Period-based positioning (1D / 5D / Since Date)
* Professional tables & charts for decision support

Designed for **derivatives traders, market analysts, and quant enthusiasts** who want structured, repeatable analysis inside Excel.

---

## 🧠 Core Concepts Covered

* Open Interest analysis
* Long Buildup / Short Buildup / Short Covering / Long Unwinding
* Futures vs underlying price behavior
* Period-based positioning analysis
* Data-driven market structure understanding

---

## 🗂️ Project Architecture

### 🔹 Entry Points (Macros / Buttons)

| Macro                | Purpose                                              |
| -------------------- | ---------------------------------------------------- |
| `Function_Call`      | End-to-end pipeline: download → process → dashboard  |
| `CreateDropdown`     | Creates stock selection dropdown                     |
| `Update_Stock_Chart` | Refreshes charts & period buildup for selected stock |

---

## 🔁 Execution Flow (High Level)

```
Function_Call
  ├─ Download & Unzip Bhavcopy
  ├─ Clean & Append Data
  ├─ Create Futures / OI / Underlying Sheets
  ├─ Pivot & Aggregate Data
  ├─ Calculate % Changes
  ├─ Generate Buildup Tables
  ├─ Sort & Format Tables
  └─ Update Dashboard
```

```
Update_Stock_Chart
  ├─ Fetch Last N Trading Days
  ├─ Prepare Chart Data
  ├─ Create Price vs OI Charts
  ├─ Create Period Buildup Table
  └─ Rearrange Sheets & Return Focus
```

---

## 📁 Sheet Structure

| Sheet             | Description                               |
| ----------------- | ----------------------------------------- |
| `Data`            | Raw appended bhavcopy data                |
| `Sheet_2 / 3 / 4` | Futures / Underlying / OI processed views |
| `chart_1_Data`    | Price vs OI chart source                  |
| `chart_2_Data`    | Basis points chart source                 |
| `Period_Buildup`  | 1D / 5D / Since Date positioning          |
| `Macro Control`   | User control panel                        |

---

## 📈 Buildup Classification Logic

| Price Change | OI Change | Classification |
| ------------ | --------- | -------------- |
| ↑            | ↑         | Long Buildup   |
| ↓            | ↑         | Short Buildup  |
| ↑            | ↓         | Short Covering |
| ↓            | ↓         | Long Unwinding |

---

## ⏱ Period Buildup Horizons

* **1 Day** → Latest vs previous trading day
* **5 Day** → Latest vs 5th previous trading day
* **Since Date** → Latest vs user-selected reference date

Computed dynamically using trading-day logic (weekends & holidays excluded).

---

## 🎨 Formatting & UX Features

* Auto-generated tables with borders
* Merged & centered titles
* Header shading
* Percentage formatting
* Dynamic column sizing
* Sheet auto-reordering
* Cursor focus returned to control sheet

---

## 🛠 Key VBA Modules & Utilities

* `DownloadLast20TradingDays`
* `GetLastNthTradingDay`
* `Create_Period_Buildup`
* `ApplyTableBorder`
* `CreatePrice_OI_ComboChart`
* `GroupAndSum`
* `KeepNearestExpiry`
* `RearrangeSheets_UsingArray`

---

## ⚙️ Requirements

* Microsoft Excel (Windows)
* VBA enabled (.xlsm)
* Internet access (for NSE Bhavcopy download)

---

## 🚀 How to Use

1. Open the Excel file
2. Enable macros
3. Go to **Macro Control** sheet
4. Select stock & parameters
5. Click **Run / Update** buttons
6. Review tables & charts

---

## ⚠️ Notes & Assumptions

* NSE Bhavcopy URL structure assumed stable
* Only futures contracts are processed
* Nearest expiry logic applied
* Designed for analysis, not live trading execution

---

## 📌 Future Enhancements

* Holiday calendar integration
* Sector-level aggregation
* Signal scoring system
* Export to CSV / PDF
* Performance optimization for large datasets
* Migration to Python / Power BI

---

## 👤 Author

**Omkar Surve**
Derivatives & Market Structure Analyst

Focused on **Open Interest analytics, market positioning, and automation**.

---

## 📜 License

This project is for **educational and analytical purposes**.
Use at your own discretion.

---

⭐ If you find this useful, consider starring the repository!

# 📊 Excel VBA Project: Report Loop – Consolidate & Automate

This project brings together multiple VBA techniques into one cohesive solution. 
It automates the **cleanup, formatting, calculation**, and **consolidation** of data from multiple worksheets 
into a single, clean **Yearly Report**.

---

## 📝 Project Overview

The workbook contains:
- **Four regional worksheets** with raw, unformatted data (missing headers and calculations)
- One empty worksheet titled **Yearly Report**

The goal is to:
1. Add and format headers on each regional worksheet
2. Insert a **SUM calculation** in the Total Expense column (Column F)
3. Copy the cleaned data from all sheets
4. Paste the combined results into the **Yearly Report** worksheet

All of this is triggered through a single macro procedure: `FinalReportLoop`.

---

## ⚙️ Key Features

- 🔁 **Loop through all worksheets** dynamically
- 📌 **Call reusable procedures** (like `AddHeaders`, `FormatHeaders`, `AutomateSum`)
- 📋 **Copy and paste** data between sheets
- ↔️ **Move between worksheets**
- 📐 **Offset to find the right row** to paste new data in the Yearly Report
- ➕ **Apply calculations** to each dataset before consolidation

---

## 📁 Files Included

- `final-report-generator.xlsm`: Macro-enabled workbook with all VBA code and worksheets
- `README.md`: This documentation


---

## 🚀 How to Use

1. Open `final-report-generator.xlsm` in Excel
2. Enable macros when prompted
3. Go to the **Developer tab** → **Macros**
4. Select `FinalReportLoop` → Click **Run**
5. Sit back while Excel:
   - Cleans up each worksheet
   - Adds calculations
   - Consolidates all data into **Yearly Report**

---

## 🧠 Skills Demonstrated

- Working with **loops, variables, and procedures** in VBA
- **Copy-paste automation** using Excel VBA
- Cross-sheet **data consolidation**
- **Dynamic cell referencing** with offset and last-row logic
- Modular programming using **procedure calls**

---

## 🖼️ Demo
This Demo shows the formatting in the 4th worksheet. The total sum is generated and the data is copied to the 
'Yearly report' worksheet.
In the 'Yearly Report' worksheet, column(C:F) is adjusted to autofit the values 
(https://github.com/DD-Dhivya/Data_Analytics_Portfolio/blob/main/Excel_projects/report_loop/Reportloop_demo.gif)

```markdown


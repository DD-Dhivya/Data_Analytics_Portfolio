# 📊 Excel Automation: Auto-Sum Total Expenses Across Worksheets

This project demonstrates how to automate the process of inserting a **SUM function** in the `Total Expense` column across multiple worksheets in an Excel workbook using VBA.

## 📝 Project Overview

The workbook contains **four worksheets**, each with a column (Column **F**) dedicated to `Total Expense`. However, the number of rows varies from sheet to sheet. This macro automatically:

- Detects the **last row** in Column F of each worksheet
- Inserts the **SUM formula** directly below the last value
- Loops through **all worksheets** in the workbook
- Executes with a **single macro** triggered from the **Developer tab**

## ⚙️ Features

- Dynamically identifies the last filled row on each sheet
- Uses **variables** to store cell references and build formulas
- Applies the **SUM formula** automatically without manual input
- Efficiently loops through all worksheets
- Activated via a button press or by running the macro `AutoAutomateSum`

## 🧠 Skills Demonstrated

- Working with **VBA loops and variables**
- Navigating through multiple worksheets
- Dynamically referencing cell ranges
- Automating formula insertion

## 🚀 How to Use

1. Open the `auto-sum-expenses.xlsm` file in Excel.
2. Make sure macros are enabled.
3. Go to the **Developer tab** → **Macros** → select `AutoAutomateSum` → click **Run**.
4. The macro will loop through each worksheet and insert the SUM formula in Column F.

## 📁 Files Included

- `auto-sum-expenses.xlsm`: The macro-enabled workbook with automation code
- `README.md`: Project documentation


## 🔧 Tools Used

- Microsoft Excel
- Visual Basic for Applications (VBA)

## 🖼️ Demo

Here is the quick preview of automation in action
![Auto sum Demo](Excel_projects/Automatesum/Automatesum_demo.gif)

## 🧾 VBA Code Preview

The macro used to automate the SUM function looks like this:

![VBA Code Screenshot](https://github.com/DD-Dhivya/Data_Analytics_Portfolio/blob/main/Excel_projects/Automatesum/Automatesum_vbacode.png))

---

✨ This is a simple yet powerful example of how a little bit of VBA can save time and reduce manual effort across repetitive tasks!

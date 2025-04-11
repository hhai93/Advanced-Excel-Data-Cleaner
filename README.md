# 🧹 Advanced Excel Data Cleaner

![GitHub](https://img.shields.io/badge/license-MIT-blue.svg) ![GitHub last commit](https://img.shields.io/github/last-commit/hhai93/Advanced-Excel-Data-Cleaner)

A VBA script for cleaning and validating Excel data, featuring blank row removal, duplicate elimination, text standardization, email validation, and custom numeric checks, all via an intuitive UserForm.

---

## ✨ Features
- 🗑️ Removes blank rows and duplicate entries.
- ✂️ Standardizes text by trimming spaces and converting to uppercase.
- 📧 Validates email formats and highlights errors.
- 📈 Checks numeric values against custom thresholds.
- 🎨 User-friendly interface for selecting cleaning tasks.
- 📊 Detailed summary of all actions performed.

## 📋 Prerequisites
- 🖥️ Microsoft Excel (2010 or later) with VBA enabled.
- 📊 An Excel file with data to clean (`.xlsx` or `.xls`).

---

## 🚀 How to Use

### 1. Prepare Your Excel File
- Ensure data starts at cell A1 with headers in row 1.
- Example:
  | Name   | Email          | Sales  | Date       |
  |--------|----------------|--------|------------|
  | John   | john@doe.com   | 500    | 01-02-2023 |
  |        | alice          | 1000000| 01/02/2023 |
  | John   | john@doe.com   | 500    | 01-02-2023 |

### 2. Add the VBA Script
- Open your Excel file.
- Press `Alt + F11` to open the VBA editor.
- Go to **File** > **Import File** and import `AdvancedDataCleanerForm.frm`.
- Insert a new module and paste the code from `AdvancedDataCleaner.vba`

### 3. Run the Script
- Run `ShowAdvancedDataCleaner` to open the UserForm.
- Select cleaning tasks (e.g., remove blank rows, validate emails).
- Specify a maximum value and column for numeric checks (optional).
- Click **Run** to clean your data.
- 🎉 Review the detailed summary report!

---

## 🛠️ Code Explanation
- **`AdvancedDataCleanerForm`**: UserForm for interactive task selection.
- **`RemoveDuplicates`**: Deletes duplicate rows based on specified columns.
- **`Trim & UCase`**: Standardizes text across all cells.
- **`Email Validation`**: Highlights emails missing "@" or ".".
- **`Numeric Check`**: Flags values exceeding a user-defined threshold.
- 🔄 Comprehensive reporting of all cleaning actions.

---

## ⚠️ Notes
- 💾 Always back up your data before running the script.
- 📬 Email validation targets columns with "Email" in the header (case-insensitive).
- 🔍 Customize validation rules in the code for specific needs (e.g., phone numbers, dates).
- 🖌️ UserForm requires importing `.frm` file—see [VBA documentation](https://docs.microsoft.com/en-us/office/vba) for guidance.

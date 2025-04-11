# 🧹 Advanced Excel Data Cleaner

![GitHub](https://img.shields.io/badge/license-MIT-blue.svg) ![GitHub last commit](https://img.shields.io/github/last-commit/hhai93/Advanced-Excel-Data-Cleaner)

A powerful VBA script for cleaning and validating Excel data, featuring blank row removal, duplicate elimination, text standardization, custom regex validation, reference-based checks, undo functionality, and more, all through an intuitive UserForm.

---

## ✨ Features
- 🗑️ Removes blank rows and duplicate entries.
- ✂️ Standardizes text by trimming spaces and converting to uppercase.
- 📧 Validates email formats and highlights errors.
- 📱 Supports regex-like validation with predefined (phone, postal) and custom patterns.
- 🔍 Checks data against a reference list for automated input validation.
- 🔄 Undo functionality to revert changes.
- 📈 Validates numeric values against custom thresholds.
- 🎨 User-friendly interface for selecting tasks.
- 📊 Detailed summary of all actions.

## 📋 Prerequisites
- 🖥️ Microsoft Excel (2010 or later) with VBA enabled.
- 📊 An Excel file with data to clean (`.xlsx` or `.xls`).
- (Optional) A "Reference" sheet with valid data (e.g., product codes in column A).

---

## 🚀 How to Use

### 1. Prepare Your Excel File
- Ensure data starts at cell A1 with headers in row 1.
- Example:
  | Name   | Email          | Sales  | Phone        | Product  |
  |--------|----------------|--------|--------------|----------|
  | John   | john@doe.com   | 500    | 0912345678   | PROD001  |
  |        | alice          | 1000000| 123456789    | prod002  |
  | John   | john@doe.com   | 500    | EMP1234      | PROD001  |

- (Optional) Create a "Reference" sheet:
  | Valid Product |
  |---------------|
  | PROD001       |
  | PROD002       |

### 2. Add the VBA Script
- Open your Excel file and press `Alt + F11` to open the VBA editor.
- **Add UserForm**:
  - Insert a new UserForm named `AdvancedDataCleanerForm`.
  - Add controls as described in [`AdvancedDataCleanerForm.vb`](AdvancedDataCleanerForm.vb) under "UserForm Layout".
  - Copy and paste the code from "UserForm Code" section into the UserForm's code window.
- **Add Module**:
  - Insert a new module and paste the code from [`ShowAdvancedDataCleaner.vba`](ShowAdvancedDataCleaner.vba).
- Save the Excel file as `.xlsm` (macro-enabled).

### 3. Run the Script
- Press `Alt + F8`, select `ShowAdvancedDataCleaner`, and run.
- In the UserForm:
  - Select cleaning tasks (e.g., remove blank rows, validate emails, standardize text).
  - Specify:
    - Maximum value and column for numeric checks.
    - Column and regex pattern (e.g., `09########` for phone, `EMP####` for custom).
    - Column to check against the "Reference" sheet.
  - Click **Run** to clean your data.
  - Use **Undo** to revert changes if needed.
- 🎉 Review the detailed summary report!

---

## 🛠️ Code Explanation
- **`AdvancedDataCleanerForm.vb`**: Defines the UserForm layout and logic for interactive task selection.
- **`ShowAdvancedDataCleaner.vba`**: Simple module to launch the UserForm.
- **Functionality**:
  - `RemoveDuplicates`: Deletes duplicate rows based on specified columns.
  - `Trim & UCase`: Standardizes text across cells.
  - `Email Validation`: Highlights emails missing "@" or ".".
  - `Regex-like Validation`: Checks formats using VBA `Like` operator (predefined or custom patterns).
  - `Reference Check`: Validates data against a "Reference" sheet.
  - `Undo`: Restores data from a hidden backup sheet.
  - `Numeric Check`: Flags values exceeding a threshold.
- 🔄 Comprehensive reporting with regex pattern details.

---

## ⚠️ Notes
- 💾 Always back up your Excel file before running.
- 📬 Email validation targets columns with "Email" in the header (case-insensitive).
- 📱 Regex validation supports:
  - "Phone": `09########` or `+84#########`.
  - "Postal": `######`.
  - "Custom": Any VBA `Like` pattern (e.g., `EMP####`, `[0-3][0-9]-[0-1][0-9]-[2][0][0-9][0-9]`).
- 🔍 Reference validation requires a "Reference" sheet with valid data in column A.
- 🔄 Undo is limited to the most recent operation.
- 🖌️ To modify the UserForm, edit the layout and code in `AdvancedDataCleanerForm.vb`.

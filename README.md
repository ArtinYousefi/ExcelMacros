# Excel Word Count Functions (VBA)

This repository provides three easy-to-use VBA functions for Microsoft Excel that accurately count words in cells. These functions utilize regular expressions to closely match Microsoft Word's built-in word-counting logic, ensuring consistency and accuracy.

## 📌 Functions Included

### 1. `WordCount(cell As Range) → Long`

**Description:**
Returns the number of words contained in a single Excel cell.

**Usage Example:**
```excel
=WordCount(A1)
```

---

### 2. `CountWordsOver(cutoff As Long, ParamArray cells() As Variant) → Long`

**Description:**
Counts how many of the provided cells have word counts greater than the specified threshold.

**Usage Example:**
```excel
=CountWordsOver(50, A1, A2, A3)
```

---

### 3. `CountWordsOverVerbose(cutoff As Long, ParamArray cells() As Variant) → String`

**Description:**
Returns both the number of cells exceeding the specified word count threshold and their respective word counts.

**Usage Example:**
```excel
=CountWordsOverVerbose(50, A1, A2, A3)
```

**Sample Output:**
```
Count: 2 — Over 50: 54, 61
```

---

## 🚀 Quick Installation Guide

Follow these steps to use the functions in your Excel workbook:

1. Open your Excel workbook.
2. Press `Alt + F11` (Windows) or `Option + Fn + F11` (Mac) to launch the VBA editor.
3. Right-click within the Project window and select `Insert > Module`.
4. Copy and paste the VBA code from this repository into the new module.
5. Save your workbook as a macro-enabled workbook (`.xlsm`).

---

## ⚙️ Compatibility

- ✅ Compatible with Microsoft Excel on both Windows and Mac.
- ✅ Requires enabling macros in Excel.
- ✅ Utilizes VBScript regular expressions (`RegExp`) for accurate word counting.

---

## 📝 License

This project is available under the MIT License.


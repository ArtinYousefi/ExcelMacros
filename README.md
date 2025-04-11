# 🧮 Excel Word Count Functions (VBA)

This repository contains three custom VBA functions for Microsoft Excel that accurately count words in cells, mimicking Microsoft Word's definition of a "word" using regular expressions.

These functions are useful for writers, researchers, or anyone analyzing text data inside Excel.

---

## 🔧 Included Functions

### 1. `WordCount(cell As Range) → Long`
Counts the number of words in a **single cell** using regex.

#### ✅ Example:
```excel
=WordCount(A1)
2. CountWordsOver(cutoff As Long, ParamArray cells() As Variant) → Long
Counts how many of the selected cells contain more than cutoff words.

✅ Example:
excel
Copy
Edit
=CountWordsOver(50, A1, A2, A3)
Returns the number of cells (e.g. 2) that exceed 50 words.

3. CountWordsOverVerbose(cutoff As Long, ParamArray cells() As Variant) → String
Does the same as above, but also returns the actual word counts that exceeded the threshold.

✅ Example:
excel
Copy
Edit
=CountWordsOverVerbose(50, A1, A2, A3)
Output:
yaml
Copy
Edit
Count: 2 — Over 50: 54, 67
📦 How to Install in Excel
Open Excel

Press Alt + F11 (or Option + Fn + F11 on Mac) to open the Visual Basic for Applications editor

From the menu: Insert > Module

Paste in the function code from this repo

Save your file as a macro-enabled workbook (.xlsm)

🎯 Features
Supports Mac and Windows

Uses VBScript regular expressions for accurate word detection (matches Word's logic)

Supports multiple, non-contiguous cells

Ignores empty cells or cells with no valid words

Automatically trims extra whitespace and handles tabs, line breaks, and punctuation properly

🧠 Word Definition
A "word" is defined as a sequence of alphanumeric characters (\w+) bounded by word boundaries (\b) — exactly how Microsoft Word defines a word.
So "Hello, world!" counts as 2 words.

🧪 Testing & Limitations
Works in any version of Excel with macro support

Requires macro permissions to be enabled when the workbook is opened

This is a read-only tool — it does not modify your data

📝 License
MIT License

🙌 Author
Artin — feel free to fork, contribute, or integrate into your own Excel-based workflows!

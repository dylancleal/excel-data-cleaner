# 🧹 Premium Tier – Excel Data Cleaner for Complex Multi-Row Headers

## 📌 **Project Overview**

This Python script **automatically cleans complex Excel files** with multi-row headers and merged cells by:

- Combining header rows into single-level column names  
- Reshaping wide-format data into a clean, long-format table  
- Extracting **Order Date, Ship Mode, Segment, and Sales** columns  
- Standardising dates to `YYYY-MM-DD` format  
- Removing duplicates and empty sales rows

---

## 📝 **Problem Statement**

Businesses often receive Excel reports with:

- Multiple header rows (e.g. product class, segment)  
- Merged cells and poor data structure  
- Scattered sales values across many columns

Manually cleaning these files for analysis is time-consuming and error-prone.

✅ **This automated solution** converts messy spreadsheets into structured datasets ready for analysis in seconds.

---

## ⚙️ **Requirements**

- Python 3.x
- pandas
- openpyxl

---

## 🛠️ **Installation**

### **1. Install Python**

**Windows:**  
- Download from [python.org](https://www.python.org/downloads/windows/)  
- Ensure **“Add Python to PATH”** is checked during installation.

**macOS:**  
- Check if installed:  
  `python3 --version`  
- If not, install via Homebrew:  
  `brew install python`

---

### **2. Install required libraries**

**Windows:**  
`pip install pandas openpyxl`

**macOS:**  
`pip3 install pandas openpyxl`

---

## 🚀 **Usage**

1. **Update these variables in the script** to match your file paths:

```python
input_file = r'C:\Users\dylan\OneDrive\Documents\Sample_CSV_Data\dirty_data_2.xlsx'   # Full path to your messy Excel file
output_file = r'C:\Users\dylan\OneDrive\Documents\Sample_CSV_Data\cleaned_data.xlsx'  # Full path for your cleaned output file
```

---

### 🔎 **What does this script do?**

✅ Reads messy Excel files **without assuming headers**  
✅ Combines **multiple header rows** into single column names  
✅ Renames first column as **‘Order Date’**  
✅ Reshapes data using **pandas melt** to extract:

| Order Date | Ship Mode | Segment | Sales |
|------------|-----------|---------|-------|

✅ Uses regex to split Ship Mode and Segment from combined headers  
✅ Standardises dates to `YYYY-MM-DD`  
✅ Removes empty sales rows and invalid dates  
✅ Outputs a **clean, analysis-ready Excel file**

---

2. **Run the script:**

**Windows:**  
`python excel_data_cleaner.py`

**macOS:**  
`python3 excel_data_cleaner.py`

---

## 📂 **Example**

### 🔹 **Input (dirty_data_2.xlsx)**

- Multi-row headers with product classes and segments  
- Order dates as first column with no explicit header  
- Sales scattered across columns with merged header cells

### 🔹 **Output (cleaned_data.xlsx)**

| Order Date | Ship Mode      | Segment    | Sales  |
|------------|----------------|------------|--------|
| 2013-03-14 | First Class    | Consumer   | 129.44 |
| 2013-03-14 | Standard Class | Corporate  | 91.056 |
| …          | …              | …          | …      |

---

## 📸 **Screenshots**

### 🔹 **Before cleaning**

![Input Excel file](images/input_excel.png)

### 🔹 **After cleaning**

![Cleaned Excel file](images/output_excel.png)

---

## 🔑 **Notes**

- The script uses regex extraction to split Ship Mode and Segment for flexible parsing.  
- Adjust regex patterns in the script if your header naming structure differs.  
- Requires consistent multi-row header structure for automated parsing.

---

## 👨‍💻 **Author**

Dylan Cleal – Python Developer | Data Automation & Web Scraping Specialist

---

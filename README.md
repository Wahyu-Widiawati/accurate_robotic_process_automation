# 🚀 Quick Start: Accurate Sales Invoice Automation

This project automates exporting **Sales Invoices** from **Accurate 5** to Excel, and uploads the cleaned data into **PostgreSQL**.

---

## 📋 Quick Steps

1. **Run Python script**:  
   `rpa_invoice_accurate_to_local.py`  
   ➡️ Automates Accurate to export `.xls` invoice reports.

2. **Run R script**:  
   `invoice_local_to_postgre.r`  
   ➡️ Cleans the `.xls` files and uploads them into PostgreSQL.

---

## 🛠️ Requirements

- Accurate 5 Desktop App
- Python 3.10+ (`pip install pyautogui`)
- R + R packages (`readxl`, `tidyverse`, `DBI`, `RPostgres`)

---

## 📂 File Structure

| File | Purpose |
|:---|:---|
| `rpa_invoice_accurate_to_local.py` | Automates Accurate export |
| `invoice_local_to_postgre.r` | Cleans and inserts data into database |

---

## 📌 Important Notes

- Make sure screenshots of Accurate buttons are ready.
- Edit folder paths and database credentials as needed.
- Run Python script first, then R script.

---

## 📈 Workflow Summary

```plaintext
Accurate ➔ Excel (.xls) ➔ Clean in R ➔ Upload to PostgreSQL
```

---

## ✨ Future Improvements

- Add error handling
- Automate scheduling
- Add notifications

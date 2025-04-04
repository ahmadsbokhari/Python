
# PDF Table Extractor

This Python script extracts tables from a PDF and exports them into an Excel file — with smart naming for each sheet based on table headers.

Perfect for automating data extraction from multi-page PDFs like reports, invoices, or research documents.

---

## Features

- Extracts tables from every page of a PDF.
- Automatically merges tables that continue across pages.
- Generates clean Excel files with:
  - Named sheets based on table headers.
  - Padded/matched rows and columns for consistency.
- Handles missing or inconsistent table headers gracefully.

---

## How It Works

1. Uses `pdfplumber` to read tables from PDF.
2. Processes each table:
   - Cleans headers.
   - Merges multi-page tables.
   - Creates safe and unique sheet names.
3. Writes the result to an `.xlsx` Excel file using `pandas`.



## Usage

### 1. Install dependencies

```bash
pip install -r requirements.txt

### 2. Run the script

python extract_tables.py
```

By default, it reads from sample-tables.pdf and writes to named_pdf_tables.xlsx. You can modify the file paths inside the script to use your own.

import pdfplumber
import pandas as pd

pdf_path = "sample-tables.pdf"
excel_path = "named_pdf_tables.xlsx"

tables = []
current_table = None
current_header = None
used_sheet_names = set()

# Safe DataFrame builder
def safe_table_to_df(header, body):
    max_len = max(len(header), max(len(row) for row in body))
    padded_header = [(h if h else f"Col{i+1}") for i, h in enumerate(header + [""] * (max_len - len(header)))]
    cleaned_body = [row + [""] * (max_len - len(row)) if len(row) < max_len else row[:max_len] for row in body]
    return pd.DataFrame(cleaned_body, columns=padded_header)

# Generate a safe and unique sheet name
def get_safe_sheet_name(header_row, used_names, default_prefix="Table"):
    base_name = next((h for h in header_row if h and h.strip()), None)
    if not base_name:
        base_name = f"{default_prefix}_{len(used_names)+1}"
    base_name = str(base_name).strip().replace("\n", " ")[:31]

    original = base_name
    counter = 1
    while base_name in used_names:
        base_name = f"{original[:27]}_{counter}"
        counter += 1

    used_names.add(base_name)
    return base_name

# Read PDF and extract/merge tables
with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages, start=1):
        page_tables = page.extract_tables()

        for i, table in enumerate(page_tables, start=1):
            if not table or len(table) < 2:
                continue

            header = table[0]
            body = table[1:]
            cleaned_header = [str(h).strip().lower() if h else "" for h in header]

            if current_table is not None:
                is_continuation = (
                    cleaned_header == current_header or all(not h for h in header)
                )
                if is_continuation:
                    df_new = safe_table_to_df(current_header, body)
                    current_table = pd.concat([current_table, df_new], ignore_index=True)
                    continue
                else:
                    sheet_name = get_safe_sheet_name(current_header, used_sheet_names)
                    tables.append((sheet_name, current_table))

            current_header = cleaned_header
            current_table = safe_table_to_df(header, body)

# Save the last table
if current_table is not None:
    sheet_name = get_safe_sheet_name(current_header, used_sheet_names)
    tables.append((sheet_name, current_table))

# Write to Excel
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    for sheet_name, df in tables:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

print(f"✅ Saved {len(tables)} named tables to {excel_path}")

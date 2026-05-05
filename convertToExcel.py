import pdfplumber
import re
import openpyxl

pdf_path = "Medical-Record-Requirements-for-Pre-Service- UHC.pdf"

def clean(text):
    if not text:
        return ""
    text = re.sub(r"CPT® is a registered trademark[^\n]*\n?", "", text)
    text = re.sub(r"© 2019 United HealthCare[^\n]*\n?", "", text)
    return text.strip()

def extract_cell_text(page, bbox):
    x0, top, x1, bottom = bbox
    cropped = page.within_bbox((x0, top, x1, bottom))
    words = cropped.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return ""
    lines = {}
    for w in words:
        y_key = round(w['top'] / 5) * 5
        lines.setdefault(y_key, []).append(w)
    return "\n".join(
        " ".join(w['text'] for w in sorted(lines[y], key=lambda w: w['x0']))
        for y in sorted(lines)
    )

COL_X = [
    (66.62,  189.0),   # Service Category
    (189.0,  373.5),   # CPT Codes
    (373.5,  732.64),  # Clinical Information
]

excel_rows = []

with pdfplumber.open(pdf_path) as pdf:
    for page_num in range(24, 26):  # pages 25-26 (0-indexed)
        page = pdf.pages[page_num]
        print(f"\n{'='*60}")
        print(f"PAGE {page_num + 1}")
        print('='*60)

        tables = page.find_tables()
        if not tables:
            print(f"No table found on page {page_num + 1}.")
            continue

        table = tables[0]
        print(f"Table bbox: {table.bbox}")
        print(f"Rows found: {len(table.rows)}\n")

        for i, row in enumerate(table.rows):
            y_top    = row.cells[0][1] if row.cells[0] else row.bbox[1]
            y_bottom = row.cells[0][3] if row.cells[0] else row.bbox[3]

            col_texts = [
                clean(extract_cell_text(page, (x0, y_top, x1, y_bottom)))
                for x0, x1 in COL_X
            ]

            if i == 0:
                row_label = "HEADER"
            elif not col_texts[0]:
                row_label = f"ROW {i} [carry-over]"
            else:
                row_label = f"ROW {i}"

            print(f"{'─'*60}")
            print(f"{row_label}")
            print(f"  [Service Category]   : {col_texts[0]!r}")
            print(f"  [CPT Codes]          : {col_texts[1]!r}")
            print(f"  [Clinical Info]      : {col_texts[2]!r}")

            if i > 0:  # skip header row
                if col_texts[0] or not excel_rows:
                    excel_rows.append(col_texts)
                else:
                    # carry-over: merge CPT codes and clinical info into previous row
                    prev = excel_rows[-1]
                    prev[1] = (prev[1] + "\n" + col_texts[1]).strip()
                    prev[2] = (prev[2] + "\n" + col_texts[2]).strip()

# Write to Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Service Category", "CPT Codes", "Clinical Information Requested"])
for row in excel_rows:
    ws.append(row)

output_path = "output.xlsx"
wb.save(output_path)
print(f"\nSaved to {output_path}")

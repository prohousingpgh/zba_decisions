import pdfplumber
import re
import csv

def extract_application_section_lines(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()

        # Dynamically locate top and bottom Y positions
        top_app = top_appear = None
        for word in words:
            if word['text'].startswith("Application:"):
                top_app = word['top']
            elif word['text'].startswith("Appearances:"):
                top_appear = word['top']
            if top_app and top_appear:
                break

        if not top_app or not top_appear:
            raise ValueError("Couldn't locate bounding keywords")

        # Create bounding box and extract lines
        y0 = min(top_app, top_appear)
        y1 = max(top_app, top_appear)
        bbox = (0, y0, page.width, y1)
        region = page.within_bbox(bbox)
        lines = region.extract_text().splitlines()

        return [line.strip() for line in lines if line.strip()]

def parse_lines_preserving_rows(lines):
    rows = []
    current_type = ""
    current_section = ""
    current_description = ""

    for line in lines:
        # New top-level row: starts with 'Variance' or 'Special Exception'
        if line in ("Variance", "Special Exception"):
            if current_description:
                rows.append([current_type, current_section, current_description.strip()])
                current_description = ""
            current_type = line
            current_section = ""
        elif re.match(r"^Section \d", line):
            if current_description:
                rows.append([current_type, current_section, current_description.strip()])
            current_section = line
            current_description = ""
        else:
            # Line is continuation of previous description
            if current_description:
                current_description += " " + line
            else:
                current_description = line

    # Final row
    if current_description:
        rows.append([current_type, current_section, current_description.strip()])

    # Post-process: Insert empty values when fields don't change
    final_rows = []
    last_type = ""
    last_section = ""
    for row in rows:
        type_val, sec_val, desc = row
        if type_val:
            last_type = type_val
        else:
            type_val = ""
        if sec_val:
            last_section = sec_val
        else:
            sec_val = ""
        final_rows.append([type_val, sec_val, desc])

    return final_rows

def save_rows_to_csv(rows, filename="extracted_table.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Type", "Section", "Description"])
        writer.writerows(rows)

# === MAIN ===
pdf_path = "2610 Maple Ave - 63 of 2025.pdf"
lines = extract_application_section_lines(pdf_path)
rows = parse_lines_preserving_rows(lines)
save_rows_to_csv(rows, filename="2610_Maple_Ave_Table.csv")

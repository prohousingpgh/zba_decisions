import pdfplumber
import os
import csv
from pdf2image import convert_from_path
import pytesseract

def extract_application_section_lines(pdf_path):
    ocr_flag = False
    with pdfplumber.open(pdf_path) as pdf:
        lines = []
        tables = []
        if pdf.pages[0].extract_text() == "":
            print(pdf_path + " needs OCR")
            ocr_flag = True
            pages = convert_from_path(pdf_path)
            for page in pages:
                # pytesseract isn't great at grouping text - easiest way to handle it is just
                # to write the OCR text as a text-based PDF and then run that through pdfplumber
                pdf = pytesseract.image_to_pdf_or_hocr(page, extension='pdf')
                with open('temp.pdf', 'w+b') as f:
                    f.write(pdf)
                with pdfplumber.open('temp.pdf') as temp_pdf:
                    for temp_pdf_page in temp_pdf.pages:
                        lines.extend(temp_pdf_page.extract_text().splitlines())
                        tables.extend(temp_pdf_page.extract_tables({
                            "vertical_strategy": "explicit",
                            "horizontal_strategy": "text",
                            "snap_tolerance": 6,  # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
                            "explicit_vertical_lines": [temp_pdf_page.width * 0.09, temp_pdf_page.width * 0.195, temp_pdf_page.width * 0.36, temp_pdf_page.width * 0.89]
                        }))
        else:
            for page in pdf.pages:
                lines.extend(page.extract_text().splitlines())
                tables.extend(page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 6  # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
                }))

        return [[line.strip() for line in lines if line.strip()], tables, ocr_flag]

def parse_lines_preserving_rows(lines, tables, ocr_flag, filename):
    rows = []
    current_type = ""
    current_section = ""
    current_description = ""

    hearing_date = ""
    decision_date = ""
    case_number = ""
    address = ""
    decision = ""
    decision_text = ""
    in_decision = False
    for line in lines:
        if line.startswith("Date of Hearing"):
            hearing_date = line[len("Date of Hearing: "):]
        if line.startswith("Date of Decision"):
            decision_date = line[len("Date of Decision: "):]
        if line.startswith("Zone Case"):
            case_number = line[len("Zone Case: "):]
        if line.startswith("Address"):
            address = line[len("Address: "):]
        if line.startswith("Decision"):
            in_decision = True
        if line.startswith("s/"):
            in_decision = False
        if in_decision:
            decision_text += line

    start_table = False
    rows_found = False
    new_section = False
    for table in tables:
        for row in table:
            if row[0].startswith("Special") or row[0].startswith("Variance"):
                start_table = True
            if start_table and len(row) == 3:
                # First row of the table can only contain Special Exception or Variance - any other value means the table is over
                if row[0] != "" and not row[0].startswith("Special") and not row[0].startswith("Variance") and not row[0].startswith("Exception"):
                    start_table = False
                    break
                if current_description != "" and ((row[0] == "" and row[1] == "" and row[2] == "" and ("Exception" in current_type or "Variance" in current_type)) or "Special" in row[0] or row[0].startswith("Variance")):
                    rows.append([filename, hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision_text, ocr_flag])
                    rows_found = True
                    new_section = True
                    current_description = ""
                if row[0].startswith("Special") or row[0].startswith("Variance"):
                    current_type = row[0]
                if row[0].startswith("Exception"):
                    current_type += " " + row[0]
                if row[1] != "":
                    if new_section:
                        current_section = row[1]
                    else:
                        current_section += " " + row[1]
                if row[2]:
                    current_description += " " + row[2]

    # Final row
    if current_description:
        rows.append([filename, hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision_text, ocr_flag])
        rows_found = True

    if not rows_found:
        rows.append([filename, hearing_date, decision_date, case_number, address, "None found", "None found", "None found", decision_text, ocr_flag])
        print("No rows found for " + filename)

    return rows

def save_rows_to_csv(rows, filename="extracted_table.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["File Name", "Hearing Date", "Decision Date", "Case Number", "Address", "Type", "Section", "Description", "Decision", "OCR Flag"])
        writer.writerows(rows)

# === MAIN ===
input_directory = "input"
rows = []
for name in os.listdir(input_directory):
    try:
        extracted_data = extract_application_section_lines(os.path.join(input_directory, name))
        rows.extend(parse_lines_preserving_rows(extracted_data[0], extracted_data[1], extracted_data[2], name))
    except Exception as e:
        print("Error while parsing file " + name, e)
save_rows_to_csv(rows, filename="output.csv")

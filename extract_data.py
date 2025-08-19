import pdfplumber
import os
import csv
from pdf2image import convert_from_path
import pytesseract

start_indicators = ["Special", "Variance", "Review", "Protest", "Appeal"]
end_indicators = ["Exception", "Appeal"]

def get_tables(pages):
    # There are a few different shapes of table that these reports can have. Try each in sequence.
    lines = []
    tables = []
    for page in pages:
        lines.extend(page.extract_text().splitlines())
        tables.extend(page.extract_tables({
            "vertical_strategy": "lines", # Identify columns by vertical lines between them
            "horizontal_strategy": "text",
            "snap_tolerance": 6  # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
        }))
    for table in tables:
        for row in table:
            if any(row[0] is not None and row[0] in str for str in start_indicators) and len(row) == 3:
                return lines, tables
    for page in pages:
        words = page.extract_words()
        min_x = 100000
        max_x = 0
        for word in words:
            min_x = min(min_x, word["x0"])
            max_x = max(max_x, word["x1"])
        tables.extend(page.extract_tables({
            "vertical_strategy": "explicit",
            "horizontal_strategy": "text",
            "snap_tolerance": 6, # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
            "explicit_vertical_lines": [min_x, page.width * 0.195,
                                        page.width * 0.36, max_x] # Explicit vertical columns
        }))
    for table in tables:
        for row in table:
            if any(row[0] is not None and row[0] in str for str in start_indicators) and len(row) == 3:
                return lines, tables
    for page in pages:
        min_x = 100000
        max_x = 0
        for word in words:
            min_x = min(min_x, word["x0"])
            max_x = max(max_x, word["x1"])
        tables.extend(page.extract_tables({
            "vertical_strategy": "explicit",
            "horizontal_strategy": "text",
            "snap_tolerance": 6, # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
            "explicit_vertical_lines": [min_x, page.width * 0.37,
                                        page.width * 0.64, max_x] # Explicit vertical columns
        }))
    return lines, tables

def extract_application_section_lines(pdf_path):
    ocr_flag = False
    with pdfplumber.open(pdf_path) as pdf:
        lines = []
        tables = []
        # If there's no text, it's a scan and we need to OCR it
        if pdf.pages[0].extract_text() == "":
            ocr_flag = True
            pages = convert_from_path(pdf_path)
            for page in pages:
                # pytesseract isn't great at grouping text - easiest way to handle it is just
                # to write the OCR text as a text-based PDF and then run that through pdfplumber
                pdf = pytesseract.image_to_pdf_or_hocr(page, extension='pdf')
                with open('temp.pdf', 'w+b') as f:
                    f.write(pdf)
                with pdfplumber.open('temp.pdf') as temp_pdf:
                    new_lines, new_tables = get_tables(temp_pdf.pages)
                    lines.extend(new_lines)
                    tables.extend(new_tables)
        else:
            new_lines, new_tables = get_tables(pdf.pages)
            lines.extend(new_lines)
            tables.extend(new_tables)

        return [line.strip() for line in lines if line.strip()], tables, ocr_flag

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
            decision_text += line + " "
    if ("APPROVES" in decision_text or "may continue" in decision_text or "APPROVED" in decision_text) and not ("DENIED" in decision_text or "DENIES" in decision_text):
        decision = "APPROVED"
    elif ("DENIED" in decision_text or "DENIES" in decision_text) and not ("APPROVES" in decision_text or "may continue" in decision_text or "APPROVED" in decision_text):
        decision = "DENIED"
    else:
        decision = "UNKNOWN"

    start_table = False
    rows_found = False
    new_section = False
    for table in tables:
        for row in table:
            if row[0] == None:
                row[0] = ""
            if row[0].startswith(tuple(start_indicators)) and len(row) == 3:
                start_table = True
            if start_table and len(row) == 3:
                # First column of the table contains particular values - any other value means the table is over
                if row[0] != "" and not row[0].startswith(tuple(start_indicators)) and not row[0].startswith(tuple(end_indicators)):
                    start_table = False
                    break
                # Write an entry to the final list if it looks like the entry is over
                if current_description != "" and ((row[0] == "" and row[1] == "" and row[2] == "" and ("Exception" in current_type or "Variance" in current_type or "Appeal" in current_type or "Review" in current_type)) or row[0].startswith(tuple(start_indicators))):
                    rows.append([filename, hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision, decision_text, ocr_flag])
                    rows_found = True
                    new_section = True
                    current_description = ""
                # Process the next line of the table
                if row[0].startswith(tuple(end_indicators)):
                    current_type += " " + row[0]
                elif row[0].startswith(tuple(start_indicators)):
                    current_type = row[0]
                if row[1] != "":
                    if new_section:
                        current_section = row[1]
                    else:
                        current_section += " " + row[1]
                if row[2]:
                    current_description += " " + row[2]

    # Final row
    if current_type and current_description:
        rows.append([filename, hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision, decision_text, ocr_flag])
        rows_found = True

    if not rows_found:
        rows.append([filename, hearing_date, decision_date, case_number, address, "None found", "None found", "None found", decision, decision_text, ocr_flag])
        print("No rows found for " + filename)

    return rows

def save_rows_to_csv(rows, filename="extracted_table.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["File Name", "Hearing Date", "Decision Date", "Case Number", "Address", "Type", "Section", "Description", "Decision", "Decision Text", "OCR Flag"])
        writer.writerows(rows)

# === MAIN ===
input_directory = "input"
rows = []
for name in os.listdir(input_directory):
    try:
        lines, tables, ocr_flag = extract_application_section_lines(os.path.join(input_directory, name))
        rows.extend(parse_lines_preserving_rows(lines, tables, ocr_flag, name))
    except Exception as e:
        print("Error while parsing file " + name, e)
        rows.append([name, "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error", "Error"])
save_rows_to_csv(rows, filename="output.csv")

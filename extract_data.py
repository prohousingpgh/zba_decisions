import pdfplumber
import os
import csv

def extract_application_section_lines(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        lines = []
        tables = []
        for page in pdf.pages:
            lines.extend(page.extract_text().splitlines())
            tables.extend(page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "text",
                "snap_tolerance": 6  # Determines the sensitivity when grouping misaligned table text into a single line. May need to tweak this to get good results.
            }))

        return [[line.strip() for line in lines if line.strip()], tables]

def parse_lines_preserving_rows(lines, tables):
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
        if line.startswith("Date of Hearing: "):
            hearing_date = line[len("Date of Hearing: "):]
        if line.startswith("Date of Decision: "):
            decision_date = line[len("Date of Decision: "):]
        if line.startswith("Zone Case: "):
            case_number = line[len("Zone Case: "):]
        if line.startswith("Address: "):
            address = line[len("Address: "):]
        if line.startswith("Decision: "):
            in_decision = True
        if line.startswith("s/"):
            in_decision = False
        if in_decision:
            decision_text += line

    for table in tables:
        for row in table:
            if len(row) == 3:
                if row[2] == "":
                    rows.append([hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision_text])
                    current_description = ""
                else:
                    if row[0] != "":
                        current_type = row[0]
                    if row[1] != "":
                        current_section = row[1]
                    if row[2]:
                        current_description = current_description + " " + row[2]

    # Final row
    if current_description:
        rows.append([hearing_date, decision_date, case_number, address, current_type, current_section, current_description.strip(), decision_text])

    return rows

def save_rows_to_csv(rows, filename="extracted_table.csv"):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Hearing Date", "Decision Date", "Case Number", "Address", "Type", "Section", "Description", "Decision"])
        writer.writerows(rows)

# === MAIN ===
input_directory = "input"
rows = []
for name in os.listdir(input_directory):
    try:
        extracted_data = extract_application_section_lines(os.path.join(input_directory, name))
        rows.extend(parse_lines_preserving_rows(extracted_data[0], extracted_data[1]))
    except Exception as e:
        print("Error while parsing file " + name, e)
save_rows_to_csv(rows, filename="output.csv")

import os
import re
import logging
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook

# -----------------------------
# CONFIGURATION
# -----------------------------
PROGRAM_NAME = "ReportToExcelWorkbook"
RUN_DIR = "d:/data/"
INPUT_FILE_NAME = "input.txt"

def timestamp():
    return datetime.now().strftime("%Y%m%d-%H%M%S")

RUNSTAMP = timestamp()

IN_FILE = f"{RUN_DIR}{INPUT_FILE_NAME}"
LOG_FILE = f"{RUN_DIR}{PROGRAM_NAME}-{INPUT_FILE_NAME}-{RUNSTAMP}.log"
OUT_FILE = f"{RUN_DIR}{PROGRAM_NAME}-{INPUT_FILE_NAME}-{RUNSTAMP}.xlsx"

# -----------------------------
# LOGGING SETUP
# -----------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

logging.info(f"Starting {PROGRAM_NAME}...")
logging.info(f"Log file created: {LOG_FILE}")

# -----------------------------
# PARSING LOGIC
# -----------------------------
def parse_text_file(text: str):
    """
    Parse the custom text file into a dictionary:
    {
        "SheetName": [ [header1, header2], [row1col1, row1col2], ... ],
        ...
    }
    """

    # Remove decorative separators
    cleaned = re.sub(r"=+\n", "", text)

    # Regex to capture sheet name + table block
    pattern = r"SHEET:\s*(.+?)\n(.*?)(?=\nSHEET:|$)"
    matches = re.findall(pattern, cleaned, flags=re.DOTALL)

    sheets = {}

    for sheet_name, block in matches:
        logging.info(f"Parsing sheet: {sheet_name}")

        rows = []
        for line in block.splitlines():
            if "|" not in line:
                continue

            parts = [col.strip() for col in line.split("|")]
            # parts = [p for p in parts if p]  # remove empty columns

            if parts:
                rows.append(parts)

        if rows:
            sheets[sheet_name] = rows
        else:
            logging.warning(f"Sheet '{sheet_name}' contained no table rows.")

    return sheets


# -----------------------------
# WRITE EXCEL 
# -----------------------------
def write_excel(sheets: dict, outfile: str):
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sheet_name, rows in sheets.items():
        safe_name = sheet_name[:31]  # Excel sheet name limit
        row_count = len(rows)        # count rows including header

        logging.info(f"Creating sheet: {safe_name} with {row_count} rows")

        ws = wb.create_sheet(title=safe_name)

        for row in rows:
            ws.append(row)

    wb.save(outfile)
    logging.info(f"Workbook saved: {outfile}")


# -----------------------------
# ENTRY POINT
# -----------------------------
if __name__ == "__main__":
    """    
    input_path = Path(RUN_DIR) / INPUT_FILE_NAME

    with open(input_path, "r", encoding="utf-8") as f:
        content = f.read()
    """
    with open(IN_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    sheets = parse_text_file(content)
    logging.info(f"Parsed {len(sheets)} sheets successfully.")

    write_excel(sheets, OUT_FILE)

    logging.info(f"{PROGRAM_NAME} completed successfully.")

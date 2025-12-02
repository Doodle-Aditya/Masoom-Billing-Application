# excel_handler/excel_manager.py
import os
import re
from openpyxl import Workbook, load_workbook
from typing import List
from PyPDF2 import PdfReader

EXCEL_PATH = os.path.join("output", "records.xlsx")
TEMPLATE_DONOR_PDF = os.path.join("pdf_generator", "Donor Name.pdf")  # uploaded file

HEADERS = [
    "Date", "Ref No", "College Name", "Class", "Student Name",
    "College Fees", "Total Fees", "Masoom Contribution", "Student Contribution",
    "Payable", "Rs in Words", "Donor Name", "Cheque Issue Name",
    "Account Holder", "Bank Name", "Account Number", "IFSC", "Cheque Number"
]

def ensure_excel():
    if not os.path.exists("output"):
        os.makedirs("output")
    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "Records"
        ws.append(HEADERS)
        wb.save(EXCEL_PATH)

def save_to_excel(data: dict):
    """
    Append a row of data to records.xlsx
    """
    ensure_excel()
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    row = [
        data.get("date", ""),
        data.get("ref_no", ""),
        data.get("college_name", ""),
        data.get("class", ""),
        data.get("student_name", ""),
        data.get("college_fees", ""),
        data.get("total_fees", ""),
        data.get("masoom_contribution", ""),
        data.get("student_contribution", ""),
        data.get("payable", ""),
        data.get("rs_in_words", ""),
        data.get("donor_name", ""),
        data.get("cheque_issue_name", ""),
        data.get("account_holder", ""),
        data.get("bank_name", ""),
        data.get("account_number", ""),
        data.get("ifsc", ""),
        data.get("cheque_number", "")
    ]
    ws.append(row)
    wb.save(EXCEL_PATH)

def get_last_ref_no() -> str:
    """
    Returns last ref_no string from the last non-header row, or empty if none.
    """
    ensure_excel()
    wb = load_workbook(EXCEL_PATH, read_only=True)
    ws = wb.active
    max_row = ws.max_row
    if max_row <= 1:
        return ""
    last_ref = ws.cell(row=max_row, column=2).value  # "Ref No" is column 2 in HEADERS
    return last_ref if last_ref else ""

def next_ref_no() -> str:
    """
    Generate next reference number based on last Excel row.
    Format: M-0001/24-25/LT (increment the numeric part)
    """
    last = get_last_ref_no()
    if not last:
        next_num = 1
    else:
        # extract first numeric section after 'M-'
        m = re.search(r"M-0*([0-9]+)", last)
        if m:
            next_num = int(m.group(1)) + 1
        else:
            next_num = 1
    return f"M-{next_num:04d}/24-25/LT"

def load_donors_from_pdf(pdf_path: str = TEMPLATE_DONOR_PDF) -> List[str]:
    """
    Extract donor names from the provided PDF (simple heuristic).
    Returns unique cleaned lines as dropdown options.
    """
    if not os.path.exists(pdf_path):
        return []
    try:
        reader = PdfReader(pdf_path)
        raw = []
        for p in reader.pages:
            text = p.extract_text()
            if text:
                raw.append(text)
        all_text = "\n".join(raw)
        # split into lines, clean
        lines = [line.strip() for line in all_text.splitlines() if line.strip()]
        # de-duplicate while preserving order
        seen = set()
        out = []
        for line in lines:
            # remove odd characters, trailing numbers, excessive spaces
            cleaned = " ".join(line.split())
            if cleaned and cleaned not in seen:
                seen.add(cleaned)
                out.append(cleaned)
        return out
    except Exception:
        return []

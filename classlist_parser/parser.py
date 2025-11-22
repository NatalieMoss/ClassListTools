"""
PCC Classlist Party Starter
---------------------------
A tool for parsing Banner SIS PDF class lists into clean spreadsheets.
Part of the PCC Classlist Party Pack.
"""

import pdfplumber
import re
import pandas as pd
from collections import defaultdict
from tkinter import Tk, filedialog, messagebox
import sys
import os

_TERM_TEXT_RE = re.compile(r'\b(Spring|Summer|Fall|Winter)\s+(20\d{2})\b', re.I)
_TERM_CODE_RE = re.compile(r'\b(20\d{2}0[1-4])\b')  # 202501..202504


def _term_from_code(code: str) -> str:
    season_map = {1: "Winter", 2: "Spring", 3: "Summer", 4: "Fall"}
    return f"{season_map.get(int(code[-1]), '')} {code[:4]}".strip()

def _detect_term_from_pdf(pdf) -> str:
    try:
        header = (pdf.pages[0].extract_text() or "")
    except Exception:
        header = ""
    m = _TERM_TEXT_RE.search(header)
    if m:
        return f"{m.group(1).title()} {m.group(2)}"
    for m in _TERM_CODE_RE.finditer(header):
        guess = _term_from_code(m.group(0))
        if guess:
            return guess
    return ""  # ok if empty

def _safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r'[^A-Za-z0-9 _\-]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

# List of allowed GEO courses, now including 242
allowed_courses = {"170", "221", "223", "240", "242", "244", "246", "248", "252", "254", "260", "265", "266", "267", "270", "280A"}

records = []
current_class = {}

try:
    root = Tk()
    root.withdraw()
    pdf_file = filedialog.askopenfilename(title="Select your class list PDF", filetypes=[("PDF files", "*.pdf")])
    if not pdf_file:
        sys.exit()

    records = []
    current_class = {}

    with pdfplumber.open(pdf_file) as pdf:
        # detect the term while the file is open
        term_text = _detect_term_from_pdf(pdf)   # "Fall 2025" / "Summer 2025" / ""

        for page in pdf.pages:
            text = (page.extract_text() or "")
            lines = text.split("\n")

            # header line: capture CRN, subject, course number, section, name
            for line in lines:
                course_match = re.match(r"\s*(\d{5})\s+(\w+)\s+(\d+[A-Z]?)\s+(\d)\s+(.*)", line)
                if course_match:
                    course_number = course_match.group(3)
                    if course_number not in allowed_courses:
                        current_class = {}
                    else:
                        current_class = {
                            "CRN": course_match.group(1),
                            "Subject": course_match.group(2),
                            "Course Number": course_number,
                            "Section": course_match.group(4),
                            "Course Name": course_match.group(5).strip(),
                        }

            # student rows
            idx = 0
            while idx < len(lines):
                line = lines[idx]
                gnum_match = re.search(r"(G\d{8})", line)
                if gnum_match and current_class:
                    try:
                        name_part = line.split(gnum_match.group(1))[0]
                        last_first = name_part.split(None, 1)[1].split(',')
                        last_name = last_first[0].strip()
                        first_name = last_first[1].strip()
                    except Exception:
                        last_name, first_name = "", ""

                    g_number = gnum_match.group(1)

                    email = ""
                    if idx + 1 < len(lines):
                        email_line = lines[idx + 1].strip()
                        if "@pcc.edu" in email_line:
                            email = email_line.split()[0]

                    records.append({
                        "First Name": first_name,
                        "Last Name": last_name,
                        "G Number": g_number,
                        "PCC email address": email,
                        "Non-PCC email": "",
                        "Class": f"{current_class.get('Subject')} {current_class.get('Course Number')}",
                        "CRN": current_class.get("CRN"),
                    })
                    idx += 2
                else:
                    idx += 1

    # build dynamic output name (now that pdf is closed)
    stem = f"GEO_Class_Lists_{term_text}" if term_text else "GEO_Class_Lists"
    out_name = _safe_filename(stem) + ".xlsx"

    def app_dir():
        # Folder of the EXE if frozen, else folder of this .py file
        return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
            else os.path.dirname(os.path.abspath(__file__))


    # Always write to: <tool folder>\Output Files
    output_dir = os.path.join(app_dir(), "Output Files")
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, out_name)  # out_name already "GEO_Class_Lists_<term>.xlsx"

    # output_dir = os.path.dirname(pdf_file)  # or wherever you prefer
    # output_path = os.path.join(output_dir, out_name)

    # group and write
    grouped = defaultdict(list)
    for rec in records:
        grouped[f"{rec['Class']}_{rec['CRN']}"].append(rec)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(records).to_excel(writer, sheet_name="Combined", index=False)
        for key, recs in grouped.items():
            sheet_name = key[:31]
            pd.DataFrame(recs).to_excel(writer, sheet_name=sheet_name, index=False)

    messagebox.showinfo("Done", f"Created:\n{output_path}")

except Exception as e:
    messagebox.showerror("Error", f"There was an error: {e}")



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

# CONFIGURATION SETTINGS FOR THIS TOOL.
# Other departments or colleges can edit these values to customize behavior.
# 🔧 TO CUSTOMIZE FOR YOUR DEPARTMENT:
# - Change department_prefix to your subject code (e.g., "MTH", "WR").
# - Change allowed_courses to your course numbers or set to None.
# - Change email_domain if you're not at PCC.

SETTINGS = {
    # If you want to restrict parsing to a single subject/department
    # (e.g., "GEO"), put the subject code here. Otherwise, use None.
    "department_prefix": None,  # e.g., "GEO" or None for all subjects

    # List of allowed course numbers as strings (e.g., ["170", "221"]).
    # Set to None to accept all course numbers.
    "allowed_courses": {"170", "221", "223", "240", "242", "244", "246",
                        "248", "252", "254", "260", "265", "266", "267",
                        "270", "280A"},

    # Email domain used for institutional emails in the PDF.
    # Other colleges can change this to their domain.
    "email_domain": "@pcc.edu",

    # Prefix for the output Excel filename (term will be appended if found).
    "output_name_prefix": "GEO_Class_Lists",

    # Subfolder where output files are written.
    "output_subfolder": "Output Files",
}


_TERM_TEXT_RE = re.compile(r'\b(Spring|Summer|Fall|Winter)\s+(20\d{2})\b', re.I)
_TERM_CODE_RE = re.compile(r'\b(20\d{2}0[1-4])\b')  # e.g. 202501..202504


def _term_from_code(code: str) -> str:
    """
    Convert a Banner-style term code (e.g. '202503') into a human-readable label.
    Args:
        code (str): The 6-digit Banner term code. The last digit encodes the season:
                    1 = Winter, 2 = Spring, 3 = Summer, 4 = Fall.
    Returns:
        str: A string like 'Spring 2025' or an empty string if the code is not valid.
    """
    season_map = {1: "Winter", 2: "Spring", 3: "Summer", 4: "Fall"}
    # Example: '202503' -> last digit is 3 -> 'Summer', year is '2025'
    return f"{season_map.get(int(code[-1]), '')} {code[:4]}".strip()


def _detect_term_from_pdf(pdf) -> str:
    """
    Detect the academic term from the first page of a Banner class list PDF.
    This function checks for:
      1. Explicit term text like 'Fall 2025', or
      2. A numeric Banner term code like '202503', which it then converts.
    Args:
        pdf: An open pdfplumber PDF object.
    Returns:
        str: A human-readable term such as 'Fall 2025', or an empty string
             if no term information can be found.
    """
    try:
        header = (pdf.pages[0].extract_text() or "")
    except Exception:
        header = ""

    # Try a direct match like 'Fall 2025'
    m = _TERM_TEXT_RE.search(header)
    if m:
        return f"{m.group(1).title()} {m.group(2)}"

    # Otherwise, look for Banner codes like '202503'
    for m in _TERM_CODE_RE.finditer(header):
        guess = _term_from_code(m.group(0))
        if guess:
            return guess

    return ""  # okay if empty; we'll fall back to a generic filename


def _safe_filename(s: str) -> str:
    """
    Clean a string so it is safe to use as a filename on most operating systems.
    This:
      * Replaces None with an empty string.
      * Removes characters that are not letters, numbers, spaces, underscores, or hyphens.
      * Collapses repeated whitespace into a single space.
    Args:
        s (str): The original string (e.g. a course title or term label).
    Returns:
        str: A simplified, filesystem-safe version of the string.
    """
    s = (s or "").strip()
    # Keep only letters, digits, spaces, underscores, and hyphens
    s = re.sub(r'[^A-Za-z0-9 _\-]', '', s)
    # Collapse multiple spaces into one
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def app_dir() -> str:
    """
    Get the folder where this tool should write its output.
    If the script is 'frozen' into an EXE (using something like PyInstaller),
    this returns the folder where the EXE lives. Otherwise, it returns the
    folder containing this .py file.
    Returns:
        str: Absolute path to the application's base directory.
    """
    if getattr(sys, "frozen", False):
        # Running as a bundled EXE
        return os.path.dirname(sys.executable)
    # Running as a normal .py file
    return os.path.dirname(os.path.abspath(__file__))


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

                    # Apply optional course filter
                    allowed = SETTINGS.get("allowed_courses")
                    if allowed is not None and course_number not in allowed:
                        # Skip this class if it's not in the allowed list
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
                        if SETTINGS["email_domain"] in email_line:
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
    prefix = SETTINGS["output_name_prefix"]
    stem = f"{prefix}_{term_text}" if term_text else prefix
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



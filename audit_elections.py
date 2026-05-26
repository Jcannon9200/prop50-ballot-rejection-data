"""
Election file auditor — scans all data files and classifies them as:
  CONFIRMED : relates to the November 2025 California Special Election
  FLAGGED   : relates to a different election (2022, 2024, etc.)
  UNKNOWN   : could not determine election from filename or content
"""

import os
import re
import csv
import zipfile
from pathlib import Path

ROOT = Path(r"D:\Ballot Rejection Data")
REPORT = ROOT / "election_file_audit.txt"

# ---------------------------------------------------------------------------
# Election detection patterns
# ---------------------------------------------------------------------------

CONFIRMED_PATTERNS = [
    r"november\s+4[,\s]+2025",
    r"nov(?:ember)?\s*[,\-]?\s*2025",
    r"11[/\-]4[/\-]2025",
    r"11[/\-]04[/\-]2025",
    # 2-digit year: 11-04-25 or 11/4/25
    r"11[/\-]0?4[/\-]25\b",
    # hyphenated: 2025-Special or 2025-special
    r"2025[\s\-_]+special",
    r"special[\s\-_]+2025",
    r"special\s+election\s+2025",
    r"\bN25\b",
    r"2025.*special.*election",
    r"special.*election.*2025",
    r"november.*2025.*election",
    # response/PRA dated 2026 for Nov 2025 election (common pattern in filenames)
    r"rejected\s+ballot.*2025",
    r"2025.*rejected\s+ballot",
    # Solano-style: 2025 + election-specific terms in filename
    r"2025.*\b(?:VBM|provisional|challenged|CVR|ballot)\b",
    r"\b(?:VBM|provisional|challenged|CVR)\b.*2025",
    # date format 11-04-25 in text body
    r"\b11[\-/]04[\-/]25\b",
    r"\b11[\-/]4[\-/]25\b",
]

FLAGGED_PATTERNS = [
    (r"\bG2024\b|\bP2024\b|general\s+2024|primary\s+2024|november\s+2024|november.*2024", "2024 election"),
    (r"\bG2022\b|\bP2022\b|general\s+2022|primary\s+2022|november\s+2022|november.*2022", "2022 election"),
    (r"\b202[0-3]\b",                                                                      "pre-2024 year reference"),
    (r"march\s+2024|june\s+2024|primary.*2024|2024.*primary",                              "2024 primary"),
    (r"march\s+2026|june\s+2026",                                                          "2026 election reference"),
]

_confirmed_re = re.compile("|".join(CONFIRMED_PATTERNS), re.IGNORECASE)
_flagged_res  = [(re.compile(p, re.IGNORECASE), label) for p, label in FLAGGED_PATTERNS]


def classify_text(text: str) -> tuple[str, str]:
    """Return (status, reason) based on text content."""
    if _confirmed_re.search(text):
        m = _confirmed_re.search(text)
        return "CONFIRMED", f"matched: '{m.group().strip()}'"
    for pattern, label in _flagged_res:
        m = pattern.search(text)
        if m:
            return "FLAGGED", f"{label} — matched: '{m.group().strip()}'"
    return "UNKNOWN", "no election date pattern found"


def classify_filename(name: str) -> tuple[str, str] | None:
    """Quick filename-only classification. Returns None if inconclusive."""
    status, reason = classify_text(name)
    if status != "UNKNOWN":
        return status, f"filename: {reason}"
    # Humboldt-style codes in filenames
    if re.search(r"[-_ ][GP]202[0-9]", name, re.IGNORECASE):
        m = re.search(r"[-_ ]([GP]202[0-9])", name, re.IGNORECASE)
        return "FLAGGED", f"filename election code: {m.group(1)}"
    return None


# ---------------------------------------------------------------------------
# Text extraction helpers
# ---------------------------------------------------------------------------

MAX_CHARS = 50_000  # limit per file to keep it fast


def read_pdf(path: Path) -> str:
    try:
        import PyPDF2
        text = []
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text.append(page.extract_text() or "")
                if sum(len(t) for t in text) > MAX_CHARS:
                    break
        return " ".join(text)
    except Exception as e:
        return f"[PDF read error: {e}]"


def read_xlsx(path: Path) -> str:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        parts = []
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                line = " ".join(str(c) for c in row if c is not None)
                if line.strip():
                    parts.append(line)
                if sum(len(p) for p in parts) > MAX_CHARS:
                    break
        return " ".join(parts)
    except Exception as e:
        return f"[XLSX read error: {e}]"


def read_xls(path: Path) -> str:
    try:
        import xlrd
        wb = xlrd.open_workbook(path)
        parts = []
        for sheet in wb.sheets():
            for rx in range(sheet.nrows):
                line = " ".join(str(v) for v in sheet.row_values(rx) if v != "")
                if line.strip():
                    parts.append(line)
                if sum(len(p) for p in parts) > MAX_CHARS:
                    break
        return " ".join(parts)
    except ImportError:
        # xlrd not installed — fall back to filename only
        return "[xls: xlrd not installed]"
    except Exception as e:
        return f"[XLS read error: {e}]"


def read_docx(path: Path) -> str:
    try:
        from docx import Document
        doc = Document(path)
        return " ".join(p.text for p in doc.paragraphs)[:MAX_CHARS]
    except ImportError:
        return "[docx: python-docx not installed]"
    except Exception as e:
        return f"[DOCX read error: {e}]"


def read_txt(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read(MAX_CHARS)
    except Exception as e:
        return f"[TXT read error: {e}]"


def read_csv(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read(MAX_CHARS)
    except Exception as e:
        return f"[CSV read error: {e}]"


def read_zip(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
        return f"[ZIP contents: {', '.join(names[:30])}]"
    except Exception as e:
        return f"[ZIP read error: {e}]"


READERS = {
    ".pdf":  read_pdf,
    ".xlsx": read_xlsx,
    ".xls":  read_xls,
    ".docx": read_docx,
    ".doc":  read_docx,
    ".txt":  read_txt,
    ".csv":  read_csv,
    ".zip":  read_zip,
}

SKIP_DIRS = {".git"}

# ---------------------------------------------------------------------------
# Main scan
# ---------------------------------------------------------------------------

results = []  # list of dicts

for dirpath, dirs, files in os.walk(ROOT):
    # Skip hidden/git dirs
    dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
    county = Path(dirpath).relative_to(ROOT).parts[0] if Path(dirpath) != ROOT else "(root)"

    for fname in sorted(files):
        if fname == REPORT.name or fname == Path(__file__).name:
            continue
        fpath = Path(dirpath) / fname
        suffix = fpath.suffix.lower()

        if suffix not in READERS:
            continue

        rel = str(fpath.relative_to(ROOT))
        print(f"Scanning: {rel}")

        # 1. Try filename classification first
        fn_result = classify_filename(fname)
        if fn_result:
            status, reason = fn_result
            results.append({"county": county, "file": rel, "status": status, "reason": reason})
            continue

        # 2. Read content
        text = READERS[suffix](fpath)
        combined = fname + " " + text

        # 3. Classify on content + filename together
        status, reason = classify_text(combined)
        results.append({"county": county, "file": rel, "status": status, "reason": reason})

# ---------------------------------------------------------------------------
# Write report
# ---------------------------------------------------------------------------

confirmed = [r for r in results if r["status"] == "CONFIRMED"]
flagged   = [r for r in results if r["status"] == "FLAGGED"]
unknown   = [r for r in results if r["status"] == "UNKNOWN"]

SEP = "=" * 70

lines = [
    SEP,
    "ELECTION FILE AUDIT REPORT",
    "Target election: November 2025 California Special Election",
    SEP,
    f"Total files scanned : {len(results)}",
    f"CONFIRMED (Nov 2025): {len(confirmed)}",
    f"FLAGGED (other elec.): {len(flagged)}",
    f"UNKNOWN             : {len(unknown)}",
    "",
]

lines += [SEP, f"CONFIRMED ({len(confirmed)} files)", SEP]
for r in confirmed:
    lines += [f"  [{r['county']}]  {r['file']}", f"    Reason: {r['reason']}", ""]

lines += [SEP, f"FLAGGED — DIFFERENT ELECTION ({len(flagged)} files)", SEP]
for r in flagged:
    lines += [f"  [{r['county']}]  {r['file']}", f"    Reason: {r['reason']}", ""]

lines += [SEP, f"UNKNOWN — COULD NOT DETERMINE ({len(unknown)} files)", SEP]
for r in unknown:
    lines += [f"  [{r['county']}]  {r['file']}", f"    Reason: {r['reason']}", ""]

report_text = "\n".join(lines)
REPORT.write_text(report_text, encoding="utf-8")
print()
print(report_text)
print(f"\nReport saved to: {REPORT}")

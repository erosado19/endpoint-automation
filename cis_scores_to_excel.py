"""
Read CIS benchmark report PDFs (e.g. CIS-CAT results), pull out the score,
and record it in a tracking spreadsheet under the right fiscal quarter.

Runs only on PDFs you already have. It makes no network calls.

Usage:
    python cis_scores_to_excel.py reports/ --workbook CIS_Scores.xlsx --fy-start-month 10
    python cis_scores_to_excel.py report1.pdf report2.pdf --workbook CIS_Scores.xlsx --dry-run
    python cis_scores_to_excel.py reports/ --workbook CIS_Scores.xlsx --quarter "FY26 Q2"
    python cis_scores_to_excel.py report.pdf --show-text      # print what the script reads, for tuning

Spreadsheet layout ("CIS Scores" sheet):
    Host | Benchmark | FY26 Q1 | FY26 Q2 | ...
An "Import log" sheet records every file read, so you can trace each number.
A backup copy of the workbook is saved next to it before every change.
"""
import argparse
import re
import shutil
import sys
from bisect import bisect_right
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

try:
    from pypdf import PdfReader
except ImportError:
    sys.exit("This script needs pypdf. Install it with:  pip install pypdf")

NUM = r"(\d{1,3}(?:\.\d+)?)"
# Patterns tried in order. The first match wins. Add your own with --score-regex.
# Specific "total/overall" wording comes before the generic "Score:" so a
# per-section score earlier in the report isn't picked by mistake.
SCORE_PATTERNS = [
    rf"\bTotal\s+Score\s*[:\-]?\s*{NUM}\s*%",
    rf"\bOverall\s+Score\s*[:\-]?\s*{NUM}\s*%",
    rf"\bBenchmark\s+Score\s*[:\-]?\s*{NUM}\s*%",
    rf"(?<![\w-])Compliant\s*[:\-]?\s*{NUM}\s*%",          # "Compliant: 78%" (not "Non-Compliant")
    rf"\bCompliance(?:\s+Score)?\s*[:\-]?\s*{NUM}\s*%",
    rf"\bPassed\s+Checks\s*[:\-]?\s*{NUM}\s*%",            # CIS dashboard: "Passed Checks 64.5% (20)"
    rf"\bTotal\s+[\d.]+\s+[\d.]+\s+{NUM}\s*%",   # summary table row: Total <pass> <fail> <pct>%
    rf"\bScore\s*[:\-]\s*{NUM}\s*%",
]
HOST_PATTERNS = [
    r"\b(?:Target\s+)?Host\s*name\s*[:\-]?\s*([A-Za-z0-9][\w.\-]*)",
    r"\bComputer\s+Name\s*[:\-]?\s*([A-Za-z0-9][\w.\-]*)",
    r"\b(?:Asset|Device|Endpoint|Machine|System)\s+Name\s*[:\-]?\s*([A-Za-z0-9][\w.\-]*)",
    r"\bDeployment\s+Name\s*[:\-]?\s*([A-Za-z0-9][^\n]{0,80})",   # may contain spaces
    r"\bTarget\s*[:\-]\s*([A-Za-z0-9][\w.\-]*)",
    r"\bHost\s*[:\-]\s*([A-Za-z0-9][\w.\-]*)",
]
DATE_LABELS = [
    r"\b(?:Assessment|Scan|Report|Start)\s+(?:Date|Time)\s*[:\-]?\s*([^\n]{6,40})",
    r"\bLast\s+(?:Updated|Scanned)\s*(?:Time|Date|On)?\s*[:\-]?\s*([^\n]{6,40})",
    r"\b(?:Scanned|Assessed|Generated)\s+(?:On|At|Date)\s*[:\-]?\s*([^\n]{6,40})",
    r"\bDate\s*[:\-]\s*([^\n]{6,40})",
]
# Each report in a multi-report PDF starts with this word. Change it with --split-on.
DEFAULT_SPLIT = "Compliant"
MONTHS = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?"
DATE_SHAPES = [
    (re.compile(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})"), "ymd"),
    (re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})"), "mdy"),
    (re.compile(r"(\d{1,2})-(\d{1,2})-(\d{4})"), "mdy"),
    (re.compile(rf"({MONTHS})\s+(\d{{1,2}}),?\s+(\d{{4}})", re.I), "Mdy"),
    (re.compile(rf"(\d{{1,2}})\s+({MONTHS}),?\s+(\d{{4}})", re.I), "dMy"),
]
BENCH_PATTERN = r"(CIS\s+[^\n]{3,90}?Benchmark(?:\s+\(?v?\d+(?:\.\d+)*\)?)?)"
QUARTER_RE = re.compile(r"^\s*FY\s*'?(\d{2}|\d{4})\s*[- ]?\s*Q([1-4])\s*$", re.I)

SCORES_SHEET, LOG_SHEET = "CIS Scores", "Import log"
LOG_HEADER = ["Imported at", "File", "Host", "Benchmark", "Report date", "Quarter", "Score %", "Action"]
BOLD = Font(bold=True)
CHANGED = PatternFill("solid", fgColor="FAEEDA")


class CisError(Exception):
    """A problem the user can fix."""


# ---------------------------------------------------------------- reading PDFs

def pdf_pages(path):
    """Return the text of each page, in order."""
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            raise CisError("PDF is password protected")
    return [(pg.extract_text() or "") for pg in reader.pages]


def pdf_text(path):
    return "\n".join(pdf_pages(path))


def split_reports(pages, marker=DEFAULT_SPLIT):
    """Cut a PDF's text into separate reports, each starting at the marker word.

    Returns a list of (first page number, report text). A PDF without the marker,
    or with marker set to "", is treated as one report. "Non-Compliant" and
    "NonCompliant" don't count as a report start.
    """
    starts, pos = [], 0
    for p in pages:
        starts.append(pos)
        pos += len(p) + 1
    text = "\n".join(pages)
    if not marker:
        return [(1, text)]
    rx = re.compile(rf"(?<![\w-]){re.escape(marker)}(?![\w])", re.IGNORECASE)
    hits = [m.start() for m in rx.finditer(text)]
    if not hits:
        return [(1, text)]
    bounds = hits + [len(text)]
    return [(bisect_right(starts, a), text[a:b]) for a, b in zip(bounds, bounds[1:])]


def first(patterns, text):
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def parse_date(s, dayfirst=False):
    """Find the first date in a string such as 'February 12, 2026 10:42 AM'."""
    if not s:
        return None
    candidates = sorted(((m.start(), m, kind) for rx, kind in DATE_SHAPES for m in rx.finditer(s)),
                        key=lambda t: t[0])
    for _, m, kind in candidates:     # earliest date in the text first; skip invalid ones
        g = m.groups()
        try:
            if kind == "ymd":
                return datetime(int(g[0]), int(g[1]), int(g[2]))
            if kind == "mdy":
                a, b, y = int(g[0]), int(g[1]), int(g[2])
                if dayfirst or a > 12:
                    a, b = b, a
                return datetime(y, a, b)
            if kind == "Mdy":
                return datetime.strptime(f"{g[0][:3]} {g[1]} {g[2]}", "%b %d %Y")
            if kind == "dMy":
                return datetime.strptime(f"{g[0]} {g[1][:3]} {g[2]}", "%d %b %Y")
        except ValueError:
            continue     # e.g. 2026-13-45: keep looking
    return None


def extract_report(text, fallback_host, score_patterns=SCORE_PATTERNS, dayfirst=False):
    """Pull host, benchmark, date and score out of report text.

    Returns (info, notes). info is None when no usable score was found.
    """
    notes = []
    score = first(score_patterns, text)
    if score is None:
        return None, ["no score found. Run with --show-text to see the text, then use --score-regex"]
    try:
        value = float(score)
    except ValueError:
        return None, [f"score '{score}' isn't a number. Check --score-regex"]
    if not 0 <= value <= 100:
        return None, [f"score {value} is outside 0-100, so it was not recorded"]

    host = first(HOST_PATTERNS, text)
    if not host:
        host = fallback_host
        notes.append(f"no host name found in PDF, used the file name '{host}'")
    bench_m = re.search(BENCH_PATTERN, text, re.IGNORECASE)
    bench = re.sub(r"\s+", " ", bench_m.group(1)).strip() if bench_m else "Unknown benchmark"
    if not bench_m:
        notes.append("no benchmark name found in PDF")
    rdate = parse_date(first(DATE_LABELS, text), dayfirst=dayfirst)
    return {"host": host, "bench": bench, "date": rdate, "score": round(value, 2)}, notes


# ---------------------------------------------------------------- quarters

def fiscal_quarter(d, fy_start_month):
    """FY is named by the year it ends in (e.g. FY starting Oct 2025 = FY26)."""
    if not 1 <= fy_start_month <= 12:
        raise CisError("--fy-start-month must be between 1 and 12.")
    offset = (d.month - fy_start_month) % 12
    q = offset // 3 + 1
    fy = d.year + (1 if fy_start_month != 1 and d.month >= fy_start_month else 0)
    return f"FY{fy % 100:02d} Q{q}"


def normalize_quarter(label):
    """'fy26 q2', 'FY2026-Q2', "FY'26 Q2" -> 'FY26 Q2'. Returns None if it isn't a quarter."""
    m = QUARTER_RE.match(str(label or ""))
    if not m:
        return None
    return f"FY{int(m.group(1)) % 100:02d} Q{m.group(2)}"


def quarter_sort_key(label):
    n = normalize_quarter(label)
    if not n:
        return (1, 0, 0, str(label))          # non-quarter columns (e.g. Notes) go last
    return (0, int(n[2:4]), int(n[-1]), "")


def to_number(v):
    """Read a score already in the sheet: 85, 85.0, '85', '85%'. None if it isn't a number."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().rstrip("%").strip())
    except ValueError:
        return None


# ---------------------------------------------------------------- workbook

def open_workbook(path):
    if path.exists():
        try:
            wb = load_workbook(path)
        except PermissionError:
            raise CisError(f"Couldn't open {path}. Close it in Excel and run again.")
        except Exception as e:
            raise CisError(f"Couldn't open {path} as an Excel workbook ({e}).")
    else:
        wb = Workbook()
        wb.active.title = SCORES_SHEET
    if SCORES_SHEET not in wb.sheetnames:
        wb.create_sheet(SCORES_SHEET, 0)
    ws = wb[SCORES_SHEET]
    a1, b1 = ws.cell(1, 1).value, ws.cell(1, 2).value
    if a1 is None and b1 is None and ws.max_row == 1:
        ws.cell(1, 1, "Host")
        ws.cell(1, 2, "Benchmark")
    elif str(a1).strip().lower() != "host" or str(b1).strip().lower() != "benchmark":
        raise CisError(f"The '{SCORES_SHEET}' sheet must start with columns 'Host' and 'Benchmark' "
                       f"(found '{a1}', '{b1}'). Use a new workbook or rename those headers.")
    if LOG_SHEET not in wb.sheetnames:
        wb.create_sheet(LOG_SHEET).append(LOG_HEADER)
    return wb


def load_table(ws):
    """Return (columns, table) where table is {(host, bench): {column: value}}.

    Quarter headers are normalized so 'fy26 q2' and 'FY26 Q2' are one column.
    Non-quarter columns (e.g. Notes) are kept as they are.
    """
    header = [c.value for c in ws[1]]
    cols, src = [], []
    for h in header[2:]:
        if h is None or str(h).strip() == "":
            src.append(None)
            continue
        name = normalize_quarter(h) or str(h).strip()
        src.append(name)
        if name not in cols:
            cols.append(name)
    table = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or (r[0] in (None, "") and (len(r) < 2 or r[1] in (None, ""))):
            continue
        key = (str(r[0]).strip(), str(r[1]).strip() if len(r) > 1 and r[1] is not None else "")
        row = table.setdefault(key, {})
        for name, v in zip(src, r[2:]):
            if name and v is not None and row.get(name) is None:
                row[name] = v
    return cols, table


def safe_text(cell):
    """Store PDF-derived text as plain text, never as a formula."""
    if cell.data_type == "f":
        cell.data_type = "s"


def save_workbook(wb, path):
    backup = None
    if path.exists():
        backup = path.with_name(f"{path.stem}.backup-{datetime.now():%Y%m%d-%H%M%S}{path.suffix}")
        shutil.copy2(path, backup)
    try:
        wb.save(path)
    except PermissionError:
        raise CisError(f"Couldn't save {path}. Close it in Excel and run again.")
    return backup


# ---------------------------------------------------------------- main flow

def collect_files(inputs):
    files, missing = [], []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            files += sorted(p for p in path.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")
        elif path.is_file():
            files.append(path)
        else:
            missing.append(str(path))
    return files, missing


def process(files, wb, fy_start_month=1, quarter=None, score_patterns=SCORE_PATTERNS,
            overwrite=False, dayfirst=False, split_on=DEFAULT_SPLIT):
    ws, log = wb[SCORES_SHEET], wb[LOG_SHEET]
    cols, table = load_table(ws)
    if quarter is not None:
        q = normalize_quarter(quarter)
        if not q:
            raise CisError(f'--quarter must look like "FY26 Q2", got "{quarter}".')
        quarter = q

    changed, problems, results = set(), [], []
    seen_this_run = {}     # (key, quarter) -> report date, to keep the newest of two scans
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    reports = []     # (label for messages and log, fallback host, text, file)
    for f in files:
        try:
            pages = pdf_pages(f)
        except Exception as e:
            problems.append(f"{f.name}: could not read PDF ({e})")
            continue
        parts = split_reports(pages, split_on)
        if len(parts) > 1:
            print(f"{f.name}: {len(pages)} pages, {len(parts)} reports found", flush=True)
            for page, text in parts:
                reports.append((f"{f.name} p.{page}", f"{f.stem} p.{page}", text, f))
        else:
            reports.append((f.name, f.stem, parts[0][1], f))

    for label, fallback_host, text, f in reports:
        info, notes = extract_report(text, fallback_host, score_patterns, dayfirst=dayfirst)
        problems += [f"{label}: {n}" for n in notes]
        if info is None:
            continue

        rdate = info["date"]
        if quarter:
            q = quarter
        else:
            if rdate is None:
                rdate = datetime.fromtimestamp(f.stat().st_mtime)
                problems.append(f"{label}: no report date found in PDF, used the file's modified date "
                                f"{rdate.date()}. Use --quarter if that's wrong.")
            q = fiscal_quarter(rdate, fy_start_month)

        if q not in cols:
            cols.append(q)
        key = (info["host"], info["bench"])
        row = table.setdefault(key, {})
        current = to_number(row.get(q))
        value = info["score"]

        slot = (key, q)
        if slot in seen_this_run:
            # Two reports for the same host/benchmark/quarter in this run: keep the newest scan.
            prev = seen_this_run[slot]
            if rdate is not None and (prev is None or rdate > prev):
                action = "updated (newer scan in this run)" if current != value else "unchanged"
                row[q] = value
                seen_this_run[slot] = rdate
                changed.add(slot)
            else:
                action = "skipped (older or same-date scan already read this run)"
        elif row.get(q) is not None and current is None:
            action = f"skipped (cell has '{row.get(q)}', not a number)"
        elif current is not None and current == value:
            action = "unchanged"
            seen_this_run[slot] = rdate
        elif current is not None and not overwrite:
            action = f"skipped (already {current:g}; use --overwrite)"
        else:
            action = "updated" if current is not None else "added"
            row[q] = value
            seen_this_run[slot] = rdate
            if action == "updated":
                changed.add(slot)
        results.append({"file": label, "host": info["host"], "bench": info["bench"],
                        "date": rdate, "quarter": q, "score": value, "action": action})
        log.append([now, label, info["host"], info["bench"],
                    rdate.strftime("%Y-%m-%d") if rdate else "", q, value, action])
        for c in log[log.max_row][:4]:
            safe_text(c)

    # Rewrite the scores sheet once: quarters in order, other columns after, rows sorted.
    cols.sort(key=quarter_sort_key)
    idx = wb.sheetnames.index(SCORES_SHEET)
    widths = {k: v.width for k, v in ws.column_dimensions.items()}
    wb.remove(ws)
    ws = wb.create_sheet(SCORES_SHEET, idx)
    ws.append(["Host", "Benchmark"] + cols)
    for key in sorted(table, key=lambda k: (k[0].lower(), k[1].lower())):
        ws.append(list(key) + [table[key].get(c) for c in cols])
        r = ws.max_row
        safe_text(ws.cell(r, 1))
        safe_text(ws.cell(r, 2))
        for j, c in enumerate(cols, start=3):
            cell = ws.cell(r, j)
            if isinstance(cell.value, (int, float)) and normalize_quarter(c):
                cell.number_format = "0.0"
            if (key, c) in changed:
                cell.fill = CHANGED

    for sheet in (ws, log):
        for c in sheet[1]:
            c.font = BOLD
        sheet.freeze_panes = "A2"
        for col_cells in sheet.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            letter = col_cells[0].column_letter
            sheet.column_dimensions[letter].width = max(min(width + 2, 60), widths.get(letter) or 0)
    return results, problems


def main(argv=None):
    p = argparse.ArgumentParser(description="Record CIS benchmark PDF scores by fiscal quarter.")
    p.add_argument("inputs", nargs="+", help="PDF files and/or folders containing PDFs")
    p.add_argument("--workbook", default="CIS_Scores.xlsx", help="Tracking spreadsheet (created if missing)")
    p.add_argument("--fy-start-month", type=int, default=1,
                   help="Month your fiscal year starts, 1-12 (1 = calendar year, 7 = July, 10 = October)")
    p.add_argument("--quarter", help='Force all files into one quarter, e.g. "FY26 Q2"')
    p.add_argument("--score-regex", help="Your own regex with one group capturing the score number")
    p.add_argument("--dayfirst", action="store_true", help="Dates like 03/04/2026 in PDFs are DD/MM/YYYY")
    p.add_argument("--overwrite", action="store_true", help="Replace a score already recorded for that quarter")
    p.add_argument("--dry-run", action="store_true", help="Show what would change without saving")
    p.add_argument("--show-text", action="store_true", help="Print the extracted PDF text and stop")
    p.add_argument("--split-on", default=DEFAULT_SPLIT,
                   help='Word that starts each report in a multi-report PDF (default "Compliant"; '
                        '"" = one report per PDF)')
    a = p.parse_args(argv)

    try:
        if not 1 <= a.fy_start_month <= 12:
            raise CisError("--fy-start-month must be between 1 and 12.")
        score_patterns = list(SCORE_PATTERNS)
        if a.score_regex:
            try:
                rx = re.compile(a.score_regex)
            except re.error as e:
                raise CisError(f"--score-regex isn't a valid pattern ({e}).")
            if rx.groups < 1:
                raise CisError("--score-regex needs one (group) around the number, e.g. "
                               r'"Compliance:\s*(\d+(?:\.\d+)?)%"')
            score_patterns.insert(0, a.score_regex)

        files, missing = collect_files(a.inputs)
        for m in missing:
            print(f"NOTE: not found: {m}")
        if not files:
            raise CisError("No PDF files found.")

        if a.show_text:
            for f in files:
                try:
                    print(f"===== {f.name} =====\n{pdf_text(f)}\n")
                except Exception as e:
                    print(f"===== {f.name} =====\n(could not read: {e})\n")
            return 0

        wb_path = Path(a.workbook)
        if wb_path.suffix.lower() != ".xlsx":
            raise CisError("--workbook must be an .xlsx file.")
        wb = open_workbook(wb_path)
        results, problems = process(files, wb, a.fy_start_month, a.quarter, score_patterns,
                                    a.overwrite, a.dayfirst, a.split_on.strip())

        print(f"{'File':<34} {'Host':<24} {'Quarter':<9} {'Score':>6}  Action")
        for r in results:
            name = r["file"] if len(r["file"]) <= 34 else "..." + r["file"][-31:]   # keep the page number
            print(f"{name:<34} {r['host'][:24]:<24} {r['quarter']:<9} {r['score']:>6.1f}  {r['action']}")
        for msg in problems:
            print("NOTE:", msg)

        if a.dry_run:
            print("\nDry run: nothing saved.")
        elif not results:
            print("\nNothing to record, workbook not changed.")
        else:
            backup = save_workbook(wb, wb_path)
            print(f"\nSaved {wb_path}" + (f"  (backup: {backup.name})" if backup else ""))
    except CisError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())

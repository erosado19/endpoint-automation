import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


@pytest.fixture
def make_pdf(tmp_path):
    """Create a simple text PDF. Needs reportlab (test-only dependency)."""
    canvas = pytest.importorskip("reportlab.pdfgen.canvas")

    def _make(name, lines, folder=None, password=None):
        """lines: text lines; a "\\f" line starts a new page (pages also break when full)."""
        folder = Path(folder or tmp_path)
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / name
        c = canvas.Canvas(str(path))
        y = 800
        for ln in lines:
            if ln == "\f" or y < 50:
                c.showPage()
                y = 800
                if ln == "\f":
                    continue
            c.drawString(50, y, ln)
            y -= 18
        c.save()
        if password:
            from pypdf import PdfReader, PdfWriter
            w = PdfWriter()
            for pg in PdfReader(str(path)).pages:
                w.add_page(pg)
            w.encrypt(password)
            with open(path, "wb") as fh:
                w.write(fh)
        return path

    return _make


def report_lines(host, date, score, bench="CIS Microsoft Windows 11 Enterprise Benchmark v3.0.0"):
    return [bench, f"Target Hostname: {host}", f"Assessment Date: {date}", f"Total Score: {score}%"]


def compliant_report(host, date, score, pages=1, bench="CIS Microsoft Windows 11 Enterprise Benchmark v3.0.0"):
    """One report in the multi-report layout: starts with 'Compliant', may run over several pages."""
    lines = ["\f", f"Compliant: {score}%", bench, f"Hostname: {host}", f"Scan Date: {date}",
             "Section 1 Account Policies  Score: 40.0%"]
    for p in range(pages - 1):
        lines += ["\f"] + [f"Check {p}.{i}  Ensure setting {i} is configured  Non-Compliant" for i in range(30)]
    return lines

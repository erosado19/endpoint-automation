import zipfile
from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

import cis_scores_to_excel as cis
from cis_scores_to_excel import CisError, fiscal_quarter, normalize_quarter, parse_date, extract_report
from conftest import report_lines


def scores(path):
    ws = load_workbook(path)[cis.SCORES_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    return rows[0], {(r[0], r[1]): dict(zip(rows[0][2:], r[2:])) for r in rows[1:]}


# ------------------------------------------------------------ quarters

@pytest.mark.parametrize("date,start,expected", [
    ("2025-10-01", 10, "FY26 Q1"), ("2025-12-31", 10, "FY26 Q1"), ("2026-01-01", 10, "FY26 Q2"),
    ("2026-06-30", 10, "FY26 Q3"), ("2026-09-30", 10, "FY26 Q4"), ("2026-10-01", 10, "FY27 Q1"),
    ("2026-01-15", 1, "FY26 Q1"), ("2026-12-31", 1, "FY26 Q4"),
    ("2025-07-01", 7, "FY26 Q1"), ("2026-06-30", 7, "FY26 Q4"),
    ("2099-03-01", 4, "FY99 Q4"), ("2099-04-01", 4, "FY00 Q1"),   # FY2100 shows as FY00
])
def test_fiscal_quarter(date, start, expected):
    assert fiscal_quarter(datetime.fromisoformat(date), start) == expected


@pytest.mark.parametrize("bad", [0, 13, -1])
def test_fiscal_quarter_bad_month(bad):
    with pytest.raises(CisError):
        fiscal_quarter(datetime(2026, 1, 1), bad)


@pytest.mark.parametrize("label,expected", [
    ("FY26 Q2", "FY26 Q2"), ("fy26 q2", "FY26 Q2"), ("FY2026-Q2", "FY26 Q2"), ("FY'26 Q2", "FY26 Q2"),
    (" FY26Q2 ", "FY26 Q2"), ("FY26 Q5", None), ("Notes", None), (None, None), (2026, None),
])
def test_normalize_quarter(label, expected):
    assert normalize_quarter(label) == expected


# ------------------------------------------------------------ parsing text

@pytest.mark.parametrize("raw,expected", [
    ("2025-11-04", "2025-11-04"),
    ("2025-11-04T10:42:00Z", "2025-11-04"),
    ("February 12, 2026 10:42 AM", "2026-02-12"),
    ("Feb 12 2026", "2026-02-12"),
    ("Sept. 3, 2026", "2026-09-03"),
    ("12 February 2026", "2026-02-12"),
    ("01/20/2026", "2026-01-20"),
    ("20/01/2026", "2026-01-20"),     # day > 12 so it must be day first
    ("2026-13-45 then 2026-02-01", "2026-02-01"),
    ("no date here", None), ("", None), (None, None),
])
def test_parse_date(raw, expected):
    got = parse_date(raw)
    assert (got.date().isoformat() if got else None) == expected


def test_parse_date_dayfirst():
    assert parse_date("03/04/2026", dayfirst=True) == datetime(2026, 4, 3)
    assert parse_date("03/04/2026") == datetime(2026, 3, 4)


def test_total_score_preferred_over_earlier_section_score():
    text = "Section 1 Score: 40%\nCIS Foo Benchmark v1.0\nHostname: WS-1\nTotal Score: 88.5%"
    info, _ = extract_report(text, "x")
    assert info["score"] == 88.5


def test_summary_table_row():
    info, _ = extract_report("Computer Name: SRV-01\nTotal 210 30 87.5%", "x")
    assert info["score"] == 87.5 and info["host"] == "SRV-01"


@pytest.mark.parametrize("text", ["Total Score: 150%", "Total Score: 999%", "nothing useful"])
def test_bad_scores_not_recorded(text):
    info, notes = extract_report(text, "x")
    assert info is None and notes


def test_update_label_is_not_mistaken_for_date():
    info, _ = extract_report("Last Update: 2020-01-01\nTotal Score: 80%\nHostname: A", "x")
    assert info["date"] is None


def test_missing_host_falls_back_to_file_name_with_note():
    info, notes = extract_report("Total Score: 80%", "ws-77_report")
    assert info["host"] == "ws-77_report"
    assert any("file name" in n for n in notes)


# ------------------------------------------------------------ end to end with real PDFs

def run(args):
    return cis.main([str(a) for a in args])


def test_add_then_rerun_is_unchanged(tmp_path, make_pdf):
    d = tmp_path / "r"
    make_pdf("a.pdf", report_lines("WS-001", "2025-11-04", "82.5"), d)
    make_pdf("b.pdf", report_lines("WS-001", "February 12, 2026", "88.1"), d)
    wb = tmp_path / "s.xlsx"
    assert run([d, "--workbook", wb, "--fy-start-month", "10"]) == 0
    header, data = scores(wb)
    assert header == ("Host", "Benchmark", "FY26 Q1", "FY26 Q2")
    key = ("WS-001", "CIS Microsoft Windows 11 Enterprise Benchmark v3.0.0")
    assert data[key] == {"FY26 Q1": 82.5, "FY26 Q2": 88.1}

    assert run([d, "--workbook", wb, "--fy-start-month", "10"]) == 0
    assert scores(wb)[1] == data
    log = list(load_workbook(wb)[cis.LOG_SHEET].iter_rows(min_row=2, values_only=True))
    assert [r[-1] for r in log[-2:]] == ["unchanged", "unchanged"]
    assert list(tmp_path.glob("s.backup-*.xlsx"))    # backup made before second save


def test_conflict_needs_overwrite(tmp_path, make_pdf):
    wb = tmp_path / "s.xlsx"
    run([make_pdf("a.pdf", report_lines("WS-1", "2026-02-01", "80")), "--workbook", wb])
    p2 = make_pdf("a2.pdf", report_lines("WS-1", "2026-02-15", "90"))
    run([p2, "--workbook", wb])
    assert list(scores(wb)[1].values())[0]["FY26 Q1"] == 80
    run([p2, "--workbook", wb, "--overwrite"])
    assert list(scores(wb)[1].values())[0]["FY26 Q1"] == 90


@pytest.mark.parametrize("order", [("old", "new"), ("new", "old")])
def test_two_scans_same_quarter_newest_wins(tmp_path, make_pdf, order):
    d = tmp_path / "r"
    files = {"old": report_lines("WS-1", "2026-01-05", "70"), "new": report_lines("WS-1", "2026-03-20", "95")}
    # file names sort in the requested order
    make_pdf(f"1_{order[0]}.pdf", files[order[0]], d)
    make_pdf(f"2_{order[1]}.pdf", files[order[1]], d)
    wb = tmp_path / "s.xlsx"
    assert run([d, "--workbook", wb]) == 0
    assert list(scores(wb)[1].values())[0]["FY26 Q1"] == 95


def test_bad_files_reported_not_fatal(tmp_path, make_pdf, capsys):
    d = tmp_path / "r"
    make_pdf("good.pdf", report_lines("WS-1", "2026-02-01", "80"), d)
    make_pdf("noscore.pdf", ["just some text"], d)
    (d / "fake.pdf").write_text("this is not a pdf", encoding="utf-8")
    make_pdf("locked.pdf", report_lines("WS-2", "2026-02-01", "70"), d, password="secret")
    (d / "notes.txt").write_text("ignore me", encoding="utf-8")
    wb = tmp_path / "s.xlsx"
    assert run([d, "--workbook", wb]) == 0
    out = capsys.readouterr().out
    assert "noscore.pdf: no score found" in out
    assert "fake.pdf: could not read PDF" in out
    assert "locked.pdf: could not read PDF" in out
    assert "notes.txt" not in out
    assert [k[0] for k in scores(wb)[1]] == ["WS-1"]


def test_uppercase_extension_found(tmp_path, make_pdf):
    d = tmp_path / "r"
    make_pdf("A.PDF", report_lines("WS-1", "2026-02-01", "80"), d)
    assert run([d, "--workbook", tmp_path / "s.xlsx"]) == 0


def test_existing_sheet_with_messy_headers_and_text_scores(tmp_path, make_pdf):
    wb_path = tmp_path / "s.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = cis.SCORES_SHEET
    bench = "CIS Microsoft Windows 11 Enterprise Benchmark v3.0.0"
    ws.append(["Host", "Benchmark", "Notes", "fy26 q2", "FY26 Q1"])
    ws.append(["WS-1", bench, "owner: IT", None, "85%"])
    wb.save(wb_path)

    run([make_pdf("a.pdf", report_lines("WS-1", "2025-11-01", "85")), "--workbook", wb_path,
         "--fy-start-month", "10"])
    header, data = scores(wb_path)
    assert header == ("Host", "Benchmark", "FY26 Q1", "FY26 Q2", "Notes")
    assert data[("WS-1", bench)]["Notes"] == "owner: IT"
    assert data[("WS-1", bench)]["FY26 Q1"] == "85%"     # '85%' read as 85 -> unchanged, left as typed


def test_wrong_layout_refused(tmp_path, make_pdf, capsys):
    wb_path = tmp_path / "s.xlsx"
    wb = Workbook()
    wb.active.title = cis.SCORES_SHEET
    wb.active.append(["Server", "Score"])
    wb.save(wb_path)
    assert run([make_pdf("a.pdf", report_lines("WS-1", "2026-02-01", "80")), "--workbook", wb_path]) == 2
    assert "must start with columns" in capsys.readouterr().err


def test_dry_run_changes_nothing(tmp_path, make_pdf):
    wb = tmp_path / "s.xlsx"
    assert run([make_pdf("a.pdf", report_lines("WS-1", "2026-02-01", "80")), "--workbook", wb, "--dry-run"]) == 0
    assert not wb.exists()


def test_forced_quarter(tmp_path, make_pdf):
    wb = tmp_path / "s.xlsx"
    run([make_pdf("a.pdf", report_lines("WS-1", "2026-02-01", "80")), "--workbook", wb, "--quarter", "fy27 q3"])
    assert scores(wb)[0][2] == "FY27 Q3"


def test_missing_date_uses_file_time_with_note(tmp_path, make_pdf, capsys):
    p = make_pdf("a.pdf", ["Hostname: WS-1", "Total Score: 80%"])
    assert run([p, "--workbook", tmp_path / "s.xlsx"]) == 0
    assert "modified date" in capsys.readouterr().out


def test_custom_score_regex(tmp_path, make_pdf):
    p = make_pdf("a.pdf", ["Hostname: WS-1", "Assessment Date: 2026-02-01", "Compliance: 77.7 percent"])
    wb = tmp_path / "s.xlsx"
    assert run([p, "--workbook", wb, "--score-regex", r"Compliance:\s*(\d+(?:\.\d+)?)\s*percent"]) == 0
    assert list(scores(wb)[1].values())[0]["FY26 Q1"] == 77.7


@pytest.mark.parametrize("extra", [
    ["--fy-start-month", "13"], ["--quarter", "Q2 2026"], ["--score-regex", "Score: \\d+"],
    ["--score-regex", "(unclosed"], ["--workbook", "scores.csv"],
])
def test_bad_options_return_2(tmp_path, make_pdf, extra):
    p = make_pdf("a.pdf", report_lines("WS-1", "2026-02-01", "80"))
    args = [p, "--workbook", tmp_path / "s.xlsx"] + extra
    assert run(args) == 2


def test_no_pdfs_returns_2(tmp_path):
    assert run([tmp_path, "--workbook", tmp_path / "s.xlsx"]) == 2


def test_formula_like_file_name_stored_as_text(tmp_path, make_pdf):
    p = make_pdf("=HYPERLINK(1).pdf", ["Total Score: 80%", "Assessment Date: 2026-02-01"])
    wb = tmp_path / "s.xlsx"
    assert run([p, "--workbook", wb]) == 0
    with zipfile.ZipFile(wb) as z:
        for name in z.namelist():
            if name.startswith("xl/worksheets/"):
                assert "<f>" not in z.read(name).decode("utf-8"), name
    assert scores(wb)[1].popitem()[0][0] == "=HYPERLINK(1)"

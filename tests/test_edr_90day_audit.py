import zipfile

import pandas as pd
import pytest
from openpyxl import load_workbook

import edr_90day_audit as audit
from edr_90day_audit import AuditError, split_identity, parse_dates, detect_dayfirst, run_audit

AS_OF = pd.Timestamp("2026-11-15")


def frame(rows, cols=("Email", "Hostname", "Last Logon")):
    return pd.DataFrame(rows, columns=list(cols)).fillna("")


def status_of(result, account):
    u = result["users"]
    return u.loc[u["Account"] == account, "Audit Status"].iloc[0]


# ------------------------------------------------------------ identities

@pytest.mark.parametrize("value,expected", [
    ("jdoe@Example.com", ("jdoe", "example.com")),
    ("  JDoe@EXAMPLE.COM  ", ("jdoe", "example.com")),
    ("Jdoe@Nycexample", ("jdoe", "nycexample")),
    ("NYC\\JDoe", ("jdoe", "nyc")),
    ("jdoe", ("jdoe", "")),
    ("John Doe <jdoe@example.com>", ("jdoe", "example.com")),
    ('"jdoe@example.com"', ("jdoe", "example.com")),
    ("", ("", "")),
    ("   ", ("", "")),
    ("@example.com", ("", "")),          # no username -> invalid
    ("jdoe@", ("jdoe", "")),
    (None, ("", "")),
    (float("nan"), ("", "")),
])
def test_split_identity(value, expected):
    assert split_identity(value) == expected


def test_same_name_different_tenants_stay_separate_and_case_merges():
    df = frame([
        ["jdoe@example.com", "WS-1", "2026-11-10"],
        ["JDoe@Example.com", "WS-2", "2026-03-02"],
        ["Jdoe@Nycexample", "NYC-1", "2026-06-01"],
    ])
    r = run_audit(df, "Email", "Last Logon", AS_OF, host_col="Hostname")
    u = r["users"].set_index("Account")
    assert set(u.index) == {"jdoe@example.com", "jdoe@nycexample"}
    assert u.loc["jdoe@example.com", "Rows in export"] == 2
    assert u.loc["jdoe@example.com", "Devices"] == "WS-1, WS-2"
    assert u.loc["jdoe@example.com", "Audit Status"] == "Active"        # newest date wins
    assert u.loc["jdoe@nycexample", "Audit Status"] == "Inactive (>90d)"


def test_newest_date_wins_regardless_of_row_order():
    df = frame([["a@x.com", "", "2026-01-01"], ["a@x.com", "", "2026-11-01"], ["a@x.com", "", "2026-05-01"]])
    r = run_audit(df, "Email", "Last Logon", AS_OF)
    assert r["users"]["Days Inactive"].iloc[0] == 14


def test_export_columns_named_domain_and_account_do_not_collide():
    df = pd.DataFrame({
        "Account": ["jdoe@example.com", "asmith@example.com"],
        "Domain": ["EXAMPLE-TENANT", "EXAMPLE-TENANT"],
        "Last Seen": ["2026-11-01", "2026-01-01"],
    })
    r = run_audit(df, "Account", "Last Seen", AS_OF, tenant_col="Domain")
    assert list(r["users"]["Account"]) == ["asmith@example.com", "jdoe@example.com"]
    assert list(r["users"]["Domain"]) == ["example.com", "example.com"]   # from the email, not the column
    assert list(r["raw"]["Domain"]) == ["EXAMPLE-TENANT", "EXAMPLE-TENANT"]  # original column untouched
    assert "As shown in export" in r["users"].columns


def test_tenant_column_used_when_no_domain():
    df = pd.DataFrame({"User": ["jdoe", "jdoe", "jdoe"], "Tenant": ["Example", "NYC", "Example"],
                       "Last Logon": ["2026-11-01", "2026-01-01", "2026-02-01"]})
    r = run_audit(df, "User", "Last Logon", AS_OF, tenant_col="Tenant")
    assert set(r["users"]["Account"]) == {"jdoe@example", "jdoe@nyc"}
    assert not any("no domain" in w for w in r["warnings"])


def test_warns_when_no_domain_and_no_tenant():
    df = pd.DataFrame({"User": ["jdoe"], "Last Logon": ["2026-11-01"]})
    r = run_audit(df, "User", "Last Logon", AS_OF)
    assert any("no domain" in w for w in r["warnings"])
    assert r["users"]["Domain"].iloc[0] == "(no domain)"


def test_blank_users_skipped_and_counted():
    df = frame([["", "WS-1", "2026-11-01"], ["   ", "WS-2", "2026-11-01"], ["a@x.com", "WS-3", "2026-11-01"]])
    r = run_audit(df, "Email", "Last Logon", AS_OF)
    assert len(r["users"]) == 1
    assert any("2 row(s) had no user" in w for w in r["warnings"])


# ------------------------------------------------------------ thresholds

@pytest.mark.parametrize("days_ago,expected", [
    (0, "Active"), (89, "Active"), (90, "Active"), (91, "Inactive (>90d)"), (400, "Inactive (>90d)"),
])
def test_threshold_boundary(days_ago, expected):
    d = (AS_OF - pd.Timedelta(days=days_ago)).date().isoformat()
    r = run_audit(frame([["a@x.com", "", d]]), "Email", "Last Logon", AS_OF)
    assert status_of(r, "a@x.com") == expected


def test_warn_days():
    df = frame([["a@x.com", "", "2026-09-01"], ["b@x.com", "", "2026-11-01"]])
    r = run_audit(df, "Email", "Last Logon", AS_OF, warn_days=60)
    assert status_of(r, "a@x.com") == "Warning (>60d)"
    assert status_of(r, "b@x.com") == "Active"
    assert len(r["review"]) == 0      # warnings are not in "Needs review"


@pytest.mark.parametrize("days,warn", [(0, None), (-5, None), (90, 90), (90, 120), (90, 0)])
def test_bad_threshold_options_rejected(days, warn):
    with pytest.raises(AuditError):
        run_audit(frame([["a@x.com", "", "2026-01-01"]]), "Email", "Last Logon", AS_OF, days=days, warn_days=warn)


def test_future_date_warns():
    r = run_audit(frame([["a@x.com", "", "2027-01-01"]]), "Email", "Last Logon", AS_OF)
    assert any("after the audit date" in w for w in r["warnings"])


# ------------------------------------------------------------ dates

@pytest.mark.parametrize("raw,expected", [
    ("2026-11-10", "2026-11-10"),
    ("2026-11-10T23:59:59Z", "2026-11-10"),
    ("2026-11-10T22:00:00-05:00", "2026-11-11"),      # converted to UTC
    ("2026-11-10 08:00:00", "2026-11-10"),
    ("Nov 10, 2026", "2026-11-10"),
    ("10 November 2026", "2026-11-10"),
    ("133442208000000000", "2023-11-12"),             # Windows FILETIME
    ("1762732800000", "2025-11-10"),                  # epoch milliseconds
    ("1762732800", "2025-11-10"),                     # epoch seconds
    ("", None), ("Never", None), ("N/A", None), ("0", None),
    ("1601-01-01", None),                             # AD "never logged on" placeholder
])
def test_parse_dates(raw, expected):
    got = parse_dates(pd.Series([raw]))[0]
    if expected is None:
        assert pd.isna(got)
    else:
        assert got == pd.Timestamp(expected)


def test_dayfirst_detection():
    assert detect_dayfirst(["25/03/2026", "01/04/2026"])[0] is True
    assert detect_dayfirst(["03/25/2026", "04/01/2026"])[0] is False
    dayfirst, warn = detect_dayfirst(["03/04/2026"])
    assert dayfirst is False and "either way" in warn
    with pytest.raises(AuditError):
        detect_dayfirst(["25/03/2026", "03/25/2026"])


def test_dayfirst_applied_to_parsing():
    df = frame([["a@x.com", "", "01/04/2026"], ["b@x.com", "", "25/03/2026"]])
    r = run_audit(df, "Email", "Last Logon", AS_OF)
    assert r["dayfirst"] is True
    u = r["users"].set_index("Account")
    assert str(u.loc["a@x.com", "Last Activity"]) == "2026-04-01"


def test_unreadable_dates_reported():
    r = run_audit(frame([["a@x.com", "", "garbage"]]), "Email", "Last Logon", AS_OF)
    assert status_of(r, "a@x.com") == "No date"
    assert any("couldn't be read" in w for w in r["warnings"])


# ------------------------------------------------------------ exclusions

def test_exclusions(tmp_path):
    f = tmp_path / "ex.txt"
    f.write_text("# service accounts\nsvc_backup@example.com\nbreakglass\n\n", encoding="utf-8")
    df = frame([
        ["svc_backup@example.com", "", "2025-01-01"],
        ["svc_backup@nycexample", "", "2025-01-01"],
        ["breakglass@example.com", "", "2025-01-01"],
        ["breakglass@nycexample", "", "2025-01-01"],
        ["jdoe@example.com", "", "2025-01-01"],
    ])
    r = run_audit(df, "Email", "Last Logon", AS_OF, exclude=audit.load_exclusions(f))
    assert set(r["users"]["Account"]) == {"svc_backup@nycexample", "jdoe@example.com"}
    assert len(r["excluded"]) == 3


def test_missing_exclude_file():
    with pytest.raises(AuditError):
        audit.load_exclusions("does-not-exist.txt")


# ------------------------------------------------------------ input files

@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-16", "cp1252"])
def test_read_csv_encodings(tmp_path, encoding):
    p = tmp_path / "e.csv"
    p.write_text("Email,Last Logon\njosé@example.com,2026-11-01\n", encoding=encoding)
    df = audit.read_csv_any(p)
    assert list(df.columns) == ["Email", "Last Logon"]
    assert df["Email"][0] == "josé@example.com"


def test_empty_and_missing_files(tmp_path):
    p = tmp_path / "empty.csv"
    p.write_text("", encoding="utf-8")
    with pytest.raises(AuditError):
        audit.read_csv_any(p)
    with pytest.raises(AuditError):
        audit.read_csv_any(tmp_path / "nope.csv")


def test_find_column_ignores_case_and_spaces():
    df = pd.DataFrame(columns=["  E-Mail ", "LAST SEEN", "Device Name"])
    df.columns = [c.strip() for c in df.columns]
    assert audit.find_column(df, None, audit.USER_COL_GUESSES, "user") == "E-Mail"
    assert audit.find_column(df, "last seen ", [], "date") == "LAST SEEN"
    with pytest.raises(AuditError):
        audit.find_column(df, "Nope", [], "date")
    with pytest.raises(AuditError):
        audit.find_column(pd.DataFrame(columns=["X"]), None, audit.DATE_COL_GUESSES, "date")


def test_email_column_preferred_over_username():
    df = pd.DataFrame(columns=["User Name", "Email", "Last Logon"])
    assert audit.find_column(df, None, audit.USER_COL_GUESSES, "user") == "Email"


def test_header_only_csv_produces_empty_report(tmp_path):
    src = tmp_path / "h.csv"
    src.write_text("Email,Last Logon\n", encoding="utf-8")
    out = tmp_path / "out.xlsx"
    assert audit.main([str(src), "--as-of", "2026-11-15", "--out", str(out)]) == 0
    assert load_workbook(out)["All users"].max_row == 1


# ------------------------------------------------------------ output safety

def test_formula_injection_is_neutralised(tmp_path):
    src = tmp_path / "evil.csv"
    src.write_text(
        'Email,Hostname,Last Logon\n'
        '"=HYPERLINK(""http://evil.example"",""click"")@x.com","=1+1",2026-01-01\n'
        'ok@x.com,"+cmd|\' /C calc\'!A0",2026-01-01\n',
        encoding="utf-8")
    out = tmp_path / "out.xlsx"
    assert audit.main([str(src), "--as-of", "2026-11-15", "--out", str(out)]) == 0
    with zipfile.ZipFile(out) as z:
        for name in z.namelist():
            if name.startswith("xl/worksheets/"):
                assert "<f>" not in z.read(name).decode("utf-8"), name
    ws = load_workbook(out)["Raw export"]
    assert ws["B2"].value == "=1+1" and ws["B2"].data_type == "s"


def test_cli_errors_return_2(tmp_path, capsys):
    src = tmp_path / "a.csv"
    src.write_text("Email,Last Logon\na@x.com,2026-11-01\n", encoding="utf-8")
    assert audit.main([str(src), "--as-of", "11/15/2026"]) == 2
    assert audit.main([str(src), "--out", str(src)]) == 2
    assert audit.main([str(src), "--user-col", "Nope"]) == 2
    assert audit.main([str(tmp_path / "missing.csv")]) == 2
    assert "ERROR" in capsys.readouterr().err


def test_cli_end_to_end(tmp_path):
    src = tmp_path / "t.csv"
    src.write_text(
        "Email,Hostname,Last Logon,OS\n"
        "jdoe@example.com,WS-001,2026-11-10T14:22:00Z,Win11\n"
        "Jdoe@Nycexample,NYC-001,2026-06-01,Win11\n"
        "mlee@nycexample,NYC-004,,Win11\n",
        encoding="utf-8")
    assert audit.main([str(src), "--as-of", "2026-11-15"]) == 0
    out = tmp_path / "t_user_audit_2026-11-15.xlsx"
    wb = load_workbook(out)
    assert wb.sheetnames == ["Summary", "By domain", "Needs review", "All users", "Raw export"]
    review = [r[0] for r in wb["Needs review"].iter_rows(min_row=2, values_only=True)]
    assert review == ["mlee@nycexample", "jdoe@nycexample"]
    by_dom = {r[0]: r[1] for r in wb["By domain"].iter_rows(min_row=2, values_only=True)}
    assert by_dom == {"nycexample": 2, "example.com": 1}
    rules = [r.formula[0] for cf in wb["All users"].conditional_formatting for r in cf.rules]
    assert any("Inactive" in f for f in rules) and any("Warning" in f for f in rules)

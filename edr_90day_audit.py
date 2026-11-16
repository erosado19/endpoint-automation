"""
90-day user audit: read a CSV exported from the EDR console and flag user
accounts whose most recent activity is older than N days (default 90).

Each account is identified by username + domain (multi-tenant safe):
    jdoe@example.com and JDoe@Example.com   -> same account
    jdoe@example.com and jdoe@nycexample    -> different accounts
A user can show up on several rows (one per device or login). The script
keeps the most recent date for each account.

Works only on the CSV you already export. It makes no API calls or network requests.

Usage:
    python edr_90day_audit.py export.csv
    python edr_90day_audit.py export.csv --user-col "Email" --date-col "Last Logon"
    python edr_90day_audit.py export.csv --as-of 2026-09-30 --exclude service_accounts.txt
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Column names commonly used in exports. Matching ignores case and spaces at the ends.
# Email-style columns come first: in a multi-tenant EDR the domain is what
# tells jdoe@example.com apart from jdoe@nycexample.com.
USER_COL_GUESSES = [
    "email", "email address", "e-mail", "mail", "upn", "userprincipalname",
    "user principal name", "user email",
    "user", "username", "user name", "user_name", "account", "account name",
    "accountname", "samaccountname",
    "logged in user", "last logged in user", "last logged on user", "last user",
    "login", "logon name",
]
TENANT_COL_GUESSES = [
    "tenant", "tenant name", "customer", "customer name", "organization",
    "organisation", "org", "site", "site name", "cid", "domain", "domain name",
]
DATE_COL_GUESSES = [
    "last logon", "last login", "last logon time", "last login time", "lastlogon",
    "lastlogontimestamp", "last sign-in", "last sign in", "last activity",
    "last active", "last seen", "last_seen", "lastseen", "last seen (utc)",
    "last check-in", "last checkin", "last contact", "last contact time",
]
HOST_COL_GUESSES = [
    "hostname", "host name", "host", "device name", "devicename",
    "computer name", "computername", "endpoint name", "device",
]

ENCODINGS = ["utf-8-sig", "utf-16", "cp1252"]
EARLIEST_REAL_DATE = pd.Timestamp("1990-01-01")   # older = "never" placeholders like 1601-01-01
FILETIME_EPOCH_OFFSET = 116444736000000000         # 100-ns ticks between 1601-01-01 and 1970-01-01
SLASH_DATE = re.compile(r"^\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\b")

RED = PatternFill("solid", fgColor="F7C1C1")
AMBER = PatternFill("solid", fgColor="FAEEDA")
BOLD = Font(bold=True)


class AuditError(Exception):
    """A problem the user can fix (bad column name, bad option, unreadable file)."""


# ---------------------------------------------------------------- reading input

def read_csv_any(path):
    """Read a CSV, trying the encodings EDR consoles and Excel commonly produce."""
    path = Path(path)
    if not path.is_file():
        raise AuditError(f"File not found: {path}")
    last = None
    for enc in ENCODINGS:
        try:
            df = pd.read_csv(path, dtype=str, encoding=enc, keep_default_na=False)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except UnicodeError as e:
            last = e
        except pd.errors.EmptyDataError:
            raise AuditError(f"{path.name} is empty.")
    raise AuditError(f"Couldn't read {path.name} as text ({last}). Re-export it as CSV (UTF-8).")


def find_column(df, wanted, guesses, kind, required=True):
    lower = {c.lower().strip(): c for c in df.columns}
    if wanted:
        key = wanted.strip().lower()
        if key in lower:
            return lower[key]
        raise AuditError(f"Column '{wanted}' not found. Columns in this file: {list(df.columns)}")
    for g in guesses:
        if g in lower:
            return lower[g]
    if not required:
        return None
    raise AuditError(
        f"Couldn't find the {kind} column automatically.\n"
        f"Columns in this file: {list(df.columns)}\n"
        f"Pass it with --{kind}-col \"<column name>\""
    )


# ---------------------------------------------------------------- identities

ANGLE = re.compile(r"<([^<>]+)>")


def split_identity(value):
    """Return (username, domain) in lowercase. Returns ("", "") for blank or invalid values.

    jdoe@Example.com               -> ("jdoe", "example.com")
    "John Doe <jdoe@example.com>"  -> ("jdoe", "example.com")
    NYC\\JDoe                      -> ("jdoe", "nyc")
    jdoe                           -> ("jdoe", "")
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "", ""
    v = str(value).strip()
    m = ANGLE.search(v)
    if m:
        v = m.group(1)
    v = v.strip().strip('"').strip().lower()
    if "@" in v:
        user, dom = v.rsplit("@", 1)
    elif "\\" in v:
        dom, user = v.split("\\", 1)
    else:
        user, dom = v, ""
    user, dom = user.strip(), dom.strip()
    if not user:
        return "", ""
    return user, dom


def account_key(user, dom):
    return f"{user}@{dom}" if dom else user


# ---------------------------------------------------------------- dates

def detect_dayfirst(values):
    """Look at slash-style dates (03/04/2026) and decide day-first vs month-first.

    Returns (dayfirst, warning). Raises AuditError if the file mixes both styles.
    """
    first_big = second_big = ambiguous = 0
    for v in values:
        m = SLASH_DATE.match(v)
        if not m:
            continue
        a, b = int(m.group(1)), int(m.group(2))
        if a > 12:
            first_big += 1
        elif b > 12:
            second_big += 1
        else:
            ambiguous += 1
    if first_big and second_big:
        raise AuditError("The date column mixes DD/MM/YYYY and MM/DD/YYYY values. "
                         "Re-export with ISO dates (YYYY-MM-DD) if the console allows it.")
    if first_big:
        return True, None
    if second_big:
        return False, None
    if ambiguous:
        return False, (f"{ambiguous} date(s) like 03/04/2026 could be read either way. "
                       f"Assumed month first (US). Use --dayfirst if your export is DD/MM/YYYY.")
    return False, None


def parse_dates(series, dayfirst=False):
    """Parse a column of mixed date formats into naive, midnight-normalized timestamps.

    Handles ISO, US/EU slash dates, text dates, timezone suffixes, Unix epoch
    seconds/milliseconds and Windows FILETIME (lastLogonTimestamp). Blanks,
    "Never", 0 and placeholder dates before 1990 become NaT (no date).
    """
    s = series.fillna("").astype(str).str.strip()
    # Microsecond resolution: nanoseconds can't hold dates before 1677, and AD
    # exports use 1601-01-01 for "never logged on".
    result = pd.Series(pd.NaT, index=s.index, dtype="datetime64[us]")

    digits = s.str.fullmatch(r"\d{9,18}")
    if digits.any():
        n = pd.to_numeric(s[digits], errors="coerce").astype("float64")
        secs = pd.Series(float("nan"), index=n.index)
        secs[n >= 1e17] = (n[n >= 1e17] - FILETIME_EPOCH_OFFSET) / 1e7      # FILETIME
        ms = (n >= 1e11) & (n < 1e14)
        secs[ms] = n[ms] / 1000                                              # epoch ms
        sec = (n >= 1e8) & (n < 1e11)
        secs[sec] = n[sec]                                                   # epoch s
        result[digits] = pd.to_datetime(secs, unit="s", errors="coerce").astype("datetime64[us]")

    text = ~digits & (s != "")
    if text.any():
        parsed = pd.to_datetime(s[text], errors="coerce", format="mixed", utc=True, dayfirst=dayfirst)
        result[text] = parsed.dt.tz_convert(None).astype("datetime64[us]")

    result = result.dt.normalize()
    result[result < EARLIEST_REAL_DATE] = pd.NaT
    return result


# ---------------------------------------------------------------- the audit

def load_exclusions(path):
    if not path:
        return set(), set()
    p = Path(path)
    if not p.is_file():
        raise AuditError(f"Exclude file not found: {p}")
    exact, names = set(), set()
    for ln in p.read_text(encoding="utf-8-sig").splitlines():
        ln = ln.strip()
        if not ln or ln.startswith("#"):
            continue
        user, dom = split_identity(ln)
        if not user:
            continue
        if dom:
            exact.add(account_key(user, dom))
        else:
            names.add(user)
    return exact, names


def run_audit(df, user_col, date_col, as_of, days=90, warn_days=None,
              host_col=None, tenant_col=None, exclude=(set(), set()), dayfirst=None):
    """Build the audit tables. Returns a dict of DataFrames plus info for the summary."""
    if days < 1:
        raise AuditError("--days must be at least 1.")
    if warn_days is not None and not (0 < warn_days < days):
        raise AuditError("--warn-days must be greater than 0 and less than --days.")

    work = df.copy()
    total_rows = len(work)
    warnings = []

    # All helper columns start with "__" so they can never clash with export columns
    # such as "Domain" or "Account".
    parts = work[user_col].map(split_identity)
    work["__name"] = parts.str[0]
    work["__domain"] = parts.str[1]
    blank_user = work["__name"] == ""
    skipped_blank = int(blank_user.sum())
    work = work[~blank_user]

    if tenant_col:
        no_dom = work["__domain"] == ""
        work.loc[no_dom, "__domain"] = work.loc[no_dom, tenant_col].fillna("").astype(str).str.strip().str.lower()
    work["__key"] = [account_key(u, d) for u, d in zip(work["__name"], work["__domain"])]
    no_domain_rows = int((work["__domain"] == "").sum())

    exact, names = exclude
    ex_mask = work["__key"].isin(exact) | work["__name"].isin(names)
    excluded = sorted(set(work.loc[ex_mask, "__key"]))
    work = work[~ex_mask]

    if dayfirst is None:
        dayfirst, warn = detect_dayfirst(work[date_col].astype(str))
        if warn:
            warnings.append(warn)
    work["__date"] = parse_dates(work[date_col], dayfirst=dayfirst)
    unparsed = work[date_col].astype(str).str.strip().ne("") & work["__date"].isna()
    if unparsed.any():
        sample = ", ".join(sorted(set(work.loc[unparsed, date_col].astype(str)))[:3])
        warnings.append(f"{int(unparsed.sum())} date value(s) couldn't be read and count as 'No date' "
                        f"(e.g. {sample}).")

    future = work["__date"] > as_of
    if future.any():
        warnings.append(f"{int(future.sum())} row(s) have a date after the audit date. Check --as-of.")

    g = work.groupby("__key", sort=False)
    users = pd.DataFrame({
        "Account": list(g.groups.keys()),
    }).set_index("Account", drop=False)
    users["Domain"] = g["__domain"].first()
    users["__date"] = g["__date"].max()
    users["Rows in export"] = g.size()
    if host_col:
        users["Devices"] = g[host_col].agg(
            lambda s: ", ".join(sorted({str(x).strip() for x in s if str(x).strip()})))
    users["As shown in export"] = g[user_col].first()
    users = users.reset_index(drop=True)

    users["Last Activity"] = users["__date"].dt.date
    users["Days Inactive"] = (as_of - users["__date"]).dt.days.astype("Int64")

    def status(d):
        if pd.isna(d):
            return "No date"
        if d > days:
            return f"Inactive (>{days}d)"
        if warn_days is not None and d > warn_days:
            return f"Warning (>{warn_days}d)"
        return "Active"

    users["Audit Status"] = users["Days Inactive"].map(status).astype(str)
    users["Domain"] = users["Domain"].replace("", "(no domain)")
    users = users.sort_values(["Days Inactive", "Account"], ascending=[False, True],
                              na_position="first", kind="stable")
    cols = ["Account", "Domain", "Last Activity", "Days Inactive", "Audit Status", "Rows in export"]
    cols += ["Devices"] if host_col else []
    cols += ["As shown in export"]
    users = users[cols].reset_index(drop=True)
    review = users[users["Audit Status"].str.startswith(("Inactive", "No date"))].reset_index(drop=True)

    statuses = ["Active"] + ([f"Warning (>{warn_days}d)"] if warn_days is not None else []) + \
               [f"Inactive (>{days}d)", "No date"]
    if len(users):
        by_domain = pd.crosstab(users["Domain"], users["Audit Status"])
    else:
        by_domain = pd.DataFrame(index=pd.Index([], name="Domain"))
    by_domain = by_domain.reindex(columns=statuses, fill_value=0)
    by_domain.insert(0, "Total users", by_domain.sum(axis=1))
    by_domain = by_domain.sort_values(["Total users"], ascending=False, kind="stable").reset_index()

    counts = users["Audit Status"].value_counts().reindex(statuses, fill_value=0)
    if no_domain_rows:
        warnings.append(f"{no_domain_rows} row(s) have no domain (and no tenant value). Accounts with the "
                        f"same name in different tenants could be merged. Use an email column for "
                        f"--user-col, or --tenant-col.")
    if skipped_blank:
        warnings.append(f"{skipped_blank} row(s) had no user value and were skipped.")

    summary = pd.DataFrame({
        "Item": ["Audit date", "Threshold (days)", "Rows in export", "Rows skipped (no user)",
                 "Accounts excluded", "Domains", "Total accounts"]
                + [f"  {k}" for k in counts.index],
        "Value": [as_of.date().isoformat(), days, total_rows, skipped_blank,
                  len(excluded), int(users["Domain"].nunique()), len(users)]
                 + [int(v) for v in counts.values],
    })
    raw = work.drop(columns=[c for c in work.columns if c.startswith("__")])
    return {"summary": summary, "by_domain": by_domain, "review": review, "users": users,
            "raw": raw, "counts": counts, "warnings": warnings, "excluded": excluded,
            "dayfirst": dayfirst}


# ---------------------------------------------------------------- writing output

def defuse_formulas(ws):
    """Make sure text that looks like a formula (=HYPERLINK(...)) is stored as plain text.

    Export values come from many tenants; a username or hostname starting with '='
    must never become a live formula when someone opens the report.
    """
    for row in ws.iter_rows():
        for c in row:
            if c.data_type == "f":
                c.data_type = "s"


def autosize(ws, sample_rows=1000):
    """Size columns from the first rows only, so big exports stay fast."""
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row, sample_rows + 1)):
        for c in row:
            if c.value is not None:
                widths[c.column] = max(widths.get(c.column, 0), len(str(c.value)))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(w + 2, 60)


def write_report(out, result, source_name):
    summary = result["summary"].copy()
    summary.loc[-1] = ["Source file", source_name]
    summary = summary.sort_index().reset_index(drop=True)
    sheets = [("Summary", summary), ("By domain", result["by_domain"]),
              ("Needs review", result["review"]), ("All users", result["users"]),
              ("Raw export", result["raw"])]
    if result["excluded"]:
        sheets.append(("Excluded", pd.DataFrame({"Account": result["excluded"]})))
    if result.get("notice") is not None:
        sheets.append(("Deletion notice", result["notice"]))
    if len(result["raw"]) > 1_048_000:
        raise AuditError("The export has more rows than Excel allows on one sheet.")
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            for name, frame in sheets:
                frame.to_excel(xw, sheet_name=name, index=False)
            for name, ws in xw.sheets.items():
                defuse_formulas(ws)
                for c in ws[1]:
                    c.font = BOLD
                ws.freeze_panes = "A2"
                if name in ("Needs review", "All users") and ws.max_row > 1:
                    # Conditional formatting: one rule instead of styling every cell,
                    # and the colors follow the rows when you sort or filter.
                    col = get_column_letter(list(result["users"].columns).index("Audit Status") + 1)
                    rng = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
                    ws.conditional_formatting.add(rng, FormulaRule(
                        formula=[f'OR(LEFT(${col}2,8)="Inactive",${col}2="No date")'], fill=RED))
                    ws.conditional_formatting.add(rng, FormulaRule(
                        formula=[f'LEFT(${col}2,7)="Warning"'], fill=AMBER))
                if name != "Summary" and ws.max_row > 1:
                    ws.auto_filter.ref = ws.dimensions
                autosize(ws)
    except PermissionError:
        raise AuditError(f"Couldn't write {out}. Close it in Excel and run again.")


# ---------------------------------------------------------------- deletion notice

def notice_table(result, delete_on):
    """Accounts that would be deleted on delete_on: everyone Inactive or with No date."""
    u = result["users"]
    t = u[u["Audit Status"].str.startswith(("Inactive", "No date"))]
    t = t[["Account", "Domain", "Last Activity", "Days Inactive", "Audit Status"]].copy()
    t["Scheduled deletion"] = delete_on.date()
    return t.reset_index(drop=True)


def write_notice(path, notice, as_of, delete_on, days, source_name):
    """Write a plain-text notice listing the accounts scheduled for deletion.

    Only writes a file. The script never deletes or changes any account.
    """
    when = f"{delete_on.date().isoformat()} ({delete_on.strftime('%A')})"
    lines = [
        "ACCOUNT DELETION NOTICE",
        "=" * 60,
        f"Audit date:      {as_of.date().isoformat()}",
        f"Deletion date:   {when}",
        f"Source export:   {source_name}",
        "",
        f"The user accounts below have not signed in for more than {days} days.",
        f"They will be deleted on {when} unless the user",
        "signs in before that date.",
        "",
        f"If you still need your account, sign in before {delete_on.date().isoformat()}",
        "or contact IT.",
        "",
    ]

    def table(rows):
        out = []
        width = max([len("Account")] + [len(a) for a in rows["Account"]]) + 2
        out.append(f"  {'Account':<{width}}{'Last sign-in':<14}Days inactive")
        out.append(f"  {'-' * (width - 2):<{width}}{'-' * 12:<14}{'-' * 13}")
        for _, r in rows.iterrows():
            last = str(r["Last Activity"]) if pd.notna(r["Days Inactive"]) else "unknown"
            d = str(r["Days Inactive"]) if pd.notna(r["Days Inactive"]) else "-"
            out.append(f"  {r['Account']:<{width}}{last:<14}{d}")
        return out

    inactive = notice[notice["Audit Status"].str.startswith("Inactive")]
    unknown = notice[notice["Audit Status"] == "No date"]
    if inactive.empty and unknown.empty:
        lines.append("No accounts are scheduled for deletion.")
    for dom, rows in inactive.groupby("Domain", sort=True):
        lines.append(f"== {dom} ({len(rows)} account{'s' if len(rows) != 1 else ''}) ==")
        lines += table(rows) + [""]
    if not unknown.empty:
        lines.append(f"== Last sign-in unknown ({len(unknown)}): confirm these before deleting ==")
        lines += table(unknown) + [""]
    lines += ["-" * 60,
              f"Total accounts scheduled for deletion: {len(notice)}",
              "This file is a notice only. No accounts were changed or deleted."]
    try:
        Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    except PermissionError:
        raise AuditError(f"Couldn't write {path}. Close it and run again.")


def parse_ymd(value, option):
    try:
        return pd.Timestamp(datetime.strptime(value, "%Y-%m-%d"))
    except ValueError:
        raise AuditError(f"{option} must look like 2026-09-30, got '{value}'.")


# ---------------------------------------------------------------- command line

def main(argv=None):
    p = argparse.ArgumentParser(description="Flag user accounts not active in N days.")
    p.add_argument("csv", help="CSV exported from the EDR console")
    p.add_argument("--days", type=int, default=90, help="Inactivity threshold in days (default 90)")
    p.add_argument("--user-col", help="User column, ideally email/UPN (auto-detected if omitted)")
    p.add_argument("--date-col", help="Last-activity column (auto-detected if omitted)")
    p.add_argument("--host-col", help="Optional hostname column, listed for context")
    p.add_argument("--tenant-col", help="Optional tenant column, used when a user value has no domain")
    p.add_argument("--as-of", help="Audit date YYYY-MM-DD (default: today)")
    p.add_argument("--warn-days", type=int, default=None, help="Also mark accounts getting close, e.g. 60")
    p.add_argument("--dayfirst", action="store_true", help="Dates like 03/04/2026 are DD/MM/YYYY")
    p.add_argument("--exclude", help="Text file of accounts to skip, one per line")
    p.add_argument("--out", help="Output .xlsx (default: <csv name>_user_audit_<date>.xlsx)")
    p.add_argument("--notice", action="store_true",
                   help="Also write a text file listing accounts to be deleted if the user doesn't sign in")
    p.add_argument("--delete-on", help="Deletion date for the notice, YYYY-MM-DD (implies --notice)")
    p.add_argument("--grace-days", type=int, default=14,
                   help="Days from the audit date to the deletion date when --delete-on isn't given (default 14)")
    a = p.parse_args(argv)

    try:
        as_of = parse_ymd(a.as_of, "--as-of") if a.as_of else pd.Timestamp(datetime.now().date())
        delete_on = None
        if a.notice or a.delete_on:
            if a.delete_on:
                delete_on = parse_ymd(a.delete_on, "--delete-on")
            elif a.grace_days < 1:
                raise AuditError("--grace-days must be at least 1.")
            else:
                delete_on = as_of + pd.Timedelta(days=a.grace_days)
            if delete_on <= as_of:
                raise AuditError(f"--delete-on ({delete_on.date()}) must be after the audit date ({as_of.date()}).")

        src = Path(a.csv)
        df = read_csv_any(src)
        user_col = find_column(df, a.user_col, USER_COL_GUESSES, "user")
        date_col = find_column(df, a.date_col, DATE_COL_GUESSES, "date")
        host_col = find_column(df, a.host_col, HOST_COL_GUESSES, "host", required=False)
        tenant_col = find_column(df, a.tenant_col, TENANT_COL_GUESSES, "tenant", required=False)
        if tenant_col == user_col:
            tenant_col = None

        result = run_audit(df, user_col, date_col, as_of, days=a.days, warn_days=a.warn_days,
                           host_col=host_col, tenant_col=tenant_col,
                           exclude=load_exclusions(a.exclude),
                           dayfirst=True if a.dayfirst else None)

        if delete_on is not None:
            result["notice"] = notice_table(result, delete_on)

        out = Path(a.out) if a.out else src.with_name(f"{src.stem}_user_audit_{as_of.date().isoformat()}.xlsx")
        if out.resolve() == src.resolve():
            raise AuditError("--out can't be the same file as the input CSV.")
        write_report(out, result, src.name)
        notice_path = None
        if delete_on is not None:
            notice_path = out.with_name(f"{src.stem}_deletion_notice_{as_of.date().isoformat()}.txt")
            write_notice(notice_path, result["notice"], as_of, delete_on, a.days, src.name)
    except AuditError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2

    print(f"User column: '{user_col}'   Date column: '{date_col}'   "
          f"Host column: '{host_col or 'none'}'   Tenant column: '{tenant_col or 'none'}'   "
          f"As of: {as_of.date()}   Dates read {'day' if result['dayfirst'] else 'month'} first")
    print(f"  {'Domains':<20} {result['by_domain'].shape[0]}")
    for k, v in result["counts"].items():
        print(f"  {k:<20} {v}")
    if result["excluded"]:
        print(f"  {'Excluded accounts':<20} {len(result['excluded'])}")
    for w in result["warnings"]:
        print("NOTE:", w)
    print(f"Report written to {out}")
    if notice_path:
        print(f"Deletion notice written to {notice_path}  "
              f"({len(result['notice'])} account(s), deletion date {delete_on.date()})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

# Endpoint automation scripts

Both scripts run locally on files you already have. They make no API calls or
network requests and don't touch the EDR itself.

## Setup (one time)

```
pip install pandas openpyxl pypdf
```

If you can't install packages, ask IT to approve these three. They're widely used
and come from PyPI.

---

## 1. 90-day user audit: `edr_90day_audit.py`

Export the user list from the EDR console as CSV as usual, then:

```
python edr_90day_audit.py edr_export.csv
```

The script finds the user and last-activity columns automatically. If it can't,
it lists the columns so you can name them yourself:

```
python edr_90day_audit.py edr_export.csv --user-col "User Name" --date-col "Last Logon"
```

How users are counted (multi-tenant safe):
- Each account is identified by **username plus domain**, ignoring capitalization.
  - `jdoe@example.com` and `JDoe@Example.com` are the same person.
  - `jdoe@example.com` and `Jdoe@Nycexample` are **different** people.
  - `NYC\jdoe` uses `nyc` as the domain.
- The script looks for an email or UPN column first. If your export has both a
  username column and an email column, it picks the email column.
- If a value has no domain at all (just `jdoe`), the script uses a tenant or
  customer column if there is one (or `--tenant-col "Tenant"`). If there isn't,
  it prints a warning, because the same name in two tenants would be merged.
- When an account appears on several rows (several devices), the most recent
  date is used, so activity on any machine counts as active.
- If the export has a hostname column, each account's devices are listed for context.

Useful options:

| Option | What it does |
|---|---|
| `--days 90` | Inactivity threshold (default 90) |
| `--warn-days 60` | Also highlight users getting close |
| `--as-of 2026-09-30` | Audit as of a specific date, e.g. quarter end |
| `--exclude exclude.txt` | Skip service or break-glass accounts listed one per line |
| `--tenant-col "Tenant"` | Tenant column to use when user values have no domain |
| `--dayfirst` | Dates like 03/04/2026 are DD/MM/YYYY (the script detects this when it can) |
| `--out audit.xlsx` | Choose the output file name |

### Deletion notice

Add `--notice` to also write a text file listing the accounts that will be deleted
if the user doesn't sign in:

```
python edr_90day_audit.py edr_export.csv --delete-on 2026-12-01
python edr_90day_audit.py edr_export.csv --notice --grace-days 30
```

| Option | What it does |
|---|---|
| `--notice` | Write the notice. The deletion date is 14 days after the audit date |
| `--delete-on 2026-12-01` | Use this exact deletion date (turns on `--notice`) |
| `--grace-days 30` | Days from the audit date to the deletion date (default 14) |

The notice (`<csv name>_deletion_notice_<date>.txt`) lists the inactive accounts
grouped by domain, with each one's last sign-in and days inactive. Accounts with
no date are listed separately under "Last sign-in unknown", so you can check them
before deleting. The same list is added to the Excel report as a **Deletion notice**
sheet. **The script never deletes or changes any account. It only writes the notice.**

### Exclude file

In the exclude file, `svc_backup@example.com` skips only that tenant's account, while a
bare `svc_backup` skips that name in every tenant. Lines starting with `#` are ignored.

The output Excel file has these sheets:
- **Summary**: counts for your audit record
- **By domain**: active, inactive and no-date counts for each tenant or domain
- **Needs review**: users inactive more than 90 days, plus users with no date (highlighted)
- **All users**: one row per account (with its domain), most inactive first
- **Raw export**: the original rows, for evidence
- **Excluded**: the accounts your exclude file removed (only when you use `--exclude`)
- **Deletion notice**: accounts scheduled for deletion and the date (only with `--notice`)

Notes on dates:
- "Inactive" means **more than** 90 days, so exactly 90 days still counts as active.
- The script reads ISO dates, US and EU slash dates, text dates, timezone suffixes,
  Unix timestamps and Windows FILETIME values (AD `lastLogonTimestamp`).
- "Never", blanks, 0 and placeholder dates such as 1601-01-01 count as **No date**
  and appear in Needs review.
- Anything the script can't read is listed in a NOTE when it finishes, so nothing is
  silently dropped.

Speed: a few thousand rows take seconds. About 200,000 rows take around 2 minutes,
most of it spent writing the Excel file.

---

## 2. CIS benchmark scores: `cis_scores_to_excel.py`

Put the quarter's CIS report PDFs in a folder, then:

```
python cis_scores_to_excel.py "C:\path\to\cis_reports" --workbook CIS_Scores.xlsx --fy-start-month 10
```

Set `--fy-start-month` to the month your fiscal year starts (1 = January, 7 = July,
10 = October). Fiscal years are named by the year they end in, so Oct 2025 to Sep 2026 is FY26.

The script reads the score, host, benchmark name and report date from each report,
works out the fiscal quarter, and writes the score in the matching column.

**PDFs with many reports:** a single PDF can hold any number of reports (35 reports
and 500+ pages take a few seconds). Each report starts with the word **Compliant**,
and the script finds them the way Ctrl+F does. The PDF is not changed or split on
disk. Each report is logged with its page number (e.g. `batch.pdf p.37`) so you can
check it. If the start word ever changes, use `--split-on "New Word"`.
"Non-Compliant" never counts as a report start. Anything before the first
"Compliant", such as a cover page, is ignored.

The result looks like this:

| Host | Benchmark | FY26 Q1 | FY26 Q2 |
|---|---|---|---|
| WS-001 | CIS Microsoft Windows 11 Enterprise Benchmark v3.0.0 | 82.5 | 88.1 |

Safety features:
- Before every save, a backup copy (`CIS_Scores.backup-<date-time>.xlsx`) is made
  next to the workbook.
- Rerunning on the same PDFs changes nothing.
- An existing score is never overwritten unless you add `--overwrite`.
- If two reports for the same host and quarter are in one run, the newest scan wins.
- Scores outside 0-100 are rejected, never recorded.
- Every file read is recorded on the **Import log** sheet.
- `--dry-run` shows what would change without saving.
- Quarter headers you typed differently (`fy26 q2`, `FY2026-Q2`) are merged into
  `FY26 Q2`. Extra columns such as `Notes` are kept and moved to the end.
- If the sheet doesn't start with `Host` and `Benchmark`, the script stops rather than guess.

Note: the script rebuilds the **CIS Scores** sheet each run. Keep charts, formulas
and pivot tables on a separate sheet that refers to it. The Excel library used
can't preserve charts in a workbook it rewrites.

**If a PDF says "no score found":** your report layout differs from what the
script expects. Run

```
python cis_scores_to_excel.py report.pdf --show-text
```

to see the text the script reads, and find the line with the score. Then either
pass a pattern such as `--score-regex "Compliance\s*:\s*(\d+(?:\.\d+)?)%"`, or share
that line (with hostnames removed) and the pattern can be added to the script.

Use `--quarter "FY26 Q2"` to put every file in a given quarter regardless of
the dates inside them.

---

## Running the tests

The `tests` folder holds 143 automated checks covering multi-tenant names,
date formats, thresholds, bad files, encrypted PDFs, formula injection and
CLI errors. To run them (they need `pytest` and `reportlab`, for testing only):

```
pip install pytest reportlab
python -m pytest tests -q
```

Run them again after any change to the scripts.

**Tip:** Run the script on a copy of your existing spreadsheet first. If your
current sheet uses a different layout, the script may need adjusting to match it.

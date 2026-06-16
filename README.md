# ABSA Home Loan Extra Screen Update

This script downloads ABSA Home Loan Legal Excel files from FTP and updates Legal Suite extra screens using the `Account Number` as `Matter.TheirRef`.

Script: [absa_home_loan_extrascreen_update.py](/home/lordwiz/Documents/Peter's%20Works/Iconis%20Stuff/Legal%20Suite/ABSA%20Write%20Back%20Feedback/absa_home_loan_extrascreen_update.py:1)

## What It Does

- Downloads the ABSA `Comments` and `PTP` workbooks for a given day from FTP.
- Groups rows by `Account Number`.
- Fetches the Legal Suite matter by `Matter.TheirRef`.
- If multiple matters are found for the same account, only matters with `FileRef` starting with `A0038/` or `ABS697/` are eligible.
- Updates the correct Legal Suite extra screen:
  - `553` for `Comments`
  - `552` for `PTP`
- Writes a JSON run report under `downloads_absa/_reports/`.

## Credentials

The script reads these values from `.env` by default:

- `FTP_HOST`
- `FTP_USER`
- `FTP_PASS`
- `LEGALSUITE_API_KEY`

## Comments Mapping

The `Comments` workbook uses the Legal Suite memo form layout:

- `field1` = `Account Number`
- Latest row 1:
  - `field2` = `Comment / Memo`
  - `field3` = `Number Dialled`
  - `field4` = `Date`
  - `field5` = blank
- Latest row 2:
  - `field6` = `Comment / Memo`
  - `field7` = `Number Dialled`
  - `field8` = `Date`
  - `field9` = blank
- Latest row 3:
  - `field10` = `Comment / Memo`
  - `field11` = `Number Dialled`
  - `field12` = `Date`
  - `field13` = blank
- Latest row 4:
  - `field14` = `Comment / Memo`
  - `field15` = `Number Dialled`
  - `field16` = `Date`
  - `field17` = blank
- Latest row 5:
  - `field18` = `Comment / Memo`
  - `field19` = `Number Dialled`
  - `field20` = `Date`
  - `field21` = blank
- Latest row 6:
  - `field22` = `Comment / Memo`
  - `field23` = `Number Dialled`
  - `field24` = `Date`
  - `field25` = blank
- `field26` = `Branch ID`

Selection rules:

- Only the latest 6 rows per `Account Number` are used.
- Rows are ordered by `Date` descending.
- If dates are tied, the lower row in Excel wins.

## PTP Mapping

The current PTP mapping is:

- `field1` = `Account Number`
- `field2` = `PTP Capture Date`
- `field3` = `PTP Due Date`
- `field4` = `PTP Amount`

## Usage

Run for today:

```bash
python3 absa_home_loan_extrascreen_update.py
```

Dry run:

```bash
python3 absa_home_loan_extrascreen_update.py --dry-run
```

Run for a specific date:

```bash
python3 absa_home_loan_extrascreen_update.py --date 20260520
```

Run for one day ago:

```bash
python3 absa_home_loan_extrascreen_update.py --day-1
```

Equivalent explicit form:

```bash
python3 absa_home_loan_extrascreen_update.py --days-ago 1
```

Only comments:

```bash
python3 absa_home_loan_extrascreen_update.py --only comments
```

Only PTP:

```bash
python3 absa_home_loan_extrascreen_update.py --only ptp
```

Skip FTP download and use existing local files:

```bash
python3 absa_home_loan_extrascreen_update.py --skip-download --date 20260520
```

Skip GET verification after update:

```bash
python3 absa_home_loan_extrascreen_update.py --no-verify
```

## Notes

- The script uses only the Python standard library plus `openpyxl`.
- No `requests` dependency is required.
- If no eligible matter matches the allowed `FileRef` prefixes, that account is skipped and recorded in the JSON report.
# absa-write-back-feedback

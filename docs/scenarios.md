# The four scenarios

Every command and output below is real: it comes from the run in [../demo/transcript.txt](../demo/transcript.txt), recorded with release candidate 1.0.0-rc1 of v1.0.0 (Pro package) on Windows 11 with Microsoft 365 Excel. Lines with local paths (`data:`, `output:`, `report:`, `log:`, `package:`, `folder:`, `golden:`) are left out here; they are in the transcript. All data is synthetic: the sales of an invented shop with three stores, *Central*, *Market* and *Port*.

Commands run from the Kit's folder with its single entry point, `bin/steadylatch-check.ps1`. Its [exit codes](exit-codes.md) are stable within 1.x.

## E1: an input column is renamed

**What the Kit does:** a data contract stops a bad input before the macro runs or anything is written; an empty required cell is reported as empty, never as 0. Runs in CI and locally.

The `Amount` column of the sales file has become `Amt`:

```text
> pwsh -File bin/steadylatch-check.ps1 -ContractPath demo/contracts/sales-csv.contract.json -ReportDirectory out/reports -DataPath demo/fixtures/e1/sales-e1a-renamed-column.csv
contract shop-sales-csv: Invalid (6 rows, 1 violations)
  - sales-e1a-renamed-column.csv: required column 'Amount' is missing. Columns found: OrderId, OrderDate, Store, Product, Quantity, UnitPrice, Amt, Discount, GiftWrap.
 EXIT 1  fail
```

Variant B, one amount is empty:

```text
  - sales-e1b-empty-amount.csv: row 3 (line 4), column 'Amount' is empty; this column requires a value.
 EXIT 1  fail
```

## E2: the API fails

**What the Kit does:** the client retries only what is worth retrying, within a time budget, never leaves a partial output, and network and contract problems get different exit codes. Runs in CI and locally.

A local mock on 127.0.0.1 serves the sales, or fails on purpose. The good run saves the file:

```text
api lumen-sales-good: Valid (GET http://127.0.0.1:18080/good/v1/sales?month=2026-09, 1 attempt(s))
  lumen-sales-good response: satisfies the contract shop-sales-json; written to sales-2026-09.json
 EXIT 0  pass
```

Then four failures:

| Case | Result | Exit |
|---|---|---|
| (a) Timeout | `gave up after 3 attempt(s) (retry.maxAttempts); last attempt: no complete response within 2 s` | 4 |
| (b) 429 with Retry-After of 2 s | waits, retries, `Valid (... 2 attempt(s))` | 0 |
| (c) HTML instead of JSON | `the body is not valid JSON at line 1: '<' is an invalid start of a value. (Content-Type: text/html; charset=utf-8)`, not retried | 4 |
| (d) Valid JSON that breaks the contract | `breaks the contract shop-sales-json (1 violations); nothing was written` | 1 |

Each attempt is one JSON line in the log, and the token is never written:

```text
Accept           Authorization X-Correlation-Id
------           ------------- ----------------
application/json [REDACTED]    51c7e16509874674bcb69dde6b5cab20
```

The saved file has the same SHA256 before and after the failures: no failing run changed it.

## E3: a macro changes

**What the Kit does:** the output of the changed macro is compared cell by cell with an approved golden. The macro runs locally in Excel, on a copy of the workbook; the comparison also runs in CI.

The original macro matches the golden:

```text
excel monthly-report: Succeeded (macro BuildMonthlyReport of lumen-monthly-report.xlsm)
  macro 'BuildMonthlyReport' of 'lumen-monthly-report.xlsm' ran on a copy; sheet 'Report' saved to monthly-report.xlsx
compare monthly-report-xlsx: Match: 0 differences (1 within tolerance)
 EXIT 0  pass
```

The copy with a bug differs in one line, a date filter that drops the last day of the month:

```text
-        If orderDate >= firstDay And orderDate <= lastDay Then
+        If orderDate >= firstDay And orderDate < lastDay Then
```

Its output against the same golden:

```text
compare monthly-report-xlsx: Mismatch: 8 differences (1 within tolerance)
  - row with Store=Central, column 'Orders' (cell B4): expected '4', found '3' (delta -1; no tolerance: exact match required).
  - row with Store=Central, column 'Units' (cell C4): expected '9', found '8' (delta -1; no tolerance: exact match required).
  - row with Store=Central, column 'Revenue' (cell D4): expected '92.71', found '60.71' (delta -32, outside the tolerance of ±0.005).
  - row with Store=Central, column 'AvgOrderValue' (cell F4): expected '23.18', found '20.239999999999998' (delta -2.940000000000002, outside the tolerance of ±0.005).
  - row with Store=Market, column 'Orders' (cell B5): expected '4', found '3' (delta -1; no tolerance: exact match required).
  - row with Store=Market, column 'Units' (cell C5): expected '13', found '11' (delta -2; no tolerance: exact match required).
  - row with Store=Market, column 'Revenue' (cell D5): expected '44.8', found '35.799999999999997' (delta -9.000000000000003, outside the tolerance of ±0.005).
  - row with Store=Market, column 'AvgOrderValue' (cell F5): expected '11.2', found '11.93' (delta 0.73, outside the tolerance of ±0.005).
 EXIT 1  fail
```

Expected, found, delta and tolerance per cell. The one difference within tolerance is listed, not counted as a regression.

## E4: a new version is delivered

**What the Kit does:** the VBA is exported to text with an environment manifest, packed with the workbook and its checksums, and the previous version is restored and checked with the runner and the golden. This is the delivery tooling of the Pro and Consultant licenses; packaging and restoring need no Excel.

Version 1.0.0 is packaged; then version 1.1.0, the workbook with the bug that E3 caught:

```text
vba lumen-monthly-report: Match: 3 component(s), 1 file(s) and the workbook 'lumen-monthly-report.xlsm' agree with the manifest (exported 2026-09-28T23:44:53Z)
package lumen-monthly-report-1.0.0: Succeeded
delivery lumen-monthly-report-1.0.0: Valid: 3 file(s) agree with SHA256SUMS and the VBA export of 'lumen-monthly-report.xlsm' agrees with its manifest
 EXIT 0  pass
```

Going back to 1.0.0 from its package, into a new folder:

```text
restore lumen-monthly-report-1.0.0: Succeeded
  restored 'lumen-monthly-report.xlsm' and 3 other file(s) of 'lumen-monthly-report-1.0.0.zip' to 'lumen-monthly-report-1.0.0' (workbook sha256 02dc67d328c4c4b4557a1df97d17959f3bb6c740032b85944f70ed5bb8d23a55)
 EXIT 0  pass
```

The restored workbook has the same SHA256 as the delivered one, and it runs in Excel and matches the approved golden again (`Match: 0 differences`). The Kit never overwrites a file: putting the restored version back in use is your step.

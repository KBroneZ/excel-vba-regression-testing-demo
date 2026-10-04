# Requirements, tested versions and limits

What the SteadyLatch Automation Reliability Kit v1.0.0 needs, where it was tested, and what it does not do. What was not tested says so.

## Requirements

- PowerShell 7.4 or later (`pwsh`). The module has no runtime dependencies. Windows PowerShell 5.1 is not supported: the module refuses to load there.
- Running macros (the Excel runner) and exporting or importing VBA need Windows with Excel desktop, with a person present. Never in hosted CI.
- Everything else, including checking a VBA export against its manifest and making, checking or restoring a delivery package, runs on any OS with PowerShell 7.4+.

## Tested versions

- Excel: Microsoft 365, version 16.0, build 20430.20092 (v1.0.0 and its release candidate), Current Channel, 64-bit, English UI, Spanish regional separators, on Windows 11 Pro 10.0.26200.
- PowerShell: 7.6.6 on that machine, and the PowerShell 7 preinstalled on GitHub's ubuntu-latest and windows-latest images (the Kit's CI). 7.4 is the minimum the module declares, not a version tested on its own.
- Not tested yet: Windows 10, 32-bit Excel, perpetual Office versions, other UI languages.

## What it does not do

- No running VBA in hosted CI (GitHub Actions and the like): Microsoft does not support automating Office without a person present.
- No Excel for Mac, Excel on the web, Office Scripts, Google Sheets or .xls.
- Values only: no comparing formats, charts, pivot tables or formulas.
- No API connector catalogue, interactive OAuth or writing to external systems: the client reads one URL with a token from an environment variable.
- Not a replacement for Rubberduck or a VBA unit-testing framework: the Kit checks results and complements them.
- No copy protection or technical license management: the license is a contract.
- No adapting the Kit to your workbook: that is a separate service.
- The Kit never changes security settings: not the Trust Center's macro settings, not "Trust access to the VBA project object model", not PowerShell's execution policy, and it never removes the "downloaded from the Internet" mark.
- CI compares what you commit: it cannot tell whether a committed output was produced by the committed workbook.

## Fixed limits

Enforced by the code: going over one is an error with a message, never a partial result.

- .xlsx/.xlsm file size up to 100 MB; up to 2,000,000 cells in one sheet.
- CSV, JSON and .xlsx files are read whole into memory: meant for hundreds to tens of thousands of rows.
- Excel run or VBA export time limit 1-3,600 s, default 120 s; up to 30 macro arguments, one macro per run.
- API: 1-10 attempts (default 3); Retry-After accepted up to a cap (default 60 s); https only, plain http only to this machine; redirects not followed.

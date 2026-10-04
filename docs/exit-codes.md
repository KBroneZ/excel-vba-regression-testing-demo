# Exit codes

Every run of `bin/steadylatch-check.ps1`, the Kit's single entry point, ends with one of these codes, in CI and locally. A code never changes meaning within 1.x.

| Code | Name | Meaning |
|---|---|---|
| 0 | Ok | Check passed, no differences outside tolerance and every contract satisfied. |
| 1 | CheckFailed | Regression found or data contract violated; a report explains what failed. |
| 2 | UsageError | Invalid usage or configuration. |
| 3 | EnvironmentUnavailable | Required environment not available, for example Excel desktop. |
| 4 | NetworkError | Network or API failure after bounded retries. |
| 5 | InternalError | Unexpected error inside the Kit (a bug). |

In the [demo](scenarios.md): a renamed column (E1) and a changed macro (E3) end with 1; an API timeout or an HTML page instead of JSON (E2) with 4; valid JSON that breaks the contract (E2) with 1, not 4.

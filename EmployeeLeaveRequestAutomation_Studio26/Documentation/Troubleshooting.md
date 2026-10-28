# Troubleshooting guide

## Package restore

If Excel or mail activities show as unknown types, open Manage Packages, select the Official feed, restore the package versions listed in `PackageVerification.md`, and reopen the project.

## Excel file errors

- Confirm the workbook exists and is not open with an exclusive lock.
- Confirm worksheet names are `LeaveRequests`, `Config`, and `LeaveResults` as applicable.
- Confirm the input header row has all 12 required columns.
- Use the separate output workbook; never point Write Range at the input workbook.

## Notification errors

`NotificationStatus` is independent from the leave decision. `Not Configured`, `Skipped`, and `Failed` are valid notification outcomes. Check provider TLS, OAuth/credential policy, and network access before increasing retries. Do not retry authentication failures indefinitely.

## Duplicate processing

The process uses an in-memory RequestID set during the run. A duplicate row is rejected and does not generate a second final decision for that RequestID. For unattended reruns, preserve and review the previous output workbook and introduce a durable idempotency store before production use.

## Build failures

Validate the failing XAML file first, then rebuild the project. Check argument names, package versions, namespace imports, and `InvokeCode` attribute encoding. This project has passed package restore, validation for all nine XAML files, project build, and headless runtime execution. A `Status=Error` row is a business-rule outcome for invalid input; it is not a workflow/runtime failure. Review `DecisionReason` in `Data/LeaveRequests_Output.xlsx` to distinguish business outcomes from execution errors.

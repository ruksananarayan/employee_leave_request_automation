# Smoke test results

## UiPath runtime smoke test

- Project: `EmployeeLeaveRequestAutomation_Studio26`
- Entry point: `Main.xaml`
- Execution: headless Studio/Helm via `uip rpa run --skip-build`
- Notifications: disabled by `Config.xlsx` (`NotificationEnabled=False`)
- Runtime verdict: passed (`Result=Success`, `hasErrors=False`, `errorMessage=null`, `debugState=null`)

Workflow log summary:

- Configuration: `OK - 16 configuration keys loaded`
- Input schema: `Schema OK`
- Processed: `12 requests`
- Errors: `1`
- Results: written
- Summary: written

## Output distribution

| Status | Count |
|---|---:|
| Approved | 3 |
| Rejected | 7 |
| Pending Manager Approval | 1 |
| Error | 1 |

## Notification distribution

| Notification status | Count |
|---|---:|
| Not Configured | 10 |
| Skipped - Pending Manager Approval | 1 |
| Skipped - Error status | 1 |
| Sent | 0 |
| Failed | 0 |

## Output evidence

`Data/LeaveRequests_Output.xlsx` contains 12 decision rows with the expected business-rule outcomes. `Data/LeaveRequests_Summary.xlsx` contains the expected aggregate counts. No `Logs/execution.log` was created because the successful run did not invoke the exception logger.

## Deterministic rule evidence

The UiPath output matches the earlier deterministic rule check: Approved=3, Rejected=7, Pending Manager Approval=1, Error=1. No email was sent.

## Clean-input runtime verification

- Input workbook: `Data/LeaveRequests_Clean_Input.xlsx` selected through the configured `InputPath` key in `Config.xlsx`.
- Configuration wiring fix: `Main.xaml` now reads the configured `InputPath` after configuration initialization and resolves the relative path against the project directory.
- Notifications: disabled by `Config.xlsx` (`NotificationEnabled=False`).
- Runtime verdict: passed (`Result=Success`, `hasErrors=False`, `errorMessage=null`, `debugState=null`).
- Workflow log summary: `Processed 5 requests; errors: 0; OK - Results written; OK - Summary written`.

### Clean-input output distribution

| Status | Count |
|---|---:|
| Approved | 5 |
| Rejected | 0 |
| Pending Manager Approval | 0 |
| Error | 0 |

### Clean-input summary workbook

`Data/LeaveRequests_Summary.xlsx` contained:

- Total Requests: 5
- Approved: 5
- Rejected: 0
- Pending Manager Approval: 0
- Errors: 0
- Notifications Sent: 0
- Notifications Failed: 0

The clean-input verification was executed in an isolated copy so the production default configuration and the 12-row failure-injection input remained unchanged.

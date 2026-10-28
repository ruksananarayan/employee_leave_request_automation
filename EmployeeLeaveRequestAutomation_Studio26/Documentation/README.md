# Employee Leave Request Automation

A modular UiPath Windows process that reads fictional leave requests from Excel, validates configurable rules, writes separate result and summary workbooks, and optionally sends employee notifications.

## Verified project

- Project folder: `EmployeeLeaveRequestAutomation_Studio26`
- Target framework: Windows; .NET: `net10.0`
- Expression language: Visual Basic
- Studio: 26.0.204.0-preview
- Entry point: `Main.xaml`
- Notifications: disabled by default

## Safety defaults

- `Config.xlsx` sets `NotificationEnabled=False`.
- The input workbook is never used as the output target.
- No password, token, or secret is stored in Excel, XAML, source control, or logs.
- Email credentials are retrieved locally from an approved UiPath Credential Asset only when notifications are enabled.
- Sample addresses use `example.com` and are fictional.
- Pending requests never receive an approval or rejection message.

## Project layout

- `Main.xaml`: orchestration.
- `Config.xlsx`: non-secret configuration.
- `Data/LeaveRequests_Input.xlsx`: failure-injection sample input used for business-rule testing.
- `Data/LeaveRequests_Clean_Input.xlsx`: separate clean sample input with five valid requests and no intentional business-rule errors.
- `Data/LeaveRequests_Output.xlsx`: separate result workbook.
- `Data/LeaveRequests_Summary.xlsx`: generated metrics workbook.
- `Data/LeaveRequests_Template.xlsx`: output schema template.
- `Workflows/`: modular child workflows.
- `Documentation/`: business rules, test plan, troubleshooting, dashboard guidance, and verification evidence.

## Verified execution

The compatible sibling project passed package restore, per-file validation for all nine XAML workflows, project-level build, and a headless runtime smoke test with notifications disabled.

Runtime evidence:

- 12 requests processed
- Approved: 3
- Rejected: 7
- Pending Manager Approval: 1
- Error: 1
- Notifications sent: 0
- Notifications failed: 0

See `Documentation/ValidationStatus.md`, `Documentation/PackageVerification.md`, and `Documentation/SmokeTestResults.md`.

## Run locally

```powershell
uip rpa run --file-path Main.xaml --project-dir "<PROJECT_ROOT>" --skip-build --output json
```

Review `Data/LeaveRequests_Output.xlsx` and `Data/LeaveRequests_Summary.xlsx` after execution. The exception log is created under `Logs/` only when an exception is logged; the verified successful run did not create one.

## Email setup

Create an approved Orchestrator Credential Asset named `EmployeeLeaveSmtpCredential`, or add an `EmailCredentialAsset=<approved asset name>` row to `Config.xlsx`. Configure the provider and sender settings, then enable notifications only after a controlled test recipient and provider authentication have been approved. Never paste a password into chat or into a workflow file. Microsoft 365 tenants should use the organization-approved Microsoft 365/Outlook or OAuth path rather than assuming basic SMTP authentication.

## Non-blocking build warnings

The verified build reports warnings for naming conventions, unused variables, missing organization Automation Hub URL policy metadata, and unavailable tenant-feed/dotnet packaging services. It reports no validation or build errors.

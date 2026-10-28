# EmployeeLeaveRequestAutomation_Studio26 agent context

<!-- PROJECT-CONTEXT:START -->
<!-- discovery-metadata: cs=0 xaml=9 deps=5 -->

- Project root: `C:\Users\Ruksana R\OneDrive\Documents\UiPath\RPA project\EmployeeLeaveRequestAutomation_Studio26`
- Windows project, .NET `net10.0`, Visual Basic expressions, Studio 26.0.204.0-preview; entry point `Main.xaml`.
- Nine XAML workflows total: `Main.xaml` plus eight files under `Workflows/`.
- Restored dependencies: `UiPath.System.Activities [26.10.0]`, `UiPath.Excel.Activities [3.6.1]`, `UiPath.Mail.Activities [2.12.10]`, `UiPath.Testing.Activities [25.10.3]`, `UiPath.UIAutomation.Activities [26.10.6]`.
- Input/config/output assets: `Config.xlsx`, `Data/LeaveRequests_Input.xlsx`, `Data/LeaveRequests_Output.xlsx`, and `Data/LeaveRequests_Summary.xlsx`.
- Notifications default to disabled. Credentials are retrieved locally from an approved UiPath Credential Asset only when notifications are enabled; never add passwords, tokens, or client secrets to source, workbooks, or logs.
- UiPath gate passed: restore, all nine per-file validations, project build, and headless runtime smoke test. Runtime distribution: Approved=3, Rejected=7, Pending Manager Approval=1, Error=1; no email was sent.
- Build warnings are non-blocking and documented in `Documentation/ValidationStatus.md` and `Documentation/PackageVerification.md`.
- Controlled live-email testing is not performed until approved provider and credential configuration exists.

<!-- PROJECT-CONTEXT:END -->

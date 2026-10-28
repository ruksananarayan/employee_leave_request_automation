# Package verification status

Target project: `EmployeeLeaveRequestAutomation_Studio26`

| Setting | Verified value |
|---|---|
| Studio | 26.0.204.0-preview |
| Framework | Windows |
| .NET | `net10.0` |
| Expressions | Visual Basic |
| CLI | UiPath CLI 1.204.0-preview.250 |

## Dependencies

| Package | Version | Intended use | Restore status |
|---|---:|---|---|
| UiPath.System.Activities | 26.10.0 | Core control flow, Invoke Code, Invoke Workflow File, logging, credential retrieval | Restored |
| UiPath.Excel.Activities | 3.6.1 | Excel workbook read/write activities | Restored |
| UiPath.Mail.Activities | 2.12.10 | Optional SMTP notification activity | Restored |
| UiPath.Testing.Activities | 25.10.3 | Scaffold dependency; not used by the process workflows | Restored |
| UiPath.UIAutomation.Activities | 26.10.6 | Scaffold dependency; not used by the process workflows | Restored |

## Gate results

- `uip rpa restore`: passed.
- `uip rpa validate`: passed for `Main.xaml` and all eight files under `Workflows/`.
- `uip rpa build`: passed with `Data.Success=True`.
- `uip rpa run --file-path Main.xaml --skip-build`: passed with `hasErrors=False` and `errorMessage=null`.

## Security repair applied

The initial build correctly rejected workflow-level `SecureString` arguments under `ST-SEC-007`. The project was repaired without downgrading to plaintext:

- `Main.xaml` has no password argument.
- `SendEmployeeNotification.xaml` retrieves the configured Orchestrator Credential Asset locally.
- Credential retrieval and SMTP use share the narrow inner retry-action scope, satisfying `ST-SEC-008`.
- No password or token is stored in source, workbooks, or logs.

## Non-blocking environment warnings

The build emitted warnings that do not fail the project: `dotnet` was unavailable for packaging/signing, tenant feeds could not be queried, the organization Automation Hub URL policy metadata is absent, and several naming/unused-variable analyzer recommendations remain.

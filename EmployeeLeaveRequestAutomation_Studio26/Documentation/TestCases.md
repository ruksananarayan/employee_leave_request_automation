# Test plan and execution record

The following cases are prepared from the sample data and failure-injection scenarios. The first twelve business-rule cases were exercised by the verified UiPath runtime smoke test with notifications disabled. Email and infrastructure-failure cases remain controlled follow-up tests.

| ID | Scenario | Expected status | Actual status | Result |
|---|---|---|---|---|
| TC-01 | Valid Annual request | Approved | Approved | Passed |
| TC-02 | Valid Sick request | Approved | Approved | Passed |
| TC-03 | Missing mandatory field | Rejected | Rejected | Passed |
| TC-04 | Invalid employee email | Rejected | Rejected | Passed |
| TC-05 | Duplicate RequestID | Rejected | Rejected | Passed |
| TC-06 | StartDate later than EndDate | Rejected | Rejected | Passed |
| TC-07 | Weekend-only request | Rejected | Rejected | Passed |
| TC-08 | Configured holiday in range | Calculated days exclude holiday | Approved; calculated days exclude holiday | Passed |
| TC-09 | Incorrect submitted TotalDays | Rejected | Rejected | Passed |
| TC-10 | Insufficient Annual balance | Rejected | Rejected | Passed |
| TC-11 | Threshold exceeded | Pending Manager Approval | Pending Manager Approval | Passed |
| TC-12 | Missing manager email | Error | Error | Passed |
| TC-13 | Email authentication failure | Decision unchanged | Not executed | Controlled test required |
| TC-14 | Excel locked during write | Decision data preserved; write error logged | Not executed | Controlled test required |
| TC-15 | Empty input workbook | Controlled Error | Not executed | Controlled test required |
| TC-16 | Single-row unexpected exception | Other rows continue | Not executed | Controlled test required |
| TC-17 | Clean five-row input with no intentional business-rule errors | Five Approved; zero Rejected, Pending Manager Approval, or Error | Five Approved; zero Rejected, Pending Manager Approval, or Error | Passed |

## Runtime evidence

The verified UiPath run processed 12 rows and produced the separate output and summary workbooks. Notifications remained disabled, so no live email provider or credential was contacted.

## Controlled notification test

Before enabling notifications, configure an approved Credential Asset, provider TLS/OAuth settings, and a controlled recipient. Confirm that `NotificationEnabled` is changed only for that test environment and review the output notification columns afterward.

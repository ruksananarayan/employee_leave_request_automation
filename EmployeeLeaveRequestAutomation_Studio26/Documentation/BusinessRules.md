# Business Rules

## Rule order

1. Required columns and row fields.
2. Request ID uniqueness.
3. Email format.
4. Leave type.
5. Date parsing, order, past-date policy, and working-day calculation.
6. Submitted TotalDays must match calculated working days.
7. Annual leave balance.
8. Approval threshold.
9. Manager information for approval-required requests.
10. Deterministic status and reason.

## Statuses

- `Approved`: all rules pass and calculated days are within the automatic approval threshold.
- `Rejected`: a business rule fails; the reason is explicit.
- `Pending Manager Approval`: all validation checks pass but calculated days exceed the configured threshold. This is not a final approval.
- `Error`: unexpected technical failure or an approval-required request has no manager email.

## Configurable keys

`ApprovalThreshold`, `AllowPastDates`, `HolidayDates`, `NotificationEnabled`, `MaxEmailRetries`, `InputPath`, `OutputPath`, and sheet names are stored in `Config.xlsx`.

## Working days

The calculation includes both start and end dates, excludes Saturday and Sunday, and excludes dates listed in `HolidayDates` as `yyyy-MM-dd` values.

These rules are portfolio examples only and must be reviewed against the organization's actual leave policy before production deployment.

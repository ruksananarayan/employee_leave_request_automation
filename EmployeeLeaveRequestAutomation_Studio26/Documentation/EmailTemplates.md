# Email templates

## Approved

Subject: `Leave Request Approved - [RequestID]`

Dear [EmployeeName],

Your leave request has been approved.

Request ID: [RequestID]
Leave Type: [LeaveType]
Start Date: [StartDate]
End Date: [EndDate]
Approved Working Days: [CalculatedDays]

Decision: Approved

Please coordinate with your team as necessary.

Regards,
HR Automation System

## Rejected

Subject: `Leave Request Update - [RequestID]`

Dear [EmployeeName],

Your leave request could not be approved based on the configured leave validation rules.

Request ID: [RequestID]
Leave Type: [LeaveType]
Start Date: [StartDate]
End Date: [EndDate]

Reason: [DecisionReason]

Please review the details and submit a corrected request if necessary.

Regards,
HR Automation System

## Pending

Subject: `Manager Approval Required - [RequestID]`

Pending requests intentionally do not receive an approval or rejection email. The implementation records `Skipped - Pending Manager Approval` in `NotificationStatus`.

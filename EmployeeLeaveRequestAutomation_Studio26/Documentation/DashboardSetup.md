# Dashboard setup

An optional dashboard source is provided as `Data/LeaveRequests_Summary.xlsx` after a successful run. It contains total requests, approved, rejected, pending, error, sent, skipped, and failed notification counts.

For UiPath Apps, bind a read-only table to the `LeaveResults` sheet and expose filters for RequestID, EmployeeID, department, leave type, status, and processed date. Connect **Start processing** to the published process and keep a confirmation step before invocation. Connect refresh/export controls to supported Apps actions or label them as prototype controls until a deployment-specific binding is configured.

Do not expose credentials or `Config.xlsx` secret fields to app users. Apply the organization's authentication, authorization, and data-minimization policies before publishing.

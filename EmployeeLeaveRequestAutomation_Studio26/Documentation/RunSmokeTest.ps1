$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$root = Split-Path -Parent $PSScriptRoot

function Read-XlsxRows([string]$path) {
    $stream = [IO.File]::OpenRead($path)
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read, $false)
    $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    if ($null -eq $entry) { throw "Worksheet part missing: $path" }
    $reader = [IO.StreamReader]::new($entry.Open())
    $xml = [xml]$reader.ReadToEnd()
    $reader.Dispose(); $zip.Dispose(); $stream.Dispose()
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('m','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $rows = @()
    foreach ($row in $xml.SelectNodes('//m:sheetData/m:row', $ns)) {
        $values = @()
        foreach ($cell in $row.SelectNodes('./m:c', $ns)) {
            $text = $cell.SelectSingleNode('./m:is/m:t', $ns)
            $values += $(if ($null -eq $text) { '' } else { $text.InnerText })
        }
        $rows += ,$values
    }
    return $rows
}

function To-Map([object[]]$rows) {
    $map = @{}
    foreach ($row in $rows | Select-Object -Skip 1) {
        if ($row.Count -ge 2 -and $row[0]) { $map[$row[0]] = $row[1] }
    }
    return $map
}

$configRows = Read-XlsxRows (Join-Path $root 'Config.xlsx')
$config = To-Map $configRows
$inputRows = Read-XlsxRows (Join-Path $root 'Data\LeaveRequests_Input.xlsx')
$headers = @($inputRows[0])
$holidaySet = [Collections.Generic.HashSet[datetime]]::new()
foreach ($h in $config['HolidayDates'].Split(';')) {
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($h.Trim(), 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { [void]$holidaySet.Add($parsed.Date) }
}
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$results = @()
$today = [datetime]'2026-10-27'
$threshold = [int]$config['ApprovalThreshold']
foreach ($values in $inputRows | Select-Object -Skip 1) {
    $row = @{}
    for ($i = 0; $i -lt $headers.Count; $i++) { $row[$headers[$i]] = if ($i -lt $values.Count) { $values[$i].Trim() } else { '' } }
    $status = 'Rejected'; $reason = ''; $days = 0
    $mandatory = @('RequestID','EmployeeID','EmployeeName','Email','LeaveType','StartDate','EndDate','TotalDays')
    foreach ($name in $mandatory) { if ([string]::IsNullOrWhiteSpace($row[$name])) { $reason = "Missing mandatory field: $name"; break } }
    if ($reason) { }
    elseif (-not $seen.Add($row['RequestID'])) { $reason = "Duplicate RequestID: $($row['RequestID'])" }
    elseif ($row['Email'] -notmatch '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') { $reason = 'Invalid employee email format' }
    elseif (@('Annual','Sick','Casual','Unpaid') -notcontains $row['LeaveType']) { $reason = "Invalid leave type: $($row['LeaveType'])" }
    else {
        $start = [datetime]::MinValue; $end = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($row['StartDate'], 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$start)) { $reason = 'Invalid or unparseable StartDate' }
        elseif (-not [datetime]::TryParseExact($row['EndDate'], 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$end)) { $reason = 'Invalid or unparseable EndDate' }
        elseif ($start.Date -gt $end.Date) { $reason = 'StartDate must not be later than EndDate' }
        elseif ($config['AllowPastDates'] -eq 'False' -and $start.Date -lt $today.Date) { $reason = 'Past-dated leave is not allowed by configuration' }
        else {
            for ($d = $start.Date; $d -le $end.Date; $d = $d.AddDays(1)) { if ($d.DayOfWeek -notin @([DayOfWeek]::Saturday,[DayOfWeek]::Sunday) -and -not $holidaySet.Contains($d)) { $days++ } }
            if ($days -le 0) { $reason = 'Requested range contains no working days' }
            elseif ([int]$row['TotalDays'] -ne $days) { $reason = 'Submitted TotalDays does not match calculated working days' }
            elseif ($row['LeaveType'] -eq 'Annual' -and $days -gt [int]$row['AnnualLeaveBalance']) { $reason = 'Requested working days exceed AnnualLeaveBalance' }
            elseif ($days -gt $threshold -and [string]::IsNullOrWhiteSpace($row['ManagerEmail'])) { $status = 'Error'; $reason = 'Manager email is missing for approval-required request' }
            elseif ($days -gt $threshold) { $status = 'Pending Manager Approval'; $reason = 'Request exceeds automatic approval threshold' }
            else { $status = 'Approved'; $reason = 'All configured validation rules passed' }
        }
    }
    if (-not $reason -and $status -eq 'Rejected') { $reason = 'Rejected by validation rules' }
    $notification = if ($status -eq 'Pending Manager Approval') { 'Skipped - Pending Manager Approval' } elseif ($status -eq 'Error') { 'Skipped - Error status' } elseif ($config['NotificationEnabled'] -eq 'False') { 'Not Configured' } elseif ($row['Email'] -notmatch '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') { 'Skipped - Invalid Email' } else { 'Ready' }
    $results += [pscustomobject]@{ RequestID=$row['RequestID']; Status=$status; CalculatedDays=$days; DecisionReason=$reason; NotificationStatus=$notification }
}

$counts = $results | Group-Object Status | Sort-Object Name
$expected = @{ Approved = 3; Rejected = 7; 'Pending Manager Approval' = 1; Error = 1 }
foreach ($key in $expected.Keys) { $actual = @($results | Where-Object Status -eq $key).Count; if ($actual -ne $expected[$key]) { throw "Expected $($expected[$key]) $key rows, got $actual" } }
$out = @('# No-email smoke test result', '', "Executed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')", 'Execution mode: standalone deterministic rule check; no UiPath Robot run and no email side effect.', '', '## Status distribution', '', '| Status | Count |', '|---|---:|')
foreach ($g in $counts) { $out += "| $($g.Name) | $($g.Count) |" }
$out += '', '## Row decisions', '', '| RequestID | Status | Days | NotificationStatus | Reason |', '|---|---|---:|---|---|'
foreach ($r in $results) { $out += "| $($r.RequestID) | $($r.Status) | $($r.CalculatedDays) | $($r.NotificationStatus) | $($r.DecisionReason) |" }
$out += '', 'Expected distribution matched: Approved=3, Rejected=7, Pending Manager Approval=1, Error=1.', 'Notifications were not sent because NotificationEnabled=False.'
[IO.File]::WriteAllLines((Join-Path $root 'Documentation\SmokeTestResults.md'), $out)
$results | Format-Table -AutoSize
exit 0

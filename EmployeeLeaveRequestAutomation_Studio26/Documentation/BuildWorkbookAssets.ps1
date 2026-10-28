$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Split-Path -Parent $PSScriptRoot
$data = Join-Path $root 'Data'

function XmlEscape([string]$value) {
    if ($null -eq $value) { return '' }
    return [System.Security.SecurityElement]::Escape($value)
}

function ColumnName([int]$number) {
    $result = ''
    while ($number -gt 0) {
        $remainder = ($number - 1) % 26
        $result = ([char](65 + $remainder)) + $result
        $number = [math]::Floor(($number - 1) / 26)
    }
    return $result
}

function New-Xlsx([string]$path, [string]$sheetName, [string[]]$headers, [object[][]]$rows) {
    $parts = @{}
    $parts['[Content_Types].xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@
    $parts['_rels/.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
    $parts['xl/workbook.xml'] = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="$(XmlEscape $sheetName)" sheetId="1" r:id="rId1"/></sheets>
</workbook>
"@
    $parts['xl/_rels/workbook.xml.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$builder.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    $allRows = New-Object System.Collections.Generic.List[object]
    [void]$allRows.Add($headers)
    foreach ($row in $rows) { [void]$allRows.Add($row) }
    for ($rowIndex = 0; $rowIndex -lt $allRows.Count; $rowIndex++) {
        $excelRow = $rowIndex + 1
        [void]$builder.Append("<row r=""$excelRow"">")
        $values = $allRows[$rowIndex]
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            $value = ''
            if ($columnIndex -lt $values.Count -and $null -ne $values[$columnIndex]) { $value = [string]$values[$columnIndex] }
            $reference = "$(ColumnName ($columnIndex + 1))$excelRow"
            [void]$builder.Append("<c r=""$reference"" t=""inlineStr""><is><t xml:space=""preserve"">$(XmlEscape $value)</t></is></c>")
        }
        [void]$builder.AppendLine('</row>')
    }
    [void]$builder.AppendLine('</sheetData></worksheet>')
    $parts['xl/worksheets/sheet1.xml'] = $builder.ToString()

    $temporaryPath = "$path.tmp"
    if (Test-Path $temporaryPath) { Remove-Item -LiteralPath $temporaryPath -Force }
    $fileStream = [System.IO.File]::Open($temporaryPath, [System.IO.FileMode]::Create)
    $archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    foreach ($part in $parts.GetEnumerator()) {
        $entry = $archive.CreateEntry($part.Key)
        $entryStream = $entry.Open()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($part.Value)
        $entryStream.Write($bytes, 0, $bytes.Length)
        $entryStream.Dispose()
    }
    $archive.Dispose()
    $fileStream.Dispose()
    Move-Item -LiteralPath $temporaryPath -Destination $path -Force
}

$configHeaders = @('Key', 'Value', 'Description')
$configRows = @(
    @('NotificationEnabled', 'False', 'Safe default. Set True only after credential and provider setup.'),
    @('EmailCredentialAsset', 'EmployeeLeaveSmtpCredential', 'Orchestrator Credential Asset name; secret is never stored here.'),
    @('EmailProvider', 'SMTP', 'Use SMTP only when approved by the organization.'),
    @('SenderDisplayName', 'HR Automation System', 'Display name for employee notifications.'),
    @('SenderEmail', 'hr-automation@example.com', 'Fictional sender placeholder.'),
    @('SmtpHost', 'smtp.office365.com', 'Provider host; verify tenant policy.'),
    @('SmtpPort', '587', 'TLS submission port.'),
    @('ApprovalThreshold', '5', 'Maximum working days for automatic approval.'),
    @('AllowPastDates', 'False', 'Reject requests whose start date is before today.'),
    @('HolidayDates', '2026-11-11;2026-12-25', 'Semicolon-separated yyyy-MM-dd dates.'),
    @('MaxEmailRetries', '2', 'Bounded retries for transient delivery failures.'),
    @('InputSheet', 'LeaveRequests', 'Input worksheet name.'),
    @('OutputSheet', 'LeaveResults', 'Output worksheet name.'),
    @('InputPath', 'Data\LeaveRequests_Input.xlsx', 'Relative to project root.'),
    @('OutputPath', 'Data\LeaveRequests_Output.xlsx', 'Separate generated output file.'),
    @('LogFolder', 'Logs', 'Relative log folder.')
)
New-Xlsx (Join-Path $root 'Config.xlsx') 'Config' $configHeaders $configRows

$inputHeaders = @('RequestID', 'EmployeeID', 'EmployeeName', 'Email', 'Department', 'LeaveType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'AnnualLeaveBalance', 'ManagerEmail')
$inputRows = @(
    @('LR-1001','EMP-1001','Alice Morgan','alice.morgan@example.com','Finance','Annual','2026-11-02','2026-11-04','3','Family travel','10','manager.finance@example.com'),
    @('LR-1002','EMP-1002','Brian Lee','brian.lee@example.com','Operations','Sick','2026-11-05','2026-11-05','1','Medical appointment','0','manager.ops@example.com'),
    @('LR-1003','EMP-1003','Carla Diaz','','Engineering','Annual','2026-11-09','2026-11-10','2','Personal time','5','manager.eng@example.com'),
    @('LR-1004','EMP-1004','David Shah','david.shah.example.com','Sales','Casual','2026-11-06','2026-11-06','1','Personal errand','0','manager.sales@example.com'),
    @('LR-1001','EMP-1005','Elena Rossi','elena.rossi@example.com','Finance','Annual','2026-11-16','2026-11-17','2','Family event','8','manager.finance@example.com'),
    @('LR-1006','EMP-1006','Farah Khan','farah.khan@example.com','HR','Annual','2026-11-20','2026-11-18','3','Incorrect date order','10','manager.hr@example.com'),
    @('LR-1007','EMP-1007','George Chen','george.chen@example.com','Operations','Casual','2026-11-07','2026-11-08','0','Weekend event','0','manager.ops@example.com'),
    @('LR-1008','EMP-1008','Hannah Smith','hannah.smith@example.com','Engineering','Annual','2026-11-10','2026-11-12','2','Personal travel spanning holiday','5','manager.eng@example.com'),
    @('LR-1009','EMP-1009','Ivan Petrov','ivan.petrov@example.com','Legal','Annual','2026-11-23','2026-11-25','4','Incorrect day count','10','manager.legal@example.com'),
    @('LR-1010','EMP-1010','Julia Park','julia.park@example.com','Finance','Annual','2026-11-26','2026-12-02','5','Extended vacation','2','manager.finance@example.com'),
    @('LR-1011','EMP-1011','Kiran Rao','kiran.rao@example.com','Engineering','Annual','2026-12-07','2026-12-14','6','Long planned leave','10','manager.eng@example.com'),
    @('LR-1012','EMP-1012','Lina Brown','lina.brown@example.com','Sales','Annual','2026-12-07','2026-12-14','6','Long leave without manager','10','')
)
New-Xlsx (Join-Path $data 'LeaveRequests_Input.xlsx') 'LeaveRequests' $inputHeaders $inputRows
$outputHeaders = $inputHeaders + @('CalculatedDays', 'Status', 'DecisionReason', 'ProcessedDate', 'NotificationStatus', 'NotificationError')
New-Xlsx (Join-Path $data 'LeaveRequests_Template.xlsx') 'LeaveResults' $outputHeaders @()
New-Xlsx (Join-Path $data 'LeaveRequests_Output.xlsx') 'LeaveResults' $outputHeaders @()

Get-ChildItem -LiteralPath $root, $data -Filter '*.xlsx' | Select-Object FullName, Length

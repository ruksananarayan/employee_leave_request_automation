$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root 'Data\LeaveRequests_Input.xlsx'
$output = Join-Path $root 'Data\LeaveRequests_Clean_Input.xlsx'

function Xml-Escape([string]$Value) {
    if ($null -eq $Value) { return '' }
    return [System.Security.SecurityElement]::Escape($Value)
}

function New-Cell([string]$Reference, [string]$Value) {
    return '<c r="' + $Reference + '" t="inlineStr"><is><t xml:space="preserve">' + (Xml-Escape $Value) + '</t></is></c>'
}

function New-Row([int]$Number, [string[]]$Values) {
    $columns = @('A','B','C','D','E','F','G','H','I','J','K','L')
    $cells = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $cells.Add((New-Cell ($columns[$i] + $Number) $Values[$i]))
    }
    return '<row r="' + $Number + '">' + ($cells -join '') + '</row>'
}

function Replace-Sheet([string]$Path, [string]$SheetXml) {
    $temporary = $Path + '.tmp'
    $inputStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    $inputZip = [System.IO.Compression.ZipArchive]::new($inputStream, [System.IO.Compression.ZipArchiveMode]::Read)
    $outputStream = [System.IO.File]::Open($temporary, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $outputZip = [System.IO.Compression.ZipArchive]::new($outputStream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $inputZip.Entries) {
            $newEntry = $outputZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $destination = $newEntry.Open()
            try {
                if ($entry.FullName -eq 'xl/worksheets/sheet1.xml') {
                    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($SheetXml)
                    $destination.Write($bytes, 0, $bytes.Length)
                } else {
                    $sourceStream = $entry.Open()
                    try { $sourceStream.CopyTo($destination) } finally { $sourceStream.Dispose() }
                }
            } finally { $destination.Dispose() }
        }
    } finally {
        $outputZip.Dispose()
        $outputStream.Dispose()
        $inputZip.Dispose()
        $inputStream.Dispose()
    }
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

$headers = @('RequestID','EmployeeID','EmployeeName','Email','Department','LeaveType','StartDate','EndDate','TotalDays','Reason','AnnualLeaveBalance','ManagerEmail')
$rows = @(
    @('CLEAN-1001','EMP-C1001','Asha Morgan','asha.morgan@example.com','Finance','Annual','2026-11-02','2026-11-04','3','Family travel','10','manager.finance@example.com'),
    @('CLEAN-1002','EMP-C1002','Ben Lee','ben.lee@example.com','Operations','Sick','2026-11-05','2026-11-05','1','Medical appointment','0','manager.ops@example.com'),
    @('CLEAN-1003','EMP-C1003','Chitra Diaz','chitra.diaz@example.com','Engineering','Casual','2026-11-09','2026-11-10','2','Personal time','5','manager.eng@example.com'),
    @('CLEAN-1004','EMP-C1004','Dev Shah','dev.shah@example.com','Sales','Annual','2026-11-10','2026-11-12','2','Travel spanning configured holiday','10','manager.sales@example.com'),
    @('CLEAN-1005','EMP-C1005','Esha Rossi','esha.rossi@example.com','HR','Unpaid','2026-11-16','2026-11-17','2','Personal leave','0','manager.hr@example.com')
)

$sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
$sheetXml += New-Row 1 $headers
for ($i = 0; $i -lt $rows.Count; $i++) { $sheetXml += New-Row ($i + 2) $rows[$i] }
$sheetXml += '</sheetData></worksheet>'

Copy-Item -LiteralPath $source -Destination $output -Force
Replace-Sheet $output $sheetXml

[pscustomobject]@{
    Output = $output
    Rows = $rows.Count
    NotificationEnabled = 'False in Config.xlsx by default'
    ExpectedStatus = 'Approved=5; Rejected=0; Pending Manager Approval=0; Error=0'
} | ConvertTo-Json
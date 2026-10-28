$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$wf = Join-Path $root 'Workflows'
New-Item -ItemType Directory -Path $wf -Force | Out-Null

function Esc([string]$value) {
    if ($null -eq $value) { return '' }
    return $value.Replace('&','&amp;').Replace('"','&quot;').Replace('<','&lt;').Replace('>','&gt;').Replace("`r",'').Replace("`n",'&#xA;')
}
function X([string]$expression) { return (Esc $expression) }
function Code([string]$source) { return (Esc $source) }

$imports = @'
  <TextExpression.NamespacesForImplementation>
    <sco:Collection x:TypeArguments="x:String">
      <x:String>System</x:String><x:String>System.Activities</x:String><x:String>System.Activities.Statements</x:String>
      <x:String>System.Collections</x:String><x:String>System.Collections.Generic</x:String><x:String>System.Collections.ObjectModel</x:String>
      <x:String>System.Data</x:String><x:String>System.Diagnostics</x:String><x:String>System.Globalization</x:String>
      <x:String>System.IO</x:String><x:String>System.Linq</x:String><x:String>System.Net.Mail</x:String>
      <x:String>System.Security</x:String><x:String>System.Text</x:String><x:String>System.Text.RegularExpressions</x:String>
      <x:String>UiPath.Core</x:String><x:String>UiPath.Core.Activities</x:String><x:String>GlobalVariablesNamespace</x:String><x:String>GlobalConstantsNamespace</x:String>
    </sco:Collection>
  </TextExpression.NamespacesForImplementation>
  <TextExpression.ReferencesForImplementation>
    <sco:Collection x:TypeArguments="AssemblyReference">
      <AssemblyReference>System.Activities</AssemblyReference><AssemblyReference>Microsoft.VisualBasic</AssemblyReference>
      <AssemblyReference>System.Private.CoreLib</AssemblyReference><AssemblyReference>System</AssemblyReference>
      <AssemblyReference>System.Data</AssemblyReference><AssemblyReference>System.Data.Common</AssemblyReference>
      <AssemblyReference>System.Data.DataSetExtensions</AssemblyReference><AssemblyReference>System.Linq</AssemblyReference>
      <AssemblyReference>System.IO</AssemblyReference><AssemblyReference>System.Net.Mail</AssemblyReference>
      <AssemblyReference>UiPath.System.Activities</AssemblyReference><AssemblyReference>UiPath.Excel.Activities</AssemblyReference>
      <AssemblyReference>UiPath.Mail.Activities</AssemblyReference>
    </sco:Collection>
  </TextExpression.ReferencesForImplementation>
'@

function Header([string]$className, [string]$members) {
    $header = @"
<Activity mc:Ignorable="sap sap2010" x:Class="$className" sap2010:WorkflowViewState.IdRef="ActivityBuilder_1"
 xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation"
 xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation"
 xmlns:sco="clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib"
 xmlns:scg="clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib"
 xmlns:sd="clr-namespace:System.Data;assembly=System.Data.Common"
 xmlns:s="clr-namespace:System;assembly=System.Private.CoreLib"
 xmlns:ss="clr-namespace:System.Security;assembly=System.Private.CoreLib"
 xmlns:ui="clr-namespace:UiPath.Core.Activities;assembly=UiPath.System.Activities"
 xmlns:uea="clr-namespace:UiPath.Excel.Activities;assembly=UiPath.Excel.Activities"
 xmlns:smtp="clr-namespace:UiPath.Mail.SMTP.Activities;assembly=UiPath.Mail.Activities"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
"@
    return $header + $imports + $members
}
function Footer { return '</Activity>' }
function Members([string]$text) { return "  <x:Members>`n$text`n  </x:Members>" }
function Prop([string]$name,[string]$type) { return "    <x:Property Name=""$name"" Type=""$type"" />" }
function WriteWorkflow([string]$file,[string]$className,[string]$members,[string]$body) {
    $content = (Header $className (Members $members)) + "`n  <Sequence DisplayName=""$className Sequence"" sap2010:WorkflowViewState.IdRef=""Sequence_1"">`n" + $body + "`n  </Sequence>`n" + (Footer)
    # Normalize the quote-doubled template text without damaging empty attributes or escaped VB strings.
    $content = [regex]::Replace($content, '""([^"\r\n]+)""', '"$1"')
    $content = $content.Replace('="""', '=""')
    $content = $content.Replace('"" />', '" />')
    $content = $content.Replace('""', '"')
    # Preserve the empty Range attribute used to read an entire worksheet.
    $content = $content.Replace('Range=" AddHeaders', 'Range="" AddHeaders')
    $destinationFolder = $wf
    if ($file -eq 'Main.xaml') { $destinationFolder = $root }
    [System.IO.File]::WriteAllText((Join-Path $destinationFolder $file), $content, [System.Text.UTF8Encoding]::new($false))
}

$initCode = @'
Dim config As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
If dtConfig Is Nothing OrElse dtConfig.Columns.Count < 2 Then
    out_ConfigText = ""
    out_Status = "Error - Config workbook is empty or missing Key/Value columns"
Else
    For Each r As DataRow In dtConfig.Rows
        If Not r.IsNull(0) Then
            Dim key = r(0).ToString().Trim()
            Dim value = If(r.IsNull(1), "", r(1).ToString().Trim())
            If key <> "" Then config(key) = value
        End If
    Next
    Dim parts As New List(Of String)
    For Each pair In config
        parts.Add(pair.Key & "=" & pair.Value.Replace("|", " ").Replace(vbCr, " ").Replace(vbLf, " "))
    Next
    out_ConfigText = String.Join("|", parts)
    out_Status = "OK - " & config.Count.ToString() & " configuration keys loaded"
End If
'@
$initMembers = (Prop 'in_ConfigPath' 'InArgument(x:String)') + "`n" + (Prop 'in_SheetName' 'InArgument(x:String)') + "`n" + (Prop 'out_ConfigText' 'OutArgument(x:String)') + "`n" + (Prop 'out_Status' 'OutArgument(x:String)')
$initBody = @"
  <Sequence.Variables><Variable x:TypeArguments=""sd:DataTable"" Name=""dtConfig"" /></Sequence.Variables>
  <uea:ReadRange DisplayName=""Read Config Workbook"" sap2010:WorkflowViewState.IdRef=""ReadRange_1"" WorkbookPath=""[$(X 'in_ConfigPath')]"" SheetName=""[$(X 'in_SheetName')]"" Range="""" AddHeaders=""True"" DataTable=""[dtConfig]"" />
  <ui:InvokeCode DisplayName=""Normalize Configuration"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $initCode)"">
    <ui:InvokeCode.Arguments>
      <scg:Dictionary x:TypeArguments=""x:String, Argument"">
        <InArgument x:TypeArguments=""sd:DataTable"" x:Key=""dtConfig"">[dtConfig]</InArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_ConfigText"">[out_ConfigText]</OutArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_Status"">[out_Status]</OutArgument>
      </scg:Dictionary>
    </ui:InvokeCode.Arguments>
  </ui:InvokeCode>
"@
WriteWorkflow 'InitializeConfig.xaml' 'Workflows_InitializeConfig' $initMembers $initBody

$readCode = @'
Dim required = New String() {"RequestID", "EmployeeID", "EmployeeName", "Email", "Department", "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", "AnnualLeaveBalance", "ManagerEmail"}
out_SchemaValid = True
out_SchemaReason = "Schema OK"
If dtInput Is Nothing OrElse dtInput.Columns.Count = 0 Then
    out_SchemaValid = False
    out_SchemaReason = "Input workbook is empty"
Else
    For Each name In required
        If Not dtInput.Columns.Contains(name) Then
            out_SchemaValid = False
            out_SchemaReason = "Missing required column: " & name
            Exit For
        End If
    Next
End If
'@
$readMembers = (Prop 'in_InputPath' 'InArgument(x:String)') + "`n" + (Prop 'in_SheetName' 'InArgument(x:String)') + "`n" + (Prop 'out_dtInput' 'OutArgument(sd:DataTable)') + "`n" + (Prop 'out_SchemaValid' 'OutArgument(x:Boolean)') + "`n" + (Prop 'out_SchemaReason' 'OutArgument(x:String)')
$readBody = @"
  <Sequence.Variables><Variable x:TypeArguments=""sd:DataTable"" Name=""dtInput"" /></Sequence.Variables>
  <uea:ReadRange DisplayName=""Read Leave Requests"" sap2010:WorkflowViewState.IdRef=""ReadRange_1"" WorkbookPath=""[$(X 'in_InputPath')]"" SheetName=""[$(X 'in_SheetName')]"" Range="""" AddHeaders=""True"" DataTable=""[dtInput]"" />
  <Assign DisplayName=""Expose Input DataTable"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""sd:DataTable""><VisualBasicReference x:TypeArguments=""sd:DataTable"" ExpressionText=""out_dtInput"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""sd:DataTable""><VisualBasicValue x:TypeArguments=""sd:DataTable"" ExpressionText=""dtInput"" /></InArgument></Assign.Value></Assign>
  <ui:InvokeCode DisplayName=""Validate Input Schema"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $readCode)"">
    <ui:InvokeCode.Arguments>
      <scg:Dictionary x:TypeArguments=""x:String, Argument"">
        <InArgument x:TypeArguments=""sd:DataTable"" x:Key=""dtInput"">[dtInput]</InArgument>
        <OutArgument x:TypeArguments=""x:Boolean"" x:Key=""out_SchemaValid"">[out_SchemaValid]</OutArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_SchemaReason"">[out_SchemaReason]</OutArgument>
      </scg:Dictionary>
    </ui:InvokeCode.Arguments>
  </ui:InvokeCode>
"@
WriteWorkflow 'ReadLeaveRequests.xaml' 'Workflows_ReadLeaveRequests' $readMembers $readBody

$validateCode = @'
Dim cfg As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
If in_configText IsNot Nothing Then
    For Each part In in_configText.Split("|"c)
        Dim p = part.IndexOf("="c)
        If p > 0 Then cfg(part.Substring(0, p)) = part.Substring(p + 1)
    Next
End If
Dim Cell = Function(column As String) As String
    If in_row Is Nothing OrElse Not in_row.Table.Columns.Contains(column) OrElse in_row.IsNull(column) Then Return ""
    Return in_row(column).ToString().Trim()
End Function
out_status = "Rejected"
out_reason = ""
out_calculatedDays = 0
out_managerRequired = False
Dim requestId = Cell("RequestID")
Dim mandatory = New String() {"RequestID", "EmployeeID", "EmployeeName", "Email", "LeaveType", "StartDate", "EndDate", "TotalDays"}
For Each name In mandatory
    If Cell(name) = "" Then
        out_reason = "Missing mandatory field: " & name
        out_seenRequestIDs = If(in_seenRequestIDs, "")
        Return
    End If
Next
Dim seen = If(in_seenRequestIDs, "")
If seen.IndexOf("|" & requestId & "|", StringComparison.OrdinalIgnoreCase) >= 0 Then
    out_reason = "Duplicate RequestID: " & requestId
    out_seenRequestIDs = seen
    Return
End If
out_seenRequestIDs = seen & requestId & "|"
If Not Regex.IsMatch(Cell("Email"), "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$") Then
    out_reason = "Invalid employee email format"
    Return
End If
Dim leaveType = Cell("LeaveType")
If Not New String() {"Annual", "Sick", "Casual", "Unpaid"}.Contains(leaveType, StringComparer.OrdinalIgnoreCase) Then
    out_reason = "Invalid leave type: " & leaveType
    Return
End If
Dim startDate As DateTime
Dim endDate As DateTime
If Not DateTime.TryParseExact(Cell("StartDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, startDate) Then
    out_reason = "Invalid or unparseable StartDate"
    Return
End If
If Not DateTime.TryParseExact(Cell("EndDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, endDate) Then
    out_reason = "Invalid or unparseable EndDate"
    Return
End If
If startDate.Date > endDate.Date Then
    out_reason = "StartDate must not be later than EndDate"
    Return
End If
Dim allowPast = False
Boolean.TryParse(If(cfg.ContainsKey("AllowPastDates"), cfg("AllowPastDates"), "False"), allowPast)
If Not allowPast AndAlso startDate.Date < DateTime.Today Then
    out_reason = "Past-dated leave is not allowed by configuration"
    Return
End If
Dim holidays As New HashSet(Of DateTime)
Dim holidayText = If(cfg.ContainsKey("HolidayDates"), cfg("HolidayDates"), "")
For Each h In holidayText.Split(";"c)
    Dim holiday As DateTime
    If DateTime.TryParseExact(h.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, holiday) Then holidays.Add(holiday.Date)
Next
Dim cursor = startDate.Date
While cursor <= endDate.Date
    If cursor.DayOfWeek <> DayOfWeek.Saturday AndAlso cursor.DayOfWeek <> DayOfWeek.Sunday AndAlso Not holidays.Contains(cursor.Date) Then out_calculatedDays += 1
    cursor = cursor.AddDays(1)
End While
If out_calculatedDays <= 0 Then
    out_reason = "Requested range contains no working days"
    Return
End If
Dim submittedDays As Integer
If Not Integer.TryParse(Cell("TotalDays"), submittedDays) Then
    out_reason = "TotalDays is not a valid integer"
    Return
End If
If submittedDays <> out_calculatedDays Then
    out_reason = "Submitted TotalDays does not match calculated working days"
    Return
End If
If leaveType.Equals("Annual", StringComparison.OrdinalIgnoreCase) Then
    Dim balance As Integer
    If Not Integer.TryParse(Cell("AnnualLeaveBalance"), balance) Then
        out_reason = "AnnualLeaveBalance is not a valid integer"
        Return
    End If
    If out_calculatedDays > balance Then
        out_reason = "Requested working days exceed AnnualLeaveBalance"
        Return
    End If
End If
Dim threshold As Integer = 0
Integer.TryParse(If(cfg.ContainsKey("ApprovalThreshold"), cfg("ApprovalThreshold"), "0"), threshold)
If out_calculatedDays > threshold Then
    out_managerRequired = True
    If Cell("ManagerEmail") = "" Then
        out_status = "Error"
        out_reason = "Manager email is missing for approval-required request"
    Else
        out_status = "Pending Manager Approval"
        out_reason = "Request exceeds automatic approval threshold"
    End If
Else
    out_status = "Approved"
    out_reason = "All configured validation rules passed"
End If
'@
$validateMembers = (Prop 'in_row' 'InArgument(sd:DataRow)') + "`n" + (Prop 'in_seenRequestIDs' 'InArgument(x:String)') + "`n" + (Prop 'in_configText' 'InArgument(x:String)') + "`n" + (Prop 'out_seenRequestIDs' 'OutArgument(x:String)') + "`n" + (Prop 'out_status' 'OutArgument(x:String)') + "`n" + (Prop 'out_reason' 'OutArgument(x:String)') + "`n" + (Prop 'out_calculatedDays' 'OutArgument(x:Int32)') + "`n" + (Prop 'out_managerRequired' 'OutArgument(x:Boolean)')
$validateBody = @"
  <ui:InvokeCode DisplayName=""Validate Leave Request"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $validateCode)"">
    <ui:InvokeCode.Arguments>
      <scg:Dictionary x:TypeArguments=""x:String, Argument"">
        <InArgument x:TypeArguments=""sd:DataRow"" x:Key=""in_row"">[in_row]</InArgument>
        <InArgument x:TypeArguments=""x:String"" x:Key=""in_seenRequestIDs"">[in_seenRequestIDs]</InArgument>
        <InArgument x:TypeArguments=""x:String"" x:Key=""in_configText"">[in_configText]</InArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_seenRequestIDs"">[out_seenRequestIDs]</OutArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_status"">[out_status]</OutArgument>
        <OutArgument x:TypeArguments=""x:String"" x:Key=""out_reason"">[out_reason]</OutArgument>
        <OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_calculatedDays"">[out_calculatedDays]</OutArgument>
        <OutArgument x:TypeArguments=""x:Boolean"" x:Key=""out_managerRequired"">[out_managerRequired]</OutArgument>
      </scg:Dictionary>
    </ui:InvokeCode.Arguments>
  </ui:InvokeCode>
"@
WriteWorkflow 'ValidateLeaveRequest.xaml' 'Workflows_ValidateLeaveRequest' $validateMembers $validateBody

$processInitCode = @'
out_dtResults = New DataTable()
Dim names = New String() {"RequestID", "EmployeeID", "EmployeeName", "Email", "Department", "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", "AnnualLeaveBalance", "ManagerEmail", "CalculatedDays", "Status", "DecisionReason", "ProcessedDate", "NotificationStatus", "NotificationError"}
For Each name In names
    out_dtResults.Columns.Add(name, GetType(String))
Next
'@
$processCountCode = @'
out_processedCount = If(dtResults Is Nothing, 0, dtResults.Rows.Count)
out_errorCount = If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("Status").ToString() = "Error"))
'@
$processMembers = (Prop 'in_dtInput' 'InArgument(sd:DataTable)') + "`n" + (Prop 'in_configText' 'InArgument(x:String)') + "`n" + (Prop 'out_dtResults' 'OutArgument(sd:DataTable)') + "`n" + (Prop 'out_processedCount' 'OutArgument(x:Int32)') + "`n" + (Prop 'out_errorCount' 'OutArgument(x:Int32)')
$processBody = @"
  <Sequence.Variables>
    <Variable x:TypeArguments=""sd:DataTable"" Name=""dtResults"" />
    <Variable x:TypeArguments=""x:String"" Name=""seenRequestIDs"" Default=""|"" />
    <Variable x:TypeArguments=""x:String"" Name=""nextSeenRequestIDs"" />
    <Variable x:TypeArguments=""x:String"" Name=""status"" />
    <Variable x:TypeArguments=""x:String"" Name=""reason"" />
    <Variable x:TypeArguments=""x:Int32"" Name=""calculatedDays"" />
    <Variable x:TypeArguments=""x:Boolean"" Name=""managerRequired"" />
  </Sequence.Variables>
  <ui:InvokeCode DisplayName=""Initialize Result Schema"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $processInitCode)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><OutArgument x:TypeArguments=""sd:DataTable"" x:Key=""out_dtResults"">[dtResults]</OutArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
  <ui:ForEachRow DisplayName=""Process Each Leave Request"" DataTable=""[in_dtInput]"" sap2010:WorkflowViewState.IdRef=""ForEachRow_1""><ui:ForEachRow.Body><ActivityAction x:TypeArguments=""sd:DataRow""><ActivityAction.Argument><DelegateInArgument x:TypeArguments=""sd:DataRow"" Name=""CurrentRow"" /></ActivityAction.Argument><Sequence DisplayName=""Row Transaction"">
    <ui:InvokeWorkflowFile DisplayName=""Validate Current Request"" WorkflowFileName=""Workflows\ValidateLeaveRequest.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_1""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataRow"" x:Key=""in_row"">[CurrentRow]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_seenRequestIDs"">[seenRequestIDs]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_configText"">[in_configText]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_seenRequestIDs"">[nextSeenRequestIDs]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_status"">[status]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_reason"">[reason]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_calculatedDays"">[calculatedDays]</OutArgument><OutArgument x:TypeArguments=""x:Boolean"" x:Key=""out_managerRequired"">[managerRequired]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <Assign DisplayName=""Update Seen Request IDs"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""seenRequestIDs"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""nextSeenRequestIDs"" /></InArgument></Assign.Value></Assign>
    <ui:AddDataRow DisplayName=""Append Decision Result"" sap2010:WorkflowViewState.IdRef=""AddDataRow_1""><ui:AddDataRow.DataTable><InOutArgument x:TypeArguments=""sd:DataTable""><VisualBasicReference x:TypeArguments=""sd:DataTable"" ExpressionText=""dtResults"" /></InOutArgument></ui:AddDataRow.DataTable><ui:AddDataRow.ArrayRow><InArgument x:TypeArguments=""s:Object[]""><VisualBasicValue x:TypeArguments=""s:Object[]"" ExpressionText=""New Object() {If(CurrentRow.IsNull(&quot;RequestID&quot;), &quot;&quot;, CurrentRow(&quot;RequestID&quot;).ToString()), If(CurrentRow.IsNull(&quot;EmployeeID&quot;), &quot;&quot;, CurrentRow(&quot;EmployeeID&quot;).ToString()), If(CurrentRow.IsNull(&quot;EmployeeName&quot;), &quot;&quot;, CurrentRow(&quot;EmployeeName&quot;).ToString()), If(CurrentRow.IsNull(&quot;Email&quot;), &quot;&quot;, CurrentRow(&quot;Email&quot;).ToString()), If(CurrentRow.IsNull(&quot;Department&quot;), &quot;&quot;, CurrentRow(&quot;Department&quot;).ToString()), If(CurrentRow.IsNull(&quot;LeaveType&quot;), &quot;&quot;, CurrentRow(&quot;LeaveType&quot;).ToString()), If(CurrentRow.IsNull(&quot;StartDate&quot;), &quot;&quot;, CurrentRow(&quot;StartDate&quot;).ToString()), If(CurrentRow.IsNull(&quot;EndDate&quot;), &quot;&quot;, CurrentRow(&quot;EndDate&quot;).ToString()), If(CurrentRow.IsNull(&quot;TotalDays&quot;), &quot;&quot;, CurrentRow(&quot;TotalDays&quot;).ToString()), If(CurrentRow.IsNull(&quot;Reason&quot;), &quot;&quot;, CurrentRow(&quot;Reason&quot;).ToString()), If(CurrentRow.IsNull(&quot;AnnualLeaveBalance&quot;), &quot;&quot;, CurrentRow(&quot;AnnualLeaveBalance&quot;).ToString()), If(CurrentRow.IsNull(&quot;ManagerEmail&quot;), &quot;&quot;, CurrentRow(&quot;ManagerEmail&quot;).ToString()), calculatedDays.ToString(), status, reason, DateTime.Now.ToString(&quot;yyyy-MM-dd HH:mm:ss&quot;), &quot;Pending Notification&quot;, &quot;&quot;}"" /></InArgument></ui:AddDataRow.ArrayRow></ui:AddDataRow>
  </Sequence></ActivityAction></ui:ForEachRow.Body></ui:ForEachRow>
  <ui:InvokeCode DisplayName=""Count Decisions"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_2"" Code=""$(Code $processCountCode)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataTable"" x:Key=""dtResults"">[dtResults]</InArgument><OutArgument x:TypeArguments=""sd:DataTable"" x:Key=""out_dtResults"">[out_dtResults]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_processedCount"">[out_processedCount]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_errorCount"">[out_errorCount]</OutArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
"@
WriteWorkflow 'ProcessLeaveRequests.xaml' 'Workflows_ProcessLeaveRequests' $processMembers $processBody

$notifyCode = @'
Dim cfg As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
If in_configText IsNot Nothing Then
    For Each part In in_configText.Split("|"c)
        Dim p = part.IndexOf("="c)
        If p > 0 Then cfg(part.Substring(0, p)) = part.Substring(p + 1)
    Next
End If
Dim value = Function(column As String) As String
    If in_row Is Nothing OrElse Not in_row.Table.Columns.Contains(column) OrElse in_row.IsNull(column) Then Return ""
    Return in_row(column).ToString().Trim()
End Function
Dim enabled = False
Boolean.TryParse(If(cfg.ContainsKey("NotificationEnabled"), cfg("NotificationEnabled"), "False"), enabled)
Dim currentStatus = value("Status")
Dim recipient = value("Email")
Dim validEmail = Regex.IsMatch(recipient, "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
out_notificationStatus = ""
out_notificationError = ""
If currentStatus = "Pending Manager Approval" Then
    out_notificationStatus = "Skipped - Pending Manager Approval"
ElseIf currentStatus = "Error" Then
    out_notificationStatus = "Skipped - Error status"
ElseIf Not enabled Then
    out_notificationStatus = "Not Configured"
ElseIf Not validEmail Then
    out_notificationStatus = "Skipped - Invalid Email"
Else
    out_notificationStatus = "Ready"
End If
out_subject = ""
out_body = ""
If out_notificationStatus = "Ready" Then
    If currentStatus = "Approved" Then
        out_subject = "Leave Request Approved - " & value("RequestID")
        out_body = "Dear " & value("EmployeeName") & "," & vbCrLf & vbCrLf & "Your leave request has been approved." & vbCrLf & vbCrLf & "Request ID: " & value("RequestID") & vbCrLf & "Leave Type: " & value("LeaveType") & vbCrLf & "Start Date: " & value("StartDate") & vbCrLf & "End Date: " & value("EndDate") & vbCrLf & "Approved Working Days: " & value("CalculatedDays") & vbCrLf & vbCrLf & "Decision: Approved" & vbCrLf & vbCrLf & "Please coordinate with your team as necessary." & vbCrLf & vbCrLf & "Regards," & vbCrLf & "HR Automation System"
    Else
        out_subject = "Leave Request Update - " & value("RequestID")
        out_body = "Dear " & value("EmployeeName") & "," & vbCrLf & vbCrLf & "Your leave request could not be approved based on the configured leave validation rules." & vbCrLf & vbCrLf & "Request ID: " & value("RequestID") & vbCrLf & "Leave Type: " & value("LeaveType") & vbCrLf & "Start Date: " & value("StartDate") & vbCrLf & "End Date: " & value("EndDate") & vbCrLf & vbCrLf & "Reason: " & value("DecisionReason") & vbCrLf & vbCrLf & "Please review the details and submit a corrected request if necessary." & vbCrLf & vbCrLf & "Regards," & vbCrLf & "HR Automation System"
    End If
End If
out_server = If(cfg.ContainsKey("SmtpHost"), cfg("SmtpHost"), "")
Integer.TryParse(If(cfg.ContainsKey("SmtpPort"), cfg("SmtpPort"), "587"), out_port)
out_sender = If(cfg.ContainsKey("SenderEmail"), cfg("SenderEmail"), "")
'@
$notifyFinal = @'
If in_row IsNot Nothing Then
    If in_row.Table.Columns.Contains("NotificationStatus") Then in_row("NotificationStatus") = out_notificationStatus
    If in_row.Table.Columns.Contains("NotificationError") Then in_row("NotificationError") = out_notificationError
End If
'@
$notifyMembers = (Prop 'in_row' 'InArgument(sd:DataRow)') + "`n" + (Prop 'in_configText' 'InArgument(x:String)') + "`n" + (Prop 'in_securePassword' 'InArgument(ss:SecureString)') + "`n" + (Prop 'out_notificationStatus' 'OutArgument(x:String)') + "`n" + (Prop 'out_notificationError' 'OutArgument(x:String)')
$notifyBody = @"
  <Sequence.Variables><Variable x:TypeArguments=""x:String"" Name=""subject"" /><Variable x:TypeArguments=""x:String"" Name=""body"" /><Variable x:TypeArguments=""x:String"" Name=""server"" /><Variable x:TypeArguments=""x:String"" Name=""sender"" /><Variable x:TypeArguments=""x:Int32"" Name=""port"" /><Variable x:TypeArguments=""x:String"" Name=""sendStatus"" /><Variable x:TypeArguments=""x:String"" Name=""sendError"" /></Sequence.Variables>
  <ui:InvokeCode DisplayName=""Prepare Notification"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $notifyCode)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataRow"" x:Key=""in_row"">[in_row]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_configText"">[in_configText]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_notificationStatus"">[out_notificationStatus]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_notificationError"">[out_notificationError]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_subject"">[subject]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_body"">[body]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_server"">[server]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_port"">[port]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_sender"">[sender]</OutArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
  <If DisplayName=""If Notification Is Configured"" sap2010:WorkflowViewState.IdRef=""If_1""><If.Condition><InArgument x:TypeArguments=""x:Boolean""><VisualBasicValue x:TypeArguments=""x:Boolean"" ExpressionText=""out_notificationStatus = &quot;Ready&quot;"" /></InArgument></If.Condition><If.Then><Sequence DisplayName=""Send Notification""><TryCatch DisplayName=""Send Email with Error Handling"" sap2010:WorkflowViewState.IdRef=""TryCatch_1""><TryCatch.Try><Sequence DisplayName=""Try""><smtp:SendMail DisplayName=""Send SMTP Email"" sap2010:WorkflowViewState.IdRef=""SendMail_1"" Server=""[server]"" Port=""[port]"" Email=""[sender]"" SecurePassword=""[in_securePassword]"" To=""[in_row(&quot;Email&quot;).ToString()]"" Subject=""[subject]"" Body=""[body]"" IsBodyHtml=""False"" SecureConnection=""StartTls"" /></Sequence></TryCatch.Try><TryCatch.Catches><Catch x:TypeArguments=""s:Exception"" sap2010:WorkflowViewState.IdRef=""Catch$([char]96)1_1""><ActivityAction x:TypeArguments=""s:Exception""><ActivityAction.Argument><DelegateInArgument x:TypeArguments=""s:Exception"" Name=""exception"" /></ActivityAction.Argument><Sequence DisplayName=""Catch""><Assign DisplayName=""Set Notification Failed"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""out_notificationStatus"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""&quot;Failed&quot;"" /></InArgument></Assign.Value></Assign><Assign DisplayName=""Capture Notification Error"" sap2010:WorkflowViewState.IdRef=""Assign_2""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""out_notificationError"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""exception.ToString()"" /></InArgument></Assign.Value></Assign></Sequence></ActivityAction></Catch></TryCatch.Catches></TryCatch></Sequence></If.Then></If>
  <ui:InvokeCode DisplayName=""Update Notification Columns"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_2"" Code=""$(Code $notifyFinal)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataRow"" x:Key=""in_row"">[in_row]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""out_notificationStatus"">[out_notificationStatus]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""out_notificationError"">[out_notificationError]</InArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
"@
WriteWorkflow 'SendEmployeeNotification.xaml' 'Workflows_SendEmployeeNotification' $notifyMembers $notifyBody

$writeMembers = (Prop 'in_dtResults' 'InArgument(sd:DataTable)') + "`n" + (Prop 'in_OutputPath' 'InArgument(x:String)') + "`n" + (Prop 'in_SheetName' 'InArgument(x:String)') + "`n" + (Prop 'out_Status' 'OutArgument(x:String)')
$writeBody = @"
  <uea:WriteRange DisplayName=""Write Leave Results"" sap2010:WorkflowViewState.IdRef=""WriteRange_1"" WorkbookPath=""[$(X 'in_OutputPath')]"" SheetName=""[$(X 'in_SheetName')]"" StartingCell=""A1"" AddHeaders=""True"" DataTable=""[$(X 'in_dtResults')]"" />
  <Assign DisplayName=""Set Write Status"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""out_Status"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""&quot;OK - Results written&quot;"" /></InArgument></Assign.Value></Assign>
"@
WriteWorkflow 'WriteResultsToExcel.xaml' 'Workflows_WriteResultsToExcel' $writeMembers $writeBody

$summaryCode = @'
Dim summary = New DataTable()
summary.Columns.Add("Metric", GetType(String))
summary.Columns.Add("Count", GetType(Integer))
Dim add = Sub(name As String, count As Integer)
    summary.Rows.Add(name, count)
End Sub
Dim total = If(dtResults Is Nothing, 0, dtResults.Rows.Count)
add("Total Requests", total)
add("Approved", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("Status").ToString() = "Approved")))
add("Rejected", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("Status").ToString() = "Rejected")))
add("Pending Manager Approval", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("Status").ToString() = "Pending Manager Approval")))
add("Errors", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("Status").ToString() = "Error")))
add("Notifications Sent", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("NotificationStatus").ToString() = "Sent")))
add("Notifications Failed", If(dtResults Is Nothing, 0, dtResults.AsEnumerable().Count(Function(r) r("NotificationStatus").ToString() = "Failed")))
out_dtSummary = summary
'@
$summaryMembers = (Prop 'in_dtResults' 'InArgument(sd:DataTable)') + "`n" + (Prop 'in_OutputPath' 'InArgument(x:String)') + "`n" + (Prop 'in_SheetName' 'InArgument(x:String)') + "`n" + (Prop 'out_Status' 'OutArgument(x:String)')
$summaryBody = @"
  <Sequence.Variables><Variable x:TypeArguments=""sd:DataTable"" Name=""dtSummary"" /></Sequence.Variables>
  <ui:InvokeCode DisplayName=""Generate Summary Data"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $summaryCode)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataTable"" x:Key=""dtResults"">[$(X 'in_dtResults')] </InArgument><OutArgument x:TypeArguments=""sd:DataTable"" x:Key=""out_dtSummary"">[dtSummary]</OutArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
  <uea:WriteRange DisplayName=""Write Summary Workbook"" sap2010:WorkflowViewState.IdRef=""WriteRange_1"" WorkbookPath=""[$(X 'in_OutputPath')]"" SheetName=""[$(X 'in_SheetName')]"" StartingCell=""A1"" AddHeaders=""True"" DataTable=""[dtSummary]"" />
  <Assign DisplayName=""Set Summary Status"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""out_Status"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""&quot;OK - Summary written&quot;"" /></InArgument></Assign.Value></Assign>
"@
WriteWorkflow 'GenerateSummary.xaml' 'Workflows_GenerateSummary' $summaryMembers $summaryBody

$logCode = @'
Dim directory = Path.GetDirectoryName(in_LogPath)
If Not String.IsNullOrWhiteSpace(directory) Then Directory.CreateDirectory(directory)
Dim safeMessage = If(in_Message, "").Replace(vbCr, " ").Replace(vbLf, " ")
Dim line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "|" & If(in_Category, "Unknown") & "|" & If(in_Workflow, "Unknown") & "|" & If(in_RequestID, "") & "|" & safeMessage & Environment.NewLine
File.AppendAllText(in_LogPath, line)
'@
$logMembers = (Prop 'in_LogPath' 'InArgument(x:String)') + "`n" + (Prop 'in_Category' 'InArgument(x:String)') + "`n" + (Prop 'in_Workflow' 'InArgument(x:String)') + "`n" + (Prop 'in_RequestID' 'InArgument(x:String)') + "`n" + (Prop 'in_Message' 'InArgument(x:String)')
$logBody = @"
  <ui:InvokeCode DisplayName=""Append Safe Exception Log"" Language=""VBNet"" sap2010:WorkflowViewState.IdRef=""InvokeCode_1"" Code=""$(Code $logCode)""><ui:InvokeCode.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""x:String"" x:Key=""in_LogPath"">[$(X 'in_LogPath')]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Category"">[$(X 'in_Category')]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Workflow"">[$(X 'in_Workflow')]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_RequestID"">[$(X 'in_RequestID')]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Message"">[$(X 'in_Message')]</InArgument></scg:Dictionary></ui:InvokeCode.Arguments></ui:InvokeCode>
"@
WriteWorkflow 'LogException.xaml' 'Workflows_LogException' $logMembers $logBody

$mainMembers = (Prop 'in_EmailSecurePassword' 'InArgument(ss:SecureString)')
$mainBody = @"
  <Sequence.Variables>
    <Variable x:TypeArguments=""x:String"" Name=""configPath"" /><Variable x:TypeArguments=""x:String"" Name=""inputPath"" /><Variable x:TypeArguments=""x:String"" Name=""outputPath"" /><Variable x:TypeArguments=""x:String"" Name=""summaryPath"" /><Variable x:TypeArguments=""x:String"" Name=""logPath"" /><Variable x:TypeArguments=""x:String"" Name=""configText"" /><Variable x:TypeArguments=""x:String"" Name=""configStatus"" /><Variable x:TypeArguments=""x:String"" Name=""readStatus"" /><Variable x:TypeArguments=""sd:DataTable"" Name=""dtInput"" /><Variable x:TypeArguments=""sd:DataTable"" Name=""dtResults"" /><Variable x:TypeArguments=""x:Boolean"" Name=""schemaValid"" /><Variable x:TypeArguments=""x:String"" Name=""schemaReason"" /><Variable x:TypeArguments=""x:Int32"" Name=""processedCount"" /><Variable x:TypeArguments=""x:Int32"" Name=""errorCount"" /><Variable x:TypeArguments=""x:String"" Name=""notificationStatus"" /><Variable x:TypeArguments=""x:String"" Name=""notificationError"" /><Variable x:TypeArguments=""x:String"" Name=""writeStatus"" /><Variable x:TypeArguments=""x:String"" Name=""summaryStatus"" />
  </Sequence.Variables>
  <Assign DisplayName=""Set Config Path"" sap2010:WorkflowViewState.IdRef=""Assign_1""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""configPath"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""Path.Combine(Environment.CurrentDirectory, &quot;Config.xlsx&quot;)"" /></InArgument></Assign.Value></Assign>
  <Assign DisplayName=""Set Input Path"" sap2010:WorkflowViewState.IdRef=""Assign_2""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""inputPath"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""Path.Combine(Environment.CurrentDirectory, &quot;Data&quot;, &quot;LeaveRequests_Input.xlsx&quot;)"" /></InArgument></Assign.Value></Assign>
  <Assign DisplayName=""Set Output Path"" sap2010:WorkflowViewState.IdRef=""Assign_3""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""outputPath"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""Path.Combine(Environment.CurrentDirectory, &quot;Data&quot;, &quot;LeaveRequests_Output.xlsx&quot;)"" /></InArgument></Assign.Value></Assign>
  <Assign DisplayName=""Set Summary Path"" sap2010:WorkflowViewState.IdRef=""Assign_4""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""summaryPath"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""Path.Combine(Environment.CurrentDirectory, &quot;Data&quot;, &quot;LeaveRequests_Summary.xlsx&quot;)"" /></InArgument></Assign.Value></Assign>
  <Assign DisplayName=""Set Log Path"" sap2010:WorkflowViewState.IdRef=""Assign_5""><Assign.To><OutArgument x:TypeArguments=""x:String""><VisualBasicReference x:TypeArguments=""x:String"" ExpressionText=""logPath"" /></OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=""x:String""><VisualBasicValue x:TypeArguments=""x:String"" ExpressionText=""Path.Combine(Environment.CurrentDirectory, &quot;Logs&quot;, &quot;execution.log&quot;)"" /></InArgument></Assign.Value></Assign>
  <TryCatch DisplayName=""Process Leave Requests"" sap2010:WorkflowViewState.IdRef=""TryCatch_1""><TryCatch.Try><Sequence DisplayName=""Try""><ui:InvokeWorkflowFile DisplayName=""Initialize Configuration"" WorkflowFileName=""Workflows\InitializeConfig.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_1""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""x:String"" x:Key=""in_ConfigPath"">[configPath]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_SheetName"">[&quot;Config&quot;]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_ConfigText"">[configText]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_Status"">[configStatus]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <ui:InvokeWorkflowFile DisplayName=""Read Leave Requests"" WorkflowFileName=""Workflows\ReadLeaveRequests.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_2""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""x:String"" x:Key=""in_InputPath"">[inputPath]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_SheetName"">[&quot;LeaveRequests&quot;]</InArgument><OutArgument x:TypeArguments=""sd:DataTable"" x:Key=""out_dtInput"">[dtInput]</OutArgument><OutArgument x:TypeArguments=""x:Boolean"" x:Key=""out_SchemaValid"">[schemaValid]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_SchemaReason"">[schemaReason]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <ui:LogMessage Level=""Info"" DisplayName=""Log Initialization"" sap2010:WorkflowViewState.IdRef=""LogMessage_1"" Message=""[$&quot;Configuration: {configStatus}; Input schema: {schemaReason}&quot;]"" />
    <ui:InvokeWorkflowFile DisplayName=""Process Decisions"" WorkflowFileName=""Workflows\ProcessLeaveRequests.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_3""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataTable"" x:Key=""in_dtInput"">[dtInput]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_configText"">[configText]</InArgument><OutArgument x:TypeArguments=""sd:DataTable"" x:Key=""out_dtResults"">[dtResults]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_processedCount"">[processedCount]</OutArgument><OutArgument x:TypeArguments=""x:Int32"" x:Key=""out_errorCount"">[errorCount]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <ui:ForEachRow DisplayName=""Send Employee Notifications"" DataTable=""[dtResults]"" sap2010:WorkflowViewState.IdRef=""ForEachRow_1""><ui:ForEachRow.Body><ActivityAction x:TypeArguments=""sd:DataRow""><ActivityAction.Argument><DelegateInArgument x:TypeArguments=""sd:DataRow"" Name=""CurrentRow"" /></ActivityAction.Argument><Sequence DisplayName=""Notification Transaction""><ui:InvokeWorkflowFile DisplayName=""Send Employee Notification"" WorkflowFileName=""Workflows\SendEmployeeNotification.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_4""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataRow"" x:Key=""in_row"">[CurrentRow]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_configText"">[configText]</InArgument><InArgument x:TypeArguments=""ss:SecureString"" x:Key=""in_securePassword"">[in_EmailSecurePassword]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_notificationStatus"">[notificationStatus]</OutArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_notificationError"">[notificationError]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile></Sequence></ActivityAction></ui:ForEachRow.Body></ui:ForEachRow>
    <ui:InvokeWorkflowFile DisplayName=""Write Results Workbook"" WorkflowFileName=""Workflows\WriteResultsToExcel.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_5""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataTable"" x:Key=""in_dtResults"">[dtResults]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_OutputPath"">[outputPath]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_SheetName"">[&quot;LeaveResults&quot;]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_Status"">[writeStatus]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <ui:InvokeWorkflowFile DisplayName=""Generate Summary"" WorkflowFileName=""Workflows\GenerateSummary.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_6""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""sd:DataTable"" x:Key=""in_dtResults"">[dtResults]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_OutputPath"">[summaryPath]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_SheetName"">[&quot;Summary&quot;]</InArgument><OutArgument x:TypeArguments=""x:String"" x:Key=""out_Status"">[summaryStatus]</OutArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile>
    <ui:LogMessage Level=""Info"" DisplayName=""Log Completion"" sap2010:WorkflowViewState.IdRef=""LogMessage_2"" Message=""[$&quot;Processed {processedCount} requests; errors: {errorCount}; {writeStatus}; {summaryStatus}&quot;]"" />
  </Sequence></TryCatch.Try><TryCatch.Catches><Catch x:TypeArguments=""s:Exception"" sap2010:WorkflowViewState.IdRef=""Catch$([char]96)1_1""><ActivityAction x:TypeArguments=""s:Exception""><ActivityAction.Argument><DelegateInArgument x:TypeArguments=""s:Exception"" Name=""exception"" /></ActivityAction.Argument><Sequence DisplayName=""Catch""><ui:InvokeWorkflowFile DisplayName=""Log Main Exception"" WorkflowFileName=""Workflows\LogException.xaml"" sap2010:WorkflowViewState.IdRef=""InvokeWorkflowFile_7""><ui:InvokeWorkflowFile.Arguments><scg:Dictionary x:TypeArguments=""x:String, Argument""><InArgument x:TypeArguments=""x:String"" x:Key=""in_LogPath"">[logPath]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Category"">[&quot;MainWorkflow&quot;]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Workflow"">[&quot;Main.xaml&quot;]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_RequestID"">[&quot;&quot;]</InArgument><InArgument x:TypeArguments=""x:String"" x:Key=""in_Message"">[exception.ToString()]</InArgument></scg:Dictionary></ui:InvokeWorkflowFile.Arguments></ui:InvokeWorkflowFile><Rethrow DisplayName=""Rethrow Main Exception"" sap2010:WorkflowViewState.IdRef=""Rethrow_1"" /></Sequence></ActivityAction></Catch></TryCatch.Catches></TryCatch>
"@
WriteWorkflow 'Main.xaml' 'Main' $mainMembers $mainBody

Get-ChildItem -LiteralPath $wf -Filter '*.xaml' | Select-Object Name, Length

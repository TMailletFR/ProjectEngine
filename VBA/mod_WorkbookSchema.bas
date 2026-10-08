Attribute VB_Name = "mod_WorkbookSchema"
Option Explicit

Public Const CURRENT_SCHEMA_VERSION As Long = 1
Public Const PE_SCHEMA_PROPERTY As String = "ProjectEngine.SchemaVersion"
Public Const PE_RELEASE_PROPERTY As String = "ProjectEngine.ReleaseId"

Public Function WorkbookSchema_Property(ByVal wb As Workbook, ByVal key As String) As Variant
    Dim item As Object
    For Each item In wb.CustomDocumentProperties
        If StrComp(item.Name, key, vbBinaryCompare) = 0 Then
            WorkbookSchema_Property = item.Value
            Exit Function
        End If
    Next item
End Function

Public Sub WorkbookSchema_SetProperty(ByVal wb As Workbook, ByVal key As String, ByVal value As Variant)
    Dim item As Object
    For Each item In wb.CustomDocumentProperties
        If StrComp(item.Name, key, vbBinaryCompare) = 0 Then
            item.Value = value
            Exit Sub
        End If
    Next item
    If VarType(value) = vbLong Or VarType(value) = vbInteger Then
        wb.CustomDocumentProperties.Add key, False, msoPropertyTypeNumber, value
    Else
        wb.CustomDocumentProperties.Add key, False, msoPropertyTypeString, CStr(value)
    End If
End Sub

Public Function WorkbookSchema_Version(ByVal wb As Workbook) As Long
    Dim value As Variant
    value = WorkbookSchema_Property(wb, PE_SCHEMA_PROPERTY)
    If IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) Then Err.Raise 5, "WorkbookSchema", "INVALID_SCHEMA_METADATA"
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < 1 Or CDbl(value) > 2147483647# Then
        Err.Raise 5, "WorkbookSchema", "INVALID_SCHEMA_METADATA"
    End If
    WorkbookSchema_Version = CLng(value)
End Function

' Only build/release tooling calls this; opening or importing never stamps a source.
Public Sub WorkbookSchema_StampBuild(Optional ByVal releaseId As String = "")
    WorkbookSchema_SetProperty ThisWorkbook, PE_SCHEMA_PROPERTY, CURRENT_SCHEMA_VERSION
    If Len(Trim$(releaseId)) > 0 Then WorkbookSchema_SetProperty ThisWorkbook, PE_RELEASE_PROPERTY, releaseId
End Sub

Public Function WorkbookSchema_WritesAllowed() As Boolean
    Dim state As String
    On Error GoTo Blocked
    If WorkbookSchema_Version(ThisWorkbook) <> CURRENT_SCHEMA_VERSION Then Exit Function
    state = CStr(WorkbookSchema_Property(ThisWorkbook, "ProjectEngine.ImportState"))
    If state = "FAILED" Or state = "IMPORT_IN_PROGRESS" Then
        If Not Migration_IsRunning() Then Exit Function
    End If
    WorkbookSchema_WritesAllowed = True
Blocked:
End Function

Public Sub WorkbookSchema_RequireWritable()
    If Not WorkbookSchema_WritesAllowed() Then
        Err.Raise vbObjectError + 5720, "WorkbookSchema", Migration_Text("MIGRATION.BLOCKED")
    End If
End Sub

Public Function WorkbookSchema_UserActionAllowed() As Boolean
    WorkbookSchema_UserActionAllowed = WorkbookSchema_WritesAllowed()
    If Not WorkbookSchema_UserActionAllowed Then
        MsgBox Migration_Text("MIGRATION.BLOCKED"), vbExclamation, Migration_Text("MIGRATION.TITLE")
    End If
End Function

Public Function WorkbookSchema_OpenAllowed() As Boolean
    WorkbookSchema_OpenAllowed = WorkbookSchema_WritesAllowed()
    If Not WorkbookSchema_OpenAllowed Then
        MsgBox Migration_Text("MIGRATION.BLOCKED"), vbExclamation, Migration_Text("MIGRATION.TITLE")
    End If
End Function

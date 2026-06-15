Attribute VB_Name = "mod_WBSButtons"
Option Explicit

'=====================================================
' mod_WBSButtons
'
' WBS main buttons + Task Type setup.
'
' Console routing:
' - aucune MsgBox directe dans ce module
' - les erreurs sont envoyées vers frmPlanningMessages
'=====================================================

Public Sub Ensure_WBS_Main_Buttons()

    Dim ws As Worksheet
    Dim startLeft As Double
    Dim topPos As Double
    Dim btnWidthGreen As Double
    Dim btnWidthRed As Double
    Dim btnWidthForced As Double
    Dim btnHeight As Double
    Dim gap As Double

    Set ws = ThisWorkbook.Worksheets("WBS")

    Ensure_WBS_TaskType_Input_Setup ws
    Ensure_CriticalPathMode_Toggle

    startLeft = ws.Range("A1").Left + 6
    topPos = 8

    btnWidthGreen = 68
    btnWidthRed = 82
    btnWidthForced = 104
    btnHeight = 34
    gap = 8

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Planning", _
        "Update" & vbCrLf & "Planning", _
        "Run_Planning_Update", _
        startLeft, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Gantt", _
        "Update" & vbCrLf & "Gantt", _
        "Run_Gantt_Update", _
        startLeft + btnWidthGreen + gap, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_SCurve", _
        "Update" & vbCrLf & "S-Curve", _
        "Run_SCurve_Update", _
        startLeft + (btnWidthGreen + gap) * 2, _
        topPos, _
        btnWidthGreen, _
        btnHeight, _
        RGB(112, 173, 71)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_ForcedPlanning", _
        "Forced" & vbCrLf & "Planning Update", _
        "Run_Forced_Planning_Update", _
        startLeft + (btnWidthGreen + gap) * 3, _
        topPos, _
        btnWidthForced, _
        btnHeight, _
        RGB(192, 0, 0)

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_Full", _
        "Full" & vbCrLf & "Update", _
        "Run_Full_Update", _
        startLeft + (btnWidthGreen + gap) * 3 + btnWidthForced + gap, _
        topPos, _
        btnWidthRed, _
        btnHeight, _
        RGB(192, 0, 0)

End Sub

Private Sub CreateOrUpdateWBSFloatingButton( _
    ByVal ws As Worksheet, _
    ByVal shpName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal btnWidth As Double, _
    ByVal btnHeight As Double, _
    ByVal fillColor As Long)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, _
            topPos, _
            btnWidth, _
            btnHeight)
        shp.Name = shpName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = btnWidth
    shp.Height = btnHeight
    shp.OnAction = macroName
    shp.Placement = xlFreeFloating

    With shp.TextFrame2
        .TextRange.Text = captionText
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 6
        .MarginRight = 6
        .MarginTop = 2
        .MarginBottom = 2
    End With

    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 1

End Sub

Private Sub Ensure_WBS_TaskType_Input_Setup(ByVal ws As Worksheet)

    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim oldEvents As Boolean

    If ws Is Nothing Then Exit Sub

    On Error GoTo SafeExit

    On Error Resume Next
    Set tbl = ws.ListObjects("tbl_WBS")
    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    If Not WBS_TableHasColumn(tbl, "Task Type") Then
        Err.Raise vbObjectError + 2310, "Ensure_WBS_TaskType_Input_Setup", _
            "Missing required WBS input column: Task Type"
    End If

    Set rng = tbl.ListColumns("Task Type").DataBodyRange

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False

    For Each cell In rng.Cells

        normalizedValue = Normalize_WBS_TaskType_Value(cell.value)

        If Trim$(CStr(cell.value)) = "" Then
            cell.value = "Task"
        ElseIf CStr(cell.value) <> normalizedValue Then
            cell.value = normalizedValue
        End If

    Next cell

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Task,Milestone,Level of Effort"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Task Type"
        .InputMessage = "Choose: Task, Milestone, or Level of Effort."
        .ErrorTitle = "Invalid Task Type"
        .errorMessage = "Allowed values: Task, Milestone, Level of Effort."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

SafeExit:
    Application.EnableEvents = oldEvents

    If Err.Number <> 0 Then
        WBSButtons_ShowConsoleError _
            "Ensure_WBS_TaskType_Input_Setup", _
            "Erreur dans Ensure_WBS_TaskType_Input_Setup : " & Err.Description, _
            "Error in Ensure_WBS_TaskType_Input_Setup: " & Err.Description
    End If

End Sub

Private Function Normalize_WBS_TaskType_Value(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            Normalize_WBS_TaskType_Value = "Task"

        Case "MILESTONE", "MS", "JALON"
            Normalize_WBS_TaskType_Value = "Milestone"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            Normalize_WBS_TaskType_Value = "Level of Effort"

        Case Else
            Normalize_WBS_TaskType_Value = CStr(rawValue)

    End Select

End Function

Private Function WBS_TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    WBS_TableHasColumn = Not col Is Nothing

End Function

Private Sub WBSButtons_ShowConsoleError( _
    ByVal procName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, _
        "STOP", _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub



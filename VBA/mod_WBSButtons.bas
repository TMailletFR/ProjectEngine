Attribute VB_Name = "mod_WBSButtons"
Option Explicit

'===============================================================================
' MODULE : mod_WBSButtons
' DOMAINE / DOMAIN : WBS
'
' FR
' Possede Reset/Armageddon, les boutons WBS, leur langue et la mise en forme des inputs.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns Reset/Armageddon, WBS buttons, their language and input formatting.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Armageddon, Reset_Planning, Ensure_WBS_Main_Buttons, WBS_ApplyLanguage, WBS_SetLanguage, WBS_CurrentLanguage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const RESET_PLANNING_EMPTY_WBS_ROWS As Long = 5
Private gWBSLanguage As String
'------------------------------------------------------------------------------
' FR: Traite la reference Armageddon sans modifier les donnees d'entree.
' EN: Handles the Armageddon reference without mutating input data.
'------------------------------------------------------------------------------

Public Sub Armageddon(Optional ByVal skipConfirmation As Boolean = False)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrHandler

    If Not skipConfirmation Then
        answer = MsgBox( _
            FormatPlanningConsoleMessageForCurrentLanguage(BiMsg( _
                "Cette action va vider le planning, le Dashboard, l'historique et les acquittements." & vbCrLf & _
                "Continuer ?", _
                "This will clear the planning, Dashboard, event history and acknowledgements." & vbCrLf & _
                "Continue?")), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            FormatPlanningConsoleMessageForCurrentLanguage(BiMsg("Full Reset", "Full Reset")))

        If answer <> vbYes Then Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ResetPlanning_PrepareEmptyWBS ws, tbl
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Full_Update
    Reset_Dashboard
    ClearPlanningWarningAcknowledgements
    ClearPlanningEventHistory
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "Armageddon", _
        "Erreur dans Full Reset : " & Err.Description, _
        "Error in Full Reset: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning dans le perimetre possede par le composant.
' EN: Resets Reset Planning within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Reset_Planning()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim answer As VbMsgBoxResult
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo ErrHandler

    answer = MsgBox( _
        FormatPlanningConsoleMessageForCurrentLanguage(BiMsg( _
            "Cette action va vider le WBS et nettoyer les sorties calcul planning." & vbCrLf & _
            "Gantt, S-Curve, Dashboard, historique et acknowledgements ne seront pas modifies." & vbCrLf & vbCrLf & _
            "Continuer ?", _
            "This will clear the WBS and clean planning calculation outputs." & vbCrLf & _
            "Gantt, S-Curve, Dashboard, history and acknowledgements will not be modified." & vbCrLf & vbCrLf & _
            "Continue?") ), _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        FormatPlanningConsoleMessageForCurrentLanguage(BiMsg("Reset Planning", "Reset Planning")))

    If answer <> vbYes Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ResetPlanning_PrepareEmptyWBS ws, tbl

CleanReset:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

    Run_Planning_Update
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0

    WBSButtons_ShowConsoleError _
        "Reset_Planning", _
        "Erreur dans Reset Planning : " & Err.Description, _
        "Error in Reset Planning: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Prepare Empty WBS dans le perimetre possede par le composant.
' EN: Resets Reset Planning Prepare Empty WBS within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_PrepareEmptyWBS( _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject)

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub

    Do While tbl.ListRows.Count > RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

    Do While tbl.ListRows.Count < RESET_PLANNING_EMPTY_WBS_ROWS
        tbl.ListRows.Add
    Loop

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
        ResetPlanning_ApplyWBSInputSetup tbl
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Input Setup dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Input Setup within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSInputSetup(ByVal tbl As ListObject)

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    ResetPlanning_ApplyListValidation tbl, "Task Type", "Task,Milestone,Level of Effort", _
        "Task Type", "Choose: Task, Milestone, or Level of Effort.", _
        "Invalid Task Type", "Allowed values: Task, Milestone, Level of Effort."

    ResetPlanning_ApplyListValidation tbl, "S", "Y,N", _
        "S", "Choose Y to show in Summary, N to hide.", _
        "Invalid S", "Allowed values: blank, Y, N."

    ResetPlanning_ApplyListValidation tbl, "Cal", CALENDAR_7D & "," & CALENDAR_6D & "," & CALENDAR_5D, _
        "Cal", "Choose: 7j/7, 6j/7, or 5j/7.", _
        "Invalid Cal", "Allowed values: blank, 7j/7, 6j/7, 5j/7."

    ResetPlanning_ApplyWBSFormats tbl

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply List Validation dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply List Validation within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyListValidation( _
    ByVal tbl As ListObject, _
    ByVal columnName As String, _
    ByVal listFormula As String, _
    ByVal inputTitle As String, _
    ByVal inputMessage As String, _
    ByVal errorTitle As String, _
    ByVal errorMessage As String)

    Dim rng As Range

    If Not WBS_TableHasColumn(tbl, columnName) Then Exit Sub
    Set rng = tbl.ListColumns(columnName).DataBodyRange
    If rng Is Nothing Then Exit Sub

    rng.NumberFormat = "@"
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=listFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = inputTitle
        .InputMessage = inputMessage
        .ErrorTitle = errorTitle
        .errorMessage = errorMessage
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Reset Planning Apply WBS Formats dans le perimetre possede par le composant.
' EN: Resets Reset Planning Apply WBS Formats within the state owned by the component.
'------------------------------------------------------------------------------

Private Sub ResetPlanning_ApplyWBSFormats(ByVal tbl As ListObject)

    If WBS_TableHasColumn(tbl, "WBS") Then tbl.ListColumns("WBS").DataBodyRange.NumberFormat = "@"
    If WBS_TableHasColumn(tbl, "ID") Then tbl.ListColumns("ID").DataBodyRange.NumberFormat = "0"
    If WBS_TableHasColumn(tbl, "Baseline Start") Then tbl.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Baseline Finish") Then tbl.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Actual Start") Then tbl.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Actual Finish") Then tbl.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Forecast Start") Then tbl.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Forecast Finish") Then tbl.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    If WBS_TableHasColumn(tbl, "Baseline Duration") Then tbl.ListColumns("Baseline Duration").DataBodyRange.NumberFormat = "0"
    If WBS_TableHasColumn(tbl, "Actual Duration") Then tbl.ListColumns("Actual Duration").DataBodyRange.NumberFormat = "0"
    If WBS_TableHasColumn(tbl, "Calculated Duration") Then tbl.ListColumns("Calculated Duration").DataBodyRange.NumberFormat = "0"

End Sub

'=====================================================
' mod_WBSButtons
'
' WBS main buttons + Task Type setup.
'
' Console routing:
' - confirmation utilisateur autorisee pour Reset Planning
' - les erreurs sont envoyées vers frmPlanningMessages
'=====================================================

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Main Buttons si necessaire.
' EN: Ensures or creates WBS Main Buttons when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_WBS_Main_Buttons()

    Dim ws As Worksheet
    Dim startLeft As Double
    Dim topPos As Double
    Dim btnWidthGreen As Double
    Dim btnWidthRed As Double
    Dim btnWidthForced As Double
    Dim btnWidthReset As Double
    Dim btnHeight As Double
    Dim gap As Double

    Set ws = ThisWorkbook.Worksheets("WBS")

    On Error Resume Next
    ws.Shapes("btn_WBS_FullReset").Delete
    On Error GoTo 0

    Ensure_WBS_TaskType_Input_Setup ws
    Ensure_CriticalPathMode_Toggle

    startLeft = ws.Range("A1").Left + 6
    topPos = 8

    btnWidthGreen = 68
    btnWidthRed = 82
    btnWidthForced = 104
    btnWidthReset = 86
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

    CreateOrUpdateWBSFloatingButton _
        ws, _
        "btn_WBS_ResetPlanning", _
        "Reset" & vbCrLf & "Planning", _
        "Reset_Planning", _
        startLeft + (btnWidthGreen + gap) * 3 + btnWidthForced + gap + btnWidthRed + gap, _
        topPos, _
        btnWidthReset, _
        btnHeight, _
        RGB(192, 0, 0)

    WBS_ApplyLanguage

End Sub

'------------------------------------------------------------------------------
' FR: Actualise WBS Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes WBS Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub WBS_ApplyLanguage(Optional ByVal languageCode As String = "")

    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("WBS")

    If Trim$(languageCode) <> "" Then
        WBS_SetLanguage languageCode
    Else
        EnsureWBSLanguageInitialized
    End If

    WBS_SetShapeText ws, "btn_WBS_Planning", WBS_L("Mettre à jour" & vbCrLf & "Planning", "Update" & vbCrLf & "Planning")
    WBS_SetShapeText ws, "btn_WBS_Gantt", WBS_L("Mettre à jour" & vbCrLf & "Gantt", "Update" & vbCrLf & "Gantt")
    WBS_SetShapeText ws, "btn_WBS_SCurve", WBS_L("Mettre à jour" & vbCrLf & "S-Curve", "Update" & vbCrLf & "S-Curve")
    WBS_SetShapeText ws, "btn_WBS_ForcedPlanning", WBS_L("MàJ forcée" & vbCrLf & "Planning", "Forced" & vbCrLf & "Planning Update")
    WBS_SetShapeText ws, "btn_WBS_Full", WBS_L("Mise à jour" & vbCrLf & "complète", "Full" & vbCrLf & "Update")
    WBS_SetShapeText ws, "btn_WBS_ResetPlanning", WBS_L("Réinitialiser" & vbCrLf & "Planning", "Reset" & vbCrLf & "Planning")

    ws.Range("L1").Value = WBS_L("Fin baseline" & vbCrLf & "projet", "Project Baseline" & vbCrLf & "Finish")
    ws.Range("M1").Value = WBS_L("Fin calculée" & vbCrLf & "projet", "Project" & vbCrLf & "Calculated Finish")
    ws.Range("N1").Value = WBS_L("Retard projet" & vbCrLf & "(jours)", "Project Delay" & vbCrLf & "(days)")

    Exit Sub

ErrHandler:
    WBSButtons_ShowConsoleError _
        "WBS_ApplyLanguage", _
        "Erreur dans WBS_ApplyLanguage : " & Err.Description, _
        "Error in WBS_ApplyLanguage: " & Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Language dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub WBS_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gWBSLanguage = "FR"
        Case "EN"
            gWBSLanguage = "EN"
        Case Else
            gWBSLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS Current Language sans modifier les donnees d'entree.
' EN: Returns the WBS Current Language value without mutating input data.
'------------------------------------------------------------------------------

Public Function WBS_CurrentLanguage() As String

    EnsureWBSLanguageInitialized
    WBS_CurrentLanguage = gWBSLanguage

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSLanguage Initialized si necessaire.
' EN: Ensures or creates WBSLanguage Initialized when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSLanguageInitialized()

    If UCase$(Trim$(gWBSLanguage)) <> "FR" And UCase$(Trim$(gWBSLanguage)) <> "EN" Then
        gWBSLanguage = "EN"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur WBS L sans modifier les donnees d'entree.
' EN: Returns the WBS L value without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_L(ByVal frText As String, ByVal enText As String) As String

    If WBS_CurrentLanguage() = "FR" Then
        WBS_L = frText
    Else
        WBS_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise WBS Set Shape Text dans l'etat runtime du composant.
' EN: Activates or initializes WBS Set Shape Text in the component runtime state.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub WBS_SetShapeText( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String)

    Dim shp As Shape

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    shp.TextFrame2.TextRange.Text = captionText

End Sub
'------------------------------------------------------------------------------
' FR: Cree Or Update WBSFloating Button pour le domaine WBS buttons.
' EN: Creates Or Update WBSFloating Button for the WBS buttons domain.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR: Verifie ou cree WBS Task Type Input Setup si necessaire.
' EN: Ensures or creates WBS Task Type Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_WBS_TaskType_Input_Setup(ByVal ws As Worksheet)

    Dim tbl As ListObject
    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim rowIndex As Long
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

        rowIndex = cell.Row - rng.Row + 1

        If Not WBSButtons_RowHasTaskIdentity(tbl, rowIndex) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        Else
            normalizedValue = Normalize_WBS_TaskType_Value(cell.value)

            If Trim$(CStr(cell.value)) = "" Then
                cell.value = "Task"
            ElseIf CStr(cell.value) <> normalizedValue Then
                cell.value = normalizedValue
            End If
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

'------------------------------------------------------------------------------
' FR: Normalise WBS Task Type Value dans un format exploitable.
' EN: Normalizes WBS Task Type Value into a usable format.
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Buttons Row Has Task IDentity sans modifier les donnees d'entree.
' EN: Returns the WBS Buttons Row Has Task IDentity reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSButtons_RowHasTaskIdentity( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long) As Boolean

    Dim idVal As String
    Dim wbsVal As String

    On Error GoTo SafeExit

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tbl.ListRows.Count Then Exit Function

    If WBS_TableHasColumn(tbl, "ID") Then
        idVal = Trim$(CStr(tbl.ListColumns("ID").DataBodyRange.Cells(rowIndex, 1).value))
    End If

    If WBS_TableHasColumn(tbl, "WBS") Then
        wbsVal = Trim$(CStr(tbl.ListColumns("WBS").DataBodyRange.Cells(rowIndex, 1).value))
    End If

    wbsVal = Replace$(wbsVal, ",", ".")
    WBSButtons_RowHasTaskIdentity = (idVal <> "" Or wbsVal <> "")

SafeExit:
End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference WBS Table Has Column sans modifier les donnees d'entree.
' EN: Returns the WBS Table Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function WBS_TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    WBS_TableHasColumn = Not col Is Nothing

End Function

'------------------------------------------------------------------------------
' FR: Projette la collection WBS Buttons Show Console Error vers l'interface autorisee par la politique runtime.
' EN: Projects the WBS Buttons Show Console Error collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

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














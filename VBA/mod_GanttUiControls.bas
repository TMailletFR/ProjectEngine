Attribute VB_Name = "mod_GanttUiControls"
Option Explicit

'===============================================================================
' MODULE : mod_GanttUiControls
' DOMAINE / DOMAIN : Gantt
'
' FR
' Cree et actualise les boutons et toggles Gantt en conservant leurs noms OnAction.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Creates and refreshes Gantt buttons and toggles while preserving their OnAction names.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Ensure_Gantt_Test_Buttons, RefreshFixedHeaderToggleVisuals
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' GANTT button/toggle creation and visual wrappers.
' Shape names and OnAction macro names are preserved.
'=====================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const TOGGLE_ROW_TOP As Long = 2
Private Const TOGGLE_ROW_BOTTOM As Long = 3
Private Const COL_WBS As Long = 1
Private Const BTN_VIEW_BG_NAME As String = "btn_Gantt_View_BG"
Private Const BTN_VIEW_KNOB_NAME As String = "btn_Gantt_View_Knob"
Private Const BTN_VIEW_LEFT_NAME As String = "btn_Gantt_View_Left"
Private Const BTN_VIEW_RIGHT_NAME As String = "btn_Gantt_View_Right"
Private Const BTN_CP_BG_NAME As String = "btn_Gantt_CP_BG"
Private Const BTN_CP_KNOB_NAME As String = "btn_Gantt_CP_Knob"
Private Const BTN_CP_LEFT_NAME As String = "btn_Gantt_CP_Left"
Private Const BTN_CP_RIGHT_NAME As String = "btn_Gantt_CP_Right"
Private Const BTN_CP_MULTI_BG_NAME As String = "shp_GANTT_CPMode_BG"
Private Const BTN_CP_MULTI_KNOB_NAME As String = "shp_GANTT_CPMode_Knob"
Private Const BTN_CP_MULTI_LEFT_NAME As String = "lbl_GANTT_CPMode_Toggle"
Private Const BTN_CONSTRAINT_BG_NAME As String = "btn_Gantt_Constraint_BG"
Private Const BTN_CONSTRAINT_KNOB_NAME As String = "btn_Gantt_Constraint_Knob"
Private Const BTN_CONSTRAINT_LEFT_NAME As String = "btn_Gantt_Constraint_Left"
Private Const BTN_SCALE_BG_NAME As String = "btn_Gantt_Scale_BG"
Private Const BTN_SCALE_KNOB_NAME As String = "btn_Gantt_Scale_Knob"
Private Const BTN_SCALE_LEFT_NAME As String = "btn_Gantt_Scale_Left"
Private Const BTN_SCALE_RIGHT_NAME As String = "btn_Gantt_Scale_Right"
Private Const BTN_SCENARIO_NAME As String = "btn_Gantt_Scenario"
Private Const BTN_SCENARIO_CAPTION As String = "Scenario"
Private Const BTN_TEST_NAME As String = "btn_Gantt_Test"
Private Const BTN_LOCK_NAME As String = "btn_Gantt_Lock"
Private Const BTN_TEST_CAPTION As String = "Test"
Private Const BTN_LOCK_CAPTION As String = "Lock"
Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

'------------------------------------------------------------------------------
' FR:
' Cree ou met a jour les boutons Scenario/Test/Lock et les toggles fixes du header GANTT.
'
' EN:
' Creates or updates Scenario/Test/Lock buttons and fixed header toggles on GANTT.
'
' Entrees / Inputs:
' - Feuille GANTT et etat UI courant.
'
' Sorties / Outputs:
' - Shapes boutons avec OnAction, toggles, libelles localises.
'
' Appele par / Called by:
' - Setup/static layout, ApplyGanttUiState, Gantt_SafeEmptyState.
'
' Notes:
' - Expose des entry points OnAction vers GanttLive et les toggles.
'------------------------------------------------------------------------------
Public Sub Ensure_Gantt_Test_Buttons()

    Dim ws As Worksheet
    Dim scenarioCell As Range
    Dim testCell As Range
    Dim lockCell As Range

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    With ws.Range("A1:H3")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone
    End With

    EnsureGanttViewInitialized

    Set scenarioCell = ws.Range("B1")
    Set testCell = ws.Range("C1")
    Set lockCell = ws.Range("D1")

    CreateOrUpdateGanttButton ws, BTN_SCENARIO_NAME, BTN_SCENARIO_CAPTION, "Run_Gantt_Scenario_Engine", scenarioCell, 0.34, True
    CreateOrUpdateGanttButton ws, BTN_TEST_NAME, BTN_TEST_CAPTION, "Run_Gantt_Test_Engine", testCell, 1, False
    CreateOrUpdateGanttButton ws, BTN_LOCK_NAME, BTN_LOCK_CAPTION, "Run_Gantt_Lock_Changes", lockCell, 1, False

    DeleteShapeIfExists ws, BTN_VIEW_BG_NAME
    DeleteShapeIfExists ws, BTN_VIEW_KNOB_NAME
    DeleteShapeIfExists ws, BTN_VIEW_LEFT_NAME
    DeleteShapeIfExists ws, BTN_VIEW_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_CP_BG_NAME
    DeleteShapeIfExists ws, BTN_CP_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CP_LEFT_NAME
    DeleteShapeIfExists ws, BTN_CP_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_SCALE_BG_NAME
    DeleteShapeIfExists ws, BTN_SCALE_KNOB_NAME
    DeleteShapeIfExists ws, BTN_SCALE_LEFT_NAME
    DeleteShapeIfExists ws, BTN_SCALE_RIGHT_NAME

    DeleteShapeIfExists ws, BTN_CP_MULTI_BG_NAME
    DeleteShapeIfExists ws, BTN_CP_MULTI_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CP_MULTI_LEFT_NAME

    DeleteShapeIfExists ws, BTN_CONSTRAINT_BG_NAME
    DeleteShapeIfExists ws, BTN_CONSTRAINT_KNOB_NAME
    DeleteShapeIfExists ws, BTN_CONSTRAINT_LEFT_NAME

    BuildFixedHeaderToggles ws
    Gantt_ApplyLanguage

End Sub

'------------------------------------------------------------------------------
' FR:
' Reconstruit les toggles de header: Detail/Summary, Day/Week/Month, Constraints, CP/LP et mode multi.
'
' EN:
' Rebuilds header toggles: Detail/Summary, Day/Week/Month, Constraints, CP/LP, and multi mode.
'
' Entrees / Inputs:
' - Feuille GANTT et etats globaux de toggles.
'
' Sorties / Outputs:
' - Shapes track/knob/label avec OnAction vers les macros de toggle.
'
' Appele par / Called by:
' - Ensure_Gantt_Test_Buttons.
'
' Notes:
' - Couple UI GANTT, CriticalPathModeToggle et etat interne du module.
'------------------------------------------------------------------------------
Private Sub BuildFixedHeaderToggles(ByVal ws As Worksheet)

    Dim x As Double
    Dim y As Double

    Dim groupGap As Double
    Dim trackGap As Double

    Dim labelW1 As Double
    Dim labelW2 As Double
    Dim labelW3 As Double
    Dim labelW4 As Double
    Dim labelW5 As Double
    Dim analyticsTrackW As Double

    Dim trackW As Double
    Dim trackH As Double
    Dim knobSize As Double

    x = ws.cells(TOGGLE_ROW_TOP, COL_WBS).Left + 5
    y = ws.cells(TOGGLE_ROW_TOP, COL_WBS).Top + 3

    trackGap = 7
    groupGap = 25
    
    'Detail / Summary
    labelW1 = 70
    
    'None / Critical Path / Longest Path
    labelW2 = 130
    
    'Day / Week / Month
    labelW3 = 78
    
    'Single / Multiple Project
    labelW4 = 95
    
    'Constraint
    labelW5 = 45
    
    trackW = 28
    analyticsTrackW = 46
    trackH = 12
    knobSize = 9

    CreateFixedHeaderToggle ws, _
        BTN_VIEW_LEFT_NAME, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, _
        x, y, labelW1, "Detail / Summary", _
        x + labelW1 + trackGap, y + 2, trackW, trackH, knobSize, _
        (GetGanttViewMode() = GANTT_VIEW_SUMMARY), "Toggle_Gantt_View"

    x = x + labelW1 + trackGap + trackW + groupGap

    CreateFixedHeaderTriToggle ws, _
        BTN_SCALE_LEFT_NAME, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, _
        x, y, labelW3, "Day / Week / Month", _
        x + labelW3 + trackGap, y + 2, analyticsTrackW, trackH, knobSize, _
        GetGanttTimelineScaleMode(), "Toggle_Gantt_Scale"

    x = x + labelW3 + trackGap + analyticsTrackW + groupGap

    CreateFixedHeaderToggle ws, _
        BTN_CONSTRAINT_LEFT_NAME, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, _
        x, y, labelW5, "Constraint", _
        x + labelW5 + trackGap, y + 2, trackW, trackH, knobSize, _
        GetGanttShowConstraints(), "Toggle_Gantt_Constraints"

    x = ws.cells(TOGGLE_ROW_BOTTOM, COL_WBS).Left + 5
    y = ws.cells(TOGGLE_ROW_BOTTOM, COL_WBS).Top + 3

    CreateFixedHeaderTriToggle ws, _
        BTN_CP_LEFT_NAME, BTN_CP_BG_NAME, BTN_CP_KNOB_NAME, _
        x, y, labelW2, "None / Critical Path / Longest Path", _
        x + labelW2 + trackGap, y + 2, analyticsTrackW, trackH, knobSize, _
        GetGanttAnalyticsPathMode(), "Toggle_Gantt_CriticalPath"

    x = x + labelW2 + trackGap + analyticsTrackW + groupGap

    CreateFixedHeaderToggle ws, _
        BTN_CP_MULTI_LEFT_NAME, BTN_CP_MULTI_BG_NAME, BTN_CP_MULTI_KNOB_NAME, _
        x, y, labelW4, "Single / Multiple Project", _
        x + labelW4 + trackGap, y + 2, trackW, trackH, knobSize, _
        IsCriticalPathMultiNetworkEnabled(), "Toggle_CriticalPathMode_FromGantt"

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub CreateFixedHeaderToggle( _
    ByVal ws As Worksheet, _
    ByVal labelName As String, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal labelLeft As Double, _
    ByVal labelTop As Double, _
    ByVal labelWidth As Double, _
    ByVal labelText As String, _
    ByVal trackLeft As Double, _
    ByVal trackTop As Double, _
    ByVal trackWidth As Double, _
    ByVal trackHeight As Double, _
    ByVal knobSize As Double, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)

    DeleteShapeIfExists ws, labelName
    DeleteShapeIfExists ws, bgName
    DeleteShapeIfExists ws, knobName

    Set shpLabel = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, labelLeft, labelTop, labelWidth, trackHeight + 4)
    shpLabel.Name = labelName

    With shpLabel
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = ""
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    shpBg.Name = bgName

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    shpKnob.Name = knobName

    With shpKnob
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    If isOn Then
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub CreateFixedHeaderTriToggle( _
    ByVal ws As Worksheet, _
    ByVal labelName As String, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal labelLeft As Double, _
    ByVal labelTop As Double, _
    ByVal labelWidth As Double, _
    ByVal labelText As String, _
    ByVal trackLeft As Double, _
    ByVal trackTop As Double, _
    ByVal trackWidth As Double, _
    ByVal trackHeight As Double, _
    ByVal knobSize As Double, _
    ByVal modeValue As String, _
    ByVal macroName As String)

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim knobTop As Double
    Dim normalizedMode As String
    Dim onColor As Long

    onColor = RGB(68, 114, 196)
    normalizedMode = UCase$(Trim$(modeValue))
    If bgName = BTN_SCALE_BG_NAME Then
        If normalizedMode <> GANTT_SCALE_WEEK And normalizedMode <> GANTT_SCALE_MONTH Then
            normalizedMode = GANTT_SCALE_DAY
        End If
    Else
        If normalizedMode <> GANTT_ANALYTICS_PATH_CP And normalizedMode <> GANTT_ANALYTICS_PATH_LP Then
            normalizedMode = GANTT_ANALYTICS_PATH_NONE
        End If
    End If

    DeleteShapeIfExists ws, labelName
    DeleteShapeIfExists ws, bgName
    DeleteShapeIfExists ws, knobName

    Set shpLabel = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, labelLeft, labelTop, labelWidth, trackHeight + 4)
    shpLabel.Name = labelName

    With shpLabel
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = ""
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    shpBg.Name = bgName

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    Select Case normalizedMode
        Case GANTT_ANALYTICS_PATH_LP, GANTT_SCALE_MONTH
            knobLeft = trackLeft + trackWidth - knobSize - 2
        Case GANTT_ANALYTICS_PATH_CP, GANTT_SCALE_WEEK
            knobLeft = trackLeft + ((trackWidth - knobSize) / 2)
        Case Else
            knobLeft = trackLeft + 2
    End Select

    Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    shpKnob.Name = knobName

    With shpKnob
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    If normalizedMode = GANTT_ANALYTICS_PATH_NONE Or normalizedMode = GANTT_SCALE_DAY Then
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    Else
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Public Sub RefreshFixedHeaderToggleVisuals(ByVal ws As Worksheet)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("RefreshFixedHeaderToggleVisuals", "Gantt UI")

    EnsureGanttViewInitialized

    RefreshFixedHeaderToggleVisual ws, BTN_VIEW_BG_NAME, BTN_VIEW_KNOB_NAME, (GetGanttViewMode() = GANTT_VIEW_SUMMARY)
    RefreshFixedHeaderTriToggleVisual ws, BTN_SCALE_BG_NAME, BTN_SCALE_KNOB_NAME, GetGanttTimelineScaleMode()
    RefreshFixedHeaderToggleVisual ws, BTN_CONSTRAINT_BG_NAME, BTN_CONSTRAINT_KNOB_NAME, GetGanttShowConstraints()
    RefreshFixedHeaderTriToggleVisual ws, BTN_CP_BG_NAME, BTN_CP_KNOB_NAME, GetGanttAnalyticsPathMode()
    RefreshFixedHeaderToggleVisual ws, BTN_CP_MULTI_BG_NAME, BTN_CP_MULTI_KNOB_NAME, IsCriticalPathMultiNetworkEnabled()

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub RefreshFixedHeaderToggleVisual( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal isOn As Boolean)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim knobLeft As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)

    On Error Resume Next
    Set shpBg = ws.Shapes(bgName)
    Set shpKnob = ws.Shapes(knobName)
    On Error GoTo 0

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    If isOn Then
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
        knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 2
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
        knobLeft = shpBg.Left + 2
    End If

    shpBg.Line.Weight = 1
    shpBg.Shadow.Visible = msoFalse
    shpKnob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpKnob.Line.ForeColor.RGB = RGB(150, 150, 150)
    shpKnob.Line.Weight = 0.75
    shpKnob.Shadow.Visible = msoFalse
    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub RefreshFixedHeaderTriToggleVisual( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal modeValue As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim normalizedMode As String
    Dim knobLeft As Double
    Dim onColor As Long

    onColor = RGB(68, 114, 196)
    normalizedMode = UCase$(Trim$(modeValue))
    If bgName = BTN_SCALE_BG_NAME Then
        If normalizedMode <> GANTT_SCALE_WEEK And normalizedMode <> GANTT_SCALE_MONTH Then
            normalizedMode = GANTT_SCALE_DAY
        End If
    Else
        If normalizedMode <> GANTT_ANALYTICS_PATH_CP And normalizedMode <> GANTT_ANALYTICS_PATH_LP Then
            normalizedMode = GANTT_ANALYTICS_PATH_NONE
        End If
    End If

    On Error Resume Next
    Set shpBg = ws.Shapes(bgName)
    Set shpKnob = ws.Shapes(knobName)
    On Error GoTo 0

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    Select Case normalizedMode
        Case GANTT_ANALYTICS_PATH_LP, GANTT_SCALE_MONTH
            knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 2
        Case GANTT_ANALYTICS_PATH_CP, GANTT_SCALE_WEEK
            knobLeft = shpBg.Left + ((shpBg.Width - shpKnob.Width) / 2)
        Case Else
            knobLeft = shpBg.Left + 2
    End Select

    If normalizedMode = GANTT_ANALYTICS_PATH_NONE Or normalizedMode = GANTT_SCALE_DAY Then
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
    Else
        shpBg.Fill.ForeColor.RGB = onColor
        shpBg.Line.ForeColor.RGB = onColor
    End If

    shpBg.Line.Weight = 1
    shpBg.Shadow.Visible = msoFalse
    shpKnob.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpKnob.Line.ForeColor.RGB = RGB(150, 150, 150)
    shpKnob.Line.Weight = 0.75
    shpKnob.Shadow.Visible = msoFalse
    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

End Sub

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then shp.Delete

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub CreateOrUpdateGanttButton( _
    ByVal ws As Worksheet, _
    ByVal shpName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal topLeftCell As Range, _
    Optional ByVal widthRatio As Double = 1, _
    Optional ByVal alignRightInCell As Boolean = False)

    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim availableWidth As Double

    Set shp = GetShapeIfExists(ws, shpName)

    If widthRatio <= 0 Then widthRatio = 1
    If widthRatio > 1 Then widthRatio = 1

    availableWidth = topLeftCell.Width - 4
    widthVal = availableWidth * widthRatio
    heightVal = topLeftCell.Height - 4

    If widthVal < 36 Then widthVal = 36

    If alignRightInCell Then
        leftPos = topLeftCell.Left + 2 + (availableWidth - widthVal)
    Else
        leftPos = topLeftCell.Left + 2
    End If

    topPos = topLeftCell.Top + 2

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
        shp.Name = shpName
    End If

    shp.Left = leftPos
    shp.Top = topPos
    shp.Width = widthVal
    shp.Height = heightVal
    shp.OnAction = macroName

    With shp
        .Adjustments.item(1) = 0.25
        .Placement = xlMoveAndSize
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 6
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 1.5
        .Shadow.Transparency = 0.45
    End With

    With shp.TextFrame2
        .TextRange.Text = captionText
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.ParagraphFormat.alignment = msoAlignCenter
    End With

    Select Case shpName
        Case BTN_SCENARIO_NAME, BTN_LOCK_NAME
            shp.Fill.ForeColor.RGB = RGB(192, 0, 0)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        Case BTN_TEST_NAME
            shp.Fill.ForeColor.RGB = RGB(112, 173, 71)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End Select

End Sub


'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Private Sub EnsureOrUpdateCompactToggleSwitch( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal leftName As String, _
    ByVal leftCell As Range, _
    ByVal rightCell As Range, _
    ByVal leftText As String, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim shpLeft As Shape

    Dim groupLeft As Double
    Dim groupTop As Double
    Dim groupWidth As Double
    Dim groupHeight As Double

    Dim leftLabelWidth As Double
    Dim gapBeforeTrack As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double

    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Set shpBg = GetShapeIfExists(ws, bgName)
    Set shpKnob = GetShapeIfExists(ws, knobName)
    Set shpLeft = GetShapeIfExists(ws, leftName)

    trackWidth = 28
    trackHeight = 14
    knobSize = 9

    groupLeft = leftCell.Left + 4
    groupTop = leftCell.Top + 2
    groupWidth = (rightCell.Left + rightCell.Width) - groupLeft - 2
    groupHeight = leftCell.Height - 4

    gapBeforeTrack = 4
    leftLabelWidth = groupWidth - trackWidth - gapBeforeTrack
    If leftLabelWidth < 26 Then leftLabelWidth = 26

    trackLeft = groupLeft + leftLabelWidth + gapBeforeTrack
    trackTop = groupTop + ((groupHeight - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 2
    Else
        knobLeft = trackLeft + 2
    End If

    If shpBg Is Nothing Then
        Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    End If

    If shpKnob Is Nothing Then
        Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    End If

    If shpLeft Is Nothing Then
        Set shpLeft = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, groupLeft, groupTop, leftLabelWidth, groupHeight)
    End If

    shpBg.Name = bgName
    shpKnob.Name = knobName
    shpLeft.Name = leftName

    With shpBg
        .Left = trackLeft
        .Top = trackTop
        .Width = trackWidth
        .Height = trackHeight
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(230, 230, 230)
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    With shpKnob
        .Left = knobLeft
        .Top = knobTop
        .Width = knobSize
        .Height = knobSize
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.5
        .Shadow.Visible = msoFalse
    End With

    With shpLeft
        .Left = groupLeft
        .Top = groupTop
        .Width = leftLabelWidth
        .Height = groupHeight
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = leftText
        .TextFrame2.TextRange.Font.Size = 8.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignLeft
    End With

    If isOn Then
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    Else
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
Private Function GetShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape

    On Error Resume Next
    Set GetShapeIfExists = ws.Shapes(shapeName)
    On Error GoTo 0

End Function


'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Private Sub EnsureOrUpdateToggleSwitch( _
    ByVal ws As Worksheet, _
    ByVal bgName As String, _
    ByVal knobName As String, _
    ByVal leftName As String, _
    ByVal rightName As String, _
    ByVal leftCell As Range, _
    ByVal rightCell As Range, _
    ByVal leftText As String, _
    ByVal rightText As String, _
    ByVal isOn As Boolean, _
    ByVal macroName As String)

    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim shpLeft As Shape
    Dim shpRight As Shape

    Dim groupLeft As Double
    Dim groupTop As Double
    Dim groupHeight As Double

    Dim leftLabelWidth As Double
    Dim rightLabelWidth As Double
    Dim gapBeforeTrack As Double
    Dim gapAfterTrack As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double

    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Set shpBg = GetShapeIfExists(ws, bgName)
    Set shpKnob = GetShapeIfExists(ws, knobName)
    Set shpLeft = GetShapeIfExists(ws, leftName)
    Set shpRight = GetShapeIfExists(ws, rightName)

    trackWidth = 42
    trackHeight = 20
    knobSize = 14

    If bgName = BTN_CP_BG_NAME Then
        groupLeft = leftCell.Left + 20
        leftLabelWidth = 88
        rightLabelWidth = 24
        gapBeforeTrack = 10
        gapAfterTrack = 8

    ElseIf bgName = BTN_SCALE_BG_NAME Then
        groupLeft = leftCell.Left + 80
        leftLabelWidth = 26
        rightLabelWidth = 40
        gapBeforeTrack = 10
        gapAfterTrack = 8

    Else
        groupLeft = leftCell.Left + 10
        leftLabelWidth = 34
        rightLabelWidth = 42
        gapBeforeTrack = 8
        gapAfterTrack = 8
    End If

    groupTop = leftCell.Top + 2
    groupHeight = leftCell.Height - 4

    trackLeft = groupLeft + leftLabelWidth + gapBeforeTrack
    trackTop = groupTop + ((groupHeight - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isOn Then
        knobLeft = trackLeft + trackWidth - knobSize - 3
    Else
        knobLeft = trackLeft + 3
    End If

    If shpBg Is Nothing Then
        Set shpBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, trackTop, trackWidth, trackHeight)
    End If

    If shpKnob Is Nothing Then
        Set shpKnob = ws.Shapes.AddShape(msoShapeOval, knobLeft, knobTop, knobSize, knobSize)
    End If

    If shpLeft Is Nothing Then
        Set shpLeft = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, groupLeft, groupTop, leftLabelWidth, groupHeight)
    End If

    If shpRight Is Nothing Then
        Set shpRight = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, trackLeft + trackWidth + gapAfterTrack, groupTop, rightLabelWidth, groupHeight)
    End If

    shpBg.Name = bgName
    shpKnob.Name = knobName
    shpLeft.Name = leftName
    shpRight.Name = rightName

    FormatToggleTrack shpBg, trackLeft, trackTop, macroName
    FormatToggleKnob shpKnob, knobLeft, knobTop, macroName
    FormatToggleLabel shpLeft, groupLeft, groupTop, leftLabelWidth, groupHeight, leftText, macroName
    FormatToggleLabel shpRight, trackLeft + trackWidth + gapAfterTrack, groupTop, rightLabelWidth, groupHeight, rightText, macroName

    If isOn Then
        shpRight.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    Else
        shpLeft.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        shpRight.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub FormatToggleTrack( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 42
        .Height = 20
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(230, 230, 230)
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub FormatToggleKnob( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 14
        .Height = 14
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 4
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 1
        .Shadow.Transparency = 0.45
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Private Sub FormatToggleLabel( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal labelText As String, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.TextRange.Text = labelText
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    End With

End Sub


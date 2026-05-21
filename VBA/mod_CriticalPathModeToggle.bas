Attribute VB_Name = "mod_CriticalPathModeToggle"
Option Explicit

'============================================================
' mod_CriticalPathModeToggle
'
' Critical Path mode switch.
'
' Objectif :
' - Même philosophie que le toggle Analytics.
' - L'utilisateur pilote le mode depuis WBS.
' - Le moteur analytics lit une fonction publique, jamais une cellule en dur.
'
' Backend / fallback cellule :
' - WBS!AD1 : label visible
' - WBS!AE1 : état texte du mode
'
' Valeurs reconnues :
' - GLOBAL
' - MULTI NETWORK
'
' Shapes UI dans WBS :
' - lbl_WBS_CPMode_Toggle
' - shp_WBS_CPMode_BG
' - shp_WBS_CPMode_Knob
'
' Comportement :
' - GLOBAL = comportement actuel : un seul project finish global.
' - MULTI NETWORK = chaque réseau indépendant porte son propre critical path.
'
' Important :
' - Changer le toggle invalide tbl_CALC_STATE en passant Run Status à DIRTY.
' - Le Run_Calc_Engine lancé ensuite force donc un recalcul complet.
'============================================================

Private Const CP_MODE_WS As String = "WBS"
Private Const CP_MODE_LABEL_CELL As String = "AD1"
Private Const CP_MODE_STATE_CELL As String = "AE1"

Private Const CP_STATE_WS As String = "CALC_STATE"
Private Const CP_STATE_TABLE As String = "tbl_CALC_STATE"

Private Const CP_MODE_GLOBAL As String = "GLOBAL"
Private Const CP_MODE_MULTI As String = "MULTI NETWORK"

Private Const CP_LABEL_SHAPE As String = "lbl_WBS_CPMode_Toggle"
Private Const CP_BG_SHAPE As String = "shp_WBS_CPMode_BG"
Private Const CP_KNOB_SHAPE As String = "shp_WBS_CPMode_Knob"

Private Const CP_GANTT_WS As String = "GANTT"

Private Const CP_GANTT_LABEL_SHAPE As String = "lbl_GANTT_CPMode_Toggle"
Private Const CP_GANTT_BG_SHAPE As String = "shp_GANTT_CPMode_BG"
Private Const CP_GANTT_KNOB_SHAPE As String = "shp_GANTT_CPMode_Knob"

Public Function IsCriticalPathMultiNetworkEnabled() As Boolean

    Dim ws As Worksheet
    Dim v As String

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    v = UCase$(Trim$(CStr(ws.Range(CP_MODE_STATE_CELL).value)))
    v = Replace$(v, "-", " ")
    v = Replace$(v, "_", " ")
    v = Replace$(v, "É", "E")

    Do While InStr(1, v, "  ", vbBinaryCompare) > 0
        v = Replace$(v, "  ", " ")
    Loop

    IsCriticalPathMultiNetworkEnabled = _
        (v = CP_MODE_MULTI Or _
         v = "MULTI" Or _
         v = "MULTINETWORK" Or _
         v = "MULTI NETWORKS" Or _
         v = "MULTI RESEAU" Or _
         v = "MULTI RESEAUX" Or _
         v = "ON" Or _
         v = "TRUE" Or _
         v = "1" Or _
         v = "YES" Or _
         v = "OUI")

    Exit Function

SafeExit:
    'Fail-safe volontaire : comportement historique.
    IsCriticalPathMultiNetworkEnabled = False

End Function

Public Sub Toggle_CriticalPathMode()

    Toggle_CriticalPathMode_Core False

End Sub

Public Sub Toggle_CriticalPathMode_FromGantt()

    Toggle_CriticalPathMode_Core True

End Sub

Public Sub Ensure_CriticalPathMode_Toggle()

    Dim ws As Worksheet
    Dim stateText As String

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    stateText = UCase$(Trim$(CStr(ws.Range(CP_MODE_STATE_CELL).value)))

    If stateText = "" Then
        SetCriticalPathModeVisualState False
    Else
        SetCriticalPathModeVisualState IsCriticalPathMultiNetworkEnabled()
    End If

    Ensure_CriticalPathMode_Shapes
    Refresh_CriticalPathMode_Toggle_Visual

SafeExit:
End Sub

Public Sub SetCriticalPathModeVisualState(ByVal multiNetworkEnabled As Boolean)

    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    'AD/AE restent le backend/fallback.
    'Le texte cellule est masqué pour éviter les doublons derrière les shapes.
    With ws.Range(CP_MODE_LABEL_CELL)
        .value = "Multi Critical Path"
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = .Interior.Color
        .Borders.LineStyle = xlNone
    End With

    With ws.Range(CP_MODE_STATE_CELL)
        .NumberFormat = "@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = .Interior.Color
        .Borders.LineStyle = xlNone

        If multiNetworkEnabled Then
            .value = CP_MODE_MULTI
        Else
            .value = CP_MODE_GLOBAL
        End If
    End With

SafeExit:
End Sub

Public Sub Refresh_CriticalPathMode_Toggle_Visual()

    Dim ws As Worksheet
    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape
    Dim isMulti As Boolean
    Dim knobLeft As Double

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)

    On Error Resume Next
    Set shpLabel = ws.Shapes(CP_LABEL_SHAPE)
    Set shpBg = ws.Shapes(CP_BG_SHAPE)
    Set shpKnob = ws.Shapes(CP_KNOB_SHAPE)
    On Error GoTo SafeExit

    If shpBg Is Nothing Then Exit Sub
    If shpKnob Is Nothing Then Exit Sub

    isMulti = IsCriticalPathMultiNetworkEnabled()

    If Not shpLabel Is Nothing Then
        With shpLabel
            .TextFrame2.TextRange.Text = "Multi Critical Path"
            .TextFrame2.TextRange.Font.Size = 9.5
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .OnAction = ""
        End With
    End If

    If isMulti Then
        shpBg.Fill.ForeColor.RGB = RGB(112, 173, 71)
        shpBg.Line.ForeColor.RGB = RGB(112, 173, 71)
        knobLeft = shpBg.Left + shpBg.Width - shpKnob.Width - 3
    Else
        shpBg.Fill.ForeColor.RGB = RGB(230, 230, 230)
        shpBg.Line.ForeColor.RGB = RGB(170, 170, 170)
        knobLeft = shpBg.Left + 3
    End If

    shpBg.OnAction = "Toggle_CriticalPathMode"
    shpKnob.OnAction = "Toggle_CriticalPathMode"

    shpKnob.Left = knobLeft
    shpKnob.Top = shpBg.Top + ((shpBg.Height - shpKnob.Height) / 2)

SafeExit:
End Sub


Private Sub Ensure_CriticalPathMode_Shapes()

    Dim ws As Worksheet
    Dim labelCell As Range
    Dim stateCell As Range

    Dim shpLabel As Shape
    Dim shpBg As Shape
    Dim shpKnob As Shape

    Dim labelLeft As Double
    Dim labelTop As Double
    Dim labelWidth As Double
    Dim labelHeight As Double

    Dim trackWidth As Double
    Dim trackHeight As Double
    Dim knobSize As Double
    Dim trackLeft As Double
    Dim trackTop As Double
    Dim knobLeft As Double
    Dim knobTop As Double

    Dim isMulti As Boolean

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_MODE_WS)
    Set labelCell = ws.Range(CP_MODE_LABEL_CELL)
    Set stateCell = ws.Range(CP_MODE_STATE_CELL)

    'Nettoyage ancienne version + version avec textes Global/Multi.
    DeleteShapeIfExists_CriticalPathMode ws, CP_LABEL_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, CP_BG_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, CP_KNOB_SHAPE
    DeleteShapeIfExists_CriticalPathMode ws, "shp_WBS_CPMode_Left"
    DeleteShapeIfExists_CriticalPathMode ws, "shp_WBS_CPMode_Right"

    isMulti = IsCriticalPathMultiNetworkEnabled()

    labelLeft = labelCell.Left + 6
    labelTop = labelCell.Top + 3
    labelWidth = labelCell.Width + 14
    labelHeight = labelCell.Height - 6

    trackWidth = 42
    trackHeight = 20
    knobSize = 14

    trackLeft = stateCell.Left + ((stateCell.Width - trackWidth) / 2)
    trackTop = stateCell.Top + ((stateCell.Height - trackHeight) / 2)
    knobTop = trackTop + ((trackHeight - knobSize) / 2)

    If isMulti Then
        knobLeft = trackLeft + trackWidth - knobSize - 3
    Else
        knobLeft = trackLeft + 3
    End If

    Set shpLabel = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        labelLeft, _
        labelTop, _
        labelWidth, _
        labelHeight)

    shpLabel.Name = CP_LABEL_SHAPE

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
        .TextFrame2.TextRange.Text = "Multi Critical Path"
        .TextFrame2.TextRange.Font.Size = 9.5
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    Set shpBg = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        trackLeft, _
        trackTop, _
        trackWidth, _
        trackHeight)

    shpBg.Name = CP_BG_SHAPE

    With shpBg
        .Adjustments.item(1) = 0.5
        .OnAction = "Toggle_CriticalPathMode"
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With

    Set shpKnob = ws.Shapes.AddShape( _
        msoShapeOval, _
        knobLeft, _
        knobTop, _
        knobSize, _
        knobSize)

    shpKnob.Name = CP_KNOB_SHAPE

    With shpKnob
        .OnAction = "Toggle_CriticalPathMode"
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

    Refresh_CriticalPathMode_Toggle_Visual

SafeExit:
End Sub

Private Sub ForceFullRecalcAfterCriticalPathModeChange()

    'On invalide le snapshot incrémental.
    'Get_Changed_TaskIds verra Run Status <> OK et forcera un full recalcul.

    SetCalcStateRunStatusDirty

End Sub

Private Sub SetCalcStateRunStatusDirty()

    Dim wsState As Worksheet
    Dim tblState As ListObject
    Dim mapState As Object

    On Error GoTo SafeExit

    Set wsState = ThisWorkbook.Worksheets(CP_STATE_WS)
    Set tblState = wsState.ListObjects(CP_STATE_TABLE)

    If tblState Is Nothing Then Exit Sub
    If tblState.DataBodyRange Is Nothing Then Exit Sub

    Set mapState = BuildColumnMap_CriticalPathMode(tblState)

    If Not mapState.Exists("Run Status") Then Exit Sub

    tblState.ListColumns("Run Status").DataBodyRange.value = "DIRTY"

SafeExit:
End Sub

Private Function BuildColumnMap_CriticalPathMode(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    If tbl Is Nothing Then
        Set BuildColumnMap_CriticalPathMode = d
        Exit Function
    End If

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMap_CriticalPathMode = d

End Function

Private Sub FormatCriticalPathModeToggleTrack( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String, _
    ByVal isOn As Boolean)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 27
        .Height = 12
        .Adjustments.item(1) = 0.5
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Line.Weight = 1
        .Shadow.Visible = msoFalse

        If isOn Then
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .Line.ForeColor.RGB = RGB(68, 114, 196)
        Else
            .Fill.ForeColor.RGB = RGB(230, 230, 230)
            .Line.ForeColor.RGB = RGB(170, 170, 170)
        End If
    End With

End Sub

Private Sub FormatCriticalPathModeToggleKnob( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal macroName As String)

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = 9
        .Height = 9
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 0.75
        .Shadow.Visible = msoFalse
    End With

End Sub

Private Sub FormatCriticalPathModeToggleLabel( _
    ByVal shp As Shape, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal labelText As String, _
    ByVal macroName As String, _
    ByVal isOn As Boolean)

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
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        If isOn Then
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(68, 114, 196)
        Else
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
        End If
    End With

End Sub

Private Sub DeleteShapeIfExists_CriticalPathMode( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then shp.Delete

End Sub

Private Sub Toggle_CriticalPathMode_Core(ByVal refreshGanttAfterCalc As Boolean)

    Dim enabledNext As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo ErrHandler

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    enabledNext = Not IsCriticalPathMultiNetworkEnabled()

    SetCriticalPathModeVisualState enabledNext
    Refresh_CriticalPathMode_Toggle_Visual

    ForceFullRecalcAfterCriticalPathModeChange

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

    Run_Calc_Engine

    If refreshGanttAfterCalc Then
        Refresh_Gantt
    End If

    Refresh_CriticalPathMode_Toggle_Visual

    Exit Sub

ErrHandler:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

    CalcBridge_ShowSingleConsoleMessage _
        "STOP", _
        "Erreur dans Toggle_CriticalPathMode : " & Err.Description, _
        "Error in Toggle_CriticalPathMode: " & Err.Description

End Sub

Public Sub Cleanup_CriticalPathMode_Gantt_Legacy_Shapes()

    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CP_GANTT_WS)

    DeleteShapeIfExists_CriticalPathMode ws, "shp_GANTT_CPMode_Left"
    DeleteShapeIfExists_CriticalPathMode ws, "shp_GANTT_CPMode_Right"

    'Ne supprime pas les 3 shapes actuelles :
    'CP_GANTT_LABEL_SHAPE / CP_GANTT_BG_SHAPE / CP_GANTT_KNOB_SHAPE.
    'Elles sont maintenant repositionnées proprement par BuildFixedHeaderToggles.

SafeExit:
End Sub

Attribute VB_Name = "mod_PEConsoleStyle"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private mLayoutMs As Double
Public Sub PEConsole_Theme(ByVal form As Object)
    form.BackColor = RGB(247, 248, 250)
    ThemeControls form
End Sub
Private Sub ThemeControls(ByVal parent As Object)
    Dim control As Object
    For Each control In parent.Controls
        Select Case TypeName(control)
            Case "TextBox", "ComboBox", "Frame"
                control.BackColor = RGB(247, 248, 250)
                If TypeName(control) = "TextBox" Or TypeName(control) = "Frame" Then
                    control.BorderStyle = fmBorderStyleNone
                    control.SpecialEffect = fmSpecialEffectFlat
                End If
                If TypeName(control) = "ComboBox" Then
                    control.TextAlign = fmTextAlignCenter
                    If control.font.size < 11 Then control.font.size = 11
                End If
            Case "CommandButton"
                If control.font.size < 11 Then control.font.size = 11
            Case "Label"
                If control.name <> "lblMessageType" Then control.BackColor = RGB(247, 248, 250)
        End Select
        If TypeName(control) = "Frame" Then ThemeControls control
    Next control
End Sub

Public Sub PEConsole_Badge(ByVal view As Object)
    Dim badge As MSForms.Label, surface As MSForms.Image, failure As String
    Dim acknowledgement As MSForms.CheckBox
    Dim width As Single, height As Single
    On Error GoTo Classic
    Set badge = view.Controls("lblMessageType")
    Set acknowledgement = view.Controls("chkWarningAck")
    acknowledgement.WordWrap = False: acknowledgement.AutoSize = True
    On Error Resume Next
    Set surface = view.Controls("peSeverity")
    On Error GoTo Classic
    If surface Is Nothing Then Set surface = view.Controls.Add("Forms.Image.1", "peSeverity", True)
    surface.BackStyle = fmBackStyleTransparent: surface.BorderStyle = fmBorderStyleNone
    surface.PictureSizeMode = fmPictureSizeModeStretch: surface.Enabled = False
    If Len(badge.caption) = 0 Then
        surface.Visible = False: badge.Visible = False
        Exit Sub
    End If
    If PEPresentation_ClassicRequested() Then GoTo Classic
    ' Keep the original native label for glyph spacing; the WMF paints only its rounded background.
    badge.WordWrap = False: badge.AutoSize = True
    badge.BorderStyle = fmBorderStyleNone: badge.SpecialEffect = fmSpecialEffectFlat
    badge.BackStyle = fmBackStyleTransparent
    width = badge.width + 16: height = badge.height + 8
    surface.Move badge.left - 8, badge.top - 4, width, height
    Set surface.picture = PEVector(vbNullString, width, height, 4, badge.BackColor, badge.ForeColor, badge.BackColor, False)
    acknowledgement.left = surface.left + surface.width + 10
    acknowledgement.top = surface.top + (surface.height - acknowledgement.height) / 2
    surface.Visible = True: surface.ZOrder 1
    badge.Visible = True: badge.ZOrder 0
    Exit Sub
Classic:
    failure = Err.description
    If Not surface Is Nothing Then surface.Visible = False
    If Not badge Is Nothing Then badge.Visible = Len(badge.caption) > 0
    If Not badge Is Nothing And Not acknowledgement Is Nothing Then
        acknowledgement.left = badge.left + badge.width + 10
        acknowledgement.top = badge.top + (badge.height - acknowledgement.height) / 2
    End If
    If Len(failure) > 0 Then PEPresentation_RecordFallback failure
End Sub
Public Sub PEConsole_ClearBadge(ByVal view As Object)
    Dim surface As MSForms.Image
    On Error Resume Next
    Set surface = view.Controls("peSeverity")
    If Not surface Is Nothing Then
        Set surface.picture = Nothing
        view.Controls.Remove "peSeverity"
        view.Controls("lblMessageType").Visible = True
    End If
    On Error GoTo 0
End Sub
Public Sub PEWelcome_Compact(ByVal view As Object)
    Dim measure As MSForms.TextBox, text As MSForms.TextBox, name As Variant
    Dim height As Single, top As Single, frameHeight As Single
    Set text = view.Controls("txtMessage")
    view.width = 608
    frameHeight = view.height - view.InsideHeight
    view.Controls("lblTitle").Move 16, 12, view.InsideWidth - 32, 24
    text.Move 16, 42, view.InsideWidth - 32, text.height
    Set measure = view.Controls.Add("Forms.TextBox.1", "peWelcomeMeasure", True)
    measure.left = -10000: measure.top = -10000: measure.TabStop = False
    measure.font.name = text.font.name: measure.font.size = text.font.size
    measure.font.bold = text.font.bold: measure.font.italic = text.font.italic
    measure.Multiline = True: measure.WordWrap = True: measure.ScrollBars = fmScrollBarsNone
    height = MeasureText(measure, text.text, text.width) + 8
    If height < 24 Then height = 24
    If height > 105 Then height = 105
    text.height = height: text.ScrollBars = fmScrollBarsVertical
    top = text.top + text.height + 8
    view.Controls("cmdPrevious").Move 16, top, 210, 32
    view.Controls("cmdNext").Move 230, top, 220, 32
    view.Controls("cmdClose").Move 454, top, 126, 32
    view.Controls("cmdClearAck").Move 16, top + 36, 210, 32
    view.Controls("cmdClearHistory").Move 230, top + 36, 220, 32
    view.height = top + 68 + 12 + frameHeight
    view.Controls.Remove "peWelcomeMeasure"
End Sub

' Same wrapped native TextBox measurement as validated 62.2 A, without its B assets or laboratory details panel.
Public Sub PEMessageLayout_Calculate(ByVal form As Object, ByVal messages As Collection, ByRef outerWidth As Single, ByRef outerHeight As Single)
    Dim measure As MSForms.TextBox, clientWidth As Single, bodyHeight As Single
    Dim maxClientHeight As Single, frameWidth As Single, frameHeight As Single
    Dim started As Double
    started = PEClockMs()
    ' Loaded MSForms can report stale client metrics until a real size change.
    form.width = form.width + 1: form.width = form.width - 1
    form.height = form.height + 1: form.height = form.height - 1
    Set measure = form.Controls.Add("Forms.TextBox.1", "peMeasure", True)
    measure.left = -10000: measure.top = -10000: measure.TabStop = False
    measure.font.name = "Segoe UI": measure.font.size = 9
    measure.Multiline = True: measure.WordWrap = True: measure.ScrollBars = fmScrollBarsNone
    clientWidth = 600
    bodyHeight = MeasurePages(measure, messages, clientWidth - 48)
    If bodyHeight > 330 Then
        clientWidth = 648: bodyHeight = MeasurePages(measure, messages, clientWidth - 48)
    End If
    If bodyHeight > 460 Then
        clientWidth = 748: bodyHeight = MeasurePages(measure, messages, clientWidth - 48)
    End If
    maxClientHeight = Application.height - 72
    If maxClientHeight > 680 Then maxClientHeight = 680
    If maxClientHeight < 230 Then maxClientHeight = 230
    bodyHeight = bodyHeight + 8
    If bodyHeight < 24 Then bodyHeight = 24
    If bodyHeight > maxClientHeight - 126 Then bodyHeight = maxClientHeight - 126
    frameWidth = form.width - form.InsideWidth: frameHeight = form.height - form.InsideHeight
    outerWidth = clientWidth + frameWidth: outerHeight = 76 + bodyHeight + 50 + frameHeight
    form.Controls.Remove "peMeasure"
    PEConsole_RecordLayout PEClockMs() - started
End Sub
Private Function MeasurePages(ByVal measure As MSForms.TextBox, ByVal messages As Collection, ByVal width As Single) As Single
    Dim item As Variant, text As String
    MeasurePages = 12
    If messages Is Nothing Then Exit Function
    For Each item In messages
        If IsObject(item) Then
            text = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))
        Else
            text = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item))
        End If
        Dim height As Single
        height = MeasureText(measure, text, width)
        If height > MeasurePages Then MeasurePages = height
    Next item
End Function
Private Function MeasureText(ByVal measure As MSForms.TextBox, ByVal text As String, ByVal width As Single) As Single
    With measure
        .AutoSize = False: .width = width: .height = 12: .text = text
        .AutoSize = True: MeasureText = .height: .AutoSize = False
    End With
End Function
Public Sub PEConsole_RecordLayout(ByVal value As Double)
    mLayoutMs = value
End Sub
Public Function PEConsole_LayoutMs() As Double
    PEConsole_LayoutMs = mLayoutMs
End Function
Public Function PEClockMs() As Double
    Dim counter As Currency, frequency As Currency
    QueryPerformanceCounter counter: QueryPerformanceFrequency frequency
    If frequency <> 0 Then PEClockMs = counter * 1000# / frequency
End Function

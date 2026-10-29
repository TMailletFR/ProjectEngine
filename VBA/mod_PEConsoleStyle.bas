Attribute VB_Name = "mod_PEConsoleStyle"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
Private mLayoutMs As Double
#If VBA7 Then
Private Type PE_UI_POINT
    X As Long
    Y As Long
End Type
Private Type PE_UI_RECT
    left As Long
    top As Long
    right As Long
    Bottom As Long
End Type
Private Type PE_UI_MESSAGE
    hwnd As LongPtr
    message As Long
#If Win64 Then
    padding As Long
#End If
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    point As PE_UI_POINT
    privateData As Long
End Type
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal kind As Long, ByVal callback As LongPtr, ByVal instance As LongPtr, ByVal thread As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hook As LongPtr, ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef process As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hwnd As LongPtr, ByVal text As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, ByRef rectangle As PE_UI_RECT) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, ByRef rectangle As PE_UI_RECT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, ByRef point As PE_UI_POINT) As Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal action As Long, ByVal param As Long, ByRef value As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Sub PEConsole_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal count As LongPtr)
#If Win64 Then
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal point As LongLong) As LongPtr
#Else
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As LongPtr
#End If
Private mWheelHook As LongPtr, mWheelWindow As LongPtr
Private mWheelView As Object, mWheelControls As Collection, mWheelObserver As Object
Private mWheelBusy As Boolean, mWheelRemainder As Long, mWheelLines As Long, mWheelHandled As Long
Private mWheelControl As String, mWheelError As String
#End If
Public Sub PEConsole_WheelAttach(ByVal view As Object, Optional ByVal observer As Object = Nothing)
#If VBA7 Then
    Dim control As Object, hwnd As LongPtr, className As String, length As Long
    On Error GoTo Failed
    PEConsole_WheelDetach
    hwnd = GetActiveWindow()
    className = String$(128, vbNullChar)
    length = GetClassName(hwnd, StrPtr(className), Len(className))
    If left$(className, 7) <> "Thunder" Or length = 0 Then Exit Sub
    Set mWheelControls = New Collection
    For Each control In view.controls
        If TypeName(control) = "ListBox" Or TypeName(control) = "ComboBox" Then mWheelControls.Add control
    Next control
    If mWheelControls.count = 0 Then Exit Sub
    mWheelLines = 3
    Call SystemParametersInfo(&H68, 0, mWheelLines, 0)
    Set mWheelView = view: Set mWheelObserver = observer: mWheelWindow = hwnd
    mWheelRemainder = 0: mWheelControl = "": mWheelError = ""
    ' Current UI thread only; no desktop-wide hook and no work outside an active form.
    mWheelHook = SetWindowsHookEx(3, AddressOf PEConsole_WheelProc, 0, GetCurrentThreadId())
    If mWheelHook = 0 Then Err.Raise 5, "PEConsole_WheelAttach", "Console wheel hook unavailable"
    Exit Sub
Failed:
    mWheelError = CStr(Err.number) & ":" & Err.description
    PEConsole_WheelDetach
#End If
End Sub

Public Sub PEConsole_WheelDetach(Optional ByVal view As Object = Nothing)
#If VBA7 Then
    If Not view Is Nothing Then
        If Not mWheelView Is Nothing Then
            If Not mWheelView Is view Then Exit Sub
        End If
    End If
    If mWheelHook <> 0 Then Call UnhookWindowsHookEx(mWheelHook)
    mWheelHook = 0: mWheelWindow = 0
    Set mWheelControls = Nothing: Set mWheelView = Nothing: Set mWheelObserver = Nothing
    mWheelRemainder = 0: mWheelControl = "": mWheelBusy = False
#End If
End Sub

Public Function PEConsole_WheelState() As String
#If VBA7 Then
    Dim count As Long, message As PE_UI_MESSAGE
    If Not mWheelControls Is Nothing Then count = mWheelControls.count
    PEConsole_WheelState = "hook=" & IIf(mWheelHook <> 0, "True", "False") & ";controls=" & CStr(count) & ";handled=" & CStr(mWheelHandled) & ";message-bytes=" & CStr(LenB(message)) & ";error=" & mWheelError
#Else
    PEConsole_WheelState = "hook=False;controls=0;handled=0;error=VBA7-required"
#End If
End Function

Public Function PEConsole_ScrollList(ByVal control As Object, ByVal lines As Long) As Boolean
    Dim top As Long, nextTop As Long, count As Long
    count = control.ListCount
    If count = 0 Then PEConsole_ScrollList = True: Exit Function
    top = CLng(CallByName(control, "TopIndex", VbGet))
    If top < 0 Then Exit Function
    nextTop = top + lines
    If nextTop < 0 Then nextTop = 0
    If nextTop >= count Then nextTop = count - 1
    ' Scroll the viewport, never the selection or the financial editor's activity.
    If nextTop <> top Then CallByName control, "TopIndex", VbLet, nextTop
    PEConsole_ScrollList = True
End Function

#If VBA7 Then
Private Function FinanceUI_SignedHighWord(ByVal value As Long) As Long
    FinanceUI_SignedHighWord = (value And &H7FFF0000) \ &H10000
    If (value And &H80000000) <> 0 Then FinanceUI_SignedHighWord = FinanceUI_SignedHighWord - &H8000&
End Function

Private Function PEConsole_WheelRoute(ByRef point As PE_UI_POINT, ByVal delta As Long) As Boolean
    Dim client As PE_UI_POINT, rectangle As PE_UI_RECT, popupRect As PE_UI_RECT
    Dim scaleX As Double, scaleY As Double, X As Double, Y As Double, control As Object, hit As Boolean
    Dim underMouse As LongPtr, root As LongPtr, process As Long, steps As Long, lines As Long
#If Win64 Then
    Dim packedPoint As LongLong
    PEConsole_CopyMemory packedPoint, point, 8
    underMouse = WindowFromPoint(packedPoint)
#Else
    underMouse = WindowFromPoint(point.X, point.Y)
#End If
    If mWheelView Is Nothing Or mWheelWindow = 0 Then Exit Function
    If Not mWheelView.visible Then PEConsole_WheelDetach: Exit Function
    If GetClientRect(mWheelWindow, rectangle) = 0 Then Exit Function
    client = point
    If ScreenToClient(mWheelWindow, client) = 0 Then Exit Function
    scaleX = (rectangle.right - rectangle.left) / CDbl(mWheelView.InsideWidth)
    scaleY = (rectangle.Bottom - rectangle.top) / CDbl(mWheelView.InsideHeight)
    If scaleX <= 0 Or scaleY <= 0 Then Exit Function
    X = client.X / scaleX: Y = client.Y / scaleY
    For Each control In mWheelControls
        If control.visible And control.enabled Then
            hit = (X >= control.left And X < control.left + control.width And Y >= control.top And Y < control.top + control.height)
            If TypeName(control) = "ComboBox" Then
                If CLng(CallByName(control, "TopIndex", VbGet)) < 0 Then GoTo NextControl
                If Not hit And Not underMouse = 0 Then
                    root = GetAncestor(underMouse, 2)
                    If root <> mWheelWindow And GetWindowThreadProcessId(root, process) = GetCurrentThreadId() Then
                        If GetWindowRect(root, popupRect) <> 0 Then
                            hit = (point.X >= popupRect.left And point.X < popupRect.right And point.Y >= popupRect.top And point.Y < popupRect.Bottom)
                        End If
                    End If
                End If
            End If
            If hit Then
                If mWheelControl <> control.name Then mWheelRemainder = 0: mWheelControl = control.name
                mWheelRemainder = mWheelRemainder + delta
                steps = Fix(mWheelRemainder / 120)
                mWheelRemainder = mWheelRemainder - steps * 120
                lines = mWheelLines
                If lines = -1 Then
                    lines = Fix(control.height / (control.font.size * 1.5))
                    If lines < 1 Then lines = 1
                End If
                If steps <> 0 And lines > 0 Then
                    If Not PEConsole_ScrollList(control, -steps * lines) Then Exit Function
                    If Not mWheelObserver Is Nothing Then mWheelObserver.PEListScrolled CStr(control.name)
                End If
                PEConsole_WheelRoute = True
                Exit Function
            End If
        End If
NextControl:
    Next control
End Function

Public Function PEConsole_WheelProc(ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim message As PE_UI_MESSAGE, point As PE_UI_POINT, packed As Long, delta As Long
    On Error GoTo Failed
    If code >= 0 And wParam = 1 And Not mWheelBusy And Not mWheelView Is Nothing Then
        PEConsole_CopyMemory message, ByVal lParam, LenB(message)
        If message.message = &H20A Then
            mWheelBusy = True
            PEConsole_CopyMemory packed, message.wParam, 4
            delta = FinanceUI_SignedHighWord(packed)
            PEConsole_CopyMemory packed, message.lParam, 4
            point.X = packed And &HFFFF&
            If point.X > &H7FFF Then point.X = point.X - &H10000
            point.Y = FinanceUI_SignedHighWord(packed)
            If PEConsole_WheelRoute(point, delta) Then
                message.message = 0
                PEConsole_CopyMemory ByVal lParam, message, LenB(message)
                mWheelHandled = mWheelHandled + 1
            End If
            mWheelBusy = False
        End If
    End If
Forward:
    PEConsole_WheelProc = CallNextHookEx(mWheelHook, code, wParam, lParam)
    Exit Function
Failed:
    mWheelBusy = False: mWheelError = CStr(Err.number) & ":" & Err.description
    PEConsole_WheelDetach
    Resume Forward
End Function
#End If

Public Sub PEConsole_Theme(ByVal form As Object)
    form.BackColor = RGB(247, 248, 250)
    ThemeControls form
End Sub
Private Sub ThemeControls(ByVal parent As Object)
    Dim control As Object
    For Each control In parent.controls
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
    Dim badge As MSForms.label, surface As MSForms.Image, failure As String
    Dim acknowledgement As MSForms.CheckBox
    Dim width As Single, height As Single
    On Error GoTo Classic
    Set badge = view.controls("lblMessageType")
    Set acknowledgement = view.controls("chkWarningAck")
    acknowledgement.WordWrap = False: acknowledgement.AutoSize = True
    On Error Resume Next
    Set surface = view.controls("peSeverity")
    On Error GoTo Classic
    If surface Is Nothing Then Set surface = view.controls.Add("Forms.Image.1", "peSeverity", True)
    surface.BackStyle = fmBackStyleTransparent: surface.BorderStyle = fmBorderStyleNone
    surface.PictureSizeMode = fmPictureSizeModeStretch: surface.enabled = False
    If Len(badge.caption) = 0 Then
        surface.visible = False: badge.visible = False
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
    surface.visible = True: surface.ZOrder 1
    badge.visible = True: badge.ZOrder 0
    Exit Sub
Classic:
    failure = Err.description
    If Not surface Is Nothing Then surface.visible = False
    If Not badge Is Nothing Then badge.visible = Len(badge.caption) > 0
    If Not badge Is Nothing And Not acknowledgement Is Nothing Then
        acknowledgement.left = badge.left + badge.width + 10
        acknowledgement.top = badge.top + (badge.height - acknowledgement.height) / 2
    End If
    If Len(failure) > 0 Then PEPresentation_RecordFallback failure
End Sub
Public Sub PEConsole_ClearBadge(ByVal view As Object)
    Dim surface As MSForms.Image
    On Error Resume Next
    Set surface = view.controls("peSeverity")
    If Not surface Is Nothing Then
        Set surface.picture = Nothing
        view.controls.Remove "peSeverity"
        view.controls("lblMessageType").visible = True
    End If
    On Error GoTo 0
End Sub
Public Sub PEWelcome_Compact(ByVal view As Object)
    Dim measure As MSForms.TextBox, text As MSForms.TextBox, name As Variant
    Dim height As Single, top As Single, frameHeight As Single
    Set text = view.controls("txtMessage")
    view.width = 608
    frameHeight = view.height - view.InsideHeight
    view.controls("lblTitle").Move 16, 12, view.InsideWidth - 32, 24
    text.Move 16, 42, view.InsideWidth - 32, text.height
    Set measure = view.controls.Add("Forms.TextBox.1", "peWelcomeMeasure", True)
    measure.left = -10000: measure.top = -10000: measure.TabStop = False
    measure.font.name = text.font.name: measure.font.size = text.font.size
    measure.font.bold = text.font.bold: measure.font.italic = text.font.italic
    measure.Multiline = True: measure.WordWrap = True: measure.ScrollBars = fmScrollBarsNone
    height = MeasureText(measure, text.text, text.width) + 8
    If height < 24 Then height = 24
    If height > 105 Then height = 105
    text.height = height: text.ScrollBars = fmScrollBarsVertical
    top = text.top + text.height + 8
    view.controls("cmdPrevious").Move 16, top, 210, 32
    view.controls("cmdNext").Move 230, top, 220, 32
    view.controls("cmdClose").Move 454, top, 126, 32
    view.controls("cmdClearAck").Move 16, top + 36, 210, 32
    view.controls("cmdClearHistory").Move 230, top + 36, 220, 32
    view.height = top + 68 + 12 + frameHeight
    view.controls.Remove "peWelcomeMeasure"
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
    Set measure = form.controls.Add("Forms.TextBox.1", "peMeasure", True)
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
    form.controls.Remove "peMeasure"
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



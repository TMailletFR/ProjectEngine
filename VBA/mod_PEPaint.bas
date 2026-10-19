Attribute VB_Name = "mod_PEPaint"
Option Explicit
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal flags As Long, ByVal text As LongPtr, ByVal characters As Long, ByVal buffer As LongPtr, ByVal capacity As Long, ByVal defaultCharacter As LongPtr, ByVal usedDefault As LongPtr) As Long
Private Declare PtrSafe Function CreateMetaFileW Lib "gdi32" (ByVal file As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseMetaFile Lib "gdi32" (ByVal dc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateEnhMetaFileW Lib "gdi32" (ByVal reference As LongPtr, ByVal file As LongPtr, ByRef frame As RECT, ByVal description As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" (ByVal dc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" (ByVal metafile As LongPtr) As Long
Private Declare PtrSafe Function GetWinMetaFileBits Lib "gdi32" (ByVal metafile As LongPtr, ByVal size As Long, ByVal buffer As LongPtr, ByVal mapMode As Long, ByVal reference As LongPtr) As Long
Private Declare PtrSafe Function SetMetaFileBitsEx Lib "gdi32" (ByVal size As Long, ByRef buffer As Any) As LongPtr
Private Declare PtrSafe Function GetMetaFileBitsEx Lib "gdi32" (ByVal metafile As LongPtr, ByVal size As Long, ByVal buffer As LongPtr) As Long
Private Declare PtrSafe Function CopyMetaFileW Lib "gdi32" (ByVal metafile As LongPtr, ByVal path As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal dc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal dc As LongPtr, ByVal capability As Long) As Long
Private Declare PtrSafe Function SetViewportExtEx Lib "gdi32" (ByVal dc As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal oldSize As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Target As Any, ByRef source As Any, ByVal size As LongPtr)
Private Declare PtrSafe Function DeleteMetaFile Lib "gdi32" (ByVal metafile As LongPtr) As Long
Private Declare PtrSafe Function SetMapMode Lib "gdi32" (ByVal dc As LongPtr, ByVal mode As Long) As Long
Private Declare PtrSafe Function SetWindowExtEx Lib "gdi32" (ByVal dc As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal oldSize As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal color As Long) As LongPtr
Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal style As Long, ByVal width As Long, ByVal color As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal dc As LongPtr, ByVal obj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal obj As LongPtr) As Long
Private Declare PtrSafe Function RoundRect Lib "gdi32" (ByVal dc As LongPtr, ByVal left As Long, ByVal top As Long, ByVal right As Long, ByVal bottom As Long, ByVal rx As Long, ByVal ry As Long) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal angle As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strike As Long, ByVal charset As Long, ByVal precision As Long, ByVal clip As Long, ByVal quality As Long, ByVal family As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal dc As LongPtr, ByVal mode As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal dc As LongPtr, ByVal color As Long) As Long
Private Declare PtrSafe Function DrawTextW Lib "user32" (ByVal dc As LongPtr, ByVal text As LongPtr, ByVal count As Long, ByRef rectangle As RECT, ByVal flags As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef description As PICTDESC, ByRef iid As GUID, ByVal owns As Long, ByRef picture As IPictureDisp) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal text As LongPtr, ByRef iid As GUID) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef ticks As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function GetGuiResources Lib "user32" (ByVal process As LongPtr, ByVal flags As Long) As Long
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Type GUID
    d1 As Long
    d2 As Integer
    d3 As Integer
    d4(0 To 7) As Byte
End Type
Private Type PICTDESC
    size As Long
    kind As Long
    handle As LongPtr
    xExt As Long
    yExt As Long
End Type
Public Function PEClock() As Double
    Dim ticks As Currency, frequency As Currency
    QueryPerformanceCounter ticks: QueryPerformanceFrequency frequency
    PEClock = ticks / frequency * 1000#
End Function
Public Function PEGdiCount() As Long
    PEGdiCount = GetGuiResources(GetCurrentProcess(), 0)
End Function
Public Function PEWMFCaptionSupported(ByVal text As String) As Boolean
    Dim count As Long, usedDefault As Long, buffer() As Byte, i As Long
    If Len(text) = 0 Then PEWMFCaptionSupported = True: Exit Function
    ' The legacy WMF text is ANSI. UTF-8 ACP is not qualified for its font mapping.
    If GetACP() = 65001 Then
        For i = 1 To Len(text)
            If AscW(Mid$(text, i, 1)) < 0 Or AscW(Mid$(text, i, 1)) > 127 Then Exit Function
        Next i
        PEWMFCaptionSupported = True
        Exit Function
    End If
    count = WideCharToMultiByte(0, &H400, StrPtr(text), Len(text), 0, 0, 0, VarPtr(usedDefault))
    If count <= 0 Then Exit Function
    ReDim buffer(0 To count - 1)
    count = WideCharToMultiByte(0, &H400, StrPtr(text), Len(text), VarPtr(buffer(0)), count, 0, VarPtr(usedDefault))
    PEWMFCaptionSupported = (count > 0 And usedDefault = 0)
End Function
Public Function PEVector(ByVal caption As String, ByVal width As Single, ByVal height As Single, _
    ByVal radius As Single, ByVal fill As Long, ByVal ink As Long, ByVal edge As Long, _
    ByVal focused As Boolean, Optional ByVal fontSize As Single = 12, _
    Optional ByVal fontName As String = "Segoe UI", Optional ByVal fontBold As Boolean = False, _
    Optional ByVal fontItalic As Boolean = False, Optional ByVal choice As Boolean = False) As IPictureDisp
    Dim dc As LongPtr, metafile As LongPtr, brush As LongPtr, pen As LongPtr, font As LongPtr
    Dim oldBrush As LongPtr, oldPen As LongPtr, oldFont As LongPtr
    Dim w As Long, h As Long, r As Long, bounds As RECT, measured As RECT
    Dim description As PICTDESC, iid As GUID, picture As IPictureDisp
    Dim hr As Long, errorNumber As Long, errorText As String, textHeight As Long
    Dim reference As LongPtr, enhanced As LongPtr, frame As RECT, bits() As Byte, byteCount As Long
    Dim px As Long, py As Long, mmX As Long, mmY As Long, resX As Long, resY As Long
    On Error GoTo Failed
    If Not PEWMFCaptionSupported(caption) Then Err.Raise vbObjectError + 625, , "WMF text conversion would lose characters"
    ' No background rectangle: the metafile contains only the rounded path and text.
    w = CLng(width * 4): h = CLng(height * 4): r = CLng(radius * 4)
    If w < 32 Or h < 32 Then Err.Raise 5, , "Rounded surface minimum size is 8 pt"
    If r < 0 Then r = 0
    If r > h / 2 Then r = h / 2
    If r > w / 2 Then r = w / 2
    reference = CreateCompatibleDC(0)
    If reference = 0 Then Err.Raise 5, , "Reference DC failed"
    px = CLng(width * GetDeviceCaps(reference, 88) / 72#)
    py = CLng(height * GetDeviceCaps(reference, 90) / 72#)
    mmX = GetDeviceCaps(reference, 4): mmY = GetDeviceCaps(reference, 6)
    resX = GetDeviceCaps(reference, 8): resY = GetDeviceCaps(reference, 10)
    If mmX <= 0 Or mmY <= 0 Or resX <= 0 Or resY <= 0 Then Err.Raise 5, , "Metafile device metrics unavailable"
    ' EMF frame describes the actual device pixels; PICTDESC below describes the requested OLE size.
    frame.right = CLng(CDbl(px) * mmX * 100# / resX)
    frame.bottom = CLng(CDbl(py) * mmY * 100# / resY)
    dc = CreateEnhMetaFileW(reference, 0, frame, 0)
    If dc = 0 Then Err.Raise 5, , "CreateEnhMetaFile failed"
    SetMapMode dc, 8: SetWindowExtEx dc, w, h, 0
    SetViewportExtEx dc, px, py, 0
    brush = CreateSolidBrush(fill): pen = CreatePen(0, IIf(focused, 6, 2), edge)
    oldBrush = SelectObject(dc, brush): oldPen = SelectObject(dc, pen)
    RoundRect dc, 5, 5, w - 5, h - 5, r * 2, r * 2
    If fontSize <= 0 Then Err.Raise 5, , "FontSize must be positive"
    SetBkMode dc, 1: SetTextColor dc, ink
    CaptionLayout dc, caption, w, h, choice, fontName, fontSize, fontBold, fontItalic, True
    SelectObject dc, oldPen: SelectObject dc, oldBrush
    DeleteObject pen: pen = 0: DeleteObject brush: brush = 0
    enhanced = CloseEnhMetaFile(dc): dc = 0
    If enhanced = 0 Then Err.Raise 5, , "CloseEnhMetaFile failed"
    byteCount = GetWinMetaFileBits(enhanced, 0, 0, 8, reference)
    If byteCount <= 0 Then Err.Raise 5, , "WMF conversion size failed"
    ReDim bits(0 To byteCount - 1)
    If GetWinMetaFileBits(enhanced, byteCount, VarPtr(bits(0)), 8, reference) <> byteCount Then Err.Raise 5, , "WMF conversion failed"
    metafile = SetMetaFileBitsEx(byteCount, bits(0))
    DeleteEnhMetaFile enhanced: enhanced = 0
    DeleteDC reference: reference = 0
    If metafile = 0 Then Err.Raise 5, , "WMF creation failed"
    description.size = LenB(description): description.kind = 2: description.handle = metafile
    ' PICTDESC.wmf extents are twips; the EMF frame above remains HIMETRIC.
    description.xExt = CLng(width * 20#): description.yExt = CLng(height * 20#)
    CLSIDFromString StrPtr("{7BF80981-BF32-101A-8BBB-00AA00300CAB}"), iid
    hr = OleCreatePictureIndirect(description, iid, 1, picture)
    If hr <> 0 Then Err.Raise hr, , "OleCreatePictureIndirect failed"
    metafile = 0
    Set PEVector = picture
    Exit Function
Failed:
    errorNumber = Err.number: errorText = Err.description
    On Error Resume Next
    If dc <> 0 Then
        If oldFont <> 0 Then SelectObject dc, oldFont
        If oldPen <> 0 Then SelectObject dc, oldPen
        If oldBrush <> 0 Then SelectObject dc, oldBrush
        enhanced = CloseEnhMetaFile(dc)
    End If
    If font <> 0 Then DeleteObject font
    If pen <> 0 Then DeleteObject pen
    If brush <> 0 Then DeleteObject brush
    If metafile <> 0 Then DeleteMetaFile metafile
    If enhanced <> 0 Then DeleteEnhMetaFile enhanced
    If reference <> 0 Then DeleteDC reference
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Function
Private Function TextBlock(ByVal dc As LongPtr, ByVal text As String, ByRef bounds As RECT, _
    ByVal name As String, ByVal size As Single, ByVal bold As Boolean, ByVal italic As Boolean, _
    ByVal centered As Boolean, ByVal draw As Boolean) As Long
    Dim font As LongPtr, previous As LongPtr, measured As RECT, flags As Long
    Dim number As Long, description As String
    On Error GoTo Failed
    font = CreateFontW(-CLng(size * 4), 0, 0, 0, IIf(bold, 700, 400), Abs(CLng(italic)), 0, 0, 1, 0, 0, 4, 0, StrPtr(name))
    If font = 0 Then Err.Raise 5, , "Caption font unavailable"
    previous = SelectObject(dc, font)
    ' OLE WMF playback retains DrawText clipping across blocks. Layout is checked before drawing.
    flags = &H910 Or Abs(CLng(centered))
    measured = bounds
    TextBlock = DrawTextW(dc, StrPtr(text), Len(text), measured, flags Or &H400)
    If Len(text) > 0 And TextBlock <= 0 Then Err.Raise 5, , "Caption measurement failed"
    If measured.right > bounds.right Then Err.Raise 5, , "Caption exceeds available width"
    If draw And Len(text) > 0 Then
        If DrawTextW(dc, StrPtr(text), Len(text), bounds, flags) <= 0 Then Err.Raise 5, , "Caption recording failed"
    End If
    SelectObject dc, previous: DeleteObject font
    Exit Function
Failed:
    number = Err.number: description = Err.description
    If previous <> 0 Then SelectObject dc, previous
    If font <> 0 Then DeleteObject font
    Err.Raise number, , description
End Function
Private Function CaptionLayout(ByVal dc As LongPtr, ByVal caption As String, ByVal w As Long, ByVal h As Long, _
    ByVal choice As Boolean, ByVal name As String, ByVal size As Single, ByVal bold As Boolean, _
    ByVal italic As Boolean, ByVal draw As Boolean) As Long
    Dim title As String, body As String, separator As Long, titleHeight As Long, bodyHeight As Long
    Dim bounds As RECT
    bounds.bottom = 32760
    If choice Then
        separator = InStr(caption, vbCrLf & vbCrLf)
        If separator > 0 Then
            title = left$(caption, separator - 1): body = Mid$(caption, separator + 4)
        Else
            title = caption
        End If
        bounds.left = 48: bounds.right = w - 48
        titleHeight = TextBlock(dc, title, bounds, name, size * 4! / 3!, True, False, False, False)
        bodyHeight = TextBlock(dc, body, bounds, name, size, False, False, False, False)
        CaptionLayout = 96 + titleHeight + 32 + bodyHeight
        If draw Then
            If CaptionLayout > h Then Err.Raise 5, , "Choice layout needs more height"
            bounds.top = 48: bounds.bottom = bounds.top + titleHeight
            TextBlock dc, title, bounds, name, size * 4! / 3!, True, False, False, True
            bounds.top = bounds.bottom + 32: bounds.bottom = bounds.top + bodyHeight
            TextBlock dc, body, bounds, name, size, False, False, False, True
        End If
    Else
        bounds.left = 28: bounds.right = w - 28
        bodyHeight = TextBlock(dc, caption, bounds, name, size, bold, italic, True, False)
        CaptionLayout = bodyHeight + 8
        If draw Then
            If CaptionLayout > h Then Err.Raise 5, , "Caption does not fit PEButton"
            bounds.top = (h - bodyHeight) \ 2: bounds.bottom = bounds.top + bodyHeight
            TextBlock dc, caption, bounds, name, size, bold, italic, True, True
        End If
    End If
End Function
Public Function PEChoiceHeight(ByVal caption As String, ByVal width As Single, ByVal fontName As String, ByVal fontSize As Single) As Single
    Dim dc As LongPtr, failure As Long, description As String
    On Error GoTo Failed
    dc = CreateCompatibleDC(0)
    If dc = 0 Then Err.Raise 5, , "Choice measurement DC failed"
    SetMapMode dc, 8
    SetWindowExtEx dc, CLng(width * 4), 4096, 0
    SetViewportExtEx dc, CLng(width * GetDeviceCaps(dc, 88) / 72#), CLng(1024 * GetDeviceCaps(dc, 90) / 72#), 0
    PEChoiceHeight = CaptionLayout(dc, caption, CLng(width * 4), 0, True, fontName, fontSize, False, False, False) / 4! + 4
    DeleteDC dc
    Exit Function
Failed:
    failure = Err.number: description = Err.description
    If dc <> 0 Then DeleteDC dc
    Err.Raise failure, , description
End Function
Private Function TextRecords(ByVal metafile As LongPtr) As Long
    Dim bits() As Byte, size As Long, position As Long, words As Long, operation As Integer
    size = GetMetaFileBitsEx(metafile, 0, 0)
    If size < 18 Then Err.Raise 5, , "WMF header missing"
    ReDim bits(0 To size - 1)
    If GetMetaFileBitsEx(metafile, size, VarPtr(bits(0))) <> size Then Err.Raise 5, , "WMF read failed"
    position = 18
    Do While position + 6 <= size
        CopyMemory words, bits(position), 4
        CopyMemory operation, bits(position + 4), 2
        If words < 3 Or CDbl(position) + CDbl(words) * 2 > size Then Err.Raise 5, , "Invalid WMF record"
        If operation = &H521 Or operation = &HA32 Then TextRecords = TextRecords + 1
        position = position + words * 2
    Loop
End Function
Public Function PETextRecordingGate() As String
    Dim dc As LongPtr, metafile As LongPtr, bounds As RECT, oldMeasure As Long, oldDraw As Long, oldRecords As Long
    Dim picture As IPictureDisp, fileCopy As LongPtr, scenario As Variant, captions As Variant, index As Long, count As Long
    bounds.right = 824: bounds.bottom = 184
    dc = CreateMetaFileW(0)
    oldMeasure = DrawTextW(dc, StrPtr("Continuer"), 9, bounds, &HC11)
    oldDraw = DrawTextW(dc, StrPtr("Continuer"), 9, bounds, &H811)
    metafile = CloseMetaFile(dc)
    oldRecords = TextRecords(metafile): DeleteMetaFile metafile
    PETextRecordingGate = "old_measure=" & oldMeasure & ";old_draw=" & oldDraw & ";old_text_records=" & oldRecords
    captions = Array("Continuer", "Précédent", "Continue", "Previous", "Continuer avec le projet en cours", "Continue with the current project", "R&D")
    For index = 0 To UBound(captions)
        Set picture = PEVector(CStr(captions(index)), 206, 46, 8, RGB(20, 105, 98), vbWhite, RGB(20, 105, 98), False)
        count = TextRecords(CLngPtr(picture.handle))
        If count = 0 Then Err.Raise 5, , "No text records: " & captions(index)
        PETextRecordingGate = PETextRecordingGate & vbCrLf & captions(index) & ":text_records=" & count
        fileCopy = CopyMetaFileW(CLngPtr(picture.handle), StrPtr(ThisWorkbook.path & "\caption-proof-" & index & ".wmf"))
        If fileCopy = 0 Then Err.Raise 5, , "WMF proof export failed"
        DeleteMetaFile fileCopy
    Next index
End Function

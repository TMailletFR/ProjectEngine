Attribute VB_Name = "mod_CalendarEngine"
Option Explicit

'=====================================================
' mod_CalendarEngine
' Pure calendar helpers for planning calculations.
' No worksheet access, no side effects.
'
' V1 calendars:
' - blank / 7j/7 = every day
' - 6j/7 = Monday-Saturday
' - 5j/7 = Monday-Friday
'
' Duration convention is inclusive:
' AddWorkingDays(start, 1, cal) = normalized start.
'=====================================================

Public Const CALENDAR_7D As String = "7j/7"
Public Const CALENDAR_6D As String = "6j/7"
Public Const CALENDAR_5D As String = "5j/7"

Public Function NormalizeCalendarType(ByVal rawValue As Variant) As String

    Dim txt As String

    txt = LCase$(Trim$(CStr(rawValue)))
    txt = Replace(txt, " ", vbNullString)

    Select Case txt
        Case vbNullString, "7j/7", "7/7", "7j", "7d/7", "7d"
            NormalizeCalendarType = CALENDAR_7D

        Case "6j/7", "6/7", "6j", "6d/7", "6d"
            NormalizeCalendarType = CALENDAR_6D

        Case "5j/7", "5/7", "5j", "5d/7", "5d"
            NormalizeCalendarType = CALENDAR_5D

        Case Else
            Err.Raise vbObjectError + 7601, "NormalizeCalendarType", _
                "Invalid calendar type: " & CStr(rawValue)
    End Select

End Function

Public Function IsValidCalendarType(ByVal rawValue As Variant) As Boolean

    On Error GoTo InvalidValue
    NormalizeCalendarType rawValue
    IsValidCalendarType = True
    Exit Function

InvalidValue:
    IsValidCalendarType = False

End Function

Public Function IsWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Boolean

    Dim calType As String
    Dim weekdayIndex As Long

    If Not HasValue(dateValue) Then Exit Function

    calType = NormalizeCalendarType(calendarType)
    weekdayIndex = Weekday(CDate(dateValue), vbMonday)

    Select Case calType
        Case CALENDAR_7D
            IsWorkingDay = True

        Case CALENDAR_6D
            IsWorkingDay = (weekdayIndex <= 6)

        Case CALENDAR_5D
            IsWorkingDay = (weekdayIndex <= 5)
    End Select

End Function

Public Function NextWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim d As Date
    Dim calType As String

    If Not HasValue(dateValue) Then
        NextWorkingDay = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = CDate(dateValue) + 1

    Do While Not IsWorkingDay(d, calType)
        d = d + 1
    Loop

    NextWorkingDay = CDbl(d)

End Function

Public Function PreviousWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim d As Date
    Dim calType As String

    If Not HasValue(dateValue) Then
        PreviousWorkingDay = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = CDate(dateValue) - 1

    Do While Not IsWorkingDay(d, calType)
        d = d - 1
    Loop

    PreviousWorkingDay = CDbl(d)

End Function

Public Function AddWorkingDays( _
    ByVal startDate As Variant, _
    ByVal durationDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim remaining As Long
    Dim d As Date

    If Not HasValue(startDate) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(durationDays) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    If Not IsNumeric(durationDays) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    remaining = CLng(durationDays)
    If remaining <= 0 Then
        AddWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = NormalizeToWorkingDayForward(startDate, calType)
    remaining = remaining - 1

    Do While remaining > 0
        d = CDate(NextWorkingDay(d, calType))
        remaining = remaining - 1
    Loop

    AddWorkingDays = CDbl(d)

End Function

Public Function SubtractWorkingDays( _
    ByVal finishDate As Variant, _
    ByVal durationDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim remaining As Long
    Dim d As Date

    If Not HasValue(finishDate) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(durationDays) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    If Not IsNumeric(durationDays) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    remaining = CLng(durationDays)
    If remaining <= 0 Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = NormalizeToWorkingDayBackward(finishDate, calType)
    remaining = remaining - 1

    Do While remaining > 0
        d = CDate(PreviousWorkingDay(d, calType))
        remaining = remaining - 1
    Loop

    SubtractWorkingDays = CDbl(d)

End Function

Public Function DateDiffWorkingDays( _
    ByVal startDate As Variant, _
    ByVal finishDate As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim startSerial As Long
    Dim finishSerial As Long

    If Not HasValue(startDate) Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(finishDate) Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    startSerial = CLng(CDbl(CDate(startDate)))
    finishSerial = CLng(CDbl(CDate(finishDate)))

    If finishSerial < startSerial Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    Select Case calType
        Case CALENDAR_7D
            DateDiffWorkingDays = finishSerial - startSerial + 1
        Case CALENDAR_6D
            DateDiffWorkingDays = CountWorkingDaysFast(startSerial, finishSerial, 6)
        Case CALENDAR_5D
            DateDiffWorkingDays = CountWorkingDaysFast(startSerial, finishSerial, 5)
        Case Else
            Err.Raise vbObjectError + 7603, "DateDiffWorkingDays", _
                "Unsupported calendar type: " & calType
    End Select

End Function

Public Function ApplyLag( _
    ByVal baseDate As Variant, _
    ByVal lagDays As Variant, _
    Optional ByVal calendarType As Variant = "", _
    Optional ByVal linkType As String = "FS") As Variant

    Dim calType As String
    Dim lagVal As Long
    Dim anchorDate As Variant

    If Not HasValue(baseDate) Then
        ApplyLag = Empty
        Exit Function
    End If

    If Not HasValue(lagDays) Then
        lagVal = 0
    ElseIf IsNumeric(lagDays) Then
        lagVal = CLng(lagDays)
    Else
        ApplyLag = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)

    Select Case UCase$(Trim$(linkType))
        Case "FS", ""
            anchorDate = NextWorkingDay(baseDate, calType)
        Case "SS", "FF"
            anchorDate = NormalizeToWorkingDayForward(baseDate, calType)
        Case Else
            Err.Raise vbObjectError + 7602, "ApplyLag", _
                "Unsupported link type for calendar lag: " & linkType
    End Select

    ApplyLag = ShiftWorkingDays(anchorDate, lagVal, calType)

End Function

Public Function OffsetWorkingDays( _
    ByVal dateValue As Variant, _
    ByVal offsetDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim offsetVal As Long
    Dim calType As String

    If Not HasValue(dateValue) Then
        OffsetWorkingDays = Empty
        Exit Function
    End If

    If HasValue(offsetDays) Then
        offsetVal = CLng(CDbl(offsetDays))
    Else
        offsetVal = 0
    End If

    calType = NormalizeCalendarType(calendarType)
    OffsetWorkingDays = ShiftWorkingDays(dateValue, offsetVal, calType)

End Function

Public Function SignedWorkingDayOffset( _
    ByVal requiredDate As Variant, _
    ByVal actualDate As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim diffVal As Variant
    Dim calType As String

    If Not HasValue(requiredDate) Then
        SignedWorkingDayOffset = Empty
        Exit Function
    End If

    If Not HasValue(actualDate) Then
        SignedWorkingDayOffset = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)

    If CLng(CDbl(CDate(actualDate))) >= CLng(CDbl(CDate(requiredDate))) Then
        diffVal = DateDiffWorkingDays(requiredDate, actualDate, calType)
        If HasValue(diffVal) Then
            SignedWorkingDayOffset = CLng(diffVal) - 1
            If CLng(CDbl(CDate(actualDate))) > CLng(CDbl(CDate(requiredDate))) Then
                If CLng(SignedWorkingDayOffset) = 0 Then SignedWorkingDayOffset = 1
            End If
        End If
    Else
        diffVal = DateDiffWorkingDays(actualDate, requiredDate, calType)
        If HasValue(diffVal) Then
            SignedWorkingDayOffset = -(CLng(diffVal) - 1)
            If CLng(SignedWorkingDayOffset) = 0 Then SignedWorkingDayOffset = -1
        End If
    End If

End Function

Private Function ShiftWorkingDays( _
    ByVal dateValue As Variant, _
    ByVal offsetDays As Long, _
    ByVal calendarType As String) As Variant

    Dim d As Date
    Dim i As Long

    If Not HasValue(dateValue) Then
        ShiftWorkingDays = Empty
        Exit Function
    End If

    If offsetDays >= 0 Then
        d = NormalizeToWorkingDayForward(dateValue, calendarType)
        For i = 1 To offsetDays
            d = CDate(NextWorkingDay(d, calendarType))
        Next i
    Else
        d = NormalizeToWorkingDayBackward(dateValue, calendarType)
        For i = 1 To Abs(offsetDays)
            d = CDate(PreviousWorkingDay(d, calendarType))
        Next i
    End If

    ShiftWorkingDays = CDbl(d)

End Function

Private Function NormalizeToWorkingDayForward( _
    ByVal dateValue As Variant, _
    ByVal calendarType As String) As Date

    Dim d As Date

    d = CDate(dateValue)
    Do While Not IsWorkingDay(d, calendarType)
        d = d + 1
    Loop

    NormalizeToWorkingDayForward = d

End Function

Private Function NormalizeToWorkingDayBackward( _
    ByVal dateValue As Variant, _
    ByVal calendarType As String) As Date

    Dim d As Date

    d = CDate(dateValue)
    Do While Not IsWorkingDay(d, calendarType)
        d = d - 1
    Loop

    NormalizeToWorkingDayBackward = d

End Function

Private Function CountWorkingDaysFast( _
    ByVal startSerial As Long, _
    ByVal finishSerial As Long, _
    ByVal workDaysPerWeek As Long) As Long

    Dim totalDays As Long
    Dim fullWeeks As Long
    Dim remainderDays As Long
    Dim i As Long
    Dim d As Date

    totalDays = finishSerial - startSerial + 1
    fullWeeks = totalDays \ 7
    remainderDays = totalDays Mod 7

    CountWorkingDaysFast = fullWeeks * workDaysPerWeek

    For i = 0 To remainderDays - 1
        d = CDate(startSerial + i)
        If Weekday(d, vbMonday) <= workDaysPerWeek Then
            CountWorkingDaysFast = CountWorkingDaysFast + 1
        End If
    Next i

End Function

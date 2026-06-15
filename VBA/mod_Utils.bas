Attribute VB_Name = "mod_Utils"
'=====================================================
' Module: mod_Utils
'
' Purpose:
' Provide shared helper functions used across the
' planning engine (WBS, CALC, Gantt, Live simulation).
'
' Scope:
' - Generic value handling (HasValue, GetCellValue, etc.)
' - WBS hierarchy utilities
' - Predecessor manipulation and expansion
' - Common comparison and formatting helpers
'
' Rules:
' - No business logic specific to a module (Gantt / CALC)
' - Functions must be reusable and side-effect free
' - Keep implementations lightweight (used in loops)
'=====================================================

Option Explicit

' Return True only when value is usable (not Error / Empty / blank).
Public Function HasValue(ByVal v As Variant) As Boolean

    If isError(v) Then
        HasValue = False
    ElseIf IsEmpty(v) Then
        HasValue = False
    ElseIf Trim(CStr(v)) = "" Then
        HasValue = False
    Else
        HasValue = True
    End If

End Function

' Return value if usable, otherwise return Empty.
Public Function GetCellValue(ByVal v As Variant) As Variant

    If HasValue(v) Then
        GetCellValue = v
    Else
        GetCellValue = Empty
    End If

End Function

' Return numeric value if valid, otherwise return 0.
Public Function GetNumericOrZero(ByVal v As Variant) As Double

    If HasValue(v) And IsNumeric(v) Then
        GetNumericOrZero = CDbl(v)
    Else
        GetNumericOrZero = 0
    End If

End Function

' Return the maximum of two Excel date values.
Public Function maxDate(ByVal d1 As Variant, ByVal d2 As Variant) As Double

    If CDbl(d1) >= CDbl(d2) Then
        maxDate = CDbl(d1)
    Else
        maxDate = CDbl(d2)
    End If

End Function

' Merge inherited and local predecessor strings without duplicates.
Public Function MergePredStrings(ByVal inheritedPreds As String, ByVal localPreds As String) As String

    Dim seen As Object
    Dim ordered As Collection

    Set seen = CreateObject("Scripting.Dictionary")
    Set ordered = New Collection

    AddPredTokensToOrdered inheritedPreds, seen, ordered
    AddPredTokensToOrdered localPreds, seen, ordered

    MergePredStrings = JoinCollectionWithSemicolon(ordered)

End Function

' Resolve full predecessor chain for an ID (including inherited WBS parents).
Public Function ResolvePredsForId( _
    ByVal id As String, _
    ByVal idToWbs As Object, _
    ByVal wbsToId As Object, _
    ByVal rawPredById As Object, _
    ByVal resolvedPredById As Object) As String

    Dim currentWBS As String
    Dim parentWbs As String
    Dim parentId As String
    Dim inheritedPreds As String
    Dim localPreds As String

    If resolvedPredById.Exists(id) Then
        ResolvePredsForId = CStr(resolvedPredById(id))
        Exit Function
    End If

    inheritedPreds = ""
    localPreds = ""

    If idToWbs.Exists(id) Then
        currentWBS = CStr(idToWbs(id))
        parentWbs = GetParentWBS(currentWBS)

        If parentWbs <> "" Then
            If wbsToId.Exists(parentWbs) Then
                parentId = CStr(wbsToId(parentWbs))
                inheritedPreds = ResolvePredsForId(parentId, idToWbs, wbsToId, rawPredById, resolvedPredById)
            End If
        End If
    End If

    If rawPredById.Exists(id) Then
        localPreds = CStr(rawPredById(id))
    End If

    ResolvePredsForId = MergePredStrings(inheritedPreds, localPreds)
    resolvedPredById(id) = ResolvePredsForId

End Function

' Return parent WBS code from a given WBS string.
Public Function GetParentWBS(ByVal wbs As String) As String

    Dim lastDotPos As Long

    wbs = Replace(Trim$(wbs), ",", ".")
    lastDotPos = InStrRev(wbs, ".")

    If lastDotPos > 0 Then
        GetParentWBS = Left$(wbs, lastDotPos - 1)
    Else
        GetParentWBS = ""
    End If

End Function

' Add unique predecessor tokens into an ordered collection.
Public Sub AddPredTokensToOrdered(ByVal predText As String, ByVal seen As Object, ByVal ordered As Collection)

    Dim arr As Variant
    Dim i As Long
    Dim token As String

    predText = Trim(CStr(predText))
    If predText = "" Then Exit Sub

    arr = Split(predText, ";")

    For i = LBound(arr) To UBound(arr)
        token = Trim(CStr(arr(i)))

        If token <> "" Then
            If Not seen.Exists(token) Then
                seen(token) = True
                ordered.Add token
            End If
        End If
    Next i

End Sub

' Join a collection of items into a semicolon-separated string.
Public Function JoinCollectionWithSemicolon(ByVal items As Collection) As String

    Dim i As Long
    Dim result As String

    result = ""

    For i = 1 To items.Count
        If result <> "" Then result = result & ";"
        result = result & CStr(items(i))
    Next i

    JoinCollectionWithSemicolon = result

End Function

' Expand predecessor list by replacing parent IDs with all their leaf descendants.
Public Function ExpandPredsToLeafIds( _
    ByVal predText As String, _
    ByVal parentIds As Object, _
    ByVal directChildrenById As Object, _
    ByVal leafDescCache As Object) As String

    Dim seen As Object
    Dim ordered As Collection
    Dim arr As Variant
    Dim i As Long
    Dim token As String
    Dim leafs As Collection
    Dim leafId As Variant

    Set seen = CreateObject("Scripting.Dictionary")
    Set ordered = New Collection

    predText = Trim(CStr(predText))
    If predText = "" Then
        ExpandPredsToLeafIds = ""
        Exit Function
    End If

    arr = Split(predText, ";")

    For i = LBound(arr) To UBound(arr)
        token = Trim(CStr(arr(i)))

        If token <> "" Then

            If parentIds.Exists(token) Then
                Set leafs = GetLeafDescendants(token, directChildrenById, leafDescCache)

                For Each leafId In leafs
                    If Not seen.Exists(CStr(leafId)) Then
                        seen(CStr(leafId)) = True
                        ordered.Add CStr(leafId)
                    End If
                Next leafId
            Else
                If Not seen.Exists(token) Then
                    seen(token) = True
                    ordered.Add token
                End If
            End If

        End If
    Next i

    ExpandPredsToLeafIds = JoinCollectionWithSemicolon(ordered)

End Function

' Recursively retrieve all leaf descendants of a given parent ID.
Public Function GetLeafDescendants( _
    ByVal parentId As String, _
    ByVal directChildrenById As Object, _
    ByVal leafDescCache As Object) As Collection

    Dim result As Collection
    Dim childId As Variant
    Dim childLeafs As Collection
    Dim leafId As Variant

    If leafDescCache.Exists(parentId) Then
        Set GetLeafDescendants = leafDescCache(parentId)
        Exit Function
    End If

    Set result = New Collection

    If Not directChildrenById.Exists(parentId) Then
        result.Add parentId
        Set leafDescCache(parentId) = result
        Set GetLeafDescendants = result
        Exit Function
    End If

    If directChildrenById(parentId).Count = 0 Then
        result.Add parentId
        Set leafDescCache(parentId) = result
        Set GetLeafDescendants = result
        Exit Function
    End If

    For Each childId In directChildrenById(parentId)
        Set childLeafs = GetLeafDescendants(CStr(childId), directChildrenById, leafDescCache)

        For Each leafId In childLeafs
            result.Add CStr(leafId)
        Next leafId
    Next childId

    Set leafDescCache(parentId) = result
    Set GetLeafDescendants = result

End Function

' Return bilingual message with explicit FR/EN sections for console language rendering.
Public Function BiMsg(ByVal frText As String, ByVal enText As String) As String

    BiMsg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

End Function

' Normalize percent input (accept 0-1 or 0-100 scale).
Public Function NormalizePercentInput(ByVal v As Variant) As Variant

    Dim x As Double

    If Not HasValue(v) Then
        NormalizePercentInput = Empty
        Exit Function
    End If

    If Not IsNumeric(v) Then
        NormalizePercentInput = Empty
        Exit Function
    End If

    x = CDbl(v)

    If x < 0 Then
        NormalizePercentInput = Empty
    ElseIf x <= 1 Then
        NormalizePercentInput = x
    Else
        NormalizePercentInput = x / 100#
    End If

End Function

' Compare two values with tolerance for numeric values.
Public Function ValuesDiffer(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean

    If Not HasValue(v1) And Not HasValue(v2) Then Exit Function

    If HasValue(v1) Xor HasValue(v2) Then
        ValuesDiffer = True
        Exit Function
    End If

    If IsNumeric(v1) And IsNumeric(v2) Then
        ValuesDiffer = (Abs(CDbl(v1) - CDbl(v2)) > 0.0000001)
    Else
        ValuesDiffer = (CStr(v1) <> CStr(v2))
    End If

End Function

' Normalize WBS string format (trim + replace comma with dot).
Public Function NormalizeWBS(ByVal rawValue As Variant) As String

    Dim s As String

    If isError(rawValue) Then
        NormalizeWBS = ""
        Exit Function
    End If

    If IsEmpty(rawValue) Then
        NormalizeWBS = ""
        Exit Function
    End If

    s = Trim$(CStr(rawValue))

    If s = "" Then
        NormalizeWBS = ""
        Exit Function
    End If

    ' Canonical separator for WBS = dot
    s = Replace$(s, ",", ".")

    ' Remove accidental spaces
    s = Replace$(s, " ", "")

    ' Defensive cleanup for repeated dots
    Do While InStr(1, s, "..", vbBinaryCompare) > 0
        s = Replace$(s, "..", ".")
    Loop

    ' Remove leading/trailing dot if ever present
    If Left$(s, 1) = "." Then s = Mid$(s, 2)
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    NormalizeWBS = s

End Function

'=================================================
' Parse a single predecessor token.
'
' INPUT
' - tokenText examples:
'   1
'   1+3
'   1FS+3
'   1SS-2
'   1FF
'
' OUTPUT
' - predWBS  : predecessor WBS
' - linkType : FS / SS / FF
' - lagVal   : integer lag
' - rawToken : normalized raw token
'
' RETURN
' - True if token parsed successfully
' - False if token is invalid
'=================================================
Public Function ParsePredecessorToken( _
    ByVal token As String, _
    ByRef predWbs As String, _
    ByRef linkType As String, _
    ByRef lagVal As Long, _
    ByRef rawToken As String, _
    ByRef errText As String) As Boolean

    Dim t As String
    Dim posType As Long
    Dim posPlus As Long
    Dim posMinus As Long
    Dim posLag As Long
    Dim suffix As String
    Dim lagText As String

    ParsePredecessorToken = False
    predWbs = ""
    linkType = "FS"
    lagVal = 0
    rawToken = ""
    errText = ""

    t = Trim$(token)

    If ContainsForbiddenWhitespace(t) Then
        errText = "Spaces are not allowed in predecessor tokens."
        Exit Function
    End If

    t = Replace$(t, ",", ".")

    If t = "" Then
        errText = "Token vide."
        Exit Function
    End If

    '----------------------------------------
    ' 1) Detect explicit link type if present
    '----------------------------------------
    posType = 0

    If InStr(1, t, "SS", vbTextCompare) > 0 Then
        linkType = "SS"
        posType = InStr(1, t, "SS", vbTextCompare)
    ElseIf InStr(1, t, "FF", vbTextCompare) > 0 Then
        linkType = "FF"
        posType = InStr(1, t, "FF", vbTextCompare)
    ElseIf InStr(1, t, "FS", vbTextCompare) > 0 Then
        linkType = "FS"
        posType = InStr(1, t, "FS", vbTextCompare)
    End If

    '----------------------------------------
    ' 2) Split WBS / lag
    '    Accepted forms:
    '    - 1.2.3
    '    - 1.2.3+4
    '    - 1.2.3-2
    '    - 1.2.3FS
    '    - 1.2.3FS+4
    '    - 1.2.3SS-2
    '----------------------------------------
    If posType > 0 Then

        predWbs = Left$(t, posType - 1)
        suffix = Mid$(t, posType + 2)

    Else

        posPlus = InStrRev(t, "+")
        posMinus = InStrRev(t, "-")

        posLag = 0
        If posPlus > 0 And posMinus > 0 Then
            If posPlus > posMinus Then
                posLag = posPlus
            Else
                posLag = posMinus
            End If
        ElseIf posPlus > 0 Then
            posLag = posPlus
        ElseIf posMinus > 0 Then
            posLag = posMinus
        End If

        If posLag > 0 Then
            predWbs = Left$(t, posLag - 1)
            suffix = Mid$(t, posLag)
        Else
            predWbs = t
            suffix = ""
        End If

    End If

    predWbs = Trim$(predWbs)

    If predWbs = "" Then
        errText = "Prédécesseur WBS vide."
        Exit Function
    End If

    If Not IsValidPureWBS(predWbs) Then
        errText = "WBS prédécesseur invalide : " & predWbs
        Exit Function
    End If

    '----------------------------------------
    ' 3) Parse lag if present
    '----------------------------------------
    If suffix <> "" Then
        If Left$(suffix, 1) <> "+" And Left$(suffix, 1) <> "-" Then
            errText = "Suffix invalide : " & suffix
            Exit Function
        End If

        lagText = suffix

        If Not IsNumeric(lagText) Then
            errText = "Lag invalide : " & lagText
            Exit Function
        End If

        lagVal = CLng(lagText)
    End If

    rawToken = predWbs & linkType
    If lagVal > 0 Then
        rawToken = rawToken & "+" & CStr(lagVal)
    ElseIf lagVal < 0 Then
        rawToken = rawToken & CStr(lagVal)
    End If

    ParsePredecessorToken = True

End Function

Private Function ContainsForbiddenWhitespace(ByVal textValue As String) As Boolean

    ContainsForbiddenWhitespace = _
        (InStr(1, textValue, " ", vbBinaryCompare) > 0) Or _
        (InStr(1, textValue, vbTab, vbBinaryCompare) > 0) Or _
        (InStr(1, textValue, Chr$(160), vbBinaryCompare) > 0)

End Function
Private Function IsValidPureWBS(ByVal wbsText As String) As Boolean

    Dim reWBS As Object

    Set reWBS = CreateObject("VBScript.RegExp")
    reWBS.Pattern = "^\d+(\.\d+)*$"

    IsValidPureWBS = reWBS.Test(Trim$(wbsText))

End Function

'=================================================
' Parse a full Predecessors WBS cell into a collection
' of dictionaries ready for later use.
'
' EACH ITEM CONTAINS
' - Succ ID
' - Succ WBS
' - Pred ID
' - Pred WBS
' - Link Type
' - Lag
' - Raw Token
'
' RETURN
' - True if all tokens were parsed successfully
' - False if at least one token failed
'
' NOTE
' - empty cell = valid, zero link
'=================================================
Public Function ParsePredecessorsText( _
    ByVal succId As String, _
    ByVal succWBS As String, _
    ByVal predText As String, _
    ByVal wbsToId As Object, _
    ByRef linksOut As Collection, _
    ByRef errText As String) As Boolean

    Dim parts As Variant
    Dim i As Long
    Dim tokenText As String

    Dim predWbs As String
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Long
    Dim rawToken As String

    Dim linkRow As Object

    Set linksOut = New Collection
    errText = ""

    predText = Replace$(Trim$(predText), ",", ".")

    If predText = "" Then
        ParsePredecessorsText = True
        Exit Function
    End If

    parts = Split(predText, ";")

    For i = LBound(parts) To UBound(parts)

        tokenText = Trim$(CStr(parts(i)))

        If tokenText = "" Then
            errText = "Empty predecessor token in: " & predText
            Exit Function
        End If

        If Not ParsePredecessorToken(tokenText, predWbs, linkType, lagVal, rawToken, errText) Then
            errText = "Successor WBS " & succWBS & " -> " & errText
            Exit Function
        End If

        If wbsToId.Exists(predWbs) Then
            predId = CStr(wbsToId(predWbs))
        Else
            predId = ""
        End If

        Set linkRow = CreateObject("Scripting.Dictionary")
        linkRow("Succ ID") = Trim$(succId)
        linkRow("Succ WBS") = NormalizeWBS(succWBS)
        linkRow("Pred ID") = predId
        linkRow("Pred WBS") = predWbs
        linkRow("Link Type") = linkType
        linkRow("Lag") = lagVal
        linkRow("Raw Token") = rawToken
        linkRow("Entered Token") = tokenText

        linksOut.Add linkRow

    Next i

    ParsePredecessorsText = True

End Function


'=================================================
' Build WBS -> ID map from tbl_WBS.
' Dedicated helper for parser/testing.
'=================================================
Public Function BuildWbsToIdMapFromTable( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim arr As Variant
    Dim r As Long
    Dim wbsVal As String
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildWbsToIdMapFromTable = d
        Exit Function
    End If

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        wbsVal = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If wbsVal <> "" And idVal <> "" Then
            d(wbsVal) = idVal
        End If
    Next r

    Set BuildWbsToIdMapFromTable = d

End Function

Private Sub Utils_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub




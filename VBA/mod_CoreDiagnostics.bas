Attribute VB_Name = "mod_CoreDiagnostics"
Option Explicit

'===============================================================================
' Structured Core diagnostic records. Run_Calc_Core remains the sole producer
' and lifecycle owner; this module only stores, queries and serializes its facts.
'===============================================================================

Public Sub CoreDiagnostics_Record( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal diagnosticCode As String, _
    Optional ByVal diagnosticArguments As Object = Nothing, _
    Optional ByVal classification As String = "ROOT", _
    Optional ByVal relatedObjectId As String = "", _
    Optional ByVal family As String = "CORE", _
    Optional ByVal specializedPayload As Object = Nothing)

    Dim records As Collection
    Dim record As Object
    Dim existing As Variant
    Dim signature As String

    If coreDiagnostics Is Nothing Then Exit Sub
    If Trim$(taskId) = "" Or Trim$(diagnosticCode) = "" Then Exit Sub

    If coreDiagnostics.Exists(CStr(taskId)) Then
        Set records = coreDiagnostics(CStr(taskId))
    Else
        Set records = New Collection
        coreDiagnostics.Add CStr(taskId), records
    End If

    signature = CoreDiagnostics_RecordSignature(diagnosticCode, diagnosticArguments, _
        classification, relatedObjectId, family)
    For Each existing In records
        If IsObject(existing) Then
            If CStr(existing("Signature")) = signature Then Exit Sub
        End If
    Next existing

    Set record = CreateObject("Scripting.Dictionary")
    record("Code") = UCase$(Trim$(diagnosticCode))
    record("Classification") = UCase$(Trim$(classification))
    record("RelatedObjectID") = Trim$(relatedObjectId)
    record("Family") = UCase$(Trim$(family))
    record("Signature") = signature
    If Not diagnosticArguments Is Nothing Then Set record("Arguments") = diagnosticArguments
    If Not specializedPayload Is Nothing Then Set record("Payload") = specializedPayload
    records.Add record

End Sub

Public Sub CoreDiagnostics_ClearScope(ByVal coreDiagnostics As Object, ByVal recalcScope As Object)

    Dim taskId As Variant

    If coreDiagnostics Is Nothing Then Exit Sub
    If recalcScope Is Nothing Then
        coreDiagnostics.RemoveAll
        Exit Sub
    End If
    For Each taskId In recalcScope.Keys
        If coreDiagnostics.Exists(CStr(taskId)) Then coreDiagnostics.Remove CStr(taskId)
    Next taskId

End Sub

Public Sub CoreDiagnostics_MergeSpecialized( _
    ByVal coreDiagnostics As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal cascadeDiagnostics As Object)

    Dim key As Variant
    Dim diag As Object
    Dim args As Object
    Dim rootErrorId As String

    If coreDiagnostics Is Nothing Then Exit Sub

    If Not constraintDiagnostics Is Nothing Then
        For Each key In constraintDiagnostics.Keys
            If IsObject(constraintDiagnostics(CStr(key))) Then
                Set diag = constraintDiagnostics(CStr(key))
                If diag.Exists("Code") Then
                    CoreDiagnostics_Record coreDiagnostics, CStr(key), CStr(diag("Code")), _
                        Nothing, "ROOT", vbNullString, "CONSTRAINT", diag
                End If
            End If
        Next key
    End If

    If Not cascadeDiagnostics Is Nothing Then
        For Each key In cascadeDiagnostics.Keys
            If IsObject(cascadeDiagnostics(CStr(key))) Then
                Set diag = cascadeDiagnostics(CStr(key))
                Set args = CreateObject("Scripting.Dictionary")
                rootErrorId = vbNullString
                If diag.Exists("RootErrorID") Then
                    rootErrorId = CStr(diag("RootErrorID"))
                    args("RootErrorID") = rootErrorId
                End If
                If diag.Exists("ParentPropagatedFrom") Then _
                    args("ParentID") = CStr(diag("ParentPropagatedFrom"))
                CoreDiagnostics_Record coreDiagnostics, CStr(key), _
                    "CORE.ERROR.BLOCKED_PREDECESSOR_CHAIN", args, "INHERITED", _
                    rootErrorId, "CASCADE", diag
            End If
        Next key
    End If

End Sub

Public Function CoreDiagnostics_ForTask(ByVal coreDiagnostics As Object, ByVal taskId As String) As Collection

    If coreDiagnostics Is Nothing Then Exit Function
    If Not coreDiagnostics.Exists(CStr(taskId)) Then Exit Function
    Set CoreDiagnostics_ForTask = coreDiagnostics(CStr(taskId))

End Function

Public Function CoreDiagnostics_TaskHasCode( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal diagnosticCode As String) As Boolean

    Dim records As Collection
    Dim record As Variant

    Set records = CoreDiagnostics_ForTask(coreDiagnostics, taskId)
    If records Is Nothing Then Exit Function
    For Each record In records
        If IsObject(record) Then
            If StrComp(CStr(record("Code")), diagnosticCode, vbTextCompare) = 0 Then
                CoreDiagnostics_TaskHasCode = True
                Exit Function
            End If
        End If
    Next record

End Function

Public Function CoreDiagnostics_TaskHasFamily( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal family As String) As Boolean

    Dim records As Collection
    Dim record As Variant

    Set records = CoreDiagnostics_ForTask(coreDiagnostics, taskId)
    If records Is Nothing Then Exit Function
    For Each record In records
        If IsObject(record) Then
            If StrComp(CStr(record("Family")), family, vbTextCompare) = 0 Then
                CoreDiagnostics_TaskHasFamily = True
                Exit Function
            End If
        End If
    Next record

End Function

Public Function CoreDiagnostics_TaskHasClassification( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal classification As String) As Boolean

    Dim records As Collection
    Dim record As Variant

    Set records = CoreDiagnostics_ForTask(coreDiagnostics, taskId)
    If records Is Nothing Then Exit Function
    For Each record In records
        If IsObject(record) Then
            If StrComp(CStr(record("Classification")), classification, vbTextCompare) = 0 Then
                CoreDiagnostics_TaskHasClassification = True
                Exit Function
            End If
        End If
    Next record

End Function

Public Function CoreDiagnostics_TaskRelatedObjectForCode( _
    ByVal coreDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal diagnosticCode As String) As String

    Dim records As Collection
    Dim record As Variant

    Set records = CoreDiagnostics_ForTask(coreDiagnostics, taskId)
    If records Is Nothing Then Exit Function
    For Each record In records
        If IsObject(record) Then
            If StrComp(CStr(record("Code")), diagnosticCode, vbTextCompare) = 0 Then
                CoreDiagnostics_TaskRelatedObjectForCode = CStr(record("RelatedObjectID"))
                Exit Function
            End If
        End If
    Next record

End Function

Private Function CoreDiagnostics_RecordSignature( _
    ByVal diagnosticCode As String, _
    ByVal diagnosticArguments As Object, _
    ByVal classification As String, _
    ByVal relatedObjectId As String, _
    ByVal family As String) As String

    Dim keys As Variant
    Dim i As Long
    Dim key As String

    CoreDiagnostics_RecordSignature = UCase$(Trim$(diagnosticCode)) & "|" & _
        UCase$(Trim$(classification)) & "|" & Trim$(relatedObjectId) & "|" & _
        UCase$(Trim$(family))
    If diagnosticArguments Is Nothing Then Exit Function
    If diagnosticArguments.Count = 0 Then Exit Function

    keys = diagnosticArguments.Keys
    CoreDiagnostics_SortStrings keys
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        CoreDiagnostics_RecordSignature = CoreDiagnostics_RecordSignature & "|" & _
            UCase$(Trim$(key)) & "=" & CStr(diagnosticArguments(key))
    Next i

End Function

Private Sub CoreDiagnostics_SortStrings(ByRef values As Variant)

    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If StrComp(CStr(values(i)), CStr(values(j)), vbTextCompare) > 0 Then
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next j
    Next i

End Sub

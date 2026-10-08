Attribute VB_Name = "mod_TextCatalog"
Option Explicit

'===============================================================================
' MODULE : mod_TextCatalog
' DOMAINE / DOMAIN : Localization foundation
'
' Resolves symbolic text keys without owning language selection. Domain owners
' pass an explicit language key. The runtime registry is built once per session
' and performs no Excel object-model access during lookup or formatting.
'===============================================================================

Public Const TEXT_LANGUAGE_EN As String = "EN"
Public Const TEXT_LANGUAGE_FR As String = "FR"

Private Const MISSING_TEXT_PREFIX As String = "[missing text:"

Private mTextsByKey As Object
Private mDiagnostics As Collection
Private mCatalogBuilt As Boolean
Private mCatalogBuilding As Boolean
Private mBuildCount As Long
Private mLookupCount As Long

'------------------------------------------------------------------------------
' Returns the requested translation, then canonical English, then a visible
' missing-text marker. Language ownership always remains with the caller.
'------------------------------------------------------------------------------
Public Function TextCatalog_Get(ByVal messageKey As String, ByVal languageKey As String) As String

    Dim normalizedKey As String
    Dim normalizedLanguage As String
    Dim translations As Object

    TextCatalog_EnsureBuilt
    mLookupCount = mLookupCount + 1

    normalizedKey = TextCatalog_NormalizeMessageKey(messageKey)
    normalizedLanguage = TextCatalog_NormalizeLanguageKey(languageKey)

    If Not mTextsByKey.Exists(normalizedKey) Then
        TextCatalog_RecordDiagnostic "UNKNOWN_KEY", normalizedKey & "/" & normalizedLanguage
        TextCatalog_Get = TextCatalog_MissingMarker(normalizedKey, normalizedLanguage)
        Exit Function
    End If

    Set translations = mTextsByKey(normalizedKey)
    If translations.Exists(normalizedLanguage) Then
        If Len(CStr(translations(normalizedLanguage))) > 0 Then
            TextCatalog_Get = CStr(translations(normalizedLanguage))
            Exit Function
        End If
    End If

    If normalizedLanguage <> TEXT_LANGUAGE_EN Then
        TextCatalog_RecordDiagnostic "LANGUAGE_FALLBACK", normalizedKey & "/" & normalizedLanguage
    Else
        TextCatalog_RecordDiagnostic "MISSING_ENGLISH", normalizedKey
    End If

    If translations.Exists(TEXT_LANGUAGE_EN) Then
        If Len(CStr(translations(TEXT_LANGUAGE_EN))) > 0 Then
            TextCatalog_Get = CStr(translations(TEXT_LANGUAGE_EN))
            Exit Function
        End If
    End If

    TextCatalog_Get = TextCatalog_MissingMarker(normalizedKey, normalizedLanguage)

End Function

'------------------------------------------------------------------------------
' Formats a localized template with named arguments. The caller creates the
' argument dictionary through TextCatalog_Arguments or an equivalent object.
'------------------------------------------------------------------------------
Public Function TextCatalog_Format( _
    ByVal messageKey As String, _
    ByVal languageKey As String, _
    ByVal namedArguments As Object) As String

    Dim localizedText As String
    Dim placeholders As Object
    Dim placeholderName As Variant
    Dim argumentName As Variant
    Dim result As String
    Dim cursor As Long
    Dim openPosition As Long
    Dim closePosition As Long
    Dim tokenName As String
    Dim tokenText As String

    localizedText = TextCatalog_Get(messageKey, languageKey)
    Set placeholders = TextCatalog_ExtractPlaceholders(localizedText)

    For Each placeholderName In placeholders.Keys
        If Not namedArguments Is Nothing Then
            If Not namedArguments.Exists(CStr(placeholderName)) Then
                TextCatalog_RecordDiagnostic "MISSING_ARGUMENT", _
                    TextCatalog_NormalizeMessageKey(messageKey) & "/" & CStr(placeholderName)
            End If
        Else
            TextCatalog_RecordDiagnostic "MISSING_ARGUMENT", _
                TextCatalog_NormalizeMessageKey(messageKey) & "/" & CStr(placeholderName)
        End If
    Next placeholderName

    If Not namedArguments Is Nothing Then
        For Each argumentName In namedArguments.Keys
            If Not placeholders.Exists(CStr(argumentName)) Then
                TextCatalog_RecordDiagnostic "EXTRA_ARGUMENT", _
                    TextCatalog_NormalizeMessageKey(messageKey) & "/" & CStr(argumentName)
            End If
        Next argumentName
    End If

    ' Read tokens only from the original template; inserted values are opaque text.
    cursor = 1
    Do
        openPosition = InStr(cursor, localizedText, "{", vbBinaryCompare)
        If openPosition = 0 Then Exit Do
        closePosition = InStr(openPosition + 1, localizedText, "}", vbBinaryCompare)
        If closePosition = 0 Then Exit Do
        result = result & Mid$(localizedText, cursor, openPosition - cursor)
        tokenName = Mid$(localizedText, openPosition + 1, closePosition - openPosition - 1)
        tokenText = Mid$(localizedText, openPosition, closePosition - openPosition + 1)
        If Len(tokenName) > 0 And Not namedArguments Is Nothing Then
            If namedArguments.Exists(tokenName) Then tokenText = CStr(namedArguments(tokenName))
        End If
        result = result & tokenText
        cursor = closePosition + 1
    Loop
    TextCatalog_Format = result & Mid$(localizedText, cursor)

End Function

'------------------------------------------------------------------------------
' Creates a binary-compare named-argument dictionary from name/value pairs.
'------------------------------------------------------------------------------
Public Function TextCatalog_Arguments(ParamArray nameValuePairs() As Variant) As Object

    Dim arguments As Object
    Dim pairCount As Long
    Dim i As Long
    Dim argumentName As String

    Set arguments = CreateObject("Scripting.Dictionary")
    arguments.CompareMode = vbBinaryCompare

    pairCount = TextCatalog_ParamArrayLength(nameValuePairs)
    If pairCount Mod 2 <> 0 Then
        Err.Raise vbObjectError + 5401, "TextCatalog_Arguments", _
            TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.ARGUMENT_PAIRS")
    End If

    For i = 0 To pairCount - 1 Step 2
        argumentName = CStr(nameValuePairs(i))
        If Len(argumentName) = 0 Then
            Err.Raise vbObjectError + 5402, "TextCatalog_Arguments", _
                TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.EMPTY_ARGUMENT_NAME")
        End If
        If arguments.Exists(argumentName) Then
            Err.Raise vbObjectError + 5403, "TextCatalog_Arguments", _
                TextCatalogBootstrap_ErrorWire( _
                    "TEXTCATALOG.ERROR.DUPLICATE_ARGUMENT", argumentName)
        End If
        arguments.Add argumentName, nameValuePairs(i + 1)
    Next i

    Set TextCatalog_Arguments = arguments

End Function

'------------------------------------------------------------------------------
' Validates uniqueness, completeness and placeholder parity for FR and EN.
'------------------------------------------------------------------------------
Public Function TextCatalog_ValidateAll() As String

    Dim validationErrors As Collection
    Dim messageKey As Variant
    Dim translations As Object
    Dim englishSignature As String
    Dim frenchSignature As String

    TextCatalog_EnsureBuilt
    Set validationErrors = New Collection

    For Each messageKey In mTextsByKey.Keys
        Set translations = mTextsByKey(CStr(messageKey))

        If Not translations.Exists(TEXT_LANGUAGE_EN) Then
            validationErrors.Add CStr(messageKey) & ": missing EN"
        ElseIf Len(CStr(translations(TEXT_LANGUAGE_EN))) = 0 Then
            validationErrors.Add CStr(messageKey) & ": empty EN"
        End If

        If Not translations.Exists(TEXT_LANGUAGE_FR) Then
            validationErrors.Add CStr(messageKey) & ": missing FR"
        ElseIf Len(CStr(translations(TEXT_LANGUAGE_FR))) = 0 Then
            validationErrors.Add CStr(messageKey) & ": empty FR"
        End If

        If translations.Exists(TEXT_LANGUAGE_EN) And translations.Exists(TEXT_LANGUAGE_FR) Then
            englishSignature = TextCatalog_PlaceholderSignature(CStr(translations(TEXT_LANGUAGE_EN)))
            frenchSignature = TextCatalog_PlaceholderSignature(CStr(translations(TEXT_LANGUAGE_FR)))
            If englishSignature <> frenchSignature Then
                validationErrors.Add CStr(messageKey) & ": placeholder mismatch EN=" & _
                    englishSignature & " FR=" & frenchSignature
            End If
        End If
    Next messageKey

    If validationErrors.Count = 0 Then
        TextCatalog_ValidateAll = "PASS"
    Else
        TextCatalog_ValidateAll = TextCatalog_JoinCollection(validationErrors, vbCrLf)
    End If

End Function

Public Function TextCatalog_BuildCount() As Long
    TextCatalog_EnsureBuilt
    TextCatalog_BuildCount = mBuildCount
End Function

Public Function TextCatalog_LookupCount() As Long
    TextCatalog_LookupCount = mLookupCount
End Function

Public Function TextCatalog_DiagnosticCount() As Long
    TextCatalog_EnsureDiagnostics
    TextCatalog_DiagnosticCount = mDiagnostics.Count
End Function

Public Function TextCatalog_Diagnostic(ByVal index As Long) As String
    TextCatalog_EnsureDiagnostics
    If index < 1 Or index > mDiagnostics.Count Then
        Err.Raise vbObjectError + 5404, "TextCatalog_Diagnostic", _
            TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.DIAGNOSTIC_INDEX")
    End If
    TextCatalog_Diagnostic = CStr(mDiagnostics(index))
End Function

' Release-audit inventory. Equal translations are not automatically defects:
' technical tokens and canonical business values may legitimately be identical.
Public Function TextCatalog_IdenticalTranslationKeys() As String

    Dim messageKey As Variant
    Dim translations As Object
    Dim result As String

    TextCatalog_EnsureBuilt
    For Each messageKey In mTextsByKey.Keys
        Set translations = mTextsByKey(CStr(messageKey))
        If translations.Exists(TEXT_LANGUAGE_EN) And _
           translations.Exists(TEXT_LANGUAGE_FR) Then
            If CStr(translations(TEXT_LANGUAGE_EN)) = CStr(translations(TEXT_LANGUAGE_FR)) Then
                If Len(result) > 0 Then result = result & vbCrLf
                result = result & CStr(messageKey)
            End If
        End If
    Next messageKey

    TextCatalog_IdenticalTranslationKeys = result
End Function

Public Function TextCatalog_IdenticalTranslationInventory() As String

    Dim messageKey As Variant
    Dim translations As Object
    Dim result As String

    TextCatalog_EnsureBuilt
    For Each messageKey In mTextsByKey.Keys
        Set translations = mTextsByKey(CStr(messageKey))
        If translations.Exists(TEXT_LANGUAGE_EN) And _
           translations.Exists(TEXT_LANGUAGE_FR) Then
            If CStr(translations(TEXT_LANGUAGE_EN)) = CStr(translations(TEXT_LANGUAGE_FR)) Then
                If Len(result) > 0 Then result = result & vbCrLf
                result = result & CStr(messageKey) & vbTab & _
                    Replace$(CStr(translations(TEXT_LANGUAGE_EN)), vbCrLf, "\\n")
            End If
        End If
    Next messageKey

    TextCatalog_IdenticalTranslationInventory = result
End Function

' Harness-only reset. Product language changes never call this procedure.
Public Sub TextCatalog_ResetForTests()
    Set mTextsByKey = Nothing
    Set mDiagnostics = Nothing
    mCatalogBuilt = False
    mCatalogBuilding = False
    mBuildCount = 0
    mLookupCount = 0
End Sub

Private Sub TextCatalog_EnsureBuilt()

    Dim candidateTexts As Object

    If mCatalogBuilt Then Exit Sub
    If mCatalogBuilding Then
        Err.Raise vbObjectError + 5405, "TextCatalog_EnsureBuilt", _
            TextCatalogBootstrap_ErrorWire("TEXTCATALOG.ERROR.RECURSIVE_BUILD")
    End If

    On Error GoTo BuildFailed
    mCatalogBuilding = True

    Set candidateTexts = CreateObject("Scripting.Dictionary")
    candidateTexts.CompareMode = vbBinaryCompare

    TextCatalog_ImportDefinitions candidateTexts, _
        "BOOTSTRAP", TextCatalogBootstrap_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "COMMON", TextCatalogCommon_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "NAV", TextCatalogNavigation_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "DIAG", TextCatalogDiagnostics_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "VISIBLE_HEADER", VisibleTableSchema_TextCatalogDefinitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "WBS", TextCatalogWBS_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "GANTT", TextCatalogGantt_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "SCURVE", TextCatalogSCurve_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "SETTINGS", TextCatalogSettings_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "CONSTRAINTS", TextCatalogConstraints_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "DASHBOARD", TextCatalogDashboard_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "RIBBON", TextCatalogRibbon_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "MIGRATION", TextCatalogMigration_Definitions()
    TextCatalog_ImportDefinitions candidateTexts, _
        "WELCOME", TextCatalogWelcome_Definitions()

    Set mTextsByKey = candidateTexts
    mCatalogBuilt = True
    mCatalogBuilding = False
    mBuildCount = mBuildCount + 1
    Exit Sub

BuildFailed:
    mCatalogBuilding = False
    Set mTextsByKey = Nothing
    Err.Raise Err.Number, "TextCatalog_EnsureBuilt", Err.Description

End Sub

Private Sub TextCatalog_ImportDefinitions( _
    ByVal targetTexts As Object, _
    ByVal providerName As String, _
    ByVal definitions As Variant)

    Dim definition As Variant
    Dim messageKey As String
    Dim translations As Object

    For Each definition In definitions
        If Not IsArray(definition) Then
            Err.Raise vbObjectError + 5406, "TextCatalog_ImportDefinitions", _
                TextCatalogBootstrap_ErrorWire( _
                    "TEXTCATALOG.ERROR.INVALID_DEFINITION", providerName)
        End If

        messageKey = TextCatalog_NormalizeMessageKey(CStr(definition(0)))
        If targetTexts.Exists(messageKey) Then
            Err.Raise vbObjectError + 5407, "TextCatalog_ImportDefinitions", _
                TextCatalogBootstrap_ErrorWire( _
                    "TEXTCATALOG.ERROR.DUPLICATE_KEY", messageKey)
        End If

        Set translations = CreateObject("Scripting.Dictionary")
        translations.CompareMode = vbBinaryCompare
        translations.Add TEXT_LANGUAGE_EN, CStr(definition(1))
        translations.Add TEXT_LANGUAGE_FR, CStr(definition(2))
        targetTexts.Add messageKey, translations
    Next definition

End Sub

Private Function TextCatalog_NormalizeMessageKey(ByVal messageKey As String) As String
    TextCatalog_NormalizeMessageKey = UCase$(Trim$(messageKey))
End Function

Private Function TextCatalog_NormalizeLanguageKey(ByVal languageKey As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(languageKey))
    If Len(normalized) = 0 Then normalized = TEXT_LANGUAGE_EN
    TextCatalog_NormalizeLanguageKey = normalized
End Function

Private Function TextCatalog_MissingMarker(ByVal messageKey As String, ByVal languageKey As String) As String
    TextCatalog_MissingMarker = MISSING_TEXT_PREFIX & messageKey & "/" & languageKey & "]"
End Function

Private Function TextCatalog_ExtractPlaceholders(ByVal value As String) As Object

    Dim result As Object
    Dim openPosition As Long
    Dim closePosition As Long
    Dim placeholderName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare
    openPosition = 1

    Do
        openPosition = InStr(openPosition, value, "{", vbBinaryCompare)
        If openPosition = 0 Then Exit Do
        closePosition = InStr(openPosition + 1, value, "}", vbBinaryCompare)
        If closePosition = 0 Then Exit Do

        placeholderName = Mid$(value, openPosition + 1, closePosition - openPosition - 1)
        If Len(placeholderName) > 0 Then
            If Not result.Exists(placeholderName) Then result.Add placeholderName, True
        End If
        openPosition = closePosition + 1
    Loop

    Set TextCatalog_ExtractPlaceholders = result

End Function

Private Function TextCatalog_PlaceholderSignature(ByVal value As String) As String

    Dim placeholders As Object
    Dim names As Variant
    Dim i As Long
    Dim j As Long
    Dim swapValue As String

    Set placeholders = TextCatalog_ExtractPlaceholders(value)
    If placeholders.Count = 0 Then Exit Function

    names = placeholders.Keys
    For i = LBound(names) To UBound(names) - 1
        For j = i + 1 To UBound(names)
            If StrComp(CStr(names(i)), CStr(names(j)), vbBinaryCompare) > 0 Then
                swapValue = CStr(names(i))
                names(i) = names(j)
                names(j) = swapValue
            End If
        Next j
    Next i
    TextCatalog_PlaceholderSignature = Join(names, "|")

End Function

Private Sub TextCatalog_RecordDiagnostic(ByVal diagnosticCode As String, ByVal detail As String)
    TextCatalog_EnsureDiagnostics
    mDiagnostics.Add diagnosticCode & ":" & detail
End Sub

Private Sub TextCatalog_EnsureDiagnostics()
    If mDiagnostics Is Nothing Then Set mDiagnostics = New Collection
End Sub

Private Function TextCatalog_ParamArrayLength(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    TextCatalog_ParamArrayLength = UBound(values) - LBound(values) + 1
    Exit Function
EmptyArray:
    TextCatalog_ParamArrayLength = 0
End Function

Private Function TextCatalog_JoinCollection(ByVal values As Collection, ByVal separator As String) As String

    Dim parts() As String
    Dim i As Long

    ReDim parts(0 To values.Count - 1)
    For i = 1 To values.Count
        parts(i - 1) = CStr(values(i))
    Next i
    TextCatalog_JoinCollection = Join(parts, separator)

End Function

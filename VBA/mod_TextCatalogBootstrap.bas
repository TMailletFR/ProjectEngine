Attribute VB_Name = "mod_TextCatalogBootstrap"
Option Explicit

' Bootstrap diagnostics belong to the common catalog, but must also be
' resolvable while that registry is still being built.
Public Function TextCatalogBootstrap_Definitions() As Variant
    Dim definitions(0 To 6) As Variant

    definitions(0) = Array("TEXTCATALOG.ERROR.ARGUMENT_PAIRS", _
        "Named arguments must be supplied as name/value pairs.", _
        "Les arguments nommés doivent être fournis par paires nom/valeur.")
    definitions(1) = Array("TEXTCATALOG.ERROR.EMPTY_ARGUMENT_NAME", _
        "A named argument cannot have an empty name.", _
        "Un argument nommé ne peut pas avoir un nom vide.")
    definitions(2) = Array("TEXTCATALOG.ERROR.DUPLICATE_ARGUMENT", _
        "Duplicate named argument: {Value}", _
        "Argument nommé en double : {Value}")
    definitions(3) = Array("TEXTCATALOG.ERROR.DIAGNOSTIC_INDEX", _
        "Diagnostic index is out of range.", _
        "L'index de diagnostic est hors limites.")
    definitions(4) = Array("TEXTCATALOG.ERROR.RECURSIVE_BUILD", _
        "Recursive catalog build detected.", _
        "Construction récursive du catalogue détectée.")
    definitions(5) = Array("TEXTCATALOG.ERROR.INVALID_DEFINITION", _
        "Invalid definition in provider {Value}", _
        "Définition invalide dans le provider {Value}")
    definitions(6) = Array("TEXTCATALOG.ERROR.DUPLICATE_KEY", _
        "Duplicate message key: {Value}", _
        "Clé de message en double : {Value}")

    TextCatalogBootstrap_Definitions = definitions
End Function

' Does not call TextCatalog_Get/Format: it is safe before registry creation.
Public Function TextCatalogBootstrap_ErrorWire( _
    ByVal messageKey As String, _
    Optional ByVal replacementValue As String = "") As String

    Dim definition As Variant
    Dim englishText As String
    Dim frenchText As String

    For Each definition In TextCatalogBootstrap_Definitions()
        If CStr(definition(0)) = messageKey Then
            englishText = Replace$(CStr(definition(1)), "{Value}", replacementValue)
            frenchText = Replace$(CStr(definition(2)), "{Value}", replacementValue)
            TextCatalogBootstrap_ErrorWire = BiMsg(frenchText, englishText)
            Exit Function
        End If
    Next definition

    TextCatalogBootstrap_ErrorWire = "[missing text:" & messageKey & "]"
End Function

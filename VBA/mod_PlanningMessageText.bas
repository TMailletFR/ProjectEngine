Attribute VB_Name = "mod_PlanningMessageText"
Option Explicit

' Adapter to existing bilingual event transport, not another catalog or owner.
' Both languages remain materialized because event hashes include both texts.
Public Function PlanningMessageText_Format( _
    ByVal messageKey As String, _
    Optional ByVal frenchArguments As Variant, _
    Optional ByVal englishArguments As Variant) As String

    Dim frenchNamedArguments As Object
    Dim englishNamedArguments As Object

    If IsObject(frenchArguments) Then Set frenchNamedArguments = frenchArguments
    If IsObject(englishArguments) Then Set englishNamedArguments = englishArguments

    PlanningMessageText_Format = BiMsg( _
        TextCatalog_Format(messageKey, TEXT_LANGUAGE_FR, frenchNamedArguments), _
        TextCatalog_Format(messageKey, TEXT_LANGUAGE_EN, englishNamedArguments))
End Function

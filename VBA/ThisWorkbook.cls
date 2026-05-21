Option Explicit

Dim gAppEvents As clsAppEvents

Private Sub Workbook_Open()
    Init_AppEvents
    Ensure_WBS_Main_Buttons
End Sub

Private Sub Workbook_Activate()
    Init_AppEvents
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Init_AppEvents
End Sub

Public Sub Init_AppEvents()

    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If

    If gAppEvents.App Is Nothing Then
        Set gAppEvents.App = Application
    End If

End Sub




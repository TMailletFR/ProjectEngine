Attribute VB_Name = "mod_MacroGuard"
'=================================================
' mod_MacroGuard
' Central macro abort / event guard
'=================================================
Option Explicit

Public gMacroRunActive As Boolean
Public gMacroAbortRequested As Boolean
Public gMacroAbortSource As String
Public gMacroAbortMessageFR As String
Public gMacroAbortMessageEN As String
Public gMacroAbortPopupShown As Boolean

Public Sub BeginMacroRun(Optional ByVal sourceName As String = "")
    gMacroRunActive = True
    gMacroAbortRequested = False
    gMacroAbortSource = sourceName
    gMacroAbortMessageFR = ""
    gMacroAbortMessageEN = ""
    gMacroAbortPopupShown = False
    BeginPlanningWorkflow sourceName
End Sub

Public Sub EndMacroRun()
    EndPlanningWorkflow
    gMacroRunActive = False
    gMacroAbortRequested = False
    gMacroAbortSource = ""
    gMacroAbortMessageFR = ""
    gMacroAbortMessageEN = ""
    gMacroAbortPopupShown = False
    EndPlanningEventRun
End Sub

Public Sub RequestMacroAbort( _
    ByVal sourceName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    gMacroAbortRequested = True
    gMacroAbortSource = sourceName
    gMacroAbortMessageFR = frText
    gMacroAbortMessageEN = enText

    On Error Resume Next
    Application.EnableEvents = False
    On Error GoTo 0
End Sub

Public Function IsMacroRunActive() As Boolean
    IsMacroRunActive = gMacroRunActive
End Function

Public Function IsMacroAbortRequested() As Boolean
    IsMacroAbortRequested = gMacroAbortRequested
End Function

Public Sub AbortIfRequested(Optional ByVal sourceName As String = "")

    If Not gMacroAbortRequested Then Exit Sub

    If Trim$(sourceName) <> "" Then
        Err.Raise vbObjectError + 2901, sourceName, _
            gMacroAbortMessageFR & vbCrLf & gMacroAbortMessageEN
    Else
        Err.Raise vbObjectError + 2901, "MacroAbort", _
            gMacroAbortMessageFR & vbCrLf & gMacroAbortMessageEN
    End If

End Sub

Public Sub ShowAbortMessageOnce()

    Dim consoleMessages As Collection

    If Not gMacroAbortRequested Then Exit Sub
    If gMacroAbortPopupShown Then Exit Sub

    gMacroAbortPopupShown = True

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        "FR:" & vbCrLf & _
        gMacroAbortMessageFR & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        gMacroAbortMessageEN

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub




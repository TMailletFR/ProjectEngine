Attribute VB_Name = "mod_WorksheetEvents"
Option Explicit

' Central dispatcher for native Excel worksheet edits.
' Runtime/macro messages stay in the CalcBridge/EventHistory pipeline.
Public Sub Handle_Worksheet_Change(ByVal Sh As Object, ByVal Target As Range)

    If Sh Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Select Case CStr(Sh.Name)
        Case "WBS"
            Handle_WBS_Change Sh, Target
        Case "GANTT"
            Handle_Gantt_Change Sh, Target
        Case "CONSTRAINTS"
            Handle_Constraints_Change Sh, Target
        Case "CALC_ALARM", "EVENT_HISTORY", "EVENT_ACK"
            Handle_EventHistory_Change Sh, Target
    End Select

End Sub

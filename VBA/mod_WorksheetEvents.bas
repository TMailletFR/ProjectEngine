Attribute VB_Name = "mod_WorksheetEvents"
Option Explicit

'===============================================================================
' MODULE : mod_WorksheetEvents
' DOMAINE / DOMAIN : Workbook / Sheet Events
'
' FR
' Dispatch les changements de feuille Excel vers WBS, Constraints ou Gantt.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Dispatches Excel sheet changes to WBS, Constraints or Gantt.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Handle_Worksheet_Change
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Central dispatcher for native Excel worksheet edits.
' Runtime/macro messages stay in the CalcBridge/EventHistory pipeline.
'------------------------------------------------------------------------------
' FR: Dispatche un changement de feuille vers le domaine fonctionnel concerne.
' EN: Dispatches a sheet change to the matching functional domain.
'------------------------------------------------------------------------------
Public Sub Handle_Worksheet_Change(ByVal Sh As Object, ByVal Target As Range)

    If Sh Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Not (Sh.Parent Is ThisWorkbook) Then Exit Sub

    Select Case CStr(Sh.Name)
        Case "WBS"
            WorkflowCommand_InputsChanged
            Handle_WBS_Change Sh, Target
            ProjectWelcome_ObserveTutorialEdit Sh, Target
        Case "GANTT"
            Handle_Gantt_Change Sh, Target
        Case "DASHBOARD"
            Dashboard_HandleSnapshotSelectionChange Sh, Target
        Case "CONSTRAINTS"
            WorkflowCommand_InputsChanged
            Handle_Constraints_Change Sh, Target
        Case "CALC_ALARM", "EVENT_HISTORY", "EVENT_ACK"
            Handle_EventHistory_Change Sh, Target
    End Select

End Sub

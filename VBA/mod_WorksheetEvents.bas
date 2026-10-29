Attribute VB_Name = "mod_WorksheetEvents"
Option Explicit

'===============================================================================
' MODULE : mod_WorksheetEvents
' DOMAINE / DOMAIN : Workbook / Sheet Events
'
' FR
' Dispatch les changements de feuille Excel vers leur domaine.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Dispatches Excel sheet changes to their functional domain.
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
Public Sub Handle_Worksheet_Change(ByVal Sh As Object, ByVal target As Range)

    If Sh Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If Not (Sh.parent Is ThisWorkbook) Then Exit Sub
    If Finance_ExpenseSheetChanged(Sh, target) Then Exit Sub

    Select Case CStr(Sh.name)
        Case "WBS"
            WorkflowCommand_InputsChanged
            Handle_WBS_Change Sh, target
            ProjectWelcome_ObserveTutorialEdit Sh, target
        Case "GANTT"
            Handle_Gantt_Change Sh, target
        Case "DASHBOARD"
            Dashboard_HandleSnapshotSelectionChange Sh, target
        Case "CONSTRAINTS"
            WorkflowCommand_InputsChanged
            Handle_Constraints_Change Sh, target
        Case "FINANCE", "FINANCE SUMMARY"
            Handle_Finance_Change Sh, target
        Case "CALC_ALARM", "EVENT_HISTORY", "EVENT_ACK"
            Handle_EventHistory_Change Sh, target
    End Select

End Sub



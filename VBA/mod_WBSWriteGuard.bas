Attribute VB_Name = "mod_WBSWriteGuard"
'=================================================
' mod_WBSWriteGuard
' Authorized engine writes into WBS calculated columns
'=================================================
Option Explicit

Private gWBSWriteAuthorized As Boolean
Private gWBSWriteSource As String
Private gWBSAllowedColumns As Object
Private gWBSWriteSourceStack As Collection
Private gWBSAllowedColumnsStack As Collection

Public Sub BeginAuthorizedWBSWrite( _
    ByVal sourceName As String, _
    ByVal allowedColumnNames As Variant)

    Dim i As Long

    EnsureWBSWriteGuardStacks

    If gWBSWriteAuthorized Then
        gWBSWriteSourceStack.Add gWBSWriteSource
        gWBSAllowedColumnsStack.Add gWBSAllowedColumns
    End If

    gWBSWriteAuthorized = True
    gWBSWriteSource = sourceName
    Set gWBSAllowedColumns = CreateObject("Scripting.Dictionary")

    For i = LBound(allowedColumnNames) To UBound(allowedColumnNames)
        gWBSAllowedColumns(CStr(allowedColumnNames(i))) = True
    Next i

End Sub

Public Sub EndAuthorizedWBSWrite()

    EnsureWBSWriteGuardStacks

    If gWBSWriteAuthorized Then
        If gWBSWriteSourceStack.Count > 0 Then
            gWBSWriteSource = CStr(gWBSWriteSourceStack(gWBSWriteSourceStack.Count))
            Set gWBSAllowedColumns = gWBSAllowedColumnsStack(gWBSAllowedColumnsStack.Count)
            gWBSWriteSourceStack.Remove gWBSWriteSourceStack.Count
            gWBSAllowedColumnsStack.Remove gWBSAllowedColumnsStack.Count
        Else
            gWBSWriteAuthorized = False
            gWBSWriteSource = ""
            Set gWBSAllowedColumns = Nothing
        End If
    End If

End Sub

Public Function IsAuthorizedWBSWriteActive() As Boolean
    IsAuthorizedWBSWriteActive = gWBSWriteAuthorized
End Function

Public Function GetAuthorizedWBSWriteSource() As String
    GetAuthorizedWBSWriteSource = gWBSWriteSource
End Function

Public Function IsAuthorizedWBSColumn(ByVal columnName As String) As Boolean

    If Not gWBSWriteAuthorized Then Exit Function
    If gWBSAllowedColumns Is Nothing Then Exit Function

    IsAuthorizedWBSColumn = gWBSAllowedColumns.Exists(CStr(columnName))

End Function

Private Sub EnsureWBSWriteGuardStacks()

    If gWBSWriteSourceStack Is Nothing Then Set gWBSWriteSourceStack = New Collection
    If gWBSAllowedColumnsStack Is Nothing Then Set gWBSAllowedColumnsStack = New Collection

End Sub


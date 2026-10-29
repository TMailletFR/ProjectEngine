VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PEFinanceRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Fixed Finance record, avoiding COM object activation for each ledger row.
Private mValues(0 To 39) As Variant
Private mDefined(0 To 39) As Boolean

Public Sub Initialize(ByVal keys As Variant)
    Dim key As Variant
    For Each key In keys: mDefined(Index(CStr(key))) = True: Next key
End Sub

Private Function Index(ByVal key As String) As Long
    Select Case key
        Case "WBS": Index = 0
        Case "TASK_NAME": Index = 1
        Case "STATE": Index = 2
        Case "PROGRESS": Index = 3
        Case "BUDGET": Index = 4
        Case "PROFILE": Index = 5
        Case "ACTUAL": Index = 6
        Case "ETC": Index = 7
        Case "EAC": Index = 8
        Case "VARIANCE": Index = 9
        Case "START": Index = 10
        Case "FINISH": Index = 11
        Case "PROJECT": Index = 12
        Case "TASK_ID": Index = 13
        Case "TASK_TYPE": Index = 14
        Case "DIAGNOSTIC": Index = 15
        Case "DATE": Index = 16
        Case "AMOUNT": Index = 17
        Case "COMMENT": Index = 18
        Case "PARENT": Index = 19
        Case "PROJECT_ID": Index = 20
        Case "NAME": Index = 21
        Case "DISPLAY_ORDER": Index = 22
        Case "LINK_ID": Index = 23
        Case "FINANCIAL_ID": Index = 24
        Case "ROW_ORDER": Index = 25
        Case "COPY_FROM": Index = 26
        Case "EXPENSE_ID": Index = 27
        Case "SCOPE_ID": Index = 28
        Case "CURRENCY": Index = 29
        Case "REFRESH_AT": Index = 30
        Case "REVISION": Index = 31
        Case "NEXT_PROJECT": Index = 32
        Case "NEXT_LINK": Index = 33
        Case "NEXT_FINANCIAL": Index = 34
        Case "NEXT_EXPENSE": Index = 35
        Case "OTHERS": Index = 36
        Case "RELATION": Index = 37
        Case "CONTRIBUTION": Index = 38
        Case "PROJECT_NAME": Index = 39
        Case Else: Err.Raise 5, "PEFinanceRow", "Unknown Finance field: " & key
    End Select
End Function

Public Property Get item(ByVal key As String) As Variant
Attribute item.VB_UserMemId = 0
    Dim i As Long: i = Index(key)
    If Not mDefined(i) Then Err.Raise 5, "PEFinanceRow", "Undefined Finance field: " & key
    item = mValues(i)
End Property

Public Property Let item(ByVal key As String, ByVal value As Variant)
    Dim i As Long: i = Index(key)
    If Not mDefined(i) Then Err.Raise 5, "PEFinanceRow", "Undefined Finance field: " & key
    mValues(i) = value
End Property

Public Sub Add(ByVal key As String, ByVal value As Variant)
    Dim i As Long: i = Index(key)
    If mDefined(i) Then Err.Raise 5, "PEFinanceRow", "Duplicate Finance field: " & key
    mDefined(i) = True: mValues(i) = value
End Sub

Public Function exists(ByVal key As String) As Boolean
    exists = mDefined(Index(key))
End Function

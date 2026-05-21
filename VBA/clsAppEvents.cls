VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
    
    Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
    
    Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
        If Sh Is Nothing Then Exit Sub
        If Target Is Nothing Then Exit Sub
        If Not (Sh.Parent Is ThisWorkbook) Then Exit Sub

        Handle_Worksheet_Change Sh, Target
    
    End Sub
    
    


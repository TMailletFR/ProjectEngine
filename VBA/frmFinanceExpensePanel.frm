VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinanceExpensePanel 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFinanceExpensePanel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinanceExpensePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPE As PEConsolePresentation
Private WithEvents mDefault As MSForms.CommandButton
Attribute mDefault.VB_VarHelpID = -1
Private WithEvents mCancel As MSForms.CommandButton
Attribute mCancel.VB_VarHelpID = -1

Public Sub Prepare(ByVal projectName As String, ByVal sheetName As String)
    FinanceUI_Init Me, Finance_Text("EXPENSES") & " | " & projectName, 400, 160
    FinanceUI_Label Me, "message", Replace$(Finance_Text("NATIVE_PANEL"), "{Sheet}", sheetName), 18, 14, 360, 66
    FinanceUI_Button Me, "confirm", Finance_Text("CONFIRM"), 48, 96, 145
    FinanceUI_Button Me, "cancel", Finance_Text("CANCEL"), 205, 96, 165, True
    Set mPE = FinanceUI_Buttons(Me, Array("confirm", "cancel"), mDefault, mCancel)
End Sub
Public Sub PEInvoke(ByVal name As String)
    On Error GoTo Failed
    Finance_ExpenseSheetAction name
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub mDefault_Click()
    mPE.InvokeDefault
End Sub
Private Sub mCancel_Click()
    mPE.InvokeCancel
End Sub
Public Sub PEHover(ByVal name As String)
    If Not mPE Is Nothing Then mPE.Hover name
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PEHover ""
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo Failed
    If CloseMode = 0 Then
        Cancel = True: Finance_ExpenseSheetCancel False
    Else
        If Not mPE Is Nothing Then mPE.Detach
        Set mPE = Nothing
    End If
    Exit Sub
Failed: FinanceUI_Error
End Sub
Private Sub UserForm_Terminate()
    Set mPE = Nothing
End Sub


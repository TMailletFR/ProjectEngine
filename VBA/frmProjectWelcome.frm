VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectWelcome
   Caption         =   ""
   ClientHeight    =   7800
   ClientLeft      =   930
   ClientTop       =   3705
   ClientWidth     =   11400
   OleObjectBlob   =   "frmProjectWelcome.frx":0000
End
Attribute VB_Name = "frmProjectWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSelectedAction As String

Public Property Get SelectedAction() As String
    SelectedAction = mSelectedAction
End Property

Public Sub LoadWelcome(ByVal title As String, ByVal message As String, _
                       ByVal importLabel As String, ByVal newLabel As String, ByVal closeLabel As String)
    mSelectedAction = ""
    Me.Caption = title
    Me.Width = 640
    Me.Height = 280
    Me.BackColor = RGB(250, 250, 250)
    Me.StartUpPosition = 1
    lblTitle.Caption = title
    lblTitle.Left = 24: lblTitle.Top = 18
    lblTitle.Width = 580: lblTitle.Height = 24
    lblTitle.Font.Name = "Segoe UI": lblTitle.Font.Size = 12: lblTitle.Font.Bold = True
    lblTitle.BackStyle = fmBackStyleTransparent
    lblCounter.Visible = False
    lblMessageType.Visible = False
    chkWarningAck.Visible = False
    cmdClearAck.Visible = False
    cmdClearHistory.Visible = False
    txtMessage.Text = message
    txtMessage.Left = 24: txtMessage.Top = 60
    txtMessage.Width = 580: txtMessage.Height = 105
    txtMessage.Multiline = True: txtMessage.WordWrap = True: txtMessage.Locked = True
    txtMessage.Font.Name = "Segoe UI": txtMessage.Font.Size = 11
    txtMessage.BackColor = RGB(255, 255, 255)
    txtMessage.SpecialEffect = fmSpecialEffectFlat
    txtMessage.ScrollBars = fmScrollBarsNone
    ConfigureButton cmdPrevious, importLabel, 24, 210
    ConfigureButton cmdNext, newLabel, 246, 220
    ConfigureButton cmdClose, closeLabel, 478, 126
    cmdNext.Default = True
    cmdClose.Cancel = True
End Sub

Private Sub ConfigureButton(ByVal button As Object, ByVal caption As String, ByVal left As Single, ByVal width As Single)
    button.Caption = caption
    button.Left = left: button.Top = 192
    button.Width = width: button.Height = 32
    button.Font.Name = "Segoe UI": button.Font.Size = 10
    button.Enabled = True
End Sub

Private Sub cmdPrevious_Click()
    mSelectedAction = "IMPORT"
    Me.Hide
End Sub

Private Sub cmdNext_Click()
    mSelectedAction = "NEW"
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerfScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mToken As Long

Public Sub Init(ByVal token As Long)
    mToken = token
End Sub

Private Sub Class_Terminate()
    If mToken <> 0 Then Profiler_EndScope mToken
    mToken = 0
End Sub

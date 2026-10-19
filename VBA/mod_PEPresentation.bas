Attribute VB_Name = "mod_PEPresentation"
Option Explicit
Private mClassic As Boolean, mInjectFailure As Boolean, mReason As String
Public Sub PEPresentation_Configure(ByVal Classic As Boolean, Optional ByVal injectFailure As Boolean = False)
    mClassic = Classic: mInjectFailure = injectFailure: mReason = ""
End Sub
Public Function PEPresentation_ClassicRequested() As Boolean
    PEPresentation_ClassicRequested = mClassic
End Function
Public Sub PEPresentation_CheckInitialization()
    If mInjectFailure Then Err.Raise vbObjectError + 626, , "Laboratory recoverable initialization failure"
End Sub
Public Sub PEPresentation_RecordFallback(ByVal reason As String)
    mReason = reason
End Sub
Public Function PEPresentation_LastFallback() As String
    PEPresentation_LastFallback = mReason
End Function
Public Sub PEStyle_Modern()
    PEPresentation_Configure False
End Sub
Public Sub PEStyle_Classic()
    PEPresentation_Configure True
End Sub
Public Sub PEStyle_Failure()
    PEPresentation_Configure False, True
End Sub

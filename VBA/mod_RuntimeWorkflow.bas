Attribute VB_Name = "mod_RuntimeWorkflow"
Option Explicit

'=================================================
' mod_RuntimeWorkflow
' Minimal parent/child runtime workflow context.
'
' Current responsibilities:
' - in-memory context
' - display ownership observation
' - targeted deferred display for Lock/recalc orchestration
' - no global root-only display enforcement
'=================================================

Private gPlanningWorkflowStack As Collection
Private gPlanningWorkflowSequence As Long
Private gPlanningWorkflowDisplayOwnerId As String
Private gPlanningWorkflowDeferredDisplayMessages As Collection
Private gPlanningWorkflowStopOnlyDisplayDepth As Long

Public Function BeginPlanningWorkflow(Optional ByVal sourceProcedure As String = "") As String

    Dim ctx As Object
    Dim parentCtx As Object
    Dim parentId As String
    Dim rootId As String
    Dim depth As Long

    EnsurePlanningWorkflowStack
    If Not ValidatePlanningWorkflowStack() Then
        ResetPlanningWorkflowContextSafe "BeginPlanningWorkflow.InvalidStack"
    End If

    If gPlanningWorkflowStack.Count > 0 Then
        Set parentCtx = gPlanningWorkflowStack(gPlanningWorkflowStack.Count)
        parentId = CStr(parentCtx("WorkflowId"))
        rootId = CStr(parentCtx("RootWorkflowId"))
        depth = CLng(parentCtx("WorkflowDepth")) + 1
    Else
        parentId = vbNullString
        depth = 1
    End If

    gPlanningWorkflowSequence = gPlanningWorkflowSequence + 1
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx("WorkflowId") = BuildPlanningWorkflowId(sourceProcedure)
    ctx("ParentWorkflowId") = parentId
    If rootId = "" Then
        ctx("RootWorkflowId") = CStr(ctx("WorkflowId"))
    Else
        ctx("RootWorkflowId") = rootId
    End If
    ctx("WorkflowDepth") = depth
    ctx("SourceProcedure") = Trim$(sourceProcedure)
    ctx("StartedAt") = Now

    If Not ValidatePlanningWorkflowContext(ctx, depth, parentId, CStr(ctx("RootWorkflowId"))) Then
        ResetPlanningWorkflowContextSafe "BeginPlanningWorkflow.InvalidNewContext"
        Exit Function
    End If

    gPlanningWorkflowStack.Add ctx
    If depth = 1 Then
        gPlanningWorkflowDisplayOwnerId = CStr(ctx("WorkflowId"))
    End If
    BeginPlanningWorkflow = CStr(ctx("WorkflowId"))

End Function

Public Sub EndPlanningWorkflow()

    EnsurePlanningWorkflowStack
    If gPlanningWorkflowStack.Count = 0 Then
        TracePlanningWorkflow "EndPlanningWorkflow ignored: stack is already empty."
        Exit Sub
    End If

    gPlanningWorkflowStack.Remove gPlanningWorkflowStack.Count
    If gPlanningWorkflowStack.Count = 0 Then
        gPlanningWorkflowDisplayOwnerId = vbNullString
    End If
    If Not ValidatePlanningWorkflowStack() Then
        ResetPlanningWorkflowContextSafe "EndPlanningWorkflow.InvalidStackAfterPop"
    End If

End Sub

Public Sub ClearPlanningWorkflowContext()

    Set gPlanningWorkflowStack = New Collection
    gPlanningWorkflowDisplayOwnerId = vbNullString
    Set gPlanningWorkflowDeferredDisplayMessages = New Collection
    gPlanningWorkflowStopOnlyDisplayDepth = 0

End Sub

Public Sub BeginPlanningWorkflowStopOnlyDisplay()

    gPlanningWorkflowStopOnlyDisplayDepth = gPlanningWorkflowStopOnlyDisplayDepth + 1

End Sub

Public Sub EndPlanningWorkflowStopOnlyDisplay()

    If gPlanningWorkflowStopOnlyDisplayDepth > 0 Then
        gPlanningWorkflowStopOnlyDisplayDepth = gPlanningWorkflowStopOnlyDisplayDepth - 1
    End If

End Sub

Public Function IsPlanningWorkflowStopOnlyDisplay() As Boolean

    IsPlanningWorkflowStopOnlyDisplay = (gPlanningWorkflowStopOnlyDisplayDepth > 0)

End Function

Public Sub ResetPlanningWorkflowContextSafe(Optional ByVal reason As String = "")

    TracePlanningWorkflow "Reset workflow context. Reason: " & Trim$(reason)
    ClearPlanningWorkflowContext

End Sub

Public Function EnsurePlanningWorkflowStarted(Optional ByVal sourceProcedure As String = "") As Boolean

    EnsurePlanningWorkflowStack
    If Not ValidatePlanningWorkflowStack() Then
        ResetPlanningWorkflowContextSafe "EnsurePlanningWorkflowStarted.InvalidStack"
    End If

    If gPlanningWorkflowStack.Count = 0 Then
        BeginPlanningWorkflow sourceProcedure
        EnsurePlanningWorkflowStarted = True
    Else
        EnsurePlanningWorkflowStarted = False
    End If

End Function

Public Function IsPlanningWorkflowActive() As Boolean

    EnsurePlanningWorkflowStack
    If Not ValidatePlanningWorkflowStack() Then
        ResetPlanningWorkflowContextSafe "IsPlanningWorkflowActive.InvalidStack"
    End If
    IsPlanningWorkflowActive = (gPlanningWorkflowStack.Count > 0)

End Function

Public Function GetCurrentPlanningWorkflow() As Object

    EnsurePlanningWorkflowStack
    If Not ValidatePlanningWorkflowStack() Then
        ResetPlanningWorkflowContextSafe "GetCurrentPlanningWorkflow.InvalidStack"
    End If

    If gPlanningWorkflowStack.Count = 0 Then
        Set GetCurrentPlanningWorkflow = Nothing
    Else
        Set GetCurrentPlanningWorkflow = gPlanningWorkflowStack(gPlanningWorkflowStack.Count)
    End If

End Function

Public Function ValidatePlanningWorkflowStack() As Boolean

    Dim i As Long
    Dim ctx As Object
    Dim previousCtx As Object
    Dim rootId As String
    Dim expectedParentId As String

    EnsurePlanningWorkflowStack
    ValidatePlanningWorkflowStack = True

    For i = 1 To gPlanningWorkflowStack.Count
        If Not IsObject(gPlanningWorkflowStack(i)) Then
            TracePlanningWorkflow "Invalid workflow stack: item is not an object at depth " & CStr(i)
            ValidatePlanningWorkflowStack = False
            Exit Function
        End If

        Set ctx = gPlanningWorkflowStack(i)
        If i = 1 Then
            expectedParentId = vbNullString
            If WorkflowContextHasKey(ctx, "WorkflowId") Then rootId = CStr(ctx("WorkflowId"))
        Else
            Set previousCtx = gPlanningWorkflowStack(i - 1)
            expectedParentId = CStr(previousCtx("WorkflowId"))
        End If

        If Not WorkflowContextHasKey(ctx, "RootWorkflowId") Then
            ValidatePlanningWorkflowStack = False
            Exit Function
        End If

        If rootId = "" Then rootId = CStr(ctx("RootWorkflowId"))

        If Not ValidatePlanningWorkflowContext(ctx, i, expectedParentId, rootId) Then
            ValidatePlanningWorkflowStack = False
            Exit Function
        End If
    Next i

End Function

Public Function IsRootPlanningWorkflow() As Boolean

    Dim ctx As Object

    Set ctx = GetCurrentPlanningWorkflow()
    If ctx Is Nothing Then
        IsRootPlanningWorkflow = True
    Else
        IsRootPlanningWorkflow = (CLng(ctx("WorkflowDepth")) <= 1)
    End If

End Function

Public Function GetPlanningWorkflowDisplayOwnerId() As String

    GetPlanningWorkflowDisplayOwnerId = gPlanningWorkflowDisplayOwnerId

End Function

Public Function CanCurrentWorkflowDisplay(Optional ByVal sourceProcedure As String = "") As Boolean

    Dim ctx As Object

    Set ctx = GetCurrentPlanningWorkflow()

    'Soft transition: child displays are still allowed for compatibility.
    'The trace makes future root-only ownership observable before enforcement.
    CanCurrentWorkflowDisplay = True

    If ctx Is Nothing Then
        TracePlanningWorkflow "Display requested without active workflow. Source=" & Trim$(sourceProcedure)
    ElseIf Not IsRootPlanningWorkflow() Then
        TracePlanningWorkflow "Child workflow display allowed in soft mode. Source=" & _
            Trim$(sourceProcedure) & " WorkflowId=" & CStr(ctx("WorkflowId")) & _
            " RootWorkflowId=" & CStr(ctx("RootWorkflowId"))
    ElseIf Trim$(gPlanningWorkflowDisplayOwnerId) <> "" _
        And CStr(ctx("WorkflowId")) <> gPlanningWorkflowDisplayOwnerId Then
        TracePlanningWorkflow "Display requested by non-owner root in soft mode. Source=" & _
            Trim$(sourceProcedure) & " WorkflowId=" & CStr(ctx("WorkflowId")) & _
            " DisplayOwner=" & gPlanningWorkflowDisplayOwnerId
    End If

End Function

Public Function ShouldDeferCurrentWorkflowDisplayToRoot(Optional ByVal sourceProcedure As String = "") As Boolean

    Dim ctx As Object
    Dim rootCtx As Object
    Dim rootSource As String

    Set ctx = GetCurrentPlanningWorkflow()
    If ctx Is Nothing Then Exit Function
    If IsRootPlanningWorkflow() Then Exit Function

    Set rootCtx = GetRootPlanningWorkflow()
    If rootCtx Is Nothing Then Exit Function

    rootSource = UCase$(Trim$(CStr(rootCtx("SourceProcedure"))))
    ShouldDeferCurrentWorkflowDisplayToRoot = (rootSource = "RUN_GANTT_LOCK_CHANGES")

    If ShouldDeferCurrentWorkflowDisplayToRoot Then
        TracePlanningWorkflow "Child workflow display deferred to root. Source=" & _
            Trim$(sourceProcedure) & " WorkflowId=" & CStr(ctx("WorkflowId")) & _
            " RootWorkflowId=" & CStr(ctx("RootWorkflowId"))
    End If

End Function

Public Sub DeferPlanningWorkflowDisplayMessages(ByVal messages As Collection)

    Dim item As Variant
    Dim deferredItem As Object

    If messages Is Nothing Then Exit Sub
    If messages.Count = 0 Then Exit Sub

    EnsurePlanningWorkflowDeferredMessages

    For Each item In messages
        Set deferredItem = CreateObject("Scripting.Dictionary")
        deferredItem("Type") = CStr(item("Type"))
        deferredItem("Message") = CStr(item("Message"))
        deferredItem("HistoryHandled") = True
        gPlanningWorkflowDeferredDisplayMessages.Add deferredItem
    Next item

End Sub

Public Sub DrainPlanningWorkflowDeferredDisplayMessages(ByVal targetMessages As Collection)

    Dim item As Variant

    If targetMessages Is Nothing Then Exit Sub

    EnsurePlanningWorkflowDeferredMessages
    If gPlanningWorkflowDeferredDisplayMessages.Count = 0 Then Exit Sub

    For Each item In gPlanningWorkflowDeferredDisplayMessages
        targetMessages.Add item
    Next item

    Set gPlanningWorkflowDeferredDisplayMessages = New Collection

End Sub

Public Function GetCurrentPlanningWorkflowId() As String

    Dim ctx As Object

    Set ctx = GetCurrentPlanningWorkflow()
    If ctx Is Nothing Then
        GetCurrentPlanningWorkflowId = vbNullString
    Else
        GetCurrentPlanningWorkflowId = CStr(ctx("WorkflowId"))
    End If

End Function

Public Function GetCurrentPlanningRootWorkflowId() As String

    Dim ctx As Object

    Set ctx = GetCurrentPlanningWorkflow()
    If ctx Is Nothing Then
        GetCurrentPlanningRootWorkflowId = vbNullString
    Else
        GetCurrentPlanningRootWorkflowId = CStr(ctx("RootWorkflowId"))
    End If

End Function

Public Function GetCurrentPlanningWorkflowDepth() As Long

    Dim ctx As Object

    Set ctx = GetCurrentPlanningWorkflow()
    If ctx Is Nothing Then
        GetCurrentPlanningWorkflowDepth = 0
    Else
        GetCurrentPlanningWorkflowDepth = CLng(ctx("WorkflowDepth"))
    End If

End Function

Private Function GetRootPlanningWorkflow() As Object

    EnsurePlanningWorkflowStack
    If gPlanningWorkflowStack.Count = 0 Then
        Set GetRootPlanningWorkflow = Nothing
    Else
        Set GetRootPlanningWorkflow = gPlanningWorkflowStack(1)
    End If

End Function

Private Sub EnsurePlanningWorkflowStack()

    If gPlanningWorkflowStack Is Nothing Then
        Set gPlanningWorkflowStack = New Collection
    End If

End Sub

Private Sub EnsurePlanningWorkflowDeferredMessages()

    If gPlanningWorkflowDeferredDisplayMessages Is Nothing Then
        Set gPlanningWorkflowDeferredDisplayMessages = New Collection
    End If

End Sub

Private Function ValidatePlanningWorkflowContext( _
    ByVal ctx As Object, _
    ByVal expectedDepth As Long, _
    ByVal expectedParentId As String, _
    ByVal expectedRootId As String) As Boolean

    ValidatePlanningWorkflowContext = False

    If ctx Is Nothing Then
        TracePlanningWorkflow "Invalid workflow context: context is Nothing."
        Exit Function
    End If

    If Not WorkflowContextHasKey(ctx, "WorkflowId") Then Exit Function
    If Not WorkflowContextHasKey(ctx, "ParentWorkflowId") Then Exit Function
    If Not WorkflowContextHasKey(ctx, "RootWorkflowId") Then Exit Function
    If Not WorkflowContextHasKey(ctx, "WorkflowDepth") Then Exit Function
    If Not WorkflowContextHasKey(ctx, "SourceProcedure") Then Exit Function
    If Not WorkflowContextHasKey(ctx, "StartedAt") Then Exit Function

    If Trim$(CStr(ctx("WorkflowId"))) = "" Then
        TracePlanningWorkflow "Invalid workflow context: empty WorkflowId."
        Exit Function
    End If

    If CLng(ctx("WorkflowDepth")) <> expectedDepth Or CLng(ctx("WorkflowDepth")) < 1 Then
        TracePlanningWorkflow "Invalid workflow context depth. Actual=" & _
            CStr(ctx("WorkflowDepth")) & " Expected=" & CStr(expectedDepth)
        Exit Function
    End If

    If CStr(ctx("ParentWorkflowId")) <> expectedParentId Then
        TracePlanningWorkflow "Invalid workflow parent. Actual=" & _
            CStr(ctx("ParentWorkflowId")) & " Expected=" & expectedParentId
        Exit Function
    End If

    If CStr(ctx("RootWorkflowId")) <> expectedRootId Or Trim$(CStr(ctx("RootWorkflowId"))) = "" Then
        TracePlanningWorkflow "Invalid workflow root. Actual=" & _
            CStr(ctx("RootWorkflowId")) & " Expected=" & expectedRootId
        Exit Function
    End If

    ValidatePlanningWorkflowContext = True

End Function

Private Function WorkflowContextHasKey(ByVal ctx As Object, ByVal keyName As String) As Boolean

    On Error GoTo MissingKey
    WorkflowContextHasKey = ctx.Exists(keyName)
    If Not WorkflowContextHasKey Then
        TracePlanningWorkflow "Invalid workflow context: missing key " & keyName
    End If
    Exit Function

MissingKey:
    WorkflowContextHasKey = False
    TracePlanningWorkflow "Invalid workflow context: cannot read key " & keyName

End Function

Private Sub TracePlanningWorkflow(ByVal messageText As String)

    Debug.Print "PlanningWorkflow: " & messageText

End Sub

Private Function BuildPlanningWorkflowId(ByVal sourceProcedure As String) As String

    Dim cleanSource As String

    cleanSource = Trim$(sourceProcedure)
    If cleanSource = "" Then cleanSource = "Workflow"

    BuildPlanningWorkflowId = _
        Format$(Now, "yyyymmdd-hhnnss") & "-" & _
        Format$(gPlanningWorkflowSequence, "0000") & "-" & _
        CleanPlanningWorkflowIdPart(cleanSource)

End Function

Private Function CleanPlanningWorkflowIdPart(ByVal value As String) As String

    Dim i As Long
    Dim ch As String
    Dim result As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i

    CleanPlanningWorkflowIdPart = result

End Function

Option Explicit

' Pure WT state engine.
' No SAP, Excel navigation, worksheet writes, Select, or Activate are allowed here.
' Remaining Qty is supplied by a later read-only reader; this module only decides
' the state from the values it receives.

Public Const WT_STATE_DONE As String = "WT_DONE"
Public Const WT_STATE_ACTIVE As String = "WT_ACTIVE"
Public Const WT_STATE_CANCELLED_PENDING As String = "WT_CANCELLED_PENDING_REMAINING"
Public Const WT_STATE_CANCELLED_REMAINING As String = "WT_CANCELLED_REMAINING"
Public Const WT_STATE_CANCELLED_ZERO As String = "WT_CANCELLED_ZERO"
Public Const WT_STATE_ERROR As String = "WT_STATE_ERROR"

Public Const WT_DECISION_OB_DONE As String = "NO_ACTION_OB_DONE"
Public Const WT_DECISION_DONE As String = "NO_ACTION_DONE"
Public Const WT_DECISION_WAIT_REMAINING As String = "WAIT_REMAINING_QTY"
Public Const WT_DECISION_REPROCESS As String = "REPROCESS_REQUIRED"
Public Const WT_DECISION_STOP_USER As String = "STOP_USER_ACTION"
Public Const WT_DECISION_STOP As String = "STOP"
Public Const WT_DECISION_ACTIVE As String = "TBD_BY_PREPLAN"
Public Const WT_POSTING_FALSE As String = "FALSE"
Public Const WT_POSTING_TBD As String = "TBD_BY_PREPLAN"

' Normalize raw OB/WT/Remaining inputs into one deterministic state.
' PostingAllowed remains False in V1.2.1 because Posting is decided later by
' the pre-plan; the decision string exposes that it is still TBD for active rows.
Public Sub WT_NormalizeState(ByVal obState As String, ByVal rawWTStatus As String, _
                             ByVal remainingQtyKnown As Boolean, ByVal remainingQty As Double, _
                             ByRef rowState As String, ByRef processAllowed As Boolean, _
                             ByRef postingAllowed As String, ByRef needRemaining As Boolean, _
                             ByRef decision As String, ByRef reasonText As String)
    Dim ob As String, wt As String

    rowState = WT_STATE_ERROR
    processAllowed = False
    postingAllowed = WT_POSTING_FALSE
    needRemaining = False
    decision = WT_DECISION_STOP
    reasonText = "WT_STATE_ERROR"

    ob = UCase$(Trim$(obState))
    wt = UCase$(Trim$(rawWTStatus))

    If ob = "OB_DONE" Then
        rowState = WT_STATE_DONE
        decision = WT_DECISION_OB_DONE
        reasonText = "OB_DONE has priority; WT is audit/display only."
        Exit Sub
    End If

    If ob = "OB_STATE_NOT_PROVEN" Or ob <> "OB_OPEN_OR_PARTIAL" Then
        reasonText = "OB_STATE_NOT_PROVEN: OB header state is not proven."
        Exit Sub
    End If

    Select Case wt
        Case "C"
            rowState = WT_STATE_DONE
            decision = WT_DECISION_DONE
            reasonText = "WT status C: WT is already confirmed."

        Case ""
            rowState = WT_STATE_ACTIVE
            processAllowed = True
            postingAllowed = WT_POSTING_TBD
            decision = WT_DECISION_ACTIVE
            reasonText = "Blank WT status: active user-picking row."

        Case "A"
            If Not remainingQtyKnown Then
                rowState = WT_STATE_CANCELLED_PENDING
                needRemaining = True
                decision = WT_DECISION_WAIT_REMAINING
                reasonText = "REMAINING_QTY_REQUIRED: cancelled WT needs a proven Remaining Qty."
            ElseIf remainingQty < 0# Then
                reasonText = "REMAINING_QTY_NOT_PROVEN: Remaining Qty cannot be negative."
            ElseIf remainingQty > 0# Then
                rowState = WT_STATE_CANCELLED_REMAINING
                processAllowed = True
                postingAllowed = WT_POSTING_TBD
                decision = WT_DECISION_REPROCESS
                reasonText = "Cancelled WT has Remaining Qty > 0; reprocess is required."
            Else
                rowState = WT_STATE_CANCELLED_ZERO
                decision = WT_DECISION_STOP_USER
                reasonText = "OB_QTY_ZERO_REQUIRES_USER_ACTION: cancelled WT has Remaining Qty=0."
            End If

        Case Else
            reasonText = "UNKNOWN_WT_STATUS: raw WT status is not A, C, or blank."
    End Select
End Sub

Public Function WT_IsProcessAllowed(ByVal obState As String, ByVal rawWTStatus As String, _
                                    ByVal remainingQtyKnown As Boolean, ByVal remainingQty As Double) As Boolean
    Dim rowState As String, postingAllowed As String, needRemaining As Boolean
    Dim decision As String, reasonText As String, result As Boolean
    WT_NormalizeState obState, rawWTStatus, remainingQtyKnown, remainingQty, _
                      rowState, result, postingAllowed, needRemaining, _
                      decision, reasonText
    WT_IsProcessAllowed = result
End Function

Public Function WT_NeedsRemainingQty(ByVal obState As String, ByVal rawWTStatus As String, _
                                     ByVal remainingQtyKnown As Boolean, ByVal remainingQty As Double) As Boolean
    Dim rowState As String, processAllowed As Boolean, postingAllowed As String
    Dim decision As String, reasonText As String, result As Boolean
    WT_NormalizeState obState, rawWTStatus, remainingQtyKnown, remainingQty, _
                      rowState, processAllowed, postingAllowed, result, _
                      decision, reasonText
    WT_NeedsRemainingQty = result
End Function

Public Function WT_GetDecision(ByVal obState As String, ByVal rawWTStatus As String, _
                               ByVal remainingQtyKnown As Boolean, ByVal remainingQty As Double) As String
    Dim rowState As String, processAllowed As Boolean, postingAllowed As String
    Dim needRemaining As Boolean, reasonText As String, result As String
    WT_NormalizeState obState, rawWTStatus, remainingQtyKnown, remainingQty, _
                      rowState, processAllowed, postingAllowed, needRemaining, _
                      result, reasonText
    WT_GetDecision = result
End Function

' Compile/acceptance probe. It is pure logic only and intentionally writes nothing.
Public Sub WT_StateEngine_CompileProbe()
    Dim rowState As String, decision As String, reasonText As String
    Dim processAllowed As Boolean, postingAllowed As String, needRemaining As Boolean

    WT_NormalizeState "OB_DONE", "", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_DONE And Not processAllowed And decision = WT_DECISION_OB_DONE, "OB_DONE"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "C", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_DONE And Not processAllowed And decision = WT_DECISION_DONE, "C"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "A", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_CANCELLED_PENDING And Not processAllowed And needRemaining _
              And decision = WT_DECISION_WAIT_REMAINING, "A pending"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "A", True, 5#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_CANCELLED_REMAINING And processAllowed And Not needRemaining _
              And decision = WT_DECISION_REPROCESS, "A remaining"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "A", True, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_CANCELLED_ZERO And Not processAllowed And decision = WT_DECISION_STOP_USER, "A zero"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_ACTIVE And processAllowed And decision = WT_DECISION_ACTIVE, "Blank"

    WT_NormalizeState "OB_OPEN_OR_PARTIAL", "X", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_ERROR And Not processAllowed And InStr(1, reasonText, "UNKNOWN_WT_STATUS", vbTextCompare) > 0, "Unknown"

    WT_NormalizeState "OB_STATE_NOT_PROVEN", "", False, 0#, rowState, processAllowed, postingAllowed, needRemaining, decision, reasonText
    WT_Assert rowState = WT_STATE_ERROR And Not processAllowed And decision = WT_DECISION_STOP, "OB not proven"
End Sub

Private Sub WT_Assert(ByVal condition As Boolean, ByVal caseName As String)
    If Not condition Then Err.Raise vbObjectError + 9780, , "WT_STATE_ENGINE_PROBE_FAIL: " & caseName
End Sub

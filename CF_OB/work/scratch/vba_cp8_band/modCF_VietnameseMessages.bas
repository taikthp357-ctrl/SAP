Attribute VB_Name = "modCF_VietnameseMessages"
Option Explicit

' Vietnamese message mapper for user-facing feedback.
' Maps Result enum (English technical) + Detail context to Vietnamese string.
' Used by Phase 1b worklist to display CHECK feedback in Vietnamese.

Public Function CF_VN_GetMessage(ByVal result As String, ByVal detail As String) As String
    Dim msg As String
    result = UCase$(Trim$(result))
    detail = Trim$(detail)
    Select Case result
        Case "PASS"
            msg = "S" & ChrW(7855) & "n s" & ChrW(224) & "ng x" & ChrW(7917) & " l" & ChrW(253)
        Case "BLOCKED_SOURCE_PLAN_INCOMPLETE"
            If InStr(1, detail, "Multi-batch", vbTextCompare) > 0 Then
                msg = "V" & ChrW(7883) & " tr" & ChrW(237) & " c" & ChrW(243) & " nhi" & ChrW(7873) & "u l" & ChrW(244) & ", c" & ChrW(7847) & "n ch" & ChrW(7885) & "n l" & ChrW(244) & " c" & ChrW(7909) & " th" & ChrW(7875)
            ElseIf InStr(1, detail, "PICK_BIN_CANDIDATE_COUNT", vbTextCompare) > 0 Then
                msg = "C" & ChrW(243) & " nhi" & ChrW(7873) & "u v" & ChrW(7883) & " tr" & ChrW(237) & " ngu" & ChrW(7891) & "n, c" & ChrW(7847) & "n ch" & ChrW(7885) & "n v" & ChrW(7883) & " tr" & ChrW(237) & " c" & ChrW(7909) & " th" & ChrW(7875)
            Else
                msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u ngu" & ChrW(7891) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911)
            End If
        Case "BLOCKED_QTY_GROUP_BASELINE"
            msg = "Nh" & ChrW(243) & "m s" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng ch" & ChrW(432) & "a x" & ChrW(225) & "c " & ChrW(273) & ChrW(7883) & "nh " & ChrW(273) & ChrW(432) & ChrW(7907) & "c baseline"
        Case "BLOCKED_INCOMPLETE_EXECUTION"
            msg = "H" & ChrW(224) & "nh " & ChrW(273) & ChrW(7897) & "ng tr" & ChrW(432) & ChrW(7899) & "c ch" & ChrW(432) & "a ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
        Case "BLOCKED_CREATE_OPEN"
            msg = "Ch" & ChrW(7913) & "ng t" & ChrW(7915) & " PRDO " & ChrW(273) & "ang m" & ChrW(7903) & ", c" & ChrW(7847) & "n " & ChrW(273) & ChrW(243) & "ng tr" & ChrW(432) & ChrW(7899) & "c khi ti" & ChrW(7871) & "p t" & ChrW(7909) & "c"
        Case "STOP"
            If InStr(1, detail, "SOURCE_PLAN_INCOMPLETE", vbTextCompare) > 0 Then
                msg = "D" & ChrW(7915) & "ng: d" & ChrW(7919) & " li" & ChrW(7879) & "u ngu" & ChrW(7891) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7911)
            ElseIf InStr(1, detail, "REGRESSION_GATE", vbTextCompare) > 0 Then
                msg = "D" & ChrW(7915) & "ng: kh" & ChrW(244) & "ng th" & ChrW(7887) & "a " & ChrW(273) & "i" & ChrW(7873) & "u ki" & ChrW(7879) & "n an to" & ChrW(224) & "n regression"
            Else
                msg = "D" & ChrW(7915) & "ng: " & detail
            End If
        Case "READY_FOR_CF"
            msg = "S" & ChrW(7855) & "n s" & ChrW(224) & "ng " & ChrW(273) & ChrW(7875) & " user x" & ChrW(225) & "c nh" & ChrW(7853) & "n WT"
        Case "DONE_BY_USER_CF"
            msg = ChrW(272) & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t b" & ChrW(7903) & "i user CF"
        Case "DONE"
            msg = "Ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
        Case "PLAN_NEEDS_TAI_POSTING"
            msg = "C" & ChrW(7847) & "n TAI POSTING " & ChrW(273) & ChrW(7875) & " ti" & ChrW(7871) & "p t" & ChrW(7909) & "c"
        Case Else
            msg = result
            If detail <> "" Then msg = msg & " (" & Left$(detail, 100) & ")"
    End Select
    CF_VN_GetMessage = msg
End Function

Public Function CF_VN_GetStageLabel(ByVal stage As String) As String
    stage = UCase$(Trim$(stage))
    Select Case stage
        Case "CHECK PREPARE":            CF_VN_GetStageLabel = "Chu" & ChrW(7849) & "n b" & ChrW(7883) & " ki" & ChrW(7875) & "m tra"
        Case "CHECK DONE":               CF_VN_GetStageLabel = "Ki" & ChrW(7875) & "m tra ho" & ChrW(224) & "n t" & ChrW(7845) & "t"
        Case "EXECUTION_CONTEXT_REUSE":  CF_VN_GetStageLabel = "T" & ChrW(225) & "i s" & ChrW(7917) & " d" & ChrW(7909) & "ng context"
        Case "CANCEL_QUEUE":             CF_VN_GetStageLabel = "H" & ChrW(224) & "ng " & ChrW(273) & ChrW(7907) & "i cancel"
        Case "CANCEL_GROUP":             CF_VN_GetStageLabel = "Cancel nh" & ChrW(243) & "m"
        Case "CANCEL_BATCH_READBACK":    CF_VN_GetStageLabel = ChrW(272) & ChrW(7885) & "c l" & ChrW(7841) & "i sau cancel"
        Case "LOCATION_GROUP":           CF_VN_GetStageLabel = "Thay " & ChrW(273) & ChrW(7893) & "i v" & ChrW(7883) & " tr" & ChrW(237)
        Case "CREATE_GROUP":             CF_VN_GetStageLabel = "T" & ChrW(7841) & "o WT m" & ChrW(7899) & "i"
        Case "FINAL_VERIFY":             CF_VN_GetStageLabel = "X" & ChrW(225) & "c minh cu" & ChrW(7889) & "i"
        Case "WAIT_USER_CF":             CF_VN_GetStageLabel = "Ch" & ChrW(7901) & " user CF"
        Case Else:                       CF_VN_GetStageLabel = stage
    End Select
End Function
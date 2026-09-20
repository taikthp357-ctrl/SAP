import os
import re

migo_path = r"g:\SAP\v184_patch\MIGO311_RunPost_V24_ErrorGuard.bas"
with open(migo_path, "r", encoding="utf-8") as f:
    code = f.read()

# 1. Update ATT_Log and ATT_History signatures & implementations
old_log_and_hist = """Private Sub ATT_Log(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = stage: ws.Cells(r, 3).value = status: ws.Cells(r, 4).value = doc: ws.Cells(r, 5).value = msg
End Sub

Private Sub ATT_History(ByVal doc As String, ByVal ticket As String, ByVal typ As String, ByVal path As String, ByVal result As String)
    Dim ws As Worksheet, r As Long: Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = doc: ws.Cells(r, 3).value = ticket: ws.Cells(r, 4).value = typ: ws.Cells(r, 5).value = path: ws.Cells(r, 6).value = result
End Sub"""

new_log_and_hist = """Private Sub ATT_Log(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal msg As String, Optional ByVal filePath As String = "", Optional ByVal runId As String = "")
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = stage
    ws.Cells(r, 3).Value = status
    ws.Cells(r, 4).Value = doc
    ws.Cells(r, 5).Value = msg
    ws.Cells(r, 6).Value = filePath
    ws.Cells(r, 7).Value = IIf(runId <> "", runId, Format$(Now, "yyyymmdd_hhnnss"))
    On Error GoTo 0
End Sub

Private Sub ATT_History(ByVal doc As String, ByVal ticket As String, ByVal typ As String, ByVal path As String, ByVal result As String, Optional ByVal sapUser As String = "", Optional ByVal note As String = "")
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = doc
    ws.Cells(r, 3).Value = ticket
    ws.Cells(r, 4).Value = typ
    ws.Cells(r, 5).Value = path
    ws.Cells(r, 6).Value = result
    ws.Cells(r, 7).Value = sapUser
    ws.Cells(r, 8).Value = note
    On Error GoTo 0
End Sub"""

if old_log_and_hist in code:
    code = code.replace(old_log_and_hist, new_log_and_hist)
    print("Replaced ATT_Log and ATT_History successfully.")
else:
    print("WARNING: Could not find exact old_log_and_hist block!")

# 2. Update SAP_FillHeadTextOnly with rich history and diagnostics
old_headtext_pattern = re.compile(r"Public Sub SAP_FillHeadTextOnly\(\).*?End Sub", re.DOTALL)

new_headtext_sub = """Public Sub SAP_FillHeadTextOnly()
    Dim ws As Worksheet, docs As Object, ses As Object, k As Variant, d As Object
    Dim t0 As Double, errText As String, currentText As String, n As Long
    Dim sapUser As String, expectedTicket As String, totalDocs As Long
    Dim matchedCount As Long, diffCount As Long, errCount As Long, runId As String
    
    If Not ToolTryBegin("DIEN HEADTEXT") Then Exit Sub
    t0 = Timer
    runId = "HEADTEXT_" & Format$(Now, "yyyymmdd_hhnnss")
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Set docs = ATT_BuildDocTickets(ws)
    If docs.Count = 0 Then Err.Raise vbObjectError + 7010, , "Khong co DOC + So phieu de xu ly."
    
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7011, , "Khong tim thay SAP session. Hay dang nhap SAP truoc."
    ATT_ActivateMigoSession ses
    
    On Error Resume Next
    sapUser = CStr(ses.Info.User)
    If sapUser = "" Then sapUser = Environ$("USERNAME")
    On Error GoTo EH
    
    ATT_Log "HEADTEXT_RUN", "START", "", "Bat dau dien HeadText cho " & docs.Count & " DOC", "", runId
    
    totalDocs = docs.Count
    For Each k In docs.keys
        Set d = docs(k)
        If d("TICKETS").Count = 0 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "NEED_SOPHIEU", "Can nhap So phieu tu MB51."
            ATT_History CStr(k), "", "HEADTEXT", "", "FAIL", sapUser, "Thieu So phieu trong INPUT (cot 4). Hay copy tu MB51 vao."
            ATT_Log "HEADTEXT", "NEED_SOPHIEU", CStr(k), "DOC thieu So phieu", "", runId
            errCount = errCount + 1
        ElseIf d("TICKETS").Count > 1 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_MULTI_SOPHIEU", "DOC co nhieu So phieu; khong tu ghi de HeadText."
            ATT_History CStr(k), ATT_JoinKeys(d("TICKETS"), "; "), "HEADTEXT", "", "SKIP", sapUser, "DOC co nhieu So phieu khac nhau. Can kiem tra tach dong thu cong."
            ATT_Log "HEADTEXT", "MULTI_SOPHIEU", CStr(k), "Nhieu so phieu: " & ATT_JoinKeys(d("TICKETS"), "; "), "", runId
            errCount = errCount + 1
        Else
            expectedTicket = ATT_FirstKey(d("TICKETS"))
            Application.StatusBar = "Dien HeadText DOC " & CStr(k) & " (" & expectedTicket & ")"
            errText = "": currentText = ""
            
            If ATT_OpenMigoDoc(ses, CStr(k), errText) Then
                If ATT_ReadHeadText(ses, currentText, errText) Then
                    If Trim$(currentText) = "" Then
                        If ATT_WriteHeadText(ses, expectedTicket, errText) Then
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_DONE", ""
                            ATT_History CStr(k), expectedTicket, "HEADTEXT", "'' -> " & expectedTicket, "DONE", sapUser, "Da dien moi va luu HeadText thanh cong tren SAP MIGO."
                            ATT_Log "HEADTEXT", "DONE", CStr(k), "Dien moi: " & expectedTicket, "", runId
                            n = n + 1
                        Else
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_FAIL", errText
                            ATT_History CStr(k), expectedTicket, "HEADTEXT", expectedTicket, "FAIL", sapUser, "Loi ghi HeadText tren SAP: " & errText & ". Kiem tra quyen ghi MIGO."
                            ATT_Log "HEADTEXT", "WRITE_FAIL", CStr(k), errText, "", runId
                            errCount = errCount + 1
                        End If
                    ElseIf ATT_SameToken(currentText, expectedTicket) Then
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_OK", "Da co dung So phieu."
                        ATT_History CStr(k), expectedTicket, "HEADTEXT", currentText, "OK", sapUser, "HeadText tren SAP da trung khop ('" & currentText & "'). Khong can ghi lai."
                        ATT_Log "HEADTEXT", "MATCH_OK", CStr(k), "SAP=" & currentText, "", runId
                        matchedCount = matchedCount + 1
                    Else
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_EXISTS_DIFFERENT", "SAP=" & currentText
                        ATT_History CStr(k), expectedTicket, "HEADTEXT", "SAP: " & currentText & " | INPUT: " & expectedTicket, "DIFF", sapUser, "HeadText tren SAP ('" & currentText & "') khac voi So phieu INPUT ('" & expectedTicket & "'). Can kiem tra lai."
                        ATT_Log "HEADTEXT", "DIFF_WARNING", CStr(k), "SAP=" & currentText & " | INPUT=" & expectedTicket, "", runId
                        diffCount = diffCount + 1
                    End If
                Else
                    ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_READ_FAIL", errText
                    ATT_History CStr(k), expectedTicket, "HEADTEXT", expectedTicket, "FAIL", sapUser, "Loi doc HeadText: " & errText & ". Kiem tra header MIGO."
                    ATT_Log "HEADTEXT", "READ_FAIL", CStr(k), errText, "", runId
                    errCount = errCount + 1
                End If
                
                ' Tra ve man hinh nhap DOC de san sang cho DOC tiep theo
                ATT_ReturnToMigoEntry ses, errText
            Else
                ATT_SetDocStatus ws, CStr(k), C_HEAD, "OPEN_DOC_FAIL", errText
                ATT_History CStr(k), expectedTicket, "HEADTEXT", expectedTicket, "FAIL", sapUser, "Loi mo DOC tren SAP: " & errText & ". Kiem tra ma DOC va quyen truy cap."
                ATT_Log "HEADTEXT", "OPEN_DOC_FAIL", CStr(k), errText, "", runId
                errCount = errCount + 1
            End If
        End If
        DoEvents
    Next k
    
    ATT_Log "HEADTEXT_RUN", "FINISH", "", "Ket thuc: Dien moi=" & n & ", Da khop=" & matchedCount & ", Khac biet=" & diffCount & ", Loi=" & errCount, "", runId
    
    MsgBox "XU LY HEADTEXT HOAN TAT!" & vbCrLf & vbCrLf & _
           "- DOC da dien moi thanh cong: " & n & vbCrLf & _
           "- DOC da co HeadText trung khop: " & matchedCount & vbCrLf & _
           "- DOC co HeadText khac INPUT: " & diffCount & vbCrLf & _
           "- DOC gap loi / thieu du lieu: " & errCount & vbCrLf & vbCrLf & _
           "Chi tiet lich su va huong dan fix loi da duoc luu tai sheet ATT_HISTORY va LOG.", _
           vbInformation, "DIEN HEADTEXT SAP"

SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0)
    ToolEnd
    Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "HEADTEXT", "FATAL", "", CStr(Err.Number) & " - " & Err.Description, "", runId
    MsgBox "DIEN HEADTEXT bi loi " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Xem chi tiet tai sheet LOG va ATT_HISTORY.", vbCritical, "LOI DIEN HEADTEXT"
    Resume SafeExit
End Sub"""

code = old_headtext_pattern.sub(new_headtext_sub, code, count=1)
print("Updated SAP_FillHeadTextOnly successfully.")

# 3. Update ATT_MarkVerified with sapUser and note
old_mark_verified = """Private Sub ATT_MarkVerified(ByVal ws As Worksheet, ByVal docInfo As Object)
    Dim k As Variant, g As Object, docNo As String
    For Each k In docInfo("GROUPS").keys
        Set g = docInfo("GROUPS")(k)
        docNo = CStr(ws.Cells(g("ROWS")(1), C_DOC).value)
        If Trim$(CStr(g("BBBG"))) <> "" Then
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "DONE"
            ATT_History docNo, CStr(g("TICKET")), "BBBG", CStr(g("BBBG")), "DONE"
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
        End If
        If Trim$(CStr(g("VAT"))) <> "" Then
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "DONE"
            ATT_History docNo, CStr(g("TICKET")), "VAT", CStr(g("VAT")), "DONE"
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "MISSING_VAT"
        End If
    Next k
End Sub"""

new_mark_verified = """Private Sub ATT_MarkVerified(ByVal ws As Worksheet, ByVal docInfo As Object, Optional ByVal sapUser As String = "")
    Dim k As Variant, g As Object, docNo As String
    For Each k In docInfo("GROUPS").keys
        Set g = docInfo("GROUPS")(k)
        docNo = CStr(ws.Cells(g("ROWS")(1), C_DOC).value)
        If Trim$(CStr(g("BBBG"))) <> "" Then
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "DONE"
            ATT_History docNo, CStr(g("TICKET")), "BBBG", CStr(g("BBBG")), "DONE", sapUser, "Attach BBBG hoan tat va xac nhan tren SAP."
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
        End If
        If Trim$(CStr(g("VAT"))) <> "" Then
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "DONE"
            ATT_History docNo, CStr(g("TICKET")), "VAT", CStr(g("VAT")), "DONE", sapUser, "Attach VAT hoan tat va xac nhan tren SAP."
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "MISSING_VAT"
        End If
    Next k
End Sub"""

if old_mark_verified in code:
    code = code.replace(old_mark_verified, new_mark_verified)
    print("Updated ATT_MarkVerified successfully.")

with open(migo_path, "w", encoding="utf-8") as f:
    f.write(code)

print("Saved updated MIGO311_RunPost_V24_ErrorGuard.bas without BOM.")

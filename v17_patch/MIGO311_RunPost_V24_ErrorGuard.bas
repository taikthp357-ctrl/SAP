Attribute VB_Name = "MIGO311_RunPost_V24_ErrorGuard"
Option Explicit

' ======================================================================
'  MIGO311_RunPost_V24_ErrorGuard  -  V1.7 SMART GOS ATT
'  Build tu baseline V1.6.4 on dinh
'  Nang cap: Smart GOS 3-step retry, SAP Security preflight/wait,
'             Unicode file dialog, logging stage day du, folder picker
'             khong khoa duong dan
' ======================================================================

#If VBA7 Then
Private Declare PtrSafe Sub ATT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub ATT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Public LastMigoRunSucceeded As Boolean
Public LastMigoProcessedCount As Long
Public LastMigoPostedCount As Long
Public LastMigoRunMessage As String
Public LastMigoTotalSeconds As Double
Private mToolBusy As Boolean
Private mStop As Boolean
Private mSourceFolder As String

Private Const SH_INPUT As String = "INPUT"
Private Const SH_SCAN As String = "FILE_SCAN"
Private Const SH_CONFIG As String = "CONFIG"
Private Const SH_HISTORY As String = "ATT_HISTORY"
Private Const SH_LOG As String = "LOG"
Private Const SH_RAW As String = "OCR_RAW"
Private Const FIRST_ROW As Long = 7
Private Const C_MAT As Long = 1
Private Const C_DOC As Long = 2
Private Const C_PO As Long = 3
Private Const C_TICKET As Long = 4
Private Const C_BBBG As Long = 5
Private Const C_VAT As Long = 6
Private Const C_OCR_PO As Long = 7
Private Const C_OCR_MAT_BBBG As Long = 8
Private Const C_OCR_TICKET As Long = 9
Private Const C_OCR_MAT_VAT As Long = 10
Private Const C_CHECK As Long = 11
Private Const C_HEAD As Long = 12
Private Const C_ATT_BBBG As Long = 13
Private Const C_ATT_VAT As Long = 14
Private Const C_ATT As Long = 15
Private Const C_NOTE As Long = 16

' Timeout defaults (seconds)
Private Const TIMEOUT_DOC_READY As Long = 30000      ' 30s
Private Const TIMEOUT_GOS_READY As Long = 15000      ' 15s
Private Const TIMEOUT_SAP_SECURITY As Long = 60000   ' 60s
Private Const TIMEOUT_ATT_DIALOG As Long = 20000     ' 20s
Private Const TIMEOUT_FILE_ACCEPTED As Long = 30000  ' 30s
Private Const TIMEOUT_RETURN_TO_DOC As Long = 20000  ' 20s

Private Const GOS_MAX_RETRY As Long = 3

' ======================================================================
' OCR / Document Read helpers (giu nguyen tu V1.6.4)
' ======================================================================
Private Function ATT_RunTesseractPass(ByVal imagePath As String, ByVal psm As String, ByRef status As String) As String
    Dim exe As String, base As String, cmd As String, rc As Long, txt As String
    exe = ATT_TesseractPath()
    If exe = "" Then status = "TESSERACT_MISSING": Exit Function
    base = ATT_TempFolder("OCR") & "\result_" & psm
    cmd = ATT_Q(exe) & " " & ATT_Q(imagePath) & " " & ATT_Q(base) & _
          " -l " & ATT_Cfg("OCR_LANGUAGE", "eng+vie") & " --psm " & psm & _
          " --dpi 300 -c preserve_interword_spaces=1"
    rc = ATT_RunWait(cmd)
    If rc <> 0 Then status = "OCR_FAIL_" & rc: Exit Function
    txt = base & ".txt"
    If Dir$(txt) = "" Then status = "OCR_NO_OUTPUT": Exit Function
    ATT_RunTesseractPass = ATT_ReadText(txt): status = "OCR_OK"
End Function

Private Function ATT_RunTesseractDigits(ByVal imagePath As String, ByVal psm As String) As String
    Dim exe As String, base As String, cmd As String, txt As String
    exe = ATT_TesseractPath()
    If exe = "" Then Exit Function
    base = ATT_TempFolder("OCRNUM") & "\digits_" & psm
    cmd = ATT_Q(exe) & " " & ATT_Q(imagePath) & " " & ATT_Q(base) & _
          " -l eng --oem 1 --psm " & psm & " --dpi 300" & _
          " -c thresholding_method=2 -c thresholding_kfactor=0.34" & _
          " -c tessedit_char_whitelist=0123456789 -c preserve_interword_spaces=1"
    If ATT_RunWait(cmd) <> 0 Then Exit Function
    txt = base & ".txt"
    If Dir$(txt) <> "" Then ATT_RunTesseractDigits = ATT_ReadText(txt)
End Function

Private Function ATT_TicketNearMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim ca As String, cb As String, i As Long, diff As Long
    ca = ATT_TicketCode(a): cb = ATT_TicketCode(b)
    If ca = "" Or cb = "" Or Len(ca) <> Len(cb) Then Exit Function
    If Len(ca) < 4 Then Exit Function
    For i = 1 To Len(ca)
        If Mid$(ca, i, 1) <> Mid$(cb, i, 1) Then diff = diff + 1
        If diff > 1 Then Exit Function
    Next i
    ATT_TicketNearMatch = (diff = 1)
End Function

Private Function ATT_AppendNote(ByVal currentText As String, ByVal newText As String) As String
    If Trim$(currentText) = "" Then ATT_AppendNote = newText Else ATT_AppendNote = currentText & " " & newText
End Function

Private Function ATT_ExtractPO(ByVal raw As String, ByVal groups As Object) As Object
    Dim found As Object, expected As Object, g As Object, poSet As Object
    Dim groupKey As Variant, poKey As Variant, re As Object
    Dim code As String, pattern As String, i As Long
    Set found = CreateObject("Scripting.Dictionary")
    Set expected = CreateObject("Scripting.Dictionary")
    For Each groupKey In groups.keys
        Set g = groups(groupKey)
        Set poSet = g("PO")
        For Each poKey In poSet.keys
            code = UCase$(ATT_Clean(CStr(poKey)))
            If Len(code) = 10 And Left$(code, 1) = "5" And IsNumeric(code) Then expected(code) = True
        Next poKey
    Next groupKey
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    For Each poKey In expected.keys
        code = CStr(poKey)
        pattern = "(^|[^0-9])"
        For i = 1 To Len(code)
            If i > 1 Then pattern = pattern & "[ -]?"
            pattern = pattern & Mid$(code, i, 1)
        Next i
        pattern = pattern & "([^0-9]|$)"
        re.pattern = pattern
        If re.Test(raw) Then found(code) = True
    Next poKey
    Set ATT_ExtractPO = found
End Function

Private Function ATT_RunTesseract(ByVal imagePath As String, ByRef status As String) As String
    Dim psm1 As String, psm2 As String, s1 As String, s2 As String, st1 As String, st2 As String
    psm1 = ATT_Cfg("OCR_PSM", "6")
    psm2 = ATT_Cfg("OCR_PSM_SECONDARY", "11")
    s1 = ATT_RunTesseractPass(imagePath, psm1, st1)
    If psm2 <> "" And psm2 <> psm1 Then s2 = ATT_RunTesseractPass(imagePath, psm2, st2)
    ATT_RunTesseract = s1
    If Trim$(s2) <> "" Then ATT_RunTesseract = ATT_RunTesseract & vbCrLf & "--- OCR PASS " & psm2 & " ---" & vbCrLf & s2
    If Trim$(ATT_RunTesseract) = "" Then status = st1 & "/" & st2 Else status = "OCR_MULTI_" & psm1 & "_" & psm2
End Function

Private Function ATT_ReadPreparedImage(ByVal imagePath As String, ByRef status As String) As String
    Dim prepared As String, out As String, stMain As String, stSecond As String, digits As String
    prepared = ATT_PreprocessImageWIA(imagePath)
    If prepared = "" Then prepared = imagePath
    out = ATT_RunTesseractPass(prepared, ATT_Cfg("OCR_PSM", "6"), stMain)
    out = out & vbCrLf & "--- IMAGE SPARSE PASS ---" & vbCrLf & _
          ATT_RunTesseractPass(prepared, ATT_Cfg("OCR_PSM_SECONDARY", "11"), stSecond)
    digits = ATT_RunTesseractDigits(prepared, ATT_Cfg("OCR_IMAGE_NUMERIC_PSM", "6"))
    If Trim$(digits) <> "" Then out = out & vbCrLf & "--- IMAGE DIGITS ONLY ---" & vbCrLf & digits
    If prepared = imagePath Then status = "IMAGE_WIA_FALLBACK" Else status = "IMAGE_WIA_PREPROCESSED"
    ATT_ReadPreparedImage = out
End Function

Private Function ATT_PreprocessImageWIA(ByVal imagePath As String) As String
    Const WIA_SCALE As String = "{4EBB0166-C18B-4065-9332-109015741711}"
    Const WIA_CONVERT As String = "{42A6E907-1D2F-4B38-AC50-31ADBE2AB3C2}"
    Const WIA_PNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Dim img As Object, proc As Object, result As Object, outFolder As String, outPath As String
    On Error GoTo EH
    outFolder = ATT_TempFolder("IMG_WIA"): outPath = outFolder & "\image_scaled.png"
    Set img = CreateObject("WIA.ImageFile"): img.LoadFile imagePath
    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add WIA_SCALE
    proc.Filters(1).Properties("MaximumWidth").value = 3600
    proc.Filters(1).Properties("MaximumHeight").value = 3600
    proc.Filters(1).Properties("PreserveAspectRatio").value = True
    proc.Filters.Add WIA_CONVERT
    proc.Filters(2).Properties("FormatID").value = WIA_PNG
    Set result = proc.Apply(img): result.SaveFile outPath
    If Dir$(outPath) <> "" Then ATT_PreprocessImageWIA = outPath
    Exit Function
EH:
    ATT_PreprocessImageWIA = ""
End Function

Private Sub ATT_MatchGroups(ByVal groups As Object, ByVal infos As Collection)
    Dim k As Variant, g As Object, info As Object, b As Collection, v As Collection
    Dim matOk As Boolean, poOk As Boolean, i As Long, bInfo As Object, ticketRead As String
    Dim ticketFuzzy As Boolean
    For Each k In groups.keys
        Set g = groups(k): Set b = New Collection: Set v = New Collection
        If Trim$(g("TICKET")) = "" Then
            g("STATUS") = "NEED_SOPHIEU": g("NOTE") = "Nhap So phieu tu MB51 roi CHECK lai.": GoTo ContinueGroup
        End If
        For i = 1 To infos.Count
            Set info = infos(i)
            If info("TYPE") <> "VAT" Then
                poOk = ATT_ContainsAll(info("PO"), g("PO"))
                matOk = ATT_Overlap(info("MAT"), g("MAT")) > 0
                ticketRead = ATT_Clean(info("TICKET"))
                If poOk And matOk Then
                    If ticketRead = "" Or ATT_TicketMatches(ticketRead, g("TICKET")) Or ATT_TicketNearMatch(ticketRead, g("TICKET")) Then b.Add info
                End If
            End If
        Next i
        If b.Count = 0 Then
            g("STATUS") = "NO_BBBG_MATCH": g("NOTE") = "Khong co BBBG khop tat ca PO/SA + Material + So phieu."
        ElseIf b.Count > 1 Then
            g("STATUS") = "MULTI_BBBG_MATCH": g("NOTE") = b.Count & " file BBBG cung khop."
        Else
            Set bInfo = b(1): g("BBBG") = bInfo("PATH")
            g("OCR_PO") = ATT_JoinKeys(bInfo("PO"), ";")
            g("OCR_MAT_B") = ATT_JoinKeys(bInfo("MAT"), ";")
            ticketRead = ATT_Clean(bInfo("TICKET"))
            ticketFuzzy = (ticketRead <> "" And Not ATT_TicketMatches(ticketRead, g("TICKET")) And ATT_TicketNearMatch(ticketRead, g("TICKET")))
            If ticketRead = "" Then
                g("OCR_TICKET") = "NOT_READ"
            ElseIf ticketFuzzy Then
                g("OCR_TICKET") = g("TICKET")
                g("NOTE") = "OCR doc " & ticketRead & "; da doi chieu PO + Material va sua thanh " & g("TICKET") & "."
            Else
                g("OCR_TICKET") = ticketRead
            End If
            For i = 1 To infos.Count
                Set info = infos(i)
                If info("TYPE") = "VAT" And ATT_FileTicketMatch(info("NAME"), g("TICKET")) Then
                    matOk = (ATT_Overlap(info("MAT"), g("MAT")) > 0 And ATT_Overlap(info("MAT"), bInfo("MAT")) > 0)
                    If matOk Then v.Add info
                End If
            Next i
            If v.Count = 0 Then
                If ATT_BoolCfg("VAT_REQUIRED", True) Then
                    g("STATUS") = "NO_VAT_MATCH": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "VAT can ten file khop So phieu va ma BIW khop BBBG.")
                Else
                    g("STATUS") = "CHECK_OK": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "Khong co VAT; CONFIG cho phep BBBG-only.")
                End If
            ElseIf v.Count > 1 Then
                g("STATUS") = "MULTI_VAT_MATCH": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), v.Count & " file VAT cung khop.")
            Else
                Set info = v(1): g("VAT") = info("PATH")
                g("OCR_MAT_V") = ATT_JoinKeys(info("MAT"), ";")
                g("STATUS") = "CHECK_OK"
                If g("OCR_TICKET") = "NOT_READ" Then g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "BBBG khong doc duoc So phieu; da xac nhan bang tat ca PO/SA + Material.")
            End If
        End If
ContinueGroup:
    Next k
End Sub

Private Function ATT_ValidateMaterialDocs(ByVal ws As Worksheet) As Boolean
    Dim r As Long, lr As Long, missingCount As Long
    Dim mat As String, doc As String, po As String, ticket As String
    lr = Application.Max(ATT_LastRow(ws, C_MAT), ATT_LastRow(ws, C_DOC), ATT_LastRow(ws, C_PO), ATT_LastRow(ws, C_TICKET))
    ATT_ValidateMaterialDocs = True
    For r = FIRST_ROW To lr
        mat = ATT_Clean(ws.Cells(r, C_MAT).value)
        doc = ATT_Clean(ws.Cells(r, C_DOC).value)
        po = ATT_Clean(ws.Cells(r, C_PO).value)
        ticket = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If doc = "" And (mat <> "" Or po <> "" Or ticket <> "") Then
            ws.Cells(r, C_CHECK).value = "INPUT_MISSING_DOC"
            ws.Cells(r, C_NOTE).value = "Thieu Material DOC cot B; hay dien DOC de MIGO co the chay."
            missingCount = missingCount + 1
        End If
    Next r
    If missingCount > 0 Then
        ATT_ValidateMaterialDocs = False
        MsgBox missingCount & " dong dang thieu Material DOC o cot B." & vbCrLf & _
               "Hay dien DOC, sau do bam CHECK lai de tool mo MIGO.", _
               vbExclamation, "CAN DIEN MATERIAL DOC"
    End If
End Function

Public Sub OCR_CheckFiles()
    Dim ws As Worksheet, groups As Object, files As Collection
    Dim folder As String, t0 As Double
    If Not ToolTryBegin("CHECK TEN FILE") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    mStop = False: ATT_ClearCheckOutput ws
    If Not ATT_ValidateMaterialDocs(ws) Then GoTo SafeExit
    Set groups = ATT_BuildGroups(ws)
    If groups.Count = 0 Then Err.Raise vbObjectError + 7001, , "Khong co dong hop le tu dong " & FIRST_ROW & "."
    folder = ATT_ResolveSourceFolder()
    If folder = "" Then GoTo SafeExit
    Set files = New Collection
    ATT_ListFiles folder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
    If files.Count = 0 Then Err.Raise vbObjectError + 7003, , "Thu muc khong co PDF/JPG/JPEG/PNG/TIF/TIFF."
    ATT_ListFilesNoOcr files
    ATT_MatchFileNamesToInput groups, files
    ATT_WriteGroups ws, groups
    ATT_Log "CHECK", "DONE_FILENAME", "", "Da doi chieu ten " & files.Count & " file voi " & groups.Count & " nhom INPUT."
SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0)
    ToolEnd
    Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "CHECK", "FAIL", "", CStr(Err.Number) & " - " & Err.Description
    MsgBox "CHECK TEN FILE loi " & Err.Number & ": " & Err.Description, vbCritical, "TOOL BBBG/VAT"
    Resume SafeExit
End Sub

Public Sub CheckFiles_NoOCR()
    OCR_CheckFiles
End Sub

Private Sub ATT_ListFilesNoOcr(ByVal files As Collection)
    Dim wsS As Worksheet, p As Variant, r As Long, typ As String
    Set wsS = ThisWorkbook.Worksheets(SH_SCAN)
    wsS.rows("2:" & wsS.rows.Count).ClearContents
    r = 2
    For Each p In files
        typ = ATT_Classify(ATT_FileName(CStr(p)), "")
        wsS.Cells(r, 1).value = r - 1
        wsS.Cells(r, 2).value = ATT_FileName(CStr(p))
        wsS.Cells(r, 3).value = CStr(p)
        wsS.Cells(r, 4).value = LCase$(ATT_Ext(CStr(p)))
        wsS.Cells(r, 5).value = typ
        wsS.Cells(r, 6).value = "FILENAME_ONLY"
        wsS.Cells(r, 9).value = ATT_TicketFromFileName(ATT_FileName(CStr(p)))
        wsS.Cells(r, 10).value = FileLen(CStr(p))
        wsS.Cells(r, 11).value = FileDateTime(CStr(p))
        r = r + 1
    Next p
End Sub

Private Sub ATT_MatchFileNamesToInput(ByVal groups As Object, ByVal files As Collection)
    Dim k As Variant, g As Object, p As Variant, inputTicket As String, fileTicket As String
    Dim b As Collection, v As Collection, u As Collection, typ As String, pathList As String
    For Each k In groups.keys
        Set g = groups(k): Set b = New Collection: Set v = New Collection: Set u = New Collection
        inputTicket = ATT_Clean(CStr(g("TICKET")))
        g("HEAD_STATUS") = "CHUA_KIEM_TRA_MIGO"
        If inputTicket = "" Then
            g("STATUS") = "NEED_SOPHIEU": g("NOTE") = "Can nhap So phieu trong INPUT de doi chieu ten file."
            GoTo ContinueGroupFilename
        End If
        Application.StatusBar = "Doi chieu ten file: " & inputTicket
        For Each p In files
            fileTicket = ATT_TicketFromFileName(ATT_FileName(CStr(p)))
            If ATT_TicketMatches(fileTicket, inputTicket) Or ATT_FileTicketMatch(ATT_FileName(CStr(p)), inputTicket) Then
                typ = ATT_Classify(ATT_FileName(CStr(p)), "")
                Select Case typ
                    Case "VAT": v.Add CStr(p)
                    Case "BBBG": b.Add CStr(p)
                    Case Else: u.Add CStr(p)
                End Select
            End If
        Next p
        If u.Count > 0 Then
            g("STATUS") = "UNKNOWN_FILE_TYPE"
            g("NOTE") = u.Count & " file dung So phieu nhung chua ro BBBG/VAT: " & inputTicket
            GoTo ContinueGroupFilename
        End If
        pathList = ""
        For Each p In b
            If pathList <> "" Then pathList = pathList & "|"
            pathList = pathList & CStr(p)
        Next p
        g("BBBG") = pathList
        If v.Count > 0 Then g("VAT") = CStr(v(1)) Else g("VAT") = ""
        If b.Count = 0 And v.Count = 0 Then
            g("STATUS") = "NO_BBBG_VAT_MATCH": g("NOTE") = "Khong co BBBG/VAT co So phieu khop INPUT: " & inputTicket
        ElseIf b.Count = 0 Then
            g("STATUS") = "PARTIAL_NO_BBBG": g("NOTE") = "Thieu BBBG; se ATT VAT neu HeadText MIGO dung."
        ElseIf v.Count = 0 Then
            g("STATUS") = "PARTIAL_NO_VAT": g("NOTE") = b.Count & " BBBG khop INPUT; thieu VAT, se ATT BBBG neu HeadText MIGO dung."
        ElseIf v.Count > 1 Then
            g("STATUS") = "PARTIAL_MULTI_VAT": g("NOTE") = b.Count & " BBBG khop; co nhieu VAT, bo qua VAT de tranh nham."
        Else
            g("STATUS") = "CHECK_OK": g("NOTE") = b.Count & " BBBG + 1 VAT khop INPUT."
        End If
        g("OCR_TICKET") = inputTicket
ContinueGroupFilename:
        DoEvents
    Next k
End Sub

Private Function ATT_TicketFromFileName(ByVal fileName As String) As String
    Dim re As Object, ms As Object, s As String
    s = UCase$(ATT_BaseName(fileName))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*[-_. ]*[ ]*([0-9]{2,24})\b"
    If re.Test(s) Then
        Set ms = re.Execute(s): ATT_TicketFromFileName = "IN - " & CStr(ms(0).SubMatches(0)): Exit Function
    End If
    re.pattern = "(^|[^0-9])([0-9]{2,24})([^0-9]|$)"
    If re.Test(s) Then
        Set ms = re.Execute(s): ATT_TicketFromFileName = "IN - " & CStr(ms(0).SubMatches(1))
    End If
End Function

Public Sub SAP_FillHeadTextOnly()
    Dim ws As Worksheet, docs As Object, ses As Object, k As Variant, d As Object
    Dim t0 As Double, errText As String, currentText As String, n As Long
    If Not ToolTryBegin("DIEN HEADTEXT") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Set docs = ATT_BuildDocTickets(ws)
    If docs.Count = 0 Then Err.Raise vbObjectError + 7010, , "Khong co DOC + So phieu de xu ly."
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7011, , "Khong tim thay SAP session. Hay dang nhap SAP truoc."
    For Each k In docs.keys
        Set d = docs(k)
        If d("TICKETS").Count = 0 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "NEED_SOPHIEU", "Can nhap So phieu tu MB51."
        ElseIf d("TICKETS").Count > 1 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_MULTI_SOPHIEU", "DOC co nhieu So phieu; khong tu ghi de HeadText."
        Else
            Application.StatusBar = "HeadText DOC " & CStr(k)
            errText = "": currentText = ""
            If ATT_OpenMigoDoc(ses, CStr(k), errText) Then
                If ATT_ReadHeadText(ses, currentText, errText) Then
                    If Trim$(currentText) = "" Then
                        If ATT_WriteHeadText(ses, ATT_FirstKey(d("TICKETS")), errText) Then
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_DONE", ""
                            n = n + 1
                        Else
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_FAIL", errText
                        End If
                    ElseIf ATT_SameToken(currentText, ATT_FirstKey(d("TICKETS"))) Then
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_OK", "Da co dung So phieu."
                    Else
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_EXISTS_DIFFERENT", "SAP=" & currentText
                    End If
                Else
                    ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_READ_FAIL", errText
                End If
            Else
                ATT_SetDocStatus ws, CStr(k), C_HEAD, "OPEN_DOC_FAIL", errText
            End If
        End If
        DoEvents
    Next k
      MsgBox "Da xu ly HeadText cho " & n & " DOC.", vbInformation, "DIEN HEADTEXT"
SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0): ToolEnd: Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "HEADTEXT", "FAIL", "", CStr(Err.Number) & " - " & Err.Description
      MsgBox "DIEN HEADTEXT loi " & Err.Number & ": " & Err.Description, vbCritical
    Resume SafeExit
End Sub

Private Function ATT_WasPreviouslyAttached(ByVal docNo As String, ByVal attachmentType As String, ByVal filePath As String) As Boolean
    Dim ws As Worksheet, lr As Long, r As Long
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    lr = ws.Cells(ws.rows.Count, 2).End(xlUp).Row
    For r = lr To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 2).value)), Trim$(docNo), vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ws.Cells(r, 4).value)), Trim$(attachmentType), vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ws.Cells(r, 5).value)), Trim$(filePath), vbTextCompare) = 0 And _
           UCase$(Trim$(CStr(ws.Cells(r, 6).value))) = "DONE" Then
            ATT_WasPreviouslyAttached = True
            Exit Function
        End If
    Next r
EH:
End Function

' ======================================================================
'  SAP_AttachCheckedFiles  -  V1.7 NANG CAP
'  Them: SAP Security preflight, GOS smart retry, logging stage day du
'        Source folder khong khoa cung, kiem tra duong dan mang + UNC
' ======================================================================
Public Sub SAP_AttachCheckedFiles()
    Dim ws As Worksheet, docs As Object, groups As Object, files As Collection, ses As Object, docKey As Variant, docInfo As Object
    Dim titles As Object, grpKey As Variant, g As Object, errText As String
    Dim expected As Object, titleB As String, titleV As String, doneDocs As Long, t0 As Double, docAttachFailed As Boolean
    Dim preflightPassed As Boolean

    If Not ToolTryBegin("ATT SAP") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)

    ' --- 1. Validate input rows ---
    If Not ATT_ValidateMaterialDocs(ws) Then GoTo SafeExit

    ' --- 2. Ket noi SAP session ---
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7021, , "Khong tim thay SAP session."
    ATT_ActivateMigoSession ses

    ' --- 3. Chon & kiem tra thu muc nguon (khong khoa cung duong dan) ---
    Dim srcFolder As String
    srcFolder = ATT_ResolveSourceFolder()
    If srcFolder = "" Then GoTo SafeExit
    mSourceFolder = srcFolder

    ' --- 4. Quet file ---
    Set files = New Collection
    ATT_ListFiles mSourceFolder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
    If files.Count = 0 Then Err.Raise vbObjectError + 7022, , "Thu muc khong co file de ATT: " & mSourceFolder

    ' --- 5. Doi chieu nhom BBBG/VAT ---
    Set groups = ATT_BuildGroups(ws)
    If groups.Count = 0 Then Err.Raise vbObjectError + 7023, , "Khong co DOC + So phieu hop le de ATT."
    ATT_ListFilesNoOcr files
    ATT_MatchFileNamesToInput groups, files
    ATT_WriteGroups ws, groups
    Set docs = ATT_BuildAttachDocs(ws)
    If docs.Count = 0 Then Err.Raise vbObjectError + 7020, , "Khong co nhom HeadText + file hop le de ATT."

    ' --- 6. SAP Security Preflight (chay truoc vong lap chinh) ---
    ATT_Log "GOS_PREFLIGHT", "START", "", "Kiem tra SAP Security truoc khi ATT hang loat."
    preflightPassed = ATT_GosSecurityPreflight(ses, errText)
    If Not preflightPassed Then
        ATT_Log "GOS_PREFLIGHT", "FAIL", "", errText
        MsgBox "Khong the mo GOS truoc khi chay hang loat." & vbCrLf & _
               "Chi tiet: " & errText & vbCrLf & vbCrLf & _
               "Hay kiem tra:" & vbCrLf & _
               "  1. SAP GUI Security - da Allow + Remember?" & vbCrLf & _
               "  2. MIGO dang o man hinh nhap DOC?" & vbCrLf & _
               "  3. Session SAP khong bi busy?", vbCritical, "GOS PREFLIGHT FAIL"
        GoTo SafeExit
    End If
    ATT_Log "GOS_PREFLIGHT", "PASS", "", "SAP Security da xac nhan. Bat dau ATT hang loat."

    ' --- 7. Vong lap chinh: ATT tung DOC ---
    For Each docKey In docs.keys
        Set docInfo = docs(docKey)
        Application.StatusBar = "ATT DOC " & CStr(docKey)
        errText = ""
        docAttachFailed = False

        ' Mo MIGO doc
        If Not ATT_OpenMigoDoc(ses, CStr(docKey), errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "MIGO_DOC_OPEN_FAIL", errText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "OPEN_DOC_FAIL", errText
            ATT_Log "ATT", "OPEN_DOC_FAIL", CStr(docKey), errText
            Exit For
        End If

        ' Doc va doi chieu HeadText
        Dim liveHeadText As String, expectedHeadText As String
        liveHeadText = "": expectedHeadText = ""
        For Each grpKey In docInfo("GROUPS").keys
            Set g = docInfo("GROUPS")(grpKey): expectedHeadText = CStr(g("TICKET")): Exit For
        Next grpKey
        If Not ATT_ReadHeadText(ses, liveHeadText, errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_READ_FAIL", errText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "HEADTEXT_READ_FAIL", errText
            ATT_Log "ATT", "HEADTEXT_READ_FAIL", CStr(docKey), errText
            Exit For
        End If
        If Not ATT_TicketMatches(liveHeadText, expectedHeadText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_MISMATCH", "SAP=" & liveHeadText & "; INPUT=" & expectedHeadText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "HEADTEXT_CHANGED", "SAP=" & liveHeadText & "; CHECK=" & expectedHeadText
            ATT_Log "ATT", "HEADTEXT_MISMATCH", CStr(docKey), "SAP=" & liveHeadText & "; INPUT=" & expectedHeadText
            Exit For
        End If

        ' Doc danh sach attachment hien co
        Set titles = ATT_GetAttachmentTitles(ses, errText)
        ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_OK", ""
        If titles Is Nothing Then
            ATT_Log "ATT", "LIST_UNREADABLE_CONTINUE", CStr(docKey), errText
            Set titles = CreateObject("Scripting.Dictionary")
        End If

        ' ATT tung file BBBG va VAT
        Set expected = CreateObject("Scripting.Dictionary")
        For Each grpKey In docInfo("GROUPS").keys
            Set g = docInfo("GROUPS")(grpKey)
            titleB = ATT_FileName(CStr(g("BBBG")))
            titleV = ATT_FileName(CStr(g("VAT")))
            If Trim$(CStr(g("BBBG"))) <> "" Then expected(ATT_NormTitle(titleB)) = titleB
            If Trim$(CStr(g("VAT"))) <> "" Then expected(ATT_NormTitle(titleV)) = titleV

            ' -- ATT BBBG --
            Dim bParts As Variant, bIdx As Long, bPath As String, bAllOk As Boolean
            If Trim$(CStr(g("BBBG"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
            ElseIf ATT_TitleExists(titles, titleB) Or ATT_WasPreviouslyAttached(CStr(docKey), "BBBG", CStr(g("BBBG"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ALREADY_ATTACHED"
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "BBBG: " & titleB
            Else
                bParts = Split(CStr(g("BBBG")), "|"): bAllOk = True
                For bIdx = LBound(bParts) To UBound(bParts)
                    bPath = Trim$(CStr(bParts(bIdx)))
                    If bPath <> "" Then
                        If Not ATT_AttachOneSmartGOS(ses, bPath, titleB, errText) Then
                            bAllOk = False: ATT_AddNoteRows ws, g("ROWS"), "BBBG: " & errText
                            ATT_Log "ATT", "ATT_FAIL_BBBG", CStr(docKey), bPath & " | " & errText
                        Else
                            ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "BBBG: " & titleB
                        End If
                    End If
                Next bIdx
                If bAllOk Then
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_SENT"
                Else
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_PARTIAL_FAIL"
                    docAttachFailed = True
                    Exit For
                End If
            End If

            ' -- ATT VAT --
            If Trim$(CStr(g("VAT"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "NOT_FOUND"
            ElseIf ATT_TitleExists(titles, titleV) Or ATT_WasPreviouslyAttached(CStr(docKey), "VAT", CStr(g("VAT"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ALREADY_ATTACHED"
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "VAT: " & titleV
            ElseIf ATT_AttachOneSmartGOS(ses, CStr(g("VAT")), titleV, errText) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_SENT"
                ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "VAT: " & titleV
            Else
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_FAIL"
                ATT_AddNoteRows ws, g("ROWS"), "VAT: " & errText
                ATT_Log "ATT", "ATT_FAIL_VAT", CStr(docKey), CStr(g("VAT")) & " | " & errText
                docAttachFailed = True
                Exit For
            End If
        Next grpKey

        If docAttachFailed Then Exit For

        ' Ket qua DOC
        If ATT_DocHasMissingVat(docInfo) Then
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_PARTIAL_NO_VAT", "Da ATT BBBG; chua co file VAT."
        Else
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_DONE", ""
            doneDocs = doneDocs + 1
        End If
        ATT_MarkVerified ws, docInfo

        ' Quay lai man hinh nhap DOC
        If Not ATT_ReturnToMigoEntry(ses, errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_DONE", "ATT da thanh cong, nhung chua ve man hinh nhap DOC: " & errText
            ATT_Log "ATT", "RETURN_TO_DOC_FAILED", CStr(docKey), errText
            Exit For
        End If
        ATT_Log "ATT", "RESULT_READY", CStr(docKey), "ATT xong DOC " & CStr(docKey)
NextDoc:
        DoEvents
    Next docKey

SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0): ToolEnd: Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "ATT", "FAIL", "", CStr(Err.Number) & " - " & Err.Description
    Resume SafeExit
End Sub

' ======================================================================
'  ATT_ResolveSourceFolder  -  V1.7 MO'I
'  Lay thu muc tu mSourceFolder hoac CONFIG!B4 (key SOURCE_FOLDER)
'  Neu khong hop le, bat user chon lai
'  Khong co fallback ve thu muc cua user cu the
' ======================================================================
Private Function ATT_ResolveSourceFolder() As String
    Dim folder As String, errMsg As String
    folder = Trim$(mSourceFolder)
    If folder = "" Then folder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))

    If folder <> "" Then
        If Dir$(folder, vbDirectory) = "" Then
            ATT_Log "FOLDER", "SOURCE_FOLDER_NOT_FOUND", "", folder
            errMsg = "Thu muc cau hinh khong ton tai hoac khong truy cap duoc:" & vbCrLf & folder & vbCrLf & vbCrLf & "Vui long chon lai thu muc BBBG/VAT."
            MsgBox errMsg, vbExclamation, "THU MUC KHONG TIM THAY"
            folder = ""
        End If
    End If

    If folder = "" Then
        ATT_Log "FOLDER", "SOURCE_FOLDER_EMPTY", "", "Chua co thu muc - bat user chon"
        folder = ATT_PickFolder()
        If folder = "" Then
            ATT_Log "FOLDER", "SOURCE_FOLDER_EMPTY", "", "User huy chon thu muc"
            ATT_ResolveSourceFolder = ""
            Exit Function
        End If
        ' Kiem tra thu muc duoc chon
        If Dir$(folder, vbDirectory) = "" Then
            ATT_Log "FOLDER", "SOURCE_FOLDER_NOT_FOUND", "", folder
            MsgBox "Thu muc vua chon khong truy cap duoc:" & vbCrLf & folder, vbCritical, "LOI THU MUC"
            ATT_ResolveSourceFolder = ""
            Exit Function
        End If
    End If

    ' Luu lai thu muc da chon
    mSourceFolder = folder
    ATT_SetCfg "SOURCE_FOLDER", folder
    ATT_ResolveSourceFolder = folder
End Function

' ======================================================================
'  ATT_GosSecurityPreflight  -  V1.7 MO'I
'  Mo GOS tren chung tu dau tien, phat hien SAP Security, cho user Allow
'  Tra ve True neu preflight PASS (GOS mo duoc, Security da xu ly)
' ======================================================================
Private Function ATT_GosSecurityPreflight(ByVal ses As Object, ByRef errText As String) As Boolean
    Dim shell As Object, w As Object, startedAt As Double, secHandled As Boolean
    On Error GoTo EH

    ' Dam bao MIGO san sang va khong busy
    If Not ATT_WaitForSapIdle(ses, TIMEOUT_DOC_READY, errText) Then
        errText = "SAP_NOT_IDLE: " & errText
        ATT_Log "GOS_PREFLIGHT", "SAP_NOT_IDLE", "", errText
        Exit Function
    End If

    Set shell = Nothing
    On Error Resume Next
    Set shell = ses.findById("wnd[0]/titl/shellcont/shell")
    On Error GoTo EH
    If shell Is Nothing Then
        errText = "GOS_BUTTON_NOT_FOUND: Shell control khong tim thay"
        ATT_Log "GOS_PREFLIGHT", "GOS_BUTTON_NOT_FOUND", "", errText
        Exit Function
    End If

    ' Thu mo GOS menu (Method 1 - Technical ID)
    On Error Resume Next
    shell.pressContextButton "%GOS_TOOLBOX"
    ATT_Wait 500
    DoEvents

    ' Kiem tra SAP Security xuat hien
    startedAt = Timer
    If modWinFileDialog.WinIsSapSecurityVisible() Then
        ATT_Log "GOS_PREFLIGHT", "SAP_SECURITY_DETECTED", "", "Cua so SAP GUI Security xuat hien"
        modWinFileDialog.WinBringSecurityDialogToFront
        ATT_Log "GOS_PREFLIGHT", "SAP_SECURITY_WAIT", "", "Cho user chon Allow + Remember My Decision"

        MsgBox "SAP GUI Security dang hoi quyen truy cap GOS." & vbCrLf & vbCrLf & _
               "Vui long:" & vbCrLf & _
               "  1. Cua so SAP Security da hien len phia truoc" & vbCrLf & _
               "  2. Chon [Allow]" & vbCrLf & _
               "  3. Tich [Remember my decision]" & vbCrLf & _
               "  4. Sau do nhan OK trong thong bao nay de tiep tuc", _
               vbInformation + vbOKOnly, "SAP GUI SECURITY - CAN XAC NHAN"

        ' Kiem tra lai xem Security da dong chua
        If modWinFileDialog.WinIsSapSecurityVisible() Then
            ATT_Log "GOS_PREFLIGHT", "SAP_SECURITY_TIMEOUT", "", "Cua so Security van con mo sau khi user bam OK"
        Else
            secHandled = True
            ATT_Log "GOS_PREFLIGHT", "SAP_SECURITY_HANDLED", "", "User da xu ly SAP Security thanh cong"
        End If
        ATT_Wait 800
    End If
    On Error GoTo EH

    ' Thu mo Create Attachment (mo ra la preflight PASS)
    Dim w2 As Object
    On Error Resume Next
    shell.selectContextMenuItem "%GOS_PCATTA_CREA"
    ATT_Wait 600
    DoEvents
    Set w2 = ses.findById("wnd[1]")
    On Error GoTo EH

    If Not w2 Is Nothing Then
        ' GOS mo duoc - dong popup thu nghiem
        On Error Resume Next
        ses.findById("wnd[1]").sendVKey 12  ' ESC dong popup thu
        ATT_Wait 300
        On Error GoTo EH
        ATT_GosSecurityPreflight = True
        Exit Function
    End If

    ' GOS chua mo duoc - thu retry Method 2
    ATT_Log "GOS_PREFLIGHT", "TRY_FAIL", "", "Method 1 khong mo duoc GOS dialog, thu Method 2"
    If ATT_GosOpenWithRetry(ses, errText) Then
        ' Dong popup thu nghiem
        On Error Resume Next
        ses.findById("wnd[1]").sendVKey 12
        ATT_Wait 300
        On Error GoTo EH
        ATT_GosSecurityPreflight = True
    Else
        ATT_Log "GOS_PREFLIGHT", "GOS_OPEN_FAILED_AFTER_RETRY", "", errText
    End If
    Exit Function
EH:
    errText = "GOS_PREFLIGHT: " & Err.Description
    ATT_Log "GOS_PREFLIGHT", "FAIL", "", errText
End Function

' ======================================================================
'  ATT_GosOpenWithRetry  -  V1.7 MO'I
'  3-step smart GOS retry voi delay tang dan
'  Tra ve True neu mo duoc wnd[1] (GOS dialog)
' ======================================================================
Private Function ATT_GosOpenWithRetry(ByVal ses As Object, ByRef errText As String) As Boolean
    Dim shell As Object, w As Object, attempt As Long, delay As Long
    Dim methods(1 To 3) As String
    methods(1) = "TECHNICAL_ID"
    methods(2) = "REFRESH_CONTROL"
    methods(3) = "MENU_FALLBACK"

    For attempt = 1 To GOS_MAX_RETRY
        delay = 300 * (2 ^ (attempt - 1))  ' 300ms, 600ms, 1200ms
        ATT_Log "GOS", "TRY_" & attempt, "", "Method: " & methods(attempt) & " | delay=" & delay & "ms"

        ' Dam bao SAP khong busy
        If Not ATT_WaitForSapIdle(ses, TIMEOUT_GOS_READY, errText) Then
            ATT_Log "GOS", "SAP_NOT_IDLE", "", "Attempt " & attempt & ": " & errText
            ATT_Wait delay
            GoTo NextAttempt
        End If

        Set shell = Nothing
        On Error Resume Next
        Set shell = ses.findById("wnd[0]/titl/shellcont/shell")
        On Error GoTo 0

        If shell Is Nothing Then
            errText = "GOS_BUTTON_NOT_FOUND"
            ATT_Log "GOS", "TRY_FAIL", "", "Attempt " & attempt & ": Shell control khong tim thay"
            ATT_Wait delay
            GoTo NextAttempt
        End If

        ' Thuc hien theo method tuong ung
        On Error Resume Next
        Select Case attempt
            Case 1
                ' Method 1: Technical ID truc tiep
                shell.pressContextButton "%GOS_TOOLBOX"
                ATT_Wait delay
                shell.selectContextMenuItem "%GOS_PCATTA_CREA"
            Case 2
                ' Method 2: Re-fetch shell, dua SAP len foreground, goi lai
                ses.findById("wnd[0]").setFocus
                ATT_Wait 200
                Set shell = ses.findById("wnd[0]/titl/shellcont/shell")
                shell.pressContextButton "%GOS_TOOLBOX"
                ATT_Wait delay
                shell.selectContextMenuItem "%GOS_PCATTA_CREA"
            Case 3
                ' Method 3: Menu fallback - tim nut GOS khac nhau
                Dim gosBtn As Object
                Set gosBtn = ATT_FindGosButton(ses)
                If Not gosBtn Is Nothing Then
                    gosBtn.press
                    ATT_Wait delay
                    ' Thu tim menu item sau khi bam nut GOS
                    Dim menuItem As Object
                    Set menuItem = ATT_FindContextMenuAttach(ses)
                    If Not menuItem Is Nothing Then menuItem.Select
                Else
                    ' Fallback cuoi: sendVKey / keyboard shortcut
                    ses.findById("wnd[0]").sendVKey 40
                    ATT_Wait delay
                End If
        End Select
        ATT_Wait 300
        DoEvents
        On Error GoTo 0

        ' Kiem tra GOS dialog da xuat hien
        Set w = Nothing
        On Error Resume Next
        Set w = ses.findById("wnd[1]")
        On Error GoTo 0
        If Not w Is Nothing Then
            ATT_Log "GOS", "ACTION_SENT", "", "Attempt " & attempt & " thanh cong - GOS dialog xuat hien"
            ATT_GosOpenWithRetry = True
            Exit Function
        End If

        ATT_Log "GOS", "TRY_FAIL", "", "Attempt " & attempt & ": GOS dialog chua xuat hien sau " & methods(attempt)
NextAttempt:
        ATT_Wait delay
    Next attempt

    errText = "GOS_OPEN_FAILED_AFTER_RETRY: khong mo duoc GOS sau " & GOS_MAX_RETRY & " lan thu voi 3 methods"
    ATT_Log "GOS", "GOS_OPEN_FAILED_AFTER_RETRY", "", errText
End Function

' ======================================================================
'  ATT_FindGosButton  -  V1.7 MO'I  (helper cho Method 3)
'  Tim toolbar button GOS theo nhieu cach
' ======================================================================
Private Function ATT_FindGosButton(ByVal ses As Object) As Object
    On Error Resume Next
    Set ATT_FindGosButton = ses.findById("wnd[0]/titl/shellcont/shell")
    If ATT_FindGosButton Is Nothing Then
        Set ATT_FindGosButton = ATT_FindChildByIdContains(ses.findById("wnd[0]", False), "GOS_TOOLBOX")
    End If
    On Error GoTo 0
End Function

' ======================================================================
'  ATT_FindContextMenuAttach  -  V1.7 MO'I  (helper cho Method 3)
'  Tim menu item Create Attachment trong context menu hien tai
' ======================================================================
Private Function ATT_FindContextMenuAttach(ByVal ses As Object) As Object
    On Error Resume Next
    Set ATT_FindContextMenuAttach = ses.findById("wnd[1]")
    On Error GoTo 0
End Function

' ======================================================================
'  ATT_AttachOneSmartGOS  -  V1.7 NANG CAP tu ATT_AttachOne
'  Goi GOS theo 3-step retry, nhan dien popup SAP + Windows native
'  Ho tro tieng Viet va duong dan UNC
' ======================================================================
Private Function ATT_AttachOneSmartGOS(ByVal ses As Object, ByVal src As String, ByVal title As String, ByRef errText As String) As Boolean
    Dim w As Object, statusBefore As String, t0 As Double
    On Error GoTo EH

    ' Kiem tra file ton tai
    If Len(Dir$(src)) = 0 Then
        errText = "FILE_NOT_FOUND: " & src
        ATT_Log "ATT", "FILE_NOT_FOUND", "", src
        Exit Function
    End If

    statusBefore = CStr(ses.findById("wnd[0]/sbar").messageType) & "|" & CStr(ses.findById("wnd[0]/sbar").text)

    ' Dam bao SAP san sang
    If Not ATT_WaitForSapIdle(ses, TIMEOUT_DOC_READY, errText) Then
        errText = "SAP_NOT_IDLE: " & errText
        ATT_Log "ATT", "SAP_NOT_IDLE", "", errText
        Exit Function
    End If

    ' Mo GOS voi Smart Retry
    If Not ATT_GosOpenWithRetry(ses, errText) Then
        ATT_Log "ATT", "GOS_NO_DIALOG_POSSIBLE_SECURITY_RULE_OR_NOT_READY", "", errText
        Exit Function
    End If

    ATT_Log "ATT", "ACTION_SENT", "", "GOS dialog mo thanh cong, dien file: " & ATT_FileName(src)

    ' Lay wnd[1]
    On Error Resume Next
    Set w = ses.findById("wnd[1]")
    On Error GoTo EH

    ' Dien file vao popup
    Dim filled As Boolean
    filled = False
    If Not w Is Nothing Then
        ' Thu dien vao SAP file dialog (DY_PATH / DY_FILENAME hoac RLGRAP-FILENAME)
        filled = ATT_SetSapFileDialogSmart(w, src)
    End If

    ' Neu SAP dialog khong duoc, thu Windows native dialog (Unicode)
    If Not filled Then
        ATT_Log "ATT", "WINDOWS_DIALOG_FOUND", "", "Thu Windows native dialog cho: " & src
        filled = modWinFileDialog.WinImportFile(src, TIMEOUT_ATT_DIALOG \ 1000)
    End If

    If Not filled Then
        errText = "ATT_DIALOG_NOT_FOUND: Khong dien duoc file path vao bat ky dialog nao"
        ATT_Log "ATT", "ATT_DIALOG_NOT_FOUND", "", src
        Exit Function
    End If

    ' Cho SAP xac nhan ATT
    If ATT_WaitForAttachment(ses, ATT_FileName(src), statusBefore, errText) Then
        ATT_AttachOneSmartGOS = True
        ATT_Log "ATT", "RESULT_READY", "", "Da ATT: " & ATT_FileName(src)
    End If
    Exit Function
EH:
    errText = "ATT_FILE: " & Err.Description
    ATT_Log "ATT", "ATT_FAIL", "", src & " | " & errText
End Function

' ======================================================================
'  ATT_SetSapFileDialogSmart  -  V1.7 NANG CAP tu ATT_SetSapFileDialog
'  Nhan dien popup SAP 2 dong (DY_PATH + DY_FILENAME) va 1 dong (RLGRAP)
'  Log stage ro rang
' ======================================================================
Private Function ATT_SetSapFileDialogSmart(ByVal w As Object, ByVal fullPath As String) As Boolean
    Dim fso As Object, p As String, n As String, o As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = fso.GetParentFolderName(fullPath)
    n = fso.GetFileName(fullPath)
    On Error Resume Next

    ' Case A: Popup SAP 2 dong (DY_PATH + DY_FILENAME) - uu tien cao nhat
    Set o = Nothing
    Set o = w.findById("usr/ctxtDY_PATH")
    If Not o Is Nothing Then
        o.text = p
        Set o = Nothing
        Set o = w.findById("usr/ctxtDY_FILENAME")
        If Not o Is Nothing Then
            o.text = n
            w.findById("tbar[0]/btn[0]").press
            If Err.Number = 0 Then
                ATT_SetSapFileDialogSmart = True
                ATT_Log "ATT", "SAP_DIALOG_DY_FOUND", "", "Dien DY_PATH=" & p & " | DY_FILENAME=" & n
            End If
            Err.Clear
        End If
        GoTo Cleanup
    End If

    ' Case B: Popup SAP 1 dong (RLGRAP-FILENAME)
    Err.Clear
    Set o = Nothing
    Set o = w.findById("usr/ctxtRLGRAP-FILENAME")
    If Not o Is Nothing Then
        o.text = fullPath
        w.findById("tbar[0]/btn[0]").press
        If Err.Number = 0 Then
            ATT_SetSapFileDialogSmart = True
            ATT_Log "ATT", "SAP_DIALOG_RLGRAP_FOUND", "", "Dien RLGRAP-FILENAME=" & fullPath
        End If
        Err.Clear
    End If

Cleanup:
    On Error GoTo 0
End Function

' ======================================================================
'  CHECK / GROUPING (giu nguyen tu V1.6.4)
' ======================================================================
Private Function ATT_BuildGroups(ByVal ws As Worksheet) As Object
    Dim groups As Object, g As Object, r As Long, lr As Long, key As String
    Dim mat As String, doc As String, po As String, ticket As String
    Set groups = CreateObject("Scripting.Dictionary")
    lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        mat = ATT_Clean(ws.Cells(r, C_MAT).value)
        doc = ATT_Clean(ws.Cells(r, C_DOC).value)
        po = ATT_Clean(ws.Cells(r, C_PO).value)
        ticket = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If mat & doc & po & ticket <> "" Then
            If doc = "" Or po = "" Or mat = "" Then
                ws.Cells(r, C_CHECK).value = "INPUT_MISSING"
                ws.Cells(r, C_NOTE).value = "Can Material + DOC + PO/SA."
            Else
                key = UCase$(doc & "|" & ticket)
                If Not groups.Exists(key) Then
                    Set g = CreateObject("Scripting.Dictionary")
                    g("DOC") = doc: g("TICKET") = ticket
                    Set g("PO") = CreateObject("Scripting.Dictionary")
                    Set g("MAT") = CreateObject("Scripting.Dictionary")
                    Set g("ROWS") = New Collection
                    g("BBBG") = "": g("VAT") = "": g("STATUS") = "": g("NOTE") = ""
                    groups.Add key, g
                End If
                Set g = groups(key)
                g("PO")(UCase$(po)) = True
                g("MAT")(ATT_NormCode(mat)) = True
                g("ROWS").Add r
            End If
        End If
    Next r
    Set ATT_BuildGroups = groups
End Function

Private Function ATT_MarkGroupsNeedTicket(ByVal groups As Object) As Long
    Dim k As Variant, g As Object
    For Each k In groups.keys
        Set g = groups(k)
        If Trim$(CStr(g("TICKET"))) = "" Then
            g("STATUS") = "NEED_SOPHIEU"
            g("NOTE") = "Nhap So phieu o cot D theo dang NCC IN - 4288 roi CHECK lai."
            ATT_MarkGroupsNeedTicket = ATT_MarkGroupsNeedTicket + 1
        End If
    Next k
End Function

Private Function ATT_ScanAndOcr(ByVal files As Collection, ByVal groups As Object) As Collection
    Dim infos As New Collection, p As Variant, info As Object, raw As String, stat As String
    Dim wsS As Worksheet, wsR As Worksheet, rows As Long, rowR As Long
    Set wsS = ThisWorkbook.Worksheets(SH_SCAN): Set wsR = ThisWorkbook.Worksheets(SH_RAW)
    wsS.rows("2:" & wsS.rows.Count).ClearContents: wsR.rows("2:" & wsR.rows.Count).ClearContents
    rows = 2: rowR = 2
    For Each p In files
        If mStop Then Exit For
        Application.StatusBar = "OCR " & (rows - 1) & "/" & files.Count & ": " & ATT_FileName(CStr(p))
        stat = "": raw = ATT_ReadDocument(CStr(p), stat)
        Set info = CreateObject("Scripting.Dictionary")
        info("PATH") = CStr(p): info("NAME") = ATT_FileName(CStr(p))
        info("RAW") = raw: info("OCR_STATUS") = stat
        info("TYPE") = ATT_Classify(info("NAME"), raw)
        Set info("PO") = ATT_ExtractPO(raw, groups)
        Set info("MAT") = ATT_ExtractMaterials(raw)
        info("TICKET") = ATT_ExtractTicket(raw)
        infos.Add info
        wsS.Cells(rows, 1).value = rows - 1
        wsS.Cells(rows, 2).value = info("NAME")
        wsS.Cells(rows, 3).value = info("PATH")
        wsS.Cells(rows, 4).value = LCase$(ATT_Ext(CStr(p)))
        wsS.Cells(rows, 5).value = info("TYPE")
        wsS.Cells(rows, 6).value = stat
        wsS.Cells(rows, 7).value = ATT_JoinKeys(info("PO"), ";")
        wsS.Cells(rows, 8).value = ATT_JoinKeys(info("MAT"), ";")
        wsS.Cells(rows, 9).value = info("TICKET")
        wsS.Cells(rows, 10).value = FileLen(CStr(p))
        wsS.Cells(rows, 11).value = FileDateTime(CStr(p))
        wsR.Cells(rowR, 1).value = rows - 1
        wsR.Cells(rowR, 2).value = info("PATH")
        wsR.Cells(rowR, 3).value = raw
        rows = rows + 1: rowR = rowR + 1
        DoEvents
    Next p
    Set ATT_ScanAndOcr = infos
End Function

Private Sub ATT_WriteGroups(ByVal ws As Worksheet, ByVal groups As Object)
    Dim k As Variant, g As Object, r As Variant
    For Each k In groups.keys
        Set g = groups(k)
        For Each r In g("ROWS")
            ws.Cells(r, C_BBBG).value = g("BBBG")
            ws.Cells(r, C_VAT).value = g("VAT")
            If g.Exists("OCR_PO") Then ws.Cells(r, C_OCR_PO).value = g("OCR_PO")
            If g.Exists("OCR_MAT_B") Then ws.Cells(r, C_OCR_MAT_BBBG).value = g("OCR_MAT_B")
            If g.Exists("OCR_TICKET") Then ws.Cells(r, C_OCR_TICKET).value = g("OCR_TICKET")
            If g.Exists("OCR_MAT_V") Then ws.Cells(r, C_OCR_MAT_VAT).value = g("OCR_MAT_V")
            If g.Exists("HEAD_STATUS") Then ws.Cells(r, C_HEAD).value = g("HEAD_STATUS")
            ws.Cells(r, C_CHECK).value = g("STATUS")
            ws.Cells(r, C_NOTE).value = g("NOTE")
        Next r
    Next k
End Sub

' ======================================================================
'  OCR / FILE (giu nguyen tu V1.6.4)
' ======================================================================
Private Function ATT_ReadDocument(ByVal path As String, ByRef status As String) As String
    Dim ext As String, tmp As String, txtExe As String, ps1 As String, f As String, out As String
    ext = LCase$(ATT_Ext(path)): status = ""
    If ext <> "pdf" Then
        ATT_ReadDocument = ATT_ReadPreparedImage(path, status): Exit Function
    End If
    tmp = ATT_TempFolder("PDF")
    txtExe = ATT_Cfg("PDFTOTEXT_EXE", "")
    If txtExe <> "" And Dir$(txtExe) <> "" Then
        f = tmp & "\pdf.txt"
        If ATT_RunWait(ATT_Q(txtExe) & " -layout " & ATT_Q(path) & " " & ATT_Q(f)) = 0 Then
            out = ATT_ReadText(f)
            If Len(Trim$(out)) >= CLng(Val(ATT_Cfg("PDF_TEXT_MIN_CHARS", "30"))) Then
                status = "PDF_TEXT": ATT_ReadDocument = out: Exit Function
            End If
        End If
    End If
    ps1 = ThisWorkbook.path & "\helpers\RenderPdf.ps1"
    If Dir$(ps1) = "" Then status = "PDF_RENDERER_MISSING": Exit Function
    If ATT_RunWait("powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & ATT_Q(ps1) & _
                   " -PdfPath " & ATT_Q(path) & " -OutputDir " & ATT_Q(tmp) & _
                   " -MaxPages " & ATT_Cfg("PDF_MAX_PAGES", "5")) <> 0 Then
        status = "PDF_RENDER_FAIL": Exit Function
    End If
    f = Dir$(tmp & "\page_*.png")
    Do While f <> ""
        out = out & vbCrLf & ATT_RunTesseract(tmp & "\" & f, status)
        f = Dir$()
    Loop
    If Len(Trim$(out)) = 0 Then status = "PDF_OCR_EMPTY" Else status = "PDF_OCR"
    ATT_ReadDocument = out
End Function

Private Function ATT_OcrPreflight(ByRef msg As String) As Boolean
    Dim p As String
    p = ATT_TesseractPath()
    If p = "" Then
        msg = "Khong tim thay tesseract.exe. Cap nhat CONFIG/TESSERACT_EXE.": Exit Function
    End If
    ATT_OcrPreflight = True
End Function

Private Function ATT_TesseractPath() As String
    Dim p As String
    p = ThisWorkbook.path & "\runtime\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = ATT_Cfg("TESSERACT_EXE", "")
    If p <> "" And Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = Environ$("LOCALAPPDATA") & "\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = Environ$("ProgramFiles") & "\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p
End Function

Private Function ATT_Classify(ByVal fileName As String, ByVal raw As String) As String
    Dim s As String
    s = ATT_NormText(fileName & " " & raw)
    If ATT_HasKeyword(fileName, ATT_Cfg("VAT_KEYWORDS", "VAT,HOADON,HOA DON,INVOICE")) Or _
       InStr(s, "HOA DON GIA TRI GIA TANG") > 0 Or InStr(s, "VAT INVOICE") > 0 Then
        ATT_Classify = "VAT"
    ElseIf ATT_HasKeyword(fileName, ATT_Cfg("BBBG_KEYWORDS", "BBBG,BBGH,BIENBAN,BIEN BAN,BAN GIAO")) Then
        ATT_Classify = "BBBG"
    Else
        ATT_Classify = "UNKNOWN"
    End If
End Function

Private Function ATT_ExtractTicket(ByVal raw As String) As String
    Dim re As Object, ms As Object, s As String
    s = UCase$(raw)
    Set re = CreateObject("VBScript.RegExp"): re.Global = False: re.IgnoreCase = True
    re.pattern = "(SO[ ]*PHIEU|PHIEU[ ]*SO|BBBG[ ]*(NO|NUMBER)?)[ ]*[:#-]?[ ]*([A-Z0-9_-]{2,24})"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        ATT_ExtractTicket = CStr(ms(0).SubMatches(ms(0).SubMatches.Count - 1))
    End If
End Function

Private Function ATT_ExtractMaterials(ByVal raw As String) As Object
    Dim d As Object, lines As Variant, x As Variant, compact As String, m As Object, i As Long, v As String
    Set d = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(raw, vbCr, vbLf), vbLf)
    Set m = CreateObject("VBScript.RegExp"): m.Global = True: m.IgnoreCase = True
    m.pattern = "\bB[I1]W[0-9A-Z]{6,24}(AA|AB)\b"
    For Each x In lines
        compact = UCase$(CStr(x))
        compact = Replace(compact, " ", ""): compact = Replace(compact, "-", ""): compact = Replace(compact, "_", "")
        If m.Test(compact) Then
            For i = 0 To m.Execute(compact).Count - 1
                v = UCase$(m.Execute(compact)(i).value)
                v = Replace(v, "B1W", "BIW")
                d(v) = True
            Next i
        End If
    Next x
    Set ATT_ExtractMaterials = d
End Function

Private Function ATT_RegexDict(ByVal text As String, ByVal pattern As String) As Object
    Dim d As Object, re As Object, ms As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.IgnoreCase = True: re.pattern = pattern
    If re.Test(text) Then
        Set ms = re.Execute(text)
        For i = 0 To ms.Count - 1: d(UCase$(ms(i).value)) = True: Next i
    End If
    Set ATT_RegexDict = d
End Function

Private Sub ATT_ListFiles(ByVal folder As String, ByVal recursive As Boolean, ByRef out As Collection)
    Dim fso As Object, fld As Object, fil As Object, subf As Object, ext As String
    Set fso = CreateObject("Scripting.FileSystemObject"): Set fld = fso.GetFolder(folder)
    For Each fil In fld.files
        ext = LCase$(fso.GetExtensionName(fil.path))
        Select Case ext
            Case "pdf", "jpg", "jpeg", "png", "tif", "tiff"
                out.Add fil.path
        End Select
    Next fil
    If recursive Then For Each subf In fld.SubFolders: ATT_ListFiles subf.path, True, out: Next subf
End Sub

Private Sub ATT_ActivateMigoSession(ByVal ses As Object)
    On Error Resume Next
    ses.findById("wnd[0]").maximize
    AppActivate "MIGO"
    On Error GoTo 0
End Sub

Private Function ATT_ReturnToMigoEntry(ByVal ses As Object, ByRef errText As String) As Boolean
    Dim o As Object, fieldId As String
    On Error GoTo EH
    ses.findById("wnd[0]").sendVKey 12
    If Not ATT_WaitForSapIdle(ses, TIMEOUT_RETURN_TO_DOC, errText) Then Exit Function
    ATT_Wait 500
    On Error Resume Next
    Set o = ses.findById("wnd[1]", False)
    On Error GoTo EH
    If Not o Is Nothing Then
        errText = "MIGO_EXIT_CONFIRMATION_REQUIRED"
        Exit Function
    End If
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")
    Set o = ATT_FindMigoDocField(ses, fieldId)
    If o Is Nothing Then
        errText = "MIGO_DOC_FIELD_NOT_READY"
        Exit Function
    End If
    ATT_ReturnToMigoEntry = True
    Exit Function
EH:
    errText = "RETURN_TO_MIGO_ENTRY: " & Err.Description
End Function

Private Function ATT_WaitForSapIdle(ByVal ses As Object, ByVal timeoutMs As Long, ByRef errText As String) As Boolean
    Dim startedAt As Double
    On Error GoTo EH
    startedAt = Timer
    Do While ses.Busy
        ATT_Wait 100
        If ATT_Elapsed(startedAt) * 1000# >= timeoutMs Then
            errText = "SAP_UI_BUSY_TIMEOUT"
            Exit Function
        End If
    Loop
    ATT_WaitForSapIdle = True
    Exit Function
EH:
    errText = "SAP_UI_IDLE_CHECK: " & Err.Description
End Function

Private Function ATT_FindChildByIdContains(ByVal obj As Object, ByVal needle As String) As Object
    Dim child As Object, found As Object
    On Error Resume Next
    If obj Is Nothing Then Exit Function
    If InStr(1, CStr(obj.id), needle, vbTextCompare) > 0 Then
        Set ATT_FindChildByIdContains = obj
        Exit Function
    End If
    For Each child In obj.Children
        Set found = ATT_FindChildByIdContains(child, needle)
        If Not found Is Nothing Then
            Set ATT_FindChildByIdContains = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

Private Function ATT_FindChildByNameExact(ByVal obj As Object, ByVal targetName As String) As Object
    Dim child As Object, found As Object
    On Error Resume Next
    If obj Is Nothing Then Exit Function
    If StrComp(Trim$(CStr(obj.Name)), targetName, vbTextCompare) = 0 Then
        Set ATT_FindChildByNameExact = obj
        Exit Function
    End If
    For Each child In obj.Children
        Set found = ATT_FindChildByNameExact(child, targetName)
        If Not found Is Nothing Then
            Set ATT_FindChildByNameExact = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

Private Function ATT_FindMigoDocField(ByVal ses As Object, ByVal fieldId As String) As Object
    Dim o As Object, root As Object
    On Error Resume Next
    Set o = ses.findById(fieldId, False)
    If o Is Nothing Then
        Set root = ses.findById("wnd[0]", False)
        Set o = ATT_FindChildByNameExact(root, "GODYNPRO-MAT_DOC")
        If o Is Nothing Then Set o = ATT_FindChildByIdContains(root, "GODYNPRO-MAT_DOC")
    End If
    On Error GoTo 0
    Set ATT_FindMigoDocField = o
End Function

Private Function ATT_WaitMigoDocLoaded(ByVal ses As Object, ByVal docNo As String, ByVal timeoutMs As Long, ByRef errText As String) As Boolean
    Dim startedAt As Double, caption As String, sType As String, sText As String
    Dim fieldId As String, docField As Object, headField As Object, root As Object
    Dim fieldText As String, headFound As Boolean
    On Error GoTo EH
    startedAt = Timer
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")
    Do
        If Not CBool(ses.Busy) Then
            caption = CStr(ses.findById("wnd[0]").text)
            sType = CStr(ses.findById("wnd[0]/sbar").messageType)
            sText = CStr(ses.findById("wnd[0]/sbar").text)
            If sType = "E" Or sType = "A" Or sType = "X" Then
                errText = "MIGO_LOAD_ERROR: " & sText
                Exit Function
            End If
            If InStr(1, ATT_NormCode(caption), ATT_NormCode(docNo), vbTextCompare) > 0 Then
                ATT_WaitMigoDocLoaded = True
                Exit Function
            End If
            Set docField = ATT_FindMigoDocField(ses, fieldId)
            fieldText = ""
            If Not docField Is Nothing Then fieldText = CStr(docField.text)
            Set root = ses.findById("wnd[0]", False)
            Set headField = ATT_FindChildByNameExact(root, "GOHEAD-BKTXT")
            headFound = Not headField Is Nothing
            If ATT_SameToken(fieldText, docNo) And headFound Then
                ATT_WaitMigoDocLoaded = True
                Exit Function
            End If
        End If
        If ATT_Elapsed(startedAt) * 1000# >= timeoutMs Then
            errText = "MIGO_DOC_LOAD_TIMEOUT: expected=" & docNo & "; field=" & fieldText & "; title=" & caption & "; sbar=" & sText
            Exit Function
        End If
        ATT_Wait 100
    Loop
EH:
    errText = "MIGO_DOC_LOAD_CHECK: " & Err.Description
End Function

Private Function ATT_OpenMigoDoc(ByVal ses As Object, ByVal docNo As String, ByRef errText As String) As Boolean
    Dim o As Object, sType As String, sText As String, fieldId As String, caption As String
    Dim requiredKey As String, currentKey As String
    On Error GoTo EH
    ses.findById("wnd[0]").maximize
    caption = CStr(ses.findById("wnd[0]").text)
    If InStr(1, caption, "Material Document", vbTextCompare) > 0 Then
        ses.findById("wnd[0]").sendVKey 12
        If Not ATT_WaitForSapIdle(ses, TIMEOUT_DOC_READY, errText) Then Exit Function
        ATT_Wait 500
        On Error Resume Next: Set o = ses.findById("wnd[1]", False): On Error GoTo EH
        If Not o Is Nothing Then
            errText = "MIGO_EXIT_CONFIRMATION_REQUIRED: Dong thay doi chua luu truoc khi mo DOC " & docNo
            Exit Function
        End If
    End If
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")
    On Error Resume Next
    Set o = ses.findById("wnd[0]").findByName("GODYNPRO-ACTION", "GuiComboBox")
    On Error GoTo EH
    requiredKey = ATT_Cfg("MIGO_ACTION_KEY", "A12")
    If Not o Is Nothing Then
        currentKey = CStr(o.key)
        If currentKey <> requiredKey Then
            o.key = requiredKey
            If Not ATT_WaitForSapIdle(ses, TIMEOUT_DOC_READY, errText) Then Exit Function
        End If
    End If
    Set o = Nothing
    On Error Resume Next
    Set o = ses.findById("wnd[0]").findByName("GODYNPRO-REFDOC", "GuiComboBox")
    On Error GoTo EH
    requiredKey = ATT_Cfg("MIGO_REFDOC_KEY", "R02")
    If Not o Is Nothing Then
        currentKey = CStr(o.key)
        If currentKey <> requiredKey Then
            o.key = requiredKey
            If Not ATT_WaitForSapIdle(ses, TIMEOUT_DOC_READY, errText) Then Exit Function
        End If
    End If
    Set o = ATT_FindMigoDocField(ses, fieldId)
    If o Is Nothing Then Err.Raise vbObjectError + 7025, , "MIGO_DOC_FIELD_NOT_FOUND: " & fieldId
    o.SetFocus
    o.text = ""
    ATT_Wait 100
    o.text = CStr(docNo)
    ATT_Wait 150
    If Not ATT_SameToken(CStr(o.text), CStr(docNo)) Then
        o.text = CStr(docNo)
        ATT_Wait 150
        If Not ATT_SameToken(CStr(o.text), CStr(docNo)) Then
            errText = "MIGO_DOC_INPUT_MISMATCH: expected=" & CStr(docNo) & "; actual=" & CStr(o.text)
            Exit Function
        End If
    End If
    On Error Resume Next: o.caretPosition = Len(CStr(docNo)): On Error GoTo EH
    ATT_Wait 150
    ses.findById("wnd[0]").sendVKey 0
    If Not ATT_WaitMigoDocLoaded(ses, docNo, TIMEOUT_DOC_READY * 1000, errText) Then Exit Function
    ATT_OpenMigoDoc = True: Exit Function
EH:
    errText = "OPEN_MIGO_DOC: " & Err.Description
End Function

Private Function ATT_MigoHasLoadedDocument(ByVal ses As Object) As Boolean
    Dim o As Object, v As String
    On Error Resume Next
    Set o = ses.findById("wnd[0]/usr", False).findByName("GOHEAD-BKTXT", "GuiTextField")
    If Not o Is Nothing Then
        v = Trim$(CStr(o.text))
        ATT_MigoHasLoadedDocument = (v <> "")
    End If
    On Error GoTo 0
End Function

Private Function ATT_ReadHeadText(ByVal ses As Object, ByRef value As String, ByRef errText As String) As Boolean
    Dim o As Object
    On Error GoTo EH
    On Error Resume Next: ses.SendCommand "MIGO_OK_HEADER_OPEN": Err.Clear: On Error GoTo EH
    Set o = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
    value = CStr(o.text): ATT_ReadHeadText = True: Exit Function
EH:
    errText = "READ_HEADTEXT: " & Err.Description
End Function

Private Function ATT_WriteHeadText(ByVal ses As Object, ByVal value As String, ByRef errText As String) As Boolean
    Dim o As Object, t As String
    On Error GoTo EH
    Set o = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
    o.text = value: ses.findById("wnd[0]").sendVKey 11: ATT_Wait 400
    t = CStr(ses.findById("wnd[0]/sbar").messageType)
    If t = "E" Or t = "A" Or t = "X" Then errText = CStr(ses.findById("wnd[0]/sbar").text): Exit Function
    ATT_WriteHeadText = True: Exit Function
EH:
    errText = "WRITE_HEADTEXT: " & Err.Description
End Function

Private Function ATT_FindAttachmentGrid(ByVal root As Object) As Object
    Dim child As Object, found As Object, typeName As String, subTypeName As String
    On Error Resume Next
    If root Is Nothing Then Exit Function
    typeName = CStr(root.Type)
    subTypeName = CStr(root.SubType)
    If typeName = "GuiGridView" Or (typeName = "GuiShell" And InStr(1, subTypeName, "Grid", vbTextCompare) > 0) Then
        Set ATT_FindAttachmentGrid = root
        Exit Function
    End If
    For Each child In root.Children
        Set found = ATT_FindAttachmentGrid(child)
        If Not found Is Nothing Then
            Set ATT_FindAttachmentGrid = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

Private Function ATT_GetAttachmentTitles(ByVal ses As Object, ByRef errText As String) As Object
    Dim d As Object, shell As Object, grid As Object, tbl As Object, w As Object, r As Long, title As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo EH
    Set shell = ses.findById("wnd[0]/titl/shellcont/shell")
    shell.pressContextButton "%GOS_TOOLBOX": shell.selectContextMenuItem "%GOS_VIEW_ATTA": ATT_Wait 350
    Set w = ses.findById("wnd[1]")
    Set grid = ATT_FindAttachmentGrid(w)
    If Not grid Is Nothing Then
        For r = 0 To CLng(grid.rowCount) - 1
            title = ATT_GridTitle(grid, r)
            If title <> "" Then d(ATT_NormTitle(title)) = title
        Next r
    Else
        Set tbl = ATT_FindType(w, "GuiTableControl")
        If tbl Is Nothing Then
            errText = "Khong tim thay bang Attachment list."
        Else
            For r = 0 To CLng(tbl.rowCount) - 1
                title = ATT_TableTitle(tbl, r)
                If title <> "" Then d(ATT_NormTitle(title)) = title
            Next r
        End If
    End If
    On Error Resume Next: w.Close: On Error GoTo 0
    Set ATT_GetAttachmentTitles = d
    Exit Function
EH:
    errText = "ATT_LIST: " & Err.Description
    On Error Resume Next: ses.findById("wnd[1]").Close: On Error GoTo 0
    Set ATT_GetAttachmentTitles = Nothing
End Function

Private Function ATT_StatusIsAttachmentSuccess(ByVal messageType As String, ByVal messageText As String) As Boolean
    Dim t As String, hasObject As Boolean, hasSuccess As Boolean
    If UCase$(Trim$(messageType)) <> "S" Then Exit Function
    t = UCase$(Trim$(messageText))
    hasObject = (InStr(1, t, "ATTACH", vbTextCompare) > 0 Or InStr(1, t, "DOCUMENT", vbTextCompare) > 0 Or InStr(1, t, "FILE", vbTextCompare) > 0)
    hasSuccess = (InStr(1, t, "CREATED", vbTextCompare) > 0 Or InStr(1, t, "SAVED", vbTextCompare) > 0 Or InStr(1, t, "STORED", vbTextCompare) > 0 Or InStr(1, t, "SUCCESS", vbTextCompare) > 0 Or InStr(1, t, "THANH CONG", vbTextCompare) > 0)
    ATT_StatusIsAttachmentSuccess = (hasObject And hasSuccess)
End Function

Private Function ATT_WaitForAttachment(ByVal ses As Object, ByVal fileName As String, ByVal statusBefore As String, ByRef errText As String) As Boolean
    Dim startedAt As Double, sType As String, sText As String, statusNow As String
    On Error GoTo EH
    startedAt = Timer
    ATT_Wait 500
    Do
        If Not ATT_WaitForSapIdle(ses, 10000, errText) Then Exit Function
        sType = CStr(ses.findById("wnd[0]/sbar").messageType)
        sText = CStr(ses.findById("wnd[0]/sbar").text)
        statusNow = sType & "|" & sText
        If (sType = "E" Or sType = "A" Or sType = "X") And statusNow <> statusBefore Then
            errText = "ATT_SAP_ERROR: " & sText
            Exit Function
        End If
        If ATT_StatusIsAttachmentSuccess(sType, sText) Then
            ATT_WaitForAttachment = True
            Exit Function
        End If
        If ATT_Elapsed(startedAt) >= 30# Then
            errText = "ATT_SUCCESS_MESSAGE_TIMEOUT: file=" & fileName & "; status=" & statusNow
            Exit Function
        End If
        ATT_Wait 250
    Loop
EH:
    errText = "ATT_WAIT_CONFIRM: " & Err.Description
End Function

Private Function ATT_FindType(ByVal root As Object, ByVal typeName As String) As Object
    Dim i As Long, c As Object, x As Object
    On Error Resume Next
    If CStr(root.Type) = typeName Then Set ATT_FindType = root: Exit Function
    For i = 0 To root.Children.Count - 1
        Set c = root.Children(i)
        Set x = ATT_FindType(c, typeName)
        If Not x Is Nothing Then Set ATT_FindType = x: Exit Function
    Next i
End Function

Private Function ATT_GridTitle(ByVal grid As Object, ByVal r As Long) As String
    Dim a As Variant, x As Variant, v As String, cols As Variant, header As String
    On Error Resume Next
    cols = grid.ColumnOrder
    For Each x In cols
        header = CStr(grid.GetColumnTitle(CStr(x)))
        If InStr(1, header, "TITLE", vbTextCompare) > 0 Or InStr(1, header, "DESCRIPTION", vbTextCompare) > 0 Then
            v = CStr(grid.GetCellValue(r, CStr(x)))
            If Trim$(v) <> "" Then ATT_GridTitle = v: Exit Function
        End If
    Next x
    On Error GoTo 0
    a = Array("BITM_DESCR", "OBJDES", "TITLE", "DESCRIPT", "FILENAME")
    For Each x In a
        On Error Resume Next: v = CStr(grid.GetCellValue(r, CStr(x))): On Error GoTo 0
        If Trim$(v) <> "" Then ATT_GridTitle = v: Exit Function
    Next x
End Function

' ======================================================================
'  DOC / ATT GROUPS (giu nguyen tu V1.6.4)
' ======================================================================
Private Function ATT_BuildDocTickets(ByVal ws As Worksheet) As Object
    Dim d As Object, x As Object, r As Long, lr As Long, doc As String, t As String
    Set d = CreateObject("Scripting.Dictionary"): lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        doc = ATT_Clean(ws.Cells(r, C_DOC).value): t = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If doc <> "" Then
            If Not d.Exists(doc) Then Set x = CreateObject("Scripting.Dictionary"): Set x("TICKETS") = CreateObject("Scripting.Dictionary"): d.Add doc, x
            Set x = d(doc): If t <> "" Then x("TICKETS")(UCase$(t)) = True
        End If
    Next r
    Set ATT_BuildDocTickets = d
End Function

Private Function ATT_BuildAttachDocs(ByVal ws As Worksheet) As Object
    Dim docs As Object, di As Object, groups As Object, g As Object, r As Long, lr As Long, key As String
    Dim doc As String, t As String, st As String
    Set docs = CreateObject("Scripting.Dictionary"): lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        st = UCase$(ATT_Clean(ws.Cells(r, C_CHECK).value))
        If st = "CHECK_OK" Or Left$(st, 8) = "PARTIAL_" Then
            doc = ATT_Clean(ws.Cells(r, C_DOC).value): t = ATT_Clean(ws.Cells(r, C_TICKET).value): key = UCase$(doc & "|" & t)
            If Not docs.Exists(doc) Then Set di = CreateObject("Scripting.Dictionary"): Set di("GROUPS") = CreateObject("Scripting.Dictionary"): docs.Add doc, di
            Set di = docs(doc): Set groups = di("GROUPS")
            If Not groups.Exists(key) Then
                Set g = CreateObject("Scripting.Dictionary"): g("TICKET") = t
                g("BBBG") = CStr(ws.Cells(r, C_BBBG).value): g("VAT") = CStr(ws.Cells(r, C_VAT).value)
                Set g("ROWS") = New Collection: groups.Add key, g
            End If
            Set g = groups(key): g("ROWS").Add r
        End If
    Next r
    Set ATT_BuildAttachDocs = docs
End Function

Private Function ATT_DocHasMissingVat(ByVal docInfo As Object) As Boolean
    Dim k As Variant, g As Object
    For Each k In docInfo("GROUPS").keys
        Set g = docInfo("GROUPS")(k)
        If Trim$(CStr(g("VAT"))) = "" Then
            ATT_DocHasMissingVat = True
            Exit Function
        End If
    Next k
End Function

Private Sub ATT_MarkVerified(ByVal ws As Worksheet, ByVal docInfo As Object)
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
End Sub

' ======================================================================
'  GENERIC HELPERS (giu nguyen tu V1.6.4)
' ======================================================================
Private Function ATT_Cfg(ByVal key As String, ByVal defaultValue As String) As String
    Dim ws As Worksheet, f As Range
    On Error GoTo x
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then ATT_Cfg = CStr(ws.Cells(f.Row, 2).value) Else ATT_Cfg = defaultValue
    Exit Function
x:  ATT_Cfg = defaultValue
End Function

Private Sub ATT_SetCfg(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet, f As Range, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG): Set f = ws.Columns(1).Find(key, LookAt:=xlWhole)
    If f Is Nothing Then r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1 Else r = f.Row
    ws.Cells(r, 1).value = key: ws.Cells(r, 2).value = value
End Sub

Private Function ATT_BoolCfg(ByVal key As String, ByVal def As Boolean) As Boolean
    Dim s As String: s = UCase$(ATT_Cfg(key, IIf(def, "TRUE", "FALSE")))
    ATT_BoolCfg = (s = "TRUE" Or s = "1" Or s = "YES")
End Function

Private Function ATT_PickFolder() As String
    Dim fd As Object: Set fd = Application.FileDialog(4)
    fd.title = "Chon thu muc chua BBBG/VAT (ho tro Local, Network, UNC)": fd.AllowMultiSelect = False
    If fd.Show = -1 Then ATT_PickFolder = fd.SelectedItems(1)
End Function

Private Function ATT_LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    ATT_LastRow = ws.Cells(ws.rows.Count, col).End(xlUp).Row
    If ATT_LastRow < FIRST_ROW Then ATT_LastRow = FIRST_ROW
End Function

Private Function ATT_Clean(ByVal v As Variant) As String
    ATT_Clean = Trim$(Replace(CStr(v), ChrW$(160), " "))
End Function

Private Function ATT_NormText(ByVal s As String) As String
    s = UCase$(s): s = Replace(s, "_", " "): s = Replace(s, "-", " "): s = Replace(s, ".", " ")
    s = Replace(s, vbTab, " "): s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    ATT_NormText = Trim$(s)
End Function

Private Function ATT_NormCode(ByVal s As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.pattern = "[^A-Za-z0-9]"
    s = UCase$(re.Replace(s, "")): ATT_NormCode = Replace(s, "B1W", "BIW")
End Function

Private Function ATT_TicketCode(ByVal ticket As String) As String
    Dim re As Object, ms As Object, s As String
    s = ATT_NormText(ticket)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*([0-9]{2,24})\b"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        ATT_TicketCode = "IN" & CStr(ms(0).SubMatches(0))
    Else
        ATT_TicketCode = ATT_NormCode(ticket)
    End If
End Function

Private Function ATT_TicketMatches(ByVal a As String, ByVal b As String) As Boolean
    Dim ca As String, cb As String
    ca = ATT_TicketCode(a): cb = ATT_TicketCode(b)
    If ca = "" Or cb = "" Then Exit Function
    ATT_TicketMatches = (ca = cb)
End Function

Private Function ATT_FileTicketMatch(ByVal fileName As String, ByVal ticket As String) As Boolean
    Dim re As Object, ms As Object, s As String, digits As String, compact As String
    s = UCase$(ATT_BaseName(fileName))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*[-_. ]*[ ]*([0-9]{2,24})\b"
    If re.Test(UCase$(ticket)) Then
        Set ms = re.Execute(UCase$(ticket))
        digits = CStr(ms(0).SubMatches(0))
        re.pattern = "(^|[^0-9])IN[ ]*[-_. ]*[ ]*0*" & digits & "([^0-9]|$)"
        If re.Test(s) Then ATT_FileTicketMatch = True: Exit Function
        re.pattern = "(^|[^0-9])0*" & digits & "([^0-9]|$)"
        ATT_FileTicketMatch = re.Test(s)
        Exit Function
    End If
    compact = ATT_NormCode(s)
    ATT_FileTicketMatch = (InStr(1, compact, ATT_TicketCode(ticket), vbTextCompare) > 0)
End Function

Private Function ATT_TextHasToken(ByVal text As String, ByVal ticket As String) As Boolean
    ATT_TextHasToken = ATT_TokenSequence(ATT_NormText(text), ATT_NormText(ticket))
End Function

Private Function ATT_TokenSequence(ByVal hay As String, ByVal needle As String) As Boolean
    Dim h As Variant, n As Variant, i As Long, j As Long, ok As Boolean
    If needle = "" Then Exit Function
    h = Split(hay, " "): n = Split(needle, " ")
    For i = LBound(h) To UBound(h) - (UBound(n) - LBound(n))
        ok = True
        For j = LBound(n) To UBound(n)
            If ATT_NormCode(CStr(h(i + j))) <> ATT_NormCode(CStr(n(j))) Then ok = False: Exit For
        Next j
        If ok Then ATT_TokenSequence = True: Exit Function
    Next i
End Function

Private Function ATT_SameToken(ByVal a As String, ByVal b As String) As Boolean
    ATT_SameToken = (ATT_NormCode(a) = ATT_NormCode(b))
End Function

Private Function ATT_ContainsAll(ByVal have As Object, ByVal need As Object) As Boolean
    Dim k As Variant
    ATT_ContainsAll = True
    For Each k In need.keys
        If Not have.Exists(UCase$(CStr(k))) Then
            ATT_ContainsAll = False
            Exit Function
        End If
    Next k
End Function

Private Function ATT_Overlap(ByVal a As Object, ByVal b As Object) As Long
    Dim k As Variant
    For Each k In b.keys
        If a.Exists(UCase$(CStr(k))) Then ATT_Overlap = ATT_Overlap + 1
    Next k
End Function

Private Function ATT_JoinKeys(ByVal d As Object, ByVal SEP As String) As String
    Dim k As Variant
    For Each k In d.keys
        If ATT_JoinKeys <> "" Then ATT_JoinKeys = ATT_JoinKeys & SEP
        ATT_JoinKeys = ATT_JoinKeys & CStr(k)
    Next k
End Function

Private Function ATT_HasKeyword(ByVal s As String, ByVal csv As String) As Boolean
    Dim x As Variant
    For Each x In Split(csv, ",")
        If InStr(1, UCase$(s), Trim$(UCase$(CStr(x))), vbTextCompare) > 0 Then
            ATT_HasKeyword = True
            Exit Function
        End If
    Next x
End Function

Private Function ATT_Title(ByVal ticket As String, ByVal kind As String) As String
    Dim fmt As String
    If UCase$(kind) = "VAT" Then fmt = ATT_Cfg("VAT_TITLE_FORMAT", "{SOPHIEU} VAT") Else fmt = ATT_Cfg("BBBG_TITLE_FORMAT", "{SOPHIEU} BBBG")
    ATT_Title = Trim$(Replace(fmt, "{SOPHIEU}", ticket, 1, -1, vbTextCompare))
End Function

Private Function ATT_NormTitle(ByVal s As String) As String
    s = ATT_BaseName(s): ATT_NormTitle = ATT_NormCode(s)
End Function

Private Function ATT_TitleExists(ByVal titles As Object, ByVal expected As String) As Boolean
    Dim k As Variant, e As String: e = ATT_NormTitle(expected)
    For Each k In titles.keys: If InStr(1, CStr(k), e, vbTextCompare) > 0 Or InStr(1, e, CStr(k), vbTextCompare) > 0 Then ATT_TitleExists = True: Exit Function
    Next k
End Function

Private Function ATT_AllTitlesExist(ByVal titles As Object, ByVal expected As Object) As Boolean
    Dim k As Variant
    ATT_AllTitlesExist = True
    For Each k In expected.keys
        If Not ATT_TitleExists(titles, expected(k)) Then
            ATT_AllTitlesExist = False
            Exit Function
        End If
    Next k
End Function

Private Function ATT_FirstKey(ByVal d As Object) As String
    Dim a As Variant
    If d Is Nothing Then Exit Function
    If d.Count = 0 Then Exit Function
    a = d.keys
    ATT_FirstKey = CStr(a(0))
End Function

Private Function ATT_FileName(ByVal p As String) As String
    ATT_FileName = CreateObject("Scripting.FileSystemObject").GetFileName(p)
End Function

Private Function ATT_BaseName(ByVal p As String) As String
    ATT_BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(p)
End Function

Private Function ATT_Ext(ByVal p As String) As String
    ATT_Ext = CreateObject("Scripting.FileSystemObject").GetExtensionName(p)
End Function

Private Function ATT_SafeName(ByVal s As String) As String
    Dim x As Variant: For Each x In Array("\", "/", ":", "*", "?", """", "<", ">", "|"): s = Replace(s, CStr(x), "_"): Next x
    ATT_SafeName = s
End Function

Private Function ATT_Q(ByVal s As String) As String
    ATT_Q = Chr$(34) & s & Chr$(34)
End Function

Private Function ATT_RunWait(ByVal cmd As String) As Long
    ATT_RunWait = CreateObject("WScript.Shell").Run(cmd, 0, True)
End Function

Private Function ATT_ReadText(ByVal p As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.LoadFromFile p: ATT_ReadText = st.ReadText: st.Close
End Function

Private Function ATT_TempFolder(ByVal prefix As String) As String
    Dim p As String: p = Environ$("TEMP") & "\" & prefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 100), "00000000")
    MkDir p: ATT_TempFolder = p
End Function

Private Sub ATT_Wait(ByVal ms As Long)
    DoEvents
    ATT_Sleep ms
    DoEvents
End Sub

Private Function ATT_Elapsed(ByVal t0 As Double) As Double
    ATT_Elapsed = Timer - t0
    If ATT_Elapsed < 0 Then ATT_Elapsed = ATT_Elapsed + 86400#
End Function

Private Sub ATT_ClearCheckOutput(ByVal ws As Worksheet)
    Dim lr As Long: lr = ATT_LastRow(ws, C_DOC)
    If lr < FIRST_ROW Then Exit Sub
    ws.Range(ws.Cells(FIRST_ROW, C_BBBG), ws.Cells(lr, C_NOTE)).ClearContents
End Sub

Private Sub ATT_SetRows(ByVal ws As Worksheet, ByVal rows As Collection, ByVal col As Long, ByVal value As String)
    Dim r As Variant
    For Each r In rows
        ws.Cells(r, col).value = value
    Next r
End Sub

Private Sub ATT_AddNoteRows(ByVal ws As Worksheet, ByVal rows As Collection, ByVal msg As String)
    Dim r As Variant
    For Each r In rows
        If ws.Cells(r, C_NOTE).value <> "" Then
            ws.Cells(r, C_NOTE).value = ws.Cells(r, C_NOTE).value & " | "
        End If
        ws.Cells(r, C_NOTE).value = ws.Cells(r, C_NOTE).value & msg
    Next r
End Sub

Private Sub ATT_SetDocStatus(ByVal ws As Worksheet, ByVal doc As String, ByVal col As Long, ByVal status As String, ByVal note As String)
    Dim r As Long, lr As Long
    lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        If ATT_Clean(ws.Cells(r, C_DOC).value) = doc Then
            ws.Cells(r, col).value = status
            If note <> "" Then ATT_AddNoteRows ws, ATT_OneRow(r), note
        End If
    Next r
End Sub

Private Function ATT_OneRow(ByVal r As Long) As Collection
    Dim c As New Collection
    c.Add r
    Set ATT_OneRow = c
End Function

Private Sub ATT_Log(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = stage: ws.Cells(r, 3).value = status: ws.Cells(r, 4).value = doc: ws.Cells(r, 5).value = msg
End Sub

Private Sub ATT_History(ByVal doc As String, ByVal ticket As String, ByVal typ As String, ByVal path As String, ByVal result As String)
    Dim ws As Worksheet, r As Long: Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = doc: ws.Cells(r, 3).value = ticket: ws.Cells(r, 4).value = typ: ws.Cells(r, 5).value = path: ws.Cells(r, 6).value = result
End Sub

' ======================================================================
'  COMPATIBILITY STUBS (giu nguyen tu V1.6.4)
' ======================================================================
Public Function ToolTryBegin(ByVal actionName As String) As Boolean
    If mToolBusy Then MsgBox "Tool dang chay. Hay doi luot hien tai ket thuc.", vbExclamation: Exit Function
    mToolBusy = True: ToolTryBegin = True
End Function

Public Sub ToolEnd()
    mToolBusy = False
End Sub

Public Function ToolRunIsActive() As Boolean
    ToolRunIsActive = mToolBusy
End Function

Public Function IsRunInProgress() As Boolean
    IsRunInProgress = mToolBusy
End Function

Public Sub ResetTransientState()
    mToolBusy = False
    mStop = False
End Sub

Public Sub GetExistingOBState(ByRef readyCount As Long, ByRef postedCount As Long, ByRef staleCount As Long)
    readyCount = 0
    postedCount = 0
    staleCount = 0
End Sub

Public Sub Run_MIGO311_Post_FirstReadyOB()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_AllowYellow()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_AllowYellow()
    ATT_DisabledOldMacro
End Sub
Public Sub Post_MIGO_311_FromOBGroups_V3()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_OneClick()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_UncheckRedOK()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_UncheckRedOK()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_PartialRescue()
    ATT_DisabledOldMacro
End Sub

Private Sub ATT_DisabledOldMacro()
    MsgBox "Macro MIGO 311 cu khong dung trong tool ATT nay.", vbInformation
End Sub

' ======================================================================
'  SelectSourceFolder_NoOCR  -  V1.7 NANG CAP
'  Khong khoa duong dan, luu vao CONFIG!B4, ho tro Local/Network/UNC
' ======================================================================
Public Sub SelectSourceFolder_NoOCR()
    Dim p As String
    p = ATT_PickFolder()
    If p = "" Then Exit Sub

    ' Kiem tra thu muc co truy cap duoc
    If Dir$(p, vbDirectory) = "" Then
        MsgBox "Thu muc vua chon khong truy cap duoc:" & vbCrLf & p, vbCritical, "LOI THU MUC"
        Exit Sub
    End If

    mSourceFolder = p
    ATT_SetCfg "SOURCE_FOLDER", p
    MsgBox "Da chon thu muc BBBG/VAT:" & vbCrLf & p, vbInformation, "BBBG/VAT - V1.7"
End Sub

Private Function ATT_TableTitle(ByVal tbl As Object, ByVal rowIndex As Long) As String
    On Error Resume Next
    ATT_TableTitle = CStr(tbl.GetCell(rowIndex, 1).text)
    If ATT_TableTitle = "" Then ATT_TableTitle = CStr(tbl.GetCell(rowIndex, 0).text)
    On Error GoTo 0
End Function

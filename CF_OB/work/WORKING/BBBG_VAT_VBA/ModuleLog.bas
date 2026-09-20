Option Explicit

Private logRow As Long
Private mLogWs As Worksheet

Private Function LogSheet() As Worksheet
If mLogWs Is Nothing Then Set mLogWs = ThisWorkbook.Sheets("LOG")
If UCase$(Trim$(CStr(mLogWs.Cells(1, 1).value & ""))) <> "THOI GIAN" Then
    mLogWs.Cells(1, 1).value = "THOI GIAN"
    mLogWs.Cells(1, 2).value = "LOAI"
    mLogWs.Cells(1, 3).value = "PRODUCT"
    mLogWs.Cells(1, 4).value = "NOI DUNG"
    mLogWs.Cells(1, 5).value = "THAM CHIEU"
    mLogWs.Cells(1, 6).value = "HUONG DAN"
End If
Set LogSheet = mLogWs
End Function

Public Sub ClearLog()
Dim ws As Worksheet
Dim lastRow As Long
Set ws = LogSheet()

lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
If lastRow < 2 Then lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
If lastRow >= 2 Then ws.Range("A2:F" & CStr(lastRow)).ClearContents

ws.Columns(1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
logRow = 1
End Sub

Public Sub WriteLog(ByVal logType As String, ByVal content As String, _
Optional ByVal refInfo As String = "", _
Optional ByVal productCode As String = "")
Dim ws As Worksheet
Set ws = LogSheet()

If logRow < 1 Then
    logRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If logRow < 1 Then logRow = 1
End If
logRow = logRow + 1
ws.Cells(logRow, 1).value = Now
ws.Cells(logRow, 2).value = logType
ws.Cells(logRow, 3).value = productCode
ws.Cells(logRow, 4).value = content
ws.Cells(logRow, 5).value = refInfo
End Sub

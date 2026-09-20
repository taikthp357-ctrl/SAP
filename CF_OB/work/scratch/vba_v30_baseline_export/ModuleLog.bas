Attribute VB_Name = "ModuleLog"
Option Explicit

Private logRow As Long
Private mLogWs As Worksheet

Private Function LogSheet() As Worksheet
If mLogWs Is Nothing Then Set mLogWs = ThisWorkbook.Sheets("RUN_LOG")
If UCase$(Trim$(CStr(mLogWs.Cells(8, 1).value & ""))) <> "TIME" Then
    mLogWs.Cells(8, 1).value = "Time"
    mLogWs.Cells(8, 2).value = "Run ID"
    mLogWs.Cells(8, 3).value = "Stage"
    mLogWs.Cells(8, 4).value = "OB"
    mLogWs.Cells(8, 5).value = "Document"
    mLogWs.Cells(8, 6).value = "Item"
    mLogWs.Cells(8, 7).value = "WH"
    mLogWs.Cells(8, 8).value = "WT"
    mLogWs.Cells(8, 9).value = "Material"
    mLogWs.Cells(8, 10).value = "Action"
    mLogWs.Cells(8, 11).value = "Status"
    mLogWs.Cells(8, 12).value = "Message"
    mLogWs.Cells(8, 13).value = "SAP Sec"
    mLogWs.Cells(8, 14).value = "Excel Sec"
    mLogWs.Cells(8, 15).value = "Session"
    mLogWs.Cells(8, 16).value = "TCode"
    mLogWs.Cells(8, 17).value = "Title"
End If
Set LogSheet = mLogWs
End Function

Public Sub ClearLog()
Dim ws As Worksheet
Dim lastRow As Long
Set ws = LogSheet()

lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
If lastRow < 2 Then lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).row
If lastRow >= 9 Then ws.Range("A9:Q" & CStr(lastRow)).ClearContents

ws.columns(1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
logRow = 8
End Sub

Public Sub WriteLog(ByVal logType As String, ByVal content As String, _
Optional ByVal refInfo As String = "", _
Optional ByVal productCode As String = "")
Dim ws As Worksheet
Set ws = LogSheet()

If logRow < 9 Then
    logRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If logRow < 8 Then logRow = 8
End If
logRow = logRow + 1
ws.Cells(logRow, 1).value = Now
ws.Cells(logRow, 2).value = ""
ws.Cells(logRow, 3).value = logType
ws.Cells(logRow, 4).value = refInfo
ws.Cells(logRow, 10).value = "MODULE_SAP"
ws.Cells(logRow, 11).value = logType
ws.Cells(logRow, 12).value = content
ws.Cells(logRow, 9).value = productCode
End Sub


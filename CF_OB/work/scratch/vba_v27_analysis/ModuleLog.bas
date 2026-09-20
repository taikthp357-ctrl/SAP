Attribute VB_Name = "ModuleLog"
Option Explicit

Private logRow As Long
Private mLogWs As Worksheet

Private Function LogSheet() As Worksheet
If mLogWs Is Nothing Then Set mLogWs = ThisWorkbook.Sheets("RUN_LOG")
If UCase$(Trim$(CStr(mLogWs.Cells(8, 1).Value & ""))) <> "TIME" Then
    mLogWs.Cells(8, 1).Value = "Time"
    mLogWs.Cells(8, 2).Value = "Run ID"
    mLogWs.Cells(8, 3).Value = "Stage"
    mLogWs.Cells(8, 4).Value = "OB"
    mLogWs.Cells(8, 5).Value = "Document"
    mLogWs.Cells(8, 6).Value = "Item"
    mLogWs.Cells(8, 7).Value = "WH"
    mLogWs.Cells(8, 8).Value = "WT"
    mLogWs.Cells(8, 9).Value = "Material"
    mLogWs.Cells(8, 10).Value = "Action"
    mLogWs.Cells(8, 11).Value = "Status"
    mLogWs.Cells(8, 12).Value = "Message"
    mLogWs.Cells(8, 13).Value = "SAP Sec"
    mLogWs.Cells(8, 14).Value = "Excel Sec"
    mLogWs.Cells(8, 15).Value = "Session"
    mLogWs.Cells(8, 16).Value = "TCode"
    mLogWs.Cells(8, 17).Value = "Title"
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
ws.Cells(logRow, 1).Value = Now
ws.Cells(logRow, 2).Value = ""
ws.Cells(logRow, 3).Value = logType
ws.Cells(logRow, 4).Value = refInfo
ws.Cells(logRow, 10).Value = "MODULE_SAP"
ws.Cells(logRow, 11).Value = logType
ws.Cells(logRow, 12).Value = content
ws.Cells(logRow, 9).Value = productCode
End Sub


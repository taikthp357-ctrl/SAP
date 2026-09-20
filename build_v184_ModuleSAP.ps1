$ErrorActionPreference = "Stop"

$inputFile = "g:\SAP\v183_audit_exported\ModuleSAP.bas"
$outputDir = "g:\SAP\v184_patch"
$outputFile = Join-Path $outputDir "ModuleSAP.bas"

Write-Host "Reading $inputFile"
$content = [System.IO.File]::ReadAllText($inputFile)

$find1 = @"
Private Function SAPX_GetSessionForTcode(ByVal tcode As String) As Object
Dim seed As Object, con As Object, ses As Object, cand As Object, ids As Object
Dim i As Long, n As Long, sid As String, t As Double
Set seed = modSupports.Get_SAPSession()
If seed Is Nothing Then Err.Raise vbObjectError + 2938, , "Khong co SAP session."
Set con = seed.parent
For i = 0 To con.Children.Count - 1
    Set ses = SAPX_CollectionObjectAt(con.Children, i)
    If Not ses Is Nothing Then If SAPX_TcodeIs(ses, tcode) Then Set SAPX_GetSessionForTcode = ses: Exit Function
Next i
Set ids = CreateObject("Scripting.Dictionary"): n = con.Children.Count
"@

$replace1 = @"
Private Function SAPX_GetSessionForTcode(ByVal tcode As String) As Object
Dim seed As Object, app As Object, con As Object, ses As Object, cand As Object, ids As Object
Dim i As Long, j As Long, n As Long, sid As String, t As Double
Set seed = modSupports.Get_SAPSession()
If seed Is Nothing Then Err.Raise vbObjectError + 2938, , "Khong co SAP session."

' User profiles do not keep the same SAP connection/session index.
' Search all connections before creating a new tab.
On Error Resume Next
Set app = seed.parent.parent
On Error GoTo 0
If Not app Is Nothing Then
    For i = 0 To app.Children.Count - 1
        Set con = SAPX_CollectionObjectAt(app.Children, i)
        If Not con Is Nothing Then
            For j = 0 To con.Children.Count - 1
                Set ses = SAPX_CollectionObjectAt(con.Children, j)
                If Not ses Is Nothing Then
                    If SAPX_TcodeIs(ses, tcode) Then
                        If SAPX_WaitIdle(ses, 20) Then
                            Set SAPX_GetSessionForTcode = ses
                            Exit Function
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End If

' No existing tcode: create a new session in the active connection.
Set con = seed.parent
Set ids = CreateObject("Scripting.Dictionary"): n = con.Children.Count
"@

$find2 = @"
SAPX_Navigate cand, tcode
If Not SAPX_TcodeIs(cand, tcode) Then Err.Raise vbObjectError + 2944
"@

$replace2 = @"
SAPX_Navigate cand, tcode
If Not SAPX_WaitIdle(cand, 20) Then Err.Raise vbObjectError + 2943, , "SAP chua san sang sau khi dieu huong."
If Not SAPX_TcodeIs(cand, tcode) Then Err.Raise vbObjectError + 2944
"@

$content = $content.Replace($find1, $replace1)
$content = $content.Replace($find2, $replace2)

if (-not (Test-Path $outputDir)) {
    Write-Host "Creating directory $outputDir"
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Write-Host "Writing $outputFile without UTF-8 BOM"
[System.IO.File]::WriteAllText($outputFile, $content, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Done."

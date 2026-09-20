# Patch Dir$ bug with Unicode Win32 API functions ATT_FolderExistsW and ATT_FileExistsW

# 1. Update modWinFileDialog.bas
$winFile = "g:\SAP\v183_patch\modWinFileDialog.bas"
$winContent = Get-Content -Path $winFile -Raw

$winHelperCode = @'

' ======================================================================
'  Unicode Win32 File/Folder Existence Helpers (No Dir$ ANSI Bug)
' ======================================================================
#If VBA7 Then
    Private Declare PtrSafe Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As LongPtr) As Long
#Else
    Private Declare Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As Long) As Long
#End If

Public Function ATT_FolderExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim$(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then
        ATT_FolderExistsW = ((attr And 16) <> 0)
    End If
End Function

Public Function ATT_FileExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim$(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then
        ATT_FileExistsW = ((attr And 16) = 0)
    End If
End Function
'@

# Clean out any old helper code if appended previously
if ($winContent -match "Unicode Win32 File/Folder Existence Helpers") {
    $winContent = $winContent.Substring(0, $winContent.IndexOf("' ======================================================================`r`n'  Unicode Win32 File/Folder Existence Helpers"))
}

$winContent = $winContent.Replace('If Len(Dir$(fullPath)) = 0 Then', 'If Not ATT_FileExistsW(fullPath) Then')
$winContent = $winContent + "`r`n" + $winHelperCode
[System.IO.File]::WriteAllText($winFile, $winContent, [System.Text.Encoding]::UTF8)
Write-Host "Updated modWinFileDialog.bas with ATT_FolderExistsW and ATT_FileExistsW" -ForegroundColor Green

# 2. Update MIGO311_RunPost_V24_ErrorGuard.bas
$migoFile = "g:\SAP\v183_patch\MIGO311_RunPost_V24_ErrorGuard.bas"
$migoContent = Get-Content -Path $migoFile -Raw

$migoContent = $migoContent.Replace('If Dir$(folder, vbDirectory) = "" Then', 'If Not modWinFileDialog.ATT_FolderExistsW(folder) Then')
$migoContent = $migoContent.Replace('If Dir$(p, vbDirectory) = "" Then', 'If Not modWinFileDialog.ATT_FolderExistsW(p) Then')
$migoContent = $migoContent.Replace('If Len(Dir$(src)) = 0 Then', 'If Not modWinFileDialog.ATT_FileExistsW(src) Then')

[System.IO.File]::WriteAllText($migoFile, $migoContent, [System.Text.Encoding]::UTF8)
Write-Host "Updated MIGO311_RunPost_V24_ErrorGuard.bas replacing Dir$ calls with Unicode Win32 API" -ForegroundColor Green

# 3. Clean BOM
powershell -File g:\SAP\fix_bom.ps1

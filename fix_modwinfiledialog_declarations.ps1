# Restructure modWinFileDialog.bas to place API Declare statements at top General Declarations section

$v17WinFile = "g:\SAP\v17_patch\modWinFileDialog.bas"
$outDir = "g:\SAP\v183_patch"
$winOutPath = Join-Path $outDir "modWinFileDialog.bas"

$winRaw = Get-Content -Path $v17WinFile -Raw

# 1. Add GetFileAttributesW into top #If VBA7 declaration block
$vba7Declarations = @"
    Private Declare PtrSafe Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As LongPtr) As Long
"@

$vba32Declarations = @"
    Private Declare Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As Long) As Long
"@

# Insert WFD_GetFileAttributesW into #If VBA7 block
$winRaw = $winRaw.Replace(
    'Private Declare PtrSafe Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongPtrW" _' + "`r`n" + '        (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr',
    'Private Declare PtrSafe Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongPtrW" _' + "`r`n" + '        (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr' + "`r`n" + $vba7Declarations
)

# Insert WFD_GetFileAttributesW into #Else block
$winRaw = $winRaw.Replace(
    'Private Declare Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" _' + "`r`n" + '        (ByVal hwnd As Long, ByVal nIndex As Long) As Long',
    'Private Declare Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" _' + "`r`n" + '        (ByVal hwnd As Long, ByVal nIndex As Long) As Long' + "`r`n" + $vba32Declarations
)

# 2. Add implementation functions at bottom (Public Functions only, no Declare statements)
$winHelperFunctions = @'

' ======================================================================
'  Unicode Win32 File/Folder Existence Helpers (No Dir$ ANSI Bug)
' ======================================================================
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

# Replace Len(Dir$(fullPath)) = 0 in WinImportFile
$winRaw = $winRaw.Replace('If Len(Dir$(fullPath)) = 0 Then', 'If Not ATT_FileExistsW(fullPath) Then')

# Save updated modWinFileDialog.bas
[System.IO.File]::WriteAllText($winOutPath, ($winRaw + "`r`n" + $winHelperFunctions), [System.Text.Encoding]::UTF8)
Write-Host "Restructured modWinFileDialog.bas declarations to top section successfully!" -ForegroundColor Green

# Strip BOM
powershell -File g:\SAP\fix_bom.ps1

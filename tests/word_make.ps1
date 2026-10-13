# ASCII-only on purpose: a .ps1 containing Chinese must be saved as UTF-8 *with BOM*,
# otherwise Windows PowerShell 5.1 reads it as ANSI and reports nonsense errors.
# Keeping this file pure ASCII sidesteps that trap entirely.
param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Add()
    $word.Selection.InsertFile($In)
    $doc.SaveAs2($Out, 16)          # 16 = wdFormatDocumentDefault (.docx)
    $doc.Close(0)
} finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Output ("OK " + $Out)

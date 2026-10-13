# ASCII-only on purpose (see word_make.ps1 header).
# Reads a .docx through real Microsoft Word and dumps its text as UTF-8.
# If Word cannot parse the file it will raise / repair; a successful round trip whose
# text matches what we intended is the end-to-end proof that the file is intact.
param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($In, $false, $true)   # ConfirmConversions=false, ReadOnly=true
    $text = $doc.Content.Text
    $paras = $doc.Paragraphs.Count
    $doc.Close(0)
} finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
[IO.File]::WriteAllText($Out, $text, [Text.UTF8Encoding]::new($false))
Write-Output ("OK paragraphs=" + $paras)

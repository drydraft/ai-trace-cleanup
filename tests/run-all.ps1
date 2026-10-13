# ASCII-only on purpose (a .ps1 with Chinese must be saved as UTF-8 *with BOM*,
# otherwise Windows PowerShell 5.1 reads it as ANSI and reports nonsense errors).
#
#   powershell -File tests/run-all.ps1
#
# The [Console]::OutputEncoding line is not optional: without it, redirecting a
# Python child that prints UTF-8 through PowerShell's `>` re-encodes it as GBK
# and silently corrupts every Chinese character in the log.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
python (Join-Path $here 'selftest.py')
$self = $LASTEXITCODE
python (Join-Path $here 'word_roundtrip.py')
$word = $LASTEXITCODE
Write-Output ("selftest=" + $self + "  word_roundtrip=" + $word)
exit ([Math]::Max([Math]::Max($self, $word), 0))

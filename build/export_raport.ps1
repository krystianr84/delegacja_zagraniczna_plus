param([string]$Pdf = (Join-Path $PSScriptRoot "raport_test.pdf"), [string]$Set = "", [string]$Sheet = "Raport",
      [string]$File = (Join-Path $PSScriptRoot "..\Kalkulator_delegacji_zagranicznej.xlsx"))
# Eksportuje arkusz (domyślnie „Raport”, -Sheet Delegacja) do PDF (bez zapisywania zmian w .xlsx). -Set jak w verify.ps1.
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($File)
    $k = $wb.Worksheets.Item("Kalkulator")
    if ($Set) { foreach ($kv in $Set.Split(";")) { $a, $v = $kv.Split("=", 2); if ($a -match "^(.+)!(.+)$") { $wb.Worksheets.Item($Matches[1]).Range($Matches[2]).Formula = $v } else { $k.Range($a).Formula = $v } } }
    $x.CalculateFull()
    $r = $wb.Worksheets.Item($Sheet)
    foreach ($c in $r.UsedRange.Cells) { if ($c.Text -match '^#') { "BŁĄD $Sheet!$($c.Address(0,0)) = $($c.Text)" } }
    $r.ExportAsFixedFormat(0, $Pdf)
    "PDF: $Pdf"
} finally { $wb.Close($false); $x.Quit() }

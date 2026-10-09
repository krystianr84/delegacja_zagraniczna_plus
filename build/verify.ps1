param([string]$File = (Join-Path $PSScriptRoot "..\Kalkulator_delegacji_zagranicznej.xlsx"), [string]$Set = "")
# Przelicza plik w Excelu (bez zapisu), szuka błędów i wypisuje wyniki. -Set "K7=...;C15=Niemcy;Kursy NBP!F8=4.25" podmienia dane wejściowe (bez nazwy arkusza = Kalkulator).
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($File)
    $k = $wb.Worksheets.Item("Kalkulator")
    if ($Set) { foreach ($kv in $Set.Split(";")) { $a, $v = $kv.Split("=", 2); if ($a -match "^(.+)!(.+)$") { $wb.Worksheets.Item($Matches[1]).Range($Matches[2]).Formula = $v } else { $k.Range($a).Formula = $v } } }
    $x.CalculateFull()
    foreach ($ws in $wb.Worksheets) {
        foreach ($c in $ws.UsedRange.Cells) {
            $t = $c.Text
            if ($t -match '^#(REF|VALUE|NAME|DIV|N/A|NUM|NULL|SPILL|CALC)') { "BŁĄD $($ws.Name)!$($c.Address(0,0)) = $t  :: $($c.Formula2)" }
        }
    }
    "Kontrola: " + $k.Range("E1").Text
    $L = Get-Content (Join-Path $PSScriptRoot "layout.json") -Raw -Encoding utf8 | ConvertFrom-Json
    "--- Wyliczenie diety (wiersze $($L.R0)-$($L.RS))"
    foreach ($r in $L.R0..$L.RS) { ($k.Range("B${r}:M${r}").Cells | ForEach-Object { $_.Text }) -join " | " }
    "--- Dieta wg walut"
    foreach ($r in ($L.V0 + 2)..$L.VL) { ($k.Range("B${r}:E${r}").Cells | ForEach-Object { $_.Text }) -join " | " }
    "--- Pozostałe koszty"
    foreach ($r in $L.KO0..$L.KT) { ($k.Range("B${r}:F${r}").Cells | ForEach-Object { $_.Text }) -join " | " }
    $n = $wb.Worksheets.Item("Kursy NBP")
    "--- Kursy: C4=" + $n.Range("C4").Text + " C5=" + $n.Range("C5").Text
    foreach ($r in 8..17) { ($n.Range("B${r}:E${r}").Cells | ForEach-Object { $_.Text }) -join " | " }
    "--- Okres " + $n.Range("C21").Text + " - " + $n.Range("C22").Text
    foreach ($r in 25..30) { ($n.Range("B${r}:L${r}").Cells | ForEach-Object { $_.Text }) -join " | " }
} finally { $wb.Close($false); $x.Quit() }

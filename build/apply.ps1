param([string]$File = (Join-Path $PSScriptRoot "..\Kalkulator_delegacji_zagranicznej.xlsx"))
$dir = $PSScriptRoot
$F = Get-Content (Join-Path $dir "formulas.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($File)
    $bad = @()
    foreach ($p in $F.PSObject.Properties) {
        $i = $p.Name.LastIndexOf("!")
        $sh = $p.Name.Substring(0, $i).Trim("'")
        $addr = $p.Name.Substring($i + 1)
        try { $wb.Worksheets.Item($sh).Range($addr).Formula2 = $p.Value }
        catch { $bad += "$($p.Name): $($_.Exception.Message)" }
    }
    $bad
    $x.CalculateFull()
    # ochrona arkuszy bez hasła: edytowalne tylko pola wejściowe (odblokowane w build.py);
    # wolno zmieniać szerokości kolumn, wysokości wierszy i formaty. Odblokowanie: Recenzja → Nie chroń arkusza.
    foreach ($ws in $wb.Worksheets) {
        $ws.Protect("", $true, $true, $true, $false, $true, $true, $true) | Out-Null
        $ws.EnableSelection = 0
    }
    $wb.Worksheets.Item("Kalkulator").Activate()
    $wb.Worksheets.Item("Kalkulator").Range("C5").Select()
    $wb.Save()
    "zapisano, formuł: " + $F.PSObject.Properties.Name.Count + ", błędów wpisu: " + $bad.Count
} finally { if ($wb) { $wb.Close($false) }; $x.Quit() }
if ($bad.Count) { exit 1 }

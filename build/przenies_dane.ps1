param([Parameter(Mandatory)][string]$Z,
      [string]$File = (Join-Path $PSScriptRoot "..\Kalkulator_delegacji_zagranicznej.xlsx"))
# Kopiuje dane wpisane przez użytkownika (pola wejściowe Kalkulatora i kurs ręczny z Kursów NBP) z pliku -Z do świeżo zbudowanego pliku.
# Oba pliki muszą mieć ten sam układ (layout.json z ostatniego build.py). Uruchamiaj po build.py + apply.ps1.
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$L = Get-Content (Join-Path $PSScriptRoot "layout.json") -Raw -Encoding utf8 | ConvertFrom-Json
$E1 = $L.RET - 1
$pola = @($L.PRAC, $L.STAN, $L.CEL, $L.NR, $L.DPOL, $L.MIEJ, $L.SROD, $L.WYJ, $L.POW, $L.DKURS, $L.TRYB, $L.ZAL, $L.TERM, $L.DOW,
          "C$($L.E0):M$E1", "D$($L.RET)", "K$($L.RET):M$($L.RET)", "C$($L.KO0):D$($L.KT - 1)")
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
    $src = $x.Workbooks.Open((Resolve-Path $Z).Path, 0, $true)
    $dst = $x.Workbooks.Open((Resolve-Path $File).Path)
    $s = $src.Worksheets.Item("Kalkulator"); $d = $dst.Worksheets.Item("Kalkulator")
    foreach ($a in $pola) {
        # przez Formula (tekst), bo przypisanie Value2 z innego skoroszytu potrafi rzucić błąd rzutowania
        foreach ($c in $s.Range($a).Cells) { if (-not $c.HasFormula) { $d.Range($c.Address()).Formula = [string]$c.Formula } }
    }
    # kurs ręczny (arkusz Kursy NBP, kolumny F:H) – tylko wiersze z wpisanym kursem
    $ks = $src.Worksheets.Item("Kursy NBP"); $kd = $dst.Worksheets.Item("Kursy NBP")
    foreach ($r in $L.KRS0..$L.KRS1) {
        if ($ks.Range("F$r").Value2 -is [double]) { foreach ($c in "F", "G", "H") { $kd.Range("$c$r").Formula = [string]$ks.Range("$c$r").Formula } }
    }
    $x.CalculateFull()
    $dst.Save()
    "Przeniesiono dane z: $Z"
    "Kontrola: " + $d.Range("E1").Text + " | Ogółem do rozliczenia: " + $d.Range("F$($L.KT)").Text
} finally { $x.Workbooks.Close(); $x.Quit() }

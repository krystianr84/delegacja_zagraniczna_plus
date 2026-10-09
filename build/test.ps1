param([string]$File = (Join-Path $PSScriptRoot "..\wydanie\Kalkulator_delegacji_zagranicznej.xlsx"))
# Testy regresji na pliku z przykładowymi danymi z build.py (domyślnie wydanie/). Plik nie jest zapisywany.
# Każdy przypadek: zmiany danych (-Set jak w verify.ps1), oczekiwany początek komunikatu kontroli E1 i kwoty (tekst jak w polskim Excelu).
# Kod wyjścia 1, gdy którykolwiek przypadek się nie zgadza. Wymaga internetu (kursy NBP).
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$File = [IO.Path]::GetFullPath($File)
$L = Get-Content (Join-Path $PSScriptRoot "layout.json") -Raw -Encoding utf8 | ConvertFrom-Json
$kontrolny = "C15=Węgry;D15=2026-09-28 12:00;E15=0;C16=;D16=;E16=;D20=2026-10-02 16:00;K7=2026-10-08;K8=w dniu daty kursu"
$przypadki = @(
    @{ N = "przykład z build.py (Niemcy + Austria)"; Set = ""; E1 = "OK"; Dieta = "930,46"; Ogolem = "2 570,96" },
    @{ N = "przypadek kontrolny (Węgry)"; Set = $kontrolny; E1 = "OK"; Dieta = "834,92" },
    @{ N = "krótki wyjazd 7,5 h → 1/3 diety"; Set = "C16=;D16=;E16=;E15=1;D20=2026-09-14 18:00"; E1 = "OK"; Dieta = "38,77" },
    @{ N = "pusty etap między krajami"; Set = "C16=;D16=;E16=;C17=Austria;D17=2026-09-16 14:00"; E1 = "Nie zostawiaj pustego etapu" },
    @{ N = "tekst zamiast daty"; Set = "D15=abc"; E1 = "Niepoprawna data" },
    @{ N = "więcej posiłków niż dób"; Set = "E15=9"; E1 = "Więcej posiłków niż dni" },
    @{ N = "powrót przed wjazdem"; Set = "D20=2026-09-15 10:00"; E1 = "Powrót przez granicę musi być po" },
    @{ N = "brak kursu (data z przyszłości)"; Set = "K7=2099-12-31"; E1 = "Brak kursu NBP" }
)
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
$bledy = 0
try {
    foreach ($p in $przypadki) {
        $wb = $x.Workbooks.Open($File, 0, $true)
        try {
            $k = $wb.Worksheets.Item("Kalkulator")
            if ($p.Set) { foreach ($kv in $p.Set.Split(";")) { $a, $v = $kv.Split("=", 2); $k.Range($a).Formula = $v } }
            $x.CalculateFull()
            $wynik = @()
            $e1 = $k.Range("E1").Text
            if (-not $e1.StartsWith($p.E1)) { $wynik += "kontrola: '$e1', oczekiwano '$($p.E1)...'" }
            if ($p.Dieta -and $k.Range("M$($L.RS)").Text -ne $p.Dieta) { $wynik += "dieta: $($k.Range("M$($L.RS)").Text), oczekiwano $($p.Dieta)" }
            if ($p.Ogolem) {
                $og = $k.Range("F$($L.KT)").Text
                if ($og -ne $p.Ogolem) { $wynik += "ogółem: $og, oczekiwano $($p.Ogolem)" }
                $d = $wb.Worksheets.Item("Delegacja")
                $og2 = $d.Range("I" + $d.Range("F1:F200").Find("Ogółem").Row).Text
                if ($og2 -ne $p.Ogolem) { $wynik += "Delegacja ogółem: $og2, oczekiwano $($p.Ogolem)" }
            }
            if ($p.E1 -eq "OK") {   # przy poprawnych danych żaden arkusz nie może mieć błędu #...
                # Value2 całego arkusza naraz (szybko); błąd Excela przychodzi jako Int32 (kod CVErr), liczby jako Double
                foreach ($ws in $wb.Worksheets) {
                    $ur = $ws.UsedRange; $v = $ur.Value2
                    if ($v -isnot [array]) { continue }
                    for ($i = 1; $i -le $v.GetLength(0); $i++) { for ($j = 1; $j -le $v.GetLength(1); $j++) {
                        if ($v[$i, $j] -is [int]) { $wynik += "$($ws.Name)!$($ur.Cells.Item($i, $j).Address(0,0)) = $($ur.Cells.Item($i, $j).Text)" }
                    } }
                }
            }
            if ($wynik) { $bledy++; "BŁĄD  $($p.N)"; $wynik | ForEach-Object { "      $_" } } else { "OK    $($p.N)" }
        } finally { $wb.Close($false) }
    }
} finally { $x.Quit() }
if ($bledy) { "Nie przeszło: $bledy z $($przypadki.Count)"; exit 1 }
"Wszystkie testy przeszły ($($przypadki.Count))."

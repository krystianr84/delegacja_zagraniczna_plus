param([string]$File = (Join-Path $PSScriptRoot "..\wydanie\Kalkulator_delegacji_zagranicznej.xlsx"))
# Testy regresji na pliku z przykładowymi danymi z build.py (domyślnie wydanie/). Plik nie jest zapisywany.
# Każdy przypadek: zmiany danych (-Set jak w verify.ps1), oczekiwany początek komunikatu kontroli E1 i kwoty (tekst jak w polskim Excelu).
# Kod wyjścia 1, gdy którykolwiek przypadek się nie zgadza. Wymaga internetu (kursy NBP).
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$File = [IO.Path]::GetFullPath($File)
$L = Get-Content (Join-Path $PSScriptRoot "layout.json") -Raw -Encoding utf8 | ConvertFrom-Json
$kontrolny = "C15=Węgry;D15=2026-09-28 12:00;E15=0;C16=;D16=;E16=;D20=2026-10-02 16:00;K7=2026-10-08;K8=w dniu daty kursu"
# kwota słownie na druku Delegacji: ogółem = dieta z przykładu (930,46) + inne wydatki, noclegi wyzerowane
$noc = "C$($L.KO0 + 2)"; $inne = "C$($L.KO0 + 4)"
$bezKraju = "C15=;D15=;E15=;C16=;D16=;E16=;$noc="
$przypadki = @(
    @{ N = "przykład z build.py (Niemcy + Austria)"; Set = ""; E1 = "OK"; Dieta = "930,46"; Ogolem = "2 570,96"; Slownie = "dwa tysiące pięćset siedemdziesiąt złotych i 96/100" },
    @{ N = "przypadek kontrolny (Węgry)"; Set = $kontrolny; E1 = "OK"; Dieta = "834,92" },
    @{ N = "krótki wyjazd 7,5 h → 1/3 diety"; Set = "C16=;D16=;E16=;E15=1;D20=2026-09-14 18:00"; E1 = "OK"; Dieta = "38,77" },
    @{ N = "pusty etap między krajami"; Set = "C16=;D16=;E16=;C17=Austria;D17=2026-09-16 14:00"; E1 = "Nie zostawiaj pustego etapu" },
    @{ N = "tekst zamiast daty"; Set = "D15=abc"; E1 = "Niepoprawna data" },
    @{ N = "więcej posiłków niż dób"; Set = "E15=9"; E1 = "Więcej posiłków niż dni" },
    @{ N = "powrót przed wjazdem"; Set = "D20=2026-09-15 10:00"; E1 = "Powrót przez granicę musi być po" },
    @{ N = "brak kursu (data z przyszłości)"; Set = "K7=2099-12-31"; E1 = "Brak kursu NBP" },
    @{ N = "kurs ręczny zamiast NBP (Węgry, 4,4)"; Set = "$kontrolny;K7=2099-12-31;Kursy NBP!F8=4.4;Kursy NBP!G8=2026-10-08;Kursy NBP!H8=196/A/NBP/2026"; E1 = "OK"; Dieta = "838,95" },
    @{ N = "słownie: 932 → złote"; Set = "$noc=;$inne=1.54"; E1 = "OK"; Slownie = "dziewięćset trzydzieści dwa złote i 00/100" },
    @{ N = "słownie: 935 → złotych"; Set = "$noc=;$inne=4.54"; E1 = "OK"; Slownie = "dziewięćset trzydzieści pięć złotych i 00/100" },
    @{ N = "słownie: 1012 → złotych (12–14)"; Set = "$noc=;$inne=81.54"; E1 = "OK"; Slownie = "jeden tysiąc dwanaście złotych i 00/100" },
    @{ N = "słownie: 1022 → złote"; Set = "$noc=;$inne=91.54"; E1 = "OK"; Slownie = "jeden tysiąc dwadzieścia dwa złote i 00/100" },
    @{ N = "słownie: 1 → złoty"; Set = "$bezKraju;$inne=1"; E1 = "Wybierz kraj"; Slownie = "jeden złoty i 00/100" },
    @{ N = "słownie: pusty druk → puste pole"; Set = $bezKraju; E1 = "Wybierz kraj"; Slownie = "" },
    @{ N = "kurs ręczny bez nr tabeli"; Set = "Kursy NBP!F8=4.4;Kursy NBP!G8=2026-10-08"; E1 = "Kurs ręczny: uzupełnij" }
)
$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
$bledy = 0
try {
    foreach ($p in $przypadki) {
        $wb = $x.Workbooks.Open($File, 0, $true)
        try {
            $k = $wb.Worksheets.Item("Kalkulator")
            if ($p.Set) { foreach ($kv in $p.Set.Split(";")) { $a, $v = $kv.Split("=", 2); if ($a -match "^(.+)!(.+)$") { $wb.Worksheets.Item($Matches[1]).Range($Matches[2]).Formula = $v } else { $k.Range($a).Formula = $v } } }
            $x.CalculateFull()
            # NBP chwilowo odrzuca zapytania po kilkuset w krótkim czasie (16 przypadków × ~20 zapytań) – przy nieudanym
            # pobraniu kursu z daty nie z przyszłości odczekaj i przelicz ponownie (do 6 × 10 s)
            $kn = $wb.Worksheets.Item("Kursy NBP")
            for ($proba = 0; $proba -lt 6; $proba++) {
                $c5 = $kn.Range("C5").Value2
                $blad = @($L.KRS0..$L.KRS1 | Where-Object { [string]$kn.Range("N$_").Value2 -eq "BLAD" }).Count
                if (-not $blad -or $c5 -isnot [double] -or $c5 -gt (Get-Date).ToOADate()) { break }
                Start-Sleep -Seconds 10; $x.CalculateFull()
            }
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
            if ($p.ContainsKey("Slownie")) {
                $d = $wb.Worksheets.Item("Delegacja")
                $sl = [string]$d.Range("I" + ($d.Range("F1:F200").Find("Ogółem").Row + 1)).Value2
                if ($sl -ne $p.Slownie) { $wynik += "słownie: '$sl', oczekiwano '$($p.Slownie)'" }
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

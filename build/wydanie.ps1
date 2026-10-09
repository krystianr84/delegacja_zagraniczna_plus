param([string]$File = (Join-Path $PSScriptRoot "..\wydanie\Kalkulator_delegacji_zagranicznej.xlsx"))
# Buduje publiczne pliki do repozytorium bez ruszania roboczego pliku z prawdziwymi danymi:
#  - $File – z przykładowymi danymi z build.py, sprawdzany testami (test.ps1),
#  - <nazwa>_pusty.xlsx – pusty szablon (build.py --pusty), sprawdzany przez verify.ps1 (brak błędów #...).
# Z obu usuwa z metadanych nazwę użytkownika, którą Excel wpisuje przy zapisie (lastModifiedBy).
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
$File = [IO.Path]::GetFullPath($File)
$Pusty = [IO.Path]::ChangeExtension($File, $null).TrimEnd(".") + "_pusty.xlsx"
New-Item -ItemType Directory -Force (Split-Path $File) | Out-Null

function Zbuduj([string]$Plik, [string[]]$Opcje) {
    python (Join-Path $PSScriptRoot "build.py") $Plik @Opcje
    if ($LASTEXITCODE) { throw "build.py zakończył się błędem" }
    & (Join-Path $PSScriptRoot "apply.ps1") -File $Plik
    if ($LASTEXITCODE) { throw "apply.ps1: nie wszystkie formuły udało się wpisać" }
}

function UsunAutora([string]$Plik) {
    Add-Type -AssemblyName System.IO.Compression
    $zip = [IO.Compression.ZipFile]::Open($Plik, "Update")
    try {
        $e = $zip.GetEntry("docProps/core.xml")
        $r = New-Object IO.StreamReader($e.Open()); $xml = $r.ReadToEnd(); $r.Close()
        $xml = $xml -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', ''
        $e.Delete()
        $w = New-Object IO.StreamWriter($zip.CreateEntry("docProps/core.xml").Open(), (New-Object Text.UTF8Encoding $false))
        $w.Write($xml); $w.Close()
    } finally { $zip.Dispose() }
}

# pusty szablon najpierw – plik z przykładem buduje się na końcu, więc formulas.json i layout.json zostają z jego budowy
Zbuduj $Pusty @("--pusty")
$bledy = & (Join-Path $PSScriptRoot "verify.ps1") -File $Pusty | Select-String "^BŁĄD"
if ($bledy) { $bledy; throw "pusty szablon ma błędy formuł – nie publikuj go" }

Zbuduj $File @()
& (Join-Path $PSScriptRoot "test.ps1") -File $File
if ($LASTEXITCODE) { throw "test.ps1: testy nie przeszły – nie publikuj tego pliku" }

UsunAutora $File
UsunAutora $Pusty
"wydanie: $File"
"pusty szablon: $Pusty"

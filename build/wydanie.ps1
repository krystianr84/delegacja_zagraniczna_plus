param([string]$File = (Join-Path $PSScriptRoot "..\wydanie\Kalkulator_delegacji_zagranicznej.xlsx"))
# Buduje publiczny plik do repozytorium (przykładowe dane z build.py) bez ruszania roboczego pliku z prawdziwymi danymi,
# sprawdza go testami (test.ps1) i usuwa z metadanych nazwę użytkownika, którą Excel wpisuje przy zapisie (lastModifiedBy).
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
$File = [IO.Path]::GetFullPath($File)
New-Item -ItemType Directory -Force (Split-Path $File) | Out-Null
python (Join-Path $PSScriptRoot "build.py") $File
if ($LASTEXITCODE) { throw "build.py zakończył się błędem" }
& (Join-Path $PSScriptRoot "apply.ps1") -File $File
if ($LASTEXITCODE) { throw "apply.ps1: nie wszystkie formuły udało się wpisać" }
& (Join-Path $PSScriptRoot "test.ps1") -File $File
if ($LASTEXITCODE) { throw "test.ps1: testy nie przeszły – nie publikuj tego pliku" }
Add-Type -AssemblyName System.IO.Compression
$zip = [IO.Compression.ZipFile]::Open($File, "Update")
try {
    $e = $zip.GetEntry("docProps/core.xml")
    $r = New-Object IO.StreamReader($e.Open()); $xml = $r.ReadToEnd(); $r.Close()
    $xml = $xml -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', ''
    $e.Delete()
    $w = New-Object IO.StreamWriter($zip.CreateEntry("docProps/core.xml").Open(), (New-Object Text.UTF8Encoding $false))
    $w.Write($xml); $w.Close()
} finally { $zip.Dispose() }
"wydanie: $File"

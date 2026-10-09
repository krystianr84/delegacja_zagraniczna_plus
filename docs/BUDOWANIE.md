# Budowanie, testy i wydawanie

Dokument dla osób, które chcą zmieniać kalkulator. Do samego korzystania z niego wystarczy [README](../README.md).

Plik `.xlsx` jest **generowany** – nie edytuj go ręcznie, zmieniaj `build/build.py`. Notatki techniczne i sprawdzone pułapki (ustawienia regionalne, parsowanie kursów NBP, ochrona arkuszy) są w [`CLAUDE.md`](../CLAUDE.md).

## Struktura repozytorium

```
├── wydanie/                 pliki publikowane w Wydaniach: z przykładowymi danymi i pusty szablon (_pusty)
├── build/
│   ├── build.py             układ, formaty, dane statyczne, walidacje → .xlsx + formulas.json + layout.json
│   ├── apply.ps1            Excel przez COM: wpisuje formuły, przelicza, chroni arkusze, zapisuje
│   ├── verify.ps1           przelicza plik, szuka błędów #..., wypisuje wyniki i kursy
│   ├── test.ps1             testy regresji z oczekiwanymi kwotami (m.in. przypadek kontrolny)
│   ├── export_raport.ps1    eksport wybranego arkusza do PDF (podgląd wydruku)
│   ├── przenies_dane.ps1    przenosi wpisane dane ze starego pliku do świeżo zbudowanego
│   ├── wydanie.ps1          buduje publiczne pliki w wydanie/, uruchamia testy, czyści metadane
│   ├── social_preview.py    obrazek podglądu repo (docs/social-preview.png) ze zrzutów w docs/
│   ├── kraje.json           120 państw: [nazwa, waluta, dieta, limit noclegu]
│   ├── formulas.json        wygenerowane formuły (składnia angielska) – widać w nich zmiany logiki
│   └── layout.json          wygenerowane adresy pól wejściowych
├── docs/                    zrzuty ekranu, obrazek podglądu repo (social-preview.png) i ten dokument
├── .github/ISSUE_TEMPLATE/  formularze zgłoszeń (błąd w wyliczeniu, plik nie działa, propozycja)
├── CHANGELOG.md             historia zmian
├── CLAUDE.md                notatki techniczne i pułapki (dla Claude Code)
└── LICENSE                  licencja MIT
```

Arkusze w pliku: **Kalkulator**, **Raport**, **Delegacja**, **Instrukcja**, **Kursy NBP**, **Kraje** (stawki diet), **Parametry** (procenty za posiłki).

## Budowanie

Potrzebne: Windows z Excelem, Python 3 z `openpyxl` (`pip install -r requirements.txt`), PowerShell 7 (`pwsh`) i internet.

Budowa ma dwa etapy i **zawsze trzeba uruchomić oba**. Formuły są wpisywane przez Excel, bo openpyxl źle zapisuje funkcje tablic dynamicznych (prefiksy `_xlfn.`, niejawne przecięcie `@`).

Publiczne pliki w `wydanie/` buduje jeden skrypt: plik z przykładowymi danymi (sprawdzany testami) i pusty szablon `_pusty.xlsx` (`build.py --pusty`, sprawdzany pod kątem błędów `#...`). Na koniec usuwa z metadanych obu plików nazwę użytkownika, którą Excel dopisuje przy zapisie. Przy błędzie formuł albo nieudanym teście skrypt przerywa pracę:

```bash
pwsh -NoProfile -File build/wydanie.ps1
```

Ręcznie, krok po kroku:

```bash
cd build
python build.py ../wydanie/Kalkulator_delegacji_zagranicznej.xlsx       # struktura → .xlsx (jeszcze bez formuł)
pwsh -NoProfile -File apply.ps1  -File "<pełna ścieżka do .xlsx>"        # formuły, przeliczenie, ochrona
pwsh -NoProfile -File verify.ps1 -File "<pełna ścieżka do .xlsx>"        # kontrola wyników i błędów #...
```

Bez argumentu `build.py` i skrypty działają na roboczym pliku `Kalkulator_delegacji_zagranicznej.xlsx` w katalogu głównym. Ten plik zawiera prawdziwe dane i nie trafia do repozytorium (`.gitignore`). Przebudowa kasuje wpisane dane, więc przed nią zrób kopię, a po `apply.ps1` przywróć dane: `pwsh -NoProfile -File przenies_dane.ps1 -Z "<kopia>.xlsx"`.

## Testy

`build/test.ps1` otwiera plik w Excelu (bez zapisywania) i sprawdza 10 przypadków: przykład z `build.py`, przypadek kontrolny (Węgry), krótki wyjazd (1/3 diety), pusty etap między krajami, tekst zamiast daty, nadmiar posiłków, powrót przed wjazdem, brak kursu oraz kurs ręczny (poprawny i bez nr tabeli). Dla poprawnych danych porównuje kwoty diety i ogółem (także na druku Delegacji) i szuka błędów `#...` we wszystkich arkuszach. Kod wyjścia 1 oznacza, że któryś test nie przeszedł.

```bash
pwsh -NoProfile -File build/test.ps1                      # domyślnie wydanie/Kalkulator_delegacji_zagranicznej.xlsx
```

Testy potrzebują Excela i internetu, więc nie da się ich uruchomić w GitHub Actions.

Test scenariusza bez zapisywania zmian w pliku (tu: przypadek kontrolny – oczekiwany wynik 4 dni 4 godz. 0 min, 190,67 EUR, 834,92 PLN). Adres bez nazwy arkusza dotyczy Kalkulatora; inne arkusze: `Kursy NBP!F8=4.4`:

```bash
pwsh -NoProfile -File verify.ps1 -File "<.xlsx>" -Set "C15=Węgry;D15=2026-09-28 12:00;E15=0;C16=;D16=;E16=;D20=2026-10-02 16:00;K7=2026-10-08;K8=w dniu daty kursu"
pwsh -NoProfile -File export_raport.ps1 -File "<.xlsx>" -Sheet Delegacja -Pdf "$TEMP/d.pdf"
```

## Wydanie nowej wersji

1. Zmień `WERSJA` / `WERSJA_DATA` w `build/build.py` i dopisz zmiany do [`CHANGELOG.md`](../CHANGELOG.md).
2. Zbuduj publiczny plik: `pwsh -NoProfile -File build/wydanie.ps1`. Jeśli zmienił się wygląd albo numer wersji (widoczny w nagłówku), odśwież zrzuty w `docs/`: eksport arkuszy Kalkulator, Raport i Delegacja z `wydanie/` przez `export_raport.ps1 -Sheet …`, potem `pdftoppm -r 110 -png` i przycięcie białych marginesów (strona 1 Kalkulatora i Raportu, obie strony Delegacji). Na koniec `python build/social_preview.py` i wgranie `docs/social-preview.png` w **Settings → General → Social preview** (GitHub nie ma do tego API).
3. Zatwierdź zmiany i wypchnij je na GitHub.
4. Na GitHubie: **Releases → Draft a new release**, tag `vX.Y`, tytuł „Wersja X.Y”, w opisie zmiany z `CHANGELOG.md`. Dołącz oba pliki z `wydanie/` (`Kalkulator_delegacji_zagranicznej.xlsx` i `Kalkulator_delegacji_zagranicznej_pusty.xlsx`) **pod tymi samymi nazwami** – linki „Pobierz” w README prowadzą zawsze do plików o tych nazwach w najnowszym wydaniu. Jeśli zmienił się przykład w `build.py` albo stawki, popraw tabelę „Przykład krok po kroku” w README.

## Aktualizacja stawek

Po zmianie przepisów popraw `build/kraje.json` (kraj, waluta, dieta dzienna, limit noclegu) i wydaj nową wersję. Procenty za posiłki są w arkuszu **Parametry**.

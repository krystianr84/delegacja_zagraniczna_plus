# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Cel projektu

Kalkulator diety zagranicznej dla serwisantów w Excelu z automatycznym wypełnianiem druku delegacji (lokalny wzór `DELEGACJA.pdf`, poza repozytorium). **Liczona** jest tylko dieta zagraniczna; przejazdy, ryczałty, noclegi, inne wydatki i zaliczka są **wpisywane** kwotami (z przeliczeniem waluty po kursie NBP) wyłącznie na potrzeby druku delegacji. Dieta krajowa, kilometrówka i automatyczne liczenie ryczałtów – poza zakresem, nie dodawaj bez prośby.

Produktem jest `Kalkulator_delegacji_zagranicznej.xlsx`. Plik jest bez makr, a wszystko liczą formuły. Kursy NBP pobiera `WEBSERVICE` z `api.nbp.pl`, a waluta i stawka diety wynikają z wybranego kraju. Użytkownicy pracują na Excelu desktopowym na Windows (Microsoft 365 / Excel 2024).

## Wersje

Numer wersji to stałe `WERSJA`/`WERSJA_DATA` w `build.py` (widoczne w nagłówku Kalkulatora i w Instrukcji). Wydanie wersji = zamrożona kopia `Kalkulator_delegacji_zagranicznej ver. X.Y.xlsx` w katalogu głównym (nie nadpisuj jej przy kolejnych zmianach).
- **1.0** – kalkulator diety zagranicznej + Raport.
- **2.4** (09.10.2026) – kurs ręczny awaryjny (Kursy NBP F:H, zastępuje kurs z WEBSERVICE; data i nr tabeli wymagane – kontrola E1), `-Set` z nazwą arkusza.
- **2.3** (09.10.2026) – przykłady dat w komunikatach w formacie pól (`14.09.2026 10:30`).
- **2.2** (09.10.2026) – kwota słownie w Delegacji (I61) ze słowami „złotych i” przed groszami (parametr `jedn` helpera `slownie()`).
- **2.1** (09.10.2026) – kontrole: pusty etap między krajami (dawał podwójną dietę), tekst w polach dat (+ walidacja dat), więcej posiłków danego rodzaju niż rozpoczętych dób etapu; `test.ps1`, `wydanie.ps1`, licencja MIT.
- **2.0** (09.10.2026) – arkusz Delegacja (druk z `DELEGACJA.pdf`), koszty przejazdów/noclegów/ryczałtów/zaliczka, godziny przyjazdu i wyjazdu z miejscowości, sekcja 1 w dwóch kolumnach, kafelki wyniku, pola wymagane i podświetlanie błędów, polskie daty, ochrona arkuszy.

## Budowanie i weryfikacja

Plik `.xlsx` jest **generowany** – nie edytuj go ręcznie, zmieniaj skrypty w `build/`. Budowa ma dwa etapy i **zawsze trzeba uruchomić oba**:

```bash
cd build
python build.py                    # openpyxl: struktura, formaty, dane statyczne, walidacje → .xlsx + formulas.json
pwsh -NoProfile -File apply.ps1    # Excel przez COM: wpisuje formuły (Formula2), przelicza, zapisuje
pwsh -NoProfile -File verify.ps1   # przelicza w Excelu, szuka błędów #..., wypisuje wyniki i kursy
pwsh -NoProfile -File test.ps1     # testy regresji (domyślnie na wydanie/…xlsx z przykładowymi danymi); kod wyjścia 1 = błąd
pwsh -NoProfile -File export_raport.ps1 -Pdf "$TEMP/r.pdf"   # eksport arkusza „Raport” do PDF (podgląd wydruku)
pwsh -NoProfile -File export_raport.ps1 -Sheet Delegacja -Pdf "$TEMP/d.pdf"   # druk delegacji (2 strony)
```

Sam `build.py` nadpisuje `.xlsx` wersją **bez formuł**. Formuły trafiają do pliku dopiero po `apply.ps1`. Opcjonalny argument to inna ścieżka wyjściowa (`python build.py ../wydanie/Kalkulator_delegacji_zagranicznej.xlsx`); `apply.ps1`, `verify.ps1` i `export_raport.ps1` przyjmują wtedy `-File "<pełna ścieżka>"`.

**Repozytorium (GitHub):** robocze `*.xlsx` w katalogu głównym (w tym kopie „ver. X.Y”) zawierają prawdziwe dane i są w `.gitignore`. Publiczny plik do pobrania to `wydanie/Kalkulator_delegacji_zagranicznej.xlsx` – czysty build z przykładowymi danymi z `build.py`; przebuduj go przy każdej zmianie: `pwsh -NoProfile -File build/wydanie.ps1` (build.py + apply + testy + usunięcie `lastModifiedBy`, które Excel wypełnia nazwiskiem użytkownika). Skrypt buduje też pusty szablon `wydanie/Kalkulator_delegacji_zagranicznej_pusty.xlsx` (`build.py --pusty` – bez przykładowych danych, zostaje tylko tryb kursu; sprawdzany przez verify.ps1 pod kątem błędów `#...`). Nowe przykładowe dane w `build.py` dodawaj tak, żeby `--pusty` je pomijał. Liczby przykładu są rozpisane w README („Przykład krok po kroku”) – po zmianie przykładu albo stawek popraw tę tabelę i odśwież zrzuty w `docs/` (eksport arkuszy do PDF → PNG), jeśli zmienił się wygląd. Użytkownicy pobierają plik z GitHub Releases (link `releases/latest/download/Kalkulator_delegacji_zagranicznej.xlsx` w README), więc nowa wersja = też nowe wydanie z tagiem `vX.Y` i plikiem pod **tą samą nazwą**; zmiany dopisz do `CHANGELOG.md` (kroki w `docs/BUDOWANIE.md`). README jest dla użytkownika (pobieranie, obsługa, FAQ), sprawy techniczne idą do `docs/BUDOWANIE.md`.

**Właściciel wpisuje w pliku prawdziwe dane – przebudowa je kasuje.** Przed `build.py` skopiuj plik, a po `apply.ps1` przywróć dane:
`pwsh -NoProfile -File przenies_dane.ps1 -Z "<kopia>.xlsx"` (kopiuje pola wejściowe wg `layout.json`; działa tylko przy tym samym układzie – po zmianie układu trzeba przenieść dane mapowaniem stary→nowy adres). Ręczne zmiany wyglądu zrobione przez właściciela w pliku najpierw przenieś do `build.py`.

Arkusze są chronione bez hasła (`apply.ps1` na końcu wywołuje `Protect`; dozwolone formatowanie, szerokości i wysokości). Edytowalne są tylko komórki odblokowane w `build.py` – `inp()` ustawia `Protection(locked=False)`, a pętla przed zapisem odblokowuje całe scalenia. Nowe pole wejściowe zawsze twórz przez `inp()`, inaczej będzie zablokowane. Skrypty z `-Set` i `przenies_dane.ps1` piszą tylko w pola wejściowe, więc działają przy włączonej ochronie.

Kalkulator ma trzy warstwy formatowania warunkowego, w kolejności pierwszeństwa: czerwone pola z błędem (`blad()` w `build.py`), pomarańczowa ramka pól wymaganych (`wymagane()` – K7, C15, D15, D20 na stałe, D16:D19 po wybraniu kraju) i wyszarzanie nieużywanych etapów. Kafelki wyniku (I1:N2): „Do wypłaty – dieta” (= dieta, bo u właściciela do wypłaty jest zawsze sama dieta), „Wydatki (z zaliczki na wydatki)” (= ogółem − dieta) i „Ogółem do rozliczenia”; pokazują „—”, gdy kontrola E1 ≠ OK. Zaliczka na wydatki jest rozliczana poza kalkulatorem. Nowa reguła w kontroli E1 (zagnieżdżone IF w `build.py`) = nowe podświetlenie `blad()` + nowy przypadek w `test.ps1`.

Daty mają format `dd.mm.yyyy hh:mm` (w pliku kody są niezależne od języka, polski Excel pokazuje dd.mm.rrrr).

Test scenariusza bez zmiany pliku (zmiany nie są zapisywane):
```bash
pwsh -NoProfile -File export_raport.ps1 -Pdf "$TEMP/r.pdf" -Set "C15=Węgry;D15=2026-09-28 12:00;E15=0;C16=;D16=;E16=;D20=2026-10-02 16:00;K7=2026-10-08;K8=w dniu daty kursu"
```
To przypadek kontrolny: oczekiwany wynik to „4 dni 4 godz. 0 min”, 190,67 EUR i 834,92 PLN (kurs 4,3789 z tabeli 196/A/NBP/2026). Komórki wejściowe (aktualne adresy zawsze w `build/layout.json`, generowanym przez `build.py`): sekcja 1 – lewa kolumna C5:C11 (pracownik, stanowisko, cel, nr polecenia, data polecenia, miejscowość wyjazdu, środki lokomocji), prawa K5:K11 (wyjazd z / powrót do miejscowości, data kursu, tryb kursu, zaliczka, termin rozliczenia, liczba dowodów); etapy C15:M19 (C kraj, D wjazd, E–G posiłki, H miejscowość, I–J przyjazd/wyjazd z miejscowości – tylko do druku, K–M przejazd: środek, koszt, waluta; N = PLN); wiersz 20 „Powrót” (D20 powrót przez granicę, K20:M20 przejazd powrotny); sekcja 5 C46:D50 (ryczałty, dojazdy, noclegi, inne wydatki z walutą; F51 = ogółem do rozliczenia).

Do skryptów `.ps1` używaj `pwsh` (PowerShell 7). Wymagany jest zainstalowany Excel i internet (NBP). Python potrzebuje `openpyxl` (nie ma `pywin32`, dlatego COM obsługuje PowerShell).

## Architektura

- `build/build.py` – jedyne źródło prawdy o układzie. Formuły nie są zapisywane przez openpyxl, tylko zbierane w słowniku `F` (`"Arkusz!Komórka" → formuła` w składni angielskiej, z przecinkami) i zrzucane do `formulas.json`. Robimy tak, bo openpyxl źle zapisuje funkcje tablic dynamicznych (`LET`, `TEXTAFTER`, `TEXTSPLIT`, `DROP`, `XLOOKUP` wymagałyby prefiksów `_xlfn.` i spowodowałyby niejawne przecięcie `@`). Numery wierszy wynikają ze stałych `ETAPY`, `E0` (etapy – dane), `RET` (wiersz „Powrót” pod etapami), `R0`/`RS` (wyliczenie i suma) i `V0` (dieta wg walut). `verify.ps1` i `przenies_dane.ps1` czytają numery wierszy z `layout.json`.
- `build/kraje.json` – 120 państw `[nazwa, waluta, dieta, limit_noclegu]` z załącznika do rozporządzenia MPiPS z 29.01.2013 (t.j. Dz.U. 2023 poz. 2190). Limit noclegu nie jest używany.
- Arkusze:
  - **Kalkulator**: sekcja 1 to dane podróży i polecenia wyjazdu (dwie kolumny pól), sekcja 2 to do 5 etapów-krajów (kraj, wjazd, posiłki, miejscowość, godziny, przejazd) plus wiersz „Powrót”, sekcja 3 liczy dietę wiersz w wiersz, sekcja 4 to dieta wg walut, sekcja 5 to pozostałe koszty i ogółem do rozliczenia.
  - **Raport**: wydruk A4 (pionowo, jedna strona). Zawiera dane podróży, tabelę diet i sekcję „Koszty podróży”. Pobiera wszystko formułami z Kalkulatora: sekcji 3 i 4 oraz arkusza Kursy NBP (data i nr tabeli). Kwoty w tekście formatuje `FIXED`, nie `TEXT` (ustawienia regionalne).
  - **Delegacja**: druk „Polecenie wyjazdu służbowego” (str. 1) + „Rachunek kosztów podróży” (str. 2) odwzorowany z `DELEGACJA.pdf`; skala 84%, ręczny podział strony przed wierszem 41. Trasa w tabeli przejazdów: dom (K5) → miejscowość 1 (przyjazd I15) → … → miejscowość n (wyjazd Jn) → dom (K6); puste godziny zastępowane czasami przekroczenia granic. Kwota słownie stoi pod kwotą Ogółem (I61:I65) – układ wybrany przez właściciela; Diety = Kalkulator!M{RS}; koszty przejazdu z kolumny N etapów, pozostałe pozycje z sekcji 5 Kalkulatora (kwoty w walutach obcych przeliczane po kursie NBP z arkusza Kursy NBP, tym samym co dla diety). Sumy w druku liczone `SUM(...)`, nie `+`, bo puste pozycje to tekst "". Kwota słownie z helpera `slownie()` w `build.py` (LET + LAMBDA; nazwy w LET nie mogą wyglądać jak adres komórki, np. `w3`, `c`, `r`). Pola podpisów i zatwierdzeń zostają kropkami. Arkusza **Raport** nie zmieniaj przy pracach nad Delegacją (prośba właściciela).
  - **Kursy NBP**: A to kurs do przeliczenia, B to kursy z okresu w formułach rozlewanych. W A kolumny C:E (kurs, data, nr tabeli – używane przez Kalkulator i Raport) biorą wartości z kursu ręcznego F:H, gdy F nie jest puste; I = źródło kursu. `przenies_dane.ps1` kopiuje F:H tylko z wierszy z wpisanym kursem. Surowy XML jest w ukrytej kolumnie N i ukrytym wierszu 23.
  - **Kraje**, **Parametry** (procenty za posiłki), **Instrukcja**.

## Pułapki (sprawdzone)

- **Nie używaj `FILTERXML` do liczb z NBP.** W polskich ustawieniach regionalnych „4.3171” daje bzdurną liczbę. Parsuj tekstowo: `NUMBERVALUE(TEXTBEFORE(TEXTAFTER(xml,"<Mid>",-1),"</Mid>"),".")`.
- **Nie używaj `TEXT(data,"yyyy-mm-dd")`.** Kody formatu zależą od ustawień regionalnych (w PL to `rrrr`). Daty do URL buduj z `YEAR`/`MONTH`/`DAY` (helper `ymd()` w `build.py`).
- Kurs „z dnia X” = zapytanie o zakres X−14…X i wybór ostatniego notowania (weekendy i święta zwracają 404). Pojedyncze zapytanie API obejmuje maksymalnie 93 dni.
- W skryptach `.ps1` nie używaj polskich cudzysłowów „ ” wewnątrz stringów w podwójnych cudzysłowach – PowerShell traktuje je jak zwykłe `"` i skrypt się nie parsuje.
- Błędy formuł przez COM szukaj w `UsedRange.Value2` (błąd = `Int32`), nie przez `.Text` komórka po komórce – to drugie trwa minuty.
- Tekst `"w dniu daty kursu"` z listy wyboru jest porównywany w formule. Po zmianie etykiety popraw też formułę.

## Reguły merytoryczne zaszyte w formułach

- Czas (§12): od przekroczenia granicy (samolot: start) do powrotu przez granicę (lądowanie). Etap kończy się wjazdem do kolejnego kraju.
- Dieta (§13): liczona osobno dla każdego etapu. Pełne doby + reszta: ≤8 h → 1/3, ≤12 h → 1/2, >12 h → 1.
- Posiłki (§14): −15% / −30% / −30% diety dziennej za każdy posiłek; dieta nie spada poniżej 0.
- Stawki są z rozporządzenia obowiązującego od 29.11.2022. W 2026 r. był projekt podwyżki. Przy zmianie przepisów aktualizuj `kraje.json`.

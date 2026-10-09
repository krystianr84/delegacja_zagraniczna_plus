# Historia zmian

Numer wersji jest widoczny w nagłówku arkusza Kalkulator i w arkuszu Instrukcja. Pliki poszczególnych wersji są w [Wydaniach (Releases)](https://github.com/krystianr84/delegacji_zagranicznej_plus/releases).

## 2.4 (09.10.2026)

**Dodano**
- **Kurs ręczny** w arkuszu Kursy NBP (kolumny F–H: kurs, data notowania, nr tabeli) – na wypadek, gdy kursy nie pobierają się automatycznie (np. sieć firmowa blokuje api.nbp.pl). Wpisany kurs zastępuje kurs z NBP w diecie, kosztach, Raporcie i Delegacji; kolumna „Źródło kursu” pokazuje, skąd jest kurs.
- Kontrolę „Kurs ręczny: uzupełnij datę notowania i nr tabeli” (czerwone podświetlenie brakującego pola). Komunikat „Brak kursu NBP” podpowiada kurs ręczny.
- **Pusty szablon** `Kalkulator_delegacji_zagranicznej_pusty.xlsx` w wydaniu – te same arkusze bez przykładowych danych (`build.py --pusty`; `wydanie.ps1` buduje oba pliki).
- Dwa testy regresji (kurs ręczny zamiast NBP, kurs ręczny bez nr tabeli); `-Set` w skryptach przyjmuje adresy z nazwą arkusza (`Kursy NBP!F8=4.4`). `przenies_dane.ps1` przenosi też kurs ręczny.

**Zmieniono**
- Kolumny z nr tabeli w arkuszu Kursy NBP są szersze (numer mieści się w całości).

## 2.3 (09.10.2026)

**Zmieniono**
- Przykłady dat w komunikatach i podpowiedziach (kontrola danych, walidacja pól z datą, Instrukcja) są w formacie wyświetlanym w polach: `14.09.2026 10:30` zamiast `2026-09-14 10:30`.

## 2.2 (09.10.2026)

**Zmieniono**
- Na druku Delegacji (str. 2) kwota słownie pod „Ogółem” ma słowa „złotych i” przed groszami, np. „dwa tysiące pięćset siedemdziesiąt złotych i 96/100”.

## 2.1 (09.10.2026)

**Naprawiono**
- Pusty etap między dwoma krajami (np. kraj w etapach 1 i 3, etap 2 pusty) dawał podwójną dietę za ten sam okres. Teraz kontrola zgłasza błąd „Nie zostawiaj pustego etapu między krajami”.

**Dodano**
- Kontrolę tekstu w polach dat (np. data wpisana z błędem, której Excel nie rozpoznał): komunikat „Niepoprawna data” i czerwone podświetlenie pola. Dotyczy dat wjazdu i powrotu, godzin przyjazdu i wyjazdu, wyjazdu i powrotu do miejscowości, daty kursu i daty polecenia.
- Walidację danych w polach z datą: Excel odrzuca wpis, który nie jest datą.
- Kontrolę posiłków: liczba śniadań, obiadów lub kolacji w etapie nie może przekraczać liczby rozpoczętych dób pobytu w danym kraju.
- Testy regresji `build/test.ps1` (8 przypadków, w tym przypadek kontrolny) i skrypt `build/wydanie.ps1`, który buduje publiczny plik w `wydanie/` i czyści z niego metadane.
- Licencję MIT oraz informację o Widoku chronionym i o licencji w arkuszu Instrukcja.

## 2.0 (09.10.2026)

**Dodano**
- Arkusz **Delegacja**: druk „Polecenie wyjazdu służbowego” (str. 1) i „Rachunek kosztów podróży” (str. 2) odwzorowany ze standardowego druku delegacji. Trasa w tabeli przejazdów prowadzi z miejscowości wyjazdu przez kolejne miejscowości z powrotem do domu. Gdy godzin przyjazdu i wyjazdu nie wpisano, druk podstawia czasy przekroczenia granic. Kwota ogółem jest też zapisana słownie.
- W sekcji 1 Kalkulatora pola polecenia wyjazdu: stanowisko, nr i data polecenia, miejscowość wyjazdu, środki lokomocji, wyjazd z miejscowości i powrót do niej, zaliczka, termin rozliczenia, liczba dowodów. Sekcja jest ułożona w dwóch kolumnach.
- W etapach: miejscowość, godziny przyjazdu do niej i wyjazdu z niej (tylko do druku) oraz przejazd do kraju (środek lokomocji, koszt, waluta). W wierszu „Powrót” jest przejazd powrotny.
- **Sekcję 5** – ryczałty, dojazdy, noclegi i inne wydatki z walutą. Kwoty w walutach obcych przelicza się po tym samym kursie NBP co dietę. Pod sekcją jest kwota ogółem do rozliczenia.
- Kafelki wyniku: „Do wypłaty – dieta”, „Wydatki (z zaliczki na wydatki)” i „Ogółem do rozliczenia”. Przy błędnych danych pokazują „—”.
- Pomarańczową ramkę pól wymaganych, czerwone podświetlenie pól z błędem, wyszarzanie nieużywanych etapów i legendę kolorów.
- Walidację kwot (liczby dziesiętne) i list wyboru walut.
- Ochronę arkuszy bez hasła: edytowalne są tylko żółte pola.
- Skrypt `build/przenies_dane.ps1`, który przenosi wpisane dane do przebudowanego pliku.

**Zmieniono**
- Daty mają format `dd.mm.rrrr gg:mm` (wcześniej `rrrr-mm-dd gg:mm`).
- W Raporcie kolumna „Czas (godz.)” nazywa się teraz „Czas pobytu”. Poza tym Raport się nie zmienił.

## 1.0

Pierwsza wersja: arkusze Kalkulator, Raport, Instrukcja, Kursy NBP, Kraje i Parametry.
- Do 5 etapów (krajów) plus wiersz „Powrót”. Dla każdego etapu podaje się chwilę wjazdu i liczbę zapewnionych posiłków.
- Dieta liczona osobno dla każdego etapu wg §12–14, w podziale na waluty, z przeliczeniem na PLN.
- Kursy NBP pobierane przez `WEBSERVICE` z dnia przed datą kursu albo z dnia daty kursu.
- Kontrola danych: brak kraju, brak daty wjazdu, kolejność etapów, data powrotu, brak daty kursu, brak kursu NBP.
- Raport na jednej stronie A4.

> Szczegóły wersji 1.0 i 2.1 pochodzą z porównania zamrożonych kopii plików. Kopii wersji 2.0 nie zachowano, dlatego podział zmian między 2.0 a 2.1 opiera się na notatkach z prac.

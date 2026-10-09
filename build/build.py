import json, os, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
PUSTY = "--pusty" in sys.argv[1:]   # pusty szablon: bez przykładowych danych (zostają tylko wartości domyślne)
OUT = ARGS[0] if ARGS else os.path.join(HERE, "..", "Kalkulator_delegacji_zagranicznej.xlsx")
WERSJA, WERSJA_DATA = "2.4", "09.10.2026"
kraje = json.load(open(os.path.join(HERE, "kraje.json"), encoding="utf-8"))
WALUTY = ["EUR", "USD", "AUD", "DKK", "JPY", "CAD", "NOK", "CHF", "SEK", "GBP"]
F = {}  # formuły wpisywane później przez Excel (Formula2)

FONT = "Arial"
f_norm = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=16, bold=True, color="1F3864")
f_sub = Font(name=FONT, size=9, italic=True, color="595959")
f_sec = Font(name=FONT, size=11, bold=True, color="FFFFFF")
f_hdr = Font(name=FONT, size=9, bold=True, color="1F3864")
f_in = Font(name=FONT, size=10, color="0000FF")
f_link = Font(name=FONT, size=10, color="008000")
f_note = Font(name=FONT, size=8, italic=True, color="7F7F7F")
f_total = Font(name=FONT, size=11, bold=True)
fill_sec = PatternFill("solid", start_color="1F3864")
fill_hdr = PatternFill("solid", start_color="D9E1F2")
fill_in = PatternFill("solid", start_color="FFF2CC")
fill_tot = PatternFill("solid", start_color="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
DT = "dd.mm.yyyy hh:mm"
D = "dd.mm.yyyy"
M2 = '#,##0.00;-#,##0.00;"-"'
M4 = '#,##0.0000##;-#,##0.0000##;"-"'


def sec(ws, row, text, last="O"):
    ws[f"B{row}"] = text
    for c in ws[f"B{row}:{last}{row}"][0]:
        c.fill = fill_sec
        c.font = f_sec
    ws.row_dimensions[row].height = 20


def hdr(ws, row, cols, labels, height=42):
    for c, l in zip(cols, labels):
        x = ws[f"{c}{row}"]
        x.value = l
        x.font = f_hdr
        x.fill = fill_hdr
        x.alignment = wrap_c
        x.border = box
    ws.row_dimensions[row].height = height


def inp(cell, value=None, fmt=None):
    if value is not None:
        cell.value = value
    cell.font = f_in
    cell.fill = fill_in
    cell.border = box
    cell.protection = Protection(locked=False)
    if fmt:
        cell.number_format = fmt


def out(cell, fmt=None, bold=False):
    cell.font = f_bold if bold else f_norm
    cell.border = box
    if fmt:
        cell.number_format = fmt


def note(ws, ref, text):
    ws[ref] = text
    ws[ref].font = f_note
    ws[ref].alignment = Alignment(vertical="center", wrap_text=False)


wb = Workbook()
kal = wb.active
kal.title = "Kalkulator"
kur = wb.create_sheet("Kursy NBP")
kra = wb.create_sheet("Kraje")
par = wb.create_sheet("Parametry")
ins = wb.create_sheet("Instrukcja")
rap = wb.create_sheet("Raport")
dele = wb.create_sheet("Delegacja")
for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows(min_row=1, max_row=120, max_col=22):
        for c in row:
            c.font = f_norm


# ---------------- Parametry ----------------
par.column_dimensions["A"].width = 2
par.column_dimensions["B"].width = 46
par.column_dimensions["C"].width = 12
par.column_dimensions["D"].width = 60
par["B1"] = "Parametry diety zagranicznej"
par["B1"].font = f_title
src_rozp = "Rozp. MPiPS z 29.01.2013, t.j. Dz.U. 2023 poz. 2190"
P = [
    (3, "Potrącenie za śniadanie (% diety)", 0.15, f"{src_rozp}, §14 ust. 2"),
    (4, "Potrącenie za obiad (% diety)", 0.30, f"{src_rozp}, §14 ust. 2"),
    (5, "Potrącenie za kolację (% diety)", 0.30, f"{src_rozp}, §14 ust. 2"),
]
for r, l, v, s in P:
    par[f"B{r}"] = l
    inp(par[f"C{r}"], v, "0%")
    note(par, f"D{r}", "Źródło: " + s)

# ---------------- Kraje ----------------
kra.column_dimensions["A"].width = 32
for c in "BC":
    kra.column_dimensions[c].width = 16
hdr(kra, 1, "ABC", ["Państwo", "Waluta", "Dieta za dobę"], 30)
for i, (n, w, d, _limit) in enumerate(kraje, start=2):
    kra[f"A{i}"] = n
    kra[f"A{i}"].border = box
    inp(kra[f"B{i}"], w)
    inp(kra[f"C{i}"], d, "#,##0")
KR_LAST = 1 + len(kraje)
note(kra, "E2", "Źródło: załącznik do " + src_rozp + " (stawki obowiązujące od 29.11.2022).")
note(kra, "E3", "Gibraltar = Wielka Brytania; Hongkong 55 USD; Tajwan 40 EUR (uwagi pod tabelą załącznika).")
note(kra, "E4", "Po zmianie rozporządzenia wystarczy poprawić kwoty w tej tabeli – kalkulator przeliczy się sam.")
kra.freeze_panes = "A2"
KR_A = f"Kraje!$A$2:$A${KR_LAST}"
KR = lambda col: f"Kraje!${col}$2:${col}${KR_LAST}"

# ---------------- Kalkulator: stałe układu ----------------
# sekcja 1: dwie kolumny pól – lewa (B: etykieta, C:G wartość), prawa (I:J etykieta, K:L wartość)
PRAC, STAN, CEL, NR, DPOL, MIEJ, SROD = (f"C{r}" for r in range(5, 12))
WYJ, POW, DKURS, TRYB, ZAL, TERM, DOW = (f"K{r}" for r in range(5, 12))
ETAPY = 5
E0 = 15                 # pierwszy wiersz etapów (dane)
E1 = E0 + ETAPY - 1
RET = E1 + 1            # wiersz „Powrót” (przekroczenie granicy PL)
R0 = RET + 4            # pierwszy wiersz wyliczeń
R1 = R0 + ETAPY - 1
RS = R1 + 1             # wiersz „DIETA RAZEM”
V0 = RS + 2             # sekcja: dieta wg walut

# ---------------- Kursy NBP ----------------
kur.column_dimensions["A"].width = 2
kur.column_dimensions["B"].width = 37.7
for c in "CDEFGHIJKL":
    kur.column_dimensions[c].width = 13
kur["B1"] = "Kursy średnie NBP (tabela A) – pobierane automatycznie z api.nbp.pl"
kur["B1"].font = f_title
sec(kur, 3, "A. Kursy do przeliczenia diety", "L")
kur["B4"] = "Data kursu (z Kalkulatora)"
out(kur["C4"], D)
kur["C4"].font = f_link
F["'Kursy NBP'!C4"] = f'=IF(Kalkulator!{DKURS}="","",Kalkulator!{DKURS})'
kur["B5"] = "Szukany kurs: ostatni opublikowany do dnia"
out(kur["C5"], D)
F["'Kursy NBP'!C5"] = f'=IF(C4="","",C4-IF(Kalkulator!{TRYB}="w dniu daty kursu",0,1))'
# C:E – kurs użyty w wyliczeniach (z NBP albo z kursu ręcznego); F:H – kurs ręczny (awaryjnie, gdy WEBSERVICE nie działa)
hdr(kur, 7, "BCDEFGHI", ["Waluta", "Kurs średni (PLN za 1 jedn.)", "Data notowania", "Nr tabeli NBP",
                         "Kurs ręczny (opcjonalnie)", "Data notowania (ręcznie)", "Nr tabeli (ręcznie)", "Źródło kursu"], 42)
KRS0, KRS1 = 8, 7 + len(WALUTY)
ymd = lambda d: f'YEAR({d})&"-"&RIGHT("0"&MONTH({d}),2)&"-"&RIGHT("0"&DAY({d}),2)'
for i, w in enumerate(WALUTY):
    r = KRS0 + i
    kur[f"B{r}"] = w
    for c in "BCDEI":
        out(kur[f"{c}{r}"])
    kur[f"C{r}"].number_format = M4
    kur[f"D{r}"].number_format = D
    inp(kur[f"F{r}"], fmt=M4)
    inp(kur[f"G{r}"], fmt=D)
    inp(kur[f"H{r}"], fmt="@")
    # surowa odpowiedź API (ukryta kolumna N) – jedno zapytanie na walutę
    F[f"'Kursy NBP'!N{r}"] = (
        f'=IF($C$5="","",IFERROR(WEBSERVICE("https://api.nbp.pl/api/exchangerates/rates/A/"&B{r}&"/"&'
        f'{ymd("($C$5-14)")}&"/"&{ymd("$C$5")}&"/?format=xml"),"BLAD"))')
    F[f"'Kursy NBP'!C{r}"] = (
        f'=IF(F{r}<>"",F{r},IF(N{r}="","",IFERROR(NUMBERVALUE(TEXTBEFORE(TEXTAFTER(N{r},"<Mid>",-1),"</Mid>"),"."),"brak kursu")))')
    F[f"'Kursy NBP'!D{r}"] = (
        f'=IF(F{r}<>"",IF(G{r}="","",G{r}),IF(N{r}="","",IFERROR(LET(s,TEXTBEFORE(TEXTAFTER(N{r},"<EffectiveDate>",-1),"</EffectiveDate>"),'
        f'DATE(LEFT(s,4),MID(s,6,2),RIGHT(s,2))),"-")))')
    F[f"'Kursy NBP'!E{r}"] = (
        f'=IF(F{r}<>"",H{r},IF(N{r}="","",IFERROR(TEXTBEFORE(TEXTAFTER(N{r},"<No>",-1),"</No>"),"-")))')
    F[f"'Kursy NBP'!I{r}"] = f'=IF(F{r}<>"","wpisany ręcznie",IF(ISNUMBER(C{r}),"NBP",""))'
KURS_W = f"'Kursy NBP'!$B${KRS0}:$B${KRS1}"
KURS_K = f"'Kursy NBP'!$C${KRS0}:$C${KRS1}"
KRF = f"'Kursy NBP'!$F${KRS0}:$F${KRS1}"
KRG = f"'Kursy NBP'!$G${KRS0}:$G${KRS1}"
KRH = f"'Kursy NBP'!$H${KRS0}:$H${KRS1}"
kur.column_dimensions["N"].hidden = True
for c in "FG":
    kur.column_dimensions[c].width = 14
kur.column_dimensions["I"].width = 15
for c in "EH":   # nr tabeli, np. 196/A/NBP/2026
    kur.column_dimensions[c].width = 15.5
note(kur, "K8", "Kurs z ostatniego dnia roboczego – zapytanie obejmuje 14 dni wstecz, więc weekendy i święta nie powodują błędu.")
note(kur, "K9", "Kursy pobierają się same po zmianie daty. Ręczne odświeżenie: Ctrl+Alt+F9.")
note(kur, "K11", "Kurs ręczny (F–H) wpisz tylko wtedy, gdy kurs nie pobiera się automatycznie (np. sieć firmowa blokuje api.nbp.pl).")
note(kur, "K12", "Kurs średni, datę i nr tabeli A znajdziesz na nbp.pl/statystyka-i-sprawozdawczosc/kursy/tabela-a. Wpisany kurs zastępuje kurs z NBP.")
note(kur, "K13", "Żeby wrócić do kursu pobieranego automatycznie, usuń kurs ręczny (kolumny F–H).")
vkr = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True,
                     showErrorMessage=True, errorTitle="Niepoprawny kurs", error="Wpisz kurs większy od zera, np. 4,2563.")
vkd = DataValidation(type="decimal", operator="between", formula1="36526", formula2="73051", allow_blank=True,
                     showErrorMessage=True, errorTitle="Niepoprawna data", error="Wpisz datę notowania, np. 14.09.2026.")
for v_, ref_ in ((vkr, f"F{KRS0}:F{KRS1}"), (vkd, f"G{KRS0}:G{KRS1}")):
    kur.add_data_validation(v_)
    v_.add(ref_)
# brakująca data / nr tabeli przy wpisanym kursie ręcznym – czerwone pole (Raport drukuje datę i nr tabeli)
for c in "GH":
    kur.conditional_formatting.add(f"{c}{KRS0}:{c}{KRS1}", FormulaRule(
        formula=[f'AND($F{KRS0}<>"",{c}{KRS0}="")'], fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")))
kur.conditional_formatting.add(f"I{KRS0}:I{KRS1}", FormulaRule(
    formula=[f'$F{KRS0}<>""'], font=Font(name=FONT, bold=True, color="C65911")))

sec(kur, 20, "B. Kursy NBP z podanego okresu (domyślnie pobyt za granicą – daty można nadpisać)", "L")
kur["B21"] = "Od dnia"
inp(kur["C21"], fmt=D)
F["'Kursy NBP'!C21"] = f'=IF(Kalkulator!D{E0}="","",INT(Kalkulator!D{E0}))'
kur["B22"] = "Do dnia"
inp(kur["C22"], fmt=D)
F["'Kursy NBP'!C22"] = f'=IF(Kalkulator!D{RET}="","",INT(Kalkulator!D{RET}))'
note(kur, "D21", "Maks. 93 dni w jednym zapytaniu (ograniczenie API NBP).")
hdr(kur, 24, ["B"] + list("CDEFGHIJKL"), ["Data notowania"] + WALUTY, 22)
for j, w in enumerate(WALUTY):
    c = "CDEFGHIJKL"[j]
    F[f"'Kursy NBP'!{c}23"] = (
        f'=IF(OR($C$21="",$C$22=""),"",IFERROR(WEBSERVICE("https://api.nbp.pl/api/exchangerates/rates/A/{w}/"&'
        f'{ymd("$C$21")}&"/"&{ymd("$C$22")}&"/?format=xml"),"BLAD"))')
    F[f"'Kursy NBP'!{c}25"] = (
        f'=IF({c}23="","",IFERROR(TRANSPOSE(NUMBERVALUE(TEXTBEFORE(DROP(TEXTSPLIT({c}23,"<Mid>"),,1),"</Mid>"),".")),"brak notowań"))')
F["'Kursy NBP'!B25"] = (
    '=IF(C23="","",IFERROR(LET(p,LEFT(DROP(TEXTSPLIT(C23,"<EffectiveDate>"),,1),10),'
    'TRANSPOSE(DATE(LEFT(p,4),MID(p,6,2),RIGHT(p,2)))),"brak notowań w tym okresie"))')
kur.row_dimensions[23].hidden = True
for r in range(25, 25 + 93):
    kur[f"B{r}"].number_format = D
    kur[f"B{r}"].alignment = Alignment(horizontal="center")
    for c in "CDEFGHIJKL":
        kur[f"{c}{r}"].number_format = "0.0000##"
kur.freeze_panes = "A4"

# ---------------- Kalkulator ----------------
widths = {"A": 2, "B": 31.6, "C": 22, "D": 19}
for c in "EFGHIJKLM":
    widths[c] = 13
widths["F"] = widths["G"] = 16.5
widths["I"] = widths["J"] = 16.5   # daty z godziną przyjazdu/wyjazdu; etykiety prawej kolumny sekcji 1 (I:J)
for k, v in widths.items():
    kal.column_dimensions[k].width = v
kal["B1"] = "Kalkulator diety zagranicznej"
kal["B1"].font = f_title
kal["B2"] = f"ver. {WERSJA} ({WERSJA_DATA}) · wg rozporządzenia MPiPS z 29.01.2013 (t.j. Dz.U. 2023 poz. 2190) · kursy NBP pobierane automatycznie"
kal["B2"].font = f_sub
kal["D1"] = "Kontrola danych:"
kal["D1"].font = f_bold
kal["D1"].alignment = Alignment(horizontal="right", vertical="center")
kal.merge_cells("E1:H1")
kal["E1"].font = f_bold
kal["E1"].alignment = Alignment(vertical="center")

sec(kal, 4, "1. Dane podróży i polecenia wyjazdu", "N")
dane = [  # (adres, etykieta, przykład, format, podpowiedź)
    (PRAC, "Pracownik (imię i nazwisko)", "Jan Kowalski", None, None),
    (STAN, "Stanowisko służbowe", "Serwisant", None, None),
    (CEL, "Cel podróży / zlecenie serwisowe", "Serwis linii pakującej – klient w Monachium i Linzu", None, None),
    (NR, "Nr polecenia wyjazdu", "1/2026", None, None),
    (DPOL, "Data polecenia (z dnia)", dt.date(2026, 9, 10), D, None),
    (MIEJ, "Miejscowość wyjazdu i powrotu", "Poznań", None, None),
    (SROD, "Środki lokomocji", "samochód służbowy", None, "Domyślny środek lokomocji w tabeli przejazdów delegacji."),
    (WYJ, "Wyjazd z miejscowości (data i godz.)", dt.datetime(2026, 9, 14, 5, 30), DT, "Wyjazd z domu / siedziby. Puste = przekroczenie granicy z etapu 1."),
    (POW, "Powrót do miejscowości (data i godz.)", dt.datetime(2026, 9, 18, 22, 0), DT, "Powrót do domu / siedziby. Puste = powrót przez granicę."),
    (DKURS, "Data kursu NBP", dt.date(2026, 9, 10), D, "Dzień wypłaty zaliczki / rozliczenia. Kurs średni NBP pobiera się automatycznie."),
    (TRYB, "Kurs NBP z dnia", "dzień przed datą kursu", None, "„dzień przed” = ostatni dzień roboczy przed datą (praktyka podatkowa); „w dniu” = §20 rozporządzenia."),
    (ZAL, "Pobrana zaliczka (zł)", None, M2, "Tylko do druku delegacji: pomniejsza kwotę do wypłaty."),
    (TERM, "Termin rozliczenia zaliczki (dni)", None, "0", None),
    (DOW, "Liczba załączonych dowodów", None, "0", "Rachunki, bilety itp. – do pola „Załączam … dowodów”."),
]
if PUSTY:
    dane = [(a, l, v if a == TRYB else None, fm_, n) for a, l, v, fm_, n in dane]
for a, l, v, fm_, n in dane:
    col, r = a[0], int(a[1:])
    lab = "B" if col == "C" else "I"
    kal[f"{lab}{r}"] = l
    if col == "C":
        kal.merge_cells(f"C{r}:G{r}")
    else:
        kal.merge_cells(f"I{r}:J{r}")
        kal.merge_cells(f"K{r}:L{r}")
    inp(kal[a], v, fm_)
    kal[a].alignment = Alignment(horizontal="left")
    if n:
        kal[a].comment = Comment(n, "Kalkulator")

# pasek sekcji 2 podzielony na grupy kolumn: dieta | miejscowość i godziny | przejazd
GRUPY = [("B", "G", "2. Pobyt za granicą – kraj, wjazd i zapewnione posiłki (dieta)", "1F3864"),
         ("H", "J", "Miejscowość i godziny (do druku delegacji)", "2E75B6"),
         ("K", "N", "Przejazd (do rachunku kosztów)", "548235")]
for c0, c1, txt_, kol in GRUPY:
    kal[f"{c0}{E0-2}"] = txt_
    for c in kal[f"{c0}{E0-2}:{c1}{E0-2}"][0]:
        c.fill = PatternFill("solid", start_color=kol)
        c.font = f_sec
    kal[f"{c1}{E0-2}"].border = Border(right=Side(style="medium", color="FFFFFF"))
kal.row_dimensions[E0 - 2].height = 20
hdr(kal, E0 - 1, "BCDEFGHIJKLMN", ["Etap", "Kraj", "Wjazd do kraju (data i godz.)",
                                 "Śniadania (szt.)", "Obiady (szt.)", "Kolacje (szt.)",
                                 "Miejscowość", "Przyjazd (data i godz.)", "Wyjazd (data i godz.)",
                                 "Środek lokomocji", "Koszt", "Waluta", "Koszt (PLN)"], 32)
kal[f"B{E0-1}"].comment = Comment("Każdy kraj to osobny etap. Etap trwa do wjazdu do następnego kraju albo do powrotu przez granicę.\n"
                                  "Kolejny wiersz uaktywnia się po wybraniu kraju w poprzednim.", "Kalkulator")
kal[f"E{E0-1}"].comment = Comment("Liczba posiłków zapewnionych w tym kraju (np. śniadania w cenie hotelu) – pomniejszają dietę (§14).", "Kalkulator")
kal[f"K{E0-1}"].comment = Comment("Puste = środek lokomocji z sekcji 1.", "Kalkulator")
przyklad = {
    E0: ["Niemcy", dt.datetime(2026, 9, 14, 10, 30), 2, 0, 0, "Monachium",
         dt.datetime(2026, 9, 14, 12, 15), dt.datetime(2026, 9, 16, 9, 0)],
    E0 + 1: ["Austria", dt.datetime(2026, 9, 16, 14, 0), 2, 0, 0, "Linz",
             dt.datetime(2026, 9, 16, 15, 30), dt.datetime(2026, 9, 18, 14, 0)],
}
if PUSTY:
    przyklad = {}
for i in range(ETAPY):
    r = E0 + i
    kal[f"B{r}"] = f"Etap {i+1}"
    kal[f"B{r}"].border = box
    kal[f"B{r}"].alignment = Alignment(horizontal="center")
    vals = przyklad.get(r, [None] * 8)
    for c, v, fm in zip("CDEFGHIJ", vals, [None, DT, "0", "0", "0", None, DT, DT]):
        inp(kal[f"{c}{r}"], v, fm)
        kal[f"{c}{r}"].alignment = Alignment(horizontal="center")
fill_off = PatternFill("solid", start_color="F2F2F2")
kal[f"B{RET}"] = "Powrót"
kal[f"C{RET}"] = "Polska"
for c in "BCEFGHIJ":
    kal[f"{c}{RET}"].border = box
    kal[f"{c}{RET}"].fill = fill_off
    kal[f"{c}{RET}"].alignment = Alignment(horizontal="center")
kal[f"B{RET}"].font = f_bold
inp(kal[f"D{RET}"], None if PUSTY else dt.datetime(2026, 9, 18, 16, 45), DT)
kal[f"D{RET}"].alignment = Alignment(horizontal="center")
kal.merge_cells(f"E{RET}:G{RET}")
kal[f"E{RET}"] = "koniec pobytu za granicą (§12)"
kal[f"E{RET}"].font = f_note
# przejazd do danego kraju (wiersz „Powrót” = przejazd powrotny); waluta obca → kurs NBP jak dla diety
pln = lambda kw, wal: (f'IF(N({kw})=0,"",IF(OR({wal}="",{wal}="PLN"),{kw},'
                       f'IFERROR(ROUND({kw}*XLOOKUP({wal},{KURS_W},{KURS_K}),2),"brak kursu")))')
for r in range(E0, RET + 1):
    inp(kal[f"K{r}"])
    inp(kal[f"L{r}"], fmt=M2)
    inp(kal[f"M{r}"], "PLN")
    kal[f"M{r}"].alignment = Alignment(horizontal="center")
    out(kal[f"N{r}"], M2)
    F[f"Kalkulator!N{r}"] = f"={pln(f'L{r}', f'M{r}')}"
kal.column_dimensions["N"].width = 13
kal[f"I{E0}"].comment = Comment("Godzina dotarcia na miejsce (hotel / klient).\nPuste = chwila wjazdu do kraju.\n"
                                "Tylko do druku delegacji – nie wpływa na dietę.", "Kalkulator")
kal[f"J{E0}"].comment = Comment("Godzina wyjazdu z miejscowości (np. z hotelu).\nPuste = wjazd do kolejnego kraju "
                                "albo powrót przez granicę.\nTylko do druku delegacji – nie wpływa na dietę.", "Kalkulator")

sec(kal, R0 - 2, "3. Wyliczenie diety (automatycznie)", "M")
hdr(kal, R0 - 1, "BCDEFGHIJKLM", [
    "Etap", "Kraj", "Waluta", "Dieta za dobę", "Od", "Do", "Czas (godz.)", "Liczba diet",
    "Potrącenie za posiłki", "Dieta należna (waluta)", "Kurs NBP", "Dieta należna (PLN)"], 42)
fm = {"C": None, "D": None, "E": "#,##0.00", "F": DT, "G": DT, "H": "0.00", "I": "0.0000",
      "J": M2, "K": M2, "L": "0.0000##", "M": M2}
for i in range(ETAPY):
    r, k = R0 + i, E0 + i
    nxt = f'IF(C{k+1}<>"",D{k+1},$D${RET})' if i < ETAPY - 1 else f"$D${RET}"
    kal[f"B{r}"] = f"Etap {i+1}"
    kal[f"B{r}"].border = box
    kal[f"B{r}"].alignment = Alignment(horizontal="center")
    for c, f_ in fm.items():
        out(kal[f"{c}{r}"], f_, bold=(c in "KM"))
    on = f'C{k}=""'
    F[f"Kalkulator!C{r}"] = f'=IF({on},"",C{k})'
    F[f"Kalkulator!D{r}"] = f'=IF({on},"",XLOOKUP(C{k},{KR_A},{KR("B")},"?"))'
    F[f"Kalkulator!E{r}"] = f'=IF({on},"",XLOOKUP(C{k},{KR_A},{KR("C")},0))'
    F[f"Kalkulator!F{r}"] = f'=IF({on},"",D{k})'
    F[f"Kalkulator!G{r}"] = f'=IF({on},"",{nxt})'
    F[f"Kalkulator!H{r}"] = f'=IF({on},"",IFERROR(MAX(0,ROUND((G{r}-F{r})*1440,0)/60),0))'
    # §13 ust. 3: pełne doby + niepełna doba: do 8 h 1/3, 8–12 h 1/2, ponad 12 h cała dieta
    F[f"Kalkulator!I{r}"] = (f'=IF({on},"",LET(h,H{r},d,INT(h/24),reszta,h-24*d,'
                             f'd+IF(reszta=0,0,IF(reszta<=8,1/3,IF(reszta<=12,1/2,1)))))')
    # §14 ust. 2: śniadanie 15%, obiad 30%, kolacja 30% diety dziennej
    F[f"Kalkulator!J{r}"] = (f'=IF({on},"",ROUND(E{r}*(N(E{k})*Parametry!$C$3+N(F{k})*Parametry!$C$4'
                             f'+N(G{k})*Parametry!$C$5),2))')
    F[f"Kalkulator!K{r}"] = f'=IF({on},"",MAX(0,ROUND(E{r}*I{r}-J{r},2)))'
    F[f"Kalkulator!L{r}"] = f'=IF({on},"",XLOOKUP(D{r},{KURS_W},{KURS_K},"brak kursu"))'
    F[f"Kalkulator!M{r}"] = f'=IF({on},"",IF(ISNUMBER(L{r}),ROUND(K{r}*L{r},2),"brak kursu"))'
kal[f"B{RS}"] = "DIETA RAZEM (PLN)"
for c in "BCDEFGHIJKLM":
    kal[f"{c}{RS}"].fill = fill_tot
    kal[f"{c}{RS}"].font = f_total
    kal[f"{c}{RS}"].border = box
F[f"Kalkulator!M{RS}"] = f"=SUM(M{R0}:M{R1})"
kal[f"M{RS}"].number_format = M2

sec(kal, V0, "4. Dieta według walut", "E")
hdr(kal, V0 + 1, "BCDE", ["Waluta", "Dieta (waluta)", "Kurs NBP", "Dieta (PLN)"], 30)
for i, w in enumerate(WALUTY):
    r = V0 + 2 + i
    kal[f"B{r}"] = w
    kal[f"B{r}"].alignment = Alignment(horizontal="center")
    for c, f_ in zip("BCDE", [None, M2, "0.0000##", M2]):
        out(kal[f"{c}{r}"], f_)
    F[f"Kalkulator!C{r}"] = f'=SUMIF($D${R0}:$D${R1},B{r},$K${R0}:$K${R1})'
    F[f"Kalkulator!D{r}"] = f'=IF(C{r}=0,"",XLOOKUP(B{r},{KURS_W},{KURS_K},"brak kursu"))'
    F[f"Kalkulator!E{r}"] = f'=IF(D{r}="",0,IF(ISNUMBER(D{r}),ROUND(C{r}*D{r},2),"brak kursu"))'
VL = V0 + 1 + len(WALUTY)

# sekcja 5: pozostałe koszty podróży (do rachunku kosztów w arkuszu „Delegacja”)
S6 = VL + 2
sec(kal, S6, "5. Pozostałe koszty podróży (do rachunku kosztów w arkuszu „Delegacja”)", "M")
hdr(kal, S6 + 1, "BCDEF", ["Pozycja", "Kwota", "Waluta", "Kurs NBP", "Kwota (PLN)"], 30)
KOSZTY = ["Ryczałty za dojazdy", "Dojazdy udokumentowane", "Noclegi wg rachunków", "Noclegi – ryczałt",
          "Inne wydatki wg załączników"]
KO0 = S6 + 2
KO = {}
przyklad6 = {} if PUSTY else {"Noclegi wg rachunków": (380, "EUR")}
for i, l in enumerate(KOSZTY):
    r = KO0 + i
    KO[l] = r
    kal[f"B{r}"] = l
    kal[f"B{r}"].border = box
    v, w = przyklad6.get(l, (None, "PLN"))
    inp(kal[f"C{r}"], v, M2)
    inp(kal[f"D{r}"], w)
    kal[f"D{r}"].alignment = Alignment(horizontal="center")
    out(kal[f"E{r}"], "0.0000##")
    out(kal[f"F{r}"], M2)
    F[f"Kalkulator!E{r}"] = f'=IF(OR(N(C{r})=0,D{r}="",D{r}="PLN"),"",XLOOKUP(D{r},{KURS_W},{KURS_K},"brak kursu"))'
    F[f"Kalkulator!F{r}"] = f"={pln(f'C{r}', f'D{r}')}"
KO1 = KO0 + len(KOSZTY) - 1
note(kal, f"G{KO['Ryczałty za dojazdy']}", "Ryczałt za dojazdy: 10% diety za każdą rozpoczętą dobę pobytu (§16 ust. 3).")
note(kal, f"G{KO['Noclegi – ryczałt']}", "Ryczałt za nocleg bez rachunku: 25% limitu noclegu danego kraju (§16 ust. 2).")
note(kal, f"G{KO['Inne wydatki wg załączników']}", "Np. parking, opłaty drogowe, wiza, telefon – wg rachunków.")
KT = KO1 + 1
kal[f"B{KT}"] = "OGÓŁEM do rozliczenia (PLN): diety + przejazdy + pozostałe koszty"
for c in "BCDEF":
    kal[f"{c}{KT}"].fill = fill_tot
    kal[f"{c}{KT}"].font = f_total
    kal[f"{c}{KT}"].border = box
F[f"Kalkulator!F{KT}"] = f"=M{RS}+SUM(N{E0}:N{RET})+SUM(F{KO0}:F{KO1})"
kal[f"F{KT}"].number_format = M2
KL = KT

# kontrola danych
F["Kalkulator!E1"] = (
    f'=IF(C{E0}="","Wybierz kraj w etapie 1",'
    # pusty etap między krajami: poprzedni kraj liczyłby się do powrotu, a kolejny drugi raz (podwójna dieta)
    f'IF(SUMPRODUCT((C{E0+1}:C{E1}<>"")*(C{E0}:C{E1-1}=""))>0,"Nie zostawiaj pustego etapu między krajami",'
    f'IF(SUMPRODUCT(--ISTEXT(D{E0}:D{RET}))+SUMPRODUCT(--ISTEXT(I{E0}:J{E1}))+ISTEXT({WYJ})+ISTEXT({POW})+ISTEXT({DKURS})+ISTEXT({DPOL})>0,'
    f'"Niepoprawna data (wpisz np. 14.09.2026 10:30) – zobacz czerwone pole",'
    f'IF(SUMPRODUCT((C{E0}:C{E1}<>"")*(D{E0}:D{E1}=""))>0,"Uzupełnij datę wjazdu w każdym etapie",'
    f'IF(SUMPRODUCT((C{E0+1}:C{E1}<>"")*(D{E0+1}:D{E1}<=D{E0}:D{E1-1}))>0,"Etapy muszą być w kolejności chronologicznej",'
    f'IF(D{RET}="","Uzupełnij datę powrotu przez granicę",'
    f'IF(D{RET}<=MAX(D{E0}:D{E1}),"Powrót przez granicę musi być po ostatnim wjeździe",'
    # posiłków jednego rodzaju nie może być więcej niż rozpoczętych dób pobytu w danym kraju
    f'IF(SUM((E{E0}:G{E1}>IFERROR(ROUNDUP(I{R0}:I{R1},0),1E9))*1)>0,"Więcej posiłków niż dni pobytu w kraju – sprawdź śniadania, obiady i kolacje",'
    f'IF({DKURS}="","Uzupełnij datę kursu",'
    # kurs ręczny bez daty notowania lub nr tabeli (Raport drukuje oba)
    f'IF(SUMPRODUCT(({KRF}<>"")*((({KRG}="")+({KRH}=""))>0))>0,"Kurs ręczny: uzupełnij datę notowania i nr tabeli (arkusz Kursy NBP)",'
    f'IF(COUNTIF(M{R0}:M{R1},"brak kursu")>0,"Brak kursu NBP – sprawdź datę kursu i internet albo wpisz kurs ręcznie (arkusz Kursy NBP)","OK")))))))))))')
kal.conditional_formatting.add("E1:H1", FormulaRule(formula=['$E$1="OK"'], fill=PatternFill("solid", start_color="C6EFCE"), font=Font(name=FONT, bold=True, color="006100")))
kal.conditional_formatting.add("E1:H1", FormulaRule(formula=['$E$1<>"OK"'], fill=PatternFill("solid", start_color="FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006")))


def dv(ws, formula, ref, prompt=None):
    v = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                       errorTitle="Niepoprawna wartość", error="Wybierz wartość z listy.")
    if prompt:
        v.promptTitle, v.prompt, v.showInputMessage = "Podpowiedź", prompt, True
    ws.add_data_validation(v)
    for x in ref.split():
        v.add(x)


dv(kal, '"dzień przed datą kursu,w dniu daty kursu"', TRYB)
WAL_LIST = '"' + ",".join(["PLN"] + WALUTY) + '"'
dv(kal, WAL_LIST, f"M{E0}:M{RET} D{KO0}:D{KO1}")
vk = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                    showErrorMessage=True, error="Wpisz kwotę ≥ 0.")
kal.add_data_validation(vk)
vk.add(f"L{E0}:L{RET}")
vk.add(f"C{KO0}:C{KO1}")
dv(kal, f"={KR_A}", f"C{E0}:C{E1}", "Wybierz kraj – waluta i stawka diety ustawią się automatycznie.")
v = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                   showErrorMessage=True, error="Wpisz liczbę całkowitą ≥ 0.")
kal.add_data_validation(v)
v.add(f"E{E0}:G{E1}")
vd = DataValidation(type="decimal", operator="between", formula1="36526", formula2="73051", allow_blank=True,
                    showErrorMessage=True, errorTitle="Niepoprawna data",
                    error="Wpisz datę (i godzinę), np. 14.09.2026 10:30.")
kal.add_data_validation(vd)
for ref in (f"D{E0}:D{RET}", f"I{E0}:J{E1}", WYJ, POW, DKURS, DPOL):
    vd.add(ref)

kal[f"D{E0}"].comment = Comment("Podróż lądowa: chwila przekroczenia granicy.\nSamolot: start z ostatniego lotniska w kraju (§12).\nKolejny etap: wjazd do kolejnego kraju.", "Kalkulator")
kal[f"D{RET}"].comment = Comment("Podróż lądowa: przekroczenie granicy w drodze powrotnej.\nSamolot: lądowanie na pierwszym lotnisku w kraju.", "Kalkulator")
# kafelki z wynikiem (wiersze 1–2, kolumny I:N)
# do wypłaty jest zawsze dieta; pozostałe wydatki serwisant pokrywa z zaliczki na wydatki (poza kalkulatorem)
KAFLE = [("I", "J", f'="Do wypłaty – dieta"&IF($E$1<>"OK","",IFERROR(" ("&TEXTJOIN(" + ",TRUE,FILTER(FIXED(C{V0+2}:C{VL},2)&" "&B{V0+2}:B{VL},C{V0+2}:C{VL}<>0))&")",""))',
          f"=M{RS}"),
         ("K", "L", '="Wydatki (z zaliczki na wydatki)"', f"=F{KT}-M{RS}"),
         ("M", "N", '="Ogółem do rozliczenia"', f"=F{KT}")]
for c0, c1, lab, wart in KAFLE:
    kal.merge_cells(f"{c0}1:{c1}1")
    kal.merge_cells(f"{c0}2:{c1}2")
    F[f"Kalkulator!{c0}1"] = lab
    F[f"Kalkulator!{c0}2"] = f'=IF($E$1<>"OK","—",{wart[1:]})'   # przy błędnych danych nie pokazuj kwot
    for c in kal[f"{c0}1:{c1}1"][0]:
        c.fill = fill_sec
        c.font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    for c in kal[f"{c0}2:{c1}2"][0]:
        c.fill = fill_tot
        c.font = Font(name=FONT, size=15, bold=True, color="1F3864")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.number_format = '#,##0.00 "zł"'
    kal[f"{c0}1"].border = Border(left=Side(style="medium", color="FFFFFF"))
    kal[f"{c0}2"].border = Border(left=Side(style="medium", color="FFFFFF"))
kal.row_dimensions[1].height = 18
kal.row_dimensions[2].height = 30

# --- pola wymagane (pomarańczowa ramka) i legenda kolorów ---
ORANGE = "ED7D31"
o_m = Side(style="medium", color=ORANGE)


def wymagane(ref):
    """Gruba pomarańczowa ramka wokół pola (także scalonego)."""
    rows = list(kal[ref if ":" in ref else f"{ref}:{ref}"])
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            b = c.border
            c.border = Border(left=o_m if j == 0 else b.left, right=o_m if j == len(row) - 1 else b.right,
                              top=o_m if i == 0 else b.top, bottom=o_m if i == len(rows) - 1 else b.bottom)


for ref in (f"{DKURS}:{chr(ord(DKURS[0]) + 1)}{DKURS[1:]}", f"C{E0}", f"D{E0}", f"D{RET}"):
    wymagane(ref)
# data wjazdu w kolejnych etapach jest wymagana dopiero po wybraniu kraju
o_t = Side(style="thin", color=ORANGE)
kal.conditional_formatting.add(f"D{E0+1}:D{E1}", FormulaRule(formula=[f'$C{E0+1}<>""'], border=Border(left=o_t, right=o_t, top=o_t, bottom=o_t)))
kal["B3"] = "Legenda:"
kal["B3"].font = f_sub
kal["B3"].alignment = Alignment(horizontal="right", vertical="center")
for ref, txt_, styl in (("C3", "wymagane", "req"), ("D3", "do wpisania (opcjonalne)", "opt"), ("E3", "liczy się samo", "out"),
                        ("F3", "błąd w danych", "err")):
    c = kal[ref]
    c.value = txt_
    c.font = Font(name=FONT, size=8, color="0000FF" if styl in ("req", "opt") else ("9C0006" if styl == "err" else "000000"))
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = {"req": fill_in, "opt": fill_in, "out": PatternFill(fill_type=None), "err": PatternFill("solid", start_color="FFC7CE")}[styl]
    c.border = Border(left=o_m, right=o_m, top=o_m, bottom=o_m) if styl == "req" else box
kal.row_dimensions[3].height = 16

# --- podświetlenie komórek z błędem (czerwone; reguły dodane przed wyszarzaniem, więc mają pierwszeństwo) ---
p_err = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
f_err = Font(name=FONT, bold=True, color="9C0006")


def blad(rng, formula):
    kal.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=p_err, font=f_err))


nast = f'IF(AND($C{E0+1}<>"",$D{E0+1}<>""),$D{E0+1},$D${RET})'   # wjazd do kolejnego kraju albo powrót przez granicę
blad(f"C{E0}", f'$C{E0}=""')
blad(f"C{E0+1}:C{E1-1}", f'AND($C{E0+1}="",COUNTIF($C{E0+2}:$C${E1},"?*")>0)')   # pusty etap między krajami
blad(f"D{E0}:D{RET}", f'ISTEXT($D{E0})')
blad(f"I{E0}:J{E1}", f'ISTEXT(I{E0})')
for a in (WYJ, POW, DKURS):
    blad(f"{a}:{chr(ord(a[0]) + 1)}{a[1:]}", f'ISTEXT(${a[0]}${a[1:]})')
blad(f"{DPOL}:G{DPOL[1:]}", f'ISTEXT(${DPOL[0]}${DPOL[1:]})')
blad(f"E{E0}:G{E1}", f'AND($C{E0}<>"",N(E{E0})>IFERROR(ROUNDUP($I{R0},0),1E9))')   # więcej posiłków niż dób
blad(f"D{E0}:D{E1}", f'AND($C{E0}<>"",$D{E0}="")')
blad(f"D{E0+1}:D{E1}", f'AND($C{E0+1}<>"",$D{E0+1}<>"",$D{E0+1}<=$D{E0})')
blad(f"D{RET}", f'AND($C${E0}<>"",OR($D${RET}="",$D${RET}<=MAX($D${E0}:$D${E1})))')
blad(f"{DKURS}:{chr(ord(DKURS[0]) + 1)}{DKURS[1:]}", f'${DKURS[0]}${DKURS[1:]}=""')
blad(f"{WYJ}:{chr(ord(WYJ[0]) + 1)}{WYJ[1:]}", f'AND(${WYJ[0]}${WYJ[1:]}<>"",$D${E0}<>"",${WYJ[0]}${WYJ[1:]}>$D${E0})')
blad(f"{POW}:{chr(ord(POW[0]) + 1)}{POW[1:]}", f'AND(${POW[0]}${POW[1:]}<>"",$D${RET}<>"",${POW[0]}${POW[1:]}<$D${RET})')
blad(f"I{E0}:I{E1}", f'AND($I{E0}<>"",$D{E0}<>"",OR($I{E0}<$D{E0},$I{E0}>{nast}))')
blad(f"J{E0}:J{E1}", f'AND($J{E0}<>"",OR($J{E0}<IF($I{E0}<>"",$I{E0},$D{E0}),$J{E0}>{nast}))')
blad(f"M{E0}:M{RET}", f'$N{E0}="brak kursu"')
blad(f"D{KO0}:D{KO1}", f'$F{KO0}="brak kursu"')

# wyszarzanie nieużywanych etapów: wiersz uaktywnia się po wybraniu kraju w poprzednim
f_off = Font(name=FONT, color="BFBFBF")
p_off = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
kal.conditional_formatting.add(f"B{E0+1}:C{E1}", FormulaRule(formula=[f'$C{E0}=""'], fill=p_off, font=f_off))
kal.conditional_formatting.add(f"D{E0}:N{E1}", FormulaRule(formula=[f'$C{E0}=""'], fill=p_off, font=f_off))
kal.conditional_formatting.add(f"B{R0}:M{R1}", FormulaRule(formula=[f'$C{E0}=""'], font=f_off))
kal.freeze_panes = "A4"
kal.page_setup.orientation = "landscape"
kal.page_setup.fitToWidth = 1
kal.page_setup.fitToHeight = 0
kal.sheet_properties.pageSetUpPr.fitToPage = True
kal.print_area = f"A1:N{KL}"

# ---------------- Raport (wydruk) ----------------
rap.column_dimensions["A"].width = 2
for c, w in zip("BCDEFGHIJK", [18, 15.5, 15.5, 19, 9, 8.5, 10.5, 10.5, 10.5, 7]):
    rap.column_dimensions[c].width = w
DTR = "dd.mm.yyyy hh:mm"
f_rt = Font(name=FONT, size=14, bold=True)
f_rl = Font(name=FONT, size=10, bold=True)
f_rv = Font(name=FONT, size=10)
f_rs = Font(name=FONT, size=10, bold=True)
line = Border(bottom=Side(style="thin", color="D9D9D9"))
fill_rs = PatternFill("solid", start_color="F2F2F2")


def r_sec(row, text):
    rap.merge_cells(f"B{row}:K{row}")
    rap[f"B{row}"] = text
    rap[f"B{row}"].font = f_rs
    rap[f"B{row}"].fill = fill_rs
    rap[f"B{row}"].alignment = Alignment(horizontal="center", vertical="center")
    rap.row_dimensions[row].height = 22


def r_kv(row, label, formula, fmt=None, bold=False):
    rap.merge_cells(f"B{row}:D{row}")
    rap.merge_cells(f"E{row}:K{row}")
    rap[f"B{row}"] = label
    rap[f"B{row}"].font = f_rl
    rap[f"E{row}"].font = f_rl if bold else f_rv
    rap[f"E{row}"].alignment = Alignment(horizontal="right", vertical="center")
    if fmt:
        rap[f"E{row}"].number_format = fmt
    for c in "BCDEFGHIJK":
        rap[f"{c}{row}"].border = line
    rap.row_dimensions[row].height = 20
    F[f"Raport!E{row}"] = formula


rap.merge_cells("B2:K2")
rap["B2"] = "Rozliczenie diety z tytułu zagranicznej podróży służbowej"
rap["B2"].font = f_rt
rap.merge_cells("B3:K3")
rap["B3"] = "wg rozporządzenia MPiPS z 29.01.2013 (t.j. Dz.U. 2023 poz. 2190) · kursy średnie NBP (tabela A)"
rap["B3"].font = f_sub
rap.merge_cells("B4:K4")
rap["B4"].font = Font(name=FONT, size=10, bold=True, color="C00000")
F["Raport!B4"] = '=IF(Kalkulator!E1="OK","","UWAGA: "&Kalkulator!E1)'

kraje_rng = f"Kalkulator!C{E0}:C{E1}"
start, koniec = f"Kalkulator!D{E0}", f"Kalkulator!D{RET}"
r_sec(6, "Dane podróży")
r_kv(7, "Pracownik:", f'=IF(Kalkulator!{PRAC}="","",Kalkulator!{PRAC})')
r_kv(8, "Cel podróży:", f'=IF(Kalkulator!{CEL}="","",Kalkulator!{CEL})')
r_kv(9, "Kraj docelowy:", f'=IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER({kraje_rng},{kraje_rng}<>""))),"")')
r_kv(10, "Data i godzina rozpoczęcia:", f'=IF({start}="","",{start})', DTR)
r_kv(11, "Data i godzina zakończenia:", f'=IF({koniec}="","",{koniec})', DTR)
r_kv(12, "Czas podróży:", (
    f'=IF(OR({start}="",{koniec}=""),"",LET(m,ROUND(({koniec}-{start})*1440,0),d,INT(m/1440),'
    f'h,INT(MOD(m,1440)/60),mi,MOD(m,60),d&IF(d=1," dzień "," dni ")&h&" godz. "&mi&" min"))'))

r_sec(14, "Diety")
RH = 15
for c, l in zip("BCDEFGHIJK", ["Kraj", "Od", "Do", "Czas pobytu", "Dieta za dobę", "Liczba diet",
                                "Wysokość diet", "Potrącenie za posiłki", "Należna dieta", "Waluta"]):
    x = rap[f"{c}{RH}"]
    x.value = l
    x.font = Font(name=FONT, size=8, bold=True)
    x.alignment = wrap_c
    x.border = Border(bottom=Side(style="thin", color="808080"))
rap.row_dimensions[RH].height = 30
for i in range(ETAPY):
    r, src = RH + 1 + i, R0 + i
    cols = {"B": ("C", None), "C": ("F", DTR), "D": ("G", DTR), "F": ("E", "#,##0.00"),
            "G": ("I", "0.0000"), "J": ("K", "#,##0.00"), "K": ("D", None)}
    for c, (sc, fmt_) in cols.items():
        F[f"Raport!{c}{r}"] = f'=IF(Kalkulator!$C{src}="","",Kalkulator!{sc}{src})'
        if fmt_:
            rap[f"{c}{r}"].number_format = fmt_
    # czas pobytu w dniach i godzinach (minuty tylko, gdy niezerowe)
    F[f"Raport!E{r}"] = (f'=IF(Kalkulator!$C{src}="","",LET(m,ROUND(Kalkulator!H{src}*60,0),d,INT(m/1440),'
                         f'h,INT(MOD(m,1440)/60),mi,MOD(m,60),d&IF(d=1," dzień "," dni ")&h&" godz."&IF(mi>0," "&mi&" min","")))')
    F[f"Raport!H{r}"] = f'=IF(Kalkulator!$C{src}="","",ROUND(Kalkulator!E{src}*Kalkulator!I{src},2))'
    F[f"Raport!I{r}"] = f'=IF(Kalkulator!$C{src}="","",Kalkulator!J{src})'
    rap[f"H{r}"].number_format = rap[f"I{r}"].number_format = "#,##0.00"
    for c in "BCDEFGHIJK":
        rap[f"{c}{r}"].font = Font(name=FONT, size=9, bold=(c == "J"))
        rap[f"{c}{r}"].border = line
        rap[f"{c}{r}"].alignment = Alignment(horizontal="left" if c == "B" else "center", vertical="center")
    rap.row_dimensions[r].height = 18

# podsumowanie wg walut (wiersze sekcji 4 Kalkulatora)
VW = f"Kalkulator!$B${V0+2}:$B${V0+1+len(WALUTY)}"
VK = f"Kalkulator!$C${V0+2}:$C${V0+1+len(WALUTY)}"
VR = f"Kalkulator!$D${V0+2}:$D${V0+1+len(WALUTY)}"
dmy = lambda d: f'RIGHT("0"&DAY({d}),2)&"."&RIGHT("0"&MONTH({d}),2)&"."&YEAR({d})'
RK = RH + ETAPY + 2
r_sec(RK, "Koszty podróży")
r_kv(RK + 1, "Należne diety w walucie:", f'=IFERROR(TEXTJOIN("; ",TRUE,FILTER(FIXED({VK},2)&" "&{VW},{VK}<>0)),"")')
r_kv(RK + 2, "Kurs średni NBP:", (
    f'=IFERROR(TEXTJOIN("; ",TRUE,FILTER(LET(w,{VW},k,{VR},'
    f'd,XLOOKUP(w,{KURS_W},\'Kursy NBP\'!$D${KRS0}:$D${KRS1}),n,XLOOKUP(w,{KURS_W},\'Kursy NBP\'!$E${KRS0}:$E${KRS1}),'
    f'"1 "&w&" = "&IF(ISNUMBER(k),FIXED(k,IF(k<0.1,6,4),TRUE),"brak kursu")&" PLN (tab. "&n&" z "&{dmy("d")}&")"),{VK}<>0)),"")'))
r_kv(RK + 3, "Koszt całkowity w PLN:", f"=Kalkulator!M{RS}", '#,##0.00 "PLN"', bold=True)
rap[f"B{RK+3}"].font = Font(name=FONT, size=11, bold=True)
rap[f"E{RK+3}"].font = Font(name=FONT, size=11, bold=True)

RP = RK + 6
rap.merge_cells(f"B{RP}:K{RP}")
F[f"Raport!B{RP}"] = (f'="Kurs NBP z "&IF(Kalkulator!{TRYB}="w dniu daty kursu","dnia ","ostatniego dnia roboczego przed dniem ")'
                      f'&IF(Kalkulator!{DKURS}="","-",{dmy("Kalkulator!" + DKURS)})&". Dieta liczona wg §13–14 rozporządzenia, osobno dla każdego kraju."')
rap[f"B{RP}"].font = f_note
rap.merge_cells(f"B{RP+1}:K{RP+1}")
F[f"Raport!B{RP+1}"] = f'="Data sporządzenia: "&{dmy("TODAY()")}'
rap[f"B{RP+1}"].font = f_note

rap.print_area = f"A1:K{RP+1}"
rap.page_setup.paperSize = rap.PAPERSIZE_A4
rap.page_setup.orientation = "portrait"
rap.page_setup.fitToWidth = 1
rap.page_setup.fitToHeight = 1
rap.sheet_properties.pageSetUpPr.fitToPage = True
rap.print_options.horizontalCentered = True
rap.page_margins.left = rap.page_margins.right = 0.5
rap.page_margins.top = rap.page_margins.bottom = 0.6

# ---------------- Delegacja (druk: polecenie wyjazdu służbowego + rachunek kosztów podróży) ----------------
for c, w in zip("ABCDEFGHI", [1, 15, 11.5, 7.5, 15, 11.5, 7.5, 15, 14.5]):
    dele.column_dimensions[c].width = w
blk = Side(style="thin", color="000000")
f_d = Font(name=FONT, size=10)
f_ds = Font(name=FONT, size=7)
f_db = Font(name=FONT, size=10, bold=True)
DOTS = lambda n: ". " * n
K = lambda a: f"Kalkulator!${a[0]}${a[1:]}"
# wartość z Kalkulatora albo kropki, gdy puste
val = lambda ref, n: f'IF({ref}="","{DOTS(n)}",{ref})'
START = f'IF({K(WYJ)}="",Kalkulator!$D${E0},{K(WYJ)})'
KONIEC = f'IF({K(POW)}="",Kalkulator!$D${RET},{K(POW)})'
DOM = f'IF({K(MIEJ)}="","Polska",{K(MIEJ)})'
MIEJSCA = f'IF(Kalkulator!$H${E0}:$H${E1}<>"",Kalkulator!$H${E0}:$H${E1},Kalkulator!$C${E0}:$C${E1})'
N_ET = f'COUNTIF(Kalkulator!$C${E0}:$C${E1},"?*")'
# przyjazd do miejscowości (puste = wjazd do kraju); wyjazd z niej (puste = wjazd do kolejnego kraju / powrót przez granicę)
PRZYJ = f'IF(Kalkulator!$I${E0}:$I${E1}="",Kalkulator!$D${E0}:$D${E1},Kalkulator!$I${E0}:$I${E1})'
WYJAZ = (f'IF(Kalkulator!$J${E0}:$J${E1}="",IF(Kalkulator!$C${E0+1}:$C${RET}<>"",Kalkulator!$D${E0+1}:$D${RET},'
         f'Kalkulator!$D${RET}),Kalkulator!$J${E0}:$J${E1})')


def slownie(x, jedn=""):
    """Kwota słownie po polsku, np. 930,46 → „dziewięćset trzydzieści 46/100”; jedn wstawiana przed groszami."""
    J = '{"","jeden","dwa","trzy","cztery","pięć","sześć","siedem","osiem","dziewięć"}'
    N = ('{"dziesięć","jedenaście","dwanaście","trzynaście","czternaście","piętnaście","szesnaście",'
         '"siedemnaście","osiemnaście","dziewiętnaście"}')
    T = ('{"","","dwadzieścia","trzydzieści","czterdzieści","pięćdziesiąt","sześćdziesiąt","siedemdziesiąt",'
         '"osiemdziesiąt","dziewięćdziesiąt"}')
    S = '{"","sto","dwieście","trzysta","czterysta","pięćset","sześćset","siedemset","osiemset","dziewięćset"}'
    return ('LET(x,ROUND(ABS(' + x + '),2),z,INT(x),g,ROUND((x-z)*100,0),'
            'tri,LAMBDA(v,LET(h,INT(v/100),t,INT(MOD(v,100)/10),u,MOD(v,10),'
            'INDEX(' + S + ',h+1)&" "&IF(t=1,INDEX(' + N + ',u+1),INDEX(' + T + ',t+1)&" "&INDEX(' + J + ',u+1)))),'
            'grp,LAMBDA(v,jd,dw,wl,IF(v=0,"",IF(v=1,jd,tri(v)&" "&'
            'IF(AND(MOD(v,10)>=2,MOD(v,10)<=4,OR(MOD(v,100)<10,MOD(v,100)>20)),dw,wl)))),'
            's,TRIM(grp(INT(z/10^6),"jeden milion","miliony","milionów")&" "&'
            'grp(INT(MOD(z,10^6)/1000),"jeden tysiąc","tysiące","tysięcy")&" "&tri(MOD(z,1000))),'
            'IF(z=0,"zero",s)&" ' + (jedn + " " if jedn else "") + '"&RIGHT("0"&g,2)&"/100")')


def dtxt(ref, text=None, font=f_d, h="left", v="center", wrap=False, merge=None, formula=None):
    if merge:
        dele.merge_cells(merge)
    c = dele[ref]
    if text is not None:
        c.value = text
    if formula:
        F[f"Delegacja!{ref}"] = formula
    c.font = font
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    return c


def frame(rng, inner=False):
    """Obramowanie zakresu: tylko obwódka albo (inner=True) siatka."""
    rows = list(dele[rng if ":" in rng else f"{rng}:{rng}"])
    for i, row in enumerate(rows):
        for j, c in enumerate(row):
            if inner:
                c.border = Border(left=blk, right=blk, top=blk, bottom=blk)
            else:
                c.border = Border(left=blk if j == 0 else c.border.left, right=blk if j == len(row) - 1 else c.border.right,
                                  top=blk if i == 0 else c.border.top, bottom=blk if i == len(rows) - 1 else c.border.bottom)


def hgt(rows, h):
    for r in rows:
        dele.row_dimensions[r].height = h


# ostrzeżenie (drukuje się tylko przy błędnych danych)
dtxt("B1", font=Font(name=FONT, size=9, bold=True, color="C00000"), merge="B1:I1", formula=(
    f'=IF(Kalkulator!E1<>"OK","UWAGA: "&Kalkulator!E1,'
    f'IF(AND({K(WYJ)}<>"",{K(WYJ)}>Kalkulator!$D${E0}),"UWAGA: wyjazd z miejscowości jest po przekroczeniu granicy",'
    f'IF(AND({K(POW)}<>"",{K(POW)}<Kalkulator!$D${RET}),"UWAGA: powrót do miejscowości jest przed przekroczeniem granicy",'
    f'IF(COUNTIF(Kalkulator!$N${E0}:$N${RET},"brak kursu")+COUNTIF(Kalkulator!$F${KO0}:$F${KO1},"brak kursu")>0,'
    f'"UWAGA: brak kursu NBP dla waluty kosztu – sprawdź datę kursu",'
    f'IF(SUMPRODUCT((Kalkulator!$I${E0}:$I${E1}<>"")*(Kalkulator!$I${E0}:$I${E1}<Kalkulator!$D${E0}:$D${E1}))'
    f'+SUMPRODUCT((Kalkulator!$J${E0}:$J${E1}<>"")*((Kalkulator!$J${E0}:$J${E1}<IF(Kalkulator!$I${E0}:$I${E1}<>"",Kalkulator!$I${E0}:$I${E1},Kalkulator!$D${E0}:$D${E1}))'
    f'+(Kalkulator!$J${E0}:$J${E1}>IF(Kalkulator!$C${E0+1}:$C${RET}<>"",Kalkulator!$D${E0+1}:$D${RET},Kalkulator!$D${RET}))))>0,'
    f'"UWAGA: godziny przyjazdu/wyjazdu z miejscowości nie zgadzają się z przekroczeniami granic – zobacz czerwone pola w Kalkulatorze","")))))'))

# ===== strona 1: polecenie wyjazdu =====
dtxt("F2", "STWIERDZENIE POBYTU SŁUŻBOWEGO", Font(name=FONT, size=9), h="center", merge="F2:I2")
frame("F2:I2")
frame("F3:I21")
dtxt("B5", "Pieczątka wysyłającego", f_ds, h="center", v="bottom", merge="B5:E5")
hgt(range(2, 6), 24)
dtxt("B6", "POLECENIE WYJAZDU", Font(name=FONT, size=15, bold=True), h="center", merge="B6:E6")
dtxt("B7", font=Font(name=FONT, size=15, bold=True), h="center", merge="B7:E7",
     formula=f'="SŁUŻBOWEGO Nr "&{val(K(NR), 8)}')
hgt([6, 7], 26)
dtxt("B8", h="center", merge="B8:E8",
     formula=f'="z dnia "&IF({K(DPOL)}="","{DOTS(14)}",{dmy(K(DPOL))})')
dtxt("B10", merge="B10:E10", formula=f'="dla "&{val(K(PRAC), 40)}')
dtxt("B11", "(imię i nazwisko)", f_ds, h="center", v="top", merge="B11:E11")
dtxt("B12", merge="B12:E12", formula=f'={val(K(STAN), 44)}')
dtxt("B13", "(stanowisko służbowe)", f_ds, h="center", v="top", merge="B13:E13")
dtxt("B14", wrap=True, merge="B14:E14", formula=(
    f'="do "&IFERROR(TEXTJOIN(", ",TRUE,UNIQUE(FILTER({MIEJSCA},Kalkulator!$C${E0}:$C${E1}<>""))),"{DOTS(40)}")'))
dtxt("B15", merge="B15:E15", formula=(
    f'="na czas od "&IF(OR({START}="",{START}=0),"{DOTS(10)}",{dmy(START)})'
    f'&" do "&IF(OR({KONIEC}="",{KONIEC}=0),"{DOTS(10)}",{dmy(KONIEC)})'))
dtxt("B16", wrap=True, v="top", merge="B16:E16", formula=f'="w celu "&{val(K(CEL), 40)}')
hgt([9], 10)
hgt([10, 12, 15], 22)
hgt([11, 13], 12)
hgt([14], 34)
hgt([16], 48)
dtxt("B17", "środki lokomocji", v="top", merge="B17:E17")
dtxt("B18", merge="B18:E18", formula=f'={val(K(SROD), 30)}')
hgt([17, 18], 18)
dtxt("B20", f"{DOTS(7)}", h="center", v="bottom")
dtxt("C20", f"{DOTS(12)}", h="center", v="bottom", merge="C20:E20")
dtxt("B21", "data", f_ds, h="center", v="top")
dtxt("C21", "podpis wysyłającego", f_ds, h="center", v="top", merge="C21:E21")
hgt([19, 20], 20)
frame("B6:E16")
frame("B17:E18")
frame("B19:E21")
# linia cięcia
for c in "ABCDEFGHI":
    dele[f"{c}23"].border = Border(bottom=Side(style="dashed", color="000000"))
ZAL_T = f'IF(N({K(ZAL)})=0,"{DOTS(8)}",FIXED({K(ZAL)},2))'
ZAL_S = f'IF(N({K(ZAL)})=0,"{DOTS(30)}",{slownie(K(ZAL))})'
dtxt("B25", wrap=True, merge="B25:I25", formula=(
    f'="Proszę o wypłacenie zaliczki w kwocie zł "&{ZAL_T}&" słownie zł "&{ZAL_S}'))
dtxt("B26", merge="B26:I26", formula=(
    f'="na pokrycie wydatków zgodnie z poleceniem wyjazdu służbowego nr "&{val(K(NR), 10)}'))
hgt([25], 30)
hgt([26], 20)
dtxt("H28", DOTS(12), h="center", v="bottom", merge="H28:I28")
dtxt("H29", "Podpis delegowanego", f_ds, h="center", v="top", merge="H29:I29")
dtxt("B31", f"Zatwierdzono na zł {DOTS(8)} słownie zł {DOTS(40)}", merge="B31:I31")
dtxt("B32", f"do wypłaty z sum {DOTS(14)}", merge="B32:I32")
hgt([31, 32], 22)
for c, l in zip("BCDEF", ["Część", "Dział", "Rozdział", "§", "Poz."]):
    dtxt(f"{c}35", l, Font(name=FONT, size=8), h="center")
frame("B35:F36", inner=True)
dtxt("G34", "Konto", Font(name=FONT, size=8), h="center", merge="G34:H34")
dtxt("G35", "Wn", Font(name=FONT, size=8), h="center")
dtxt("H35", "Ma", Font(name=FONT, size=8), h="center")
dtxt("I34", "Nr dowodu", Font(name=FONT, size=8), h="center", merge="I34:I35")
frame("G34:I36", inner=True)
hgt([34, 35, 36], 16)
dtxt("F38", DOTS(7), h="center", v="bottom")
dtxt("G38", DOTS(14), h="center", v="bottom", merge="G38:I38")
dtxt("F39", "data", f_ds, h="center", v="top")
dtxt("G39", "podpisy sprawdzających", f_ds, h="center", v="top", merge="G39:I39")
dtxt("I40", "str. 1", Font(name=FONT, size=8), h="right")

# ===== strona 2: rachunek kosztów podróży =====
P2 = 41
dtxt(f"B{P2}", "RACHUNEK KOSZTÓW PODRÓŻY", Font(name=FONT, size=13, bold=True), h="center", merge=f"B{P2}:I{P2}")
hgt([P2], 30)
H1, H2 = P2 + 2, P2 + 3
f8 = Font(name=FONT, size=8)
dtxt(f"B{H1}", "W Y J A Z D", f8, h="center", merge=f"B{H1}:D{H1}")
dtxt(f"E{H1}", "P R Z Y J A Z D", f8, h="center", merge=f"E{H1}:G{H1}")
dtxt(f"H{H1}", "Środki lokomocji*)", f8, h="center", merge=f"H{H1}:H{H2}")
dtxt(f"I{H1}", "Koszty przejazdu", f8, h="center")
for c, l in zip("BCDEFGI", ["miejscowość", "data", "godz.", "miejscowość", "data", "godz.", "zł i gr"]):
    dtxt(f"{c}{H2}", l, f8, h="center")
T0, T1 = H2 + 1, H2 + 8
frame(f"B{H1}:I{T1}", inner=True)
hgt(range(T0, T1 + 1), 17)
for i in range(1, 9):
    r = T0 + i - 1
    L = (f'LET(n,{N_ET},i,{i},p,{MIEJSCA},dom,{DOM},pocz,{START},kon,{KONIEC},puste,OR(n=0,i>n+1),'
         f'prz,{PRZYJ},wyj,{WYJAZ},')
    # trasa: dom → miejscowość 1 → … → miejscowość n → dom
    wyj_dt = 'IF(i=1,pocz,INDEX(wyj,i-1))'
    prz_dt = 'IF(i<=n,INDEX(prz,i),kon)'
    F[f"Delegacja!B{r}"] = f'={L}IF(puste,"",IF(i=1,dom,INDEX(p,i-1))))'
    F[f"Delegacja!E{r}"] = f'={L}IF(puste,"",IF(i<=n,INDEX(p,i),dom)))'
    for cd, ch, e in (("C", "D", wyj_dt), ("F", "G", prz_dt)):
        F[f"Delegacja!{cd}{r}"] = f'={L}v,IF(puste,0,N({e})),IF(v=0,"",INT(v)))'
        F[f"Delegacja!{ch}{r}"] = f'={L}v,IF(puste,0,N({e})),IF(v=0,"",MOD(v,1)))'
        dele[f"{cd}{r}"].number_format = "dd.mm.yyyy"
        dele[f"{ch}{r}"].number_format = "hh:mm"
    idx = f"IF(i<=n,i,{ETAPY + 1})"
    F[f"Delegacja!H{r}"] = f'={L}s,IF(puste,"",INDEX(Kalkulator!$K${E0}:$K${RET},{idx})&""),IF(puste,"",IF(s="",{K(SROD)}&"",s)))'
    F[f"Delegacja!I{r}"] = f'={L}IF(puste,"",INDEX(Kalkulator!$N${E0}:$N${RET},{idx})))'
    dele[f"I{r}"].number_format = M2
    for c in "BCDEFGHI":
        dele[f"{c}{r}"].font = Font(name=FONT, size=8)
        dele[f"{c}{r}"].alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)

# zestawienie kosztów (prawa strona) i pola sprawdzających (lewa strona)
Z = T1 + 1
pozycje = ["Ryczałty za dojazdy", "Dojazdy udokumentowane", "Razem przejazdy, dojazdy", "Diety",
           "Noclegi wg rachunków", "Noclegi – ryczałt", "Inne wydatki wg załączników"]
for i, l in enumerate(pozycje):
    r = Z + i
    dtxt(f"F{r}", l, f8, merge=f"F{r}:H{r}")
    frame(f"F{r}:H{r}")
    dele[f"I{r}"].number_format = M2
    dele[f"I{r}"].font = f8
    frame(f"I{r}")
hgt(range(Z, Z + 7), 15)
RAZ, DIET, OG = Z + 2, Z + 3, Z + 7
F[f"Delegacja!I{RAZ}"] = f"=SUM(I{T0}:I{T1},I{Z},I{Z+1})"
F[f"Delegacja!I{DIET}"] = f"=Kalkulator!M{RS}"
for r_, poz in ((Z, "Ryczałty za dojazdy"), (Z + 1, "Dojazdy udokumentowane"), (Z + 4, "Noclegi wg rachunków"),
                (Z + 5, "Noclegi – ryczałt"), (Z + 6, "Inne wydatki wg załączników")):
    F[f"Delegacja!I{r_}"] = f'=IF(Kalkulator!F{KO[poz]}="","",Kalkulator!F{KO[poz]})'
dele[f"I{DIET}"].font = Font(name=FONT, size=8, bold=True)
dtxt(f"B{Z}", "Rachunek sprawdzono pod względem", f8, h="center", merge=f"B{Z}:E{Z+1}")
dtxt(f"B{Z+2}", "Merytorycznym", f8, h="center", merge=f"B{Z+2}:C{Z+3}")
dtxt(f"D{Z+2}", "Formalnym\ni rachunkowym", f8, h="center", wrap=True, merge=f"D{Z+2}:E{Z+3}")
for a, b in (("B", "C"), ("D", "E")):
    dtxt(f"{a}{Z+5}", f"{DOTS(5)}   {DOTS(6)}", f8, h="center", v="bottom", merge=f"{a}{Z+5}:{b}{Z+5}")
    dtxt(f"{a}{Z+6}", "Data            podpis", f_ds, h="center", v="top", merge=f"{a}{Z+6}:{b}{Z+6}")
    frame(f"{a}{Z+2}:{b}{Z+3}")
    frame(f"{a}{Z+4}:{b}{Z+6}")
frame(f"B{Z}:E{Z+1}")
# Ogółem + słownie
dtxt(f"F{OG}", "Ogółem", Font(name=FONT, size=9, bold=True), h="right", merge=f"F{OG}:H{OG}")
dtxt(f"F{OG+1}", "Słownie złotych:", f8, v="top", merge=f"F{OG+1}:H{OG+5}")
dtxt(f"I{OG}", font=Font(name=FONT, size=10, bold=True), h="right", v="top",
     formula=f"=SUM(I{RAZ},I{DIET},I{DIET+1}:I{DIET+3})")
dtxt(f"I{OG+1}", font=Font(name=FONT, size=10), h="center", v="top", wrap=True, merge=f"I{OG+1}:I{OG+5}",
     formula=f"={slownie(f'I{OG}', 'złotych i')}")
dele[f"I{OG}"].number_format = M2
frame(f"F{OG}:H{OG+5}")
frame(f"I{OG}:I{OG+5}")
dtxt(f"B{OG}", f"Zatwierdzono na zł {DOTS(10)}", f8, h="center", merge=f"B{OG}:E{OG}")
dtxt(f"B{OG+1}", f"słownie {DOTS(32)}", f8, merge=f"B{OG+1}:E{OG+1}")
dtxt(f"B{OG+2}", DOTS(40), f8, merge=f"B{OG+2}:E{OG+2}")
dtxt(f"B{OG+4}", f"{DOTS(7)}        {DOTS(16)}", f8, h="center", v="bottom", merge=f"B{OG+4}:E{OG+4}")
dtxt(f"B{OG+5}", "data                              podpisy zatwierdzających", f_ds, h="center", v="top", merge=f"B{OG+5}:E{OG+5}")
frame(f"B{OG}:E{OG+5}")
hgt(range(OG, OG + 6), 17)
# zaliczka / do wypłaty
ZW = OG + 6
dtxt(f"F{ZW}", "Załączam", f8, v="top")
dtxt(f"F{ZW+1}", font=f_ds, h="center", v="bottom", formula=f'=IF(N({K(DOW)})=0,"{DOTS(5)}",{K(DOW)})&" dowodów"')
dtxt(f"G{ZW}", "Pobrano zaliczkę", f8, merge=f"G{ZW}:H{ZW}")
dtxt(f"G{ZW+1}", "Do wypłaty - zwrotu", f8, merge=f"G{ZW+1}:H{ZW+1}")
F[f"Delegacja!I{ZW}"] = f'=IF(N({K(ZAL)})=0,"",{K(ZAL)})'
F[f"Delegacja!I{ZW+1}"] = f"=I{OG}-N(I{ZW})"
for r in (ZW, ZW + 1):
    dele[f"I{r}"].number_format = M2
    dele[f"I{r}"].font = Font(name=FONT, size=9, bold=(r == ZW + 1))
    frame(f"G{r}:H{r}")
    frame(f"I{r}")
frame(f"F{ZW}:F{ZW+1}")
dtxt(f"F{ZW+2}", "Niniejszy rachunek przedkładam", f8, v="top", merge=f"F{ZW+2}:I{ZW+2}")
dtxt(f"F{ZW+4}", DOTS(10), f8, h="center", v="bottom", merge=f"F{ZW+4}:G{ZW+4}")
dtxt(f"H{ZW+4}", DOTS(12), f8, h="center", v="bottom", merge=f"H{ZW+4}:I{ZW+4}")
dtxt(f"F{ZW+5}", "data", f_ds, h="center", v="top", merge=f"F{ZW+5}:G{ZW+5}")
dtxt(f"H{ZW+5}", "podpis", f_ds, h="center", v="top", merge=f"H{ZW+5}:I{ZW+5}")
frame(f"F{ZW+2}:I{ZW+5}")
dtxt(f"B{ZW+1}", f"Kwituję odbiór zł {DOTS(12)}", f8, merge=f"B{ZW+1}:E{ZW+1}")
dtxt(f"B{ZW+2}", f"Słownie zł {DOTS(30)}", f8, merge=f"B{ZW+2}:E{ZW+2}")
dtxt(f"B{ZW+3}", DOTS(40), f8, merge=f"B{ZW+3}:E{ZW+3}")
dtxt(f"B{ZW+4}", f"{DOTS(8)}        {DOTS(14)}", f8, h="center", v="bottom", merge=f"B{ZW+4}:E{ZW+4}")
dtxt(f"B{ZW+5}", "data                              podpis", f_ds, h="center", v="top", merge=f"B{ZW+5}:E{ZW+5}")
frame(f"B{ZW}:E{ZW+5}")
hgt(range(ZW, ZW + 6), 17)
# oświadczenie o zaliczce
OS = ZW + 7
dtxt(f"B{OS}", wrap=True, v="top", merge=f"B{OS}:I{OS+1}", formula=(
    f'="Zaliczkę w kwocie zł "&{ZAL_T}&"   słownie zł "&{ZAL_S}&"   otrzymałem i zobowiązuję się rozliczyć z niej '
    f'w terminie "&IF(N({K(TERM)})=0,". . . .",{K(TERM)})&" dni po zakończeniu podróży, upoważniając równocześnie '
    f'zakład pracy do potrącenia kwoty nierozliczonej zaliczki z najbliższej wypłaty wynagrodzenia."'))
hgt([OS, OS + 1], 22)
dtxt(f"B{OS+3}", h="center", v="bottom", merge=f"B{OS+3}:D{OS+3}", formula=f'={val(K(PRAC), 24)}')
dtxt(f"G{OS+3}", DOTS(22), h="center", v="bottom", merge=f"G{OS+3}:I{OS+3}")
dtxt(f"B{OS+4}", "imię i nazwisko delegowanego", f_ds, h="center", v="top", merge=f"B{OS+4}:D{OS+4}")
dtxt(f"G{OS+4}", "data i podpis delegowanego", f_ds, h="center", v="top", merge=f"G{OS+4}:I{OS+4}")
hgt([OS + 3], 28)
dtxt(f"B{OS+6}", ("*) Wymienić środek lokomocji, klasę, rodzaj biletu (bezpłatny, ulgowy, normalny). W podróżach przy użyciu "
                  "własnych środków lokomocji (samochód, motocykl, motorower) podać również ilość km i stawkę za 1 km."),
     f_ds, wrap=True, v="top", merge=f"B{OS+6}:I{OS+7}")
DL = OS + 9
dtxt(f"I{DL}", "str. 2", Font(name=FONT, size=8), h="right")

from openpyxl.worksheet.pagebreak import Break
dele.row_breaks.append(Break(id=P2 - 1))
dele.print_area = f"A1:I{DL}"
dele.page_setup.paperSize = dele.PAPERSIZE_A4
dele.page_setup.orientation = "portrait"
dele.page_setup.scale = 84
dele.print_options.horizontalCentered = True
dele.page_margins.left = dele.page_margins.right = 0.4
dele.page_margins.top = 0.5
dele.page_margins.bottom = 0.4

# ---------------- Instrukcja ----------------
ins.column_dimensions["A"].width = 2
ins.column_dimensions["B"].width = 120
txt = [
    ("Jak korzystać z kalkulatora diety zagranicznej", f_title),
    (f"Wersja {WERSJA} z {WERSJA_DATA}: kurs ręczny awaryjny; kontrola pustych etapów, posiłków i dat; druk delegacji (polecenie wyjazdu + rachunek kosztów), "
     "koszty przejazdów, noclegów i ryczałtów, kafelki z wynikiem, oznaczenie pól wymaganych i błędów, ochrona arkuszy.", f_sub),
    ("", None),
    ("1. Arkusz „Kalkulator” – wypełniaj tylko żółte pola z niebieskim tekstem. Pozostałe komórki liczą się same.", None),
    ("2. Sekcja 2: każdy kraj to osobny etap. Wpisz chwilę przekroczenia granicy (samolot: start z ostatniego lotniska w kraju), np. 14.09.2026 10:30.", None),
    ("    Etap kończy się wjazdem do kolejnego kraju. W wierszu „Powrót” wpisz przekroczenie granicy PL (samolot: lądowanie w kraju).", None),
    ("3. Waluta i stawka diety ustawiają się automatycznie po wybraniu kraju (arkusz „Kraje”).", None),
    ("4. Kursy średnie NBP pobierają się automatycznie z api.nbp.pl (arkusz „Kursy NBP”). Wymagany internet i Excel z Microsoft 365 lub Excel 2024 na Windows.", None),
    ("    Plik pobrany z internetu otwiera się w Widoku chronionym – kliknij „Włącz edytowanie” (i „Włącz zawartość”, jeśli Excel zapyta), inaczej kursy się nie pobiorą.", None),
    ("    Jeśli kursy się nie odświeżyły: Ctrl+Alt+F9.", None),
    ("    Gdy kursy nie pobierają się wcale (np. sieć firmowa blokuje api.nbp.pl): w arkuszu „Kursy NBP” wpisz kurs ręczny z datą notowania i nr tabeli (kolumny F–H) – zastąpi kurs z NBP.", None),
    ("5. „Kontrola danych” (góra arkusza) pokazuje OK albo opis błędu w danych. Kafelki obok pokazują kwotę do wypłaty (zawsze dieta), wydatki pokrywane z zaliczki na wydatki i ogółem do rozliczenia.", None),
    ("    Arkusze są chronione bez hasła – można wpisywać tylko w żółte pola. Żeby zmienić coś innego: Recenzja → Nie chroń arkusza.", None),
    ("6. Wydruk: przejdź do arkusza „Raport” i naciśnij Ctrl+P (albo Plik → Eksportuj → PDF). Raport mieści się na jednej stronie A4.", None),
    ("7. Druk delegacji: arkusz „Delegacja” (polecenie wyjazdu + rachunek kosztów podróży, 2 strony A4) wypełnia się sam z Kalkulatora.", None),
    ("    Godziny przyjazdu do miejscowości i wyjazdu z niej: kolumny I–J etapów. Koszty przejazdu: kolumny K–N (wiersz „Powrót” = przejazd powrotny); ryczałty, noclegi i inne wydatki: sekcja 5 (kwoty w walucie przeliczają się po kursie NBP).", None),
    ("    Dodatkowe dane (nr polecenia, stanowisko, miejscowość wyjazdu, środki lokomocji, zaliczka) – sekcja 1 Kalkulatora; miejscowości – kolumna H etapów.", None),
    ("", None),
    ("Zasady obliczeń (rozporządzenie MPiPS z 29.01.2013, t.j. Dz.U. 2023 poz. 2190)", f_bold),
    ("• Czas (§12): od przekroczenia granicy w drodze za granicę do przekroczenia jej w drodze powrotnej (samolot: start–lądowanie).", None),
    ("• Dieta (§13): za każdą pełną dobę 100%; niepełna doba: do 8 h – 1/3, 8–12 h – 50%, ponad 12 h – 100%. Liczona osobno dla każdego etapu.", None),
    ("• Posiłki (§14): śniadanie –15%, obiad –30%, kolacja –30% diety.", None),
    ("• Kurs: średni kurs NBP (tabela A) – z dnia poprzedzającego datę kursu albo z dnia daty kursu (§20); przy dniu wolnym brany jest ostatni wcześniejszy kurs.", None),
    ("", None),
    ("Aktualizacja stawek: diety – arkusz „Kraje”, procenty za posiłki – arkusz „Parametry”.", f_bold),
    ("", None),
    ("Licencja MIT · Copyright (c) 2026 Krystian Rutkowski. Wolno używać, kopiować, zmieniać i rozpowszechniać, "
     "także komercyjnie, z zachowaniem tej informacji. Oprogramowanie dostarczane „tak jak jest”, bez żadnej gwarancji.", f_sub),
]
for i, (t, fnt) in enumerate(txt, start=1):
    ins[f"B{i}"] = t
    ins[f"B{i}"].font = fnt or f_norm

for ws in wb.worksheets:
    for mr in ws.merged_cells.ranges:
        if not ws.cell(mr.min_row, mr.min_col).protection.locked:
            for row in ws.iter_rows(min_row=mr.min_row, max_row=mr.max_row, min_col=mr.min_col, max_col=mr.max_col):
                for c in row:
                    c.protection = Protection(locked=False)
wb._sheets = [kal, rap, dele, ins, kur, kra, par]
wb.active = 0
wb.save(OUT)
json.dump(F, open(os.path.join(HERE, "formulas.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
json.dump({"KRS0": KRS0, "KRS1": KRS1, "E0": E0, "RET": RET, "R0": R0, "RS": RS, "V0": V0, "VL": VL, "KO0": KO0, "KT": KT,
           **{k: v for k, v in dict(PRAC=PRAC, STAN=STAN, CEL=CEL, NR=NR, DPOL=DPOL, MIEJ=MIEJ, SROD=SROD, WYJ=WYJ,
                                    POW=POW, DKURS=DKURS, TRYB=TRYB, ZAL=ZAL, TERM=TERM, DOW=DOW).items()}},
          open(os.path.join(HERE, "layout.json"), "w", encoding="utf-8"), indent=0)
print("ok", len(F), "formulas; R0", R0, "RS", RS, "V0", V0, "VL", VL)

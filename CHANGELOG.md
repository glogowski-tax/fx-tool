# CHANGELOG — FX_TOOL

## 2026-06-14 (14)
- **Kurs celny: wybór miesiąca z interfejsu (plik bez dat).** W trybie kursu celnego doszedł przełącznik „Skąd miesiąc kursu celnego": *Z kolumny z datami* (jak w (13)) albo *Wybierz miesiąc ręcznie*. Przy ręcznym wyborze wskazujesz miesiąc + rok, a kurs celny (przedostatnia środa miesiąca poprzedniego) stosowany jest do **wszystkich wierszy** — **kolumna z datami nie jest potrzebna** (główna zaleta kursu miesięcznego: plik wejściowy nie musi zawierać dat).
  - Gdy brak kolumny dat, kolumna z kursem (+ „PLN przeliczone") jest **dopisywana na końcu** arkusza; nagłówek zawiera miesiąc obowiązywania, np. `kurs celny NBP EUR/PLN (czerwiec 2026)`.
  - Działa łącznie z wyborem waluty (jedna / per-wiersz) i z PLN (kurs 1). `process_workbook` przyjmuje `fixed_ref_date`; `col_idx` może być `None`.
  - Pod maską: stała `POLISH_MONTHS`, picker miesiąc/rok (domyślnie bieżący), podgląd „kurs na <miesiąc> = kurs NBP z <data>".

## 2026-06-14 (13)
- **Kurs celny — stały na miesiąc (zasady celne).** Nowy przełącznik „Podstawa kursu": *Kurs z dnia poprzedzającego (zasady VAT)* — dotychczasowe działanie — albo *Kurs celny — stały na miesiąc (zasady celne)*. Intrastat dopuszcza obie metody (instrukcja GUS + rozporządzenie MF / UKC art. 53, 146).
  - Kurs celny = kurs NBP tabeli A z **przedostatniej środy miesiąca poprzedzającego** miesiąc transakcji; obowiązuje przez **cały miesiąc kalendarzowy**. Aplikacja dobiera go **per miesiąc daty transakcji** (plik obejmujący kilka miesięcy → każdy miesiąc swój kurs; plik jednomiesięczny → jeden kurs dla wszystkich wierszy).
  - **Zawsze z NBP** — EBC nie publikuje kursu celnego (opcja źródła ECB ukryta w tym trybie). Działa też dla wielu walut (każda ma kurs z tej samej środy) i z PLN (kurs 1). Nagłówek kolumny: `kurs celny NBP …/PLN`.
  - Jeśli przedostatnia środa była dniem wolnym (brak notowania) — kurs z ostatniego dnia roboczego przed nią (`rate_on_or_before`).
  - Weryfikacja na żywych danych: kurs celny czerwca 2026 (EUR 4,2546 / USD 3,6709 / GBP 4,9126) = NBP tab. A z 2026-05-20. Zgodne.
  - **Znane ograniczenie (świadome):** nie obsługujemy wyjątku UKC o korekcie kursu w trakcie miesiąca przy wahaniach ≥5% (art. 146 UKC-RW) — rzadki przypadek; kurs traktujemy jako stały na cały miesiąc.

## 2026-06-14 (12)
- **Waluta z kolumny (per wiersz).** Nowy przełącznik „Sposób doboru waluty": *Jedna waluta dla całego pliku* (dotychczasowe działanie) lub *Waluta z kolumny (per wiersz)* — wskazujesz kolumnę z kodem/nazwą waluty i dla **każdego wiersza** kurs dobierany jest wg jego waluty. Kolumna z walutą auto-wykrywana (nazwa zawiera „waluta"/„currency").
  - `normalize_currency()` rozpoznaje kody ISO (EUR, USD…) i częste nazwy PL/EN (`euro`, `dolar`, `funt`, `frank`, `zł`…); nierozpoznane → status **„Nieznana waluta"** (wiersz bez przeliczenia, bez crasha).
  - **PLN** obsłużone specjalnie: kurs 1, kwota PLN = kwota oryginalna, bez zapytania do API (NBP nie ma endpointu PLN/PLN).
  - Kursy pobierane **hurtowo, jedno zapytanie na każdą napotkaną walutę** (nie per wiersz). **ECB tylko dla EUR** — w trybie per-wiersz wybór źródła dotyczy wyłącznie wierszy w EUR, pozostałe zawsze z NBP. Nagłówek kolumny kursu bez kodu waluty (`kurs NBP/PLN`).
  - **Pre-skan** przed generowaniem: lista wykrytych walut z licznościami i ostrzeżenie o nierozpoznanych wartościach (wzorowane na skanie kodów CN).

## 2026-06-13 (11)
- **Ikonki pomocy „?" przy każdej opcji.** Po najechaniu myszką rozwija się opis działania funkcji (Streamlit `help=`). Dodane do: wyboru pliku, arkusza, kolumny z datami, waluty, źródła kursu (NBP/EBC), kolumny z kwotami, kolumny z kodem kraju. VAT, CN i ucinanie tekstu miały już opis.

## 2026-06-13 (10)
- **Walidacja CN także dla kodów 10-cyfrowych (TARIC).** Dotychczas tylko 8 cyfr; 10-cyfrowe wpadały w „błędny format". Teraz kod 10-cyfrowy walidowany po **prefiksie 8 cyfr** (= kod CN wewnątrz TARIC). Nowe statusy informujące, że sprawdzono prefiks: `OK (10-cyfr → sprawdzono prefiks 8)` oraz `nieaktualny (10-cyfr → prefiks 8 poza edycją)`. Kody 8-cyfrowe bez zmian.
- `normalize_cn`: uzupełnianie zgubionego przez Excel wiodącego zera rozszerzone na 10-cyfrowe (9 → 10, analogicznie do 7 → 8); float bez części ułamkowej (np. `84713000.0`) sprowadzany do int, by „.0" nie fałszowało liczby cyfr.
- Refaktor spójności: jedna funkcja `cn_is_valid(status)` jako wspólna bramka „poprawny / nie" we wszystkich trzech miejscach (early-return `cn_outcome`, akceptacja zamiennika, lista niepoprawnych w UI) — UI i zapisany plik nie mogą się rozjechać. Zamiennik 10-cyfrowy też akceptowany.

## 2026-06-10 (9)
- **Obsługa plików CSV** obok Excela. Uploader przyjmuje `xlsx` i `csv`. Wczytywanie CSV: auto-wykrywanie kodowania (`utf-8-sig` → `cp1250` → `latin-1`) i separatora (`;` / tab / `,`). Wartości konwertowane sensownie: czyste liczby → int, przecinek dziesiętny `1234,56` → float, **kody z wiodącym zerem zostają tekstem** (nie psuje CN/ID). Eksport: jeśli wejście było CSV → wynik też CSV (UTF-8 z BOM, separator dziesiętny dopasowany do separatora pól); jeśli Excel → po staremu xlsx. Bez nowych zależności (stdlib `csv`/`io`).

## 2026-06-10 (8)
- Przeliczanie PLN: powrót do **pełnych złotych** (round bez miejsc po przecinku), format komórki `0` (bez zer po przecinku, bez separatora tysięcy). Cofnięcie zmiany z (1)/(2) na życzenie.

## 2026-06-10 (7)
- Korekta CN: zły zamiennik (nadal nieaktualny / błędny format) NIE jest zapisywany — w komórce zostaje oryginalny kod i jego status. Tylko poprawny zamiennik podmienia kod (status "poprawiony").

## 2026-06-10 (6)
- **Interaktywna korekta kodów CN.** Po wskazaniu kolumny CN apka od razu listuje niepoprawne kody (z liczbą wystąpień). Obok każdego pole na kod zastępczy — wpisany kod podmienia błędny w kolumnie CN. Przycisk „🔁 Sprawdź ponownie kody" przelicza statusy na żywo; obok „Pozostało niepoprawnych / poprawionych". Drugą opcją jest zwykły zapis pliku. Status CN po operacji: `OK` / `nieaktualny` (niepoprawiony lub poprawiony na inny zły) / `poprawiony` / `błędny format`.
- Wspólna funkcja `cn_outcome(oryginał, zamiennik, lista) -> (status, wartość)` używana i w UI, i przy zapisie — podgląd nie może rozjechać się z plikiem.
- Hardening: odczyt pliku przez `getvalue()` (odporne na częste przeładowania UI).

## 2026-06-10 (5)
- Nowa opcja (sekcja Fluiconnecto): **walidacja kodów CN** względem obowiązującej edycji **CN 2026**. Z kolumny z kodami CN powstaje nowa kolumna ze statusem: `OK` / `nieaktualny` / `błędny format`. Obsługa zera wiodącego (Excel zapisuje 01012100 jako liczbę 1012100 — uzupełniane). Lista 9791 kodów wbudowana w repo (`cn_2026.txt`), źródło: GUS (stat.gov.pl), ładowana raz (cache). Rok edycji widoczny w UI i nagłówku kolumny.
  - **UWAGA (utrzymanie):** lista CN aktualizowana co roku — w styczniu pobrać nowy plik CN z GUS, wygenerować `cn_<rok>.txt`, podbić `CN_EDITION_YEAR` w app.py.

## 2026-06-10 (4)
- Nowa opcja: **kod kraju z numeru VAT** — z kolumny z numerami VAT kontrahentów wyciągany 2-literowy prefiks kraju do nowej kolumny obok (`... (kod kraju)`). `EL` (prefiks VAT Grecji) → `GR`; pozostałe zgodne z ISO 3166. `XI` (Irlandia Płn.) zachowane. Brak czytelnego prefiksu → pusta komórka. Kolumna VAT auto-wykrywana (nazwa zawiera „vat").
- **Front-end:** opcje specyficzne dla klienta **Fluiconnecto** (normalizacja kraju, kod kraju z VAT, ucinanie wielolinijkowego tekstu) zgrupowane w wyróżnionej, obramowanej sekcji „🔶 Obróbka dla Fluiconnecto" — operator od razu widzi, że to obróbka kliencka.
- Refaktor: kolumny pochodne (kraj, VAT) wstawiane przez wspólny helper z bezpieczną korektą wszystkich indeksów (eliminuje błędy przesunięcia przy wielu wstawianych kolumnach).

## 2026-06-10 (3)
- Nowa opcja (checkbox, domyślnie włączony): **ucinanie wielolinijkowego tekstu** do pierwszej linii we wszystkich komórkach arkusza. Komórki ze złamaniem linii (Alt+Enter) skracane do pierwszej linii; jednolinijkowe bez zmian. Przydatne do Intrastat.

## 2026-06-10 (2)
- Nowa opcja: **normalizacja kodów krajów** (kolumna `OrigCountryRegionId`) z formatu 3-literowego (ISO 3166 alpha-3, np. `ITA`, `DEU`, `BLR`) na 2-literowy (alpha-2, np. `IT`, `DE`, `BY`). Wynik trafia do **nowej kolumny obok** (`... (ISO-2)`), oryginał zostaje. Pełna mapa ISO (249 krajów). Kolumna wybierana w UI, z auto-wykryciem `OrigCountryRegionId`. Nieznane/niepasujące wartości przepisywane bez zmian. Wartości, nie formuły (zgodnie z konwencją projektu).

## 2026-06-10
- Przeliczanie na PLN: wynik mnożenia zaokrąglany teraz do **dwóch miejsc po przecinku** (wcześniej do pełnych złotych) i zapisywany w komórce z **formatem liczbowym** `0.00` (dwa miejsca po przecinku, bez separatora tysięcy; zamiast formatu „Ogólne").

## 2026-06-10 (7)
- Korekta CN: zły zamiennik (nadal nieaktualny / błędny format) NIE jest zapisywany — w komórce zostaje oryginalny kod i jego status. Tylko poprawny zamiennik podmienia kod (status "poprawiony").

# CHANGELOG — FX_TOOL

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

# CHANGELOG — FX_TOOL

## 2026-06-10 (2)
- Nowa opcja: **normalizacja kodów krajów** (kolumna `OrigCountryRegionId`) z formatu 3-literowego (ISO 3166 alpha-3, np. `ITA`, `DEU`, `BLR`) na 2-literowy (alpha-2, np. `IT`, `DE`, `BY`). Wynik trafia do **nowej kolumny obok** (`... (ISO-2)`), oryginał zostaje. Pełna mapa ISO (249 krajów). Kolumna wybierana w UI, z auto-wykryciem `OrigCountryRegionId`. Nieznane/niepasujące wartości przepisywane bez zmian. Wartości, nie formuły (zgodnie z konwencją projektu).

## 2026-06-10
- Przeliczanie na PLN: wynik mnożenia zaokrąglany teraz do **dwóch miejsc po przecinku** (wcześniej do pełnych złotych) i zapisywany w komórce z **formatem liczbowym** `0.00` (dwa miejsca po przecinku, bez separatora tysięcy; zamiast formatu „Ogólne").

# Generator dyzurow pracownikow (OR-Tools CP-SAT)

Wersja v1.2 jest reuzywalna: dane sa w Excelu, walidator tworzy raport diagnostyczny,
a kod jest rozbity na moduly (loader/walidator/solver).

## Pliki
- `generator_dyzurow_czerwiec_2026.py` - glowny runner CLI,
- `harmonogram/excel_io.py` - odczyt/zapis Excel + raport diagnostyczny,
- `harmonogram/validator.py` - walidacje techniczne i biznesowe,
- `harmonogram/solver.py` - model CP-SAT,
- `harmonogram/sample_data.py` - przykladowa konfiguracja,
- `harmonogram/models.py` - dataclass i stale,
- `requirements.txt` - zaleznosci Pythona,
- `dane_wejsciowe_czerwiec_2026.xlsx` - dane wejsciowe (tworzone automatycznie, jesli brak),
- `yyyy.MM - Lista Zastępstw.xlsx` - wynikowy plik koncowy.

## Instalacja
```bash
python3 -m pip install -r requirements.txt
```

## Szybki start
Pierwsze uruchomienie wygeneruje przykladowy plik danych wejsciowych:

```bash
python3 generator_dyzurow_czerwiec_2026.py
```

Nastepnie edytuj `dane_wejsciowe_czerwiec_2026.xlsx` i uruchom ponownie ten sam command.

## Przydatne komendy
```bash
python3 generator_dyzurow_czerwiec_2026.py --create-template template_dane.xlsx
python3 generator_dyzurow_czerwiec_2026.py --export-sample-config dane_wejsciowe_czerwiec_2026.xlsx
python3 generator_dyzurow_czerwiec_2026.py --config dane_wejsciowe_czerwiec_2026.xlsx --validate-only
python3 generator_dyzurow_czerwiec_2026.py --config dane_wejsciowe_czerwiec_2026.xlsx --validate-only --validation-report raport_bledow.xlsx
python3 generator_dyzurow_czerwiec_2026.py --config dane_wejsciowe_czerwiec_2026.xlsx --output "2026.06 - Lista Zastępstw.xlsx"
python3 generator_dyzurow_czerwiec_2026.py --config dane_wejsciowe_czerwiec_2026.xlsx --output-template "2026.06 - Lista Zastępstw.xlsx"
python3 generator_dyzurow_czerwiec_2026.py --use-built-in-june-config
```

## Walidator v1.2
Przed uruchomieniem solvera skrypt sprawdza m.in.:
- obecność wymaganych arkuszy,
- poprawność nagłówków,
- duplikaty pracowników,
- poprawność dat i zakresów urlopów,
- poprawność parametrów solvera i wag celu,
- konflikty wymuszeń,
- podstawowe konflikty biznesowe jeszcze przed solverem.

W przypadku błędów skrypt:
- wypisuje listę problemów z nazwą arkusza i numerem wiersza,
- zapisuje raport `*_diagnostyka.xlsx` (domyslnie obok pliku wejsciowego),
- koloruje `BLAD` na czerwono i `OSTRZEZENIE` na zolto.

## Struktura danych w Excelu
Skrypt odczytuje arkusze:
- `Meta` (`year`, `month`),
- `Pracownicy` (`Pracownik`, `Bez_glownego`, `Preferuj_mniej_dyzurow`),
- `Swieta` (`Data`),
- `Niedostepnosc_tygodniowa` (`Pracownik`, `Dzien_tygodnia`),
- `Wyjatki_dostepnosci` (`Pracownik`, `Data`),
- `Urlopy` (`Pracownik`, `Data_od`, `Data_do`),
- `Wymuszenia` (`Data`, `Typ_dyzuru`, `Pracownik`),
- `Parametry` (`solver_max_time_seconds`, `solver_num_workers`),
- `Wagi_celu` (`total`, `main`, `extra`, `friday`, `weekday`, `prefer_less_penalty`),
- `Instrukcja` (tworzona automatycznie jako pomoc dla użytkownika).

## Nazwy pracownikow
- W wyniku pracownicy sa zapisywani zgodnie z nazwa z arkusza `Pracownicy`, ale po oczyszczeniu przypadkowych enterow i nadmiarowych spacji.
- Mozesz podac dowolne nazwy pracownikow (np. imie i nazwisko, skroty, inicjaly), byle nie byly puste.
- Duplikaty sa blokowane walidacja, rowniez wtedy, gdy roznia sie tylko wielkoscia liter.
- Przy porownaniach i walidacji skrypt ignoruje przypadkowe spacje, taby i entery w polach tekstowych.


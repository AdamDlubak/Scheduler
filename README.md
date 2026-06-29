# Generator dyzurow pracownikow (OR-Tools CP-SAT)

Wersja v1.2 jest reuzywalna: dane sa w Excelu, walidator tworzy raport diagnostyczny,
a kod jest rozbity na moduly (loader/walidator/solver).

## Pliki
- `generator.py` - glowny runner CLI,
- `scheduler/excel_io.py` - odczyt/zapis Excel + raport diagnostyczny,
- `scheduler/validator.py` - walidacje techniczne i biznesowe,
- `scheduler/solver.py` - model CP-SAT,
- `scheduler/sample_data.py` - przykladowa konfiguracja,
- `scheduler/models.py` - dataclass i stale,
- `requirements.txt` - zaleznosci Pythona,
- plik wejsciowy XLSX z danymi (tworzony poleceniem `--create-template` lub `--export-sample-config`),
- wynikowy plik XLSX z schedulerem.

## Instalacja
```bash
python3 -m pip install -r requirements.txt
```

## Budowanie paczki EXE na Windows
W repo jest manualny workflow: `.github/workflows/build-windows-exe.yml`.

Jak zbudowac nowa wersje paczki:
1. Wypchnij aktualny kod do GitHub.
2. Wejdz w zakladke **Actions** -> **Build Windows EXE**.
3. Kliknij **Run workflow** i podaj `release_label` (np. `v1.3.0`).
4. Opcja `publish_release=true` opublikuje ten sam ZIP w zakladce **Releases**.
5. Po zakonczeniu pobierz artefakt `windows-exe-<release_label>`.
6. W artefakcie bedzie ZIP `Scheduler-Dyzurow-<release_label>.zip` gotowy do wyslania uzytkownikowi Windows.

Workflow wykonuje smoke test uruchomienia EXE (sprawdza `http://127.0.0.1:5001`),
wiec build zatrzyma sie, jesli paczka jest niekompletna.

Zawartosc paczki dla uzytkownika:
- `Scheduler-Dyzurow.exe`
- `Start.bat` (najprostsze uruchomienie)
- `INSTRUKCJA.txt`
- pliki wymagane przez aplikacje (`templates` i runtime PyInstaller).

Przykladowa struktura po rozpakowaniu ZIP:
```text
Scheduler-Dyzurow.exe
Start.bat
INSTRUKCJA.txt
templates/
```

## Troubleshooting EXE (Windows)
Jesli po uruchomieniu EXE pojawia sie traceback w `ortools` / `cp_model.py`,
to najczesciej brakuje natywnych bibliotek OR-Tools w paczce.

W workflow jest to naprawione przez:

```text
--collect-all ortools
```

Po tej poprawce uruchom nowy build i pobierz nowy ZIP.

## Szybki start
Pierwsze uruchomienie wygeneruje przykladowy plik danych wejsciowych:

```bash
python3 generator.py --export-sample-config dane_wejsciowe.xlsx
```

Nastepnie edytuj `dane_wejsciowe.xlsx` i uruchom walidacje/generowanie.

## Przydatne komendy
```bash
python3 generator.py --create-template template_dane.xlsx
python3 generator.py --export-sample-config dane_wejsciowe.xlsx
python3 generator.py --config dane_wejsciowe.xlsx --validate-only
python3 generator.py --config dane_wejsciowe.xlsx --validate-only --validation-report raport_bledow.xlsx
python3 generator.py --config dane_wejsciowe.xlsx --output "2026.07 - Lista Zastepstw.xlsx"
python3 generator.py --config dane_wejsciowe.xlsx --output-template "2026.07 - Lista Zastepstw.xlsx"
python3 generator.py --use-built-in-june-config
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


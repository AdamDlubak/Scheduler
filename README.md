# Generator dyzurow pracownikow (OR-Tools CP-SAT)

Wersja v1.2 jest reuzywalna: dane sa w Excelu, walidator tworzy raport diagnostyczny,
a kod jest rozbity na moduly (loader/walidator/solver).

## Pliki
- `generator.py` - glowny runner CLI,
- `harmonogram/excel_io.py` - odczyt/zapis Excel + raport diagnostyczny,
- `harmonogram/validator.py` - walidacje techniczne i biznesowe,
- `harmonogram/solver.py` - model CP-SAT,
- `harmonogram/sample_data.py` - przykladowa konfiguracja,
- `harmonogram/models.py` - dataclass i stale,
- `requirements.txt` - zaleznosci Pythona,
- plik wejsciowy XLSX z danymi (tworzony poleceniem `--create-template` lub `--export-sample-config`),
- wynikowy plik XLSX z harmonogramem.

## Instalacja
```bash
python3 -m pip install -r requirements.txt
```

## Budowanie paczki EXE na Windows (z macOS przez GitHub)
W repo jest manualny workflow: `.github/workflows/build-windows-exe.yml`.

Jak zbudowac nowa wersje paczki:
1. Wypchnij aktualny kod do GitHub.
2. Wejdz w zakladke **Actions** -> **Build Windows EXE**.
3. Kliknij **Run workflow** i podaj `release_label` (np. `v1.3.0`).
4. Opcja `publish_release=true` opublikuje ten sam ZIP w zakladce **Releases**.
5. Po zakonczeniu pobierz artefakt `windows-exe-<release_label>`.
6. W artefakcie bedzie ZIP `Harmonogram-Dyzurow-<release_label>.zip` gotowy do wyslania uzytkownikowi Windows.

Zawartosc paczki dla uzytkownika:
- `Harmonogram-Dyzurow.exe`
- `Start.bat` (najprostsze uruchomienie)
- `INSTRUKCJA.txt`
- pliki wymagane przez aplikacje (`templates` i runtime PyInstaller).

Przykladowa struktura po rozpakowaniu ZIP:
```text
Harmonogram-Dyzurow.exe
Start.bat
INSTRUKCJA.txt
templates/
```

## SmartScreen / "Uruchom mimo to" (Windows)
- Bezplatnie nie da sie tego gwarantowanie usunac dla EXE pobranego z internetu.
- Ten komunikat wynika z braku platnego podpisu kodu i reputacji SmartScreen.
- Najprostsza opcja bez kosztow: uruchamianie wersji webowej (uzytkownik tylko otwiera link w przegladarce).

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


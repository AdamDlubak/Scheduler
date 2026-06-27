# Generator dyzurow pracownikow (OR-Tools CP-SAT)

Skrypt buduje grafik dyzurow na dni robocze miesiaca, uwzgledniajac:
- niedostepnosci tygodniowe,
- urlopy,
- wymuszone dyzury,
- twarde reguly sekwencji,
- miekkie kryteria wyrownywania obciazen.

## Pliki
- `generator_dyzurow_czerwiec_2026.py` - glowny skrypt (latwo rozszerzalny dla kolejnych miesiecy),
- `requirements.txt` - zaleznosci Pythona,
- `grafik_czerwiec_2026.xlsx` - wynik (po uruchomieniu skryptu).

## Uruchomienie
```bash
python3 -m pip install -r requirements.txt
python3 generator_dyzurow_czerwiec_2026.py
```

## Jak dostosowac do innego miesiaca
W funkcji `june_2026_config()` podmien:
- `year`, `month`,
- `employees`,
- `holidays`,
- `weekly_unavailability`,
- `weekly_availability_exceptions`,
- `vacations`,
- `forced_assignments`.

Reszta logiki modelu pozostaje bez zmian.


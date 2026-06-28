#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
from calendar import monthrange
from datetime import date
from pathlib import Path

from harmonogram.excel_io import (
    export_config_to_excel,
    export_schedule_to_excel,
    export_validation_report,
    load_config_from_excel,
    write_template_excel,
)
from harmonogram.models import InputValidationError
from harmonogram.sample_data import june_2026_config
from harmonogram.solver import build_workdays, print_table, solve_schedule


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generator dyzurow (OR-Tools CP-SAT), wersja v1.2.")
    parser.add_argument(
        "--config",
        default="dane_wejsciowe_czerwiec_2026.xlsx",
        help="Plik Excela z danymi wejsciowymi.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Sciezka wyjsciowego pliku XLSX. Domyslnie: yyyy.MM - Lista Zastępstw.xlsx",
    )
    parser.add_argument(
        "--output-template",
        default="",
        help="Opcjonalny plik szablonu XLSX dla wyniku (domyslnie: 2026.06 - Lista Zastępstw.xlsx).",
    )
    parser.add_argument(
        "--create-template",
        default="",
        help="Zapisz pusty szablon danych wejsciowych XLSX i zakoncz.",
    )
    parser.add_argument(
        "--export-sample-config",
        default="",
        help="Zapisz przykladowa konfiguracje (czerwiec 2026) do XLSX i zakoncz.",
    )
    parser.add_argument(
        "--use-built-in-june-config",
        action="store_true",
        help="Uzyj konfiguracji zaszytej w kodzie (awaryjnie, bez Excela).",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Sprawdz tylko poprawnosc pliku Excel, bez uruchamiania solvera.",
    )
    parser.add_argument(
        "--validation-report",
        default="",
        help="Opcjonalna sciezka raportu diagnostycznego XLSX. Domyslnie: <config>_diagnostyka.xlsx.",
    )
    return parser


def print_startup_diagnostics(config) -> None:
    print("\nDIAGNOSTYKA WEJSCIA")
    print(f"- Miesiac: {config.year}-{config.month:02d}")

    holidays_in_month = sorted(d for d in config.holidays if d.year == config.year and d.month == config.month)
    if holidays_in_month:
        print(f"- Swieta ({len(holidays_in_month)}): " + ", ".join(d.isoformat() for d in holidays_in_month))
    else:
        print("- Swieta (0): brak")

    all_days = [date(config.year, config.month, d) for d in range(1, monthrange(config.year, config.month)[1] + 1)]
    weekends = [d for d in all_days if d.weekday() >= 5]
    workdays = build_workdays(config.year, config.month, config.holidays)
    holiday_weekdays = [d for d in holidays_in_month if d.weekday() < 5]

    print(f"- Wszystkie dni miesiaca: {len(all_days)}")
    print(f"- Weekendy (wykluczone): {len(weekends)}")
    print(f"- Swieta robocze (wykluczone): {len(holiday_weekdays)}")
    print(f"- Dni planowane: {len(workdays)}")

    if workdays:
        print("- Zakres planowania: " + f"{workdays[0].isoformat()} -> {workdays[-1].isoformat()}")


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()

    if args.create_template:
        write_template_excel(args.create_template)
        print(f"Zapisano szablon danych: {args.create_template}")
        return

    if args.export_sample_config:
        export_config_to_excel(june_2026_config(), args.export_sample_config)
        print(f"Zapisano przykladowa konfiguracje: {args.export_sample_config}")
        return

    warnings: list[str] = []

    if args.use_built_in_june_config:
        config = june_2026_config()
    else:
        config_path = Path(args.config)
        if not config_path.exists():
            export_config_to_excel(june_2026_config(), str(config_path))
            print(
                f"Nie znaleziono pliku {config_path}. Utworzono przykladowe dane wejscowe. "
                "Dostosuj je i uruchom skrypt ponownie."
            )
            return

        report_path = args.validation_report or str(config_path.with_name(f"{config_path.stem}_diagnostyka.xlsx"))

        try:
            load_result = load_config_from_excel(str(config_path))
        except InputValidationError as exc:
            export_validation_report(str(config_path), report_path, exc.errors, exc.warnings)
            print(exc.format_message())
            print(f"\nZapisano raport diagnostyczny: {report_path}")
            return

        config = load_result.config
        warnings = load_result.warnings

    if warnings:
        print("Ostrzezenia w danych wejsciowych:")
        for warning in warnings:
            print(f"- {warning}")

    print_startup_diagnostics(config)

    if args.validate_only:
        print("Walidacja zakonczona powodzeniem - plik wejsciowy jest poprawny.")
        return

    result = solve_schedule(config)

    print(f"Status solvera: {result.status_name}")
    if result.objective_value is not None:
        print(f"Wartosc funkcji celu: {result.objective_value:.2f}")

    if not result.schedule_rows:
        print("Brak grafiku do zapisania.")
        return

    print("\nGRAFIK")
    schedule_for_print = [(day.isoformat(), weekday, main, extra) for day, weekday, main, extra in result.schedule_rows]
    print_table(["Data", "Dzien", "Glowny", "Dodatkowy"], schedule_for_print)

    print("\nPODSUMOWANIE")
    print_table(["Pracownik", "Lacznie", "Glowne", "Dodatkowe", "Piatki"], result.summary_rows)

    output_path = args.output or f"{config.year}.{config.month:02d} - Lista Zastępstw.xlsx"
    template_path = args.output_template or None
    export_schedule_to_excel(result, output_path, config=config, template_path=template_path)
    print(f"\nZapisano plik Excel: {output_path}")


if __name__ == "__main__":
    main()




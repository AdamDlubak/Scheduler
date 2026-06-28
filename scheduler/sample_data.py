from __future__ import annotations

from datetime import date

from scheduler.models import SHIFT_MAIN, SchedulerConfig, VacationRange


def june_2026_config() -> SchedulerConfig:
    employees = [
        "Ewa",
        "Kasia",
        "Kinga",
        "Kamila",
        "Ola",
        "Karolina",
        "Teresa",
        "Aneta Sz.",
        "Aneta Ł.",
        "Ania",
    ]

    holidays = {date(2026, 6, 4)}

    weekly_unavailability = {
        "Ewa": {1, 3},
        "Kasia": {0, 1},
        "Kinga": {0, 3},
        "Kamila": {0, 2},
        "Ola": {0, 2},
        "Karolina": {0, 2},
        "Teresa": {0, 2},
        "Aneta Sz.": {1, 3},
        "Aneta Ł.": {1, 2},
        "Ania": {1},
    }

    weekly_availability_exceptions = {
        "Teresa": {date(2026, 6, 1)},
    }

    vacations = {
        "Ewa": [VacationRange(date(2026, 6, 1), date(2026, 6, 14))],
        "Karolina": [VacationRange(date(2026, 6, 15), date(2026, 6, 30))],
        "Aneta Sz.": [VacationRange(date(2026, 6, 15), date(2026, 6, 30))],
        "Kinga": [
            VacationRange(date(2026, 6, 2), date(2026, 6, 8)),
            VacationRange(date(2026, 6, 26), date(2026, 6, 26)),
            VacationRange(date(2026, 6, 30), date(2026, 6, 30)),
        ],
        "Kamila": [VacationRange(date(2026, 6, 10), date(2026, 6, 12))],
        "Ania": [
            VacationRange(date(2026, 6, 8), date(2026, 6, 10)),
            VacationRange(date(2026, 6, 22), date(2026, 6, 30)),
        ],
        "Aneta Ł.": [VacationRange(date(2026, 6, 25), date(2026, 6, 26))],
        "Ola": [VacationRange(date(2026, 6, 1), date(2026, 6, 5))],
        "Teresa": [VacationRange(date(2026, 6, 5), date(2026, 6, 8))],
    }

    forced_assignments = {
        (date(2026, 6, 1), SHIFT_MAIN): "Teresa",
        (date(2026, 6, 2), SHIFT_MAIN): "Karolina",
    }

    return SchedulerConfig(
        year=2026,
        month=6,
        employees=employees,
        holidays=holidays,
        weekly_unavailability=weekly_unavailability,
        weekly_availability_exceptions=weekly_availability_exceptions,
        vacations=vacations,
        forced_assignments=forced_assignments,
        no_main_shift_employees={"Ania"},
        prefer_less_duties_employees={"Ania"},
    )


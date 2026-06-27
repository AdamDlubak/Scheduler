#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

from calendar import monthrange
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, Iterable, List, Sequence, Set, Tuple

from openpyxl import Workbook
from ortools.sat.python import cp_model


SHIFT_MAIN = "glowny"
SHIFT_EXTRA = "dodatkowy"
SHIFTS = (SHIFT_MAIN, SHIFT_EXTRA)
WEEKDAY_NAMES_PL = {
    0: "pon",
    1: "wt",
    2: "sr",
    3: "czw",
    4: "pt",
    5: "sob",
    6: "niedz",
}


@dataclass(frozen=True)
class VacationRange:
    start: date
    end: date

    def contains(self, day: date) -> bool:
        return self.start <= day <= self.end


@dataclass
class SchedulerConfig:
    year: int
    month: int
    employees: List[str]
    holidays: Set[date]
    weekly_unavailability: Dict[str, Set[int]]
    weekly_availability_exceptions: Dict[str, Set[date]] = field(default_factory=dict)
    vacations: Dict[str, List[VacationRange]] = field(default_factory=dict)
    forced_assignments: Dict[Tuple[date, str], str] = field(default_factory=dict)


@dataclass
class ScheduleResult:
    status_name: str
    schedule_rows: List[Tuple[date, str, str, str]]
    summary_rows: List[Tuple[str, int, int, int, int]]
    objective_value: float | None = None


@dataclass
class ModelArtifacts:
    model: cp_model.CpModel
    x: Dict[Tuple[str, date, str], cp_model.IntVar]
    totals: Dict[str, cp_model.LinearExpr]
    totals_main: Dict[str, cp_model.LinearExpr]
    totals_extra: Dict[str, cp_model.LinearExpr]
    fridays: Dict[str, cp_model.LinearExpr]


def daterange(start: date, end: date) -> Iterable[date]:
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def build_workdays(year: int, month: int, holidays: Set[date]) -> List[date]:
    _, last_day = monthrange(year, month)
    days: List[date] = []
    for d in range(1, last_day + 1):
        day = date(year, month, d)
        if day.weekday() >= 5:
            continue
        if day in holidays:
            continue
        days.append(day)
    return days


def is_unavailable(config: SchedulerConfig, employee: str, day: date) -> bool:
    if day in config.weekly_availability_exceptions.get(employee, set()):
        return False

    if day.weekday() in config.weekly_unavailability.get(employee, set()):
        return True

    for vr in config.vacations.get(employee, []):
        if vr.contains(day):
            return True
    return False


def static_consistency_checks(config: SchedulerConfig, workdays: Sequence[date]) -> List[str]:
    errors: List[str] = []

    for (day, shift), employee in config.forced_assignments.items():
        if day not in workdays:
            errors.append(f"Wymuszenie {employee} {shift} {day} wypada poza dniem roboczym grafiku.")
            continue
        if is_unavailable(config, employee, day):
            errors.append(f"Wymuszenie {employee} {shift} {day} koliduje z niedostepnoscia.")

    # Szybka diagnostyka: czy kazdego dnia sa min. 2 osoby dostepne.
    for day in workdays:
        available = [e for e in config.employees if not is_unavailable(config, e, day)]
        if len(available) < 2:
            errors.append(
                f"Dzien {day}: dostepnych osob = {len(available)}, a potrzebne sa 2 dyzury."
            )

    return errors


def add_fairness_objective(
    model: cp_model.CpModel,
    config: SchedulerConfig,
    workdays: Sequence[date],
    totals: Dict[str, cp_model.LinearExpr],
    totals_main: Dict[str, cp_model.LinearExpr],
    totals_extra: Dict[str, cp_model.LinearExpr],
    fridays: Dict[str, cp_model.LinearExpr],
    duties_by_weekday: Dict[Tuple[str, int], cp_model.LinearExpr],
) -> None:
    employees = config.employees
    max_total = len(workdays)

    objective_terms: List[cp_model.LinearExpr] = []

    def pairwise_abs(expr_a: cp_model.LinearExpr, expr_b: cp_model.LinearExpr, name: str) -> cp_model.IntVar:
        diff = model.NewIntVar(0, max_total, name)
        model.Add(diff >= expr_a - expr_b)
        model.Add(diff >= expr_b - expr_a)
        return diff

    for i in range(len(employees)):
        for j in range(i + 1, len(employees)):
            e1, e2 = employees[i], employees[j]
            objective_terms.append(8 * pairwise_abs(totals[e1], totals[e2], f"d_tot_{i}_{j}"))
            objective_terms.append(5 * pairwise_abs(totals_main[e1], totals_main[e2], f"d_main_{i}_{j}"))
            objective_terms.append(5 * pairwise_abs(totals_extra[e1], totals_extra[e2], f"d_extra_{i}_{j}"))
            objective_terms.append(4 * pairwise_abs(fridays[e1], fridays[e2], f"d_fr_{i}_{j}"))
            for wd in range(5):
                objective_terms.append(
                    2 * pairwise_abs(
                        duties_by_weekday[(e1, wd)],
                        duties_by_weekday[(e2, wd)],
                        f"d_wd_{wd}_{i}_{j}",
                    )
                )

    # Ania powinna miec mniej dyzurow niz pozostali (miekka kara za naruszenie).
    ania = "Ania"
    if ania in employees:
        for emp in employees:
            if emp == ania:
                continue
            penalty = model.NewIntVar(0, max_total, f"penalty_ania_vs_{emp}")
            model.Add(penalty >= totals[ania] - totals[emp] + 1)
            model.Add(penalty >= 0)
            objective_terms.append(20 * penalty)

    model.Minimize(sum(objective_terms))


def build_model(config: SchedulerConfig, workdays: Sequence[date]) -> ModelArtifacts:
    model = cp_model.CpModel()
    x: Dict[Tuple[str, date, str], cp_model.IntVar] = {}

    for e in config.employees:
        for day in workdays:
            for s in SHIFTS:
                x[(e, day, s)] = model.NewBoolVar(f"x_{e}_{day}_{s}")

    # Dokladnie 1 glowny i 1 dodatkowy dziennie.
    for day in workdays:
        for s in SHIFTS:
            model.Add(sum(x[(e, day, s)] for e in config.employees) == 1)

    # Max 1 dyzur dziennie na osobe.
    for e in config.employees:
        for day in workdays:
            model.Add(sum(x[(e, day, s)] for s in SHIFTS) <= 1)

    # Niedostepnosci.
    for e in config.employees:
        for day in workdays:
            if is_unavailable(config, e, day):
                for s in SHIFTS:
                    model.Add(x[(e, day, s)] == 0)

    # Wymuszenia.
    for (day, shift), employee in config.forced_assignments.items():
        if day in workdays:
            model.Add(x[(employee, day, shift)] == 1)

    # Ania nigdy nie ma dyzuru glownego.
    if "Ania" in config.employees:
        for day in workdays:
            model.Add(x[("Ania", day, SHIFT_MAIN)] == 0)

    # Zakaz glownych dzien po dniu (tylko rzeczywista kolejnosc kalendarzowa).
    for e in config.employees:
        for i in range(len(workdays) - 1):
            d1, d2 = workdays[i], workdays[i + 1]
            if (d2 - d1).days == 1:
                model.Add(x[(e, d1, SHIFT_MAIN)] + x[(e, d2, SHIFT_MAIN)] <= 1)

    # Zakaz dyzurow przez 3 kolejne dni kalendarzowe.
    for e in config.employees:
        for i in range(len(workdays) - 2):
            d1, d2, d3 = workdays[i], workdays[i + 1], workdays[i + 2]
            if (d2 - d1).days == 1 and (d3 - d2).days == 1:
                model.Add(
                    sum(x[(e, d, s)] for d in (d1, d2, d3) for s in SHIFTS) <= 2
                )

    totals: Dict[str, cp_model.LinearExpr] = {}
    totals_main: Dict[str, cp_model.LinearExpr] = {}
    totals_extra: Dict[str, cp_model.LinearExpr] = {}
    fridays: Dict[str, cp_model.LinearExpr] = {}
    duties_by_weekday: Dict[Tuple[str, int], cp_model.LinearExpr] = {}

    for e in config.employees:
        totals[e] = sum(x[(e, day, s)] for day in workdays for s in SHIFTS)
        totals_main[e] = sum(x[(e, day, SHIFT_MAIN)] for day in workdays)
        totals_extra[e] = sum(x[(e, day, SHIFT_EXTRA)] for day in workdays)
        fridays[e] = sum(x[(e, day, s)] for day in workdays if day.weekday() == 4 for s in SHIFTS)
        for wd in range(5):
            duties_by_weekday[(e, wd)] = sum(
                x[(e, day, s)] for day in workdays if day.weekday() == wd for s in SHIFTS
            )

    add_fairness_objective(
        model,
        config,
        workdays,
        totals,
        totals_main,
        totals_extra,
        fridays,
        duties_by_weekday,
    )

    return ModelArtifacts(model=model, x=x, totals=totals, totals_main=totals_main, totals_extra=totals_extra, fridays=fridays)


def solve_schedule(config: SchedulerConfig) -> ScheduleResult:
    workdays = build_workdays(config.year, config.month, config.holidays)
    pre_errors = static_consistency_checks(config, workdays)
    if pre_errors:
        joined = "\n- " + "\n- ".join(pre_errors)
        return ScheduleResult(
            status_name=f"INFEASIBLE (wykryto konflikt juz przed solverem):{joined}",
            schedule_rows=[],
            summary_rows=[],
        )

    artifacts = build_model(config, workdays)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 240
    solver.parameters.num_search_workers = 8

    status = solver.Solve(artifacts.model)
    status_name = solver.StatusName(status)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        details = [
            "Solver nie znalazl rozwiazania.",
            "Mozliwe konflikty: zbyt mala liczba dostepnych osob w pojedynczych dniach,",
            "kolidujace wymuszenia, albo zbyt restrykcyjne nakladanie niedostepnosci/urlopow.",
        ]
        return ScheduleResult(
            status_name=f"{status_name} | " + " ".join(details),
            schedule_rows=[],
            summary_rows=[],
        )

    schedule_rows: List[Tuple[date, str, str, str]] = []
    for day in workdays:
        main = next(e for e in config.employees if solver.Value(artifacts.x[(e, day, SHIFT_MAIN)]) == 1)
        extra = next(e for e in config.employees if solver.Value(artifacts.x[(e, day, SHIFT_EXTRA)]) == 1)
        schedule_rows.append((day, WEEKDAY_NAMES_PL[day.weekday()], main, extra))

    summary_rows: List[Tuple[str, int, int, int, int]] = []
    for e in config.employees:
        total = int(solver.Value(artifacts.totals[e]))
        main = int(solver.Value(artifacts.totals_main[e]))
        extra = int(solver.Value(artifacts.totals_extra[e]))
        fri = int(solver.Value(artifacts.fridays[e]))
        summary_rows.append((e, total, main, extra, fri))

    summary_rows.sort(key=lambda r: (-r[1], r[0]))

    return ScheduleResult(
        status_name=status_name,
        schedule_rows=schedule_rows,
        summary_rows=summary_rows,
        objective_value=solver.ObjectiveValue(),
    )


def print_table(headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(str(cell)))

    def fmt_row(row: Sequence[object]) -> str:
        return " | ".join(str(c).ljust(widths[i]) for i, c in enumerate(row))

    print(fmt_row(headers))
    print("-+-".join("-" * w for w in widths))
    for r in rows:
        print(fmt_row(r))


def export_to_excel(result: ScheduleResult, output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Grafik"

    ws.append(["Data", "Dzien", "Dyzur glowny", "Dyzur dodatkowy"])
    for day, day_name, main, extra in result.schedule_rows:
        ws.append([day.isoformat(), day_name, main, extra])

    ws2 = wb.create_sheet("Podsumowanie")
    ws2.append(["Pracownik", "Lacznie", "Glowne", "Dodatkowe", "Piatki"])
    for row in result.summary_rows:
        ws2.append(list(row))

    wb.save(output_path)


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
        "Kinga": [VacationRange(date(2026, 6, 2), date(2026, 6, 8)), VacationRange(date(2026, 6, 26), date(2026, 6, 26)), VacationRange(date(2026, 6, 30), date(2026, 6, 30))],
        "Kamila": [VacationRange(date(2026, 6, 10), date(2026, 6, 12))],
        "Ania": [VacationRange(date(2026, 6, 8), date(2026, 6, 10)), VacationRange(date(2026, 6, 22), date(2026, 6, 30))],
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
    )


def main() -> None:
    config = june_2026_config()
    result = solve_schedule(config)

    print(f"Status solvera: {result.status_name}")
    if result.objective_value is not None:
        print(f"Wartosc funkcji celu: {result.objective_value:.2f}")

    if not result.schedule_rows:
        print("Brak grafiku do zapisania.")
        return

    print("\nGRAFIK")
    schedule_for_print = [(d.isoformat(), wd, m, ex) for d, wd, m, ex in result.schedule_rows]
    print_table(["Data", "Dzien", "Glowny", "Dodatkowy"], schedule_for_print)

    print("\nPODSUMOWANIE")
    print_table(["Pracownik", "Lacznie", "Glowne", "Dodatkowe", "Piatki"], result.summary_rows)

    output_path = "grafik_czerwiec_2026.xlsx"
    export_to_excel(result, output_path)
    print(f"\nZapisano plik Excel: {output_path}")


if __name__ == "__main__":
    main()






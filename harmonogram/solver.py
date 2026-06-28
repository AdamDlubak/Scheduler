from __future__ import annotations

from calendar import monthrange
from datetime import date
from typing import Dict, List, Sequence, Set, Tuple

from ortools.sat.python import cp_model

from harmonogram.models import (
    ModelArtifacts,
    SHIFTS,
    SHIFT_EXTRA,
    SHIFT_MAIN,
    WEEKDAY_NAMES_PL,
    ScheduleResult,
    SchedulerConfig,
)


def build_workdays(year: int, month: int, holidays: Set[date]) -> List[date]:
    _, last_day = monthrange(year, month)
    days: List[date] = []
    for day_index in range(1, last_day + 1):
        day = date(year, month, day_index)
        if day.weekday() >= 5:
            continue
        if day in holidays:
            continue
        days.append(day)
    return days


def is_unavailable(config: SchedulerConfig, employee: str, day: date) -> bool:
    has_weekly_exception = day in config.weekly_availability_exceptions.get(employee, set())
    if day.weekday() in config.weekly_unavailability.get(employee, set()) and not has_weekly_exception:
        return True

    for vacation in config.vacations.get(employee, []):
        if vacation.contains(day):
            return True
    return False


def static_consistency_checks(config: SchedulerConfig, workdays: Sequence[date]) -> List[str]:
    errors: List[str] = []

    for (day, shift), employee in config.forced_assignments.items():
        if day not in workdays:
            errors.append(f"Wymuszenie {employee} {shift} {day} wypada poza dniem roboczym grafiku.")
            continue
        if is_unavailable(config, employee, day):
            errors.append(f"Wymuszenie {employee} {shift} {day} koliduje z niedostepnoscia lub urlopem.")

    for day in workdays:
        available = [e for e in config.employees if not is_unavailable(config, e, day)]
        if len(available) < 2:
            errors.append(f"Dzien {day}: dostepnych osob = {len(available)}, a potrzebne sa 2 dyzury.")

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
    weights = {
        "total": int(config.objective_weights.get("total", 8)),
        "main": int(config.objective_weights.get("main", 5)),
        "extra": int(config.objective_weights.get("extra", 5)),
        "friday": int(config.objective_weights.get("friday", 4)),
        "weekday": int(config.objective_weights.get("weekday", 2)),
        "prefer_less_penalty": int(config.objective_weights.get("prefer_less_penalty", 20)),
    }

    objective_terms: List[cp_model.LinearExpr] = []

    def pairwise_abs(expr_a: cp_model.LinearExpr, expr_b: cp_model.LinearExpr, name: str) -> cp_model.IntVar:
        diff = model.NewIntVar(0, max_total, name)
        model.Add(diff >= expr_a - expr_b)
        model.Add(diff >= expr_b - expr_a)
        return diff

    for i in range(len(employees)):
        for j in range(i + 1, len(employees)):
            e1, e2 = employees[i], employees[j]
            objective_terms.append(weights["total"] * pairwise_abs(totals[e1], totals[e2], f"d_tot_{i}_{j}"))
            objective_terms.append(weights["main"] * pairwise_abs(totals_main[e1], totals_main[e2], f"d_main_{i}_{j}"))
            objective_terms.append(weights["extra"] * pairwise_abs(totals_extra[e1], totals_extra[e2], f"d_extra_{i}_{j}"))
            objective_terms.append(weights["friday"] * pairwise_abs(fridays[e1], fridays[e2], f"d_fr_{i}_{j}"))
            for weekday in range(5):
                objective_terms.append(
                    weights["weekday"]
                    * pairwise_abs(
                        duties_by_weekday[(e1, weekday)],
                        duties_by_weekday[(e2, weekday)],
                        f"d_wd_{weekday}_{i}_{j}",
                    )
                )

    for pref_emp in sorted(config.prefer_less_duties_employees):
        if pref_emp not in employees:
            continue
        for emp in employees:
            if emp == pref_emp:
                continue
            penalty = model.NewIntVar(0, max_total, f"penalty_pref_less_{pref_emp}_vs_{emp}")
            model.Add(penalty >= totals[pref_emp] - totals[emp] + 1)
            model.Add(penalty >= 0)
            objective_terms.append(weights["prefer_less_penalty"] * penalty)

    model.Minimize(sum(objective_terms))


def build_model(config: SchedulerConfig, workdays: Sequence[date]) -> ModelArtifacts:
    model = cp_model.CpModel()
    x: Dict[Tuple[str, date, str], cp_model.IntVar] = {}

    for employee in config.employees:
        for day in workdays:
            for shift in SHIFTS:
                x[(employee, day, shift)] = model.NewBoolVar(f"x_{employee}_{day}_{shift}")

    for day in workdays:
        for shift in SHIFTS:
            model.Add(sum(x[(employee, day, shift)] for employee in config.employees) == 1)

    for employee in config.employees:
        for day in workdays:
            model.Add(sum(x[(employee, day, shift)] for shift in SHIFTS) <= 1)

    for employee in config.employees:
        for day in workdays:
            if is_unavailable(config, employee, day):
                for shift in SHIFTS:
                    model.Add(x[(employee, day, shift)] == 0)

    for (day, shift), employee in config.forced_assignments.items():
        if day in workdays:
            model.Add(x[(employee, day, shift)] == 1)

    for employee in sorted(config.no_main_shift_employees):
        if employee in config.employees:
            for day in workdays:
                model.Add(x[(employee, day, SHIFT_MAIN)] == 0)

    for employee in config.employees:
        for i in range(len(workdays) - 1):
            d1, d2 = workdays[i], workdays[i + 1]
            if (d2 - d1).days == 1:
                model.Add(x[(employee, d1, SHIFT_MAIN)] + x[(employee, d2, SHIFT_MAIN)] <= 1)

    for employee in config.employees:
        for i in range(len(workdays) - 2):
            d1, d2, d3 = workdays[i], workdays[i + 1], workdays[i + 2]
            if (d2 - d1).days == 1 and (d3 - d2).days == 1:
                model.Add(sum(x[(employee, d, shift)] for d in (d1, d2, d3) for shift in SHIFTS) <= 2)

    totals: Dict[str, cp_model.LinearExpr] = {}
    totals_main: Dict[str, cp_model.LinearExpr] = {}
    totals_extra: Dict[str, cp_model.LinearExpr] = {}
    fridays: Dict[str, cp_model.LinearExpr] = {}
    duties_by_weekday: Dict[Tuple[str, int], cp_model.LinearExpr] = {}

    for employee in config.employees:
        totals[employee] = sum(x[(employee, day, shift)] for day in workdays for shift in SHIFTS)
        totals_main[employee] = sum(x[(employee, day, SHIFT_MAIN)] for day in workdays)
        totals_extra[employee] = sum(x[(employee, day, SHIFT_EXTRA)] for day in workdays)
        fridays[employee] = sum(
            x[(employee, day, shift)] for day in workdays if day.weekday() == 4 for shift in SHIFTS
        )
        for weekday in range(5):
            duties_by_weekday[(employee, weekday)] = sum(
                x[(employee, day, shift)] for day in workdays if day.weekday() == weekday for shift in SHIFTS
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

    return ModelArtifacts(
        model=model,
        x=x,
        totals=totals,
        totals_main=totals_main,
        totals_extra=totals_extra,
        fridays=fridays,
    )


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
    solver.parameters.max_time_in_seconds = max(1, int(config.solver_max_time_seconds))
    solver.parameters.num_search_workers = max(1, int(config.solver_num_workers))

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
        main = next(employee for employee in config.employees if solver.Value(artifacts.x[(employee, day, SHIFT_MAIN)]) == 1)
        extra = next(employee for employee in config.employees if solver.Value(artifacts.x[(employee, day, SHIFT_EXTRA)]) == 1)
        schedule_rows.append((day, WEEKDAY_NAMES_PL[day.weekday()], main, extra))

    summary_rows: List[Tuple[str, int, int, int, int]] = []
    for employee in config.employees:
        total = int(solver.Value(artifacts.totals[employee]))
        main = int(solver.Value(artifacts.totals_main[employee]))
        extra = int(solver.Value(artifacts.totals_extra[employee]))
        friday = int(solver.Value(artifacts.fridays[employee]))
        summary_rows.append((employee, total, main, extra, friday))

    summary_rows.sort(key=lambda row: (-row[1], row[0]))
    return ScheduleResult(
        status_name=status_name,
        schedule_rows=schedule_rows,
        summary_rows=summary_rows,
        objective_value=solver.ObjectiveValue(),
    )


def print_table(headers: Sequence[str], rows: Sequence[Sequence[object]]) -> None:
    widths = [len(header) for header in headers]
    for row in rows:
        for i, cell in enumerate(row):
            widths[i] = max(widths[i], len(str(cell)))

    def fmt_row(row: Sequence[object]) -> str:
        return " | ".join(str(cell).ljust(widths[i]) for i, cell in enumerate(row))

    print(fmt_row(headers))
    print("-+-".join("-" * width for width in widths))
    for row in rows:
        print(fmt_row(row))


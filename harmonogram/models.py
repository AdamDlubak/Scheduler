from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Set, Tuple

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
    no_main_shift_employees: Set[str] = field(default_factory=set)
    prefer_less_duties_employees: Set[str] = field(default_factory=set)
    solver_max_time_seconds: int = 240
    solver_num_workers: int = 8
    objective_weights: Dict[str, int] = field(
        default_factory=lambda: {
            "total": 8,
            "main": 5,
            "extra": 5,
            "friday": 4,
            "weekday": 2,
            "prefer_less_penalty": 20,
        }
    )


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


@dataclass
class ExcelLoadResult:
    config: SchedulerConfig
    warnings: List[str] = field(default_factory=list)


class InputValidationError(Exception):
    def __init__(self, errors: List[str], warnings: List[str] | None = None):
        self.errors = list(errors)
        self.warnings = list(warnings or [])
        super().__init__(self.format_message())

    def format_message(self) -> str:
        parts = ["Wykryto bledy w pliku wejsciowym:"]
        parts.extend(f"- {error}" for error in self.errors)
        if self.warnings:
            parts.append("\nOstrzezenia:")
            parts.extend(f"- {warning}" for warning in self.warnings)
        return "\n".join(parts)


#!/usr/bin/env python3
from __future__ import annotations

from harmonogram.sample_data import june_2026_config
from harmonogram.solver import solve_schedule


def main() -> int:
    result = solve_schedule(june_2026_config())
    print(f"Status: {result.status_name}")
    if not result.schedule_rows:
        print("Brak harmonogramu.")
        return 1
    print(f"Liczba dni w grafiku: {len(result.schedule_rows)}")
    print(f"Liczba pracownikow w podsumowaniu: {len(result.summary_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


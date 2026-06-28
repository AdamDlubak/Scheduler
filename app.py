#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import os
import json
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from datetime import datetime

from harmonogram.excel_io import (
    export_config_to_excel,
    export_schedule_to_excel,
    export_validation_report,
    load_config_from_excel,
    write_template_excel,
)
from harmonogram.models import InputValidationError
from harmonogram.sample_data import june_2026_config
from harmonogram.solver import solve_schedule, build_workdays, print_table
from harmonogram.diagnostics import (
    create_diagnostic_report,
    get_diagnostics_folder,
    log_info,
    log_error,
)
from calendar import monthrange
import platform
import subprocess

# Konfiguracja
BASE_DIR = Path(__file__).parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
DOWNLOAD_FOLDER = BASE_DIR / "downloads"
ALLOWED_EXTENSIONS = {"xlsx"}

# Tworzenie folderów
UPLOAD_FOLDER.mkdir(exist_ok=True)
DOWNLOAD_FOLDER.mkdir(exist_ok=True)

# Flask app
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["DOWNLOAD_FOLDER"] = DOWNLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB
CORS(app)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_validation_errors(config_path: str) -> dict:
    """Waliduj plik i zwróć błędy/ostrzeżenia"""
    try:
        load_result = load_config_from_excel(config_path)
        return {"success": True, "warnings": load_result.warnings, "errors": []}
    except InputValidationError as exc:
        return {
            "success": False,
            "errors": exc.errors,
            "warnings": exc.warnings,
        }


def localize_solver_status(status_name: str) -> str:
    """Zamień techniczne statusy solvera na prosty, polski opis dla UI."""
    status_map = {
        "OPTIMAL": "Gotowe (najlepszy wynik)",
        "FEASIBLE": "Gotowe (poprawny grafik)",
        "INFEASIBLE": "Brak rozwiązania",
        "MODEL_INVALID": "Błąd konfiguracji modelu",
        "UNKNOWN": "Brak jednoznacznego wyniku",
    }
    for code, label in status_map.items():
        if status_name.startswith(code):
            return status_name.replace(code, label, 1)
    return status_name


@app.route("/")
def index():
    """Strona główna"""
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_file():
    """Obsługa uploadu pliku"""
    if "file" not in request.files:
        return jsonify({"error": "Brak pliku"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Pusta nazwa pliku"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Tylko pliki .xlsx są akceptowane"}), 400

    filepath: Path | None = None
    try:
        # Zapisz plik
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"config_{timestamp}.xlsx"
        filepath = UPLOAD_FOLDER / filename

        file.save(filepath)

        # Wstepna walidacja (nie blokuje przejscia do kroku "Sprawdz i kontynuuj").
        try:
            validation = get_validation_errors(str(filepath))
        except Exception as exc:
            validation = {
                "success": False,
                "warnings": [],
                "errors": [
                    "Nie udalo sie odczytac pliku jako poprawny Excel. "
                    "Kliknij 'Sprawdz i kontynuuj', aby zobaczyc szczegoly.",
                    str(exc),
                ],
            }

        return jsonify({
            "success": True,
            "filename": filename,
            "filepath": str(filepath),
            "validation": validation,
        })

    except Exception as e:
        if filepath and filepath.exists():
            return jsonify({
                "success": True,
                "filename": filepath.name,
                "filepath": str(filepath),
                "validation": {
                    "success": False,
                    "warnings": [],
                    "errors": [
                        "Plik zostal zapisany, ale wystapil problem z jego wstepnym sprawdzeniem.",
                        str(e),
                    ],
                },
            })
        return jsonify({"error": str(e)}), 500


@app.route("/api/validate", methods=["POST"])
def validate():
    """Waliduj plik bez generowania harmonogramu"""
    data = request.get_json()
    filepath = data.get("filepath")

    if not filepath or not Path(filepath).exists():
        return jsonify({"error": "Plik nie znaleziony"}), 400

    try:
        validation = get_validation_errors(filepath)

        if validation["success"]:
            # Załaduj konfigurację do wyświetlenia diagnostyki
            load_result = load_config_from_excel(filepath)
            config = load_result.config

            # Diagnostyka
            all_days = [
                datetime(config.year, config.month, d).date()
                for d in range(1, monthrange(config.year, config.month)[1] + 1)
            ]
            weekends = [d for d in all_days if d.weekday() >= 5]
            workdays = build_workdays(config.year, config.month, config.holidays)
            holidays_in_month = sorted(
                d for d in config.holidays
                if d.year == config.year and d.month == config.month
            )
            holiday_weekdays = [d for d in holidays_in_month if d.weekday() < 5]

            diagnostics = {
                "month": f"{config.year}-{config.month:02d}",
                "employees_count": len(config.employees),
                "holidays_count": len(holidays_in_month),
                "all_days": len(all_days),
                "weekends": len(weekends),
                "holiday_weekdays": len(holiday_weekdays),
                "workdays": len(workdays),
                "solver_max_time": config.solver_max_time_seconds,
                "solver_workers": config.solver_num_workers,
            }
            # Przetłumacz na bardziej ludzki format dla diagnostyki
            diagnostics["display_holidays"] = diagnostics.get("holidays_count", 0)
            diagnostics["display_skip_days"] = diagnostics["weekends"] + diagnostics.get("holiday_weekdays", 0)

            log_info(f"Walidacja powodzenie: {filepath}")

            return jsonify({
                "success": True,
                "warnings": validation["warnings"],
                "diagnostics": diagnostics,
            })
        else:
            log_error(f"Walidacja błędy: {filepath} - {validation['errors']}")
            return jsonify({
                "success": False,
                "errors": validation["errors"],
                "warnings": validation["warnings"],
            })

    except Exception as e:
        log_error(f"Błąd walidacji: {str(e)}")
        # Stwórz raport diagnostyczny
        create_diagnostic_report(
            e,
            input_file_path=filepath,
            context={
                "operation": "validate",
                "input_file": filepath,
            }
        )
        return jsonify({
            "error": f"Błąd walidacji: {str(e)}. Raport diagnostyki został zapisany.",
            "has_diagnostic_report": True,
        }), 500


@app.route("/api/generate", methods=["POST"])
def generate_schedule():
    """Generuj harmonogram"""
    data = request.get_json()
    filepath = data.get("filepath")

    if not filepath or not Path(filepath).exists():
        return jsonify({"error": "Plik nie znaleziony"}), 400

    try:
        # Waliduj
        validation = get_validation_errors(filepath)
        if not validation["success"]:
            return jsonify({
                "error": "Plik zawiera błędy - nie można wygenerować harmonogramu",
                "errors": validation["errors"],
            }), 400

        # Załaduj i rozwiąż
        load_result = load_config_from_excel(filepath)
        config = load_result.config

        result = solve_schedule(config)

        # Przygotuj wynik: pelny miesiac, z pustymi dyzurami dla dni wolnych.
        assignments_by_day = {
            day: (main or "", extra or "")
            for day, _weekday, main, extra in result.schedule_rows
        }
        day_names = ["Pon", "Wt", "Śr", "Czw", "Pt", "Sob", "Niedz"]
        last_day = monthrange(config.year, config.month)[1]
        schedule_data = []
        for d in range(1, last_day + 1):
            day = datetime(config.year, config.month, d).date()
            main, extra = assignments_by_day.get(day, ("", ""))
            schedule_data.append({
                "date": day.isoformat(),
                "weekday": day_names[day.weekday()],
                "main": main,
                "extra": extra,
            })

        summary_data = [
            {
                "employee": row[0],
                "total": row[1],
                "main": row[2],
                "extra": row[3],
                "friday": row[4],
            }
            for row in result.summary_rows
        ]

        # Zapisz Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"harmonogram_{config.year}.{config.month:02d}_{timestamp}.xlsx"
        output_path = DOWNLOAD_FOLDER / output_filename

        export_schedule_to_excel(result, str(output_path), config=config)

        log_info(f"Harmonogram wygenerowany: {output_filename}")

        return jsonify({
            "success": True,
            "status": localize_solver_status(result.status_name),
            "objective_value": float(result.objective_value) if result.objective_value else None,
            "schedule": schedule_data,
            "summary": summary_data,
            "download_filename": output_filename,
        })

    except Exception as e:
        log_error(f"Błąd generowania harmonogramu: {str(e)}")
        # Stwórz raport diagnostyczny
        create_diagnostic_report(
            e,
            input_file_path=filepath,
            context={
                "operation": "generate_schedule",
                "input_file": filepath,
            }
        )
        return jsonify({
            "error": f"Błąd generowania: {str(e)}. Raport diagnostyki został zapisany.",
            "has_diagnostic_report": True,
        }), 500


@app.route("/api/download/<filename>", methods=["GET"])
def download_file(filename: str):
    """Pobierz wygenerowany plik"""
    try:
        filepath = DOWNLOAD_FOLDER / filename
        if not filepath.exists():
            return jsonify({"error": "Plik nie znaleziony"}), 404

        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/template", methods=["GET"])
def get_template():
    """Pobierz szablon"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        template_filename = f"szablon_{timestamp}.xlsx"
        template_path = DOWNLOAD_FOLDER / template_filename

        write_template_excel(str(template_path))

        return jsonify({
            "success": True,
            "download_filename": template_filename,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/sample-config", methods=["GET"])
def get_sample_config():
    """Pobierz przykładową konfigurację"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        config_filename = f"przyklad_konfiguracja_{timestamp}.xlsx"
        config_path = DOWNLOAD_FOLDER / config_filename

        export_config_to_excel(june_2026_config(), str(config_path))

        return jsonify({
            "success": True,
            "download_filename": config_filename,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/diagnostics-folder", methods=["GET"])
def open_diagnostics_folder():
    """Otwórz folder diagnostyki na systemie użytkownika"""
    try:
        diag_folder = get_diagnostics_folder()
        diag_folder.mkdir(exist_ok=True)
        
        system = platform.system()
        if system == "Windows":
            os.startfile(str(diag_folder))
        elif system == "Darwin":  # macOS
            subprocess.Popen(["open", str(diag_folder)])
        elif system == "Linux":
            subprocess.Popen(["xdg-open", str(diag_folder)])
        
        return jsonify({
            "success": True,
            "message": f"Folder diagnostyki: {diag_folder}",
            "path": str(diag_folder),
        })
    except Exception as e:
        log_error(f"Błąd otwierania folderu diagnostyki: {str(e)}")
        return jsonify({"error": str(e)}), 500


@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Nie znaleziono"}), 404


@app.errorhandler(500)
def server_error(error):
    return jsonify({"error": "Błąd serwera"}), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5001)





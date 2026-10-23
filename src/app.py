import os
import sys
import socket
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, render_template, request, send_file
from waitress import serve

import philrice_forecasting as model


def resource_path(relative_path: str) -> Path:
    """Return bundled resource path both in development and PyInstaller builds."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).resolve().parent / relative_path


TEMPLATE_DIR = resource_path("templates")
STATIC_DIR = resource_path("static")

app = Flask(
    __name__,
    template_folder=str(TEMPLATE_DIR),
    static_folder=str(STATIC_DIR),
)

RUNTIME_DIR = Path(tempfile.gettempdir()) / "PhilRice_Forecast"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_DIR = RUNTIME_DIR / "output"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

RESULT_FILE = OUTPUT_DIR / "PhilRice_Forecast_Results.xlsx"


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/forecast", methods=["POST"])
def forecast():
    uploaded_file = request.files.get("file")

    if not uploaded_file or uploaded_file.filename == "":
        return render_template(
            "index.html",
            error="Please select an Excel file first."
        )

    if not uploaded_file.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template(
            "index.html",
            error="Please upload an Excel workbook (.xlsx or .xlsm)."
        )

    # Clear prior upload(s) so the forecasting engine sees only the selected workbook.
    for file in UPLOAD_DIR.glob("*"):
        if file.is_file():
            try:
                file.unlink()
            except OSError:
                pass

    input_path = UPLOAD_DIR / Path(uploaded_file.filename).name
    uploaded_file.save(input_path)

    try:
        source, results, skipped, input_mode = model.run_forecasting_from_folder(
            UPLOAD_DIR
        )

        model.verify_results(results)

        if RESULT_FILE.exists():
            RESULT_FILE.unlink()

        model.write_results_excel(results, RESULT_FILE)

        return render_template(
            "index.html",
            success=True,
            results=results,
            skipped=skipped,
            input_mode=input_mode
        )

    except Exception as exc:
        return render_template(
            "index.html",
            error=str(exc)
        )


@app.route("/download")
def download():
    if not RESULT_FILE.exists():
        return "No forecast result is available yet.", 404

    return send_file(
        RESULT_FILE,
        as_attachment=True,
        download_name="PhilRice_Forecast_Results.xlsx"
    )


def find_free_port() -> int:
    """Find a free local TCP port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def open_browser(url: str):
    webbrowser.open(url)


def main():
    port = find_free_port()
    url = f"http://127.0.0.1:{port}"

    threading.Timer(1.2, open_browser, args=(url,)).start()

    serve(
        app,
        host="127.0.0.1",
        port=port,
        threads=4,
    )


if __name__ == "__main__":
    main()

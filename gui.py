"""Browser-based GUI for collecting date range and shift information."""
from __future__ import annotations

import threading
import time
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from flask import Flask, current_app, redirect, render_template, request, url_for, has_app_context
from werkzeug.serving import make_server

app = Flask(__name__, template_folder="templates", static_folder="static")

_DEFAULT_HOST = "127.0.0.1"
_DEFAULT_PORT = 5000


@dataclass
class _FormState:
    result_event: threading.Event
    result: Optional[Tuple[datetime, datetime, str, int]] = None
    server_thread: Optional["_ServerThread"] = None
    review_event: threading.Event = field(default_factory=threading.Event)
    review_data: Optional[Dict[str, List[Dict[str, str]]]] = None
    continue_event: threading.Event = field(default_factory=threading.Event)
    processing_started: bool = False


class _ServerThread(threading.Thread):
    """Runs the Flask development server in a background thread."""

    def __init__(self, flask_app: Flask, host: str, port: int) -> None:
        super().__init__(daemon=True)
        self._app = flask_app
        self._host = host
        self._port = port
        self._server = make_server(host, port, flask_app)

    def run(self) -> None:
        ctx = self._app.app_context()
        ctx.push()
        try:
            self._server.serve_forever()
        finally:
            ctx.pop()

    def shutdown(self) -> None:
        self._server.shutdown()


def _get_form_state() -> _FormState:
    config = current_app.config if has_app_context() else app.config
    state = config.get("FORM_STATE")
    if state is None:
        raise RuntimeError("FORM_STATE is not configured.")
    return state


def _default_form_values() -> Dict[str, str]:
    now = datetime.now()
    kw = now.isocalendar()[1] + 1
    if kw > 53:
        kw = 1
    return {
        "start_date": now.strftime("%Y-%m-%d"),
        "start_hour": now.strftime("%H"),
        "start_minute": "00",
        "end_date": now.strftime("%Y-%m-%d"),
        "end_hour": now.strftime("%H"),
        "end_minute": now.strftime("%M"),
        "kalenderwoche": f"{kw:02d}",
        "schichten_text": "",
    }


def _gather_form_values(form_data) -> Dict[str, str]:
    return {
        "start_date": (form_data.get("start_date") or "").strip(),
        "start_hour": (form_data.get("start_hour") or "").strip(),
        "start_minute": (form_data.get("start_minute") or "").strip(),
        "end_date": (form_data.get("end_date") or "").strip(),
        "end_hour": (form_data.get("end_hour") or "").strip(),
        "end_minute": (form_data.get("end_minute") or "").strip(),
        "kalenderwoche": (form_data.get("kalenderwoche") or "").strip(),
        "schichten_text": (form_data.get("schichten_text") or "").strip(),
    }


def _build_datetime(date_str: str, hour_str: str, minute_str: str) -> datetime:
    if not date_str:
        raise ValueError("Bitte ein Datum auswaehlen.")
    if not hour_str or not minute_str:
        raise ValueError("Bitte Stunde und Minute angeben.")
    try:
        hour = int(hour_str)
        minute = int(minute_str)
    except ValueError as exc:
        raise ValueError("Stunde und Minute muessen Zahlen sein.") from exc
    if not 0 <= hour <= 23:
        raise ValueError("Stunde muss zwischen 0 und 23 liegen.")
    if not 0 <= minute <= 59:
        raise ValueError("Minute muss zwischen 0 und 59 liegen.")
    try:
        return datetime.strptime(f"{date_str} {hour:02d}:{minute:02d}", "%Y-%m-%d %H:%M")
    except ValueError as exc:
        raise ValueError("Ungueltige Datums- oder Zeitangabe.") from exc


def _parse_kw(value: str) -> int:
    if not value:
        raise ValueError("Bitte eine Kalenderwoche angeben.")
    try:
        kw = int(value)
    except ValueError as exc:
        raise ValueError("Kalenderwoche muss eine Zahl sein.") from exc
    if not 1 <= kw <= 53:
        raise ValueError("Kalenderwoche muss zwischen 1 und 53 liegen.")
    return kw


def _trigger_server_shutdown(state: _FormState) -> None:
    server = state.server_thread
    if not server:
        return

    def _delayed_shutdown() -> None:
        time.sleep(1)
        server.shutdown()

    threading.Thread(target=_delayed_shutdown, daemon=True).start()


@app.route("/", methods=["GET"])
def show_form():
    values = _default_form_values()
    return render_template("index.html", values=values, error=None, success=False)


@app.route("/submit", methods=["POST"])
def submit_form():
    state = _get_form_state()
    values = _gather_form_values(request.form)
    try:
        start = _build_datetime(values["start_date"], values["start_hour"], values["start_minute"])
        end = _build_datetime(values["end_date"], values["end_hour"], values["end_minute"])
        if end < start:
            raise ValueError("Enddatum darf nicht vor dem Startdatum liegen.")
        kw = _parse_kw(values["kalenderwoche"])
    except ValueError as exc:
        return (
            render_template("index.html", values=values, error=str(exc), success=False),
            400,
        )

    state.processing_started = False
    state.review_data = None
    state.review_event.clear()
    state.continue_event.clear()
    state.result = (start, end, values["schichten_text"], kw)
    state.result_event.set()
    return redirect(url_for("review_mails"))


@app.route("/review", methods=["GET"])
def review_mails():
    state = _get_form_state()
    state.review_event.wait()
    data = state.review_data or {"unknown": [], "known": []}
    return render_template(
        "review.html",
        unknown_mails=data["unknown"],
        known_mails=data["known"],
        processing_started=state.processing_started,
    )


@app.route("/review/continue", methods=["POST"])
def review_continue():
    state = _get_form_state()
    state.processing_started = True
    state.continue_event.set()
    data = state.review_data or {"unknown": [], "known": []}
    return render_template(
        "review.html",
        unknown_mails=data["unknown"],
        known_mails=data["known"],
        processing_started=True,
    )


def zeitraum_abfragen(host: str = _DEFAULT_HOST, port: int = _DEFAULT_PORT):
    """Start the browser UI, wait for user input, and return the collected data."""
    state = _FormState(result_event=threading.Event())
    app.config["FORM_STATE"] = state

    server_thread = _ServerThread(app, host, port)
    state.server_thread = server_thread
    server_thread.start()

    url = f"http://{host}:{port}/"
    try:
        webbrowser.open_new(url)
    except Exception:
        pass

    state.result_event.wait()
    return state.result


def zeige_mail_uebersicht(mails) -> None:
    """Expose fetched mails in a review page and wait for user confirmation."""
    state = _get_form_state()
    state.review_event.clear()
    state.continue_event.clear()
    state.processing_started = False

    def _format_entry(mail) -> Dict[str, str]:
        received = mail.empfangen or datetime.min
        empfangen_display = (
            received.strftime("%d.%m.%Y %H:%M") if mail.empfangen else "Kein Datum"
        )
        return {
            "mailadresse": str(mail.mailadresse or "").strip(),
            "betreff": str(mail.betreff or ""),
            "mailtext": str(mail.mailtext or ""),
            "empfangen_display": empfangen_display,
        }

    sorted_mails = sorted(
        mails,
        key=lambda mail: mail.empfangen or datetime.min,
        reverse=True,
    )

    unknown = [_format_entry(mail) for mail in sorted_mails if not (mail.kurzname or "").strip()]
    known = [_format_entry(mail) for mail in sorted_mails if (mail.kurzname or "").strip()]

    state.review_data = {"unknown": unknown, "known": known}
    state.review_event.set()

    state.continue_event.wait()


def stop_gui_server() -> None:
    """Shutdown the background Flask server if it is still running."""
    state = app.config.get("FORM_STATE")
    if not state:
        return

    _trigger_server_shutdown(state)
    server_thread = state.server_thread
    if server_thread and server_thread.is_alive():
        server_thread.join(timeout=5)

    app.config.pop("FORM_STATE", None)

"""Flask web application for handling the Swaxi mail workflow."""
from __future__ import annotations

import asyncio
import os
from dataclasses import dataclass, field
from datetime import datetime
from collections.abc import Iterable
from typing import Dict, List, Optional

from flask import Flask, redirect, render_template, request, url_for

from llm import build_instructions, process_mails_parallel
from mail import MailJSON, lade_mails_aus_outlook, kuerze_mailtexte, weise_kurznamen_zu
from mitarbeiter import Mitarbeiter, anonymisiere_mailtexte, lade_mitarbeiter
from shifts import erzeuge_schicht_tabellen

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")


@dataclass
class WorkflowState:
    start: Optional[datetime] = None
    end: Optional[datetime] = None
    kalenderwoche: Optional[int] = None
    available_shifts: str = ""
    mitarbeiter: List[Mitarbeiter] = field(default_factory=list)
    mails: List[MailJSON] = field(default_factory=list)


_STATE = WorkflowState()


def _reset_state() -> None:
    global _STATE
    _STATE = WorkflowState()


def _get_state() -> WorkflowState:
    return _STATE


def _default_form_values(state: Optional[WorkflowState] = None) -> Dict[str, str]:
    now = datetime.now()
    kw = now.isocalendar()[1] + 1
    if kw > 53:
        kw = 1
    values = {
        "start_date": now.strftime("%Y-%m-%d"),
        "start_hour": now.strftime("%H"),
        "start_minute": "00",
        "end_date": now.strftime("%Y-%m-%d"),
        "end_hour": now.strftime("%H"),
        "end_minute": now.strftime("%M"),
        "kalenderwoche": f"{kw:02d}",
        "schichten_text": "",
    }
    if not state:
        return values
    if state.start:
        values["start_date"] = state.start.strftime("%Y-%m-%d")
        values["start_hour"] = state.start.strftime("%H")
        values["start_minute"] = state.start.strftime("%M")
    if state.end:
        values["end_date"] = state.end.strftime("%Y-%m-%d")
        values["end_hour"] = state.end.strftime("%H")
        values["end_minute"] = state.end.strftime("%M")
    if state.kalenderwoche:
        values["kalenderwoche"] = f"{state.kalenderwoche:02d}"
    if state.available_shifts:
        values["schichten_text"] = state.available_shifts
    return values


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


def _format_datetime(dt: Optional[datetime]) -> str:
    if not dt:
        return "--"
    return dt.strftime("%d.%m.%Y %H:%M")


def _format_review_lists(mails: List[MailJSON]) -> Dict[str, List[Dict[str, str]]]:
    def _format_entry(mail: MailJSON) -> Dict[str, str]:
        received = mail.empfangen or datetime.min
        received_label = mail.empfangen.strftime("%d.%m.%Y %H:%M") if mail.empfangen else "Kein Datum"
        return {
            "mailadresse": str(mail.mailadresse or "").strip(),
            "betreff": str(mail.betreff or ""),
            "mailtext": str(mail.mailtext or ""),
            "empfangen_display": received_label,
            "kurzname": str(getattr(mail, "kurzname", "") or ""),
            "mail_id": mail.mail_id,
            "sort_key": received,
        }

    formatted = [_format_entry(mail) for mail in mails]
    formatted.sort(key=lambda item: (item["sort_key"], item["mail_id"]), reverse=True)
    unknown = [item for item in formatted if not item["kurzname"].strip()]
    known = [item for item in formatted if item["kurzname"].strip()]
    for item in formatted:
        item.pop("sort_key", None)
    return {"unknown": unknown, "known": known}


def _value_has_content(value) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, dict):
        return any(_value_has_content(v) for v in value.values())
    if isinstance(value, Iterable):
        return any(_value_has_content(item) for item in value)
    if isinstance(value, bool):
        return value
    return True


def _mail_has_schichtwuensche(mail: MailJSON) -> bool:
    schichtwuensche = getattr(mail, "schichtwuensche", None)
    if not isinstance(schichtwuensche, dict):
        return False
    return any(_value_has_content(value) for value in schichtwuensche.values())


def _format_mail_entries(mails: List[MailJSON]) -> List[Dict[str, str]]:
    entries = []
    for mail in mails:
        received = mail.empfangen or datetime.min
        mail_id_value = getattr(mail, "mail_id", "")
        try:
            mail_id_sort = int(mail_id_value)
        except (TypeError, ValueError):
            mail_id_sort = 0
        entries.append(
            {
                "mailadresse": str(mail.mailadresse or "").strip(),
                "betreff": str(mail.betreff or ""),
                "mailtext": str(mail.mailtext or ""),
                "empfangen_display": mail.empfangen.strftime("%d.%m.%Y %H:%M") if mail.empfangen else "Kein Datum",
                "mail_id": str(mail_id_value),
                "kurzname": str(getattr(mail, "kurzname", "") or ""),
                "sort_key": received,
                "mail_id_sort": mail_id_sort,
            }
        )
    entries.sort(key=lambda item: (item["sort_key"], item.get("mail_id_sort", 0)), reverse=True)
    for entry in entries:
        entry.pop("sort_key", None)
        entry.pop("mail_id_sort", None)
    return entries


def _render_review(state: WorkflowState, error: Optional[str] = None, status: Optional[str] = None, status_level: str = "info"):
    lists = _format_review_lists(state.mails)
    return render_template(
        "review.html",
        unknown_mails=lists["unknown"],
        known_mails=lists["known"],
        error=error,
        status=status,
        status_level=status_level,
        timeframe=f"{_format_datetime(state.start)} – {_format_datetime(state.end)}",
        kalenderwoche=state.kalenderwoche,
    )


@app.route("/", methods=["GET"])
def show_form():
    state = _get_state()
    if request.args.get("reset") == "1":
        _reset_state()
        state = _get_state()
    values = _default_form_values(state)
    return render_template("index.html", values=values, error=None, success=False)


@app.route("/submit", methods=["POST"])
def submit_form():
    state = _get_state()
    values = _gather_form_values(request.form)
    try:
        start = _build_datetime(values["start_date"], values["start_hour"], values["start_minute"])
        end = _build_datetime(values["end_date"], values["end_hour"], values["end_minute"])
        if end < start:
            raise ValueError("Enddatum darf nicht vor dem Startdatum liegen.")
        kw = _parse_kw(values["kalenderwoche"])
    except ValueError as exc:
        return render_template("index.html", values=values, error=str(exc), success=False), 400

    state.start = start
    state.end = end
    state.kalenderwoche = kw
    state.available_shifts = values["schichten_text"]
    state.mails.clear()

    try:
        mitarbeiter = lade_mitarbeiter()
        mails = lade_mails_aus_outlook(start, end)
        mails = kuerze_mailtexte(mails)
        mails = weise_kurznamen_zu(mails, mitarbeiter)
    except Exception as exc:  # noqa: BLE001 - want to surface full error for manual handling
        app.logger.exception("Fehler bei der Mailverarbeitung")
        return (
            render_template(
                "index.html",
                values=values,
                error=f"Fehler beim Einlesen der Mails: {exc}",
                success=False,
            ),
            500,
        )

    state.mitarbeiter = mitarbeiter
    state.mails = mails

    return redirect(url_for("review_mails"))


@app.route("/review", methods=["GET"])
def review_mails():
    state = _get_state()
    if not state.mails:
        return redirect(url_for("show_form"))
    return _render_review(state)


@app.route("/review/continue", methods=["POST"])
def review_continue():
    state = _get_state()
    if not state.mails:
        return redirect(url_for("show_form"))

    mails_processable = [mail for mail in state.mails if getattr(mail, "kurzname", "").strip()]
    if not mails_processable:
        return _render_review(state, error="Keine zu verarbeitenden Mails gefunden.")

    try:
        mails_processable = anonymisiere_mailtexte(mails_processable, state.mitarbeiter)
        instructions = build_instructions(state.available_shifts)
        asyncio.run(process_mails_parallel(instructions, mails_processable))
        state.mails = list(mails_processable)
    except Exception as exc:  # noqa: BLE001 - propagate meaningful failure message to the UI
        app.logger.exception("Fehler in der Verarbeitung")
        return _render_review(state, error=f"Verarbeitung fehlgeschlagen: {exc}")

    return redirect(url_for("show_results"))


@app.route("/results", methods=["GET"])
def show_results():
    state = _get_state()
    if not state.mails:
        return redirect(url_for("review_mails"))
    sorted_mails = sorted(
        state.mails,
        key=lambda mail: (mail.empfangen or datetime.min, getattr(mail, "mail_id", 0)),
        reverse=True,
    )

    processed_raw = [
        mail for mail in sorted_mails if getattr(mail, "ist_schichtwunsch", False) and _mail_has_schichtwuensche(mail)
    ]
    recognized_not_processed_raw = [
        mail
        for mail in sorted_mails
        if getattr(mail, "ist_schichtwunsch", False)
        and not _mail_has_schichtwuensche(mail)
        and not _value_has_content(getattr(mail, "kommentar", None))
    ]
    non_shift_raw = [mail for mail in sorted_mails if not getattr(mail, "ist_schichtwunsch", False)]

    shifts_table_raw, mail_ids_table_raw = erzeuge_schicht_tabellen(sorted_mails)
    shifts_headers = list(shifts_table_raw[0]) if shifts_table_raw else []
    shifts_rows = []
    if shifts_table_raw:
        for row_index in range(1, len(shifts_table_raw)):
            row = []
            for col_index, _ in enumerate(shifts_headers):
                cell_value = shifts_table_raw[row_index][col_index]
                cell_mail_id = ""
                if row_index < len(mail_ids_table_raw) and col_index < len(mail_ids_table_raw[row_index]):
                    raw_mail_id = mail_ids_table_raw[row_index][col_index]
                    cell_mail_id = str(raw_mail_id or "")
                    if cell_mail_id == "None":
                        cell_mail_id = ""
                row.append({"value": cell_value, "mail_id": cell_mail_id})
            shifts_rows.append(row)

    has_shift_entries = any(cell["value"] for row in shifts_rows for cell in row)

    return render_template(
        "results.html",
        processed_mails=_format_mail_entries(processed_raw),
        recognized_not_processed_mails=_format_mail_entries(recognized_not_processed_raw),
        non_shift_mails=_format_mail_entries(non_shift_raw),
        shifts_headers=shifts_headers,
        shifts_rows=shifts_rows,
        has_shift_entries=has_shift_entries,
    )


@app.route("/reset", methods=["POST"])
def reset_session():
    _reset_state()
    return redirect(url_for("show_form"))

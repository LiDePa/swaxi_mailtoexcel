"""Hilfsfunktionen für Schichtwunsch-Tabellen."""

from collections import defaultdict
from typing import Dict, Iterable, List, Sequence, Tuple


_WOCHENTAGE: Tuple[str, ...] = (
	"Mittwoch",
	"Donnerstag",
	"Freitag",
	"Samstag",
	"Sonntag",
	"Montag",
	"Dienstag",
)


def _to_sequence(value) -> Sequence[str]:
	"""Normalisiert beliebige Schichtwerte zu einer Sequenz von Strings."""
	if value is None:
		return ()
	if isinstance(value, str):
		parts = [p.strip() for p in value.replace(";", ",").split(",")]
		clean = [p for part in parts for p in part.split() if p]
		return tuple(clean)
	if isinstance(value, dict):
		return tuple(str(k).strip() for k, v in value.items() if v)
	if isinstance(value, Iterable):
		return tuple(str(item).strip() for item in value if str(item).strip())
	return (str(value).strip(),)


def _sanitize_shift(shift: str) -> str:
	return shift.replace(".", "").strip()


def erzeuge_schicht_tabellen(mails) -> Tuple[List[List[str]], List[List[str]], List[Dict[str, str]]]:
	"""Erstellt Schicht- und Mail-ID-Tabellen sowie Kommentarlisten."""
	spalteninhalte = defaultdict(list)
	mail_id_spalten = defaultdict(list)
	kommentar_map: Dict[str, str] = {}
	kommentare_liste: List[Dict[str, str]] = []

	for mail in mails or []:
		if not getattr(mail, "ist_schichtwunsch", False):
			continue

		schichtwuensche = getattr(mail, "schichtwuensche", {}) or {}
		if not isinstance(schichtwuensche, dict):
			continue
		kurzname = str(getattr(mail, "kurzname", "")).strip()
		mail_id = getattr(mail, "mail_id", "")

		kommentar = str(getattr(mail, "kommentar", "") or "").strip()
		if kommentar:
			if kurzname:
				kommentar_map[kurzname] = kommentar
			kommentare_liste.append({"kurzname": kurzname, "kommentar": kommentar, "mail_id": str(mail_id)})

		for tag in _WOCHENTAGE:
			raw_shift_values = schichtwuensche.get(tag)
			shifts = []
			for item in _to_sequence(raw_shift_values):
				cleaned = _sanitize_shift(item)
				if cleaned:
					shifts.append(cleaned)
			if not shifts:
				continue

			eintrag = f"{kurzname} {','.join(shifts)}".strip()
			spalteninhalte[tag].append(eintrag)
			mail_id_spalten[tag].append(str(mail_id))

	max_zeilen = max((len(spalteninhalte[tag]) for tag in _WOCHENTAGE), default=0)

	shifts_table = [list(_WOCHENTAGE)]
	mail_ids_table = [list(_WOCHENTAGE)]

	for zeile in range(max_zeilen):
		shifts_row = []
		mail_ids_row = []
		for tag in _WOCHENTAGE:
			if zeile < len(spalteninhalte[tag]):
				shifts_row.append(spalteninhalte[tag][zeile])
				mail_ids_row.append(mail_id_spalten[tag][zeile])
			else:
				shifts_row.append("")
				mail_ids_row.append("")
		shifts_table.append(shifts_row)
		mail_ids_table.append(mail_ids_row)

	return shifts_table, mail_ids_table, kommentare_liste


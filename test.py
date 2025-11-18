"""Utility script to build a dummy MailJSON list for local testing."""

from __future__ import annotations

import hashlib
import pickle
from datetime import datetime
from pathlib import Path

from mail import lade_mails_aus_outlook, kuerze_mailtexte, weise_kurznamen_zu
from mitarbeiter import anonymisiere_mailtexte, lade_mitarbeiter

try:
	import pywintypes
except ImportError:  # pragma: no cover - pywintypes always available in runtime env
	pywintypes = None


# Placeholder timeframe; adjust both values before running the script.
START_DATETIME = datetime(2025, 7, 30, 14, 25)
END_DATETIME = datetime(2025, 8, 4, 11, 30)

OUTPUT_FILE = Path(__file__).resolve().parent / "dummymails.pkl"


def main() -> None:
	print(f"Lade Mitarbeiterliste...")
	mitarbeiter = lade_mitarbeiter()

	print(f"Lese Mails aus Outlook zwischen {START_DATETIME} und {END_DATETIME}...")
	mails = lade_mails_aus_outlook(START_DATETIME, END_DATETIME)
	if not mails:
		print("Keine Mails im angegebenen Zeitraum gefunden. Abbruch ohne Datei.")
		return

	print("Bereinige Mailtexte...")
	mails = kuerze_mailtexte(mails)
	mails = weise_kurznamen_zu(mails, mitarbeiter)

	print("Anonymisiere Mailinhalte...")
	mails = anonymisiere_mailtexte(mails, mitarbeiter)

	_hash_cache_mail: dict[str, str] = {}
	_hash_cache_kurzname: dict[str, str] = {}

	def _hash_value(value: str, prefix: str, cache: dict[str, str]) -> str:
		key = (value or "").strip().lower()
		if key in cache:
			return cache[key]
		if not key:
			cache[key] = ""
			return ""
		digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:8]
		label = f"{prefix}_{digest}"
		cache[key] = label
		return label

	for mail in mails:
		if getattr(mail, "mailtext_anonym", None):
			mail.mailtext = mail.mailtext_anonym
		mail.mailadresse = _hash_value(str(mail.mailadresse or ""), "mail", _hash_cache_mail) or "mailadress_anon"
		mail.kurzname = _hash_value(str(getattr(mail, "kurzname", "")), "anon", _hash_cache_kurzname) or getattr(mail, "kurzname", "")

	if pywintypes is not None:
		for mail in mails:
			received = getattr(mail, "empfangen", None)
			if isinstance(received, pywintypes.TimeType):
				mail.empfangen = datetime(
					received.year,
					received.month,
					received.day,
					received.hour,
					received.minute,
					received.second,
					received.microsecond,
				)

	print(f"Speichere {len(mails)} Mails nach {OUTPUT_FILE}...")
	with OUTPUT_FILE.open("wb") as handle:
		breakpoint()
		pickle.dump(mails, handle)

	print("Fertig.")


if __name__ == "__main__":
	main()


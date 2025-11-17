import re
import locale
import win32com.client
import pywintypes
from collections import OrderedDict
from datetime import datetime


class OutlookAccessError(RuntimeError):
	"""Raised when Outlook cannot be accessed or required folders are missing."""



class MailJSON:
	def __init__(self, mail_id, mailadresse, betreff, mailtext, empfangen=None):
		self.mail_id = mail_id
		self.kurzname = ""
		self.mailadresse = mailadresse
		self.betreff = betreff
		self.mailtext = mailtext
		self.mailtext_anonym = ""
		self.ist_schichtwunsch = True
		self.schichtwuensche = {}
		self.kommentar = ""
		self.empfangen = empfangen
	def __repr__(self):
		return f"<MailJSON {self.mail_id}: {self.betreff} von {self.mailadresse}>"


def _detect_outlook_datetime_format():
	"""Map Windows locale to an Outlook Restrict-friendly datetime pattern."""
	try:
		original = locale.setlocale(locale.LC_TIME)
	except locale.Error:
		original = None
	try:
		user_locale = locale.setlocale(locale.LC_TIME, '')
	except locale.Error:
		user_locale = original or ''
	finally:
		if original:
			try:
				locale.setlocale(locale.LC_TIME, original)
			except locale.Error:
				pass
	locale_token = (user_locale or '').lower()
	if locale_token.startswith(('de', 'german')):
		return "%d.%m.%Y %H:%M"
	return "%m/%d/%Y %I:%M %p"


_OUTLOOK_DATETIME_FORMAT = _detect_outlook_datetime_format()

# Hilfsfunktion: Suche Ordner nach Name (case-insensitive, optional rekursiv)
def finde_ordner(parent_folder, ordnername, rekursiv=False):
	ordnername = ordnername.lower()
	for folder in parent_folder.Folders:
		if folder.Name.lower() == ordnername:
			return folder
		if rekursiv:
			found = finde_ordner(folder, ordnername, rekursiv)
			if found:
				return found
	return None

def lade_mails_aus_outlook(startdatum, enddatum):
	try:
		outlook_app = win32com.client.Dispatch("Outlook.Application")
	except pywintypes.com_error as exc:
		raise OutlookAccessError("Outlook konnte nicht gestartet werden. Bitte Outlook oeffnen und erneut versuchen.") from exc

	try:
		outlook = outlook_app.GetNamespace("MAPI")
	except pywintypes.com_error as exc:
		raise OutlookAccessError("Outlook-Postfach konnte nicht geladen werden. Bitte Outlook neu starten und erneut versuchen.") from exc
	# Freigegebenes Postfach holen
	konto = None
	for acc in outlook.Folders:
		if acc.Name.lower() == "swaxi dispo":
			konto = acc
			break
	if not konto:
		raise OutlookAccessError("Postfach 'Swaxi Dispo' wurde nicht gefunden. Bitte Berechtigungen pruefen und Outlook neu starten.")
	# Nur Mails aus Inbox und Archive > Dienstwünsche / FZ
	mails = []

	# Outlook Restrict-Filter für Datum (ReceivedTime)
	def restrict_filter(start, end):
		start_str = start.strftime(_OUTLOOK_DATETIME_FORMAT)
		end_str = end.strftime(_OUTLOOK_DATETIME_FORMAT)
		return f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"

	# Inbox suchen (case-insensitive)
	inbox = finde_ordner(konto, "Inbox") or finde_ordner(konto, "Posteingang")
	if inbox:
		try:
			items = inbox.Items.Restrict(restrict_filter(startdatum, enddatum))
			items.Sort("ReceivedTime", True)
			for item in items:
				if hasattr(item, "ReceivedTime"):
					empfangen = item.ReceivedTime
					# Offset-naive Vergleich
					if empfangen.tzinfo is not None:
						empfangen = empfangen.replace(tzinfo=None)
					if startdatum <= empfangen <= enddatum:
						mail_id = len(mails) + 1
						mailadresse = item.SenderEmailAddress
						betreff = item.Subject
						mailtext = item.Body
						mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext, empfangen))
		except pywintypes.com_error as exc:
			raise OutlookAccessError("Auf die Outlook-Inbox konnte nicht zugegriffen werden. Bitte Outlook neu starten und erneut versuchen.") from exc
	else:
		raise OutlookAccessError("Ordner 'Inbox' bzw. 'Posteingang' wurde im Postfach 'Swaxi Dispo' nicht gefunden. Bitte Outlook-Konfiguration pruefen.")

	# Archive > Dienstwünsche / FZ suchen (rekursiv)
	dienstwuensche_fz = None
	archive = finde_ordner(konto, "Archive")
	if archive:
		dienstwuensche_fz = finde_ordner(archive, "Dienstwünsche / FZ")
	if dienstwuensche_fz:
		try:
			items = dienstwuensche_fz.Items.Restrict(restrict_filter(startdatum, enddatum))
			items.Sort("ReceivedTime", True)
			for item in items:
				if hasattr(item, "ReceivedTime"):
					empfangen = item.ReceivedTime
					if empfangen.tzinfo is not None:
						empfangen = empfangen.replace(tzinfo=None)
					if startdatum <= empfangen <= enddatum:
						mail_id = len(mails) + 1
						mailadresse = item.SenderEmailAddress
						betreff = item.Subject
						mailtext = item.Body
						mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext, empfangen))
		except pywintypes.com_error as exc:
			raise OutlookAccessError("Auf den Ordner 'Archive > Dienstwünsche / FZ' konnte nicht zugegriffen werden. Bitte Outlook neu starten und erneut versuchen.") from exc
	else:
		print("Ordner 'Archive > Dienstwünsche / FZ' nicht gefunden!")

	print(f"{len(mails)} Mails im Zeitraum gefunden.")
	return mails


def kuerze_mailtexte(mailjson_liste):
	"""Entfernt Footer-Anteile und uebermaessige Leerzeilen aus allen Mailtexten."""
	if not mailjson_liste:
		return mailjson_liste

	kuerzungs_pattern = re.compile(r"^(.*(?:Swaxi Dispo|swaxi-dispo@sw-augsburg\.de).*)$", flags=re.MULTILINE | re.IGNORECASE)
	zeilenumbruch_pattern = re.compile(r"(\r\n|\n|\r){2,}")

	for mail in mailjson_liste:
		mail_text_bearbeitet = mail.mailtext or ""

		match = kuerzungs_pattern.search(mail_text_bearbeitet)
		if match:
			mail_text_bearbeitet = mail_text_bearbeitet[:match.start()]

		mail_text_bearbeitet = "\n".join(
			line for line in mail_text_bearbeitet.splitlines() if "https://" not in line
		)

		mail_text_bearbeitet = zeilenumbruch_pattern.sub("\n", mail_text_bearbeitet)
		mail.mailtext = mail_text_bearbeitet.strip()

	return mailjson_liste


# Neue Funktion: Weise jedem MailJSON-Objekt den kurzname des passenden Mitarbeiters zu
def weise_kurznamen_zu(mailjson_liste, mitarbeiter_liste):
	"""
	Ergänzt für jedes MailJSON-Objekt den kurzname des zugehörigen Mitarbeiters anhand der mailadresse
	und führt mehrere Mails desselben Mitarbeiters zusammen. Neuere Inhalte werden ans Ende des
	bereits bereinigten mailtextes angehängt.
	"""
	if not mailjson_liste:
		return mailjson_liste

	mail_map = {}
	for m in mitarbeiter_liste:
		if not m.mailadressen:
			continue
		for addr in str(m.mailadressen).replace(';', ',').split(','):
			addr = addr.strip().lower()
			if addr:
				mail_map[addr] = m.kurzname

	for mail in mailjson_liste:
		mail_addr = str(mail.mailadresse).strip().lower()
		if mail_addr in mail_map:
			mail.kurzname = mail_map[mail_addr]
		else:
			for k, v in mail_map.items():
				if k and k in mail_addr:
					mail.kurzname = v
					break

	# Sortiere nach Empfangszeit, damit jüngere Mails später angehängt werden.
	# Der Enumerationsindex stellt sicher, dass wir bei fehlender Zeitangabe stabil bleiben.
	indexed_sorted_mails = sorted(
		enumerate(mailjson_liste),
		key=lambda pair: (
			pair[1].empfangen or datetime.min,
			pair[0],
		),
	)

	combined = OrderedDict()
	separator = "\n---\n"

	for _, mail in indexed_sorted_mails:
		# Verwende IMMER die Mailadresse als Schlüssel (lowercased, stripped)
		key = str(mail.mailadresse).strip().lower()
		if key not in combined:
			combined[key] = mail
			continue

		existing = combined[key]
		if mail.mailtext_anonym:
			if existing.mailtext_anonym:
				existing.mailtext_anonym = f"{existing.mailtext_anonym}{separator}{mail.mailtext_anonym}"
			else:
				existing.mailtext_anonym = mail.mailtext_anonym
		if mail.mailtext:
			if existing.mailtext:
				existing.mailtext = f"{existing.mailtext}{separator}{mail.mailtext}"
			else:
				existing.mailtext = mail.mailtext

	combined_list = list(combined.values())
	for new_id, mail in enumerate(combined_list, start=1):
		mail.mail_id = new_id

	return combined_list





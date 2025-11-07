import win32com.client
from collections import OrderedDict
from datetime import datetime


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
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	# Freigegebenes Postfach holen
	konto = None
	for acc in outlook.Folders:
		if acc.Name.lower() == "swaxi dispo":
			konto = acc
			break
	if not konto:
		print("Postfach 'Swaxi Dispo' nicht gefunden!")
		return []
	# Nur Mails aus Inbox und Archive > Dienstwünsche / FZ
	mails = []

	# Outlook Restrict-Filter für Datum (ReceivedTime)
	# Nur Datumsteil (mm/dd/yyyy), Uhrzeit wird in Python gefiltert
	def restrict_filter(start):
		start_str = start.strftime('%m/%d/%Y')
		return f"[ReceivedTime] >= '{start_str}'"

	# Inbox suchen (case-insensitive)
	inbox = finde_ordner(konto, "Inbox")
	if inbox:
		try:
			items = inbox.Items.Restrict(restrict_filter(startdatum))
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
		except Exception as e:
			print(f"Fehler beim Zugriff auf Inbox: {e}")
	else:
		print("Ordner 'Inbox' nicht gefunden!")

	# Archive > Dienstwünsche / FZ suchen (rekursiv)
	dienstwuensche_fz = None
	archive = finde_ordner(konto, "Archive")
	if archive:
		dienstwuensche_fz = finde_ordner(archive, "Dienstwünsche / FZ")
	if dienstwuensche_fz:
		try:
			items = dienstwuensche_fz.Items.Restrict(restrict_filter(startdatum))
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
		except Exception as e:
			print(f"Fehler beim Zugriff auf Dienstwünsche / FZ: {e}")
	else:
		print("Ordner 'Archive > Dienstwünsche / FZ' nicht gefunden!")

	print(f"{len(mails)} Mails im Zeitraum gefunden.")
	return mails


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
		# Verwende kurzname als Schlüssel, falle auf mailadresse zurück, falls kein Match.
		key = mail.kurzname or str(mail.mailadresse).strip().lower()
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





import win32com.client
from datetime import datetime


class MailJSON:
	def __init__(self, mail_id, mailadresse, betreff, mailtext):
		self.Mail_ID = mail_id
		self.Kurzname = ""
		self.Mailadresse = mailadresse
		self.Betreff = betreff
		self.Mailtext = mailtext
		self.Mailtext_anonym = ""
		self.Schichtwünsche = ""
		self.Kommentar = ""
	def __repr__(self):
		return f"<MailJSON {self.Mail_ID}: {self.Betreff} von {self.Mailadresse}>"

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
						mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext))
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
						mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext))
		except Exception as e:
			print(f"Fehler beim Zugriff auf Dienstwünsche / FZ: {e}")
	else:
		print("Ordner 'Archive > Dienstwünsche / FZ' nicht gefunden!")

	print(f"{len(mails)} Mails im Zeitraum gefunden.")
	return mails


# Neue Funktion: Weise jedem MailJSON-Objekt den Kurzname des passenden Mitarbeiters zu
def weise_kurznamen_zu(mailjson_liste, mitarbeiter_liste):
	"""
	Ergänzt für jedes MailJSON-Objekt den Kurzname des zugehörigen Mitarbeiters anhand der Mailadresse.
	Die Zuordnung erfolgt case-insensitive und unterstützt ggf. mehrere Mailadressen pro Mitarbeiter (Komma/Strichpunkt-getrennt).
	"""
	# Mapping: Mailadresse (lower) -> Kurzname
	mail_map = {}
	for m in mitarbeiter_liste:
		if not m.mailadressen:
			continue
		# Mailadressen können durch Komma oder Semikolon getrennt sein
		for addr in str(m.mailadressen).replace(';', ',').split(','):
			addr = addr.strip().lower()
			if addr:
				mail_map[addr] = m.kurzname
	for mail in mailjson_liste:
		mail_addr = str(mail.Mailadresse).strip().lower()
		# Versuche exakten Match
		if mail_addr in mail_map:
			mail.Kurzname = mail_map[mail_addr]
		else:
			# Fallback: Teilstring-Match (z.B. falls Outlook-Adressen abweichen)
			for k, v in mail_map.items():
				if k and k in mail_addr:
					mail.Kurzname = v
					break
	return mailjson_liste





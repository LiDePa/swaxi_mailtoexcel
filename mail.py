import win32com.client
from datetime import datetime



# Beispielhafte Zeitpunkte (können später durch GUI ersetzt werden)
START_DATUM = datetime(2025, 10, 14, 0, 0)
END_DATUM = datetime(2025, 10, 18, 23, 59)



class MailJSON:
	def __init__(self, mail_id, mailadresse, betreff, mailtext):
		self.Mail_ID = mail_id
		self.Fahrernummer = ""
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

def lade_mails_aus_outlook():
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
	# Inbox suchen (case-insensitive)
	inbox = finde_ordner(konto, "Inbox")
	if inbox:
		for item in inbox.Items:
			if hasattr(item, "ReceivedTime"):
				empfangen = item.ReceivedTime
				# Offset-naive Vergleich
				if empfangen.tzinfo is not None:
					empfangen = empfangen.replace(tzinfo=None)
				if START_DATUM <= empfangen <= END_DATUM:
					mail_id = len(mails) + 1
					mailadresse = item.SenderEmailAddress
					betreff = item.Subject
					mailtext = item.Body
					mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext))
	else:
		print("Ordner 'Inbox' nicht gefunden!")

	# Archive > Dienstwünsche / FZ suchen (rekursiv)
	dienstwuensche_fz = None
	archive = finde_ordner(konto, "Archive")
	if archive:
		dienstwuensche_fz = finde_ordner(archive, "Dienstwünsche / FZ")
	if dienstwuensche_fz:
		for item in dienstwuensche_fz.Items:
			if hasattr(item, "ReceivedTime"):
				empfangen = item.ReceivedTime
				# Offset-naive Vergleich
				if empfangen.tzinfo is not None:
					empfangen = empfangen.replace(tzinfo=None)
				if START_DATUM <= empfangen <= END_DATUM:
					mail_id = len(mails) + 1
					mailadresse = item.SenderEmailAddress
					betreff = item.Subject
					mailtext = item.Body
					mails.append(MailJSON(mail_id, mailadresse, betreff, mailtext))
	else:
		print("Ordner 'Archive > Dienstwünsche / FZ' nicht gefunden!")

	print(f"{len(mails)} Mails im Zeitraum gefunden.")
	return mails





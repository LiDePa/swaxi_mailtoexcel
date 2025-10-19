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
# Outlook-Mails aus freigegebenem Postfach und Unterordner einlesen
import win32com.client
from datetime import datetime

# Beispielhafte Zeitpunkte (können später durch GUI ersetzt werden)
START_DATUM = datetime(2025, 10, 14, 0, 0)
END_DATUM = datetime(2025, 10, 18, 23, 59)

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
# Mitarbeiterdaten laden und fehlende Kurznamen ergänzen
from openpyxl import load_workbook

def lade_mitarbeiter():
	MITARBEITER_DATEI = "Mitarbeiter.xlsx"
	MITARBEITER_PATH = get_resource_path(MITARBEITER_DATEI)
	wb = load_workbook(MITARBEITER_PATH)
	ws = wb.active
	mitarbeiter_liste = []
	for row in ws.iter_rows(min_row=2):  # Erste Zeile ist Header
		# Stoppe, wenn die Zeile komplett leer ist
		if all(cell.value is None for cell in row):
			break
		fahrernummer = row[0].value  # Spalte A: "Titel"
		vorname = row[1].value      # Spalte B: "Vorname"
		nachname = row[2].value     # Spalte C: "Nachname"
		mailadressen = row[5].value # Spalte F: "E-Mail Privat"
		kurzname = row[10].value    # Spalte K: "Kurzname"
		# Validierung
		if not fahrernummer or (not vorname and not nachname):
			print(f"Fehlerhafte Zeile: Fahrernummer oder Name fehlt. Ignoriert: {vorname} {nachname}")
			continue
		# Kurzname generieren falls leer
		if not kurzname:
			kurzname = generiere_kurzname(vorname, nachname)
			row[10].value = kurzname
		mitarbeiter = Mitarbeiter(fahrernummer, vorname, nachname, mailadressen, kurzname)
		mitarbeiter_liste.append(mitarbeiter)
	wb.save(MITARBEITER_PATH)
	return mitarbeiter_liste

# Kurzname generieren
def generiere_kurzname(vorname, nachname):
	vornamen = str(vorname).split()
	nachnamen = str(nachname).split()
	initialen = [f"{v[0]}." for v in vornamen if v]
	kurzname = " ".join(initialen + nachnamen)
	return kurzname.strip()

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
# Modul für Mail-Handling und Mail_JSON
import os
import sys

# Hilfsfunktion für korrekten Pfad (auch für PyInstaller)
def get_resource_path(filename):
	if hasattr(sys, '_MEIPASS'):
		return os.path.join(sys._MEIPASS, filename)
	return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

class Mitarbeiter:
	def __init__(self, fahrernummer, vorname, nachname, mailadressen, kurzname):
		self.fahrernummer = fahrernummer
		self.vorname = vorname
		self.nachname = nachname
		self.mailadressen = mailadressen
		self.kurzname = kurzname
	def __repr__(self):
		return f"<Mitarbeiter {self.kurzname} ({self.fahrernummer})>"
# Modul für Mail-Handling und Mail_JSON

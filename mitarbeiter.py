from openpyxl import load_workbook
import os
import sys



class Mitarbeiter:
	def __init__(self, fahrernummer, vorname, nachname, mailadressen, kurzname):
		self.fahrernummer = fahrernummer
		self.vorname = vorname
		self.nachname = nachname
		self.mailadressen = mailadressen
		self.kurzname = kurzname
	def __repr__(self):
		return f"<Mitarbeiter {self.kurzname} ({self.fahrernummer})>"
	
# Hilfsfunktion für korrekten Pfad (auch für PyInstaller)
def get_resource_path(filename):
	if hasattr(sys, '_MEIPASS'):
		return os.path.join(sys._MEIPASS, filename)
	return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

# Mitarbeiterdaten laden und fehlende Kurznamen ergänzen
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
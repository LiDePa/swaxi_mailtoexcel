

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
try:
	from tkcalendar import DateEntry
except ImportError:
	DateEntry = None  # Fallback, falls tkcalendar nicht installiert ist

class ZeitraumDialog:

	def __init__(self, master):
		self.master = master
		master.title("Zeitraum auswählen")


		now = datetime.now()
		now_str = now.strftime("%Y-%m-%d %H:%M")

		# Startdatum: Kalender und Uhr
		tk.Label(master, text="Startdatum").grid(row=0, column=0, sticky="w")
		if DateEntry:
			self.start_date = DateEntry(master, date_pattern="yyyy-mm-dd")
		else:
			self.start_date = tk.Entry(master)
			self.start_date.insert(0, now.strftime("%Y-%m-%d"))
		self.start_date.grid(row=0, column=1)

		# Stunde und Minute Spinboxen (größer anzeigen)
		spin_font = ("Arial", 14)  # größere Schrift für bessere Sichtbarkeit
		tk.Label(master, text="Stunde").grid(row=0, column=2)
		self.start_hour = tk.Spinbox(master, from_=0, to=23, width=3, format="%02.0f", font=spin_font)
		self.start_hour.grid(row=0, column=3)
		self.start_hour.delete(0, tk.END)
		self.start_hour.insert(0, now.strftime("%H"))

		tk.Label(master, text=":").grid(row=0, column=4)
		tk.Label(master, text="Minute").grid(row=0, column=5)
		self.start_minute = tk.Spinbox(master, from_=0, to=59, width=3, format="%02.0f", font=spin_font)
		self.start_minute.grid(row=0, column=6)
		self.start_minute.delete(0, tk.END)
		self.start_minute.insert(0, "00")


		# Enddatum: Kalender und Uhr
		tk.Label(master, text="Enddatum").grid(row=1, column=0, sticky="w")
		if DateEntry:
			self.end_date = DateEntry(master, date_pattern="yyyy-mm-dd")
		else:
			self.end_date = tk.Entry(master)
			self.end_date.insert(0, now.strftime("%Y-%m-%d"))
		self.end_date.grid(row=1, column=1)

		tk.Label(master, text="Stunde").grid(row=1, column=2)
		self.end_hour = tk.Spinbox(master, from_=0, to=23, width=3, format="%02.0f", font=spin_font)
		self.end_hour.grid(row=1, column=3)
		self.end_hour.delete(0, tk.END)
		self.end_hour.insert(0, now.strftime("%H"))

		tk.Label(master, text=":").grid(row=1, column=4)
		tk.Label(master, text="Minute").grid(row=1, column=5)
		self.end_minute = tk.Spinbox(master, from_=0, to=59, width=3, format="%02.0f", font=spin_font)
		self.end_minute.grid(row=1, column=6)
		self.end_minute.delete(0, tk.END)
		self.end_minute.insert(0, now.strftime("%M"))

		self.button = tk.Button(master, text="Weiter", command=self.on_submit)
		self.button.grid(row=2, column=0, columnspan=7)

		self.result = None

		# Setze Fokus auf Startdatum (Datum oder Stunde)
		if DateEntry:
			self.start_date.focus_set()
		else:
			self.start_date.focus_set()

	def on_submit(self):
		try:
			# Startdatum zusammensetzen
			if DateEntry:
				start_date_str = self.start_date.get_date().strftime("%Y-%m-%d")
				end_date_str = self.end_date.get_date().strftime("%Y-%m-%d")
			else:
				start_date_str = self.start_date.get()
				end_date_str = self.end_date.get()
			start_hour = self.start_hour.get()
			start_minute = self.start_minute.get()
			end_hour = self.end_hour.get()
			end_minute = self.end_minute.get()
			start_str = f"{start_date_str} {start_hour}:{start_minute}"
			end_str = f"{end_date_str} {end_hour}:{end_minute}"
			start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
			end = datetime.strptime(end_str, "%Y-%m-%d %H:%M")
			self.result = (start, end)
			self.master.quit()
		except Exception as e:
			messagebox.showerror("Fehler", f"Ungültiges Datum: {e}")

def zeitraum_abfragen():
	root = tk.Tk()
	dialog = ZeitraumDialog(root)
	root.mainloop()
	return dialog.result

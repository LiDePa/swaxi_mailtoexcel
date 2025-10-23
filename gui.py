

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

		spin_font = ("Arial", 14)  # gleiche Schrift wie Stunde/Minute

		# Stunde und Minute Spinboxen (größer anzeigen)
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

		# Kalenderwoche (KW) Label und Spinbox (editierbar, gleiche Größe wie Stunde/Minute)
		tk.Label(master, text="Kalenderwoche").grid(row=0, column=7, sticky="w", padx=(10,0))
		self.kw_var = tk.StringVar()
		# Initial KW berechnen: nächste KW
		kw = now.isocalendar()[1] + 1
		if kw > 52:
			kw = 1
		self.kw_var.set(str(kw))
		self.kw_spinbox = tk.Spinbox(master, from_=1, to=53, width=3, format="%02.0f", font=spin_font, textvariable=self.kw_var, wrap=True)
		self.kw_spinbox.grid(row=0, column=8)

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

		# Workaround für tkcalendar.DateEntry Overlay-Bug:
		if DateEntry:
			def _init_dateentry_overlay():
				# Enddatum-DateEntry einmal kurz öffnen und schließen
				self.end_date.event_generate('<Button-1>')
				self.master.after(100, lambda: self.master.focus())
			self.master.after(200, _init_dateentry_overlay)



		# Textfeld für verfügbare Schichten (800 Pixel Standardhöhe, resizable)
		tk.Label(master, text="Verügbare Schichten (aus Mail kopieren):").grid(row=3, column=0, columnspan=9, sticky="w", pady=(10,0))
		self.schichten_text = tk.Text(master, height=1, width=80, wrap="word")
		self.schichten_text.grid(row=4, column=0, columnspan=9, sticky="nsew")
		# Setze die Höhe auf 800 Pixel
		self.schichten_text.update_idletasks()
		self.schichten_text.config(height=1)
		self.schichten_text.config(height=int(800/self.schichten_text.winfo_fpixels('1c')))
		# Fenstergröße anpassen und resizable machen
		master.minsize(800, 600)
		master.resizable(True, True)
		master.grid_rowconfigure(4, weight=1)
		master.grid_columnconfigure(8, weight=1)

		self.button = tk.Button(master, text="Weiter", command=self.on_submit)
		self.button.grid(row=5, column=0, columnspan=9, pady=(10,0))

		self.result = None

		# Setze Fokus auf Startdatum (Datum oder Stunde)
		self.start_date.focus_set()

	def on_submit(self):
		try:
			# Startdatum zusammensetzen
			if DateEntry:
				start_date_obj = self.start_date.get_date()
				start_date_str = start_date_obj.strftime("%Y-%m-%d")
				end_date_str = self.end_date.get_date().strftime("%Y-%m-%d")
			else:
				start_date_str = self.start_date.get()
				start_date_obj = datetime.strptime(start_date_str, "%Y-%m-%d")
				end_date_str = self.end_date.get()
			start_hour = self.start_hour.get()
			start_minute = self.start_minute.get()
			end_hour = self.end_hour.get()
			end_minute = self.end_minute.get()
			start_str = f"{start_date_str} {start_hour}:{start_minute}"
			end_str = f"{end_date_str} {end_hour}:{end_minute}"
			start = datetime.strptime(start_str, "%Y-%m-%d %H:%M")
			end = datetime.strptime(end_str, "%Y-%m-%d %H:%M")
			schichten = self.schichten_text.get("1.0", "end-1c")
			# KW aus Spinbox übernehmen
			try:
				kw = int(self.kw_var.get())
			except Exception:
				kw = start_date_obj.isocalendar()[1]
			self.result = (start, end, schichten, kw)
			self.master.quit()
		except Exception as e:
			messagebox.showerror("Fehler", f"Ungültiges Datum: {e}")

def zeitraum_abfragen():
	root = tk.Tk()
	dialog = ZeitraumDialog(root)
	root.mainloop()
	return dialog.result

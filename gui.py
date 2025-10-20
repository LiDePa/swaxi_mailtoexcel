
import tkinter as tk
from tkinter import ttk
from datetime import datetime

class ZeitraumDialog:
	def __init__(self, master):
		self.master = master
		master.title("Zeitraum auswählen")

		tk.Label(master, text="Startdatum (YYYY-MM-DD)").grid(row=0, column=0)
		self.start_entry = tk.Entry(master)
		self.start_entry.grid(row=0, column=1)

		tk.Label(master, text="Enddatum (YYYY-MM-DD)").grid(row=1, column=0)
		self.end_entry = tk.Entry(master)
		self.end_entry.grid(row=1, column=1)

		self.button = tk.Button(master, text="Weiter", command=self.on_submit)
		self.button.grid(row=2, column=0, columnspan=2)

		self.result = None


	def on_submit(self):
		try:
			start = datetime.strptime(self.start_entry.get(), "%Y-%m-%d")
			start = start.replace(hour=0, minute=0)
			end = datetime.strptime(self.end_entry.get(), "%Y-%m-%d")
			end = end.replace(hour=23, minute=59)
			self.result = (start, end)
			self.master.quit()
		except Exception as e:
			tk.messagebox.showerror("Fehler", f"Ungültiges Datum: {e}")

def zeitraum_abfragen():
	root = tk.Tk()
	dialog = ZeitraumDialog(root)
	root.mainloop()
	return dialog.result

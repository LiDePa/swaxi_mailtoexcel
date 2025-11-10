import asyncio

from mail import lade_mails_aus_outlook, kuerze_mailtexte, weise_kurznamen_zu
from mitarbeiter import lade_mitarbeiter, anonymisiere_mailtexte
from gui import stop_gui_server, zeige_mail_uebersicht, zeitraum_abfragen
from llm import build_instructions, process_mails_parallel


if __name__ == "__main__":
    mitarbeiter = lade_mitarbeiter()
    try:
        zeitraum = zeitraum_abfragen()
        if not zeitraum:
            print("Kein Zeitraum auswaehlt. Programm beendet.")
            exit()
        # Rückgabe aus index.html einlesen
        start, end, available_shifts, kalenderwoche = zeitraum

        # Mails einlesen und aufbereiten
        mails = lade_mails_aus_outlook(start, end)
        mails = kuerze_mailtexte(mails)
        mails = weise_kurznamen_zu(mails, mitarbeiter)

        # review.html
        zeige_mail_uebersicht(mails)

        # Mails anonymisieren, prompt erstellen und alles an LLM senden
        mails = anonymisiere_mailtexte(mails, mitarbeiter)
        instructions = build_instructions(available_shifts)
        if mails:
            asyncio.run(process_mails_parallel(instructions, mails))

        print(available_shifts)
        for mail in mails:
            print(f"mailtext anonymisiert:\n{mail.mailtext_anonym}\n")
            print(f"Schichtwunsch erkannt: {mail.ist_schichtwunsch}")
            print(f"schichtwuensche: {mail.schichtwuensche}")
            print(f"kommentar: {mail.kommentar}\n{'-'*40}\n")
    finally:
        stop_gui_server()


        





        
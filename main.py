

from mail import lade_mails_aus_outlook, weise_kurznamen_zu
from mitarbeiter import lade_mitarbeiter, anonymisiere_mailtexte
from gui import zeitraum_abfragen


if __name__ == "__main__":
    mitarbeiter = lade_mitarbeiter()
    zeitraum = zeitraum_abfragen()
    if zeitraum:
        # Rückgabe aus GUI 
        start, end, schichten_text, kalenderwoche = zeitraum

        # Mails einlesen und aufbereiten
        mails = lade_mails_aus_outlook(start, end)
        mails = anonymisiere_mailtexte(mails, mitarbeiter)
        mails = weise_kurznamen_zu(mails, mitarbeiter)

        # Prompt erstellen und mails an LLM übergeben
        


        print(f"Verfügbare Schichten:\n{schichten_text}\n")
        for mail in mails:
            print(mail.Mailtext_anonym)


from mail import lade_mails_aus_outlook, weise_kurznamen_zu
from mitarbeiter import lade_mitarbeiter, anonymisiere_mailtexte
from gui import zeitraum_abfragen


if __name__ == "__main__":
    mitarbeiter = lade_mitarbeiter()
    zeitraum = zeitraum_abfragen()
    if zeitraum:
        # Neue Rückgabe: (start, end, schichten_text, kw)
        start, end, schichten_text, kalenderwoche = zeitraum
        # Übergib start und end an lade_mails_aus_outlook
        mails = lade_mails_aus_outlook(start, end)
        mails = weise_kurznamen_zu(mails, mitarbeiter)
        mails = anonymisiere_mailtexte(mails, mitarbeiter)
        print(f"Verfügbare Schichten:\n{schichten_text}\n")
        for mail in mails:
            print(mail.Mailtext_anonym)

from mail import lade_mails_aus_outlook, START_DATUM, END_DATUM
from mitarbeiter import lade_mitarbeiter, anonymisiere_mailtexte
from gui import zeitraum_abfragen
import mail


if __name__ == "__main__":
    mitarbeiter = lade_mitarbeiter()
    zeitraum = zeitraum_abfragen()
    if zeitraum:
        mail.START_DATUM, mail.END_DATUM = zeitraum
        mails = lade_mails_aus_outlook()
        mails = anonymisiere_mailtexte(mails, mitarbeiter)
        print(f"Mails: {mails}")
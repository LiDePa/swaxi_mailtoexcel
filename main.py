from mail import lade_mails_aus_outlook
from mitarbeiter import lade_mitarbeiter, anonymisiere_mailtexte

if __name__ == "__main__":
    mails = lade_mails_aus_outlook()
    mitarbeiter = lade_mitarbeiter()
    mails = anonymisiere_mailtexte(mails, mitarbeiter)
    print(f"Mails: {mails}")
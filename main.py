
from mail import lade_mitarbeiter, lade_mails_aus_outlook

if __name__ == "__main__":
    mails = lade_mails_aus_outlook()
    print(f"Geladene Mails: {mails}")
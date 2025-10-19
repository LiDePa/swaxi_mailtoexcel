
from mail import lade_mitarbeiter, debug_mail_json

if __name__ == "__main__":
    mitarbeiter = lade_mitarbeiter()
    print(f"Geladene Mitarbeiter: {mitarbeiter}")
    debug_mail_json()
# Modul für OpenAI-Anbindung

import os
from dotenv import load_dotenv
from openai import AsyncOpenAI
from pydantic import BaseModel
import asyncio

# .env laden
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# AsyncOpenAI-Client initialisieren
client = AsyncOpenAI(api_key=OPENAI_API_KEY)


class MailExtraction(BaseModel):
    ist_schichtwunsch: bool
    schichtwuensche: dict[str, list[str]]  # e.g., {"Montag": ["D20", "D30"], "Dienstag": ["D50"]}
    kommentar: str


#TODO: schichtwünsche nach tagen unterscheiden
#TODO: schichten in klammern wenn entweder an einem oder anderen tag möglich
#TODO: beides in prompts erklären

examples = """
<beispiel>
  <input>Verfügbare Schichten: Mo Früh, Di Spät. Mailtext: 'Hallo, ich hätte gerne Mo Früh und optional Di Spät. Danke!'</input>
  <output>{
  "ist_schichtwunsch": true,
  "schichtwuensche": {
    "Montag": ["Frühschicht", "Mittagsschicht"],
    "Dienstag": ["Spätschicht"],
    "Mittwoch": [],
	"Donnerstag": [],
	"Freitag": [],
	"Samstag": [],
	"Sonntag": []
  },
  "kommentar": "Keine besonderen Wünsche für Mittwoch."
}</output>
</beispiel>
<beispiel>
  <input>Verfügbare Schichten: Mi Nacht, Do Früh. Mailtext: 'Kann diese Woche leider nicht arbeiten.'</input>
  <output>{
  "ist_schichtwunsch": true,
  "schichtwuensche": {
    "Montag": [],
    "Dienstag": [],
    "Mittwoch": [],
	"Donnerstag": [],
	"Freitag": [],
	"Samstag": [],
	"Sonntag": []
  },
  "kommentar": "Keine besonderen Wünsche für Mittwoch."
}</output>
</beispiel>
<beispiel>
  <input>Verfügbare Schichten: Fr Spät, Sa Früh. Mailtext: 'Bitte tragt mich für Fr Spät ein, Samstag kann ich nur, wenn es dringend ist.'</input>
  <output>{
  "ist_schichtwunsch": true,
  "schichtwuensche": {
    "Montag": [],
    "Dienstag": [],
    "Mittwoch": [],
    "Donnerstag": [],
    "Freitag": [],
    "Samstag": [],
    "Sonntag": []
  },
  "kommentar": "Keine besonderen Wünsche für Mittwoch."
}</output>
</beispiel>
"""

def build_instructions(available_shifts, examples=examples):
	
	return f"""
Du bist Teil eines automatisierten KI-Workflows. In einer Angebots-Mail hat ein Schichtleiter seinen Mitarbeitern verfügbare Schichten für die kommende Woche angeboten und du extrahierst nun alle Schichtwünsche aus deren Antworten. Dazu erhältst du die vom Schichtleiter angebotene Liste der verfügbaren Schichten, sowie den anonymisierten Text einer Mail die nach Absenden der Angebots-Mail im Postfach des Schichtleiters eingegangen ist.

Verfügbare Schichten:
<verfügbare_schichten>{available_shifts}</verfügbare_schichten>

Anweisungen zur Extraktion:

* Analysiere den Mailtext und bestimme zunächst, ob es sich dabei überhaupt um eine Antwort auf die Angebots-Mail des Schichtleiters handelt. Falls nicht, setze ist_schichtwunsch auf False und lasse die Felder 'schichtwuensche' und 'kommentar' leer.
* Falls der Mailtext wiederum Informationen zur Verfügbarkeit des Mitarbeiters oder eben konkrete Schichtwünsche beinhaltet, setze ist_schichtwunsch=True, und entnehme dem Mailtext sowohl die vom Mitarbeiter gewünschten Schichten für das Feld 'schichtwuensche', als auch relevante Kommentare zur Verfügbarkeit des Mitarbeiters, für das Feld 'kommentar'. Wenn keine Schichtwünsche oder keine relevanten Kommentare vorhanden sind, lass das entsprechende Feld leer. Sortiere die Schichtwünsche im Feld 'schichtwuensche' nach Wochentagen.
* Oft wünschen sich die Mitarbeiter nicht alle gelisteten Schichten zusammen, sondern entweder bestimmte Schichten an einem bestimmten Tag oder andere Schichten an einem anderen Tag (logische XOR-Beziehung). Überprüfe explizit, ob im Mailtext solche XOR-Schichten formuliert sind, und setze diese im Feld 'schichtwuensche' in Klammern. Beachte dabei: Es kann immer maximal eine Schicht pro Tag vom Mitarbeiter übernommen werden, somit stellen entweder-oder Formulierungen innerhalb des selben Tages keine XOR-Schichten dar.

Beispiele zur Veranschaulichung:
<beispiele>{examples}</beispiele>
"""

EXAMPLE_BLOCKS = [
	{
		"input": "Verfügbare Schichten: Mo Früh, Di Spät. Mailtext: 'Hallo, ich hätte gerne Mo Früh und optional Di Spät. Danke!'",
		"output": "ist_schichtwunsch: True; schichtwuensche: 'Mo Früh, optional Di Spät'; kommentar: 'Keine weiteren Angaben'",
	},
	{
		"input": "Verfügbare Schichten: Mi Nacht, Do Früh. Mailtext: 'Kann diese Woche leider nicht arbeiten.'",
		"output": "ist_schichtwunsch: False; schichtwuensche: ''; kommentar: ''",
	},
	{
		"input": "Verfügbare Schichten: Fr Spät, Sa Früh. Mailtext: 'Bitte tragt mich für Fr Spät ein, Samstag kann ich nur, wenn es dringend ist.'",
		"output": "ist_schichtwunsch: True; schichtwuensche: 'Fr Spät, Sa Früh bei Bedarf'; kommentar: 'Samstag nur im Notfall'",
	},
]

# Asynchrone Funktion für LLM-Request
async def process_mail_anonym(mail, available_shifts):
	instructions = build_instructions(available_shifts)
	prompt = mail.Mailtext_anonym
	try:
		response = await client.responses.parse(
			model="gpt-4o-2024-08-06",
			input=[
				{"role": "developer", "content": instructions},
				{"role": "user", "content": prompt},
			],
			text_format=MailExtraction,
		)
		parsed = getattr(response, "parsed", None)
		if isinstance(parsed, list):
			parsed = parsed[0] if parsed else None
		data = parsed.model_dump() if parsed else {"ist_schichtwunsch": False, "schichtwuensche": "", "kommentar": ""}
	except Exception:
		data = {"ist_schichtwunsch": False, "schichtwuensche": "", "kommentar": ""}
	return {"mail_id": mail.Mail_ID, **data}

# Parallele Verarbeitung einer Liste von MailJSON-Objekten
async def process_mails_parallel(mails, available_shifts=""):
	tasks = [process_mail_anonym(mail, available_shifts=available_shifts) for mail in mails]
	return await asyncio.gather(*tasks)

# Beispiel für die Nutzung:
# from mail import MailJSON
# mails = [MailJSON(...), ...]
# results = asyncio.run(process_mails_parallel(mails))
# for result in results:
#     # result['mail_id'], result['schichtwuensche'], result['kommentar']
#     # -> Dem passenden MailJSON-Objekt zuweisen

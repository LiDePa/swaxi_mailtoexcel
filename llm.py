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


class output_JSON(BaseModel):
    ist_schichtwunsch: bool
    schichtwuensche: dict[str, list[str]]  # e.g., {"Montag": ["D20", "D30"], "Dienstag": ["D50"]}
    kommentar: str


examples = """
<example id="example-1">
<verfügbare_Schichten>
Mittwoch 08.10.

D. 41

D. 60



Donnerstag

D. 41

D. 60



Freitag

D. 55

D. 43



Samstag

D. 20

D. 90

D. 91



Dienstag 14.10.

D. 40

D. 33
</verfügbare_Schichten>
<input>
Anon Anon fahrernummer 6397
Mittwoch 08.10
D.41 oder D60
Und
Donnerstag 09.10
D.41
Oder
Freitag 10.10
D.43
Oder
Samstag 11.10
D.90
Oder
Dienstag 14.10
D.33
Mit freundlichen Grüßen
Anon Anon
---
Anon Anon fahrernummer 6397
korrigiere, am Dienstag kann ich doch nicht.
Mit freundlichen Grüßen
Anon Anon
</input>
<output>{
"ist_schichtwunsch": True,
  "schichtwuensche": {
    "Mittwoch 8.10.": ["D.41", "D.60"],
    "Donnerstag": ["(D.41)"],
    "Freitag": ["(D.43)"],
    "Samstag": ["(D.90)"],
    "Sonntag": [],
    "Dienstag 14.10.": []
  },
  "kommentar": ""}
</output>
</example>

<example id="example-2">
<verfügbare_Schichten>
Sonntag

D. 41



Montag 13.10.

D. 40
</verfügbare_Schichten>
<input>
Hallo, 
ich muss leider am 22.10. meine Schicht stornieren, da ich von meinem Hauptjob kurzfristig wegfahren muss. Entschuldigung.
Fahrernummer 6325
Gruß Anon Anon
</input>
<output>
{"ist_schichtwunsch": False, "schichtwuensche": {}, "kommentar": ""}
</output>
</example>

<example id="example-3">
<verfügbare_Schichten>
Montag 20.10.

D.13

D.40

D.60

D.33

 

Dienstag 21.10.

D.13

D.41

D.60

D.33
</verfügbare_Schichten>
<input>
Hallo Anon
Ich würde gerne kW 42 entweder 20.10 oder 21.10 D.13(17.45-21-45) gerne übernehmen wenn es möglich ist beide..Danke..
LG
Anon
</input>
<output>{
"ist_schichtwunsch": True,
  "schichtwuensche": {
    "Mittwoch": [],
    "Donnerstag": [],
    "Freitag": [],
    "Samstag": [],
    "Sonntag": [],
    "Montag": [D.13],
    "Dienstag": [D.21]
  },
  "kommentar": ""}
</output>
</example>

<example id="example-4">
<verfügbare_Schichten>
Donnerstag

D. 55

D. 43

 

Samstag

D. 55

D. 43
</verfügbare_Schichten>
<input>
Servus,
ich könnte am Samstag eine Schicht übernehmen, falls noch welche verfügbar sind.
Mit freundlichen Grüßen
</input>
<output>{
"ist_schichtwunsch": True,
  "schichtwuensche": {
    "Mittwoch": [],
    "Donnerstag": [],
    "Freitag": [],
    "Samstag": ["EGAL"],
    "Sonntag": [],
    "Montag": [],
    "Dienstag": []
  },
  "kommentar": ""}
</output>
</example>
"""

def build_instructions(available_shifts, examples=examples):
	
	return f"""
Du bist Teil eines automatisierten KI-Workflows und erstellst eine "output_JSON". In einer Angebots-Mail hat ein Schichtleiter seinen Mitarbeitern eine Liste "verfügbare_Schichten" für die kommende Woche angeboten. Du extrahierst nun alle Schichtwünsche aus deren Antworten und schreibst sie in output_JSON. Dazu erhältst du die vom Schichtleiter angebotene Liste der verfügbaren Schichten, sowie den anonymisierten Text einer Mail die nach Absenden der Angebots-Mail im Postfach des Schichtleiters eingegangen ist. Sollte der selbe Mitarbeiter mehrere Mails geschickt haben (z.B. um sich zu korrigieren), erhältst du trotzdem einen einzelnen Mailtext der alle seine Antworten zusammenfasst.

# Anweisungen zur Extraktion

* Analysiere den Mailtext und bestimme zunächst, ob es sich dabei überhaupt um eine Antwort auf die Angebots-Mail des Schichtleiters handelt. Falls nicht, setze ist_schichtwunsch auf False und lasse die Felder 'schichtwuensche' und 'kommentar' leer. Du kannst in diesem Fall alle anderen Schritte überspringen. Falls du unsicher bist, ob es sich um einen Schichtwunsch handelt, entscheide dich im Zweifelsfall für ist_schichtwunsch=True.
* Falls der Mailtext wiederum Schichtwünsche beinhält, setze ist_schichtwunsch=True, und entnehme dem Mailtext sowohl die vom Mitarbeiter gewünschten Schichten für das Feld 'schichtwuensche', als auch wichtige Kommentare zur Verfügbarkeit des Mitarbeiters, für das Feld 'kommentar' in output_JSON. Beachte dabei folgendes:
    * Erstelle im Feld 'schitchtwuensche' als dictionary immer eine vollständige Liste der Wochentage, in der Reihenfolge, wie sie in verfügbare_Schichten genannt werden. Sollte in verfügbare_Schichten einer der sieben Wochentage nicht genannt werden, oder einen offensichtlichen Schreibfehler haben, dann füge ihn in output_JSON an der passenden Stelle selbst ein, sodass das dictionary in 'schichtwünsche' immer aus genau 7 Schlüsseln von aufeinander folgenden Wochentagen besteht. Meistens beginnt die Liste mit Mittwoch und endet mit Dienstag.
    * Oft wünschen sich die Mitarbeiter nicht alle gelisteten Schichten zusammen, sondern entweder bestimmte Schichten an einem bestimmten Tag oder andere Schichten an einem anderen Tag (logische XOR-Beziehung). Überprüfe explizit, ob im Mailtext solche XOR-Schichten formuliert sind, und setze diese im Feld 'schichtwuensche' in Klammern. Beachte dabei: Es kann immer maximal eine Schicht pro Tag vom Mitarbeiter übernommen werden, somit stellen entweder-oder Formulierungen innerhalb des selben Tages keine XOR-Schichten dar.
    * Sollte sich ein Mitarbeiter in seinem Mailtext für alle möglichen Schichten eines Tages bereit erklären, so trage am entsprechenden Tag nicht die Bezeichnungen der Schichten sondern "EGAL" ein. Sollte er sich für alle Schichten aller Tage bereit erklären, trage an allen Tagen "EGAL" ein.
* Sollte ist_schichtwunsch=True sein, aber du kannst keine konkreten Schichtwünsche oder Kommentare extrahieren, oder sollte sich der Mitarbeiter auch nur eine einzige Schicht wünschen, die nicht in der vom Schichtleiter angebotenen Liste verfügbare_Schichten enthalten ist, dann erstelle kein dictionary und lasse die Felder 'schichtwuensche' und 'kommentar' leer. Tu das selbe wann immer du dir bei etwas unsicher bist.

# verfügbare_Schichten

<verfügbare_Schichten>{available_shifts}</verfügbare_Schichten>

# Beispiele zur Veranschaulichung

<beispiele>{examples}</beispiele>
"""

# Asynchrone Funktion für LLM-Request
async def process_mail_anonym(instructions, prompt):
	try:
		response = await client.responses.parse(
			model="gpt-5-2025-08-07",
			input=[
				{"role": "developer", "content": instructions},
				{"role": "user", "content": prompt},
			],
			text_format=output_JSON,
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

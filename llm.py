# Modul für OpenAI-Anbindung

import os
import asyncio
import copy
from typing import Any, Dict, Iterable, List, Mapping

from dotenv import load_dotenv
from openai import AsyncOpenAI
from pydantic import BaseModel, Field

# .env laden
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# AsyncOpenAI-Client initialisieren
client = AsyncOpenAI(api_key=OPENAI_API_KEY)


class Schichtwuensche(BaseModel):
    Mittwoch: List[str] = Field(default_factory=list)
    Donnerstag: List[str] = Field(default_factory=list)
    Freitag: List[str] = Field(default_factory=list)
    Samstag: List[str] = Field(default_factory=list)
    Sonntag: List[str] = Field(default_factory=list)
    Montag: List[str] = Field(default_factory=list)
    Dienstag: List[str] = Field(default_factory=list)

class OutputJSON(BaseModel):
  ist_schichtwunsch: bool
  schichtwuensche: Schichtwuensche
  kommentar: str


_DEFAULT_OUTPUT = OutputJSON(
  ist_schichtwunsch=False,
  schichtwuensche=Schichtwuensche(),
  kommentar="",
).model_dump()


def _get_value(mail_item: Any, key: str, default: Any = None) -> Any:
  """Read a value from a dict-like or attribute-based mail representation."""
  if isinstance(mail_item, Mapping):
    return mail_item.get(key, default)
  return getattr(mail_item, key, default)


def build_instructions(available_shifts):
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
  Hinweis: Der Zeitraum ist vorlesungsfreie Zeit i.S.d. Werkstudierendenregelung. Daher sind hier mehr als 20h pro Woche möglich.
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
      "Mittwoch": ["D.41", "D.60"],
      "Donnerstag": ["(D.41)"],
      "Freitag": ["(D.43)"],
      "Samstag": ["(D.90)"],
      "Sonntag": [],
      "Dienstag": []
    },
    "kommentar": "Mehr als 20 h pro Woche möglich, da vorlesungsfreie Zeit."}
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
  {"ist_schichtwunsch": False,
    "schichtwuensche": {
      "Mittwoch": [],
      "Donnerstag": [],
      "Freitag": [],
      "Samstag": [],
      "Sonntag": [],
      "Montag": [],
      "Dienstag": []
    }, 
    "kommentar": ""}
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
      "Montag": ["D.13"],
      "Dienstag": ["D.13"]
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

  <example id="example-5">
  <verfügbare_Schichten>
  Mittwoch 06.08.
  D.33
  Donnerstag 07.08.
  D.14
  D.40
  D.37
  Freitag 08.08.
  D.80
  Samstag 09.08.
  D.33
  Sonntag 10.08.
  D.21
  D.41
  D.39
  Montag 11.08.
  D.14
  D.16
  Dienstag 12.08.
  D.2
  D.14
  D.16
  </verfügbare_Schichten>
  <input>
  Servus Anon, 
  in KW32 (vorlesungsfreie Zeit) würde ich gerne noch folgende Dienste zusätzlich übernehmen: 
  Mittwoch 06.08.: D.33
  Donnerstag 07.08.: D.40 oder auch D.37
  Sonntag: 10.08.: D.41 oder notfalls auch D.39
  Montag 11.08.: D.16 
  Gib mir gerne so viele Dienste wie möglich. 
  Beste Grüße 
  Anon Anon
  </input>
  <output>{
  "ist_schichtwunsch": True,
    "schichtwuensche": {
      "Mittwoch": ["D.33"],
      "Donnerstag": ["D.40", "D.37"],
      "Freitag": [],
      "Samstag": [],
      "Sonntag": ["D.41", "D.39"],
      "Montag": ["D.16"],
      "Dienstag": []
    },
    "kommentar": "D.39 am Sonntag nur notfalls. Wünscht sich so viele Dienste wie möglich."}
  </output>
  </example>"""
	
  return f"""
Du erzeugst genau eine JSON im Schema "OutputJSON" anhand der Liste <verfügbare_Schichten> und eines anonymisierten Mailtexts.

# Kurzüberblick
* Bestimme, ob die Mail ein Schichtwunsch ist. Wenn nein: setze ist_schichtwunsch = False und lasse 'schichtwuensche' sowie 'kommentar' leer.
* Falls unklar: entscheide zugunsten ist_schichtwunsch = True.
* Mehrere zusammengefügte Mails desselben Mitarbeiters werden als ein Text behandelt. Neuere Nachrichten sind weiter unten und überschreiben vorherige Nachrichten.

# Extraktionsregeln
* Normalisiere Schicht-IDs ins Format 'D.41' (kein Leerzeichen).
* Klammern = XOR über verschiedene Tage: Schichten, die an verschiedenen Tagen per „oder“ verbunden sind, sollen als '"(D.41)"' markiert werden (z. B. Mittwoch D.41 oder Donnerstag D.41 => "Mittwoch": ["(D.41)"], "Donnerstag": ["(D.41)"]).
* „oder“ innerhalb desselben Tages: Liste die Schichten ohne Klammern (z. B. Mittwoch: entweder D.41 oder D.60 => ["D.41","D.60"]).
* „und“ trennt unabhängige Wünsche.
* Wenn alle Schichten eines Tages akzeptabel sind dann setze '"EGAL"'.
* 'kommentar': Alle sonstigen für die Vergabe relevanten Angaben (z. B. Kann weniger/mehr Stunden machen als normalerweise, Priorisierungen zwischen gewünschten Schichten). In 'kommentar' lieber zu viel als zu wenig reinschreiben. Du entscheidest, welche Informationen für die Schichtvergabe relevant sind.

# Schwierige Fälle / Sonderregeln
* Wenn ist_schichtwunsch = True, aber sich keine verlässlichen Wünsche extrahieren lassen, dann belasse 'schichtwuensche' und 'kommentar' leer — das kennzeichnet den Fall zur manuellen Prüfung.
* Falls nur eine gewünschte Schicht genannt wird, die nicht in <verfügbare_Schichten> steht, gilt dasselbe (leere Felder).

# verfügbare_Schichten
<verfügbare_Schichten>{available_shifts}</verfügbare_Schichten>

# Beispiele zur Veranschaulichung
<beispiele>{examples}</beispiele>
"""

# Asynchrone Funktion für LLM-Request
async def process_mail_anonym(instructions: str, mail_item: Any) -> Dict[str, Any]:
  """Run a single anonymised mail through the LLM and return the parsed response."""

  prompt = _get_value(mail_item, "Mailtext_anonym", "") or ""
  mail_id = _get_value(mail_item, "Mail_ID") or _get_value(mail_item, "mail_id")

  try:
    response = await client.responses.parse(
      model="gpt-5-2025-08-07",
      input=[
        {"role": "developer", "content": instructions},
        {"role": "user", "content": prompt},
      ],
      text_format=OutputJSON,
    )
    parsed = getattr(response, "parsed", None)
    if isinstance(parsed, list):
      parsed = parsed[0] if parsed else None
    data = parsed.model_dump() if parsed else copy.deepcopy(_DEFAULT_OUTPUT)
  except Exception:
    data = copy.deepcopy(_DEFAULT_OUTPUT)

  schichtwuensche = copy.deepcopy(data.get("schichtwuensche", {}))
  kommentar = data.get("kommentar", "")
  ist_schichtwunsch = data.get("ist_schichtwunsch", False)

  result = {
    "Mail_ID": mail_id,
    "mail_id": mail_id,
    "ist_schichtwunsch": ist_schichtwunsch,
    "ist_Schichtwunsch": ist_schichtwunsch,
    "schichtwuensche": schichtwuensche,
    "Schichtwuensche": schichtwuensche,
    "kommentar": kommentar,
    "Kommentar": kommentar,
  }

  if isinstance(mail_item, dict):
    mail_item.update(result)
  else:
    if mail_id is not None and hasattr(mail_item, "Mail_ID"):
      mail_item.Mail_ID = mail_id
    if hasattr(mail_item, "Schichtwuensche"):
      mail_item.Schichtwuensche = schichtwuensche
    if hasattr(mail_item, "Kommentar"):
      mail_item.Kommentar = kommentar
    if hasattr(mail_item, "ist_Schichtwunsch"):
      mail_item.ist_Schichtwunsch = ist_schichtwunsch
    if hasattr(mail_item, "ist_schichtwunsch"):
      mail_item.ist_schichtwunsch = ist_schichtwunsch

  return {"mail_id": mail_id, **data}

# Parallele Verarbeitung einer Liste von MailJSON-Objekten
async def process_mails_parallel(instructions: str, mails: Iterable[Any]):
  tasks = [process_mail_anonym(instructions, mail) for mail in mails]
  return await asyncio.gather(*tasks)

# Beispiel für die Nutzung:
# from mail import MailJSON
# mails = [MailJSON(...), ...]
# results = asyncio.run(process_mails_parallel(instructions, mails))
# for result in results:
#     # result['mail_id'], result['schichtwuensche'], result['kommentar']
#     # -> Dem passenden MailJSON-Objekt zuweisen

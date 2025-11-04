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
	schichtwuensche: str
	kommentar: str

# Beispiel-Prompt mit Platzhalter

instructions = """
Du bist ein hilfreiches KI-Modell. Analysiere den folgenden anonymisierten E-Mail-Text und extrahiere daraus die Schichtwünsche sowie einen eventuellen Kommentar. Fülle die Felder 'schichtwuensche' und 'kommentar' im strukturierten Ergebnis passend aus.
"""

PROMPT_TEMPLATE = """Mailtext_anonym:\n{mailtext_anonym}\n"""

# Asynchrone Funktion für LLM-Request
async def process_mail_anonym(mail, prompt_template=PROMPT_TEMPLATE):
	prompt = prompt_template.format(mailtext_anonym=mail.Mailtext_anonym)
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
		data = parsed.model_dump() if parsed else {"schichtwuensche": "", "kommentar": ""}
	except Exception:
		data = {"schichtwuensche": "", "kommentar": ""}
	return {"mail_id": mail.Mail_ID, **data}

# Parallele Verarbeitung einer Liste von MailJSON-Objekten
async def process_mails_parallel(mails):
	tasks = [process_mail_anonym(mail) for mail in mails]
	return await asyncio.gather(*tasks)

# Beispiel für die Nutzung:
# from mail import MailJSON
# mails = [MailJSON(...), ...]
# results = asyncio.run(process_mails_parallel(mails))
# for result in results:
#     # result['mail_id'], result['schichtwuensche'], result['kommentar']
#     # -> Dem passenden MailJSON-Objekt zuweisen

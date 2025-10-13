# main.py

import os
from dotenv import load_dotenv
from embedding import update_embeddings

# Load environment variables
load_dotenv("settings.env")
load_dotenv("llm.env")

MAILNAME_COSSIM = float(os.getenv("MAILNAME_COSSIM", 0.6))
PRESIDIO_COSSIM = float(os.getenv("PRESIDIO_COSSIM", 0.6))

if __name__ == "__main__":
    print("Starting the workflow...")
    # Update embeddings based on Mitarbeitendendaten.xlsx
    update_embeddings()
    # GUI and further processing will be added here
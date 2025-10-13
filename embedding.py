# embedding.py

import os
import pickle
import pandas as pd
from sentence_transformers import SentenceTransformer
from presidio_analyzer import AnalyzerEngine, PatternRecognizer
from presidio_anonymizer import AnonymizerEngine

EMBEDDINGS_FILE = "embeddings.pkl"
EMPLOYEE_DATA_FILE = "Mitarbeitendendaten.xlsx"
MODEL_NAME = "all-MiniLM-L6-v2"

# Initialize the sentence transformer model
model = SentenceTransformer(MODEL_NAME)

# Initialize Presidio
analyzer = AnalyzerEngine()
anonymizer = AnonymizerEngine()

def update_embeddings():
    """Update embeddings.pkl based on Mitarbeitendendaten.xlsx."""
    # Load existing embeddings or create a new list
    if os.path.exists(EMBEDDINGS_FILE):
        with open(EMBEDDINGS_FILE, "rb") as f:
            embeddings = pickle.load(f)
    else:
        embeddings = []

    # Load employee data
    employee_data = pd.read_excel(EMPLOYEE_DATA_FILE)
    existing_names = {entry["Klarname"] for entry in embeddings}

    # Update embeddings
    for _, row in employee_data.iterrows():
        klarname = f"{row['Vorname']} {row['Nachname']}"
        if klarname not in existing_names:
            embedding = model.encode(klarname.lower())
            embeddings.append({
                "Klarname": klarname,
                "embedding": embedding,
                "Fahrernummer": row["Fahrernummer"],
                "mailaddresses": []
            })

    # Save updated embeddings
    with open(EMBEDDINGS_FILE, "wb") as f:
        pickle.dump(embeddings, f)

    # Update Presidio recognizers
    klarname_list = [entry["Klarname"] for entry in embeddings]
    recognizer = PatternRecognizer(supported_entity="Klarname", patterns=[
        {"name": "KlarnamePattern", "pattern": name, "score": 1.0} for name in klarname_list
    ])
    analyzer.registry.add_recognizer(recognizer)

    print("Embeddings updated successfully.")
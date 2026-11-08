# LetsMeet Migration - Akt 1, 2 und 3
# Baut die Zieldatenbank aus Excel, MongoDB und XML auf.
# Kann mehrmals laufen ohne Fehler, weil vorher alles gedroppt wird.

import re
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
import pandas as pd
from pymongo import MongoClient
from sqlalchemy import create_engine, text

# Ordner in dem dieses Script liegt (fuer Vorlage-Dateien)
try:
    SCRIPT_DIR = Path(__file__).resolve().parent
except NameError:
    SCRIPT_DIR = Path.cwd()

# Verbindungsdaten - pg8000 weil auf dem Schulserver kein psycopg2 da ist
DB = "postgresql+pg8000://user:secret@127.0.0.1:5432/lf8_lets_meet_db"
MONGO_URI = "mongodb://127.0.0.1:27017/"
MONGO_DB = "LetsMeet"

# Trenner in den zusammengesetzten Spalten: Komma + ein Leerzeichen
TRENNER = ", "
# Spaltenname der Hobby-Sammelspalte in der Excel
HOBBY_COL = "Hobby1 %Prio1%; Hobby2 %Prio2%; Hobby3 %Prio3%; Hobby4 %Prio4%; Hobby5 %Prio5%;"
# Regex holt aus "Hobbyname %42%" den Namen und die Zahl raus
HOBBY_PAT = re.compile(r"([^%;]+?)\s*%(\d+)%\s*;?")
# erlaubter Bereich fuer priority
PRIO_MIN = -100
PRIO_MAX = 100

# ----------------------------------------------------------------------
# Vorlagen fuer das Aenderungspaket (Akt 3).
# Diese drei Dateien liegen nicht immer auf dem Server - das Script legt
# sie bei Bedarf selbst an. Die Encoding-Dateien enthalten absichtlich
# kaputte Bytes (Latin-1 bzw. Mojibake), daher als bytes definiert.
# ----------------------------------------------------------------------
VORLAGE_CHANGE_REQUEST = (
    '<?xml version="1.0" encoding="UTF-8"?>\n'
    '<transferpack>\n'
    '  <records>\n'
    '    <like email="abdel.sabah@1mal1.te" target_email="transfer.orphan@letsmeet.invalid"/>\n'
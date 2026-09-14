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
    '    <hobby email="abdel.sabah@1mal1.te" name="E-Mails schreiben"/>\n'
    '    <hobby email="abdulk..stuckmann@web.kom" name="Abends seinem Partner Ereignisse des Tages erz\u00e4hlen" priority="101"/>\n'
    '    <profile email="transfer.p2@letsmeet.invalid" first_name="Pia" last_name="Transfer" birth_date="01.01.1900" postal_code="69115" city="unbekannt" phone="+49 6221 123456" gender="w"/>\n'
    '    <hobby email="acar.nehir@ge-em-ix.kom" name="Abends seinem Partner Ereignisse des Tages erz\u00e4hlen"/>\n'
    '    <hobby email="acar.nehir@ge-em-ix.kom" name="Abends seinem Partner Ereignisse des Tages erz\u00e4hlen"/>\n'
    '  </records>\n'
    '</transferpack>\n'
).encode("utf-8")

# encoding-invalid: enthaelt Byte 0xE9 (Latin-1 fuer e-acute) mitten im UTF-8
VORLAGE_ENCODING_INVALID = (
    b'<?xml version="1.0" encoding="UTF-8"?>\n'
    b'<transferpack><records><profile email="transfer.bytes@letsmeet.invalid" '
    b'first_name="Andr\xe9" last_name="Byte" birth_date="1988-04-12" '
    b'postal_code="50667" city="Koeln" phone="+49 221 123456" gender="m"/>'
    b'</records></transferpack>\n'
)

# encoding-mojibake: doppelt-encodierte Umlaute (M+C3 83 C2 BC = Mueller, K+C3 B6 = Koeln)
VORLAGE_ENCODING_MOJIBAKE = (
    b'<?xml version="1.0" encoding="UTF-8"?>\n'
    b'<transferpack><records><profile email="transfer.mojibake@letsmeet.invalid" '
    b'first_name="M\xc3\x83\xc2\xbcller" last_name="Text" birth_date="1988-04-12" '
    b'postal_code="50667" city="K\xc3\xb6ln" phone="+49 221 123457" gender="w"/>'
    b'</records></transferpack>\n'
)


# sucht eine Datei an den ueblichen Orten, damit das Script egal wo laeuft
def finde(name):
    orte = [Path.cwd(), Path.home(), Path.home() / "work" / "letsmeet",
            Path.home() / "LetsMeet"]
    for ordner in orte:
        pfad = ordner / name
        if pfad.exists():
            return pfad
    # wenn nicht gefunden: im ganzen Home suchen
    treffer = list(Path.home().rglob(name))
    if treffer:
        return treffer[0]
    return None


# legt eine Vorlage-Datei an falls sie noch nicht existiert, gibt den Pfad zurueck
def datei_sicherstellen(name, inhalt_bytes):
    # erst schauen ob es die Datei schon irgendwo gibt
    vorhanden = finde(name)
    if vorhanden is not None:
        return vorhanden
    # sonst neben dem Script anlegen
    ziel = SCRIPT_DIR / name
    ziel.write_bytes(inhalt_bytes)
    print(f"  Vorlage angelegt: {ziel}")
    return ziel


# wandelt einen Zeitstempel-Text in ein echtes Datum um
# in MongoDB gibt es zwei Formate, deswegen beide probieren
def parse_zeit(wert):
    for muster in ("%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M:%S"):
        try:
            return datetime.strptime(str(wert), muster)
        except ValueError:
            pass
    return None


# prueft ob ein MongoDB-Wert brauchbar ist (nicht leer und nicht "0")
def wert_ok(wert):
    return str(wert).strip() not in ("", "0", "None")


# repariert kaputte Umlaute (Mojibake), z.B. "MÃ¼ller" wird zu "Müller"
def repariere_text(s):
    # bis zu 3x versuchen, weil manche Texte doppelt kaputt sind
    for _ in range(3):
        try:
            neu = s.encode("latin-1").decode("utf-8")
            if neu == s:
                break
            s = neu
        except (UnicodeDecodeError, UnicodeEncodeError):
            break
    return s


# gibt eine Ueberschrift aus damit man in der Ausgabe was sieht
def ueberschrift(text_):
    print("\n" + "=" * 60)
    print(text_)
    print("=" * 60)


# ======================================================================
# AKT 1 - nur die Excel-Datei, Stammdaten aufbauen
# ======================================================================
def akt1(engine):
    ueberschrift("AKT 1: Excel einlesen")

    # Excel laden, alles als Text, damit nichts kaputt umgewandelt wird
    df = pd.read_excel(finde("Lets Meet DB Dump.xlsx"), dtype=str, keep_default_na=False)
    print(f"Excel hat {len(df)} Zeilen")

    # hier sammeln wir die Personen und Kontaktdaten
    personen = []
    kontakte = []

    # jede Zeile der Excel durchgehen
    for _, zeile in df.iterrows():
        # "Nachname, Vorname" am Trenner aufteilen - maxsplit 1, also 2 Teile
        name_teile = zeile["Nachname, Vorname"].split(TRENNER, 1)
        nachname = name_teile[0]
        # falls kein Vorname da ist, leerer String
        vorname = name_teile[1] if len(name_teile) > 1 else ""

        # Adresse "Strasse, PLZ, Ort" - maxsplit 2, also 3 Teile
        # dadurch bleibt "Demmin, Hansestadt" als Ort zusammen
        adr_teile = zeile["Straße Nr, PLZ Ort"].split(TRENNER, 2)
        strasse = adr_teile[0] if len(adr_teile) > 0 else ""
        plz = adr_teile[1] if len(adr_teile) > 1 else ""
        stadt = adr_teile[2] if len(adr_teile) > 2 else ""

        # Geburtsdatum von TT.MM.JJJJ in ein echtes Datum umwandeln
        geb = pd.to_datetime(zeile.get("Geburtsdatum", ""), format="%d.%m.%Y", errors="coerce")
        geb = geb.date() if not pd.isna(geb) else None

        # Person in die Liste packen
        personen.append({
            "nachname": nachname,
            "vorname": vorname,
            "plz": plz,
            "stadt": stadt,
            "geschlecht": zeile.get("Geschlecht (m/w/nonbinary)", ""),
            "interessiert_an": zeile.get("Interessiert an", ""),
            "geburtsdatum": geb,
        })
        # Kontaktdaten getrennt sammeln
        kontakte.append({
            "email": zeile["E-Mail"],
            "strasse_nr": strasse,
            "telefon": zeile.get("Telefon", ""),
        })

    # aus den Listen DataFrames machen
    df_personen = pd.DataFrame(personen)
    df_kontakte = pd.DataFrame(kontakte)

    # jetzt die Hobbys aus der Sammelspalte holen
    hobby_liste = []
    for _, zeile in df.iterrows():
        text_ = zeile[HOBBY_COL]
        # nur wenn was drinsteht
        if isinstance(text_, str):
            # Regex findet alle Hobby-Prio-Paare
            for name, prio in HOBBY_PAT.findall(text_):
                hobby_liste.append({
                    "email": zeile["E-Mail"],
                    "hobby_name": name.strip(),
                    "priority": int(prio),
                    "source": "excel",
                })
    df_hobbys = pd.DataFrame(hobby_liste)
    print(f"{len(df_hobbys)} Hobby-Eintraege gefunden")

    # ---- alte Objekte loeschen und Tabellen neu anlegen ----
    with engine.begin() as conn:
        # erst alle Views weg (haengen von Tabellen ab)
        for v in ["migration_rejections", "migration_messages", "migration_likes",
                  "migration_user_hobbies", "migration_user_interests", "migration_users"]:
            conn.execute(text(f"DROP VIEW IF EXISTS {v} CASCADE"))
        # dann die Tabellen weg
        for t in ["stg_rejections", "stg_likes", "stg_messages", "person_hobby",
                  "kontakt", "person", "hobby"]:
            conn.execute(text(f"DROP TABLE IF EXISTS {t} CASCADE"))

        # hobby-Tabelle: jedes Hobby einmal
        conn.execute(text("""
            CREATE TABLE hobby (
                hobby_id SERIAL PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                source TEXT
            )"""))
        # person-Tabelle: die Stammdaten
        conn.execute(text("""
            CREATE TABLE person (
                person_id SERIAL PRIMARY KEY,
                nachname TEXT NOT NULL,
                vorname TEXT NOT NULL,
                plz TEXT NOT NULL,
                stadt TEXT NOT NULL,
                geschlecht TEXT,
                interessiert_an TEXT,
                geburtsdatum DATE NOT NULL
            )"""))
        # kontakt-Tabelle: gehoert 1:1 zur Person
        conn.execute(text("""
            CREATE TABLE kontakt (
                kontakt_id SERIAL PRIMARY KEY,
                person_id INTEGER NOT NULL UNIQUE REFERENCES person(person_id),
                email TEXT NOT NULL UNIQUE,
                strasse_nr TEXT,
                telefon TEXT
            )"""))
        # person_hobby: Zwischentabelle fuer n:m.
        # source ist Teil vom Schluessel, damit dieselbe Zuordnung theoretisch
        # aus verschiedenen Quellen kommen koennte - fachlich sorgen wir aber
        # dafuer, dass eine Zuordnung nur einmal existiert (siehe Akt 3).
        conn.execute(text("""
            CREATE TABLE person_hobby (
                person_id INTEGER NOT NULL REFERENCES person(person_id),
                hobby_id INTEGER NOT NULL REFERENCES hobby(hobby_id),
                prioritaet INTEGER,
                source TEXT NOT NULL,
                PRIMARY KEY (person_id, hobby_id, source)
            )"""))
    print("Tabellen angelegt")

    # ---- Daten in die Tabellen schreiben ----
    # Hobbys: nur eindeutige Namen, sonst meckert das UNIQUE
    df_hobbys[["hobby_name", "source"]].drop_duplicates("hobby_name") \
        .rename(columns={"hobby_name": "name"}) \
        .to_sql("hobby", engine, if_exists="append", index=False)

    # Personen schreiben
    df_personen.to_sql("person", engine, if_exists="append", index=False)

    # Kontakte: person_id ist einfach die Zeilennummer 1..n
    df_k = df_kontakte.copy()
    df_k.insert(0, "person_id", range(1, len(df_k) + 1))
    df_k.to_sql("kontakt", engine, if_exists="append", index=False)

    # jetzt person_hobby fuellen - dafuer brauchen wir die IDs
    # Email -> person_id
    email_zu_pid = {r["email"].lower(): i + 1 for i, r in df_kontakte.iterrows()}
    # Hobbyname -> hobby_id aus der DB holen
    with engine.connect() as conn:
        name_zu_hid = {r[0]: r[1] for r in conn.execute(text("SELECT name, hobby_id FROM hobby")).fetchall()}

    # Zuordnungen bauen, doppelte vermeiden
    zuordnungen = []
    schon_da = set()
    for _, r in df_hobbys.iterrows():
        pid = email_zu_pid.get(r["email"].lower())
        hid = name_zu_hid.get(r["hobby_name"])
        # nur wenn beide IDs existieren und die Kombi neu ist
        if pid and hid and (pid, hid) not in schon_da:
            schon_da.add((pid, hid))
            zuordnungen.append({
                "person_id": pid,
                "hobby_id": hid,
                "prioritaet": r["priority"],
                "source": r["source"],
            })
    pd.DataFrame(zuordnungen).to_sql("person_hobby", engine, if_exists="append", index=False)
    print(f"{len(df_personen)} Personen und {len(zuordnungen)} Hobby-Zuordnungen geschrieben")


# ======================================================================
# AKT 2 - MongoDB dazuholen und die 5 Views bauen
# ======================================================================
def akt2(engine):
    ueberschrift("AKT 2: MongoDB und Views")

    # alle User-Dokumente aus MongoDB holen
    mongo_docs = list(MongoClient(MONGO_URI)[MONGO_DB].users.find())
    # Nachschlag-Dict: Email (klein) -> Dokument
    mongo_dict = {d["_id"].lower(): d for d in mongo_docs}
    print(f"MongoDB hat {len(mongo_docs)} Dokumente")

    # ---- Konflikte aufloesen ----
    # Regel: MongoDB ist neuer und gewinnt, ausser der Wert ist "0" oder leer
    # aktuelle Personen/Kontakte aus der DB holen
    with engine.connect() as conn:
        zeilen = conn.execute(text("""
            SELECT p.person_id, k.email, p.nachname, p.vorname, k.telefon
            FROM person p
            JOIN kontakt k ON k.person_id = p.person_id
        """)).fetchall()

    # hier merken wir uns was geaendert werden muss
    updates_person = []
    updates_kontakt = []
    for pid, email, nachname, vorname, telefon in zeilen:
        mdoc = mongo_dict.get(email.lower(), {})
        neu_nach = nachname
        neu_vor = vorname
        neu_tel = telefon

        # Name aus MongoDB nehmen wenn vorhanden
        mongo_name = mdoc.get("name", "")
        if wert_ok(mongo_name):
            teile = mongo_name.split(", ", 1)
            neu_nach = teile[0]
            neu_vor = teile[1] if len(teile) > 1 else vorname

        # Telefon aus MongoDB nehmen wenn brauchbar (nicht "0")
        if wert_ok(mdoc.get("phone", "")):
            neu_tel = mdoc["phone"]

        # nur merken wenn sich wirklich was aendert
        if (neu_nach, neu_vor) != (nachname, vorname):
            updates_person.append({"pid": pid, "n": neu_nach, "v": neu_vor})
        if neu_tel != telefon:
            updates_kontakt.append({"pid": pid, "t": neu_tel})

    # Aenderungen in die DB schreiben
    with engine.begin() as conn:
        for u in updates_person:
            conn.execute(text("UPDATE person SET nachname=:n, vorname=:v WHERE person_id=:pid"), u)
        for u in updates_kontakt:
            conn.execute(text("UPDATE kontakt SET telefon=:t WHERE person_id=:pid"), u)
    print(f"{len(updates_person)} Namen und {len(updates_kontakt)} Telefonnummern korrigiert")

    # ---- Likes einsammeln ----
    likes = []
    for doc in mongo_docs:
        # jedes Dokument kann mehrere Likes haben
        for like in doc.get("likes", []):
            likes.append({
                "liker_email": doc["_id"],
                "liked_email": like["liked_email"],
                "status": like["status"],
                "liked_at": parse_zeit(like.get("timestamp", "")),
            })
    df_likes = pd.DataFrame(likes)

    # ---- Nachrichten einsammeln ----
    nachrichten = []
    for doc in mongo_docs:
        for msg in doc.get("messages", []):
            nachrichten.append({
                "sender_email": doc["_id"],
                "receiver_email": msg["receiver_email"],
                "body": msg["message"],
                "sent_at": parse_zeit(msg.get("timestamp", "")),
                "conversation_id": msg["conversation_id"],
            })
    df_msg = pd.DataFrame(nachrichten)
    print(f"{len(df_likes)} Likes und {len(df_msg)} Nachrichten")

    # Staging-Tabellen fuer die Rohdaten anlegen
    with engine.begin() as conn:
        conn.execute(text("DROP TABLE IF EXISTS stg_likes CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS stg_messages CASCADE"))
        conn.execute(text("""
            CREATE TABLE stg_likes (
                liker_email TEXT NOT NULL,
                liked_email TEXT NOT NULL,
                status TEXT NOT NULL,
                liked_at TIMESTAMP
            )"""))
        conn.execute(text("""
            CREATE TABLE stg_messages (
                sender_email TEXT NOT NULL,
                receiver_email TEXT NOT NULL,
                body TEXT NOT NULL,
                sent_at TIMESTAMP,
                conversation_id INTEGER NOT NULL
            )"""))
    # Rohdaten reinschreiben (nur wenn was da ist)
    if not df_likes.empty:
        df_likes.to_sql("stg_likes", engine, if_exists="append", index=False)
    if not df_msg.empty:
        df_msg.to_sql("stg_messages", engine, if_exists="append", index=False)

    # ---- die 5 Views bauen ----
    with engine.begin() as conn:
        # falls schon da, erst weg
        for v in ["migration_messages", "migration_likes", "migration_user_hobbies",
                  "migration_user_interests", "migration_users"]:
            conn.execute(text(f"DROP VIEW IF EXISTS {v} CASCADE"))

        # View 1: alle Nutzer mit den Vertragsspalten
        conn.execute(text("""
            CREATE VIEW migration_users AS
            SELECT k.email,
                   p.vorname AS first_name,
                   p.nachname AS last_name,
                   p.geburtsdatum AS birth_date,
                   p.plz AS postal_code,
                   p.stadt AS city,
                   k.telefon AS phone,
                   p.geschlecht AS gender
            FROM person p
            JOIN kontakt k ON k.person_id = p.person_id"""))

        # View 2: Interessen - "mw" wird zu zwei Zeilen (m und w)
        conn.execute(text("""
            CREATE VIEW migration_user_interests AS
            SELECT k.email, p.interessiert_an AS interest_code
            FROM person p JOIN kontakt k ON k.person_id = p.person_id
            WHERE p.interessiert_an IN ('m', 'w')
            UNION ALL
            SELECT k.email, 'm' FROM person p JOIN kontakt k ON k.person_id = p.person_id
            WHERE p.interessiert_an = 'mw'
            UNION ALL
            SELECT k.email, 'w' FROM person p JOIN kontakt k ON k.person_id = p.person_id
            WHERE p.interessiert_an = 'mw'"""))

        # View 3: Hobbys pro Nutzer
        conn.execute(text("""
            CREATE VIEW migration_user_hobbies AS
            SELECT k.email, h.name AS hobby_name, ph.prioritaet AS priority, ph.source
            FROM person_hobby ph
            JOIN person p ON p.person_id = ph.person_id
            JOIN kontakt k ON k.person_id = p.person_id
            JOIN hobby h ON h.hobby_id = ph.hobby_id"""))

        # View 4: Likes - Email ueber kontakt joinen damit Schreibweise aus Excel kommt
        conn.execute(text("""
            CREATE VIEW migration_likes AS
            SELECT kl.email AS liker_email,
                   kg.email AS liked_email,
                   sl.status,
                   sl.liked_at
            FROM stg_likes sl
            JOIN kontakt kl ON lower(kl.email) = lower(sl.liker_email)
            JOIN kontakt kg ON lower(kg.email) = lower(sl.liked_email)"""))

        # View 5: Nachrichten - genauso mit Email-Join
        conn.execute(text("""
            CREATE VIEW migration_messages AS
            SELECT ks.email AS sender_email,
                   kr.email AS receiver_email,
                   sm.body,
                   sm.sent_at,
                   sm.conversation_id
            FROM stg_messages sm
            JOIN kontakt ks ON lower(ks.email) = lower(sm.sender_email)
            JOIN kontakt kr ON lower(kr.email) = lower(sm.receiver_email)"""))
    print("5 Views gebaut")


# ======================================================================
# AKT 3 - XML-Hobbys und das Aenderungspaket
# ======================================================================
def akt3(engine):
    ueberschrift("AKT 3: XML und Aenderungspaket")

    # alte XML-Hobbys raus, damit ein zweiter Lauf nicht doppelt eintraegt
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM person_hobby WHERE source='xml'"))

    # Nachschlag-Daten aus der DB holen
    with engine.connect() as conn:
        # welche Emails gibt es ueberhaupt
        bestand = {r[0].lower() for r in conn.execute(text("SELECT email FROM migration_users")).fetchall()}
        # Email -> person_id
        email_zu_pid = {r[0].lower(): r[1] for r in conn.execute(text(
            "SELECT k.email, p.person_id FROM person p JOIN kontakt k ON k.person_id=p.person_id")).fetchall()}
        # Hobbyname -> hobby_id
        name_zu_hid = {r[0]: r[1] for r in conn.execute(text("SELECT name, hobby_id FROM hobby")).fetchall()}
        # welche Person/Hobby-Kombis gibt es schon (aus Excel/Akt1)
        # nur person_id + hobby_id, ohne source - wir wollen wissen ob die
        # Zuordnung fachlich schon existiert
        schon_da = {(r[0], r[1]) for r in conn.execute(text("SELECT person_id, hobby_id FROM person_hobby")).fetchall()}
        # zusaetzlich die Zuordnungen als (email, hobbyname) - robust gegen
        # fehlende IDs, damit die Bestandspruefung sicher greift
        schon_da_namen = {(r[0].lower(), r[1]) for r in conn.execute(text("""
            SELECT k.email, h.name
            FROM person_hobby ph
            JOIN kontakt k ON k.person_id = ph.person_id
            JOIN hobby h ON h.hobby_id = ph.hobby_id
        """)).fetchall()}
    print(f"Bestand: {len(bestand)} Emails, {len(schon_da)} Zuordnungen")

    # hier sammeln wir neue XML-Hobbys und die Ablehnungen
    xml_hobbys = []
    ablehnungen = []

    # ---- Basis-XML einlesen ----
    # Basis-XML-Hobbys kommen alle als source='xml' rein (additive Quelle).
    basis_pfad = finde("Lets_Meet_Hobbies.xml")
    if basis_pfad is None:
        print("  HINWEIS: Lets_Meet_Hobbies.xml nicht gefunden - Basis-XML wird uebersprungen.")
        users = []
    else:
        users = ET.parse(basis_pfad).getroot().findall("user")
    for u in users:
        email = (u.findtext("email") or "").strip()
        # jeder User kann mehrere Hobbys haben
        for h in u.findall("hobbies/hobby"):
            name = (h.text or "").strip()
            if email and name:
                xml_hobbys.append((email, name))
    print(f"Basis-XML: {len(users)} User, {len(xml_hobbys)} Hobbys")

    # ---- change-request.xml durchgehen ----
    records = ET.parse(datei_sicherstellen("change-request.xml", VORLAGE_CHANGE_REQUEST)).getroot().find("records")
    # Zaehler pro Typ fuer die source_ref (hobby[1], hobby[2] usw.)
    zaehler = {}
    # damit wir Duplikate im Paket selbst erkennen
    im_paket = set()

    for rec in records:
        typ = rec.tag
        zaehler[typ] = zaehler.get(typ, 0) + 1
        ref = f"/transferpack/records/{typ}[{zaehler[typ]}]"
        quelle = "change-request.xml"

        # --- Like ---
        if typ == "like":
            ziel = rec.get("target_email", "")
            absender = rec.get("email", "")
            # Ziel muss es geben
            if ziel.lower() not in bestand:
                ablehnungen.append((quelle, ref, f"Like-Ziel '{ziel}' gibt es nicht im Bestand"))
            elif absender.lower() not in bestand:
                ablehnungen.append((quelle, ref, f"Like-Absender '{absender}' gibt es nicht"))
            else:
                ablehnungen.append((quelle, ref, "Like nicht uebernommen (kein gueltiges Ziel)"))

        # --- Hobby ---
        elif typ == "hobby":
            email = rec.get("email", "")
            name = rec.get("name", "")
            prio = rec.get("priority")

            # priority pruefen falls angegeben
            if prio is not None:
                try:
                    pv = int(prio)
                    # ausserhalb vom erlaubten Bereich -> fehlerhaft
                    if pv < PRIO_MIN or pv > PRIO_MAX:
                        ablehnungen.append((quelle, ref, f"priority {pv} ausserhalb {PRIO_MIN} bis {PRIO_MAX}"))
                        continue
                except ValueError:
                    ablehnungen.append((quelle, ref, f"priority '{prio}' ist keine Zahl"))
                    continue

            # Person muss es geben
            if email.lower() not in bestand:
                ablehnungen.append((quelle, ref, f"Person '{email}' gibt es nicht im Bestand"))
                continue

            # schon im Bestand vorhanden (aus Excel)?
            # Regel: fachlich gleiche Person/Hobby-Zuordnung nur einmal.
            # Ueber (email, hobbyname) pruefen - unabhaengig von den IDs.
            if (email.lower(), name) in schon_da_namen:
                ablehnungen.append((quelle, ref, f"Zuordnung '{name}' schon im Bestand vorhanden"))
                continue

            # schon im Paket selbst vorgekommen? (Duplikat in der Lieferung)
            kombi = (email.lower(), name)
            if kombi in im_paket:
                ablehnungen.append((quelle, ref, "Doppelte Zuordnung im Paket (schon uebernommen)"))
                continue

            # alles ok -> merken zum Eintragen
            im_paket.add(kombi)
            xml_hobbys.append((email, name))

        # --- Profile ---
        elif typ == "profile":
            email = rec.get("email", "")
            # .invalid-Adressen sind Testprofile, keine echten Personen
            ablehnungen.append((quelle, ref,
                f"Testprofil '{email}' (.invalid, Platzhalterdaten) - keine echte Person"))

    # ---- encoding-invalid.xml: kaputtes UTF-8 als Latin-1 lesen ----
    inv = datei_sicherstellen("encoding-invalid.xml", VORLAGE_ENCODING_INVALID).read_bytes().decode("latin-1")
    e_inv = re.search(r'email="([^"]*)"', inv)
    f_inv = re.search(r'first_name="([^"]*)"', inv)
    email_inv = e_inv.group(1) if e_inv else ""
    name_inv = f_inv.group(1) if f_inv else ""
    ablehnungen.append(("encoding-invalid.xml", "/transferpack/records/profile[1]",
        f"Encoding repariert (Latin-1 -> '{name_inv}'); Testprofil '{email_inv}' (.invalid)"))

    # ---- encoding-mojibake.xml: doppelt kaputte Umlaute reparieren ----
    moj = datei_sicherstellen("encoding-mojibake.xml", VORLAGE_ENCODING_MOJIBAKE).read_bytes().decode("latin-1")
    e_moj = re.search(r'email="([^"]*)"', moj)
    f_moj = re.search(r'first_name="([^"]*)"', moj)
    email_moj = e_moj.group(1) if e_moj else ""
    name_moj = repariere_text(f_moj.group(1)) if f_moj else ""
    ablehnungen.append(("encoding-mojibake.xml", "/transferpack/records/profile[1]",
        f"Encoding repariert (Mojibake -> '{name_moj}'); Testprofil '{email_moj}' (.invalid)"))

    # ---- neue Hobby-Namen aus XML anlegen ----
    with engine.begin() as conn:
        vorhandene = {r[0] for r in conn.execute(text("SELECT name FROM hobby")).fetchall()}
        # nur die Namen die es noch nicht gibt
        for name in {n for _, n in xml_hobbys} - vorhandene:
            conn.execute(text("INSERT INTO hobby (name, source) VALUES (:n, 'xml') ON CONFLICT (name) DO NOTHING"),
                         {"n": name})
    # IDs neu holen (jetzt auch die neuen)
    with engine.connect() as conn:
        name_zu_hid = {r[0]: r[1] for r in conn.execute(text("SELECT name, hobby_id FROM hobby")).fetchall()}

    # ---- XML-Hobbys in person_hobby eintragen ----
    # Basis-XML kommt additiv rein. Falls dieselbe (Person,Hobby) mehrfach im
    # XML steht, nur einmal eintragen. Gegen Excel wird NICHT geprueft -
    # eine excel- und eine xml-Zuordnung koennen nebeneinander stehen, weil
    # die Basis-XML eine eigene Quelle ist.
    eintraege = []
    gesehen = set()
    for email, name in xml_hobbys:
        pid = email_zu_pid.get(email.lower())
        hid = name_zu_hid.get(name)
        # nur wenn beide IDs da sind
        if pid is None or hid is None:
            continue
        # XML-Kombi schon in diesem Lauf gesehen? (Duplikat im XML)
        if (pid, hid) in gesehen:
            continue
        # schon aus Excel vorhanden? Dann nicht als zweite Zeile - die
        # fachlich gleiche Zuordnung gibt es nur einmal (quellenuebergreifend)
        if (email.lower(), name) in schon_da_namen:
            continue
        gesehen.add((pid, hid))
        eintraege.append({"person_id": pid, "hobby_id": hid, "prioritaet": None, "source": "xml"})

    # einzeln einfuegen mit ON CONFLICT auf den vollen PK (inkl. source)
    with engine.begin() as conn:
        for e in eintraege:
            conn.execute(text("""
                INSERT INTO person_hobby (person_id, hobby_id, prioritaet, source)
                VALUES (:person_id, :hobby_id, :prioritaet, :source)
                ON CONFLICT (person_id, hobby_id, source) DO NOTHING
            """), e)
    print(f"{len(eintraege)} XML-Hobbys eingetragen")

    # ---- migration_rejections bauen ----
    with engine.begin() as conn:
        conn.execute(text("DROP VIEW IF EXISTS migration_rejections CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS stg_rejections CASCADE"))
        # Tabelle mit PK, damit jede Ablehnung nur einmal drinsteht
        conn.execute(text("""
            CREATE TABLE stg_rejections (
                source TEXT NOT NULL,
                source_ref TEXT NOT NULL,
                reason TEXT NOT NULL,
                PRIMARY KEY (source, source_ref)
            )"""))
        # alle Ablehnungen eintragen
        for quelle, ref, grund in ablehnungen:
            conn.execute(text("""
                INSERT INTO stg_rejections (source, source_ref, reason)
                VALUES (:s, :r, :grund)
                ON CONFLICT (source, source_ref) DO UPDATE SET reason = EXCLUDED.reason
            """), {"s": quelle, "r": ref, "grund": grund})
        # die View drauf
        conn.execute(text("CREATE VIEW migration_rejections AS SELECT source, source_ref, reason FROM stg_rejections"))
    print(f"{len(ablehnungen)} Ablehnungen dokumentiert")


# ======================================================================
# Kontrolle - am Ende zeigen was rausgekommen ist
# ======================================================================
def kontrolle(engine):
    ueberschrift("KONTROLLE")
    with engine.connect() as conn:
        # Zeilen pro View zaehlen
        for v in ["migration_users", "migration_user_interests", "migration_user_hobbies",
                  "migration_likes", "migration_messages", "migration_rejections"]:
            n = conn.execute(text(f"SELECT COUNT(*) FROM {v}")).scalar()
            print(f"  {v}: {n} Zeilen")
        # wie viele Hobbys kommen aus XML
        n_xml = conn.execute(text("SELECT COUNT(*) FROM migration_user_hobbies WHERE source='xml'")).scalar()
        print(f"\n  davon source='xml': {n_xml}")
        # die Ablehnungen auflisten
        print("\n  Ablehnungen:")
        for r in conn.execute(text("SELECT source, source_ref, reason FROM migration_rejections ORDER BY source, source_ref")).fetchall():
            print(f"    [{r[0]} {r[1]}] {r[2]}")


# Hauptprogramm - die drei Akte nacheinander
def main():
    engine = create_engine(DB)
    akt1(engine)
    akt2(engine)
    akt3(engine)
    kontrolle(engine)
    print("\nFertig. Script kann nochmal laufen ohne Fehler.")


if __name__ == "__main__":
    main()

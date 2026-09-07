"""
Schema:

  ort(ort_id PK, plz TEXT NOT NULL, stadt TEXT NOT NULL)
      - eine Zeile je (PLZ, Stadt)-Kombination.
      - (plz, stadt) zusammen UNIQUE.

  hobby(hobby_id PK, name TEXT NOT NULL UNIQUE)
      - eine Zeile je unterschiedlichem Hobby.

  person(person_id PK, install, imp,
         nachname NOT NULL, vorname NOT NULL,
         strasse_nr, telefon,
         email NOT NULL UNIQUE,
         geschlecht, interessiert_an,
         geburtsdatum NOT NULL,
         ort_id NOT NULL FK -> ort)
      - NOT NULL auf allen View-Pflichtfeldern.

  person_hobby(person_id FK, hobby_id FK, prioritaet)
      - zusammengesetzter PK (person_id, hobby_id).

  VIEW migration_users(email, first_name, last_name,
                       birth_date, postal_code, city)
      - alle Felder NOT NULL (via person + ort).

"""

import re                                          # Regex fuer das Zerlegen der Hobby-Spalte
import pandas as pd                                # Einlesen und Umformen der Excel-Daten
from pathlib import Path                           # plattformunabhaengige Pfade
from sqlalchemy import create_engine, text         # DB-Verbindung + rohes SQL ausfuehren

FILENAME = "Lets Meet DB Dump.xlsx"                # Name der Quelldatei, die gesucht wird
HOBBY_COL = "Hobby1 %Prio1%; Hobby2 %Prio2%; Hobby3 %Prio3%; Hobby4 %Prio4%; Hobby5 %Prio5%;"   # exakter Spaltenname in der Excel
HOBBY_PATTERN = re.compile(r"([^%;]+?)\s*%(\d+)%\s*;?")   # Gruppe 1 = Hobbyname, Gruppe 2 = Prioritaetszahl zwischen %%
TRENNER = ", "   # exakt: Komma + ein Leerzeichen, kein strip()   # Trennzeichen fuer Name/Adresse

DB_USER = "user"                                   # DB-Benutzername
DB_PASSWORD = "secret"                             # DB-Passwort
DB_HOST = "localhost"                              # DB-Host
DB_PORT = 5432                                     # PostgreSQL-Standardport
DB_NAME = "lf8_lets_meet_db"                       # Zieldatenbank

try:
    SCRIPT_DIR = Path(__file__).resolve().parent   # Ordner der Skriptdatei (normaler Aufruf)
except NameError:
    SCRIPT_DIR = Path.cwd()                        # Fallback: __file__ fehlt (Notebook/REPL) -> aktuelles Verzeichnis


def finde_xlsx(dateiname=FILENAME):
    kandidaten = [                                 # Liste wahrscheinlicher Speicherorte, in Pruefreihenfolge
        SCRIPT_DIR / dateiname, SCRIPT_DIR.parent / dateiname,      # neben dem Skript und eine Ebene darueber
        Path.cwd() / dateiname, Path.cwd().parent / dateiname,      # im Arbeitsverzeichnis und darueber
        Path.home() / dateiname,                                    # im Home-Verzeichnis
        Path.home() / "work" / dateiname,                           # ~/work
        Path.home() / "work" / "letsmeet" / dateiname,              # ~/work/letsmeet
        Path.home() / "LetsMeet" / dateiname,                       # ~/LetsMeet
    ]
    for pfad in kandidaten:                        # Kandidaten der Reihe nach durchgehen
        if pfad.exists():                          # erste existierende Datei gewinnt
            return pfad                            # gefundenen Pfad zurueckgeben
    treffer = list(Path.home().rglob(dateiname))   # Notfall: rekursiv das ganze Home durchsuchen (langsam)
    if treffer:                                    # wenn dabei etwas gefunden wurde
        return treffer[0]                          # den ersten Treffer nehmen
    orte = "\n".join(f"  - {k}" for k in kandidaten)   # Kandidatenliste fuer die Fehlermeldung formatieren
    raise FileNotFoundError(f"'{dateiname}' nicht gefunden:\n{orte}")   # Abbruch mit Hinweis, wo gesucht wurde


def section(title):
    print("\n" + "=" * 70)                         # Leerzeile + Trennlinie
    print(title)                                   # Ueberschrift des Abschnitts
    print("=" * 70)                                # Trennlinie darunter


def main():
    xlsx_path = finde_xlsx()                       # Quelldatei lokalisieren
    # dtype=str + keep_default_na=False: Rohtext exakt erhalten
    df = pd.read_excel(xlsx_path, dtype=str, keep_default_na=False)   # alles als Text, leere Zellen bleiben "" statt NaN
    print(f"Eingelesen: {xlsx_path} ({df.shape[0]} Zeilen)")          # Kontrollausgabe: Pfad + Zeilenzahl

    
    section("1. Namen aufteilen (Split an ', ', KEIN strip)")
    name_split = df["Nachname, Vorname"].str.split(TRENNER, n=1, expand=True)   # max. 1 Split -> 2 Spalten
    df["nachname"] = name_split[0]                 # Teil vor dem Komma
    df["vorname"]  = name_split[1]                 # Teil nach dem Komma

    section("2. Adresse aufteilen (Split an ', ' maxsplit=2, KEIN strip)")
    adr_split = df["Straße Nr, PLZ Ort"].str.split(TRENNER, n=2, expand=True)   # max. 2 Splits -> 3 Spalten
    df["strasse_nr"] = adr_split[0]                # "Musterweg 12"
    df["plz"]        = adr_split[1]                # Postleitzahl
    df["stadt"]      = adr_split[2]                # Ortsname (Rest der Zeichenkette)

    section("3. Hobbys parsen -> Hobby1..5 / Prio1..5")
    parsed = df[HOBBY_COL].apply(                  # je Zeile eine Liste aus (Hobby, Prio)-Paaren erzeugen
        lambda t: HOBBY_PATTERN.findall(t) if isinstance(t, str) else []   # Nicht-Strings -> leere Liste
    )
    for i in range(1, 6):                          # fuer die Hobby-Plaetze 1 bis 5
        df[f"Hobby{i}"] = parsed.apply(
            lambda p, i=i: p[i-1][0].strip() if len(p) >= i else None      # Name des i-ten Hobbys, sonst None
        )
        df[f"Prio{i}"] = parsed.apply(
            lambda p, i=i: int(p[i-1][1]) if len(p) >= i else None         # Prioritaet des i-ten Hobbys als int
        )
    print("Hobby-Spalten erstellt.")               # Statusmeldung

    section("4. Geburtsdatum parsen und Schluessel ergaenzen")
    df["geburtsdatum"] = pd.to_datetime(
        df["Geburtsdatum"], format="%d.%m.%Y", errors="coerce"   # deutsches Datumsformat; ungueltige Werte -> NaT
    ).dt.date                                                    # nur das Datum ohne Uhrzeit behalten
    nicht_parsebar = df["geburtsdatum"].isna().sum()             # Anzahl fehlgeschlagener Umwandlungen
    if nicht_parsebar:                                           # falls es welche gab
        print(f"WARNUNG: {nicht_parsebar} Geburtsdaten nicht parsebar (werden NULL).")   # warnen (verletzt spaeter NOT NULL)

    df.insert(0, "install", range(1, len(df) + 1))   # laufende Nummer als erste Spalte
    df.insert(1, "imp", "Excel")                     # Herkunftskennzeichen der Daten

    section("5. ort-Tabelle: eindeutige (plz, stadt)-Kombinationen")
    ort = (df[["plz", "stadt"]].drop_duplicates()               # Duplikate entfernen
           .sort_values(["plz", "stadt"]).reset_index(drop=True))   # sortieren und Index neu durchnummerieren
    ort.insert(0, "ort_id", range(1, len(ort) + 1))             # kuenstlichen Primaerschluessel vergeben
    ort_lookup = {(r.plz, r.stadt): r.ort_id for r in ort.itertuples()}   # Nachschlagetabelle (plz, stadt) -> ort_id
    print(f"{len(ort)} unterschiedliche Orte")                  # Kontrollausgabe

    section("6. hobby-Tabelle: eindeutige Hobby-Namen")
    alle_hobbys = pd.unique(
        pd.concat([df[f"Hobby{i}"] for i in range(1, 6)]).dropna()   # alle 5 Hobby-Spalten untereinander, ohne None
    )
    hobby = pd.DataFrame({"hobby_id": range(1, len(alle_hobbys) + 1),   # fortlaufende IDs
                           "name": alle_hobbys})                        # zugehoerige Namen
    hobby_lookup = {r.name: r.hobby_id for r in hobby.itertuples()}     # Nachschlagetabelle Name -> hobby_id
    print(f"{len(hobby)} unterschiedliche Hobbys")                      # Kontrollausgabe

    section("7. person-Tabelle (ohne Kontaktfelder)")
    person = pd.DataFrame({
        "person_id":      range(1, len(df) + 1),        # Primaerschluessel, Zeilenreihenfolge der Excel
        "install":        df["install"].values,         # laufende Nummer aus Schritt 4
        "imp":            df["imp"].values,             # Herkunft "Excel"
        "nachname":       df["nachname"].values,        # aus Schritt 1
        "vorname":        df["vorname"].values,         # aus Schritt 1
        "geschlecht":     df["Geschlecht (m/w/nonbinary)"].values,   # unveraendert aus der Excel
        "interessiert_an":df["Interessiert an"].values,              # unveraendert aus der Excel
        "geburtsdatum":   df["geburtsdatum"].values,                 # aus Schritt 4
        "ort_id":         [ort_lookup[(p, s)]                        # Fremdschluessel per Lookup aufloesen
                           for p, s in zip(df["plz"], df["stadt"])],
    })
    print(f"{len(person)} Personen")                    # Kontrollausgabe

    section("7b. kontakt-Tabelle: email, strasse_nr, telefon")
    kontakt = pd.DataFrame({
        "kontakt_id": range(1, len(df) + 1),            # eigener Primaerschluessel
        "person_id":  range(1, len(df) + 1),            # 1:1-Bezug zur person-Tabelle
        "email":      df["E-Mail"].values,              # E-Mail unveraendert
        "strasse_nr": df["strasse_nr"].values,          # aus Schritt 2
        "telefon":    df["Telefon"].values,             # Telefonnummer unveraendert
    })
    print(f"{len(kontakt)} Kontakt-Eintraege")          # Kontrollausgabe

    section("8. person_hobby-Tabelle: n:m-Zuordnung")
    zuordnungen = []                                    # Sammelliste fuer die Zuordnungszeilen
    for zeilen_idx, row in df.iterrows():               # jede Personenzeile durchgehen
        pid = zeilen_idx + 1                            # Index 0-basiert -> person_id 1-basiert
        for i in range(1, 6):                           # die bis zu 5 Hobby-Plaetze pruefen
            h = row[f"Hobby{i}"]                        # Hobbyname
            p = row[f"Prio{i}"]                         # zugehoerige Prioritaet
            if pd.notna(h):                             # nur belegte Plaetze uebernehmen
                zuordnungen.append({
                    "person_id": pid,                                   # Verweis auf person
                    "hobby_id":  hobby_lookup[h],                       # Verweis auf hobby
                    "prioritaet": int(p) if pd.notna(p) else None,      # Prioritaet als int oder NULL
                })
    person_hobby = pd.DataFrame(zuordnungen)            # Liste in DataFrame umwandeln
    print(f"{len(person_hobby)} Person-Hobby-Zuordnungen")   # Kontrollausgabe

    section("9. PostgreSQL: Objekte ersetzen und neu anlegen")
    engine = create_engine(                             # Verbindungsobjekt zur Datenbank bauen
        f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"    # Treiber + Zugangsdaten
        f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"                   # Host, Port, Datenbankname
    )

    with engine.begin() as conn:                        # Transaktion: am Ende automatisch COMMIT
        conn.execute(text("DROP VIEW  IF EXISTS migration_users CASCADE"))   # View zuerst weg
        conn.execute(text("DROP TABLE IF EXISTS letsmeet      CASCADE"))     # evtl. alte Rohtabelle entfernen
        conn.execute(text("DROP TABLE IF EXISTS person_hobby  CASCADE"))     # Reihenfolge: Kind- vor Elterntabellen
        conn.execute(text("DROP TABLE IF EXISTS kontakt       CASCADE"))     # haengt an person
        conn.execute(text("DROP TABLE IF EXISTS person        CASCADE"))     # haengt an ort
        conn.execute(text("DROP TABLE IF EXISTS hobby         CASCADE"))     # Stammtabelle
        conn.execute(text("DROP TABLE IF EXISTS ort           CASCADE"))     # Stammtabelle

        conn.execute(text("""
            CREATE TABLE ort (
                ort_id  INTEGER PRIMARY KEY,       -- kuenstlicher Schluessel
                plz     TEXT    NOT NULL,          -- PLZ als Text (fuehrende Nullen bleiben erhalten)
                stadt   TEXT    NOT NULL,          -- Ortsname
                UNIQUE (plz, stadt)                -- jede Kombination nur einmal
            )
        """))
        conn.execute(text("""
            CREATE TABLE hobby (
                hobby_id INTEGER PRIMARY KEY,      -- kuenstlicher Schluessel
                name     TEXT    NOT NULL UNIQUE   -- Hobbyname, keine Dubletten
            )
        """))
        conn.execute(text("""
            CREATE TABLE person (
                person_id       INTEGER PRIMARY KEY,   -- kuenstlicher Schluessel
                install         INTEGER,               -- laufende Nummer aus dem Import
                imp             TEXT,                  -- Herkunft der Daten
                nachname        TEXT    NOT NULL,      -- Pflichtfeld der View
                vorname         TEXT    NOT NULL,      -- Pflichtfeld der View
                geschlecht      TEXT,                  -- optional
                interessiert_an TEXT,                  -- optional
                geburtsdatum    DATE    NOT NULL,      -- Pflichtfeld der View
                ort_id          INTEGER NOT NULL REFERENCES ort(ort_id)   -- Fremdschluessel auf ort
            )
        """))
        conn.execute(text("""
            CREATE TABLE kontakt (
                kontakt_id INTEGER PRIMARY KEY,                                  -- kuenstlicher Schluessel
                person_id  INTEGER NOT NULL UNIQUE REFERENCES person(person_id), -- UNIQUE erzwingt 1:1 zu person
                email      TEXT    NOT NULL UNIQUE,                              -- Pflichtfeld, keine Dubletten
                strasse_nr TEXT,                                                 -- optional
                telefon    TEXT                                                  -- optional
            )
        """))
        conn.execute(text("""
            CREATE TABLE person_hobby (
                person_id   INTEGER NOT NULL REFERENCES person(person_id),   -- FK auf person
                hobby_id    INTEGER NOT NULL REFERENCES hobby(hobby_id),     -- FK auf hobby
                prioritaet  INTEGER,                                         -- Rang 1..5, optional
                PRIMARY KEY (person_id, hobby_id)                            -- jedes Hobby je Person nur einmal
            )
        """))
    print("Tabellen ort, hobby, person, person_hobby angelegt.")   # Statusmeldung

    ort.to_sql("ort",           engine, if_exists="append", index=False)   # Stammdaten zuerst (FK-Reihenfolge)
    hobby.to_sql("hobby",       engine, if_exists="append", index=False)   # Hobby-Stammdaten
    person.to_sql("person",     engine, if_exists="append", index=False)   # braucht ort
    kontakt.to_sql("kontakt",   engine, if_exists="append", index=False)   # braucht person
    person_hobby.to_sql("person_hobby", engine, if_exists="append", index=False)   # braucht person und hobby
    print("Alle Daten eingefuegt.")                    # Statusmeldung

    section("10. View migration_users auf normalisiertem Schema")
    with engine.begin() as conn:                       # neue Transaktion
        conn.execute(text("""
            CREATE VIEW migration_users AS             -- Zielformat fuer die Migration
            SELECT
                k.email         AS email,              -- aus kontakt
                p.vorname       AS first_name,         -- aus person
                p.nachname      AS last_name,          -- aus person
                p.geburtsdatum  AS birth_date,         -- aus person
                o.plz           AS postal_code,        -- aus ort
                o.stadt         AS city                -- aus ort
            FROM person p
            JOIN ort o     ON p.ort_id    = o.ort_id       -- Ort je Person
            JOIN kontakt k ON k.person_id = p.person_id    -- Kontaktdaten je Person
        """))
    print("View 'migration_users' angelegt.")          # Statusmeldung

    section("Kontrolle")
    with engine.connect() as conn:                     # nur lesende Verbindung
        for t in ["ort", "hobby", "person", "kontakt", "person_hobby"]:   # alle Tabellen durchgehen
            n = conn.execute(text(f"SELECT COUNT(*) FROM {t}")).scalar()  # Zeilen zaehlen
            print(f"  {t}: {n} Zeilen")                                   # Ergebnis ausgeben

        n_view = conn.execute(
            text("SELECT COUNT(*) FROM migration_users")).scalar()            # Zeilen in der View
        n_email = conn.execute(
            text("SELECT COUNT(DISTINCT email) FROM migration_users")).scalar()   # eindeutige E-Mails (Dublettenpruefung)
        print(f"  migration_users (View): {n_view} Zeilen, "
              f"{n_email} eindeutige E-Mails")                                # beides ausgeben

        nulls = conn.execute(text("""
            SELECT
                COUNT(*) FILTER (WHERE email IS NULL)        AS email_null,   -- je Pflichtfeld die NULL-Werte zaehlen
                COUNT(*) FILTER (WHERE first_name IS NULL)   AS fn_null,
                COUNT(*) FILTER (WHERE last_name IS NULL)    AS ln_null,
                COUNT(*) FILTER (WHERE birth_date IS NULL)   AS bd_null,
                COUNT(*) FILTER (WHERE postal_code IS NULL)  AS plz_null,
                COUNT(*) FILTER (WHERE city IS NULL)         AS city_null
            FROM migration_users
        """)).fetchone()                                                       # eine Ergebniszeile holen
        print(f"\n  NULL-Check View-Pflichtfelder: {dict(zip(nulls._fields, nulls))}")   # Spaltennamen mit Werten paaren

        print("\n  Beispiel-Join (Person 1 mit Ort und Hobbys):")   # Ueberschrift der Stichprobe
        ergebnis = conn.execute(text("""
            SELECT p.vorname, p.nachname, o.plz, o.stadt,
                   h.name AS hobby, ph.prioritaet          -- Person mit Ort und allen Hobbys
            FROM person p
            JOIN ort o ON p.ort_id = o.ort_id                       -- Ort dazu
            JOIN person_hobby ph ON ph.person_id = p.person_id      -- Zuordnungstabelle
            JOIN hobby h ON h.hobby_id = ph.hobby_id                -- Hobbynamen aufloesen
            WHERE p.person_id = 1                                   -- nur die erste Person
            ORDER BY ph.prioritaet DESC                             -- hoechste Prioritaetszahl zuerst
        """)).fetchall()                                            # alle Zeilen holen
        for zeile in ergebnis:                                      # Ergebniszeilen ausgeben
            print(f"    {zeile}")


if __name__ == "__main__":     # nur bei direktem Aufruf, nicht beim Import
    main()                     # Programm starten

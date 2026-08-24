"""
LetsMeet - Normalisierung in 3NF (mehrere Tabellen mit PK/FK)
====================================================================
Startet direkt bei 'Lets Meet DB Dump.xlsx' und baut das normalisierte
Schema in einem Durchlauf auf.

Ersetzt alle bisherigen Objekte (letsmeet, migration_users, ort, hobby,
person, person_hobby) durch das normalisierte Schema.

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

Rohdaten werden exakt nach Akt-1-Regeln aufgeteilt:
Split an ', ' (Komma + genau ein Leerzeichen), KEIN strip().

Verbindung: localhost:5432, DB lf8_lets_meet_db, User user.
Benoetigt: pip install sqlalchemy psycopg2-binary --break-system-packages
Aufruf:    python3 normalisierung.py
"""

import re
import pandas as pd
from pathlib import Path
from sqlalchemy import create_engine, text

FILENAME = "Lets Meet DB Dump.xlsx"
HOBBY_COL = "Hobby1 %Prio1%; Hobby2 %Prio2%; Hobby3 %Prio3%; Hobby4 %Prio4%; Hobby5 %Prio5%;"
HOBBY_PATTERN = re.compile(r"([^%;]+?)\s*%(\d+)%\s*;?")
TRENNER = ", "   # exakt: Komma + ein Leerzeichen, kein strip()

DB_USER = "user"
DB_PASSWORD = "secret"
DB_HOST = "localhost"
DB_PORT = 5432
DB_NAME = "lf8_lets_meet_db"

try:
    SCRIPT_DIR = Path(__file__).resolve().parent
except NameError:
    SCRIPT_DIR = Path.cwd()


def finde_xlsx(dateiname=FILENAME):
    kandidaten = [
        SCRIPT_DIR / dateiname, SCRIPT_DIR.parent / dateiname,
        Path.cwd() / dateiname, Path.cwd().parent / dateiname,
        Path.home() / dateiname,
        Path.home() / "work" / dateiname,
        Path.home() / "work" / "letsmeet" / dateiname,
        Path.home() / "LetsMeet" / dateiname,
    ]
    for pfad in kandidaten:
        if pfad.exists():
            return pfad
    treffer = list(Path.home().rglob(dateiname))
    if treffer:
        return treffer[0]
    orte = "\n".join(f"  - {k}" for k in kandidaten)
    raise FileNotFoundError(f"'{dateiname}' nicht gefunden:\n{orte}")


def section(title):
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def main():
    xlsx_path = finde_xlsx()
    # dtype=str + keep_default_na=False: Rohtext exakt erhalten
    df = pd.read_excel(xlsx_path, dtype=str, keep_default_na=False)
    print(f"Eingelesen: {xlsx_path} ({df.shape[0]} Zeilen)")

    # ------------------------------------------------------------------
    section("1. Namen aufteilen (Split an ', ', KEIN strip)")
    name_split = df["Nachname, Vorname"].str.split(TRENNER, n=1, expand=True)
    df["nachname"] = name_split[0]    # kein .strip()
    df["vorname"]  = name_split[1]    # kein .strip()

    # ------------------------------------------------------------------
    section("2. Adresse aufteilen (Split an ', ' maxsplit=2, KEIN strip)")
    adr_split = df["Straße Nr, PLZ Ort"].str.split(TRENNER, n=2, expand=True)
    df["strasse_nr"] = adr_split[0]
    df["plz"]        = adr_split[1]
    df["stadt"]      = adr_split[2]   # Ort = alles nach 2. Trenner

    # ------------------------------------------------------------------
    section("3. Hobbys parsen -> Hobby1..5 / Prio1..5")
    parsed = df[HOBBY_COL].apply(
        lambda t: HOBBY_PATTERN.findall(t) if isinstance(t, str) else []
    )
    for i in range(1, 6):
        df[f"Hobby{i}"] = parsed.apply(
            lambda p, i=i: p[i-1][0].strip() if len(p) >= i else None
        )
        df[f"Prio{i}"] = parsed.apply(
            lambda p, i=i: int(p[i-1][1]) if len(p) >= i else None
        )
    print("Hobby-Spalten erstellt.")

    # ------------------------------------------------------------------
    section("4. Geburtsdatum parsen und Schluessel ergaenzen")
    df["geburtsdatum"] = pd.to_datetime(
        df["Geburtsdatum"], format="%d.%m.%Y", errors="coerce"
    ).dt.date
    nicht_parsebar = df["geburtsdatum"].isna().sum()
    if nicht_parsebar:
        print(f"WARNUNG: {nicht_parsebar} Geburtsdaten nicht parsebar (werden NULL).")

    df.insert(0, "install", range(1, len(df) + 1))
    df.insert(1, "imp", "Excel")

    # ------------------------------------------------------------------
    section("5. ort-Tabelle: eindeutige (plz, stadt)-Kombinationen")
    ort = (df[["plz", "stadt"]].drop_duplicates()
           .sort_values(["plz", "stadt"]).reset_index(drop=True))
    ort.insert(0, "ort_id", range(1, len(ort) + 1))
    ort_lookup = {(r.plz, r.stadt): r.ort_id for r in ort.itertuples()}
    print(f"{len(ort)} unterschiedliche Orte")

    # ------------------------------------------------------------------
    section("6. hobby-Tabelle: eindeutige Hobby-Namen")
    alle_hobbys = pd.unique(
        pd.concat([df[f"Hobby{i}"] for i in range(1, 6)]).dropna()
    )
    hobby = pd.DataFrame({"hobby_id": range(1, len(alle_hobbys) + 1),
                           "name": alle_hobbys})
    hobby_lookup = {r.name: r.hobby_id for r in hobby.itertuples()}
    print(f"{len(hobby)} unterschiedliche Hobbys")

    # ------------------------------------------------------------------
    section("7. person-Tabelle")
    person = pd.DataFrame({
        "person_id":      range(1, len(df) + 1),
        "install":        df["install"].values,
        "imp":            df["imp"].values,
        "nachname":       df["nachname"].values,
        "vorname":        df["vorname"].values,
        "strasse_nr":     df["strasse_nr"].values,
        "telefon":        df["Telefon"].values,
        "email":          df["E-Mail"].values,
        "geschlecht":     df["Geschlecht (m/w/nonbinary)"].values,
        "interessiert_an":df["Interessiert an"].values,
        "geburtsdatum":   df["geburtsdatum"].values,
        "ort_id":         [ort_lookup[(p, s)]
                           for p, s in zip(df["plz"], df["stadt"])],
    })
    print(f"{len(person)} Personen")

    # ------------------------------------------------------------------
    section("8. person_hobby-Tabelle: n:m-Zuordnung")
    zuordnungen = []
    for zeilen_idx, row in df.iterrows():
        pid = zeilen_idx + 1
        for i in range(1, 6):
            h = row[f"Hobby{i}"]
            p = row[f"Prio{i}"]
            if pd.notna(h):
                zuordnungen.append({
                    "person_id": pid,
                    "hobby_id":  hobby_lookup[h],
                    "prioritaet": int(p) if pd.notna(p) else None,
                })
    person_hobby = pd.DataFrame(zuordnungen)
    print(f"{len(person_hobby)} Person-Hobby-Zuordnungen")

    # ------------------------------------------------------------------
    section("9. PostgreSQL: Objekte ersetzen und neu anlegen")
    engine = create_engine(
        f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}"
        f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"
    )

    with engine.begin() as conn:
        # In Abhaengigkeitsreihenfolge droppen
        conn.execute(text("DROP VIEW  IF EXISTS migration_users CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS letsmeet      CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS person_hobby  CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS person        CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS hobby         CASCADE"))
        conn.execute(text("DROP TABLE IF EXISTS ort           CASCADE"))

        conn.execute(text("""
            CREATE TABLE ort (
                ort_id  INTEGER PRIMARY KEY,
                plz     TEXT    NOT NULL,
                stadt   TEXT    NOT NULL,
                UNIQUE (plz, stadt)
            )
        """))
        conn.execute(text("""
            CREATE TABLE hobby (
                hobby_id INTEGER PRIMARY KEY,
                name     TEXT    NOT NULL UNIQUE
            )
        """))
        conn.execute(text("""
            CREATE TABLE person (
                person_id       INTEGER PRIMARY KEY,
                install         INTEGER,
                imp             TEXT,
                nachname        TEXT    NOT NULL,
                vorname         TEXT    NOT NULL,
                strasse_nr      TEXT,
                telefon         TEXT,
                email           TEXT    NOT NULL UNIQUE,
                geschlecht      TEXT,
                interessiert_an TEXT,
                geburtsdatum    DATE    NOT NULL,
                ort_id          INTEGER NOT NULL REFERENCES ort(ort_id)
            )
        """))
        conn.execute(text("""
            CREATE TABLE person_hobby (
                person_id   INTEGER NOT NULL REFERENCES person(person_id),
                hobby_id    INTEGER NOT NULL REFERENCES hobby(hobby_id),
                prioritaet  INTEGER,
                PRIMARY KEY (person_id, hobby_id)
            )
        """))
    print("Tabellen ort, hobby, person, person_hobby angelegt.")

    # In Reihenfolge der Abhaengigkeiten einfuegen
    ort.to_sql("ort",           engine, if_exists="append", index=False)
    hobby.to_sql("hobby",       engine, if_exists="append", index=False)
    person.to_sql("person",     engine, if_exists="append", index=False)
    person_hobby.to_sql("person_hobby", engine, if_exists="append", index=False)
    print("Alle Daten eingefuegt.")

    # ------------------------------------------------------------------
    section("10. View migration_users auf normalisiertem Schema")
    # Alle Felder NOT NULL: email/vorname/nachname/geburtsdatum via
    # person (NOT NULL-Constraints), plz/stadt via ort (NOT NULL).
    # INNER JOIN garantiert, dass kein NULL durch fehlenden ort entsteht.
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE VIEW migration_users AS
            SELECT
                p.email         AS email,
                p.vorname       AS first_name,
                p.nachname      AS last_name,
                p.geburtsdatum  AS birth_date,
                o.plz           AS postal_code,
                o.stadt         AS city
            FROM person p
            JOIN ort o ON p.ort_id = o.ort_id
        """))
    print("View 'migration_users' angelegt.")

    # ------------------------------------------------------------------
    section("Kontrolle")
    with engine.connect() as conn:
        for t in ["ort", "hobby", "person", "person_hobby"]:
            n = conn.execute(text(f"SELECT COUNT(*) FROM {t}")).scalar()
            print(f"  {t}: {n} Zeilen")

        n_view = conn.execute(
            text("SELECT COUNT(*) FROM migration_users")).scalar()
        n_email = conn.execute(
            text("SELECT COUNT(DISTINCT email) FROM migration_users")).scalar()
        print(f"  migration_users (View): {n_view} Zeilen, "
              f"{n_email} eindeutige E-Mails")

        # NULL-Check auf allen Pflichtfeldern der View
        nulls = conn.execute(text("""
            SELECT
                COUNT(*) FILTER (WHERE email IS NULL)        AS email_null,
                COUNT(*) FILTER (WHERE first_name IS NULL)   AS fn_null,
                COUNT(*) FILTER (WHERE last_name IS NULL)    AS ln_null,
                COUNT(*) FILTER (WHERE birth_date IS NULL)   AS bd_null,
                COUNT(*) FILTER (WHERE postal_code IS NULL)  AS plz_null,
                COUNT(*) FILTER (WHERE city IS NULL)         AS city_null
            FROM migration_users
        """)).fetchone()
        print(f"\n  NULL-Check View-Pflichtfelder: {dict(zip(nulls._fields, nulls))}")

        print("\n  Beispiel-Join (Person 1 mit Ort und Hobbys):")
        ergebnis = conn.execute(text("""
            SELECT p.vorname, p.nachname, o.plz, o.stadt,
                   h.name AS hobby, ph.prioritaet
            FROM person p
            JOIN ort o ON p.ort_id = o.ort_id
            JOIN person_hobby ph ON ph.person_id = p.person_id
            JOIN hobby h ON h.hobby_id = ph.hobby_id
            WHERE p.person_id = 1
            ORDER BY ph.prioritaet DESC
        """)).fetchall()
        for zeile in ergebnis:
            print(f"    {zeile}")


if __name__ == "__main__":
    main()
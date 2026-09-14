#!/bin/bash
# check-schulserver.sh — Prüfbefehl auf dem Schulserver mit durchgereichten
# Zusatzargumenten (z. B. --snapshot-out/--snapshot-compare für Akt 3).
#
# Der vorhandene Wrapper `letsmeet check <version>` reicht keine weiteren
# Argumente an die Node-CLI durch. Dieses Skript ruft dieselbe CLI mit
# denselben Umgebungsvariablen und Defaults wie der Wrapper auf und gibt
# alle Argumente nach der Vertragsversion unverändert weiter, samt Exitcode.
# Es setzt keine Daten zurück und installiert nichts.
#
# Aufruf:
#   bash scripts/check-schulserver.sh V3 --snapshot-out /pfad/snapshot.json
#   bash scripts/check-schulserver.sh V3 --snapshot-compare /pfad/snapshot.json
set -uo pipefail

VERSION="${1:-}"
if [ -z "$VERSION" ]; then
  echo "Version angeben, z. B.: bash scripts/check-schulserver.sh V1" >&2
  exit 2
fi
shift
case "$VERSION" in
  V1|V2|V3) ;;
  *) echo "Unbekannte Version: $VERSION (erlaubt: V1, V2, V3)" >&2; exit 2 ;;
esac

SHARED="${LETSMEET_SHARED:-/opt/letsmeet}"
NODE_DIR="${LETSMEET_NODE:-$SHARED/env}"
APP_DIR="$SHARED/app"
HOME_LM="${LETSMEET_HOME:-$HOME/work/letsmeet}"

PGHOST=127.0.0.1
PGPORT="${LETSMEET_PG_PORT:-5432}"
PGDATABASE=lf8_lets_meet_db
PGUSER=user
PGPASSWORD=secret
CHECK_HISTORY_PATH="${CHECK_HISTORY_PATH:-$HOME_LM/data/check-history.jsonl}"

NODE_BIN="$NODE_DIR/bin/node"
if [ ! -x "$NODE_BIN" ]; then
  echo "Node nicht gefunden: $NODE_BIN" >&2
  exit 2
fi
if [ ! -d "$APP_DIR" ]; then
  echo "App-Verzeichnis nicht gefunden: $APP_DIR" >&2
  exit 2
fi

cd "$APP_DIR" || exit 2

PGHOST="$PGHOST" PGPORT="$PGPORT" PGDATABASE="$PGDATABASE" PGUSER="$PGUSER" PGPASSWORD="$PGPASSWORD" \
CONTRACT_VERSION="$VERSION" CHECK_HISTORY_PATH="$CHECK_HISTORY_PATH" \
exec "$NODE_BIN" server/dist/cli.js "$@"

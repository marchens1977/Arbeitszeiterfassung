#!/bin/bash
# Startet alle automatischen Prüfungen nacheinander.  Start: bash werkzeuge/alle-tests.sh
cd "$(dirname "$0")/.." || exit 1
git fetch -q origin main 2>/dev/null
for t in werkzeuge/*.test.js; do
  echo "===== $t ====="
  node "$t" || { echo "FEHLER in $t"; exit 1; }
done
echo "===== Alle Prüfungen durchgelaufen ====="

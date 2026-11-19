#!/bin/zsh
# Doppelklick startet das Titelscreening mit der kriterien.toml aus diesem Ordner.
cd "$(dirname "$0")"
PY="$HOME/.venvs/screening/bin/python"

if [ ! -x "$PY" ]; then
  echo "Die Python-Umgebung fehlt. Einmalig im Terminal einrichten mit:"
  echo "  python3 -m venv ~/.venvs/screening"
  echo "  ~/.venvs/screening/bin/pip install openpyxl"
  echo
  read -k1 "?Taste drücken zum Schließen …"
  exit 1
fi

tomls=(*.toml(N))
if (( ${#tomls} > 1 )); then
  echo "Welche Kriterien-Datei? (Nummer eingeben)"
  select k in $tomls; do [[ -n $k ]] && break; done
else
  k=kriterien.toml
fi
"$PY" screening.py "" "$k"
echo
read -k1 "?Fertig – Taste drücken zum Schließen …"

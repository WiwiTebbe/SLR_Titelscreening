# SLR Titelscreening

Regelbasiertes Titel- und Abstract-Screening für systematische Literaturreviews (SLR).
Das Skript sortiert die Treffer einer Datenbankrecherche nach Ein- und Ausschlusskriterien,
die in einer einfachen Textdatei stehen, begründet jede Entscheidung und liefert die Zahlen
für das PRISMA-Flussdiagramm.

*Rule-based title/abstract screening for systematic literature reviews.
Interface and documentation are in German.*

## Funktionen

- Liest Exporte aus Scopus, Web of Science, PubMed, Rayyan u. a. (CSV/TSV/TXT, XLSX, XLS)
  sowie BibTeX-Dateien, z. B. aus der ACM Digital Library.
- Prüft jeden Treffer gegen deine Kriterien: erlaubte oder verbotene Werte (Sprache,
  Dokumenttyp), Jahresbereich, Suchbegriffe mit Platzhalter (`adolescen*`) in Titel,
  Abstract und Keywords, Pflichtfelder.
- Erkennt Duplikate über DOI und Titel.
- Schreibt zwei Dateien neben die Treffer-Datei:
  - `<name>_eingeschlossen.csv` bzw. `.xlsx`: nur die eingeschlossenen Treffer
  - `<name>_markiert.xlsx`: alle Treffer farbig markiert (grün = eingeschlossen,
    gelb = eingeschlossen, aber eine Angabe fehlte, rot = ausgeschlossen mit Begründung),
    dazu ein Blatt „Zusammenfassung“ mit den PRISMA-Zahlen und allen verwendeten Kriterien

## Voraussetzungen

- Python 3.11 oder neuer. Prüfen mit `python3 --version`. Das auf macOS vorinstallierte
  Python ist oft älter; dann eine aktuelle Version von python.org oder über Homebrew installieren.
- Das Paket `openpyxl`, für alte `.xls`-Dateien zusätzlich `xlrd`

## Einrichtung

Repository klonen oder als ZIP herunterladen, dann im Terminal in den Ordner wechseln:

```bash
python3 -m venv ~/.venvs/screening
~/.venvs/screening/bin/pip install -r requirements.txt
cp Beispiel/kriterien_beispiel.toml kriterien.toml
```

Die Startdatei für macOS sucht Python unter `~/.venvs/screening`. Wer nur im Terminal
arbeitet, kann die Umgebung auch woanders anlegen.

## Ausprobieren

```bash
~/.venvs/screening/bin/python screening.py Beispiel/beispiel_treffer.csv Beispiel/kriterien_beispiel.toml
```

Die 15 erfundenen Beispieltreffer decken alle Fälle ab: Duplikate, falsche Sprache,
zu altes Erscheinungsjahr, ausgeschlossener Dokumenttyp, Thema verfehlt, Tierstudie und
ein gelb markierter Treffer ohne Sprachangabe.

## Benutzen

1. `kriterien.toml` in einem Texteditor öffnen und die eigenen Kriterien eintragen.
   Oben in der Datei sind alle Prüfarten erklärt.
2. **macOS:** Doppelklick auf `Screening starten.command`, die Treffer-Datei ins
   Terminal-Fenster ziehen, Enter. Liegen mehrere `.toml`-Dateien im Ordner, fragt die
   Startdatei, welche gelten soll. Falls macOS den ersten Start blockiert:
   Rechtsklick → Öffnen.
3. **Terminal (alle Systeme)**, mit dem Python aus der eingerichteten Umgebung:

   ```bash
   python screening.py treffer.csv kriterien.toml
   python screening.py export.xlsx kriterien.toml --blatt "Tabelle1"
   python screening.py acm_export.bib kriterien.toml --ausgabe ergebnisse
   ```

   `--ausgabe` legt die Ergebnisse in einem anderen Ordner ab, `--blatt` wählt ein
   Tabellenblatt einer Excel-Datei. Ohne Kriterien-Datei nimmt das Skript `kriterien.toml`
   aus seinem eigenen Ordner.

## Hinweise zu BibTeX

- Die Spalte „Document Type“ enthält den BibTeX-Eintragstyp (`article`, `inproceedings` …),
  nicht „Conference Paper“ wie bei Scopus. Wer Konferenzbeiträge ausschließen will,
  ergänzt `"inproceedings"` bei den verbotenen Dokumenttypen.
- ACM-Exporte enthalten keine Sprachangabe; das Sprachkriterium wird dann übersprungen.
- Ergebnisse von BibTeX-Dateien werden immer als `.xlsx` gespeichert.

## Mehrere Reviews

Für jede SLR eine eigene Kriterien-Datei anlegen, z. B. `kriterien_review2.toml`.
Das Blatt „Zusammenfassung“ hält fest, welche Datei mit welchen Bedingungen verwendet wurde.
Für die Dokumentation die `.toml` zusammen mit den Rohdaten der Recherche aufbewahren.

## Grenzen

Das Tool arbeitet mit Wortlisten, nicht mit inhaltlichem Verständnis. Es übernimmt die
regelbasierte Vorauswahl, ersetzt aber kein inhaltliches Screening. Gelb markierte Treffer
von Hand prüfen und die Suchbegriffe vorab an einigen bekannten relevanten Arbeiten testen.

## Entstehung

Der Code entstand mit Unterstützung des KI-Assistenten Claude (Anthropic).

## Lizenz

MIT, siehe [LICENSE](LICENSE).

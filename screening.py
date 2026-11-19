#!/usr/bin/env python3
"""
Literatur-Screening nach Ein- und Ausschlusskriterien
=====================================================

Liest die Treffer einer Literaturrecherche (CSV/TSV/TXT, XLSX, XLS oder BibTeX) und prüft
jeden Treffer gegen die Kriterien aus einer TOML-Datei (Vorlage: kriterien.toml).
Ein Treffer wird eingeschlossen, wenn er alle Kriterien erfüllt, sonst ausgeschlossen.

Ausgabe (neben der Treffer-Datei oder im Ordner von --ausgabe):
  <name>_eingeschlossen.<endung>  nur die eingeschlossenen Treffer
                                  (CSV bleibt CSV, Excel wird .xlsx)
  <name>_markiert.xlsx            alle Treffer in Originalreihenfolge, zeilenweise eingefärbt:
                                    grün = eingeschlossen
                                    gelb = eingeschlossen, aber eine Angabe fehlte → bitte prüfen
                                    rot  = ausgeschlossen (Begründung in Spalte B)
                                  plus Blatt „Zusammenfassung“ mit Zahlen für das PRISMA-Diagramm

Aufruf:
  python3 screening.py treffer.csv kriterien.toml
  python3 screening.py treffer.xlsx kriterien.toml --ausgabe ergebnisse --blatt "Tabelle1"
  python3 screening.py acm_export.bib kriterien.toml
  python3 screening.py              (fragt nach der Treffer-Datei)

Benötigt Python 3.11+ und openpyxl (für alte .xls-Dateien zusätzlich xlrd).
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import re
import sys
import tomllib
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Das Paket 'openpyxl' fehlt. Installieren mit:  pip install openpyxl")


STATUS_EIN = "Eingeschlossen"
STATUS_PRUEFEN = "Eingeschlossen – bitte prüfen"
STATUS_AUS = "Ausgeschlossen"
SPALTE_STATUS = "Screening: Status"
SPALTE_GRUND = "Screening: Begründung"
NAME_DUPLIKAT = "Duplikat"

FUELLUNG = {
    STATUS_EIN: PatternFill("solid", start_color="C6EFCE"),
    STATUS_PRUEFEN: PatternFill("solid", start_color="FFEB9C"),
    STATUS_AUS: PatternFill("solid", start_color="FFC7CE"),
}
KOPF_FUELLUNG = PatternFill("solid", start_color="D9D9D9")
FETT = Font(bold=True)

# Erlaubte Angaben je Prüfart (zusätzlich zu name, typ, spalten, wenn_leer)
SCHLUESSEL = {
    "erlaubte_werte": {"werte"},
    "verbotene_werte": {"werte"},
    "zahlenbereich": {"min", "max"},
    "enthaelt": {"begriffe", "modus"},
    "enthaelt_nicht": {"begriffe"},
    "nicht_leer": set(),
}
GEMEINSAME_SCHLUESSEL = {"name", "typ", "spalten", "wenn_leer"}

# Mögliche Werte für wenn_leer und wie sie in der Zusammenfassung heißen
WENN_LEER = {
    "einschliessen": "einschliessen (gelb)",
    "ausschliessen": "ausschliessen",
    "ignorieren": "ignorieren (bleibt grün)",
}

# Gängige Spaltennamen aus Scopus, Web of Science, PubMed, Rayyan, Zotero und EndNote.
# Findet das Skript für einen in kriterien.toml angegebenen Namen keine Spalte, nimmt es
# stattdessen ein Synonym aus derselben Gruppe, das in der Datei vorkommt.
SYNONYME = [
    ["Title", "Article Title", "Document Title", "Primary Title", "TI", "Titel"],
    ["Abstract", "Abstract Note", "AB", "Zusammenfassung"],
    ["Author Keywords", "Keywords", "Index Keywords", "Keywords Plus", "MeSH Terms", "MeSH", "DE", "KW",
     "Manual Tags", "Schlagwörter", "Stichwörter"],
    ["Language of Original Document", "Language", "Lang", "LA", "Sprache"],
    ["Year", "Publication Year", "Pub Year", "PY", "Jahr", "Erscheinungsjahr"],
    ["Document Type", "Publication Type", "Item Type", "Reference Type", "Type", "DT", "PT", "Dokumenttyp"],
    ["DOI", "DI"],
    ["Authors", "Author", "AU", "Autoren"],
    ["Source Title", "Journal", "Booktitle", "Source", "Publication Title", "SO", "Zeitschrift"],
]


class Fehler(Exception):
    """Fehler mit einer verständlichen Meldung für den Nutzer."""


# --------------------------------------------------------------------------- Einlesen

@dataclass
class Tabelle:
    kopf: list[str]
    zeilen: list[list]
    trennzeichen: str | None = None  # gesetzt, wenn die Quelle eine CSV-Datei war


def als_text(wert) -> str:
    if wert is None:
        return ""
    if isinstance(wert, float) and wert.is_integer():
        return str(int(wert))
    if isinstance(wert, (dt.datetime, dt.date)):
        return wert.isoformat()
    return str(wert).strip()


def lese_tabelle(pfad: Path, blatt: str | None) -> Tabelle:
    if not pfad.is_file():
        raise Fehler(f"Datei nicht gefunden: {pfad}")
    endung = pfad.suffix.lower()
    if endung in (".csv", ".tsv", ".txt"):
        return _lese_csv(pfad)
    if endung in (".xlsx", ".xlsm"):
        return _lese_xlsx(pfad, blatt)
    if endung == ".xls":
        return _lese_xls(pfad, blatt)
    if endung in (".bib", ".bibtex"):
        return _lese_bib(pfad)
    raise Fehler(f"Dateiformat „{endung}“ wird nicht unterstützt (möglich: .csv, .tsv, .txt, .xlsx, .xls, .bib).")


def _lese_text(pfad: Path) -> str:
    roh = pfad.read_bytes()
    for kodierung in ("utf-8-sig", "cp1252"):
        try:
            return roh.decode(kodierung)
        except UnicodeDecodeError:
            pass
    return roh.decode("latin-1")  # kann nie scheitern


def _lese_csv(pfad: Path) -> Tabelle:
    text = _lese_text(pfad)
    # Trennzeichen = das häufigste Kandidatenzeichen in der Kopfzeile (Komma, Semikolon, Tab, |)
    kopfzeile = text.lstrip("\r\n").split("\n", 1)[0]
    trennzeichen = max(",;\t|", key=kopfzeile.count)
    csv.field_size_limit(2**31 - 1)
    zeilen = list(csv.reader(io.StringIO(text, newline=""), delimiter=trennzeichen))
    return _baue_tabelle(zeilen, trennzeichen)


def _lese_xlsx(pfad: Path, blatt: str | None) -> Tabelle:
    try:
        mappe = load_workbook(pfad, read_only=True, data_only=True)
    except Exception as e:
        raise Fehler(f"Excel-Datei konnte nicht gelesen werden: {e}") from None
    try:
        if blatt is None:
            ws = mappe.worksheets[0]
        elif blatt in mappe.sheetnames:
            ws = mappe[blatt]
        else:
            raise Fehler(f"Tabellenblatt „{blatt}“ gibt es nicht. Vorhanden: {', '.join(mappe.sheetnames)}")
        zeilen = [list(z) for z in ws.iter_rows(values_only=True)]
    finally:
        mappe.close()
    return _baue_tabelle(zeilen)


def _lese_xls(pfad: Path, blatt: str | None) -> Tabelle:
    try:
        import xlrd
    except ImportError:
        raise Fehler("Für alte .xls-Dateien wird das Paket 'xlrd' benötigt (pip install xlrd). "
                     "Alternativ die Datei in Excel als .xlsx speichern.") from None
    mappe = xlrd.open_workbook(pfad)
    if blatt is None:
        ws = mappe.sheet_by_index(0)
    elif blatt in mappe.sheet_names():
        ws = mappe.sheet_by_name(blatt)
    else:
        raise Fehler(f"Tabellenblatt „{blatt}“ gibt es nicht. Vorhanden: {', '.join(mappe.sheet_names())}")
    return _baue_tabelle([ws.row_values(i) for i in range(ws.nrows)])


# ----- BibTeX (z. B. ACM Digital Library, IEEE Xplore, Google Scholar, Zotero) -----

# Diese Felder stehen vorne, alle übrigen folgen in der Reihenfolge ihres ersten Auftretens.
BIB_REIHENFOLGE = ["author", "title", "year", "journal", "booktitle", "abstract", "keywords", "doi", "language"]
BIB_TYP = "Document Type"  # Eintragstyp wie article, inproceedings – passt zur Synonymgruppe Dokumenttyp
BIB_SCHLUESSEL = "Citation Key"


def _lese_bib(pfad: Path) -> Tabelle:
    eintraege = _parse_bibtex(_lese_text(pfad))
    if not eintraege:
        raise Fehler("In der BibTeX-Datei wurde kein Eintrag gefunden (Einträge beginnen mit @article{, "
                     "@inproceedings{ …).")
    felder: list[str] = []
    for e in eintraege:
        felder += [f for f in e["felder"] if f not in felder]
    felder = [f for f in BIB_REIHENFOLGE if f in felder] + [f for f in felder if f not in BIB_REIHENFOLGE]
    kopf = [BIB_TYP, BIB_SCHLUESSEL] + felder
    zeilen = [kopf] + [[e["typ"], e["schluessel"]] + [e["felder"].get(f, "") for f in felder] for e in eintraege]
    return _baue_tabelle(zeilen)


def _klammer_ende(text: str, start: int) -> int:
    """Position der Klammer, die die Klammer bei `start` schließt – geschweifte Klammern zählen mit."""
    oeffnend = text[start]
    schliessend = "}" if oeffnend == "{" else ")"
    tiefe, i = 0, start + 1
    while i < len(text):
        z = text[i]
        if z == "\\":
            i += 2
            continue
        if z == "{":
            tiefe += 1
        elif z == "}" and tiefe > 0:
            tiefe -= 1
        elif z == schliessend and tiefe == 0:
            return i
        i += 1
    raise Fehler(f"BibTeX: Eintrag ab Zeichen {start} wird nicht geschlossen (fehlt eine Klammer?).")


def _bib_wert(text: str, i: int, makros: dict[str, str]) -> tuple[str, int]:
    """Liest einen Feldwert ab Position i: {…}, "…", Zahl oder @string-Makro, auch mit # verkettet."""
    teile = []
    while True:
        while i < len(text) and text[i].isspace():
            i += 1
        if i >= len(text):
            break
        if text[i] == "{":
            ende = _klammer_ende(text, i)
            teile.append(text[i + 1:ende])
            i = ende + 1
        elif text[i] == '"':
            tiefe, j = 0, i + 1
            while j < len(text) and not (text[j] == '"' and tiefe == 0 and text[j - 1] != "\\"):
                tiefe += {"{": 1, "}": -1}.get(text[j], 0)
                j += 1
            teile.append(text[i + 1:j])
            i = j + 1
        else:
            m = re.compile(r"[^\s,#}]+").match(text, i)
            if not m:
                break
            wort = m.group()
            teile.append(makros.get(wort.lower(), wort))
            i = m.end()
        while i < len(text) and text[i].isspace():
            i += 1
        if i < len(text) and text[i] == "#":
            i += 1
            continue
        break
    return "".join(teile), i


def _bib_felder(text: str, makros: dict[str, str]) -> dict[str, str]:
    felder: dict[str, str] = {}
    i = 0
    name_re = re.compile(r"[\s,]*([A-Za-z][\w\-:.+]*)\s*=\s*")
    while (m := name_re.match(text, i)):
        wert, i = _bib_wert(text, m.end(), makros)
        felder.setdefault(m.group(1).lower(), wert)
    return felder


def _parse_bibtex(text: str) -> list[dict]:
    eintraege, makros = [], {"jan": "January", "feb": "February", "mar": "March", "apr": "April",
                              "may": "May", "jun": "June", "jul": "July", "aug": "August",
                              "sep": "September", "oct": "October", "nov": "November", "dec": "December"}
    kopf_re = re.compile(r"@\s*([A-Za-z]+)\s*([{(])")
    i = 0
    while (m := kopf_re.search(text, i)):
        typ = m.group(1).lower()
        ende = _klammer_ende(text, m.end() - 1)
        inhalt = text[m.end():ende]
        i = ende + 1
        if typ in ("comment", "preamble"):
            continue
        if typ == "string":
            makros.update({k: _latex_zu_text(v) for k, v in _bib_felder(inhalt, makros).items()})
            continue
        schluessel, _, rest = inhalt.partition(",")
        if "=" in schluessel:  # Eintrag ohne Zitierschlüssel
            schluessel, rest = "", inhalt
        felder = {k: _latex_zu_text(v) for k, v in _bib_felder(rest, makros).items()}
        if "author" in felder:  # „A and B and C“ → „A; B; C“ wie bei Scopus
            felder["author"] = "; ".join(a.strip() for a in re.split(r"\s+and\s+", felder["author"]) if a.strip())
        if "doi" in felder:
            felder["doi"] = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", felder["doi"], flags=re.I)
        eintraege.append({"typ": typ, "schluessel": schluessel.strip(), "felder": felder})
    return eintraege


AKZENTE = {'"': "\u0308", "'": "\u0301", "`": "\u0300", "^": "\u0302", "~": "\u0303", "=": "\u0304",
           ".": "\u0307", "c": "\u0327", "v": "\u030c", "u": "\u0306", "H": "\u030b", "k": "\u0328",
           "r": "\u030a"}
SONDERZEICHEN = {"ss": "ß", "o": "ø", "O": "Ø", "aa": "å", "AA": "Å", "ae": "æ", "AE": "Æ", "oe": "œ",
                 "OE": "Œ", "l": "ł", "L": "Ł", "i": "i", "j": "j"}


def _latex_zu_text(s: str) -> str:
    """Macht aus LaTeX-Schreibweisen lesbaren Text: {\\"u} → ü, \\& → &, {Augmented Reality} → Augmented Reality."""
    import unicodedata

    def akzent(m: re.Match) -> str:
        buchstabe = SONDERZEICHEN.get(m.group(2).lstrip("\\"), m.group(2).lstrip("\\"))
        return buchstabe + AKZENTE[m.group(1)]

    s = re.sub(r"\\([\"'`^~=.])\s*\{?\s*(\\?[A-Za-z])\s*\}?", akzent, s)
    s = re.sub(r"\\([cvuHkr])(?:\s*\{\s*(\\?[A-Za-z])\s*\}|\s+(?=[A-Za-z])([A-Za-z]))",
               lambda m: akzent(re.match(r"(.)(.*)", m.group(1) + (m.group(2) or m.group(3)))), s)
    s = re.sub(r"\\(ss|aa|AA|ae|AE|oe|OE|o|O|l|L|i|j)(?![A-Za-z])\s*", lambda m: SONDERZEICHEN[m.group(1)], s)
    s = unicodedata.normalize("NFC", s)
    for alt, neu in (("\\&", "&"), ("\\%", "%"), ("\\_", "_"), ("\\$", "$"), ("\\#", "#"),
                     ("\\{", "\x00"), ("\\}", "\x01"), ("---", "—"), ("--", "–"), ("``", "“"), ("''", "”")):
        s = s.replace(alt, neu)
    s = re.sub(r"(?<!\\)~", " ", s)
    s = re.sub(r"\\[A-Za-z]+\*?\s*", "", s)  # übrige Befehle wie \textit, \emph: Befehl weg, Inhalt bleibt
    s = s.replace("{", "").replace("}", "").replace("\x00", "{").replace("\x01", "}").replace("$", "")
    return re.sub(r"\s+", " ", s).strip()


def _baue_tabelle(rohzeilen: list[list], trennzeichen: str | None = None) -> Tabelle:
    zeilen = [list(z) for z in rohzeilen if any(als_text(w) for w in z)]
    if len(zeilen) < 2:
        raise Fehler("Die Datei enthält keine Treffer (leer oder nur Kopfzeile).")
    # Leere Spalten am rechten Rand abschneiden, kürzere Zeilen auffüllen
    breite = max(i + 1 for z in zeilen for i, w in enumerate(z) if als_text(w))
    zeilen = [["" if w is None else w for w in z[:breite]] + [""] * (breite - len(z)) for z in zeilen]
    # Kopfzeile = erste Zeile, die mindestens zur Hälfte gefüllt ist. So wird z. B. die
    # Titelzeile übersprungen, die Apple Numbers beim Excel-Export voranstellt.
    start = next((nr for nr, z in enumerate(zeilen[:20]) if 2 * sum(bool(als_text(w)) for w in z) >= breite), 0)
    zeilen = zeilen[start:]
    if len(zeilen) < 2:
        raise Fehler("Die Datei enthält keine Treffer (leer oder nur Kopfzeile).")
    kopf = [als_text(w) or f"Spalte {i + 1}" for i, w in enumerate(zeilen[0])]
    return Tabelle(kopf, zeilen[1:], trennzeichen)


def _spaltenschluessel(name: str) -> str:
    """'Author Keywords', 'author_keywords' und 'AuthorKeywords' gelten als dieselbe Spalte."""
    return re.sub(r"[\W_]+", "", name.casefold())


def spalten_indizes(kopf: list[str], namen: list[str]) -> list[int]:
    """Alle Spalten aus `namen`, die es in der Datei gibt – leer, wenn keine passt.

    Namen werden nach Synonymgruppe gebündelt (Titel, Abstract, Keywords …). Kommt aus einer
    Gruppe keiner der angegebenen Namen in der Datei vor, zählt jedes Synonym der Gruppe.
    """
    index: dict[str, int] = {}
    for i, k in enumerate(kopf):
        if s := _spaltenschluessel(k):
            index.setdefault(s, i)
    gruppen = [[_spaltenschluessel(n) for n in g] for g in SYNONYME]

    buendel: dict[int | str, list[str]] = {}
    for n in map(_spaltenschluessel, namen):
        gruppe = next((nr for nr, g in enumerate(gruppen) if n in g), n)
        buendel.setdefault(gruppe, []).append(n)

    gefunden = []
    for gruppe, gewuenscht in buendel.items():
        ersatz = gruppen[gruppe] if isinstance(gruppe, int) else []
        gefunden += [index[s] for s in gewuenscht if s in index] or [index[s] for s in ersatz if s in index]
    return list(dict.fromkeys(gefunden))


# --------------------------------------------------------------------------- Kriterien

@dataclass
class Kriterium:
    name: str
    typ: str
    spalten: list[str]
    eintraege: list[str] = field(default_factory=list)  # werte bzw. begriffe wie angegeben
    muster: list[re.Pattern] = field(default_factory=list)
    minimum: float | None = None
    maximum: float | None = None
    alle_begriffe: bool = False
    wenn_leer: str = "einschliessen"  # oder "ausschliessen" / "ignorieren"
    indizes: list[int] = field(default_factory=list)

    def beschreibung(self) -> str:
        liste = ", ".join(self.eintraege)
        if self.typ == "erlaubte_werte":
            return f"muss einer sein von: {liste}"
        if self.typ == "verbotene_werte":
            return f"darf keiner sein von: {liste}"
        if self.typ == "enthaelt":
            return f"{'alle' if self.alle_begriffe else 'mindestens einer'} von: {liste}"
        if self.typ == "enthaelt_nicht":
            return f"keiner von: {liste}"
        if self.typ == "nicht_leer":
            return "muss ausgefüllt sein"
        teile = []
        if self.minimum is not None:
            teile.append(f"mindestens {_zahl_text(self.minimum)}")
        if self.maximum is not None:
            teile.append(f"höchstens {_zahl_text(self.maximum)}")
        return ", ".join(teile)


def _zahl_text(zahl: float) -> str:
    return str(int(zahl)) if float(zahl).is_integer() else str(zahl)


def _normiere(wort: str) -> str:
    """'Enthält nicht' → 'enthaelt_nicht', damit Schreibvarianten in der TOML-Datei funktionieren."""
    for alt, neu in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss"), ("-", "_"), (" ", "_")):
        wort = wort.replace(alt, neu)
    return wort.strip().lower()


def _liste(wert, schluessel: str, wo: str) -> list[str]:
    if isinstance(wert, (str, int, float)) and not isinstance(wert, bool):
        wert = [wert]
    eintraege = []
    if isinstance(wert, list):
        eintraege = [str(w).strip() for w in wert
                     if isinstance(w, (str, int, float)) and not isinstance(w, bool) and str(w).strip()]
    if not eintraege:
        raise Fehler(f'{wo}: „{schluessel}“ fehlt oder ist leer – erwartet z. B. {schluessel} = ["…", "…"]')
    return eintraege


def _wert_muster(wert: str) -> re.Pattern:
    """Passt auf den ganzen Zellwert; * steht für beliebige Zeichen."""
    teile = (re.escape(t) for t in re.sub(r"\s+", " ", wert).split("*"))
    return re.compile(".*".join(teile), re.IGNORECASE)


def _begriff_muster(begriff: str) -> re.Pattern:
    """Passt auf ganze Wörter bzw. Wortfolgen im Text; * steht für beliebige Buchstaben."""
    woerter = (re.escape(w).replace(r"\*", r"\w*") for w in begriff.split())
    return re.compile(r"(?<!\w)" + r"[\s\-]+".join(woerter) + r"(?!\w)", re.IGNORECASE)


def lade_kriterien(pfad: Path) -> tuple[list[Kriterium], list[str]]:
    """Liefert die Kriterien und die Spalten für den Dublettenabgleich (leer = aus)."""
    if not pfad.is_file():
        raise Fehler(f"Kriterien-Datei nicht gefunden: {pfad}")
    try:
        daten = tomllib.loads(pfad.read_text(encoding="utf-8-sig"))
    except tomllib.TOMLDecodeError as e:
        raise Fehler(f"Formatfehler in {pfad.name}: {e}") from None

    unbekannt = set(daten) - {"kriterium", "duplikate"}
    if unbekannt:
        raise Fehler(f"Unbekannter Eintrag in {pfad.name}: {', '.join(sorted(unbekannt))} "
                     "(erlaubt sind [duplikate] und [[kriterium]]-Blöcke)")

    duplikate = daten.get("duplikate", {})
    if not isinstance(duplikate, dict) or set(duplikate) - {"entfernen", "spalten"}:
        raise Fehler("[duplikate] darf nur „entfernen“ und „spalten“ enthalten.")
    duplikat_spalten = _liste(duplikate.get("spalten"), "spalten", "[duplikate]") if duplikate.get("entfernen") else []

    roh = daten.get("kriterium", [])
    if not isinstance(roh, list):
        raise Fehler("Kriterien bitte als [[kriterium]] (doppelte eckige Klammern) angeben.")
    kriterien = [_baue_kriterium(eintrag, nr) for nr, eintrag in enumerate(roh, 1)]
    if not kriterien and not duplikat_spalten:
        raise Fehler(f"In {pfad.name} ist kein Kriterium aktiv.")
    return kriterien, duplikat_spalten


def _baue_kriterium(eintrag: dict, nr: int) -> Kriterium:
    name = str(eintrag.get("name") or f"Kriterium {nr}")
    wo = f"Kriterium „{name}“"
    typ = _normiere(str(eintrag.get("typ", "")))
    if typ not in SCHLUESSEL:
        raise Fehler(f"{wo}: unbekannter typ „{eintrag.get('typ', '')}“. Möglich: {', '.join(SCHLUESSEL)}")
    unbekannt = set(eintrag) - GEMEINSAME_SCHLUESSEL - SCHLUESSEL[typ]
    if unbekannt:
        raise Fehler(f"{wo}: unbekannte Angabe {', '.join(sorted(unbekannt))} für typ „{typ}“. "
                     f"Möglich: {', '.join(sorted(GEMEINSAME_SCHLUESSEL | SCHLUESSEL[typ]))}")

    k = Kriterium(name=name, typ=typ, spalten=_liste(eintrag.get("spalten"), "spalten", wo))

    k.wenn_leer = _normiere(str(eintrag.get("wenn_leer", "einschliessen")))
    if k.wenn_leer not in WENN_LEER:
        raise Fehler(f'{wo}: wenn_leer muss "einschliessen", "ausschliessen" oder "ignorieren" sein.')

    if typ in ("erlaubte_werte", "verbotene_werte"):
        k.eintraege = _liste(eintrag.get("werte"), "werte", wo)
        k.muster = [_wert_muster(w) for w in k.eintraege]
    elif typ in ("enthaelt", "enthaelt_nicht"):
        k.eintraege = _liste(eintrag.get("begriffe"), "begriffe", wo)
        k.muster = [_begriff_muster(b) for b in k.eintraege]
        modus = _normiere(str(eintrag.get("modus", "mindestens_eins")))
        if modus not in ("mindestens_eins", "alle"):
            raise Fehler(f'{wo}: modus muss "mindestens_eins" oder "alle" sein.')
        k.alle_begriffe = modus == "alle"
    elif typ == "zahlenbereich":
        k.minimum, k.maximum = eintrag.get("min"), eintrag.get("max")
        for wert in (k.minimum, k.maximum):
            if wert is not None and (isinstance(wert, bool) or not isinstance(wert, (int, float))):
                raise Fehler(f"{wo}: min und max müssen Zahlen ohne Anführungszeichen sein, z. B. min = 2010")
        if k.minimum is None and k.maximum is None:
            raise Fehler(f"{wo}: bitte min und/oder max angeben.")
    return k


# --------------------------------------------------------------------------- Prüfung

ZAHL = re.compile(r"-?\d+(?:[.,]\d+)?")


def _einzelwerte(werte: list[str]):
    """Ganzer Zellinhalt plus die Teile von Mehrfachangaben wie „English; German“."""
    for w in werte:
        w = re.sub(r"\s+", " ", w)
        yield w
        yield from (t.strip() for t in re.split(r"[;|,]", w) if t.strip())


def pruefe(k: Kriterium, zeile: list) -> tuple[str, str]:
    """Liefert ("ok" | "verletzt" | "leer", Detail für die Begründung)."""
    werte = [t for i in k.indizes if (t := als_text(zeile[i]))]
    if k.typ == "nicht_leer":
        return ("ok", "") if werte else ("verletzt", "Angabe fehlt")
    if not werte:
        return "leer", "Angabe fehlt"

    if k.typ == "erlaubte_werte":
        if any(m.fullmatch(t) for t in _einzelwerte(werte) for m in k.muster):
            return "ok", ""
        return "verletzt", f"Wert: {' | '.join(werte)}"

    if k.typ == "verbotene_werte":
        verboten = next((t for t in _einzelwerte(werte) if any(m.fullmatch(t) for m in k.muster)), None)
        return ("verletzt", f"Wert: {verboten}") if verboten else ("ok", "")

    if k.typ == "zahlenbereich":
        treffer = next((m.group() for w in werte if (m := ZAHL.search(w))), None)
        if treffer is None:
            return "leer", f"keine Zahl in „{werte[0]}“"
        zahl = float(treffer.replace(",", "."))
        if (k.minimum is not None and zahl < k.minimum) or (k.maximum is not None and zahl > k.maximum):
            return "verletzt", f"Wert: {_zahl_text(zahl)}"
        return "ok", ""

    text = "\n".join(werte)
    gefunden = [b for b, m in zip(k.eintraege, k.muster) if m.search(text)]
    if k.typ == "enthaelt_nicht":
        return ("verletzt", f"gefunden: {', '.join(gefunden)}") if gefunden else ("ok", "")
    if k.alle_begriffe:
        fehlend = [b for b in k.eintraege if b not in gefunden]
        return ("verletzt", f"fehlt: {', '.join(fehlend)}") if fehlend else ("ok", "")
    return ("ok", "") if gefunden else ("verletzt", "keiner der Begriffe gefunden")


def _dublettenschluessel(wert) -> str:
    """DOI in Kleinbuchstaben ohne https://doi.org/, sonst Text ohne Satzzeichen (z. B. Titel)."""
    s = als_text(wert).lower()
    s = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", s)
    if re.match(r"10\.\d{4,9}/\S+$", s):
        return s
    s = re.sub(r"[\W_]+", "", s)
    return s if len(s) >= 20 else ""  # sehr kurze Titel wie „Editorial“ nicht abgleichen


@dataclass
class Ergebnis:
    status: str
    begruendung: str
    verletzt: list[str]  # Namen der verletzten Kriterien in Prüfreihenfolge


def screene(tab: Tabelle, kriterien: list[Kriterium], duplikat_indizes: list[int]) -> list[Ergebnis]:
    ergebnisse = []
    erstes_vorkommen: dict[str, int] = {}
    for nr, zeile in enumerate(tab.zeilen):
        gruende, hinweise, verletzt = [], [], []

        if duplikat_indizes:
            schluessel = {s for i in duplikat_indizes if (s := _dublettenschluessel(zeile[i]))}
            original = min((erstes_vorkommen[s] for s in schluessel if s in erstes_vorkommen), default=None)
            if original is not None:
                gruende.append(f"Duplikat von Zeile {original + 2}")  # +2: Excel-Zeile inkl. Kopfzeile
                verletzt.append(NAME_DUPLIKAT)
            for s in schluessel:
                erstes_vorkommen.setdefault(s, nr if original is None else original)

        for k in kriterien:
            ergebnis, detail = pruefe(k, zeile)
            if ergebnis == "verletzt" or (ergebnis == "leer" and k.wenn_leer == "ausschliessen"):
                gruende.append(f"{k.name} ({detail})")
                verletzt.append(k.name)
            elif ergebnis == "leer" and k.wenn_leer == "einschliessen":
                hinweise.append(f"{k.name} ({detail})")

        if gruende:
            ergebnisse.append(Ergebnis(STATUS_AUS, "; ".join(gruende), verletzt))
        elif hinweise:
            ergebnisse.append(Ergebnis(STATUS_PRUEFEN, "Nicht prüfbar: " + "; ".join(hinweise), []))
        else:
            ergebnisse.append(Ergebnis(STATUS_EIN, "", []))
    return ergebnisse


# --------------------------------------------------------------------------- Ausgabe

GANZZAHL = re.compile(r"-?(0|[1-9]\d{0,14})")


def _excel_wert(wert, aus_csv: bool):
    if isinstance(wert, str):
        if aus_csv and GANZZAHL.fullmatch(wert.strip()):
            return int(wert)  # z. B. Jahreszahlen aus CSV als Zahl, damit Excel-Filter funktionieren
        return ILLEGAL_CHARACTERS_RE.sub("", wert)[:32767]
    return wert


def _schreibe_blatt(ws, kopf: list[str], zeilen: list[list], aus_csv: bool, fuellungen=None) -> None:
    ws.append(kopf)
    for zelle in ws[1]:
        zelle.font = FETT
        zelle.fill = KOPF_FUELLUNG
    for r, zeile in enumerate(zeilen, start=2):
        fuellung = fuellungen[r - 2] if fuellungen else None
        for c, wert in enumerate(zeile, start=1):
            zelle = ws.cell(row=r, column=c, value=_excel_wert(wert, aus_csv))
            if zelle.data_type == "f":  # Text, der mit „=“ beginnt, nicht als Formel auswerten
                zelle.data_type = "s"
            if fuellung:
                zelle.fill = fuellung
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for c, name in enumerate(kopf):
        laenge = max([len(name)] + [len(als_text(z[c])) for z in zeilen[:300]])
        ws.column_dimensions[get_column_letter(c + 1)].width = min(max(laenge + 2, 10), 60)


def _speichere(mappe: Workbook, pfad: Path) -> None:
    try:
        mappe.save(pfad)
    except PermissionError:
        raise Fehler(f"{pfad.name} konnte nicht gespeichert werden – ist die Datei noch in Excel geöffnet?") from None


def schreibe_eingeschlossen(pfad: Path, tab: Tabelle, ergebnisse: list[Ergebnis]) -> None:
    kopf = [SPALTE_STATUS, SPALTE_GRUND] + tab.kopf
    zeilen = [[e.status, e.begruendung] + z for e, z in zip(ergebnisse, tab.zeilen) if e.status != STATUS_AUS]
    if tab.trennzeichen is None:
        mappe = Workbook()
        mappe.active.title = "Eingeschlossen"
        _schreibe_blatt(mappe.active, kopf, zeilen, aus_csv=False)
        _speichere(mappe, pfad)
        return
    try:
        # utf-8-sig, damit Excel Umlaute beim Doppelklick richtig anzeigt
        with pfad.open("w", encoding="utf-8-sig", newline="") as f:
            schreiber = csv.writer(f, delimiter=tab.trennzeichen)
            schreiber.writerow(kopf)
            schreiber.writerows(zeilen)
    except PermissionError:
        raise Fehler(f"{pfad.name} konnte nicht gespeichert werden – ist die Datei noch geöffnet?") from None


def _spalten_text(tab: Tabelle, indizes: list[int]) -> str:
    return ", ".join(tab.kopf[i] for i in indizes) if indizes else "übersprungen – keine passende Spalte"


def schreibe_markiert(pfad: Path, tab: Tabelle, ergebnisse: list[Ergebnis], kriterien: list[Kriterium],
                      duplikat_spalten: list[str], duplikat_indizes: list[int],
                      treffer: Path, kriteriendatei: Path) -> None:
    mappe = Workbook()
    ws = mappe.active
    ws.title = "Screening"
    kopf = [SPALTE_STATUS, SPALTE_GRUND] + tab.kopf
    zeilen = [[e.status, e.begruendung] + z for e, z in zip(ergebnisse, tab.zeilen)]
    _schreibe_blatt(ws, kopf, zeilen, tab.trennzeichen is not None, [FUELLUNG[e.status] for e in ergebnisse])

    ws = mappe.create_sheet("Zusammenfassung")

    def zeile(*werte, fett=False, fuellung=None):
        ws.append(list(werte) or [None])
        for zelle in ws[ws.max_row]:
            if fett:
                zelle.font = FETT
        if fuellung:
            ws.cell(row=ws.max_row, column=1).fill = fuellung

    anzahl = Counter(e.status for e in ergebnisse)
    zeile("Literatur-Screening – Zusammenfassung")
    ws["A1"].font = Font(bold=True, size=14)
    zeile()
    zeile("Treffer-Datei", treffer.name)
    zeile("Kriterien-Datei", kriteriendatei.name)
    zeile("Erstellt am", dt.datetime.now().strftime("%d.%m.%Y %H:%M"))
    zeile()
    zeile("Treffer gesamt", len(ergebnisse), fett=True)
    zeile("Eingeschlossen", anzahl[STATUS_EIN] + anzahl[STATUS_PRUEFEN], fuellung=FUELLUNG[STATUS_EIN])
    zeile("   davon bitte prüfen", anzahl[STATUS_PRUEFEN], fuellung=FUELLUNG[STATUS_PRUEFEN])
    zeile("Ausgeschlossen", anzahl[STATUS_AUS], fuellung=FUELLUNG[STATUS_AUS])
    zeile()
    zeile("Ausschlussgrund", "Treffer, die das Kriterium verletzen", "davon als erster Grund gezählt", fett=True)
    verletzt = Counter(name for e in ergebnisse for name in e.verletzt)
    erster_grund = Counter(e.verletzt[0] for e in ergebnisse if e.verletzt)
    for name in ausschlussgruende([k for k in kriterien if k.indizes], duplikat_indizes):
        zeile(name, verletzt[name], erster_grund[name])
    zeile()
    zeile("Ein Treffer kann mehrere Kriterien verletzen. Die Spalte „davon als erster Grund gezählt“ "
          "ergibt zusammen genau die Zahl der Ausgeschlossenen (geeignet für das PRISMA-Flussdiagramm).")
    zeile()
    zeile("Farben", fett=True)
    zeile(STATUS_EIN, "erfüllt alle Kriterien", fuellung=FUELLUNG[STATUS_EIN])
    zeile(STATUS_PRUEFEN, "erfüllt alle prüfbaren Kriterien, aber eine Angabe fehlte – bitte von Hand prüfen",
          fuellung=FUELLUNG[STATUS_PRUEFEN])
    zeile(STATUS_AUS, "verletzt mindestens ein Kriterium (siehe Begründung)", fuellung=FUELLUNG[STATUS_AUS])
    zeile()
    zeile("Verwendete Kriterien", "Prüfart", "Geprüfte Spalten", "Bedingung", "Wenn Angabe fehlt", fett=True)
    if duplikat_spalten:
        zeile(NAME_DUPLIKAT, "duplikate", _spalten_text(tab, duplikat_indizes),
              "gleiche DOI bzw. gleicher Titel wie ein Treffer weiter oben", "–",
              fuellung=None if duplikat_indizes else FUELLUNG[STATUS_PRUEFEN])
    for k in kriterien:
        zeile(k.name, k.typ, _spalten_text(tab, k.indizes), k.beschreibung(),
              WENN_LEER[k.wenn_leer],
              fuellung=None if k.indizes else FUELLUNG[STATUS_PRUEFEN])
    for spalte, breite in zip("ABCDE", (38, 40, 35, 60, 22)):
        ws.column_dimensions[spalte].width = breite

    _speichere(mappe, pfad)


def ausschlussgruende(kriterien: list[Kriterium], duplikat_indizes: list[int]) -> list[str]:
    namen = ([NAME_DUPLIKAT] if duplikat_indizes else []) + [k.name for k in kriterien]
    return list(dict.fromkeys(namen))


def drucke_zuordnung(tab: Tabelle, kriterien: list[Kriterium], duplikat_spalten: list[str],
                     duplikat_indizes: list[int]) -> None:
    """Zeigt, welche Spalte für welches Kriterium verwendet wird und was übersprungen wird."""
    zuordnung = ([(NAME_DUPLIKAT, duplikat_indizes)] if duplikat_spalten else []) + [(k.name, k.indizes) for k in kriterien]
    breite = max(len(name) for name, _ in zuordnung)
    for name, indizes in zuordnung:
        zeichen = "✓" if indizes else "–"
        print(f"  {zeichen} {name:<{breite}}  → {_spalten_text(tab, indizes)}")
    if not all(indizes for _, indizes in zuordnung):
        print(f"\n  Vorhandene Spalten: {', '.join(tab.kopf)}")
    if not any(indizes for _, indizes in zuordnung):
        print("\nAchtung: Kein Kriterium passt zu dieser Datei – alle Treffer gelten als eingeschlossen.")


def drucke_zusammenfassung(ergebnisse: list[Ergebnis], kriterien: list[Kriterium], duplikat_indizes: list[int]):
    anzahl = Counter(e.status for e in ergebnisse)
    verletzt = Counter(name for e in ergebnisse for name in e.verletzt)
    erster_grund = Counter(e.verletzt[0] for e in ergebnisse if e.verletzt)
    namen = ausschlussgruende(kriterien, duplikat_indizes)
    breite = max(len(n) for n in namen + ["Ausschlussgrund"])

    print(f"\nTreffer gesamt:  {len(ergebnisse):>6}")
    print(f"Eingeschlossen:  {anzahl[STATUS_EIN] + anzahl[STATUS_PRUEFEN]:>6}"
          + (f"   (davon {anzahl[STATUS_PRUEFEN]} gelb – Angabe fehlte, bitte prüfen)" if anzahl[STATUS_PRUEFEN] else ""))
    print(f"Ausgeschlossen:  {anzahl[STATUS_AUS]:>6}\n")
    print(f"  {'Ausschlussgrund':<{breite}}  verletzt  davon erster Grund")
    for name in namen:
        print(f"  {name:<{breite}}  {verletzt[name]:>8}  {erster_grund[name]:>18}")


# --------------------------------------------------------------------------- Start

def _pfad_abfragen(frage: str) -> Path:
    # Beim Hineinziehen ins Terminal kommen Pfade mit Anführungszeichen oder „\ “ an
    eingabe = input(frage).strip().strip("'\"").replace("\\ ", " ")
    if not eingabe:
        raise Fehler("Keine Datei angegeben.")
    return Path(eingabe).expanduser()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sortiert die Treffer einer Literaturrecherche nach Ein- und Ausschlusskriterien.")
    parser.add_argument("treffer", nargs="?", help="CSV-, Excel- oder BibTeX-Datei mit den Treffern")
    parser.add_argument("kriterien", nargs="?",
                        help="TOML-Datei mit den Kriterien (Standard: kriterien.toml neben diesem Skript)")
    parser.add_argument("-a", "--ausgabe", help="Ordner für die Ergebnisdateien (Standard: Ordner der Treffer-Datei)")
    parser.add_argument("-b", "--blatt", help="Tabellenblatt bei Excel-Dateien (Standard: erstes Blatt)")
    args = parser.parse_args()

    try:
        if args.treffer:
            treffer = Path(args.treffer).expanduser()
        else:
            treffer = _pfad_abfragen("Treffer-Datei (CSV/Excel/BibTeX) – Pfad eingeben oder ins Fenster ziehen: ")
        if args.kriterien:
            kriteriendatei = Path(args.kriterien).expanduser()
        elif (standard := Path(__file__).with_name("kriterien.toml")).is_file():
            kriteriendatei = standard
        else:
            kriteriendatei = _pfad_abfragen("Kriterien-Datei (.toml): ")

        kriterien, duplikat_spalten = lade_kriterien(kriteriendatei)
        tab = lese_tabelle(treffer, args.blatt)
        # Kriterien ohne passende Spalte in dieser Datei werden übersprungen
        for k in kriterien:
            k.indizes = spalten_indizes(tab.kopf, k.spalten)
        duplikat_indizes = spalten_indizes(tab.kopf, duplikat_spalten)
        aktiv = [k for k in kriterien if k.indizes]
        print(f"Treffer:   {treffer.name} ({len(tab.zeilen)} Zeilen, {len(tab.kopf)} Spalten)")
        print(f"Kriterien: {kriteriendatei.name}")
        drucke_zuordnung(tab, kriterien, duplikat_spalten, duplikat_indizes)

        ergebnisse = screene(tab, aktiv, duplikat_indizes)

        ziel = Path(args.ausgabe).expanduser() if args.ausgabe else treffer.parent
        ziel.mkdir(parents=True, exist_ok=True)
        endung = treffer.suffix.lower() if tab.trennzeichen else ".xlsx"
        pfad_eingeschlossen = ziel / f"{treffer.stem}_eingeschlossen{endung}"
        pfad_markiert = ziel / f"{treffer.stem}_markiert.xlsx"
        schreibe_eingeschlossen(pfad_eingeschlossen, tab, ergebnisse)
        schreibe_markiert(pfad_markiert, tab, ergebnisse, kriterien, duplikat_spalten, duplikat_indizes,
                          treffer, kriteriendatei)
    except Fehler as e:
        sys.exit(f"Fehler: {e}")

    drucke_zusammenfassung(ergebnisse, aktiv, duplikat_indizes)
    print(f"\nGespeichert:\n  {pfad_eingeschlossen}\n  {pfad_markiert}")


if __name__ == "__main__":
    main()

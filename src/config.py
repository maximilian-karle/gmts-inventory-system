"""
config.py
=========
Zentrale Konfiguration fuer das GMTS-Projekt (Global Manufacturing Transition Strategy).

Dieses Modul ist die EINZIGE Stelle, an der die bekannten Produkttechnologien
und die daraus abgeleiteten Pfade (Input-/Output-Verzeichnisse je Technologie)
definiert werden. Der gesamte uebrige Code (data_loader, stock_analysis,
report_builder, main) bleibt technologie-agnostisch.

Hintergrund: Das Projekt ist als Template angelegt. main.py verarbeitet
standardmaessig ALLE Technologien in einem Lauf - es muss hier also nichts
mehr manuell ausgewaehlt werden. Fuer den gezielten Nachlauf einer einzelnen
Technologie siehe main_single.py, das ueber ein Kommandozeilenargument
gesteuert wird.

Zwei Input-Modi (siehe main.py fuer die Auswahllogik):
    1. GESAMTEXPORT-MODUS (bevorzugt, seit 24.06.2026): EIN ZMLAG-Export unter
       INPUT_ROOT / ZMLAG_FULL_EXPORT_FILENAME enthaelt alle Technologien
       gemeinsam. data_loader.split_by_technology() trennt anschliessend nach
       der SAP-Spalte 'Technologie'. Technologien, die nicht in
       KNOWN_TECHNOLOGIES gelistet sind, werden NICHT verworfen, sondern
       erhalten automatisch einen generierten Slug (siehe slugify()).
    2. EINZELORDNER-MODUS (Fallback, frueheres Verhalten): falls kein
       Gesamtexport vorliegt, wird weiterhin je Technologie ein eigener
       Ordner input_data/<slug>/ mit einem technologie-spezifischen Export
       erwartet. Beide Modi werden pro Lauf NICHT gemischt (siehe main.py).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd


# ---------------------------------------------------------------------------
# Bekannte Technologien im Unternehmen
# ---------------------------------------------------------------------------
# Jede Technologie wird im Kern weiterhin EINZELN ausgewertet (kein Vermischen
# der Kennzahlen mehrerer Technologien) - main.py iteriert lediglich automatisch
# ueber alle vorkommenden Technologien und erzeugt zusaetzlich einen
# konsolidierten Gesamt-Report (siehe main.py, report_builder.py).
#
# "slug" ist die dateisystemsichere Kurzform (fuer Ordnernamen), "label" ist
# die Originalschreibweise (fuer Report-Titel, Excel-Ueberschriften etc.).
#
# Hinweis: Diese Liste ist seit Einfuehrung des Gesamtexport-Modus keine
# harte Begrenzung mehr, sondern dient als "bekannte/erwartete" Zuordnung
# fuer huebsche Slugs und Labels. Taucht im SAP-Export eine Technologie auf,
# die hier NICHT gelistet ist, wird trotzdem ein Slug generiert (siehe
# slugify()) und die Technologie ganz normal mitverarbeitet - sie bekommt
# dann lediglich keinen "huebschen", sondern einen automatisch abgeleiteten
# Slug/Ordnernamen.
KNOWN_TECHNOLOGIES = {
    "cat_4": "CAT 4",
    "dcs_1_0": "DCS 1.0",
    "dcs_2_0": "DCS 2.0",
    "el_net": "EL.NET",
    "dcs_0_5": "DCS 0.5",
    "el_motion": "EL.MOTION",
    "el_industry": "EL.INDUSTRY",
    "el_capture": "EL.CAPTURE",
}

# Umgekehrtes Mapping (Label -> Slug) fuer den Gesamtexport-Modus: dort wird
# pro Material das tatsaechliche SAP-Label gelesen und muss auf den
# passenden, "huebschen" Slug zurueckgefuehrt werden, statt blind zu slugifyen
# (sonst wuerde z.B. "DCS 2.0" zu "dcs_2_0" neu generiert werden, was hier
# zufaellig identisch ist, bei anderen Labels aber nicht garantiert waere).
_LABEL_TO_SLUG = {label: slug for slug, label in KNOWN_TECHNOLOGIES.items()}


# ---------------------------------------------------------------------------
# Projekt-Basispfade
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_ROOT = PROJECT_ROOT / "input_data"
OUTPUT_ROOT = PROJECT_ROOT / "output_data"

# Erwarteter Dateiname des Gesamtexports im Gesamtexport-Modus (siehe
# Modul-Docstring). Liegt direkt unter INPUT_ROOT, da er nicht mehr einer
# einzelnen Technologie zugeordnet ist.
ZMLAG_FULL_EXPORT_FILENAME = "zmlag_export.xlsx"
ZMLAG_FULL_EXPORT_PATH = INPUT_ROOT / ZMLAG_FULL_EXPORT_FILENAME

# ---------------------------------------------------------------------------
# Weitere Datenquellen (Phase 2, ab 24.06.2026)
# ---------------------------------------------------------------------------
# Drei dauerhaft GETRENNTE Quellen (bewusste Architekturentscheidung, siehe
# Projektstatus.md Abschnitt "Datenquellen-Strategie"):
#   1. ZMLAG       - eigenstaendige Transaktion, nicht mergebar -> Einzelexport
#   2. MVER        - zu gross/eigenstaendig fuer SE16XXL-Merge -> Einzelexport
#   3. SE16XXL     - alle uebrigen Materialstamm-/Bestandsfelder, die sich
#                    per Mehrfachtabellen-Merge gewinnen lassen
# Diese Quellen werden NICHT zu einer einzigen Tabelle zusammengefasst,
# sondern in den jeweiligen Loader-Funktionen separat gelesen und erst auf
# Material-Ebene gejoint (siehe data_loader.py).

# MVER (36-Monats-Verbrauchshistorie, Gesamtexport ueber alle Technologien).
MVER_FULL_EXPORT_FILENAME = "mver_export.xlsx"
MVER_FULL_EXPORT_PATH = INPUT_ROOT / MVER_FULL_EXPORT_FILENAME

# SE16XXL-Merge-Report "Inventory" (liefert u.a. Ist-Bestandsmenge in Stueck
# ueber B~GesBestand sowie den Materialstatus A~MatStatus). Fuer Phase 2 wird
# ausschliesslich dieser zweite Report benoetigt.
SE16XXL_INVENTORY_EXPORT_FILENAME = "se16xxl_inventory_export.xlsx"
SE16XXL_INVENTORY_EXPORT_PATH = INPUT_ROOT / SE16XXL_INVENTORY_EXPORT_FILENAME

# SE16XXL-Merge-Report "Dispo_ABCXYZ" (grosser Stammdaten-Report: Planlieferzeit,
# WE-Bearbeitungszeit, Meldebestand, Sicherheitsbestand, Preise, etc.).
# Wird fuer Phase 2 noch NICHT eingelesen - relevant erst fuer Phase 4/5
# (Wiederbeschaffungszeit, Fruehwarnsystem). Pfad wird hier bereits definiert,
# damit der Dateiname projektweit konsistent ist, sobald ein Loader dafuer
# gebaut wird.
SE16XXL_DISPO_ABCXYZ_EXPORT_FILENAME = "se16xxl_dispo_abcxyz_export.xlsx"
SE16XXL_DISPO_ABCXYZ_EXPORT_PATH = INPUT_ROOT / SE16XXL_DISPO_ABCXYZ_EXPORT_FILENAME

# SE16XXL-Merge-Report "Open_Orders" (offene Bestellungen/Lieferplan-
# einteilungen: C~-Block = Bestellkopf/-position, D~-Block = Einteilungszeile
# mit Liefertermin und Mengenfeldern). Grundlage fuer Baustein B (Reichweite
# "Stock + bestaetigte Zugaenge", bisher zurueckgestellt, siehe
# Projektstatus.md Abschnitt 3/4c). Neu ab 01.07.2026, siehe
# open_orders_loader.py.
SE16XXL_OPEN_ORDERS_EXPORT_FILENAME = "se16xxl_open_orders_export.xlsx"
SE16XXL_OPEN_ORDERS_EXPORT_PATH = INPUT_ROOT / SE16XXL_OPEN_ORDERS_EXPORT_FILENAME


# ---------------------------------------------------------------------------
# Service-Level-Matrix nach ABC/XYZ-Klasse (Phase 5, ab 29.06.2026)
# ---------------------------------------------------------------------------
# Zentrale fachliche Annahme fuer das Fruehwarnsystem (safety_stock.py):
# welcher Ziel-Servicegrad (Wahrscheinlichkeit, waehrend der
# Wiederbeschaffungszeit NICHT in eine Fehlmenge zu laufen) je ABC/XYZ-
# Kombination angestrebt wird. Aus diesem Wert leitet sich der z-Wert fuer
# die Safety-Stock-Formel ab (SS = z * sigma * sqrt(L)).
#
# Herkunft: A-Zeile (AX=99%, AY=98%, AZ=97%) entspricht 1:1 dem von Max
# bereitgestellten Referenz-Template (Reiter '06_Inv_Calculation' in
# '00_Forecast_Template_Safety_Stock_Preissteuerung.xlsx'). B- und C-Zeile
# sind eine plausible, monoton fallende Fortsetzung (Entscheidung Claude,
# 29.06.2026, von Max zur Pruefung freigegeben) - GERINGERER Wertanteil
# (B/C) UND volatilerer Verbrauch (Y/Z) rechtfertigen jeweils einen
# niedrigeren Ziel-Servicegrad, da das Kapitalbindungsrisiko eines zu hohen
# Sicherheitsbestands bei diesen Materialien staerker wiegt als bei A/X.
#
# Bewusst hier in config.py statt in safety_stock.py: dies ist eine
# fachliche Geschaeftsregel (kann sich aendern, sobald Joachim/Max andere
# Ziel-Servicegrade vorgeben), keine Implementierungslogik - analog dazu,
# dass KNOWN_TECHNOLOGIES ebenfalls hier zentral liegt statt in einem
# einzelnen Berechnungsmodul.
#
# WICHTIG: Diese Matrix ist eine sichtbare Annahme, kein "stiller Default".
# Aenderung hier wirkt sich beim naechsten Lauf direkt aus - keine
# Code-Anpassung in safety_stock.py notwendig.
SERVICE_LEVEL_MATRIX: dict[str, dict[str, float]] = {
    "A": {"X": 0.99, "Y": 0.98, "Z": 0.97},
    "B": {"X": 0.98, "Y": 0.96, "Z": 0.94},
    "C": {"X": 0.96, "Y": 0.93, "Z": 0.90},
}

# Naeherungsweise Tage-pro-Monat-Umrechnung fuer die Wiederbeschaffungszeit
# (dispo_abcxyz_loader liefert Wiederbeschaffungszeit_Tage, Phase 3/Baustein
# A rechnen durchgaengig in Monaten - siehe trend_analysis.py,
# simulation_analysis.py). Bewusst ein einfacher Faktor (30) statt
# kalendergenauer Monatslaengen, da die Wiederbeschaffungszeit selbst schon
# eine SAP-seitige Naeherung ist (Planlieferzeit + WE-Bearbeitungszeit in
# Tagen) - eine kalendergenaue Umrechnung wuerde hier Praezision suggerieren,
# die in den Ausgangsdaten nicht vorhanden ist.
DAYS_PER_MONTH = 30.0


# ---------------------------------------------------------------------------
# Reichweite-vs-Wiederbeschaffung-Matrix (HTML-Dashboard, 30.06.2026)
# ---------------------------------------------------------------------------
# Bucket-Grenzen (Monate) fuer die Kreuztabelle Wiederbeschaffungszeit
# (Zeilen) x Reichweite (Spalten) im HTML-Dashboard (siehe html_dashboard.py,
# Kachel "Reichweite-vs-Wiederbeschaffung-Matrix"). Zentral hier gepflegt,
# da sowohl der Python-Payload-Aufbau als auch die JS-seitige Einordnung
# (window.GMTS_BUCKET_EDGES) dieselben Grenzen brauchen. EDGES sind die
# oberen Trennwerte zwischen den Buckets (< edges[i] -> Bucket i, sonst der
# letzte, offene Bucket); LABELS eine Zeile laenger als EDGES (letzter
# Bucket ist nach oben offen).
REICHWEITE_WBZ_BUCKET_EDGES = [3, 6, 9, 15, 20, 30]
REICHWEITE_WBZ_BUCKET_LABELS = [
    "[0-3]", "[3-6]", "[6-9]", "[9-15]", "[15-20]", "[20-30]", "[30-999]",
]


# ---------------------------------------------------------------------------
# Finanzparameter fuer Working Capital/ROI (Phase 6, ab 30.06.2026)
# ---------------------------------------------------------------------------
# Zentrale fachliche Annahmen fuer working_capital.py (analog zu
# SERVICE_LEVEL_MATRIX oben: sichtbare Annahme, kein stiller Default -
# Aenderung hier wirkt sich beim naechsten Lauf direkt aus, keine
# Code-Anpassung in working_capital.py notwendig).
#
# Herkunft: alle drei Werte direkt von Max vorgegeben (Klaerung 30.06.2026,
# siehe Projektstatus.md Abschnitt 4i) - das Whitepaper
# Inventory_Management_I.pdf (Kap. 8) nennt nur die Formelstruktur und
# Variablennamen (holding_cost_rate, order_cost,
# EMERGENCY_FREIGHT_COST_PER_UNIT), KEINE konkreten Zahlenwerte fuer E&L.
#
# IMPLEMENTATION_COST/ANNUAL_MAINTENANCE_COST (Whitepaper Kap. 8.7/8.8)
# bleiben bewusst nicht definiert - Klaerung mit Max, 30.06.2026: keine
# ROI-Berechnung gegen Implementierungskosten, siehe working_capital.py
# fuer die stattdessen verwendete ROI-Definition (Annual_Savings / Working_
# Capital_Aktuell).

# Lagerhaltungskostensatz (% p.a. auf den Bestandswert) - deckt Kapital-
# kosten, Lagerflaeche, Versicherung, Veralterungsrisiko ab.
HOLDING_COST_RATE = 0.20

# Losfixkosten je Bestellvorgang (EUR), unabhaengig von Bestellmenge.
# ORDER_COST ist Stand 30.06.2026 NOCH NICHT in working_capital.py
# verwendet (siehe Modul-Docstring dort: Ordering_Cost benoetigt eine
# Bestellanzahl, die ohne Policy-Simulation (Baustein C, vorgemerkt) nicht
# verfuegbar ist) - bereits hier definiert, damit die Konstante projektweit
# bereitsteht, sobald Baustein C umgesetzt wird.
ORDER_COST = 75.0

# Aufschlag auf den Standardpreis bei Fehlmengen-Notfallbeschaffung (z.B.
# Luftfracht-/Express-Zuschlag) - bewusst RELATIV statt als fixer EUR-Betrag
# je Stueck, da E&L-Materialien stark unterschiedliche Stueckwerte haben
# (Entscheidung Claude, 30.06.2026, von Max bestaetigt). Ebenfalls noch
# NICHT in working_capital.py verwendet (siehe Modul-Docstring dort:
# Shortage_Cost benoetigt eine erwartete Fehlmengen-MENGE in Stueck, die
# Baustein A/simulation_analysis.py Stand 30.06.2026 nicht ausweist,
# sondern nur eine Fehlmengen-WAHRSCHEINLICHKEIT) - bereits hier definiert
# fuer die spaetere Erweiterung.
EMERGENCY_FREIGHT_SURCHARGE_RATE = 0.30


# ---------------------------------------------------------------------------
# Materialnummern-Normalisierung (ab 24.06.2026)
# ---------------------------------------------------------------------------
# SAP speichert Materialnummern intern als 18-stelliges Feld (MATNR, CHAR18),
# zeigt sie aber je nach Transaktion/Layout UNTERSCHIEDLICH lang an: MVER,
# SE16XXL ("Inventory" und "Dispo_ABCXYZ") liefern die vollen 18 Stellen mit
# fuehrenden Nullen (z.B. "000000000000385325"), waehrend ZMLAG eine gekuerzte
# 8-stellige Ansicht liefert (z.B. "00385325"). Ohne Normalisierung liefe ein
# Merge auf 'Material' zwischen diesen Quellen fuer JEDE Zeile ins Leere -
# genau das Symptom, das den Reichweite-Reiter leer aussehen liess (siehe
# Projektstatus.md, Aenderungshistorie 24.06.2026).
#
# Entscheidung Max, 24.06.2026: einheitlich auf die SAP-native 18-stellige
# Form normalisieren (statt umgekehrt auf 8 Stellen zu kuerzen), da MVER und
# SE16XXL bereits in diesem Format vorliegen und ZMLAG die Ausnahme ist.
# Dies gilt durchgaengig - auch fuer die Anzeige im Excel-Report (bewusst
# KEIN separates Kurzformat fuer die Ausgabe, um die Konsistenz ueber alle
# Quellen UND alle Reiter hinweg zu wahren).
MATERIAL_NUMBER_LENGTH = 18


def normalize_material_number(material: object) -> object:
    """Normalisiert eine Materialnummer auf die SAP-native 18-stellige Form
    (fuehrende Nullen aufgefuellt), damit ZMLAG, MVER und SE16XXL ueber
    denselben Join-Key zusammengefuehrt werden koennen (siehe Modul-
    Abschnitt oben).

    Erwartet eine rein numerische Zeichenkette (ggf. mit fuehrenden Nullen,
    die beim Einlesen als Text erhalten bleiben muessen - siehe dtype=str in
    den jeweiligen Loadern). Nicht-numerische oder leere Werte (NaN) werden
    unveraendert zurueckgegeben, damit z.B. fehlende Materialnummern weiterhin
    von der bestehenden Validierung (_validate_material_numbers) erkannt
    werden, statt hier bereits stillschweigend "verschluckt" zu werden.

    Beispiel: "00385325" -> "000000000000385325", "385325" -> "000000000000385325".

    Args:
        material: Rohwert aus einer der drei Quellen (str erwartet, aber
            defensiv gegenueber NaN/None/bereits korrekt formatierten Werten).

    Returns:
        18-stelliger String mit fuehrenden Nullen, oder der unveraenderte
        Eingabewert, falls er nicht normalisierbar ist (NaN, leer, nicht-
        numerisch).
    """
    if material is None:
        return material
    try:
        if pd.isna(material):
            return material
    except (TypeError, ValueError):
        pass

    text = str(material).strip()
    if text == "" or not text.isdigit():
        return material

    return text.zfill(MATERIAL_NUMBER_LENGTH)


def slugify(label: str) -> str:
    """Erzeugt einen dateisystemsicheren Slug aus einem beliebigen SAP-Label.

    Wird im Gesamtexport-Modus fuer Technologien verwendet, die NICHT in
    KNOWN_TECHNOLOGIES gelistet sind (z.B. eine neu hinzugekommene 9.
    Technologie oder ein Tippfehler in SAP) - diese werden bewusst nicht
    verworfen, sondern bekommen automatisch einen abgeleiteten Ordnernamen.

    Beispiel: "EL.SENSOR" -> "el_sensor", "DCS 3.0 Beta" -> "dcs_3_0_beta".

    Logik: kleinschreiben, alles ausser a-z/0-9 zu Unterstrich, mehrfache
    Unterstriche zusammenfassen, fuehrende/abschliessende Unterstriche entfernen.
    """
    normalized = re.sub(r"[^a-z0-9]+", "_", label.strip().lower())
    return normalized.strip("_") or "unbekannt"


def resolve_slug_for_label(label: str) -> str:
    """Liefert den Slug fuer ein SAP-Technologie-Label.

    Bekannte Labels (siehe KNOWN_TECHNOLOGIES) liefern ihren festen,
    "huebschen" Slug zurueck. Unbekannte Labels werden automatisch
    slugifiziert (siehe slugify()) - sie werden NICHT verworfen.
    """
    return _LABEL_TO_SLUG.get(label, slugify(label))


def require_columns(
    df: pd.DataFrame,
    required_columns,
    *,
    caller: str,
    arg_name: str,
) -> None:
    """Prueft, ob df alle required_columns enthaelt, sonst ValueError mit
    einheitlicher Fehlermeldung.

    Fasst das Validierungsmuster zusammen, das bisher in den
    _merge_inputs()-Funktionen von safety_stock.py, simulation_analysis.py
    und working_capital.py identisch wiederholt war (Modul-Konsolidierung,
    01.07.2026, siehe Abschnitt 6 Projektstatus.md - "Merge-Pattern-
    Helper"). Die eigentliche Merge-/Join-Logik sowie die fachlich je
    Modul unterschiedlichen Konsolen-Warnungen bei herausfallenden
    Materialien bleiben BEWUSST individuell in den jeweiligen Modulen
    (zu unterschiedlich in Wortlaut/Kontext, um sinnvoll generalisiert zu
    werden, ohne an Aussagekraft fuer die COO-Kommunikation zu verlieren -
    siehe dortige Docstrings).

    Args:
        df: zu pruefendes DataFrame.
        required_columns: Menge oder Liste der Pflichtspalten.
        caller: Modul- und Funktionsname fuer die Fehlermeldung, z.B.
            "safety_stock._merge_inputs()".
        arg_name: Name des Parameters in der aufrufenden Funktion, z.B.
            "coverage_df", fuer die Fehlermeldung.

    Raises:
        ValueError: falls eine oder mehrere Spalten fehlen. Meldungsformat
            identisch zum bisherigen, individuellen Code in den drei
            Modulen (keine Verhaltensaenderung).
    """
    missing = set(required_columns) - set(df.columns)
    if missing:
        raise ValueError(
            f"{caller} erwartet in {arg_name} die Spalten "
            f"{sorted(required_columns)}, es fehlen: {sorted(missing)}."
        )


@dataclass(frozen=True)
class RunConfig:
    """Container fuer alle laufabhaengigen Einstellungen eines Analyse-Durchlaufs."""

    slug: str
    label: str
    output_dir: Path
    # input_dir ist nur im Einzelordner-Modus gesetzt (siehe Modul-Docstring).
    # Im Gesamtexport-Modus entfaellt ein technologie-spezifischer Input-Pfad,
    # da alle Technologien gemeinsam aus ZMLAG_FULL_EXPORT_PATH gelesen werden.
    input_dir: Optional[Path] = None

    @property
    def report_output_path(self) -> Path:
        """Zielpfad fuer den generierten Excel-Report dieses Laufs."""
        return self.output_dir / f"{self.slug}_bestandsreport.xlsx"

    @property
    def dashboard_html_path(self) -> Path:
        """VERALTET (seit 29.06.2026, Designueberarbeitung HTML-Dashboard,
        siehe Projektstatus.md Abschnitt 4h): Frueher Zielpfad fuer ein
        EIGENES HTML-Dashboard je Technologie. Ersetzt durch EINE zentrale
        Datei mit Technologie-Dropdown (config.CONSOLIDATED_DASHBOARD_HTML_
        PATH, siehe html_dashboard.write_combined_dashboard()) - main.py/
        main_single.py rufen diese Property nicht mehr auf. Bewusst NICHT
        entfernt (nur als veraltet markiert), falls externer Code noch
        darauf zugreift."""
        return self.output_dir / f"{self.slug}_dashboard.html"


# Zielpfad fuer den konsolidierten Gesamt-Report ueber alle Technologien
# (siehe main.py / report_builder.build_consolidated_report). Liegt bewusst
# direkt unter OUTPUT_ROOT statt in einem einzelnen Technologie-Ordner, da
# er technologieuebergreifend ist.
CONSOLIDATED_REPORT_PATH = OUTPUT_ROOT / "alle_technologien_bestandsreport.xlsx"

# Analoger Zielpfad fuer das konsolidierte Plotly-HTML-Dashboard (Phase 6,
# ab 29.06.2026, siehe html_dashboard.py) - liegt neben CONSOLIDATED_REPORT_
# PATH, gleicher Dateiname-Stamm wie der Excel-Report.
CONSOLIDATED_DASHBOARD_HTML_PATH = OUTPUT_ROOT / "alle_technologien_dashboard.html"


def get_config_for_slug(slug: str) -> RunConfig:
    """Liefert die RunConfig fuer einen explizit angegebenen, BEKANNTEN
    Technologie-Slug (siehe KNOWN_TECHNOLOGIES). Setzt zusaetzlich input_dir,
    da diese Funktion ausschliesslich fuer den Einzelordner-Modus (Fallback)
    verwendet wird - im Gesamtexport-Modus wird stattdessen
    get_config_for_label() genutzt.

    Wird von main.py / main_single.py im Einzelordner-Fallback aufgerufen.

    Raises:
        ValueError: wenn slug nicht in KNOWN_TECHNOLOGIES enthalten ist.
    """
    if slug not in KNOWN_TECHNOLOGIES:
        valid = ", ".join(KNOWN_TECHNOLOGIES.keys())
        raise ValueError(f"Unbekannte Technologie '{slug}'. Gueltige Werte sind: {valid}")

    label = KNOWN_TECHNOLOGIES[slug]
    return RunConfig(
        slug=slug,
        label=label,
        input_dir=INPUT_ROOT / slug,
        output_dir=OUTPUT_ROOT / slug,
    )


def get_config_for_label(label: str) -> RunConfig:
    """Liefert die RunConfig fuer ein SAP-Technologie-Label im Gesamtexport-
    Modus (siehe Modul-Docstring). Im Unterschied zu get_config_for_slug()
    ist hier KEINE Beschraenkung auf KNOWN_TECHNOLOGIES notwendig - jedes im
    Export vorkommende Label wird verarbeitet, bekannte Labels erhalten ihren
    festen Slug, unbekannte einen automatisch generierten (siehe
    resolve_slug_for_label()). input_dir bleibt unbesetzt (None), da im
    Gesamtexport-Modus kein technologie-spezifischer Input-Ordner existiert.
    """
    slug = resolve_slug_for_label(label)
    return RunConfig(
        slug=slug,
        label=label,
        output_dir=OUTPUT_ROOT / slug,
    )


def ensure_run_directories(cfg: RunConfig) -> None:
    """Stellt sicher, dass Input- (falls vorhanden) und Output-Verzeichnis
    fuer den Lauf existieren."""
    if cfg.input_dir is not None:
        cfg.input_dir.mkdir(parents=True, exist_ok=True)
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

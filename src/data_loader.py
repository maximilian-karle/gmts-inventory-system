"""
data_loader.py
===============
Einlesen und Validieren von SAP-Exporten fuer das GMTS-Projekt.

Aktuell unterstuetzt: ZMLAG (Materialstamm-Snapshot mit Bestandswerten,
ABC/XYZ-Klassifizierung und Technologie-Zuordnung).

Zwei Einlese-Modi (siehe config.py fuer den Hintergrund):
    1. GESAMTEXPORT-MODUS (bevorzugt, seit 24.06.2026):
       load_full_zmlag_export() liest EINEN ZMLAG-Export mit allen
       Technologien gemeinsam und gibt einen einzigen DataFrame zurueck.
       split_by_technology() trennt diesen anschliessend anhand der SAP-
       Spalte 'Technologie' in einzelne DataFrames je Technologie auf -
       unbekannte Technologien werden dabei NICHT verworfen.
    2. EINZELORDNER-MODUS (Fallback, frueheres Verhalten):
       load_zmlag_export_for_technology() liest weiterhin einen
       technologie-spezifischen Export aus einem eigenen Input-Ordner.

Beide Modi teilen sich dieselbe Kernvalidierung (Spalten, Header-Erkennung,
Summenzeilen-Filter, Materialnummern-Eindeutigkeit) ueber die gemeinsame
Hilfsfunktion _read_and_validate_zmlag_file().

Weitere Exporttypen (MVER fuer Verbrauchshistorie, offene Bestellungen, etc.)
werden in eigenen Funktionen ergaenzt, sobald die entsprechenden Beispieldaten
vorliegen - die Grundstruktur (Funktion je Exporttyp, gemeinsames Rueckgabeformat
als pandas DataFrame) bleibt dabei gleich.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import config
from config import RunConfig


# ---------------------------------------------------------------------------
# Erwartete Spalten im ZMLAG-Export
# ---------------------------------------------------------------------------
# Basis: reale Beispielzeile vom 23.06.2026.
# Die drei Stichtagsspalten (z.B. "202512", "202604", "202605") sind variabel
# (aendern sich je nach Exportzeitpunkt) und werden daher per Muster (6-stellig,
# numerisch, JJJJMM) erkannt statt hart codiert.
ZMLAG_FIXED_COLUMNS = [
    "Disponent",
    "Material",
    "Materialkurztext",
    "Normbezeichnung",
    "Produkthierarchie Dispo",
    "EL_01178 Bezeichnung Merkmalwert",
    "Laufender Wert",
    "ABC-Kennzeichen",
    "ABCXYZ-Kennzeichen",
]

# Spalte, die die Technologie-Zugehoerigkeit traegt (z.B. "DCS 2.0")
TECHNOLOGY_COLUMN = "EL_01178 Bezeichnung Merkmalwert"

# Erste fixe Spaltenueberschrift der eigentlichen Tabelle. Wird genutzt, um die
# Header-Zeile im Export zu finden (siehe _find_header_row).
HEADER_ANCHOR_COLUMN = "Disponent"

# Obergrenze, wie viele Zeilen am Anfang der Datei nach der Header-Zeile
# durchsucht werden. SAP liefert hier aktuell 4 Freitext-/Metadatenzeilen
# (Berichtstitel, Bewertungskreis, Bewertungsart, Leerzeile) vor der
# eigentlichen Tabelle - dieser Wert liegt bewusst mit Puffer darueber, damit
# kleinere Abweichungen im Export (z.B. eine zusaetzliche Metadatenzeile)
# nicht sofort zum Fehler fuehren.
MAX_HEADER_SEARCH_ROWS = 15


def _is_period_column(column_name: str) -> bool:
    """Erkennt Stichtags-Spalten im Format JJJJMM (z.B. '202512').

    SAP liefert hier variable, exportabhaengige Spaltennamen (die jeweils
    aktuellen Stichtage), daher Erkennung per Muster statt fester Liste.
    """
    name = str(column_name).strip()
    return len(name) == 6 and name.isdigit() and name[:2] in ("19", "20")


def _find_header_row(file_path: Path) -> int:
    """Sucht die Zeile im Excel-Export, die die echten Spaltenueberschriften
    enthaelt, und gibt ihren 0-basierten Zeilenindex zurueck (passend fuer
    den header=-Parameter von pd.read_excel).

    Hintergrund: ZMLAG-Exporte enthalten am Anfang mehrere Freitext-/
    Metadatenzeilen (Berichtstitel, Bewertungskreis, Bewertungsart, Leerzeile),
    bevor die eigentliche Tabelle mit Spaltenueberschriften wie 'Disponent',
    'Material' usw. beginnt. Statt diese Zeilenanzahl hart zu codieren, wird
    die Zeile gesucht, deren erste Spalte den Wert HEADER_ANCHOR_COLUMN
    ('Disponent') enthaelt - das macht den Loader robust gegenueber kleineren
    Abweichungen in der Anzahl der Kopfzeilen zwischen Exporten.

    Raises:
        ValueError: wenn innerhalb von MAX_HEADER_SEARCH_ROWS keine Zeile mit
            HEADER_ANCHOR_COLUMN in der ersten Spalte gefunden wird.
    """
    preview = pd.read_excel(
        file_path, header=None, nrows=MAX_HEADER_SEARCH_ROWS
    )

    first_column = preview.iloc[:, 0]
    matches = first_column[
        first_column.astype(str).str.strip() == HEADER_ANCHOR_COLUMN
    ]

    if matches.empty:
        raise ValueError(
            f"ZMLAG-Export '{file_path.name}': Konnte die Header-Zeile nicht "
            f"finden (gesucht: erste Spalte = '{HEADER_ANCHOR_COLUMN}' "
            f"innerhalb der ersten {MAX_HEADER_SEARCH_ROWS} Zeilen).\n"
            f"Pruefe, ob sich das Exportformat geaendert hat."
        )

    return int(matches.index[0])


def find_zmlag_file(input_dir: Path) -> Path:
    """Sucht die erste Excel-Datei im gegebenen Input-Verzeichnis.

    Es wird bewusst kein fester Dateiname erwartet, da SAP-Exporte je nach
    Nutzer/Zeitpunkt unterschiedlich benannt sein koennen. Wird nur im
    Einzelordner-Modus (Fallback) verwendet - im Gesamtexport-Modus ist der
    Dateiname fest (siehe config.ZMLAG_FULL_EXPORT_PATH).

    Raises:
        FileNotFoundError: wenn kein .xlsx im Verzeichnis liegt.
    """
    if not input_dir.exists():
        raise FileNotFoundError(
            f"Input-Verzeichnis existiert nicht: {input_dir}\n"
            f"Bitte ZMLAG-Export dort ablegen."
        )

    candidates = sorted(input_dir.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError(
            f"Kein ZMLAG-Export (.xlsx) gefunden in: {input_dir}\n"
            f"Bitte Export aus Transaktion ZMLAG dort ablegen."
        )

    return candidates[0]


def full_export_exists() -> bool:
    """Prueft, ob ein Gesamtexport (siehe config.ZMLAG_FULL_EXPORT_PATH)
    vorliegt. Wird von main.py / main_single.py genutzt, um zwischen
    Gesamtexport-Modus und Einzelordner-Fallback zu entscheiden (siehe
    config.py, Modul-Docstring) - die beiden Modi werden bewusst nicht pro
    Lauf gemischt.
    """
    return config.ZMLAG_FULL_EXPORT_PATH.exists()


def _read_and_validate_zmlag_file(file_path: Path) -> pd.DataFrame:
    """Gemeinsame Kernlogik zum Einlesen und fachlichen Validieren eines
    ZMLAG-Exports - unabhaengig davon, ob es sich um einen technologie-
    spezifischen Einzelexport oder den technologieuebergreifenden
    Gesamtexport handelt.

    Schritte: Header-Zeile finden, einlesen (Material als Text wg. fuehrender
    Nullen), Spalten validieren, SAP-Summenzeilen entfernen, Materialnummern
    auf Eindeutigkeit pruefen, 'Technologie'-Spalte materialgenau aus SAP
    setzen.

    Raises:
        FileNotFoundError: wenn file_path nicht existiert.
        ValueError: wenn erwartete Spalten fehlen oder Materialnummern doppelt sind.
    """
    header_row = _find_header_row(file_path)
    df = pd.read_excel(file_path, header=header_row, dtype={"Material": str})

    _validate_columns(df, file_path)
    df = _drop_summary_rows(df, file_path)

    # Normalisierung auf die SAP-native 18-stellige Form (siehe config.py,
    # Abschnitt "Materialnummern-Normalisierung"): ZMLAG liefert eine
    # gekuerzte 8-stellige Ansicht, MVER und SE16XXL dagegen bereits die
    # vollen 18 Stellen. Ohne diesen Schritt wuerde der spaetere Merge in
    # coverage_analysis.py fuer JEDE Zeile ins Leere laufen. Bewusst NACH
    # _drop_summary_rows() (Summenzeilen haben kein Material, NaN bleibt
    # durch normalize_material_number() unveraendert) und VOR der
    # Eindeutigkeitspruefung, damit diese bereits auf dem finalen Format prueft.
    df["Material"] = df["Material"].apply(config.normalize_material_number)

    _validate_material_numbers(df, file_path)

    df["Technologie"] = df[TECHNOLOGY_COLUMN]

    return df


def load_zmlag_export_for_technology(cfg: RunConfig) -> pd.DataFrame:
    """Liest den ZMLAG-Export fuer EINE Technologie aus ihrem eigenen
    Input-Ordner ein und validiert ihn (EINZELORDNER-MODUS / Fallback,
    siehe Modul-Docstring und config.py).

    Args:
        cfg: aktive RunConfig (siehe config.py); cfg.input_dir muss gesetzt sein.

    Returns:
        DataFrame mit einer Zeile pro Materialnummer, validierten Spalten,
        und einer Spalte 'Technologie'. Diese wird MATERIALGENAU aus der
        SAP-Spalte TECHNOLOGY_COLUMN uebernommen - NICHT pauschal aus
        cfg.label gesetzt (siehe fachlicher Hinweis zu
        _validate_technology_assignment unten).

        cfg.label bleibt weiterhin relevant: es definiert, welche Technologie
        fuer DIESEN Input-Ordner *erwartet* wird, und wird zur Plausibilitaets-
        pruefung genutzt (siehe _validate_technology_assignment).

    Raises:
        FileNotFoundError: wenn keine Exportdatei gefunden wird oder
            cfg.input_dir nicht gesetzt ist.
        ValueError: wenn erwartete Spalten fehlen oder Materialnummern doppelt sind.
    """
    if cfg.input_dir is None:
        raise FileNotFoundError(
            f"RunConfig fuer '{cfg.label}' hat keinen input_dir gesetzt - "
            f"load_zmlag_export_for_technology() ist nur fuer den "
            f"Einzelordner-Modus vorgesehen."
        )

    file_path = find_zmlag_file(cfg.input_dir)
    df = _read_and_validate_zmlag_file(file_path)
    _validate_technology_assignment(df, cfg, file_path)

    return df


def load_full_zmlag_export() -> pd.DataFrame:
    """Liest den technologieuebergreifenden ZMLAG-GESAMTEXPORT ein und
    validiert ihn (GESAMTEXPORT-MODUS, bevorzugt - siehe Modul-Docstring
    und config.py).

    Anders als load_zmlag_export_for_technology() wird hier KEINE
    Plausibilitaetspruefung gegen eine "erwartete" Technologie durchgefuehrt -
    im Gesamtexport ist das Vorkommen mehrerer Technologien der Normalfall,
    nicht ein Hinweis auf einen falsch abgelegten Export. Die Aufteilung
    nach Technologie erfolgt separat in split_by_technology().

    Returns:
        Ein einziger DataFrame mit allen Technologien, validierten Spalten,
        und einer Spalte 'Technologie' (materialgenau aus SAP).

    Raises:
        FileNotFoundError: wenn config.ZMLAG_FULL_EXPORT_PATH nicht existiert.
        ValueError: wenn erwartete Spalten fehlen oder Materialnummern doppelt sind
            (gilt ueber den GESAMTEN Export, da Materialnummern laut Bestaetigung
            unternehmensweit eindeutig genau einer Technologie zugeordnet sind).
    """
    file_path = config.ZMLAG_FULL_EXPORT_PATH
    if not file_path.exists():
        raise FileNotFoundError(
            f"Gesamtexport nicht gefunden: {file_path}\n"
            f"Bitte ZMLAG-Export mit allen Technologien dort ablegen "
            f"(Dateiname: '{config.ZMLAG_FULL_EXPORT_FILENAME}')."
        )

    return _read_and_validate_zmlag_file(file_path)


def split_by_technology(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Teilt einen Gesamtexport-DataFrame (siehe load_full_zmlag_export)
    anhand der Spalte 'Technologie' in einzelne DataFrames je Technologie auf.

    Bekannte Technologien (siehe config.KNOWN_TECHNOLOGIES) werden ueber
    ihr Label auf ihren festen, "huebschen" Slug abgebildet. Technologien,
    die NICHT bekannt sind, werden NICHT verworfen - sie erhalten
    automatisch einen generierten Slug (siehe config.slugify()) und werden
    ganz normal mitverarbeitet. Eine Konsolen-Meldung weist in diesem Fall
    darauf hin, damit Max das bei Bedarf nachpflegen kann
    (config.KNOWN_TECHNOLOGIES erweitern).

    Args:
        df: vollstaendig eingelesener und validierter Gesamtexport-DataFrame
            (mit Spalte 'Technologie').

    Returns:
        Mapping Slug -> DataFrame (nur die Zeilen dieser Technologie).
        Materialien mit fehlender Technologie-Kennzeichnung (NaN) werden
        herausgefiltert und als Konsolen-Warnung gemeldet, da sie keinem
        Report zugeordnet werden koennen.
    """
    missing_technology = df["Technologie"].isna()
    if missing_technology.any():
        affected = df.loc[missing_technology, "Material"].tolist()
        print(
            f"  Warnung: {len(affected)} Material(ien) ohne Technologie-"
            f"Kennzeichnung werden in keinem Einzelreport beruecksichtigt: "
            f"{affected}"
        )
        df = df[~missing_technology]

    result: dict[str, pd.DataFrame] = {}
    unknown_labels: set[str] = set()

    for label, group in df.groupby("Technologie"):
        slug = config.resolve_slug_for_label(label)
        if label not in config.KNOWN_TECHNOLOGIES.values():
            unknown_labels.add(label)
        result[slug] = group.reset_index(drop=True)

    if unknown_labels:
        print(
            f"  Hinweis: {len(unknown_labels)} im Export vorkommende "
            f"Technologie(n) sind nicht in config.KNOWN_TECHNOLOGIES "
            f"gelistet, werden aber trotzdem verarbeitet (automatisch "
            f"generierter Ordnername): {sorted(unknown_labels)}"
        )

    return result


def _drop_summary_rows(df: pd.DataFrame, file_path: Path) -> pd.DataFrame:
    """Entfernt abschliessende Summenzeilen, die SAP an manche ZMLAG-Exporte
    anhaengt (z.B. eine Gesamtsumme ueber alle Stichtags-Wertspalten).

    Erkennungsmerkmal: 'Material' ist leer, obwohl die Zeile sonst Werte
    enthaelt (insbesondere in den Stichtags-Spalten). Eine Zeile, die *gar
    keine* Werte enthaelt (komplett leer), wird hier nicht herausgefiltert -
    das waere ein anderes Problem und soll in _validate_material_numbers
    weiterhin auffallen.

    Jede entfernte Zeile wird auf der Konsole ausgegeben, damit nichts
    unbemerkt verschwindet, falls es sich doch nicht um eine harmlose
    Summenzeile handelt.
    """
    is_summary_row = df["Material"].isna() & df.drop(columns=["Material"]).notna().any(axis=1)

    if is_summary_row.any():
        removed = df[is_summary_row]
        print(
            f"  Hinweis: {len(removed)} Zeile(n) ohne Materialnummer "
            f"(vermutlich SAP-Summenzeile) aus '{file_path.name}' entfernt:"
        )
        for idx, row in removed.iterrows():
            preview = {col: row[col] for col in get_period_columns(df) if pd.notna(row[col])}
            print(f"    Zeile {idx}: {preview}")

    return df[~is_summary_row].reset_index(drop=True)


def _validate_columns(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft, ob alle fix erwarteten Spalten vorhanden sind."""
    missing = [col for col in ZMLAG_FIXED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"ZMLAG-Export '{file_path.name}' fehlen erwartete Spalten: {missing}\n"
            f"Vorhandene Spalten: {list(df.columns)}"
        )

    period_columns = [c for c in df.columns if _is_period_column(c)]
    if not period_columns:
        raise ValueError(
            f"ZMLAG-Export '{file_path.name}' enthaelt keine erkennbaren "
            f"Stichtags-Spalten (Format JJJJMM, z.B. '202512')."
        )


def _validate_material_numbers(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft auf fehlende oder doppelte Materialnummern.

    Da ZMLAG laut Bestaetigung genau eine Zeile pro Material liefert (Snapshot)
    und eine Materialnummer laut fachlicher Bestaetigung unternehmensweit
    eindeutig genau einer Technologie zugeordnet ist, sind Duplikate ein
    Hinweis auf einen fehlerhaften oder unerwartet strukturierten Export und
    werden daher hart abgefangen statt nur gewarnt. Dies gilt unveraendert
    auch im Gesamtexport-Modus (Pruefung ueber den GESAMTEN Export).
    """
    if df["Material"].isna().any():
        raise ValueError(
            f"ZMLAG-Export '{file_path.name}' enthaelt Zeilen ohne Materialnummer."
        )

    duplicates = df["Material"][df["Material"].duplicated()].unique()
    if len(duplicates) > 0:
        raise ValueError(
            f"ZMLAG-Export '{file_path.name}' enthaelt doppelte Materialnummern: "
            f"{list(duplicates)}\n"
            f"Erwartet wird eine Zeile pro Material (Snapshot)."
        )


def _validate_technology_assignment(
    df: pd.DataFrame, cfg: RunConfig, file_path: Path
) -> None:
    """Prueft, ob die in der Datei tatsaechlich vorkommenden Technologien
    (Spalte TECHNOLOGY_COLUMN) zur fuer diesen Input-Ordner erwarteten
    Technologie (cfg.label) passen.

    NUR relevant im EINZELORDNER-MODUS (Fallback) - im Gesamtexport-Modus
    gibt es keinen "erwarteten" Input-Ordner mehr und diese Pruefung
    entfaellt (siehe load_full_zmlag_export).

    Fachlicher Hintergrund: Pro Input-Ordner (z.B. input_data/dcs_2_0/) wird
    grundsaetzlich genau ein ZMLAG-Export mit Materialien einer Technologie
    erwartet. Da die Technologie-Spalte materialgenau aus den SAP-Daten
    uebernommen wird, wuerde ein falsch abgelegter Export (z.B. DCS-1.0-Datei
    im dcs_2_0-Ordner) sonst unbemerkt durchlaufen und im Report mit der
    (korrekten, aber unerwarteten) SAP-Technologie auftauchen.

    Dies wird bewusst NICHT hart abgefangen (kein Raise), da einzelne
    Materialien durchaus technologieuebergreifend genutzt werden koennen
    und das kein Datenfehler sein muss - es wird lediglich eine Konsolen-
    Warnung ausgegeben, damit Max das pruefen kann.
    """
    actual_technologies = set(df["Technologie"].dropna().unique())
    unexpected = actual_technologies - {cfg.label}

    if unexpected:
        print(
            f"  Warnung: '{file_path.name}' (Ordner fuer '{cfg.label}') "
            f"enthaelt Material(ien) mit abweichender Technologie-Kennzeichnung: "
            f"{sorted(unexpected)}. Bitte pruefen, ob der Export im richtigen "
            f"Input-Ordner liegt."
        )


def get_period_columns(df: pd.DataFrame) -> list[str]:
    """Hilfsfunktion: liefert alle erkannten Stichtags-Spalten eines geladenen DataFrames,
    sortiert chronologisch (aufsteigend nach JJJJMM)."""
    period_columns = [c for c in df.columns if _is_period_column(c)]
    return sorted(period_columns, key=lambda c: str(c))

"""
open_orders_loader.py
======================
Einlesen des SAP SE16XXL-Mehrfachtabellen-Merge-Exports "Open_Orders" -
Grundlage fuer Baustein B (Reichweite "Stock + bestaetigte Zugaenge",
bisher bewusst zurueckgestellt, siehe Projektstatus.md Abschnitt 3/4c).
Eigenstaendiges Modul neben dispo_abcxyz_loader.py/se16xxl_loader.py
(gleiches Muster: eigene Spaltenliste/Validierung/Normalisierung je Report,
keine gemeinsame Basis - Entscheidung Max, 25.06.2026, hier fortgefuehrt).

Der Export enthaelt neben den bekannten Stammdatenfeldern (A~/B~-Bloecke,
identisch zu Dispo_ABCXYZ/Inventory) zwei neue Bloecke:
    - C~: Bestellkopf/-position (C~EinkBeleg = Bestellnummer, C~Position =
      Position im Beleg, C~LoeschKz = Loeschkennzeichen)
    - D~: Lieferplan-Einteilungszeile (D~Einteilung, D~LiefDat =
      Liefertermin, D~EintMeng. = eingeteilte Menge, D~WE-Menge = bereits
      gebuchter Wareneingang zu dieser Einteilung)

Eine Bestellposition (C~EinkBeleg + C~Position) kann MEHRERE
Einteilungszeilen (Teillieferungen) haben - im Unterschied zu
Dispo_ABCXYZ/Inventory ist dieser Export daher NICHT eindeutig auf
Materialebene, sondern eine Zeile je Einteilung. Kein Zusammenfassen auf
Material-Ebene in diesem Loader (bleibt Aufgabe des aufrufenden Bausteins,
z.B. Summierung der offenen Menge je Material fuer die Reichweite-
Berechnung) - der Loader liefert die Rohdaten auf Einteilungs-Ebene, damit
keine Information (z.B. einzelne Liefertermine) vorzeitig verloren geht.

Fachliche Kernberechnung (offene Menge je Einteilung):
    Offene_Menge = D~EintMeng. - D~WE-Menge

Nur der offene Rest ist ein kuenftiger Zugang - bereits gebuchter
Wareneingang steckt schon im Bestand (ZMLAG/SE16XXL-Inventory) und darf
nicht doppelt gezaehlt werden (sonst Ueberschaetzung der Reichweite).

Designentscheidungen (Klaerung mit Max, 01.07.2026):
    - Stornierte Positionen (C~LoeschKz gesetzt) werden NICHT herausgefiltert,
      sondern behalten und zusaetzlich ueber die Spalte 'Storniert' (bool)
      markiert - die Entscheidung, ob eine stornierte Position bei der
      Zugangsberechnung ignoriert wird, bleibt bewusst beim aufrufenden
      Baustein (z.B. Baustein B filtert dort explizit), statt sie hier
      unsichtbar zu verlieren.
    - Voll gelieferte Einteilungen (Offene_Menge <= 0) werden NICHT
      herausgefiltert, sondern mit Offene_Menge = 0 behalten (Claude-
      Entscheidung, 01.07.2026: konsistent zum Projektprinzip "keine
      stillschweigenden Auffuellungen/Verluste" - die Zeile bleibt fuer
      spaetere Auswertungen sichtbar, z.B. Nachbestellhistorie, siehe
      Projektstatus.md Abschnitt 3, traegt aber korrekt mit 0 zur Summe der
      offenen Zugaenge bei). Negative Werte (Overdelivery: WE-Menge >
      EintMeng.) werden ebenfalls auf 0 gekappt, nicht als negativer Zugang
      gefuehrt.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import config


# Erwartete Spalten im SE16XXL-Open_Orders-Export (Reihenfolge wie im realen
# Export von Max, 01.07.2026). A~/B~-Bloecke wie im Dispo_ABCXYZ-/Inventory-
# Export bereits bekannt - hier bewusst mit eingelesen/validiert (identische
# Reports koennten sich zwischen den drei SE16XXL-Exporten fachlich leicht
# unterscheiden, z.B. Stichtag), aber fuer Baustein B NICHT weiterverarbeitet
# (nur der C~/D~-Block ist neu und wird tatsaechlich extrahiert).
OPEN_ORDERS_COLUMNS = [
    "A~Material",
    "A~Werk",
    "A~MatStatus",
    "A~ABC",
    "A~Planlief",
    "A~WE-BearbZt",
    "A~GesWZeit",
    "A~BeschArt",
    "A~Meldebest",
    "A~SichBest",
    "A~Feste LGr",
    "A~RundWert",
    "A~Losf. Kst.",
    "A~NachfMat",
    "A~LiefGrad",
    "B~GesBestand",
    "B~PrsStrg",
    "B~/",
    "B~GLD-Preis",
    "B~StdPreis",
    "C~Mandant",
    "C~EinkBeleg",
    "C~Position",
    "C~LöschKz",
    "C~Herkunft",
    "C~Material",
    "C~Werk",
    "D~Mandant",
    "D~EinkBeleg",
    "D~Position",
    "D~Einteilung",
    "D~LiefDat",
    "D~EintMeng.",
    "D~Vor.Menge",
    "D~WE-Menge",
    "D~Ausgegeben",
    "D~Banf",
    "D~Best.Datum",
]

MATERIAL_COLUMN = "A~Material"
EINKBELEG_COLUMN = "C~EinkBeleg"
POSITION_COLUMN = "C~Position"
EINTEILUNG_COLUMN = "D~Einteilung"
LOESCHKZ_COLUMN = "C~LöschKz"
LIEFDAT_COLUMN = "D~LiefDat"
EINTMENGE_COLUMN = "D~EintMeng."
WE_MENGE_COLUMN = "D~WE-Menge"

# Felder, die fuer Baustein B tatsaechlich benoetigt werden. Mapping:
# SAP-Rohspalte -> sprechender Zielspaltenname im Ergebnis-DataFrame.
_NUMERIC_FIELDS = {
    EINTMENGE_COLUMN: "Eingeteilte_Menge",
    WE_MENGE_COLUMN: "WE_Menge",
}

_TEXT_FIELDS = {
    EINKBELEG_COLUMN: "Einkaufsbeleg",
    POSITION_COLUMN: "Bestellposition",
    EINTEILUNG_COLUMN: "Einteilungszeile",
}


def _parse_german_number(value: object) -> float:
    """Wandelt einen Wert in float um - robust gegenueber bereits numerischen
    Werten und Text im deutschen Zahlenformat ('1.003' = 1003.0, '1,003' waere
    1.003 - siehe Hinweis unten). Analog zur gleichnamigen Funktion in
    dispo_abcxyz_loader.py/mver_loader.py/se16xxl_loader.py (bewusst
    dupliziert statt in einer gemeinsamen Util-Datei zusammengefasst - siehe
    Begruendung in se16xxl_loader.py).

    Hinweis Mengenfelder (D~EintMeng., D~WE-Menge): SAP liefert diese im
    realen Export als Text mit Punkt als Tausendertrennzeichen (z.B.
    '1.003' = 1003 Stueck, nicht 1,003). Gleiche deutsche Zahlenformat-
    Konvention wie in den uebrigen Loadern (Punkt=Tausender, Komma=Dezimal).
    """
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return float("nan")

    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return float("nan")


def _validate_open_orders_columns(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft, ob alle erwarteten Spalten im Open_Orders-Export vorhanden sind."""
    missing = [col for col in OPEN_ORDERS_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"SE16XXL-Open_Orders-Export '{file_path.name}' fehlen erwartete "
            f"Spalten: {missing}\nVorhandene Spalten: {list(df.columns)}"
        )


def full_open_orders_export_exists() -> bool:
    """Prueft, ob der SE16XXL-Open_Orders-Export vorliegt."""
    return config.SE16XXL_OPEN_ORDERS_EXPORT_PATH.exists()


def load_full_open_orders_export() -> pd.DataFrame:
    """Liest den SE16XXL-Open_Orders-Export ein und bereitet die fuer
    Baustein B benoetigten Felder auf.

    WICHTIG: Im Unterschied zu load_full_dispo_abcxyz_export() liefert diese
    Funktion KEIN eindeutig auf Material aggregiertes DataFrame, sondern eine
    Zeile JE LIEFERPLAN-EINTEILUNG (eine Bestellposition kann mehrere
    Einteilungen/Teillieferungen haben, siehe Modul-Docstring). Ein Material
    kann daher mehrfach vorkommen - das ist hier KEIN Duplikat-Fehler,
    sondern fachlich korrekt (mehrere offene Liefertermine desselben
    Materials).

    Returns:
        DataFrame mit Spalte 'Material' (18-stellig normalisiert),
        'Einkaufsbeleg', 'Bestellposition', 'Einteilungszeile' (Text/str,
        unveraendert aus SAP), 'Liefertermin' (datetime), 'Eingeteilte_Menge'/
        'WE_Menge' (float, aus _NUMERIC_FIELDS), 'Offene_Menge' (float,
        siehe berechne_offene_menge()) und 'Storniert' (bool, aus
        C~LoeschKz).

    Raises:
        FileNotFoundError: wenn config.SE16XXL_OPEN_ORDERS_EXPORT_PATH nicht
            existiert.
        ValueError: wenn erwartete Spalten fehlen.
    """
    file_path = config.SE16XXL_OPEN_ORDERS_EXPORT_PATH
    if not file_path.exists():
        raise FileNotFoundError(
            f"SE16XXL-Open_Orders-Export nicht gefunden: {file_path}\n"
            f"Bitte Export dort ablegen (Dateiname: "
            f"'{config.SE16XXL_OPEN_ORDERS_EXPORT_FILENAME}')."
        )

    df = pd.read_excel(file_path, dtype={MATERIAL_COLUMN: str})
    _validate_open_orders_columns(df, file_path)

    result = pd.DataFrame()
    # Normalisierung auf die SAP-native 18-stellige Form (siehe config.py),
    # analog zu den uebrigen Loadern - macht den spaeteren Merge mit
    # ZMLAG/MVER/Inventory/Dispo_ABCXYZ robust.
    result["Material"] = df[MATERIAL_COLUMN].apply(config.normalize_material_number)

    for sap_col, target_col in _TEXT_FIELDS.items():
        result[target_col] = df[sap_col].astype(str)

    # Loeschkennzeichen: SAP liefert i.d.R. 'X' bei gesetztem Kennzeichen,
    # sonst leer/NaN. Robust gegenueber beiden Faellen statt exaktem
    # String-Vergleich.
    result["Storniert"] = df[LOESCHKZ_COLUMN].notna() & (
        df[LOESCHKZ_COLUMN].astype(str).str.strip() != ""
    )

    result["Liefertermin"] = pd.to_datetime(
        df[LIEFDAT_COLUMN], format="%d.%m.%Y", errors="coerce"
    )

    for sap_col, target_col in _NUMERIC_FIELDS.items():
        result[target_col] = df[sap_col].apply(_parse_german_number)

    return result


def berechne_offene_menge(df: pd.DataFrame) -> pd.DataFrame:
    """Berechnet die offene Menge je Lieferplan-Einteilung:

        Offene_Menge = Eingeteilte_Menge - WE_Menge

    Werte <= 0 (voll geliefert oder Overdelivery) werden auf 0 gekappt statt
    als negativer Zugang gefuehrt zu werden (Entscheidung Claude, 01.07.2026,
    siehe Modul-Docstring: Zeile bleibt sichtbar, traegt aber korrekt mit 0
    zur Summe bei - kein Herausfiltern, kein negativer "Zugang").

    Args:
        df: DataFrame mit Spalten 'Eingeteilte_Menge', 'WE_Menge' (siehe
            load_full_open_orders_export()).

    Returns:
        Kopie von df, erweitert um 'Offene_Menge' (float, >= 0). NaN, falls
        einer der beiden Bestandteile fehlt (kein stillschweigendes
        Auffuellen mit 0 - eine fehlende Menge ist ein Datenproblem, keine
        volle Lieferung).
    """
    result = df.copy()
    offene_menge = result["Eingeteilte_Menge"] - result["WE_Menge"]
    result["Offene_Menge"] = offene_menge.clip(lower=0)
    # NaN bleibt NaN (clip() aendert NaN nicht) - fehlende Eingangsdaten
    # werden nicht stillschweigend als "voll geliefert" (0) interpretiert.
    return result


def build_open_orders_table(df: pd.DataFrame) -> pd.DataFrame:
    """Komplettlauf der Offene-Menge-Berechnung: laedt NICHT selbst (siehe
    load_full_open_orders_export()), sondern erwartet das bereits geladene
    DataFrame und wendet berechne_offene_menge() an.

    Dies ist die Hauptfunktion, die main.py/main_single.py fuer Baustein B
    aufrufen werden.

    Args:
        df: Ergebnis von load_full_open_orders_export().

    Returns:
        DataFrame mit allen Spalten aus df, erweitert um 'Offene_Menge'.
    """
    return berechne_offene_menge(df)


def aggregate_offene_menge_je_material(
    df: pd.DataFrame, *, exclude_storniert: bool = True
) -> pd.DataFrame:
    """Fasst die offene Menge auf Material-Ebene zusammen (Summe ueber alle
    Einteilungen/Bestellpositionen eines Materials) - fuer die Reichweite-
    Berechnung (Baustein B, "Stock + bestaetigte Zugaenge") wird pro Material
    EIN Gesamtwert der offenen Menge benoetigt, nicht die Einteilungs-Ebene.

    Args:
        df: Ergebnis von build_open_orders_table() (muss 'Offene_Menge'
            enthalten).
        exclude_storniert: Wenn True (Default), werden stornierte Positionen
            (Storniert == True) vor der Summierung herausgefiltert - eine
            stornierte Einteilung ist kein realer kuenftiger Zugang. Wenn
            False, werden stornierte Positionen mitgezaehlt (z.B. fuer eine
            Uebersicht "alle offenen Positionen inkl. stornierter, zur
            manuellen Pruefung").

    Returns:
        DataFrame mit Spalten 'Material', 'Offene_Menge_Summe' (float),
        'Anzahl_Einteilungen' (int), 'Naechster_Liefertermin' (datetime,
        fruehester Liefertermin unter den beruecksichtigten Einteilungen).
    """
    working = df[~df["Storniert"]] if exclude_storniert else df

    aggregated = (
        working.groupby("Material")
        .agg(
            Offene_Menge_Summe=("Offene_Menge", "sum"),
            Anzahl_Einteilungen=("Offene_Menge", "size"),
            Naechster_Liefertermin=("Liefertermin", "min"),
        )
        .reset_index()
    )
    return aggregated


def filter_by_materials(df: pd.DataFrame, materials: set) -> pd.DataFrame:
    """Filtert die (technologieuebergreifend geladene) Open-Orders-Tabelle
    auf die Materialien EINER Technologie - analog zum bestehenden
    Filterprinzip in dispo_abcxyz_loader.py (Dispo_ABCXYZ ist wie MVER,
    SE16XXL-Inventory und Open_Orders ein technologieuebergreifender
    Gesamtexport, der EINMAL fuer den gesamten Lauf gelesen und
    anschliessend je Technologie gefiltert wird).

    Args:
        df: Ergebnis von load_full_open_orders_export() (ggf. nach
            build_open_orders_table()/aggregate_offene_menge_je_material()).
        materials: Menge der Materialnummern (18-stellig normalisiert) der
            aktuell verarbeiteten Technologie, z.B. set(zmlag_df['Material']).

    Returns:
        Gefilterte Kopie von df, nur Zeilen mit Material in materials.
    """
    return df[df["Material"].isin(materials)].copy()

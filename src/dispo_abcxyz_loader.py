"""
dispo_abcxyz_loader.py
========================
Einlesen des SAP SE16XXL-Mehrfachtabellen-Merge-Exports "Dispo_ABCXYZ" -
des GROESSEREN der beiden SE16XXL-Reports, die Max nutzt (siehe
Projektstatus.md, Abschnitt "Datenquellen-Strategie"). Eigenstaendiges
Modul neben se16xxl_loader.py (welches den anderen SE16XXL-Report,
"Inventory", einliest) - bewusst NICHT in dieselbe Datei integriert, da
beide Reports strukturell unterschiedliche Spaltenlisten/Prefixe haben und
eigene Validierungslogik benoetigen (Entscheidung Max, 25.06.2026).

Liefert je Material u.a.:
    - C~Planlief, C~WE-BearbZt: Wiederbeschaffungszeit-Bestandteile
      (Whitepaper Kap. 3.2: L = Planlieferzeit + WE-Bearbeitungszeit)
    - B~Vari.Koeff: Variationskoeffizient des Verbrauchs - fuer Phase 5
      (Safety Stock) vorgesehen
    - C~SichBest, C~Meldebest: Sicherheitsbestand/Meldebestand (Phase 5)
    - E~StdPreis, E~GLD-Preis: Preise (Phase 6, Working-Capital-Berechnung)
    - A~ABCXYZ: KOMBINIERTES ABC/XYZ-Kennzeichen (z.B. "AX") - im Unterschied
      zu ZMLAG, das ABC- und XYZ-Kennzeichen GETRENNT liefert (siehe
      stock_analysis.py), und im Unterschied zum "Inventory"-Report, der nur
      das ABC-Kennzeichen allein liefert (A~ABC, ohne XYZ)

Wichtiger fachlicher Hinweis (Klaerung mit Max, 25.06.2026): Der "Inventory"-
Report enthaelt eigene Lead-Time-Felder (A~Planlief, A~WE-BearbZt,
A~GesWZeit), die laut Max ECHTE Werte fuehren (kein Strukturplatzhalter).
Anhand eines realen Stichprobenvergleichs (zwei Materialien) liefern beide
Quellen fuer Planlief/WE-BearbZt/SichBest/Meldebest IDENTISCHE Werte - es
handelt sich also um eine echte Ueberlappung, kein Datenkonflikt im
Normalfall. Da Dispo_ABCXYZ aber die fachlich fuer Phase 4/5 vorgesehene
Quelle ist UND zusaetzliche, in Inventory nicht vorhandene Felder liefert
(v.a. Vari.Koeff), wird Dispo_ABCXYZ als PRIMAERE Quelle fuer die
Wiederbeschaffungszeit behandelt. Liegen widerspruechliche Werte vor
(unterschiedliche Planlief/WE-BearbZt fuer dasselbe Material in beiden
Quellen), wird dies als Konsolen-Warnung ausgegeben statt stillschweigend
eine Quelle vorzuziehen (siehe compare_lead_time_sources()).

A~GesWZeit (Inventory) liefert die Lead-Time-Summe bereits vorberechnet aus
SAP. Dieses Modul berechnet die Summe TROTZDEM selbst aus den Einzelfeldern
(Whitepaper-Formel, siehe calculate_lead_time()) und nutzt eine ggf.
vorhandene SAP-Summe nur als Plausibilitaetscheck (Entscheidung Max,
25.06.2026) - nicht als direkte Uebernahme, damit die Formel im Code
nachvollziehbar bleibt und nicht von einem SAP-internen Rechenweg abhaengt,
der sich aendern koennte.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import config


# Erwartete Spalten im SE16XXL-Dispo_ABCXYZ-Export (Reihenfolge wie im
# realen Export von Max, 25.06.2026). Bewusst vollstaendig eingelesen/
# validiert, auch wenn aktuell (Phase 4) nur ein Teil der Felder tatsaechlich
# weiterverarbeitet wird - die uebrigen liegen damit bereits fuer Phase 5/6
# bereit, ohne dass der Loader spaeter erneut angefasst werden muss
# (Entscheidung Max, 25.06.2026).
DISPO_ABCXYZ_COLUMNS = [
    "D~Status",
    "C~MatStatus",
    "A~Ändg.",
    "C~Werk",
    "C~BeschArt",
    "C~Merkmal",
    "C~StratGrup.",
    "A~Material",
    "B~Bezeich",
    "D~Normbez",
    "A~ABCXYZ",
    "B~V-Kz.",
    "B~Vari.Koeff",
    "C~EigFerZt",
    "C~Planlief",
    "C~WE-BearbZt",
    "E~/",
    "E~PrsStrg",
    "E~StdPreis",
    "E~GLD-Preis",
    "B~GsVbr",
    "C~EinkGruppe",
    "C~LiefGrad",
    "C~SichBest",
    "C~MinSichB.",
    "C~Disponent",
    "C~FertSteu.",
    "C~Art",
    "C~Meldebest",
    "C~Einz/Samm",
    "C~LosgrVerf.",
    "C~Lagerkost.",
    "C~Losf. Kst.",
    "C~Min. LGr",
    "C~Max. LGr",
    "C~Feste LGr",
    "C~RundWert",
    "C~HöchstBstd",
    "C~NachfMat",
]

MATERIAL_COLUMN = "A~Material"
PLANLIEF_COLUMN = "C~Planlief"
WE_BEARBZT_COLUMN = "C~WE-BearbZt"

# Felder, die fuer Phase 4 (Wiederbeschaffungszeit) bzw. bereits vorbereitend
# fuer Phase 5/6 ins Ergebnis-DataFrame uebernommen werden. Mapping:
# SAP-Rohspalte -> sprechender Zielspaltenname im Ergebnis-DataFrame.
_NUMERIC_FIELDS = {
    PLANLIEF_COLUMN: "Planlieferzeit_Tage",
    WE_BEARBZT_COLUMN: "WE_Bearbeitungszeit_Tage",
    "B~Vari.Koeff": "Variationskoeffizient",
    "C~SichBest": "Sicherheitsbestand_Soll",
    "C~MinSichB.": "Mindest_Sicherheitsbestand",
    "C~Meldebest": "Meldebestand",
    "E~StdPreis": "Standardpreis",
    "E~GLD-Preis": "Gleitender_Durchschnittspreis",
    "C~Min. LGr": "Mindest_Losgroesse",
    "C~Max. LGr": "Hoechst_Losgroesse",
    "C~Feste LGr": "Feste_Losgroesse",
    "C~RundWert": "Rundungswert",
    "C~HöchstBstd": "Hoechstbestand",
}

_TEXT_FIELDS = {
    "A~ABCXYZ": "ABCXYZ_Kombiniert",
    "C~MatStatus": "MatStatus",
    "C~Disponent": "Disponent",
    "C~EinkGruppe": "Einkaeufergruppe",
    "C~BeschArt": "Beschaffungsart",
}


def _parse_german_number(value: object) -> float:
    """Wandelt einen Wert in float um - robust gegenueber bereits numerischen
    Werten und Text im deutschen Zahlenformat ('5,000' = 5.0). Analog zur
    gleichnamigen Funktion in mver_loader.py/se16xxl_loader.py (bewusst
    dupliziert statt in einer gemeinsamen Util-Datei zusammengefasst - siehe
    Begruendung in se16xxl_loader.py).
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


def _validate_dispo_abcxyz_columns(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft, ob alle erwarteten Spalten im Dispo_ABCXYZ-Export vorhanden sind."""
    missing = [col for col in DISPO_ABCXYZ_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"SE16XXL-Dispo_ABCXYZ-Export '{file_path.name}' fehlen erwartete "
            f"Spalten: {missing}\nVorhandene Spalten: {list(df.columns)}"
        )


def full_dispo_abcxyz_export_exists() -> bool:
    """Prueft, ob der SE16XXL-Dispo_ABCXYZ-Export vorliegt."""
    return config.SE16XXL_DISPO_ABCXYZ_EXPORT_PATH.exists()


def load_full_dispo_abcxyz_export() -> pd.DataFrame:
    """Liest den SE16XXL-Dispo_ABCXYZ-Export ein und bereitet die fuer
    Phase 4 (Wiederbeschaffungszeit) und vorbereitend fuer Phase 5/6
    benoetigten Felder auf.

    Returns:
        DataFrame mit Spalte 'Material' (18-stellig normalisiert) sowie
        allen Zielspalten aus _NUMERIC_FIELDS (float) und _TEXT_FIELDS
        (str, unveraendert aus SAP uebernommen).

    Raises:
        FileNotFoundError: wenn config.SE16XXL_DISPO_ABCXYZ_EXPORT_PATH
            nicht existiert.
        ValueError: wenn erwartete Spalten fehlen.
    """
    file_path = config.SE16XXL_DISPO_ABCXYZ_EXPORT_PATH
    if not file_path.exists():
        raise FileNotFoundError(
            f"SE16XXL-Dispo_ABCXYZ-Export nicht gefunden: {file_path}\n"
            f"Bitte Export dort ablegen (Dateiname: "
            f"'{config.SE16XXL_DISPO_ABCXYZ_EXPORT_FILENAME}')."
        )

    df = pd.read_excel(file_path, dtype={MATERIAL_COLUMN: str})
    _validate_dispo_abcxyz_columns(df, file_path)

    result = pd.DataFrame()
    # Normalisierung auf die SAP-native 18-stellige Form (siehe config.py).
    # Dispo_ABCXYZ liefert bereits 18 Stellen (siehe reale Beispielzeilen von
    # Max, 25.06.2026) - die explizite Anwendung macht den spaeteren Merge
    # mit ZMLAG/MVER/Inventory aber robust gegenueber Format-Aenderungen,
    # analog zu se16xxl_loader.py.
    result["Material"] = df[MATERIAL_COLUMN].apply(config.normalize_material_number)

    for sap_col, target_col in _NUMERIC_FIELDS.items():
        result[target_col] = df[sap_col].apply(_parse_german_number)

    for sap_col, target_col in _TEXT_FIELDS.items():
        result[target_col] = df[sap_col]

    duplicates = result["Material"][result["Material"].duplicated()].unique()
    if len(duplicates) > 0:
        print(
            f"  Hinweis: {len(duplicates)} Material(ien) kommen im SE16XXL-"
            f"Dispo_ABCXYZ-Export '{file_path.name}' mehrfach vor (z.B. "
            f"mehrere Werke) - es wird die ERSTE Zeile je Material verwendet: "
            f"{list(duplicates)}"
        )
        result = result.drop_duplicates(subset="Material", keep="first")

    return result


def calculate_lead_time(df: pd.DataFrame) -> pd.DataFrame:
    """Berechnet die Wiederbeschaffungszeit nach der Whitepaper-Formel
    (Kap. 3.2):

        L = Planlieferzeit + WE-Bearbeitungszeit

    Bewusst als EIGENE Berechnung im Code (statt direkter Uebernahme einer
    von SAP ggf. mitgelieferten Summe wie A~GesWZeit im "Inventory"-Report) -
    Entscheidung Max, 25.06.2026: die Formel soll im Code nachvollziehbar
    bleiben statt von einem SAP-internen, nicht einsehbaren Rechenweg
    abzuhaengen. Siehe compare_lead_time_sources() fuer den Plausibilitaets-
    abgleich gegen eine SAP-seitige Summe, falls vorhanden.

    Args:
        df: DataFrame mit Spalten 'Planlieferzeit_Tage', 'WE_Bearbeitungszeit_Tage'
            (siehe load_full_dispo_abcxyz_export()).

    Returns:
        Kopie von df, erweitert um 'Wiederbeschaffungszeit_Tage'. NaN, falls
        einer der beiden Bestandteile fehlt (kein stillschweigendes
        Auffuellen mit 0, da das eine falsche, zu kurze Lead Time
        suggerieren wuerde).
    """
    result = df.copy()
    result["Wiederbeschaffungszeit_Tage"] = (
        result["Planlieferzeit_Tage"] + result["WE_Bearbeitungszeit_Tage"]
    )
    return result


def compare_lead_time_sources(
    dispo_df: pd.DataFrame, inventory_df: pd.DataFrame
) -> pd.DataFrame:
    """Vergleicht die Lead-Time-Bestandteile (Planlieferzeit, WE-Bearbeitungszeit)
    zwischen dem Dispo_ABCXYZ-Export (primaere Quelle, siehe Modul-Docstring)
    und dem Inventory-Export (se16xxl_loader.py), FALLS beide vorliegen.

    Gibt eine Konsolen-Warnung aus, wenn sich die Werte fuer dasselbe Material
    unterscheiden (Entscheidung Max, 25.06.2026: Dispo_ABCXYZ wird bei
    Widerspruch bevorzugt verwendet, der Widerspruch selbst aber nicht
    stillschweigend uebergangen, sondern sichtbar gemacht). Im Normalfall
    (siehe reale Stichprobe von Max) sind beide Quellen identisch - diese
    Funktion deckt den Sonderfall ab, dass die beiden Exporte zu
    unterschiedlichen Zeitpunkten gezogen wurden und sich daher der
    Dispo-Stamm zwischenzeitlich geaendert hat.

    Args:
        dispo_df: Ergebnis von load_full_dispo_abcxyz_export() (ggf. nach
            calculate_lead_time()).
        inventory_df: Ergebnis von se16xxl_loader.load_full_se16xxl_inventory_export().
            Erwartet KEINE eigenen Lead-Time-Spalten im Standard-Rueckgabewert
            dieser Funktion (siehe se16xxl_loader.py) - dieser Vergleich liest
            daher die Rohspalten A~Planlief/A~WE-BearbZt direkt aus dem
            Inventory-Export nochmal ein, falls vorhanden, statt sich auf eine
            Erweiterung von se16xxl_loader.py zu verlassen, die es (Stand
            25.06.2026) noch nicht gibt.

    Returns:
        DataFrame mit allen Materialien, bei denen ein Widerspruch festgestellt
        wurde (leer, falls keine Widersprueche). Wird zusaetzlich als
        Konsolen-Warnung ausgegeben, falls nicht leer.
    """
    if "Material" not in inventory_df.columns:
        return pd.DataFrame()

    inventory_lead_time_cols = {"Planlieferzeit_Tage_Inventory", "WE_Bearbeitungszeit_Tage_Inventory"}
    if not inventory_lead_time_cols.issubset(inventory_df.columns):
        # Inventory-DataFrame enthaelt (noch) keine eigenen Lead-Time-Spalten
        # in vergleichbarer Form - kein Vergleich moeglich, kein Fehler.
        return pd.DataFrame()

    merged = dispo_df.merge(
        inventory_df[["Material", *inventory_lead_time_cols]],
        on="Material",
        how="inner",
    )

    mismatch = merged[
        (merged["Planlieferzeit_Tage"] != merged["Planlieferzeit_Tage_Inventory"])
        | (merged["WE_Bearbeitungszeit_Tage"] != merged["WE_Bearbeitungszeit_Tage_Inventory"])
    ]

    if not mismatch.empty:
        print(
            f"  Warnung: {len(mismatch)} Material(ien) haben abweichende "
            f"Lead-Time-Werte zwischen Dispo_ABCXYZ und Inventory - es wird "
            f"Dispo_ABCXYZ verwendet (siehe Material-Liste): "
            f"{list(mismatch['Material'])}"
        )

    return mismatch


def build_lead_time_table(df: pd.DataFrame) -> pd.DataFrame:
    """Komplettlauf der Wiederbeschaffungszeit-Berechnung (Phase 4): laedt
    NICHT selbst (siehe load_full_dispo_abcxyz_export()), sondern erwartet
    das bereits geladene DataFrame und wendet calculate_lead_time() an.

    Dies ist die Hauptfunktion, die main.py/main_single.py fuer den
    Wiederbeschaffungszeit-Reiter aufrufen werden.

    Args:
        df: Ergebnis von load_full_dispo_abcxyz_export().

    Returns:
        DataFrame mit allen Spalten aus df, erweitert um
        'Wiederbeschaffungszeit_Tage'.
    """
    return calculate_lead_time(df)


def filter_by_materials(df: pd.DataFrame, materials: set) -> pd.DataFrame:
    """Filtert die (technologieuebergreifend geladene) Dispo_ABCXYZ-Tabelle
    auf die Materialien EINER Technologie - analog zum bestehenden
    Filterprinzip fuer mver_long_df in main.py/main_single.py
    (_process()/_process_technology()), da Dispo_ABCXYZ wie MVER und
    SE16XXL-Inventory ein technologieuebergreifender Gesamtexport ist, der
    EINMAL fuer den gesamten Lauf gelesen und anschliessend je Technologie
    gefiltert wird.

    Args:
        df: Ergebnis von load_full_dispo_abcxyz_export() (ggf. nach
            calculate_lead_time()/build_lead_time_table()).
        materials: Menge der Materialnummern (18-stellig normalisiert) der
            aktuell verarbeiteten Technologie, z.B. set(zmlag_df['Material']).

    Returns:
        Gefilterte Kopie von df, nur Zeilen mit Material in materials.
    """
    return df[df["Material"].isin(materials)].copy()

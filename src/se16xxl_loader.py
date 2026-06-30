"""
se16xxl_loader.py
===================
Einlesen von SAP SE16XXL-Mehrfachtabellen-Merge-Exporten als dritte,
eigenstaendige Datenquelle neben ZMLAG und MVER (siehe Projektstatus.md,
Abschnitt "Datenquellen-Strategie").

Aktuell unterstuetzt: der Report "Inventory" (zweiter von zwei SE16XXL-
Reports, die Max nutzt). Liefert je Material/Werk u.a.:
    - B~GesBestand: Ist-Bestandsmenge in Stueck (benoetigt fuer die
      Reichweitenberechnung in Phase 2, da ZMLAG nur Bestandswerte in Euro
      liefert, MVER-Verbrauch aber in Stueck vorliegt - siehe Klaerung mit
      Max, 24.06.2026)
    - A~MatStatus: Lebenszyklus-/Freigabestatus des Materials (F=Freigegeben,
      G/L/X=verschiedene Sperrarten, A=Laeuft aus, D=Abgekuendigt, etc. -
      vollstaendige Codeliste siehe _MATERIAL_STATUS_LABELS unten)

Der erste, groessere SE16XXL-Report ("Dispo_ABCXYZ", mit Planlieferzeit,
Meldebestand, Sicherheitsbestand, Preisen etc.) wird hier bewusst NOCH NICHT
eingelesen - relevant erst fuer Phase 4/5 (Wiederbeschaffungszeit,
Fruehwarnsystem). config.py definiert den Dateipfad dafuer bereits, damit
der Dateiname projektweit konsistent ist.

Wichtiger fachlicher Hinweis (siehe Klaerung mit Max, 24.06.2026):
A~MatStatus ist ein GENERELLER Lebenszyklus-Status des Materials (ein Wert
pro Material), NICHT die in den "Initiale Informationen" geforderte
Aufteilung der BESTANDSMENGE nach Bestandsarten (frei verwendbar / in
Q-Pruefung / gesperrt, mit jeweils eigener Teilmenge). Diese feinere
Aufteilung ist weiterhin ein offener Punkt (siehe Projektstatus.md).
A~MatStatus wird hier nur als Datenqualitaets-/Plausibilitaets-Hinweis
genutzt (z.B. Warnung bei Reichweitenberechnung fuer gesperrte Materialien).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import config


# Erwartete Spalten im SE16XXL-Inventory-Export (Reihenfolge wie im realen
# Export von Max, 24.06.2026). Nicht alle werden fuer Phase 2 benoetigt, aber
# vollstaendig validiert, damit Strukturaenderungen im Export frueh auffallen.
SE16XXL_INVENTORY_COLUMNS = [
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
]

# Spalten, die fuer Phase 2 tatsaechlich gebraucht werden (Bestandsmenge +
# Status). Die uebrigen werden eingelesen/validiert, aber im Rueckgabe-
# DataFrame nicht weiter verarbeitet - sie liegen fuer spaetere Phasen
# (Wiederbeschaffungszeit etc.) bereits vor, sobald sie benoetigt werden.
MATERIAL_COLUMN = "A~Material"
STOCK_QUANTITY_COLUMN = "B~GesBestand"
MATERIAL_STATUS_COLUMN = "A~MatStatus"
# Phase 4 (25.06.2026): bereits validierte, bisher aber ignorierte Rohspalten -
# siehe load_full_se16xxl_inventory_export() Docstring fuer den Verwendungszweck
# (Plausibilitaetsabgleich, NICHT primaere Lead-Time-Quelle).
LEAD_TIME_PLANLIEF_COLUMN = "A~Planlief"
LEAD_TIME_WE_BEARBZT_COLUMN = "A~WE-BearbZt"

# Vollstaendige Codeliste fuer A~MatStatus (von Max bestaetigt, 24.06.2026).
# Wird genutzt, um im Report einen lesbaren Klartext statt nur des Kuerzels
# anzuzeigen, und um "kritische" Stati (gesperrt/abgekuendigt) zu erkennen.
MATERIAL_STATUS_LABELS = {
    "A": "Laeuft aus",
    "D": "Abgekuendigt",
    "E": "Ersatzteil",
    "F": "Freigegeben",
    "G": "Gesperrt",
    "K": "Konstruktion",
    "L": "Gesperrt mit Loeschkennz.",
    "M": "Messeexponate (nur CO)",
    "P": "Nullserie/Prototyp",
    "U": "Material neu unvollstaendig",
    "X": "Gesperrt tech. Aenderung",
    "Y": "Klaerung noetig - Einkauf",
}

# Stati, bei denen eine berechnete Reichweite mit Vorsicht zu interpretieren
# ist (Material nicht regulaer verfuegbar/aktiv disponiert). Wird fuer eine
# Konsolen-Warnung in Phase 2 genutzt, kein Hard-Fail.
CRITICAL_MATERIAL_STATUSES = {"G", "L", "X", "D", "A"}


def _parse_german_number(value: object) -> float:
    """Wandelt einen Wert in float um - robust gegenueber bereits numerischen
    Werten und Text im deutschen Zahlenformat ('5,000' = 5.0). Analog zur
    gleichnamigen Funktion in mver_loader.py (bewusst dupliziert, da beide
    Module unabhaengig voneinander nutzbar bleiben sollen - kein Koppeln
    ueber eine gemeinsame Util-Datei fuer nur eine kleine Hilfsfunktion).
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


def _validate_se16xxl_inventory_columns(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft, ob alle erwarteten Spalten im SE16XXL-Inventory-Export vorhanden sind."""
    missing = [col for col in SE16XXL_INVENTORY_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"SE16XXL-Inventory-Export '{file_path.name}' fehlen erwartete "
            f"Spalten: {missing}\nVorhandene Spalten: {list(df.columns)}"
        )


def full_se16xxl_inventory_export_exists() -> bool:
    """Prueft, ob der SE16XXL-Inventory-Export vorliegt."""
    return config.SE16XXL_INVENTORY_EXPORT_PATH.exists()


def load_full_se16xxl_inventory_export() -> pd.DataFrame:
    """Liest den SE16XXL-Inventory-Export ein und bereitet die fuer Phase 2
    benoetigten Felder auf.

    Returns:
        DataFrame mit Spalten:
            - Material (str, mit fuehrenden Nullen)
            - Bestandsmenge (float, Stueck) - aus B~GesBestand
            - MatStatus (str, Rohcode) - aus A~MatStatus
            - MatStatus_Bezeichnung (str, Klartext via MATERIAL_STATUS_LABELS)
            - MatStatus_Kritisch (bool) - True bei gesperrten/auslaufenden Materialien
            - Planlieferzeit_Tage_Inventory (float) - aus A~Planlief; NUR fuer den
              Plausibilitaetsabgleich gegen die gleichnamigen Felder aus dem
              SE16XXL-Dispo_ABCXYZ-Export (siehe dispo_abcxyz_loader.
              compare_lead_time_sources(), Phase 4, 25.06.2026). Dispo_ABCXYZ
              bleibt die PRIMAERE Quelle fuer die Wiederbeschaffungszeit -
              diese beiden Spalten werden NICHT fuer die eigentliche
              Lead-Time-Berechnung verwendet.
            - WE_Bearbeitungszeit_Tage_Inventory (float) - aus A~WE-BearbZt,
              analog zu Planlieferzeit_Tage_Inventory.

    Raises:
        FileNotFoundError: wenn config.SE16XXL_INVENTORY_EXPORT_PATH nicht existiert.
        ValueError: wenn erwartete Spalten fehlen.
    """
    file_path = config.SE16XXL_INVENTORY_EXPORT_PATH
    if not file_path.exists():
        raise FileNotFoundError(
            f"SE16XXL-Inventory-Export nicht gefunden: {file_path}\n"
            f"Bitte Export dort ablegen (Dateiname: "
            f"'{config.SE16XXL_INVENTORY_EXPORT_FILENAME}')."
        )

    df = pd.read_excel(file_path, dtype={MATERIAL_COLUMN: str})
    _validate_se16xxl_inventory_columns(df, file_path)

    result = pd.DataFrame()
    # Normalisierung auf die SAP-native 18-stellige Form (siehe config.py,
    # Abschnitt "Materialnummern-Normalisierung"). SE16XXL liefert bereits
    # 18 Stellen, daher hier im Normalfall ein No-Op - die explizite
    # Anwendung macht den Merge mit ZMLAG (siehe coverage_analysis.py) aber
    # robust gegenueber zukuenftigen Format-Aenderungen.
    result["Material"] = df[MATERIAL_COLUMN].apply(config.normalize_material_number)
    result["Bestandsmenge"] = df[STOCK_QUANTITY_COLUMN].apply(_parse_german_number)
    result["MatStatus"] = df[MATERIAL_STATUS_COLUMN]
    result["MatStatus_Bezeichnung"] = result["MatStatus"].map(MATERIAL_STATUS_LABELS).fillna(
        "Unbekannter Status"
    )
    result["MatStatus_Kritisch"] = result["MatStatus"].isin(CRITICAL_MATERIAL_STATUSES)

    # Phase 4 (25.06.2026): A~Planlief/A~WE-BearbZt sind bereits seit Phase 2
    # in SE16XXL_INVENTORY_COLUMNS validiert, wurden aber bisher nicht ins
    # Ergebnis-DataFrame uebernommen. Dienen hier AUSSCHLIESSLICH dem
    # Plausibilitaetsabgleich gegen den Dispo_ABCXYZ-Export (siehe
    # dispo_abcxyz_loader.compare_lead_time_sources()) - nicht der
    # eigentlichen Lead-Time-Berechnung, da Dispo_ABCXYZ dafuer die fachlich
    # vorgesehene, primaere Quelle ist (Entscheidung Max, 25.06.2026).
    result["Planlieferzeit_Tage_Inventory"] = df[LEAD_TIME_PLANLIEF_COLUMN].apply(
        _parse_german_number
    )
    result["WE_Bearbeitungszeit_Tage_Inventory"] = df[LEAD_TIME_WE_BEARBZT_COLUMN].apply(
        _parse_german_number
    )

    duplicates = result["Material"][result["Material"].duplicated()].unique()
    if len(duplicates) > 0:
        print(
            f"  Hinweis: {len(duplicates)} Material(ien) kommen im SE16XXL-"
            f"Inventory-Export '{file_path.name}' mehrfach vor (z.B. mehrere "
            f"Werke) - es wird die ERSTE Zeile je Material verwendet: "
            f"{list(duplicates)}"
        )
        result = result.drop_duplicates(subset="Material", keep="first")

    return result

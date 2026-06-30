"""
stock_analysis.py
==================
Berechnungslogik fuer Bestand und Bestandsentwicklung auf Material-Nr.-Ebene.

Aktueller Umfang (Basis: ZMLAG-Export):
    - Bestandswert-Entwicklung vom ersten Stichtag bis zum tagesaktuellen
      "Laufenden Wert" (siehe fachlicher Hinweis unten)
    - Veraenderung absolut und prozentual zwischen Start und tagesaktuellem Endpunkt
    - Einfache Risikoindikation auf Basis ABC/XYZ-Klassifizierung

Fachlicher Hinweis zu "Laufender Wert" (Klaerung mit Max, 24.06.2026):
Die drei JJJJMM-Spalten sind Monats-Stichtage (vermutlich jeweils Ultimo).
"Laufender Wert" ist der Bestandswert zum tatsaechlichen Exportzeitpunkt
(z.B. 23.06.2026) und damit IMMER aktueller als der letzte Monats-Stichtag.
Fuer eine reale Bestandsbewertung ist "Laufender Wert" der wichtigste, weil
tagesaktuellste Datenpunkt. Er wird daher hier als "Bestandswert_Ende" der
Entwicklungsrechnung verwendet, NICHT der letzte Stichtag. Der letzte
Stichtag bleibt zusaetzlich als eigene Spalte ('Bestandswert_Letzter_Stichtag')
erhalten, um diese Zwischenstufe nicht zu verlieren.

Hinweis: Eine echte Trend- und Reichweitenanalyse (gleitende 36-Monats-Historie,
Reichweite/Reichweitentage, Sondereffekt-Erkennung) erfordert den MVER-Export
und wird in einem spaeteren Ausbauschritt (Phase 3 des Projektplans) ergaenzt.
Die hier berechneten Kennzahlen sind bewusst auf das beschraenkt, was aus
ZMLAG allein robust ableitbar ist - keine Annahmen ueber nicht vorhandene Daten.
"""

from __future__ import annotations

import pandas as pd

from data_loader import get_period_columns


def calculate_stock_development(df: pd.DataFrame) -> pd.DataFrame:
    """Berechnet die Bestandswert-Entwicklung vom fruehesten verfuegbaren
    Stichtag bis zum tagesaktuellen "Laufenden Wert" je Material.

    Args:
        df: geladener ZMLAG-DataFrame (siehe data_loader.load_zmlag_export),
            muss die Spalte 'Laufender Wert' enthalten.

    Returns:
        Kopie von df, erweitert um:
            - 'Bestandswert_Start': Wert am fruehesten verfuegbaren Stichtag
            - 'Bestandswert_Letzter_Stichtag': Wert am spaetesten Monats-
              Stichtag (Zwischenstufe, NICHT der Endpunkt der Entwicklung -
              siehe fachlicher Hinweis im Modul-Docstring)
            - 'Bestandswert_Ende': tagesaktueller Wert (= 'Laufender Wert'
              aus dem ZMLAG-Export, zum Exportzeitpunkt)
            - 'Veraenderung_Absolut': Ende - Start
            - 'Veraenderung_Prozent': (Ende - Start) / Start
              als Dezimalwert (z.B. -0.2907 fuer -29,07 %), damit das
              Excel-Zellformat '0.00%' im Report korrekt greift.
              NaN, falls Start = 0, um Division durch 0 zu vermeiden.
    """
    result = df.copy()
    period_columns = get_period_columns(result)

    if len(period_columns) < 1:
        raise ValueError(
            f"Fuer eine Entwicklungsberechnung wird mindestens 1 Stichtags-"
            f"Spalte als Startpunkt benoetigt, gefunden: {period_columns}"
        )

    if "Laufender Wert" not in result.columns:
        raise ValueError(
            "Spalte 'Laufender Wert' fehlt im DataFrame - sie wird als "
            "tagesaktueller Endpunkt der Bestandswert-Entwicklung benoetigt."
        )

    earliest_col = period_columns[0]
    latest_period_col = period_columns[-1]

    result["Bestandswert_Start"] = result[earliest_col]
    result["Bestandswert_Letzter_Stichtag"] = result[latest_period_col]
    result["Bestandswert_Ende"] = result["Laufender Wert"]
    result["Veraenderung_Absolut"] = (
        result["Bestandswert_Ende"] - result["Bestandswert_Start"]
    )

    result["Veraenderung_Prozent"] = result.apply(
        lambda row: (
            (row["Veraenderung_Absolut"] / row["Bestandswert_Start"])
            if row["Bestandswert_Start"] not in (0, None) and pd.notna(row["Bestandswert_Start"])
            else float("nan")
        ),
        axis=1,
    )

    return result


def classify_risk_by_abc_xyz(df: pd.DataFrame) -> pd.DataFrame:
    """Ordnet jedem Material eine grobe Risiko-Prioritaet auf Basis der
    ABCXYZ-Klassifizierung zu.

    Diese Einstufung ist ein einfacher, transparenter Platzhalter fuer die
    vollwertige Risikoklassifizierung in Phase 5 (Fruehwarnsystem), die
    zusaetzlich Reichweite, Wiederbeschaffungszeit und reale Zugaenge
    einbezieht. Hier dient sie nur der ersten Priorisierung im Bestandsreport.

    Logik:
        X-Materialien (stabiler Verbrauch) -> geringeres Risiko
        Z-Materialien (volatiler Verbrauch) -> hoeheres Risiko
        A-Materialien (hoher Wertanteil) -> Aufmerksamkeit unabhaengig von XYZ

    Args:
        df: DataFrame mit Spalte 'ABCXYZ-Kennzeichen'

    Returns:
        Kopie von df, erweitert um 'Risiko_Einstufung' (Text: 'Niedrig',
        'Mittel', 'Hoch', oder 'Unbekannt' falls Kennzeichen fehlt/unerwartet).
    """
    result = df.copy()

    def _classify(code: object) -> str:
        if not isinstance(code, str) or len(code) != 2:
            return "Unbekannt"

        abc, xyz = code[0], code[1]

        if xyz == "Z":
            return "Hoch"
        if abc == "A" and xyz == "Y":
            return "Mittel"
        if xyz == "X":
            return "Niedrig"
        return "Mittel"

    result["Risiko_Einstufung"] = result["ABCXYZ-Kennzeichen"].apply(_classify)

    return result

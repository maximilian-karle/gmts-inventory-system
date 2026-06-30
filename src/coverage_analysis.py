"""
coverage_analysis.py
======================
Berechnung der Reichweite (Bestandsreichweite in Monaten) je Material -
Phase 2 des Projektplans.

Fachlicher Hintergrund: Reichweite = Bestand / Verbrauchsrate. Damit diese
Rechnung sinnvoll ist, MUESSEN Bestand und Verbrauch in derselben Einheit
vorliegen. Da ZMLAG nur Bestandswerte in Euro liefert (siehe stock_analysis.py),
MVER-Verbrauch aber in Stueck, wird fuer die Reichweite eine DRITTE Quelle
benoetigt: die Bestandsmenge in Stueck aus dem SE16XXL-Inventory-Export
(siehe se16xxl_loader.py, B~GesBestand). Diese Architekturentscheidung wurde
mit Max am 24.06.2026 final geklaert.

Umfang dieses Moduls (Phase 2, Stand 24.06.2026):
    - Reichweite "Stock-only": aktueller Bestand (Stueck) / Ø-Monatsverbrauch
      (12 Monate rollierend, siehe mver_loader.calculate_average_monthly_
      consumption_12m)
    - Frei editierbares Planungsfeld: optionaler manueller Override-Wert fuer
      den Verbrauch je Material, der den historischen Durchschnitt ersetzt
      (z.B. bei bekanntem zukuenftigem Bedarfssprung). Wird NICHT automatisch
      befuellt - leer = historischer Durchschnitt wird verwendet.

BEWUSST NICHT in Phase 2 enthalten (siehe Projektstatus.md, Offene Punkte):
    - Reichweite "Stock + bestaetigte Zugaenge": die SAP-Quelle fuer offene
      Bestellungen/Lieferplaene ist noch nicht eingebunden. Dieser Punkt wird
      erst angegangen, nachdem Phase 2 mit der Stock-only-Variante
      abgeschlossen ist (Entscheidung Max, 24.06.2026).
"""

from __future__ import annotations

import pandas as pd


# Name der Spalte fuer das frei editierbare Planungsfeld im finalen DataFrame/
# Report. Bleibt standardmaessig leer (NaN) - der Anwender (z.B. Disponent)
# kann hier nach Bedarf manuell einen abweichenden Monatsverbrauch eintragen,
# der dann statt des historischen Durchschnitts fuer die Reichweite verwendet
# wird (siehe apply_planning_override()).
PLANNING_FIELD_COLUMN = "Planung_Verbrauch_Override"


def merge_stock_and_consumption(
    zmlag_df: pd.DataFrame,
    consumption_df: pd.DataFrame,
    stock_quantity_df: pd.DataFrame,
) -> pd.DataFrame:
    """Fuehrt die drei Datenquellen auf Material-Ebene zusammen:
        - zmlag_df: ZMLAG-Bestandsdaten (Materialstamm, Bestandswerte in Euro)
        - consumption_df: Ø-Monatsverbrauch aus MVER (siehe mver_loader)
        - stock_quantity_df: Bestandsmenge in Stueck aus SE16XXL (siehe se16xxl_loader)

    Es wird ein LEFT JOIN auf zmlag_df durchgefuehrt - jedes Material aus dem
    aktuellen ZMLAG-Lauf bleibt erhalten, auch wenn fuer MVER oder SE16XXL
    (noch) keine Daten vorliegen (z.B. neues Material ohne Verbrauchshistorie).
    Fehlende Werte werden NICHT mit 0 aufgefuellt, sondern bleiben als NaN
    erhalten, damit eine fehlende Datengrundlage im Report sichtbar bleibt
    und nicht faelschlich als "kein Verbrauch"/"kein Bestand" interpretiert
    werden kann.

    Args:
        zmlag_df: DataFrame mit Spalte 'Material' (z.B. Ergebnis von
            stock_analysis.calculate_stock_development).
        consumption_df: DataFrame mit Spalten 'Material', 'Ø_Verbrauch_Fenster'
            (generischer Name unabhaengig von der gewaehlten Fenstergroesse,
            siehe mver_loader.calculate_average_monthly_consumption),
            'Anzahl_Monate_Verfuegbar'.
        stock_quantity_df: DataFrame mit Spalten 'Material', 'Bestandsmenge',
            'MatStatus', 'MatStatus_Bezeichnung', 'MatStatus_Kritisch'.

    Returns:
        Zusammengefuehrter DataFrame, eine Zeile pro Material aus zmlag_df.
    """
    merged = zmlag_df.merge(consumption_df, on="Material", how="left")
    merged = merged.merge(stock_quantity_df, on="Material", how="left")
    return merged


def apply_planning_override(df: pd.DataFrame) -> pd.DataFrame:
    """Stellt sicher, dass die Planungsfeld-Spalte (PLANNING_FIELD_COLUMN)
    existiert, und berechnet 'Verbrauch_Fuer_Reichweite' als den Wert, der
    tatsaechlich fuer die Reichweitenberechnung verwendet wird:
        - Planungsfeld gesetzt (nicht leer/NaN) -> Planungsfeld-Wert wird verwendet
        - Planungsfeld leer -> historischer Ø-Verbrauch (12 Monate) wird verwendet

    Das Planungsfeld wird, falls noch nicht vorhanden, leer (NaN) angelegt -
    es ist bewusst ein manuelles Eingabefeld im Excel-Report und wird NICHT
    automatisch durch den Loader befuellt.

    Args:
        df: DataFrame nach merge_stock_and_consumption(), muss 'Ø_Verbrauch_Fenster'
            enthalten.

    Returns:
        Kopie von df, erweitert um PLANNING_FIELD_COLUMN (falls nicht
        vorhanden) und 'Verbrauch_Fuer_Reichweite'.
    """
    result = df.copy()

    if PLANNING_FIELD_COLUMN not in result.columns:
        result[PLANNING_FIELD_COLUMN] = pd.NA

    override = pd.to_numeric(result[PLANNING_FIELD_COLUMN], errors="coerce")
    result["Verbrauch_Fuer_Reichweite"] = override.combine_first(result["Ø_Verbrauch_Fenster"])

    return result


def calculate_coverage_stock_only(df: pd.DataFrame) -> pd.DataFrame:
    """Berechnet die Reichweite (Stock-only) in Monaten je Material:

        Reichweite_Monate_StockOnly = Bestandsmenge / Verbrauch_Fuer_Reichweite

    Sonderfaelle:
        - Verbrauch_Fuer_Reichweite = 0 oder NaN -> Reichweite = NaN (Division
          durch 0 wird vermieden; ein Verbrauch von 0 wuerde sonst eine
          unendliche Reichweite suggerieren, was im Report irrefuehrend waere)
        - Bestandsmenge fehlt (NaN, z.B. kein SE16XXL-Datensatz fuer dieses
          Material) -> Reichweite = NaN

    Args:
        df: DataFrame nach apply_planning_override(), muss 'Bestandsmenge'
            und 'Verbrauch_Fuer_Reichweite' enthalten.

    Returns:
        Kopie von df, erweitert um 'Reichweite_Monate_StockOnly'.
    """
    result = df.copy()

    def _safe_divide(row: pd.Series) -> float:
        verbrauch = row["Verbrauch_Fuer_Reichweite"]
        bestand = row["Bestandsmenge"]
        if pd.isna(verbrauch) or verbrauch == 0 or pd.isna(bestand):
            return float("nan")
        return bestand / verbrauch

    result["Reichweite_Monate_StockOnly"] = result.apply(_safe_divide, axis=1)

    return result


def build_coverage_table(
    zmlag_df: pd.DataFrame,
    consumption_df: pd.DataFrame,
    stock_quantity_df: pd.DataFrame,
) -> pd.DataFrame:
    """Komplettlauf der Reichweitenberechnung (Phase 2, Stock-only): Merge der
    drei Quellen, Anwendung des Planungsfelds, Berechnung der Reichweite.

    Dies ist die Hauptfunktion, die main.py fuer den Reichweite-Reiter aufruft.

    Returns:
        DataFrame mit (u.a.) Material, Bestandsmenge, MatStatus_Bezeichnung,
        MatStatus_Kritisch, Ø_Verbrauch_12M, Anzahl_Monate_Verfuegbar,
        Planung_Verbrauch_Override, Verbrauch_Fuer_Reichweite,
        Reichweite_Monate_StockOnly.
    """
    merged = merge_stock_and_consumption(zmlag_df, consumption_df, stock_quantity_df)
    merged = apply_planning_override(merged)
    merged = calculate_coverage_stock_only(merged)
    return merged

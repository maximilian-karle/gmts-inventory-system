"""
executive_summary.py
======================
Executive-Summary-Tabelle - Phase 6 des Projektplans.

Fachlicher Hintergrund: Mit Phasen 1-5 sind sieben Detail-Reiter entstanden
(ZMLAG_Bestand, Reichweite, Trend_MVER, Trend_Details, Wiederbeschaffung,
Simulation_Verbrauch, Fruehwarnsystem, Forecast_MVER), jeder mit eigenem
Join/eigenen Pflichtquellen. Max' Rueckmeldung (29.06.2026): zu viele
Informationen ohne einen "roten Faden" auf einen Blick. Dieses Modul
verdichtet die wichtigsten Kennzahlen JE MATERIAL zu EINER Zeile - kein
Ersatz fuer die Detail-Reiter (die bleiben fuer Nachvollziehbarkeit
bestehen), sondern ein vorangestellter Ueberblick, analog zu Max'
Referenz-Template ('17_Executive_Summary'/'18_Executive_Dashboard').

Spaltenauswahl (Klaerung mit Max, 29.06.2026, siehe Projektstatus.md
Abschnitt 4f):
    - Aus ZMLAG_Bestand: Bestandswert (€, "Laufender Wert"), ABC/XYZ-Klasse
    - Aus Reichweite: Bestandsmenge (Stueck), Reichweite (Monate), MatStatus
    - Aus Trend_MVER: Trend_Einstufung
    - Aus Fruehwarnsystem: Meldebestand_ROP, Delta_Meldebestand (SAP-Vergleich)
    - Aus Simulation_Verbrauch: Fehlmengen_Wahrscheinlichkeit_Horizont

Join-Strategie: LEFT JOIN auf ZMLAG (jedes Material aus dem aktuellen
ZMLAG-Lauf bleibt erhalten), analog zu coverage_analysis.py - fehlende
Werte aus optionalen Phasen werden NICHT mit 0/Platzhalter aufgefuellt,
sondern bleiben NaN, falls die jeweilige Phase fuer ein Material nicht
berechnet werden konnte (z.B. fehlende Quelle fuer Fruehwarnsystem/
Simulation). Jede Quelle ist unabhaengig optional (Default None) - ein
Material erscheint immer, auch wenn z.B. nur Phase 1+2 vorliegen.

Zusaetzliche abgeleitete Spalte 'Prioritaet' (einfache Kategorisierung,
KEIN neues statistisches Modell): kombiniert Fehlmengen-Risiko und
ABC-Klasse zu einer lesbaren Vier-Stufen-Einordnung fuer die schnelle
visuelle Priorisierung im Dashboard (siehe report_builder.py, Reiter
'Executive_Summary'/'Dashboard'). Bewusst regelbasiert und transparent
nachvollziehbar (keine Blackbox-Gewichtung) - siehe _classify_priority().

Oeffentliche Hauptfunktion:
    build_executive_summary(zmlag_df, coverage_df=None, trend_df=None,
        safety_stock_df=None, simulation_df=None) -> DataFrame mit einer
        Zeile pro Material.
"""

from __future__ import annotations

import pandas as pd

import config

# Schwelle fuer "kritisches" Fehlmengen-Risiko (Anteil der Monte-Carlo-
# Laeufe mit Fehlmenge im Planungshorizont) - oberhalb dieser Schwelle wird
# ein Material unabhaengig von der ABC-Klasse als 'Kritisch' eingestuft.
# Wert in Abstimmung mit Max waehlbar; 0.5 als nachvollziehbarer Startwert
# ("in mehr als der Haelfte der simulierten Zukuenfte tritt eine Fehlmenge
# ein").
CRITICAL_SHORTAGE_PROBABILITY = 0.5

# Schwelle fuer "erhoehtes" Fehlmengen-Risiko, unterhalb von
# CRITICAL_SHORTAGE_PROBABILITY.
ELEVATED_SHORTAGE_PROBABILITY = 0.2


def _classify_priority(row: pd.Series) -> str:
    """Ordnet ein Material einer von vier Prioritaetsstufen zu, basierend
    auf Fehlmengen-Wahrscheinlichkeit (Baustein A) und ABC-Klasse (ZMLAG).

    Regelbasiert, bewusst einfach und ohne neue Gewichtung/Schwellenwert-
    Optimierung (kein eigenstaendiges Modell) - dient nur der schnellen
    visuellen Vorsortierung im Dashboard, ersetzt nicht die Detailanalyse
    in den Phasen 1-5.

    Regeln (in dieser Reihenfolge ausgewertet):
        1. Fehlmengen-Wahrscheinlichkeit fehlt (keine Simulation vorhanden)
           UND ABC/Reichweite fehlen ebenfalls -> 'Unbekannt' (zu wenig
           Datengrundlage fuer eine Aussage)
        2. Fehlmengen-Wahrscheinlichkeit >= CRITICAL_SHORTAGE_PROBABILITY
           -> 'Kritisch' (unabhaengig von der ABC-Klasse - eine hohe
           Fehlmengen-Wahrscheinlichkeit ist immer relevant)
        3. ABC-Klasse 'A' UND Fehlmengen-Wahrscheinlichkeit >=
           ELEVATED_SHORTAGE_PROBABILITY -> 'Kritisch' (A-Material mit
           bereits erhoehtem Risiko wiegt staerker als bei B/C)
        4. Fehlmengen-Wahrscheinlichkeit >= ELEVATED_SHORTAGE_PROBABILITY
           -> 'Erhoeht'
        5. sonst -> 'Unauffaellig'

    Args:
        row: Zeile der zusammengefuehrten Summary-Tabelle, muss
            'Fehlmengen_Wahrscheinlichkeit_Horizont' und 'ABC-Kennzeichen'
            enthalten (Werte duerfen NaN sein).

    Returns:
        'Kritisch', 'Erhoeht', 'Unauffaellig' oder 'Unbekannt'.
    """
    shortage_probability = row.get("Fehlmengen_Wahrscheinlichkeit_Horizont")
    abc_class = row.get("ABC-Kennzeichen")

    has_shortage_data = pd.notna(shortage_probability)
    has_abc_data = pd.notna(abc_class)

    if not has_shortage_data and not has_abc_data:
        return "Unbekannt"

    if not has_shortage_data:
        return "Unbekannt"

    if shortage_probability >= CRITICAL_SHORTAGE_PROBABILITY:
        return "Kritisch"

    if abc_class == "A" and shortage_probability >= ELEVATED_SHORTAGE_PROBABILITY:
        return "Kritisch"

    if shortage_probability >= ELEVATED_SHORTAGE_PROBABILITY:
        return "Erhoeht"

    return "Unauffaellig"


def build_executive_summary(
    zmlag_df: pd.DataFrame,
    coverage_df: pd.DataFrame | None = None,
    trend_df: pd.DataFrame | None = None,
    safety_stock_df: pd.DataFrame | None = None,
    simulation_df: pd.DataFrame | None = None,
    working_capital_df: pd.DataFrame | None = None,
    lead_time_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Baut die Executive-Summary-Tabelle: eine Zeile pro Material aus
    zmlag_df, angereichert um die wichtigsten Kennzahlen aus den optional
    uebergebenen Phasen-Ergebnissen.

    LEFT JOIN auf zmlag_df (analog zu coverage_analysis.merge_stock_and_
    consumption()) - jedes Material aus dem aktuellen ZMLAG-Lauf bleibt
    erhalten, auch wenn einzelne Phasen fuer dieses Material nicht
    berechnet werden konnten (z.B. fehlende Dispo_ABCXYZ-Zeile fuer
    Fruehwarnsystem). Fehlende Werte bleiben NaN, kein Platzhalter.

    Args:
        zmlag_df: vollstaendig angereicherter ZMLAG-DataFrame (nach
            stock_analysis.calculate_stock_development() und
            classify_risk_by_abc_xyz()) - Pflichtquelle, da Phase 1 immer
            vorliegt. Muss mindestens 'Material', 'Materialkurztext',
            'Technologie', 'Laufender Wert', 'ABC-Kennzeichen',
            'ABCXYZ-Kennzeichen' enthalten.
        coverage_df: optionales Ergebnis von coverage_analysis.
            build_coverage_table() (Phase 2) - liefert Bestandsmenge,
            Reichweite, MatStatus.
        trend_df: optionales Ergebnis von trend_analysis.build_trend_table()
            (erster Tupel-Eintrag, Phase 3) - liefert Trend_Einstufung.
            Nur die Kennzahlenspalten werden uebernommen, NICHT die rohen
            Monatsspalten (die bleiben Trend_MVER vorbehalten).
        safety_stock_df: optionales Ergebnis von safety_stock.
            build_safety_stock_table() (Phase 5) - liefert Meldebestand_ROP
            und (falls vorhanden) Delta_Meldebestand.
        simulation_df: optionales Ergebnis von simulation_analysis.
            build_simulation_table() (Baustein A) - liefert
            Fehlmengen_Wahrscheinlichkeit_Horizont.
        working_capital_df: optionales Ergebnis von working_capital.
            build_working_capital_table() (Phase 6, ab 30.06.2026) - liefert
            Working_Capital_Reduktion und ROI_Prozent. Siehe Projektstatus.md
            Abschnitt 4i fuer die Designklaerung.
        lead_time_df: optionales Ergebnis von dispo_abcxyz_loader.
            build_lead_time_table() (Phase 4) - liefert
            Wiederbeschaffungszeit_Monate (aus Wiederbeschaffungszeit_Tage
            / config.DAYS_PER_MONTH), Voraussetzung fuer die Reichweite-
            vs-WBZ-Matrix im HTML-Dashboard (Abschnitt 4k Projektstatus.md,
            30.06.2026). Rueckwaertskompatibel: ohne lead_time_df fehlt die
            Spalte einfach (NaN ueberall via LEFT JOIN), kein Fehler.

    Returns:
        DataFrame mit einer Zeile pro Material, Spalten:
            - Material, Materialkurztext, Technologie (aus ZMLAG)
            - Bestandswert_EUR (= 'Laufender Wert', umbenannt fuer
              Eindeutigkeit im Summary-Kontext)
            - ABC_Kennzeichen, ABCXYZ_Kennzeichen (aus ZMLAG)
            - Bestandsmenge_Stueck, Reichweite_Monate, MatStatus_Bezeichnung
              (aus coverage_df, falls vorhanden, sonst NaN)
            - Trend_Einstufung (aus trend_df, falls vorhanden, sonst NaN)
            - Meldebestand_ROP, Delta_Meldebestand (aus safety_stock_df,
              falls vorhanden, sonst NaN; Delta_Meldebestand zusaetzlich
              NaN, falls safety_stock_df keine SAP-Vergleichsspalten hatte)
            - Fehlmengen_Wahrscheinlichkeit_Horizont (aus simulation_df,
              falls vorhanden, sonst NaN)
            - Working_Capital_Reduktion, ROI_Prozent (aus working_capital_df,
              falls vorhanden, sonst NaN; siehe working_capital.py)
            - Prioritaet (siehe _classify_priority(), regelbasiert)

    Raises:
        ValueError: wenn zmlag_df leer ist oder die Pflichtspalten fehlen.
    """
    required_zmlag_cols = {
        "Material", "Materialkurztext", "Technologie", "Laufender Wert",
        "ABC-Kennzeichen", "ABCXYZ-Kennzeichen",
    }
    missing = required_zmlag_cols - set(zmlag_df.columns)
    if missing:
        raise ValueError(
            f"executive_summary.build_executive_summary() erwartet in "
            f"zmlag_df die Spalten {sorted(required_zmlag_cols)}, es "
            f"fehlen: {sorted(missing)}."
        )
    if zmlag_df.empty:
        raise ValueError(
            "executive_summary.build_executive_summary() hat einen leeren "
            "zmlag_df erhalten - keine Materialien zum Zusammenfassen "
            "vorhanden."
        )

    result = zmlag_df[
        ["Material", "Materialkurztext", "Technologie", "Laufender Wert",
         "ABC-Kennzeichen", "ABCXYZ-Kennzeichen"]
    ].copy()
    result = result.rename(columns={"Laufender Wert": "Bestandswert_EUR"})

    if coverage_df is not None:
        coverage_cols = [
            col for col in [
                "Material", "Bestandsmenge", "Reichweite_Monate_StockOnly",
                "MatStatus_Bezeichnung",
            ]
            if col in coverage_df.columns
        ]
        result = result.merge(coverage_df[coverage_cols], on="Material", how="left")
        result = result.rename(columns={
            "Bestandsmenge": "Bestandsmenge_Stueck",
            "Reichweite_Monate_StockOnly": "Reichweite_Monate",
        })

    if trend_df is not None and "Trend_Einstufung" in trend_df.columns:
        result = result.merge(
            trend_df[["Material", "Trend_Einstufung"]], on="Material", how="left"
        )

    if safety_stock_df is not None:
        safety_stock_cols = [
            col for col in ["Material", "Meldebestand_ROP", "Delta_Meldebestand"]
            if col in safety_stock_df.columns
        ]
        result = result.merge(
            safety_stock_df[safety_stock_cols], on="Material", how="left"
        )

    if simulation_df is not None and (
        "Fehlmengen_Wahrscheinlichkeit_Horizont" in simulation_df.columns
    ):
        result = result.merge(
            simulation_df[["Material", "Fehlmengen_Wahrscheinlichkeit_Horizont"]],
            on="Material", how="left",
        )

    if working_capital_df is not None:
        working_capital_cols = [
            col for col in ["Material", "Working_Capital_Reduktion", "ROI_Prozent"]
            if col in working_capital_df.columns
        ]
        result = result.merge(
            working_capital_df[working_capital_cols], on="Material", how="left"
        )

    if lead_time_df is not None and "Wiederbeschaffungszeit_Tage" in lead_time_df.columns:
        lead_time_slim = lead_time_df[["Material", "Wiederbeschaffungszeit_Tage"]].copy()
        lead_time_slim["Wiederbeschaffungszeit_Monate"] = (
            lead_time_slim["Wiederbeschaffungszeit_Tage"] / config.DAYS_PER_MONTH
        )
        result = result.merge(
            lead_time_slim[["Material", "Wiederbeschaffungszeit_Monate"]],
            on="Material", how="left",
        )

    result["Prioritaet"] = result.apply(_classify_priority, axis=1)

    return result

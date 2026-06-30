"""
working_capital.py
====================
Working Capital / ROI-Berechnung - Fortsetzung von Phase 6 (Gesamtreport &
COO-Praesentation).

Fachlicher Hintergrund (siehe Projektstatus.md Abschnitt 4i fuer die
vollstaendige Designklaerung mit Max, 30.06.2026, sowie Whitepaper
Inventory_Management_I.pdf, Kap. 8 "Economic Impact and ROI Analysis", und
Ivanov, "Global Supply Chain and Operations Management", Kap. 13.10/15):

    Working_Capital = Bestandsmenge x Stueckpreis

Das Whitepaper definiert "Working Capital optimal" als Ergebnis einer
vollstaendigen (s,q)/(s,S)-Policy-Optimierungssimulation (Whitepaper Kap. 7).
GMTS hat dafuer (Stand 30.06.2026) keine Entsprechung - Baustein A
(simulation_analysis.py) simuliert nur den kuenftigen Verbrauch, keine
Bestandsdynamik mit Wiederauffuellung. Entscheidung (Klaerung mit Max,
30.06.2026): "Optimal" wird vereinfacht als die bereits vorhandene
Meldebestand_ROP-Zielgroesse aus Phase 5 (safety_stock.py, geschlossene
Formel) verwendet, NICHT als Ergebnis einer neuen Optimierungssimulation.
Eine echte Policy-Simulation mit tatsaechlichen Zugaengen ist als
eigenstaendiger, spaeterer Baustein vorgemerkt (Arbeitstitel "Baustein C",
gekoppelt an Baustein B - siehe Projektstatus.md Abschnitt 4i).

Finanzparameter (Vorgabe Max, 30.06.2026, siehe config.py):
    HOLDING_COST_RATE                  = 0.20  (20% p.a.)
    ORDER_COST                         = 75.0  (EUR je Bestellung)
    EMERGENCY_FREIGHT_SURCHARGE_RATE   = 0.30  (30% Aufschlag auf Standardpreis)

Scope-Eingrenzung v1 (Klaerung mit Max, 30.06.2026): Baustein A liefert
KEINE erwartete Fehlmengen-MENGE in Stueck (nur eine Fehlmengen-
WAHRSCHEINLICHKEIT), daher ist Shortage_Cost (Whitepaper 8.5) hier NICHT
berechenbar, ohne simulation_analysis.py zu erweitern. Ebenso fehlt eine
Bestellanzahl ohne Policy-Simulation, wodurch Ordering_Cost (Whitepaper 8.4)
entfaellt. Beide bewusst zurueckgestellt, NICHT verworfen - vorgemerkt fuer
eine spaetere Erweiterung (siehe Projektstatus.md Abschnitt 4i).

IMPLEMENTATION_COST/ANNUAL_MAINTENANCE_COST (Whitepaper 8.7/8.8) bleiben
irrelevant (Klaerung mit Max, 30.06.2026) - keine klassische ROI-Berechnung
gegen Implementierungskosten. Stattdessen: ROI_Prozent als Kapitalrendite-
Kennzahl ("jaehrliche Ersparnis im Verhaeltnis zum aktuell gebundenen
Kapital"), als Anteil/Bruch gespeichert (z.B. 0.08 fuer 8%) - konsistent
mit allen uebrigen Prozent-Spalten im Projekt (Ziel_Servicegrad,
Fehlmengen_Wahrscheinlichkeit_Horizont), damit das Excel-Prozentformat
'0.00%' korrekt greift. Payback-Periode entfaellt vollstaendig (ohne
Investitionssumme nicht sinnvoll definierbar).

Materialauswahl: INNER JOIN aus coverage_df (Bestandsmenge, Phase 2),
safety_stock_df (Meldebestand_ROP, Phase 5) und lead_time_df (Standardpreis,
Phase 4) - analog zu safety_stock._merge_inputs(). Ein Material ohne ALLE
DREI Quellen wird NICHT berechnet (keine Default-Annahmen, Konsolen-Hinweis
mit Anzahl) - konsistent mit dem Prinzip aus Phase 2/3/5/Baustein A.

Oeffentliche Hauptfunktion:
    build_working_capital_table(coverage_df, safety_stock_df, lead_time_df)
        -> DataFrame mit einer Zeile pro Material und den unten
        beschriebenen Kennzahlen (siehe build_working_capital_table()-
        Docstring).
"""

from __future__ import annotations

import pandas as pd

import config


def _merge_inputs(
    coverage_df: pd.DataFrame,
    safety_stock_df: pd.DataFrame,
    lead_time_df: pd.DataFrame,
) -> pd.DataFrame:
    """Fuehrt coverage_df (Bestandsmenge), safety_stock_df (Meldebestand_ROP)
    und lead_time_df (Standardpreis) auf Material-Ebene zusammen - INNER
    JOIN, siehe Modul-Docstring fuer die Begruendung. Gibt eine Konsolen-
    Warnung aus, falls dabei Materialien herausfallen (analog zu
    safety_stock._merge_inputs()/simulation_analysis._merge_inputs()).

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table(),
            muss 'Material' und 'Bestandsmenge' enthalten.
        safety_stock_df: Ergebnis von safety_stock.build_safety_stock_table(),
            muss 'Material' und 'Meldebestand_ROP' enthalten.
        lead_time_df: Ergebnis von dispo_abcxyz_loader.build_lead_time_table(),
            muss 'Material' und 'Standardpreis' enthalten.

    Returns:
        DataFrame mit einer Zeile pro Material, das in ALLEN DREI
        Eingabe-DataFrames vorkommt.

    Raises:
        ValueError: wenn einer der drei Eingabe-DataFrames die erwarteten
            Pflichtspalten nicht enthaelt.
    """
    required_coverage_cols = {"Material", "Bestandsmenge"}
    missing_coverage = required_coverage_cols - set(coverage_df.columns)
    if missing_coverage:
        raise ValueError(
            f"working_capital._merge_inputs() erwartet in coverage_df die "
            f"Spalten {sorted(required_coverage_cols)}, es fehlen: "
            f"{sorted(missing_coverage)}."
        )

    required_safety_stock_cols = {"Material", "Meldebestand_ROP"}
    missing_safety_stock = required_safety_stock_cols - set(safety_stock_df.columns)
    if missing_safety_stock:
        raise ValueError(
            f"working_capital._merge_inputs() erwartet in safety_stock_df "
            f"die Spalten {sorted(required_safety_stock_cols)}, es fehlen: "
            f"{sorted(missing_safety_stock)}."
        )

    required_lead_time_cols = {"Material", "Standardpreis"}
    missing_lead_time = required_lead_time_cols - set(lead_time_df.columns)
    if missing_lead_time:
        raise ValueError(
            f"working_capital._merge_inputs() erwartet in lead_time_df die "
            f"Spalten {sorted(required_lead_time_cols)}, es fehlen: "
            f"{sorted(missing_lead_time)}."
        )

    merged = coverage_df[["Material", "Bestandsmenge"]].merge(
        safety_stock_df[["Material", "Meldebestand_ROP"]], on="Material", how="inner"
    )
    merged = merged.merge(
        lead_time_df[["Material", "Standardpreis"]], on="Material", how="inner"
    )

    n_coverage = coverage_df["Material"].nunique()
    n_safety_stock = safety_stock_df["Material"].nunique()
    n_lead_time = lead_time_df["Material"].nunique()
    n_merged = merged["Material"].nunique()
    n_excluded = n_coverage - n_merged
    if n_excluded > 0:
        print(
            f"  Hinweis (Working Capital): {n_excluded} Material(ien) aus "
            f"Reichweite ({n_coverage}) haben keine gemeinsame Datengrundlage "
            f"mit Fruehwarnsystem ({n_safety_stock}) und/oder Wiederbeschaffung "
            f"({n_lead_time}) - werden NICHT in Working_Capital berechnet "
            f"(Working_Capital benoetigt alle drei Quellen)."
        )

    missing_price = merged[merged["Standardpreis"].isna()]
    if not missing_price.empty:
        print(
            f"  Hinweis (Working Capital): {len(missing_price)} Material(ien) "
            f"haben keinen Standardpreis hinterlegt und werden mit NaN fuer "
            f"alle Working-Capital-Kennzahlen ausgewiesen: "
            f"{list(missing_price['Material'])}"
        )

    return merged


def build_working_capital_table(
    coverage_df: pd.DataFrame,
    safety_stock_df: pd.DataFrame,
    lead_time_df: pd.DataFrame,
) -> pd.DataFrame:
    """Komplettlauf der Working-Capital-/ROI-Berechnung (Phase 6, Fortsetzung):
    merged die drei Eingabequellen und berechnet je Material den aktuellen
    und den ueber Meldebestand_ROP definierten "optimalen" Bestandswert,
    die daraus resultierende Kapitalbindungs-Reduktion, die jaehrliche
    Holding-Cost-Ersparnis sowie eine Kapitalrendite-Kennzahl (ROI_Prozent).

    Siehe Modul-Docstring fuer die Designklaerung (insb. die Vereinfachung
    "Optimal = Meldebestand_ROP" statt einer vollwertigen Policy-
    Optimierungssimulation, sowie die bewusste Eingrenzung auf Holding Cost
    ohne Shortage_Cost/Ordering_Cost in v1).

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table()
            fuer dieselbe Technologie.
        safety_stock_df: Ergebnis von safety_stock.build_safety_stock_table()
            fuer dieselbe Technologie.
        lead_time_df: Ergebnis von dispo_abcxyz_loader.build_lead_time_table(),
            bereits auf die Materialien dieser Technologie gefiltert.

    Returns:
        DataFrame mit einer Zeile pro Material (nur Materialien, die in
        ALLEN DREI Pflichtquellen vorkommen - siehe _merge_inputs()), mit
        folgenden Spalten, als Herleitungskette von links nach rechts
        angeordnet (analog zum "roter Faden"-Prinzip aus Phase 5):
            - Material
            - Bestandsmenge_Stueck (= 'Aktuell', aus coverage_df)
            - Meldebestand_ROP_Stueck (= 'Optimal', aus safety_stock_df)
            - Standardpreis (aus lead_time_df)
            - Working_Capital_Aktuell (= Bestandsmenge_Stueck * Standardpreis)
            - Working_Capital_Optimal (= Meldebestand_ROP_Stueck * Standardpreis)
            - Working_Capital_Reduktion (= Aktuell - Optimal; positiv =
              Kapital freisetzbar, negativ = Soll-Bestand hoeher als Ist,
              kein Risiko verschweigen)
            - Holding_Cost_Aktuell, Holding_Cost_Optimal (= jeweiliger
              Working-Capital-Wert * config.HOLDING_COST_RATE)
            - Annual_Savings (= Holding_Cost_Aktuell - Holding_Cost_Optimal)
            - ROI_Prozent (= Annual_Savings / Working_Capital_Aktuell; als
              Anteil/Bruch wie alle uebrigen Prozent-Spalten im Projekt, z.B.
              0.08 fuer 8% - konsistent mit Ziel_Servicegrad/Fehlmengen_
              Wahrscheinlichkeit_Horizont, damit das Excel-Prozentformat
              '0.00%' korrekt greift; NaN, falls Working_Capital_Aktuell 0
              oder NaN ist, um Division durch 0 zu vermeiden)

    Raises:
        ValueError: wenn einer der drei Eingabe-DataFrames die erwarteten
            Spalten nicht enthaelt (siehe _merge_inputs()).
    """
    merged = _merge_inputs(coverage_df, safety_stock_df, lead_time_df)

    result = merged.rename(columns={
        "Bestandsmenge": "Bestandsmenge_Stueck",
        "Meldebestand_ROP": "Meldebestand_ROP_Stueck",
    })

    result["Working_Capital_Aktuell"] = (
        result["Bestandsmenge_Stueck"] * result["Standardpreis"]
    )
    result["Working_Capital_Optimal"] = (
        result["Meldebestand_ROP_Stueck"] * result["Standardpreis"]
    )
    result["Working_Capital_Reduktion"] = (
        result["Working_Capital_Aktuell"] - result["Working_Capital_Optimal"]
    )

    result["Holding_Cost_Aktuell"] = (
        result["Working_Capital_Aktuell"] * config.HOLDING_COST_RATE
    )
    result["Holding_Cost_Optimal"] = (
        result["Working_Capital_Optimal"] * config.HOLDING_COST_RATE
    )
    result["Annual_Savings"] = (
        result["Holding_Cost_Aktuell"] - result["Holding_Cost_Optimal"]
    )

    result["ROI_Prozent"] = result.apply(
        lambda row: (
            row["Annual_Savings"] / row["Working_Capital_Aktuell"]
            if pd.notna(row["Working_Capital_Aktuell"]) and row["Working_Capital_Aktuell"] != 0
            else float("nan")
        ),
        axis=1,
    )

    column_order = [
        "Material",
        "Bestandsmenge_Stueck",
        "Meldebestand_ROP_Stueck",
        "Standardpreis",
        "Working_Capital_Aktuell",
        "Working_Capital_Optimal",
        "Working_Capital_Reduktion",
        "Holding_Cost_Aktuell",
        "Holding_Cost_Optimal",
        "Annual_Savings",
        "ROI_Prozent",
    ]
    return result[column_order]


def aggregate_working_capital(working_capital_df: pd.DataFrame) -> dict:
    """Aggregiert die Working-Capital-Tabelle auf Technologie-/Gesamtebene -
    fuer die COO-Kommunikation (Executive Summary/Dashboard), analog zu
    Whitepaper Kap. 8.9 "Aggregation Across All Materials", aber OHNE
    Total_Cost_Current/Total_Cost_Optimal (siehe Modul-Docstring fuer die
    Scope-Eingrenzung - Ordering_Cost/Shortage_Cost sind in v1 nicht
    Teil der Berechnung).

    Args:
        working_capital_df: Ergebnis von build_working_capital_table().

    Returns:
        Dict mit den aggregierten Summen 'Working_Capital_Aktuell_Summe',
        'Working_Capital_Optimal_Summe', 'Working_Capital_Reduktion_Summe',
        'Annual_Savings_Summe' sowie 'ROI_Prozent_Gesamt' (= Annual_Savings_
        Summe / Working_Capital_Aktuell_Summe, als Anteil/Bruch wie
        ROI_Prozent in build_working_capital_table(), NaN bei Summe 0).
        Leeres Dict (alle Werte NaN), falls working_capital_df leer ist.
    """
    if working_capital_df.empty:
        return {
            "Working_Capital_Aktuell_Summe": float("nan"),
            "Working_Capital_Optimal_Summe": float("nan"),
            "Working_Capital_Reduktion_Summe": float("nan"),
            "Annual_Savings_Summe": float("nan"),
            "ROI_Prozent_Gesamt": float("nan"),
        }

    wc_aktuell_summe = working_capital_df["Working_Capital_Aktuell"].sum()
    wc_optimal_summe = working_capital_df["Working_Capital_Optimal"].sum()
    wc_reduktion_summe = working_capital_df["Working_Capital_Reduktion"].sum()
    annual_savings_summe = working_capital_df["Annual_Savings"].sum()

    roi_prozent_gesamt = (
        annual_savings_summe / wc_aktuell_summe
        if wc_aktuell_summe not in (0, None) and pd.notna(wc_aktuell_summe)
        else float("nan")
    )

    return {
        "Working_Capital_Aktuell_Summe": float(wc_aktuell_summe),
        "Working_Capital_Optimal_Summe": float(wc_optimal_summe),
        "Working_Capital_Reduktion_Summe": float(wc_reduktion_summe),
        "Annual_Savings_Summe": float(annual_savings_summe),
        "ROI_Prozent_Gesamt": float(roi_prozent_gesamt),
    }

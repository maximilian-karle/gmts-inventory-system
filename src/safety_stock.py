"""
safety_stock.py
================
Fruehwarnsystem (Phase 5): Safety Stock, Meldebestand (ROP) und
Service-Level-Mapping nach ABC/XYZ-Klasse.

Fachlicher Hintergrund (siehe Projektstatus.md Abschnitt 6 "Naechste
Schritte" sowie Ivanov, "Global Supply Chain and Operations Management",
Kap. 13.3/13.5/13.6, und Ivanov, "Introduction to Supply Chain Analytics",
Kap. 2.2):

    Safety_Stock = z * sigma * sqrt(L)
    ROP          = d * L + Safety_Stock

    z      = Verteilungsquantil zum angestrebten Servicegrad (Service Level),
             abhaengig von der ABC/XYZ-Klasse (siehe config.SERVICE_LEVEL_MATRIX)
    sigma  = Standardabweichung des monatlichen Verbrauchs, abgeleitet aus
             der MAD (Median Absolute Deviation) aus Phase 3:
             sigma = Verbrauch_Streuung_MAD / 0.6745 (identische Konvention
             wie in trend_analysis.py und simulation_analysis.py)
    L      = Wiederbeschaffungszeit in MONATEN (Phase 4 liefert Tage,
             siehe config.DAYS_PER_MONTH fuer die Umrechnung)
    d      = mittlerer monatlicher Verbrauch (Ø_Verbrauch_Fenster bzw.
             Verbrauch_Fuer_Reichweite aus Phase 2, falls ein Planungs-
             Override gesetzt ist)

Dieses Modul ist damit - im Unterschied zu simulation_analysis.py
(Baustein A, stochastische Simulation des KUENFTIGEN Verbrauchs) - die
KLASSISCHE, geschlossene Safety-Stock-Formel: kein Zufallsmodell, dafuer
sofort nachvollziehbar und direkt mit den SAP-eigenen Werten
(Sicherheitsbestand_Soll, Meldebestand aus Dispo_ABCXYZ) vergleichbar.

Materialauswahl (Entscheidung Claude, 29.06.2026, analog zu
simulation_analysis._merge_inputs()): INNER JOIN aus coverage_df
(Ø-Verbrauch, Phase 2), trend_df (Streuung, Phase 3) und lead_time_df
(Wiederbeschaffungszeit + ABC/XYZ-Klasse, Phase 4). Ein Material ohne ALLE
DREI Quellen wird NICHT berechnet (keine Default-Annahmen, Konsolen-
Hinweis mit Anzahl) - konsistent mit dem Prinzip aus Phase 2/3/Baustein A.

ABC/XYZ-Quelle (Klaerung mit Max, 29.06.2026): Dispo_ABCXYZ
('ABCXYZ_Kombiniert') ist die PRIMAERE Quelle fuer die Service-Level-
Zuordnung - analog zur bereits etablierten Entscheidung in
dispo_abcxyz_loader.py, Dispo_ABCXYZ als primaere Quelle fuer die
Wiederbeschaffungszeit zu behandeln (dort: Plausibilitaetsabgleich gegen
den Inventory-Report). Hier: optionaler Plausibilitaetsabgleich gegen
ZMLAG ('ABCXYZ-Kennzeichen', siehe stock_analysis.py), falls zmlag_df
uebergeben wird - bei Abweichung Konsolen-Warnung mit betroffenen
Materialien, OHNE den Lauf abzubrechen (Dispo_ABCXYZ bleibt massgeblich).

Oeffentliche Hauptfunktion:
    build_safety_stock_table(coverage_df, trend_df, lead_time_df, zmlag_df)
        -> DataFrame mit einer Zeile pro Material und den unten
        beschriebenen Kennzahlen (siehe build_safety_stock_table()-Docstring).
"""

from __future__ import annotations

import math

import pandas as pd

import config

# Umrechnung Median Absolute Deviation (MAD) -> Standardabweichung (sigma),
# identische Konvention wie trend_analysis.py/simulation_analysis.py
# (Iglewicz und Hoaglin, 1993).
MAD_TO_SIGMA_FACTOR = 0.6745


# ---------------------------------------------------------------------------
# Inverse Standardnormalverteilung ohne scipy-Abhaengigkeit
# ---------------------------------------------------------------------------
# Entscheidung Claude, 29.06.2026: das Projekt nutzt bislang ausschliesslich
# pandas/numpy/openpyxl (siehe alle uebrigen Module) - scipy wuerde als
# einzige Stelle im gesamten Projekt eine zusaetzliche Abhaengigkeit
# einfuehren, nur fuer eine einzelne Funktion (Quantil der Normalverteilung).
# Stattdessen: Peter Acklams rationale Approximation des Quantils der
# Standardnormalverteilung (oeffentlich dokumentierter Algorithmus,
# relative Genauigkeit < 1.15e-9 im Bereich (0,1)) - fuer Safety-Stock-
# z-Werte um Groessenordnungen genauer als notwendig. Gegen scipy.stats.
# norm.ppf() verifiziert (siehe Testlauf 29.06.2026): Abweichung < 1e-5
# fuer alle in config.SERVICE_LEVEL_MATRIX vorkommenden Servicegrade.
def _inverse_normal_cdf(probability: float) -> float:
    """Quantil der Standardnormalverteilung zu einer Wahrscheinlichkeit
    (0, 1) - entspricht scipy.stats.norm.ppf(probability), ohne scipy.

    Args:
        probability: Wahrscheinlichkeit in (0, 1), z.B. 0.98 fuer einen
            98%-Servicegrad.

    Returns:
        z-Wert (Quantil der Standardnormalverteilung).

    Raises:
        ValueError: wenn probability nicht in (0, 1) liegt.
    """
    if not (0.0 < probability < 1.0):
        raise ValueError(
            f"_inverse_normal_cdf() erwartet eine Wahrscheinlichkeit in (0, 1), "
            f"erhalten: {probability!r}."
        )

    # Koeffizienten der Acklam-Approximation.
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
         6.680131188771972e+01, -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
         -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00,
         3.754408661907416e+00]
    p_low = 0.02425
    p_high = 1 - p_low

    if probability < p_low:
        q = math.sqrt(-2 * math.log(probability))
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
               ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    if probability <= p_high:
        q = probability - 0.5
        r = q * q
        return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / \
               (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1)
    q = math.sqrt(-2 * math.log(1 - probability))
    return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
            ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)


def _merge_inputs(
    coverage_df: pd.DataFrame,
    trend_df: pd.DataFrame,
    lead_time_df: pd.DataFrame,
) -> pd.DataFrame:
    """Fuehrt coverage_df (Ø-Verbrauch), trend_df (Streuung) und
    lead_time_df (Wiederbeschaffungszeit + ABC/XYZ-Klasse) auf Material-
    Ebene zusammen - INNER JOIN, siehe Modul-Docstring fuer die Begruendung.
    Gibt eine Konsolen-Warnung aus, falls dabei Materialien herausfallen
    (analog zu simulation_analysis._merge_inputs()).

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table(),
            muss 'Material' und 'Verbrauch_Fuer_Reichweite' enthalten.
        trend_df: Ergebnis von trend_analysis.build_trend_table() (erster
            Tupel-Eintrag), muss 'Material' und 'Verbrauch_Streuung_MAD'
            enthalten.
        lead_time_df: Ergebnis von dispo_abcxyz_loader.build_lead_time_table(),
            muss 'Material', 'Wiederbeschaffungszeit_Tage' und
            'ABCXYZ_Kombiniert' enthalten. Optional (falls vorhanden):
            'Sicherheitsbestand_Soll', 'Meldebestand' fuer den SAP-Vergleich.

    Returns:
        DataFrame mit einer Zeile pro Material, das in ALLEN DREI
        Eingabe-DataFrames vorkommt.
    """
    required_coverage_cols = {"Material", "Verbrauch_Fuer_Reichweite"}
    missing_coverage = required_coverage_cols - set(coverage_df.columns)
    if missing_coverage:
        raise ValueError(
            f"safety_stock._merge_inputs() erwartet in coverage_df die "
            f"Spalten {sorted(required_coverage_cols)}, es fehlen: "
            f"{sorted(missing_coverage)}."
        )

    required_trend_cols = {"Material", "Verbrauch_Streuung_MAD"}
    missing_trend = required_trend_cols - set(trend_df.columns)
    if missing_trend:
        raise ValueError(
            f"safety_stock._merge_inputs() erwartet in trend_df die Spalten "
            f"{sorted(required_trend_cols)}, es fehlen: {sorted(missing_trend)}."
        )

    required_lead_time_cols = {"Material", "Wiederbeschaffungszeit_Tage", "ABCXYZ_Kombiniert"}
    missing_lead_time = required_lead_time_cols - set(lead_time_df.columns)
    if missing_lead_time:
        raise ValueError(
            f"safety_stock._merge_inputs() erwartet in lead_time_df die "
            f"Spalten {sorted(required_lead_time_cols)}, es fehlen: "
            f"{sorted(missing_lead_time)}."
        )

    coverage_materials = set(coverage_df["Material"])
    trend_materials = set(trend_df["Material"])
    lead_time_materials = set(lead_time_df["Material"])
    common_materials = coverage_materials & trend_materials & lead_time_materials

    for label, materials in (
        ("Reichweite (Phase 2)", coverage_materials),
        ("Trend (Phase 3)", trend_materials),
        ("Wiederbeschaffung (Phase 4)", lead_time_materials),
    ):
        missing_here = materials - common_materials
        if missing_here:
            print(
                f"  Hinweis (Fruehwarnsystem): {len(missing_here)} Material(ien) "
                f"aus '{label}' fehlen in mindestens einer der beiden anderen "
                f"Quellen und werden NICHT berechnet."
            )

    optional_lead_time_cols = [
        col for col in ("Sicherheitsbestand_Soll", "Meldebestand")
        if col in lead_time_df.columns
    ]

    merged = (
        coverage_df[["Material", "Verbrauch_Fuer_Reichweite"]]
        .merge(
            trend_df[["Material", "Verbrauch_Streuung_MAD"]],
            on="Material",
            how="inner",
        )
        .merge(
            lead_time_df[
                ["Material", "Wiederbeschaffungszeit_Tage", "ABCXYZ_Kombiniert"]
                + optional_lead_time_cols
            ],
            on="Material",
            how="inner",
        )
    )
    return merged


def _check_abc_xyz_consistency(merged: pd.DataFrame, zmlag_df: pd.DataFrame | None) -> None:
    """Vergleicht die primaere ABC/XYZ-Klasse (Dispo_ABCXYZ,
    'ABCXYZ_Kombiniert') mit der ZMLAG-Klasse ('ABCXYZ-Kennzeichen'), falls
    zmlag_df uebergeben wird, und gibt bei Abweichung eine Konsolen-Warnung
    aus - analog zu dispo_abcxyz_loader.compare_lead_time_sources(). Bricht
    den Lauf NICHT ab; Dispo_ABCXYZ bleibt massgeblich (Klaerung mit Max,
    29.06.2026, siehe Modul-Docstring).

    Tut nichts, falls zmlag_df None ist oder die erwartete Spalte fehlt -
    der Abgleich ist eine zusaetzliche Plausibilitaetspruefung, keine
    Voraussetzung fuer die Berechnung selbst.
    """
    if zmlag_df is None or "ABCXYZ-Kennzeichen" not in zmlag_df.columns:
        return

    comparison = merged[["Material", "ABCXYZ_Kombiniert"]].merge(
        zmlag_df[["Material", "ABCXYZ-Kennzeichen"]],
        on="Material",
        how="inner",
    )
    mismatch = comparison[
        comparison["ABCXYZ_Kombiniert"] != comparison["ABCXYZ-Kennzeichen"]
    ]
    if not mismatch.empty:
        print(
            f"  WARNUNG (Fruehwarnsystem): {len(mismatch)} Material(ien) haben "
            f"abweichende ABC/XYZ-Klassen zwischen Dispo_ABCXYZ (massgeblich) "
            f"und ZMLAG: {list(mismatch['Material'])}"
        )


def _resolve_service_level(abc_xyz_code: object) -> tuple[float, float]:
    """Loest fuer einen kombinierten ABC/XYZ-Code (z.B. 'AX') den Ziel-
    Servicegrad und den zugehoerigen z-Wert auf (siehe
    config.SERVICE_LEVEL_MATRIX).

    Args:
        abc_xyz_code: 2-Zeichen-String (ABC-Komponente, XYZ-Komponente),
            z.B. 'AX', 'BZ'. Identisches Format wie in
            stock_analysis.classify_risk_by_abc_xyz() (code[0] = ABC,
            code[1] = XYZ).

    Returns:
        (service_level, z_value). (NaN, NaN), falls abc_xyz_code kein
        gueltiger 2-Zeichen-Code ist oder die Kombination nicht in
        SERVICE_LEVEL_MATRIX vorkommt (z.B. unbekanntes Kennzeichen) -
        bewusst KEIN stiller Default-Servicegrad, siehe Modul-Docstring-
        Prinzip "keine Default-Annahmen".
    """
    if not isinstance(abc_xyz_code, str) or len(abc_xyz_code) != 2:
        return float("nan"), float("nan")

    abc, xyz = abc_xyz_code[0], abc_xyz_code[1]
    service_level = config.SERVICE_LEVEL_MATRIX.get(abc, {}).get(xyz)
    if service_level is None:
        return float("nan"), float("nan")

    return service_level, _inverse_normal_cdf(service_level)


def build_safety_stock_table(
    coverage_df: pd.DataFrame,
    trend_df: pd.DataFrame,
    lead_time_df: pd.DataFrame,
    zmlag_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Komplettlauf der Safety-Stock-/ROP-Berechnung (Phase 5): merged die
    drei Eingabequellen, loest je Material den Ziel-Servicegrad ueber die
    ABC/XYZ-Klasse auf und berechnet Safety Stock sowie Meldebestand (ROP).
    Stellt zusaetzlich den SAP-eigenen Sicherheitsbestand/Meldebestand
    (sofern in lead_time_df vorhanden) als Vergleichsspalten gegenueber
    (Klaerung mit Max, 29.06.2026 - "roter Faden" fuer die COO-Kommunikation:
    zeigt auf einen Blick, was die Analyse gegenueber dem SAP-Status-quo
    aendert).

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table()
            fuer dieselbe Technologie.
        trend_df: Ergebnis von trend_analysis.build_trend_table() (erster
            Tupel-Eintrag) fuer dieselbe Technologie.
        lead_time_df: Ergebnis von dispo_abcxyz_loader.build_lead_time_table(),
            bereits auf die Materialien dieser Technologie gefiltert.
        zmlag_df: optionaler vollstaendig angereicherter ZMLAG-DataFrame
            (wie er auch an report_builder.build_report() uebergeben wird)
            fuer den ABC/XYZ-Plausibilitaetsabgleich (siehe
            _check_abc_xyz_consistency()). Falls None (Default), wird KEIN
            Abgleich durchgefuehrt - haelt den Aufruf abwaertskompatibel.

    Returns:
        DataFrame mit einer Zeile pro Material (nur Materialien, die in
        ALLEN DREI Pflichtquellen vorkommen - siehe _merge_inputs()), mit
        folgenden Spalten, bewusst als Herleitungskette von links nach
        rechts angeordnet (Klaerung mit Max, 29.06.2026 - "roter Faden"):
            - Material
            - ABCXYZ_Kombiniert (massgebliche Klasse, aus Dispo_ABCXYZ)
            - Ziel_Servicegrad (z.B. 0.98, aus config.SERVICE_LEVEL_MATRIX)
            - Z_Wert (Verteilungsquantil zum Ziel-Servicegrad)
            - Verbrauch_Streuung_Sigma (= Verbrauch_Streuung_MAD / 0.6745)
            - Wiederbeschaffungszeit_Monate (= Wiederbeschaffungszeit_Tage / 30)
            - Wurzel_Wiederbeschaffungszeit (sqrt(Wiederbeschaffungszeit_Monate))
            - Verbrauch_Fuer_Reichweite (= 'd', Ø-Monatsverbrauch bzw.
              Planungs-Override aus Phase 2)
            - Safety_Stock (= Z_Wert * Verbrauch_Streuung_Sigma * Wurzel_Wiederbeschaffungszeit)
            - Meldebestand_ROP (= Verbrauch_Fuer_Reichweite * Wiederbeschaffungszeit_Monate + Safety_Stock)
            - Berechnungsbasis (Klartext-Zusammenfassung je Material, z.B.
              "AX -> SL 99% -> z=2.33; sigma aus MAD; L=1.73 Mon.", fuer
              die COO-Nachvollziehbarkeit - analog zum Zweck von
              Trend_Details in Phase 3)
            - SAP_Sicherheitsbestand_Soll, SAP_Meldebestand (nur falls in
              lead_time_df vorhanden - siehe _merge_inputs())
            - Delta_Safety_Stock, Delta_Meldebestand (= unsere Werte minus
              SAP-Werte, nur falls die SAP-Spalten vorhanden sind)

    Raises:
        ValueError: wenn einer der drei Pflicht-DataFrames die erwarteten
            Spalten nicht enthaelt (siehe _merge_inputs()).
    """
    merged = _merge_inputs(coverage_df, trend_df, lead_time_df)
    _check_abc_xyz_consistency(merged, zmlag_df)

    result = merged.copy()

    service_levels_and_z = result["ABCXYZ_Kombiniert"].apply(_resolve_service_level)
    result["Ziel_Servicegrad"] = service_levels_and_z.apply(lambda t: t[0])
    result["Z_Wert"] = service_levels_and_z.apply(lambda t: t[1])

    unresolved = result[result["Ziel_Servicegrad"].isna()]
    if not unresolved.empty:
        print(
            f"  Hinweis (Fruehwarnsystem): {len(unresolved)} Material(ien) "
            f"haben eine ABC/XYZ-Klasse ohne hinterlegten Ziel-Servicegrad "
            f"(siehe config.SERVICE_LEVEL_MATRIX) und werden mit NaN fuer "
            f"Safety_Stock/Meldebestand_ROP ausgewiesen: "
            f"{list(unresolved['Material'])}"
        )

    result["Verbrauch_Streuung_Sigma"] = (
        result["Verbrauch_Streuung_MAD"] / MAD_TO_SIGMA_FACTOR
    )
    result["Wiederbeschaffungszeit_Monate"] = (
        result["Wiederbeschaffungszeit_Tage"] / config.DAYS_PER_MONTH
    )
    result["Wurzel_Wiederbeschaffungszeit"] = result["Wiederbeschaffungszeit_Monate"].apply(
        lambda x: math.sqrt(x) if pd.notna(x) and x >= 0 else float("nan")
    )

    result["Safety_Stock"] = (
        result["Z_Wert"]
        * result["Verbrauch_Streuung_Sigma"]
        * result["Wurzel_Wiederbeschaffungszeit"]
    )
    result["Meldebestand_ROP"] = (
        result["Verbrauch_Fuer_Reichweite"] * result["Wiederbeschaffungszeit_Monate"]
        + result["Safety_Stock"]
    )

    def _build_basis_text(row: pd.Series) -> str:
        if pd.isna(row["Ziel_Servicegrad"]):
            return f"{row['ABCXYZ_Kombiniert']} -> kein Ziel-Servicegrad hinterlegt"
        return (
            f"{row['ABCXYZ_Kombiniert']} -> SL {row['Ziel_Servicegrad']:.0%} -> "
            f"z={row['Z_Wert']:.2f}; sigma aus MAD; "
            f"L={row['Wiederbeschaffungszeit_Monate']:.2f} Mon."
        )

    result["Berechnungsbasis"] = result.apply(_build_basis_text, axis=1)

    if "Sicherheitsbestand_Soll" in result.columns:
        result = result.rename(columns={"Sicherheitsbestand_Soll": "SAP_Sicherheitsbestand_Soll"})
        result["Delta_Safety_Stock"] = result["Safety_Stock"] - result["SAP_Sicherheitsbestand_Soll"]
    if "Meldebestand" in result.columns:
        result = result.rename(columns={"Meldebestand": "SAP_Meldebestand"})
        result["Delta_Meldebestand"] = result["Meldebestand_ROP"] - result["SAP_Meldebestand"]

    return result

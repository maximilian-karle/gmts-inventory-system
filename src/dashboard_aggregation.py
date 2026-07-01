"""
dashboard_aggregation.py
=========================
Gemeinsame Aggregations-/Ranking-Logik fuer die Dashboard-Charts, die
sowohl vom Excel-Dashboard-Reiter (report_builder.py, Reiter 'Dashboard')
als auch vom interaktiven HTML-Dashboard (html_dashboard.py) verwendet
wird.

Historie: Bis 01.07.2026 war diese Logik inhaltsgleich in beiden Modulen
dupliziert (report_builder._aggregate_for_dashboard() /
html_dashboard._aggregate_for_html()). Im Rahmen der Modul-Konsolidierung
(Abschnitt 6 Projektstatus.md, "Code-/Aufbau-Verschlankung") hier
zusammengefuehrt, da es sich - anders als z.B. bei main.py/main_single.py
oder den *_loader.py-Modulen - um reine, zustandslose Aggregationslogik
ohne Modul-spezifische Kopplung handelt; beide Dashboard-Module bleiben
weiterhin unabhaengig voneinander (keine gegenseitigen Imports), importieren
aber gemeinsam dieses schlanke, selbst zustandslose Hilfsmodul.

Achtung Spaltenreihenfolge 'wert_vs_reichweite': enthaelt bewusst
'Materialkurztext' (fuer die Hover-Beschriftung im Plotly-Scatter), da dies
der Obermenge beider bisherigen Versionen entspricht. report_builder.py
benoetigt fuer den nativen Excel-ScatterChart eine FESTE Spaltenreihenfolge
ohne diese Zusatzspalte (positionelle Zellreferenzen) und waehlt die
benoetigten vier Spalten daher explizit am Aufrufort nach.
"""
from __future__ import annotations

import pandas as pd

# Bisher an zwei Stellen als Konstante/Literal gepflegt (report_builder.py:
# .head(15), html_dashboard.py: TOP_N_RISK_MATERIALS = 15) - hier einmalig
# als Default definiert, per Parameter ueberschreibbar.
DEFAULT_TOP_N_RISK_MATERIALS = 15


def aggregate_dashboard_data(
    summary_df: pd.DataFrame,
    top_n_risk: int = DEFAULT_TOP_N_RISK_MATERIALS,
) -> dict[str, pd.DataFrame]:
    """Berechnet die kleinen Aggregations-/Ranking-Tabellen, die als
    Datenquelle fuer die Dashboard-Charts dienen (Excel-Reiter 'Dashboard'
    UND HTML-Dashboard). Eigenstaendige, zustandslose Funktion, damit
    dieselbe Aggregationslogik fuer Einzel- UND konsolidierten Report
    verwendet werden kann (Aufrufer entscheidet, ob summary_df bereits
    eine 'Technologie'-Spalte mit mehreren Werten enthaelt).

    Args:
        summary_df: Ergebnis von executive_summary.build_executive_summary(),
            ggf. zusaetzlich um 'Technologie' ergaenzt (konsolidierter Fall).
        top_n_risk: Anzahl der Materialien in 'top_risiko_materialien'
            (Default 15, siehe DEFAULT_TOP_N_RISK_MATERIALS).

    Returns:
        Dict mit folgenden Eintraegen (jeweils ein DataFrame, leer falls
        die zugrunde liegende Spalte fehlt oder keine gueltigen Werte hat):
            - 'bestandswert_je_technologie': Technologie, Bestandswert_EUR
              (nur im konsolidierten Fall sinnvoll mit >1 Zeile - im
              Einzelreport entsteht hier eine einzeilige Tabelle)
            - 'prioritaet_verteilung': Prioritaet, Anzahl
            - 'trend_verteilung': Trend_Einstufung, Anzahl
            - 'top_risiko_materialien': Material, Materialkurztext,
              Fehlmengen_Wahrscheinlichkeit_Horizont (Top N, absteigend
              sortiert, NUR Zeilen mit vorhandenem Wert)
            - 'wert_vs_reichweite': Material, Materialkurztext,
              Bestandswert_EUR, Reichweite_Monate, ABC-Kennzeichen (fuer
              Scatter/Risikomatrix, NUR Zeilen mit beiden Werten vorhanden)
    """
    result: dict[str, pd.DataFrame] = {}

    if "Technologie" in summary_df.columns:
        result["bestandswert_je_technologie"] = (
            summary_df.groupby("Technologie", as_index=False)["Bestandswert_EUR"]
            .sum()
            .sort_values("Bestandswert_EUR", ascending=False)
        )
    else:
        result["bestandswert_je_technologie"] = pd.DataFrame(
            columns=["Technologie", "Bestandswert_EUR"]
        )

    result["prioritaet_verteilung"] = (
        summary_df["Prioritaet"].value_counts().rename_axis("Prioritaet")
        .reset_index(name="Anzahl")
    )

    if "Trend_Einstufung" in summary_df.columns:
        trend_counts = summary_df["Trend_Einstufung"].dropna()
        result["trend_verteilung"] = (
            trend_counts.value_counts().rename_axis("Trend_Einstufung")
            .reset_index(name="Anzahl")
        )
    else:
        result["trend_verteilung"] = pd.DataFrame(columns=["Trend_Einstufung", "Anzahl"])

    if "Fehlmengen_Wahrscheinlichkeit_Horizont" in summary_df.columns:
        top_risk = summary_df.dropna(subset=["Fehlmengen_Wahrscheinlichkeit_Horizont"])
        top_risk = top_risk.sort_values(
            "Fehlmengen_Wahrscheinlichkeit_Horizont", ascending=False
        ).head(top_n_risk)
        result["top_risiko_materialien"] = top_risk[
            ["Material", "Materialkurztext", "Fehlmengen_Wahrscheinlichkeit_Horizont"]
        ]
    else:
        result["top_risiko_materialien"] = pd.DataFrame(
            columns=["Material", "Materialkurztext", "Fehlmengen_Wahrscheinlichkeit_Horizont"]
        )

    if "Reichweite_Monate" in summary_df.columns:
        scatter = summary_df.dropna(subset=["Bestandswert_EUR", "Reichweite_Monate"])
        result["wert_vs_reichweite"] = scatter[
            ["Material", "Materialkurztext", "Bestandswert_EUR",
             "Reichweite_Monate", "ABC-Kennzeichen"]
        ]
    else:
        result["wert_vs_reichweite"] = pd.DataFrame(
            columns=["Material", "Materialkurztext", "Bestandswert_EUR",
                     "Reichweite_Monate", "ABC-Kennzeichen"]
        )

    return result

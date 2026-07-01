"""
html_dashboard.py
====================
Interaktives Plotly-HTML-Dashboard - Phase 6 des Projektplans, Ergaenzung
zum Excel-Dashboard-Reiter (siehe report_builder.py, Reiter 'Dashboard').

Fachlicher Hintergrund: Max moechte zusaetzlich zum Excel-Workbook eine
eigenstaendige HTML-Datei, die bei jedem Lauf automatisch im Output-Ordner
entsteht - interaktiv (Hover/Zoom/Legendenfilter ueber plotly.express),
unabhaengig oeffenbar (kein Excel, kein Server notwendig).

Designueberarbeitung (29.06.2026, Klaerung mit Max anhand eines Referenz-
Screenshots, Plotly-Dash-Enterprise-Beispiel 'Medicare Provider Charges'):
    - EINE HTML-Datei statt bisher 9 (8 Einzeltechnologien + konsolidiert),
      mit einem Technologie-DROPDOWN oben.
    - KPI-Kacheln oben, Charts im Karten-Layout, Detailtabelle unten.
    - Bewusst REIN CLIENTSEITIG, kein Dash-Server: das Dropdown schaltet
      zwischen BEREITS BEIM ERZEUGEN vorberechneten Panels um (CSS
      display:none/block), kein Live-Nachrechnen einzelner Plotly-Traces.

Erweiterung (30.06.2026, Klaerung mit Max - Anforderungsrunde "Dashboard
uebersichtlicher UND detailreicher"):
    1. Detailtabelle 'Materialdetails' mit CLIENTSEITIGER Filterung
       (Volltextsuche + Dropdowns Prioritaet/ABC/Materialstatus) und
       SORTIERUNG ueber klickbare Spaltenkoepfe. Alle Materialien je
       Technologie werden eingebettet (vorher auf 100 Zeilen gekappt) -
       die Begrenzung ergab mit clientseitiger Filterung keinen Sinn mehr.
    2. Chart 'Top-15 nach Bestandsrisiko (EUR)': verbindet die
       Fehlmengen-Wahrscheinlichkeit mit dem gebundenen Bestandswert
       (Bestandsrisiko_EUR = Fehlmengen_Wahrscheinlichkeit * Bestandswert_EUR)
       je AKTUELL gewaehlter Technologie - ersetzt den frueheren reinen
       Wahrscheinlichkeits-Balken (Wahrscheinlichkeit allein sagt nichts
       ueber die wirtschaftliche Tragweite einer Fehlmenge).
    3. Neue Kachel 'Reichweite-vs-Wiederbeschaffung-Matrix' (Max'
       Referenz-Pivot, Screenshot 30.06.2026): Kreuztabelle
       Wiederbeschaffungszeit (Zeilen) x Reichweite (Spalten) in
       Monatsbuckets, Zellwert = gebundener Bestandswert (EUR). Diagonale
       (Reichweite ~ WBZ) gelb, oberhalb gruen (Reichweite > WBZ ->
       Abbaupotential/Ueberbestand), unterhalb rot (Reichweite < WBZ ->
       Aufbaubedarf/Versorgungsrisiko). Je Zeile 'Potential Abbau'
       (Summe gruen) und 'Potential Aufbau' (Summe rot), unten ein
       'Gesamtpotential' (Abbau - Aufbau). Mit ABC-/XYZ-Filter (Slicer wie
       im Excel-Pivot) - die Matrix wird dafuer als HTML-Tabelle
       CLIENTSEITIG aus eingebetteten Material-Daten neu gerechnet (die
       einzige Stelle mit clientseitigem Recompute; bewusst als eigene
       HTML-Tabelle statt Plotly-Heatmap, um die EUR-Zellwerte und die
       Potential-Spalten exakt wie im Excel-Pivot abzubilden).
    4. Neue KPI-Kachel 'Kapital in Reichweite > Ø': gebundenes Kapital,
       das in ueberdurchschnittlicher Reichweite steckt (pro-rata-Anteil
       oberhalb des technologie-eigenen Durchschnitts) - siehe
       _excess_capital_above_mean().
    5. Globale, technologieunabhaengige Klappabschnitte oben:
       'Kennzahlen erklaert' (Glossar ROI/WIP/Working Capital/Reichweite/
       Wiederbeschaffungszeit/Bestandsrisiko etc.) und 'Anforderungs-
       abdeckung' (Mapping der initialen Anforderungsliste auf den
       aktuellen Abdeckungsstand, ✓/⚠/✗).

Bewusst EIGENSTAENDIGES Modul, NICHT von report_builder.py importiert
(Modul-Unabhaengigkeit). config wird fuer die zentral gepflegten Matrix-
Bucket-Grenzen importiert (REICHWEITE_WBZ_BUCKET_EDGES/-LABELS) - eine
fachliche Annahme, die wie SERVICE_LEVEL_MATRIX nur an EINER Stelle stehen
soll.

Oeffentliche Hauptfunktion:
    write_combined_dashboard(summary_by_label, output_path,
        page_title, all_technologies_label) -> schreibt EINE HTML-Datei.
"""

from __future__ import annotations

import html as _html
import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

import config

# ---------------------------------------------------------------------------
# Workaround: narwhals-Plugin-Discovery auf manchen Windows/Python-3.14-
# Installationen defekt (Praxisproblem bei Max, 29.06.2026)
# ---------------------------------------------------------------------------
# plotly.express nutzt intern 'narwhals'. Bei JEDEM px-Aufruf fragt narwhals
# ueber importlib.metadata.entry_points() alle Pakete nach Plugins ab; in
# Max' venv (OneDrive-Pfad mit Leerzeichen) schlaegt das mit 'OSError:
# [Errno 22]' fehl. Loesung: _discover_entrypoints() defensiv ueberschreiben
# (leere Liste bei OSError). Siehe Projektstatus.md Abschnitt 4g.
try:
    import narwhals.plugins as _narwhals_plugins

    _original_discover_entrypoints = _narwhals_plugins._discover_entrypoints

    def _safe_discover_entrypoints():
        try:
            return _original_discover_entrypoints()
        except OSError:
            from importlib.metadata import EntryPoints
            return EntryPoints([])

    _narwhals_plugins._discover_entrypoints = _safe_discover_entrypoints
except ImportError:
    pass


# Farben der Prioritaets-Kategorien (an PRIORITY_FILL_COLORS in
# report_builder.py angelehnt: Rot=Kritisch, Orange=Erhoeht, Gruen=
# Unauffaellig, Grau=Unbekannt).
PRIORITY_COLORS = {
    "Kritisch": "#D9534F",
    "Erhoeht": "#F0AD4E",
    "Unauffaellig": "#5CB85C",
    "Unbekannt": "#AAAAAA",
}

TOP_N_RISK_MATERIALS = 15

# Akzent-/Karten-Token (Dash-Enterprise-Stil aus dem Referenz-Screenshot).
ACCENT_COLOR = "#2563EB"
ACCENT_COLOR_DARK = "#1E3A8A"
BG_COLOR = "#F1F4F8"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#1F2A33"
MUTED_TEXT_COLOR = "#6B7785"
BORDER_COLOR = "#E2E6EB"

# Matrix-Zellfarben (Karten-/Heatmap-Stil, angelehnt an Max' Excel-Pivot:
# Diagonale gelb, oberhalb gruen, unterhalb rot).
MATRIX_DIAGONAL_BG = "#FFF8E1"
MATRIX_OVER_BG = "#E8F5E9"   # Reichweite > WBZ (Abbaupotential)
MATRIX_UNDER_BG = "#FDECEA"  # Reichweite < WBZ (Aufbaubedarf/Risiko)
MATRIX_OVER_TEXT = "#1B7F3B"
MATRIX_UNDER_TEXT = "#C0392B"

# Fuenf KPI-Kacheln (Reihenfolge = Anordnung). Jede wird aus dem jeweiligen
# summary_df berechnet (siehe _compute_kpis()).
KPI_DEFINITIONS = [
    ("materialien", "Materialien", "#2563EB",
     "Anzahl der Materialien dieser Technologie im aktuellen Lauf."),
    ("bestandswert", "Bestandswert gesamt", "#16A34A",
     "Summe des gebundenen Bestandswerts (Laufender Wert in EUR)."),
    ("reichweite", "Ø Reichweite (Monate)", "#D97706",
     "Durchschnittliche Reichweite = Bestandsmenge / Ø-Monatsverbrauch."),
    ("kapital_ueber_schnitt", "Kapital in Reichweite > Ø", "#7C3AED",
     "Gebundenes Kapital, das in ueberdurchschnittlicher Reichweite steckt "
     "(pro-rata-Anteil oberhalb des Durchschnitts dieser Technologie) - "
     "Indikator fuer Ueberbestand/Abbaupotential."),
    ("kritisch", "Kritische Materialien", "#DC2626",
     "Anzahl Materialien mit Prioritaet 'Kritisch' (hohe Fehlmengen-"
     "Wahrscheinlichkeit bzw. A-Material mit erhoehtem Risiko)."),
]


def _bucket_index(value: float) -> int | None:
    """Ordnet einen Monatswert einem Matrix-Bucket-Index zu (siehe
    config.REICHWEITE_WBZ_BUCKET_EDGES/-LABELS). None, falls value NaN.

    Spiegelt 1:1 die JS-Funktion gmtsBucketIndex() (clientseitige Matrix) -
    bewusst auch in Python verfuegbar, damit die Bucket-Logik testbar ist.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    for i, edge in enumerate(config.REICHWEITE_WBZ_BUCKET_EDGES):
        if value < edge:
            return i
    return len(config.REICHWEITE_WBZ_BUCKET_EDGES)


def _aggregate_for_html(summary_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Berechnet die Aggregations-/Ranking-Tabellen fuer die Charts EINER
    Technologie (bzw. des zusammengefuehrten 'Alle Technologien'-DataFrames).
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

    # Top-N nach Bestandsrisiko (EUR) = Fehlmengen-Wahrscheinlichkeit *
    # Bestandswert: verbindet Fehlmengenrisiko mit wirtschaftlicher
    # Tragweite (Klaerung mit Max, 30.06.2026).
    if (
        "Fehlmengen_Wahrscheinlichkeit_Horizont" in summary_df.columns
        and "Bestandswert_EUR" in summary_df.columns
    ):
        risk = summary_df.dropna(
            subset=["Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"]
        ).copy()
        risk["Bestandsrisiko_EUR"] = (
            risk["Fehlmengen_Wahrscheinlichkeit_Horizont"] * risk["Bestandswert_EUR"]
        )
        risk = risk[risk["Bestandsrisiko_EUR"] > 0]
        risk = risk.sort_values("Bestandsrisiko_EUR", ascending=False).head(
            TOP_N_RISK_MATERIALS
        )
        result["top_risiko_materialien"] = risk[
            ["Material", "Materialkurztext", "Bestandsrisiko_EUR",
             "Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"]
        ]
    else:
        result["top_risiko_materialien"] = pd.DataFrame(
            columns=["Material", "Materialkurztext", "Bestandsrisiko_EUR",
                     "Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"]
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


def _excess_capital_above_mean(summary_df: pd.DataFrame) -> float:
    """Berechnet das gebundene Kapital, das in ueberdurchschnittlicher
    Reichweite steckt (KPI 'Kapital in Reichweite > Ø').

    Methodik (Klaerung mit Max, 30.06.2026): fuer jedes Material mit
    Reichweite > Ø (Durchschnitt DIESER Technologie) wird der pro-rata-
    Kapitalanteil oberhalb des Durchschnitts gezaehlt:

        Ueberschuss_Material = Bestandswert * (Reichweite - Ø) / Reichweite

    Das entspricht "dem in den Reichweite-Monaten ab dem Durchschnitt nach
    oben gebundenen Kapital" - nicht der gesamte Bestandswert ueber-
    durchschnittlicher Materialien, sondern nur der Anteil, der den
    Durchschnitts-Deckungsgrad uebersteigt.

    Returns:
        Kapitalbetrag (EUR) als float; NaN, falls keine Reichweite-Daten
        vorliegen.
    """
    if "Reichweite_Monate" not in summary_df.columns:
        return float("nan")

    sub = summary_df.dropna(subset=["Reichweite_Monate"])
    sub = sub[sub["Reichweite_Monate"] > 0]
    if sub.empty:
        return float("nan")

    mean_rw = sub["Reichweite_Monate"].mean()
    above = sub[sub["Reichweite_Monate"] > mean_rw].dropna(subset=["Bestandswert_EUR"])
    if above.empty:
        return 0.0

    excess = (
        above["Bestandswert_EUR"]
        * (above["Reichweite_Monate"] - mean_rw)
        / above["Reichweite_Monate"]
    ).sum()
    return float(excess)


def _compute_kpis(summary_df: pd.DataFrame) -> dict[str, str]:
    """Berechnet die fuenf KPI-Kachel-Werte (siehe KPI_DEFINITIONS) als
    fertig formatierte Strings."""
    n_materialien = len(summary_df)
    bestandswert_gesamt = summary_df["Bestandswert_EUR"].sum()

    if "Reichweite_Monate" in summary_df.columns:
        reichweite_werte = summary_df["Reichweite_Monate"].dropna()
        reichweite_mittel = reichweite_werte.mean() if len(reichweite_werte) > 0 else float("nan")
    else:
        reichweite_mittel = float("nan")

    kapital_ueber = _excess_capital_above_mean(summary_df)
    n_kritisch = int((summary_df["Prioritaet"] == "Kritisch").sum())

    return {
        "materialien": f"{n_materialien:,}".replace(",", "."),
        "bestandswert": f"{bestandswert_gesamt:,.0f} €".replace(",", "."),
        "reichweite": (
            f"{reichweite_mittel:.1f}" if pd.notna(reichweite_mittel) else "–"
        ),
        "kapital_ueber_schnitt": (
            f"{kapital_ueber:,.0f} €".replace(",", ".") if pd.notna(kapital_ueber) else "–"
        ),
        "kritisch": f"{n_kritisch:,}".replace(",", "."),
    }


def _style_figure(fig: go.Figure) -> go.Figure:
    """Einheitliches, schlankes Karten-Layout fuer alle Plotly-Figuren."""
    fig.update_layout(
        margin=dict(t=48, b=36, l=56, r=24),
        paper_bgcolor=CARD_BG,
        plot_bgcolor=CARD_BG,
        font=dict(family="Inter, Arial, sans-serif", size=12, color=TEXT_COLOR),
        title=dict(font=dict(size=14, color=TEXT_COLOR)),
        legend=dict(font=dict(size=11)),
    )
    fig.update_xaxes(gridcolor=BORDER_COLOR, zeroline=False)
    fig.update_yaxes(gridcolor=BORDER_COLOR, zeroline=False)
    return fig


def _build_figures(
    summary_df: pd.DataFrame, bestandswert_referenz_df: pd.DataFrame | None = None
) -> list[tuple[str, go.Figure]]:
    """Baut die vier Plotly-Figuren fuer EIN Technologie-Panel. Figuren mit
    leerer Aggregation werden NICHT erzeugt (analog zur Excel-Variante)."""
    aggregates = _aggregate_for_html(summary_df)
    figures: list[tuple[str, go.Figure]] = []

    if bestandswert_referenz_df is not None:
        bwt_df = bestandswert_referenz_df
        highlight_label = (
            summary_df["Technologie"].iloc[0] if "Technologie" in summary_df.columns else None
        )
    else:
        bwt_df = aggregates["bestandswert_je_technologie"]
        highlight_label = None

    if len(bwt_df) > 0:
        colors = [
            ACCENT_COLOR_DARK if (highlight_label is not None and tech == highlight_label)
            else ACCENT_COLOR
            for tech in bwt_df["Technologie"]
        ]
        fig = px.bar(
            bwt_df, x="Technologie", y="Bestandswert_EUR",
            title="Bestandswert je Technologie (EUR)",
            labels={"Bestandswert_EUR": "Bestandswert (EUR)"},
            text_auto=".2s",
        )
        fig.update_traces(marker_color=colors)
        figures.append(("Bestandswert je Technologie", _style_figure(fig)))

    prio_df = aggregates["prioritaet_verteilung"]
    if len(prio_df) > 0:
        fig = px.pie(
            prio_df, names="Prioritaet", values="Anzahl",
            title="Prioritaets-Verteilung",
            color="Prioritaet", color_discrete_map=PRIORITY_COLORS,
            hole=0.45,
        )
        fig.update_traces(textinfo="label+percent")
        figures.append(("Prioritaets-Verteilung", _style_figure(fig)))

    risk_df = aggregates["top_risiko_materialien"]
    if len(risk_df) > 0:
        fig = px.bar(
            risk_df.sort_values("Bestandsrisiko_EUR"),
            x="Bestandsrisiko_EUR", y="Material",
            orientation="h",
            title=f"Top-{TOP_N_RISK_MATERIALS} nach Bestandsrisiko (EUR)",
            labels={
                "Bestandsrisiko_EUR": "Bestandsrisiko (EUR) = Wahrscheinlichkeit × Bestandswert",
            },
            color="Fehlmengen_Wahrscheinlichkeit_Horizont",
            color_continuous_scale="Reds",
            hover_data={
                "Materialkurztext": True,
                "Fehlmengen_Wahrscheinlichkeit_Horizont": ":.0%",
                "Bestandswert_EUR": ":,.0f",
                "Bestandsrisiko_EUR": ":,.0f",
            },
        )
        fig.update_layout(coloraxis_colorbar=dict(title="Wahrsch.", tickformat=".0%"))
        figures.append(("Top-Bestandsrisiko", _style_figure(fig)))

    scatter_df = aggregates["wert_vs_reichweite"]
    if len(scatter_df) > 0:
        fig = px.scatter(
            scatter_df, x="Reichweite_Monate", y="Bestandswert_EUR",
            color="ABC-Kennzeichen",
            title="Bestandswert vs. Reichweite",
            labels={
                "Reichweite_Monate": "Reichweite (Monate)",
                "Bestandswert_EUR": "Bestandswert (EUR)",
            },
            hover_data=["Material", "Materialkurztext"],
        )
        figures.append(("Bestandswert vs. Reichweite", _style_figure(fig)))

    return figures


# Spalten der Detail-Tabelle (Reihenfolge = Spaltenreihenfolge). Tupel:
# (df-Spalte, Anzeige-Label, Sortiertyp 'num'|'text'). Nur Spalten, die in
# summary_df vorkommen, werden geschrieben.
TABLE_COLUMNS = [
    ("Material", "Material", "text"),
    ("Materialkurztext", "Bezeichnung", "text"),
    ("ABC-Kennzeichen", "ABC", "text"),
    ("Bestandswert_EUR", "Bestandswert (EUR)", "num"),
    ("Bestandsmenge_Stueck", "Menge (Stk.)", "num"),
    ("Reichweite_Monate", "Reichweite (Mon.)", "num"),
    ("Wiederbeschaffungszeit_Monate", "WBZ (Mon.)", "num"),
    ("MatStatus_Bezeichnung", "Status", "text"),
    ("Fehlmengen_Wahrscheinlichkeit_Horizont", "Fehlmengen-Risiko", "num"),
    ("Working_Capital_Reduktion", "WC-Reduktion (EUR)", "num"),
    ("ROI_Prozent", "ROI", "num"),
    ("Prioritaet", "Priorität", "text"),
]


def _format_table_cell(column: str, value) -> str:
    """Formatiert einen Tabellenwert fuer die Anzeige."""
    if pd.isna(value):
        return "–"
    if column in ("Bestandswert_EUR", "Working_Capital_Reduktion"):
        return f"{value:,.0f} €".replace(",", ".")
    if column in ("Reichweite_Monate", "Wiederbeschaffungszeit_Monate", "Bestandsmenge_Stueck"):
        return f"{value:,.1f}".replace(",", ".")
    if column in ("Fehlmengen_Wahrscheinlichkeit_Horizont", "ROI_Prozent"):
        return f"{value * 100:.0f}%"
    return _html.escape(str(value))


def _sort_value(column: str, value, sort_type: str) -> str:
    """Liefert den maschinenlesbaren Sortierwert (data-sort) fuer eine Zelle.
    Numerische Spalten als roher float-String (NaN -> leer = ans Ende),
    Textspalten als Klartext."""
    if pd.isna(value):
        return ""
    if sort_type == "num":
        return f"{float(value):.6f}"
    return str(value)


def _build_table_html(summary_df: pd.DataFrame, panel_id: str) -> str:
    """Baut die Detail-Tabelle (HTML) fuer EIN Panel - mit Filter-Toolbar
    (Volltextsuche + Prioritaet/ABC/Status) und sortierbaren Spaltenkoepfen.
    Alle Materialien werden eingebettet (clientseitige Filterung), initial
    nach Bestandsrisiko (Wahrscheinlichkeit) absteigend vorsortiert.
    """
    if summary_df.empty:
        return "<p class='empty-hint'>Keine Materialien für diese Technologie.</p>"

    available_columns = [
        (col, label, stype) for col, label, stype in TABLE_COLUMNS
        if col in summary_df.columns
    ]

    # Vorsortierung: hoechstes Fehlmengen-Risiko zuerst (wie bisher) -
    # clientseitige Sortierung kann das anschliessend ueberschreiben.
    if "Fehlmengen_Wahrscheinlichkeit_Horizont" in summary_df.columns:
        sorted_df = summary_df.sort_values(
            "Fehlmengen_Wahrscheinlichkeit_Horizont", ascending=False, na_position="last"
        )
    else:
        sorted_df = summary_df

    table_id = f"table-{panel_id}"

    # --- Filter-Toolbar (Optionen aus den tatsaechlich vorhandenen Werten) ---
    def _options(series_name: str) -> list[str]:
        if series_name not in summary_df.columns:
            return []
        vals = sorted(
            str(v) for v in summary_df[series_name].dropna().unique()
        )
        return vals

    prio_opts = _options("Prioritaet")
    abc_opts = _options("ABC-Kennzeichen")
    status_opts = _options("MatStatus_Bezeichnung")

    def _select(select_id: str, label: str, options: list[str]) -> str:
        if not options:
            return ""
        opt_html = "<option value='__ALLE__'>Alle</option>" + "".join(
            f"<option value='{_html.escape(o)}'>{_html.escape(o)}</option>" for o in options
        )
        return (
            f"<label class='filter-label'>{_html.escape(label)}"
            f"<select id='{select_id}' onchange=\"gmtsFilterTable('{panel_id}')\">"
            f"{opt_html}</select></label>"
        )

    toolbar = (
        "<div class='table-toolbar'>"
        f"<input type='search' id='search-{panel_id}' class='table-search' "
        f"placeholder='Material / Bezeichnung suchen …' "
        f"oninput=\"gmtsFilterTable('{panel_id}')\">"
        f"{_select(f'fprio-{panel_id}', 'Priorität', prio_opts)}"
        f"{_select(f'fabc-{panel_id}', 'ABC', abc_opts)}"
        f"{_select(f'fstatus-{panel_id}', 'Status', status_opts)}"
        f"<span class='table-count' id='count-{panel_id}'></span>"
        "</div>"
    )

    # --- Kopfzeile (sortierbar) ---
    header_cells = []
    for idx, (_col, label, stype) in enumerate(available_columns):
        header_cells.append(
            f"<th data-type='{stype}' "
            f"onclick=\"gmtsSortTable('{table_id}', {idx}, '{stype}', this)\">"
            f"{_html.escape(label)}<span class='sort-ind'></span></th>"
        )
    header_html = "".join(header_cells)

    # --- Datenzeilen ---
    rows_html = []
    for _, row in sorted_df.iterrows():
        search_text = " ".join(
            str(row.get(c, "")) for c in ("Material", "Materialkurztext")
        ).lower()
        data_attrs = (
            f"data-search=\"{_html.escape(search_text)}\" "
            f"data-prio=\"{_html.escape(str(row.get('Prioritaet', '')))}\" "
            f"data-abc=\"{_html.escape(str(row.get('ABC-Kennzeichen', '')))}\" "
            f"data-status=\"{_html.escape(str(row.get('MatStatus_Bezeichnung', '')))}\""
        )
        cells = []
        for col, _label, stype in available_columns:
            value = row[col]
            sort_val = _html.escape(_sort_value(col, value, stype))
            if col == "Prioritaet" and pd.notna(value):
                color = PRIORITY_COLORS.get(value, "#AAAAAA")
                inner = (
                    f"<span class='pill' style='background:{color}22;"
                    f"color:{color};border:1px solid {color}55;'>"
                    f"{_html.escape(str(value))}</span>"
                )
            else:
                inner = _format_table_cell(col, value)
            cells.append(f"<td data-sort='{sort_val}'>{inner}</td>")
        rows_html.append(f"<tr {data_attrs}>{''.join(cells)}</tr>")

    return (
        f"{toolbar}"
        f"<div class='table-wrap'><table id='{table_id}'>"
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody></table></div>"
    )


def _matrix_payload(summary_df: pd.DataFrame) -> tuple[str, list[str], list[str], int, int]:
    """Bereitet die eingebetteten Material-Daten fuer die clientseitige
    Reichweite-vs-Wiederbeschaffung-Matrix vor.

    Returns:
        (json_str, abc_options, xyz_options, n_platzierbar, n_gesamt)
        - json_str: Liste {rw, wbz, val, abc, xyz} als JSON (nur Materialien
          mit Reichweite UND Wiederbeschaffungszeit UND Bestandswert).
        - abc_options/xyz_options: vorhandene Klassen fuer die Slicer.
        - n_platzierbar / n_gesamt: fuer den Abdeckungs-Hinweis unter der
          Matrix (wie viele Materialien ueberhaupt einsortierbar sind).
    """
    n_gesamt = len(summary_df)
    needed = {"Reichweite_Monate", "Wiederbeschaffungszeit_Monate", "Bestandswert_EUR"}
    if not needed.issubset(summary_df.columns):
        return "[]", [], [], 0, n_gesamt

    df = summary_df.dropna(
        subset=["Reichweite_Monate", "Wiederbeschaffungszeit_Monate", "Bestandswert_EUR"]
    )

    records = []
    abc_set: set[str] = set()
    xyz_set: set[str] = set()
    for _, row in df.iterrows():
        abc = row.get("ABC-Kennzeichen")
        abc = str(abc) if pd.notna(abc) else ""
        # XYZ aus dem kombinierten ABCXYZ-Kennzeichen ableiten (2-stellig,
        # z.B. "AX" -> "X"); siehe stock_analysis.classify_risk_by_abc_xyz().
        abcxyz = row.get("ABCXYZ-Kennzeichen")
        xyz = ""
        if isinstance(abcxyz, str) and len(abcxyz) == 2:
            xyz = abcxyz[1]
        if abc:
            abc_set.add(abc)
        if xyz:
            xyz_set.add(xyz)
        records.append({
            "rw": round(float(row["Reichweite_Monate"]), 4),
            "wbz": round(float(row["Wiederbeschaffungszeit_Monate"]), 4),
            "val": float(row["Bestandswert_EUR"]),
            "abc": abc,
            "xyz": xyz,
        })

    return (
        json.dumps(records),
        sorted(abc_set),
        sorted(xyz_set),
        len(records),
        n_gesamt,
    )


def _build_matrix_html(panel_id: str, summary_df: pd.DataFrame) -> str:
    """Baut die Kachel 'Reichweite-vs-Wiederbeschaffung-Matrix' fuer EIN
    Panel: ABC-/XYZ-Slicer, ein Container (clientseitig gefuellt) und ein
    eingebettetes Daten-Skript. Siehe Modul-Docstring Punkt 3."""
    payload, abc_opts, xyz_opts, n_platzierbar, n_gesamt = _matrix_payload(summary_df)

    if n_platzierbar == 0:
        return (
            "<div class='table-card'>"
            "<h2 class='table-title'>Reichweite-vs-Wiederbeschaffung-Matrix</h2>"
            "<p class='empty-hint'>Für diese Technologie liegen keine Materialien "
            "mit gleichzeitig Reichweite UND Wiederbeschaffungszeit vor "
            "(beide Achsen erforderlich).</p></div>"
        )

    def _select(select_id: str, label: str, options: list[str]) -> str:
        opt_html = "<option value='__ALLE__'>(Alle)</option>" + "".join(
            f"<option value='{_html.escape(o)}'>{_html.escape(o)}</option>" for o in options
        )
        return (
            f"<label class='filter-label'>{label}"
            f"<select id='{select_id}' onchange=\"gmtsRenderMatrix('{panel_id}')\">"
            f"{opt_html}</select></label>"
        )

    note = (
        "<p class='table-note'>Zeilen = Wiederbeschaffungszeit, Spalten = Reichweite "
        "(jeweils Monate). Zellwert = gebundener Bestandswert (EUR). "
        "<span class='legend-dot' style='background:" + MATRIX_DIAGONAL_BG + "'></span> Diagonale "
        "(ausgewogen) · <span class='legend-dot' style='background:" + MATRIX_OVER_BG + "'></span> "
        "Reichweite &gt; WBZ → Abbaupotential · "
        "<span class='legend-dot' style='background:" + MATRIX_UNDER_BG + "'></span> "
        "Reichweite &lt; WBZ → Aufbaubedarf/Risiko. "
        f"Einsortierbar: {n_platzierbar} von {n_gesamt} Materialien.</p>"
    )

    data_script = (
        f"<script>window.GMTS_MATRIX=window.GMTS_MATRIX||{{}};"
        f"window.GMTS_MATRIX['{panel_id}']={payload};</script>"
    )

    return (
        "<div class='table-card'>"
        "<h2 class='table-title'>Reichweite-vs-Wiederbeschaffung-Matrix "
        "<span class='info' title='Kreuztabelle Wiederbeschaffungszeit (Zeilen) "
        "vs. Reichweite (Spalten) in Monatsklassen. Grün oberhalb der Diagonale = "
        "Überbestand/Abbaupotential, rot unterhalb = Versorgungsrisiko/Aufbaubedarf.'>ⓘ</span>"
        "</h2>"
        "<div class='matrix-controls'>"
        f"{_select(f'mxabc-{panel_id}', 'ABC', abc_opts)}"
        f"{_select(f'mxxyz-{panel_id}', 'XYZ', xyz_opts)}"
        "</div>"
        f"<div class='matrix-wrap' id='matrix-{panel_id}'></div>"
        f"{note}{data_script}"
        "</div>"
    )


def _build_kpi_html(summary_df: pd.DataFrame) -> str:
    """Baut die KPI-Kachel-Leiste (siehe KPI_DEFINITIONS / _compute_kpis())."""
    kpi_values = _compute_kpis(summary_df)
    tiles = []
    for key, label, color, tooltip in KPI_DEFINITIONS:
        tiles.append(
            f"<div class='kpi-tile' title='{_html.escape(tooltip)}'>"
            f"<div class='kpi-icon' style='background:{color}'></div>"
            f"<div class='kpi-text'>"
            f"<div class='kpi-value'>{kpi_values[key]}</div>"
            f"<div class='kpi-label'>{_html.escape(label)} <span class='info'>ⓘ</span></div>"
            f"</div></div>"
        )
    return f"<div class='kpi-row'>{''.join(tiles)}</div>"


def _build_panel_html(
    label: str,
    panel_id: str,
    summary_df: pd.DataFrame,
    bestandswert_referenz_df: pd.DataFrame | None,
    is_first: bool,
) -> tuple[str, bool]:
    """Baut EIN vollstaendiges Technologie-Panel (KPI-Kacheln + 4 Charts +
    Matrix-Kachel + Detail-Tabelle) als per CSS ein-/ausblendbaren Block."""
    display_style = "" if is_first else "display:none;"

    if summary_df.empty:
        body = (
            "<p class='empty-hint'>Für diese Technologie liegen aktuell "
            "keine Daten vor.</p>"
        )
        return (
            f"<section class='tech-panel' id='{panel_id}' style='{display_style}'>"
            f"{body}</section>",
            False,
        )

    kpi_html = _build_kpi_html(summary_df)
    figures = _build_figures(summary_df, bestandswert_referenz_df=bestandswert_referenz_df)

    chart_blocks = []
    for i, (_chart_title, fig) in enumerate(figures):
        include_js = "inline" if (is_first and i == 0) else False
        chart_html = fig.to_html(full_html=False, include_plotlyjs=include_js)
        chart_blocks.append(f"<div class='chart-card'>{chart_html}</div>")

    matrix_html = _build_matrix_html(panel_id, summary_df)
    table_html = _build_table_html(summary_df, panel_id)

    return (
        f"<section class='tech-panel' id='{panel_id}' style='{display_style}'>"
        f"{kpi_html}"
        f"<div class='chart-grid'>{''.join(chart_blocks)}</div>"
        f"{matrix_html}"
        f"<div class='table-card'>"
        f"<h2 class='table-title'>Materialdetails</h2>"
        f"{table_html}"
        f"</div>"
        f"</section>",
        True,
    )


# ---------------------------------------------------------------------------
# Globale (technologieunabhaengige) Abschnitte: Glossar + Anforderungs-
# abdeckung. Einmal gerendert, ueber allen Panels, als aufklappbare
# <details>-Bloecke (kein JS noetig).
# ---------------------------------------------------------------------------
_GLOSSARY_ITEMS = [
    ("Working Capital (gebundenes Kapital)",
     "Im Bestand gebundenes Kapital = Bestandsmenge × Stückpreis. Dieses "
     "Geld liegt im Lager und steht nicht für anderes zur Verfügung. Je "
     "niedriger bei gesicherter Versorgung, desto besser."),
    ("WIP (Work in Progress / Umlaufbestand)",
     "Wert halbfertiger Erzeugnisse, die sich gerade in der Fertigung "
     "befinden. Hinweis: GMTS weist WIP aktuell NICHT separat aus — der hier "
     "gezeigte Bestandswert umfasst den Lagerbestand (Bestand × Preis), "
     "keine Fertigungsstufen. Als Begriff hier nur zur Einordnung."),
    ("Holding Cost (Lagerhaltungskosten)",
     "Jährlicher Kostensatz auf den Bestandswert (im Projekt 20% p.a.). "
     "Deckt Kapitalkosten, Lagerfläche, Versicherung und Veralterungsrisiko."),
    ("ROI (Kapitalrendite)",
     "Im Projekt definiert als jährliche Holding-Cost-Ersparnis im "
     "Verhältnis zum aktuell gebundenen Kapital (Annual_Savings / "
     "Working_Capital_Aktuell). KEINE klassische ROI-Rechnung gegen "
     "Implementierungskosten."),
    ("Reichweite (Monate)",
     "Bestandsmenge ÷ durchschnittlicher Monatsverbrauch. Gibt an, wie viele "
     "Monate der aktuelle Bestand bei unverändertem Verbrauch reicht."),
    ("Wiederbeschaffungszeit / WBZ (Monate)",
     "Planlieferzeit + WE-Bearbeitungszeit — wie lange Nachschub dauert. "
     "Wird für die Matrix von Tagen in Monate umgerechnet."),
    ("Bestandsrisiko (EUR)",
     "Fehlmengen-Wahrscheinlichkeit × Bestandswert — der erwartete "
     "wertmäßige Risikobeitrag eines Materials. Verbindet 'wie wahrscheinlich "
     "ist eine Fehlmenge' mit 'wie viel Kapital steht dabei auf dem Spiel'."),
    ("Fehlmengen-Wahrscheinlichkeit",
     "Anteil der Monte-Carlo-Simulationsläufe (Baustein A), in denen das "
     "Material im Planungshorizont in eine Fehlmenge läuft."),
    ("Kapital in Reichweite > Ø",
     "Gebundenes Kapital, das in überdurchschnittlicher Reichweite steckt: "
     "pro Material der Anteil oberhalb des Technologie-Durchschnitts "
     "(Bestandswert × (Reichweite − Ø)/Reichweite). Indikator für "
     "Überbestand/Abbaupotential."),
    ("Potential Abbau / Aufbau (Matrix)",
     "Abbau = Bestandskapital oberhalb der Diagonale (Reichweite > WBZ, "
     "Überbestand, freisetzbar). Aufbau = Bestandskapital unterhalb der "
     "Diagonale (Reichweite < WBZ, Versorgungsrisiko). Gesamtpotential = "
     "Abbau − Aufbau."),
]


def _build_glossary_html() -> str:
    items = "".join(
        f"<div class='glossary-item'><div class='glossary-term'>{_html.escape(term)}</div>"
        f"<div class='glossary-def'>{_html.escape(definition)}</div></div>"
        for term, definition in _GLOSSARY_ITEMS
    )
    return (
        "<details class='global-section'>"
        "<summary>Kennzahlen erklärt (ROI, WIP, Working Capital …)</summary>"
        f"<div class='glossary-grid'>{items}</div>"
        "</details>"
    )


# Mapping der initialen Anforderungsliste ('Initiale_Informationen') auf den
# aktuellen Abdeckungsstand. Status: 'ok' (✓), 'partial' (⚠), 'gap' (✗).
_REQUIREMENTS = [
    ("Bestand in Menge und Wert inkl. Stichtag", "ok",
     "KPI Bestandswert + Spalte Menge (ZMLAG €, SE16XXL Stück)."),
    ("Aufteilung nach Bestandsarten (frei / Q-Prüfung / gesperrt)", "gap",
     "Datenlücke: SAP liefert nur den Materialstatus, keine mengenmäßige "
     "Aufteilung."),
    ("Rahmenverträge / Lieferpläne (Laufzeit, Restmenge, Abruflogik)", "ok",
     "In SAP vorhanden; im Dashboard (noch) nicht dargestellt."),
    ("Offene Bestellungen / bestätigte Liefertermine", "partial",
     "Quelle vorhanden, Integration offen (Baustein B)."),
    ("Reale Zugänge (wann, welche Menge)", "gap",
     "Derzeit nicht möglich (nur bestätigter Liefertermin abbildbar)."),
    ("Verbrauch 36 Monate + Trends/Sondereffekte", "ok",
     "MVER-Historie + Trend-Einstufung (Phase 3)."),
    ("Reichweite auf Basis Ø-Verbrauch 12 Monate", "ok",
     "Phase 2, Spalte Reichweite (Monate)."),
    ("Reichweite mit frei editierbarem Planungsfeld", "ok",
     "Verbrauchs-Override je Material im Report."),
    ("Reichweite Stock-only / Stock + bestätigte Zugänge", "partial",
     "Stock-only umgesetzt; Stock + Zugänge offen (Baustein B)."),
    ("Wiederbeschaffungszeit als Dispo-/Risikobasis", "ok",
     "Phase 4; jetzt zusätzlich in der Reichweite-vs-WBZ-Matrix."),
    ("Nachbestellungen seit Listenbeginn 2024", "gap",
     "SAP-Quelle für Nachbestellhistorie noch zu klären."),
    ("Bestandsentwicklung der letzten 2–3 Jahre", "ok",
     "Aus ZMLAG-Stichtagen abgeleitet."),
]

_REQ_BADGE = {
    "ok": ("✓", "#16A34A", "abgedeckt"),
    "partial": ("⚠", "#D97706", "teilweise"),
    "gap": ("✗", "#DC2626", "Datenlücke"),
}


def _build_requirements_html() -> str:
    n_ok = sum(1 for _t, s, _d in _REQUIREMENTS if s == "ok")
    n_partial = sum(1 for _t, s, _d in _REQUIREMENTS if s == "partial")
    n_gap = sum(1 for _t, s, _d in _REQUIREMENTS if s == "gap")

    rows = []
    for term, status, detail in _REQUIREMENTS:
        symbol, color, word = _REQ_BADGE[status]
        rows.append(
            f"<tr><td><span class='req-badge' style='background:{color}1A;"
            f"color:{color};border:1px solid {color}55;'>{symbol} {word}</span></td>"
            f"<td>{_html.escape(term)}</td>"
            f"<td class='req-detail'>{_html.escape(detail)}</td></tr>"
        )

    summary = (
        f"<span class='req-summary'>"
        f"<b style='color:#16A34A'>{n_ok}</b> abgedeckt · "
        f"<b style='color:#D97706'>{n_partial}</b> teilweise · "
        f"<b style='color:#DC2626'>{n_gap}</b> Datenlücken</span>"
    )

    return (
        "<details class='global-section'>"
        "<summary>Anforderungsabdeckung (initiale Anforderungsliste) "
        f"{summary}</summary>"
        "<div class='table-wrap'><table class='req-table'><thead><tr>"
        "<th>Status</th><th>Anforderung</th><th>Anmerkung</th>"
        "</tr></thead><tbody>"
        f"{''.join(rows)}</tbody></table></div>"
        "</details>"
    )


_PAGE_CSS = f"""
:root {{
  --accent: {ACCENT_COLOR};
  --accent-dark: {ACCENT_COLOR_DARK};
  --bg: {BG_COLOR};
  --card-bg: {CARD_BG};
  --text: {TEXT_COLOR};
  --muted: {MUTED_TEXT_COLOR};
  --border: {BORDER_COLOR};
  --mx-diag: {MATRIX_DIAGONAL_BG};
  --mx-over: {MATRIX_OVER_BG};
  --mx-under: {MATRIX_UNDER_BG};
  --mx-over-text: {MATRIX_OVER_TEXT};
  --mx-under-text: {MATRIX_UNDER_TEXT};
}}
* {{ box-sizing: border-box; }}
body {{
  font-family: 'Inter', Arial, Helvetica, sans-serif;
  margin: 0;
  background: var(--bg);
  color: var(--text);
}}
header.topbar {{
  background: var(--card-bg);
  border-bottom: 1px solid var(--border);
  padding: 16px 32px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 12px;
  position: sticky;
  top: 0;
  z-index: 20;
}}
header.topbar h1 {{ font-size: 19px; margin: 0; font-weight: 600; }}
.tech-select-wrap label {{
  font-size: 12px; color: var(--muted); display: block; margin-bottom: 4px;
}}
select#tech-select {{
  font-size: 14px; padding: 8px 12px; border-radius: 6px;
  border: 1px solid var(--border); background: var(--card-bg);
  color: var(--text); min-width: 240px;
}}
main {{ padding: 24px 32px 48px; }}
.info {{ color: var(--muted); font-size: 11px; cursor: help; }}

/* Globale Klappabschnitte */
.global-section {{
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; padding: 4px 18px; margin-bottom: 16px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.global-section > summary {{
  cursor: pointer; font-weight: 600; font-size: 14px; padding: 12px 0;
  list-style: none; display: flex; align-items: center; gap: 10px;
}}
.global-section > summary::-webkit-details-marker {{ display: none; }}
.global-section > summary::before {{ content: '▸'; color: var(--muted); }}
.global-section[open] > summary::before {{ content: '▾'; }}
.req-summary {{ font-weight: 400; font-size: 12px; color: var(--muted); margin-left: auto; }}
.glossary-grid {{
  display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 14px; padding: 8px 0 18px;
}}
.glossary-term {{ font-weight: 600; font-size: 13px; margin-bottom: 3px; }}
.glossary-def {{ font-size: 12.5px; color: var(--muted); line-height: 1.45; }}
.req-table td, .req-table th {{ font-size: 12.5px; }}
.req-badge {{
  display: inline-block; padding: 2px 8px; border-radius: 10px;
  font-size: 11px; font-weight: 600; white-space: nowrap;
}}
.req-detail {{ color: var(--muted); }}

/* KPI-Kacheln */
.kpi-row {{
  display: grid; grid-template-columns: repeat(auto-fit, minmax(168px, 1fr));
  gap: 16px; margin-bottom: 24px;
}}
.kpi-tile {{
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; padding: 18px 20px; display: flex;
  align-items: center; gap: 14px; box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.kpi-icon {{ width: 36px; height: 36px; border-radius: 8px; flex-shrink: 0; }}
.kpi-value {{ font-size: 22px; font-weight: 700; line-height: 1.1; }}
.kpi-label {{ font-size: 12px; color: var(--muted); margin-top: 2px; }}

/* Charts */
.chart-grid {{
  display: grid; grid-template-columns: repeat(2, minmax(360px, 1fr));
  gap: 16px; margin-bottom: 16px;
}}
.chart-card, .table-card {{
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.chart-card {{ padding: 8px; }}
.table-card {{ padding: 20px 24px; margin-bottom: 16px; }}
.table-title {{ font-size: 15px; margin: 0 0 14px; font-weight: 600; }}

/* Tabellen allgemein */
.table-wrap {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
thead th {{
  text-align: left; color: var(--muted); font-weight: 600; font-size: 11px;
  text-transform: uppercase; letter-spacing: 0.03em;
  border-bottom: 1px solid var(--border); padding: 10px 12px; white-space: nowrap;
}}
#table-thead th, table[id^='table-'] thead th {{ cursor: pointer; user-select: none; }}
.sort-ind {{ font-size: 9px; color: var(--accent); margin-left: 4px; }}
tbody td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); white-space: nowrap; }}
tbody tr:hover {{ background: var(--bg); }}
.pill {{
  display: inline-block; padding: 3px 10px; border-radius: 12px;
  font-size: 12px; font-weight: 600;
}}
.table-note {{ font-size: 12px; color: var(--muted); margin-top: 10px; line-height: 1.5; }}
.legend-dot {{
  display: inline-block; width: 11px; height: 11px; border-radius: 3px;
  border: 1px solid var(--border); vertical-align: middle; margin: 0 2px;
}}
.empty-hint {{ color: var(--muted); padding: 40px 0; text-align: center; }}

/* Filter-Toolbar */
.table-toolbar {{
  display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end;
  margin-bottom: 14px;
}}
.table-search {{
  flex: 1 1 240px; min-width: 200px; padding: 8px 12px; font-size: 13px;
  border: 1px solid var(--border); border-radius: 6px; background: var(--card-bg);
  color: var(--text);
}}
.filter-label {{
  display: flex; flex-direction: column; font-size: 11px; color: var(--muted); gap: 4px;
}}
.filter-label select {{
  padding: 7px 10px; font-size: 13px; border: 1px solid var(--border);
  border-radius: 6px; background: var(--card-bg); color: var(--text); min-width: 120px;
}}
.table-count {{ font-size: 12px; color: var(--muted); align-self: center; margin-left: auto; }}

/* Matrix */
.matrix-controls {{ display: flex; gap: 14px; margin-bottom: 12px; flex-wrap: wrap; }}
.matrix-wrap {{ overflow-x: auto; }}
table.matrix {{ border-collapse: collapse; font-size: 12px; min-width: 720px; }}
table.matrix th, table.matrix td {{
  border: 1px solid var(--border); padding: 7px 9px; text-align: right; white-space: nowrap;
}}
table.matrix thead th {{
  text-transform: none; letter-spacing: 0; font-size: 11px; text-align: center;
  background: var(--bg); color: var(--text);
}}
table.matrix th.row-head {{ text-align: left; background: var(--bg); font-weight: 600; }}
table.matrix td.mx-diag {{ background: var(--mx-diag); font-weight: 600; }}
table.matrix td.mx-over {{ background: var(--mx-over); color: var(--mx-over-text); }}
table.matrix td.mx-under {{ background: var(--mx-under); color: var(--mx-under-text); }}
table.matrix td.mx-zero {{ color: #C2C8CF; }}
table.matrix .mx-total {{ font-weight: 600; background: #EEF2F7; }}
table.matrix th.col-abbau {{ background: #16A34A; color: #fff; }}
table.matrix th.col-aufbau {{ background: #DC2626; color: #fff; }}
table.matrix td.col-abbau {{ color: var(--mx-over-text); font-weight: 600; }}
table.matrix td.col-aufbau {{ color: var(--mx-under-text); font-weight: 600; }}
.mx-gesamtpotential {{
  margin-top: 12px; font-size: 13px; display: flex; gap: 24px; flex-wrap: wrap;
}}
.mx-gesamtpotential b {{ font-size: 15px; }}

@media (max-width: 900px) {{
  .chart-grid {{ grid-template-columns: 1fr; }}
}}
"""

# Bucket-Grenzen/-Labels fuer die clientseitige Matrix (aus config, eine
# Quelle der Wahrheit). Als JS-Globals injiziert.
_MATRIX_CONFIG_JS = (
    f"window.GMTS_BUCKET_EDGES={json.dumps(config.REICHWEITE_WBZ_BUCKET_EDGES)};"
    f"window.GMTS_BUCKET_LABELS={json.dumps(config.REICHWEITE_WBZ_BUCKET_LABELS)};"
)

_PAGE_SCRIPT = """
var GMTS_EUR = new Intl.NumberFormat('de-DE', { maximumFractionDigits: 0 });

function gmtsShowTechnology(panelId) {
  document.querySelectorAll('.tech-panel').forEach(function (panel) {
    panel.style.display = (panel.id === panelId) ? '' : 'none';
  });
  // Matrix des nun sichtbaren Panels (neu) rendern, damit Layoutbreiten stimmen.
  gmtsRenderMatrix(panelId);
}

function gmtsBucketIndex(v) {
  if (v === null || v === undefined || isNaN(v)) return -1;
  var edges = window.GMTS_BUCKET_EDGES;
  for (var i = 0; i < edges.length; i++) {
    if (v < edges[i]) return i;
  }
  return edges.length;
}

function gmtsRenderMatrix(panelId) {
  var container = document.getElementById('matrix-' + panelId);
  if (!container) return;
  var data = (window.GMTS_MATRIX || {})[panelId];
  if (!data) return;

  var abcSel = document.getElementById('mxabc-' + panelId);
  var xyzSel = document.getElementById('mxxyz-' + panelId);
  var abc = abcSel ? abcSel.value : '__ALLE__';
  var xyz = xyzSel ? xyzSel.value : '__ALLE__';

  var labels = window.GMTS_BUCKET_LABELS;
  var n = labels.length;
  var grid = [];
  for (var r = 0; r < n; r++) { grid.push(new Array(n).fill(0)); }

  data.forEach(function (rec) {
    if (abc !== '__ALLE__' && rec.abc !== abc) return;
    if (xyz !== '__ALLE__' && rec.xyz !== xyz) return;
    var rwIdx = gmtsBucketIndex(rec.rw);   // Spalte = Reichweite
    var wbzIdx = gmtsBucketIndex(rec.wbz);  // Zeile = Wiederbeschaffungszeit
    if (rwIdx < 0 || wbzIdx < 0) return;
    grid[wbzIdx][rwIdx] += rec.val;
  });

  // Aufbau (Tabelle)
  var html = '<table class="matrix"><thead><tr>';
  html += '<th class="row-head">WBZ \\u2193 / RW \\u2192</th>';
  for (var c = 0; c < n; c++) { html += '<th>' + labels[c] + '</th>'; }
  html += '<th class="mx-total">Gesamt</th>';
  html += '<th class="col-abbau">Pot. Abbau</th>';
  html += '<th class="col-aufbau">Pot. Aufbau</th></tr></thead><tbody>';

  var colTotals = new Array(n).fill(0);
  var grandTotal = 0, totalAbbau = 0, totalAufbau = 0;

  for (var rr = 0; rr < n; rr++) {
    html += '<tr><th class="row-head">' + labels[rr] + '</th>';
    var rowTotal = 0, rowAbbau = 0, rowAufbau = 0;
    for (var cc = 0; cc < n; cc++) {
      var val = grid[rr][cc];
      rowTotal += val; colTotals[cc] += val;
      var cls;
      if (cc === rr) { cls = 'mx-diag'; }
      else if (cc > rr) { cls = 'mx-over'; rowAbbau += val; }   // Reichweite > WBZ
      else { cls = 'mx-under'; rowAufbau += val; }              // Reichweite < WBZ
      if (val === 0) { cls = 'mx-zero'; }
      html += '<td class="' + cls + '">' + (val === 0 ? '\\u2013' : GMTS_EUR.format(val) + ' \\u20ac') + '</td>';
    }
    grandTotal += rowTotal; totalAbbau += rowAbbau; totalAufbau += rowAufbau;
    html += '<td class="mx-total">' + GMTS_EUR.format(rowTotal) + ' \\u20ac</td>';
    html += '<td class="col-abbau">' + (rowAbbau ? GMTS_EUR.format(rowAbbau) + ' \\u20ac' : '\\u2013') + '</td>';
    html += '<td class="col-aufbau">' + (rowAufbau ? GMTS_EUR.format(rowAufbau) + ' \\u20ac' : '\\u2013') + '</td></tr>';
  }

  // Summenzeile
  html += '<tr><th class="row-head mx-total">Gesamt</th>';
  for (var ct = 0; ct < n; ct++) {
    html += '<td class="mx-total">' + GMTS_EUR.format(colTotals[ct]) + ' \\u20ac</td>';
  }
  html += '<td class="mx-total">' + GMTS_EUR.format(grandTotal) + ' \\u20ac</td>';
  html += '<td class="col-abbau">' + GMTS_EUR.format(totalAbbau) + ' \\u20ac</td>';
  html += '<td class="col-aufbau">' + GMTS_EUR.format(totalAufbau) + ' \\u20ac</td></tr>';
  html += '</tbody></table>';

  var netto = totalAbbau - totalAufbau;
  html += '<div class="mx-gesamtpotential">';
  html += '<span>Gesamtpotential Abbau: <b style="color:' + getComputedStyle(document.documentElement).getPropertyValue('--mx-over-text') + '">' + GMTS_EUR.format(totalAbbau) + ' \\u20ac</b></span>';
  html += '<span>Gesamtpotential Aufbau: <b style="color:' + getComputedStyle(document.documentElement).getPropertyValue('--mx-under-text') + '">' + GMTS_EUR.format(totalAufbau) + ' \\u20ac</b></span>';
  html += '<span>Netto-Gesamtpotential: <b>' + GMTS_EUR.format(netto) + ' \\u20ac</b></span>';
  html += '</div>';

  container.innerHTML = html;
}

function gmtsFilterTable(panelId) {
  var table = document.getElementById('table-' + panelId);
  if (!table) return;
  var search = (document.getElementById('search-' + panelId) || {}).value || '';
  search = search.trim().toLowerCase();
  var prioEl = document.getElementById('fprio-' + panelId);
  var abcEl = document.getElementById('fabc-' + panelId);
  var statusEl = document.getElementById('fstatus-' + panelId);
  var prio = prioEl ? prioEl.value : '__ALLE__';
  var abc = abcEl ? abcEl.value : '__ALLE__';
  var status = statusEl ? statusEl.value : '__ALLE__';

  var rows = table.tBodies[0].rows;
  var shown = 0;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var ok = true;
    if (search && row.getAttribute('data-search').indexOf(search) === -1) ok = false;
    if (ok && prio !== '__ALLE__' && row.getAttribute('data-prio') !== prio) ok = false;
    if (ok && abc !== '__ALLE__' && row.getAttribute('data-abc') !== abc) ok = false;
    if (ok && status !== '__ALLE__' && row.getAttribute('data-status') !== status) ok = false;
    row.style.display = ok ? '' : 'none';
    if (ok) shown++;
  }
  var counter = document.getElementById('count-' + panelId);
  if (counter) counter.textContent = shown + ' von ' + rows.length + ' Materialien';
}

function gmtsSortTable(tableId, colIndex, type, thEl) {
  var table = document.getElementById(tableId);
  if (!table) return;
  var tbody = table.tBodies[0];
  var rows = Array.prototype.slice.call(tbody.rows);
  var asc = thEl.getAttribute('data-asc') !== 'true';

  rows.sort(function (a, b) {
    var x = a.cells[colIndex].getAttribute('data-sort');
    var y = b.cells[colIndex].getAttribute('data-sort');
    if (type === 'num') {
      var nx = (x === '' ? NaN : parseFloat(x));
      var ny = (y === '' ? NaN : parseFloat(y));
      if (isNaN(nx)) nx = asc ? Infinity : -Infinity;
      if (isNaN(ny)) ny = asc ? Infinity : -Infinity;
      return asc ? nx - ny : ny - nx;
    }
    x = (x || '').toLowerCase(); y = (y || '').toLowerCase();
    return asc ? x.localeCompare(y) : y.localeCompare(x);
  });
  rows.forEach(function (r) { tbody.appendChild(r); });

  var ths = thEl.parentElement.querySelectorAll('th');
  ths.forEach(function (th) {
    th.removeAttribute('data-asc');
    var ind = th.querySelector('.sort-ind');
    if (ind) ind.textContent = '';
  });
  thEl.setAttribute('data-asc', asc ? 'true' : 'false');
  var indicator = thEl.querySelector('.sort-ind');
  if (indicator) indicator.textContent = asc ? '\\u25B2' : '\\u25BC';
}

document.addEventListener('DOMContentLoaded', function () {
  // Alle Matrizen initial rendern und Tabellenzaehler setzen.
  Object.keys(window.GMTS_MATRIX || {}).forEach(function (pid) {
    gmtsRenderMatrix(pid);
  });
  document.querySelectorAll("table[id^='table-']").forEach(function (t) {
    var pid = t.id.replace('table-', '');
    gmtsFilterTable(pid);
  });
});
"""


def write_combined_dashboard(
    summary_by_label: dict[str, pd.DataFrame],
    output_path: Path,
    page_title: str = "GMTS Dashboard",
    all_technologies_label: str = "Alle Technologien",
) -> None:
    """Schreibt EINE interaktive HTML-Datei mit Technologie-Dropdown,
    globalen Klappabschnitten (Glossar + Anforderungsabdeckung) und je
    Technologie einem vorberechneten Panel (KPI-Kacheln + 4 Charts +
    Reichweite-vs-WBZ-Matrix + filter-/sortierbare Detailtabelle).

    Args:
        summary_by_label: Dict {Anzeigename: summary_df}. Reihenfolge der
            Keys = Dropdown-Reihenfolge; das erste Element ist initial
            sichtbar.
        output_path: Zielpfad (siehe config.CONSOLIDATED_DASHBOARD_HTML_PATH).
        page_title: Seitentitel/Ueberschrift.
        all_technologies_label: Anzeigename des zusammengefuehrten Eintrags
            (bekommt keinen Bestandswert-Vergleichsbalken, siehe unten).

    Tut nichts (Konsolen-Hinweis), falls summary_by_label leer ist oder ALLE
    enthaltenen summary_df leer sind.
    """
    output_path = Path(output_path)

    non_empty_labels = [
        label for label, df in summary_by_label.items() if not df.empty
    ]
    if not summary_by_label or not non_empty_labels:
        print(
            f"  Hinweis: Keine Technologie mit Daten in summary_by_label - "
            f"HTML-Dashboard wird NICHT erzeugt ({output_path})."
        )
        return

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Bestandswert-Vergleichsbasis (alle Einzeltechnologien) fuer den
    # Vergleichsbalken im Einzel-Panel.
    per_technology_dfs = [
        df for label, df in summary_by_label.items()
        if label != all_technologies_label and not df.empty and "Technologie" in df.columns
    ]
    if per_technology_dfs:
        bestandswert_referenz_df = (
            pd.concat(per_technology_dfs, ignore_index=True)
            .groupby("Technologie", as_index=False)["Bestandswert_EUR"]
            .sum()
            .sort_values("Bestandswert_EUR", ascending=False)
        )
    else:
        bestandswert_referenz_df = None

    options_html = []
    panels_html = []
    for i, (label, summary_df) in enumerate(summary_by_label.items()):
        panel_id = f"panel-{i}"
        is_first = (i == 0)
        options_html.append(
            f"<option value='{panel_id}'>{_html.escape(label)}</option>"
        )
        referenz = None if label == all_technologies_label else bestandswert_referenz_df
        panel_html, _has_content = _build_panel_html(
            label, panel_id, summary_df, referenz, is_first
        )
        panels_html.append(panel_html)

    global_sections = _build_glossary_html() + _build_requirements_html()

    html_doc = (
        "<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        f"<title>{_html.escape(page_title)}</title>"
        f"<style>{_PAGE_CSS}</style>"
        "</head><body>"
        "<header class='topbar'>"
        f"<h1>{_html.escape(page_title)}</h1>"
        "<div class='tech-select-wrap'>"
        "<label for='tech-select'>Technologie</label>"
        "<select id='tech-select' onchange='gmtsShowTechnology(this.value)'>"
        f"{''.join(options_html)}"
        "</select></div>"
        "</header>"
        f"<main>{global_sections}{''.join(panels_html)}</main>"
        f"<script>{_MATRIX_CONFIG_JS}</script>"
        f"<script>{_PAGE_SCRIPT}</script>"
        "</body></html>"
    )

    output_path.write_text(html_doc, encoding="utf-8")
    print(f"  HTML-Dashboard gespeichert unter: {output_path}")

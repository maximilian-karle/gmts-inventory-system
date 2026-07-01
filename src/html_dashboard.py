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
    - KPI-Kacheln oben (grosse Zahl + Label).
    - Datentabelle unten (Executive-Summary-Zeilen).
    - Bewusst REIN CLIENTSEITIG, kein Dash-Server: das Dropdown schaltet
      zwischen BEREITS BEIM ERZEUGEN vorberechneten Panels um (CSS
      display:none/block), kein Live-Nachrechnen der Plotly-Charts.

Anforderungsrunde Max (30.06.2026, Abschnitt 4k Projektstatus.md) - sieben
Erweiterungen gegenueber der ersten Fassung, umgesetzt in dieser Version:
    1. Materialdetails-Tabelle filter- und sortierbar (Volltextsuche +
       Prioritaet/ABC/Status-Dropdowns, klickbare Spaltenkoepfe, KEINE
       Zeilenkappung mehr - alle Materialien je Technologie eingebettet).
    2. Chart 'Top-15 nach Bestandsrisiko (EUR)' ersetzt den reinen
       Wahrscheinlichkeits-Balken: Bestandsrisiko_EUR = Fehlmengen_
       Wahrscheinlichkeit_Horizont * Bestandswert_EUR.
    3. Mehr Interaktion: ABC-/XYZ-Slicer an der Matrix, Tabellen-
       Filterleiste - bewusst NUR Tabelle + Matrix clientseitig
       interaktiv, die vier Plotly-Charts bleiben vorberechnet.
    4. Globaler, klappbarer Abschnitt 'Anforderungsabdeckung' (einmal
       gerendert, technologieunabhaengig).
    5. Neue Kachel 'Reichweite-vs-Wiederbeschaffung-Matrix' (Kreuztabelle
       WBZ x Reichweite in Monatsklassen, siehe config.
       REICHWEITE_WBZ_BUCKET_EDGES/-LABELS), mit Potential-Abbau/-Aufbau-
       Spalten und ABC-/XYZ-Slicer.
    6. Neue KPI-Kachel 'Kapital in Reichweite > Ø' (siehe
       _excess_capital_above_mean()).
    7. Globaler, klappbarer Glossar-Abschnitt 'Kennzahlen erklaert' plus
       Tooltips an KPI-Kacheln/Matrix.

Architekturprinzipien bleiben gewahrt: Panels werden weiterhin beim
Erzeugen VORBERECHNET (Plotly-Charts statisch, kein Live-Recompute).
Bewusst EIGENSTAENDIGES Modul, NICHT von report_builder.py importiert
(Modul-Unabhaengigkeit, analoges Prinzip zu trend_analysis.py/
mver_loader.py). Die Aggregationslogik fuer die ERSTEN BEIDEN Charts
(Bestandswert je Technologie, Prioritaets-Verteilung) ist seit 01.07.2026
mit report_builder.py ueber dashboard_aggregation.py geteilt (Modul-
Konsolidierung) - das Bestandsrisiko-Chart, die Matrix und die Tabelle
sind HTML-spezifisch (im Excel-Dashboard nicht vorhanden, siehe
Projektstatus.md Abschnitt 4k) und bleiben lokal in diesem Modul. Neu
importiert wird 'config' (fuer die zentral gepflegten Matrix-Bucket-
Grenzen, siehe config.REICHWEITE_WBZ_BUCKET_EDGES/-LABELS).

Vier Diagramme je Technologie-Panel:
    1. Bestandswert je Technologie (Balken) - im Einzel-Panel NUR der
       Balken der gewaehlten Technologie im Kontext der anderen
       (dunklere Farbe), im 'Alle Technologien'-Panel alle gleichzeitig.
    2. Prioritaets-Verteilung (Kreisdiagramm).
    3. Top-15 nach Bestandsrisiko (EUR) - siehe Punkt 2 oben.
    4. Bestandswert vs. Reichweite (Scatter, "Risikomatrix"), farblich
       nach ABC-Kennzeichen.

Oeffentliche Hauptfunktion:
    write_combined_dashboard(summary_by_label, output_path,
        all_technologies_label) -> schreibt EINE HTML-Datei mit Dropdown
        ueber alle uebergebenen Technologien.
"""

from __future__ import annotations

import html as _html
import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

import config
from dashboard_aggregation import aggregate_dashboard_data

# ---------------------------------------------------------------------------
# Workaround: narwhals-Plugin-Discovery auf manchen Windows/Python-3.14-
# Installationen defekt (Praxisproblem bei Max, 29.06.2026)
# ---------------------------------------------------------------------------
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


PRIORITY_COLORS = {
    "Kritisch": "#D9534F",
    "Erhoeht": "#F0AD4E",
    "Unauffaellig": "#5CB85C",
    "Unbekannt": "#AAAAAA",
}

TOP_N_RISK_MATERIALS = 15

ACCENT_COLOR = "#2563EB"
ACCENT_COLOR_DARK = "#1E3A8A"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#1F2A33"
BORDER_COLOR = "#E2E6EB"

KPI_DEFINITIONS = [
    ("materialien", "Materialien", "#2563EB",
     "Anzahl der Materialien dieser Technologie im aktuellen Lauf."),
    ("bestandswert", "Bestandswert gesamt", "#16A34A",
     "Summe des gebundenen Bestandswerts (Laufender Wert in EUR)."),
    ("reichweite", "Ø Reichweite (Monate)", "#D97706",
     "Durchschnittliche Reichweite = Bestandsmenge / Ø-Monatsverbrauch."),
    ("kapital_ueber_oe", "Kapital in Reichweite > Ø", "#7C3AED",
     "Gebundenes Kapital, das in ueberdurchschnittlicher Reichweite steckt "
     "(pro-rata-Anteil oberhalb des Durchschnitts dieser Technologie) - "
     "Indikator fuer Ueberbestand/Abbaupotential."),
    ("kritisch", "Kritische Materialien", "#DC2626",
     "Anzahl Materialien mit Prioritaet 'Kritisch' (hohe Fehlmengen-"
     "Wahrscheinlichkeit bzw. A-Material mit erhoehtem Risiko)."),
]

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

DEFAULT_SORT_COLUMN = "Bestandswert_EUR"


def _excess_capital_above_mean(summary_df: pd.DataFrame) -> float:
    """Berechnet die KPI 'Kapital in Reichweite > Ø': das gebundene
    Kapital, das PRO-RATA oberhalb des Technologie-Durchschnitts der
    Reichweite gebunden ist.

    Je Material mit Reichweite_Monate > Durchschnitt:
        Bestandswert_EUR * (Reichweite_Monate - Durchschnitt) / Reichweite_Monate
    summiert. Materialien mit Reichweite <= Durchschnitt tragen 0 bei.
    """
    if (
        "Reichweite_Monate" not in summary_df.columns
        or "Bestandswert_EUR" not in summary_df.columns
    ):
        return 0.0

    valid = summary_df.dropna(subset=["Reichweite_Monate", "Bestandswert_EUR"])
    valid = valid[valid["Reichweite_Monate"] > 0]
    if valid.empty:
        return 0.0

    mean_reichweite = valid["Reichweite_Monate"].mean()
    above = valid[valid["Reichweite_Monate"] > mean_reichweite]
    if above.empty:
        return 0.0

    excess_share = (above["Reichweite_Monate"] - mean_reichweite) / above["Reichweite_Monate"]
    return float((above["Bestandswert_EUR"] * excess_share).sum())


def _compute_kpis(summary_df: pd.DataFrame) -> dict[str, str]:
    """Berechnet die fuenf KPI-Kachel-Werte (siehe KPI_DEFINITIONS)."""
    n_materialien = len(summary_df)
    bestandswert_gesamt = summary_df["Bestandswert_EUR"].sum()

    if "Reichweite_Monate" in summary_df.columns:
        reichweite_werte = summary_df["Reichweite_Monate"].dropna()
        reichweite_mittel = reichweite_werte.mean() if len(reichweite_werte) > 0 else float("nan")
    else:
        reichweite_mittel = float("nan")

    kapital_ueber_oe = _excess_capital_above_mean(summary_df)
    n_kritisch = int((summary_df["Prioritaet"] == "Kritisch").sum())

    return {
        "materialien": f"{n_materialien:,}".replace(",", "."),
        "bestandswert": f"{bestandswert_gesamt:,.0f} €".replace(",", "."),
        "reichweite": (
            f"{reichweite_mittel:.1f}" if pd.notna(reichweite_mittel) else "–"
        ),
        "kapital_ueber_oe": f"{kapital_ueber_oe:,.0f} €".replace(",", "."),
        "kritisch": f"{n_kritisch:,}".replace(",", "."),
    }


def _style_figure(fig: go.Figure) -> go.Figure:
    """Wendet ein einheitliches, schlankes Layout auf eine Plotly-Figur an."""
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
    """Baut die vier Plotly-Figuren fuer EIN Technologie-Panel.

    Chart 1+2 nutzen die mit report_builder.py geteilte
    aggregate_dashboard_data(). Chart 3 (Bestandsrisiko EUR) und Chart 4
    sind HTML-spezifisch (im Excel-Dashboard nicht vorhanden).
    """
    aggregates = aggregate_dashboard_data(summary_df)
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

    # --- Chart 3: Top-15 nach Bestandsrisiko (EUR) --------------------------
    risk_cols = {"Material", "Materialkurztext", "Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"}
    if risk_cols.issubset(summary_df.columns):
        risk_base = summary_df.dropna(
            subset=["Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"]
        ).copy()
        if not risk_base.empty:
            risk_base["Bestandsrisiko_EUR"] = (
                risk_base["Fehlmengen_Wahrscheinlichkeit_Horizont"] * risk_base["Bestandswert_EUR"]
            )
            top_risk = risk_base.sort_values("Bestandsrisiko_EUR", ascending=False).head(
                TOP_N_RISK_MATERIALS
            )
            if len(top_risk) > 0:
                fig = px.bar(
                    top_risk.sort_values("Bestandsrisiko_EUR"),
                    x="Bestandsrisiko_EUR", y="Material",
                    orientation="h",
                    color="Fehlmengen_Wahrscheinlichkeit_Horizont",
                    color_continuous_scale="Reds",
                    title=f"Top-{TOP_N_RISK_MATERIALS} nach Bestandsrisiko (EUR)",
                    labels={
                        "Bestandsrisiko_EUR": "Bestandsrisiko (EUR) = Wahrscheinlichkeit × Bestandswert",
                    },
                    hover_data=["Materialkurztext", "Fehlmengen_Wahrscheinlichkeit_Horizont", "Bestandswert_EUR"],
                )
                fig.update_layout(
                    coloraxis_colorbar=dict(title="Wahrsch.", tickformat=".0%")
                )
                figures.append(("Top-Bestandsrisiko", _style_figure(fig)))

    scatter_df = aggregates["wert_vs_reichweite"]
    if len(scatter_df) > 0:
        fig = px.scatter(
            scatter_df, x="Reichweite_Monate", y="Bestandswert_EUR",
            color="ABC-Kennzeichen",
            title="Bestandswert vs. Reichweite (Risikomatrix)",
            labels={
                "Reichweite_Monate": "Reichweite (Monate)",
                "Bestandswert_EUR": "Bestandswert (EUR)",
            },
            hover_data=["Material", "Materialkurztext"],
        )
        figures.append(("Bestandswert vs. Reichweite", _style_figure(fig)))

    return figures


def _build_matrix_payload(summary_df: pd.DataFrame) -> list[dict]:
    """Baut die Rohdaten fuer die Reichweite-vs-Wiederbeschaffung-Matrix
    als kompakte Liste von Records - wird als JSON in die Seite
    eingebettet und dort clientseitig in Buckets aggregiert
    (gmtsRenderMatrix() in _APP_JS), damit der ABC-/XYZ-Slicer ohne
    Server-Rundreise funktioniert.

    NUR Materialien mit vorhandenem Bestandswert, Reichweite UND
    Wiederbeschaffungszeit werden aufgenommen.
    """
    required = {"Bestandswert_EUR", "Reichweite_Monate", "Wiederbeschaffungszeit_Monate"}
    if not required.issubset(summary_df.columns):
        return []

    valid = summary_df.dropna(
        subset=["Bestandswert_EUR", "Reichweite_Monate", "Wiederbeschaffungszeit_Monate"]
    )
    if valid.empty:
        return []

    records = []
    combined_col = "ABCXYZ-Kennzeichen" if "ABCXYZ-Kennzeichen" in valid.columns else None
    for _, row in valid.iterrows():
        combined = str(row[combined_col]) if combined_col and pd.notna(row[combined_col]) else ""
        abc = combined[0] if len(combined) >= 1 else ""
        xyz = combined[1] if len(combined) >= 2 else ""
        records.append({
            "rw": round(float(row["Reichweite_Monate"]), 4),
            "wbz": round(float(row["Wiederbeschaffungszeit_Monate"]), 4),
            "val": round(float(row["Bestandswert_EUR"]), 2),
            "abc": abc,
            "xyz": xyz,
        })
    return records


def _format_table_cell(column: str, value) -> str:
    """Formatiert einen einzelnen Tabellenwert fuer die ANZEIGE."""
    if pd.isna(value):
        return "–"
    if column in ("Bestandswert_EUR", "Working_Capital_Reduktion"):
        return f"{value:,.0f} €".replace(",", ".")
    if column in ("Bestandsmenge_Stueck", "Reichweite_Monate", "Wiederbeschaffungszeit_Monate"):
        return f"{value:,.1f}".replace(",", ".")
    if column in ("Fehlmengen_Wahrscheinlichkeit_Horizont", "ROI_Prozent"):
        return f"{value * 100:.0f}%"
    return _html.escape(str(value))


def _sort_value(column: str, sort_type: str, value) -> str:
    """Liefert den ROHEN Wert fuer das data-sort-Attribut."""
    if pd.isna(value):
        return ""
    if sort_type == "num":
        return f"{float(value):.6f}"
    return _html.escape(str(value))


def _build_table_html(summary_df: pd.DataFrame, panel_id: str) -> str:
    """Baut die Materialdetails-Tabelle (HTML-Markup) fuer EIN Technologie-
    Panel: Volltextsuche + Prioritaet/ABC/Status-Filter, klickbare
    Spaltenkoepfe, ALLE Materialien eingebettet (keine Zeilenkappung).
    """
    if summary_df.empty:
        return ""

    available_columns = [
        (col, label, sort_type) for col, label, sort_type in TABLE_COLUMNS
        if col in summary_df.columns
    ]

    if DEFAULT_SORT_COLUMN in summary_df.columns:
        sorted_df = summary_df.sort_values(DEFAULT_SORT_COLUMN, ascending=False, na_position="last")
    else:
        sorted_df = summary_df

    def _options_html(values: list[str], all_label: str = "Alle") -> str:
        opts = [f"<option value='__ALLE__'>{_html.escape(all_label)}</option>"]
        for v in values:
            opts.append(f"<option value='{_html.escape(v)}'>{_html.escape(v)}</option>")
        return "".join(opts)

    prio_values = sorted(summary_df["Prioritaet"].dropna().unique()) if "Prioritaet" in summary_df.columns else []
    abc_values = sorted(summary_df["ABC-Kennzeichen"].dropna().unique()) if "ABC-Kennzeichen" in summary_df.columns else []
    status_values = (
        sorted(summary_df["MatStatus_Bezeichnung"].dropna().unique())
        if "MatStatus_Bezeichnung" in summary_df.columns else []
    )

    toolbar = (
        f"<div class='table-toolbar'>"
        f"<input type='search' id='search-{panel_id}' class='table-search' "
        f"placeholder='Material / Bezeichnung suchen …' "
        f"oninput=\"gmtsFilterTable('{panel_id}')\">"
        f"<label class='filter-label'>Priorität"
        f"<select id='fprio-{panel_id}' onchange=\"gmtsFilterTable('{panel_id}')\">"
        f"{_options_html(prio_values)}</select></label>"
        f"<label class='filter-label'>ABC"
        f"<select id='fabc-{panel_id}' onchange=\"gmtsFilterTable('{panel_id}')\">"
        f"{_options_html(abc_values)}</select></label>"
        f"<label class='filter-label'>Status"
        f"<select id='fstatus-{panel_id}' onchange=\"gmtsFilterTable('{panel_id}')\">"
        f"{_options_html(status_values)}</select></label>"
        f"<span class='table-count' id='count-{panel_id}'>{len(summary_df)} von "
        f"{len(summary_df)} Materialien</span>"
        f"</div>"
    )

    header_cells = []
    for idx, (col, label, sort_type) in enumerate(available_columns):
        is_default = (col == DEFAULT_SORT_COLUMN)
        asc_attr = " data-asc='false'" if is_default else ""
        indicator = "▼" if is_default else ""
        header_cells.append(
            f"<th data-type='{sort_type}' "
            f"onclick=\"gmtsSortTable('table-{panel_id}', {idx}, '{sort_type}', this)\""
            f"{asc_attr}>{_html.escape(label)}<span class='sort-ind'>{indicator}</span></th>"
        )

    rows_html = []
    for _, row in sorted_df.iterrows():
        material = str(row.get("Material", ""))
        bezeichnung = str(row.get("Materialkurztext", "")) if pd.notna(row.get("Materialkurztext")) else ""
        search_key = f"{material} {bezeichnung}".strip().lower()
        prio_attr = _html.escape(str(row["Prioritaet"])) if "Prioritaet" in row and pd.notna(row["Prioritaet"]) else ""
        abc_attr = _html.escape(str(row["ABC-Kennzeichen"])) if "ABC-Kennzeichen" in row and pd.notna(row["ABC-Kennzeichen"]) else ""
        status_attr = _html.escape(str(row["MatStatus_Bezeichnung"])) if "MatStatus_Bezeichnung" in row and pd.notna(row["MatStatus_Bezeichnung"]) else ""

        cells = []
        for col, _label, sort_type in available_columns:
            value = row.get(col)
            sort_val = _sort_value(col, sort_type, value)
            if col == "Prioritaet" and pd.notna(value):
                color = PRIORITY_COLORS.get(value, "#AAAAAA")
                cells.append(
                    f"<td data-sort='{sort_val}'><span class='pill' "
                    f"style='background:{color}22;color:{color};border:1px solid {color}55;'>"
                    f"{_html.escape(str(value))}</span></td>"
                )
            else:
                cells.append(f"<td data-sort='{sort_val}'>{_format_table_cell(col, value)}</td>")
        rows_html.append(
            f"<tr data-search='{_html.escape(search_key)}' data-prio='{prio_attr}' "
            f"data-abc='{abc_attr}' data-status='{status_attr}'>{''.join(cells)}</tr>"
        )

    return (
        f"{toolbar}"
        f"<div class='table-wrap'>"
        f"<table id='table-{panel_id}'>"
        f"<thead><tr>{''.join(header_cells)}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody>"
        f"</table></div>"
    )


def _build_matrix_html(panel_id: str) -> str:
    """Baut das statische Geruest (Ueberschrift, ABC-/XYZ-Slicer, leerer
    Container) fuer die Reichweite-vs-Wiederbeschaffung-Matrix EINES
    Panels. Der Inhalt wird clientseitig von gmtsRenderMatrix() aus den
    _build_matrix_payload()-Rohdaten aufgebaut.
    """
    return (
        f"<div class='table-card'>"
        f"<h2 class='table-title'>Reichweite-vs-Wiederbeschaffung-Matrix "
        f"<span class='info' title='Kreuztabelle Wiederbeschaffungszeit (Zeilen) vs. "
        f"Reichweite (Spalten) in Monatsklassen. Grün oberhalb der Diagonale = "
        f"Überbestand/Abbaupotential, rot unterhalb = Versorgungsrisiko/"
        f"Aufbaubedarf.'>ⓘ</span></h2>"
        f"<div class='matrix-controls'>"
        f"<label class='filter-label'>ABC"
        f"<select id='mxabc-{panel_id}' onchange=\"gmtsRenderMatrix('{panel_id}')\">"
        f"<option value='__ALLE__'>(Alle)</option>"
        f"<option value='A'>A</option><option value='B'>B</option><option value='C'>C</option>"
        f"</select></label>"
        f"<label class='filter-label'>XYZ"
        f"<select id='mxxyz-{panel_id}' onchange=\"gmtsRenderMatrix('{panel_id}')\">"
        f"<option value='__ALLE__'>(Alle)</option>"
        f"<option value='X'>X</option><option value='Y'>Y</option><option value='Z'>Z</option>"
        f"</select></label>"
        f"</div>"
        f"<div class='matrix-wrap' id='matrix-{panel_id}'></div>"
        f"</div>"
    )


def _build_kpi_html(summary_df: pd.DataFrame) -> str:
    """Baut die KPI-Kachel-Leiste (HTML-Markup) fuer EIN Technologie-Panel."""
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


# ---------------------------------------------------------------------------
# Glossar - EINMAL global gerendert, technologieunabhaengig.
# ---------------------------------------------------------------------------
_GLOSSARY_ITEMS = [
    ("Working Capital (gebundenes Kapital)",
     "Im Bestand gebundenes Kapital = Bestandsmenge × Stückpreis. Dieses "
     "Geld liegt im Lager und steht nicht für anderes zur Verfügung. Je "
     "niedriger bei gesicherter Versorgung, desto besser."),
    ("WIP (Work in Progress / Umlaufbestand)",
     "Wert halbfertiger Erzeugnisse, die sich gerade in der Fertigung "
     "befinden. Hinweis: GMTS weist WIP aktuell NICHT separat aus — der "
     "hier gezeigte Bestandswert umfasst den Lagerbestand (Bestand × "
     "Preis), keine Fertigungsstufen. Als Begriff hier nur zur Einordnung."),
    ("Holding Cost (Lagerhaltungskosten)",
     "Jährlicher Kostensatz auf den Bestandswert (im Projekt 20% p.a.). "
     "Deckt Kapitalkosten, Lagerfläche, Versicherung und Veralterungsrisiko."),
    ("ROI (Kapitalrendite)",
     "Im Projekt definiert als jährliche Holding-Cost-Ersparnis im "
     "Verhältnis zum aktuell gebundenen Kapital (Annual_Savings / "
     "Working_Capital_Aktuell). KEINE klassische ROI-Rechnung gegen "
     "Implementierungskosten."),
    ("Reichweite (Monate)",
     "Bestandsmenge ÷ durchschnittlicher Monatsverbrauch. Gibt an, wie "
     "viele Monate der aktuelle Bestand bei unverändertem Verbrauch reicht."),
    ("Wiederbeschaffungszeit / WBZ (Monate)",
     "Planlieferzeit + WE-Bearbeitungszeit — wie lange Nachschub dauert. "
     "Wird für die Matrix von Tagen in Monate umgerechnet."),
    ("Bestandsrisiko (EUR)",
     "Fehlmengen-Wahrscheinlichkeit × Bestandswert — der erwartete "
     "wertmäßige Risikobeitrag eines Materials. Verbindet 'wie "
     "wahrscheinlich ist eine Fehlmenge' mit 'wie viel Kapital steht "
     "dabei auf dem Spiel'."),
    ("Fehlmengen-Wahrscheinlichkeit",
     "Anteil der Monte-Carlo-Simulationsläufe (Baustein A), in denen das "
     "Material im Planungshorizont in eine Fehlmenge läuft."),
    ("Kapital in Reichweite > Ø",
     "Gebundenes Kapital, das in überdurchschnittlicher Reichweite "
     "steckt: pro Material der Anteil oberhalb des Technologie-"
     "Durchschnitts (Bestandswert × (Reichweite − Ø)/Reichweite). "
     "Indikator für Überbestand/Abbaupotential."),
    ("Potential Abbau / Aufbau (Matrix)",
     "Abbau = Bestandskapital oberhalb der Diagonale (Reichweite > WBZ, "
     "Überbestand, freisetzbar). Aufbau = Bestandskapital unterhalb der "
     "Diagonale (Reichweite < WBZ, Versorgungsrisiko). Gesamtpotential = "
     "Abbau − Aufbau."),
]

_REQUIREMENT_COVERAGE = [
    ("ok", "Bestand in Menge und Wert inkl. Stichtag",
     "KPI Bestandswert + Spalte Menge (ZMLAG €, SE16XXL Stück)."),
    ("luecke", "Aufteilung nach Bestandsarten (frei / Q-Prüfung / gesperrt)",
     "Datenlücke: SAP liefert nur den Materialstatus, keine mengenmäßige Aufteilung."),
    ("ok", "Rahmenverträge / Lieferpläne (Laufzeit, Restmenge, Abruflogik)",
     "In SAP vorhanden; im Dashboard (noch) nicht dargestellt."),
    ("teilweise", "Offene Bestellungen / bestätigte Liefertermine",
     "Quelle vorhanden, Integration offen (Baustein B)."),
    ("luecke", "Reale Zugänge (wann, welche Menge)",
     "Derzeit nicht möglich (nur bestätigter Liefertermin abbildbar)."),
    ("ok", "Verbrauch 36 Monate + Trends/Sondereffekte",
     "MVER-Historie + Trend-Einstufung (Phase 3)."),
    ("ok", "Reichweite auf Basis Ø-Verbrauch 12 Monate",
     "Phase 2, Spalte Reichweite (Monate)."),
    ("ok", "Reichweite mit frei editierbarem Planungsfeld",
     "Verbrauchs-Override je Material im Report."),
    ("teilweise", "Reichweite Stock-only / Stock + bestätigte Zugänge",
     "Stock-only umgesetzt; Stock + Zugänge offen (Baustein B)."),
    ("ok", "Wiederbeschaffungszeit als Dispo-/Risikobasis",
     "Phase 4; jetzt zusätzlich in der Reichweite-vs-WBZ-Matrix."),
    ("luecke", "Nachbestellungen seit Listenbeginn 2024",
     "SAP-Quelle für Nachbestellhistorie noch zu klären."),
    ("ok", "Bestandsentwicklung der letzten 2–3 Jahre",
     "Aus ZMLAG-Stichtagen abgeleitet."),
]

_REQ_STATUS_STYLE = {
    "ok": ("#16A34A", "✓ abgedeckt"),
    "teilweise": ("#D97706", "⚠ teilweise"),
    "luecke": ("#DC2626", "✗ Datenlücke"),
}


def _build_glossary_html() -> str:
    """Baut den globalen, klappbaren Glossar-Abschnitt (EINMAL, nicht je
    Panel) - siehe _GLOSSARY_ITEMS."""
    items = []
    for term, definition in _GLOSSARY_ITEMS:
        items.append(
            f"<div class='glossary-item'>"
            f"<div class='glossary-term'>{_html.escape(term)}</div>"
            f"<div class='glossary-def'>{_html.escape(definition)}</div>"
            f"</div>"
        )
    return (
        "<details class='global-section'>"
        "<summary>Kennzahlen erklärt (ROI, WIP, Working Capital …)</summary>"
        f"<div class='glossary-grid'>{''.join(items)}</div>"
        "</details>"
    )


def _build_requirement_coverage_html() -> str:
    """Baut den globalen, klappbaren Abschnitt 'Anforderungsabdeckung'
    (EINMAL, nicht je Panel) - siehe _REQUIREMENT_COVERAGE."""
    n_ok = sum(1 for status, *_ in _REQUIREMENT_COVERAGE if status == "ok")
    n_teilweise = sum(1 for status, *_ in _REQUIREMENT_COVERAGE if status == "teilweise")
    n_luecke = sum(1 for status, *_ in _REQUIREMENT_COVERAGE if status == "luecke")

    rows = []
    for status, requirement, detail in _REQUIREMENT_COVERAGE:
        color, badge_label = _REQ_STATUS_STYLE[status]
        rows.append(
            f"<tr><td><span class='req-badge' style='background:{color}1A;"
            f"color:{color};border:1px solid {color}55;'>{badge_label}</span></td>"
            f"<td>{_html.escape(requirement)}</td>"
            f"<td class='req-detail'>{_html.escape(detail)}</td></tr>"
        )

    summary_span = (
        "<span class='req-summary'>"
        f"<b style='color:#16A34A'>{n_ok}</b> abgedeckt · "
        f"<b style='color:#D97706'>{n_teilweise}</b> teilweise · "
        f"<b style='color:#DC2626'>{n_luecke}</b> Datenlücken</span>"
    )

    return (
        "<details class='global-section'>"
        f"<summary>Anforderungsabdeckung (initiale Anforderungsliste) {summary_span}</summary>"
        "<div class='table-wrap'><table class='req-table'>"
        "<thead><tr><th>Status</th><th>Anforderung</th><th>Anmerkung</th></tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
        "</details>"
    )


def _build_panel_html(
    label: str,
    panel_id: str,
    summary_df: pd.DataFrame,
    bestandswert_referenz_df: pd.DataFrame | None,
    is_first: bool,
) -> tuple[str, bool]:
    """Baut EIN vollstaendiges Technologie-Panel (KPI-Kacheln + 4 Charts +
    Matrix + Materialdetails-Tabelle)."""
    display_style = "" if is_first else "display:none;"

    if summary_df.empty:
        return (
            f"<section class='tech-panel' id='{panel_id}' style='{display_style}'>"
            f"<div class='empty-hint'>Keine Daten für {_html.escape(label)}.</div>"
            f"</section>",
            False,
        )

    kpi_html = _build_kpi_html(summary_df)
    figures = _build_figures(summary_df, bestandswert_referenz_df)
    chart_cards = []
    for i, (_title, fig) in enumerate(figures):
        div_id = f"chart-{panel_id}-{i}"
        fig_html = fig.to_html(
            full_html=False, include_plotlyjs=False, div_id=div_id,
            config={"displaylogo": False, "responsive": True},
        )
        chart_cards.append(f"<div class='chart-card'>{fig_html}</div>")

    matrix_html = _build_matrix_html(panel_id)
    table_html = _build_table_html(summary_df, panel_id)
    table_section = (
        f"<div class='table-card'><h2 class='table-title'>Materialdetails</h2>"
        f"{table_html}</div>" if table_html else ""
    )

    panel = (
        f"<section class='tech-panel' id='{panel_id}' style='{display_style}'>"
        f"{kpi_html}"
        f"<div class='chart-grid'>{''.join(chart_cards)}</div>"
        f"{matrix_html}"
        f"{table_section}"
        f"</section>"
    )
    return panel, True


_PAGE_CSS = """
:root {
  --accent: #2563EB;
  --accent-dark: #1E3A8A;
  --bg: #F1F4F8;
  --card-bg: #FFFFFF;
  --text: #1F2A33;
  --muted: #6B7785;
  --border: #E2E6EB;
  --mx-diag: #FFF8E1;
  --mx-over: #E8F5E9;
  --mx-under: #FDECEA;
  --mx-over-text: #1B7F3B;
  --mx-under-text: #C0392B;
}
* { box-sizing: border-box; }
body {
  font-family: 'Inter', Arial, Helvetica, sans-serif;
  margin: 0;
  background: var(--bg);
  color: var(--text);
}
header.topbar {
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
}
header.topbar h1 { font-size: 19px; margin: 0; font-weight: 600; }
.tech-select-wrap label {
  font-size: 12px; color: var(--muted); display: block; margin-bottom: 4px;
}
select#tech-select {
  font-size: 14px; padding: 8px 12px; border-radius: 6px;
  border: 1px solid var(--border); background: var(--card-bg);
  color: var(--text); min-width: 240px;
}
main { padding: 24px 32px 48px; }
.info { color: var(--muted); font-size: 11px; cursor: help; }

/* Globale Klappabschnitte */
.global-section {
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; padding: 4px 18px; margin-bottom: 16px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}
.global-section > summary {
  cursor: pointer; font-weight: 600; font-size: 14px; padding: 12px 0;
  list-style: none; display: flex; align-items: center; gap: 10px;
}
.global-section > summary::-webkit-details-marker { display: none; }
.global-section > summary::before { content: '▸'; color: var(--muted); }
.global-section[open] > summary::before { content: '▾'; }
.req-summary { font-weight: 400; font-size: 12px; color: var(--muted); margin-left: auto; }
.glossary-grid {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 14px; padding: 8px 0 18px;
}
.glossary-term { font-weight: 600; font-size: 13px; margin-bottom: 3px; }
.glossary-def { font-size: 12.5px; color: var(--muted); line-height: 1.45; }
.req-table td, .req-table th { font-size: 12.5px; }
.req-badge {
  display: inline-block; padding: 2px 8px; border-radius: 10px;
  font-size: 11px; font-weight: 600; white-space: nowrap;
}
.req-detail { color: var(--muted); }

/* KPI-Kacheln */
.kpi-row {
  display: grid; grid-template-columns: repeat(auto-fit, minmax(168px, 1fr));
  gap: 16px; margin-bottom: 24px;
}
.kpi-tile {
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; padding: 18px 20px; display: flex;
  align-items: center; gap: 14px; box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}
.kpi-icon { width: 36px; height: 36px; border-radius: 8px; flex-shrink: 0; }
.kpi-value { font-size: 22px; font-weight: 700; line-height: 1.1; }
.kpi-label { font-size: 12px; color: var(--muted); margin-top: 2px; }

/* Charts */
.chart-grid {
  display: grid; grid-template-columns: repeat(2, minmax(360px, 1fr));
  gap: 16px; margin-bottom: 16px;
}
.chart-card, .table-card {
  background: var(--card-bg); border: 1px solid var(--border);
  border-radius: 8px; box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}
.chart-card { padding: 8px; }
.table-card { padding: 20px 24px; margin-bottom: 16px; }
.table-title { font-size: 15px; margin: 0 0 14px; font-weight: 600; }

/* Tabellen allgemein */
.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
thead th {
  text-align: left; color: var(--muted); font-weight: 600; font-size: 11px;
  text-transform: uppercase; letter-spacing: 0.03em;
  border-bottom: 1px solid var(--border); padding: 10px 12px; white-space: nowrap;
}
#table-thead th, table[id^='table-'] thead th { cursor: pointer; user-select: none; }
.sort-ind { font-size: 9px; color: var(--accent); margin-left: 4px; }
tbody td { padding: 10px 12px; border-bottom: 1px solid var(--border); white-space: nowrap; }
tbody tr:hover { background: var(--bg); }
.pill {
  display: inline-block; padding: 3px 10px; border-radius: 12px;
  font-size: 12px; font-weight: 600;
}
.table-note { font-size: 12px; color: var(--muted); margin-top: 10px; line-height: 1.5; }
.legend-dot {
  display: inline-block; width: 11px; height: 11px; border-radius: 3px;
  border: 1px solid var(--border); vertical-align: middle; margin: 0 2px;
}
.empty-hint { color: var(--muted); padding: 40px 0; text-align: center; }

/* Filter-Toolbar */
.table-toolbar {
  display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end;
  margin-bottom: 14px;
}
.table-search {
  flex: 1 1 240px; min-width: 200px; padding: 8px 12px; font-size: 13px;
  border: 1px solid var(--border); border-radius: 6px; background: var(--card-bg);
  color: var(--text);
}
.filter-label {
  display: flex; flex-direction: column; font-size: 11px; color: var(--muted); gap: 4px;
}
.filter-label select {
  padding: 7px 10px; font-size: 13px; border: 1px solid var(--border);
  border-radius: 6px; background: var(--card-bg); color: var(--text); min-width: 120px;
}
.table-count { font-size: 12px; color: var(--muted); align-self: center; margin-left: auto; }

/* Matrix */
.matrix-controls { display: flex; gap: 14px; margin-bottom: 12px; flex-wrap: wrap; }
.matrix-wrap { overflow-x: auto; }
table.matrix { border-collapse: collapse; font-size: 12px; min-width: 720px; }
table.matrix th, table.matrix td {
  border: 1px solid var(--border); padding: 7px 9px; text-align: right; white-space: nowrap;
}
table.matrix thead th {
  text-transform: none; letter-spacing: 0; font-size: 11px; text-align: center;
  background: var(--bg); color: var(--text);
}
table.matrix th.row-head { text-align: left; background: var(--bg); font-weight: 600; }
table.matrix td.mx-diag { background: var(--mx-diag); font-weight: 600; }
table.matrix td.mx-over { background: var(--mx-over); color: var(--mx-over-text); }
table.matrix td.mx-under { background: var(--mx-under); color: var(--mx-under-text); }
table.matrix td.mx-zero { color: #C2C8CF; }
table.matrix .mx-total { font-weight: 600; background: #EEF2F7; }
table.matrix th.col-abbau { background: #16A34A; color: #fff; }
table.matrix th.col-aufbau { background: #DC2626; color: #fff; }
table.matrix td.col-abbau { color: var(--mx-over-text); font-weight: 600; }
table.matrix td.col-aufbau { color: var(--mx-under-text); font-weight: 600; }
.mx-gesamtpotential {
  margin-top: 12px; font-size: 13px; display: flex; gap: 24px; flex-wrap: wrap;
}
.mx-gesamtpotential b { font-size: 15px; }

@media (max-width: 900px) {
  .chart-grid { grid-template-columns: 1fr; }
}
"""

_APP_JS = r"""
var GMTS_EUR = new Intl.NumberFormat('de-DE', { maximumFractionDigits: 0 });

function gmtsShowTechnology(panelId) {
  document.querySelectorAll('.tech-panel').forEach(function (panel) {
    panel.style.display = (panel.id === panelId) ? '' : 'none';
  });
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
    var rwIdx = gmtsBucketIndex(rec.rw);
    var wbzIdx = gmtsBucketIndex(rec.wbz);
    if (rwIdx < 0 || wbzIdx < 0) return;
    grid[wbzIdx][rwIdx] += rec.val;
  });

  var html = '<table class="matrix"><thead><tr>';
  html += '<th class="row-head">WBZ \u2193 / RW \u2192</th>';
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
      else if (cc > rr) { cls = 'mx-over'; rowAbbau += val; }
      else { cls = 'mx-under'; rowAufbau += val; }
      if (val === 0) { cls = 'mx-zero'; }
      html += '<td class="' + cls + '">' + (val === 0 ? '\u2013' : GMTS_EUR.format(val) + ' \u20ac') + '</td>';
    }
    grandTotal += rowTotal; totalAbbau += rowAbbau; totalAufbau += rowAufbau;
    html += '<td class="mx-total">' + GMTS_EUR.format(rowTotal) + ' \u20ac</td>';
    html += '<td class="col-abbau">' + (rowAbbau ? GMTS_EUR.format(rowAbbau) + ' \u20ac' : '\u2013') + '</td>';
    html += '<td class="col-aufbau">' + (rowAufbau ? GMTS_EUR.format(rowAufbau) + ' \u20ac' : '\u2013') + '</td></tr>';
  }

  html += '<tr><th class="row-head mx-total">Gesamt</th>';
  for (var ct = 0; ct < n; ct++) {
    html += '<td class="mx-total">' + GMTS_EUR.format(colTotals[ct]) + ' \u20ac</td>';
  }
  html += '<td class="mx-total">' + GMTS_EUR.format(grandTotal) + ' \u20ac</td>';
  html += '<td class="col-abbau">' + GMTS_EUR.format(totalAbbau) + ' \u20ac</td>';
  html += '<td class="col-aufbau">' + GMTS_EUR.format(totalAufbau) + ' \u20ac</td></tr>';
  html += '</tbody></table>';

  var netto = totalAbbau - totalAufbau;
  html += '<div class="mx-gesamtpotential">';
  html += '<span>Gesamtpotential Abbau: <b style="color:' + getComputedStyle(document.documentElement).getPropertyValue('--mx-over-text') + '">' + GMTS_EUR.format(totalAbbau) + ' \u20ac</b></span>';
  html += '<span>Gesamtpotential Aufbau: <b style="color:' + getComputedStyle(document.documentElement).getPropertyValue('--mx-under-text') + '">' + GMTS_EUR.format(totalAufbau) + ' \u20ac</b></span>';
  html += '<span>Netto-Gesamtpotential: <b>' + GMTS_EUR.format(netto) + ' \u20ac</b></span>';
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
  if (indicator) indicator.textContent = asc ? '\u25B2' : '\u25BC';
}

document.addEventListener('DOMContentLoaded', function () {
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
    """Schreibt EINE interaktive HTML-Datei mit einem Technologie-Dropdown
    oben, das zwischen vorberechneten Panels (KPI-Kacheln + 4 Charts +
    Matrix + Detail-Tabelle je Technologie) umschaltet.

    Globale, technologieunabhaengige Abschnitte (Glossar,
    Anforderungsabdeckung) werden EINMAL oberhalb aller Panels gerendert.

    Args:
        summary_by_label: Dict {Anzeigename: summary_df}, EINE Zeile pro
            Technologie. Reihenfolge der Dict-Keys bestimmt die Dropdown-
            Reihenfolge; das erste Element ist beim Oeffnen sichtbar.
        output_path: Zielpfad fuer die HTML-Datei.
        page_title: Seitentitel/Ueberschrift in der Kopfzeile.
        all_technologies_label: Anzeigename fuer den zusammengefuehrten
            'Alle Technologien'-Eintrag, falls in summary_by_label
            vorhanden.

    Erstellt das Zielverzeichnis, falls es noch nicht existiert. Tut
    nichts (gibt nur einen Konsolen-Hinweis aus), falls summary_by_label
    leer ist oder ALLE enthaltenen summary_df leer sind.
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
    matrix_data_by_panel: dict[str, list[dict]] = {}

    for i, (label, summary_df) in enumerate(summary_by_label.items()):
        panel_id = f"panel-{i}"
        is_first = (i == 0)
        options_html.append(
            f"<option value='{panel_id}'>{_html.escape(label)}</option>"
        )
        referenz = None if label == all_technologies_label else bestandswert_referenz_df
        panel_html, has_content = _build_panel_html(
            label, panel_id, summary_df, referenz, is_first
        )
        panels_html.append(panel_html)
        if has_content:
            matrix_data_by_panel[panel_id] = _build_matrix_payload(summary_df)

    # Plotly.js wird EINMAL global eingebunden (jeder Chart-Aufruf nutzt
    # include_plotlyjs=False) - eingebettet statt CDN, damit die Datei
    # OHNE Internetverbindung funktioniert.
    import plotly.offline as pyo
    plotly_js = pyo.get_plotlyjs()

    matrix_data_json = json.dumps(matrix_data_by_panel, ensure_ascii=False)
    bucket_edges_json = json.dumps(config.REICHWEITE_WBZ_BUCKET_EDGES)
    bucket_labels_json = json.dumps(config.REICHWEITE_WBZ_BUCKET_LABELS, ensure_ascii=False)

    data_script = (
        f"window.GMTS_BUCKET_EDGES={bucket_edges_json};"
        f"window.GMTS_BUCKET_LABELS={bucket_labels_json};"
        f"window.GMTS_MATRIX={matrix_data_json};"
    )

    html_doc = (
        "<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1'>"
        f"<title>{_html.escape(page_title)}</title>"
        f"<style>{_PAGE_CSS}</style>"
        f"<script>{plotly_js}</script>"
        "</head><body>"
        "<header class='topbar'>"
        f"<h1>{_html.escape(page_title)}</h1>"
        "<div class='tech-select-wrap'>"
        "<label for='tech-select'>Technologie</label>"
        "<select id='tech-select' onchange='gmtsShowTechnology(this.value)'>"
        f"{''.join(options_html)}"
        "</select></div>"
        "</header>"
        "<main>"
        f"{_build_glossary_html()}"
        f"{_build_requirement_coverage_html()}"
        f"{''.join(panels_html)}"
        "</main>"
        f"<script>{data_script}</script>"
        f"<script>{_APP_JS}</script>"
        "</body></html>"
    )

    output_path.write_text(html_doc, encoding="utf-8")
    print(f"  HTML-Dashboard gespeichert unter: {output_path}")

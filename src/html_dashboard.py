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
      mit einem Technologie-DROPDOWN oben - ersetzt die vorherige
      Datei-pro-Technologie-Struktur vollstaendig.
    - KPI-Kacheln oben (grosse Zahl + Label, analog zum Referenz-Screenshot:
      'Hospitals'/'Procedures'/'Avg Charge'/'Total Discharges' dort ->
      hier 'Materialien'/'Bestandswert gesamt'/'Ø Reichweite'/
      'Kritische Materialien', siehe KPI_DEFINITIONS).
    - Datentabelle unten (Executive-Summary-Zeilen, sortiert nach Risiko)
      analog zur 'Provider Charges Details'-Tabelle im Referenz-Screenshot.
    - Bewusst REIN CLIENTSEITIG, kein Dash-Server (Klaerung mit Max,
      29.06.2026: "auch wenn es einmal ausgewertet worden ist, finde ich
      das eine ganz gute Loesung" - das Dropdown schaltet zwischen BEREITS
      BEIM ERZEUGEN vorberechneten Ansichten um, kein Live-Nachrechnen).
      Dazu wird fuer JEDE Technologie ein vollstaendiges, fertiges Panel
      (KPI-Kacheln + 4 Charts + Tabelle) vorgebaut; ein kleines Stueck
      JavaScript blendet beim Dropdown-Wechsel zwischen den Panels um
      (CSS display:none/block) - kein Live-Aufbau einzelner Plotly-Traces
      aus JSON im Browser, das waere fehleranfaelliger und schwerer zu
      pruefen als vollstaendig in Python vorbereitete, statische Panels.

Bewusst EIGENSTAENDIGES Modul, NICHT von report_builder.py importiert
(Modul-Unabhaengigkeit, analoges Prinzip zu trend_analysis.py/
mver_loader.py - siehe dortige Docstrings): die Aggregationslogik
(_aggregate_for_html()) ist inhaltlich identisch zu report_builder.
_aggregate_for_dashboard(), aber bewusst hier eigenstaendig dupliziert.

Vier Diagramme je Technologie-Panel (inhaltlich wie zuvor, jetzt visuell
ueberarbeitet im Kachel-/Karten-Stil):
    1. Bestandswert je Technologie (Balken) - im Einzel-Panel zeigt dies
       NUR den einen Balken der gewaehlten Technologie zum Vergleich mit
       einer grauen Referenzlinie (Durchschnitt aller Technologien); im
       'Alle Technologien'-Panel alle Technologien im Vergleich.
    2. Prioritaets-Verteilung (Kreisdiagramm)
    3. Top-15-Fehlmengen-Risiko (Balken) - INNERHALB der gewaehlten
       Technologie (siehe Projektstatus.md Abschnitt 3a/4g - im
       Einzel-Panel technologie-spezifisch, im 'Alle Technologien'-Panel
       global ueber alle Technologien, wie bereits zuvor).
    4. Bestandswert vs. Reichweite (Scatter, "Risikomatrix"), farblich
       nach ABC-Kennzeichen.

Oeffentliche Hauptfunktion:
    write_combined_dashboard(summary_by_label, output_path,
        all_technologies_label) -> schreibt EINE HTML-Datei mit Dropdown
        ueber alle uebergebenen Technologien (+ optional 'Alle
        Technologien' als zusaetzliche Dropdown-Option). Ersetzt die
        vorherige write_html_dashboard()-Funktion (ein DataFrame, eine
        Datei) - siehe Aenderungshistorie in Projektstatus.md Abschnitt 4h.
"""

from __future__ import annotations

import html as _html
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# Workaround: narwhals-Plugin-Discovery auf manchen Windows/Python-3.14-
# Installationen defekt (Praxisproblem bei Max, 29.06.2026)
# ---------------------------------------------------------------------------
# plotly.express nutzt intern 'narwhals', um verschiedene DataFrame-Typen
# einheitlich zu behandeln. Bei JEDEM px.bar()/px.pie()/px.scatter()-Aufruf
# fragt narwhals ueber importlib.metadata.entry_points() ALLE installierten
# Pakete nach optionalen Plugins ab. In Max' venv (Pfad unter OneDrive mit
# Leerzeichen/Sonderzeichen) schlaegt das mit 'OSError: [Errno 22] Invalid
# argument' fehl. Ein Fix auf Aufrufer-Seite (z.B. anderer Eingabetyp) wirkt
# NICHT, da der fehlerhafte Codepfad unabhaengig vom Eingabetyp bei jedem
# Aufruf erreicht wird (siehe Projektstatus.md Abschnitt 4g fuer die
# vollstaendige Fehleranalyse inkl. gescheitertem ersten Loesungsversuch).
#
# Loesung: narwhals.plugins._discover_entrypoints() wird hier defensiv
# ueberschrieben - bei einem OSError wird eine LEERE EntryPoints-Liste
# zurueckgegeben statt die Exception weiterzureichen. GMTS nutzt kein
# narwhals-Drittanbieter-Plugin, eine leere Liste ist inhaltlich folgenlos.
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


# Dieselben Schwellen wie in executive_summary.py fuer die Farbgebung der
# Prioritaets-Kategorien - hier separat gepflegt (keine Modul-Abhaengigkeit,
# siehe Modul-Docstring), inhaltlich an PRIORITY_FILL_COLORS in
# report_builder.py angelehnt (Rot=Kritisch, Orange=Erhoeht, Gruen=
# Unauffaellig, Grau=Unbekannt).
PRIORITY_COLORS = {
    "Kritisch": "#D9534F",
    "Erhoeht": "#F0AD4E",
    "Unauffaellig": "#5CB85C",
    "Unbekannt": "#AAAAAA",
}

TOP_N_RISK_MATERIALS = 15

# Hauptakzentfarbe (Dash-Enterprise-Blau aus dem Referenz-Screenshot) sowie
# abgeleitete Token fuer den Kachel-/Karten-Stil dieses Dashboards.
ACCENT_COLOR = "#2563EB"
ACCENT_COLOR_DARK = "#1E3A8A"
BG_COLOR = "#F1F4F8"
CARD_BG = "#FFFFFF"
TEXT_COLOR = "#1F2A33"
MUTED_TEXT_COLOR = "#6B7785"
BORDER_COLOR = "#E2E6EB"

# Vier KPI-Kacheln (Klaerung mit Max, 29.06.2026) - Reihenfolge bestimmt die
# Anordnung in der Kachel-Leiste. Jede Kachel wird aus dem jeweiligen
# summary_df einer Technologie (bzw. dem zusammengefuehrten 'Alle
# Technologien'-DataFrame) berechnet, siehe _compute_kpis().
KPI_DEFINITIONS = [
    ("materialien", "Materialien", "#2563EB"),
    ("bestandswert", "Bestandswert gesamt", "#16A34A"),
    ("reichweite", "Ø Reichweite (Monate)", "#D97706"),
    ("kritisch", "Kritische Materialien", "#DC2626"),
]


def _aggregate_for_html(summary_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Berechnet die vier Aggregations-/Ranking-Tabellen fuer die Charts
    EINER Technologie (bzw. des zusammengefuehrten 'Alle Technologien'-
    DataFrames) - siehe Modul-Docstring fuer die Liste der vier Diagramme.
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

    if "Fehlmengen_Wahrscheinlichkeit_Horizont" in summary_df.columns:
        top_risk = summary_df.dropna(subset=["Fehlmengen_Wahrscheinlichkeit_Horizont"])
        top_risk = top_risk.sort_values(
            "Fehlmengen_Wahrscheinlichkeit_Horizont", ascending=False
        ).head(TOP_N_RISK_MATERIALS)
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


def _compute_kpis(summary_df: pd.DataFrame) -> dict[str, str]:
    """Berechnet die vier KPI-Kachel-Werte (siehe KPI_DEFINITIONS) fuer EINE
    Technologie bzw. das zusammengefuehrte 'Alle Technologien'-DataFrame.

    Bewusst als fertig formatierte Strings zurueckgegeben - die Formatierung
    (Tausenderpunkt, Einheit, Nachkommastellen) ist Teil der Kachel-
    Definition und soll nicht im Template wiederholt werden muessen.
    """
    n_materialien = len(summary_df)
    bestandswert_gesamt = summary_df["Bestandswert_EUR"].sum()

    if "Reichweite_Monate" in summary_df.columns:
        reichweite_werte = summary_df["Reichweite_Monate"].dropna()
        reichweite_mittel = reichweite_werte.mean() if len(reichweite_werte) > 0 else float("nan")
    else:
        reichweite_mittel = float("nan")

    n_kritisch = int((summary_df["Prioritaet"] == "Kritisch").sum())

    return {
        "materialien": f"{n_materialien:,}".replace(",", "."),
        "bestandswert": f"{bestandswert_gesamt:,.0f} €".replace(",", "."),
        "reichweite": (
            f"{reichweite_mittel:.1f}" if pd.notna(reichweite_mittel) else "–"
        ),
        "kritisch": f"{n_kritisch:,}".replace(",", "."),
    }


def _style_figure(fig: go.Figure) -> go.Figure:
    """Wendet ein einheitliches, schlankes Layout auf eine Plotly-Figur an
    (Kachel-/Karten-Stil, angelehnt an den Referenz-Screenshot: weisser
    Hintergrund, dezente Gitterlinien, kompakte Raender). Eine Stelle fuer
    alle vier Diagrammtypen, statt das Layout in jeder Diagrammfunktion
    einzeln zu wiederholen.
    """
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
    """Baut die vier Plotly-Figuren fuer EIN Technologie-Panel. Figuren,
    deren zugrunde liegende Aggregation leer ist, werden NICHT erzeugt -
    analog zum 'nur verfuegbare Diagramme schreiben'-Prinzip der Excel-
    Variante (siehe report_builder._build_dashboard_sheet()).

    Args:
        summary_df: summary_df DIESER EINEN Technologie (bzw. der
            zusammengefuehrte 'Alle Technologien'-DataFrame).
        bestandswert_referenz_df: optionale Vergleichsbasis fuer den
            'Bestandswert je Technologie'-Balken - im Einzel-Panel die
            Bestandswert-Aggregation UEBER ALLE Technologien (damit der
            eine Balken der gewaehlten Technologie im Kontext der anderen
            gezeigt werden kann, siehe Modul-Docstring Punkt 1). Im 'Alle
            Technologien'-Panel None (dort zeigt summary_df bereits alle
            Technologien gleichzeitig).

    Returns:
        Liste von (chart_title, Figure)-Tupeln in fester Reihenfolge.
    """
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
            risk_df.sort_values("Fehlmengen_Wahrscheinlichkeit_Horizont"),
            x="Fehlmengen_Wahrscheinlichkeit_Horizont", y="Material",
            orientation="h",
            title=f"Top-{TOP_N_RISK_MATERIALS}-Materialien nach Fehlmengen-Wahrscheinlichkeit",
            labels={
                "Fehlmengen_Wahrscheinlichkeit_Horizont": "Fehlmengen-Wahrscheinlichkeit (Horizont)",
            },
            hover_data=["Materialkurztext"],
        )
        fig.update_traces(marker_color="#D9534F")
        fig.update_layout(xaxis_tickformat=".0%")
        figures.append(("Top-Fehlmengen-Risiko", _style_figure(fig)))

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


# Spalten der Detail-Tabelle (siehe _build_table_html()) - Reihenfolge
# bestimmt die Spaltenreihenfolge in der Tabelle, analog zur 'Provider
# Charges Details'-Tabelle im Referenz-Screenshot. Nur Spalten, die auch
# tatsaechlich in summary_df vorkommen, werden geschrieben (siehe
# _build_table_html() - "nur verfuegbare Spalten schreiben"-Prinzip).
TABLE_COLUMNS = [
    ("Material", "Material"),
    ("Materialkurztext", "Bezeichnung"),
    ("ABC-Kennzeichen", "ABC"),
    ("Bestandswert_EUR", "Bestandswert (EUR)"),
    ("Reichweite_Monate", "Reichweite (Mon.)"),
    ("Fehlmengen_Wahrscheinlichkeit_Horizont", "Fehlmengen-Risiko"),
    ("Prioritaet", "Priorität"),
]

# Maximale Anzahl Zeilen in der Detail-Tabelle je Panel - bei mehreren
# hundert Materialien je Technologie (siehe Max' echter Lauf: 466 fuer
# CAT 4) waere eine vollstaendige Tabelle in einer statischen HTML-Datei
# unhandlich gross; sortiert nach Risiko zeigt diese Begrenzung ohnehin die
# relevantesten Zeilen zuerst (siehe _build_table_html()).
TABLE_MAX_ROWS = 100


def _format_table_cell(column: str, value) -> str:
    """Formatiert einen einzelnen Tabellenwert fuer die Anzeige - eine
    Stelle fuer alle zahlenwertspezifischen Formate, statt das in der
    Aufrufschleife zu wiederholen.
    """
    if pd.isna(value):
        return "–"
    if column == "Bestandswert_EUR":
        return f"{value:,.0f} €".replace(",", ".")
    if column == "Reichweite_Monate":
        return f"{value:,.1f}".replace(",", ".")
    if column == "Fehlmengen_Wahrscheinlichkeit_Horizont":
        return f"{value * 100:.0f}%"
    return _html.escape(str(value))


def _build_table_html(summary_df: pd.DataFrame) -> str:
    """Baut die Detail-Tabelle (HTML-Markup) fuer EIN Technologie-Panel,
    sortiert nach Fehlmengen-Wahrscheinlichkeit (absteigend, NaN zuletzt) -
    die risikoreichsten Materialien stehen damit oben, analog zum Zweck der
    'Top-Fehlmengen-Risiko'-Grafik. Begrenzt auf TABLE_MAX_ROWS Zeilen
    (siehe Konstante).

    Die Spalte 'Prioritaet' erhaelt eine farbige Pille (Hintergrundfarbe
    aus PRIORITY_COLORS), analog zur Ampel-Faerbung im Excel-Dashboard
    (report_builder._apply_priority_highlighting()) - visuelle Konsistenz
    zwischen Excel- und HTML-Variante.

    Returns:
        Fertiges <table>...</table>-Markup als String. Leerer String, falls
        summary_df keine Zeilen enthaelt (Panel zeigt dann keine Tabelle).
    """
    if summary_df.empty:
        return ""

    available_columns = [
        (col, label) for col, label in TABLE_COLUMNS if col in summary_df.columns
    ]
    if "Fehlmengen_Wahrscheinlichkeit_Horizont" in summary_df.columns:
        sorted_df = summary_df.sort_values(
            "Fehlmengen_Wahrscheinlichkeit_Horizont", ascending=False, na_position="last"
        )
    else:
        sorted_df = summary_df
    sorted_df = sorted_df.head(TABLE_MAX_ROWS)

    header_cells = "".join(f"<th>{_html.escape(label)}</th>" for _col, label in available_columns)
    rows_html = []
    for _, row in sorted_df.iterrows():
        cells = []
        for col, _label in available_columns:
            value = row[col]
            if col == "Prioritaet" and pd.notna(value):
                color = PRIORITY_COLORS.get(value, "#AAAAAA")
                cells.append(
                    f"<td><span class='pill' style='background:{color}22;"
                    f"color:{color};border:1px solid {color}55;'>"
                    f"{_html.escape(str(value))}</span></td>"
                )
            else:
                cells.append(f"<td>{_format_table_cell(col, value)}</td>")
        rows_html.append(f"<tr>{''.join(cells)}</tr>")

    footer_note = ""
    if len(summary_df) > TABLE_MAX_ROWS:
        footer_note = (
            f"<p class='table-note'>Zeigt die {TABLE_MAX_ROWS} Materialien mit dem "
            f"höchsten Fehlmengen-Risiko von insgesamt {len(summary_df)}.</p>"
        )

    return (
        f"<div class='table-wrap'><table><thead><tr>{header_cells}</tr></thead>"
        f"<tbody>{''.join(rows_html)}</tbody></table></div>{footer_note}"
    )


def _build_kpi_html(summary_df: pd.DataFrame) -> str:
    """Baut die KPI-Kachel-Leiste (HTML-Markup) fuer EIN Technologie-Panel
    (siehe KPI_DEFINITIONS und _compute_kpis())."""
    kpi_values = _compute_kpis(summary_df)
    tiles = []
    for key, label, color in KPI_DEFINITIONS:
        tiles.append(
            f"<div class='kpi-tile'>"
            f"<div class='kpi-icon' style='background:{color}'></div>"
            f"<div class='kpi-text'>"
            f"<div class='kpi-value'>{kpi_values[key]}</div>"
            f"<div class='kpi-label'>{_html.escape(label)}</div>"
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
    Detail-Tabelle) als HTML-Markup-Block, der per CSS (display:none/block)
    ein- bzw. ausgeblendet wird (siehe Modul-Docstring zur Dropdown-Logik).

    Args:
        label: Anzeigename der Technologie (bzw. 'Alle Technologien').
        panel_id: HTML-id dieses Panels, referenziert vom Dropdown-Skript.
        summary_df: summary_df dieser Technologie.
        bestandswert_referenz_df: siehe _build_figures().
        is_first: True fuer das initial sichtbare Panel (erste Option im
            Dropdown) - alle anderen Panels starten mit display:none.

    Returns:
        (panel_html, has_content) - has_content ist False, falls summary_df
        leer war (Panel wird dann trotzdem erzeugt, zeigt aber einen
        Hinweistext statt Kacheln/Charts/Tabelle - haelt das Dropdown
        konsistent befuellt, auch wenn fuer eine Technologie z.B. keine
        Materialien vorliegen).
    """
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

    table_html = _build_table_html(summary_df)

    return (
        f"<section class='tech-panel' id='{panel_id}' style='{display_style}'>"
        f"{kpi_html}"
        f"<div class='chart-grid'>{''.join(chart_blocks)}</div>"
        f"<div class='table-card'>"
        f"<h2 class='table-title'>Materialdetails</h2>"
        f"{table_html}"
        f"</div>"
        f"</section>",
        True,
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
}}
header.topbar h1 {{
  font-size: 19px;
  margin: 0;
  font-weight: 600;
}}
.tech-select-wrap label {{
  font-size: 12px;
  color: var(--muted);
  display: block;
  margin-bottom: 4px;
}}
select#tech-select {{
  font-size: 14px;
  padding: 8px 12px;
  border-radius: 6px;
  border: 1px solid var(--border);
  background: var(--card-bg);
  color: var(--text);
  min-width: 240px;
}}
main {{ padding: 24px 32px 48px; }}
.kpi-row {{
  display: grid;
  grid-template-columns: repeat(4, minmax(180px, 1fr));
  gap: 16px;
  margin-bottom: 24px;
}}
.kpi-tile {{
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 18px 20px;
  display: flex;
  align-items: center;
  gap: 14px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.kpi-icon {{
  width: 36px;
  height: 36px;
  border-radius: 8px;
  flex-shrink: 0;
}}
.kpi-value {{ font-size: 24px; font-weight: 700; line-height: 1.1; }}
.kpi-label {{ font-size: 12px; color: var(--muted); margin-top: 2px; }}
.chart-grid {{
  display: grid;
  grid-template-columns: repeat(2, minmax(360px, 1fr));
  gap: 16px;
  margin-bottom: 24px;
}}
.chart-card {{
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 8px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.table-card {{
  background: var(--card-bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 20px 24px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
}}
.table-title {{ font-size: 15px; margin: 0 0 14px; font-weight: 600; }}
.table-wrap {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
thead th {{
  text-align: left;
  color: var(--muted);
  font-weight: 600;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  border-bottom: 1px solid var(--border);
  padding: 10px 12px;
  white-space: nowrap;
}}
tbody td {{
  padding: 10px 12px;
  border-bottom: 1px solid var(--border);
  white-space: nowrap;
}}
tbody tr:hover {{ background: var(--bg); }}
.pill {{
  display: inline-block;
  padding: 3px 10px;
  border-radius: 12px;
  font-size: 12px;
  font-weight: 600;
}}
.table-note {{ font-size: 12px; color: var(--muted); margin-top: 10px; }}
.empty-hint {{ color: var(--muted); padding: 40px 0; text-align: center; }}
@media (max-width: 900px) {{
  .kpi-row {{ grid-template-columns: repeat(2, 1fr); }}
  .chart-grid {{ grid-template-columns: 1fr; }}
}}
"""

_PAGE_SCRIPT = """
function gmtsShowTechnology(panelId) {
  document.querySelectorAll('.tech-panel').forEach(function (panel) {
    panel.style.display = (panel.id === panelId) ? '' : 'none';
  });
}
"""


def write_combined_dashboard(
    summary_by_label: dict[str, pd.DataFrame],
    output_path: Path,
    page_title: str = "GMTS Dashboard",
    all_technologies_label: str = "Alle Technologien",
) -> None:
    """Schreibt EINE interaktive HTML-Datei mit einem Technologie-Dropdown
    oben, das zwischen vorberechneten Panels (KPI-Kacheln + 4 Charts +
    Detail-Tabelle je Technologie) umschaltet - siehe Modul-Docstring.

    Ersetzt die vorherige Datei-pro-Technologie-Struktur (Klaerung mit Max,
    29.06.2026): bisher entstanden 9 einzelne HTML-Dateien (8 Einzel-
    technologien + 1 konsolidierte), jetzt EINE Datei mit Dropdown.

    Args:
        summary_by_label: Dict {Anzeigename: summary_df}, EINE Zeile pro
            Technologie - z.B. {'CAT 4': cat4_summary_df, 'DCS 1.0': ...}.
            Reihenfolge der Dict-Keys bestimmt die Dropdown-Reihenfolge;
            das erste Element ist beim Oeffnen der Seite sichtbar.
        output_path: Zielpfad fuer die HTML-Datei (siehe config.
            CONSOLIDATED_DASHBOARD_HTML_PATH - ersetzt die vorherigen
            einzelnen dashboard_html_path-Werte je Technologie).
        page_title: Seitentitel/Ueberschrift in der Kopfzeile.
        all_technologies_label: Anzeigename fuer den zusammengefuehrten
            'Alle Technologien'-Eintrag, falls in summary_by_label
            vorhanden (siehe main.py - der Aufrufer fuegt diesen Eintrag
            VOR dem Aufruf bereits in summary_by_label ein, dieses Modul
            behandelt ihn nicht anders als jede andere Technologie, ausser
            beim Bestandswert-Vergleichsbalken, siehe _build_figures()).

    Erstellt das Zielverzeichnis, falls es noch nicht existiert.

    Tut nichts (gibt nur einen Konsolen-Hinweis aus), falls summary_by_label
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

    # Bestandswert-Vergleichsbasis (alle Technologien) fuer den
    # Einzel-Panel-Balken (siehe _build_figures()) - aus jeder NICHT
    # zusammengefuehrten Technologie gebildet (all_technologies_label wird
    # hier bewusst ausgeschlossen, sonst wuerde der Gesamt-Eintrag sich
    # selbst doppelt referenzieren).
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
        # 'Alle Technologien' bekommt KEINEN Vergleichsbalken (zeigt bereits
        # alle Technologien gleichzeitig im Bestandswert-Chart selbst).
        referenz = None if label == all_technologies_label else bestandswert_referenz_df
        panel_html, _has_content = _build_panel_html(
            label, panel_id, summary_df, referenz, is_first
        )
        panels_html.append(panel_html)

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
        f"<main>{''.join(panels_html)}</main>"
        f"<script>{_PAGE_SCRIPT}</script>"
        "</body></html>"
    )

    output_path.write_text(html_doc, encoding="utf-8")
    print(f"  HTML-Dashboard gespeichert unter: {output_path}")

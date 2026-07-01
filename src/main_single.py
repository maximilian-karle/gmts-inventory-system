"""
main_single.py
===============
Einstiegspunkt fuer einen Analyse-Lauf ueber GENAU EINE Technologie - fuer
den gezielten Nachlauf, z.B. nach Korrektur eines einzelnen Exports, ohne
dafuer den vollstaendigen Lauf ueber alle Technologien (main.py) erneut
anzustossen.

Die Technologie wird ueber ein Kommandozeilenargument (Slug) ausgewaehlt,
NICHT ueber eine feste Variable in config.py - das macht den Aufruf
selbsterklaerend und verhindert das Vertippen/Liegenlassen einer falschen
Einstellung in config.py (siehe Projektstatus.md, dokumentierter Vorfall
mit doppelter ACTIVE_TECHNOLOGY_SLUG-Zuweisung).

Zwei Input-Modi (siehe config.py / data_loader.py fuer den Hintergrund):
    1. GESAMTEXPORT-MODUS (bevorzugt): liegt ein Gesamtexport vor, wird
       dieser gelesen, nach Technologie aufgeteilt und nur der angefragte
       Slug daraus weiterverarbeitet.
    2. EINZELORDNER-MODUS (Fallback): liegt KEIN Gesamtexport vor, wird wie
       bisher der technologie-spezifische Input-Ordner gelesen. Hierfuer
       muss der Slug in config.KNOWN_TECHNOLOGIES gelistet sein.

Ablauf (identisch zum Einzel-Technologie-Teil von main.py):
    1. RunConfig fuer die per Argument angegebene Technologie ermitteln
    2. Bestandsentwicklung und Risikoeinstufung berechnen (stock_analysis.py)
    3. Reichweite (Phase 2) berechnen, FALLS MVER- und SE16XXL-Inventory-
       Gesamtexport vorhanden sind (coverage_analysis.py) - sonst wird dieser
       Schritt mit Konsolen-Hinweis uebersprungen, ohne den ZMLAG-Teil
       (Phase 1) zu beeintraechtigen
    4. Einzel-Report erzeugen (report_builder.build_report), inkl.
       Reichweite-Reiter, falls Schritt 3 Daten geliefert hat

Hinweis: Dieser Lauf aktualisiert NICHT den konsolidierten Gesamt-Report
(config.CONSOLIDATED_REPORT_PATH) - dieser wird ausschliesslich von main.py
neu erzeugt, basierend auf allen Technologien, die zum Zeitpunkt des
main.py-Laufs erfolgreich verarbeitet werden konnten.

Aufruf (aus dem src/-Ordner, bei aktivierter venv):
    python main_single.py dcs_2_0
    python main_single.py cat_4

Im Gesamtexport-Modus kann der Slug auch ein automatisch generierter Slug
einer (noch) nicht in config.KNOWN_TECHNOLOGIES gelisteten Technologie sein
(siehe config.slugify()) - im Einzelordner-Fallback ist dagegen nur ein
Slug aus config.KNOWN_TECHNOLOGIES gueltig, da dort kein Gesamtexport
existiert, aus dem ein unbekannter Slug abgeleitet werden koennte.
"""

from __future__ import annotations

import sys

import config
import data_loader
import mver_loader
import se16xxl_loader
import dispo_abcxyz_loader
import stock_analysis
import coverage_analysis
import trend_analysis
import simulation_analysis
import safety_stock
import forecast_analysis
import working_capital
import executive_summary
import html_dashboard
import report_builder


def _load_phase2_sources() -> tuple:
    """Laedt, FALLS vorhanden, die beiden Gesamtexporte fuer Phase 2 (MVER
    und SE16XXL-Inventory) je EINMAL fuer diesen Lauf.

    Identisch zur gleichnamigen Funktion in main.py - bewusst dupliziert
    statt importiert, damit main_single.py weiterhin unabhaengig von main.py
    lauffaehig bleibt (siehe Modul-Docstring).

    Gibt zusaetzlich den rohen mver_long_df zurueck (ab Phase 3,
    25.06.2026) - siehe main.py._load_phase2_sources() fuer den Hintergrund.

    Returns:
        Tuple (consumption_df, stock_quantity_df, mver_long_df) - alle drei
        None, falls einer der beiden Exports fehlt (Phase 2 UND Phase 3
        werden dann fuer diesen Lauf uebersprungen, OHNE den ZMLAG-Teil zu
        beeintraechtigen).
    """
    if not mver_loader.full_mver_export_exists():
        print(
            f"Hinweis: Kein MVER-Export gefunden ('{config.MVER_FULL_EXPORT_FILENAME}') "
            f"- Phase 2 (Reichweite) und Phase 3 (Trend) werden fuer diesen Lauf uebersprungen.\n"
        )
        return None, None, None

    if not se16xxl_loader.full_se16xxl_inventory_export_exists():
        print(
            f"Hinweis: Kein SE16XXL-Inventory-Export gefunden "
            f"('{config.SE16XXL_INVENTORY_EXPORT_FILENAME}') - Phase 2 "
            f"(Reichweite) wird fuer diesen Lauf uebersprungen.\n"
        )
        return None, None, None

    print("Phase 2: lese MVER-Gesamtexport (Verbrauchshistorie) ...")
    mver_long_df = mver_loader.load_full_mver_export()
    consumption_df = mver_loader.calculate_average_monthly_consumption(mver_long_df)
    print(f"  Ø-Verbrauch (12 Monate) fuer {len(consumption_df)} Material(ien) berechnet.\n")

    print("Phase 2: lese SE16XXL-Inventory-Export (Bestandsmenge + Status) ...")
    stock_quantity_df = se16xxl_loader.load_full_se16xxl_inventory_export()
    print(f"  Bestandsmenge fuer {len(stock_quantity_df)} Material(ien) geladen.\n")

    return consumption_df, stock_quantity_df, mver_long_df


def _load_phase4_source():
    """Laedt, FALLS vorhanden, den Dispo_ABCXYZ-Gesamtexport fuer Phase 4
    (Wiederbeschaffungszeit) EINMAL fuer diesen Lauf.

    Identisch zur gleichnamigen Funktion in main.py - bewusst dupliziert
    statt importiert, damit main_single.py weiterhin unabhaengig von main.py
    lauffaehig bleibt (siehe Modul-Docstring). Bewusst getrennt von
    _load_phase2_sources(): Phase 4 haengt nur von ZMLAG und Dispo_ABCXYZ
    ab, nicht von MVER/SE16XXL-Inventory (Entscheidung Max, 25.06.2026).

    Returns:
        DataFrame (Ergebnis von dispo_abcxyz_loader.build_lead_time_table()),
        oder None, falls config.SE16XXL_DISPO_ABCXYZ_EXPORT_PATH nicht
        existiert (Phase 4 wird dann fuer diesen Lauf uebersprungen, OHNE
        Phase 1/2/3 zu beeintraechtigen).
    """
    if not dispo_abcxyz_loader.full_dispo_abcxyz_export_exists():
        print(
            f"Hinweis: Kein SE16XXL-Dispo_ABCXYZ-Export gefunden "
            f"('{config.SE16XXL_DISPO_ABCXYZ_EXPORT_FILENAME}') - Phase 4 "
            f"(Wiederbeschaffung) wird fuer diesen Lauf uebersprungen.\n"
        )
        return None

    print("Phase 4: lese SE16XXL-Dispo_ABCXYZ-Export (Wiederbeschaffungszeit) ...")
    dispo_df = dispo_abcxyz_loader.load_full_dispo_abcxyz_export()
    lead_time_df = dispo_abcxyz_loader.build_lead_time_table(dispo_df)
    print(f"  Wiederbeschaffungszeit fuer {len(lead_time_df)} Material(ien) berechnet.\n")

    return lead_time_df


def _run_from_full_export(
    slug: str, consumption_df=None, stock_quantity_df=None, mver_long_df=None, lead_time_df=None
) -> None:
    """GESAMTEXPORT-MODUS: liest den Gesamtexport, teilt nach Technologie
    auf und verarbeitet ausschliesslich die per Slug angefragte Technologie.

    Raises:
        ValueError: wenn der Slug im aktuellen Gesamtexport nicht vorkommt.
    """
    print(f"Gesamtexport-Modus: lese '{config.ZMLAG_FULL_EXPORT_FILENAME}' ...")
    full_df = data_loader.load_full_zmlag_export()
    print(f"  {len(full_df)} Materialien insgesamt geladen.")

    splits = data_loader.split_by_technology(full_df)
    if slug not in splits:
        valid = ", ".join(sorted(splits))
        raise ValueError(
            f"Technologie-Slug '{slug}' kommt im aktuellen Gesamtexport nicht "
            f"vor. Im Export vorhandene Slugs: {valid}"
        )

    raw_df = splits[slug]
    label = raw_df["Technologie"].iloc[0]
    cfg = config.get_config_for_label(label)
    _process(cfg, raw_df, consumption_df, stock_quantity_df, mver_long_df, lead_time_df)


def _run_from_single_folder(
    slug: str, consumption_df=None, stock_quantity_df=None, mver_long_df=None, lead_time_df=None
) -> None:
    """EINZELORDNER-MODUS (Fallback): liest den technologie-spezifischen
    Export aus seinem eigenen Input-Ordner wie bisher.
    """
    cfg = config.get_config_for_slug(slug)
    config.ensure_run_directories(cfg)

    print(f"  Input-Verzeichnis:  {cfg.input_dir}")
    print("Lese ZMLAG-Export ein ...")
    raw_df = data_loader.load_zmlag_export_for_technology(cfg)
    print(f"  {len(raw_df)} Materialien geladen.")

    _process(cfg, raw_df, consumption_df, stock_quantity_df, mver_long_df, lead_time_df)


def _process(
    cfg: config.RunConfig,
    raw_df,
    consumption_df=None,
    stock_quantity_df=None,
    mver_long_df=None,
    lead_time_df=None,
) -> None:
    """Gemeinsame Endstrecke (Berechnung + Report) fuer beide Input-Modi.

    Args:
        consumption_df, stock_quantity_df: Phase-2-Quellen (siehe
            _load_phase2_sources()), bereits EINMAL fuer diesen Lauf geladen.
            Sind beide None (Exports fehlen), wird KEIN Reichweite-Reiter
            erzeugt - der ZMLAG-Teil (Phase 1) ist davon unbeeinflusst.
        mver_long_df: Phase-3-Quelle (roher MVER long_df, siehe
            _load_phase2_sources()), bereits EINMAL fuer diesen Lauf
            geladen. Wird hier auf die Materialien dieser Technologie
            gefiltert, bevor trend_analysis.build_trend_table() sie
            auswertet. None, falls Phase 2/3 uebersprungen wurden.
        lead_time_df: Phase-4-Quelle (siehe _load_phase4_source()), bereits
            EINMAL fuer diesen Lauf geladen. Wird hier auf die Materialien
            dieser Technologie gefiltert (dispo_abcxyz_loader.
            filter_by_materials()). None, falls Phase 4 uebersprungen wurde.

    Baustein A (Simulation_Verbrauch, 25.06.2026): wird automatisch
    berechnet, sobald sowohl coverage_df ALS AUCH trend_df fuer diese
    Technologie vorliegen - keine eigene Vorbedingung/Quelle, da
    simulation_analysis.py ausschliesslich auf bereits hier erzeugten
    Zwischenergebnissen (coverage_df, trend_df) aufbaut.

    Phase 5 (Fruehwarnsystem, 29.06.2026): wird automatisch berechnet,
    sobald coverage_df, trend_df UND technology_lead_time_df fuer diese
    Technologie vorliegen - safety_stock.py benoetigt alle drei Quellen
    (Ø-Verbrauch, Streuung, Wiederbeschaffungszeit + ABC/XYZ-Klasse).

    Forecast_MVER (Modellauswahl-Forecast, 29.06.2026): wird automatisch
    berechnet, sobald mver_long_df vorliegt - EIGENSTAENDIG neben trend_df
    (gleicher mver_long_df-Filter, aber KEINE Abhaengigkeit von trend_df
    selbst, siehe forecast_analysis.py-Modul-Docstring). Wird, falls
    vorhanden, zusaetzlich an simulation_analysis.build_simulation_table()
    weitergereicht, damit Baustein A den praeziseren Erwartungswert nutzen
    kann.

    Executive Summary/Dashboard (Phase 6, 29.06.2026): wird IMMER berechnet
    (Phase 1 reicht als Grundlage, siehe executive_summary.py - alle
    weiteren Quellen sind dort optional) und erzeugt zusaetzlich zu den
    bisherigen Reitern 'Executive_Summary' und 'Dashboard' im Einzel-
    Report dieser Technologie.
    """
    config.ensure_run_directories(cfg)
    print(f"  Output-Verzeichnis: {cfg.output_dir}")

    print("Berechne Bestandsentwicklung ...")
    df = stock_analysis.calculate_stock_development(raw_df)

    print("Klassifiziere Risiko nach ABC/XYZ ...")
    df = stock_analysis.classify_risk_by_abc_xyz(df)

    coverage_df = None
    if consumption_df is not None and stock_quantity_df is not None:
        print("Berechne Reichweite (Phase 2) ...")
        coverage_df = coverage_analysis.build_coverage_table(
            df[["Material"]], consumption_df, stock_quantity_df
        )

    trend_df = None
    outlier_cells = None
    details_df = None
    if mver_long_df is not None:
        print("Berechne Trend/Sondereffekte/Saisonalitaet (Phase 3) ...")
        technology_materials = set(df["Material"])
        technology_mver_long_df = mver_long_df[
            mver_long_df["Material"].isin(technology_materials)
        ]
        if technology_mver_long_df.empty:
            print(
                "  Hinweis: Keine MVER-Verbrauchsdaten fuer diese Technologie "
                "gefunden - Trend_MVER-Reiter wird uebersprungen.\n"
            )
        else:
            trend_df, outlier_cells, details_df = trend_analysis.build_trend_table(
                technology_mver_long_df
            )

    forecast_df = None
    if mver_long_df is not None:
        print("Berechne Modellauswahl-Forecast (Naive/SES/Holt/Holt-Winters) ...")
        technology_materials = set(df["Material"])
        technology_mver_long_df = mver_long_df[
            mver_long_df["Material"].isin(technology_materials)
        ]
        if technology_mver_long_df.empty:
            print(
                "  Hinweis: Keine MVER-Verbrauchsdaten fuer diese Technologie "
                "gefunden - Forecast_MVER-Reiter wird uebersprungen.\n"
            )
        else:
            forecast_df = forecast_analysis.build_forecast_table(technology_mver_long_df)
            if forecast_df.empty:
                print(
                    "  Hinweis: Kein Material dieser Technologie hat genug "
                    "Historie fuer einen Forecast - Forecast_MVER-Reiter wird "
                    "uebersprungen.\n"
                )
                forecast_df = None
            else:
                n_materials_forecast = forecast_df["Material"].nunique()
                print(f"  Forecast fuer {n_materials_forecast} Material(ien) berechnet.\n")

    simulation_df = None
    if coverage_df is not None and trend_df is not None:
        print("Berechne Monte-Carlo-Verbrauchsprognose (Baustein A) ...")
        simulation_df = simulation_analysis.build_simulation_table(
            coverage_df, trend_df, forecast_df=forecast_df
        )
        if simulation_df.empty:
            print(
                "  Hinweis: Keine gemeinsamen Materialien zwischen Reichweite "
                "und Trend fuer diese Technologie - Simulation_Verbrauch-Reiter "
                "wird uebersprungen.\n"
            )
            simulation_df = None
        else:
            print(f"  Simulation fuer {len(simulation_df)} Material(ien) abgeschlossen.\n")

    technology_lead_time_df = None
    if lead_time_df is not None:
        print("Filtere Wiederbeschaffungszeit (Phase 4) ...")
        technology_materials = set(df["Material"])
        technology_lead_time_df = dispo_abcxyz_loader.filter_by_materials(
            lead_time_df, technology_materials
        )
        if technology_lead_time_df.empty:
            print(
                "  Hinweis: Keine Dispo_ABCXYZ-Daten fuer diese Technologie "
                "gefunden - Wiederbeschaffung-Reiter wird uebersprungen.\n"
            )
            technology_lead_time_df = None

    safety_stock_df = None
    if coverage_df is not None and trend_df is not None and technology_lead_time_df is not None:
        print("Berechne Safety Stock/Meldebestand (Phase 5) ...")
        safety_stock_df = safety_stock.build_safety_stock_table(
            coverage_df, trend_df, technology_lead_time_df, zmlag_df=df
        )
        if safety_stock_df.empty:
            print(
                "  Hinweis: Keine gemeinsamen Materialien zwischen Reichweite, "
                "Trend und Wiederbeschaffung fuer diese Technologie - "
                "Fruehwarnsystem-Reiter wird uebersprungen.\n"
            )
            safety_stock_df = None
        else:
            print(f"  Safety Stock fuer {len(safety_stock_df)} Material(ien) berechnet.\n")

    working_capital_df = None
    if coverage_df is not None and safety_stock_df is not None and technology_lead_time_df is not None:
        print("Berechne Working Capital/ROI (Phase 6 Fortsetzung) ...")
        working_capital_df = working_capital.build_working_capital_table(
            coverage_df, safety_stock_df, technology_lead_time_df
        )
        if working_capital_df.empty:
            print(
                "  Hinweis: Keine gemeinsamen Materialien zwischen Reichweite, "
                "Fruehwarnsystem und Wiederbeschaffung fuer diese Technologie - "
                "Working_Capital-Reiter wird uebersprungen.\n"
            )
            working_capital_df = None
        else:
            print(f"  Working Capital fuer {len(working_capital_df)} Material(ien) berechnet.\n")

    print("Erstelle Executive Summary (Phase 6) ...")
    summary_df = executive_summary.build_executive_summary(
        df,
        coverage_df=coverage_df,
        trend_df=trend_df,
        safety_stock_df=safety_stock_df,
        simulation_df=simulation_df,
        working_capital_df=working_capital_df,
        lead_time_df=technology_lead_time_df,
    )
    print(f"  Executive Summary fuer {len(summary_df)} Material(ien) erstellt.\n")

    print("Erzeuge Excel-Report ...")
    report_builder.build_report(
        df,
        cfg,
        coverage_df=coverage_df,
        trend_df=trend_df,
        outlier_cells=outlier_cells,
        details_df=details_df,
        lead_time_df=technology_lead_time_df,
        simulation_df=simulation_df,
        safety_stock_df=safety_stock_df,
        forecast_df=forecast_df,
        working_capital_df=working_capital_df,
        summary_df=summary_df,
    )

    print(f"Fertig. Report gespeichert unter: {cfg.report_output_path}")

    print("Erzeuge interaktives HTML-Dashboard (Phase 6) ...")
    # write_combined_dashboard() statt der fruehren write_html_dashboard()
    # (Designueberarbeitung 29.06.2026, siehe html_dashboard.py-Docstring
    # und Projektstatus.md Abschnitt 4h) - hier mit nur EINEM Dict-Eintrag
    # (kein Dropdown-Mehrwert bei nur einer Technologie, aber dieselbe
    # Funktion bleibt die einzige Stelle, die das HTML-Layout erzeugt).
    # Bewusst WEITERHIN unter dem technologie-spezifischen
    # cfg.dashboard_html_path gespeichert (NICHT unter config.
    # CONSOLIDATED_DASHBOARD_HTML_PATH), damit ein main_single.py-Lauf
    # nicht versehentlich das von main.py erzeugte Gesamt-Dashboard mit nur
    # dieser einen Technologie ueberschreibt - analog zum bestehenden
    # Prinzip, dass main_single.py den konsolidierten Report nicht anfasst.
    html_dashboard.write_combined_dashboard(
        {cfg.label: summary_df},
        cfg.dashboard_html_path,
        page_title=f"GMTS Dashboard - {cfg.label}",
    )

    print(
        "\nHinweis: Der konsolidierte Gesamt-Report wurde hierbei NICHT "
        "aktualisiert - dafuer 'python main.py' (alle Technologien) ausfuehren."
    )


def run(slug: str) -> None:
    print(f"Analyse-Lauf fuer Technologie-Slug: {slug}")

    consumption_df, stock_quantity_df, mver_long_df = _load_phase2_sources()
    lead_time_df = _load_phase4_source()

    if data_loader.full_export_exists():
        _run_from_full_export(slug, consumption_df, stock_quantity_df, mver_long_df, lead_time_df)
    else:
        if slug not in config.KNOWN_TECHNOLOGIES:
            valid = ", ".join(config.KNOWN_TECHNOLOGIES.keys())
            raise ValueError(
                f"Kein Gesamtexport vorhanden, und Slug '{slug}' ist nicht in "
                f"config.KNOWN_TECHNOLOGIES gelistet (Einzelordner-Fallback "
                f"kennt nur bekannte Technologien). Gueltige Werte: {valid}"
            )
        _run_from_single_folder(slug, consumption_df, stock_quantity_df, mver_long_df, lead_time_df)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(
            "Aufruf: python main_single.py <technologie_slug>\n"
            "Im Gesamtexport-Modus: Slug muss im aktuellen Export vorkommen.\n"
            "Im Einzelordner-Fallback gueltige Slugs: "
            f"{', '.join(config.KNOWN_TECHNOLOGIES.keys())}"
        )
        sys.exit(1)

    try:
        run(sys.argv[1])
    except (ValueError, FileNotFoundError) as exc:
        print(f"Fehler: {exc}")
        sys.exit(1)

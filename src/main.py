"""
main.py
=======
Einstiegspunkt fuer einen vollstaendigen Analyse-Lauf ueber ALLE im Export
vorkommenden Technologien in einem Durchgang - kein manuelles Umstellen
einer Konfigurationsvariable mehr noetig.

Zwei Input-Modi (siehe config.py / data_loader.py fuer den Hintergrund):
    1. GESAMTEXPORT-MODUS (bevorzugt): liegt ein Gesamtexport unter
       config.ZMLAG_FULL_EXPORT_PATH vor, wird dieser EINMAL gelesen und
       anschliessend per data_loader.split_by_technology() nach der SAP-
       Spalte 'Technologie' aufgeteilt. Technologien, die nicht in
       config.KNOWN_TECHNOLOGIES gelistet sind, werden automatisch mit
       generiertem Slug mitverarbeitet statt verworfen.
    2. EINZELORDNER-MODUS (Fallback): liegt KEIN Gesamtexport vor, wird wie
       bisher je Technologie ein eigener Input-Ordner input_data/<slug>/
       erwartet und einzeln gelesen.
Beide Modi werden pro Lauf NICHT gemischt - entweder wird ausschliesslich
der Gesamtexport verwendet, oder ausschliesslich die Einzelordner.

Ablauf je Technologie (in beiden Modi identisch):
    1. RunConfig fuer die Technologie ermitteln
    2. Bestandsentwicklung und Risikoeinstufung berechnen (stock_analysis.py)
    3. Einzel-Report erzeugen (report_builder.build_report), unter
       output_data/<slug>/<slug>_bestandsreport.xlsx

Phase 2 (Reichweite, ab 24.06.2026):
    Zusaetzlich zu ZMLAG wird, FALLS vorhanden, der MVER-Gesamtexport
    (Verbrauchshistorie, siehe mver_loader.py) und der SE16XXL-Inventory-
    Export (Bestandsmenge in Stueck + Materialstatus, siehe se16xxl_loader.py)
    EINMAL fuer den gesamten Lauf gelesen - beide sind technologieuebergreifende
    Gesamtexporte, genau wie der ZMLAG-Gesamtexport. Aus den drei Quellen wird
    je Technologie der Reichweite-Reiter berechnet (coverage_analysis.py) und
    zusaetzlich zum ZMLAG_Bestand-Reiter in denselben Einzel-Report geschrieben.
    Liegt einer der beiden Exports NICHT vor, wird Phase 2 fuer den gesamten
    Lauf automatisch uebersprungen (mit Konsolen-Hinweis) - der ZMLAG-Teil
    (Phase 1) laeuft davon UNBEEINFLUSST weiter. Die "Stock + bestaetigte
    Zugaenge"-Variante der Reichweite ist bewusst NICHT enthalten (siehe
    coverage_analysis.py-Docstring fuer den Hintergrund).

Zusaetzlich am Ende:
    4. Konsolidierten Gesamt-Report ueber alle erfolgreich verarbeiteten
       Technologien erzeugen (report_builder.build_consolidated_report),
       unter config.CONSOLIDATED_REPORT_PATH. Enthaelt den Reiter
       'Alle_Technologien' (ZMLAG-Bestand) und, FALLS Phase 2 fuer den Lauf
       aktiv war, zusaetzlich einen technologieuebergreifenden Reiter
       'Reichweite' (alle je Technologie berechneten Reichweite-Tabellen in
       einer gemeinsamen Tabelle, ergaenzt um die Spalte 'Technologie' zum
       Filtern - seit 24.06.2026, vorher nur in den Einzelreports enthalten).

Phase 3 (Trend/MVER, ab 25.06.2026):
    Derselbe MVER-Gesamtexport wie in Phase 2 wird zusaetzlich mit einem
    eigenen, typischerweise GROESSEREN Fenster (Default 36 statt 12 Monate)
    ausgewertet (trend_analysis.py) - rohe monatliche Verbrauchs-Zeitreihe
    je Material plus Trend-, Sondereffekt- und Saisonalitaets-Kennzahlen.
    Ergebnis landet als zusaetzlicher Reiter 'Trend_MVER' im selben Einzel-
    Report wie ZMLAG_Bestand/Reichweite, sowie technologieuebergreifend im
    konsolidierten Report. Ist bereits Phase 2 uebersprungen (kein MVER-
    Export vorhanden), wird Phase 3 automatisch mit uebersprungen, da beide
    dieselbe Datenquelle benoetigen.

Phase 4 (Wiederbeschaffungszeit, ab 25.06.2026):
    Zusaetzlich wird, FALLS vorhanden, der SE16XXL-Dispo_ABCXYZ-Export
    (siehe dispo_abcxyz_loader.py) EINMAL fuer den gesamten Lauf gelesen
    und die Wiederbeschaffungszeit (Planlieferzeit + WE-Bearbeitungszeit,
    Whitepaper Kap. 3.2) berechnet. Bewusst UNABHAENGIG von Phase 2/3
    (Entscheidung Max, 25.06.2026): Phase 4 haengt nur von ZMLAG (Material-
    liste je Technologie) und Dispo_ABCXYZ ab, nicht von MVER/SE16XXL-
    Inventory - fehlt einer der Phase-2/3-Exporte, laeuft Phase 4 trotzdem,
    sofern Dispo_ABCXYZ vorliegt, und umgekehrt. Ergebnis landet als
    zusaetzlicher Reiter 'Wiederbeschaffung' im selben Einzel-Report sowie
    technologieuebergreifend im konsolidierten Report.

Im Einzelordner-Modus gilt weiterhin: Technologien, fuer die kein Input-
Ordner existiert oder die (noch) keinen ZMLAG-Export enthalten, werden NICHT
als Fehler behandelt und brechen den Gesamtlauf nicht ab - sie werden mit
einer Konsolenmeldung uebersprungen. Im Gesamtexport-Modus entfaellt dieser
Fall (alle im Export vorkommenden Technologien werden automatisch
verarbeitet). In beiden Modi werden fachliche Fehler bei einzelnen
Technologien (z.B. fehlende Spalten im Export) abgefangen, sodass eine
fehlerhafte Technologie nicht den gesamten Lauf stoppt.

Fuer den gezielten Nachlauf einer einzelnen Technologie (z.B. nach Korrektur
eines einzelnen Exports) siehe main_single.py.

Aufruf (aus dem src/-Ordner, bei aktivierter venv):
    python main.py
"""

from __future__ import annotations

import pandas as pd

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
    und SE16XXL-Inventory) je EINMAL fuer den gesamten Lauf.

    Gibt zusaetzlich den rohen mver_long_df zurueck (ab Phase 3,
    25.06.2026) - dieselbe MVER-Quelle wie fuer Phase 2, aber Phase 3
    (trend_analysis.build_trend_table) wertet sie mit einem eigenen,
    typischerweise GROESSEREN Fenster (Default 36 statt 12 Monate) erneut
    aus. Einmaliges Einlesen fuer beide Phasen, um die Datei nicht doppelt
    von der Platte zu laden.

    Returns:
        Tuple (consumption_df, stock_quantity_df, mver_long_df) - alle drei
        None, falls einer der beiden Exports fehlt (Phase 2 UND Phase 3
        werden dann fuer den gesamten Lauf uebersprungen, OHNE den ZMLAG-
        Teil zu beeintraechtigen).
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
    (Wiederbeschaffungszeit) EINMAL fuer den gesamten Lauf.

    Bewusst eine EIGENE Funktion, getrennt von _load_phase2_sources() -
    Phase 4 haengt NUR von ZMLAG (Materialliste je Technologie) und
    Dispo_ABCXYZ ab, NICHT von MVER/SE16XXL-Inventory (Entscheidung Max,
    25.06.2026). Fehlt der MVER- oder Inventory-Export, wird Phase 2/3
    uebersprungen, Phase 4 laeuft davon unbeeinflusst weiter, sofern
    Dispo_ABCXYZ vorliegt - und umgekehrt.

    Returns:
        DataFrame (Ergebnis von dispo_abcxyz_loader.build_lead_time_table()),
        oder None, falls config.SE16XXL_DISPO_ABCXYZ_EXPORT_PATH nicht
        existiert (Phase 4 wird dann fuer den gesamten Lauf uebersprungen,
        OHNE Phase 1/2/3 zu beeintraechtigen).
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


def _process_technology(
    raw_df,
    per_technology_dfs: dict,
    failed: list,
    consumption_df=None,
    stock_quantity_df=None,
    per_technology_coverage_dfs: dict | None = None,
    mver_long_df=None,
    per_technology_trend_dfs: dict | None = None,
    per_technology_outlier_cells: dict | None = None,
    per_technology_details_dfs: dict | None = None,
    lead_time_df=None,
    per_technology_lead_time_dfs: dict | None = None,
    per_technology_simulation_dfs: dict | None = None,
    per_technology_safety_stock_dfs: dict | None = None,
    per_technology_forecast_dfs: dict | None = None,
    per_technology_working_capital_dfs: dict | None = None,
    per_technology_summary_dfs: dict | None = None,
) -> None:
    """Verarbeitet einen bereits eingelesenen, technologie-spezifischen
    DataFrame weiter (Berechnung + Einzel-Report) und sammelt das Ergebnis
    in per_technology_dfs bzw. Fehler in failed. Gemeinsame Endstrecke fuer
    beide Input-Modi (Gesamtexport und Einzelordner), damit Berechnungs- und
    Report-Logik nicht doppelt gepflegt werden muss.

    Die RunConfig wird einheitlich ueber das tatsaechliche SAP-Label
    (raw_df['Technologie']) ermittelt statt ueber einen separat mitgefuehrten
    Slug - get_config_for_label() loest bekannte Labels auf ihren festen Slug
    und unbekannte automatisch auf einen generierten Slug auf (siehe config.py).

    Args:
        consumption_df, stock_quantity_df: Phase-2-Quellen (siehe
            _load_phase2_sources()), bereits EINMAL fuer den gesamten Lauf
            geladen. Falls beide None sind (Exports fehlen), wird fuer diese
            Technologie KEIN Reichweite-Reiter erzeugt - der ZMLAG-Teil
            (Phase 1) ist davon unbeeinflusst.
        per_technology_coverage_dfs: optionales Dict, in das die fuer diese
            Technologie berechnete coverage_df zusaetzlich eingetragen wird
            (Schluessel: cfg.slug), damit main.py sie im Anschluss an
            report_builder.build_consolidated_report() fuer den
            technologieuebergreifenden Reichweite-Reiter uebergeben kann.
            Bleibt das Dict leer/None, hat das auf den Einzel-Report keinen
            Einfluss - es ist rein additiv fuer den konsolidierten Report.
        mver_long_df: Phase-3-Quelle (roher MVER long_df, siehe
            _load_phase2_sources()), bereits EINMAL fuer den gesamten Lauf
            geladen. Wird hier auf die Materialien DIESER Technologie
            gefiltert, bevor trend_analysis.build_trend_table() sie
            auswertet (analog zum Filterprinzip von
            coverage_analysis.build_coverage_table()). Falls None, wird
            fuer diese Technologie KEIN Trend_MVER-Reiter erzeugt.
        per_technology_trend_dfs, per_technology_outlier_cells: optionale
            Dicts, analog zu per_technology_coverage_dfs, fuer den
            technologieuebergreifenden Trend_MVER-Reiter im konsolidierten
            Report.
        per_technology_details_dfs: optionales Dict, analog zu
            per_technology_trend_dfs, fuer den technologieuebergreifenden
            Trend_Details-Reiter (Transparenz-Erweiterung, ab 25.06.2026)
            im konsolidierten Report.
        lead_time_df: Phase-4-Quelle (siehe _load_phase4_source()), bereits
            EINMAL fuer den gesamten Lauf geladen. Wird hier auf die
            Materialien DIESER Technologie gefiltert (dispo_abcxyz_loader.
            filter_by_materials()), analog zum Filterprinzip von mver_long_df.
            Falls None, wird fuer diese Technologie KEIN Wiederbeschaffung-
            Reiter erzeugt.
        per_technology_lead_time_dfs: optionales Dict, analog zu
            per_technology_coverage_dfs, fuer den technologieuebergreifenden
            Wiederbeschaffung-Reiter im konsolidierten Report.
        per_technology_simulation_dfs: optionales Dict, analog zu
            per_technology_coverage_dfs, fuer den technologieuebergreifenden
            Simulation_Verbrauch-Reiter im konsolidierten Report (Baustein A,
            25.06.2026). Wird NUR befuellt, wenn fuer diese Technologie
            SOWOHL coverage_df ALS AUCH trend_df vorliegen (siehe
            simulation_analysis.py - beide Quellen werden benoetigt).
        per_technology_safety_stock_dfs: optionales Dict, analog zu
            per_technology_coverage_dfs, fuer den technologieuebergreifenden
            Fruehwarnsystem-Reiter im konsolidierten Report (Phase 5, ab
            29.06.2026). Wird NUR befuellt, wenn fuer diese Technologie ALLE
            DREI Quellen (coverage_df, trend_df, technology_lead_time_df)
            vorliegen (siehe safety_stock.py - alle drei Quellen werden
            benoetigt).
        per_technology_forecast_dfs: optionales Dict, analog zu
            per_technology_trend_dfs, fuer den technologieuebergreifenden
            Forecast_MVER-Reiter im konsolidierten Report (Modellauswahl-
            Forecast, ab 29.06.2026, siehe forecast_analysis.py).
            EIGENSTAENDIG neben trend_df berechnet (gleicher
            mver_long_df-Filter, aber KEINE Abhaengigkeit von trend_df
            selbst, siehe forecast_analysis.py-Modul-Docstring). Wird, falls
            vorhanden, zusaetzlich an simulation_analysis.
            build_simulation_table() als forecast_df-Parameter
            weitergereicht (siehe dort), damit die Monte-Carlo-Simulation
            den praeziseren Erwartungswert nutzen kann.
        per_technology_working_capital_dfs: optionales Dict, analog zu
            per_technology_coverage_dfs, fuer den technologieuebergreifenden
            Working_Capital-Reiter im konsolidierten Report (Phase 6
            Fortsetzung, ab 30.06.2026, siehe working_capital.py). Wird NUR
            befuellt, wenn fuer diese Technologie ALLE DREI Quellen
            (coverage_df, safety_stock_df, technology_lead_time_df)
            vorliegen (siehe working_capital.py - alle drei Quellen werden
            benoetigt).
        per_technology_summary_dfs: optionales Dict, analog zu
            per_technology_coverage_dfs, fuer den technologieuebergreifenden
            Executive_Summary-/Dashboard-Reiter im konsolidierten Report
            (Phase 6, ab 29.06.2026, siehe executive_summary.py). Wird
            IMMER befuellt (Phase 1 reicht als Grundlage, siehe
            executive_summary.build_executive_summary() - alle anderen
            Quellen sind dort optional), unabhaengig davon, welche der
            uebrigen Phasen fuer diese Technologie vorliegen.
    """
    label = raw_df["Technologie"].iloc[0]
    cfg = config.get_config_for_label(label)
    config.ensure_run_directories(cfg)

    try:
        print(f"  {len(raw_df)} Materialien fuer diese Technologie.")

        print("  Berechne Bestandsentwicklung ...")
        df = stock_analysis.calculate_stock_development(raw_df)

        print("  Klassifiziere Risiko nach ABC/XYZ ...")
        df = stock_analysis.classify_risk_by_abc_xyz(df)

        coverage_df = None
        if consumption_df is not None and stock_quantity_df is not None:
            print("  Berechne Reichweite (Phase 2) ...")
            coverage_df = coverage_analysis.build_coverage_table(
                df[["Material"]], consumption_df, stock_quantity_df
            )

        trend_df = None
        outlier_cells = None
        details_df = None
        if mver_long_df is not None:
            print("  Berechne Trend/Sondereffekte/Saisonalitaet (Phase 3) ...")
            technology_materials = set(df["Material"])
            technology_mver_long_df = mver_long_df[
                mver_long_df["Material"].isin(technology_materials)
            ]
            if technology_mver_long_df.empty:
                print(
                    "    Hinweis: Keine MVER-Verbrauchsdaten fuer diese Technologie "
                    "gefunden - Trend_MVER-Reiter wird fuer diese Technologie uebersprungen.\n"
                )
            else:
                trend_df, outlier_cells, details_df = trend_analysis.build_trend_table(
                    technology_mver_long_df
                )

        forecast_df = None
        if mver_long_df is not None:
            print("  Berechne Modellauswahl-Forecast (Naive/SES/Holt/Holt-Winters) ...")
            technology_materials = set(df["Material"])
            technology_mver_long_df = mver_long_df[
                mver_long_df["Material"].isin(technology_materials)
            ]
            if technology_mver_long_df.empty:
                print(
                    "    Hinweis: Keine MVER-Verbrauchsdaten fuer diese Technologie "
                    "gefunden - Forecast_MVER-Reiter wird fuer diese Technologie "
                    "uebersprungen.\n"
                )
            else:
                forecast_df = forecast_analysis.build_forecast_table(technology_mver_long_df)
                if forecast_df.empty:
                    print(
                        "    Hinweis: Kein Material dieser Technologie hat genug "
                        "Historie fuer einen Forecast - Forecast_MVER-Reiter wird "
                        "uebersprungen.\n"
                    )
                    forecast_df = None
                else:
                    n_materials_forecast = forecast_df["Material"].nunique()
                    print(f"    Forecast fuer {n_materials_forecast} Material(ien) berechnet.\n")

        simulation_df = None
        if coverage_df is not None and trend_df is not None:
            print("  Berechne Monte-Carlo-Verbrauchsprognose (Baustein A) ...")
            simulation_df = simulation_analysis.build_simulation_table(
                coverage_df, trend_df, forecast_df=forecast_df
            )
            if simulation_df.empty:
                print(
                    "    Hinweis: Keine gemeinsamen Materialien zwischen "
                    "Reichweite und Trend fuer diese Technologie - "
                    "Simulation_Verbrauch-Reiter wird uebersprungen.\n"
                )
                simulation_df = None
            else:
                print(f"    Simulation fuer {len(simulation_df)} Material(ien) abgeschlossen.\n")

        technology_lead_time_df = None
        if lead_time_df is not None:
            print("  Filtere Wiederbeschaffungszeit (Phase 4) ...")
            technology_materials = set(df["Material"])
            technology_lead_time_df = dispo_abcxyz_loader.filter_by_materials(
                lead_time_df, technology_materials
            )
            if technology_lead_time_df.empty:
                print(
                    "    Hinweis: Keine Dispo_ABCXYZ-Daten fuer diese Technologie "
                    "gefunden - Wiederbeschaffung-Reiter wird fuer diese Technologie "
                    "uebersprungen.\n"
                )
                technology_lead_time_df = None

        safety_stock_df = None
        if coverage_df is not None and trend_df is not None and technology_lead_time_df is not None:
            print("  Berechne Safety Stock/Meldebestand (Phase 5) ...")
            safety_stock_df = safety_stock.build_safety_stock_table(
                coverage_df, trend_df, technology_lead_time_df, zmlag_df=df
            )
            if safety_stock_df.empty:
                print(
                    "    Hinweis: Keine gemeinsamen Materialien zwischen Reichweite, "
                    "Trend und Wiederbeschaffung fuer diese Technologie - "
                    "Fruehwarnsystem-Reiter wird uebersprungen.\n"
                )
                safety_stock_df = None
            else:
                print(f"    Safety Stock fuer {len(safety_stock_df)} Material(ien) berechnet.\n")

        working_capital_df = None
        if coverage_df is not None and safety_stock_df is not None and technology_lead_time_df is not None:
            print("  Berechne Working Capital/ROI (Phase 6 Fortsetzung) ...")
            working_capital_df = working_capital.build_working_capital_table(
                coverage_df, safety_stock_df, technology_lead_time_df
            )
            if working_capital_df.empty:
                print(
                    "    Hinweis: Keine gemeinsamen Materialien zwischen Reichweite, "
                    "Fruehwarnsystem und Wiederbeschaffung fuer diese Technologie - "
                    "Working_Capital-Reiter wird uebersprungen.\n"
                )
                working_capital_df = None
            else:
                print(f"    Working Capital fuer {len(working_capital_df)} Material(ien) berechnet.\n")

        print("  Erstelle Executive Summary (Phase 6) ...")
        summary_df = executive_summary.build_executive_summary(
            df,
            coverage_df=coverage_df,
            trend_df=trend_df,
            safety_stock_df=safety_stock_df,
            simulation_df=simulation_df,
            working_capital_df=working_capital_df,
            lead_time_df=technology_lead_time_df,
        )
        print(f"    Executive Summary fuer {len(summary_df)} Material(ien) erstellt.\n")

        print("  Erzeuge Einzel-Report ...")
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
        print(f"    Gespeichert unter: {cfg.report_output_path}\n")

        per_technology_dfs[cfg.slug] = df
        if coverage_df is not None and per_technology_coverage_dfs is not None:
            per_technology_coverage_dfs[cfg.slug] = coverage_df
        if trend_df is not None and per_technology_trend_dfs is not None:
            per_technology_trend_dfs[cfg.slug] = trend_df
            if per_technology_outlier_cells is not None:
                per_technology_outlier_cells[cfg.slug] = outlier_cells or {}
        if details_df is not None and per_technology_details_dfs is not None:
            per_technology_details_dfs[cfg.slug] = details_df
        if technology_lead_time_df is not None and per_technology_lead_time_dfs is not None:
            per_technology_lead_time_dfs[cfg.slug] = technology_lead_time_df
        if simulation_df is not None and per_technology_simulation_dfs is not None:
            per_technology_simulation_dfs[cfg.slug] = simulation_df
        if safety_stock_df is not None and per_technology_safety_stock_dfs is not None:
            per_technology_safety_stock_dfs[cfg.slug] = safety_stock_df
        if forecast_df is not None and per_technology_forecast_dfs is not None:
            per_technology_forecast_dfs[cfg.slug] = forecast_df
        if working_capital_df is not None and per_technology_working_capital_dfs is not None:
            per_technology_working_capital_dfs[cfg.slug] = working_capital_df
        if per_technology_summary_dfs is not None:
            per_technology_summary_dfs[cfg.slug] = summary_df

    except (ValueError, KeyError) as exc:
        print(f"  Fehler bei '{cfg.label}', Technologie wird uebersprungen: {exc}\n")
        failed.append((cfg.slug, str(exc)))


def _run_full_export_mode(
    consumption_df=None, stock_quantity_df=None, mver_long_df=None, lead_time_df=None
) -> tuple[dict, dict, dict, dict, dict, dict, dict, dict, dict, dict, dict, list, list]:
    """GESAMTEXPORT-MODUS: liest config.ZMLAG_FULL_EXPORT_PATH einmal,
    teilt nach Technologie auf und verarbeitet jede Technologie weiter.
    """
    per_technology_dfs: dict[str, object] = {}
    per_technology_coverage_dfs: dict[str, object] = {}
    per_technology_trend_dfs: dict[str, object] = {}
    per_technology_outlier_cells: dict[str, object] = {}
    per_technology_details_dfs: dict[str, object] = {}
    per_technology_lead_time_dfs: dict[str, object] = {}
    per_technology_simulation_dfs: dict[str, object] = {}
    per_technology_safety_stock_dfs: dict[str, object] = {}
    per_technology_forecast_dfs: dict[str, object] = {}
    per_technology_working_capital_dfs: dict[str, object] = {}
    per_technology_summary_dfs: dict[str, object] = {}
    skipped: list[str] = []
    failed: list[tuple[str, str]] = []

    print(f"Gesamtexport-Modus: lese '{config.ZMLAG_FULL_EXPORT_FILENAME}' ...")
    full_df = data_loader.load_full_zmlag_export()
    print(f"  {len(full_df)} Materialien insgesamt geladen.\n")

    splits = data_loader.split_by_technology(full_df)
    print(f"In {len(splits)} Technologie(n) aufgeteilt: {', '.join(sorted(splits))}\n")

    for raw_df in splits.values():
        label = raw_df["Technologie"].iloc[0]
        print(f"--- {label} ---")
        _process_technology(
            raw_df,
            per_technology_dfs,
            failed,
            consumption_df,
            stock_quantity_df,
            per_technology_coverage_dfs,
            mver_long_df,
            per_technology_trend_dfs,
            per_technology_outlier_cells,
            per_technology_details_dfs,
            lead_time_df,
            per_technology_lead_time_dfs,
            per_technology_simulation_dfs,
            per_technology_safety_stock_dfs,
            per_technology_forecast_dfs,
            per_technology_working_capital_dfs,
            per_technology_summary_dfs,
        )

    return (
        per_technology_dfs,
        per_technology_coverage_dfs,
        per_technology_trend_dfs,
        per_technology_outlier_cells,
        per_technology_details_dfs,
        per_technology_lead_time_dfs,
        per_technology_simulation_dfs,
        per_technology_safety_stock_dfs,
        per_technology_forecast_dfs,
        per_technology_working_capital_dfs,
        per_technology_summary_dfs,
        skipped,
        failed,
    )


def _run_single_folder_mode(
    consumption_df=None, stock_quantity_df=None, mver_long_df=None, lead_time_df=None
) -> tuple[dict, dict, dict, dict, dict, dict, dict, dict, dict, dict, dict, list, list]:
    """EINZELORDNER-MODUS (Fallback): wie bisher - ein Input-Ordner je
    bekannter Technologie (config.KNOWN_TECHNOLOGIES), jede einzeln gelesen.
    """
    per_technology_dfs: dict[str, object] = {}
    per_technology_coverage_dfs: dict[str, object] = {}
    per_technology_trend_dfs: dict[str, object] = {}
    per_technology_outlier_cells: dict[str, object] = {}
    per_technology_details_dfs: dict[str, object] = {}
    per_technology_lead_time_dfs: dict[str, object] = {}
    per_technology_simulation_dfs: dict[str, object] = {}
    per_technology_safety_stock_dfs: dict[str, object] = {}
    per_technology_forecast_dfs: dict[str, object] = {}
    per_technology_working_capital_dfs: dict[str, object] = {}
    per_technology_summary_dfs: dict[str, object] = {}
    skipped: list[str] = []
    failed: list[tuple[str, str]] = []

    print(
        "Kein Gesamtexport gefunden - Einzelordner-Modus (Fallback) ueber "
        f"{len(config.KNOWN_TECHNOLOGIES)} bekannte Technologien.\n"
    )

    for slug in config.KNOWN_TECHNOLOGIES:
        cfg = config.get_config_for_slug(slug)
        config.ensure_run_directories(cfg)
        print(f"--- {cfg.label} ({cfg.slug}) ---")

        try:
            data_loader.find_zmlag_file(cfg.input_dir)
        except FileNotFoundError:
            print(f"  Kein ZMLAG-Export gefunden in {cfg.input_dir} - wird uebersprungen.\n")
            skipped.append(cfg.slug)
            continue

        try:
            print("  Lese ZMLAG-Export ein ...")
            raw_df = data_loader.load_zmlag_export_for_technology(cfg)
            print(f"    {len(raw_df)} Materialien geladen.")
        except (ValueError, KeyError) as exc:
            print(f"  Fehler bei '{cfg.label}', Technologie wird uebersprungen: {exc}\n")
            failed.append((cfg.slug, str(exc)))
            continue

        _process_technology(
            raw_df,
            per_technology_dfs,
            failed,
            consumption_df,
            stock_quantity_df,
            per_technology_coverage_dfs,
            mver_long_df,
            per_technology_trend_dfs,
            per_technology_outlier_cells,
            per_technology_details_dfs,
            lead_time_df,
            per_technology_lead_time_dfs,
            per_technology_simulation_dfs,
            per_technology_safety_stock_dfs,
            per_technology_forecast_dfs,
            per_technology_working_capital_dfs,
            per_technology_summary_dfs,
        )

    return (
        per_technology_dfs,
        per_technology_coverage_dfs,
        per_technology_trend_dfs,
        per_technology_outlier_cells,
        per_technology_details_dfs,
        per_technology_lead_time_dfs,
        per_technology_simulation_dfs,
        per_technology_safety_stock_dfs,
        per_technology_forecast_dfs,
        per_technology_working_capital_dfs,
        per_technology_summary_dfs,
        skipped,
        failed,
    )


def run() -> None:
    consumption_df, stock_quantity_df, mver_long_df = _load_phase2_sources()
    lead_time_df = _load_phase4_source()

    if data_loader.full_export_exists():
        (
            per_technology_dfs,
            per_technology_coverage_dfs,
            per_technology_trend_dfs,
            per_technology_outlier_cells,
            per_technology_details_dfs,
            per_technology_lead_time_dfs,
            per_technology_simulation_dfs,
            per_technology_safety_stock_dfs,
            per_technology_forecast_dfs,
            per_technology_working_capital_dfs,
            per_technology_summary_dfs,
            skipped,
            failed,
        ) = _run_full_export_mode(consumption_df, stock_quantity_df, mver_long_df, lead_time_df)
    else:
        (
            per_technology_dfs,
            per_technology_coverage_dfs,
            per_technology_trend_dfs,
            per_technology_outlier_cells,
            per_technology_details_dfs,
            per_technology_lead_time_dfs,
            per_technology_simulation_dfs,
            per_technology_safety_stock_dfs,
            per_technology_forecast_dfs,
            per_technology_working_capital_dfs,
            per_technology_summary_dfs,
            skipped,
            failed,
        ) = _run_single_folder_mode(consumption_df, stock_quantity_df, mver_long_df, lead_time_df)

    if per_technology_dfs:
        print("Erzeuge konsolidierten Gesamt-Report ueber alle verarbeiteten Technologien ...")
        report_builder.build_consolidated_report(
            per_technology_dfs,
            config.CONSOLIDATED_REPORT_PATH,
            per_technology_coverage_dfs=per_technology_coverage_dfs,
            per_technology_trend_dfs=per_technology_trend_dfs,
            per_technology_outlier_cells=per_technology_outlier_cells,
            per_technology_details_dfs=per_technology_details_dfs,
            per_technology_lead_time_dfs=per_technology_lead_time_dfs,
            per_technology_simulation_dfs=per_technology_simulation_dfs,
            per_technology_safety_stock_dfs=per_technology_safety_stock_dfs,
            per_technology_forecast_dfs=per_technology_forecast_dfs,
            per_technology_working_capital_dfs=per_technology_working_capital_dfs,
            per_technology_summary_dfs=per_technology_summary_dfs,
        )
        print(f"  Gespeichert unter: {config.CONSOLIDATED_REPORT_PATH}\n")

        if per_technology_summary_dfs:
            print("Erzeuge interaktives HTML-Dashboard (Phase 6) mit Technologie-Dropdown ...")
            # 'Alle Technologien' zuerst im Dict (siehe write_combined_
            # dashboard()-Docstring: das erste Element ist beim Oeffnen der
            # Seite sichtbar) - Klaerung mit Max, 29.06.2026: EINE Datei mit
            # Dropdown statt bisher 9 einzelner Dateien (siehe Abschnitt 4h).
            summary_parts = []
            for slug, summary_df in per_technology_summary_dfs.items():
                part = summary_df.copy()
                part["Technologie"] = per_technology_dfs[slug]["Technologie"].iloc[0]
                summary_parts.append(part)
            combined_summary = pd.concat(summary_parts, ignore_index=True, sort=False)

            all_technologies_label = "Alle Technologien"
            summary_by_label = {all_technologies_label: combined_summary}
            # Je Technologie unter ihrem SAP-Label (nicht Slug) anzeigen,
            # sortiert nach Bestandswert absteigend - die groessten
            # Technologien erscheinen oben im Dropdown.
            label_and_value = [
                (per_technology_dfs[slug]["Technologie"].iloc[0], summary_df["Bestandswert_EUR"].sum(), summary_df)
                for slug, summary_df in per_technology_summary_dfs.items()
            ]
            for label, _value, summary_df in sorted(label_and_value, key=lambda t: t[1], reverse=True):
                summary_by_label[label] = summary_df

            html_dashboard.write_combined_dashboard(
                summary_by_label,
                config.CONSOLIDATED_DASHBOARD_HTML_PATH,
                page_title="GMTS Dashboard",
                all_technologies_label=all_technologies_label,
            )
    else:
        print("Keine Technologie erfolgreich verarbeitet - kein konsolidierter Report erzeugt.\n")

    print("Zusammenfassung:")
    print(f"  Erfolgreich verarbeitet: {len(per_technology_dfs)} ({', '.join(per_technology_dfs) or '-'})")
    if skipped:
        print(f"  Uebersprungen (kein Export vorhanden): {len(skipped)} ({', '.join(skipped) or '-'})")
    if failed:
        print(f"  Fehlgeschlagen: {len(failed)}")
        for slug, msg in failed:
            print(f"    - {slug}: {msg}")


if __name__ == "__main__":
    run()

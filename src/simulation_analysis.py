"""
simulation_analysis.py
========================
Monte-Carlo-Simulation der zukuenftigen Verbrauchsentwicklung je Material -
"Baustein A" der strategischen Erweiterung (Klaerung mit Max, 25.06.2026).

Fachlicher Hintergrund: Max moechte zwei methodisch unterschiedliche
Bausteine GETRENNT abbilden, um sie spaeter unabhaengig voneinander
weiterzuentwickeln bzw. zu kombinieren:
    - Baustein A (dieses Modul): rein modellbasierte Projektion des
      KUENFTIGEN VERBRAUCHS, ausgehend von der bereits vorhandenen
      Trend-/Streuungsanalyse aus Phase 3 (trend_analysis.py). Beantwortet
      die Frage "wie wahrscheinlich reicht der aktuelle Bestand wie lange,
      wenn der Verbrauch weiter schwankt wie in der Vergangenheit beobachtet?"
    - Baustein B (separat, noch nicht umgesetzt): tatsaechlich bereits
      getaetigte, bestaetigte offene Bestellungen (Menge + Eintreffdatum)
      aus SAP - rein deterministisch, kein Zufallsmodell noetig, da Termin
      und Menge feststehen. Bewusst NICHT Teil dieses Moduls.

Methodische Grundlage (siehe Projektstatus.md, Abschnitt 4a/4b sowie die
Designklaerung vom 25.06.2026):
    - Whitepaper Inventory_Management_I.pdf, Kap. 5 (Digital Inventory
      Twin: Monte Carlo Simulation): stochastisches Nachfragemodell
      Dt = y_hat_t + eps_t, eps_t ~ N(0, sigma^2); Inventarfortschreibung
      "inventory -= demand", Lost-Sales-Modell (kein Backorder: bei
      inventory < 0 wird die Fehlmenge vermerkt und der Bestand auf 0
      gesetzt, kein negativer Bestand).
    - Ivanov, "Introduction to Supply Chain Analytics", Kap. 2.2.6
      (Continuous Inventory Control) und Kap. 3.3.3 (Development of a
      Continuous Inventory Control Model in AnyLogic): fachliche
      Einordnung der kontinuierlichen Bestandsfortschreibung unter
      stochastischer Nachfrage.
    - Ivanov, "Global Supply Chain and Operations Management", Kap. 13.6
      (Inventory Control Policies, insb. 13.6.2 Dynamic View): dynamische
      Sicht auf die Bestandsentwicklung als Gegenstueck zu den
      statischen/deterministischen Kennzahlen aus Phase 2/4.

Detailgrad des Startwerts (Klaerung mit Max, 25.06.2026 - "so detailreich
wie moeglich"): statt eines einzelnen, ueber den Planungshorizont konstanten
Erwartungswerts wird PRO ZUKUNFTSMONAT ein eigener Erwartungswert
verwendet. Es gibt dabei ZWEI moegliche Quellen (Klaerung mit Max,
29.06.2026 - siehe forecast_analysis.py):

    1. STANDARD (kein forecast_df uebergeben): lineare Trendfortschreibung
       aus Phase 3 (trend_analysis.py), unveraendert gegenueber der
       bisherigen Logik:

           erwarteter_verbrauch(k) = mittelwert + k * Trend_Steigung_Stueck_Monat

       (k = 1..Planungshorizont). Bei Trend_Einstufung 'Stabil' ist die
       Steigung ~0, wodurch sich die Formel automatisch auf den schlichten
       Ø-Verbrauch reduziert.

    2. OPTIONAL (forecast_df uebergeben, ab 29.06.2026): der explizite
       Monats-Punktforecast aus forecast_analysis.build_forecast_table()
       wird direkt als Erwartungswert je Zukunftsmonat verwendet, ANSTATT
       der linearen Trendfortschreibung. Das macht den Erwartungswert
       praeziser fuer Materialien mit Saisonalitaet (Holt-Winters) oder
       nichtlinearem Trend (Holt), die eine lineare Fortschreibung
       systematisch falsch einschaetzen wuerde.

    Beide Quellen sind UNABHAENGIG voneinander nutzbar (forecast_df ist ein
    rein optionaler Parameter, Default None -> bisheriges Verhalten,
    Backwards Compatibility).

Die Streuung (sigma) wird WEITERHIN aus Phase 3 uebernommen: sigma = MAD /
0.6745 (Iglewicz/Hoaglin-Konvention, konsistent mit der bereits in
trend_analysis._detect_outlier_periods() verwendeten Umrechnung) - UNABHAENGIG
davon, ob der Erwartungswert aus trend_df oder forecast_df stammt. Grund
(Klaerung mit Max, 29.06.2026): forecast_analysis.py liefert zusaetzlich
RMSE als alternatives Streuungsmass (Testset-Gueteschaetzung des gewaehlten
Forecast-Modells), Max wollte sich aber noch NICHT festlegen, welches der
beiden Streuungsmasse fachlich vorzuziehen ist. Daher exportiert
build_simulation_table() bei vorhandenem forecast_df BEIDE Werte als eigene
Spalten (Verbrauch_Streuung_MAD weiterhin, zusaetzlich
Verbrauch_Streuung_RMSE) - nebeneinander sichtbar und vergleichbar im
Report, OHNE dass eine der beiden Varianten verdeckt vorausgewaehlt wird
(Konsistenz mit "No silent defaults"). Die tatsaechlich fuer die
Zufallskomponente verwendete Spalte bleibt MAD, bis eine bewusste
Entscheidung fuer RMSE getroffen wird - siehe SIGMA_SOURCE_FOR_SIMULATION.

Start-Lagerbestand (Klaerung mit Max, 25.06.2026): aktuelle Bestandsmenge
aus coverage_df (Stock-only, OHNE offene Bestellungen) - bewusst einfach
gehalten und spaeter erweiterbar, sobald Baustein B (reale offene
Bestellungen) vorliegt; siehe Pipeline-Begriff IPt = OnHand + Pipeline -
Backorders im Whitepaper, hier zunaechst nur OnHand.

Materialauswahl (Entscheidung Claude, 25.06.2026, von Max freigegeben):
INNER JOIN zwischen coverage_df (liefert Bestandsmenge) und trend_df
(liefert Trend/Streuung) - ein Material ohne beide Quellen hat keine
ausreichende Datengrundlage fuer eine sinnvolle Simulation und wird
NICHT mit Default-Annahmen aufgefuellt (Konsistenz mit dem Prinzip aus
Phase 2/3: fehlende Datengrundlage bleibt sichtbar, statt stillschweigend
durch eine Annahme ersetzt zu werden). Anzahl ausgeschlossener Materialien
wird als Konsolen-Hinweis ausgegeben. Ist zusaetzlich forecast_df gesetzt,
gilt dieselbe Logik ein drittes Mal: ein Material ohne forecast_df-Eintrag
behaelt die lineare Trendfortschreibung (Fallback, KEIN Ausschluss aus der
Simulation) - forecast_analysis.py selbst schliesst Materialien mit zu
kurzer Historie bereits aus (siehe dort), eine zusaetzliche Simulation mit
einer Annahme fuer diese Materialien waere hier nicht sinnvoll, sie liefen
ohnehin schon mit trend_df weiter, sofern dort vorhanden.

Oeffentliche Hauptfunktion:
    build_simulation_table(coverage_df, trend_df, forecast_df=None, ...) ->
        DataFrame mit einer Zeile pro Material und den unten beschriebenen
        Kennzahlen.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

# Planungshorizont in Monaten (Klaerung mit Max, 25.06.2026).
DEFAULT_HORIZON_MONTHS = 12

# Anzahl Monte-Carlo-Laeufe je Material (Klaerung mit Max, 25.06.2026 -
# bewusst hoeher als der Whitepaper-Default von 200, fuer praezisere
# Perzentil-Schaetzungen in Kauf genommene laengere Laufzeit).
DEFAULT_N_SIMULATIONS = 1000

# Umrechnung Median Absolute Deviation (MAD) -> Standardabweichung (sigma)
# unter Annahme einer Normalverteilung, identische Konvention wie in
# trend_analysis.py (Iglewicz und Hoaglin, 1993).
MAD_TO_SIGMA_FACTOR = 0.6745

# Welche Streuungsquelle die Zufallskomponente der Simulation tatsaechlich
# nutzt, wenn forecast_df verfuegbar ist und damit BEIDE Quellen (MAD aus
# Phase 3, RMSE aus forecast_analysis.py) existieren - siehe Modul-
# Docstring fuer die Begruendung, warum dies bewusst noch nicht final
# entschieden ist (Klaerung mit Max, 29.06.2026: "beides anbieten, spaeter
# im Vergleich entscheiden"). Erlaubte Werte: 'mad' oder 'rmse'. Eine
# explizite Konstante statt eines stillschweigenden Hardcodings an der
# Verwendungsstelle, damit ein spaeterer Wechsel eine bewusste, sichtbare
# Aenderung an EINER Stelle bleibt statt eine versteckte Annahme im Code.
SIGMA_SOURCE_FOR_SIMULATION = "mad"

# Perzentile, die je Material aus den 1000 Simulationslaeufen berichtet
# werden (Reichweite in Monaten, bis zum ersten Fehlmengen-Monat bzw.
# Horizont-Ende, falls keine Fehlmenge auftrat).
REPORTED_PERCENTILES = (10, 50, 90)

# Reproduzierbarkeit der Zufallszahlen je Lauf (gleiche Eingabedaten ->
# gleiches Ergebnis, wichtig fuer Nachvollziehbarkeit/Testbarkeit
# gegenueber Max). Kein fachlicher Bezug, rein technische Wahl.
RANDOM_SEED = 42


def _merge_inputs(coverage_df: pd.DataFrame, trend_df: pd.DataFrame) -> pd.DataFrame:
    """Fuehrt coverage_df (Bestandsmenge) und trend_df (Trend/Streuung) auf
    Material-Ebene zusammen - INNER JOIN, siehe Modul-Docstring fuer die
    Begruendung. Gibt eine Konsolen-Warnung aus, falls Materialien aus
    coverage_df oder trend_df dabei herausfallen.

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table(),
            muss 'Material' und 'Bestandsmenge' enthalten.
        trend_df: Ergebnis von trend_analysis.build_trend_table() (erster
            Tupel-Eintrag), muss 'Material', 'Trend_Steigung_Stueck_Monat'
            und 'Verbrauch_Streuung_MAD' enthalten. Die Trend-Mittelwerte
            werden hier NICHT direkt aus trend_df uebernommen (dort nicht
            als eigene Spalte exportiert), sondern aus den rohen, um
            Sondereffekte bereinigten Monatsspalten neu berechnet - siehe
            _extract_mean_consumption().

    Returns:
        DataFrame mit einer Zeile pro Material, das in BEIDEN
        Eingabe-DataFrames vorkommt.
    """
    required_coverage_cols = {"Material", "Bestandsmenge"}
    missing_coverage = required_coverage_cols - set(coverage_df.columns)
    if missing_coverage:
        raise ValueError(
            f"simulation_analysis._merge_inputs() erwartet in coverage_df "
            f"die Spalten {sorted(required_coverage_cols)}, es fehlen: "
            f"{sorted(missing_coverage)}."
        )

    required_trend_cols = {"Material", "Trend_Steigung_Stueck_Monat", "Verbrauch_Streuung_MAD"}
    missing_trend = required_trend_cols - set(trend_df.columns)
    if missing_trend:
        raise ValueError(
            f"simulation_analysis._merge_inputs() erwartet in trend_df die "
            f"Spalten {sorted(required_trend_cols)}, es fehlen: "
            f"{sorted(missing_trend)}."
        )

    coverage_materials = set(coverage_df["Material"])
    trend_materials = set(trend_df["Material"])
    only_coverage = coverage_materials - trend_materials
    only_trend = trend_materials - coverage_materials

    if only_coverage:
        print(
            f"  Hinweis (Simulation): {len(only_coverage)} Material(ien) aus "
            f"der Reichweite-Tabelle haben keine Trend-Daten (Phase 3) und "
            f"werden NICHT simuliert."
        )
    if only_trend:
        print(
            f"  Hinweis (Simulation): {len(only_trend)} Material(ien) aus "
            f"der Trend-Tabelle haben keine Bestandsmenge (Reichweite) und "
            f"werden NICHT simuliert."
        )

    merged = coverage_df[["Material", "Bestandsmenge"]].merge(
        trend_df[["Material", "Trend_Steigung_Stueck_Monat", "Verbrauch_Streuung_MAD"]],
        on="Material",
        how="inner",
    )
    return merged


def _extract_mean_consumption(trend_df: pd.DataFrame) -> pd.Series:
    """Berechnet je Material den mittleren monatlichen Verbrauch ueber das
    Trend-Fenster, auf Basis der um Sondereffekte bereinigten Rohwerte
    (identische Bereinigungslogik wie in trend_analysis.build_trend_table(),
    dort jedoch nur als Zwischenwert fuer die Trendgerade verwendet und
    nicht als eigene Spalte exportiert - daher hier aus den vorhandenen
    Rohspalten neu abgeleitet statt trend_analysis.py um eine zusaetzliche
    Exportspalte zu erweitern).

    Vereinfachung bewusst in Kauf genommen: hier wird der einfache Mittelwert
    UEBER ALLE rohen Monatsspalten gebildet (inkl. der als Sondereffekt
    erkannten Monate), nicht die bereinigte Version. Das ist fuer den
    Simulationsstartwert ausreichend, da die Trendgerade selbst (Steigung)
    bereits bereinigt ist und die Sondereffekt-Information separat im
    Trend_Details-Reiter einsehbar bleibt - eine zweite, hier neu
    durchgefuehrte Bereinigung wuerde nur unnoetige Komplexitaet ohne
    spuerbaren Mehrwert fuer die Simulation bedeuten.

    Args:
        trend_df: Ergebnis von trend_analysis.build_trend_table() (erster
            Tupel-Eintrag), enthaelt 'Material' sowie die rohen
            Monatsspalten im Format 'YYYY-MM'.

    Returns:
        Series, indiziert nach Position (RangeIndex von trend_df), mit dem
        mittleren monatlichen Verbrauch je Zeile.
    """
    known_non_month_columns = {
        "Material",
        "Trend_Steigung_Stueck_Monat",
        "Trend_Einstufung",
        "Sondereffekt_Monate",
        "Anzahl_Sondereffekte",
        "Verbrauch_Streuung_MAD",
        "Saisonalitaet_Hinweis",
        "Anzahl_Monate_Verfuegbar",
    }
    month_columns = [col for col in trend_df.columns if col not in known_non_month_columns]
    return trend_df[month_columns].mean(axis=1, skipna=True)


def _expected_consumption_path(
    material: str,
    mean_monthly_consumption: float,
    trend_slope: float,
    horizon_months: int,
    forecast_by_material: dict[str, np.ndarray] | None,
) -> np.ndarray:
    """Liefert den erwarteten Verbrauch je Zukunftsmonat (Array der Laenge
    horizon_months), VOR der Zufallskomponente - die fachliche Quellenwahl
    zwischen linearer Trendfortschreibung und forecast_analysis.py
    (Klaerung mit Max, 29.06.2026, siehe Modul-Docstring).

    Zwei Faelle:
        1. forecast_by_material ist None ODER enthaelt 'material' NICHT:
           lineare Trendfortschreibung wie bisher (Phase 3-Basiswert):
               erwarteter_verbrauch(k) = mean_monthly_consumption + k * trend_slope
           Ein Material ohne forecast_df-Eintrag wird NICHT aus der
           Simulation ausgeschlossen (siehe Modul-Docstring) - es faellt
           lediglich auf die bisherige Methode zurueck.
        2. forecast_by_material enthaelt 'material': der Array mit den
           12 Forecast_Verbrauch-Werten aus forecast_analysis.py wird DIREKT
           uebernommen (bereits dort auf 0 gekappt, siehe
           forecast_analysis.build_forecast_table()).

    In BEIDEN Faellen wird das Ergebnis hier zusaetzlich erneut auf 0
    gekappt (Sicherheitsnetz, falls horizon_months zwischen
    forecast_analysis.py und simulation_analysis.py auseinanderlaeuft - in
    diesem Fall wird der forecast_by_material-Array stattdessen ignoriert
    und auf die lineare Fortschreibung zurueckgefallen, siehe unten).

    Args:
        material: Material-Schluessel, fuer den Lookup in
            forecast_by_material.
        mean_monthly_consumption: mittlerer monatlicher Verbrauch (Stueck),
            Basiswert fuer Fall 1.
        trend_slope: Trend_Steigung_Stueck_Monat aus Phase 3, Basiswert fuer
            Fall 1.
        horizon_months: Planungshorizont in Monaten - MUSS mit der Laenge
            des forecast_by_material-Arrays uebereinstimmen, sonst Fallback
            auf Fall 1 (siehe oben).
        forecast_by_material: optionales Mapping Material -> Array der
            Forecast_Verbrauch-Werte (Laenge = horizon_months im
            Forecast-Lauf), oder None (kein forecast_df uebergeben).

    Returns:
        Array der Laenge horizon_months, auf 0 gekappt.
    """
    if forecast_by_material is not None and material in forecast_by_material:
        forecast_path = forecast_by_material[material]
        if len(forecast_path) == horizon_months:
            return np.clip(forecast_path, a_min=0.0, a_max=None)
        # Laengen-Mismatch (z.B. weil forecast_df mit einem anderen
        # horizon_months erzeugt wurde als der aktuelle Simulationslauf) -
        # bewusst KEIN Crash und KEIN stilles Zurechtschneiden/Auffuellen
        # (das wuerde Werte erzeugen, die forecast_analysis.py nie
        # tatsaechlich berechnet hat), sondern Fallback auf die lineare
        # Trendfortschreibung fuer dieses Material, mit Konsolen-Hinweis.
        print(
            f"  Hinweis (Simulation): forecast_df fuer Material '{material}' hat "
            f"{len(forecast_path)} Monate, erwartet werden {horizon_months} - "
            f"falle fuer dieses Material auf die lineare Trendfortschreibung zurueck."
        )

    month_indices = np.arange(1, horizon_months + 1)
    expected = mean_monthly_consumption + month_indices * trend_slope
    return np.clip(expected, a_min=0.0, a_max=None)


def _simulate_material(
    start_stock: float,
    expected_consumption_per_month: np.ndarray,
    sigma: float,
    horizon_months: int,
    n_simulations: int,
    rng: np.random.Generator,
) -> dict:
    """Fuehrt die Monte-Carlo-Simulation fuer EIN Material durch.

    Pro Simulationslauf wird der Bestand monatsweise fortgeschrieben
    (Lost-Sales-Modell wie im Whitepaper, Kap. 4.4: keine Backorders, bei
    Unterschreitung von 0 wird der Bestand auf 0 gesetzt und der Monat als
    "Fehlmenge" markiert). Der erwartete Verbrauch pro Zukunftsmonat wird
    NICHT mehr innerhalb dieser Funktion berechnet (Aenderung 29.06.2026,
    siehe Modul-Docstring) - er wird bereits FERTIG BERECHNET uebergeben
    (expected_consumption_per_month), damit diese Funktion unabhaengig
    davon ist, OB der Erwartungswert aus der linearen Trendfortschreibung
    (Phase 3) oder aus forecast_analysis.py stammt. Die eigentliche
    Quellenwahl passiert dadurch ausschliesslich in build_simulation_table()
    (_expected_consumption_path()), nicht hier.

    Um eine Zufallskomponente eps ~ N(0, sigma^2) ergaenzt. Negativer
    Verbrauch wird auf 0 gekappt (ein Material kann nicht "negativ"
    verbraucht werden).

    Args:
        start_stock: Bestandsmenge zu Simulationsbeginn (Stock-only, ohne
            offene Bestellungen - siehe Modul-Docstring).
        expected_consumption_per_month: Array der Laenge horizon_months,
            erwarteter Verbrauch je Zukunftsmonat VOR der Zufallskomponente,
            bereits auf 0 gekappt (siehe _expected_consumption_path()).
        sigma: Standardabweichung des monatlichen Verbrauchs (Quelle haengt
            von SIGMA_SOURCE_FOR_SIMULATION ab, siehe Modul-Docstring).
            Falls NaN oder negativ, wird OHNE Zufallskomponente simuliert
            (deterministischer Pfad, sigma=0 entspricht "kein beobachtbares
            Rauschen in der Historie").
        horizon_months: Planungshorizont in Monaten.
        n_simulations: Anzahl Monte-Carlo-Laeufe.
        rng: numpy Zufallsgenerator (zentral uebergeben statt pro Material
            neu erzeugt, fuer Performance und kontrollierte
            Reproduzierbarkeit ueber den gesamten Lauf).

    Returns:
        Dict mit den Kennzahlen fuer dieses Material (siehe
        build_simulation_table() Docstring fuer die vollstaendige Liste).
    """
    effective_sigma = 0.0 if (pd.isna(sigma) or sigma < 0) else sigma

    # Zufallskomponente: ein (n_simulations x horizon_months)-Array, jede
    # Zeile ist ein vollstaendiger 12-Monats-Pfad fuer EINEN Simulationslauf.
    noise = rng.normal(loc=0.0, scale=effective_sigma, size=(n_simulations, horizon_months))
    simulated_consumption = np.clip(expected_consumption_per_month + noise, a_min=0.0, a_max=None)

    # Bestandsfortschreibung (Lost-Sales): kumulativer Verbrauch je Pfad,
    # vektorisiert statt monatsweiser Python-Schleife (Performance bei
    # n_simulations=1000 x 8 Technologien x mehreren hundert Materialien).
    cumulative_consumption = np.cumsum(simulated_consumption, axis=1)
    remaining_stock = start_stock - cumulative_consumption

    # Erster Monat (1-basiert) je Pfad, in dem der Bestand erstmals unter 0
    # faellt - das ist der Lost-Sales-/Fehlmengen-Eintrittsmonat. Pfade ohne
    # Fehlmenge erhalten horizon_months + 1 (= "haelt laenger als der
    # Planungshorizont"), damit Median/Perzentil-Berechnung nicht durch NaN
    # verzerrt wird, gleichzeitig aber von echten Fehlmengen-Monaten
    # unterscheidbar bleibt (siehe Kennzahl Reichweite_Monate_P*).
    shortage_occurred = remaining_stock < 0
    first_shortage_month = np.where(
        shortage_occurred.any(axis=1),
        shortage_occurred.argmax(axis=1) + 1,
        horizon_months + 1,
    )

    fehlmengen_wahrscheinlichkeit = float(np.mean(shortage_occurred.any(axis=1)))

    percentile_values = np.percentile(first_shortage_month, REPORTED_PERCENTILES)

    # Bestand am Ende des Horizonts, GEKAPPT auf 0 (Lost-Sales-Modell - ein
    # negativer "Bestand" existiert real nicht, siehe Whitepaper Kap. 4.4).
    end_of_horizon_stock = np.clip(remaining_stock[:, -1], a_min=0.0, a_max=None)

    return {
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[0]}": float(percentile_values[0]),
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[1]}": float(percentile_values[1]),
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[2]}": float(percentile_values[2]),
        "Fehlmengen_Wahrscheinlichkeit_Horizont": fehlmengen_wahrscheinlichkeit,
        "Erwarteter_Bestand_Horizont_Ende": float(np.mean(end_of_horizon_stock)),
    }


def _forecast_lookups(
    forecast_df: pd.DataFrame, horizon_months: int
) -> tuple[dict[str, np.ndarray], pd.Series]:
    """Wandelt forecast_df (Langformat aus forecast_analysis.build_forecast_table(),
    eine Zeile pro Material+Forecast_Monat) in die zwei Nachschlage-
    Strukturen um, die build_simulation_table()/_expected_consumption_path()
    benoetigen.

    Materialien werden hier NICHT gefiltert oder validiert (z.B. ob exakt
    horizon_months Zeilen vorhanden sind) - diese Pruefung passiert je
    Material einzeln in _expected_consumption_path(), da ein Mismatch fuer
    EIN Material nicht den gesamten forecast_df verwerfen soll.

    Args:
        forecast_df: Ergebnis von forecast_analysis.build_forecast_table().
        horizon_months: Planungshorizont des AKTUELLEN Simulationslaufs -
            nur fuer die Sortierung der Forecast_Monat-Werte je Material
            relevant (chronologische Reihenfolge), nicht fuer eine Laengen-
            Validierung hier (siehe oben).

    Returns:
        Tuple (forecast_by_material, rmse_by_material):
            - forecast_by_material: Dict Material -> chronologisch
              sortiertes Array der Forecast_Verbrauch-Werte.
            - rmse_by_material: Series, indiziert nach Material, mit dem
              RMSE-Wert aus forecast_df (konstant je Material ueber alle
              Forecast-Monate, daher hier mit 'first' aggregiert).
    """
    sorted_forecast = forecast_df.sort_values(["Material", "Forecast_Monat"])
    forecast_by_material = {
        material: group["Forecast_Verbrauch"].to_numpy()
        for material, group in sorted_forecast.groupby("Material")
    }
    rmse_by_material = forecast_df.groupby("Material")["RMSE"].first()
    return forecast_by_material, rmse_by_material


def build_simulation_table(
    coverage_df: pd.DataFrame,
    trend_df: pd.DataFrame,
    forecast_df: pd.DataFrame | None = None,
    horizon_months: int = DEFAULT_HORIZON_MONTHS,
    n_simulations: int = DEFAULT_N_SIMULATIONS,
    random_seed: int = RANDOM_SEED,
) -> pd.DataFrame:
    """Komplettlauf der Monte-Carlo-Verbrauchsprognose-Simulation (Baustein
    A): Merge der Inputs, Simulation je Material, Aufbau der Ergebnistabelle.

    Dies ist die Hauptfunktion, die main.py/main_single.py fuer den
    Simulation-Reiter aufrufen werden (analog zu coverage_analysis.
    build_coverage_table() und trend_analysis.build_trend_table()).

    Args:
        coverage_df: Ergebnis von coverage_analysis.build_coverage_table()
            (Phase 2) - liefert die Bestandsmenge als Startwert.
        trend_df: Ergebnis von trend_analysis.build_trend_table() (Phase 3,
            erster Tupel-Eintrag) - liefert Trend-Steigung, Streuung (MAD)
            und die rohen Monatsspalten (fuer den Mittelwert). WEITERHIN
            Pflichtquelle fuer die Streuung (sigma), UNABHAENGIG davon, ob
            forecast_df gesetzt ist (siehe Modul-Docstring).
        forecast_df: optionales Ergebnis von forecast_analysis.
            build_forecast_table() (ab 29.06.2026). Falls gesetzt, wird der
            Erwartungswert je Zukunftsmonat fuer Materialien MIT
            Forecast-Eintrag aus forecast_df entnommen statt linear aus
            trend_df fortgeschrieben (siehe Modul-Docstring,
            _expected_consumption_path()). Materialien OHNE
            Forecast-Eintrag (z.B. zu kurze MVER-Historie, siehe
            forecast_analysis.py) fallen automatisch auf die lineare
            Trendfortschreibung zurueck - KEIN Ausschluss aus der
            Simulation. Default None -> bisheriges Verhalten unveraendert
            (Backwards Compatibility).
        horizon_months: Planungshorizont in Monaten (Default 12, siehe
            DEFAULT_HORIZON_MONTHS).
        n_simulations: Anzahl Monte-Carlo-Laeufe je Material (Default 1000,
            siehe DEFAULT_N_SIMULATIONS).
        random_seed: Seed fuer den Zufallsgenerator (Reproduzierbarkeit).

    Returns:
        DataFrame mit einer Zeile pro Material (nur Materialien, die in
        BEIDEN Pflicht-Eingabe-DataFrames coverage_df/trend_df vorkommen -
        siehe Modul-Docstring), mit Spalten:
            - Material
            - Bestandsmenge_Start (Kopie des Startwerts, zur Nachvollziehbarkeit)
            - Verbrauch_Mittelwert_Basis (mittlerer Verbrauch vor Trendfortschreibung,
              UNABHAENGIG davon, ob fuer dieses Material tatsaechlich die
              lineare Fortschreibung oder forecast_df verwendet wurde -
              bleibt als Referenzwert erhalten)
            - Trend_Steigung_Stueck_Monat (Kopie aus trend_df, zur Einordnung)
            - Verbrauch_Streuung_MAD (Kopie aus trend_df, zur Einordnung)
            - Verbrauch_Streuung_RMSE (NUR vorhanden, wenn forecast_df
              uebergeben wurde - RMSE aus forecast_analysis.py, als
              ALTERNATIVES Streuungsmass zum Vergleich neben MAD, siehe
              Modul-Docstring. NICHT die tatsaechlich fuer die
              Zufallskomponente verwendete Spalte, solange
              SIGMA_SOURCE_FOR_SIMULATION='mad' bleibt - rein informativ.
              NaN fuer Materialien ohne forecast_df-Eintrag.)
            - Erwartungswert_Quelle (NUR vorhanden, wenn forecast_df
              uebergeben wurde - 'Forecast' oder 'Trend_Linear' je
              Material, fuer Transparenz, welche der zwei Methoden
              tatsaechlich verwendet wurde)
            - Reichweite_Monate_P10/_P50/_P90 (Monat des ersten Fehlmengen-
              Eintritts, ueber die n_simulations Laeufe; horizon_months + 1
              bedeutet "haelt laenger als der Planungshorizont")
            - Fehlmengen_Wahrscheinlichkeit_Horizont (Anteil der Laeufe mit
              mindestens einem Fehlmengen-Monat innerhalb des Horizonts)
            - Erwarteter_Bestand_Horizont_Ende (Mittelwert des Endbestands
              ueber alle Laeufe, gekappt auf 0)
            - Planungshorizont_Monate, Anzahl_Simulationen (Metadaten-
              Spalten, konstant ueber alle Zeilen, fuer Nachvollziehbarkeit
              im Report ohne Zugriff auf den Python-Code)

    Raises:
        ValueError: bei fehlenden Pflichtspalten in coverage_df/trend_df
            (siehe _merge_inputs()).
    """
    merged = _merge_inputs(coverage_df, trend_df)
    if merged.empty:
        print(
            "  Hinweis (Simulation): keine gemeinsamen Materialien zwischen "
            "Reichweite- und Trend-Tabelle gefunden - Simulation wird "
            "uebersprungen."
        )
        return pd.DataFrame()

    mean_consumption_by_material = _extract_mean_consumption(trend_df).set_axis(trend_df["Material"])
    merged["Verbrauch_Mittelwert_Basis"] = merged["Material"].map(mean_consumption_by_material)

    forecast_by_material: dict[str, np.ndarray] | None = None
    rmse_by_material: pd.Series | None = None
    if forecast_df is not None and not forecast_df.empty:
        forecast_by_material, rmse_by_material = _forecast_lookups(forecast_df, horizon_months)
        materials_with_forecast = set(forecast_by_material.keys())
        materials_in_simulation = set(merged["Material"])
        only_in_forecast = materials_with_forecast - materials_in_simulation
        if only_in_forecast:
            print(
                f"  Hinweis (Simulation): {len(only_in_forecast)} Material(ien) "
                f"haben einen Forecast (forecast_analysis.py), aber keine "
                f"gemeinsame Bestandsmenge/Trend-Daten - werden NICHT "
                f"simuliert (unveraendert gegenueber dem bisherigen "
                f"coverage_df/trend_df-Ausschluss)."
            )

    rng = np.random.default_rng(random_seed)

    records = []
    for _, row in merged.iterrows():
        material = row["Material"]
        sigma = row["Verbrauch_Streuung_MAD"] / MAD_TO_SIGMA_FACTOR if pd.notna(row["Verbrauch_Streuung_MAD"]) else float("nan")

        expected_consumption_per_month = _expected_consumption_path(
            material=material,
            mean_monthly_consumption=row["Verbrauch_Mittelwert_Basis"],
            trend_slope=row["Trend_Steigung_Stueck_Monat"],
            horizon_months=horizon_months,
            forecast_by_material=forecast_by_material,
        )

        result = _simulate_material(
            start_stock=row["Bestandsmenge"],
            expected_consumption_per_month=expected_consumption_per_month,
            sigma=sigma,
            horizon_months=horizon_months,
            n_simulations=n_simulations,
            rng=rng,
        )
        result["Material"] = material
        result["Bestandsmenge_Start"] = row["Bestandsmenge"]
        result["Verbrauch_Mittelwert_Basis"] = row["Verbrauch_Mittelwert_Basis"]
        result["Trend_Steigung_Stueck_Monat"] = row["Trend_Steigung_Stueck_Monat"]
        result["Verbrauch_Streuung_MAD"] = row["Verbrauch_Streuung_MAD"]

        if forecast_by_material is not None:
            used_forecast = material in forecast_by_material and len(
                forecast_by_material[material]
            ) == horizon_months
            result["Erwartungswert_Quelle"] = "Forecast" if used_forecast else "Trend_Linear"
            result["Verbrauch_Streuung_RMSE"] = (
                float(rmse_by_material[material]) if material in rmse_by_material.index else float("nan")
            )

        records.append(result)

    result_df = pd.DataFrame.from_records(records)
    result_df["Planungshorizont_Monate"] = horizon_months
    result_df["Anzahl_Simulationen"] = n_simulations

    column_order = [
        "Material",
        "Bestandsmenge_Start",
        "Verbrauch_Mittelwert_Basis",
        "Trend_Steigung_Stueck_Monat",
        "Verbrauch_Streuung_MAD",
    ]
    if forecast_by_material is not None:
        column_order += ["Verbrauch_Streuung_RMSE", "Erwartungswert_Quelle"]
    column_order += [
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[0]}",
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[1]}",
        f"Reichweite_Monate_P{REPORTED_PERCENTILES[2]}",
        "Fehlmengen_Wahrscheinlichkeit_Horizont",
        "Erwarteter_Bestand_Horizont_Ende",
        "Planungshorizont_Monate",
        "Anzahl_Simulationen",
    ]
    return result_df[column_order]

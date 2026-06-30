"""
forecast_analysis.py
======================
Modellauswahl-basierter Verbrauchsforecast je Material - neue, eigenstaendige
Erweiterung neben Phase 3 (trend_analysis.py).

Fachlicher Hintergrund (Klaerung mit Max, 29.06.2026): Max hat ein
Referenz-Snippet aus seinem Python-in-Excel-Inventory-Tool gezeigt, das pro
Material+Werk automatisch zwischen vier Forecast-Modellen waehlt (Naive,
SES, Holt, Holt-Winters) - per Train/Test-Split und RMSE als Auswahlkriterium
- und einen expliziten 12-Monats-Punktforecast liefert, inkl. Modell-
Transparenz (gewaehltes Modell, RMSE, Bias, Glaettungsparameter Alpha/Beta/
Gamma). GMTS deckte diese Frage bisher nur indirekt ab: trend_analysis.py
liefert eine lineare Regressions-Steigung ueber das GESAMTE Fenster, aber
keinen expliziten Punktforecast und keinen Modellvergleich.

Bewusste Designentscheidungen (Klaerung mit Max, 29.06.2026):
    - EIGENSTAENDIGES, PARALLELES Modul: trend_analysis.py bleibt UNVERAENDERT.
      Dieses Modul ersetzt nichts, sondern ergaenzt einen zweiten, alternativen
      Blick auf denselben MVER-Verbrauch.
    - Direkter Zugriff auf das MVER long_df (Material, Periode, Verbrauch),
      NICHT auf trend_df. Begruendung (siehe auch Empfehlung an Max):
        1. Modul-Unabhaengigkeit ist bestehendes Prinzip (siehe
           trend_analysis.py-Docstring sowie dispo_abcxyz_loader.py: Phase 4
           bewusst unabhaengig von Phase 2/3).
        2. trend_df's Kennzahlen (Trend_Steigung, Saisonalitaet) basieren auf
           um Sondereffekte BEREINIGTEN Werten - dieses Modul soll selbst
           entscheiden koennen, wie es mit Ausreissern umgeht (oder bewusst
           NICHT bereinigt, wie das Excel-Vorbild es auch nicht tut), statt
           sich auf eine fremde Bereinigungslogik zu verlassen.
        3. trend_df ist fix auf DEFAULT_TREND_WINDOW_MONTHS (36) zugeschnitten,
           waehrend der Train/Test-Split eine FLEXIBLE, moeglichst lange
           Historie braucht (Excel-Vorbild: total_months < 12 -> kein
           Forecast, 12-23 Monate -> 6 Testmonate, >=24 Monate -> 12
           Testmonate, Holt-Winters nur ab 24 Monaten Training).
    - Eigene Datenquelle bedeutet etwas Code-Duplikation bei long_df-Zugriff -
      bewusst in Kauf genommen, analog zu main_single.py, das Phase-Loading-
      Logik aus main.py bewusst dupliziert statt importiert (Prinzip:
      Unabhaengigkeit vor Wiederverwendung).
    - Neue Abhaengigkeit 'statsmodels' (Klaerung mit Max, 29.06.2026) - bisher
      nutzte das Projekt nur pandas/numpy/openpyxl (siehe z.B. safety_stock.py,
      das den z-Wert bewusst OHNE scipy berechnet). Fuer Holt/Holt-Winters/SES
      ist eine eigene Implementierung jedoch unverhaeltnismaessig aufwaendig
      gegenueber der etablierten statsmodels-Bibliothek - hier daher bewusst
      eine neue Dependency akzeptiert statt nachgebaut.
    - Planungshorizont: 12 Monate (wie im Excel-Vorbild), siehe
      DEFAULT_HORIZON_MONTHS.
    - Spaetere Verknuepfung mit Baustein A (simulation_analysis.py) als
      praeziserer Erwartungswert vorbereitet, aber NICHT in diesem Modul
      umgesetzt (siehe simulation_analysis.py fuer den optionalen
      forecast_df-Parameter).

Vier Modelle, identisch zum Excel-Vorbild:
    - Naive: letzter beobachteter Wert wird konstant fortgeschrieben.
    - SES (Simple Exponential Smoothing): nur Niveau, kein Trend/Saisonalitaet
      - geeignet fuer stabilen Verbrauch ohne Richtung.
    - Holt (Exponential Smoothing mit additivem Trend): faengt eine
      Richtung ein, aber keine Saisonalitaet.
    - Holt-Winters (Exponential Smoothing mit additivem Trend UND additiver
      Saisonalitaet, Periodenlaenge 12): nur nutzbar, wenn genug
      Trainingshistorie vorliegt (siehe MIN_TRAIN_MONTHS_FOR_HOLT_WINTERS).

Modellauswahl je Material (identisch zum Excel-Vorbild):
    1. Zu wenig Historie (< MIN_TOTAL_MONTHS_FOR_FORECAST Monate) -> Material
       wird NICHT forecastet (kein Ersatzwert, keine Default-Annahme - siehe
       "No silent defaults"-Prinzip).
    2. Sonst: Train/Test-Split (letzte TEST_MONTHS_SHORT/_LONG Monate als
       Testset, je nach Gesamthistorie), alle infrage kommenden Modelle auf
       train gefittet, RMSE+Bias auf test berechnet.
    3. Modell mit niedrigstem RMSE auf dem Testset gewinnt
       (Best_Model/RMSE/Bias/Alpha/Beta/Gamma werden aus DIESEM
       Train/Test-Lauf uebernommen - Transparenz ueber die Modell-GUETE).
    4. Der finale 12-Monats-Forecast wird anschliessend mit demselben
       Modelltyp, aber auf der VOLLEN Historie (nicht nur train) neu
       gefittet - identisch zur Logik im Excel-Vorbild ("FINAL FORECAST"-
       Abschnitt), da die Gueteschaetzung (Schritt 2/3) und der
       produktive Forecast (Schritt 4) unterschiedliche Zwecke haben.

Oeffentliche Hauptfunktion:
    build_forecast_table(long_df, horizon_months=DEFAULT_HORIZON_MONTHS,
        reference_date=None) -> DataFrame mit einer Zeile pro
        Material+Forecast-Monat (Langformat, analog zum Excel-Vorbild),
        inkl. Modell-Transparenzspalten.
"""

from __future__ import annotations

import datetime as _dt
import warnings

import numpy as np
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing, SimpleExpSmoothing

# Planungshorizont in Monaten (Klaerung mit Max, 29.06.2026 - identisch zum
# Excel-Vorbild).
DEFAULT_HORIZON_MONTHS = 12

# Mindestanzahl an Monaten Gesamthistorie, ab der ueberhaupt ein Forecast
# versucht wird (identisch zum Excel-Vorbild: "if total_months < 12:
# continue"). Materialien mit kuerzerer Historie werden NICHT forecastet -
# kein Ersatzwert, keine Default-Annahme (Konsistenz mit dem Prinzip aus
# Phase 2/3/Baustein A: fehlende Datengrundlage bleibt sichtbar).
MIN_TOTAL_MONTHS_FOR_FORECAST = 12

# Schwelle, ab der die LANGE Testset-Groesse (TEST_MONTHS_LONG statt
# TEST_MONTHS_SHORT) verwendet wird (identisch zum Excel-Vorbild:
# "test_months = 12 if total_months >= 24 else 6").
TOTAL_MONTHS_THRESHOLD_FOR_LONG_TEST = 24
TEST_MONTHS_SHORT = 6
TEST_MONTHS_LONG = 12

# Mindestanzahl an TRAININGSmonaten, ab der Holt-Winters (additive
# Saisonalitaet, Periodenlaenge 12) ueberhaupt versucht wird (identisch zum
# Excel-Vorbild: "if train_months >= 24"). Mit weniger Trainingsmonaten
# liesse sich ein 12-Monats-Saisonmuster nicht zuverlaessig schaetzen (es
# braeuchte mindestens zwei volle Zyklen).
MIN_TRAIN_MONTHS_FOR_HOLT_WINTERS = 24

# Saisonalitaets-Periodenlaenge fuer Holt-Winters (12 Kalendermonate).
SEASONAL_PERIODS = 12

# Modellnamen, identisch zum Excel-Vorbild (werden 1:1 in Best_Model
# uebernommen, daher als Konstanten statt als freier String).
MODEL_NAIVE = "Naive"
MODEL_SES = "SES"
MODEL_HOLT = "Holt"
MODEL_HOLT_WINTERS = "Holt-Winters"


def _month_sequence_between(start: _dt.date, end: _dt.date) -> list[_dt.date]:
    """Liefert die vollstaendige, lueckenlose Liste der Monatsersten
    zwischen start und end (inklusive). Analog zu
    trend_analysis._month_sequence(), hier aber eigenstaendig implementiert
    (Modul-Unabhaengigkeit, siehe Modul-Docstring).
    """
    months = []
    current = start
    while current <= end:
        months.append(current)
        if current.month == 12:
            current = _dt.date(current.year + 1, 1, 1)
        else:
            current = _dt.date(current.year, current.month + 1, 1)
    return months


def _build_material_series(group: pd.DataFrame) -> pd.Series:
    """Wandelt die long_df-Zeilen EINES Materials in eine lueckenlose,
    chronologisch sortierte pandas-Series um (Index: erster Tag des Monats
    als Timestamp, Werte: Verbrauch).

    Im Gegensatz zu trend_analysis.py wird hier NICHT auf ein festes Fenster
    zugeschnitten - die gesamte verfuegbare Historie des Materials wird
    verwendet (siehe Modul-Docstring, Begruendung 3). Fehlende Monate
    innerhalb der Historie (z.B. eine Datenluecke) werden NICHT
    stillschweigend uebersprungen, sondern als NaN in der Zeitreihe
    sichtbar gemacht (ueber eine lueckenlose Monatssequenz von erstem bis
    letztem beobachteten Monat) - Konsistenz mit dem "No silent defaults"-
    Prinzip. NaN-Werte werden im Anschluss in build_forecast_table() bei der
    Bestimmung von total_months mitgezaehlt, aber von statsmodels nicht
    automatisch interpoliert (siehe _fit_and_forecast - NaNs werden vor dem
    Fit per linearer Interpolation aufgefuellt, sonst lehnen mehrere
    statsmodels-Modelle den Fit ab).

    Args:
        group: Teilmenge von long_df fuer EIN Material (Spalten Periode,
            Verbrauch), bereits nach Future-Monaten gefiltert (siehe
            build_forecast_table()).

    Returns:
        pandas Series, DatetimeIndex (Monatsanfang, freq='MS'), float-Werte.
    """
    group = group.sort_values("Periode")
    first_month = group["Periode"].min()
    last_month = group["Periode"].max()
    month_seq = _month_sequence_between(first_month, last_month)

    by_month = group.set_index("Periode")["Verbrauch"]
    full_index = pd.DatetimeIndex([pd.Timestamp(m) for m in month_seq], freq="MS")
    series = pd.Series(
        [by_month.get(m, float("nan")) for m in month_seq],
        index=full_index,
        dtype=float,
    )
    return series


def _rmse(actual: np.ndarray, forecast: np.ndarray) -> float:
    """Root Mean Squared Error - eigenstaendig implementiert statt
    sklearn.metrics.mean_squared_error zu importieren (Excel-Vorbild nutzt
    sklearn, hier bewusst vermieden: identisches Ergebnis, aber keine
    zusaetzliche Dependency fuer eine einzeilige Formel - anders als bei
    Holt/Holt-Winters/SES, wo statsmodels echten Mehrwert bietet, siehe
    Modul-Docstring zur statsmodels-Entscheidung).
    """
    return float(np.sqrt(np.mean((actual - forecast) ** 2)))


def _fit_and_forecast(
    series: pd.Series, model_name: str, n_periods: int
) -> tuple[np.ndarray, dict[str, float | None]]:
    """Fittet EIN benanntes Modell auf 'series' und liefert den Forecast der
    naechsten n_periods Monate sowie die Glaettungsparameter (Alpha/Beta/
    Gamma, je nach Modelltyp None).

    NaN-Werte in 'series' (Datenluecken innerhalb der Historie, siehe
    _build_material_series) werden vor dem Fit per linearer Interpolation
    aufgefuellt - statsmodels' ExponentialSmoothing/SimpleExpSmoothing lehnt
    NaN-Werte sonst komplett ab. Dies betrifft nur Luecken INNERHALB der
    Historie eines Materials (zwischen erstem und letztem beobachteten
    Monat), nicht Future-Monate (die werden bereits vor Aufruf dieser
    Funktion aus der Historie entfernt, siehe build_forecast_table()).

    Raises:
        ValueError: wenn model_name unbekannt ist (Programmierfehler, keine
            Nutzereingabe - daher hartes Fehlschlagen statt Fallback).
    """
    clean_series = series.interpolate(limit_direction="both")

    params: dict[str, float | None] = {"Alpha": None, "Beta": None, "Gamma": None}

    if model_name == MODEL_NAIVE:
        forecast = np.repeat(clean_series.iloc[-1], n_periods)
        return forecast, params

    # statsmodels gibt bei kurzen/grenzwertigen Reihen haeufig Konvergenz-
    # oder Frequenz-Warnungen aus, die fuer unseren Zweck (automatisierter
    # Massenlauf ueber viele Materialien) keine Handlungsrelevanz haben -
    # bewusst unterdrueckt, um die Konsolenausgabe bei z.B. 8 Technologien x
    # mehreren hundert Materialien nicht zu ueberfluten. Echte Fehler (z.B.
    # nicht konvergierter Fit) werfen weiterhin eine Exception und werden
    # von _select_best_model() als Fallback auf Naive abgefangen.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")

        if model_name == MODEL_SES:
            fitted = SimpleExpSmoothing(clean_series, initialization_method="estimated").fit()
            forecast = fitted.forecast(n_periods).to_numpy()
            params["Alpha"] = float(fitted.params.get("smoothing_level"))
            return forecast, params

        if model_name == MODEL_HOLT:
            fitted = ExponentialSmoothing(
                clean_series, trend="add", initialization_method="estimated"
            ).fit()
            forecast = fitted.forecast(n_periods).to_numpy()
            params["Alpha"] = float(fitted.params.get("smoothing_level"))
            params["Beta"] = float(fitted.params.get("smoothing_trend"))
            return forecast, params

        if model_name == MODEL_HOLT_WINTERS:
            fitted = ExponentialSmoothing(
                clean_series,
                trend="add",
                seasonal="add",
                seasonal_periods=SEASONAL_PERIODS,
                initialization_method="estimated",
            ).fit()
            forecast = fitted.forecast(n_periods).to_numpy()
            params["Alpha"] = float(fitted.params.get("smoothing_level"))
            params["Beta"] = float(fitted.params.get("smoothing_trend"))
            params["Gamma"] = float(fitted.params.get("smoothing_seasonal"))
            return forecast, params

    raise ValueError(f"forecast_analysis._fit_and_forecast(): unbekanntes Modell '{model_name}'.")


def _select_best_model(
    series: pd.Series, test_months: int
) -> tuple[str, float, float, dict[str, float | None], int]:
    """Train/Test-Split + Modellvergleich fuer EIN Material (Schritt 2/3 im
    Modul-Docstring). Probiert alle infrage kommenden Modelle auf dem
    Trainingsteil, bewertet sie per RMSE auf dem Testteil, liefert das beste
    Modell samt dessen Guete-/Parameterwerten aus genau diesem
    Train/Test-Lauf zurueck.

    Modelle, deren Fit fehlschlaegt (z.B. Konvergenzfehler bei sehr kurzen
    oder degenerierten Reihen), fallen auf den Naive-Forecast dieses Modells
    zurueck (identisch zum Excel-Vorbild: "except: forecasts[...] =
    forecasts['Naive']") - Naive selbst kann nicht fehlschlagen (reine
    Wertfortschreibung), ist also immer als Mindeststandard verfuegbar.

    Args:
        series: vollstaendige Verbrauchs-Zeitreihe des Materials (siehe
            _build_material_series), VOR dem Split.
        test_months: Anzahl der letzten Monate, die als Testset
            zurueckgehalten werden (siehe TEST_MONTHS_SHORT/_LONG).

    Returns:
        Tuple (best_model_name, rmse, bias, params, train_months):
            - best_model_name: Name des Modells mit dem niedrigsten RMSE
            - rmse: RMSE dieses Modells auf dem Testset
            - bias: mittlere Abweichung (Ist - Forecast) auf dem Testset,
              identisch zum Excel-Vorbild ("np.mean(test.values - fc)") -
              positiver Bias bedeutet systematische UNTERschaetzung durch
              das Modell, negativer Bias eine UEBERschaetzung
            - params: Alpha/Beta/Gamma DIESES Modells aus dem Train/Test-Fit
            - train_months: Anzahl Trainingsmonate (fuer die spaetere
              Holt-Winters-Verfuegbarkeitspruefung beim finalen Forecast)
    """
    train = series.iloc[:-test_months]
    test = series.iloc[-test_months:]
    train_months = len(train)

    candidate_models = [MODEL_NAIVE, MODEL_SES, MODEL_HOLT]
    if train_months >= MIN_TRAIN_MONTHS_FOR_HOLT_WINTERS:
        candidate_models.append(MODEL_HOLT_WINTERS)

    test_values = test.to_numpy()
    naive_forecast = np.repeat(train.iloc[-1], test_months)

    model_rmse: dict[str, float] = {}
    model_bias: dict[str, float] = {}
    model_params: dict[str, dict[str, float | None]] = {}

    for name in candidate_models:
        try:
            forecast, params = _fit_and_forecast(train, name, test_months)
        except Exception:
            # Fallback auf Naive bei Fit-Fehlschlag, identisch zum
            # Excel-Vorbild (siehe Docstring oben). model_params bleibt dann
            # leer (None/None/None), da der Naive-Ersatz keine eigenen
            # Glaettungsparameter hat.
            forecast = naive_forecast
            params = {"Alpha": None, "Beta": None, "Gamma": None}

        model_rmse[name] = _rmse(test_values, forecast)
        model_bias[name] = float(np.mean(test_values - forecast))
        model_params[name] = params

    best_model = min(model_rmse, key=model_rmse.get)
    return (
        best_model,
        model_rmse[best_model],
        model_bias[best_model],
        model_params[best_model],
        train_months,
    )


def build_forecast_table(
    long_df: pd.DataFrame,
    horizon_months: int = DEFAULT_HORIZON_MONTHS,
    reference_date: _dt.date | None = None,
) -> pd.DataFrame:
    """Komplettlauf der Modellauswahl-basierten Verbrauchsprognose je
    Material (siehe Modul-Docstring fuer die vollstaendige Methodik).

    Materialien mit weniger als MIN_TOTAL_MONTHS_FOR_FORECAST Monaten
    Historie werden NICHT forecastet und tauchen NICHT im Ergebnis auf
    (kein Ersatzwert - siehe "No silent defaults"-Prinzip). Die Anzahl
    ausgeschlossener Materialien wird als Konsolen-Hinweis ausgegeben,
    analog zu simulation_analysis._merge_inputs().

    Args:
        long_df: MVER-Verbrauchshistorie im Langformat (Material, Periode,
            Verbrauch), z.B. Ergebnis von mver_loader.load_full_mver_export().
            Future-Monate (NaN, siehe mver_loader-Docstring) werden hier vor
            der Verarbeitung herausgefiltert, damit sie nicht faelschlich
            als "Datenluecke innerhalb der Historie" interpoliert werden.
        horizon_months: Anzahl der zu prognostizierenden Zukunftsmonate
            (Default 12, siehe DEFAULT_HORIZON_MONTHS).
        reference_date: Stichtag, ab dem der Forecast-Horizont beginnt
            (erster Forecast-Monat = naechster Kalendermonat nach
            reference_date). Sollte mit dem reference_date von
            mver_loader.load_full_mver_export() uebereinstimmen. Default:
            heutiges Datum.

    Returns:
        DataFrame im LANGFORMAT (eine Zeile pro Material+Forecast-Monat,
        identisch zur Struktur des Excel-Vorbilds), mit Spalten:
            - Material
            - Forecast_Monat (datetime.date, erster Tag des Monats)
            - Forecast_Verbrauch (float, auf 0 gekappt - ein negativer
              Verbrauch ist fachlich nicht sinnvoll, identisch zum
              Excel-Vorbild "np.clip(forecast, 0, None)")
            - Best_Model (str: 'Naive'/'SES'/'Holt'/'Holt-Winters')
            - RMSE, Bias (float, aus dem Train/Test-Vergleich, siehe
              _select_best_model())
            - Alpha, Beta, Gamma (float oder None, Glaettungsparameter des
              gewaehlten Modells aus dem finalen Fit auf der VOLLEN
              Historie - siehe Hinweis unten)
            - Total_Data_Points (int, Gesamtanzahl Monate Historie)
            - Train_Months, Test_Months (int, aus dem Train/Test-Split)

        Hinweis zu Alpha/Beta/Gamma: Die Modellauswahl (Best_Model/RMSE/Bias)
        basiert auf dem Train/Test-Split, der finale Forecast UND damit auch
        die hier ausgegebenen Alpha/Beta/Gamma stammen jedoch aus einem
        ZWEITEN Fit desselben Modelltyps auf der vollen Historie (siehe
        Modul-Docstring, Schritt 4) - identisch zur Logik des Excel-Vorbilds.
        RMSE/Bias bleiben dagegen die Testset-Gueteschaetzung, da fuer den
        vollen Fit kein Testset mehr existiert, gegen das erneut geprueft
        werden koennte.

    Raises:
        ValueError: wenn long_df leer ist oder die erwarteten Spalten
            (Material, Periode, Verbrauch) fehlen.
    """
    required_columns = {"Material", "Periode", "Verbrauch"}
    missing = required_columns - set(long_df.columns)
    if missing:
        raise ValueError(
            f"forecast_analysis.build_forecast_table() erwartet die Spalten "
            f"{sorted(required_columns)}, es fehlen: {sorted(missing)}."
        )

    if long_df.empty:
        raise ValueError(
            "forecast_analysis.build_forecast_table() hat einen leeren "
            "long_df erhalten - keine MVER-Daten zum Auswerten vorhanden."
        )

    if reference_date is None:
        reference_date = _dt.date.today()

    # Future-Monate (von mver_loader.py bereits als NaN markiert, siehe
    # Modul-Docstring) werden hier vollstaendig entfernt statt interpoliert -
    # sie gehoeren nicht zur tatsaechlichen Historie und duerfen weder in
    # total_months noch in den Modell-Fit einfliessen.
    historical = long_df.dropna(subset=["Verbrauch"]).copy()

    forecast_start = reference_date.replace(day=1)
    # Erster Forecast-Monat = naechster Kalendermonat nach dem letzten
    # ABGESCHLOSSENEN Monat (konsistent mit trend_analysis._window_bounds()
    # und mver_loader's Future-Monat-Logik: reference_date selbst gilt
    # bereits als "laufender, nicht abgeschlossener Monat").
    forecast_months = _month_sequence_between(
        forecast_start, _add_months(forecast_start, horizon_months - 1)
    )

    skipped_too_short = 0
    records = []

    for material, group in historical.groupby("Material"):
        series = _build_material_series(group[["Periode", "Verbrauch"]])
        total_months = len(series)

        if total_months < MIN_TOTAL_MONTHS_FOR_FORECAST:
            skipped_too_short += 1
            continue

        test_months = (
            TEST_MONTHS_LONG
            if total_months >= TOTAL_MONTHS_THRESHOLD_FOR_LONG_TEST
            else TEST_MONTHS_SHORT
        )

        best_model, rmse, bias, _train_test_params, train_months = _select_best_model(
            series, test_months
        )

        # Finaler Forecast: dasselbe Modell, aber auf der VOLLEN Historie neu
        # gefittet (siehe Docstring, Schritt 4 / Hinweis zu Alpha/Beta/Gamma).
        # Ein erneuter Fit-Fehlschlag hier (z.B. weil die volle Historie
        # numerisch anders konvergiert als der Trainingsteil) faellt
        # ebenfalls auf Naive zurueck, identisch zur Fallback-Logik in
        # _select_best_model().
        try:
            final_forecast, final_params = _fit_and_forecast(series, best_model, horizon_months)
        except Exception:
            final_forecast = np.repeat(series.iloc[-1], horizon_months)
            final_params = {"Alpha": None, "Beta": None, "Gamma": None}

        final_forecast = np.clip(final_forecast, a_min=0.0, a_max=None)

        for month, value in zip(forecast_months, final_forecast):
            records.append({
                "Material": material,
                "Forecast_Monat": month,
                "Forecast_Verbrauch": round(float(value), 2),
                "Best_Model": best_model,
                "RMSE": rmse,
                "Bias": bias,
                "Alpha": final_params["Alpha"],
                "Beta": final_params["Beta"],
                "Gamma": final_params["Gamma"],
                "Total_Data_Points": total_months,
                "Train_Months": train_months,
                "Test_Months": test_months,
            })

    if skipped_too_short:
        print(
            f"  Hinweis (Forecast): {skipped_too_short} Material(ien) mit "
            f"weniger als {MIN_TOTAL_MONTHS_FOR_FORECAST} Monaten Historie "
            f"werden NICHT forecastet."
        )

    return pd.DataFrame.from_records(records)


def _add_months(start: _dt.date, n: int) -> _dt.date:
    """Addiert n Kalendermonate auf 'start' (Tag wird auf 1 normalisiert).
    Kleine Hilfsfunktion, eigenstaendig statt pandas.DateOffset, um den
    Modul-Stil (reines datetime.date, kein zusaetzlicher pandas-Typwechsel)
    konsistent zu trend_analysis.py/mver_loader.py zu halten.
    """
    total_month_index = start.month - 1 + n
    year = start.year + total_month_index // 12
    month = total_month_index % 12 + 1
    return _dt.date(year, month, 1)

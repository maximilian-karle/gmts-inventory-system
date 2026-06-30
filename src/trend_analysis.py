"""
trend_analysis.py
===================
Trend-, Sondereffekt- und Saisonalitaetsanalyse auf Basis der MVER-
Verbrauchshistorie - Phase 3 des Projektplans.

Fachlicher Hintergrund (siehe "Initiale Informationen", Abschnitt "Verbrauch
der letzten 3 Jahre"): Der Kollege fordert dort explizit "36 Monate Historie
(monatlich bitte in einem getrennten Reiter), damit Trends, Sondereffekte und
echte Entwicklung erkannt werden und nicht nur einen Jahreswert interpretiert
wird". Dieses Modul deckt daher zwei Dinge ab, die gemeinsam in einem Reiter
landen (siehe report_builder.py):
    1. Die ROHE monatliche Verbrauchs-Zeitreihe selbst (pivotiert, ein Monat
       pro Spalte) - damit Max/Kollegen die Zahlen auch ohne Python pruefen
       koennen, nicht nur die abgeleiteten Kennzahlen.
    2. Abgeleitete Kennzahlen je Material: Trend, Sondereffekte (Ausreisser-
       Monate), Saisonalitaets-Hinweis, Streuung.

Methodische Grundlage (Klaerung mit Max, 24.06.2026 - siehe Projektstatus.md,
Abschnitt 4a): Ivanov, "Global Supply Chain and Operations Management",
Kap. 11.3 (Statistical Methods) nennt lineare Regression als Standardmethode
fuer Trend-Erkennung in der Bedarfsplanung. Das alte Python-in-Excel-
Whitepaper (Inventory_Management_I.pdf, Kap. 2.7) nutzt RMSE als
Streuungsmass - hier wird stattdessen die Standardabweichung ueber das
Fenster verwendet (aequivalent fuer unsere Zwecke, da kein Forecast-Modell
mit Testset existiert, gegen das ein Vorhersagefehler berechnet werden
koennte). Dieses Streuungsmass ist bewusst so angelegt, dass es in Phase 5
(Safety Stock, SS = z*sigma*sqrt(L)) direkt wiederverwendet werden kann.

Drei Kennzahlen-Bausteine:
    - TREND (lineare Regression ueber das Fenster):
        Steigung in Stueck/Monat ueber numpy.polyfit(Grad 1). Klassifikation
        'Steigend'/'Fallend'/'Stabil' ueber einen Schwellenwert RELATIV zum
        mittleren Monatsverbrauch (nicht absolut), da die Materialien stark
        unterschiedliche Verbrauchsgroessenordnungen haben (siehe
        TREND_RELATIVE_THRESHOLD).
    - SONDEREFFEKTE (Ausreisser-Monate):
        Robuste Ausreisser-Erkennung ueber Median + MAD (Median Absolute
        Deviation) statt Mittelwert/Standardabweichung - bewusst robust
        gewaehlt, weil die Ausreisser selbst sonst die Vergleichsbasis
        verzerren wuerden (ein einzelner sehr hoher Monat wuerde den
        Mittelwert/die Std nach oben ziehen und sich dadurch selbst
        "verstecken"). Modifizierter Z-Score nach Iglewicz/Hoaglin,
        Standard-Schwelle 3.5 (siehe _detect_outlier_periods).
    - SAISONALITAET (einfacher Monatsmuster-Index):
        Je Kalendermonat (Jan..Dez) das Verhaeltnis von dessen Durchschnitts-
        verbrauch zum Gesamtdurchschnitt ueber das Fenster. Ein Wert von 1.2
        fuer Dezember bedeutet "Dezember liegt im Schnitt 20% ueber dem
        Jahresdurchschnitt". Erfordert MINDESTENS 2 volle Jahre im Fenster,
        um ueberhaupt wiederkehrende Muster von Einmaleffekten unterscheiden
        zu koennen (siehe MIN_YEARS_FOR_SEASONALITY) - sonst wird kein Index
        berechnet (NaN), um keine unbegruendete Aussage zu suggerieren.

Oeffentliche Hauptfunktion:
    build_trend_table(long_df, window_months=DEFAULT_TREND_WINDOW_MONTHS,
        reference_date=None) -> DataFrame mit einer Zeile pro Material,
        rohen Monatsspalten (chronologisch, 'YYYY-MM') und den oben
        beschriebenen Kennzahlen.
"""

from __future__ import annotations

import datetime as _dt

import numpy as np
import pandas as pd


# Default-Fenstergroesse (Monate), analog zu mver_loader.DEFAULT_WINDOW_MONTHS
# fuer Phase 2 - flexibel ueber Parameter aenderbar (z.B. 48 Monate fuer eine
# laengere Historie), Default 36 entspricht der expliziten Anforderung aus
# "Initiale Informationen" ("36 Monate Historie").
DEFAULT_TREND_WINDOW_MONTHS = 36

# Schwelle fuer die Trend-Klassifikation, RELATIV zum mittleren Monats-
# verbrauch des Materials (nicht absolut in Stueck/Monat), da die acht
# Technologien stark unterschiedliche Verbrauchsgroessenordnungen haben.
# Bewertet wird die GESAMTVERAENDERUNG ueber das gesamte Fenster (Steigung
# * Anzahl Monate), nicht die Steigung je einzelnem Monat - ein Material mit
# Ø 20 Stueck/Monat, das sich ueber 36 Monate insgesamt um mehr als 5% des
# Mittelwerts veraendert hat, gilt als nicht mehr 'Stabil'. Eine Bewertung
# nur der monatlichen Steigung wuerde bei langen Fenstern selbst deutliche
# Trends faelschlich als 'Stabil' einstufen, da die Monatssteigung allein
# bei vielen Monaten naturgemaess klein ist.
TREND_RELATIVE_THRESHOLD = 0.05

# Schwelle fuer die modifizierte Z-Score-Ausreisser-Erkennung (Iglewicz und
# Hoaglin, 1993) - Standardwert aus der Literatur, gaengige Heuristik fuer
# "auffaellig genug, um als Sondereffekt zu gelten", ohne bei normaler
# Verbrauchsschwankung staendig Fehlalarme zu erzeugen.
OUTLIER_MODIFIED_Z_THRESHOLD = 3.5

# Mindestanzahl voller Kalenderjahre im Fenster, ab der ein Saisonalitaets-
# Index ueberhaupt sinnvoll interpretierbar ist (ein wiederkehrendes Muster
# erfordert per Definition mehr als einen Durchlauf). Bei weniger Jahren wird
# bewusst KEIN Index berechnet (NaN), statt eine nicht belastbare Zahl zu
# suggerieren.
MIN_YEARS_FOR_SEASONALITY = 2

# Anzahl Nachkommastellen fuer die rohen Monats-Stueckzahlen im pivotierten
# Reiter (Konsistenz mit COVERAGE_DECIMAL_FORMAT in report_builder.py).
_MONTH_COLUMN_DATE_FORMAT = "%Y-%m"


def _window_bounds(
    window_months: int, reference_date: _dt.date | None
) -> tuple[_dt.date, _dt.date]:
    """Berechnet Start- und Endmonat (jeweils 1. Tag des Monats) des
    rollierenden Fensters der letzten 'window_months' ABGESCHLOSSENEN
    Kalendermonate vor reference_date.

    Identische Fensterlogik wie mver_loader.calculate_average_monthly_
    consumption() (bewusst dieselbe Berechnung, damit Trend-Fenster und
    Ø-Verbrauch-Fenster aus Phase 2 bei gleichem reference_date konsistent
    sind), hier aber eigenstaendig implementiert, da trend_analysis.py nicht
    von mver_loader.py abhaengen soll (Modul-Unabhaengigkeit, siehe Stil in
    se16xxl_loader.py / mver_loader.py fuer _parse_german_number).
    """
    if reference_date is None:
        reference_date = _dt.date.today()

    first_day_of_ref_month = reference_date.replace(day=1)
    last_completed_month_end = first_day_of_ref_month - _dt.timedelta(days=1)
    window_end = last_completed_month_end.replace(day=1)

    window_start_year = window_end.year
    window_start_month = window_end.month - (window_months - 1)
    while window_start_month <= 0:
        window_start_month += 12
        window_start_year -= 1
    window_start = _dt.date(window_start_year, window_start_month, 1)

    return window_start, window_end


def _month_sequence(window_start: _dt.date, window_end: _dt.date) -> list[_dt.date]:
    """Liefert die vollstaendige, lueckenlose Liste der Monatsersten
    zwischen window_start und window_end (inklusive), unabhaengig davon, ob
    fuer jeden Monat tatsaechlich ein Wert in long_df vorliegt.

    Wichtig fuer den Pivot-Schritt: Fehlt ein Monat in den MVER-Rohdaten
    (z.B. neueres Material, das erst seit kuerzerer Zeit bewirtschaftet
    wird), soll die entsprechende Spalte im Report trotzdem erscheinen (mit
    NaN) statt stillschweigend zu fehlen - eine fehlende Spalte waere sonst
    leicht mit "kein Verbrauch" zu verwechseln.
    """
    months = []
    current = window_start
    while current <= window_end:
        months.append(current)
        if current.month == 12:
            current = _dt.date(current.year + 1, 1, 1)
        else:
            current = _dt.date(current.year, current.month + 1, 1)
    return months


def _pivot_monthly_columns(
    windowed: pd.DataFrame, month_sequence: list[_dt.date]
) -> pd.DataFrame:
    """Pivotiert die lange Zeitreihe (Material, Periode, Verbrauch) in eine
    Zeile pro Material mit einer Spalte pro Monat im Fenster.

    Erfuellt die explizite Anforderung aus "Initiale Informationen"
    ("36 Monate Historie monatlich in einem getrennten Reiter") - die rohen
    Werte bleiben hier sichtbar, nicht nur die abgeleiteten Kennzahlen.

    Args:
        windowed: long_df, bereits auf das Fenster eingeschraenkt.
        month_sequence: lueckenlose Liste der Monatsersten im Fenster (siehe
            _month_sequence) - definiert sowohl die Spaltenreihenfolge als
            auch, dass fehlende Monate als NaN-Spalte erscheinen statt zu
            fehlen.

    Returns:
        DataFrame, eine Zeile pro Material, Spalten im Format 'YYYY-MM'
        (chronologisch sortiert), Werte = Verbrauch (float, NaN falls fuer
        diesen Monat kein Wert vorliegt).
    """
    pivoted = windowed.pivot_table(
        index="Material", columns="Periode", values="Verbrauch", aggfunc="first"
    )

    # Sicherstellen, dass JEDER Monat des Fensters als Spalte existiert, auch
    # wenn fuer KEIN Material in diesem Monat ein MVER-Datensatz vorlag.
    for month in month_sequence:
        if month not in pivoted.columns:
            pivoted[month] = float("nan")

    pivoted = pivoted[month_sequence]
    pivoted.columns = [m.strftime(_MONTH_COLUMN_DATE_FORMAT) for m in pivoted.columns]
    pivoted = pivoted.reset_index()

    return pivoted


def _linear_trend(values: np.ndarray) -> tuple[float, float]:
    """Berechnet Steigung und Achsenabschnitt einer linearen Regression
    (Trend ueber die Zeit, x = 0..n-1 in Monaten) ueber die uebergebenen
    Werte. NaN-Werte (fehlende Monate) werden vor der Regression entfernt,
    die Monatsindizes (x-Werte) bleiben dabei aber an der urspruenglichen
    Position verankert, damit Luecken die Steigung nicht verzerren.

    Methodische Grundlage: Ivanov, "Global Supply Chain and Operations
    Management", Kap. 11.3.1 (Linear Regression) - Standardmethode fuer
    Trend-Erkennung in der Bedarfsplanung (siehe Modul-Docstring).

    Args:
        values: Verbrauchswerte ueber das Fenster, chronologisch, ggf. NaN.

    Returns:
        Tuple (steigung, achsenabschnitt). (NaN, NaN), falls weniger als 2
        gueltige (Nicht-NaN) Werte vorliegen - eine Regression durch einen
        oder null Punkte ist nicht aussagekraeftig.
    """
    x = np.arange(len(values))
    mask = ~np.isnan(values)

    if mask.sum() < 2:
        return float("nan"), float("nan")

    slope, intercept = np.polyfit(x[mask], values[mask], deg=1)
    return float(slope), float(intercept)


def _classify_trend(slope: float, mean_value: float, months_in_window: int) -> str:
    """Klassifiziert eine Trend-Steigung als 'Steigend'/'Fallend'/'Stabil',
    auf Basis der GESAMTVERAENDERUNG ueber das gesamte Fenster (slope *
    months_in_window), RELATIV zum mittleren Monatsverbrauch (siehe
    TREND_RELATIVE_THRESHOLD und Modul-Docstring fuer die fachliche
    Begruendung). Ein Vergleich nur der monatlichen Steigung mit dem
    Mittelwert wuerde bei langen Fenstern (z.B. 36 Monate) selbst klare
    Trends systematisch unterschaetzen.

    Args:
        slope: Steigung aus _linear_trend() (Stueck/Monat).
        mean_value: mittlerer Monatsverbrauch des Materials im Fenster.
        months_in_window: Anzahl der Monate, ueber die slope berechnet
            wurde (Laenge des Fensters, nicht nur die Anzahl gueltiger
            Werte - die Trendgerade ist ueber die volle Fensterlaenge
            definiert).

    Returns:
        'Steigend', 'Fallend', 'Stabil', oder 'Unbekannt' (falls slope oder
        mean_value NaN sind, z.B. zu wenige Datenpunkte oder Verbrauch
        durchgehend 0).
    """
    if pd.isna(slope) or pd.isna(mean_value) or mean_value == 0:
        return "Unbekannt"

    total_change_over_window = slope * months_in_window
    relative_change_over_window = total_change_over_window / mean_value

    if relative_change_over_window > TREND_RELATIVE_THRESHOLD:
        return "Steigend"
    if relative_change_over_window < -TREND_RELATIVE_THRESHOLD:
        return "Fallend"
    return "Stabil"


def _detect_outlier_periods(
    periods: list[str], values: np.ndarray
) -> tuple[list[str], float, float, np.ndarray]:
    """Erkennt Ausreisser-Monate (Sondereffekte, z.B. Einmalbestellung,
    Lieferunterbrechung) ueber den modifizierten Z-Score nach Iglewicz und
    Hoaglin (1993): robust gegenueber den Ausreissern selbst, da Median und
    MAD (Median Absolute Deviation) statt Mittelwert/Standardabweichung
    verwendet werden (siehe Modul-Docstring fuer die Begruendung).

    Modifizierter Z-Score: 0.6745 * (x - median) / MAD
    (0.6745 skaliert die MAD so, dass sie bei normalverteilten Daten mit der
    Standardabweichung vergleichbar ist - Standardkonstante aus der
    Literatur, kein frei gewaehlter Wert.)

    Erweitert (25.06.2026, Transparenz-Anforderung Max): gibt zusaetzlich
    den Median sowie das VOLLE Z-Score-Array (ein Wert je Monat im Fenster,
    nicht nur die als Sondereffekt erkannten) zurueck, damit der neue
    Reiter 'Trend_Details' (siehe report_builder.py) den kompletten
    Rechenweg je Material und je Monat nachvollziehbar macht - bewusst als
    Erweiterung dieser bestehenden Funktion statt einer separaten
    Neuberechnung, um zu garantieren, dass Trend_MVER (Sondereffekt-
    Erkennung) und Trend_Details (Z-Score je Zelle) immer exakt
    konsistent sind.

    Args:
        periods: Spaltennamen ('YYYY-MM'), gleiche Reihenfolge wie values.
        values: Verbrauchswerte ueber das Fenster, chronologisch, ggf. NaN.

    Returns:
        Tuple (sondereffekt_monate, mad, median, z_scores):
        - sondereffekt_monate: Liste der als Sondereffekt erkannten Monate
          (Format 'YYYY-MM', chronologisch)
        - mad: berechneter MAD-Wert (fuer Transparenz/Nachvollziehbarkeit
          im Report)
        - median: berechneter Median der gueltigen Werte (NaN, falls nicht
          berechenbar - identische Bedingung wie mad)
        - z_scores: np.ndarray gleicher Laenge wie values, modifizierter
          Z-Score je Monat (NaN, wo values NaN ist oder Median/MAD nicht
          berechenbar sind)

        mad/median = NaN, leere Liste und z_scores = alle NaN, falls
        weniger als 3 gueltige Werte vorliegen (Median/MAD ueber 1-2 Punkte
        nicht aussagekraeftig). mad = 0, median = Median der Werte, leere
        Liste und z_scores = alle NaN, falls MAD = 0 (alle gueltigen Werte
        identisch - dann gibt es per Definition keine Ausreisser, ein
        Z-Score waere hier eine Division durch 0).
    """
    mask = ~np.isnan(values)
    valid_values = values[mask]
    nan_scores = np.full(len(values), float("nan"))

    if len(valid_values) < 3:
        return [], float("nan"), float("nan"), nan_scores

    median = float(np.median(valid_values))
    mad = np.median(np.abs(valid_values - median))

    if mad == 0:
        return [], 0.0, median, nan_scores

    modified_z_scores = np.where(mask, 0.6745 * (values - median) / mad, float("nan"))
    outlier_mask = np.where(mask, np.abs(modified_z_scores) > OUTLIER_MODIFIED_Z_THRESHOLD, False)

    outlier_periods = [periods[i] for i in range(len(periods)) if outlier_mask[i]]
    return outlier_periods, float(mad), median, modified_z_scores


def _calculate_seasonality_index(
    month_sequence: list[_dt.date], values: np.ndarray
) -> dict[int, float] | None:
    """Berechnet einen einfachen Saisonalitaets-Index je Kalendermonat
    (1 = Januar, ..., 12 = Dezember): Verhaeltnis des Durchschnittsverbrauchs
    dieses Kalendermonats zum Gesamtdurchschnitt ueber das Fenster.

    Erfordert MINDESTENS MIN_YEARS_FOR_SEASONALITY volle Kalenderjahre im
    Fenster (siehe Modul-Docstring) - bei weniger Jahren liefert diese
    Funktion None, statt eine nicht belastbare Zahl zu suggerieren (ein
    einzelnes Vorkommen eines Monats kann ein Sondereffekt sein, kein
    wiederkehrendes Muster).

    Args:
        month_sequence: lueckenlose Liste der Monatsersten im Fenster.
        values: Verbrauchswerte in identischer Reihenfolge zu month_sequence,
            ggf. NaN.

    Returns:
        Dict {Kalendermonat (1-12): Index}, oder None, falls die Mindest-
        anzahl Jahre nicht erreicht wird oder der Gesamtdurchschnitt 0/NaN
        ist (Division durch 0 vermieden).
    """
    distinct_years = {m.year for m in month_sequence}
    if len(distinct_years) < MIN_YEARS_FOR_SEASONALITY:
        return None

    overall_mean = np.nanmean(values) if not np.all(np.isnan(values)) else float("nan")
    if pd.isna(overall_mean) or overall_mean == 0:
        return None

    index_by_month: dict[int, float] = {}
    for calendar_month in range(1, 13):
        month_values = [
            values[i]
            for i, m in enumerate(month_sequence)
            if m.month == calendar_month and not np.isnan(values[i])
        ]
        if not month_values:
            continue
        index_by_month[calendar_month] = float(np.mean(month_values) / overall_mean)

    return index_by_month or None


def _format_seasonality_hint(index_by_month: dict[int, float] | None) -> str:
    """Wandelt den Saisonalitaets-Index in einen kompakten, lesbaren Text um
    fuer die Report-Spalte 'Saisonalitaet_Hinweis' - die rohen Index-Werte
    je Kalendermonat sind als Excel-Spalte unhandlich (12 zusaetzliche
    Spalten je Material), ein zusammenfassender Text genuegt fuer die erste
    Einordnung. Detailauswertung bleibt ueber die rohen Monatsspalten im
    selben Reiter weiterhin moeglich.

    Nennt nur Kalendermonate, die deutlich (>= 15%) vom Durchschnitt
    abweichen - bei geringerer Abweichung wird 'Kein ausgepraegtes Muster'
    zurueckgegeben, um Rauschen nicht als Muster zu ueberinterpretieren.

    Args:
        index_by_month: Ergebnis von _calculate_seasonality_index().

    Returns:
        Lesbarer Text, z.B. 'Dez +24%, Jan -18%', oder ein erklaerender
        Platzhaltertext, falls kein Index berechnet werden konnte oder kein
        Monat deutlich abweicht.
    """
    if index_by_month is None:
        return f"Zu wenig Historie (< {MIN_YEARS_FOR_SEASONALITY} Jahre)"

    month_names = {
        1: "Jan", 2: "Feb", 3: "Mär", 4: "Apr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Dez",
    }

    notable = [
        (calendar_month, index_value)
        for calendar_month, index_value in index_by_month.items()
        if abs(index_value - 1.0) >= 0.15
    ]
    if not notable:
        return "Kein ausgepraegtes Muster"

    notable.sort(key=lambda item: abs(item[1] - 1.0), reverse=True)
    parts = []
    for calendar_month, index_value in notable[:4]:
        deviation_percent = (index_value - 1.0) * 100
        sign = "+" if deviation_percent >= 0 else ""
        parts.append(f"{month_names[calendar_month]} {sign}{deviation_percent:.0f}%")

    return ", ".join(parts)


def build_trend_table(
    long_df: pd.DataFrame,
    window_months: int = DEFAULT_TREND_WINDOW_MONTHS,
    reference_date: _dt.date | None = None,
) -> tuple[pd.DataFrame, dict[str, list[str]], pd.DataFrame]:
    """Hauptfunktion fuer Phase 3: baut die vollstaendige Trend-Tabelle, eine
    Zeile pro Material, mit rohen Monatsspalten UND abgeleiteten Kennzahlen.

    Dies ist die Funktion, die main.py/main_single.py fuer den Reiter
    'Trend_MVER' aufrufen (siehe report_builder.py).

    Erweitert (25.06.2026, Transparenz-Anforderung Max): liefert zusaetzlich
    'details_df' fuer den neuen Reiter 'Trend_Details' - macht die
    Sondereffekt-Erkennung je Material UND je Monat vollstaendig
    nachvollziehbar (Median, MAD, modifizierter Z-Score je Zelle), statt
    nur das Endergebnis (Sondereffekt_Monate, Verbrauch_Streuung_MAD) zu
    zeigen. Hintergrund: Max moechte Analyseergebnisse gegenueber dem COO
    fundiert erklaeren koennen, nicht nur das Ergebnis selbst zeigen.

    Args:
        long_df: Ergebnis von mver_loader.load_full_mver_export() (Material,
            Periode, Verbrauch) - dieselbe Quelle wie fuer Phase 2, hier aber
            mit einem typischerweise GROESSEREN Fenster (Default 36 statt 12
            Monate) erneut ausgewertet.
        window_months: Groesse des rollierenden Fensters in Monaten (siehe
            DEFAULT_TREND_WINDOW_MONTHS - flexibel aenderbar, analog zu
            Phase 2, siehe Modul-Docstring).
        reference_date: Stichtag, ab dem rueckwirkend gerechnet wird. Sollte
            mit dem reference_date von mver_loader.load_full_mver_export()
            uebereinstimmen (siehe _window_bounds), damit Future-Monat-
            Markierung und Fensterberechnung konsistent sind. Default:
            heutiges Datum.

    Returns:
        Tuple (trend_df, outlier_cells, details_df):

        trend_df: DataFrame mit Spalten:
            - Material
            - <Monatsspalten 'YYYY-MM'>, chronologisch (ROHE, unbereinigte
              Verbrauchswerte, NaN fuer fehlende Monate) - erfuellt die
              Anforderung "36 Monate Historie monatlich in einem getrennten
              Reiter". Sondereffekt-Monate sind hier bewusst NICHT
              herausgerechnet, nur in den folgenden Kennzahlen.
            - Trend_Steigung_Stueck_Monat (float, Stueck/Monat, lineare
              Regression, OHNE die als Sondereffekt erkannten Monate -
              siehe Klaerung mit Max, 24.06.2026)
            - Trend_Einstufung ('Steigend'/'Fallend'/'Stabil'/'Unbekannt'),
              ebenfalls auf Basis der um Sondereffekte bereinigten Werte
            - Sondereffekt_Monate (str, kommagetrennte Liste der als
              Ausreisser erkannten Monate, leer falls keine)
            - Anzahl_Sondereffekte (int)
            - Verbrauch_Streuung_MAD (float, Median Absolute Deviation -
              Streuungsmass, robust, fuer spaetere Wiederverwendung in
              Phase 5 als Safety-Stock-Input)
            - Saisonalitaet_Hinweis (str, lesbare Kurzzusammenfassung,
              ebenfalls auf Basis der um Sondereffekte bereinigten Werte)
            - Anzahl_Monate_Verfuegbar (int, wie viele der window_months
              Monate tatsaechlich einen Rohwert hatten - unabhaengig von
              der Sondereffekt-Bereinigung)

        outlier_cells: Dict {Material: [Liste der Monatsspalten-Namen, die
            als Sondereffekt erkannt wurden]}, z.B. {'00...385325':
            ['2024-05']}. Dient report_builder.py als Grundlage fuer die
            bedingte Formatierung (farbliche Markierung) der betroffenen
            Zellen in den rohen Monatsspalten - die Zahl selbst bleibt
            unveraendert sichtbar, nur optisch hervorgehoben (Klaerung mit
            Max, 24.06.2026: Peak soll im Report erkennbar bleiben, auch
            wenn er fuer Trend/Saisonalitaet herausgerechnet wurde).

        details_df: DataFrame fuer den neuen Reiter 'Trend_Details'
            (25.06.2026, Transparenz-Anforderung), eine Zeile pro Material,
            mit Spalten:
            - Material
            - Verbrauch_Median (float, Median der gueltigen Rohwerte im
              Fenster - Basiswert fuer den Z-Score, bisher nur intern in
              _detect_outlier_periods() berechnet, nicht exportiert)
            - Verbrauch_Streuung_MAD (float, identisch zu trend_df - hier
              wiederholt, damit der Reiter eigenstaendig lesbar ist, ohne
              auf Trend_MVER zurueckspringen zu muessen)
            - Je Monat im Fenster ZWEI Spalten statt einer: '<YYYY-MM>'
              (= identischer Rohwert wie in trend_df) und
              '<YYYY-MM>_ZScore' (modifizierter Z-Score dieses Monats,
              NaN wo nicht berechenbar). Damit ist je Zelle nachrechenbar:
              ZScore = 0.6745 * (Wert - Verbrauch_Median) / Verbrauch_Streuung_MAD,
              und ersichtlich, warum ein Monat ueber/unter der Schwelle
              OUTLIER_MODIFIED_Z_THRESHOLD (3.5) liegt.

    Raises:
        ValueError: wenn long_df leer ist oder die erwarteten Spalten
            (Material, Periode, Verbrauch) fehlen.
    """
    required_columns = {"Material", "Periode", "Verbrauch"}
    missing = required_columns - set(long_df.columns)
    if missing:
        raise ValueError(
            f"trend_analysis.build_trend_table() erwartet die Spalten "
            f"{sorted(required_columns)}, es fehlen: {sorted(missing)}."
        )

    if long_df.empty:
        raise ValueError(
            "trend_analysis.build_trend_table() hat einen leeren long_df "
            "erhalten - keine MVER-Daten zum Auswerten vorhanden."
        )

    window_start, window_end = _window_bounds(window_months, reference_date)
    month_sequence = _month_sequence(window_start, window_end)

    windowed = long_df[
        (long_df["Periode"] >= window_start) & (long_df["Periode"] <= window_end)
    ]

    pivoted = _pivot_monthly_columns(windowed, month_sequence)
    month_columns = [m.strftime(_MONTH_COLUMN_DATE_FORMAT) for m in month_sequence]

    records = []
    outlier_cells: dict[str, list[str]] = {}
    detail_records = []
    for _, row in pivoted.iterrows():
        material = row["Material"]
        raw_values = row[month_columns].to_numpy(dtype=float)

        # Sondereffekte ZUERST auf den rohen Werten erkennen (siehe
        # _detect_outlier_periods) - die rohen Monatsspalten im Output
        # bleiben davon unberuehrt, nur diese Zwischenvariable wird fuer
        # Trend/Saisonalitaet bereinigt (Klaerung mit Max, 24.06.2026: ein
        # einzelner Sondereffekt-Monat soll Trend/Saisonalitaet nicht
        # verzerren, da sonst nicht die "echte Entwicklung" sichtbar wird,
        # sondern nur der groesste Einmaleffekt). Kein Ersatzwert (z.B.
        # Median) wird eingesetzt - der Monat wird wie eine fehlende
        # Messung behandelt (NaN), die Trendgerade/Saisonalitaet stuetzen
        # sich nur auf die uebrigen, unauffaelligen Monate.
        outlier_periods, mad, median, z_scores = _detect_outlier_periods(
            month_columns, raw_values
        )
        if outlier_periods:
            outlier_cells[material] = outlier_periods
        outlier_period_set = set(outlier_periods)
        cleaned_values = np.array([
            float("nan") if month_columns[i] in outlier_period_set else raw_values[i]
            for i in range(len(raw_values))
        ])

        slope, _intercept = _linear_trend(cleaned_values)
        mean_value = (
            float(np.nanmean(cleaned_values))
            if not np.all(np.isnan(cleaned_values))
            else float("nan")
        )
        trend_label = _classify_trend(slope, mean_value, months_in_window=len(month_columns))

        seasonality_index = _calculate_seasonality_index(month_sequence, cleaned_values)
        seasonality_hint = _format_seasonality_hint(seasonality_index)

        months_available = int(np.sum(~np.isnan(raw_values)))

        records.append({
            "Material": material,
            "Trend_Steigung_Stueck_Monat": slope,
            "Trend_Einstufung": trend_label,
            "Sondereffekt_Monate": ", ".join(outlier_periods),
            "Anzahl_Sondereffekte": len(outlier_periods),
            "Verbrauch_Streuung_MAD": mad,
            "Saisonalitaet_Hinweis": seasonality_hint,
            "Anzahl_Monate_Verfuegbar": months_available,
        })

        # Transparenz-Reiter 'Trend_Details' (25.06.2026): je Monat zwei
        # Spalten (Rohwert + Z-Score), damit jede Sondereffekt-Erkennung
        # einzeln nachrechenbar ist (siehe Docstring). Verbrauch_Median wird
        # hier zusaetzlich zu Verbrauch_Streuung_MAD gefuehrt, da beide
        # zusammen den Z-Score je Zelle ergeben.
        detail_row = {"Material": material, "Verbrauch_Median": median, "Verbrauch_Streuung_MAD": mad}
        for i, month_col in enumerate(month_columns):
            detail_row[month_col] = raw_values[i]
            detail_row[f"{month_col}_ZScore"] = z_scores[i]
        detail_records.append(detail_row)

    metrics_df = pd.DataFrame.from_records(records)
    details_df = pd.DataFrame.from_records(detail_records)

    result = pivoted.merge(metrics_df, on="Material", how="left")
    return result, outlier_cells, details_df

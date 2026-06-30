"""
mver_loader.py
================
Einlesen der SAP-Tabelle MVER (monatliche Verbrauchshistorie je Material)
als eigenstaendige, dritte Datenquelle neben ZMLAG und SE16XXL (siehe
Projektstatus.md, Abschnitt "Datenquellen-Strategie" - bewusste
Architekturentscheidung, diese drei Quellen NICHT zu einer Tabelle zu
vermischen).

Format des MVER-Exports (Breitformat, eine Zeile pro Material UND Geschaeftsjahr):
    Mandant | Material | Werk | GeschJahr | PerKennz. | Anzahl
        | GesVerbr (Jan) | GesVerbr (Feb) | ... | GesVerbr (Dez)

Wichtige Eigenheiten dieses Exports (Stand 24.06.2026, anhand Beispieldaten
von Max verifiziert):
    - Alle 12 Verbrauchsspalten heissen identisch "GesVerbr" - sie sind daher
      NUR ueber ihre POSITION nach den Fixspalten unterscheidbar (1. = Januar,
      ..., 12. = Dezember). Eine namensbasierte Spaltenauswahl (wie bei den
      ZMLAG-Stichtagsspalten) ist hier NICHT moeglich.
    - 'GeschJahr' entspricht dem Kalenderjahr (Januar-Dezember, keine
      Abweichung), bestaetigt von Max.
    - Anzahl der Geschaeftsjahr-Zeilen pro Material ist NICHT fix. Max kann
      wahlweise 3, 4, 5, 6+ Jahre Historie exportieren (Klaerung 24.06.2026):
      ein aktuell laufendes Jahr ist dabei typischerweise nur teilweise
      befuellt (bereits abgelaufene Monate mit echtem Wert, noch nicht
      abgelaufene Monate als 0 statt leer/NaN - SAP unterscheidet hier NICHT
      zwischen "kein Verbrauch" und "Monat noch nicht erreicht"). Es gibt
      daher KEINE feste erwartete Zeilenanzahl mehr (siehe Aenderungshistorie
      Projektstatus.md - EXPECTED_YEARS_PER_MATERIAL wurde entfernt).
    - 'Anzahl' wird bewusst NICHT uebernommen (Bedeutung fachlich nicht
      geklaert, fuer unsere Berechnungen nicht benoetigt).
    - Werte koennen sowohl als echte Zahl (von Excel/SAP bereits korrekt als
      float gespeichert) als auch als Text im deutschen Zahlenformat
      ("179,000" = 179.0, Komma als Dezimaltrennzeichen) vorliegen - siehe
      _parse_german_number(). Es wird defensiv beides unterstuetzt, da das
      tatsaechliche Exportverhalten je nach SAP-/Excel-Version variieren kann.

Flexible Fenstergroesse statt fixer Jahreszahl (Klaerung mit Max, 24.06.2026):
    Max kann je nach Bedarf (eigene Analyse vs. Wunsch des Vorgesetzten nach
    laengerer Historie fuer robustere Stochastik) 36, 48 oder mehr Monate
    Historie exportieren UND auswerten wollen. calculate_average_monthly_
    consumption() nimmt daher 'window_months' als Parameter (Default 12 fuer
    die "Ø-Verbrauch letzte 12 Monate"-Anforderung aus Phase 2), und
    detect_consumption_start() erkennt je Material den vermutlichen Beginn
    der echten Verbrauchshistorie (erster Monat mit Wert != 0) - so wird ein
    neueres Material, das schlicht noch keine 36/48 Monate alt ist, NICHT
    faelschlich als Datenluecke gemeldet, waehrend ein aelteres Material mit
    einer echten Lieferunterbrechung weiterhin erkennbar bleibt.

Future-Monate (Klaerung mit Max, 24.06.2026):
    Da SAP fuer noch nicht abgelaufene Monate des laufenden Geschaeftsjahres
    eine 0 statt eines leeren Werts liefert, werden Monate, die relativ zu
    reference_date in der Zukunft liegen, beim Einlesen explizit als NaN
    markiert (nicht als 0 belassen) - sonst wuerden sie faelschlich als
    "kein Verbrauch" in den Durchschnitt einfliessen, sobald sie zufaellig
    ins Berechnungsfenster fallen.

Oeffentliche Funktionen:
    load_full_mver_export(reference_date) -> langes DataFrame (Material,
        Periode [erster Tag des Monats], Verbrauch); Future-Monate ab
        reference_date werden zu NaN.
    calculate_average_monthly_consumption(long_df, window_months, reference_date)
        -> je Material der rollierende Durchschnittsverbrauch der letzten
        'window_months' ABGESCHLOSSENEN Kalendermonate vor reference_date
        (Entscheidung Max, 24.06.2026 - bewusst rollierend ueber
        Jahresgrenzen hinweg, nicht an Kalenderjahre gebunden).
    detect_consumption_start(long_df) -> je Material der frueheste Monat mit
        Verbrauch != 0 (Hilfsfunktion fuer Datenqualitaets-Einordnung).
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pandas as pd

import config


# Fixspalten am Anfang der MVER-Tabelle, in dieser Reihenfolge erwartet.
MVER_FIXED_COLUMNS = ["Mandant", "Material", "Werk", "GeschJahr", "PerKennz.", "Anzahl"]

# Anzahl der monatlichen Verbrauchsspalten nach den Fixspalten (Jan..Dez).
MVER_MONTH_COUNT = 12

# Default-Fenstergroesse fuer die Durchschnittsverbrauchsberechnung (Phase 2:
# "Ø-Verbrauch der letzten 12 Monate"). Ueber den Parameter window_months in
# calculate_average_monthly_consumption() jederzeit auf 36, 48 etc. aenderbar,
# falls eine laengere Historie fuer robustere Stochastik gewuenscht ist
# (Klaerung mit Max, 24.06.2026 - z.B. auf Wunsch des Vorgesetzten).
DEFAULT_WINDOW_MONTHS = 12


def _parse_german_number(value: object) -> float:
    """Wandelt einen Verbrauchswert in einen float um - robust gegenueber
    beiden moeglichen Eingabeformen:
        - bereits numerisch (int/float), z.B. von pandas automatisch erkannt
        - Text im deutschen Zahlenformat, z.B. '179,000' (Komma als Dezimal-
          trennzeichen, Punkt als Tausendertrennzeichen)

    Leere Werte / NaN werden zu float('nan').
    """
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return float("nan")

    # Deutsches Zahlenformat: Punkt entfernen (Tausendertrennzeichen),
    # Komma zu Punkt (Dezimaltrennzeichen).
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return float("nan")


def _validate_mver_columns(df: pd.DataFrame, file_path: Path) -> None:
    """Prueft, ob die erwarteten Fixspalten vorhanden sind und mindestens
    MVER_MONTH_COUNT weitere Spalten (die 12 GesVerbr-Monatsspalten) danach
    folgen.

    Raises:
        ValueError: wenn Fixspalten fehlen oder zu wenige Monatsspalten
            vorhanden sind.
    """
    missing = [col for col in MVER_FIXED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            f"MVER-Export '{file_path.name}' fehlen erwartete Spalten: {missing}\n"
            f"Vorhandene Spalten: {list(df.columns)}"
        )

    month_columns = [c for c in df.columns if c not in MVER_FIXED_COLUMNS]
    if len(month_columns) < MVER_MONTH_COUNT:
        raise ValueError(
            f"MVER-Export '{file_path.name}' enthaelt nur {len(month_columns)} "
            f"Verbrauchsspalten nach den Fixspalten, erwartet werden "
            f"{MVER_MONTH_COUNT} (Januar bis Dezember)."
        )


def _wide_to_long(
    df: pd.DataFrame, file_path: Path, reference_date: _dt.date
) -> pd.DataFrame:
    """Wandelt das MVER-Breitformat (eine Zeile pro Material+Geschaeftsjahr,
    12 Monatsspalten) in ein langes Zeitreihenformat um: eine Zeile pro
    Material+Periode.

    Die 12 Monatsspalten werden POSITIONSBASIERT als Januar..Dezember
    interpretiert (siehe Modul-Docstring - sie sind alle identisch benannt
    und daher nicht ueber den Spaltennamen unterscheidbar).

    Monate, die relativ zu reference_date in der Zukunft liegen (siehe
    Modul-Docstring, "Future-Monate"), werden bewusst als NaN statt als 0
    abgelegt - SAP liefert hier ebenfalls 0, das aber "Monat noch nicht
    abgelaufen" bedeutet, nicht "kein Verbrauch".

    Returns:
        DataFrame mit Spalten: Material, Periode (erster Tag des Monats,
        als datetime.date), Verbrauch (float, NaN fuer Future-Monate).
    """
    month_columns = [c for c in df.columns if c not in MVER_FIXED_COLUMNS][:MVER_MONTH_COUNT]
    first_day_of_current_month = reference_date.replace(day=1)

    records = []
    for _, row in df.iterrows():
        material = row["Material"]
        jahr = int(row["GeschJahr"])
        for month_idx, col in enumerate(month_columns, start=1):
            periode = _dt.date(jahr, month_idx, 1)
            if periode >= first_day_of_current_month:
                # Zukunftsmonat (oder der laufende, noch nicht abgeschlossene
                # Monat) - SAP-Wert hier ignorieren, da er nicht zwischen
                # "kein Verbrauch" und "noch nicht erreicht" unterscheidet.
                verbrauch = float("nan")
            else:
                verbrauch = _parse_german_number(row[col])
            records.append({"Material": material, "Periode": periode, "Verbrauch": verbrauch})

    long_df = pd.DataFrame.from_records(records)

    if long_df["Verbrauch"].isna().all():
        raise ValueError(
            f"MVER-Export '{file_path.name}': nach der Umwandlung sind ALLE "
            f"Verbrauchswerte leer/nicht interpretierbar - bitte Zahlenformat "
            f"im Export pruefen (oder reference_date, falls dadurch alle "
            f"Monate als Zukunft markiert wurden)."
        )

    return long_df


def full_mver_export_exists() -> bool:
    """Prueft, ob der MVER-Gesamtexport vorliegt."""
    return config.MVER_FULL_EXPORT_PATH.exists()


def load_full_mver_export(reference_date: _dt.date | None = None) -> pd.DataFrame:
    """Liest den MVER-Gesamtexport ein und wandelt ihn in ein langes
    Zeitreihenformat um (eine Zeile pro Material+Periode).

    Akzeptiert eine BELIEBIGE Anzahl an Geschaeftsjahr-Zeilen pro Material
    (3, 4, 5, 6+ Jahre) - es gibt keine feste Erwartung mehr (siehe Modul-
    Docstring). Monate in der Zukunft relativ zu reference_date werden zu
    NaN (siehe _wide_to_long).

    Args:
        reference_date: Stichtag fuer die Future-Monat-Erkennung. Default:
            heutiges Datum (datetime.date.today()).

    Returns:
        DataFrame mit Spalten: Material (str), Periode (datetime.date, erster
        Tag des jeweiligen Monats), Verbrauch (float, Stueck, NaN fuer
        Future-Monate oder nicht interpretierbare Werte).

    Raises:
        FileNotFoundError: wenn config.MVER_FULL_EXPORT_PATH nicht existiert.
        ValueError: wenn erwartete Spalten fehlen oder keine Werte interpretierbar sind.
    """
    if reference_date is None:
        reference_date = _dt.date.today()

    file_path = config.MVER_FULL_EXPORT_PATH
    if not file_path.exists():
        raise FileNotFoundError(
            f"MVER-Export nicht gefunden: {file_path}\n"
            f"Bitte MVER-Export dort ablegen (Dateiname: "
            f"'{config.MVER_FULL_EXPORT_FILENAME}')."
        )

    df = pd.read_excel(file_path, dtype={"Material": str})
    _validate_mver_columns(df, file_path)

    # Normalisierung auf die SAP-native 18-stellige Form (siehe config.py,
    # Abschnitt "Materialnummern-Normalisierung"). MVER liefert bereits 18
    # Stellen, daher hier im Normalfall ein No-Op - die explizite Anwendung
    # macht den Merge mit ZMLAG (siehe coverage_analysis.py) aber robust
    # gegenueber zukuenftigen Format-Aenderungen, statt sich stillschweigend
    # auf das aktuelle Exportformat zu verlassen.
    df["Material"] = df["Material"].apply(config.normalize_material_number)

    long_df = _wide_to_long(df, file_path, reference_date)

    return long_df


def detect_consumption_start(long_df: pd.DataFrame) -> pd.DataFrame:
    """Ermittelt je Material den fruehesten Monat mit einem Verbrauchswert
    ungleich 0 - als Annaeherung an den Beginn der "echten" Verbrauchshistorie.

    Hintergrund (Klaerung mit Max, 24.06.2026): Ein Material kann in der
    Exporthistorie absichtlich mit weniger Jahren auftauchen, einfach weil es
    erst seit kuerzerer Zeit ueberhaupt bewirtschaftet wird - das ist KEIN
    Datenfehler. Diese Funktion liefert die Grundlage, um das von einer
    echten Verbrauchsluecke (Material existiert lange, hat aber zwischendurch
    keinen Verbrauch) zu unterscheiden, OHNE eine feste Jahreszahl
    vorauszusetzen.

    Args:
        long_df: Ergebnis von load_full_mver_export().

    Returns:
        DataFrame mit Spalten: Material, Erster_Verbrauch_Monat (datetime.date
        oder NaT, falls ein Material durchgehend 0/NaN hat).
    """
    nonzero = long_df[long_df["Verbrauch"].fillna(0) != 0]
    result = (
        nonzero.groupby("Material")["Periode"]
        .min()
        .reset_index()
        .rename(columns={"Periode": "Erster_Verbrauch_Monat"})
    )
    return result


def calculate_average_monthly_consumption(
    long_df: pd.DataFrame,
    window_months: int = DEFAULT_WINDOW_MONTHS,
    reference_date: _dt.date | None = None,
) -> pd.DataFrame:
    """Berechnet je Material den durchschnittlichen Monatsverbrauch ueber die
    letzten 'window_months' ABGESCHLOSSENEN Kalendermonate vor reference_date
    (rollierendes Fenster, bewusst ueber Jahresgrenzen hinweg - Entscheidung
    Max, 24.06.2026, NICHT an Kalenderjahre gebunden).

    Die Fenstergroesse ist bewusst FLEXIBEL (siehe Modul-Docstring) - Default
    12 Monate fuer die Phase-2-Anforderung "Ø-Verbrauch letzte 12 Monate",
    aber z.B. auch 36 oder 48 nutzbar, falls eine laengere Historie fuer
    robustere stochastische Auswertungen gewuenscht ist.

    Beispiel: reference_date = Juni 2026, window_months=12 -> letzter
    abgeschlossener Monat ist Mai 2026 -> Fenster = Juni 2025 bis Mai 2026.

    Args:
        long_df: Ergebnis von load_full_mver_export() (Material, Periode,
            Verbrauch).
        window_months: Groesse des rollierenden Fensters in Monaten.
        reference_date: Stichtag, ab dem rueckwirkend gerechnet wird. Default:
            heutiges Datum (datetime.date.today()). Sollte mit dem
            reference_date von load_full_mver_export() uebereinstimmen, damit
            Future-Monat-Markierung und Fensterberechnung konsistent sind.

    Returns:
        DataFrame mit Spalten:
            - Material
            - Ø_Verbrauch_<window_months>M (Stueck/Monat)
            - Anzahl_Monate_Verfuegbar (wie viele der window_months Monate
              tatsaechlich einen Wert hatten - bei Werten kleiner
              window_months ist der Durchschnitt mit Vorsicht zu interpretieren,
              z.B. weil das Material noch nicht so lange bewirtschaftet wird)
            - Ø_Verbrauch_Fenster (generischer Alias derselben Kennzahl,
              unabhaengig von window_months - fuer nachgelagerten Code wie
              coverage_analysis.py, der die Fenstergroesse nicht kennen muss)
    """
    if reference_date is None:
        reference_date = _dt.date.today()

    # Letzter ABGESCHLOSSENER Kalendermonat vor reference_date.
    first_day_of_ref_month = reference_date.replace(day=1)
    last_completed_month_end = first_day_of_ref_month - _dt.timedelta(days=1)
    window_end = last_completed_month_end.replace(day=1)

    # window_months Monate rueckwirkend ab window_end, inklusive.
    window_start_year = window_end.year
    window_start_month = window_end.month - (window_months - 1)
    while window_start_month <= 0:
        window_start_month += 12
        window_start_year -= 1
    window_start = _dt.date(window_start_year, window_start_month, 1)

    windowed = long_df[(long_df["Periode"] >= window_start) & (long_df["Periode"] <= window_end)]

    avg_col = f"Ø_Verbrauch_{window_months}M"
    result = (
        windowed.groupby("Material")
        .agg(**{avg_col: ("Verbrauch", "mean"), "Anzahl_Monate_Verfuegbar": ("Verbrauch", "count")})
        .reset_index()
    )

    # Generischer Spaltenname zusaetzlich, damit nachgelagerter Code (z.B.
    # coverage_analysis.py) nicht von window_months abhaengen muss.
    result["Ø_Verbrauch_Fenster"] = result[avg_col]

    return result

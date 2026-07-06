# Projektanforderungen – GMTS (Global Manufacturing Transition Strategy)

## Ziel

Integriertes Dispositions- und Fruehwarnsystem zur Bewertung von Lagerbestand,
Versorgungsrisiko und Bestandsentwicklung auf Material-Nr.-Ebene, als
Entscheidungsgrundlage fuer die COO-Kommunikation. Vollstaendiger Neuaufbau in
Python auf Basis von SAP-Exporten (siehe README.md, "Datenquellen-Strategie").

## Umfang

- Technologie-agnostischer Standard-Lauf ueber alle acht im Unternehmen
  vorhandenen Produkttechnologien (CAT 4, DCS 0.5/1.0/2.0, EL.NET, EL.MOTION,
  EL.INDUSTRY, EL.CAPTURE) plus automatischer Erfassung neuer/unbekannter
  Technologien.
- Konsolidierte Excel-Reports und ein interaktives HTML-Dashboard je
  Technologie sowie ein Gesamt-Report ueber alle Technologien.
- Kennzahlen: Bestandsbewertung, Reichweite, Trend-/Forecast-Analyse,
  Monte-Carlo-Simulation, Safety Stock/Reorder Point, Fruehwarnsystem,
  Working Capital/ROI.

## Stakeholder

- **Auftraggeber/fachliche Freigabe:** Max (Klaerung fachlicher Annahmen,
  z.B. Service-Level-Matrix, Finanzparameter – siehe `kpis/kpi_definitions.md`).
- **Weitere fachliche Abstimmung:** Joachim (Ziel-Servicegrade).

## Nicht-Ziele

- Kein Uebernehmen von Logik aus frueheren Excel/Python-Prototypen (siehe
  README.md, "Methodische Referenz") – vollstaendiger Neuaufbau.
- Keine ROI-Berechnung gegen Implementierungskosten (siehe config.py,
  Abschnitt "Finanzparameter").

## Laufender Status

Siehe `docs/Projektstatus.md` fuer Methodik, offene Punkte und vollstaendige
Aenderungshistorie.

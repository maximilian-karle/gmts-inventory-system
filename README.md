# Global Manufacturing Transition Strategy (GMTS)

Integriertes Dispositions- und Fruehwarnsystem zur Bewertung von Lagerbestand,
Versorgungsrisiko und Bestandsentwicklung auf Material-Nr.-Ebene. Vollstaendiger
Neuaufbau in Python auf Basis von SAP-Exporten, mit konsolidierten Excel-Reports
und interaktivem HTML-Dashboard.

## Template-Charakter

Dieses Projekt ist bewusst **technologie-agnostisch** aufgebaut. Ein Standard-Lauf
(`python main.py`) verarbeitet automatisch alle acht im Unternehmen vorhandenen
Produkttechnologien in einem Durchgang und erzeugt zusaetzlich einen konsolidierten
Gesamt-Report:

- CAT 4
- DCS 0.5
- DCS 1.0
- DCS 2.0
- EL Capture
- EL Motion
- EL.NET
- (unzugeordneter Rest-Bucket)

Fuer den gezielten Nachlauf einer einzelnen Technologie steht `main_single.py`
zur Verfuegung.

## Datenquellen-Strategie

Das Projekt nutzt bewusst **drei dauerhaft getrennte SAP-Datenquellen**, die nicht
zu einer einzigen Tabelle zusammengefuehrt werden, sondern erst auf Material-Ebene
im Python-Code per Join verbunden werden:

| Quelle | Liefert |
|---|---|
| **ZMLAG** | Materialstamm, Bestandswert (3 Stichtage + Laufender Wert), ABC/XYZ-Kennzeichen, Technologie-Zuordnung |
| **MVER** | 36+ Monate Verbrauchshistorie (Breitformat → Long-Format) |
| **SE16XXL "Inventory"** | Ist-Bestandsmenge in Stueck, Materialstatus (Lebenszyklus) |
| **SE16XXL "Dispo_ABCXYZ"** | Wiederbeschaffungszeit, Sicherheitsbestand, Meldebestand, Preise, Variationskoeffizient, kombiniertes ABCXYZ-Kennzeichen |

Bevorzugt wird je Quelle ein einzelner Gesamtexport ueber alle Technologien
(`input_data/<quelle>_export.xlsx`), intern anhand der SAP-Spalte `Technologie`
aufgeteilt. Der aeltere Einzelordner-Ansatz (`input_data/<technologie_slug>/`)
bleibt als Fallback erhalten.

## Projektstruktur

```
gmts/
├── input_data/                   # SAP-Exporte (lokal, nicht versioniert)
├── output_data/                  # generierte Excel-Reports & HTML-Dashboards (lokal, nicht versioniert)
├── src/
│   ├── config.py                  # zentrale Konfiguration (Pfade, Konstanten, Service-Level-Matrix)
│   ├── data_loader.py             # ZMLAG-Einlesen & Validierung
│   ├── mver_loader.py             # MVER-Einlesen (Verbrauchshistorie)
│   ├── se16xxl_loader.py          # SE16XXL "Inventory"-Einlesen
│   ├── dispo_abcxyz_loader.py     # SE16XXL "Dispo_ABCXYZ"-Einlesen (Lead Time etc.)
│   ├── stock_analysis.py          # Basis-Bestandskennzahlen
│   ├── coverage_analysis.py       # Reichweite (Bestand ÷ Verbrauch)
│   ├── trend_analysis.py          # Trend-/Sondereffekt-/Saisonalitaetsanalyse
│   ├── forecast_analysis.py       # Modellauswahl-Forecast (Naive/SES/Holt/Holt-Winters)
│   ├── simulation_analysis.py     # Monte-Carlo-Verbrauchsprognose (Baustein A)
│   ├── safety_stock.py            # Safety Stock & Reorder Point (ROP)
│   ├── executive_summary.py       # Regelbasierte Prioritaetsklassifikation
│   ├── working_capital.py         # Working Capital / ROI-Berechnung
│   ├── report_builder.py          # Excel-Report-Erstellung (alle Reiter)
│   ├── html_dashboard.py          # Interaktives Plotly-HTML-Dashboard
│   ├── main.py                    # Orchestrierung: alle Technologien + konsolidiert
│   └── main_single.py             # Orchestrierung: einzelne Technologie
├── docs/
│   └── Projektstatus.md           # laufend gepflegter Projektstatus (Methodik, Entscheidungen, Historie)
├── requirements.txt
└── README.md
```

## Setup (Windows / PowerShell)

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
python -m venv venv
.\venv\Scripts\activate
python -m pip install -r requirements.txt
```

Hinweis: In PowerShell ist `pip` haeufig nicht direkt im PATH registriert -
`python -m pip` umgeht das zuverlaessig.

## Ablauf eines Analyse-Laufs

1. SAP-Exporte in `input_data/` ablegen (siehe Datenquellen-Strategie oben).
2. Ausfuehren:
   ```powershell
   cd src
   python main.py
   ```
3. Ergebnis je Technologie: `output_data/<technologie_slug>/<technologie_slug>_bestandsreport.xlsx`
   sowie `<technologie_slug>_dashboard.html`.
4. Konsolidierter Gesamt-Report: `output_data/alle_technologien_bestandsreport.xlsx`
   sowie `output_data/alle_technologien_dashboard.html`.

Fuer den Nachlauf einer einzelnen Technologie: `python main_single.py`.

## Aktueller Stand

Phasen 1-6 sind abgeschlossen (Stand 30.06.2026): Bestandsbewertung, Reichweite,
Trend-/Forecast-Analyse, Monte-Carlo-Simulation, Wiederbeschaffungszeit, Safety
Stock/ROP, Fruehwarnsystem, sowie der konsolidierte COO-Report inkl. Executive
Summary, nativen Excel-Dashboards, interaktivem HTML-Dashboard und Working
Capital/ROI-Berechnung. Details zu Methodik, offenen Punkten und vollstaendiger
Aenderungshistorie siehe `docs/Projektstatus.md`.

## Methodische Referenz

Die fachliche Methodik (Safety Stock, ROP, Monte-Carlo-Simulation, Working
Capital/ROI) orientiert sich an Dmitry Ivanov, *Global Supply Chain and
Operations Management* (2019) und *Introduction to Supply Chain Analytics*
(2021), sowie einem internen Whitepaper zu Forecast-Driven Inventory
Optimization. Der Code in diesem Projekt ist ein vollstaendiger Neuaufbau und
uebernimmt keine Logik aus frueheren Excel/Python-Prototypen - diese dienen
ausschliesslich als methodische Referenz und sind bewusst nicht Teil dieses
Repositories.

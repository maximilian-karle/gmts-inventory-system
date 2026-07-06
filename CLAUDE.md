# CLAUDE.md

Hinweise fuer Claude-Sessions in diesem Repository.

## Projektstruktur

- `00_Charter/` – Projektauftrag/-anforderungen.
- `kpis/` – lesbare Referenz der Kennzahlen-Definitionen (Quelle: `src/config.py`).
- `methodology/` – methodische Referenzliteratur, lokal, nicht versioniert.
- `analysis/`, `dashboards/`, `reports/` – Platzhalter fuer Ad-hoc-Analysen,
  Dashboard-Spezifikationen bzw. finale Stakeholder-Reports (aktuell leer).
- `data/input/` – SAP-Exporte, lokal, nicht versioniert.
- `data/output/` – generierte Excel-Reports & HTML-Dashboards, lokal, nicht versioniert.
- `src/` – der eigentliche Anwendungscode (siehe README.md fuer Modulübersicht).
- `docs/Projektstatus.md` – laufend gepflegter Projektstatus (Methodik,
  Entscheidungen, vollstaendige Aenderungshistorie). **Vor groesseren
  Aenderungen lesen.**

## Arbeitsweise in diesem Projekt

- `src/config.py` ist die EINZIGE Stelle fuer Pfade, Technologie-Liste und
  fachliche Konstanten (Service-Level-Matrix, Finanzparameter, Bucket-Grenzen).
  Neue Konstanten dort ergaenzen, nicht in einzelnen Modulen duplizieren.
- Aenderungen an fachlichen Annahmen (Service-Level, Finanzparameter) sind
  bewusst sichtbar in `config.py` gepflegt – nicht als stille Defaults in
  Berechnungsmodulen. Bei Aenderungen auch `kpis/kpi_definitions.md`
  nachziehen.
- Neue Entscheidungen/Klaerungen mit Stakeholdern in
  `docs/Projektstatus.md` protokollieren (Format: Datum + Abschnitt), nicht
  nur im Commit.
- `input_dir`/`output_dir` NIE hartkodiert in einem Modul referenzieren –
  immer ueber `config.INPUT_ROOT` / `config.OUTPUT_ROOT` bzw. `RunConfig`.

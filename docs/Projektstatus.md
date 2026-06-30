# Projektstatus: Global Manufacturing Transition Strategy

**Erstellt:** 17.06.2026  
**Zuletzt aktualisiert:** 23.06.2026  
**Projektverantwortlich:** Max Karle  
**Status:** 🟡 In Aufbau

---

## 1. Projektziel

Aufbau eines integrierten Dispositions- und Frühwarnsystems zur strategischen Bewertung von Lagerbeständen, Versorgungsrisiken und Bestandsentwicklung. Vollständiger Neuaufbau in Python auf Basis von SAP-Exporten (primär via SE16XXL-Mehrfachtabellen-Merge), mit Versionierung der Projektdokumentation via GitHub.

**Template-Charakter (neu, 23.06.2026):** Das Projekt wird als **technologie-agnostisches Template** aufgebaut, nicht als Einzellösung für eine Produktlinie. Im Unternehmen existieren acht relevante Produkttechnologien (CAT 4, DCS 1.0, DCS 2.0, DCS 0.5, EL.NET, EL.MOTION, EL.INDUSTRY, EL.CAPTURE), die jeweils **einzeln** analysiert werden — kein automatisches Zusammenfassen mehrerer Technologien in einem Lauf. Jeder Analyse-Lauf bezieht sich auf genau eine Technologie; welche das ist, wird zentral über eine Konfigurationsvariable festgelegt, sodass derselbe Code ohne Anpassung für jede Technologie wiederverwendet werden kann.

**Hinweis zur Anforderungsherkunft:** Die ursprüngliche Anforderungsliste ("Initiale Informationen") wurde als Word-Dokument mit Kommentaren eines Kollegen übergeben, nicht von Max selbst verfasst. Mehrere dort als "nicht möglich" markierte Datenpunkte haben sich nach Prüfung durch Max als **machbar** herausgestellt — die ursprüngliche Einschätzung war an mehreren Stellen zu konservativ. Der Datenstatus in Abschnitt 2 spiegelt den korrigierten, von Max bestätigten Stand wider.

**Zeitrahmen:** ca. 2 Wochen (10 Arbeitstage) für ein vollständiges Gesamtpaket, siehe Phasenplan in Abschnitt 5.

---

## 2. Datenverfügbarkeit (Stand: 23.06.2026)

| Datenpunkt | Verfügbarkeit | Quelle | Anmerkung |
|---|---|---|---|
| Bestand (Menge & Wert, inkl. Stichtag) | ✅ Vorhanden | SAP (Transaktion ZMLAG bestätigt) | Snapshot-Export, eine Zeile pro Material, Bestandswert zu 3 Stichtagen |
| Bestandsaufteilung (frei / Q-Prüfung / gesperrt) | ✅ Mergebar | SAP (SE16XXL) | Ursprünglich als "nicht möglich" eingestuft — korrigiert |
| Rahmenverträge / Lieferpläne | ✅ Vorhanden | SAP | inkl. Laufzeit, Restmenge, Abruflogik |
| Offene Bestellungen / bestätigte Liefertermine | ✅ Mergebar | SAP (SE16XXL) | |
| Prognostizierte reale Zugänge (Menge & Datum) | ⚠️ Teilweise möglich | SAP | Nur über bestätigten Liefertermin abbildbar; tatsächliches Eintreffen bleibt Restunschärfe — bewusst akzeptiertes Planungsrisiko |
| Verbrauch (36 Monate, monatlich) | ✅ Vorhanden | SAP (Tabelle MVER) | Direkter Export statt Excel-Zwischenschritt |
| Trendanalyse Verbrauch | 🔄 In Aufbau | – | Wird in Phase 3 auf MVER-Basis entwickelt |
| Reichweite auf Basis Ø-Verbrauch (12 Monate) | 🔄 Neuaufbau in Python | SAP | Vorherige Excel-Logik wird nicht übernommen, sondern neu konzipiert |
| Reichweite mit frei editierbarem Planungsfeld | 🔄 In Aufbau | – | Wird in Phase 2 mitgedacht |
| Reichweite Stock-only & Stock + Zugänge | 🔄 Neuaufbau in Python | SAP | |
| Wiederbeschaffungszeit | 🔄 Methode vorhanden, Automatisierung offen | SAP | Bisher einmalig manuell in Excel berechnet; wird in Phase 4 als wiederholbare Funktion auf Bestellhistorie umgesetzt |
| Nachbestellungshistorie (seit 2024) | ⚠️ Vermutlich verfügbar, ungeprüft | SAP | Max geht davon aus, dass dies in SAP getrackt wird; konkrete Tabelle/Export noch zu identifizieren |
| ABC/XYZ-Klassifizierung | ✅ Vorhanden | SAP (Transaktion ZMLAG bestätigt) | Direkt im ZMLAG-Export enthalten (Spalten ABC-Kennzeichen, ABCXYZ-Kennzeichen); wird für Risikobewertung (Phase 5) genutzt |
| Technologie-/Produktlinienzuordnung je Material | ✅ Vorhanden | SAP (Transaktion ZMLAG bestätigt) | Merkmalfeld "EL_01178 Bezeichnung Merkmalwert"; Basis für die technologie-agnostische Template-Logik |
| Bestandsentwicklung (letzte 2–3 Jahre) | ✅ Vorhanden | SAP | Aus ZMLAG aktuell nur 3 Stichtage ableitbar (kein durchgängiger Monatsverlauf); vollständige 2-3-Jahres-Entwicklung erfordert weiterhin MVER oder einen erweiterten ZMLAG-Export |

**Verbleibende echte Lücken (Stand 23.06.2026):**
- Tatsächliches Eintreffdatum von Zugängen (nur Plan-/Bestätigungstermin verfügbar)
- Exakte SAP-Quelle für Nachbestellhistorie noch nicht identifiziert (aber als grundsätzlich vorhanden eingeschätzt)
- ZMLAG liefert nur 3 Bestandswert-Stichtage (Snapshot), keine durchgängige Monatshistorie — für die volle Bestandsentwicklung über 2-3 Jahre bleibt MVER bzw. ein erweiterter Export erforderlich

---

## 3. Offene Punkte / Lücken

- [x] Bestandsaufteilung nach Bestandsarten aus SAP ermöglichen *(gelöst: SE16XXL-Merge)*
- [ ] Reale Zugänge: Restunschärfe zwischen bestätigtem Termin und tatsächlichem Eintreffen dokumentieren und im Risikomodell berücksichtigen
- [ ] Trendanalyse auf Basis 36-Monats-Verbrauch (MVER) entwickeln
- [ ] Reichweite mit frei editierbarem Planungsfeld umsetzen
- [ ] Wiederbeschaffungszeit als automatisierte Funktion (statt manueller Einmalberechnung) umsetzen
- [ ] SAP-Quelle für Nachbestellungshistorie (ab 2024) identifizieren und Export aufbauen
- [x] GitHub-Repository aufsetzen
- [x] Python-Projekt für Dokumentations-Push zu GitHub strukturieren
- [x] Neue Python-Projektstruktur aufgesetzt *(flache Struktur statt urspr. fachlich gegliedert — siehe Abschnitt 4, Entscheidung Max: `input_data/`, `output_data/`, `src/`)*
- [x] Ersten SAP-Testexport (Struktur + Beispieldaten) bereitgestellt *(ZMLAG-Beispielzeile, siehe Abschnitt 4)*
- [ ] Vollständigen ZMLAG-Export (nicht nur Beispielzeile) einlesen und Validierungslogik gegen reale Datenmenge prüfen
- [ ] Altes GitHub-Sync-Setup (`config.py` mit Token-Platzhaltern, `github_sync.py`) im Projekt-Root entfernen und durch neuen, zur flachen Struktur passenden Sync-Workflow ersetzen *(Entscheidung Max: später)*
- [ ] Status von `docs/`-Ordner im Projekt-Root klären (vermutlich Ablageort dieses Dokuments) und ggf. in neue Struktur einordnen
- [ ] venv-Einrichtung in der finalen Projektumgebung abschließen *(Aktivierung in PowerShell funktioniert, Pakete noch zu installieren — wird in separater Sitzung behandelt)*

---

## 4. Technische Infrastruktur

| Komponente | Status | Anmerkung |
|---|---|---|
| SAP (SE16XXL Mehrfachtabellen-Merge) | ✅ Aktiv | Zentrale Datenquelle für nahezu alle Datenpunkte |
| SAP Tabelle MVER (Verbrauchshistorie) | ✅ Identifiziert | Für 36-Monats-Trendanalyse |
| SAP Transaktion ZMLAG | ✅ Erster Export erfolgreich gezogen | Liefert Materialstamm-Snapshot: Disponent, Materialkurztext, Normbezeichnung, Produkthierarchie, Technologie-Merkmal, Bestandswert zu 3 Stichtagen, ABC/XYZ-Kennzeichen. Eine Zeile pro Material. Dient als Anker-Export (Materialliste) für eine Technologie |
| Excel (historische Reichweitenlogik) | ⚪ Wird nicht übernommen | Dient nur als fachliche Referenz, kein Code-Reuse |
| Altes Python-in-Excel-Whitepaper (Forecast-Driven Inventory Optimization) | 📚 Referenz | Liefert Methodik (ROP, EOQ, Safety Stock, Monte-Carlo-Simulation), aber kein direkter Codeimport — Neuaufbau von Grund auf |
| GitHub Repository | ✅ Aktiv | `github.com/maximilian-karle/Global-Manufacturing-Transition-Strategy` (privat) |
| Python-Projekt (`gmts/`) | ✅ Neue flache Struktur umgesetzt | `input_data/<technologie>/`, `output_data/<technologie>/`, `src/` (config, data_loader, stock_analysis, report_builder, main). Ersetzt die zuvor angedachte fachlich gegliederte Struktur (forecasting/inventory/simulation/...) — Entscheidung Max: Template-Ansatz mit einer Technologie pro Lauf, einfache flache Hierarchie |
| Altes GitHub-Sync-Setup (`config.py`, `github_sync.py` im Root) | 🗑️ Entfernt | Wird später durch neuen, zur flachen Struktur passenden Sync-Workflow ersetzt (Entscheidung Max) |
| Lauffähiges Grundgerüst (Stand 23.06.2026) | ✅ Erfolgreich getestet | ZMLAG-Einlesen inkl. Validierung (fehlende/doppelte Materialnummern, fehlende Stichtagsspalten), Bestandswert-Entwicklung, einfache ABC/XYZ-Risikoeinstufung als Platzhalter, 3-reitriger Excel-Report (Übersicht/Bestand/Risiko). Getestet mit Beispielzeile + 2 Testfällen (inkl. Edge Case Bestandswert=0). Führende Nullen in Materialnummern werden im Report als Text-Zellformat erhalten |

---

## 5. Phasenplan (2 Wochen / 10 Arbeitstage)

| Phase | Tage | Ziel | Status |
|---|---|---|---|
| 1. Datenfundament | 1–2 | SAP-Exporte definieren, einlesen, validieren | 🔄 Läuft — ZMLAG-Beispielzeile erfolgreich verarbeitet, vollständiger Export steht noch aus |
| 2. Bestand & Reichweite | 3–4 | Reichweitenlogik (Ø-Verbrauch, Stock-only/+Zugänge, Planungsfeld) in Python | ⬜ Offen |
| 3. Trend & Verbrauchsanalyse | 5–6 | 36-Monats-Trendanalyse auf MVER-Basis | ⬜ Offen |
| 4. Lead Time & Nachbestellhistorie | 7–8 | Automatisierte Wiederbeschaffungszeit, Nachbestell-Tracking | ⬜ Offen |
| 5. Frühwarnsystem & Risikobewertung | 9 | Verknüpfung zu Risikoklassifizierung, inkl. ABC/XYZ-Gewichtung | ⬜ Offen |
| 6. Gesamtreport & COO-Präsentation | 10 | Konsolidierter Report + Management-Zusammenfassung | ⬜ Offen |

---

## 6. Nächste Schritte

1. Vollständigen ZMLAG-Export (nicht nur Beispielzeile) in `input_data/<technologie>/` ablegen und Einlese-/Validierungslogik gegen die reale Datenmenge testen
2. Prüfen, ob bei vollem Export weiterhin keine doppelten Materialnummern auftreten und alle Stichtagsspalten korrekt erkannt werden
3. Python-Pakete in der neu eingerichteten venv installieren (`pip install -r requirements.txt`) und ersten Lauf (`python main.py`) in der echten VS-Code-Umgebung durchführen
4. Neuen GitHub-Sync-Workflow für die flache Projektstruktur aufsetzen (ersetzt das entfernte alte Setup)
5. SAP-Quelle für Nachbestellhistorie klären
6. Status des `docs/`-Ordners im Projekt-Root klären und einordnen
7. Sobald MVER-Export vorliegt: Trendanalyse-Modul (Phase 3) auf Basis der 36-Monats-Historie beginnen

---

## 7. Änderungshistorie

| Datum | Autor | Änderung |
|---|---|---|
| 17.06.2026 | Claude | Initiale Erstellung auf Basis COO-Dokument |
| 22.06.2026 | Claude | Datenstatus korrigiert (Bestandsaufteilung, Verbrauch, ABC/XYZ als verfügbar bestätigt); Hinweis zur Anforderungsherkunft ergänzt; Phasenplan (2 Wochen) aufgenommen; Klarstellung, dass altes Excel/Whitepaper nur als methodische Referenz dient, kein Code-Reuse; verbleibende echte Lücken (reale Zugänge, Nachbestellhistorie-Quelle) präzisiert |
| 23.06.2026 | Claude | Template-Charakter ergänzt: Projekt deckt alle 8 Produkttechnologien (CAT 4, DCS 1.0/2.0/0.5, EL.NET, EL.MOTION, EL.INDUSTRY, EL.CAPTURE) einzeln ab, ein Lauf = eine Technologie, zentral konfigurierbar; ersten ZMLAG-Export (Beispielzeile) ausgewertet und als Materialstamm-Anker-Quelle bestätigt (Bestandswert zu 3 Stichtagen, ABC/XYZ, Technologie-Merkmal); neue flache Projektstruktur (`input_data/`, `output_data/`, `src/`) anstelle der zuvor angedachten fachlich gegliederten Struktur umgesetzt und als lauffähiges Grundgerüst getestet (Einlesen, Bestandsentwicklung, Risikoeinstufung, 3-reitriger Excel-Report); altes GitHub-Sync-Setup (`config.py`, `github_sync.py`) im Projekt-Root entfernt, Neuaufsetzung verschoben; offene Punkte zu `docs/`-Ordner und venv-Einrichtung dokumentiert |

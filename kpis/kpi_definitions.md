# KPI-Definitionen

Diese Datei fasst die fachlichen Kennzahlen-Definitionen und Schwellenwerte
zusammen, die im Code als "sichtbare Annahmen" (keine stillen Defaults) in
`src/config.py` gepflegt werden. Aenderungen an den Werten erfolgen dort –
diese Datei ist die lesbare Referenz dazu, keine zweite Quelle der Wahrheit.

## Service-Level-Matrix (Safety Stock/ROP)

Ziel-Servicegrad je ABC/XYZ-Klasse (Wahrscheinlichkeit, waehrend der
Wiederbeschaffungszeit nicht in eine Fehlmenge zu laufen). Leitet den
z-Wert fuer die Safety-Stock-Formel ab (SS = z * sigma * sqrt(L)).

| ABC \ XYZ | X | Y | Z |
|---|---|---|---|
| A | 99% | 98% | 97% |
| B | 98% | 96% | 94% |
| C | 96% | 93% | 90% |

Herkunft A-Zeile: Referenz-Template von Max. B/C-Zeile: plausible,
monoton fallende Fortsetzung (Entscheidung Claude, 29.06.2026, von Max
zur Pruefung freigegeben).

## Reichweite-vs-Wiederbeschaffung-Bucket-Grenzen

Bucket-Grenzen (Monate) fuer die Kreuztabelle im HTML-Dashboard:
`[0-3]`, `[3-6]`, `[6-9]`, `[9-15]`, `[15-20]`, `[20-30]`, `[30-999]`.

## Finanzparameter (Working Capital/ROI)

| Parameter | Wert | Bedeutung |
|---|---|---|
| `HOLDING_COST_RATE` | 20% p.a. | Lagerhaltungskostensatz auf den Bestandswert |
| `ORDER_COST` | 75 EUR | Losfixkosten je Bestellvorgang (noch nicht in working_capital.py verwendet) |
| `EMERGENCY_FREIGHT_SURCHARGE_RATE` | 30% | Relativer Aufschlag bei Fehlmengen-Notfallbeschaffung (noch nicht verwendet) |

Herkunft: alle drei Werte direkt von Max vorgegeben (Klaerung 30.06.2026).
ROI-Definition: `Annual_Savings / Working_Capital_Aktuell` (keine
Berechnung gegen Implementierungskosten).

## Sonstige Konstanten

- `DAYS_PER_MONTH = 30.0` – Naeherungsfaktor Tage↔Monate fuer die
  Wiederbeschaffungszeit.
- `MATERIAL_NUMBER_LENGTH = 18` – einheitliches Zielformat fuer
  Materialnummern (SAP-native 18-stellige Form) beim Join von ZMLAG,
  MVER und SE16XXL.

## Quelle

Vollstaendige, kommentierte Definitionen: `src/config.py`. Hintergrund und
Entscheidungshistorie: `docs/Projektstatus.md`.

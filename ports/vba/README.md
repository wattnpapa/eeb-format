# VBA-Dekoder für Excel-Arbeitsmappen

Drei Module, mit denen eine Excel-Mappe EEB-Codes über einen USB-Handscanner
(Tastatur-Emulation) oder einen eingefügten Link einliest. Reines VBA ohne
Systembibliotheken: Base41/Base64url, Deflate (RFC 1951), Binärstrom.

| Modul | Aufgabe |
| --- | --- |
| `m_digitalerEEBParsing.bas` | Dekoder: Scan-Text → `TEebBogen`. Alle Schemata 2 bis `SCHEMA_MAX`, mehrteilige Codes (`EEBS.`-Serien mit FNV-1a-Prüfsumme), Vorlagen-Marker `V.`. Gegenstück zu `decodeBinaer`/`parseSegmentUrl`/`segmentePayload` in `src/codec.ts`. |
| `m_digitalerEEBVokabular.bas` | Klartexte zu den Zahlencodes (THW-Einheitstypen, -Funktionen, -Fahrzeugtypen, Funkruf-Kennwörter, Hierarchie-Ebenen). **Generiert** aus `bos-vokabulare` — nicht von Hand ändern, sondern im Repo erfassungsbogen `npm run vokabular:excel-vba` laufen lassen (dort liegen Format und Vokabular zusammen als Submodule). |
| `m_digitalerEEB.bas` | Beispiel einer Mappenanbindung (Formular `ufr_digitalerEEBEingabe`, Blatt `t_staerke`, feste Spaltenbelegung). An die eigene Mappe anpassen; zeigt den Ablauf für Einzel- und mehrteilige Codes. |

Die Dateien sind in Windows-1252 kodiert (der VBA-Import erwartet ANSI) und
werden im VBA-Editor über *Datei → Datei importieren* eingebunden.

## Was der Dekoder kann und was nicht

- Liest Schema 2 bis 10; Felder jüngerer Versionen bleiben bei älteren Codes
  leer/0 (Stand minutengenau ab 7, weitere Fahrerlaubnisklassen ab 8,
  Sitzplätze ab 9, abweichende Dienststelle je Person ab 10).
- Berufe stehen im Format als Freitext mit Präfix `Beruf: ` in den
  Zusatzqualifikationen; der Dekoder trennt sie nach `TEebPerson.berufe`.
- Mehrteilige Codes: `EebIstSegmentCode` vor `EebBogenLesen` prüfen, Teile mit
  `EebSegmentSammeln` aufnehmen, bei Vollständigkeit `EebSegmenteZuBogen`,
  nach dem Schreiben `EebSammlungVerwerfen`.
- Signierte Container (`EEB2C`) werden gelesen, die Signatur wird
  **übersprungen und nicht geprüft** (`signaturstufen` > 0 zeigt es an).
- THW-Codes mit `standortRef` (OV-Nummer statt ausgeschriebener Hierarchie)
  werden als „OV-Nr. 123" ausgewiesen; das OV-Verzeichnis ist nicht enthalten.
- Unbekannte Vokabular-Codes erscheinen als `#123`; Vokabulare anderer
  Organisationen als THW sind nicht hinterlegt.

## Prüfen gegen die Testvektoren

Es gibt kein Headless-VBA, deshalb von Hand: aus
`../testvektoren/testvektoren.json` die `url` jedes Vektors (und bei
`voll-schema-10` die drei `segmente.urls` in beliebiger Reihenfolge) in das
Eingabefeld der Mappe geben bzw. `EebBogenAlsText(url)` im Direktfenster
aufrufen und mit `erwartet` vergleichen. Mindestens prüfen:

- kein Fehler „nicht unterstützte Schema-Version",
- `voll-schema-10`: `personal(1).dienststelle = "OV Delmenhorst"`,
  `personal(1).berufe = "Elektroniker/in, Krankenpfleger/in"`,
  `personal(1).fahrerlaubnis = "B+CE+A"`, `fahrzeuge(1).sitzplaetze = 7`,
  `fahrzeuge(2).typ = "Mannschaftstransportwagen Ortsverband"`,
- `alt-schema-2/8/9`: Personen, Funktionen und Fahrzeuge wie in `erwartet`.

Nach jeder Formatänderung ist dieser Ordner im selben Commit nachzuziehen
(siehe `../../CLAUDE.md`); die CI hält `SCHEMA_MAX` gegen `SCHEMA_VERSION`.

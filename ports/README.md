# Portierungen des EEB-Formats

Die Referenz des Formats ist der TypeScript-Code unter `src/`. Hier liegen
Leser in anderen Sprachen, die dieselben QR-Codes und Links dekodieren:

| Ordner | Sprache | Umfang | Tests |
| --- | --- | --- | --- |
| `vba/` | VBA (Excel) | Lesen aller Schemata ab 2, Einzel- und mehrteilige Codes, Signatur-Container wird übersprungen | manuell gegen die Testvektoren (kein Headless-VBA) |

Jede Portierung hat ein `PORT.json` (Schema-Stand, Fähigkeiten, Verweis auf
die Versionskonstante im Quelltext) und ein eigenes `README.md`.

## Testvektoren

`testvektoren/testvektoren.json` wird mit `npm run ports:testvektoren` aus der
Referenz erzeugt und ist eingecheckt. Je Vektor:

- `schemaImStrom` — die Version, die im Bytestrom steht (kann kleiner sein als
  die aktuelle, weil der Encoder die kleinste tragende Version schreibt).
- `binaerHex` — der unkomprimierte Binärstrom (nur bei neu erzeugten Vektoren).
- `payloadHex` — der komprimierte Payload mit `EEB2`-Magic, wie er im QR steht.
- `url` — der Link, wie ihn ein Scanner liefert (Base41, Marker `B.`),
  `urlBase64` — derselbe Payload Base64url-kodiert (ältere Codes).
- `segmente` — bei mehrteiligen Vektoren die `EEBS.`-URLs und die FNV-1a-`id`.
- `erwartet` — der Bogen nach Migration (`schemaVersion` = aktuelle Version),
  Zeiten als Tage bzw. Minuten seit 2020-01-01.
- `alt-schema-*` — eingefrorene Codes aus `src/codec.migration.test.ts`; sie
  belegen, dass ältere Lesepositionen weiter stimmen.

`npm run ports:pruefen` (CI) prüft, dass die Datei dem Codec entspricht und
jede Portierung die aktuelle `SCHEMA_VERSION` nennt. Was bei einer
Formatänderung zu tun ist, steht in `../CLAUDE.md`.

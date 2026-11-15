# Hinweise für Coding-Agenten (Claude Code u. a.)

Dieses Repository ist die **Referenz des EEB-Transportformats** (Modell, Codec,
Segmentierung, Signaturkette). Neben der TypeScript-Referenz liegen unter
`ports/` **Portierungen in andere Sprachen**, die dasselbe Format lesen. Sie
werden nicht aus dem TypeScript generiert, sondern von Hand gepflegt — und
veralten deshalb stillschweigend, wenn niemand sie nachzieht.

## Regel: Formatänderung = alle Portierungen im selben Commit nachziehen

Berührt eine Änderung eine der folgenden Stellen, **muss jede vorhandene
Portierung unter `ports/<sprache>/` im selben Arbeitsschritt angepasst
werden** — nicht als Folge-Ticket, nicht „später":

| Auslöser | Was in den Portierungen nachzuziehen ist |
| --- | --- |
| `SCHEMA_VERSION` in `src/model.ts` steigt, neues oder entfallenes Feld, geänderte Leseposition im Binärstrom (`encodeBinaer`/`decodeBinaer` in `src/codec.ts`) | Lesereihenfolge der Portierung ergänzen (versionsabhängig, ältere Schemata bleiben lesbar), Datenstruktur erweitern, Schema-Konstante heben (VBA: `SCHEMA_MAX`), `PORT.json` → `schemaVersion` heben |
| Flag-Bits (Personal-, Fahrzeug-, Kontakt-, Hierarchie-, Sofortbedarf-Flags) | dieselben Bits in der Portierung |
| Transportformat: Base41-Alphabet, Marker (`B.`, `V.`, `EEBS.`), Segment-Kopf, Prüfsumme (FNV-1a), `MAX_STUFEN`, `EEB2`/`EEB2C`-Magic, Deflate | entsprechende Stelle der Portierung |
| Enumerationen (`OrganisationsTyp`, `Fahrerlaubnis`, `StaerkeRolle`, `Geschlecht`, `Ernaehrung`, `KontaktArt`) | Klartext-Tabellen der Portierung |
| Neue Konvention im Freitext (z. B. Präfix `Beruf: ` in den Zusatzqualifikationen) | Portierung trennt/zeigt sie gleich |
| Vokabular-Codes (liegen in `bos-vokabulare`, nicht hier) | generierte Vokabular-Module der Portierungen neu erzeugen (VBA: `npm run vokabular:excel-vba` im Repo erfassungsbogen) |

Arbeitsschritte bei jeder solchen Änderung:

1. Referenz (TypeScript) ändern, Tests ergänzen — insbesondere eine
   eingefrorene Golden-Fixture des **alten** Schemas in
   `src/codec.migration.test.ts`, wenn die Leseposition sich verschiebt.
2. `npm run ports:testvektoren` — erzeugt `ports/testvektoren/testvektoren.json`
   neu. Dort ggf. den Vektor `voll-schema-N` in `scripts/testvektoren.mts`
   um das neue Feld erweitern, damit es überhaupt in den Vektoren vorkommt.
3. Jede Portierung unter `ports/` anpassen (siehe deren `README.md` und
   `PORT.json`); Schema-Konstante und `PORT.json` heben; in der Portierung
   den Hinweis „ab Schema N" an der Lesestelle lassen.
4. `npm run ports:pruefen` muss grün sein. Der Wächter läuft auch in der CI
   und fällt, wenn eine Portierung oder die Testvektoren hinter
   `SCHEMA_VERSION` zurückbleiben.
5. Portierungen ohne automatisch ausführbare Tests (VBA) gegen die
   Testvektoren von Hand prüfen und das in der Commit-Nachricht sagen; was
   nicht geprüft werden konnte, ebenfalls sagen.

Ist das Nachziehen einer Portierung größer als die eigentliche Änderung,
trotzdem im selben Commit erledigen. Geht es wirklich nicht, die Portierung
**sichtbar** als veraltet markieren: `PORT.json` → `"veraltet": "<Grund>"`,
Hinweis im Kopf des Quelltexts, und der Wächter bleibt rot, bis es erledigt
ist. Lieber ein roter Prüfschritt als ein Leser, der ab Schema N falsch liest.

## Regel: Schemaänderung = alle drei Apps im selben Arbeitsschritt nachziehen

Dieses Format lesen drei Apps, die den Kern als gepinnten Stand einbinden:

- **Erfassungsbogen** (<https://github.com/wattnpapa/erfassungsbogen>)
- **Meldekopf-App** (<https://github.com/wattnpapa/meldekopf>)
- **S1-Control** (<https://github.com/wattnpapa/S1-Control>)

Bleibt eine davon hinter dem Schema zurück, lehnt sie neue Bögen ab oder zeigt
sie falsch an, und zwar am Meldekopf, also dort, wo der Bogen gebraucht wird.

Deshalb gilt bei **jeder Anpassung des Schemas** (`SCHEMA_VERSION`, neues oder
entfallenes Feld, Flag-Bit, Enumeration, neue Freitext-Konvention; auch ohne
höhere Schema-Version wie bei `einheit.einzelperson`):

- Die Änderung ist erst fertig, wenn **Erfassungsbogen, Meldekopf und
  S1-Control** den neuen Stand dieses Repos übernommen haben und weiter
  funktionieren: Submodul- bzw. Paketstand heben, dort Typprüfung, Tests und
  Verhaltensszenarien laufen lassen, und alles anpassen, was das neue Feld
  lesen, anzeigen, exportieren oder weiterreichen muss.
- Das geschieht im selben Arbeitsschritt, nicht als Folgeaufgabe. Reihenfolge:
  dieses Repo und die übrigen Kern-Repos (`bos-meldekopf`, `bos-eeb-satz`,
  `bos-taktische-zeichen`, `bos-vokabulare`), dann Erfassungsbogen, dann
  Meldekopf und S1-Control.
- Fehlt der Zugriff auf eines der Repos, das **vor** der Schemaänderung sagen
  und den Zugriff anfordern. Geht es trotzdem nicht, am Ende ausdrücklich
  melden, welche App noch auf dem alten Stand ist. Nie stillschweigend nur das
  Format ändern.
- Alte Bögen bleiben in allen drei Apps lesbar (Aufnahmeregel 4, ab Schema 2).

Dieselbe Regel steht in der `CLAUDE.md` des Erfassungsbogens.

## Neue Portierung aufnehmen

- Ordner `ports/<sprache>/` mit `README.md`, `PORT.json` (Felder wie in
  `ports/vba/PORT.json`; `schemaKonstante` zeigt auf die Versionskonstante im
  Quelltext, damit das Manifest nicht ohne den Code gehoben werden kann).
- Die Portierung liest **alle** Schemata ab 2 (Aufnahmeregel 4), auch
  mehrteilige Codes (`EEBS.`) und den Signatur-Container (`EEB2C`, mindestens
  überspringen).
- Sie prüft sich gegen `ports/testvektoren/testvektoren.json`; hat die Sprache
  einen Testläufer, der in GitHub Actions läuft, den Lauf in `ci.yml`
  eintragen.
- Keine Portierung gehört in das npm-Paket (`files` in `package.json` bleibt
  bei `dist` und `README.md`).

## Übrige Regeln

Die Aufnahmeregeln ADR-003 (plattformneutral, additiv, abwärtskompatibel ab
Schema 2, gepinnte Submodul-Commits) stehen in `README.md` und sind per
ESLint, zwei tsconfigs und doppeltem Testlauf (node + jsdom) hinterlegt.

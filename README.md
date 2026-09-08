# @bos/eeb-format

Geteilter, plattformneutraler TypeScript-Kern für BOS-Anwendungen.

Dieses Repository ist **kein eigenständiges Produkt**, sondern hängt als
git-Submodul unter `vendor/eeb-format` in zwei Produkten:

- **einheitenerfassungsbogen** — PWA mit Capacitor und Electron, gebaut mit Vite
- **S1-Control v2** — Electron-Anwendung mit npm-Workspaces

Beide binden das Paket über `"@bos/eeb-format": "file:vendor/eeb-format"` ein.

Der derzeitige Inhalt ist **absichtlich fast leer**: `kernVersion()` und
`inhaltsHash()` sind der Verdrahtungsnachweis, mit dem Einbindung, Bau, Typen,
Lint und Testlauf in beiden Produkten belegt werden. Die eigentliche Extraktion
fachlicher Bausteine folgt in einem späteren Arbeitspaket.

## Aufnahmeregeln (ADR-003)

Was in diesen Kern darf, entscheiden diese sechs Regeln:

1. Aufnahme nur, wenn beide Produkte den Baustein aufrufen.
2. Keine `node:`-, DOM- oder React-Importe; geprüft per ESLint und durch Testlauf unter `node` und `jsdom`.
3. Keine Rückimporte aus `@s1/*` oder aus der Erfassungsbogen-App.
4. Änderungen additiv; Schema-Abwärtskompatibilität bleibt Pflicht (QR-Codes und Dateien ab Schema 2 lesbar).
5. Bundle-Budget im CI von erfassungsbogen.app; der Kern darf die PWA nicht schwerer machen.
6. Gepinnte Submodul-Commits; kein automatisches Folgen von `main`.

Regel 2 und 3 sind maschinell hinterlegt:

- `eslint.config.mjs` verbietet per `no-restricted-imports` alle `node:`-Module
  und deren bloße Namen (`fs`, `path`, `crypto`, …), Electron, Capacitor, React,
  DOM-nahe Pakete sowie Rückimporte aus `@s1/*` und dem Erfassungsbogen.
- `tsconfig.json` lässt `"DOM"` bewusst aus `lib` weg und setzt `"types": []`.
  Damit scheitert schon die Typprüfung an `document`, `window`, `process` oder
  `Buffer` — auch dann, wenn kein Import im Spiel ist.
- `vitest.config.ts` fährt dieselben Testdateien in zwei Projekten: einmal unter
  `node`, einmal unter `jsdom`.

## Rückweg

Blockiert oder verzögert der geteilte Kern zweimal in drei Monaten ein Release
des Schwesterprodukts, wird das Vendoring eingefroren: Der Stand wird als Kopie
nach `packages/kern-vendor/` übernommen, der Herkunfts-Commit dort festgehalten,
danach werden beide Seiten getrennt gepflegt. Der Rückweg ist ausdrücklich
vorgesehen und kein Scheitern — er kostet weniger als ein Kern, der beide
Produkte ausbremst.

## Lokal bauen und prüfen

Voraussetzung: Node 24 (siehe `.nvmrc`). Die `.npmrc` zeigt bewusst auf die
öffentliche npm-Registry, weil global eine CodeArtifact-Registry konfiguriert
ist.

```bash
npm install      # installiert und baut über das prepare-Skript nach dist/
npm run build    # dist/index.js plus dist/index.d.ts
npm run typecheck
npm run lint
npm test         # ein Lauf, zwei Umgebungen: node und jsdom
```

### Zwei TypeScript-Versionen nebeneinander

Gebaut und typgeprüft wird mit TypeScript 7 (`@typescript/native`, liefert das
`tsc` auf dem Pfad). TypeScript 7 bringt jedoch keine JavaScript-Compiler-API
mehr mit, und `typescript-eslint` bricht damit beim Laden ab. Deshalb ist
zusätzlich die von Microsoft dafür vorgesehene Kompatibilitätsschiene
installiert: `"typescript": "npm:@typescript/typescript6@^6.0.2"` stellt die
alte API bereit, die ESLint braucht. Sobald `typescript-eslint` TypeScript 7
unterstützt, fällt der zweite Eintrag ersatzlos weg.

## Wie das Paket konsumiert wird — und warum

Gewählt ist **Weg (b): `exports` zeigt auf das gebaute `dist/`**, gebaut per
`tsc`, ausgelöst über das `prepare`-Skript beim `npm install` des Konsumenten.

Die Alternative — `exports` direkt auf `./src/index.ts` — wurde ausprobiert und
verworfen. Sie sieht zunächst bequem aus und funktioniert sogar, solange der
`file:`-Pfad als Symlink im `node_modules` des Konsumenten liegt: Node löst dann
den echten Pfad außerhalb von `node_modules` auf und darf die Typen strippen.
Sobald das Paket als echtes Verzeichnis unter `node_modules` landet — nach
`npm pack`, in einem CI-Cache oder beim Packen einer Electron-Anwendung —
scheitert derselbe Import hart:

```
Error [ERR_UNSUPPORTED_NODE_MODULES_TYPE_STRIPPING]
```

Node strippt Typen grundsätzlich nicht für Abhängigkeiten unterhalb von
`node_modules`. Damit wäre der Kern von einer Eigenschaft der Installationsform
abhängig — genau die Sorte Fehler, die erst im Release auffällt. Weg (b) liefert
stattdessen gewöhnliches ESM plus Deklarationsdateien: Vite und esbuild können
vorbündeln, `tsc -b` in den Workspaces von S1-Control liest `dist/index.d.ts`
wie bei jedem anderen Paket, und Electron lädt zur Laufzeit reines JavaScript.

Beide Wege wurden mit `tsc -b` und mit einem Node-Skript gegen eine
`file:`-Installation geprüft; nur Weg (b) hält in allen Fällen.

**Zu beachten beim Konsumenten:** Neuere npm-Versionen führen Install-Skripte
von Abhängigkeiten nicht mehr ungefragt aus. Bleibt `dist/` nach dem
`npm install` leer, ist das `prepare`-Skript blockiert worden; dann entweder
`npm install-scripts approve @bos/eeb-format` oder schlicht

```bash
npm --prefix vendor/eeb-format install && npm --prefix vendor/eeb-format run build
```

## Lizenz

EUPL-1.2

/**
 * Testvektoren für die Portierungen des Formats unter ports/.
 *
 * Erzeugt aus dem TypeScript-Codec (der Referenz) einen Satz Bögen in allen
 * Transportformen — Einzelcode (Base41 und Base64url), mehrteilige Serie
 * (EEBS.), eingefrorene Alt-Schemata — und schreibt sie zusammen mit dem
 * erwarteten Dekodierergebnis nach ports/testvektoren/testvektoren.json.
 * Jede Portierung (VBA, später Python, C#, …) prüft sich gegen diese Datei;
 * siehe ports/README.md.
 *
 *   npm run ports:testvektoren   neu erzeugen (baut vorher dist/)
 *   npm run ports:pruefen        CI-Wächter, ohne zu schreiben:
 *                                 1. die eingecheckte Datei entspricht dem Codec
 *                                 2. jede Portierung (ports/<name>/PORT.json)
 *                                    nennt die aktuelle SCHEMA_VERSION
 *                                 3. VBA: SCHEMA_MAX im Modul stimmt überein
 *
 * Fällt der Wächter, hat sich das Format geändert, ohne dass die Portierungen
 * nachgezogen wurden — oder umgekehrt. Beides ist ein Fehler im selben Commit,
 * nie ein Folge-Ticket (CLAUDE.md, Abschnitt „ports/").
 *
 * Liest bewusst aus dist/ (nicht aus src/): Node führt .mts ohne Typen aus,
 * löst aber die ./model.js-Spezifizierer der Quelle nicht auf .ts um.
 */

import { readFileSync, writeFileSync, readdirSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const wurzel = join(dirname(fileURLToPath(import.meta.url)), "..");
const codec = await import(join(wurzel, "dist/codec.js"));
const model = await import(join(wurzel, "dist/model.js"));
const { nodeKompressor } = await import(join(wurzel, "dist/qr-node.js"));

const SCHEMA_VERSION: number = model.SCHEMA_VERSION;
const zieldatei = join(wurzel, "ports/testvektoren/testvektoren.json");

// ------------------------------------------------------------------ Bögen

const tag = (iso: string): number => model.datumAusIso(iso);
const minute = (iso: string, hh: number, mm: number): number => tag(iso) * model.MINUTEN_JE_TAG + hh * 60 + mm;

/** Kleinstmöglicher Bogen: nur Pflichtfelder, keine Person, kein Fahrzeug. */
const minimal = {
  schemaVersion: SCHEMA_VERSION,
  stand: tag("2026-10-27") * model.MINUTEN_JE_TAG,
  einheit: { organisation: 2, organisationName: "Freiwillige Feuerwehr Wardenburg", einheitsTyp: { freitext: "Löschzug" }, hierarchie: [{ bezeichnung: { code: 1 }, name: "Wardenburg" }] },
  einsatz: { zeitraumVon: tag("2026-10-27"), zeitraumBis: tag("2026-10-27"), ortAuftrag: "Bereitstellung" },
  personalErfassung: 1,
  staerkeManuell: { fuehrer: 1, unterfuehrer: 2, mannschaft: 6, gesamt: 9 },
  personal: [],
  fahrzeuge: [],
};

/** Voller Bogen: nutzt jedes Feld bis einschließlich Schema 10. */
const voll = {
  schemaVersion: SCHEMA_VERSION,
  stand: minute("2026-10-27", 8, 30),
  uebung: true,
  einheit: {
    organisation: 1,
    einheitsTyp: { code: 4 },
    hierarchie: [
      { bezeichnung: { code: 1 }, name: "Oldenburg", telefon: "04411234567", email: "ov-oldenburg@thw.de", kurz: "OL" },
      { bezeichnung: { code: 2 }, name: "Oldenburg" },
    ],
  },
  einsatz: {
    zeitraumVon: tag("2026-10-27"),
    zeitraumBis: tag("2026-10-29"),
    ortAuftrag: "Deichverteidigung Elsfleth",
    einsatzbeginn: minute("2026-10-27", 10, 0),
    einsatzende: minute("2026-10-29", 18, 0),
    zeitraumVonIso: "2026-10-27T08:00:00+01:00",
    zeitraumBisIso: "2026-10-29T18:00:00+01:00",
  },
  personalErfassung: 0,
  unterbringungManuell: { m: 1, w: 1, d: 0 },
  verpflegungManuell: { vegetarisch: 1, vegan: 0 },
  personal: [
    {
      vorname: "Anna", nachname: "Müller", fahrerlaubnis: 5, geschlecht: 1, staerkeRolle: 2, ernaehrung: 1,
      weitereFahrerlaubnisse: [10, 4], dienststelle: "OV Delmenhorst",
      funktionen: [{ code: 3 }, { freitext: "Deichfachberaterin" }],
      kontakte: [
        { art: 0, dienstlich: true, wert: "01701234567" },
        { art: 2, dienstlich: false, wert: "anna.mueller@example.org" },
      ],
      zusatzqualifikationen: [{ code: 32 }, { freitext: "Beruf: Elektroniker/in" }, { freitext: "Beruf: Krankenpfleger/in" }],
    },
    {
      vorname: "Ben", nachname: "Schulz", fahrerlaubnis: 9, geschlecht: 0, staerkeRolle: 0, ernaehrung: 0,
      funktionen: [{ code: 5 }], kontakte: [{ art: 1, dienstlich: false, wert: "04419876" }],
      zusatzqualifikationen: [{ freitext: "Rettungsschwimmer" }],
    },
  ],
  fahrzeuge: [
    {
      typ: { code: 4 }, kennzeichen: "THW-81234", stanKonform: true, sitzplaetze: 7, aenderungen: "Anhängerkupplung nachgerüstet",
      funkrufname: { kennwort: { code: 1 }, eigenerStandort: true, teile: [22, 51] },
    },
    { typ: { code: 27 }, kennzeichen: "THW-81299", stanKonform: false, funkrufname: { kennwort: { code: 1 }, eigenerStandort: false, ort: "Delmenhorst", teile: [86, 25] } },
    { typ: { freitext: "Privat-Pkw" } },
  ],
  sofortbedarf: { verpflegungPersonen: 2, dieselLiter: 120, benzinLiter: 0, gemischLiter: 5, unterbringung: true, ruhezeitErforderlich: true },
  sonstiges: "Testbogen Schema 11 für die Portierungen",
};

/** Schema 5: nichts, was eine höhere Version fordert — so schreibt der Encoder noch heute Altcodes. */
const schema5 = {
  ...minimal,
  einheit: { organisation: 5, einheitsTyp: { freitext: "SEG Sanität" }, hierarchie: [{ bezeichnung: { code: 1 }, name: "Oldenburg" }, { bezeichnung: { code: 2 }, name: "Oldenburg-Land" }] },
  personalErfassung: 0,
  staerkeManuell: undefined,
  personal: [
    { vorname: "Cem", nachname: "Yilmaz", fahrerlaubnis: 5, geschlecht: 0, staerkeRolle: 1, ernaehrung: 2, funktionen: [{ freitext: "Gruppenführer" }], kontakte: [], zusatzqualifikationen: [] },
  ],
  fahrzeuge: [{ typ: { freitext: "KTW" }, kennzeichen: "OL-RK 123" }],
};

// ---------------------------------------- Eingefrorene Alt-Codes (Golden)
// Dieselben Fixtures wie in src/codec.migration.test.ts — dort ist belegt,
// dass der Codec sie liest. Hier bekommen die Portierungen sie als Vektor.

const ALTCODES: Array<{ name: string; schema: number; url: string; hinweis: string }> = [
  {
    name: "alt-schema-2",
    schema: 2,
    hinweis: "Person ohne Ernährungs-Byte, Sofortbedarf mit davonVegetarisch-Byte, Base64url ohne Marker.",
    url: "https://erfassungsbogen.app/#RUVCMmNawc7IoC3oH6bgn5OSmpdUWpSuoKvgl8nIysiHKsLN4mjCyMrP4u_v4sqwgn01u6xbalFebmpOSmpSYunhPUmleekK3olJqTlJOZklVQzMHF75GYl5eanF7EGlKfk5BRmdjOxMDNyMBUImAfJcjAwsjnl5iazhqUmpRQIMDAzMIfm57G6ZxckZqUWsDAwMjMKMayezMoowM4AAAA",
  },
  {
    name: "alt-schema-8",
    schema: 8,
    hinweis: "Weitere Fahrerlaubnisklassen vorhanden, noch ohne Sitzplatzzahl (Flag 32) und Dienststelle.",
    url: "https://erfassungsbogen.app/#B.LMAD3ATPYEAN0RTSMUSDVJNUHOUFCOFDM$1NNWDEOQ77QU4NBCC19D-*5FDM5GNHHC0VDFDM$FNPENAMN-JVBOU9WM37ML3N-XDFP45MNXAO92MU4ULNURNULETZ/TF3VMHY2NQ2YE*KG5KVRJUZZTIG3VWDJID/VM6MN38NAMN7NQJDI:-BN7G9O25M0T:GG6IK4N5$LS/DA$JXTJ1550JU9RA7/0J*0",
  },
  {
    name: "alt-schema-9",
    schema: 9,
    hinweis: "Sitzplatzzahl am Fahrzeug, noch ohne Dienststelle je Person — hinter dem Fahrerlaubnis-Zähler folgt direkt die Funktionsliste.",
    url: "https://erfassungsbogen.app/#B.LMAD3A/RYY-HAPM/4M/RNP6MH8MDP9DR5OA7$HRB8E-ED2A7K5U83YTUM-C5SL59BJ1L-2JLD:TCMYH751-L-WUG-0S92$I-G/T1VS*2F6KOASM9YNBEL6-L:0$4Y6BDL13A6C0PEL8W6JKIP/K0FY-8BWOAIDL9Y1E*1HQ1ZVNXS4*DS4*58:M:0$APB9Y1YNB00",
  },
  {
    name: "alt-schema-10",
    schema: 10,
    hinweis: "Einsatzbeginn und -ende gesetzt, noch ohne Uhrzeit zum Zeitraum (Einsatz-Flags 4/8 ab Schema 11). Ehemals Vektor voll-schema-10.",
    url: "https://erfassungsbogen.app/#B.LMAD3ABXB80Y764T63JQVOYG$UGQCK5K2TW1UPG0X2AWMDXSJTO8$FLYJLDXQUUM2$C*QXVF1FRE-KGPY:73230UBLT-XAAR09YZNMMNAHJJ8Y4KTQJ6I100YN76VR6UA/HKZTGYWHTFYQLT73ZUX8TSM-/ZF-WAHJVKHW$00ZOX6-OLX:-G0-VA2I*F4XCVK4G9PPSLELQ41I9NY7GJ$N8174DSI27OY26MAGJO/6L/AK35O06TM70QPQ6*TAMJ7SFS7LPLAVRPKPZ4L3N06E5T6IIPRAO5908PA$H326V1OU$AO3PFU/PUCOCSKI4NXVY47HNF89-5BU-AQXHH0EA:*XRWPF52ENDPHMFAP3B927JKXHWZR10PKB9DZ7S$55ED:UZ*1335NYW7BE3TOLSL-SY7:$20SCGY4TKMJ09ZY:W-$2MQQZR-3:TS*K6P9JN-8H:0ERR0M0G1I1A-GBH-S07F-H0ZK484XF3ZG2BA8WNKZ91P6EQAXRXNQWOTDQM5N5M5X67BX1VQLLW-MW*LOL0NPD4K8AOKZ",
  },
];

// ------------------------------------------------------------- Erzeugung

function normalisiert(b: unknown): unknown {
  return JSON.parse(JSON.stringify(b));
}

function vektorAusBogen(name: string, bogen: Record<string, unknown>, segmente: number, hinweis: string) {
  const binaer: Uint8Array = codec.encodeBinaer(bogen);
  const payload: Uint8Array = codec.encodePayload(bogen, nodeKompressor);
  const url: string = codec.encodePayloadUrl(bogen, nodeKompressor);
  const erwartet = codec.decodePayloadUrl(url, nodeKompressor);
  const vektor: Record<string, unknown> = {
    name,
    hinweis,
    schemaImStrom: binaer[0],
    binaerHex: Buffer.from(binaer).toString("hex"),
    payloadHex: Buffer.from(payload).toString("hex"),
    url,
    urlBase64: codec.EEB_URL_PREFIX + Buffer.from(payload).toString("base64url"),
    erwartet: normalisiert(erwartet),
  };
  if (segmente >= 2) {
    const urls: string[] = codec.segmentPayloadUrls(payload, segmente);
    // Gegenprobe des Referenz-Codecs selbst, in verdrehter Reihenfolge.
    const teile = [...urls].reverse().map((u: string) => codec.parseSegmentUrl(u));
    const zurueck = codec.segmenteZuBogen(teile, nodeKompressor);
    if (JSON.stringify(normalisiert(zurueck)) !== JSON.stringify(normalisiert(erwartet))) {
      throw new Error(`Segmentierung von ${name} liefert nicht den Ausgangsbogen`);
    }
    vektor.segmente = { anzahl: segmente, id: teile[0].id, urls };
  }
  return vektor;
}

function erzeugen(): string {
  const vektoren = [
    vektorAusBogen("minimal", minimal, 0, "Nur Pflichtfelder, manuelle Stärke, Feuerwehr mit Freitext-Einheitstyp."),
    vektorAusBogen("schema-5", schema5, 0, "Inhalt fordert nur Schema 5 — der heutige Encoder schreibt dann auch Schema 5 (transportSchemaVersion)."),
    vektorAusBogen("voll-schema-11", voll, 3, "Jedes Feld bis Schema 11 belegt: Zeitraum mit Uhrzeit und Offset (ISO 8601, 08:00–18:00 +01:00), Dienststelle, Berufe-Präfix, weitere Fahrerlaubnisse, Sitzplätze, Funkrufnamen, Sofortbedarf. Zusätzlich als dreiteilige EEBS.-Serie."),
    ...ALTCODES.map((a) => {
      const erwartet = codec.decodePayloadUrl(a.url, nodeKompressor);
      const payload: Uint8Array = codec.datenDekodieren(codec.fragmentInhalt(a.url));
      return {
        name: a.name,
        hinweis: a.hinweis,
        schemaImStrom: a.schema,
        payloadHex: Buffer.from(payload).toString("hex"),
        url: a.url,
        erwartet: normalisiert(erwartet),
      };
    }),
  ];
  const datei = {
    _hinweis:
      "GENERIERT aus dem TypeScript-Codec (scripts/testvektoren.mts) — nicht von Hand ändern. " +
      "Jede Portierung unter ports/ muss jeden Vektor hier so dekodieren, dass die Felder `erwartet` entsprechen. " +
      "`erwartet` ist der Bogen NACH Migration (schemaVersion = aktuelle SCHEMA_VERSION); `schemaImStrom` ist die Version im Bytestrom.",
    schemaVersion: SCHEMA_VERSION,
    epoche: "2020-01-01 (EebDatum = Tage, EebZeitpunkt = Minuten seit Epoche, lokale Wandzeit)",
    vektoren,
  };
  return JSON.stringify(datei, null, 2) + "\n";
}

// --------------------------------------------------------------- Wächter

function portsPruefen(): string[] {
  const fehler: string[] = [];
  const portsDir = join(wurzel, "ports");
  const ports = readdirSync(portsDir, { withFileTypes: true })
    .filter((d) => d.isDirectory() && d.name !== "testvektoren")
    .map((d) => d.name);
  if (ports.length === 0) fehler.push("ports/: keine Portierung gefunden");
  for (const p of ports) {
    const manifestPfad = join(portsDir, p, "PORT.json");
    if (!existsSync(manifestPfad)) {
      fehler.push(`ports/${p}: PORT.json fehlt`);
      continue;
    }
    const manifest = JSON.parse(readFileSync(manifestPfad, "utf8")) as { schemaVersion?: number; schemaKonstante?: { datei: string; muster: string } };
    if (manifest.schemaVersion !== SCHEMA_VERSION) {
      fehler.push(`ports/${p}/PORT.json: schemaVersion ${manifest.schemaVersion} ≠ SCHEMA_VERSION ${SCHEMA_VERSION} — Portierung nachziehen und Manifest heben`);
    }
    // Zweite Quelle im Quelltext der Portierung selbst, damit das Manifest
    // nicht ohne die Portierung gehoben werden kann.
    if (manifest.schemaKonstante) {
      const quelle = readFileSync(join(portsDir, p, manifest.schemaKonstante.datei), "latin1");
      const treffer = new RegExp(manifest.schemaKonstante.muster).exec(quelle);
      const wert = treffer?.[1] == null ? NaN : Number(treffer[1]);
      if (wert !== SCHEMA_VERSION) {
        fehler.push(`ports/${p}/${manifest.schemaKonstante.datei}: Schema-Konstante ${Number.isNaN(wert) ? "nicht gefunden" : wert} ≠ SCHEMA_VERSION ${SCHEMA_VERSION}`);
      }
    }
  }
  return fehler;
}

const pruefen = process.argv.includes("--pruefen");
const inhalt = erzeugen();
if (pruefen) {
  const fehler = portsPruefen();
  const vorhanden = existsSync(zieldatei) ? readFileSync(zieldatei, "utf8") : "";
  if (vorhanden !== inhalt) {
    fehler.push("ports/testvektoren/testvektoren.json ist veraltet — `npm run ports:testvektoren` laufen lassen und die Portierungen gegen die neuen Vektoren prüfen");
  }
  if (fehler.length > 0) {
    console.error("Portierungen nicht auf Stand:\n  - " + fehler.join("\n  - "));
    process.exit(1);
  }
  console.log(`ports/: alle Portierungen auf Schema ${SCHEMA_VERSION}, Testvektoren aktuell.`);
} else {
  writeFileSync(zieldatei, inhalt);
  console.log(`geschrieben: ${zieldatei}`);
}

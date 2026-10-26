import { describe, it, expect } from "vitest";
// pako statt node:zlib: der Kern muss auch im Testlauf ohne Node-Kernmodule
// auskommen, damit derselbe Test unter jsdom laufen kann (ADR-003, Regel 2).
import { deflateRaw, inflateRaw } from "pako";
import { decodePayloadUrl, type Kompressor } from "./codec.js";
import { Ernaehrung, Fahrerlaubnis, SCHEMA_VERSION } from "./model.js";

const kompressor: Kompressor = {
  deflateRaw: (d) => deflateRaw(d, { level: 9 }),
  inflateRaw: (d) => inflateRaw(d),
};

/**
 * EINGEFRORENE Golden-Fixture: ein echter QR-Payload im ALTEN Schema v2,
 * erzeugt mit dem v2-Wire-Format (Person ohne Ernährungs-Byte; Sofortbedarf
 * mit "davonVegetarisch"-Byte). Diese Bytes dürfen sich NIE ändern — sie
 * stehen stellvertretend für alle QR-Codes/Dateien, die Nutzer vor der
 * Schema-v3-Umstellung erzeugt haben. Bricht dieser Test, ist die
 * Abwärtskompatibilität verletzt (siehe docs/datenmodell.md, Migrationspolicy).
 *
 * Inhalt (v2): THW OV Oldenburg, FGr K (A), 3 Personen, 1 Fahrzeug,
 * Sofortbedarf für 20 Personen mit davonVegetarisch = 3.
 */
const V2_QR_URL =
  "https://erfassungsbogen.app/#RUVCMmNawc7IoC3oH6bgn5OSmpdUWpSuoKvgl8nIysiHKsLN4mjCyMrP4u_v4sqwgn01u6xbalFebmpOSmpSYunhPUmleekK3olJqTlJOZklVQzMHF75GYl5eanF7EGlKfk5BRmdjOxMDNyMBUImAfJcjAwsjnl5iazhqUmpRQIMDAzMIfm57G6ZxckZqUWsDAwMjMKMayezMoowM4AAAA";

describe("Abwärtskompatibilität: alter v2-QR-Code", () => {
  const bogen = decodePayloadUrl(V2_QR_URL, kompressor);

  it("wird auf die aktuelle Schema-Version gehoben", () => {
    expect(bogen.schemaVersion).toBe(SCHEMA_VERSION);
  });

  it("liest Einheit, Personal und Fahrzeug korrekt", () => {
    expect(bogen.einheit.hierarchie[0]!.name).toBe("Oldenburg - Ni");
    expect(bogen.einheit.einheitsTyp).toEqual({ code: 43 });
    expect(bogen.personal.map((p) => `${p.vorname} ${p.nachname}`)).toEqual([
      "Johannes Rudolph",
      "Anna Weber",
      "Tom Fischer",
    ]);
    expect(bogen.fahrzeuge[0]).toMatchObject({ typ: { code: 1 }, stanKonform: true });
  });

  it("lässt das Kennzeichen des alten THW-Sonderformats leer", () => {
    // Ab Schema 4 gibt es nur noch ein Kennzeichen-Feld als String. Die alte
    // THW-Nummer wird beim Decodieren bewusst verworfen (bewusste Entscheidung,
    // siehe docs/datenmodell.md) — der übrige Datenstrom bleibt lesbar.
    // Dass der Sofortbedarf hinter dem Fahrzeug noch stimmt (siehe unten),
    // belegt: die alte Varint-Nummer wurde sauber übersprungen.
    expect(bogen.fahrzeuge[0]!.kennzeichen).toBeUndefined();
  });

  it("setzt für v2-Personen die fehlende Ernährungsform auf FLEISCH", () => {
    // v2 kannte kein Ernährungsfeld — Migration muss einen sinnvollen Default setzen.
    expect(bogen.personal.every((p) => p.ernaehrung === Ernaehrung.FLEISCH)).toBe(true);
  });

  it("rettet die alte Aggregatzahl 'davonVegetarisch' in verpflegungManuell", () => {
    // In v2 stand die Vegetarier-Zahl im Sofortbedarf; ab v3 wird sie aus dem
    // Personal abgeleitet. Damit die Angabe nicht verlorengeht, migriert sie
    // in verpflegungManuell.
    expect(bogen.verpflegungManuell).toEqual({ vegetarisch: 3, vegan: 0 });
    expect(bogen.sofortbedarf?.verpflegungPersonen).toBe(20);
  });

  it("behält die erste Person mit dienstlichem E-Mail-Template und Mobilnummer", () => {
    const johannes = bogen.personal[0]!;
    expect(johannes.kontakte).toEqual([
      { art: 0, dienstlich: false, wert: "01701234501" },
      { art: 2, dienstlich: false, emailTemplate: 1 },
    ]);
  });
});

/**
 * EINGEFRORENE Golden-Fixture: QR-Payload im Schema v8, erzeugt vor der
 * Umstellung auf v9 (Sitzplatzzahl am Fahrzeug, Flag 32 im Fahrzeug-Byte).
 * Auch diese Bytes dürfen sich nie ändern. Sie belegen, dass das neue Flag
 * den Fahrzeug-Datenstrom nicht verschoben hat: bricht der Test, lesen alle
 * vor v9 erzeugten Codes ihre Fahrzeuge falsch.
 *
 * Inhalt (v8): THW OV Oldenburg, FGr K (A), eine Person mit zwei
 * Fahrerlaubnisklassen (CE + A), zwei Fahrzeuge (FüKomKw mit Funkrufname und
 * Änderungen, MzKW als Freitext), Sofortbedarf für 20 Personen.
 */
const V8_QR_URL =
  "https://erfassungsbogen.app/#B.LMAD3ATPYEAN0RTSMUSDVJNUHOUFCOFDM$1NNWDEOQ77QU4NBCC19D-*5FDM5GNHHC0VDFDM$FNPENAMN-JVBOU9WM37ML3N-XDFP45MNXAO92MU4ULNURNULETZ/TF3VMHY2NQ2YE*KG5KVRJUZZTIG3VWDJID/VM6MN38NAMN7NQJDI:-BN7G9O25M0T:GG6IK4N5$LS/DA$JXTJ1550JU9RA7/0J*0";

describe("Abwärtskompatibilität: v8-QR-Code ohne Sitzplatzzahl", () => {
  const bogen = decodePayloadUrl(V8_QR_URL, kompressor);

  it("wird auf die aktuelle Schema-Version gehoben", () => {
    expect(bogen.schemaVersion).toBe(SCHEMA_VERSION);
  });

  it("liest beide Fahrzeuge unverschoben — ohne erfasste Sitzplatzzahl", () => {
    expect(bogen.fahrzeuge[0]).toEqual({
      typ: { code: 3 },
      kennzeichen: "THW-84397",
      funkrufname: { kennwort: { code: 1 }, eigenerStandort: true, teile: [18, 13] },
      stanKonform: true,
      aenderungen: "Zusatzbeladung \u00d6lwehr",
    });
    expect(bogen.fahrzeuge[1]).toEqual({ typ: { freitext: "MzKW" }, kennzeichen: "OL-FW 2041" });
    expect(bogen.fahrzeuge.every((f) => f.sitzplaetze === undefined)).toBe(true);
  });

  it("liest die Felder hinter den Fahrzeugen weiter richtig", () => {
    // Der Beweis, dass nichts verrutscht ist: was im Datenstrom NACH den
    // Fahrzeugen steht, muss unverändert ankommen.
    expect(bogen.sofortbedarf).toMatchObject({ verpflegungPersonen: 20, dieselLiter: 200, unterbringung: true });
    expect(bogen.personal[0]!.weitereFahrerlaubnisse).toEqual([Fahrerlaubnis.A]);
  });
});

/**
 * EINGEFRORENE v9-Fixture (erzeugt mit dem Schema-9-Encoder vor der Einführung
 * von `Person.dienststelle` in Schema 10). Sie belegt, dass die Dienststelle
 * hinter dem Fahrerlaubnis-Zähler nur ab Schema 10 gelesen wird: Für v9-Bögen
 * darf dort kein String erwartet werden, sonst verrutschen Funktionen,
 * Kontakte und alles dahinter.
 *
 * Inhalt (v9): THW OV Oldenburg, FGr K (A), 2 Personen (erste mit weiterer
 * Fahrerlaubnis A), 1 Fahrzeug mit 7 Sitzplätzen, Sofortbedarf 2 Personen.
 */
const V9_QR_URL =
  "https://erfassungsbogen.app/#B.LMAD3A/RYY-HAPM/4M/RNP6MH8MDP9DR5OA7$HRB8E-ED2A7K5U83YTUM-C5SL59BJ1L-2JLD:TCMYH751-L-WUG-0S92$I-G/T1VS*2F6KOASM9YNBEL6-L:0$4Y6BDL13A6C0PEL8W6JKIP/K0FY-8BWOAIDL9Y1E*1HQ1ZVNXS4*DS4*58:M:0$APB9Y1YNB00";

describe("Abwärtskompatibilität: v9-QR-Code ohne Dienststelle je Person", () => {
  const bogen = decodePayloadUrl(V9_QR_URL, kompressor);

  it("wird auf die aktuelle Schema-Version gehoben", () => {
    expect(bogen.schemaVersion).toBe(SCHEMA_VERSION);
  });

  it("liest die Personen unverschoben — ohne Dienststelle", () => {
    expect(bogen.personal.map((p) => `${p.vorname} ${p.nachname}`)).toEqual(["Johannes Rudolph", "Anna Weber"]);
    expect(bogen.personal.every((p) => p.dienststelle === undefined)).toBe(true);
    expect(bogen.personal[0]!.weitereFahrerlaubnisse).toEqual([Fahrerlaubnis.A]);
    expect(bogen.personal[0]!.funktionen).toEqual([{ code: 7 }]);
    expect(bogen.personal[0]!.kontakte).toEqual([{ art: 0, dienstlich: true, wert: "01701234501" }]);
    expect(bogen.personal[0]!.zusatzqualifikationen).toEqual([{ code: 12 }]);
    expect(bogen.personal[1]!.ernaehrung).toBe(Ernaehrung.VEGETARISCH);
  });

  it("liest die Felder hinter dem Personal weiter richtig", () => {
    expect(bogen.fahrzeuge[0]).toEqual({ typ: { code: 3 }, kennzeichen: "THW-84397", sitzplaetze: 7 });
    expect(bogen.sofortbedarf).toMatchObject({ verpflegungPersonen: 2, dieselLiter: 50, ruhezeitErforderlich: true });
  });
});

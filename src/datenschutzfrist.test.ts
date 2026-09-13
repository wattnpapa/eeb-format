import { describe, it, expect } from "vitest";
import {
  ANONYM_BEZEICHNUNG,
  DATENSCHUTZFRIST_TAGE,
  UHR_SPRUNG_TAGE,
  bogenAnonymisiert,
  datenschutzfristAbgelaufen,
  datenschutzfristEnde,
  tageBisAnonymisierung,
  uhrPruefen,
} from "./datenschutzfrist.js";
import {
  Ernaehrung,
  Fahrerlaubnis,
  Geschlecht,
  KontaktArt,
  MINUTEN_JE_TAG,
  OrganisationsTyp,
  PersonalErfassung,
  SCHEMA_VERSION,
  StaerkeRolle,
  staerke,
  unterbringungMWD,
  verpflegung,
  type Erfassungsbogen,
  type Person,
} from "./model.js";

const TAG = MINUTEN_JE_TAG;
const STAND = 2000 * TAG;

function person(nachname: string, rolle: StaerkeRolle, over: Partial<Person> = {}): Person {
  return {
    vorname: "Anna",
    nachname,
    staerkeRolle: rolle,
    funktionen: [{ code: 3 }],
    fahrerlaubnis: Fahrerlaubnis.CE,
    weitereFahrerlaubnisse: [Fahrerlaubnis.A],
    geschlecht: Geschlecht.W,
    ernaehrung: Ernaehrung.VEGAN,
    kontakte: [{ art: KontaktArt.MOBIL, dienstlich: true, wert: "01711234567" }],
    zusatzqualifikationen: [{ freitext: "Atemschutz" }],
    ...over,
  };
}

function bogen(over: Partial<Erfassungsbogen> = {}): Erfassungsbogen {
  return {
    schemaVersion: SCHEMA_VERSION,
    stand: STAND,
    einheit: {
      organisation: OrganisationsTyp.THW,
      einheitsTyp: { code: 1 },
      hierarchie: [{ bezeichnung: { code: 1 }, name: "Oldenburg - Ni", kurz: "OODE", telefon: "0441123", email: "ov@example.org" }],
    },
    einsatz: { zeitraumVon: 2000, zeitraumBis: 2001, ortAuftrag: "Hochwasser" },
    personalErfassung: PersonalErfassung.VOLLSTAENDIG,
    personal: [
      person("Mannschaft", StaerkeRolle.MANNSCHAFT, { geschlecht: Geschlecht.M, ernaehrung: Ernaehrung.FLEISCH }),
      person("Führer", StaerkeRolle.FUEHRER),
      person("Unterführer", StaerkeRolle.UNTERFUEHRER, { ernaehrung: Ernaehrung.VEGETARISCH }),
    ],
    fahrzeuge: [{ typ: { code: 2 }, kennzeichen: "THW-84397" }],
    sonstiges: "Rückfragen an Anna Führer",
    ...over,
  };
}

describe("Datenschutzfrist", () => {
  it("endet 90 Tage nach dem Stand", () => {
    expect(datenschutzfristEnde(bogen())).toBe(STAND + DATENSCHUTZFRIST_TAGE * TAG);
  });

  it("gilt nicht für Übungsbögen", () => {
    const b = bogen({ uebung: true });
    expect(datenschutzfristEnde(b)).toBeNull();
    expect(datenschutzfristAbgelaufen(b, STAND + 10_000 * TAG)).toBe(false);
    expect(tageBisAnonymisierung(b, STAND + 10_000 * TAG)).toBeNull();
  });

  it("läuft genau mit dem Fristende ab", () => {
    const ende = STAND + 90 * TAG;
    expect(datenschutzfristAbgelaufen(bogen(), ende - 1)).toBe(false);
    expect(datenschutzfristAbgelaufen(bogen(), ende)).toBe(true);
  });

  it("kündigt die Anonymisierung erst in den letzten 14 Tagen an", () => {
    expect(tageBisAnonymisierung(bogen(), STAND + 75 * TAG)).toBeNull();
    expect(tageBisAnonymisierung(bogen(), STAND + 76 * TAG)).toBe(14);
    expect(tageBisAnonymisierung(bogen(), STAND + 89 * TAG + 1)).toBe(1);
    expect(tageBisAnonymisierung(bogen(), STAND + 120 * TAG)).toBe(0);
  });
});

describe("bogenAnonymisiert()", () => {
  const anonym = bogenAnonymisiert(bogen());

  it("entfernt Namen, Funktionen, Qualifikationen, Fahrerlaubnisse und Erreichbarkeiten", () => {
    for (const p of anonym.personal) {
      expect(p.vorname).toBe("");
      expect(p.nachname).toMatch(new RegExp(`^${ANONYM_BEZEICHNUNG} \\d+$`));
      expect(p.funktionen).toEqual([]);
      expect(p.zusatzqualifikationen).toEqual([]);
      expect(p.kontakte).toEqual([]);
      expect(p.fahrerlaubnis).toBe(Fahrerlaubnis.NONE);
      expect(p.weitereFahrerlaubnisse).toBeUndefined();
    }
    const text = JSON.stringify(anonym);
    for (const spur of ["Anna", "Führer", "Atemschutz", "01711234567", "0441123", "ov@example.org"]) {
      expect(text).not.toContain(spur);
    }
    expect(anonym.sonstiges).toBeUndefined();
  });

  it("erhält Stärke, Unterbringung und Verpflegung", () => {
    const b = bogen();
    expect(staerke(anonym)).toEqual(staerke(b));
    expect(unterbringungMWD(anonym)).toEqual(unterbringungMWD(b));
    expect(verpflegung(anonym)).toEqual(verpflegung(b));
  });

  it("erhält Einheit, Einsatz und Fahrzeuge", () => {
    expect(anonym.einheit.hierarchie).toEqual([{ bezeichnung: { code: 1 }, name: "Oldenburg - Ni", kurz: "OODE" }]);
    expect(anonym.einsatz).toEqual(bogen().einsatz);
    expect(anonym.fahrzeuge).toEqual(bogen().fahrzeuge);
    expect(anonym.stand).toBe(STAND);
  });

  it("sortiert nach Stärkerolle und nummeriert neu", () => {
    expect(anonym.personal.map((p) => [p.nachname, p.staerkeRolle])).toEqual([
      [`${ANONYM_BEZEICHNUNG} 1`, StaerkeRolle.FUEHRER],
      [`${ANONYM_BEZEICHNUNG} 2`, StaerkeRolle.UNTERFUEHRER],
      [`${ANONYM_BEZEICHNUNG} 3`, StaerkeRolle.MANNSCHAFT],
    ]);
  });

  it("ist idempotent und verändert das Original nicht", () => {
    const b = bogen();
    const vorher = JSON.stringify(b);
    const einmal = bogenAnonymisiert(b);
    expect(JSON.stringify(bogenAnonymisiert(einmal))).toBe(JSON.stringify(einmal));
    expect(JSON.stringify(b)).toBe(vorher);
  });
});

describe("uhrPruefen()", () => {
  it("übernimmt beim ersten Start die Geräteuhr", () => {
    expect(uhrPruefen(STAND, null)).toEqual({ zeitpunkt: STAND, stand: { zuletzt: STAND } });
  });

  it("folgt der Uhr im normalen Betrieb, auch nach langer Pause", () => {
    const r = uhrPruefen(STAND + UHR_SPRUNG_TAGE * TAG, { zuletzt: STAND });
    expect(r.zeitpunkt).toBe(STAND + UHR_SPRUNG_TAGE * TAG);
  });

  it("folgt einer zurückgestellten Uhr, statt Daten vorzeitig zu löschen", () => {
    expect(uhrPruefen(STAND - 500 * TAG, { zuletzt: STAND }).zeitpunkt).toBe(STAND - 500 * TAG);
  });

  it("hält einen unplausiblen Sprung nach vorn zurück, bis er sich einen Tag später bestätigt", () => {
    const sprung = STAND + 3000 * TAG;
    const erst = uhrPruefen(sprung, { zuletzt: STAND });
    expect(erst.zeitpunkt).toBe(STAND);
    expect(erst.stand).toEqual({ zuletzt: STAND, sprung });

    const gleichDanach = uhrPruefen(sprung + 60, erst.stand);
    expect(gleichDanach.zeitpunkt).toBe(STAND);
    expect(gleichDanach.stand.sprung).toBe(sprung);

    const tagDanach = uhrPruefen(sprung + TAG, gleichDanach.stand);
    expect(tagDanach).toEqual({ zeitpunkt: sprung + TAG, stand: { zuletzt: sprung + TAG } });
  });

  it("verwirft einen Sprungkandidaten, zu dem die Uhr nicht mehr passt", () => {
    const erst = uhrPruefen(STAND + 3000 * TAG, { zuletzt: STAND });
    const andererSprung = uhrPruefen(STAND + 5000 * TAG, erst.stand);
    expect(andererSprung.zeitpunkt).toBe(STAND);
    expect(andererSprung.stand.sprung).toBe(STAND + 5000 * TAG);
  });

  it("kehrt nach einem zurückgenommenen Sprung zum Normalbetrieb zurück", () => {
    const erst = uhrPruefen(STAND + 3000 * TAG, { zuletzt: STAND });
    expect(uhrPruefen(STAND + 60, erst.stand)).toEqual({ zeitpunkt: STAND + 60, stand: { zuletzt: STAND + 60 } });
  });
});

/**
 * Datenschutzfrist: Die personenbezogenen Angaben eines Bogens verfallen
 * 90 Tage nach seiner letzten inhaltlichen Änderung (`stand`). Übungsbögen sind
 * ausgenommen.
 *
 * Was die Frist leistet — und was nicht: Die App hält die Angaben nach Ablauf
 * nicht mehr vor. Gespeicherte Bögen werden dauerhaft überschrieben, ein
 * eingelesener Bogen wird gleich anonymisiert übernommen. Der QR-Code selbst
 * bleibt unverändert: Das Format ist offen und nicht verschlüsselt, ein
 * eigener Decoder liest ihn weiter. Die Frist begrenzt also die Speicherung auf
 * den Geräten (Art. 5 Abs. 1 lit. e DSGVO), sie schützt nicht den Transportweg.
 *
 * Bewusst KEIN Schemafeld: Die Frist leitet sich aus `stand` und `uebung` ab.
 * Damit gilt sie auch für alle schon verteilten QR-Codes, kostet keine QR-Bytes
 * und ältere App-Stände lesen die Bögen weiter. Eine Bearbeitung setzt `stand`
 * neu, und die Frist beginnt von vorn — ein Bogen, an dem noch gearbeitet wird,
 * verfällt nicht.
 *
 * Plattformneutral wie model.ts.
 */

import {
  Fahrerlaubnis,
  MINUTEN_JE_TAG,
  type EebZeitpunkt,
  type Erfassungsbogen,
  type HierarchieEbene,
  type Person,
} from "./model.js";

/** Tage nach der letzten Änderung, nach denen die Personaldaten verfallen. */
export const DATENSCHUTZFRIST_TAGE = 90;

/** So viele Tage vor Ablauf kündigt die App die Anonymisierung an. */
export const DATENSCHUTZFRIST_HINWEIS_TAGE = 14;

/** Nachname, unter dem eine anonymisierte Person geführt wird („Einsatzkraft 3"). */
export const ANONYM_BEZEICHNUNG = "Einsatzkraft";

/** Ende der Frist als Zeitpunkt — `null` bei Übungsbögen, die keine Frist haben. */
export function datenschutzfristEnde(b: Pick<Erfassungsbogen, "uebung" | "stand">): EebZeitpunkt | null {
  return b.uebung ? null : b.stand + DATENSCHUTZFRIST_TAGE * MINUTEN_JE_TAG;
}

export function datenschutzfristAbgelaufen(
  b: Pick<Erfassungsbogen, "uebung" | "stand">,
  jetzt: EebZeitpunkt,
): boolean {
  const ende = datenschutzfristEnde(b);
  return ende != null && jetzt >= ende;
}

/**
 * Ganze Tage bis zur Anonymisierung, sobald die Ankündigung fällig ist, sonst
 * `null`. Kein Wert bei Übungsbögen; 0 heißt „abgelaufen".
 */
export function tageBisAnonymisierung(
  b: Pick<Erfassungsbogen, "uebung" | "stand">,
  jetzt: EebZeitpunkt,
): number | null {
  const ende = datenschutzfristEnde(b);
  if (ende == null) return null;
  const rest = ende - jetzt;
  if (rest > DATENSCHUTZFRIST_HINWEIS_TAGE * MINUTEN_JE_TAG) return null;
  return Math.max(0, Math.ceil(rest / MINUTEN_JE_TAG));
}

/**
 * Kopie des Bogens ohne Personenbezug.
 *
 * Entfernt: Namen, Funktionen, Zusatzqualifikationen, Fahrerlaubnisse,
 * Erreichbarkeiten der Personen und der Hierarchie-Ebenen sowie den Freitext
 * „Sonstiges", in dem Namen stehen können.
 *
 * Erhalten: Stärkerolle, Geschlecht und Ernährung je Person, weil Stärke,
 * Unterbringung und Verpflegung daraus abgeleitet werden. Dazu Einheit,
 * Einsatz, Fahrzeuge und Sofortbedarf, die keine Personen beschreiben.
 *
 * Die Personen werden nach Stärkerolle umsortiert und neu durchnummeriert —
 * die ursprüngliche Reihenfolge würde sonst verraten, wer an welcher Stelle
 * stand. Idempotent: ein zweiter Durchlauf ändert nichts mehr.
 */
export function bogenAnonymisiert(b: Erfassungsbogen): Erfassungsbogen {
  const personal = b.personal
    .map((p, i) => ({ p, i }))
    .sort((x, y) => y.p.staerkeRolle - x.p.staerkeRolle || x.i - y.i)
    .map(
      ({ p }, n): Person => ({
        nachname: `${ANONYM_BEZEICHNUNG} ${n + 1}`,
        vorname: "",
        staerkeRolle: p.staerkeRolle,
        funktionen: [],
        fahrerlaubnis: Fahrerlaubnis.NONE,
        geschlecht: p.geschlecht,
        ernaehrung: p.ernaehrung,
        kontakte: [],
        zusatzqualifikationen: [],
      }),
    );
  const hierarchie = b.einheit.hierarchie.map(
    (e): HierarchieEbene => ({
      bezeichnung: e.bezeichnung,
      name: e.name,
      ...(e.kurz != null ? { kurz: e.kurz } : {}),
    }),
  );
  const ergebnis: Erfassungsbogen = { ...b, einheit: { ...b.einheit, hierarchie }, personal };
  delete ergebnis.sonstiges;
  return ergebnis;
}

// ------------------------------------------------------------ Uhr-Plausibilität

/**
 * Was sich die App über die Geräteuhr merkt: den zuletzt akzeptierten
 * Zeitpunkt und gegebenenfalls einen noch unbestätigten Sprung nach vorn.
 */
export interface Uhrstand {
  /** Zuletzt akzeptierter Zeitpunkt. */
  zuletzt: EebZeitpunkt;
  /** Erster Zeitpunkt nach einem unplausiblen Sprung — wartet auf Bestätigung. */
  sprung?: EebZeitpunkt;
}

/** Ab diesem Vorsprung gegenüber dem letzten Start gilt die Uhr als gesprungen. */
export const UHR_SPRUNG_TAGE = 366;

/** Innerhalb dieses Fensters muss sich ein Sprung bestätigen, sonst beginnt die Prüfung neu. */
const SPRUNG_FENSTER_TAGE = 30;

/**
 * Welche Uhrzeit gilt für die Frist?
 *
 * Die Anonymisierung lässt sich nicht rückgängig machen. Eine Uhr, die falsch
 * nach vorn springt — leerer Akku, Werkszustand, eine Zeitsynchronisation, die
 * erst nach dem Start greift —, darf deshalb nicht mitten im Einsatz alle
 * Personaldaten löschen. Liegt die Uhr mehr als {@link UHR_SPRUNG_TAGE} Tage
 * vor dem letzten Start, gilt weiter der letzte akzeptierte Zeitpunkt. Den
 * neuen übernimmt die App erst, wenn die Uhr frühestens einen Tag später noch
 * dazu passt.
 *
 * Eine zurückgestellte Uhr wird NICHT abgefangen. Sie kann ohnehin nur
 * verhindern, was beim letzten Start noch nicht fällig war — dafür genügt es
 * schon, die App nicht zu öffnen. Die Frist ist kein Schutz gegen Absicht
 * (siehe Modulkopf).
 */
export function uhrPruefen(
  jetzt: EebZeitpunkt,
  bisher: Uhrstand | null,
): { zeitpunkt: EebZeitpunkt; stand: Uhrstand } {
  if (!bisher || jetzt - bisher.zuletzt <= UHR_SPRUNG_TAGE * MINUTEN_JE_TAG) {
    return { zeitpunkt: jetzt, stand: { zuletzt: jetzt } };
  }
  const s = bisher.sprung;
  const passtZumSprung = s != null && jetzt >= s && jetzt - s <= SPRUNG_FENSTER_TAGE * MINUTEN_JE_TAG;
  if (passtZumSprung && jetzt - s >= MINUTEN_JE_TAG) {
    return { zeitpunkt: jetzt, stand: { zuletzt: jetzt } };
  }
  return {
    zeitpunkt: bisher.zuletzt,
    stand: { zuletzt: bisher.zuletzt, sprung: passtZumSprung ? s : jetzt },
  };
}

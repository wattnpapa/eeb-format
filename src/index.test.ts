import { describe, expect, it } from "vitest";

import { inhaltsHash, kernVersion } from "./index.js";

// Diese Datei läuft in beiden Vitest-Projekten: einmal unter „node", einmal
// unter „jsdom". Dass beide Läufe dieselben Werte liefern, ist der Beleg für
// die Plattformneutralität des Kerns.
describe("@bos/eeb-format", () => {
  it("meldet die Kernversion", () => {
    expect(kernVersion()).toBe("0.0.0");
  });

  it("hasht denselben Text immer gleich", () => {
    expect(inhaltsHash("Bergungsgruppe")).toBe(inhaltsHash("Bergungsgruppe"));
  });

  it("unterscheidet verschiedene Texte", () => {
    expect(inhaltsHash("Zugtrupp")).not.toBe(inhaltsHash("Fachgruppe"));
  });

  it("liefert acht Hexziffern", () => {
    expect(inhaltsHash("")).toMatch(/^[0-9a-f]{8}$/);
    expect(inhaltsHash("Überörtliche Hilfe")).toMatch(/^[0-9a-f]{8}$/);
  });

  it("behandelt Umlaute unabhängig von der Unicode-Kodierung gleich", () => {
    const vorkomponiert = "Gef\u00e4hrdung";
    const zerlegt = "Gefa\u0308hrdung";
    expect(vorkomponiert).not.toBe(zerlegt);
    expect(inhaltsHash(zerlegt)).toBe(inhaltsHash(vorkomponiert));
  });

  it("bleibt über Umgebungen hinweg bei festen Werten", () => {
    // Feste Erwartungswerte (FNV-1a, 32 Bit). Weichen Node- und jsdom-Lauf
    // voneinander ab, fällt genau hier eine Umgebungsabhängigkeit auf.
    expect(inhaltsHash("")).toBe("811c9dc5");
    expect(inhaltsHash("abc")).toBe("1a47e90b");
    expect(inhaltsHash("S1-Control")).toBe("05d6ffd1");
  });
});

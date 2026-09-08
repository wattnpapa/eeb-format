/**
 * `@bos/eeb-format` — das Austauschformat des Einheiten-Erfassungsbogens.
 *
 * Drei Bausteine, aufeinander aufbauend:
 *  - `model`    — der Bogen selbst: Typen, Zählregeln, Schema-Migration.
 *                 Importiert nichts.
 *  - `codec`    — Bogen ⇄ Bytes ⇄ QR-URL: Base41, Segmentierung, Kompression
 *                 (die Kompression wird hineingereicht, siehe `Kompressor`).
 *  - `signatur` — die Ed25519-Kette über einem kodierten Bogen.
 *
 * Die Node-Implementierung (`qr-node`) hängt nicht hier drin, sondern am
 * Einstieg `@bos/eeb-format/node`. Sie darf `node:zlib`, `qrcode` und `Buffer`
 * benutzen; alles in dieser Datei darf das nicht — Aufnahmeregel 2 aus ADR-003
 * gilt in einem Format-Repo je Sprachimplementierung, nicht über das Repo
 * hinweg.
 *
 * Alles hier drin muss ohne `node:`-, DOM- oder Framework-Zugriffe auskommen
 * und in Node wie im Browser bitgleiche Ergebnisse liefern.
 */

export * from "./model.js";
export * from "./codec.js";
export * from "./signatur.js";

/** Version des Kerns, gepflegt im Gleichklang mit `package.json`. */
const VERSION = "0.0.0";

/**
 * Liefert die Version des geteilten Kerns.
 *
 * Dient den Konsumenten als Beleg, dass der Submodul-Stand tatsächlich geladen
 * wurde, und taugt später für Diagnoseausgaben.
 */
export function kernVersion(): string {
  return VERSION;
}

/**
 * Kodiert einen Text nach UTF-8, ohne `TextEncoder`.
 *
 * `TextEncoder` ist zwar in Node und in Browsern vorhanden, kommt typseitig
 * aber nur über `lib.dom` oder `@types/node` — beides ist hier bewusst
 * abgeschaltet. Die Handkodierung hält den Kern frei von Plattformtypen und
 * liefert in jeder Umgebung dieselben Bytes.
 */
function nachUtf8(text: string): number[] {
  const bytes: number[] = [];
  for (const zeichen of text) {
    const punkt = zeichen.codePointAt(0);
    if (punkt === undefined) continue;
    if (punkt < 0x80) {
      bytes.push(punkt);
    } else if (punkt < 0x800) {
      bytes.push(0xc0 | (punkt >> 6), 0x80 | (punkt & 0x3f));
    } else if (punkt < 0x10000) {
      bytes.push(0xe0 | (punkt >> 12), 0x80 | ((punkt >> 6) & 0x3f), 0x80 | (punkt & 0x3f));
    } else {
      bytes.push(
        0xf0 | (punkt >> 18),
        0x80 | ((punkt >> 12) & 0x3f),
        0x80 | ((punkt >> 6) & 0x3f),
        0x80 | (punkt & 0x3f),
      );
    }
  }
  return bytes;
}

/**
 * Bildet einen stabilen Inhalts-Hash über einen Text (FNV-1a, 32 Bit).
 *
 * Rein, ohne Zufall, ohne Uhrzeit, ohne Plattform-API: derselbe Text ergibt in
 * Node und im Browser dieselben acht Hex-Zeichen. Gedacht zum Erkennen
 * unveränderter Inhalte (Zwischenspeicher, Abgleich), ausdrücklich **nicht**
 * für kryptografische Zwecke.
 *
 * Der Text wird vor dem Hashen unicode-normalisiert (NFC), damit gleich
 * aussehende Umlaute unabhängig von ihrer Kodierung denselben Hash ergeben.
 */
export function inhaltsHash(text: string): string {
  let hash = 0x811c9dc5;
  for (const byte of nachUtf8(text.normalize("NFC"))) {
    hash ^= byte;
    // FNV-Primzahl 16777619, per Verschiebungen multipliziert, damit die
    // Rechnung im 32-Bit-Bereich bleibt.
    hash = (hash + ((hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24))) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

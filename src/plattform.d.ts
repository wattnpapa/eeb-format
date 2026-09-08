/**
 * Die zwei Web-Standard-Globals, die der Codec braucht — hier von Hand
 * deklariert.
 *
 * `TextEncoder` und `TextDecoder` gibt es in jeder Zielumgebung (Node seit 11,
 * jeder Browser, jede WebView). Ihre TYPEN kommen jedoch nur über `lib.dom`
 * oder `@types/node`, und beides ist in `tsconfig.json` bewusst abgeschaltet:
 * fehlt DOM in `lib` und ist `types` leer, scheitert schon die Typprüfung an
 * `document`, `window`, `process` oder `Buffer` (ADR-003, Aufnahmeregel 2).
 *
 * Statt dieser Wache ein Loch zu schlagen, stehen hier genau die beiden
 * Schnittstellen, und zwar nur mit dem, was der Codec tatsächlich aufruft. Wer
 * hier etwas ergänzt, sollte begründen können, warum das Ergänzte in Node und
 * im Browser dasselbe tut.
 *
 * Im Bau (`tsconfig.build.json`) ist diese Datei ausgenommen: dort liefert
 * `@types/node` die Deklarationen, und zwei Deklarationen desselben Namens
 * wären ein Fehler.
 */

declare class TextEncoder {
  encode(eingabe?: string): Uint8Array;
}

declare class TextDecoder {
  constructor(kennung?: string);
  decode(eingabe?: Uint8Array): string;
}

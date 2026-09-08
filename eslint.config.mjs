import js from "@eslint/js";
import tseslint from "typescript-eslint";

// Die Aufnahmeregeln aus ADR-003 sind hier maschinell hinterlegt. Was der Kern
// nicht importieren darf, fällt beim Lint durch — nicht erst beim Konsumenten.
const VERBOTENE_MUSTER = [
  {
    group: ["node:*"],
    message:
      "Der geteilte Kern muss plattformneutral bleiben: keine Node-Kernmodule (ADR-003, Aufnahmeregel 2).",
  },
  {
    group: ["electron", "electron/*", "electron-updater"],
    message: "Electron gehört in die Anwendungsschicht von S1-Control, nicht in den Kern (ADR-003).",
  },
  {
    group: ["@capacitor/*", "@capacitor-community/*"],
    message: "Capacitor gehört in die Anwendungsschicht des Erfassungsbogens, nicht in den Kern (ADR-003).",
  },
  {
    group: ["react", "react-dom", "react/*", "react-dom/*", "@testing-library/*"],
    message: "Keine React- oder DOM-Testbibliotheken im Kern (ADR-003, Aufnahmeregel 2).",
  },
  {
    group: ["@s1/*", "s1-control", "s1-control/*"],
    message: "Keine Rückimporte aus S1-Control (ADR-003, Aufnahmeregel 3).",
  },
  {
    group: ["erfassungsbogen", "erfassungsbogen/*", "@erfassungsbogen/*"],
    message: "Keine Rückimporte aus der Erfassungsbogen-App (ADR-003, Aufnahmeregel 3).",
  },
];

// Die bloßen Kernmodulnamen ohne node:-Präfix. Als Namensliste statt als Muster,
// damit etwa ein späteres Paket „node-fetch" nicht versehentlich mitgefangen wird.
const NODE_KERNMODULE = [
  "assert", "async_hooks", "buffer", "child_process", "cluster", "console",
  "constants", "crypto", "dgram", "diagnostics_channel", "dns", "domain",
  "events", "fs", "fs/promises", "http", "http2", "https", "inspector",
  "module", "net", "os", "path", "path/posix", "path/win32", "perf_hooks",
  "process", "punycode", "querystring", "readline", "repl", "stream",
  "string_decoder", "sys", "timers", "tls", "trace_events", "tty", "url",
  "util", "v8", "vm", "wasi", "worker_threads", "zlib",
].map((name) => ({
  name,
  message:
    "Der geteilte Kern muss plattformneutral bleiben: keine Node-Kernmodule (ADR-003, Aufnahmeregel 2).",
}));

// Browser-eigene Pakete, die typische Einfallstore für DOM-Abhängigkeiten sind.
const BROWSER_PAKETE = [
  "jsdom", "jsqr", "qrcode", "zxing-wasm", "pdfmake", "workbox-window",
  "virtual:pwa-register",
].map((name) => ({
  name,
  message: "Browser-/DOM-nahe Pakete gehören in die Anwendungsschicht, nicht in den Kern (ADR-003).",
}));

export default tseslint.config(
  { ignores: ["dist/**", "coverage/**", "node_modules/**"] },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
  {
    files: ["src/**/*.ts"],
    rules: {
      "no-restricted-imports": "off",
      "@typescript-eslint/no-restricted-imports": [
        "error",
        {
          paths: [...NODE_KERNMODULE, ...BROWSER_PAKETE],
          patterns: VERBOTENE_MUSTER,
        },
      ],
    },
  },
  {
    // Konfigurationsdateien liegen außerhalb des ausgelieferten Kerns und
    // dürfen daher Werkzeugpakete einbinden.
    files: ["*.mjs", "*.ts", "vitest.config.ts"],
    ignores: ["src/**"],
    ...tseslint.configs.disableTypeChecked,
  },
);

import { defineConfig } from "vitest/config";

// Ein Lauf, zwei Umgebungen: dieselben Testdateien werden einmal unter Node und
// einmal unter jsdom ausgeführt. Damit ist maschinell belegt, dass der Kern
// weder Node-Globals noch ein DOM voraussetzt — Aufnahmeregel 2 aus ADR-003.
export default defineConfig({
  test: {
    projects: [
      {
        test: {
          name: "node",
          environment: "node",
          include: ["src/**/*.test.ts"],
        },
      },
      {
        test: {
          name: "jsdom",
          environment: "jsdom",
          include: ["src/**/*.test.ts"],
        },
      },
    ],
  },
});

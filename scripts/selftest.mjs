// Сборка и запуск самопроверки в Node: node scripts/selftest.mjs <справочник.xlsx> "kind|in.xlsx|out.xlsx" ...
import { build } from "esbuild";
import { spawnSync } from "node:child_process";
await build({ entryPoints: ["scripts/selftest-entry.ts"], bundle: true, platform: "node", format: "esm", outfile: "node_modules/.selftest/run.mjs", logLevel: "warning" });
const r = spawnSync(process.execPath, ["node_modules/.selftest/run.mjs", ...process.argv.slice(2)], { stdio: "inherit" });
process.exit(r.status ?? 1);

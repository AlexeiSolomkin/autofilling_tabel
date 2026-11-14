// Сверка сервиса с табелями, заполненными вручную:
//   node scripts/report.mjs <справочник.xlsx> <папка с табелями> <отчёт.md>
import { build } from "esbuild";
import { spawnSync } from "node:child_process";
await build({ entryPoints: ["scripts/report-entry.ts"], bundle: true, platform: "node", format: "esm", outfile: "node_modules/.selftest/report.mjs", logLevel: "warning" });
const r = spawnSync(process.execPath, ["node_modules/.selftest/report.mjs", ...process.argv.slice(2)], { stdio: "inherit" });
process.exit(r.status ?? 1);

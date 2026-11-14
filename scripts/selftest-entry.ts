// Самопроверка: прогон сервиса в Node на файлах из каталога, переданного аргументом.
import { readFileSync, writeFileSync } from "node:fs";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { Workbook, type XmlEnv } from "../src/xlsx/workbook";
import { buildReference } from "../src/reference";
import { fillTabel, type TabelKind } from "../src/tabel";

const env: XmlEnv = {
  parse: (t) => new DOMParser().parseFromString(t, "application/xml") as unknown as Document,
  serialize: (d) => new XMLSerializer().serializeToString(d as any),
};

const [refPath, ...jobs] = process.argv.slice(2);
let t0 = performance.now();
const ref = buildReference(new Workbook(readFileSync(refPath), env));
console.log(`Справочник: ${(performance.now() - t0).toFixed(0)} мс;`,
  Object.values(ref.sheets).map((ix) => `${ix!.sheetName}: ${ix!.clients.length} записей, ${ix!.markingCount} маркировок`).join("; "),
  ref.errors.length ? `ошибки: ${ref.errors.join("; ")}` : "");

for (const job of jobs) {
  const [kind, inp, out] = job.split("|") as [TabelKind, string, string];
  t0 = performance.now();
  const res = fillTabel(new Workbook(readFileSync(inp), env), kind, ref);
  console.log(`\n##### ${kind} ${inp} (${(performance.now() - t0).toFixed(0)} мс)`, JSON.stringify(res.stats));
  for (const g of res.general) console.log(`  [${g.level}] ${g.text}`);
  for (const t of res.tables) {
    console.log(`  --- ${t.title}`);
    for (const e of t.entries) console.log(`    [${e.level}] ${e.text}`);
  }
  if (res.output) writeFileSync(out, res.output);
}

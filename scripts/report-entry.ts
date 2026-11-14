/**
 * Сверка сервиса с табелями, заполненными вручную.
 *
 *   node scripts/report.mjs <справочник.xlsx> <папка с табелями> <отчёт.md>
 *
 * Для каждого заполненного табеля: в обрабатываемых таблицах стираются «№ клиента», «Клиент», «Склад»,
 * табель заполняется сервисом, результат построчно сравнивается с исходными значениями.
 * Пустые шаблоны из <папка>/empty_tabels («… пустой.xlsx») заполняются как есть и сравниваются
 * с заполненной парой (тот же файл без «пустой»).
 */
import { readFileSync, readdirSync, writeFileSync, existsSync } from "node:fs";
import { join, basename } from "node:path";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { Workbook, type XmlEnv } from "../src/xlsx/workbook";
import { buildReference, type Reference } from "../src/reference";
import { fillTabel, findHeaders, dataRows, pickSheet, type TabelKind, type LogEntry, type HeaderInfo } from "../src/tabel";
import type { Sheet } from "../src/xlsx/workbook";

const env: XmlEnv = {
  parse: (t) => new DOMParser().parseFromString(t, "application/xml") as unknown as Document,
  serialize: (d) => new XMLSerializer().serializeToString(d as any),
};

interface RowVals { num: string; name: string; wh: string; marking: string; extra: string }
interface RowRef { h: HeaderInfo; r: number }

function kindOf(file: string): TabelKind {
  return /(^|[\s_])г([\s_.]|$)/i.test(basename(file)) ? "holland" : "import";
}

function tableRows(sh: Sheet, kind: TabelKind): RowRef[] {
  const headers = findHeaders(sh);
  const list = kind === "import" ? headers.slice(1) : headers;
  const out: RowRef[] = [];
  for (const h of list) {
    const next = headers[headers.indexOf(h) + 1]?.row ?? sh.maxRow + 1;
    for (const r of dataRows(sh, h, next)) out.push({ h, r });
  }
  return out;
}

function vals(sh: Sheet, { h, r }: RowRef, kind: TabelKind): RowVals {
  const t = (c: number | null) => (c ? sh.text(c, r).trim() : "");
  return {
    num: t(h.number), name: t(h.client), wh: t(h.wh), marking: t(h.mark),
    extra: kind === "holland" ? t(h.supplier) : t(h.awb),
  };
}

const nName = (s: string) => s.toLowerCase().replace(/ё/g, "е").replace(/[^\p{L}\p{N}]+/gu, " ").trim();
const nWh = (s: string) => {
  const x = s.toLowerCase().trim();
  if (x === "к" || x.startsWith("котел")) return "котельники";
  if (x === "х" || x.startsWith("химк")) return "химки";
  return x;
};
const nNum = (s: string) => s.replace(/\s+/g, "");

interface Cmp {
  file: string;
  kind: TabelKind;
  rows: number;
  exact: number;
  diffs: string[];
  notFilled: string[];
  origEmpty: string[];
  oddities: string[];
  general: string[];
  diffCount: number;
  uncertainOk: number;
  uncertainBad: number;
  defaultOk: number;
  defaultBad: number;
}

function compare(file: string, kind: TabelKind, sh: Sheet, rows: RowRef[], orig: Map<number, RowVals>, log: LogEntry[], general: LogEntry[]): Cmp {
  const byRow = new Map<number, LogEntry>();
  for (const e of log) if (e.row && e.outcome) byRow.set(e.row, e);
  const c: Cmp = { file, kind, rows: 0, exact: 0, diffs: [], notFilled: [], origEmpty: [], oddities: [], general: [], diffCount: 0, uncertainOk: 0, uncertainBad: 0, defaultOk: 0, defaultBad: 0 };
  for (const g of general) if (g.level !== "info") c.general.push(g.text);

  for (const rr of rows) {
    const o = orig.get(rr.r)!;
    const n = vals(sh, rr, kind);
    if (!o.marking || sh.hasFormula(rr.h.mark, rr.r)) continue;
    c.rows++;
    const e = byRow.get(rr.r);
    const where = `стр. ${rr.r}, ${kind === "holland" ? (o.extra || "—") : `AWB ${o.extra || "—"}`}, **${o.marking}**`;
    const origTxt = `№${o.num || "—"} «${o.name || "—"}», ${o.wh || "—"}`;
    const ourTxt = `№${n.num || "—"} «${n.name || "—"}», ${n.wh || "—"}`;
    const src = e?.refSheet ? ` [лист ${e.refSheet}${e.uncertain ? ", неточное совпадение" : ""}]` : "";

    if (!o.num && !o.name) {
      c.origEmpty.push(`- ${where}: у тебя строка не заполнена; сервис: ${n.num ? ourTxt + src : "ничего — " + (e?.text.split(": ").slice(1).join(": ") ?? "")}`);
      continue;
    }
    if (!n.num && !n.name) {
      const reason = e ? e.text.replace(/^стр\. \d+(, [^,]+)?, маркировка «[^»]*»:?\s*/, "") : "строка не обработана";
      c.notFilled.push(`- ${where}: у тебя ${origTxt}. Сервис не заполнил: ${reason}`);
      continue;
    }
    const problems: string[] = [];
    if (nNum(o.num) !== nNum(n.num)) problems.push("номер");
    if (nName(o.name) !== nName(n.name)) problems.push("клиент");
    if (o.wh && n.wh && nWh(o.wh) !== nWh(n.wh)) problems.push("склад");
    if (!o.wh && n.wh) c.oddities.push(`- ${where}: у тебя склад не указан, сервис поставил «${n.wh}»`);
    if (!problems.length) {
      c.exact++;
      if (e?.defaulted) c.defaultOk++;
      else if (e?.uncertain) c.uncertainOk++;
      continue;
    }
    c.diffCount++;
    if (e?.defaulted) c.defaultBad++;
    else if (e?.uncertain) c.uncertainBad++;
    c.diffs.push(`- ${where} — отличается: ${problems.join(", ")}. У тебя ${origTxt}; у сервиса ${ourTxt}${src}`);
  }
  return c;
}

function runFilled(ref: Reference, path: string): Cmp {
  const kind = kindOf(path);
  const wb = new Workbook(readFileSync(path), env);
  const sh = pickSheet(wb);
  const rows = tableRows(sh, kind);
  const orig = new Map<number, RowVals>();
  for (const rr of rows) {
    orig.set(rr.r, vals(sh, rr, kind));
    for (const col of [rr.h.number, rr.h.client, rr.h.wh]) if (col) sh.clear(col, rr.r);
  }
  const res = fillTabel(wb, kind, ref);
  const log = res.tables.flatMap((t) => t.entries);
  return compare(basename(path), kind, sh, rows, orig, log, res.general);
}

function runEmpty(ref: Reference, emptyPath: string, filledPath: string): Cmp {
  const kind = kindOf(filledPath);
  const fw = new Workbook(readFileSync(filledPath), env);
  const fsh = pickSheet(fw);
  const frows = tableRows(fsh, kind);
  const orig = new Map<number, RowVals>();
  for (const rr of frows) orig.set(rr.r, vals(fsh, rr, kind));

  const wb = new Workbook(readFileSync(emptyPath), env);
  const res = fillTabel(wb, kind, ref);
  const sh = pickSheet(wb);
  const rows = tableRows(sh, kind);
  const c = compare(`${basename(emptyPath)} (сравнение с «${basename(filledPath)}»)`, kind, sh, rows.filter((rr) => orig.has(rr.r)), orig, res.tables.flatMap((t) => t.entries), res.general);
  for (const rr of rows) {
    const m = sh.text(rr.h.mark, rr.r).trim();
    const om = orig.get(rr.r)?.marking ?? "";
    if (m && om && m.toLowerCase() !== om.toLowerCase()) c.oddities.push(`- стр. ${rr.r}: в пустом шаблоне маркировка «${m}», а в заполненном «${om}» — строки не совпадают`);
    if (m && !orig.has(rr.r)) c.oddities.push(`- стр. ${rr.r}: строка с маркировкой «${m}» есть только в пустом шаблоне`);
  }
  return c;
}

function section(c: Cmp): string {
  const out: string[] = [];
  out.push(`## ${c.file}`);
  out.push(`${c.kind === "holland" ? "Голландия" : "Импорт"} · строк с маркировкой: ${c.rows} · полностью совпало: ${c.exact} · отличается: ${c.diffCount} · не заполнено сервисом: ${c.notFilled.length}${c.origEmpty.length ? ` · у тебя не заполнено: ${c.origEmpty.length}` : ""}`);
  out.push("");
  if (c.general.length) out.push("**Общее:**", ...c.general.map((g) => `- ${g}`), "");
  if (c.diffs.length) out.push("**Расхождения (сервис заполнил иначе):**", ...c.diffs, "");
  if (c.notFilled.length) out.push("**Сервис не заполнил:**", ...c.notFilled, "");
  if (c.origEmpty.length) out.push("**У тебя строка пустая:**", ...c.origEmpty, "");
  if (c.oddities.length) out.push("**Странности:**", ...c.oddities, "");
  if (!c.diffs.length && !c.notFilled.length && !c.oddities.length && !c.origEmpty.length) out.push("Всё совпало.", "");
  return out.join("\n");
}

const [refPath, dir, outPath] = process.argv.slice(2);
const ref = buildReference(new Workbook(readFileSync(refPath), env));
const files = readdirSync(dir).filter((f) => f.endsWith(".xlsx") && f.startsWith("Табель")).sort();
const results: Cmp[] = [];
for (const f of files) results.push(runFilled(ref, join(dir, f)));
const emptyDir = join(dir, "empty_tabels");
if (existsSync(emptyDir)) {
  for (const f of readdirSync(emptyDir).filter((x) => x.endsWith(".xlsx")).sort()) {
    const filled = join(dir, f.replace(/\s*пустой(?=\.xlsx$)/, ""));
    if (existsSync(filled)) results.push(runEmpty(ref, join(emptyDir, f), filled));
  }
}

const total = results.reduce((a, c) => ({ rows: a.rows + c.rows, exact: a.exact + c.exact, diff: a.diff + c.diffCount, nf: a.nf + c.notFilled.length }), { rows: 0, exact: 0, diff: 0, nf: 0 });
const md: string[] = [
  "# Сверка автозаполнения с табелями, заполненными вручную",
  "",
  `Справочник: ${basename(refPath)} · табелей: ${results.length}`,
  "",
  `Итого строк: ${total.rows} · полностью совпало: ${total.exact} · отличается: ${total.diff} · не заполнено сервисом: ${total.nf}`,
  "",
  "Сравнение: номер — точно; клиент — без учёта регистра, скобок и знаков препинания («МФ ( УФА )» = «МФ УФА»); склад — «К» = Котельники, «Х» = Химки.",
  "",
  "| Табель | Строк | Совпало | Отличается | Не заполнено |",
  "|---|---|---|---|---|",
  ...results.map((c) => `| ${c.file} | ${c.rows} | ${c.exact} | ${c.diffCount} | ${c.notFilled.length} |`),
  "",
  ...results.map(section),
];
writeFileSync(outPath, md.join("\n"));
const uniqRes = results.filter((c) => !c.file.includes("пустой"));
const sum = (f: (c: Cmp) => number) => uniqRes.reduce((a, c) => a + f(c), 0);
console.log("UNCERTAIN ok/bad:", sum((c) => c.uncertainOk), sum((c) => c.uncertainBad), "DEFAULT ok/bad:", sum((c) => c.defaultOk), sum((c) => c.defaultBad));
console.log(md.slice(0, 12 + results.length).join("\n"));

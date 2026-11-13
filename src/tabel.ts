/**
 * Заполнение табелей (импорт / Голландия).
 */
import type { Sheet, Workbook } from "./xlsx/workbook";
import { addr } from "./xlsx/address";
import { normalizeHeader, namesCompatible, isMultiWarehouse } from "./normalize";
import { lookup, type ClientEntry, type Reference, type RefSheetIndex } from "./reference";

export type TabelKind = "import" | "holland";
export type LogLevel = "filled" | "ok" | "info" | "warn" | "error";

export interface LogEntry {
  level: LogLevel;
  text: string;
}

export interface TableLog {
  title: string;
  entries: LogEntry[];
}

export interface Stats {
  rows: number;
  filled: number;
  already: number;
  warnings: number;
  notFound: number;
}

export interface TabelResult {
  kind: TabelKind;
  general: LogEntry[];
  tables: TableLog[];
  stats: Stats;
  output: Uint8Array | null;
}

interface HeaderInfo {
  row: number;
  mark: number;
  client: number;
  number: number;
  wh: number | null;
  whViaComments: boolean;
  whHeaderText: string;
  awb: number | null;
  supplier: number | null;
  minCol: number;
  maxCol: number;
  hasAwb: boolean;
  hasSupplier: boolean;
}

function readHeader(sh: Sheet, r: number): HeaderInfo | null {
  const cols: [string, number][] = [];
  for (let c = 1; c <= sh.maxCol; c++) {
    const h = normalizeHeader(sh.text(c, r));
    if (h) cols.push([h, c]);
  }
  const find = (pred: (h: string) => boolean) => cols.find(([h]) => pred(h))?.[1] ?? null;
  const mark = find((h) => h === "маркировка");
  const client = find((h) => h === "клиент");
  const number = find((h) => h.startsWith("№") && h.includes("клиент")) ?? find((h) => h === "номер клиента");
  if (!mark || !client || !number) return null;

  let wh = find((h) => h.startsWith("склад"));
  let whViaComments = false;
  let whHeaderText = "";
  if (!wh) {
    const comments = find((h) => h.startsWith("комментари"));
    if (comments) {
      wh = comments + 1;
      whViaComments = true;
      whHeaderText = sh.text(wh, r).trim();
    }
  }
  const awb = find((h) => h.includes("awb"));
  const supplier = find((h) => h.startsWith("поставщик"));
  const used = [mark, client, number, ...(wh ? [wh] : [])];
  const colNums = cols.map(([, c]) => c);
  return {
    row: r, mark, client, number, wh, whViaComments, whHeaderText, awb, supplier,
    minCol: Math.min(...used, ...colNums), maxCol: Math.max(...used, ...colNums),
    hasAwb: !!awb, hasSupplier: !!supplier,
  };
}

function findHeaders(sh: Sheet): HeaderInfo[] {
  const out: HeaderInfo[] = [];
  for (let r = 1; r <= sh.maxRow; r++) {
    const h = readHeader(sh, r);
    if (h) out.push(h);
  }
  return out;
}

/** Строки данных таблицы: от заголовка до строки итогов (формула при пустых ключевых ячейках) или следующего заголовка. */
function dataRows(sh: Sheet, h: HeaderInfo, nextHeaderRow: number): number[] {
  const rows: number[] = [];
  const end = Math.min(nextHeaderRow - 1, sh.maxRow);
  for (let r = h.row + 1; r <= end; r++) {
    const keyEmpty = !sh.text(h.mark, r).trim() && !sh.text(h.client, r).trim() && !sh.text(h.number, r).trim();
    if (keyEmpty) {
      let formula = false;
      for (let c = h.minCol; c <= h.maxCol && !formula; c++) formula = sh.hasFormula(c, r);
      if (formula) break; // строка итогов — конец таблицы
      continue;
    }
    rows.push(r);
  }
  return rows;
}

/** Склады, которых больше нет, и чем их заменять при вставке. */
const WAREHOUSE_REPLACE: Record<string, string> = { можайск: "Химки" };

/**
 * Склад для вставки: если в справочнике указано несколько — берётся первый;
 * упразднённые склады заменяются (Можайск → Химки).
 */
export function resolveWarehouse(raw: string): { value: string; notes: string[] } {
  const notes: string[] = [];
  let value = raw.trim();
  if (isMultiWarehouse(value)) {
    const first = value.split(/[,;\n/]|\sи\s/i).map((s) => s.trim()).find(Boolean) ?? value;
    notes.push(`в справочнике несколько складов «${value.replace(/\s*\n\s*/g, ", ")}» — взят первый`);
    value = first;
  }
  const repl = WAREHOUSE_REPLACE[value.toLowerCase()];
  if (repl) {
    notes.push(`${value} → ${repl}`);
    value = repl;
  }
  return { value, notes };
}

const asNumberIfDigits =(s: string): string | number => (/^\d+$/.test(s) ? Number(s) : s);

function describe(e: ClientEntry): string {
  return `№${e.number || "—"} «${e.name || "без имени"}»${e.warehouse ? `, ${e.warehouse}` : ""} (стр. ${e.row})`;
}

function sameNumber(a: string, b: string): boolean {
  return a.replace(/\s+/g, "") === b.replace(/\s+/g, "");
}

interface RowCtx {
  sh: Sheet;
  h: HeaderInfo;
  r: number;
  ref: Reference;
  ix: RefSheetIndex;
  log: LogEntry[];
  stats: Stats;
  label: string;
}

function styleSource(sh: Sheet, h: HeaderInfo, col: number, r: number) {
  // ближайшая ячейка того же столбца внутри таблицы, иначе ячейка маркировки этой строки
  if (r - 1 > h.row && sh.exists(col, r - 1)) return { col, row: r - 1 };
  if (sh.exists(col, r + 1)) return { col, row: r + 1 };
  return { col: h.mark, row: r };
}

function processRow(ctx: RowCtx) {
  const { sh, h, r, ix, log, stats, label } = ctx;
  if (sh.hasFormula(h.mark, r)) return;
  const marking = sh.text(h.mark, r).trim();
  if (!marking) {
    log.push({ level: "warn", text: `${label}: пустая маркировка — пропущено` });
    stats.warnings++;
    return;
  }
  stats.rows++;
  const curNum = sh.text(h.number, r).trim();
  const curName = sh.text(h.client, r).trim();
  const curWh = h.wh ? sh.text(h.wh, r).trim() : "";
  const tag = `${label}, маркировка «${marking}»`;

  const res = lookup(ix, marking);

  if (res.status === "notfound") {
    stats.notFound++;
    log.push({ level: "error", text: `${tag}: не найдена в справочнике (лист ${ix.sheetName})` });
    return;
  }
  if (res.status === "ambiguous") {
    stats.warnings++;
    log.push({
      level: "warn",
      text: `${tag}: маркировка есть у нескольких клиентов — ничего не вставлено. Варианты: ${res.entries.map(describe).join("; ")}`,
    });
    return;
  }
  if (res.status === "partial") {
    stats.warnings++;
    log.push({
      level: "warn",
      text: `${tag}: точного совпадения нет, есть похожее (часть ячейки с несколькими маркировками или запись с префиксом в справочнике) — проверь вручную, ничего не вставлено. Кандидаты: ${res.entries.map(describe).join("; ")}`,
    });
    return;
  }

  const e = res.entry;
  const notes: string[] = [];
  if (res.viaPrefix) notes.push(`найдено без префикса «${res.viaPrefix}»`);

  if (!e.number) {
    stats.warnings++;
    log.push({ level: "warn", text: `${tag}: найден клиент ${describe(e)}, но в справочнике у него нет номера — ничего не вставлено` });
    return;
  }

  // Номер уже стоит — сверяем
  if (curNum && !sameNumber(curNum, e.number)) {
    stats.warnings++;
    log.push({
      level: "error",
      text: `${tag}: в табеле стоит №${curNum}, а по справочнику это ${describe(e)} — строка не изменена`,
    });
    return;
  }

  // Имя уже вписано — сверяем с номером
  if (curName && !curNum) {
    const known = ctx.ref.namesByNumber.get(e.number) ?? [];
    const candidates = e.name ? [e.name, ...known] : known;
    if (!candidates.some((n) => namesCompatible(curName, n))) {
      stats.warnings++;
      log.push({
        level: "error",
        text: `${tag}: в табеле клиент «${curName}», а найденный номер ${e.number} принадлежит «${candidates.join(" / ") || "?"}» — не соотносится, строка не изменена`,
      });
      return;
    }
  }

  const written: string[] = [];
  if (!curNum) {
    sh.set(h.number, r, asNumberIfDigits(e.number), styleSource(sh, h, h.number, r));
    written.push(`№ ${e.number}`);
  }
  if (!curName) {
    if (e.name) {
      sh.set(h.client, r, e.name, styleSource(sh, h, h.client, r));
      written.push(`клиент «${e.name}»`);
    } else if (res.nameConflict.length) {
      notes.push(`имя не вставлено — у номера разные имена в справочнике: ${res.nameConflict.join(" / ")}`);
    }
  }
  if (h.wh && !curWh) {
    if (!e.warehouse) {
      notes.push(
        res.whConflict.length
          ? `склад не вставлен — в справочнике разные склады: ${res.whConflict.join(" / ")}`
          : "склад не вставлен — в справочнике он не указан",
      );
    } else {
      const { value, notes: whNotes } = resolveWarehouse(e.warehouse);
      sh.set(h.wh, r, value, styleSource(sh, h, h.wh, r));
      written.push(`склад «${value}»`);
      notes.push(...whNotes);
    }
  }

  const hasProblems = notes.some((n) => n.includes("не вставлен"));
  if (written.length) {
    stats.filled++;
    log.push({
      level: hasProblems ? "warn" : "filled",
      text: `${tag} → ${written.join(", ")}${notes.length ? ` (${notes.join("; ")})` : ""}`,
    });
    if (hasProblems) stats.warnings++;
  } else {
    stats.already++;
    log.push({
      level: hasProblems ? "warn" : "ok",
      text: `${tag}: уже заполнено, совпадает со справочником (№${e.number})${notes.length ? ` (${notes.join("; ")})` : ""}`,
    });
    if (hasProblems) stats.warnings++;
  }
}

function pickSheet(wb: Workbook): Sheet {
  // В табеле один рабочий лист; если их несколько — берём первый, где есть таблица с «Маркировка»
  for (const n of wb.sheetNames) {
    const sh = wb.sheet(n);
    if (findHeaders(sh).length) return sh;
  }
  throw new Error("Не найдена ни одна таблица со столбцами «№ Клиента», «Клиент», «Маркировка»");
}

export function fillTabel(wb: Workbook, kind: TabelKind, ref: Reference): TabelResult {
  const general: LogEntry[] = [];
  const tables: TableLog[] = [];
  const stats: Stats = { rows: 0, filled: 0, already: 0, warnings: 0, notFound: 0 };

  const ix = kind === "import" ? ref.import : ref.holland;
  if (!ix) {
    general.push({ level: "error", text: `В справочнике нет листа ${kind === "import" ? "ИМПОРТ" : "ГОЛЛАНДИЯ"}` });
    return { kind, general, tables, stats, output: null };
  }

  const sh = pickSheet(wb);
  const headers = findHeaders(sh);

  // Защита от перепутанных файлов
  const looksHolland = headers.some((h) => h.hasSupplier) && !headers.some((h) => h.hasAwb);
  const looksImport = headers.some((h) => h.hasAwb);
  if (kind === "import" && looksHolland) {
    general.push({ level: "warn", text: "Похоже, это голландский табель (есть «Поставщик», нет «AWB») — проверь, в то ли окошко загружен файл" });
  }
  if (kind === "holland" && looksImport) {
    general.push({ level: "warn", text: "Похоже, это табель импорта (есть столбец AWB) — проверь, в то ли окошко загружен файл" });
  }

  // В импорте первая (верхняя) таблица заполняется вручную — пропускаем её
  const toProcess = kind === "import" ? headers.slice(1) : headers;
  if (kind === "import") {
    general.push({ level: "info", text: `Найдено таблиц: ${headers.length}. Верхняя (стр. ${headers[0].row}) — пропущена, обрабатываю нижних: ${toProcess.length}` });
  } else {
    general.push({ level: "info", text: `Найдено таблиц: ${headers.length}` });
  }

  toProcess.forEach((h, i) => {
    const next = headers[headers.indexOf(h) + 1]?.row ?? sh.maxRow + 1;
    const rows = dataRows(sh, h, next);
    const log: LogEntry[] = [];
    let title: string;
    if (kind === "import") {
      const awbs = [...new Set(rows.map((r) => (h.awb ? sh.text(h.awb, r).trim() : "")).filter(Boolean))];
      title = `Разбивка ${i + 1} (стр. ${h.row}) — AWB ${awbs.join(", ") || "не указан"}`;
    } else {
      const zone = sh.text(3, h.row + 1).trim(); // «Холод» / «Тепло» в столбце C
      title = `Таблица ${i + 1} (стр. ${h.row})${zone ? ` — ${zone}` : ""}`;
    }
    if (h.whViaComments) {
      log.push({
        level: "info",
        text: `Столбец склада подписан как «${h.whHeaderText || "пусто"}» — пишу склад в столбец после «Комментарии» (${addr(h.wh!, h.row).replace(/\d+$/, "")})`,
      });
    } else if (!h.wh) {
      log.push({ level: "warn", text: "Не найден столбец «Склад» — склад заполняться не будет" });
    }
    if (!rows.length) {
      log.push({ level: "info", text: "Таблица пустая" });
    }
    for (const r of rows) {
      const parts = [`стр. ${r}`];
      if (kind === "holland" && h.supplier) {
        const s = sh.text(h.supplier, r).trim();
        if (s) parts.push(s);
      }
      processRow({ sh, h, r, ref, ix, log, stats, label: parts.join(", ") });
    }
    tables.push({ title, entries: log });
  });

  return { kind, general, tables, stats, output: sh.dirty ? wb.save() : null };
}

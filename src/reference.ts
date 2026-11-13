/**
 * Справочник «Клиенты и маркировки».
 *
 * Структура листов ГОЛЛАНДИЯ / ИМПОРТ:
 *   строка заголовков: «№» | «Клиент» | «Маркировки» (несколько столбцов по видам) | … | «Склад»
 *   далее блоки клиентов: первая строка блока содержит № и/или имя клиента,
 *   следующие строки без № и имени — дополнительные маркировки того же клиента.
 */
import type { Sheet, Workbook } from "./xlsx/workbook";
import { markingCandidates, markingKey, normalizeHeader, splitMarkingCell, namesCompatible } from "./normalize";

export interface ClientEntry {
  number: string;
  name: string;
  warehouse: string;
  sheet: string;
  row: number;
}

export interface RefSheetIndex {
  sheetName: string;
  clients: ClientEntry[];
  exact: Map<string, ClientEntry[]>;
  weak: Map<string, ClientEntry[]>;
  markingCount: number;
}

export type LookupResult =
  | { status: "found"; entry: ClientEntry; key: string; viaPrefix: string | null; nameConflict: string[]; whConflict: string[] }
  | { status: "ambiguous"; entries: ClientEntry[]; key: string }
  | { status: "partial"; entries: ClientEntry[]; key: string }
  | { status: "notfound" };

function push(map: Map<string, ClientEntry[]>, key: string, e: ClientEntry) {
  const arr = map.get(key);
  if (!arr) map.set(key, [e]);
  else if (!arr.includes(e)) arr.push(e);
}

function findHeader(sh: Sheet) {
  for (let r = 1; r <= Math.min(10, sh.maxRow); r++) {
    const cols = new Map<string, number>();
    for (let c = 1; c <= sh.maxCol; c++) {
      const h = normalizeHeader(sh.text(c, r));
      if (h) cols.set(h, c);
    }
    const num = [...cols].find(([h]) => h.startsWith("№"))?.[1];
    const name = cols.get("клиент");
    const mark = [...cols].find(([h]) => h.startsWith("маркировк"))?.[1];
    if (num && name && mark) {
      const wh = [...cols].find(([h]) => h.startsWith("склад"))?.[1] ?? null;
      const nextHeader = [...cols.values()].filter((c) => c > mark).sort((a, b) => a - b)[0];
      const markEnd = nextHeader ? nextHeader - 1 : mark;
      return { row: r, num, name, markStart: mark, markEnd, wh };
    }
  }
  return null;
}

export function indexReferenceSheet(sh: Sheet): RefSheetIndex {
  const hdr = findHeader(sh);
  if (!hdr) throw new Error(`На листе «${sh.name}» не найдена строка заголовков (№ / Клиент / Маркировки)`);

  const idx: RefSheetIndex = { sheetName: sh.name, clients: [], exact: new Map(), weak: new Map(), markingCount: 0 };
  let cur: ClientEntry | null = null;

  for (let r = hdr.row + 1; r <= sh.maxRow; r++) {
    const num = sh.text(hdr.num, r).trim();
    const name = sh.text(hdr.name, r).trim();
    const wh = hdr.wh ? sh.text(hdr.wh, r).trim() : "";
    if (num || name) {
      cur = { number: num, name, warehouse: wh, sheet: sh.name, row: r };
      idx.clients.push(cur);
    } else if (cur && !cur.warehouse && wh) {
      cur.warehouse = wh;
    }
    if (!cur) continue; // подзаголовки («Цветущее», «Горшки», «Зелень»)

    for (let c = hdr.markStart; c <= hdr.markEnd; c++) {
      if (sh.hasFormula(c, r)) continue;
      const cell = sh.text(c, r).trim();
      if (!cell) continue;
      idx.markingCount++;
      const { exact, weak } = splitMarkingCell(cell);
      for (const k of exact) push(idx.exact, k, cur);
      for (const k of weak) push(idx.weak, k, cur);
    }
  }
  return idx;
}

/** Поиск листа по названию (без учёта регистра и пробелов). */
export function findSheetName(wb: Workbook, needle: string): string | null {
  const n = needle.toLowerCase();
  return wb.sheetNames.find((s) => s.toLowerCase().replace(/\s+/g, "").includes(n)) ?? null;
}

export interface Reference {
  holland: RefSheetIndex | null;
  import: RefSheetIndex | null;
  /** Все имена, под которыми номер встречается в справочнике (для сверки уже вписанных имён). */
  namesByNumber: Map<string, string[]>;
}

export function buildReference(wb: Workbook): Reference {
  const hName = findSheetName(wb, "голланд");
  const iName = findSheetName(wb, "импорт");
  if (!hName && !iName) {
    throw new Error("В справочнике не найдены листы «ГОЛЛАНДИЯ» и «ИМПОРТ»");
  }
  const holland = hName ? indexReferenceSheet(wb.sheet(hName)) : null;
  const imp = iName ? indexReferenceSheet(wb.sheet(iName)) : null;
  const namesByNumber = new Map<string, string[]>();
  for (const ix of [holland, imp]) {
    for (const c of ix?.clients ?? []) {
      if (!c.number || !c.name) continue;
      const arr = namesByNumber.get(c.number) ?? [];
      if (!arr.includes(c.name)) arr.push(c.name);
      namesByNumber.set(c.number, arr);
    }
  }
  return { holland, import: imp, namesByNumber };
}

function uniq<T>(arr: T[]): T[] {
  return [...new Set(arr)];
}

/** Найти клиента по маркировке из табеля. */
export function lookup(ix: RefSheetIndex, rawMarking: string): LookupResult {
  const cands = markingCandidates(rawMarking);
  for (const { key, stripped } of cands) {
    const hits = ix.exact.get(key);
    if (!hits?.length) continue;
    const numbers = uniq(hits.map((h) => h.number));
    if (numbers.length > 1) return { status: "ambiguous", entries: hits, key };
    // Один и тот же номер (клиент встречается в справочнике несколькими блоками)
    const names = uniq(hits.map((h) => h.name).filter(Boolean));
    const whs = uniq(hits.map((h) => h.warehouse).filter(Boolean));
    const nameConflict = names.length > 1 && !names.every((n) => namesCompatible(n, names[0])) ? names : [];
    const entry: ClientEntry = {
      ...hits[0],
      name: names.length === 1 ? names[0] : nameConflict.length ? "" : names[0] ?? "",
      warehouse: whs.length === 1 ? whs[0] : "",
    };
    return { status: "found", entry, key, viaPrefix: stripped, nameConflict, whConflict: whs.length > 1 ? whs : [] };
  }
  for (const { key } of cands) {
    const hits = ix.weak.get(key);
    if (hits?.length) return { status: "partial", entries: hits, key };
  }
  return { status: "notfound" };
}

export { markingKey };

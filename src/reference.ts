/**
 * Справочник «Клиенты и маркировки».
 *
 * Листы ГОЛЛАНДИЯ / ИМПОРТ («блоки клиентов»):
 *   строка заголовков: «№» | «Клиент» | «Маркировки» (несколько столбцов по видам) | … | «Склад»
 *   первая строка блока содержит № и/или имя клиента, следующие строки без № и имени —
 *   дополнительные маркировки того же клиента.
 *
 * Лист Jaaz: подтаблицы «Маркировка | Назначение» (маркировки самого Джаза, № 85,
 *   клиент = назначение, напр. «ДЖАЗ ПИТЕР») и «Маркировка | Клиент | Номер клиента»
 *   (наши клиенты, которые едут через Джаз). Красная заливка → склад Химки, иначе Котельники.
 *
 * Лист MF: правая таблица «Маркировка | Направление | Комментарий | №».
 *   Клиент = комментарий; для № 102 к нему добавляется «МФ » («МСК» → «МФ МСК»).
 *   Зелёная заливка маркировки → склад Химки, иначе Котельники.
 */
import { isGreenColor, isRedColor, type Sheet, type Workbook } from "./xlsx/workbook";
import {
  markingCandidates, markingKey, normalizeHeader, splitMarkingCell, mfStem, MF_SHORT_LEN, markingTokens,
} from "./normalize";

export type RefKind = "holland" | "import" | "jaaz" | "mf";

export const REF_TITLES: Record<RefKind, string> = {
  holland: "ГОЛЛАНДИЯ",
  import: "ИМПОРТ",
  jaaz: "Джаз",
  mf: "МФ",
};

export interface ClientEntry {
  number: string;
  name: string;
  warehouse: string;
  sheet: string;
  row: number;
}

export interface RefSheetIndex {
  kind: RefKind;
  sheetName: string;
  clients: ClientEntry[];
  exact: Map<string, ClientEntry[]>;
  weak: Map<string, ClientEntry[]>;
  /** МФ: основа маркировки без «MF-» и полная основа → записи */
  stem: Map<string, ClientEntry[]>;
  markingCount: number;
}

export type LookupResult =
  | {
      status: "found";
      entry: ClientEntry;
      key: string;
      /** Как нашлось, если не прямым совпадением (для консоли) */
      via: string | null;
      /** Совпадение не точное — заполнено, но стоит проверить */
      uncertain: boolean;
      nameConflict: string[];
      whConflict: string[];
      ix: RefSheetIndex;
    }
  | { status: "ambiguous"; entries: ClientEntry[]; key: string; via: string | null }
  | { status: "partial"; entries: ClientEntry[]; key: string }
  | { status: "notfound" };

function push(map: Map<string, ClientEntry[]>, key: string, e: ClientEntry) {
  if (!key) return;
  const arr = map.get(key);
  if (!arr) map.set(key, [e]);
  else if (!arr.includes(e)) arr.push(e);
}

function emptyIndex(kind: RefKind, sheetName: string): RefSheetIndex {
  return { kind, sheetName, clients: [], exact: new Map(), weak: new Map(), stem: new Map(), markingCount: 0 };
}

function addMarkingCell(idx: RefSheetIndex, cell: string, e: ClientEntry) {
  idx.markingCount++;
  const { exact, weak } = splitMarkingCell(cell);
  for (const k of exact) push(idx.exact, k, e);
  for (const k of weak) push(idx.weak, k, e);
  if (idx.kind === "mf") {
    // В МФ-маркировках пробелов не бывает, поэтому в ячейке разделители ещё и пробелы/табы/кавычки:
    // «MF-SEVILLA  MF-ZARAGOZA\tGR-GARDENA», «MF-MAKSBEROK "MRJ238»
    for (const part of cell.split(/[,;\n\r/\s"«»]+/)) {
      const k = markingKey(part);
      if (!k) continue;
      push(idx.exact, k, e);
      const st = mfStem(part);
      if (st) push(idx.stem, st, e);
    }
  }
}

function headerCells(sh: Sheet, r: number): [string, number][] {
  const out: [string, number][] = [];
  for (let c = 1; c <= sh.maxCol; c++) {
    const h = normalizeHeader(sh.text(c, r));
    if (h) out.push([h, c]);
  }
  return out;
}

// ---------- ГОЛЛАНДИЯ / ИМПОРТ ----------

function findBlocksHeader(sh: Sheet) {
  for (let r = 1; r <= Math.min(15, sh.maxRow); r++) {
    const cols = headerCells(sh, r);
    const num = cols.find(([h]) => h.startsWith("№"))?.[1];
    const name = cols.find(([h]) => h === "клиент")?.[1];
    const mark = cols.find(([h]) => h.startsWith("маркировк"))?.[1];
    if (num && name && mark) {
      const wh = cols.find(([h]) => h.startsWith("склад"))?.[1] ?? null;
      const nextHeader = cols.map(([, c]) => c).filter((c) => c > mark).sort((a, b) => a - b)[0];
      const markEnd = nextHeader ? nextHeader - 1 : mark;
      return { row: r, num, name, markStart: mark, markEnd, wh };
    }
  }
  return null;
}

export function indexBlocksSheet(sh: Sheet, kind: "holland" | "import"): RefSheetIndex {
  const hdr = findBlocksHeader(sh);
  if (!hdr) throw new Error(`На листе «${sh.name}» не найдена строка заголовков (№ / Клиент / Маркировки)`);
  const idx = emptyIndex(kind, sh.name);
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
      if (cell) addMarkingCell(idx, cell, cur);
    }
  }
  return idx;
}

// ---------- Джаз ----------

export const JAAZ_NUMBER = "85";

export function indexJaazSheet(sh: Sheet): RefSheetIndex {
  // Строка заголовков: несколько «Маркировка», за каждой — «Назначение» или «Клиент» + «Номер клиента»
  let hdrRow = 0;
  let cols: [string, number][] = [];
  for (let r = 1; r <= Math.min(30, sh.maxRow); r++) {
    const hc = headerCells(sh, r);
    if (hc.filter(([h]) => h === "маркировка").length >= 2) {
      hdrRow = r;
      cols = hc;
      break;
    }
  }
  if (!hdrRow) throw new Error(`На листе «${sh.name}» не найдена строка заголовков с «Маркировка»`);

  type Sub = { mark: number; dest: number | null; client: number | null; number: number | null };
  const subs: Sub[] = [];
  for (const [h, c] of cols) {
    if (h !== "маркировка") continue;
    const after = cols.filter(([, cc]) => cc > c);
    const nextMark = after.find(([hh]) => hh === "маркировка")?.[1] ?? Infinity;
    const inSub = after.filter(([, cc]) => cc < nextMark);
    subs.push({
      mark: c,
      dest: inSub.find(([hh]) => hh.startsWith("назначени"))?.[1] ?? null,
      client: inSub.find(([hh]) => hh === "клиент")?.[1] ?? null,
      number: inSub.find(([hh]) => hh.startsWith("номер") || hh.startsWith("№"))?.[1] ?? null,
    });
  }

  const idx = emptyIndex("jaaz", sh.name);
  for (const sub of subs) {
    for (let r = hdrRow + 1; r <= sh.maxRow; r++) {
      const cell = sh.text(sub.mark, r).trim();
      if (!cell || sh.hasFormula(sub.mark, r)) continue;
      const nameCol = sub.client ?? sub.dest;
      const red = isRedColor(sh.fillColor(sub.mark, r)) || (nameCol ? isRedColor(sh.fillColor(nameCol, r)) : false);
      let entry: ClientEntry;
      if (sub.client) {
        entry = {
          number: sub.number ? sh.text(sub.number, r).trim() : "",
          name: sh.text(sub.client, r).trim(),
          warehouse: red ? "Химки" : "Котельники",
          sheet: sh.name,
          row: r,
        };
      } else {
        entry = {
          number: JAAZ_NUMBER,
          name: (sub.dest ? sh.text(sub.dest, r).trim() : "") || "ДЖАЗ",
          warehouse: red ? "Химки" : "Котельники",
          sheet: sh.name,
          row: r,
        };
      }
      idx.clients.push(entry);
      addMarkingCell(idx, cell, entry);
    }
  }
  return idx;
}

// ---------- МФ ----------

export const MF_NUMBER = "102";

export function mfClientName(comment: string, number: string): string {
  const c = comment.trim();
  if (number !== MF_NUMBER) return c;
  // «МСК» — это сам МФ (Москва): пишется просто «МФ»
  if (!c || /^(мск|москва)$/i.test(c)) return "МФ";
  if (/^мф(\s|$)/i.test(c)) return c.toUpperCase();
  return `МФ ${c.toUpperCase()}`;
}

export function indexMfSheet(sh: Sheet): RefSheetIndex {
  // Правая таблица: «Маркировка | Направление | Комментарий | №»
  let hdr: { row: number; mark: number; comment: number; number: number } | null = null;
  for (let r = 1; r <= Math.min(40, sh.maxRow) && !hdr; r++) {
    const hc = headerCells(sh, r);
    const comment = hc.find(([h]) => h.startsWith("комментари"))?.[1];
    if (!comment) continue;
    const mark = hc.filter(([h, c]) => h === "маркировка" && c < comment).map(([, c]) => c).pop();
    const number = hc.find(([h, c]) => (h === "№" || h.startsWith("номер")) && c > comment)?.[1];
    if (mark && number) hdr = { row: r, mark, comment, number };
  }
  if (!hdr) throw new Error(`На листе «${sh.name}» не найдена таблица «Маркировка / Направление / Комментарий / №»`);

  const idx = emptyIndex("mf", sh.name);
  for (let r = hdr.row + 1; r <= sh.maxRow; r++) {
    const cell = sh.text(hdr.mark, r).trim();
    if (!cell || sh.hasFormula(hdr.mark, r)) continue;
    const number = sh.text(hdr.number, r).trim();
    const entry: ClientEntry = {
      number,
      name: mfClientName(sh.text(hdr.comment, r), number),
      warehouse: isGreenColor(sh.fillColor(hdr.mark, r)) ? "Химки" : "Котельники",
      sheet: sh.name,
      row: r,
    };
    idx.clients.push(entry);
    addMarkingCell(idx, cell, entry);
  }
  return idx;
}

// ---------- Справочник целиком ----------

/** Поиск листа по названию (без учёта регистра и пробелов). */
export function findSheetName(wb: Workbook, ...needles: string[]): string | null {
  return (
    wb.sheetNames.find((s) => {
      const n = s.toLowerCase().replace(/\s+/g, "");
      return needles.some((x) => n.includes(x));
    }) ?? null
  );
}

export interface Reference {
  sheets: Partial<Record<RefKind, RefSheetIndex>>;
  errors: string[];
  /** Все имена, под которыми номер встречается в справочнике (для сверки уже вписанных имён). */
  namesByNumber: Map<string, string[]>;
}

export function buildReference(wb: Workbook): Reference {
  const ref: Reference = { sheets: {}, errors: [], namesByNumber: new Map() };
  const specs: [RefKind, string[], (sh: Sheet) => RefSheetIndex][] = [
    ["holland", ["голланд"], (sh) => indexBlocksSheet(sh, "holland")],
    ["import", ["импорт"], (sh) => indexBlocksSheet(sh, "import")],
    ["jaaz", ["jaaz", "джаз"], indexJaazSheet],
    ["mf", ["mf-", "mf", "мф"], indexMfSheet],
  ];
  for (const [kind, needles, build] of specs) {
    const name = findSheetName(wb, ...needles);
    if (!name) {
      ref.errors.push(`Не найден лист ${REF_TITLES[kind]}`);
      continue;
    }
    try {
      ref.sheets[kind] = build(wb.sheet(name));
    } catch (e) {
      ref.errors.push((e as Error).message);
    }
  }
  if (!ref.sheets.holland && !ref.sheets.import) {
    throw new Error(`В справочнике не найдены листы «ГОЛЛАНДИЯ» и «ИМПОРТ». ${ref.errors.join("; ")}`);
  }
  for (const ix of Object.values(ref.sheets)) {
    for (const c of ix?.clients ?? []) {
      if (!c.number || !c.name) continue;
      const arr = ref.namesByNumber.get(c.number) ?? [];
      if (!arr.includes(c.name)) arr.push(c.name);
      ref.namesByNumber.set(c.number, arr);
    }
  }
  return ref;
}

// ---------- Поиск ----------

const uniq = <T,>(arr: T[]): T[] => [...new Set(arr)];
const normName = (s: string) => s.toLowerCase().replace(/ё/g, "е").replace(/[^\p{L}\p{N}]+/gu, " ").trim();

function resolveHits(
  ix: RefSheetIndex, hits: ClientEntry[], key: string, via: string | null, uncertain: boolean,
): LookupResult {
  const numbers = uniq(hits.map((h) => h.number));
  if (numbers.length > 1) return { status: "ambiguous", entries: hits, key, via };
  const names = uniq(hits.map((h) => h.name).filter(Boolean));
  const distinctNames = uniq(names.map(normName));
  const whs = uniq(hits.map((h) => h.warehouse).filter(Boolean));
  const nameConflict = distinctNames.length > 1 ? names : [];
  // Неточное совпадение без однозначного клиента — не угадываем
  if (uncertain && nameConflict.length) return { status: "ambiguous", entries: hits, key, via };
  const entry: ClientEntry = {
    ...hits[0],
    name: nameConflict.length ? "" : names[0] ?? "",
    warehouse: whs.length === 1 ? whs[0] : "",
  };
  return { status: "found", entry, key, via, uncertain, nameConflict, whConflict: whs.length > 1 ? whs : [], ix };
}

/** Найти клиента по маркировке из табеля на одном листе справочника. */
export interface LookupOptions {
  /** Джаз: искать по отдельным словам составной маркировки (только голландские строки Jaaz) */
  jaazWords?: boolean;
}

export function lookup(ix: RefSheetIndex, rawMarking: string, opts: LookupOptions = {}): LookupResult {
  const cands = markingCandidates(rawMarking);

  // 1. Точное совпадение (с последовательным отбрасыванием префиксов b-/QW/FLT)
  for (const { key, stripped } of cands) {
    const hits = ix.exact.get(key);
    if (hits?.length) return resolveHits(ix, hits, key, stripped ? `без префикса «${stripped}»` : null, false);
  }

  // 2. МФ: без «MF-», затем по первым 6 символам (так маркировки пишет Waibel)
  if (ix.kind === "mf") {
    const st = mfStem(rawMarking);
    if (st) {
      const hits = ix.stem.get(st);
      if (hits?.length) return resolveHits(ix, hits, st, "без «MF-»", false);
      // Waibel обрезает маркировку до ~6 символов («BIGBUK» = MF-BIGBUKETNSK, «AGAVA» = MF-AGAVA24)
      if (st.length >= 4 && st.length <= MF_SHORT_LEN) {
        const hits = [...ix.stem].filter(([k]) => k.startsWith(st)).flatMap(([, v]) => v);
        if (hits.length) return resolveHits(ix, uniq(hits), st, `по началу маркировки («${st.toUpperCase()}…»)`, false);
      }
    }
  }

  // 3. Джаз: составные маркировки «ферма + код клиента» («ATS KULESH», «PARMA MFTC»).
  //    Код клиента обычно последнее слово, поэтому слова проверяются справа налево:
  //    берётся последнее слово, для которого есть запись в справочнике.
  if (ix.kind === "jaaz" && opts.jaazWords) {
    const tokens = markingTokens(rawMarking);
    // в т.ч. «APOLO PL» → «APOLO» (служебное «PL» в словах не учитывается)
    if (tokens.length > 1 || (tokens.length === 1 && tokens[0] !== markingKey(rawMarking))) {
      for (const t of [...tokens].reverse()) {
        const hits = ix.exact.get(t);
        if (hits?.length) return resolveHits(ix, hits, t, `по слову «${t.toUpperCase()}»`, true);
      }
    }
    // Маркировка табеля — часть записи справочника («BATHURST» ↔ «BATHURST DON»)
    for (const t of [markingKey(rawMarking), ...[...tokens].reverse()]) {
      const hits = ix.weak.get(t);
      if (hits?.length) return resolveHits(ix, hits, t, `по части записи справочника («${t.toUpperCase()}»)`, true);
    }
  }

  // 4. Похожее (часть ячейки с несколькими маркировками / запись с префиксом) — только подсказка
  for (const { key } of cands) {
    const hits = ix.weak.get(key);
    if (hits?.length) return { status: "partial", entries: hits, key };
  }
  return { status: "notfound" };
}

/**
 * Поиск сразу на нескольких листах. Если нашлось на разных листах у клиентов с разными
 * номерами — неоднозначно. Иначе берётся первый лист по порядку.
 */
export function lookupMany(ixs: RefSheetIndex[], rawMarking: string, opts: LookupOptions = {}): LookupResult {
  const results = ixs.map((ix) => ({ ix, res: lookup(ix, rawMarking, opts) }));
  const amb = results.find((r) => r.res.status === "ambiguous");
  if (amb) return amb.res;
  const found = results.filter((r) => r.res.status === "found");
  if (found.length) {
    // Приоритет: точное совпадение → с отброшенным префиксом / по началу → по слову (неточное)
    const rank = (x: LookupResult) => (x.status !== "found" ? 9 : x.uncertain ? 2 : x.via ? 1 : 0);
    const best = Math.min(...found.map((r) => rank(r.res)));
    const pool = found.filter((r) => rank(r.res) === best);
    const numbers = uniq(pool.map((r) => (r.res.status === "found" ? r.res.entry.number : "")));
    if (numbers.length > 1) {
      const entries = pool.flatMap((r) => (r.res.status === "found" ? [r.res.entry] : []));
      return { status: "ambiguous", entries, key: markingKey(rawMarking), via: "на разных листах" };
    }
    return pool[0].res;
  }
  const partial = results.filter((r) => r.res.status === "partial");
  if (partial.length) {
    return {
      status: "partial",
      entries: partial.flatMap((r) => (r.res.status === "partial" ? r.res.entries : [])),
      key: markingKey(rawMarking),
    };
  }
  return { status: "notfound" };
}

export { markingKey };

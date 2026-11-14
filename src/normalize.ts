/**
 * Нормализация маркировок и имён клиентов.
 */

/** Кириллические буквы, которые выглядят как латинские (частая опечатка в маркировках). */
const HOMOGLYPHS: Record<string, string> = {
  а: "a", в: "b", е: "e", ё: "e", к: "k", м: "m", н: "h", о: "o",
  р: "p", с: "c", т: "t", у: "y", х: "x", і: "i",
};

export function prepare(s: string): string {
  return s
    .toLowerCase()
    .replace(/[‐-―−]/g, "-") // разные виды тире → дефис
    .replace(/[   ]/g, " ") // неразрывные пробелы
    .replace(/[авекмнорстухёі]/g, (ch) => HOMOGLYPHS[ch] ?? ch)
    .replace(/\s+/g, " ")
    .trim();
}

/** Ключ для сравнения маркировок: нижний регистр, без пробелов. */
export function markingKey(s: string): string {
  return prepare(s).replace(/ /g, "");
}

/**
 * Префиксы, которые можно отбросить, если маркировка целиком не нашлась:
 *   «b-» / «b » / «b -»,  «QW » / «QW.»,  «FLT ».
 * Префикс отбрасывается только если за ним идёт разделитель,
 * поэтому «QWERTY» или «FLTX» не пострадают.
 */
const PREFIX_RE = /^(?:b ?- ?|b |qw ?\. ?|qw |flt )(?=\S)/;

/**
 * Варианты маркировки для поиска по порядку: сначала как есть,
 * затем с последовательно отброшенными префиксами.
 */
export function markingCandidates(raw: string): { key: string; stripped: string | null }[] {
  const out: { key: string; stripped: string | null }[] = [];
  let p = prepare(raw);
  const seen = new Set<string>();
  let stripped = "";
  while (p) {
    const key = p.replace(/ /g, "");
    if (!seen.has(key)) {
      seen.add(key);
      out.push({ key, stripped: stripped ? stripped.trim() : null });
    }
    const m = PREFIX_RE.exec(p);
    if (!m) break;
    stripped += m[0];
    p = p.slice(m[0].length);
  }
  return out;
}

/** Разбить ячейку справочника на отдельные маркировки (пока в одной ячейке их бывает несколько). */
export function splitMarkingCell(cell: string): { exact: string[]; weak: string[] } {
  const exact = new Set<string>();
  const weak = new Set<string>();
  const whole = markingKey(cell);
  if (whole) exact.add(whole);
  for (const part of cell.split(/[,;\n\r/]+/)) {
    const k = markingKey(part);
    if (!k) continue;
    exact.add(k);
    // Если в справочнике маркировка записана с префиксом (B-QW.MOG001), вариант без префикса
    // годится только как подсказка — автоматически по нему не заполняем.
    for (const c of markingCandidates(part).slice(1)) weak.add(c.key);
    const tokens = prepare(part).split(" ");
    if (tokens.length > 1) {
      for (const t of tokens) {
        const tk = t.replace(/[()]/g, "");
        if (tk.length >= 3 && tk !== "flt") weak.add(tk);
      }
    }
  }
  for (const k of exact) weak.delete(k);
  return { exact: [...exact], weak: [...weak] };
}

const NAME_STOP = new Set(["ооо", "ип", "зао", "оао", "пао", "конс", "консолидация", "город", "ltd", "llc"]);

function nameTokens(s: string): string[] {
  return s
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim()
    .split(" ")
    .filter((t) => t.length >= 3 && !NAME_STOP.has(t));
}

/**
 * Соотносится ли имя клиента, уже вписанное в табель, с именем из справочника.
 * Совпадение целиком, вложение одного в другое или хотя бы одно общее значимое слово.
 */
export function namesCompatible(a: string, b: string): boolean {
  const ta = nameTokens(a);
  const tb = nameTokens(b);
  if (!ta.length || !tb.length) return false;
  const ja = ta.join(" ");
  const jb = tb.join(" ");
  if (ja === jb || ja.includes(jb) || jb.includes(ja)) return true;
  const setB = new Set(tb);
  return ta.some((t) => setB.has(t));
}

export function normalizeHeader(s: string): string {
  return s.toLowerCase().replace(/\s+/g, " ").trim();
}

/** Похоже ли значение склада на несколько складов сразу («Котельники, Можайск»). */
export function isMultiWarehouse(s: string): boolean {
  return /[,;\n/]|\sи\s/i.test(s.trim());
}

/**
 * Основа МФ-маркировки: без префикса «MF-»/«MF »/«MF_» и без знаков препинания.
 * «MF-IP_JIRKOV» → «ipjirkov». Waibel в голландских табелях пишет первые 6 символов основы («IPJIRK»).
 */
export function mfStem(s: string): string {
  return prepare(s)
    .replace(/^(?:mf|мф)\s*[-_ .]\s*/, "")
    .replace(/[^a-z0-9а-яё]/g, "");
}

export const MF_SHORT_LEN = 6;

/** Начинается ли маркировка с MF-/MRJ (признак клиента МФ). */
export function isMfMarking(s: string): boolean {
  return /^(?:mf|мф)\s*[-_ .]|^mrj\d/.test(prepare(s));
}

/** Слова составной маркировки («ATS KULESH» → [ats, kulesh]); «PL» и короткие не учитываются. */
export function markingTokens(s: string): string[] {
  return prepare(s)
    .split(/[\s,;/]+/)
    .map((t) => t.replace(/[()]/g, ""))
    .filter((t) => t.length >= 3 && t !== "flt" && t !== "pl");
}

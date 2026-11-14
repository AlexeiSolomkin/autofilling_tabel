/**
 * Минимальный «хирургический» редактор xlsx.
 *
 * Файл распаковывается, нужные XML-части разбираются в DOM, меняются только
 * значения конкретных ячеек, остальное (стили, «умные таблицы», формулы,
 * условное форматирование, ширины столбцов и т.д.) остаётся байт-в-байт.
 * Поэтому табель после заполнения выглядит ровно так же, как до.
 */
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import { splitAddr, addr as mkAddr } from "./address";

export type CellValue = string | number | null;

/** Абстракция над DOMParser/XMLSerializer: в браузере — нативные, в тестах — xmldom. */
export interface XmlEnv {
  parse(text: string): Document;
  serialize(doc: Document): string;
}

export const browserXmlEnv = (): XmlEnv => ({
  parse: (t) => new DOMParser().parseFromString(t, "application/xml"),
  serialize: (d) => new XMLSerializer().serializeToString(d),
});

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

function elementChildren(el: Element | Document, localName?: string): Element[] {
  const out: Element[] = [];
  const nodes = el.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    if (n.nodeType === 1 && (!localName || (n as Element).localName === localName)) out.push(n as Element);
  }
  return out;
}

function firstChild(el: Element, localName: string): Element | null {
  return elementChildren(el, localName)[0] ?? null;
}

function descendants(el: Element | Document, localName: string): Element[] {
  const out: Element[] = [];
  const walk = (node: Element | Document) => {
    for (const ch of elementChildren(node)) {
      if (ch.localName === localName) out.push(ch);
      walk(ch);
    }
  };
  walk(el);
  return out;
}

/** Текст строки из sharedStrings / inlineStr (учитывает rich text, игнорирует фонетику rPh). */
function richText(el: Element): string {
  let s = "";
  for (const ch of elementChildren(el)) {
    if (ch.localName === "t") s += ch.textContent ?? "";
    else if (ch.localName === "r") {
      const t = firstChild(ch, "t");
      if (t) s += t.textContent ?? "";
    }
  }
  return s;
}

/** Стандартная палитра Excel для индексированных цветов (самые частые). */
const INDEXED: Record<number, string> = {
  0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF", 5: "FFFF00", 6: "FF00FF", 7: "00FFFF",
  8: "000000", 9: "FFFFFF", 10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF", 15: "00FFFF",
  16: "800000", 17: "008000", 18: "000080", 19: "808000", 20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080",
  42: "CCFFCC", 43: "FFFF99", 50: "99CC00", 51: "FFCC00", 52: "FF9900", 53: "FF6600", 57: "339966",
};

function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const ch = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = ch.map((c) => Math.round(tint < 0 ? c * (1 + tint) : c + (255 - c) * tint));
  return out.map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function resolveColor(el: Element, theme: string[]): string | null {
  const rgb = el.getAttribute("rgb");
  const tint = Number(el.getAttribute("tint") ?? 0);
  if (rgb) return applyTint(rgb.slice(-6).toUpperCase(), tint);
  const th = el.getAttribute("theme");
  if (th !== null && theme[Number(th)]) return applyTint(theme[Number(th)], tint);
  const idx = el.getAttribute("indexed");
  if (idx !== null && INDEXED[Number(idx)]) return applyTint(INDEXED[Number(idx)], tint);
  return null;
}

/** Похож ли цвет на красный / зелёный (для пометок складов в справочнике). */
export function isRedColor(hex: string | null): boolean {
  if (!hex) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return r >= 180 && g <= 110 && b <= 110;
}
export function isGreenColor(hex: string | null): boolean {
  if (!hex) return false;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return g >= 110 && g > r + 35 && g > b + 35;
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

export class Sheet {
  readonly name: string;
  readonly path: string;
  private doc: Document;
  private wb: Workbook;
  private sheetData: Element;
  private rows = new Map<number, Element>();
  private cells = new Map<string, Element>();
  dirty = false;
  maxRow = 0;
  maxCol = 0;

  constructor(wb: Workbook, name: string, path: string, doc: Document) {
    this.wb = wb;
    this.name = name;
    this.path = path;
    this.doc = doc;
    const sd = descendants(doc, "sheetData")[0];
    if (!sd) throw new Error(`Лист «${name}» не содержит данных`);
    this.sheetData = sd;
    for (const row of elementChildren(sd, "row")) {
      const r = Number(row.getAttribute("r"));
      this.rows.set(r, row);
      if (r > this.maxRow) this.maxRow = r;
      for (const c of elementChildren(row, "c")) {
        const a = c.getAttribute("r");
        if (!a) continue;
        this.cells.set(a, c);
        const { col } = splitAddr(a);
        if (col > this.maxCol) this.maxCol = col;
      }
    }
  }

  getXml(): Document {
    return this.doc;
  }

  /** Значение ячейки (кэшированное значение для формул). */
  get(address: string): CellValue {
    const c = this.cells.get(address);
    if (!c) return null;
    const t = c.getAttribute("t");
    if (t === "inlineStr") {
      const is = firstChild(c, "is");
      return is ? richText(is) : null;
    }
    const v = firstChild(c, "v");
    if (!v) return null;
    const raw = v.textContent ?? "";
    if (t === "s") return this.wb.sharedString(Number(raw));
    if (t === "str" || t === "e") return raw;
    if (t === "b") return raw === "1" ? "TRUE" : "FALSE";
    if (raw === "") return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : raw;
  }

  getAt(col: number, row: number): CellValue {
    return this.get(mkAddr(col, row));
  }

  /** Текстовое представление значения (пустая строка, если ячейка пуста). */
  text(col: number, row: number): string {
    const v = this.getAt(col, row);
    return v === null ? "" : String(v);
  }

  /** Цвет заливки ячейки в виде "RRGGBB" (null — без заливки). */
  fillColor(col: number, row: number): string | null {
    const c = this.cells.get(mkAddr(col, row));
    const s = c?.getAttribute("s");
    return s ? this.wb.styleFill(Number(s)) : null;
  }

  /** Очистить значение ячейки, сохранив её стиль. */
  clear(col: number, row: number): void {
    const c = this.cells.get(mkAddr(col, row));
    if (!c) return;
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    this.dirty = true;
  }

  /** Есть ли ячейка в XML (даже пустая, но со стилем). */
  exists(col: number, row: number): boolean {
    return this.cells.has(mkAddr(col, row));
  }

  hasFormula(col: number, row: number): boolean {
    const c = this.cells.get(mkAddr(col, row));
    return !!c && !!firstChild(c, "f");
  }

  /**
   * Записать значение в ячейку. Если ячейки нет в XML — она создаётся,
   * стиль копируется из styleFrom (обычно соседняя ячейка того же столбца).
   */
  set(col: number, row: number, value: string | number, styleFrom?: { col: number; row: number }): void {
    const address = mkAddr(col, row);
    let c = this.cells.get(address);
    const ns = this.sheetData.namespaceURI;
    const prefix = this.sheetData.prefix;
    const make = (local: string) => this.doc.createElementNS(ns, prefix ? `${prefix}:${local}` : local);

    if (!c) {
      const rowEl = this.ensureRow(row);
      c = make("c");
      c.setAttribute("r", address);
      if (styleFrom) {
        const src = this.cells.get(mkAddr(styleFrom.col, styleFrom.row));
        const s = src?.getAttribute("s");
        if (s) c.setAttribute("s", s);
      }
      const after = elementChildren(rowEl, "c").find((x) => splitAddr(x.getAttribute("r")!).col > col);
      rowEl.insertBefore(c, after ?? null);
      rowEl.removeAttribute("spans");
      this.cells.set(address, c);
    }

    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    c.removeAttribute("cm");
    c.removeAttribute("vm");

    if (typeof value === "number") {
      const v = make("v");
      v.textContent = String(value);
      c.appendChild(v);
    } else {
      const idx = this.wb.addSharedString(value);
      if (idx !== null) {
        c.setAttribute("t", "s");
        const v = make("v");
        v.textContent = String(idx);
        c.appendChild(v);
      } else {
        c.setAttribute("t", "inlineStr");
        const is = make("is");
        const t = make("t");
        if (value !== value.trim()) t.setAttributeNS(XML_NS, "xml:space", "preserve");
        t.textContent = value;
        is.appendChild(t);
        c.appendChild(is);
      }
    }
    this.dirty = true;
  }

  private ensureRow(r: number): Element {
    const existing = this.rows.get(r);
    if (existing) return existing;
    const ns = this.sheetData.namespaceURI;
    const prefix = this.sheetData.prefix;
    const row = this.doc.createElementNS(ns, prefix ? `${prefix}:row` : "row");
    row.setAttribute("r", String(r));
    const after = elementChildren(this.sheetData, "row").find((x) => Number(x.getAttribute("r")) > r);
    this.sheetData.insertBefore(row, after ?? null);
    this.rows.set(r, row);
    if (r > this.maxRow) this.maxRow = r;
    return row;
  }
}

export class Workbook {
  private files: Record<string, Uint8Array>;
  private env: XmlEnv;
  private sst: string[] = [];
  private sstIndex = new Map<string, number>();
  private sstDoc: Document | null = null;
  private sstPath: string | null = null;
  private sstDirty = false;
  private sheetsByName = new Map<string, Sheet>();
  private sheetOrder: { name: string; path: string }[] = [];
  private workbookPath = "xl/workbook.xml";

  constructor(data: Uint8Array, env: XmlEnv) {
    this.env = env;
    try {
      this.files = unzipSync(data);
    } catch {
      throw new Error("Файл не похож на .xlsx (возможно, это старый формат .xls — пересохрани его как .xlsx)");
    }
    this.init();
  }

  private readXml(path: string): Document {
    const bytes = this.files[path];
    if (!bytes) throw new Error(`В файле нет части ${path}`);
    return this.env.parse(strFromU8(bytes));
  }

  private init() {
    const rootRels = this.readXml("_rels/.rels");
    const officeRel = descendants(rootRels, "Relationship").find((r) =>
      (r.getAttribute("Type") ?? "").endsWith("/officeDocument"),
    );
    if (officeRel) this.workbookPath = resolvePath("", officeRel.getAttribute("Target")!);

    const relsPath = this.workbookPath.replace(/([^/]+)$/, "_rels/$1.rels");
    const rels = new Map<string, { type: string; target: string }>();
    for (const r of descendants(this.readXml(relsPath), "Relationship")) {
      rels.set(r.getAttribute("Id")!, {
        type: r.getAttribute("Type") ?? "",
        target: resolvePath(this.workbookPath, r.getAttribute("Target")!),
      });
    }

    for (const rel of rels.values()) {
      if (rel.type.endsWith("/styles")) this.stylesPath = rel.target;
      if (rel.type.endsWith("/theme")) this.themePath = rel.target;
      if (rel.type.endsWith("/sharedStrings")) {
        this.sstPath = rel.target;
        this.sstDoc = this.readXml(rel.target);
        const root = this.sstDoc.documentElement;
        elementChildren(root, "si").forEach((si, i) => {
          const s = richText(si);
          this.sst.push(s);
          if (!this.sstIndex.has(s)) this.sstIndex.set(s, i);
        });
      }
    }

    const wbDoc = this.readXml(this.workbookPath);
    for (const s of descendants(wbDoc, "sheet")) {
      const rid = s.getAttributeNS(REL_NS, "id") ?? s.getAttribute("r:id");
      const rel = rid ? rels.get(rid) : undefined;
      if (rel) this.sheetOrder.push({ name: s.getAttribute("name") ?? "", path: rel.target });
    }
  }

  private stylesPath: string | null = null;
  private themePath: string | null = null;
  private xfFills: (string | null)[] | null = null;

  /** Цвет заливки для стиля ячейки (индекс в cellXfs) → "RRGGBB" или null. */
  styleFill(xfIndex: number): string | null {
    if (!this.xfFills) this.xfFills = this.parseFills();
    return this.xfFills[xfIndex] ?? null;
  }

  private parseFills(): (string | null)[] {
    if (!this.stylesPath || !this.files[this.stylesPath]) return [];
    const theme = this.parseTheme();
    const doc = this.readXml(this.stylesPath);
    const fillsEl = descendants(doc, "fills")[0];
    const fills: (string | null)[] = fillsEl
      ? elementChildren(fillsEl, "fill").map((f) => {
          const p = firstChild(f, "patternFill");
          if (!p || (p.getAttribute("patternType") ?? "none") === "none") return null;
          const fg = firstChild(p, "fgColor");
          return fg ? resolveColor(fg, theme) : null;
        })
      : [];
    const xfsEl = descendants(doc, "cellXfs")[0];
    if (!xfsEl) return [];
    return elementChildren(xfsEl, "xf").map((xf) => fills[Number(xf.getAttribute("fillId") ?? 0)] ?? null);
  }

  /** Цвета темы в порядке индексов Excel: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink. */
  private parseTheme(): string[] {
    if (!this.themePath || !this.files[this.themePath]) return [];
    const scheme = descendants(this.readXml(this.themePath), "clrScheme")[0];
    if (!scheme) return [];
    const byName = new Map<string, string>();
    for (const el of elementChildren(scheme)) {
      const v = elementChildren(el)[0];
      if (!v) continue;
      byName.set(el.localName, (v.getAttribute("lastClr") ?? v.getAttribute("val") ?? "000000").toUpperCase());
    }
    return ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
      .map((n) => byName.get(n) ?? "000000");
  }

  get sheetNames(): string[] {
    return this.sheetOrder.map((s) => s.name);
  }

  sheet(name: string): Sheet {
    let sh = this.sheetsByName.get(name);
    if (sh) return sh;
    const meta = this.sheetOrder.find((s) => s.name === name);
    if (!meta) throw new Error(`Нет листа «${name}»`);
    sh = new Sheet(this, meta.name, meta.path, this.readXml(meta.path));
    this.sheetsByName.set(name, sh);
    return sh;
  }

  sharedString(i: number): string {
    return this.sst[i] ?? "";
  }

  /** Добавляет строку в sharedStrings; null — если в книге нет sharedStrings (тогда пишем inlineStr). */
  addSharedString(s: string): number | null {
    if (!this.sstDoc) return null;
    const existing = this.sstIndex.get(s);
    if (existing !== undefined) return existing;
    const root = this.sstDoc.documentElement;
    const ns = root.namespaceURI;
    const prefix = root.prefix;
    const si = this.sstDoc.createElementNS(ns, prefix ? `${prefix}:si` : "si");
    const t = this.sstDoc.createElementNS(ns, prefix ? `${prefix}:t` : "t");
    if (s !== s.trim()) t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.textContent = s;
    si.appendChild(t);
    root.appendChild(si);
    const idx = this.sst.length;
    this.sst.push(s);
    this.sstIndex.set(s, idx);
    this.sstDirty = true;
    return idx;
  }

  private writeXml(path: string, doc: Document) {
    const text = this.env.serialize(doc).replace(/^\s*<\?xml[^?]*\?>\s*/, "");
    this.files[path] = strToU8(XML_DECL + text);
  }

  /** Собрать изменённый xlsx. */
  save(): Uint8Array {
    let changed = false;
    for (const sh of this.sheetsByName.values()) {
      if (sh.dirty) {
        this.writeXml(sh.path, sh.getXml());
        changed = true;
      }
    }
    if (this.sstDirty && this.sstDoc && this.sstPath) {
      const root = this.sstDoc.documentElement;
      root.setAttribute("uniqueCount", String(this.sst.length));
      // count — общее число ссылок; точное значение Excel не проверяет, но оно не должно быть меньше uniqueCount
      const count = Number(root.getAttribute("count") ?? 0);
      if (count < this.sst.length) root.setAttribute("count", String(this.sst.length));
      this.writeXml(this.sstPath, this.sstDoc);
    }
    if (changed) {
      // Попросить Excel пересчитать формулы при открытии (например, столбцы, ссылающиеся на «Клиент»)
      const wbDoc = this.readXml(this.workbookPath);
      const calcPr = descendants(wbDoc, "calcPr")[0];
      if (calcPr) {
        calcPr.setAttribute("fullCalcOnLoad", "1");
        this.writeXml(this.workbookPath, wbDoc);
      }
    }
    return zipSync(this.files, { level: 6 });
  }
}

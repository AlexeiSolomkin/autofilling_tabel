/**
 * UI сервиса автозаполнения табелей.
 */
import { Workbook, browserXmlEnv } from "./xlsx/workbook";
import { buildReference, REF_TITLES, type Reference, type RefKind } from "./reference";
import { fillTabel, type TabelKind, type TabelResult, type LogEntry } from "./tabel";

const env = browserXmlEnv();

interface Slot {
  key: "import" | "holland" | "ref";
  drop: HTMLElement;
  input: HTMLInputElement;
  file: { name: string; bytes: Uint8Array } | null;
  error: string | null;
  meta: string;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const slots: Record<Slot["key"], Slot> = {
  import: { key: "import", drop: $("dropImport"), input: $("fileImport"), file: null, error: null, meta: "" },
  holland: { key: "holland", drop: $("dropHolland"), input: $("fileHolland"), file: null, error: null, meta: "" },
  ref: { key: "ref", drop: $("dropRef"), input: $("fileRef"), file: null, error: null, meta: "" },
};

let reference: Reference | null = null;
let results: { kind: TabelKind; fileName: string; res: TabelResult }[] = [];

const runBtn = $<HTMLButtonElement>("runBtn");
const downloadBtn = $<HTMLButtonElement>("downloadBtn");
const consoleEl = $("console");
const problemsOnly = $<HTMLInputElement>("problemsOnly");

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

function renderSlot(s: Slot) {
  const d = s.drop;
  d.replaceChildren();
  d.classList.toggle("loaded", !!s.file && !s.error);
  d.classList.toggle("bad", !!s.error);
  if (!s.file) {
    d.append(el("div", "icon", d.dataset.icon), el("div", "title", "Перетащи файл или нажми для выбора"), el("div", "hint", d.dataset.hint));
    return;
  }
  const rm = el("button", "remove", "×");
  rm.title = "Убрать файл";
  rm.addEventListener("click", (ev) => {
    ev.stopPropagation();
    s.file = null;
    s.error = null;
    s.meta = "";
    if (s.key === "ref") reference = null;
    renderSlot(s);
    invalidate();
  });
  d.append(
    rm,
    el("div", "icon", s.error ? "⚠️" : "✅"),
    el("div", "fname", s.file.name),
    el("div", "meta", s.error ?? s.meta),
  );
}

function updateButtons() {
  const haveTabel = (!!slots.import.file && !slots.import.error) || (!!slots.holland.file && !slots.holland.error);
  runBtn.disabled = !(haveTabel && reference);
  downloadBtn.disabled = !results.some((r) => r.res.output);
}

/** Любое изменение входных файлов делает прошлый результат неактуальным. */
function invalidate() {
  if (results.length) {
    results = [];
    consoleEl.replaceChildren(el("div", "empty", "Файлы изменились — нажми «Заполнить табеля» ещё раз"));
  }
  updateButtons();
}

async function acceptFile(s: Slot, f: File) {
  s.file = { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) };
  s.error = null;
  s.meta = `${(f.size / 1024).toFixed(0)} КБ`;
  if (!/\.xlsx?m?$/i.test(f.name)) s.error = "Нужен файл .xlsx";
  if (s.key === "ref" && !s.error) {
    try {
      const t0 = performance.now();
      reference = buildReference(new Workbook(s.file.bytes, env));
      const ms = Math.round(performance.now() - t0);
      const parts = (["holland", "import", "jaaz", "mf"] as RefKind[])
        .filter((k) => reference!.sheets[k])
        .map((k) => `${REF_TITLES[k]}: ${reference!.sheets[k]!.markingCount} маркировок`);
      s.meta = `${parts.join(" · ")} · разобран за ${ms} мс`;
      if (reference.errors.length) s.meta += ` · ⚠️ ${reference.errors.join("; ")}`;
    } catch (e) {
      reference = null;
      s.error = (e as Error).message;
    }
  }
  renderSlot(s);
  invalidate();
}

for (const s of Object.values(slots)) {
  renderSlot(s);
  s.drop.addEventListener("click", () => s.input.click());
  s.input.addEventListener("change", () => {
    const f = s.input.files?.[0];
    if (f) void acceptFile(s, f);
    s.input.value = "";
  });
  s.drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    s.drop.classList.add("over");
  });
  s.drop.addEventListener("dragleave", () => s.drop.classList.remove("over"));
  s.drop.addEventListener("drop", (e) => {
    e.preventDefault();
    s.drop.classList.remove("over");
    const f = e.dataTransfer?.files?.[0];
    if (f) void acceptFile(s, f);
  });
}
// Чтобы случайно брошенный мимо файл не открывался в браузере
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => e.preventDefault());

function outName(name: string): string {
  return name.replace(/(\.xlsx?m?)?$/i, (ext) => ` (заполнен)${ext || ".xlsx"}`);
}

function download(name: string, bytes: Uint8Array) {
  const blob = new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = el("a");
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function lines(entries: LogEntry[], into: HTMLElement) {
  for (const e of entries) into.append(el("div", `line ${e.level}`, e.text));
}

function renderResults(errors: { kind: TabelKind; fileName: string; message: string }[]) {
  consoleEl.replaceChildren();
  const order: TabelKind[] = ["import", "holland"];
  for (const kind of order) {
    const r = results.find((x) => x.kind === kind);
    const err = errors.find((x) => x.kind === kind);
    if (!r && !err) continue;
    const sec = el("div", "section");
    const head = el("div", "sec-head");
    head.append(el("span", `sec-title ${kind === "import" ? "imp" : "hol"}`, kind === "import" ? "Импорт" : "Голландия"));
    head.append(el("span", "chip", (r ?? err)!.fileName));
    sec.append(head);

    if (err) {
      lines([{ level: "error", text: `Не удалось обработать файл: ${err.message}` }], sec);
      consoleEl.append(sec);
      continue;
    }
    const { res } = r!;
    const chips = el("div", "chips");
    chips.append(
      el("span", "chip", `строк: ${res.stats.rows}`),
      el("span", "chip filled", `заполнено: ${res.stats.filled}`),
      el("span", "chip", `уже было: ${res.stats.already}`),
      el("span", "chip warn", `проверить: ${res.stats.warnings}`),
      el("span", "chip err", `не найдено: ${res.stats.notFound}`),
    );
    if (res.stats.uncertain) chips.append(el("span", "chip warn", `неточно: ${res.stats.uncertain}`));
    if (res.stats.defaulted) chips.append(el("span", "chip warn", `по умолчанию: ${res.stats.defaulted}`));
    head.append(chips);
    if (res.output) {
      const one = el("button", "dl-one", "⬇ скачать этот");
      one.addEventListener("click", () => download(outName(r!.fileName), res.output!));
      head.append(one);
    }

    lines(res.general, sec);
    if (!res.output) lines([{ level: "info", text: "Изменений нет — файл для скачивания не формируется" }], sec);
    for (const t of res.tables) {
      const det = el("details", "tbl");
      det.open = true;
      const problems = t.entries.filter((e) => e.level === "warn" || e.level === "error").length;
      const sum = el("summary");
      sum.append(t.title + " ", el("span", "cnt", problems ? `· проблем: ${problems}` : "· без проблем"));
      det.append(sum);
      lines(t.entries, det);
      sec.append(det);
    }
    consoleEl.append(sec);
  }
}

runBtn.addEventListener("click", () => {
  if (!reference) return;
  runBtn.disabled = true;
  results = [];
  const errors: { kind: TabelKind; fileName: string; message: string }[] = [];
  for (const kind of ["import", "holland"] as TabelKind[]) {
    const s = slots[kind];
    if (!s.file || s.error) continue;
    try {
      // Каждый раз читаем исходные байты заново — повторное нажатие даёт тот же результат
      const res = fillTabel(new Workbook(s.file.bytes, env), kind, reference);
      results.push({ kind, fileName: s.file.name, res });
    } catch (e) {
      errors.push({ kind, fileName: s.file.name, message: (e as Error).message });
    }
  }
  renderResults(errors);
  updateButtons();
});

downloadBtn.addEventListener("click", () => {
  results.forEach((r, i) => {
    // небольшая пауза между файлами, чтобы браузер не склеил загрузки
    if (r.res.output) setTimeout(() => download(outName(r.fileName), r.res.output!), i * 400);
  });
});

problemsOnly.addEventListener("change", () => consoleEl.classList.toggle("problems-only", problemsOnly.checked));

const modal = $("helpModal");
$("helpBtn").addEventListener("click", () => modal.classList.add("open"));
$("helpClose").addEventListener("click", () => modal.classList.remove("open"));
modal.addEventListener("click", (e) => {
  if (e.target === modal) modal.classList.remove("open");
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") modal.classList.remove("open");
});

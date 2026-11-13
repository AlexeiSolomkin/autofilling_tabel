/** Утилиты для адресов ячеек вида "AB12". */

export function colToNum(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function splitAddr(addr: string): { col: number; row: number } {
  const m = /^([A-Z]+)(\d+)$/i.exec(addr);
  if (!m) throw new Error(`Некорректный адрес ячейки: ${addr}`);
  return { col: colToNum(m[1]), row: Number(m[2]) };
}

export function addr(col: number, row: number): string {
  return numToCol(col) + row;
}

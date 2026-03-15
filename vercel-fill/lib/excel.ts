import * as XLSX from 'xlsx';

export const CAR_COL_CAND = ['車號/最小成本單位', '車號', '最小成本單位'];
export const TIME_COL_CAND = ['檢修日期', '檢查時間', '複查時間', '時間', '紀錄時間'];

const RX14 = /^\d{14}$/;
const RXYMD = /^\d{4}[/-]\d{1,2}[/-]\d{1,2}(\s+\d{1,2}:\d{2}(:\d{2})?)?$/;

export function isDate1904(wb: XLSX.WorkBook): boolean {
  // @ts-ignore
  const wbProps = wb?.Workbook?.WBProps;
  return !!wbProps?.date1904;
}

export function excelSerialToDate(serial: number, date1904: boolean): Date | null {
  if (typeof serial !== 'number' || !isFinite(serial)) return null;
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const ms = serial * 24 * 3600 * 1000;
  return new Date(base + ms);
}

export function dateToExcelSerial(d: Date, date1904: boolean): number {
  const y = d.getUTCFullYear(), m = d.getUTCMonth(), day = d.getUTCDate();
  const only = Date.UTC(y, m, day);
  const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (only - base) / (24 * 3600 * 1000);
}

function normalizeAndParseTextDate(s: string): Date | null {
  let t = s.trim();
  if (!t) return null;
  t = t.replace(/[年/.-]/g, '/').replace(/[日]/g, '').replace(/\s*(上午|下午|AM|PM)\s*/gi, ' ');
  if (RXYMD.test(t)) {
    const [datePart, timePart] = t.split(/\s+/);
    const [Y, M, D] = datePart.split('/').map(n => parseInt(n, 10));
    let hh = 0, mm = 0, ss = 0;
    if (timePart) {
      const tt = timePart.split(':').map(n => parseInt(n, 10));
      hh = tt[0] ?? 0; mm = tt[1] ?? 0; ss = tt[2] ?? 0;
    }
    return new Date(Date.UTC(Y, (M || 1) - 1, D || 1, hh, mm, ss));
  }
  return null;
}

export function parseAnyDate(val: any, date1904: boolean): Date | null {
  if (val == null) return null;

  if (val instanceof Date) {
    return new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate(), val.getHours(), val.getMinutes(), val.getSeconds()));
  }
  if (typeof val === 'number' && isFinite(val)) {
    return excelSerialToDate(val, date1904);
  }
  const s = String(val).trim();
  if (!s) return null;

  if (RX14.test(s)) {
    const Y = +s.slice(0, 4), M = +s.slice(4, 6), D = +s.slice(6, 8);
    const h = +s.slice(8, 10), m = +s.slice(10, 12), sec = +s.slice(12, 14);
    return new Date(Date.UTC(Y, M - 1, D, h, m, sec));
  }
  if (RXYMD.test(s) || /上午|下午|AM|PM|年|月|日/.test(s)) {
    return normalizeAndParseTextDate(s);
  }
  return null;
}

export function readWb(buf: Buffer): XLSX.WorkBook {
  return XLSX.read(buf, { type: 'buffer', cellDates: true, cellNF: true, cellText: false });
}

export function headerIndex(headerRow: any[], label: string): number {
  const idx = headerRow.findIndex(v => String(v ?? '').trim() === label);
  return idx >= 0 ? idx + 1 : -1;
}

export function headerIndexOneOf(headerRow: any[], labels: string[]): number {
  for (const lab of labels) {
    const idx = headerIndex(headerRow, lab);
    if (idx > 0) return idx;
  }
  return -1;
}

export function buildLatestByCar(wb: XLSX.WorkBook): Record<string, Date> {
  const date1904 = isDate1904(wb);
  const latest: Record<string, Date> = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, raw: true });
    if (aoa.length < 2) continue;
    const header = aoa[0].map(v => String(v ?? '').trim());
    const carIdx = headerIndexOneOf(header, CAR_COL_CAND);
    if (carIdx < 1) continue;
    const timeIdxs = TIME_COL_CAND.map(lab => headerIndex(header, lab)).filter(i => i > 0);
    if (!timeIdxs.length) continue;

    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r];
      const carVal = row[carIdx - 1];
      const car = String(carVal ?? '').trim();
      if (!car) continue;

      let best: Date | null = null;
      for (const c of timeIdxs) {
        const v = row[c - 1];
        const dt = parseAnyDate(v, date1904);
        if (dt && (!best || dt.getTime() > best.getTime())) best = dt;
      }
      if (!best) continue;
      if (!latest[car] || best.getTime() > latest[car].getTime()) latest[car] = best;
    }
  }
  return latest;
}

export function writeDateOnlyCell(cell: XLSX.CellObject, d: Date, date1904: boolean) {
  const serial = dateToExcelSerial(d, date1904);
  // @ts-ignore
  cell.t = 'n';
  // @ts-ignore
  cell.v = serial;
  // @ts-ignore
  cell.z = 'yyyy/mm/dd';
  // @ts-ignore
  if (cell.f) delete cell.f;
}

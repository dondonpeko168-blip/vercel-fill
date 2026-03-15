import { NextRequest } from 'next/server';
import * as XLSX from 'xlsx';
import {
  readWb, isDate1904, buildLatestByCar,
  headerIndex, writeDateOnlyCell
} from '@/lib/excel';

export const runtime = 'nodejs';

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const source = form.get('source') as File | null;
    const target = form.get('target') as File | null;
    const sheetName = String(form.get('sheet') ?? '鋼輪計算_115').trim();
    const overwrite = String(form.get('overwrite') ?? 'true') === 'true';

    if (!source || !target) {
      return new Response('missing files: source or target', { status: 400 });
    }

    const srcBuf = Buffer.from(await source.arrayBuffer());
    const tgtBuf = Buffer.from(await target.arrayBuffer());

    const srcWb = readWb(srcBuf);
    const latestByCar = buildLatestByCar(srcWb);

    const tgtWb = readWb(tgtBuf);
    const date1904 = isDate1904(tgtWb);

    const ws = tgtWb.Sheets[sheetName] ?? tgtWb.Sheets[tgtWb.SheetNames[0]];
    if (!ws) return new Response('target workbook has no sheets', { status: 400 });

    const aoa = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, raw: true });
    if (aoa.length < 2) return new Response('target sheet has no data', { status: 400 });

    const header = aoa[0].map(v => String(v ?? '').trim());
    const carCol = headerIndex(header, '車號');
    const dateCol = headerIndex(header, '檢修日期');
    if (carCol < 1 || dateCol < 1) {
      return new Response('target sheet missing 車號 or 檢修日期 header', { status: 400 });
    }

    for (let r = 2; r <= aoa.length; r++) {
      const carRef = XLSX.utils.encode_cell({ r: r - 1, c: carCol - 1 });
      const dateRef = XLSX.utils.encode_cell({ r: r - 1, c: dateCol - 1 });
      const carCell = ws[carRef];
      if (!carCell) continue;

      const car = String(carCell.v ?? '').trim();
      if (!car) continue;

      const chosen = latestByCar[car];
      if (!chosen) continue;

      const cell = ws[dateRef] ?? (ws[dateRef] = {} as any);
      const hasFormula = !!(cell as any).f;
      if (!overwrite && hasFormula) continue;

      writeDateOnlyCell(cell, chosen, date1904);
    }

    const outWbBuf = XLSX.write(tgtWb, { type: 'buffer', bookType: 'xlsx' });

    return new Response(outWbBuf, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="filled.xlsx"'
      }
    });
  } catch (err: any) {
    console.error(err);
    return new Response('internal error: ' + (err?.message || String(err)), { status: 500 });
  }
}

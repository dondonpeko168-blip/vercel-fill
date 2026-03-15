'use client';
import React, { useState } from 'react';

export default function Page() {
  const [sheet, setSheet] = useState('鋼輪計算_115');
  const [overwrite, setOverwrite] = useState(true);
  const [downUrl, setDownUrl] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function onSubmit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    setDownUrl(null);
    setBusy(true);
    const fd = new FormData(e.currentTarget);
    fd.set('sheet', sheet);
    fd.set('overwrite', String(overwrite));
    const res = await fetch('/api/fill', { method: 'POST', body: fd });
    setBusy(false);
    if (!res.ok) {
      const txt = await res.text();
      alert(`Failed: ${res.status}
${txt}`);
      return;
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    setDownUrl(url);
  }

  return (
    <main style={{ maxWidth: 720, margin: '40px auto', fontFamily: 'system-ui', lineHeight: 1.5 }}>
      <h1>回填測試（來源 → 目標）</h1>
      <p>上傳 <code>source.xlsx</code>（第 1 步網站匯出）與 <code>target.xlsx</code>（原始目標），本服務將依規則回填「檢修日期」。</p>
      <form onSubmit={onSubmit}>
        <div>
          <label>來源檔（source.xlsx）</label><br />
          <input name="source" type="file" accept=".xlsx" required />
        </div>
        <div style={{ marginTop: 12 }}>
          <label>目標檔（target.xlsx）</label><br />
          <input name="target" type="file" accept=".xlsx" required />
        </div>
        <div style={{ marginTop: 12 }}>
          <label>目標分頁（預設：鋼輪計算_115）</label><br />
          <input value={sheet} onChange={e => setSheet(e.target.value)} />
        </div>
        <div style={{ marginTop: 12 }}>
          <label>
            <input type="checkbox" checked={overwrite} onChange={e => setOverwrite(e.target.checked)} />
            覆蓋公式（取消勾選＝保留公式、略過該列）
          </label>
        </div>
        <div style={{ marginTop: 16 }}>
          <button disabled={busy} type="submit">{busy ? '處理中…' : '開始回填'}</button>
        </div>
      </form>
      {downUrl && (
        <p style={{ marginTop: 16 }}>
          ✅ 完成：<a href={downUrl} download="filled.xlsx">下載 filled.xlsx</a>
        </p>
      )}
    </main>
  );
}

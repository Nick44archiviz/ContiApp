// Import ed export Excel (ExcelJS, caricato solo quando serve).
// Il file esportato ha lo stesso schema del foglio originale: un foglio per mese
// (Descrizione, Dettagli, 4 conti, Totale, riga Apertura e Chiusura) e un foglio "Totale AAAA" con formule.
// Lo stesso schema viene letto in importazione, quindi si può esportare, modificare da PC e ricaricare.

export const MN = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
const pad = n => String(n).padStart(2, '0');
const r2 = n => Math.round(n * 100) / 100 + 0;
const sum = a => a.reduce((x, y) => x + y, 0);

let loading = null;
export function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (loading) return loading;
  loading = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
    s.onload = () => res(window.ExcelJS);
    s.onerror = () => { loading = null; rej(new Error('Non riesco a caricare il modulo Excel. Controlla la connessione e riprova.')); };
    document.head.appendChild(s);
  });
  return loading;
}

/* ------------------------------------------------------------------ export */
const NUM = '#,##0.00;[Red]-#,##0.00';
const q = s => "'" + s.replace(/'/g, "''") + "'";
const lit = s => '"' + s.replace(/"/g, '""') + '"';
const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const F = (formula, result) => ({ formula, result });

function head(row, from, to) {
  for (let c = from; c <= to; c++) {
    const cell = row.getCell(c);
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E5BEA' } };
    cell.alignment = { horizontal: c <= 3 ? 'left' : 'right' };
  }
}

// ctx: { year, c (risultato di calc), cfg:{acc,cats,inc,goal}, month: null | 0..11 }
export async function buildWorkbook(ExcelJS, ctx) {
  const { year, c, cfg } = ctx;
  const all = ctx.month == null;
  const idx = all ? [...Array(12).keys()] : [ctx.month];
  const names = i => `${MN[i]} ${year}`;
  const closeRow = i => 4 + c.M[i].rows.length;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Conti 2026'; wb.created = new Date();
  const cats = cfg.inc.concat(cfg.cats);

  idx.forEach((i, pos) => {
    const m = c.M[i];
    const ws = wb.addWorksheet(names(i), { views: [{ state: 'frozen', ySplit: 2 }] });
    const rows = m.rows.slice().sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : a.o - b.o);
    const last = 3 + rows.length, cr = closeRow(i);
    ws.getCell('A1').value = names(i); ws.getCell('A1').font = { bold: true, size: 14 };
    const h = ws.getRow(2);
    h.values = [null, 'Descrizione', 'Dettagli', ...cfg.acc, 'Totale'];
    h.getCell(1).value = 'Data';
    head(h, 1, 8);
    // apertura
    ws.getCell('A3').value = 'Apertura'; ws.getCell('A3').font = { bold: true };
    for (let k = 0; k < 4; k++) {
      const cell = ws.getCell(3, 4 + k);
      cell.value = pos === 0 ? r2(m.open[k]) : F(`${q(names(i - 1))}!${colL(4 + k)}${closeRow(i - 1)}`, r2(m.open[k]));
      cell.numFmt = NUM;
    }
    ws.getCell('H3').value = F('SUM(D3:G3)', r2(sum(m.open))); ws.getCell('H3').numFmt = NUM;
    // movimenti
    rows.forEach((r, n) => {
      const rr = 4 + n, row = ws.getRow(rr);
      const [y, mo, d] = r.d.split('-').map(Number);
      row.getCell(1).value = new Date(Date.UTC(y, mo - 1, d)); row.getCell(1).numFmt = 'dd/mm/yyyy';
      row.getCell(2).value = r.c; row.getCell(3).value = r.t || '';
      r.v.forEach((x, k) => { if (Math.abs(x) > 0.004) { row.getCell(4 + k).value = r2(x); } row.getCell(4 + k).numFmt = NUM; });
      row.getCell(8).value = F(`SUM(D${rr}:G${rr})`, r2(sum(r.v))); row.getCell(8).numFmt = NUM;
    });
    // chiusura
    const cl = ws.getRow(cr);
    cl.getCell(1).value = 'Chiusura';
    for (let k = 0; k < 4; k++) { cl.getCell(4 + k).value = F(`SUM(${colL(4 + k)}3:${colL(4 + k)}${last})`, r2(m.close[k])); cl.getCell(4 + k).numFmt = NUM; }
    cl.getCell(8).value = F(`SUM(H3:H${last})`, r2(sum(m.close))); cl.getCell(8).numFmt = NUM;
    for (let cc = 1; cc <= 8; cc++) { cl.getCell(cc).font = { bold: true }; cl.getCell(cc).border = { top: { style: 'thin' } }; }
    // elenco categorie a tendina
    for (let rr = 4; rr <= last + 30 && rr !== cr; rr++) {
      ws.getCell(rr, 2).dataValidation = { type: 'list', allowBlank: true, formulae: [`Liste!$A$1:$A$${cats.length}`] };
    }
    // riepilogo per categoria (sostituisce la tabella pivot)
    ws.getCell('J2').value = 'Categoria'; ws.getCell('K2').value = 'Totale mese'; head(ws.getRow(2), 10, 11);
    cats.forEach((cat, n) => {
      const rr = 3 + n, isInc = cfg.inc.includes(cat);
      ws.getCell(rr, 10).value = cat;
      const val = isInc ? (m.byCat[cat] || 0) : -(m.byCat[cat] || 0);
      ws.getCell(rr, 11).value = F(`SUMIF($B$4:$B$${last},J${rr},$H$4:$H$${last})`, r2(val)); ws.getCell(rr, 11).numFmt = NUM;
    });
    const endCat = 2 + cats.length;
    const known = cats.reduce((s, cat) => s + (cfg.inc.includes(cat) ? (m.byCat[cat] || 0) : -(m.byCat[cat] || 0)), 0);
    const ok = Math.abs(sum(m.open) + known - sum(m.close)) < 0.005;
    ws.getCell('M2').value = 'Quadratura'; ws.getCell('M2').font = { bold: true };
    ws.getCell('M3').value = F(`IF(ROUND(H3+SUM(K3:K${endCat})-H${cr},2)=0,"ok","no")`, ok ? 'ok' : 'no');
    [[1, 12], [2, 26], [3, 36], [4, 17], [5, 17], [6, 17], [7, 17], [8, 15], [9, 3], [10, 26], [11, 15], [12, 3], [13, 12]].forEach(([n, w]) => ws.getColumn(n).width = w);
  });

  if (all) {
    const ws = wb.addWorksheet(`Totale ${year}`);
    ws.getCell('A1').value = `Anno ${year}`; ws.getCell('A1').font = { bold: true, size: 14 };
    const incF = i => cfg.inc.length ? cfg.inc.map(k => `SUMIF(${q(names(i))}!$B:$B,${lit(k)},${q(names(i))}!$H:$H)`).join('+') : '0';
    const r = ws.getRow(2); r.values = [null, 'Entrate', 'Uscite', 'Utile/(Perdita)']; r.getCell(1).value = 'Mese'; head(r, 1, 4);
    c.M.forEach((m, i) => {
      const rr = 3 + i;
      ws.getCell(rr, 1).value = MN[i];
      ws.getCell(rr, 2).value = F(incF(i), r2(m.ent));
      ws.getCell(rr, 3).value = F(`(${q(names(i))}!H${closeRow(i)}-${q(names(i))}!H3)-B${rr}`, r2(-m.exp));
      ws.getCell(rr, 4).value = F(`B${rr}+C${rr}`, r2(m.utile));
      [2, 3, 4].forEach(k => ws.getCell(rr, k).numFmt = NUM);
    });
    ws.getCell('A15').value = 'Totale'; ws.getCell('A15').font = { bold: true };
    [['B', c.ent], ['C', -c.exp], ['D', c.utile]].forEach(([L, v]) => { const x = ws.getCell(L + '15'); x.value = F(`SUM(${L}3:${L}14)`, r2(v)); x.numFmt = NUM; x.font = { bold: true }; });
    ws.getCell('A16').value = '% sulle entrate';
    ws.getCell('C16').value = F('IF(B15=0,0,-C15/B15)', c.ent ? c.exp / c.ent : 0); ws.getCell('D16').value = F('IF(B15=0,0,D15/B15)', c.ent ? c.utile / c.ent : 0);
    ws.getCell('C16').numFmt = '0.0%'; ws.getCell('D16').numFmt = '0.0%';

    // uscite per categoria x mese
    let row = 19;
    const tbl = (title, list, isInc) => {
      const hr = ws.getRow(row); hr.values = [null]; hr.getCell(1).value = title;
      MN.forEach((mn, i) => hr.getCell(2 + i).value = mn.slice(0, 3)); hr.getCell(14).value = 'Totale'; head(hr, 1, 14);
      const first = row + 1;
      list.forEach((cat, n) => {
        const rr = first + n; ws.getCell(rr, 1).value = cat;
        for (let i = 0; i < 12; i++) {
          const ref = `${q(names(i))}!$B:$B,$A${rr},${q(names(i))}!$H:$H`;
          const v = isInc ? (c.M[i].byCat[cat] || 0) : (c.M[i].byCat[cat] || 0);
          ws.getCell(rr, 2 + i).value = F(`${isInc ? '' : '-'}SUMIF(${ref})`, r2(v)); ws.getCell(rr, 2 + i).numFmt = NUM;
        }
        ws.getCell(rr, 14).value = F(`SUM(B${rr}:M${rr})`, r2(sum(c.M.map(m => m.byCat[cat] || 0)))); ws.getCell(rr, 14).numFmt = NUM; ws.getCell(rr, 14).font = { bold: true };
      });
      const tr = first + list.length;
      ws.getCell(tr, 1).value = isInc ? 'Totale entrate' : 'Totale uscite';
      for (let k = 2; k <= 14; k++) {
        const L = colL(k);
        const res = k === 14 ? sum(list.map(cat => sum(c.M.map(m => m.byCat[cat] || 0)))) : sum(list.map(cat => c.M[k - 2].byCat[cat] || 0));
        ws.getCell(tr, k).value = F(list.length ? `SUM(${L}${first}:${L}${tr - 1})` : '0', r2(res)); ws.getCell(tr, k).numFmt = NUM;
      }
      for (let k = 1; k <= 14; k++) { ws.getCell(tr, k).font = { bold: true }; ws.getCell(tr, k).border = { top: { style: 'thin' } }; }
      row = tr + 2;
      return { tr, total: sum(list.map(cat => sum(c.M.map(m => m.byCat[cat] || 0)))) };
    };
    const ex = tbl('USCITE', cfg.cats, false);
    const en = tbl('ENTRATE', cfg.inc, true);

    // controlli
    ws.getCell(row, 1).value = 'CONTROLLI'; ws.getCell(row, 1).font = { bold: true }; row++;
    const eq = (a, b) => Math.abs(a - b) < 0.005 ? 'Ok' : 'No';
    const chk = (label, formula, res) => { ws.getCell(row, 1).value = label; ws.getCell(row, 2).value = F(formula, res); row++; };
    chk('Entrate: mesi = categorie', `IF(ROUND(N${en.tr}-B15,2)=0,"Ok","No")`, eq(en.total, c.ent));
    chk('Uscite: mesi = categorie', `IF(ROUND(N${ex.tr}+C15,2)=0,"Ok","No")`, eq(ex.total, c.exp));
    const closeDec = `${q(names(11))}!H${closeRow(11)}`, openGen = `${q(names(0))}!H3`;
    chk('Utile = chiusura - apertura', `IF(ROUND(D15-(${closeDec}-${openGen}),2)=0,"Ok","No")`, eq(c.utile, c.closeT - c.openT));
    row++;
    ws.getCell(row, 1).value = `Saldo di apertura 01.01.${year}`; ws.getCell(row, 2).value = F(openGen, r2(c.openT)); ws.getCell(row, 2).numFmt = NUM; const ro = row++;
    ws.getCell(row, 1).value = `Saldo di chiusura 31.12.${year}`; ws.getCell(row, 2).value = F(closeDec, r2(c.closeT)); ws.getCell(row, 2).numFmt = NUM; row++;
    ws.getCell(row, 1).value = 'Obiettivo di fine anno'; ws.getCell(row, 2).value = F(`B${ro}+${cfg.goal || 0}`, r2(c.openT + (cfg.goal || 0))); ws.getCell(row, 2).numFmt = NUM;
    ws.getColumn(1).width = 30; for (let k = 2; k <= 14; k++) ws.getColumn(k).width = 13;
  }
  const li = wb.addWorksheet('Liste');
  cats.forEach((k, n) => li.getCell(n + 1, 1).value = k);
  cfg.acc.forEach((k, n) => li.getCell(n + 1, 3).value = k);
  li.state = 'hidden';
  return wb.xlsx.writeBuffer();
}

/* ------------------------------------------------------------------ import */
const monthIdx = name => MN.findIndex(m => m.toLowerCase() === String(name).toLowerCase());
function plain(v) {
  if (v == null) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('result' in v) return plain(v.result);
    if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
    if ('text' in v) return v.text;
    if ('error' in v) return null;
    return null;
  }
  return v;
}
const numOf = v => { v = plain(v); if (typeof v === 'number') return isFinite(v) ? v : 0; if (typeof v === 'string') { const n = parseFloat(v.replace(/\./g, (m, o, s) => s.includes(',') ? '' : m).replace(',', '.')); return isFinite(n) ? n : 0; } return 0; };
function dateOf(v) {
  v = plain(v);
  if (v instanceof Date) return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
  if (typeof v === 'string') { const m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(v.trim()); if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(+m[2])}-${pad(+m[1])}`; }
  return null;
}

// Legge un file .xlsx con lo schema dei fogli mensili. Restituisce i mesi trovati e le informazioni per il confronto.
export async function readWorkbook(ExcelJS, buffer, known) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const canon = {}; (known || []).forEach(k => canon[k.toLowerCase()] = k);
  const months = {}, info = {}, warnings = [];
  let acc = null;
  const found = [];
  wb.eachSheet(ws => {
    const m = /^\s*([A-Za-zÀ-ÿ]+)\s+(\d{4})\s*$/.exec(ws.name);
    if (!m) return;
    const mi = monthIdx(m[1]); if (mi < 0) return;
    const y = +m[2];
    // intestazione e righe di apertura/chiusura
    let ap = null, ch = null;
    const maxR = ws.rowCount;
    for (let r = 1; r <= Math.min(8, maxR); r++) if (String(plain(ws.getCell(r, 1).value) || '').trim().toLowerCase() === 'apertura') { ap = r; break; }
    for (let r = maxR; r >= 1; r--) if (String(plain(ws.getCell(r, 1).value) || '').trim().toLowerCase() === 'chiusura') { ch = r; break; }
    if (!ap || !ch || ch <= ap) { warnings.push(`Foglio "${ws.name}" saltato: non trovo le righe Apertura e Chiusura.`); return; }
    const hdr = [4, 5, 6, 7].map(c => String(plain(ws.getCell(2, c).value) || '').trim());
    if (hdr.every(Boolean)) acc = hdr; // vale l'ultimo foglio (il più recente)
    const open = [4, 5, 6, 7].map(c => r2(numOf(ws.getCell(ap, c).value)));
    const rows = {}; let d = null, n = 0;
    for (let r = ap + 1; r < ch; r++) {
      const dd = dateOf(ws.getCell(r, 1).value); if (dd) d = dd;
      const cat0 = plain(ws.getCell(r, 2).value);
      const v = [4, 5, 6, 7].map(c => r2(numOf(ws.getCell(r, c).value)));
      if (!v.some(x => Math.abs(x) > 0.004)) continue;
      const cat = String(cat0 == null ? '' : cat0).trim();
      n++;
      const det = plain(ws.getCell(r, 3).value);
      const row = { d: d || `${y}-${pad(mi + 1)}-01`, c: canon[cat.toLowerCase()] || cat, v, o: n };
      if (det != null && String(det).trim()) row.t = String(det).trim();
      rows[`i${pad(mi + 1)}_${String(n).padStart(3, '0')}`] = row;
    }
    const id = `${y}-${pad(mi + 1)}`;
    months[id] = { rows };
    info[id] = { open, sheet: ws.name };
    found.push(id);
  });
  found.sort();
  return { months, info, ids: found, acc, warnings };
}

// Confronta i mesi del file con quelli presenti nell'app.
export function diffMonths(parsed, docs) {
  const sig = r => [r.d, r.c, r.t || '', r.v.map(x => r2(x).toFixed(2)).join(',')].join('|');
  return parsed.ids.map(id => {
    const mine = Object.values((docs[id] && docs[id].rows) || {}).filter(r => r && !r.x);
    const theirs = Object.values(parsed.months[id].rows);
    const cnt = new Map(); mine.forEach(r => cnt.set(sig(r), (cnt.get(sig(r)) || 0) + 1));
    let added = 0; theirs.forEach(r => { const s = sig(r); if (cnt.get(s)) cnt.set(s, cnt.get(s) - 1); else added++; });
    let removed = 0; cnt.forEach(v => removed += v);
    return { id, file: theirs.length, app: mine.length, added, removed, same: added === 0 && removed === 0 };
  });
}

import { firebaseConfig } from './firebase-config.js';
import { createLocalBackend } from './backend-local.js';
import { MN, loadExcelJS, buildWorkbook, readWorkbook, diffMonths } from './excel.js';

const MS = MN.map(m => m.slice(0, 3));
const DEF = {
  acc: ['Contanti', 'Conto principale', 'Risparmio', 'Carta'],
  cats: ['Benzina', 'Auto', 'ATM/Trenord', 'Imposte e tasse', 'Vacanze', 'Shopping', 'Alimentari', 'Bar/Ristoranti', 'Iliad', 'Salute', 'Varie', 'Crypto', 'Bollette', 'Sport', 'Spotify/Apple', 'Materiale lavoro extra', 'Mutuo', 'Casa Inzago'],
  inc: ['Stipendio', 'Entrate varie'],
  goal: 20000, open: [0, 0, 0, 0]
};
let ACC = DEF.acc, CATS = DEF.cats, INC = DEF.inc;
const pad = n => String(n).padStart(2, '0');
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const r2 = n => Math.round(n * 100) / 100 + 0;
const sum = a => a.reduce((x, y) => x + y, 0);
const eur = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' });
const eur0 = new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const num0 = new Intl.NumberFormat('it-IT', { maximumFractionDigits: 0 });
const pct = new Intl.NumberFormat('it-IT', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fE = n => eur.format(r2(n)), fE0 = n => eur0.format(Math.round(n) + 0);
const sgn = n => (n > 0 ? '+' : '') + fE(n);
const dayFmt = new Intl.DateTimeFormat('it-IT', { weekday: 'short', day: 'numeric', month: 'short' });
const parseD = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const todayIso = () => { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); };
const ls = { get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }, set(k, v) { try { localStorage.setItem(k, v); } catch (e) { } } };
const parseAmt = s => { const n = parseFloat(String(s).replace(/\./g, (m, o, str) => str.includes(',') ? '' : m).replace(',', '.')); return isFinite(n) ? n : 0; };
const lines = s => [...new Set(String(s).split('\n').map(x => x.trim()).filter(Boolean))];

const now = new Date();
const S = {
  user: undefined, who: null, tab: ls.get('conti.tab') || 'home', year: now.getFullYear(), month: now.getMonth(), cat: '', q: '',
  docs: {}, cfg: null, profile: null, shared: [], ready: false, sync: 'wait', sheet: null, modal: null, err: ''
};
if (!['home', 'mesi', 'cat', 'chk'].includes(S.tab)) S.tab = 'home';
let B = null, C = null, unsubs = [], sharedUnsub = null, profUnsub = null;
const readOnly = () => !S.user || S.who !== S.user.uid;

/* ---------------------------------------------------------------- dati */
function applyCfg() {
  const c = S.cfg || {};
  ACC = (c.acc && c.acc.length === 4) ? c.acc : DEF.acc;
  CATS = c.cats || DEF.cats; INC = c.inc || DEF.inc;
}
function calc(year) {
  const yrs = new Set(Object.keys(S.docs).map(i => +i.slice(0, 4))); yrs.add(year);
  const cfg = S.cfg || {};
  const baseY = cfg.startYear || Math.min(...yrs);
  const startY = Math.min(baseY, ...yrs);
  const open = (cfg.open && cfg.open.length === 4) ? cfg.open.map(Number) : [0, 0, 0, 0];
  let bal = startY === baseY ? open.slice() : [0, 0, 0, 0];
  const M = Array.from({ length: 12 }, (_, i) => ({ i, rows: [], ent: 0, exp: 0, byCat: {}, open: null, close: null }));
  const cm = {}; let badCat = 0, badDate = 0;
  for (let y = startY; y <= year; y++) {
    for (let m = 0; m < 12; m++) {
      const d = S.docs[`${y}-${pad(m + 1)}`];
      const rows = [];
      if (d && d.rows) for (const id in d.rows) {
        const r = d.rows[id]; if (!r || r.x) continue;
        rows.push({ id, y, m, d: r.d || '', c: r.c || '', t: r.t || '', o: r.o || 0, v: [0, 1, 2, 3].map(k => +((r.v || [])[k]) || 0) });
      }
      const mo = y === year ? M[m] : null;
      if (mo) { mo.open = bal.slice(); mo.rows = rows; }
      for (const r of rows) {
        for (let k = 0; k < 4; k++) bal[k] += r.v[k];
        if (!mo) continue;
        const s = sum(r.v), isInc = INC.includes(r.c);
        if (!isInc && !CATS.includes(r.c)) badCat++;
        const dd = parseD(r.d); if (!dd || dd.getFullYear() !== year || dd.getMonth() !== m) badDate++;
        if (isInc) mo.ent += s; else mo.exp -= s;
        const val = isInc ? s : -s;
        mo.byCat[r.c] = (mo.byCat[r.c] || 0) + val;
        (cm[r.c] = cm[r.c] || Array(12).fill(0))[m] += val;
      }
      if (mo) { mo.close = bal.slice(); mo.utile = mo.ent - mo.exp; mo.has = rows.length > 0; }
    }
  }
  const ent = sum(M.map(m => m.ent)), exp = sum(M.map(m => m.exp));
  const catTot = {}; for (const k in cm) catTot[k] = sum(cm[k]);
  const openT = sum(M[0].open);
  return { rows: M.flatMap(m => m.rows), M, cm, catTot, ent, exp, utile: ent - exp, open: M[0].open, openT, close: bal, closeT: sum(bal), badCat, badDate, last: M.reduce((a, m) => m.has ? m.i : a, -1) };
}
function yearList() {
  const ys = new Set(Object.keys(S.docs).filter(k => Object.keys((S.docs[k] || {}).rows || {}).length).map(i => +i.slice(0, 4)));
  ys.add(now.getFullYear()); ys.add(S.year);
  return [...ys].sort();
}
function startWatch() {
  unsubs.forEach(f => { try { f(); } catch (e) { } }); unsubs = [];
  S.ready = false; S.docs = {}; S.cfg = null; S.err = ''; render();
  const id = S.who;
  unsubs.push(B.watchMonths(id, (docs, cache) => {
    S.docs = docs; S.ready = true; S.sync = cache ? 'wait' : 'ok';
    if (document.activeElement && document.activeElement.id === 'q') { C = calc(S.year); renderList(); } else render();
  }, e => { S.sync = 'off'; S.err = (e && e.code === 'permission-denied') ? 'Non hai accesso a questi conti.' : 'Connessione al database non riuscita.'; S.ready = true; render(); }));
  unsubs.push(B.watchConfig(id, c => { S.cfg = c; applyCfg(); render(); }));
}

/* ---------------------------------------------------------------- helpers UI */
function toast(msg) {
  const h = document.getElementById('toastHost'); h.innerHTML = '<div class="toast" role="status">' + esc(msg) + '</div>';
  clearTimeout(toast.t); toast.t = setTimeout(() => h.innerHTML = '', 3200);
}
const ICON = {
  home: '<path d="M4 11.5 12 5l8 6.5V19a1 1 0 0 1-1 1h-4v-5H9v5H5a1 1 0 0 1-1-1z"/>',
  mesi: '<rect x="4" y="5" width="16" height="15" rx="3"/><path d="M8 3v4M16 3v4M4 10h16"/>',
  cat: '<rect x="4" y="4" width="7" height="7" rx="2"/><rect x="13" y="4" width="7" height="7" rx="2"/><rect x="4" y="13" width="7" height="7" rx="2"/><rect x="13" y="13" width="7" height="7" rx="2"/>',
  chk: '<circle cx="12" cy="12" r="8.5"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>'
};
const svg = k => '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICON[k] + '</svg>';
const TABS = [['home', 'Panoramica'], ['mesi', 'Mesi'], ['cat', 'Categorie'], ['chk', 'Controlli']];
function renderNav() {
  const on = !!S.user;
  document.getElementById('tabs').innerHTML = on ? TABS.map(([k, l]) => '<button class="tab" role="tab" data-tab="' + k + '" aria-selected="' + (S.tab === k) + '">' + svg(k) + '<span>' + l + '</span></button>').join('') : '';
  document.getElementById('side').innerHTML = on ? '<div class="brand">Conti<small>' + esc(S.user.name || S.user.email) + '</small></div>' + TABS.map(([k, l]) => '<button class="sidenav" role="tab" data-tab="' + k + '" aria-selected="' + (S.tab === k) + '">' + svg(k) + '<span>' + l + '</span></button>').join('') : '';
  document.getElementById('fab').hidden = !on || readOnly() || !S.ready;
}
function go(tab) { S.tab = tab; ls.set('conti.tab', tab); renderNav(); render(); window.scrollTo({ top: 0 }); }
function head(title) {
  const st = S.sync === 'ok' ? '<span class="sync" title="Sincronizzato"><i></i><em>Sincronizzato</em></span>' : S.sync === 'off' ? '<span class="sync off" title="Non connesso"><i></i><em>Non connesso</em></span>' : '<span class="sync wait" title="Carico"><i></i><em>Carico…</em></span>';
  const ys = yearList(), maxY = Math.max(...ys);
  const who = S.shared.length ? '<select id="who" class="who" aria-label="Di chi sono i conti">' + '<option value="' + esc(S.user.uid) + '"' + (!readOnly() ? ' selected' : '') + '>I miei conti</option>' + S.shared.map(p => '<option value="' + esc(p.uid) + '"' + (S.who === p.uid ? ' selected' : '') + '>Conti di ' + esc(p.name) + '</option>').join('') + '</select>' : '';
  return '<header class="top"><h1>' + esc(title) + '<span class="sub">Anno ' + S.year + '</span></h1>' + who +
    '<div class="yr" role="group" aria-label="Anno"><button id="yprev" aria-label="Anno precedente">‹</button><b class="num">' + S.year + '</b><button id="ynext" aria-label="Anno successivo"' + (S.year >= maxY + 1 ? ' disabled' : '') + '>›</button></div>' +
    st + (readOnly() ? '' : '<button class="iconbtn" id="gear" aria-label="Impostazioni">' + svg('gear') + '</button>') + '</header>' +
    (readOnly() && S.user ? '<div class="banner" style="margin-top:4px">Stai guardando i conti di ' + esc((S.shared.find(p => p.uid === S.who) || {}).name || 'un’altra persona') + ' in sola lettura.</div>' : '') +
    (S.err ? '<div class="banner bad">' + esc(S.err) + '</div>' : '');
}

/* ---------------------------------------------------------------- grafici */
function niceScale(lo, hi) {
  if (hi <= lo) hi = lo + 1000;
  const span = hi - lo, raw = span / 4, p = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map(s => s * p).find(s => s >= raw);
  const a = Math.floor(lo / step) * step, b = Math.ceil(hi / step) * step, t = [];
  for (let v = a; v <= b + step / 2; v += step) t.push(Math.round(v * 100) / 100);
  return { lo: a, hi: b, ticks: t };
}
const fK = v => Math.abs(v) >= 1000 ? num0.format(v / 1000) + ' k' : num0.format(v);
function drawBars(el, cfg) {
  const W = Math.max(260, el.clientWidth || 320), H = cfg.h || 210, L = 36, R = 4, T = 8, Bm = 22;
  const all = cfg.series.flatMap(s => s.values);
  const sc = niceScale(Math.min(0, ...all), Math.max(0, ...all));
  const iw = W - L - R, ih = H - T - Bm, gw = iw / 12;
  const y = v => T + ih - (v - sc.lo) / (sc.hi - sc.lo) * ih;
  const n = cfg.series.length, bw = Math.min(n > 1 ? 13 : 22, gw * (n > 1 ? .34 : .55)), gap = 2;
  let g = '';
  for (const t of sc.ticks) g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '" stroke="var(--line)" stroke-width="' + (t === 0 ? 1.4 : 1) + '"/><text x="' + (L - 6) + '" y="' + (y(t) + 3.5) + '" text-anchor="end">' + fK(t) + '</text>';
  for (let i = 0; i < 12; i++) {
    const cx = L + gw * i + gw / 2, tw = n * bw + (n - 1) * gap;
    g += '<text x="' + cx + '" y="' + (H - 6) + '" text-anchor="middle">' + MS[i] + '</text>';
    cfg.series.forEach((s, k) => {
      const v = s.values[i]; if (!v) return;
      const x = cx - tw / 2 + k * (bw + gap), y0 = y(0), y1 = y(v), top = Math.min(y0, y1), h = Math.max(2, Math.abs(y0 - y1));
      const col = s.color ? s.color : (v >= 0 ? 'var(--pos)' : 'var(--neg)');
      g += '<rect x="' + x + '" y="' + top + '" width="' + bw + '" height="' + h + '" rx="' + Math.min(5, bw / 2) + '" fill="' + col + '"/>';
    });
    g += '<rect data-i="' + i + '" x="' + (L + gw * i) + '" y="' + T + '" width="' + gw + '" height="' + ih + '" fill="transparent"/>';
  }
  el.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="' + esc(cfg.label) + '">' + g + '</svg><div class="readout" aria-live="polite">' + esc(cfg.idle || '') + '</div>';
  const ro = el.querySelector('.readout'), sv = el.querySelector('svg');
  const show = ev => { const r = ev.target.closest('rect[data-i]'); if (!r) return; const i = +r.dataset.i; ro.innerHTML = '<b>' + MN[i] + '</b>' + cfg.read(i); };
  sv.addEventListener('pointermove', show); sv.addEventListener('click', show);
}
function drawDonut(el, a, b) {
  const R = 58, Cc = 2 * Math.PI * R, tot = a + b, pa = tot > 0 ? Math.max(0, Math.min(1, a / tot)) : 0;
  const seg = (len, off, col) => '<circle cx="80" cy="80" r="' + R + '" fill="none" stroke="' + col + '" stroke-width="20" stroke-dasharray="' + Math.max(0, len - 2) + ' ' + Cc + '" stroke-dashoffset="' + (-off) + '" transform="rotate(-90 80 80)"/>';
  const ring = tot > 0 ? seg(pa * Cc, 0, 'var(--neg)') + seg((1 - pa) * Cc, pa * Cc, 'var(--pos)') : '<circle cx="80" cy="80" r="' + R + '" fill="none" stroke="var(--surface2)" stroke-width="20"/>';
  el.innerHTML = '<svg viewBox="0 0 160 160" width="160" height="160" role="img" aria-label="Quota di uscite e di utile sulle entrate">' + ring + '<text x="80" y="78" text-anchor="middle" style="fill:var(--ink);font:700 22px var(--f-display)">' + (tot > 0 ? pct.format(1 - pa) : '–') + '</text><text x="80" y="96" text-anchor="middle" style="font-size:11px">risparmiato</text></svg>';
}

/* ---------------------------------------------------------------- viste */
function rankHtml(tot, base, inMonth) {
  const items = CATS.concat(Object.keys(tot).filter(k => !CATS.includes(k) && !INC.includes(k))).map(k => [k, tot[k] || 0]).filter(x => Math.abs(x[1]) > 0.004).sort((a, b) => b[1] - a[1]);
  if (!items.length) return '<div class="empty-state">Nessuna spesa registrata.</div>';
  const mx = Math.max(...items.map(i => Math.abs(i[1])));
  return items.map(([k, v]) => '<button class="rk" data-rank="' + esc(k) + '" aria-pressed="' + (inMonth && S.cat === k) + '"><span class="n">' + esc(k) + '</span><span class="a num">' + fE(v) + (base > 0 ? '<span class="p">' + pct.format(v / base) + '</span>' : '') + '</span><span class="t"><span style="width:' + Math.max(2, Math.abs(v) / mx * 100) + '%"></span></span></button>').join('');
}
function welcome() {
  return '<section class="card" style="margin-top:12px;text-align:center;padding:30px 20px"><h2 style="font-size:22px">Benvenuto</h2><p class="hint" style="margin:8px auto 18px;max-width:34ch">Parti dal tuo file Excel oppure inizia da zero impostando i saldi dei conti.</p><div class="actions" style="justify-content:center;flex-wrap:wrap"><button class="btn pri" style="flex:0 1 auto;padding-inline:22px" data-act="import">Importa il file Excel</button><button class="btn" style="flex:0 1 auto;padding-inline:22px" data-act="settings">Inizia da zero</button></div></section>';
}
function homeView(c) {
  const lastM = c.last < 0 ? null : c.M[c.last];
  const goal = (S.cfg && S.cfg.goal) || DEF.goal, target = c.openT + goal;
  const prog = Math.max(0, Math.min(1, c.utile / goal));
  const [ip, dp] = eur.format(r2(c.closeT)).replace(/\s?€/, '').split(',');
  let h = head('Panoramica');
  h += '<section class="card hero" style="margin-top:6px"><div class="eyebrow">Saldo totale' + (lastM ? ' · aggiornato a ' + MN[c.last].toLowerCase() : '') + '</div>' +
    '<div class="big num">' + ip + '<small>,' + (dp || '00') + ' €</small></div>' +
    '<div class="row">' + ACC.map((a, k) => '<span class="chip num">' + esc(a) + ' <b>' + fE0(c.close[k]) + '</b></span>').join('') + '</div>' +
    '<div class="goal"><div class="eyebrow">Obiettivo ' + S.year + '</div><div style="margin:6px 0 8px;font:700 18px var(--f-display)" class="num">' + fE0(target) + ' <span style="font-weight:500;font-size:14px;opacity:.8">a fine anno</span></div>' +
    '<div class="bar" role="progressbar" aria-valuenow="' + Math.round(prog * 100) + '" aria-valuemin="0" aria-valuemax="100"><span style="width:' + (prog * 100).toFixed(1) + '%"></span></div>' +
    '<p class="num">' + (c.utile >= goal ? 'Obiettivo raggiunto: +' + fE0(c.utile - goal) : 'Messi da parte ' + fE0(c.utile) + ' su ' + fE0(goal) + ' · ne mancano ' + fE0(goal - c.utile)) + '</p></div></section>';
  if (!Object.keys(S.docs).length) {
    return h + (readOnly() ? '<div class="empty-state"><b>Ancora nessun dato</b>Questi conti sono vuoti.</div>' : welcome());
  }
  const inz = c.catTot['Casa Inzago'] || 0;
  h += '<div class="kpis" style="margin-top:14px">' +
    '<div class="kpi"><div class="eyebrow">Entrate</div><div class="v num pos">' + fE0(c.ent) + '</div><div class="s">da gennaio</div></div>' +
    '<div class="kpi"><div class="eyebrow">Uscite</div><div class="v num neg">' + fE0(c.exp) + '</div><div class="s num">' + (c.ent ? pct.format(c.exp / c.ent) + ' delle entrate' : '–') + '</div></div>' +
    '<div class="kpi"><div class="eyebrow">Utile</div><div class="v num ' + (c.utile >= 0 ? 'pos' : 'neg') + '">' + fE0(c.utile) + '</div><div class="s num">' + (c.ent ? pct.format(c.utile / c.ent) + ' delle entrate' : '–') + '</div></div>' +
    (inz ? '<div class="kpi"><div class="eyebrow">Senza casa Inzago</div><div class="v num">' + fE0(c.exp - inz) + '</div><div class="s num">' + (c.ent ? pct.format((c.exp - inz) / c.ent) + ' delle entrate' : '–') + '</div></div>' :
      '<div class="kpi"><div class="eyebrow">Media uscite al mese</div><div class="v num">' + fE0(c.last >= 0 ? c.exp / (c.last + 1) : 0) + '</div><div class="s">sui mesi con dati</div></div>') + '</div>';
  h += '<div class="grid two-col">' +
    '<section class="card"><h2>Entrate e uscite</h2><p class="hint">Mese per mese. Tocca una colonna per i valori.</p><div class="legend"><span><i style="background:var(--pos)"></i>Entrate</span><span><i style="background:var(--neg)"></i>Uscite</span></div><div class="chart" id="ch1"></div></section>' +
    '<section class="card"><h2>Utile o perdita</h2><p class="hint">Quanto resta ogni mese dopo le uscite.</p><div class="chart" id="ch2"></div></section>' +
    '<section class="card"><h2>Dove va il reddito</h2><p class="hint">Uscite e utile in percentuale sulle entrate dell’anno.</p><div class="donut"><div id="ch3"></div><div class="dl num"><div><i style="background:var(--neg)"></i><span>Uscite <b>' + (c.ent ? pct.format(c.exp / c.ent) : '–') + '</b><br><span style="color:var(--muted);font-size:13px">' + fE0(c.exp) + '</span></span></div><div><i style="background:var(--pos)"></i><span>Utile <b>' + (c.ent ? pct.format(c.utile / c.ent) : '–') + '</b><br><span style="color:var(--muted);font-size:13px">' + fE0(c.utile) + '</span></span></div></div></div></section>' +
    '<section class="card"><h2>Uscite per categoria</h2><p class="hint">Totale ' + S.year + '. Tocca per vedere le voci.</p><div class="ranks">' + rankHtml(c.catTot, c.exp, false) + '</div></section></div>';
  const bad = checks(c).filter(x => x.st !== 'ok').length;
  h += '<button class="banner ' + (bad ? 'bad' : 'ok') + '" style="display:block;width:100%;border:0;text-align:left" data-tab="chk">' + (bad ? bad + (bad === 1 ? ' controllo da guardare' : ' controlli da guardare') + ' →' : 'Tutti i controlli sono ok →') + '</button>';
  return h;
}
function monthView(c) {
  const m = c.M[S.month];
  let h = head('Mesi');
  h += '<div class="months" id="mchips" role="group" aria-label="Mese">' + MS.map((l, i) => '<button class="mchip' + (c.M[i].has ? '' : ' empty') + '" data-m="' + i + '" aria-pressed="' + (S.month === i) + '">' + l + '</button>').join('') + '</div>';
  h += '<section class="card" style="margin-top:6px"><div class="eyebrow">' + MN[S.month] + ' ' + S.year + '</div>' +
    '<div style="margin-top:6px;font:700 34px/1.1 var(--f-display);letter-spacing:-.02em" class="num">' + fE(m.utile) + '</div><div style="color:var(--muted);font-size:13px">' + (m.utile >= 0 ? 'messi da parte nel mese' : 'spesi in più delle entrate') + '</div>' +
    '<div class="msum num"><div><span>Entrate</span><b class="pos">' + fE0(m.ent) + '</b></div><div><span>Uscite</span><b class="neg">' + fE0(m.exp) + '</b></div><div><span>Saldo a fine mese</span><b>' + fE0(sum(m.close)) + '</b></div></div>' +
    '<div class="acc4 num">' + ACC.map((a, k) => '<div><span>' + esc(a) + '</span><b>' + fE(m.close[k]) + '</b></div>').join('') + '</div></section>';
  h += '<div class="grid two-col"><section class="card"><h2>Per categoria</h2><p class="hint">Spese del mese. Tocca per filtrare le voci.</p><div class="ranks">' + rankHtml(m.byCat, m.exp, true) + '</div></section>' +
    '<section class="card"><h2>Movimenti</h2><div class="filters"><input class="field" id="q" type="search" placeholder="Cerca nei dettagli" value="' + esc(S.q) + '" autocomplete="off"><select class="field" id="fc" aria-label="Categoria"><option value="">Tutte le categorie</option>' + INC.concat(CATS).map(k => '<option' + (S.cat === k ? ' selected' : '') + '>' + esc(k) + '</option>').join('') + '</select></div><div id="txl"></div></section></div>';
  return h;
}
function txList(c) {
  const m = c.M[S.month], q = S.q.trim().toLowerCase();
  const rows = m.rows.filter(r => (!S.cat || r.c === S.cat) && (!q || (r.t + ' ' + r.c).toLowerCase().includes(q)));
  rows.sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : a.o - b.o);
  if (!m.has) return '<div class="empty-state"><b>Nessun movimento a ' + MN[S.month].toLowerCase() + '</b>' + (readOnly() ? '' : 'Tocca “Nuova voce” per aggiungere il primo.') + '</div>';
  if (!rows.length) return '<div class="empty-state"><b>Nessun risultato</b>Cambia ricerca o categoria.</div>';
  const days = [];
  for (const r of rows) { const l = days[days.length - 1]; if (l && l.d === r.d) l.r.push(r); else days.push({ d: r.d, r: [r] }); }
  days.reverse();
  return days.map(g => {
    const dd = parseD(g.d), tot = sum(g.r.map(r => sum(r.v)));
    return '<div class="day"><span>' + (dd ? dayFmt.format(dd) : 'Senza data') + '</span><span class="num">' + (Math.abs(tot) > 0.004 ? sgn(tot) : '') + '</span></div>' + g.r.slice().reverse().map(txRow).join('');
  }).join('');
}
function txRow(r) {
  const nz = r.v.map((x, k) => [x, k]).filter(p => Math.abs(p[0]) > 0.004);
  const s = sum(r.v), mv = Math.abs(s) < 0.005 && nz.length >= 2, isIn = !mv && s > 0;
  const acc = nz.map(p => ACC[p[1]]).join(' + ');
  const mono = (r.c.split(/[\s/]+/).map(w => w[0]).join('').slice(0, 2) || '?').toUpperCase();
  const amt = mv ? Math.max(...r.v) : s;
  return '<button class="tx ' + (isIn ? 'in' : '') + (mv ? ' mv' : '') + '" data-edit="' + esc(r.id) + '"><span class="mono">' + esc(mono) + '</span><span style="min-width:0"><div class="t1">' + esc(r.t || r.c) + '</div><div class="t2">' + (mv ? 'Spostamento · ' : esc(r.c) + ' · ') + esc(acc || '–') + '</div></span><span class="am num ' + (isIn ? 'pos' : '') + '">' + (mv ? '⇄ ' + fE(amt) : sgn(s)) + '</span></button>';
}
function catView(c) {
  const exps = CATS.concat(Object.keys(c.cm).filter(k => !CATS.includes(k) && !INC.includes(k)));
  const maxV = Math.max(1, ...exps.flatMap(k => c.cm[k] || [0]).map(Math.abs));
  const cell = (k, i, v, isInc) => {
    if (Math.abs(v) < 0.005) return '<td class="z c" data-cell="' + esc(k) + '|' + i + '">·</td>';
    const a = Math.min(1, Math.sqrt(Math.abs(v) / maxV)), col = isInc ? 'var(--pos)' : 'var(--neg)';
    return '<td class="c num" data-cell="' + esc(k) + '|' + i + '" title="' + esc(fE(v)) + '" style="background:color-mix(in srgb,' + col + ' ' + Math.round(a * 38) + '%,transparent)">' + num0.format(Math.round(v)) + '</td>';
  };
  const row = (k, isInc) => { const a = c.cm[k] || Array(12).fill(0); return '<tr><td>' + esc(k) + '</td>' + a.map((v, i) => cell(k, i, v, isInc)).join('') + '<td class="tot num">' + fE(c.catTot[k] || 0) + '</td></tr>'; };
  const sumRow = (label, arr, cls) => '<tr class="sum"><td>' + label + '</td>' + arr.map(v => '<td class="num ' + (cls ? (v >= 0 ? 'pos' : 'neg') : '') + '">' + (Math.abs(v) < 0.005 ? '·' : num0.format(Math.round(v))) + '</td>').join('') + '<td class="num ' + (cls ? (sum(arr) >= 0 ? 'pos' : 'neg') : '') + '">' + fE(sum(arr)) + '</td></tr>';
  let t = '<table class="heat"><thead><tr><th>' + S.year + '</th>' + MS.map(x => '<th>' + x + '</th>').join('') + '<th>Totale</th></tr></thead><tbody>';
  t += '<tr class="sec"><td colspan="14">Entrate</td></tr>' + INC.map(k => row(k, true)).join('') + sumRow('Totale entrate', c.M.map(m => m.ent));
  t += '<tr class="sec"><td colspan="14">Uscite</td></tr>' + exps.map(k => row(k, false)).join('') + sumRow('Totale uscite', c.M.map(m => m.exp));
  t += '<tr class="sec"><td colspan="14">Risultato</td></tr>' + sumRow('Utile / (perdita)', c.M.map(m => m.utile), true) + '</tbody></table>';
  return head('Categorie') + '<section class="card" style="margin-top:6px"><h2>Totale ' + S.year + '</h2><p class="hint">Importi in euro, arrotondati nelle celle. Tocca una cella per aprire le voci di quel mese.</p><div class="tbl-wrap">' + t + '</div></section>';
}
function checks(c) {
  const eq = (a, b) => Math.abs(a - b) < 0.005;
  const entCat = sum(INC.map(k => c.catTot[k] || 0)), expCat = sum(CATS.map(k => c.catTot[k] || 0));
  const neg = c.close.map((v, k) => [v, k]).filter(p => p[0] < -0.005);
  return [
    { n: 'Entrate: mesi e categorie coincidono', st: eq(c.ent, entCat) ? 'ok' : 'bad', d: 'Somma dei mesi ' + fE(c.ent) + ' · somma delle categorie ' + fE(entCat) },
    { n: 'Uscite: mesi e categorie coincidono', st: eq(c.exp, expCat) ? 'ok' : 'bad', d: 'Somma dei mesi ' + fE(c.exp) + ' · somma delle categorie ' + fE(expCat) },
    { n: 'Utile uguale alla variazione dei saldi', st: eq(c.utile, c.closeT - c.openT) ? 'ok' : 'bad', d: 'Utile ' + fE(c.utile) + ' · saldo finale meno iniziale ' + fE(c.closeT - c.openT) },
    { n: 'Saldo finale = saldo iniziale + utile', st: eq(c.openT + c.utile, c.closeT) ? 'ok' : 'bad', d: fE(c.openT) + ' + ' + fE(c.utile) + ' = ' + fE(c.openT + c.utile) + ' · saldo dei conti ' + fE(c.closeT) },
    { n: 'Categorie valide', st: c.badCat === 0 ? 'ok' : 'bad', d: c.badCat === 0 ? 'Tutte le voci usano una categoria dell’elenco.' : c.badCat + ' voci con una categoria non presente nell’elenco (Impostazioni).' },
    { n: 'Date nel mese giusto', st: c.badDate === 0 ? 'ok' : 'warn', d: c.badDate === 0 ? 'Ogni voce ha una data dentro il proprio mese.' : c.badDate + ' voci con data mancante o fuori dal mese.' },
    { n: 'Nessun conto sotto zero', st: neg.length ? 'warn' : 'ok', d: neg.length ? neg.map(p => ACC[p[1]] + ' ' + fE(p[0])).join(' · ') : 'Tutti i conti sono in positivo.' }
  ];
}
function chkView(c) {
  const ck = checks(c), bad = ck.filter(x => x.st === 'bad').length, warn = ck.filter(x => x.st === 'warn').length;
  const lab = { ok: 'Ok', bad: 'No', warn: 'Attenzione' };
  let h = head('Controlli');
  h += '<div class="banner ' + (bad ? 'bad' : warn ? '' : 'ok') + '">' + (bad ? 'Ci sono numeri che non tornano.' : warn ? 'Tutto torna, con qualche avviso.' : 'Tutto torna.') + '</div>';
  h += '<section class="card" style="margin-top:14px"><h2>Chiusura dei conti</h2><p class="hint">Scarica il file Excel con le formule, nello stesso schema del tuo foglio. Puoi modificarlo da PC e ricaricarlo dalle Impostazioni.</p>' +
    '<div class="two form"><label>Mese<select class="field" id="xm">' + MN.map((m, i) => '<option value="' + i + '"' + (i === S.month ? ' selected' : '') + '>' + m + '</option>').join('') + '</select></label></div>' +
    '<div class="actions" style="flex-wrap:wrap"><button class="btn" data-act="xmonth">Scarica il mese</button><button class="btn pri" data-act="xyear">Scarica tutto il ' + S.year + '</button></div></section>';
  h += '<section class="card" style="margin-top:14px"><h2>Verifiche di quadratura</h2><p class="hint">Gli stessi confronti del foglio Totale, ricalcolati sulle voci inserite.</p>' +
    ck.map(x => '<div class="chk"><span class="pill ' + (x.st === 'ok' ? '' : x.st) + '">' + lab[x.st] + '</span><b>' + esc(x.n) + '</b><span class="d num">' + esc(x.d) + '</span></div>').join('') + '</section>';
  h += '<section class="card" style="margin-top:14px"><h2>Saldi dei conti</h2><p class="hint">Apertura al 1 gennaio ' + S.year + ' e saldo attuale.</p><div class="tbl-wrap"><table class="heat" style="min-width:420px"><thead><tr><th>Conto</th><th>Apertura</th><th>Movimenti</th><th>Saldo</th></tr></thead><tbody>' +
    ACC.map((a, k) => '<tr><td>' + esc(a) + '</td><td class="num">' + fE(c.open[k]) + '</td><td class="num">' + sgn(c.close[k] - c.open[k]) + '</td><td class="num tot">' + fE(c.close[k]) + '</td></tr>').join('') +
    '<tr class="sum"><td>Totale</td><td class="num">' + fE(c.openT) + '</td><td class="num">' + sgn(c.closeT - c.openT) + '</td><td class="num">' + fE(c.closeT) + '</td></tr></tbody></table></div></section>';
  return h;
}
function loginView() {
  const demo = B && B.kind === 'local';
  return '<div class="login"><div class="logo">' + svg('home').replace('<svg', '<svg class="lg"') + '</div><h1>Conti</h1><p>Entrate, uscite e saldi di casa, sempre aggiornati su telefono e PC. Ognuno ha i propri conti, e si può guardare quelli dell’altro in sola lettura.</p>' +
    (demo ? '<button class="btn pri" id="login">Entra in modalità demo</button><p class="small">Firebase non è ancora configurato: i dati restano solo su questo dispositivo. Segui il README per collegarlo.</p>' :
      '<button class="btn pri" id="login">Accedi con Google</button><p class="small">Usiamo solo il tuo indirizzo email per riconoscerti.</p>') + '<div class="err" id="lerr" role="alert"></div></div>';
}

/* ---------------------------------------------------------------- render */
function render() {
  const main = document.getElementById('main');
  renderNav();
  if (S.user === undefined) { main.innerHTML = '<div class="empty-state"><b>Carico…</b></div>'; return; }
  if (!S.user) { main.innerHTML = loginView(); return; }
  if (!S.ready) { main.innerHTML = head('Conti') + '<div class="empty-state"><b>Carico i conti…</b>Un attimo, sto leggendo i dati.</div>'; return; }
  C = calc(S.year);
  main.innerHTML = S.tab === 'home' ? homeView(C) : S.tab === 'mesi' ? monthView(C) : S.tab === 'cat' ? catView(C) : chkView(C);
  drawCharts();
  if (S.tab === 'mesi') renderList();
}
function renderList() { const el = document.getElementById('txl'); if (el) el.innerHTML = txList(C); }
function drawCharts() {
  if (S.tab !== 'home' || !C) return;
  const e1 = document.getElementById('ch1'), e2 = document.getElementById('ch2'), e3 = document.getElementById('ch3');
  if (e1) drawBars(e1, { label: 'Entrate e uscite per mese', series: [{ values: C.M.map(m => m.ent), color: 'var(--pos)' }, { values: C.M.map(m => m.exp), color: 'var(--neg)' }], read: i => '<span>Entrate <b class="num pos">' + fE(C.M[i].ent) + '</b></span><span>Uscite <b class="num neg">' + fE(C.M[i].exp) + '</b></span>', idle: 'Tocca un mese per vedere gli importi.' });
  if (e2) drawBars(e2, { label: 'Utile o perdita per mese', series: [{ values: C.M.map(m => m.has ? m.utile : 0) }], h: 190, read: i => '<span>Utile <b class="num ' + (C.M[i].utile >= 0 ? 'pos' : 'neg') + '">' + sgn(C.M[i].utile) + '</b></span>', idle: 'Tocca un mese per vedere l’importo.' });
  if (e3) drawDonut(e3, C.exp, Math.max(0, C.utile));
}
let rs; window.addEventListener('resize', () => { clearTimeout(rs); rs = setTimeout(drawCharts, 120); });

/* ---------------------------------------------------------------- eventi principali */
function setYear(y) { S.year = y; S.month = y === now.getFullYear() ? now.getMonth() : 0; S.cat = ''; S.q = ''; render(); }
function jump(tab, month, cat) { S.tab = tab; S.month = month; S.cat = cat; S.q = ''; ls.set('conti.tab', tab); renderNav(); render(); window.scrollTo({ top: 0 }); }
document.addEventListener('click', e => { const t = e.target.closest('[data-tab]'); if (t && t.dataset.tab) go(t.dataset.tab); });
document.getElementById('fab').addEventListener('click', () => openSheet(null));
const mainEl = document.getElementById('main');
mainEl.addEventListener('click', async e => {
  const id = e.target.id || (e.target.closest('button') || {}).id;
  if (id === 'login') { const er = document.getElementById('lerr'); try { await B.signIn(); } catch (x) { er.textContent = 'Accesso non riuscito: ' + (x.code || x.message); } return; }
  if (id === 'gear') return openSettings();
  if (id === 'yprev') return setYear(S.year - 1);
  if (id === 'ynext') return setYear(S.year + 1);
  const act = (e.target.closest('[data-act]') || {}).dataset;
  if (act) {
    if (act.act === 'import') return pickFile();
    if (act.act === 'settings') return openSettings();
    if (act.act === 'xmonth') return doExport(+document.getElementById('xm').value);
    if (act.act === 'xyear') return doExport(null);
  }
  const mc = e.target.closest('[data-m]'); if (mc) { S.month = +mc.dataset.m; S.cat = ''; S.q = ''; return render(); }
  const rk = e.target.closest('[data-rank]');
  if (rk) { const k = rk.dataset.rank; if (S.tab === 'home') return jump('mesi', C.last < 0 ? S.month : C.last, k); S.cat = S.cat === k ? '' : k; return render(); }
  const ce = e.target.closest('[data-cell]'); if (ce) { const [k, i] = ce.dataset.cell.split('|'); return jump('mesi', +i, k); }
  const ed = e.target.closest('[data-edit]'); if (ed && !readOnly()) { const r = C.rows.find(x => x.id === ed.dataset.edit && x.m === S.month); if (r) openSheet(r); }
});
mainEl.addEventListener('input', e => { if (e.target.id === 'q') { S.q = e.target.value; renderList(); } });
mainEl.addEventListener('change', e => {
  if (e.target.id === 'fc') { S.cat = e.target.value; render(); }
  if (e.target.id === 'who') { S.who = e.target.value; S.cat = ''; S.q = ''; startWatch(); }
});

/* ---------------------------------------------------------------- scheda voce */
function openSheet(r) {
  const edit = !!r, nz = edit ? r.v.filter(x => Math.abs(x) > 0.004).length : 0, s = edit ? sum(r.v) : 0;
  const t = todayIso();
  const F = {
    edit, id: edit ? r.id : null, docId: edit ? `${r.y}-${pad(r.m + 1)}` : null, o: edit ? r.o : null,
    type: edit ? (nz >= 2 && Math.abs(s) < 0.005 ? 'mov' : s > 0 ? 'in' : 'out') : 'out', split: edit && nz >= 2,
    amount: edit && nz === 1 ? Math.abs(s) : '', acc: edit && nz === 1 ? r.v.findIndex(x => Math.abs(x) > 0.004) : 1, from: 1, to: 0,
    date: edit ? r.d : (t.startsWith(S.year + '-') ? t : S.year + '-' + pad(S.month + 1) + '-01'),
    cat: edit ? r.c : '', det: edit ? r.t : '', v: edit ? r.v.slice() : [0, 0, 0, 0]
  };
  if (F.type === 'mov' && edit) { F.from = r.v.findIndex(x => x < -0.004); F.to = r.v.findIndex(x => x > 0.004); }
  S.sheet = F; drawSheet();
}
const opt = (arr, sel) => arr.map((x, i) => '<option value="' + i + '"' + (i === sel ? ' selected' : '') + '>' + esc(x) + '</option>').join('');
function drawSheet() {
  const F = S.sheet, h = document.getElementById('sheetHost');
  if (!F) { h.innerHTML = ''; return; }
  const catSel = F.cat || (F.type === 'in' ? INC[0] : '');
  const extra = F.cat && !CATS.includes(F.cat) && !INC.includes(F.cat) ? '<option selected>' + esc(F.cat) + '</option>' : '';
  const catOpts = '<option value="">Scegli…</option><optgroup label="Entrate">' + INC.map(k => '<option' + (k === catSel ? ' selected' : '') + '>' + esc(k) + '</option>').join('') + '</optgroup><optgroup label="Uscite">' + CATS.map(k => '<option' + (k === catSel ? ' selected' : '') + '>' + esc(k) + '</option>').join('') + extra + '</optgroup>';
  let body = '';
  if (F.split) {
    body = '<div class="form"><div class="hint" style="color:var(--muted);font-size:13px">Importo per conto. Negativo se esce, positivo se entra.</div><div class="two">' + ACC.map((a, k) => '<label>' + esc(a) + '<input class="field num" id="s' + k + '" inputmode="decimal" value="' + (F.v[k] ? String(F.v[k]).replace('.', ',') : '') + '" placeholder="0,00"></label>').join('') + '</div><button class="link" id="unsplit" type="button">Torna all’importo unico</button></div>';
  } else {
    body = '<input class="amt num" id="amt" inputmode="decimal" placeholder="0,00" aria-label="Importo in euro" value="' + (F.amount !== '' ? String(F.amount).replace('.', ',') : '') + '">' +
      (F.type === 'mov' ? '<div class="two form"><label>Da<select class="field" id="from">' + opt(ACC, F.from) + '</select></label><label>A<select class="field" id="to">' + opt(ACC, F.to) + '</select></label></div>' : '<div class="form"><label>Conto<select class="field" id="acc">' + opt(ACC, F.acc) + '</select></label></div>');
  }
  const title = F.edit ? 'Modifica voce' : 'Nuova voce';
  h.innerHTML = '<div class="scrim" id="scrim"><div class="sheet" role="dialog" aria-modal="true" aria-label="' + title + '"><h2>' + title + '</h2>' +
    (F.split ? '' : '<div class="seg" role="group" aria-label="Tipo"><button type="button" data-t="out" aria-pressed="' + (F.type === 'out') + '">Uscita</button><button type="button" data-t="in" aria-pressed="' + (F.type === 'in') + '">Entrata</button><button type="button" data-t="mov" aria-pressed="' + (F.type === 'mov') + '">Spostamento</button></div>') + body +
    '<div class="form" style="margin-top:12px"><div class="two"><label>Data<input class="field" type="date" id="date" value="' + esc(F.date) + '"></label>' +
    (F.type === 'mov' && !F.split ? '<label>Dettagli<input class="field" id="det" value="' + esc(F.det || '') + '" placeholder="Spostamento" autocomplete="off"></label>' : '<label>Categoria<select class="field" id="cat">' + catOpts + '</select></label>') + '</div>' +
    (F.type === 'mov' && !F.split ? '' : '<label>Dettagli<input class="field" id="det" value="' + esc(F.det || '') + '" placeholder="Es. Iperal, benzina, stipendio" autocomplete="off"></label>') +
    (F.split || F.type === 'mov' ? '' : '<button class="link" id="dosplit" type="button">Dividi su più conti</button>') +
    '<div class="err" id="err" role="alert"></div><div class="actions">' + (F.edit ? '<button class="btn del" id="del" type="button">Elimina</button>' : '') + '<button class="btn" id="cancel" type="button">Annulla</button><button class="btn pri" id="save" type="button">Salva</button></div></div></div></div>';
  const f = h.querySelector('#amt'); if (f && !F.edit) setTimeout(() => { try { f.focus(); } catch (e) { } }, 50);
}
function readSheet() {
  const F = S.sheet, g = id => document.getElementById(id); if (!F) return;
  if (g('date')) F.date = g('date').value; if (g('cat')) F.cat = g('cat').value; if (g('det')) F.det = g('det').value;
  if (g('amt')) F.amount = g('amt').value; if (g('acc')) F.acc = +g('acc').value; if (g('from')) F.from = +g('from').value; if (g('to')) F.to = +g('to').value;
  if (F.split) F.v = [0, 1, 2, 3].map(k => g('s' + k) ? parseAmt(g('s' + k).value) : 0);
}
document.getElementById('sheetHost').addEventListener('click', async e => {
  const F = S.sheet; if (!F) return;
  if (e.target.id === 'scrim' || e.target.id === 'cancel') { S.sheet = null; drawSheet(); return; }
  const tb = e.target.closest('[data-t]');
  if (tb) { readSheet(); F.type = tb.dataset.t; if (F.type === 'in' && !F.cat) F.cat = INC[0]; if (F.type === 'mov') F.cat = 'Varie'; else if (F.cat === 'Varie' && F.det === 'Spostamento') F.cat = ''; drawSheet(); return; }
  if (e.target.id === 'dosplit') { readSheet(); const a = parseAmt(F.amount); F.v = [0, 0, 0, 0]; F.v[F.acc] = F.type === 'in' ? Math.abs(a) : -Math.abs(a); F.split = true; drawSheet(); return; }
  if (e.target.id === 'unsplit') { readSheet(); const nz = F.v.filter(x => Math.abs(x) > 0.004); if (nz.length <= 1) { const k = F.v.findIndex(x => Math.abs(x) > 0.004); F.acc = k < 0 ? 1 : k; F.amount = Math.abs(F.v[F.acc] || 0) || ''; F.type = (F.v[F.acc] || 0) > 0 ? 'in' : 'out'; F.split = false; drawSheet(); } else document.getElementById('err').textContent = 'Più conti con importo: lascia solo un conto per tornare all’importo unico.'; return; }
  if (e.target.id === 'del') {
    if (!F.confirmDel) { F.confirmDel = true; e.target.textContent = 'Conferma'; return; }
    try { await B.delRow(S.who, F.docId, F.id); S.sheet = null; drawSheet(); toast('Voce eliminata'); } catch (er) { document.getElementById('err').textContent = 'Non riesco a eliminare la voce.'; }
    return;
  }
  if (e.target.id === 'save') {
    readSheet();
    const err = document.getElementById('err'), dd = parseD(F.date);
    if (!dd) { err.textContent = 'Scegli una data valida.'; return; }
    let v = [0, 0, 0, 0], cat = F.cat, det = (F.det || '').trim();
    if (F.split) { v = F.v.map(r2); if (!v.some(x => Math.abs(x) > 0.004)) { err.textContent = 'Inserisci almeno un importo.'; return; } }
    else {
      const a = Math.abs(parseAmt(F.amount));
      if (!(a > 0)) { err.textContent = 'Inserisci un importo maggiore di zero.'; return; }
      if (F.type === 'mov') { if (F.from === F.to) { err.textContent = 'Scegli due conti diversi.'; return; } v[F.from] = -a; v[F.to] = a; cat = 'Varie'; det = det || 'Spostamento'; }
      else v[F.acc] = F.type === 'in' ? a : -a;
    }
    if (!cat) { err.textContent = 'Scegli una categoria.'; return; }
    const docId = `${dd.getFullYear()}-${pad(dd.getMonth() + 1)}`;
    const id = F.edit ? F.id : 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
    const row = { d: F.date, c: cat, t: det, v: v.map(r2), o: F.edit && F.docId === docId && F.o ? F.o : Date.now() };
    const btn = e.target; btn.disabled = true; btn.textContent = 'Salvo…';
    try {
      if (F.edit && F.docId !== docId) await B.delRow(S.who, F.docId, F.id);
      await B.putRow(S.who, docId, id, row);
      S.sheet = null; drawSheet();
      if (dd.getFullYear() !== S.year) S.year = dd.getFullYear();
      S.month = dd.getMonth(); render();
      toast(F.edit ? 'Voce aggiornata' : 'Voce aggiunta a ' + MN[dd.getMonth()].toLowerCase());
    } catch (er) { btn.disabled = false; btn.textContent = 'Salva'; err.textContent = 'Non riesco a salvare. Controlla la connessione e di essere nei tuoi conti.'; }
  }
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') { if (S.sheet) { S.sheet = null; drawSheet(); } else if (S.modal) closeModal(); } });

/* ---------------------------------------------------------------- impostazioni */
function closeModal() { S.modal = null; document.getElementById('modalHost').innerHTML = ''; }
function openSettings() {
  const cfg = S.cfg || {};
  const open = (cfg.open && cfg.open.length === 4) ? cfg.open : [0, 0, 0, 0];
  const sy = cfg.startYear || (C && Math.min(...yearList())) || now.getFullYear();
  const v = S.profile && S.profile.viewers ? S.profile.viewers : [];
  S.modal = { kind: 'settings' };
  document.getElementById('modalHost').innerHTML = '<div class="scrim" id="mscrim"><div class="sheet" role="dialog" aria-modal="true" aria-label="Impostazioni"><h2>Impostazioni</h2><div class="form">' +
    '<div class="eyebrow">Conti e saldi iniziali</div>' +
    '<p class="hint" style="margin:0">Saldo di ogni conto al 1 gennaio dell’anno in cui parti. Gli anni dopo ripartono dal saldo finale dell’anno prima.</p>' +
    '<div class="two">' + [0, 1, 2, 3].map(k => '<label>Conto ' + (k + 1) + '<input class="field" id="an' + k + '" value="' + esc(ACC[k]) + '"></label>').join('') + '</div>' +
    '<div class="two">' + [0, 1, 2, 3].map(k => '<label>Saldo iniziale<input class="field num" id="ao' + k + '" inputmode="decimal" value="' + String(open[k] || '').replace('.', ',') + '" placeholder="0,00"></label>').join('') + '</div>' +
    '<label>Anno di partenza<input class="field num" id="sy" inputmode="numeric" value="' + sy + '"></label>' +
    '<label>Obiettivo di risparmio annuale (€)<input class="field num" id="goal" inputmode="decimal" value="' + String(cfg.goal == null ? DEF.goal : cfg.goal).replace('.', ',') + '"></label>' +
    '<div class="eyebrow" style="margin-top:8px">Categorie</div>' +
    '<label>Uscite (una per riga)<textarea class="field" id="cats" rows="7">' + esc(CATS.join('\n')) + '</textarea></label>' +
    '<label>Entrate (una per riga)<textarea class="field" id="incs" rows="3">' + esc(INC.join('\n')) + '</textarea></label>' +
    '<p class="hint" style="margin:0">Se rinomini una categoria, le voci già inserite restano con il vecchio nome e compaiono nei Controlli.</p>' +
    '<div class="actions"><button class="btn pri" id="ssave" type="button">Salva impostazioni</button></div>' +
    '<div class="eyebrow" style="margin-top:14px">Excel</div>' +
    '<p class="hint" style="margin:0">Ricarica il file se l’hai modificato da PC: vedrai le differenze mese per mese prima di applicarle.</p>' +
    '<div class="actions"><button class="btn" id="simport" type="button">Importa o aggiorna da Excel</button></div>' +
    '<div class="eyebrow" style="margin-top:14px">Sola lettura per altri</div>' +
    '<p class="hint" style="margin:0">Chi inserisci qui, con l’email del suo account Google, potrà guardare i tuoi conti senza modificarli.</p>' +
    '<div id="vlist">' + viewersHtml(v) + '</div>' +
    '<div class="two"><input class="field" id="vmail" type="email" placeholder="email@gmail.com" autocomplete="off"><button class="btn" id="vadd" type="button">Aggiungi</button></div>' +
    '<div class="err" id="serr" role="alert"></div>' +
    '<div class="eyebrow" style="margin-top:14px">Account</div><p class="hint" style="margin:0">' + esc(S.user.email) + (B.kind === 'local' ? ' · modalità demo' : '') + '</p>' +
    '<div class="actions"><button class="btn" id="sclose" type="button">Chiudi</button><button class="btn del" id="slogout" type="button" style="flex:1">Esci</button></div>' +
    '</div></div></div>';
}
const viewersHtml = v => v.length ? '<div class="vchips">' + v.map(m => '<span class="vchip">' + esc(m) + '<button data-vdel="' + esc(m) + '" aria-label="Rimuovi ' + esc(m) + '">×</button></span>').join('') + '</div>' : '<p class="hint" style="margin:0">Nessuno per ora.</p>';
document.getElementById('modalHost').addEventListener('click', async e => {
  if (!S.modal) return;
  const id = e.target.id, g = k => document.getElementById(k);
  if (id === 'mscrim' || id === 'sclose' || id === 'icancel') return closeModal();
  if (S.modal.kind === 'settings') {
    if (id === 'ssave') {
      const open = [0, 1, 2, 3].map(k => r2(parseAmt(g('ao' + k).value)));
      const acc = [0, 1, 2, 3].map(k => g('an' + k).value.trim() || DEF.acc[k]);
      const cats = lines(g('cats').value), inc = lines(g('incs').value);
      const sy = parseInt(g('sy').value, 10) || now.getFullYear();
      try { await B.setConfig(S.who, { acc, open, startYear: sy, goal: r2(parseAmt(g('goal').value)), cats, inc }); toast('Impostazioni salvate'); closeModal(); } catch (er) { g('serr').textContent = 'Non riesco a salvare le impostazioni.'; }
      return;
    }
    if (id === 'simport') { closeModal(); return pickFile(); }
    if (id === 'slogout') { closeModal(); try { await B.signOut(); } catch (er) { } if (B.kind === 'local') toast('In modalità demo non c’è un account da chiudere.'); return; }
    if (id === 'vadd') {
      const m = g('vmail').value.trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m)) { g('serr').textContent = 'Scrivi un indirizzo email valido.'; return; }
      const list = [...new Set([...(S.profile && S.profile.viewers || []), m])];
      try { await B.setViewers(S.user.uid, list); g('vmail').value = ''; g('serr').textContent = ''; } catch (er) { g('serr').textContent = 'Non riesco ad aggiungerlo.'; }
      return;
    }
    const del = e.target.closest('[data-vdel]');
    if (del) { const list = (S.profile.viewers || []).filter(x => x !== del.dataset.vdel); try { await B.setViewers(S.user.uid, list); } catch (er) { g('serr').textContent = 'Non riesco a rimuoverlo.'; } return; }
  }
  if (S.modal.kind === 'import' && id === 'iapply') applyImport();
});
document.getElementById('modalHost').addEventListener('change', e => {
  if (S.modal && S.modal.kind === 'import' && e.target.dataset.imp != null) S.modal.sel[e.target.dataset.imp] = e.target.checked;
});

/* ---------------------------------------------------------------- Excel */
function download(name, buf) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], name, { type: blob.type });
  if (/iPhone|iPad/.test(navigator.userAgent) && navigator.canShare && navigator.canShare({ files: [file] })) {
    navigator.share({ files: [file], title: name }).catch(() => { });
    return;
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}
async function doExport(month) {
  toast('Preparo il file Excel…');
  try {
    const E = await loadExcelJS();
    const cfg = { acc: ACC, cats: CATS, inc: INC, goal: (S.cfg && S.cfg.goal != null) ? S.cfg.goal : DEF.goal };
    const buf = await buildWorkbook(E, { year: S.year, c: C, cfg, month });
    const who = readOnly() ? '_' + ((S.shared.find(p => p.uid === S.who) || {}).name || 'altro').replace(/[^\w-]+/g, '') : '';
    download(month == null ? `Conti_${S.year}${who}.xlsx` : `Conti_${S.year}-${pad(month + 1)}_${MN[month]}${who}.xlsx`, buf);
    toast(month == null ? 'File dell’anno pronto' : 'File di ' + MN[month].toLowerCase() + ' pronto');
  } catch (e) { toast(e.message || 'Esportazione non riuscita.'); }
}
function pickFile() {
  if (readOnly()) return;
  const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; inp.id = 'xfile';
  inp.style.display = 'none'; document.body.appendChild(inp);
  inp.addEventListener('change', async () => { const f = inp.files[0]; inp.remove(); if (f) await handleFile(f); });
  inp.click();
}
async function handleFile(file) {
  toast('Leggo il file…');
  try {
    const E = await loadExcelJS();
    const known = CATS.concat(INC);
    const parsed = await readWorkbook(E, await file.arrayBuffer(), known);
    if (!parsed.ids.length) { toast('Non trovo fogli mensili (es. “Gennaio ' + S.year + '”) in questo file.'); return; }
    const diff = diffMonths(parsed, S.docs);
    const sel = {}; diff.forEach(d => sel[d.id] = !d.same && d.file > 0);
    // categorie nuove: se nel file hanno saldo positivo sono entrate
    const net = {};
    parsed.ids.forEach(id => Object.values(parsed.months[id].rows).forEach(r => { net[r.c] = (net[r.c] || 0) + sum(r.v); }));
    const fresh = Object.keys(net).filter(k => k && !CATS.includes(k) && !INC.includes(k));
    const newInc = fresh.filter(k => net[k] > 0), newExp = fresh.filter(k => net[k] <= 0);
    const first = parsed.ids[0];
    S.modal = { kind: 'import', parsed, diff, sel, newInc, newExp, canOpen: first.endsWith('-01'), openVals: parsed.info[first].open, firstId: first, fresh: !S.cfg, fileName: file.name };
    drawImport();
  } catch (e) { toast('Non riesco a leggere il file: ' + (e.message || 'formato non riconosciuto')); }
}
function drawImport() {
  const M = S.modal, any = Object.values(M.sel).some(Boolean);
  document.getElementById('modalHost').innerHTML = '<div class="scrim" id="mscrim"><div class="sheet" role="dialog" aria-modal="true" aria-label="Importa da Excel"><h2>Importa da Excel</h2><div class="form">' +
    '<p class="hint" style="margin:0">' + esc(M.fileName) + '. I mesi selezionati sostituiscono quelli dell’app. Gli altri restano come sono.</p>' +
    '<div class="imp">' + M.diff.map(d => { const [y, m] = d.id.split('-').map(Number); const st = d.file === 0 && d.app === 0 ? '<span class="pill">Vuoto</span>' : d.app === 0 ? '<span class="pill">Nuovo</span>' : d.same ? '<span class="pill">Uguale</span>' : '<span class="pill warn">+' + d.added + ' / −' + d.removed + '</span>';
      return '<label class="irow"><input type="checkbox" data-imp="' + d.id + '"' + (M.sel[d.id] ? ' checked' : '') + '><span><b>' + MN[m - 1] + ' ' + y + '</b><small>' + d.file + ' voci nel file · ' + d.app + ' nell’app</small></span>' + st + '</label>'; }).join('') + '</div>' +
    (M.parsed.warnings.length ? '<div class="banner">' + M.parsed.warnings.map(esc).join('<br>') + '</div>' : '') +
    (M.newInc.length || M.newExp.length ? '<p class="hint" style="margin:0">Nuove categorie aggiunte: ' + esc(M.newExp.concat(M.newInc).join(', ')) + (M.newInc.length ? ' (entrate: ' + esc(M.newInc.join(', ')) + ')' : '') + '. Puoi cambiarle dalle Impostazioni.</p>' : '') +
    (M.canOpen ? '<label class="irow"><input type="checkbox" id="iopen"' + (M.fresh ? ' checked' : '') + '><span><b>Usa i saldi iniziali del file</b><small>' + (M.parsed.acc || ACC).map((a, k) => esc(a) + ' ' + fE0(M.openVals[k])).join(' · ') + '</small></span></label>' : '') +
    ((M.parsed.acc && M.parsed.acc.join() !== ACC.join()) ? '<label class="irow"><input type="checkbox" id="inames"' + (M.fresh ? ' checked' : '') + '><span><b>Usa i nomi dei conti del file</b><small>' + M.parsed.acc.map(esc).join(' · ') + '</small></span></label>' : '') +
    '<div class="err" id="ierr" role="alert"></div><div class="actions"><button class="btn" id="icancel" type="button">Annulla</button><button class="btn pri" id="iapply" type="button"' + (any ? '' : ' disabled') + '>Importa</button></div></div></div></div>';
}
async function applyImport() {
  const M = S.modal, g = k => document.getElementById(k), btn = g('iapply');
  const chosen = M.parsed.ids.filter(id => M.sel[id]);
  if (!chosen.length) return;
  btn.disabled = true; btn.textContent = 'Importo…';
  try {
    const map = {}; chosen.forEach(id => { map[id] = M.parsed.months[id]; });
    const cfg = {};
    const useOpen = g('iopen') && g('iopen').checked, useNames = g('inames') && g('inames').checked;
    if (useOpen) { cfg.open = M.openVals; cfg.startYear = +M.firstId.slice(0, 4); }
    if (useNames) cfg.acc = M.parsed.acc;
    if (M.newInc.length || M.newExp.length) { cfg.cats = CATS.concat(M.newExp); cfg.inc = INC.concat(M.newInc); }
    if (!S.cfg && !cfg.startYear) cfg.startYear = +M.firstId.slice(0, 4);
    await B.replaceMonths(S.who, map);
    if (Object.keys(cfg).length) await B.setConfig(S.who, cfg);
    const y = +chosen[chosen.length - 1].slice(0, 4);
    closeModal();
    S.year = y; S.month = +chosen[chosen.length - 1].slice(5) - 1; render();
    toast('Importati ' + chosen.length + (chosen.length === 1 ? ' mese' : ' mesi'));
  } catch (e) { btn.disabled = false; btn.textContent = 'Importa'; g('ierr').textContent = 'Importazione non riuscita: ' + (e.message || e.code || 'errore'); }
}

/* ---------------------------------------------------------------- avvio */
(async function boot() {
  renderNav(); render();
  const demo = new URLSearchParams(location.search).has('demo');
  const configured = firebaseConfig.apiKey && !/INCOLLA/.test(firebaseConfig.apiKey);
  try {
    if (configured && !demo) { const { createFirebaseBackend } = await import('./backend-firebase.js'); B = await createFirebaseBackend(); }
    else B = createLocalBackend();
  } catch (e) { B = createLocalBackend(); S.err = 'Firebase non raggiungibile: uso la modalità demo.'; }
  B.onAuth(async u => {
    if (sharedUnsub) { try { sharedUnsub(); } catch (e) { } sharedUnsub = null; }
    unsubs.forEach(f => { try { f(); } catch (e) { } }); unsubs = [];
    if (profUnsub) { try { profUnsub(); } catch (e) { } profUnsub = null; }
    if (!u) { S.user = null; S.ready = false; render(); return; }
    if (u.verified === false) { S.user = null; render(); document.getElementById('lerr') && (document.getElementById('lerr').textContent = 'L’indirizzo email del tuo account non è verificato.'); return; }
    S.user = u; S.who = u.uid;
    try { await B.ensureProfile(u); } catch (e) { S.err = 'Non riesco a creare il tuo spazio nel database: controlla le regole di Firestore.'; }
    profUnsub = B.watchProfile(u.uid, p => { S.profile = p; const m = document.getElementById('vlist'); if (m && p) m.innerHTML = viewersHtml(p.viewers || []); });
    sharedUnsub = B.watchShared(u.email, list => { S.shared = list; if (S.who !== u.uid && !list.some(p => p.uid === S.who)) { S.who = u.uid; startWatch(); } else if (S.ready) render(); });
    startWatch();
  });
})();

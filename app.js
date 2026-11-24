/* Leietakere — kort med data, redigering og gruppering etter gate. */
(function () {
  'use strict';

  /* ---------- Kolonner ----------
   * `header` er overskriften som brukes ved eksport til Excel (samme som i den opprinnelige filen).
   * Ved import finnes kolonnene ut fra overskriften; ellers brukes kolonnebokstaven.
   */
  const FIELDS = [
    { key: 'art',     col: 'A', header: 'ARTnr',                headers: ['artnr'],                 label: 'ARTnr', width: 8 },
    { key: 'abb',     col: 'B', header: 'Abb.nr.',              headers: ['abb.nr.', 'abbnr'],      label: 'Abb.nr.', width: 9 },
    { key: 'kunde',   col: 'C', header: 'Kundenr. ',            headers: ['kundenr.', 'kundenr'],   label: 'Kundenr.', width: 11 },
    { key: 'status',  col: 'D', header: 'Status',               headers: ['status'],                label: 'Status', type: 'status', width: 8 },
    { key: 'belop1',  col: 'E', header: 'Beløp1',               headers: ['beløp1'],                label: 'Beløp 1', type: 'money', width: 14 },
    { key: 'belop2',  col: 'F', header: 'Beløp2',               headers: ['beløp2'],                label: 'Beløp 2', type: 'money', width: 14 },
    { key: 'navn',    col: 'G', header: 'Navn',                 headers: ['navn'],                  label: 'Navn', width: 28 },
    { key: 'init',    col: 'H', header: 'Kolumna1',             headers: ['kolumna1', 'initialer'], label: 'Initialer', type: 'initials', width: 10 },
    { key: 'bosted',  col: 'I', header: 'Bosted',               headers: ['bosted'],                label: 'Bosted', width: 16 },
    { key: 'bolig',   col: 'J', header: 'Bolig/leilighet',      headers: ['bolig/leilighet'],       label: 'Bolig/leilighet', width: 20 },
    { key: 'adresse', col: 'K', header: 'Adresse',              headers: ['adresse'],               label: 'Adresse', width: 22 },
    { key: 'inn',     col: 'L', header: 'Dato innflytting',     headers: ['datoinnflytting'],       label: 'Dato innflytting', type: 'date', width: 14 },
    { key: 'ut',      col: 'M', header: 'Dato utflytting',      headers: ['datoutflytting'],        label: 'Dato utflytting', type: 'date', width: 14 },
    { key: 'reg',     col: 'N', header: 'Neste regulering',     headers: ['nesteregulering'],       label: 'Neste regulering', type: 'date', width: 14 },
    { key: 'status2', col: 'O', header: 'Status2',              headers: ['status2'],               label: 'Status 2', width: 16 },
    { key: 'komm',    col: 'P', header: 'Kommentar/telefomnr.', headers: ['kommentar/telefomnr.', 'kommentar/telefonnr.'], label: 'Kommentar / telefonnr.', type: 'note', width: 26 },
  ];
  const FIELD = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

  // Ekstra kolonner (valgfrie ved import, alltid med ved eksport).
  const GENDER_HEADERS = ['kjønn', 'kjonn', 'kjønn(m/k)', 'kjønn(k/m)', 'gender', 'płeć', 'plec'];
  const NOTE_HEADERS = ['merknad', 'merknader', 'notat'];

  // Søkbare felt = kolonnene + merknaden.
  const SEARCH_KEYS = FIELDS.map((f) => f.key).concat('merknad');
  const LABEL = Object.assign(Object.fromEntries(FIELDS.map((f) => [f.key, f.label])), { merknad: 'Merknad' });

  // Hva som kan vises på kortet (rekkefølgen på kortet er alltid som her).
  const SHOW_OPTIONS = [
    { id: 'init',    label: 'Initialer',            short: 'Initialer' },
    { id: 'navn',    label: 'Fullt navn',           short: 'Navn' },
    { id: 'adresse', label: 'Adresse',              short: 'Adresse' },
    { id: 'nr',      label: 'Kundenr. og Abb.nr.',  short: 'Numre' },
  ];
  const DEFAULT_SHOW = ['adresse'];

  // De tre datoene som alltid vises nederst på hvert kort.
  const DATE_ROWS = [
    { key: 'inn', short: 'Inn',  title: 'Dato innflytting' },
    { key: 'ut',  short: 'Ut',   title: 'Dato utflytting (kontraktslutt)' },
    { key: 'reg', short: 'Reg.', title: 'Neste regulering (KPI)' },
  ];

  const DUE = {
    m1: { label: 'under 1 md.', hint: 'mindre enn 1 måned igjen' },
    m2: { label: '1–2 md.',     hint: '1–2 måneder igjen' },
    m3: { label: '2–3 md.',     hint: '2–3 måneder igjen' },
  };
  const DUE_GROUPS = [
    { kind: 'ut',  title: 'Kontrakt utløper', short: 'Kontrakt', pastLabel: 'Utløpt',  pastHint: 'kontrakten har allerede utløpt' },
    { kind: 'reg', title: 'KPI-regulering',   short: 'KPI',      pastLabel: 'Forfalt', pastHint: 'reguleringsdatoen er passert' },
  ];
  const DUE_ORDER = ['m1', 'm2', 'm3', 'past'];

  const SECTIONS = [
    { title: 'Bolig',                fields: ['adresse', 'bosted', 'bolig'] },
    { title: 'Kontrakt og datoer',   fields: ['inn', 'ut', 'reg', 'status', 'status2'] },
    { title: 'Betaling',             fields: ['belop1', 'belop2'] },
    { title: 'Kontakt og merknader', fields: ['komm', 'merknad'] },
  ];

  // Felt som tømmes når en bolig markeres som ledig (adressen og boligen beholdes).
  const TENANT_KEYS = ['navn', 'init', 'kunde', 'abb', 'inn', 'ut', 'reg', 'status2', 'komm'];

  const SORTS = ['file', 'alpha', 'ut', 'reg'];
  const STORE_KEY = 'leietakere.data.v1';

  /* ---------- Tilstand ---------- */

  const state = {
    meta: {},
    records: [],
    nextUid: 1,
    show: loadPref('show', DEFAULT_SHOW).filter((id) => SHOW_OPTIONS.some((o) => o.id === id)),
    sort: loadPref('sort', 'file'),
    showVacant: loadPref('showVacant', true),
    view: loadPref('view', 'kort') === 'gater' ? 'gater' : 'kort',
    street: null,
    filters: new Set(),
    query: '',
    tokens: [],
    openId: null,
    flash: new Set(),
    undo: [],
    hasStore: false,
  };
  if (!state.show.length) state.show = DEFAULT_SHOW.slice();
  if (!SORTS.includes(state.sort)) state.sort = 'file';

  const $ = (id) => document.getElementById(id);
  const els = {
    subtitle: $('subtitle'), source: $('source'),
    fileInput: $('fileInput'), exportBtn: $('exportBtn'), moreBtn: $('moreBtn'), moreMenu: $('moreMenu'),
    importCmd: $('importCmd'), resetCmd: $('resetCmd'),
    search: $('search'), showBtn: $('showBtn'), showValue: $('showValue'), showMenu: $('showMenu'),
    sort: $('sort'), vacant: $('showVacant'), vacantCount: $('vacantCount'), count: $('count'),
    overview: $('overview'), viewbar: $('viewbar'), viewKort: $('viewKort'), viewGater: $('viewGater'),
    crumbs: $('crumbs'), addBtn: $('addBtn'),
    grid: $('grid'), empty: $('empty'), emptyText: $('emptyText'), clearSearch: $('clearSearch'),
    loader: $('loader'), loaderTitle: $('loaderTitle'), loaderText: $('loaderText'), pickBtn: $('pickBtn'),
    toolbar: $('toolbar'), toast: $('toast'),
    editor: $('editor'), editorForm: $('editorForm'), editorTitle: $('editorTitle'), editorSub: $('editorSub'),
    editorBody: $('editorBody'), vacateBtn: $('vacateBtn'), deleteBtn: $('deleteBtn'), cancelBtn: $('cancelBtn'),
    editorClose: $('editorClose'), saveBtn: $('saveBtn'),
  };

  /* ---------- Innstillinger ---------- */

  function loadPref(key, fallback) {
    try {
      const raw = localStorage.getItem('leietakere.' + key);
      return raw == null ? fallback : JSON.parse(raw);
    } catch (e) { return fallback; }
  }
  function savePref(key, value) {
    try { localStorage.setItem('leietakere.' + key, JSON.stringify(value)); } catch (e) { /* ingen lagring */ }
  }

  /* ---------- Formatering ---------- */

  const moneyFmt = new Intl.NumberFormat('nb-NO', { style: 'currency', currency: 'NOK', minimumFractionDigits: 2 });
  const relFmt = new Intl.RelativeTimeFormat('nb', { numeric: 'auto' });
  const collator = new Intl.Collator('nb', { numeric: true, sensitivity: 'base' });

  const isEmpty = (v) => v == null || (typeof v === 'string' && v.trim() === '');
  const isDate = (v) => v instanceof Date && !isNaN(v);
  const plural = (n, one, many) => (n === 1 ? one : many);
  const pad = (n) => String(n).padStart(2, '0');

  function formatDate(d) {
    return pad(d.getUTCDate()) + '.' + pad(d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
  }
  function isoDate(d) {
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }
  function formatStamp(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return '';
    return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear() + ' kl. ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  function today() {
    const now = new Date();
    return Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  }

  function addMonths(utcMs, n) {
    const d = new Date(utcMs);
    const y = d.getUTCFullYear(), m = d.getUTCMonth() + n;
    const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    return Date.UTC(y, m, Math.min(d.getUTCDate(), last));
  }

  const daysUntil = (d) => Math.round((d.getTime() - today()) / 86400000);

  // 'past' (passert), 'm1' (< 1 md.), 'm2' (< 2 md.), 'm3' (< 3 md.), 'ok' eller null.
  function dueStatus(d) {
    if (!isDate(d)) return null;
    const t = today(), ms = d.getTime();
    if (ms < t) return 'past';
    if (ms < addMonths(t, 1)) return 'm1';
    if (ms < addMonths(t, 2)) return 'm2';
    if (ms < addMonths(t, 3)) return 'm3';
    return 'ok';
  }

  // «i dag», «i morgen», «om 7 dager», «for 3 dager siden» (Intl bruker «døgn»).
  function dayPhrase(days) {
    if (days === 0) return 'i dag';
    if (days === 1) return 'i morgen';
    if (days === -1) return 'i går';
    return days > 0 ? 'om ' + days + ' dager' : 'for ' + -days + ' dager siden';
  }

  function relativeDate(d) {
    const days = daysUntil(d);
    if (Math.abs(days) < 45) return dayPhrase(days);
    const months = Math.round(days / 30.44);
    if (Math.abs(months) < 18) return relFmt.format(months, 'month');
    return relFmt.format(Math.round(days / 365.25), 'year');
  }

  function dueSentence(kind, status, d) {
    const date = formatDate(d);
    const when = dayPhrase(daysUntil(d));
    if (kind === 'ut') {
      return status === 'past'
        ? 'Kontrakten utløp ' + relativeDate(d) + ' (' + date + ')'
        : 'Kontrakten utløper ' + when + ' (' + date + ')';
    }
    return status === 'past'
      ? 'KPI-reguleringen er forfalt – datoen ' + date + ' er passert (' + relativeDate(d) + ')'
      : 'Neste KPI-regulering ' + when + ' (' + date + ')';
  }

  function display(key, v) {
    if (isEmpty(v)) return null;
    const f = FIELD[key];
    if (v instanceof Date) return isNaN(v) ? null : formatDate(v);
    if (f.type === 'money' && typeof v === 'number') return moneyFmt.format(v);
    if (f.type === 'status' && typeof v === 'boolean') return v ? 'Aktiv' : 'Inaktiv';
    if (typeof v === 'boolean') return v ? 'Ja' : 'Nei';
    if (f.type === 'initials') return String(v).replace(/[\s.]+/g, '').toUpperCase();
    return String(v).trim();
  }

  // Dato skrevet som tekst (f.eks. «31.04.2027» eller «2027-04-30»): gyldig → Date, ellers null.
  function parseDateText(s) {
    let m = /^\s*(\d{4})-(\d{2})-(\d{2})\s*$/.exec(s);
    let y, mo, d;
    if (m) { y = +m[1]; mo = +m[2] - 1; d = +m[3]; }
    else {
      m = /^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})\s*$/.exec(s);
      if (!m) return null;
      y = +m[3]; if (y < 100) y += 2000; mo = +m[2] - 1; d = +m[1];
    }
    const date = new Date(Date.UTC(y, mo, d));
    return date.getUTCMonth() === mo && date.getUTCDate() === d ? date : null;
  }

  /* ---------- Søk ---------- */

  const SPECIAL = { 'æ': 'ae', 'ø': 'o', 'å': 'a', 'ß': 'ss', 'ł': 'l', 'đ': 'd', 'œ': 'oe' };

  function foldChar(ch) {
    const lower = ch.toLowerCase();
    if (SPECIAL[lower]) return SPECIAL[lower];
    return lower.normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  function foldWithMap(text, dropSpaces) {
    let out = '';
    const map = [];
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (dropSpaces && /\s/.test(ch)) continue;
      const f = foldChar(ch);
      for (let k = 0; k < f.length; k++) { out += f[k]; map.push(i); }
    }
    return { out, map };
  }

  const fold = (text) => foldWithMap(text, false).out;
  const foldCompact = (text) => foldWithMap(text, true).out;
  const tokenize = (query) => fold(query).split(/\s+/).filter(Boolean);
  const searchIndex = (hay) => ({ folded: fold(hay), compact: foldCompact(hay) });

  function matchRanges(text, tokens) {
    if (!text || !tokens.length) return [];
    const ranges = [];
    for (const dropSpaces of [false, true]) {
      const { out, map } = foldWithMap(text, dropSpaces);
      for (const t of tokens) {
        let i = out.indexOf(t);
        while (i !== -1) {
          ranges.push([map[i], map[i + t.length - 1] + 1]);
          i = out.indexOf(t, i + 1);
        }
      }
    }
    ranges.sort((a, b) => a[0] - b[0]);
    const merged = [];
    for (const r of ranges) {
      const last = merged[merged.length - 1];
      if (last && r[0] <= last[1]) last[1] = Math.max(last[1], r[1]);
      else merged.push(r.slice());
    }
    return merged;
  }

  function fieldMatches(rec, key, token) {
    const s = rec.search[key];
    return !!s && (s.folded.includes(token) || s.compact.includes(token));
  }

  const recordMatches = (rec, tokens) => tokens.every((t) => SEARCH_KEYS.some((k) => fieldMatches(rec, k, t)));

  /* ---------- Kjønn, gate og initialer ---------- */

  function parseGender(v) {
    if (isEmpty(v)) return null;
    const s = fold(String(v)).trim();
    if (/^(org|tjeneste)/.test(s)) return 'org';
    if (/^(k|kv|kvinne|f|female|w|woman|kobieta|d|dame)\b/.test(s)) return 'K';
    if (/^(m|mann|male|man|mezczyzna|herre|gutt)\b/.test(s)) return 'M';
    return null;
  }

  const guessGender = (navn) => (navn ? NavnKjonn.gjett(String(navn), fold) : null);

  // «Eldridhågen 27» → «Eldridhågen», «Nordfoldveien 143 A/B» → «Nordfoldveien».
  function streetOf(addr) {
    if (isEmpty(addr)) return 'Uten adresse';
    const s = String(addr).trim();
    const m = /^(.*?\D)[\s,]+\d/.exec(s);
    return (m ? m[1] : s).trim();
  }

  function houseOf(addr, street) {
    if (isEmpty(addr)) return '';
    const s = String(addr).trim();
    return s.slice(street.length).replace(/^[\s,]+/, '');
  }

  // «Adreassen Oddny Karoline» → «A O K» (samme form som i Excel-filen).
  function autoInitials(navn) {
    if (isEmpty(navn)) return '';
    return String(navn).split(/[\s\-–.]+/).filter((w) => /\p{L}/u.test(w))
      .map((w) => w.match(/\p{L}/u)[0].toUpperCase()).join(' ');
  }

  /* ---------- Data: avledede verdier ---------- */

  function derive(rec) {
    rec.text = {};
    rec.search = {};
    rec.badDate = {};
    for (const f of FIELDS) {
      const v = rec.values[f.key];
      if (f.type === 'date' && typeof v === 'string' && v.trim()) rec.badDate[f.key] = true;
      const txt = display(f.key, v);
      rec.text[f.key] = txt;
      if (txt) {
        let hay = txt;
        if (typeof v === 'number') hay += ' ' + v;
        if (f.type === 'initials') hay += ' ' + v;
        rec.search[f.key] = searchIndex(hay);
      }
    }
    rec.text.merknad = rec.merknad || null;
    if (rec.merknad) rec.search.merknad = searchIndex(rec.merknad);
    rec.vacant = isEmpty(rec.values.navn) && isEmpty(rec.values.kunde);
    rec.due = { ut: dueStatus(rec.values.ut), reg: dueStatus(rec.values.reg) };
    if (rec.vacant) rec.gender = { value: null, source: null };
    else if (rec.kjonn) rec.gender = { value: rec.kjonn, source: 'manual' };
    else {
      const g = guessGender(rec.values.navn);
      rec.gender = { value: g, source: g ? 'guess' : null };
    }
    rec.street = streetOf(rec.values.adresse);
    rec.streetKey = fold(rec.street);
    rec.house = houseOf(rec.values.adresse, rec.street);
    return rec;
  }

  function emptyValues() {
    const v = {};
    for (const f of FIELDS) v[f.key] = null;
    return v;
  }

  function newRecord(values, extra) {
    const rec = Object.assign({ uid: state.nextUid++, values, kjonn: null, merknad: '', changed: null }, extra || {});
    return derive(rec);
  }

  const byUid = (uid) => state.records.find((r) => r.uid === uid);

  /* ---------- Lagring ---------- */

  function toPlain(rec) {
    const values = {};
    for (const f of FIELDS) {
      const v = rec.values[f.key];
      values[f.key] = v instanceof Date ? isoDate(v) : (v === undefined ? null : v);
    }
    return { uid: rec.uid, values, kjonn: rec.kjonn || null, merknad: rec.merknad || '', changed: rec.changed || null };
  }

  function fromPlain(p) {
    const values = emptyValues();
    for (const f of FIELDS) {
      let v = p.values ? p.values[f.key] : null;
      if (v === undefined || v === '') v = null;
      if (f.type === 'date' && typeof v === 'string') v = parseDateText(v) || v;
      values[f.key] = v;
    }
    return derive({ uid: p.uid, values, kjonn: p.kjonn || null, merknad: p.merknad || '', changed: p.changed || null });
  }

  function snapshot() {
    return JSON.stringify({ meta: state.meta, records: state.records.map(toPlain) });
  }

  function applyData(data) {
    state.meta = Object.assign({ title: '', sheetName: 'Leietakere', source: '', headers: null, savedAt: null, exportedAt: null }, data.meta || {});
    state.records = (data.records || []).map(fromPlain);
    // Sørg for unike id-er.
    let next = state.records.reduce((m, r) => Math.max(m, Number(r.uid) || 0), 0) + 1;
    const seen = new Set();
    for (const r of state.records) {
      if (!r.uid || seen.has(r.uid)) r.uid = next++;
      seen.add(r.uid);
    }
    state.nextUid = next;
  }

  function persist() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ version: 1, meta: state.meta, records: state.records.map(toPlain) }));
      state.hasStore = true;
      return true;
    } catch (e) {
      toast('Endringene kunne ikke lagres i nettleseren – bruk «Eksporter til Excel» for å ta vare på dem.');
      return false;
    }
  }

  function readStore() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return null;
      const data = JSON.parse(raw);
      return data && Array.isArray(data.records) ? data : null;
    } catch (e) { return null; }
  }

  function clearStore() {
    try { localStorage.removeItem(STORE_KEY); } catch (e) { /* ingenting */ }
    state.hasStore = false;
  }

  // Alle endringer går gjennom denne: lagrer, tegner på nytt og gir mulighet til å angre.
  function mutate(message, fn, opts) {
    const before = snapshot();
    const now = new Date().toISOString();
    const result = fn(now);
    state.meta.savedAt = now;
    state.undo.push(before);
    if (state.undo.length > 30) state.undo.shift();
    persist();
    renderAll();
    toast(message, { action: 'Angre', run: undo, ms: 7000 });
    if (opts && opts.after) opts.after(result);
    return result;
  }

  function undo() {
    const before = state.undo.pop();
    if (!before) return;
    applyData(JSON.parse(before));
    if (state.openId != null && !byUid(state.openId)) state.openId = null;
    persist();
    renderAll();
    toast('Endringen er angret.');
  }

  /* ---------- Import fra Excel ---------- */

  function headerKey(v) {
    return typeof v === 'string' ? v.toLowerCase().replace(/\s+/g, '') : '';
  }

  function recordsFromSheet(sheet) {
    const rows = sheet.rows;
    let h = rows.findIndex((r, i) => i < 40 &&
      r.some((v) => headerKey(v) === 'navn') && r.some((v) => headerKey(v) === 'adresse'));
    if (h < 0) h = 2;
    const headerRow = rows[h] || [];

    const colOf = {};
    const headers = {};
    for (const f of FIELDS) {
      let idx = headerRow.findIndex((v) => f.headers.includes(headerKey(v)));
      if (idx < 0) idx = XlsxReader.colIndex(f.col);
      colOf[f.key] = idx;
      headers[f.key] = typeof headerRow[idx] === 'string' && headerRow[idx].trim() ? headerRow[idx] : f.header;
    }
    const genderCol = headerRow.findIndex((v) => GENDER_HEADERS.includes(headerKey(v)));
    const noteCol = headerRow.findIndex((v) => NOTE_HEADERS.includes(headerKey(v)));

    let title = '';
    for (let i = 0; i < h && !title; i++) {
      const cell = (rows[i] || []).find((v) => typeof v === 'string' && v.trim());
      if (cell) title = cell.trim();
    }

    const plain = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || FIELDS.every((f) => isEmpty(row[colOf[f.key]]))) continue;
      const values = {};
      for (const f of FIELDS) {
        let v = row[colOf[f.key]];
        if (typeof v === 'string') v = v.trim() || null;
        if (f.type === 'date' && typeof v === 'string') v = parseDateText(v) || v;
        if (f.key === 'art' && typeof v === 'number') v = String(v).padStart(5, '0');
        values[f.key] = v instanceof Date ? isoDate(v) : (v ?? null);
      }
      // Kommentarer i cellene blir til en merknad.
      const notes = [];
      if (noteCol >= 0 && !isEmpty(row[noteCol])) notes.push(String(row[noteCol]).trim());
      for (let c = 0; c < row.length + 20; c++) {
        const cm = sheet.comments.get(r + ',' + c);
        if (cm) notes.push(cm.text + (cm.author ? ' (' + cm.author + ')' : ''));
      }
      plain.push({
        uid: plain.length + 1,
        values,
        kjonn: genderCol >= 0 ? parseGender(row[genderCol]) : null,
        merknad: notes.join('\n'),
        changed: null,
      });
    }
    return {
      meta: { title, sheetName: sheet.sheetName, headers: FIELDS.map((f) => headers[f.key]) },
      records: plain,
    };
  }

  async function importBuffer(buffer, sourceName) {
    const sheet = await XlsxReader.read(buffer);
    const data = recordsFromSheet(sheet);
    if (!data.records.length) throw new Error('Fant ingen rader med data i arket.');
    data.meta.source = sourceName;
    return data;
  }

  /* ---------- Innlasting ved oppstart ---------- */

  async function init() {
    const stored = readStore();
    if (stored) {
      state.hasStore = true;
      applyData(stored);
      return showData();
    }
    if (window.INNEBYGDE_DATA) {
      applyData(window.INNEBYGDE_DATA);
      return showData();
    }
    try {
      const res = await fetch('Data.xlsx', { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      applyData(await importBuffer(await res.arrayBuffer(), 'Data.xlsx'));
      return showData();
    } catch (err) {
      showLoader('Velg Excel-fil', 'Fant ingen data. Velg Excel-filen med leietakere, eller dra den hit.');
    }
  }

  function showData() {
    els.loader.hidden = true;
    document.body.classList.add('has-data');
    onSearch();
  }

  async function loadFile(file) {
    if (!file) return;
    try {
      const data = await importBuffer(await file.arrayBuffer(), file.name);
      if (!state.records.length) {
        applyData(data);
        persist();
        showData();
        toast('Lastet inn: ' + file.name);
        return;
      }
      mutate('Lastet inn ' + data.records.length + ' boliger fra ' + file.name + '.', (now) => {
        applyData(data);
        state.meta.savedAt = now;
        state.openId = null;
        state.street = null;
      });
    } catch (err) {
      if (!state.records.length) showLoader('Kunne ikke lese filen', err.message || String(err));
      else toast('Kunne ikke lese filen: ' + (err.message || err));
    }
  }

  function showLoader(title, text) {
    els.loaderTitle.textContent = title;
    els.loaderText.textContent = text;
    els.pickBtn.hidden = false;
    els.loader.hidden = false;
    els.loader.classList.remove('is-loading');
  }

  /* ---------- Eksport til Excel ---------- */

  function exportExcel() {
    const headers = (state.meta.headers && state.meta.headers.length === FIELDS.length
      ? state.meta.headers : FIELDS.map((f) => f.header)).concat('Kjønn', 'Merknad');
    const rows = state.records.map((rec) => {
      const cells = FIELDS.map((f) => {
        const v = rec.values[f.key];
        if (f.type === 'money') return { v, style: 'money' };
        if (f.key === 'art') return { v: v == null ? null : String(v), style: 'text' };
        return v;
      });
      cells.push(rec.kjonn === 'K' ? 'K' : rec.kjonn === 'M' ? 'M' : rec.kjonn === 'org' ? 'Tjeneste' : null);
      cells.push(rec.merknad || null);
      return cells;
    });
    const blob = XlsxWriter.write({
      sheetName: state.meta.sheetName || 'Leietakere',
      title: state.meta.title || 'Leietakere',
      headers,
      rows,
      colWidths: FIELDS.map((f) => f.width).concat(10, 34),
    });
    const now = new Date();
    const name = 'Leietakere ' + now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate()) + '.xlsx';
    const url = URL.createObjectURL(blob);
    const a = el('a');
    a.href = url;
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    state.meta.exportedAt = now.toISOString();
    if (state.hasStore) persist();
    updateSource();
    toast('Eksportert: ' + name);
  }

  /* ---------- DOM-hjelpere ---------- */

  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  function hl(text, cls) {
    const span = el('span', cls);
    const ranges = matchRanges(text, state.tokens);
    let pos = 0;
    for (const [a, b] of ranges) {
      if (a > pos) span.append(text.slice(pos, a));
      span.append(el('mark', null, text.slice(a, b)));
      pos = b;
    }
    if (pos < text.length) span.append(text.slice(pos));
    return span;
  }

  const ICONS = {
    copy: '<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V6a2 2 0 0 1 2-2h9"/>',
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    close: '<path d="M6 6l12 12M18 6L6 18"/>',
    phone: '<path d="M6.6 3h3l1.6 4.2-2 1.3a11 11 0 0 0 6.3 6.3l1.3-2 4.2 1.6v3a2 2 0 0 1-2.2 2A17 17 0 0 1 4.6 5.2 2 2 0 0 1 6.6 3z"/>',
    note: '<path d="M5 4h14v12l-4 4H5z"/><path d="M15 20v-4h4"/>',
    female: '<circle cx="12" cy="9" r="5.2"/><path d="M12 14.2V22M8.5 18.5h7"/>',
    male: '<circle cx="10" cy="14" r="5.2"/><path d="M13.8 10.2L20 4M15 4h5v5"/>',
    unknown: '<path d="M9.2 9a2.9 2.9 0 1 1 4.1 2.6c-.8.4-1.3 1.1-1.3 2v.6"/><path d="M12 17.6v.2"/>',
    org: '<path d="M4 21V8l8-5 8 5v13"/><path d="M9 21v-5h6v5M8.5 10.5h1M14.5 10.5h1"/>',
    alert: '<path d="M12 4l9 16H3z"/><path d="M12 10v4M12 17v.2"/>',
    clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
    pencil: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="M13.5 6.5l4 4"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    back: '<path d="M15 5l-7 7 7 7"/>',
    street: '<path d="M4 20L9 4M20 20L15 4"/><path d="M12 6v2M12 11v2M12 16v2"/>',
  };

  function iconSvg(name, cls) {
    const span = el('span', 'icon' + (cls ? ' ' + cls : ''));
    span.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICONS[name] + '</svg>';
    return span;
  }

  const GENDER_INFO = {
    K:   { icon: 'female', cls: 'g-k',   text: 'Kvinne' },
    M:   { icon: 'male',   cls: 'g-m',   text: 'Mann' },
    org: { icon: 'org',    cls: 'g-org', text: 'Tjeneste / organisasjon' },
  };

  function genderTitle(rec) {
    const g = rec.gender;
    if (!g.value) return 'Kjønn ukjent – velg kjønn med blyanten';
    const info = GENDER_INFO[g.value];
    if (g.value === 'org') return info.text;
    return info.text + (g.source === 'guess' ? ' (antatt ut fra fornavn)' : '');
  }

  function genderBadge(rec, where) {
    const info = GENDER_INFO[rec.gender.value] || { icon: 'unknown', cls: 'g-unknown' };
    const b = iconSvg(info.icon, 'gender ' + info.cls + (where ? ' gender-' + where : ''));
    b.title = genderTitle(rec);
    b.setAttribute('role', 'img');
    b.setAttribute('aria-label', genderTitle(rec));
    return b;
  }

  function recordName(rec) {
    const addr = rec.text.adresse || 'Boligen';
    return rec.text.bolig && rec.text.bolig !== addr ? addr + ' (' + rec.text.bolig + ')' : addr;
  }

  /* ---------- Utvalg ---------- */

  function baseRecords() {
    let list = state.records.filter((r) => state.showVacant || !r.vacant);
    if (state.tokens.length) list = list.filter((r) => recordMatches(r, state.tokens));
    return list;
  }

  // Det utvalget som gjelder nå (alle, eller bare én gate).
  function scopeRecords() {
    const base = baseRecords();
    return state.view === 'gater' && state.street != null ? base.filter((r) => r.streetKey === state.street) : base;
  }

  function passesFilters(rec) {
    if (!state.filters.size) return true;
    for (const f of state.filters) {
      const [kind, status] = f.split(':');
      if (rec.due[kind] === status) return true;
    }
    return false;
  }

  function sortKey() {
    const primary = SHOW_OPTIONS.find((o) => state.show.includes(o.id)).id;
    return primary === 'nr' ? 'kunde' : primary;
  }

  function sortList(list) {
    if (state.sort === 'alpha') {
      const key = sortKey();
      return list.slice().sort((a, b) => {
        const x = a.text[key], y = b.text[key];
        if (!x) return y ? 1 : 0;
        if (!y) return -1;
        return collator.compare(x, y);
      });
    }
    if (state.sort === 'ut' || state.sort === 'reg') {
      const k = state.sort, t0 = today();
      const key = (r) => {
        const d = r.values[k];
        if (!isDate(d)) return [2, 0];
        return d.getTime() >= t0 ? [0, d.getTime()] : [1, -d.getTime()];
      };
      return list.slice().sort((a, b) => {
        const x = key(a), y = key(b);
        return x[0] - y[0] || x[1] - y[1];
      });
    }
    return list;
  }

  function shownFieldKeys() {
    const keys = [];
    for (const o of SHOW_OPTIONS) {
      if (!state.show.includes(o.id)) continue;
      if (o.id === 'nr') keys.push('kunde', 'abb');
      else keys.push(o.id);
    }
    return keys;
  }

  /* ---------- Kort (sammenslått) ---------- */

  function duePill(kind, status, d) {
    const pill = el('span', 'due due-' + status);
    if (status === 'past') {
      pill.textContent = kind === 'ut' ? 'utløpt' : 'forfalt';
    } else {
      const days = daysUntil(d);
      if (days === 0) pill.textContent = 'i dag';
      else pill.append(el('span', 'hide-sm', 'om '), days + ' d.');
    }
    return pill;
  }

  function buildDates(rec) {
    const box = el('span', 'dates');
    for (const row of DATE_ROWS) {
      const line = el('span', 'date-row');
      line.title = row.title;
      line.append(el('span', 'date-label', row.short));
      const v = rec.values[row.key];
      const txt = rec.text[row.key];
      if (txt) line.append(hl(txt, 'date-val' + (rec.badDate[row.key] ? ' is-bad' : '')));
      else line.append(el('span', 'date-val is-missing', '—'));
      // På smale skjermer vises kjønnssymbolet her i stedet for øverst.
      if (row.key === 'inn' && !rec.vacant) line.append(genderBadge(rec, 'row'));
      const st = row.key === 'inn' ? null : rec.due[row.key];
      if (st && st !== 'ok') {
        line.title = dueSentence(row.key, st, v);
        line.append(duePill(row.key, st, v));
      } else if (rec.badDate[row.key]) {
        line.append(el('span', 'due due-bad', 'ugyldig'));
        line.title = row.title + ': ugyldig dato («' + txt + '»)';
      }
      box.append(line);
    }
    return box;
  }

  function buildFace(rec) {
    const face = el('button', 'face');
    face.type = 'button';
    face.setAttribute('aria-expanded', 'false');

    const top = el('span', 'face-top');
    const main = el('span', 'face-main');
    let first = true;
    for (const o of SHOW_OPTIONS) {
      if (!state.show.includes(o.id)) continue;
      const role = first ? ' is-primary' : '';
      first = false;
      if (o.id === 'nr') {
        const box = el('span', 'f-nr' + role);
        for (const k of ['kunde', 'abb']) {
          const line = el('span', 'f-nr-line');
          line.append(el('span', 'f-nr-label', FIELD[k].label));
          line.append(rec.text[k] ? hl(rec.text[k], 'f-nr-val') : el('span', 'f-nr-val is-missing', '—'));
          box.append(line);
        }
        main.append(box);
        continue;
      }
      const txt = rec.text[o.id];
      if (!txt && rec.vacant && (o.id === 'navn' || o.id === 'init')) {
        main.append(el('span', 'f-' + o.id + role + ' is-missing', o.id === 'init' ? '—' : 'Ledig bolig'));
        continue;
      }
      main.append(txt ? hl(txt, 'f-' + o.id + role) : el('span', 'f-' + o.id + role + ' is-missing', '—'));
    }
    top.append(main);
    if (rec.vacant) top.append(el('span', 'badge', 'Ledig'));
    else top.append(genderBadge(rec, 'top'));
    face.append(top);

    if (state.tokens.length) {
      const shown = shownFieldKeys().concat(DATE_ROWS.map((r) => r.key));
      const hidden = state.tokens.filter((t) => !shown.some((k) => fieldMatches(rec, k, t)));
      const used = new Set();
      for (const t of hidden) {
        if (used.size >= 2) break;
        const key = SEARCH_KEYS.find((k) => !shown.includes(k) && !used.has(k) && fieldMatches(rec, k, t));
        if (!key) continue;
        used.add(key);
        const snip = el('span', 'snippet');
        snip.append(el('span', 'snippet-label', LABEL[key] + ': '));
        snip.append(hl(rec.text[key] || '', null));
        face.append(snip);
      }
    }

    face.append(buildDates(rec));
    face.addEventListener('click', () => toggle(rec.uid));
    return face;
  }

  function editButton(rec, withText) {
    const b = el('button', withText ? 'btn btn-ghost btn-sm edit-text' : 'edit-btn');
    b.type = 'button';
    b.title = 'Rediger';
    b.setAttribute('aria-label', 'Rediger ' + recordName(rec));
    b.append(iconSvg('pencil'));
    if (withText) b.append(el('span', null, 'Rediger'));
    b.addEventListener('click', (e) => {
      e.stopPropagation();
      openEditor({ mode: 'edit', rec });
    });
    return b;
  }

  /* ---------- Kort (utvidet) ---------- */

  function copyButton(value, label) {
    const btn = el('button', 'copy');
    btn.type = 'button';
    btn.title = 'Kopier ' + label;
    btn.setAttribute('aria-label', 'Kopier ' + label);
    btn.append(iconSvg('copy'));
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      try {
        await navigator.clipboard.writeText(value);
        btn.replaceChildren(iconSvg('check'));
        btn.classList.add('is-done');
        toast('Kopierte ' + label + ': ' + value);
        setTimeout(() => { btn.replaceChildren(iconSvg('copy')); btn.classList.remove('is-done'); }, 1400);
      } catch (err) {
        toast('Kunne ikke kopiere til utklippstavlen');
      }
    });
    return btn;
  }

  function noteValue(text) {
    const digits = /^\+?[\d\s]{8,}$/.test(text) ? text.replace(/\s+/g, '') : null;
    const n = digits ? digits.replace('+', '').length : 0;
    if (digits && n >= 8 && n <= 12) {
      const a = el('a', 'tel');
      a.href = 'tel:' + digits;
      a.append(iconSvg('phone'), hl(text, null));
      return a;
    }
    return hl(text, null);
  }

  function buildAlerts(rec) {
    const box = el('div', 'alerts');
    for (const g of DUE_GROUPS) {
      const st = rec.due[g.kind];
      if (!st || st === 'ok') continue;
      const a = el('div', 'alert alert-' + st);
      a.append(iconSvg(g.kind === 'ut' ? 'alert' : 'clock'));
      a.append(el('span', null, dueSentence(g.kind, st, rec.values[g.kind])));
      box.append(a);
    }
    return box.childNodes.length ? box : null;
  }

  function buildDetail(rec) {
    const d = el('div', 'detail');
    d.setAttribute('role', 'region');

    const head = el('div', 'detail-head');
    const mono = el('div', 'monogram', rec.text.init || (rec.vacant ? '—' : '?'));
    mono.setAttribute('aria-hidden', 'true');
    if (!rec.vacant) mono.append(genderBadge(rec));
    const titles = el('div', 'detail-titles');
    const h2 = el('h2', 'detail-name');
    h2.tabIndex = -1;
    h2.append(rec.text.navn ? hl(rec.text.navn, null) : (rec.vacant ? 'Ledig bolig' : 'Uten navn'));
    titles.append(h2);
    const sub = el('p', 'detail-sub');
    const parts = [rec.text.adresse, rec.text.bosted].filter(Boolean);
    parts.forEach((p, i) => {
      if (i) sub.append(el('span', 'dot', ' · '));
      sub.append(hl(p, null));
    });
    if (!rec.vacant) {
      if (parts.length) sub.append(el('span', 'dot', ' · '));
      sub.append(el('span', 'gender-text', genderTitle(rec)));
    }
    if (sub.childNodes.length) titles.append(sub);
    const actions = el('div', 'detail-actions');
    const close = el('button', 'close');
    close.type = 'button';
    close.setAttribute('aria-label', 'Lukk kortet');
    close.title = 'Lukk (Esc)';
    close.append(iconSvg('close'));
    close.addEventListener('click', () => toggle(rec.uid));
    actions.append(editButton(rec, true), close);
    head.append(mono, titles, actions);
    d.append(head);

    const alerts = buildAlerts(rec);
    if (alerts) d.append(alerts);

    const keys = el('div', 'keynums');
    for (const k of ['kunde', 'abb', 'art']) {
      const box = el('div', 'keynum');
      box.append(el('span', 'keynum-label', FIELD[k].label));
      const row = el('div', 'keynum-row');
      if (rec.text[k]) {
        row.append(hl(rec.text[k], 'keynum-val'));
        row.append(copyButton(rec.text[k], FIELD[k].label));
      } else {
        row.append(el('span', 'keynum-val is-missing', '—'));
      }
      box.append(row);
      keys.append(box);
    }
    d.append(keys);

    const grid = el('div', 'sections');
    for (const s of SECTIONS) {
      const sec = el('section', 'section');
      sec.append(el('h3', null, s.title));
      const dl = el('dl');
      for (const k of s.fields) {
        const item = el('div', 'field');
        const dd = el('dd');
        const raw = k === 'merknad' ? rec.merknad : rec.values[k];
        const txt = rec.text[k];
        if (!txt) {
          dd.append(el('span', 'is-missing', '—'));
        } else if (k === 'merknad') {
          const c = el('p', 'cell-comment');
          c.append(iconSvg('note'), hl(txt, 'cell-comment-text'));
          dd.append(c);
        } else if (FIELD[k].type === 'note') {
          dd.append(noteValue(txt));
        } else if (FIELD[k].type === 'status') {
          const pill = el('span', 'pill ' + (raw === true ? 'pill-on' : 'pill-off'));
          pill.append(hl(txt, null));
          dd.append(pill);
        } else {
          dd.append(hl(txt, null));
          if (isDate(raw)) dd.append(el('span', 'rel', relativeDate(raw)));
          const st = rec.due[k];
          if (st && st !== 'ok') dd.append(duePill(k, st, raw));
          if (rec.badDate[k]) dd.append(el('span', 'due due-bad', 'ugyldig dato'));
        }
        item.append(el('dt', null, LABEL[k]), dd);
        dl.append(item);
      }
      sec.append(dl);
      grid.append(sec);
    }
    d.append(grid);

    const foot = el('div', 'detail-foot');
    foot.append(el('span', null, rec.changed ? 'Sist endret ' + formatStamp(rec.changed) : 'Ikke endret i appen'));
    const same = el('button', 'link-btn');
    same.type = 'button';
    same.append(iconSvg('plus'), el('span', null, 'Ny bolig på samme adresse'));
    same.addEventListener('click', () => openEditor({ mode: 'new', prefill: prefillFrom(rec) }));
    foot.append(same);
    d.append(foot);
    return d;
  }

  function buildTile(rec) {
    const li = el('li', 'tile');
    li.dataset.uid = rec.uid;
    li.style.viewTransitionName = 'tile-' + rec.uid;
    if (rec.vacant) li.classList.add('is-vacant');
    if (rec.due.ut && rec.due.ut !== 'ok') li.classList.add('ut-' + rec.due.ut);
    if (state.flash.has(rec.uid)) li.classList.add('is-flash');
    if (state.openId === rec.uid) {
      li.classList.add('is-open');
      li.append(buildDetail(rec));
    } else {
      li.append(buildFace(rec), editButton(rec, false));
    }
    return li;
  }

  /* ---------- Gater ---------- */

  function groupStreets(list) {
    const groups = new Map();
    for (const r of list) {
      let g = groups.get(r.streetKey);
      if (!g) groups.set(r.streetKey, (g = { key: r.streetKey, name: r.street, units: [] }));
      g.units.push(r);
    }
    return [...groups.values()].sort((a, b) => collator.compare(a.name, b.name));
  }

  const slug = (s) => fold(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'x';

  function statusChips(units, kind) {
    const box = el('span', 'mini-chips');
    for (const st of DUE_ORDER) {
      const n = units.filter((u) => u.due[kind] === st).length;
      if (!n) continue;
      const g = DUE_GROUPS.find((x) => x.kind === kind);
      const c = el('span', 'mini-chip');
      c.title = g.title + ': ' + n + ' ' + plural(n, 'bolig', 'boliger') + ' – ' + (st === 'past' ? g.pastHint : DUE[st].hint);
      c.append(el('span', 'sw sw-' + st), String(n));
      box.append(c);
    }
    return box.childNodes.length ? box : null;
  }

  function buildStreetTile(g, index) {
    const li = el('li', 'tile street-tile');
    li.style.viewTransitionName = 'street-' + slug(g.key) + '-' + index;
    const worst = DUE_ORDER.slice(0, 3).find((st) => g.units.some((u) => u.due.ut === st));
    if (worst) li.classList.add('st-' + worst);

    const face = el('button', 'face street-face');
    face.type = 'button';
    face.setAttribute('aria-label', g.name + ', ' + g.units.length + ' ' + plural(g.units.length, 'bolig', 'boliger'));
    const top = el('span', 'face-top');
    const main = el('span', 'face-main');
    main.append(hl(g.name, 'street-name is-primary'));
    const tenants = g.units.filter((u) => !u.vacant).length;
    const vacant = g.units.length - tenants;
    main.append(el('span', 'street-sub',
      tenants + ' ' + plural(tenants, 'leietaker', 'leietakere') + (vacant ? ' · ' + vacant + ' ledig' : '')));
    top.append(main, el('span', 'street-count', String(g.units.length)));
    face.append(top);

    const houses = [...new Set(g.units.map((u) => u.house).filter(Boolean))].sort(collator.compare);
    if (houses.length) face.append(el('span', 'street-houses', 'Nr. ' + houses.join(', ')));

    const rows = el('span', 'street-status');
    for (const grp of DUE_GROUPS) {
      const chips = statusChips(g.units, grp.kind);
      if (!chips) continue;
      const line = el('span', 'street-status-row');
      line.append(el('span', 'date-label', grp.short), chips);
      rows.append(line);
    }
    if (rows.childNodes.length) face.append(rows);

    face.addEventListener('click', () => enterStreet(g.key));
    li.append(face);
    return li;
  }

  function enterStreet(key) {
    withTransition(() => {
      state.street = key;
      state.openId = null;
      render();
    }).then(() => {
      const top = els.viewbar.getBoundingClientRect().top;
      if (top < 0) els.viewbar.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
  }

  function leaveStreet() {
    withTransition(() => {
      state.street = null;
      state.openId = null;
      render();
    });
  }

  function setView(view) {
    if (view === state.view && state.street == null) return;
    withTransition(() => {
      state.view = view;
      state.street = null;
      state.openId = null;
      savePref('view', view);
      render();
    });
  }

  /* ---------- Oversikt over frister ---------- */

  function renderOverview(scope) {
    els.overview.replaceChildren();
    for (const g of DUE_GROUPS) {
      const group = el('div', 'ov-group');
      group.setAttribute('role', 'group');
      group.setAttribute('aria-label', g.title);
      group.append(el('span', 'ov-title', g.title));
      for (const st of DUE_ORDER) {
        const id = g.kind + ':' + st;
        const n = scope.filter((r) => r.due[g.kind] === st).length;
        const chip = el('button', 'chip');
        chip.type = 'button';
        chip.setAttribute('aria-pressed', String(state.filters.has(id)));
        chip.disabled = !n && !state.filters.has(id);
        chip.title = g.title + ': ' + (st === 'past' ? g.pastHint : DUE[st].hint) + ' – klikk for å filtrere';
        chip.append(el('span', 'sw sw-' + st));
        chip.append(el('span', null, st === 'past' ? g.pastLabel : DUE[st].label));
        chip.append(el('b', 'chip-n', String(n)));
        chip.addEventListener('click', () => {
          if (state.filters.has(id)) state.filters.delete(id);
          else state.filters.add(id);
          render();
        });
        group.append(chip);
      }
      els.overview.append(group);
    }
    if (state.filters.size) {
      const reset = el('button', 'ov-reset', 'Vis alle');
      reset.type = 'button';
      reset.addEventListener('click', () => { state.filters.clear(); render(); });
      els.overview.append(reset);
    }
    els.overview.hidden = false;
  }

  function renderViewbar(scope) {
    els.viewbar.hidden = false;
    els.viewKort.setAttribute('aria-pressed', String(state.view === 'kort'));
    els.viewGater.setAttribute('aria-pressed', String(state.view === 'gater'));
    els.crumbs.replaceChildren();
    const inStreet = state.view === 'gater' && state.street != null;
    els.crumbs.hidden = !inStreet;
    if (inStreet) {
      const back = el('button', 'crumb-back');
      back.type = 'button';
      back.append(iconSvg('back'), el('span', null, 'Alle gater'));
      back.addEventListener('click', leaveStreet);
      const anyUnit = state.records.find((r) => r.streetKey === state.street);
      const name = anyUnit ? anyUnit.street : state.street;
      const cur = el('span', 'crumb-cur', name);
      els.crumbs.append(back, el('span', 'crumb-sep', '/'), cur,
        el('span', 'crumb-n', scope.length + ' ' + plural(scope.length, 'bolig', 'boliger')));
    }
    els.addBtn.querySelector('.add-label').textContent = inStreet ? 'Ny bolig i denne gaten' : 'Ny bolig';
  }

  /* ---------- Tegning ---------- */

  function render() {
    const scope = scopeRecords();
    const list = scope.filter(passesFilters);
    renderOverview(scope);
    renderViewbar(scope);

    const frag = document.createDocumentFragment();
    const streetList = state.view === 'gater' && state.street == null;
    let shown = 0;
    if (streetList) {
      const groups = groupStreets(list);
      groups.forEach((g, i) => frag.append(buildStreetTile(g, i)));
      shown = groups.length;
    } else {
      for (const rec of sortList(list)) frag.append(buildTile(rec));
      shown = list.length;
    }
    els.grid.replaceChildren(frag);
    els.grid.classList.toggle('streets', streetList);
    els.grid.classList.toggle('only-initials', state.show.length === 1 && state.show[0] === 'init');

    const total = state.records.filter((r) => state.showVacant || !r.vacant).length;
    if (streetList) {
      els.count.textContent = shown + ' ' + plural(shown, 'gate', 'gater') + ' · ' + list.length + ' ' + plural(list.length, 'bolig', 'boliger');
    } else {
      els.count.textContent = state.tokens.length || state.filters.size || state.street != null
        ? 'Viser ' + shown + ' av ' + total
        : total + ' kort';
    }

    const noResults = state.records.length > 0 && !shown;
    els.empty.hidden = !noResults;
    if (noResults) {
      els.emptyText.textContent = state.query.trim()
        ? 'Ingen treff for «' + state.query.trim() + '».'
        : state.filters.size ? 'Ingen kort passer til filteret.' : 'Ingen boliger her ennå.';
    }

    if (state.flash.size) {
      const ids = [...state.flash];
      setTimeout(() => ids.forEach((id) => state.flash.delete(id)), 50);
    }
  }

  function updateSource() {
    const tenants = state.records.filter((r) => !r.vacant).length;
    const vacant = state.records.length - tenants;
    els.subtitle.textContent = state.meta.title || 'Leietakere';
    let origin;
    if (state.hasStore && state.meta.savedAt) origin = 'Lagret i denne nettleseren ' + formatStamp(state.meta.savedAt);
    else origin = 'Innebygde data' + (state.meta.source ? ' (' + state.meta.source + ')' : '');
    els.source.textContent = origin + ' · ' + tenants + ' ' + plural(tenants, 'leietaker', 'leietakere') +
      (vacant ? ' · ' + vacant + ' ' + plural(vacant, 'ledig bolig', 'ledige boliger') : '');
    els.vacantCount.textContent = vacant ? '(' + vacant + ')' : '';
    const pending = state.hasStore && state.meta.savedAt &&
      (!state.meta.exportedAt || state.meta.exportedAt < state.meta.savedAt);
    els.exportBtn.classList.toggle('has-dot', !!pending);
    els.exportBtn.title = pending
      ? 'Det finnes endringer som ikke er eksportert til Excel ennå'
      : 'Last ned alle data som Excel-fil';
    els.resetCmd.disabled = !state.hasStore || !window.INNEBYGDE_DATA;
  }

  function renderAll() {
    updateSource();
    onSearch();
  }

  // Jevn animasjon (View Transitions API) der nettleseren støtter det.
  const vtStyle = document.createElement('style');
  document.head.append(vtStyle);

  function withTransition(update, morphIds) {
    const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (!document.startViewTransition || reduce) {
      update();
      return Promise.resolve();
    }
    vtStyle.textContent = (morphIds || []).filter((id) => id != null).map((id) =>
      '::view-transition-old(tile-' + id + '){animation:vt-fade-out .14s ease-out both;mix-blend-mode:normal}' +
      '::view-transition-new(tile-' + id + '){animation:vt-fade-in .16s ease-out both;mix-blend-mode:normal}'
    ).join('');
    const vt = document.startViewTransition(update);
    return vt.finished.catch(() => {}).then(() => { vtStyle.textContent = ''; });
  }

  function toggle(uid) {
    const opening = state.openId !== uid;
    const prev = state.openId;
    const done = withTransition(() => {
      state.openId = opening ? uid : null;
      for (const pid of [prev, uid]) {
        if (pid == null) continue;
        const li = els.grid.querySelector('.tile[data-uid="' + pid + '"]');
        const rec = byUid(pid);
        if (li && rec) li.replaceWith(buildTile(rec));
      }
    }, [prev, uid]);
    done.then(() => {
      const li = els.grid.querySelector('.tile[data-uid="' + uid + '"]');
      if (!li) return;
      if (opening) {
        li.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        const h = li.querySelector('.detail-name');
        if (h) h.focus({ preventScroll: true });
      } else {
        const face = li.querySelector('.face');
        if (face) face.focus({ preventScroll: true });
      }
    });
  }

  function scrollToUid(uid) {
    requestAnimationFrame(() => {
      const li = els.grid.querySelector('.tile[data-uid="' + uid + '"]');
      if (li) li.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    });
  }

  /* ---------- Redigering ---------- */

  const EDIT_SECTIONS = [
    { title: 'Leietaker', fields: [
      { key: 'navn', wide: true, placeholder: 'Etternavn Fornavn' },
      { key: 'init', placeholder: 'Fylles ut automatisk' },
      { key: 'kjonn', kind: 'select', label: 'Kjønn' },
      { key: 'status', kind: 'checkbox', label: 'Aktiv' },
      { key: 'kunde', inputmode: 'numeric' },
      { key: 'abb', inputmode: 'numeric' },
      { key: 'status2' },
      { key: 'komm', label: 'Telefon / kommentar', wide: true },
    ] },
    { title: 'Bolig', cls: 'ed-sec-bolig', fields: [
      { key: 'adresse', wide: true, required: true, placeholder: 'Gate og nummer, f.eks. Eldridhågen 24' },
      { key: 'bosted' },
      { key: 'bolig' },
      { key: 'art', label: 'ARTnr' },
      { key: 'antall', kind: 'count', label: 'Antall boliger', createOnly: true },
    ] },
    { title: 'Datoer', fields: [
      { key: 'inn', kind: 'date' },
      { key: 'ut', kind: 'date', label: 'Dato utflytting (kontraktslutt)' },
      { key: 'reg', kind: 'date', label: 'Neste regulering (KPI)' },
    ] },
    { title: 'Betaling', fields: [
      { key: 'belop1', kind: 'money' },
      { key: 'belop2', kind: 'money' },
    ] },
    { title: 'Merknad', fields: [
      { key: 'merknad', kind: 'textarea', wide: true, label: 'Merknad', placeholder: 'F.eks. «Deler med NAV»' },
    ] },
  ];

  const ed = { mode: null, rec: null, inputs: {}, initTouched: false };

  function buildEditorForm() {
    for (const s of EDIT_SECTIONS) {
      const fs = el('fieldset', 'ed-section' + (s.cls ? ' ' + s.cls : ''));
      fs.append(el('legend', null, s.title));
      const grid = el('div', 'ed-grid');
      for (const f of s.fields) {
        const id = 'ed-' + f.key;
        const label = f.label || LABEL[f.key] || f.key;
        let wrap, input;
        if (f.kind === 'checkbox') {
          wrap = el('label', 'ed-check');
          input = el('input');
          input.type = 'checkbox';
          wrap.append(input, el('span', 'switch-track'), el('span', null, label));
        } else {
          wrap = el('label', 'ed-field' + (f.wide ? ' is-wide' : ''));
          wrap.append(el('span', 'ed-label', label + (f.required ? ' *' : '')));
          if (f.kind === 'select') {
            input = el('select');
          } else if (f.kind === 'textarea') {
            input = el('textarea');
            input.rows = 2;
          } else {
            input = el('input');
            input.type = f.kind === 'date' ? 'date' : f.kind === 'count' ? 'number' : 'text';
            if (f.kind === 'count') { input.min = '1'; input.max = '50'; input.step = '1'; }
            if (f.kind === 'money') { input.inputMode = 'decimal'; input.placeholder = '0,00'; }
            if (f.inputmode) input.inputMode = f.inputmode;
          }
          if (f.placeholder) input.placeholder = f.placeholder;
          if (f.kind === 'money') {
            const box = el('span', 'ed-suffix');
            box.append(input, el('span', null, 'kr'));
            wrap.append(box);
          } else {
            wrap.append(input);
          }
          const hint = el('span', 'ed-hint');
          hint.id = id + '-hint';
          wrap.append(hint);
        }
        input.id = id;
        input.name = f.key;
        input.autocomplete = 'off';
        if (f.createOnly) wrap.classList.add('create-only');
        ed.inputs[f.key] = input;
        grid.append(wrap);
      }
      fs.append(grid);
      els.editorBody.append(fs);
    }

    ed.inputs.navn.addEventListener('input', () => {
      if (!ed.initTouched) ed.inputs.init.value = autoInitials(ed.inputs.navn.value);
      if (ed.inputs.navn.value.trim() && !ed.inputs.status.checked && ed.wasVacant) ed.inputs.status.checked = true;
      updateGenderOptions();
    });
    ed.inputs.init.addEventListener('input', () => {
      ed.initTouched = ed.inputs.init.value.trim() !== '' &&
        ed.inputs.init.value !== autoInitials(ed.inputs.navn.value);
    });
    for (const k of ['inn', 'ut', 'reg']) {
      ed.inputs[k].addEventListener('input', () => { ed.inputs[k].dataset.touched = '1'; setHint(k, ''); });
    }
  }

  function setHint(key, text, isError) {
    const h = document.getElementById('ed-' + key + '-hint');
    if (!h) return;
    h.textContent = text || '';
    h.classList.toggle('is-error', !!isError);
    const input = ed.inputs[key];
    if (input) input.setAttribute('aria-invalid', isError ? 'true' : 'false');
  }

  function updateGenderOptions() {
    const sel = ed.inputs.kjonn;
    const keep = sel.value;
    const guess = guessGender(ed.inputs.navn.value);
    const guessText = guess ? ' (antatt: ' + GENDER_INFO[guess].text.toLowerCase() + ')' : '';
    sel.replaceChildren();
    for (const [value, text] of [['', 'Automatisk' + guessText], ['K', 'Kvinne'], ['M', 'Mann'], ['org', 'Tjeneste / organisasjon']]) {
      const o = el('option', null, text);
      o.value = value;
      sel.append(o);
    }
    sel.value = keep;
  }

  const moneyToInput = (v) => (typeof v === 'number' ? String(v).replace('.', ',') : (v == null ? '' : String(v)));

  function parseMoney(s) {
    const t = s.replace(/\s|kr|nok/gi, '').replace(/ /g, '');
    if (!t) return { ok: true, v: null };
    let norm = t;
    if (/,\d{1,2}$/.test(norm)) norm = norm.replace(/\./g, '').replace(',', '.');
    else norm = norm.replace(/,/g, '');
    const n = Number(norm);
    return isFinite(n) ? { ok: true, v: n } : { ok: false };
  }

  function prefillFrom(rec) {
    return {
      adresse: rec.values.adresse, bosted: rec.values.bosted, bolig: rec.values.bolig,
      art: rec.values.art, belop1: rec.values.belop1 ?? rec.values.belop2,
    };
  }

  function prefillForStreet() {
    if (state.view !== 'gater' || state.street == null) return {};
    const units = state.records.filter((r) => r.streetKey === state.street);
    if (!units.length) return {};
    const count = new Map();
    for (const u of units) if (u.values.bosted) count.set(u.values.bosted, (count.get(u.values.bosted) || 0) + 1);
    const bosted = [...count.entries()].sort((a, b) => b[1] - a[1])[0];
    return { adresse: units[0].street + ' ', bosted: bosted ? bosted[0] : null };
  }

  function openEditor({ mode, rec, prefill }) {
    ed.mode = mode;
    ed.rec = rec || null;
    const values = rec ? rec.values : Object.assign(emptyValues(), prefill || {});
    ed.wasVacant = rec ? rec.vacant : true;
    els.editor.classList.toggle('is-create', mode === 'new');

    els.editorTitle.textContent = mode === 'new' ? 'Ny bolig' : 'Rediger ' + (rec.text.adresse || 'bolig');
    els.editorSub.textContent = mode === 'new'
      ? 'Fyll inn adressen. Uten navn blir boligen registrert som ledig.'
      : (rec.vacant ? 'Ledig bolig' : (rec.text.navn || '')) + (rec.text.bolig ? ' · ' + rec.text.bolig : '');

    for (const f of FIELDS) {
      const input = ed.inputs[f.key];
      if (!input) continue;
      const v = values[f.key];
      setHint(f.key, '');
      input.dataset.touched = '';
      if (f.type === 'status') input.checked = v === true;
      else if (f.type === 'date') {
        input.value = isDate(v) ? isoDate(v) : '';
        if (typeof v === 'string' && v.trim()) setHint(f.key, 'Ugyldig dato i dataene: «' + v + '». Velg en ny dato.', true);
      } else if (f.type === 'money') input.value = moneyToInput(v);
      else input.value = v == null ? '' : String(v);
    }
    ed.inputs.merknad.value = rec ? rec.merknad || '' : '';
    ed.inputs.antall.value = '1';
    setHint('antall', 'Flere boliger med samme adresse opprettes på én gang.');
    setHint('init', 'Fylles ut automatisk fra navnet.');
    ed.initTouched = !!ed.inputs.init.value && ed.inputs.init.value !== autoInitials(ed.inputs.navn.value);
    updateGenderOptions();
    ed.inputs.kjonn.value = rec && rec.kjonn ? rec.kjonn : '';

    els.vacateBtn.hidden = mode === 'new' || rec.vacant;
    els.deleteBtn.hidden = mode === 'new';
    els.saveBtn.textContent = mode === 'new' ? 'Opprett' : 'Lagre';

    setMenu(false);
    setMoreMenu(false);
    els.editor.showModal();
    els.editorBody.scrollTop = 0;
    const focusKey = mode === 'new' ? 'adresse' : 'navn';
    // På berøringsskjermer åpnes ikke tastaturet automatisk.
    if (window.matchMedia('(pointer: coarse)').matches) return;
    requestAnimationFrame(() => {
      const input = ed.inputs[focusKey];
      input.focus();
      if (mode === 'new' && input.value) input.setSelectionRange(input.value.length, input.value.length);
    });
  }

  function closeEditor() {
    if (els.editor.open) els.editor.close();
  }

  function readForm() {
    const out = {};
    let firstError = null;
    const fail = (key, msg) => { setHint(key, msg, true); if (!firstError) firstError = key; };

    for (const f of FIELDS) {
      const input = ed.inputs[f.key];
      const old = ed.rec ? ed.rec.values[f.key] : null;
      if (f.type === 'status') { out[f.key] = input.checked; continue; }
      const raw = input.value.trim();
      if (f.type === 'date') {
        if (!raw) {
          // Ugyldig tekst fra Excel beholdes til en ny dato er valgt.
          out[f.key] = typeof old === 'string' && !input.dataset.touched ? old : null;
        } else {
          const d = parseDateText(raw);
          if (!d) fail(f.key, 'Ugyldig dato.');
          out[f.key] = d;
        }
        continue;
      }
      if (f.type === 'money') {
        const m = parseMoney(raw);
        if (!m.ok) fail(f.key, 'Ugyldig beløp – skriv f.eks. 12500 eller 9321,75.');
        out[f.key] = m.v;
        continue;
      }
      if (!raw) { out[f.key] = null; continue; }
      if ((f.key === 'kunde' || f.key === 'abb') && /^\d{1,15}$/.test(raw)) { out[f.key] = Number(raw); continue; }
      if (f.key === 'komm' && typeof old === 'number' && raw === String(old)) { out[f.key] = old; continue; }
      out[f.key] = raw;
    }
    if (isEmpty(out.adresse)) fail('adresse', 'Adressen må fylles ut.');
    if (!isEmpty(out.navn) && isEmpty(out.init)) out.init = autoInitials(out.navn);

    const count = Math.round(Number(ed.inputs.antall.value) || 1);
    if (ed.mode === 'new' && (count < 1 || count > 50)) fail('antall', 'Velg mellom 1 og 50.');

    return {
      values: out,
      kjonn: ed.inputs.kjonn.value || null,
      merknad: ed.inputs.merknad.value.trim(),
      count: Math.min(50, Math.max(1, count)),
      firstError,
    };
  }

  function saveEditor() {
    const form = readForm();
    if (form.firstError) {
      ed.inputs[form.firstError].focus();
      return;
    }
    const wasVacantHidden = !state.showVacant;

    if (ed.mode === 'new') {
      const created = mutate(form.count === 1 ? 'Ny bolig er opprettet.' : form.count + ' nye boliger er opprettet.', (now) => {
        const list = [];
        for (let i = 1; i <= form.count; i++) {
          const values = Object.assign({}, form.values);
          if (form.count > 1) values.bolig = values.bolig ? values.bolig + ' – ' + i : 'Leilighet ' + i;
          const rec = newRecord(values, { kjonn: form.kjonn, merknad: form.merknad, changed: now });
          state.records.push(rec);
          list.push(rec);
        }
        list.forEach((r) => state.flash.add(r.uid));
        if (state.view === 'gater') state.street = list[0].streetKey;
        if (list.some((r) => r.vacant) && wasVacantHidden) showVacant(true);
        state.filters.clear();
        return list;
      });
      closeEditor();
      clearSearchIfHiding(created);
      scrollToUid(created[0].uid);
      return;
    }

    const rec = ed.rec;
    mutate('Endringene i ' + recordName(rec) + ' er lagret.', (now) => {
      rec.values = form.values;
      rec.kjonn = form.kjonn;
      rec.merknad = form.merknad;
      rec.changed = now;
      derive(rec);
      state.flash.add(rec.uid);
      if (rec.vacant && wasVacantHidden) showVacant(true);
      // Ny gate? Følg med til den nye gaten.
      if (state.view === 'gater' && state.street != null && rec.streetKey !== state.street) state.street = rec.streetKey;
    });
    closeEditor();
    scrollToUid(rec.uid);
  }

  function vacate() {
    const rec = ed.rec;
    if (!rec) return;
    const wasVacantHidden = !state.showVacant;
    mutate(recordName(rec) + ' er markert som ledig.', (now) => {
      const v = rec.values;
      // Leien beholdes som «Beløp 1», slik som for de andre ledige boligene.
      if (v.belop1 == null && typeof v.belop2 === 'number') v.belop1 = v.belop2;
      v.belop2 = null;
      for (const k of TENANT_KEYS) v[k] = null;
      v.status = false;
      rec.kjonn = null;
      rec.merknad = '';
      rec.changed = now;
      derive(rec);
      state.flash.add(rec.uid);
      if (wasVacantHidden) showVacant(true);
    });
    closeEditor();
    scrollToUid(rec.uid);
  }

  function removeRecord() {
    const rec = ed.rec;
    if (!rec) return;
    mutate(recordName(rec) + ' er slettet.', () => {
      state.records = state.records.filter((r) => r !== rec);
      if (state.openId === rec.uid) state.openId = null;
    });
    closeEditor();
  }

  function showVacant(on) {
    state.showVacant = on;
    els.vacant.checked = on;
    savePref('showVacant', on);
  }

  // Hvis et søk skjuler de nye kortene, tømmes søket.
  function clearSearchIfHiding(list) {
    if (state.tokens.length && !list.some((r) => recordMatches(r, state.tokens))) {
      els.search.value = '';
      onSearch();
    }
  }

  /* ---------- Valg av felt på kortene ---------- */

  function buildShowMenu() {
    els.showMenu.replaceChildren();
    for (const o of SHOW_OPTIONS) {
      const label = el('label', 'menu-item');
      const cb = el('input');
      cb.type = 'checkbox';
      cb.value = o.id;
      cb.checked = state.show.includes(o.id);
      cb.addEventListener('change', () => {
        const next = SHOW_OPTIONS.map((x) => x.id).filter((id) =>
          id === o.id ? cb.checked : state.show.includes(id));
        if (!next.length) { cb.checked = true; return; }
        withTransition(() => {
          state.show = next;
          savePref('show', next);
          updateShowValue();
          render();
        });
      });
      label.append(cb, el('span', 'menu-check'), el('span', null, o.label));
      els.showMenu.append(label);
    }
    els.showMenu.append(el('p', 'menu-hint', 'Du kan velge flere felt samtidig. Datoene vises alltid.'));
    updateShowValue();
  }

  function updateShowValue() {
    const chosen = SHOW_OPTIONS.filter((o) => state.show.includes(o.id));
    els.showValue.textContent = chosen.length === 1 ? chosen[0].label : chosen.map((o) => o.short).join(' + ');
  }

  function setMenu(open) {
    els.showMenu.hidden = !open;
    els.showBtn.setAttribute('aria-expanded', String(open));
    if (open) {
      const firstInput = els.showMenu.querySelector('input');
      if (firstInput) firstInput.focus();
    }
  }

  function setMoreMenu(open) {
    els.moreMenu.hidden = !open;
    els.moreBtn.setAttribute('aria-expanded', String(open));
    if (open) {
      const first = els.moreMenu.querySelector('button:not(:disabled)');
      if (first) first.focus();
    }
  }

  /* ---------- Melding ---------- */

  let toastTimer = null;
  function toast(msg, opts) {
    els.toast.replaceChildren(el('span', null, msg));
    if (opts && opts.action) {
      const b = el('button', 'toast-action', opts.action);
      b.type = 'button';
      b.addEventListener('click', () => {
        els.toast.classList.remove('is-visible');
        opts.run();
      });
      els.toast.append(b);
    }
    els.toast.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove('is-visible'), (opts && opts.ms) || 2600);
  }

  /* ---------- Hendelser ---------- */

  function onSearch() {
    state.query = els.search.value;
    const tokens = tokenize(state.query);
    const phrase = tokens.join(' ');
    const phraseHit = tokens.length > 1 && state.records.some((r) =>
      (state.showVacant || !r.vacant) && SEARCH_KEYS.some((k) => fieldMatches(r, k, phrase)));
    state.tokens = phraseHit ? [phrase] : tokens;
    render();
  }

  els.search.addEventListener('input', onSearch);
  els.clearSearch.addEventListener('click', () => {
    els.search.value = '';
    state.filters.clear();
    onSearch();
    els.search.focus();
  });

  els.showBtn.addEventListener('click', () => setMenu(els.showMenu.hidden));
  els.moreBtn.addEventListener('click', () => setMoreMenu(els.moreMenu.hidden));
  document.addEventListener('click', (e) => {
    if (!els.showMenu.hidden && !e.target.closest('.show-picker')) setMenu(false);
    if (!els.moreMenu.hidden && !e.target.closest('.more')) setMoreMenu(false);
  });

  els.sort.value = state.sort;
  els.sort.addEventListener('change', () => {
    withTransition(() => {
      state.sort = els.sort.value;
      savePref('sort', state.sort);
      render();
    });
  });

  els.vacant.checked = !!state.showVacant;
  els.vacant.addEventListener('change', () => {
    showVacant(els.vacant.checked);
    onSearch();
  });

  els.viewKort.addEventListener('click', () => setView('kort'));
  els.viewGater.addEventListener('click', () => setView('gater'));
  els.addBtn.addEventListener('click', () => openEditor({ mode: 'new', prefill: prefillForStreet() }));

  els.exportBtn.addEventListener('click', exportExcel);
  els.importCmd.addEventListener('click', () => { setMoreMenu(false); els.fileInput.click(); });
  els.resetCmd.addEventListener('click', () => {
    setMoreMenu(false);
    if (!window.INNEBYGDE_DATA) return;
    if (!confirm('Alle endringer som er gjort i denne nettleseren, forkastes, og de innebygde dataene lastes inn på nytt. Fortsette?')) return;
    mutate('De innebygde dataene er lastet inn på nytt.', () => {
      applyData(window.INNEBYGDE_DATA);
      state.openId = null;
      state.street = null;
    });
    clearStore();
    state.meta.savedAt = null;
    updateSource();
  });

  els.fileInput.addEventListener('change', () => {
    loadFile(els.fileInput.files[0]);
    els.fileInput.value = '';
  });
  els.pickBtn.addEventListener('click', () => els.fileInput.click());

  els.editorForm.addEventListener('submit', (e) => { e.preventDefault(); saveEditor(); });
  els.cancelBtn.addEventListener('click', closeEditor);
  els.editorClose.addEventListener('click', closeEditor);
  els.vacateBtn.addEventListener('click', vacate);
  els.deleteBtn.addEventListener('click', removeRecord);
  els.editor.addEventListener('click', (e) => { if (e.target === els.editor) closeEditor(); });

  document.addEventListener('keydown', (e) => {
    if (els.editor.open) return; // dialogen håndterer selv Esc
    if (e.key === 'Escape') {
      if (!els.showMenu.hidden) { setMenu(false); els.showBtn.focus(); return; }
      if (!els.moreMenu.hidden) { setMoreMenu(false); els.moreBtn.focus(); return; }
      if (state.openId != null) { toggle(state.openId); return; }
      if (document.activeElement === els.search && els.search.value) { els.search.value = ''; onSearch(); return; }
      if (state.view === 'gater' && state.street != null) leaveStreet();
      return;
    }
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);
    if (e.key === '/' && !typing && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      els.search.focus();
      els.search.select();
    }
  });

  let dragDepth = 0;
  window.addEventListener('dragenter', (e) => {
    if (!e.dataTransfer || !Array.from(e.dataTransfer.types).includes('Files')) return;
    dragDepth++;
    document.body.classList.add('is-dragging');
  });
  window.addEventListener('dragleave', () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) document.body.classList.remove('is-dragging');
  });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    document.body.classList.remove('is-dragging');
    const file = e.dataTransfer && e.dataTransfer.files[0];
    if (file) loadFile(file);
  });

  if ('ResizeObserver' in window) {
    new ResizeObserver(() => {
      document.documentElement.style.setProperty('--toolbar-h', els.toolbar.offsetHeight + 'px');
    }).observe(els.toolbar);
  }

  const narrow = window.matchMedia('(max-width: 640px)');
  const setPlaceholder = () => {
    els.search.placeholder = narrow.matches
      ? 'Søk: navn, adresse, nummer…'
      : 'Søk: navn, initialer, adresse, nummer, telefon…';
  };
  setPlaceholder();
  if (narrow.addEventListener) narrow.addEventListener('change', setPlaceholder);

  buildShowMenu();
  buildEditorForm();
  init().then(updateSource);

  // For testing og feilsøking.
  window.__leietakere = { snapshot: () => JSON.parse(snapshot()), importBuffer };
})();

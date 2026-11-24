/* Leietakere — kort med data fra Excel-filen (Data.xlsx). */
(function () {
  'use strict';

  /* ---------- Kolonner ----------
   * Kolonnene finnes ut fra overskriften; hvis overskriften endres,
   * brukes kolonnebokstaven som reserve.
   */
  const FIELDS = [
    { key: 'art',     col: 'A', headers: ['artnr'],                          label: 'ARTnr' },
    { key: 'abb',     col: 'B', headers: ['abb.nr.', 'abbnr'],               label: 'Abb.nr.' },
    { key: 'kunde',   col: 'C', headers: ['kundenr.', 'kundenr'],            label: 'Kundenr.' },
    { key: 'status',  col: 'D', headers: ['status'],                         label: 'Status', type: 'status' },
    { key: 'belop1',  col: 'E', headers: ['beløp1'],                         label: 'Beløp 1', type: 'money' },
    { key: 'belop2',  col: 'F', headers: ['beløp2'],                         label: 'Beløp 2', type: 'money' },
    { key: 'navn',    col: 'G', headers: ['navn'],                           label: 'Navn' },
    { key: 'init',    col: 'H', headers: ['kolumna1', 'initialer'],          label: 'Initialer', type: 'initials' },
    { key: 'bosted',  col: 'I', headers: ['bosted'],                         label: 'Bosted' },
    { key: 'bolig',   col: 'J', headers: ['bolig/leilighet'],                label: 'Bolig/leilighet' },
    { key: 'adresse', col: 'K', headers: ['adresse'],                        label: 'Adresse' },
    { key: 'inn',     col: 'L', headers: ['datoinnflytting'],                label: 'Dato innflytting', type: 'date' },
    { key: 'ut',      col: 'M', headers: ['datoutflytting'],                 label: 'Dato utflytting', type: 'date' },
    { key: 'reg',     col: 'N', headers: ['nesteregulering'],                label: 'Neste regulering', type: 'date' },
    { key: 'status2', col: 'O', headers: ['status2'],                        label: 'Status 2' },
    { key: 'komm',    col: 'P', headers: ['kommentar/telefomnr.', 'kommentar/telefonnr.'], label: 'Kommentar / telefonnr.', type: 'note' },
  ];
  const FIELD = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

  // Valgfri kolonne «Kjønn» (M/K) — overstyrer gjetningen ut fra fornavn.
  const GENDER_HEADERS = ['kjønn', 'kjonn', 'kjønn(m/k)', 'kjønn(k/m)', 'gender', 'płeć', 'plec'];

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

  // Frister: hvor lenge er det igjen til datoen?
  const DUE = {
    m1:   { label: 'under 1 md.', hint: 'mindre enn 1 måned igjen' },
    m2:   { label: '1–2 md.',     hint: '1–2 måneder igjen' },
    m3:   { label: '2–3 md.',     hint: '2–3 måneder igjen' },
    past: { label: null,          hint: null },
  };
  const DUE_GROUPS = [
    { kind: 'ut',  title: 'Kontrakt utløper', pastLabel: 'Utløpt',  pastHint: 'kontrakten har allerede utløpt' },
    { kind: 'reg', title: 'KPI-regulering',   pastLabel: 'Forfalt', pastHint: 'reguleringsdatoen er passert' },
  ];

  // Seksjoner i det utvidede kortet.
  const SECTIONS = [
    { title: 'Bolig',                fields: ['adresse', 'bosted', 'bolig'] },
    { title: 'Kontrakt og datoer',   fields: ['inn', 'ut', 'reg', 'status', 'status2'] },
    { title: 'Betaling',             fields: ['belop1', 'belop2'] },
    { title: 'Kontakt og merknader', fields: ['komm'] },
  ];

  const SORTS = ['file', 'alpha', 'ut', 'reg'];

  /* ---------- Tilstand ---------- */

  const state = {
    records: [],
    show: loadPref('show', DEFAULT_SHOW).filter((id) => SHOW_OPTIONS.some((o) => o.id === id)),
    sort: loadPref('sort', 'file'),
    showVacant: loadPref('showVacant', false),
    filters: new Set(),   // f.eks. 'ut:m1', 'reg:past'
    query: '',
    tokens: [],
    openId: null,
    sheetName: '',
    hasGenderColumn: false,
  };
  if (!state.show.length) state.show = DEFAULT_SHOW.slice();
  if (!SORTS.includes(state.sort)) state.sort = 'file';

  const $ = (id) => document.getElementById(id);
  const els = {
    subtitle: $('subtitle'),
    source: $('source'),
    fileInput: $('fileInput'),
    search: $('search'),
    showBtn: $('showBtn'),
    showValue: $('showValue'),
    showMenu: $('showMenu'),
    sort: $('sort'),
    vacant: $('showVacant'),
    vacantCount: $('vacantCount'),
    count: $('count'),
    overview: $('overview'),
    grid: $('grid'),
    empty: $('empty'),
    emptyText: $('emptyText'),
    clearSearch: $('clearSearch'),
    loader: $('loader'),
    loaderTitle: $('loaderTitle'),
    loaderText: $('loaderText'),
    pickBtn: $('pickBtn'),
    toolbar: $('toolbar'),
    toast: $('toast'),
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

  // Dagens dato som UTC-midnatt (datoene fra Excel er lagret slik).
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

  function daysUntil(d) {
    return Math.round((d.getTime() - today()) / 86400000);
  }

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
    const abs = Math.abs(days);
    if (abs < 45) return dayPhrase(days);
    const months = Math.round(days / 30.44);
    if (Math.abs(months) < 18) return relFmt.format(months, 'month');
    return relFmt.format(Math.round(days / 365.25), 'year');
  }

  // Merkelapp på kortet: «om 7 d.» / «utløpt» / «forfalt».
  function duePill(kind, status, d) {
    const pill = el('span', 'due due-' + status);
    if (status === 'past') {
      pill.textContent = kind === 'ut' ? 'utløpt' : 'forfalt';
    } else {
      const days = daysUntil(d);
      if (days === 0) pill.textContent = 'i dag';
      else pill.append(el('span', 'hide-sm', 'om\u00a0'), days + ' d.');
    }
    return pill;
  }

  // Lengre tekst til det utvidede kortet / verktøytips.
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

  // Tekst som vises for en verdi (null = ingen verdi).
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

  // Dato skrevet som tekst i Excel (f.eks. «31.04.2027»): gyldig → Date, ellers null.
  function parseDateText(s) {
    const m = /^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})\s*$/.exec(s);
    if (!m) return null;
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const mo = parseInt(m[2], 10) - 1, d = parseInt(m[1], 10);
    const date = new Date(Date.UTC(y, mo, d));
    return date.getUTCMonth() === mo && date.getUTCDate() === d ? date : null;
  }

  /* ---------- Søk ----------
   * Sammenligning uten store/små bokstaver og diakritiske tegn (ø→o, æ→ae, å→a …),
   * og for numre også uten mellomrom («45254373» finner «452 54 373»).
   */
  const SPECIAL = { 'æ': 'ae', 'ø': 'o', 'å': 'a', 'ß': 'ss', 'ł': 'l', 'đ': 'd', 'œ': 'oe' };

  function foldChar(ch) {
    const lower = ch.toLowerCase();
    if (SPECIAL[lower]) return SPECIAL[lower];
    return lower.normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  // Returnerer «foldet» tekst og et kart tilbake til posisjonene i originalteksten.
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

  // Områder [start, slutt) i originalteksten som matcher et av søkeordene.
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

  function recordMatches(rec, tokens) {
    return tokens.every((t) => FIELDS.some((f) => fieldMatches(rec, f.key, t)));
  }

  /* ---------- Kjønn ---------- */

  function parseGender(v) {
    if (isEmpty(v)) return null;
    const s = fold(String(v)).trim();
    if (/^(k|kv|kvinne|f|female|w|woman|kobieta|d|dame)\b/.test(s)) return 'K';
    if (/^(m|mann|male|man|mezczyzna|herre|gutt)\b/.test(s)) return 'M';
    return null;
  }

  function genderOf(rec, excelValue) {
    if (rec.vacant) return { value: null, source: null };
    const fromExcel = parseGender(excelValue);
    if (fromExcel) return { value: fromExcel, source: 'excel' };
    const guess = rec.values.navn ? NavnKjonn.gjett(String(rec.values.navn), fold) : null;
    return { value: guess, source: guess ? 'guess' : null };
  }

  /* ---------- Data ---------- */

  function headerKey(v) {
    return typeof v === 'string' ? v.toLowerCase().replace(/\s+/g, '') : '';
  }

  function buildRecords(sheet) {
    const rows = sheet.rows;
    let h = rows.findIndex((r, i) => i < 40 &&
      r.some((v) => headerKey(v) === 'navn') && r.some((v) => headerKey(v) === 'adresse'));
    if (h < 0) h = 2; // standard: rad 3
    const headerRow = rows[h] || [];

    const colOf = {};
    for (const f of FIELDS) {
      let idx = headerRow.findIndex((v) => f.headers.includes(headerKey(v)));
      if (idx < 0) idx = XlsxReader.colIndex(f.col);
      colOf[f.key] = idx;
    }
    const genderCol = headerRow.findIndex((v) => GENDER_HEADERS.includes(headerKey(v)));

    let title = '';
    for (let i = 0; i < h && !title; i++) {
      const cell = (rows[i] || []).find((v) => typeof v === 'string' && v.trim());
      if (cell) title = cell.trim();
    }

    const records = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || FIELDS.every((f) => isEmpty(row[colOf[f.key]]))) continue;
      const rec = { id: records.length, excelRow: r + 1, values: {}, text: {}, comments: {}, search: {}, badDate: {} };
      for (const f of FIELDS) {
        let v = row[colOf[f.key]];
        if (typeof v === 'string') v = v.trim() || null;
        if (f.type === 'date' && typeof v === 'string') {
          const parsed = parseDateText(v);
          if (parsed) v = parsed;
          else rec.badDate[f.key] = true;
        }
        rec.values[f.key] = v ?? null;
        const txt = display(f.key, v);
        rec.text[f.key] = txt;
        const cm = sheet.comments.get(r + ',' + colOf[f.key]);
        if (cm) rec.comments[f.key] = cm;
        if (txt) {
          // Til søket legger vi også til råtallet (f.eks. beløp «13092») og initialer med mellomrom.
          let hay = txt;
          if (typeof v === 'number') hay += ' ' + v;
          if (f.type === 'initials') hay += ' ' + v;
          if (cm) hay += ' ' + cm.text;
          rec.search[f.key] = { folded: fold(hay), compact: foldCompact(hay) };
        } else if (cm) {
          rec.search[f.key] = { folded: fold(cm.text), compact: foldCompact(cm.text) };
        }
      }
      rec.vacant = isEmpty(rec.values.navn) && isEmpty(rec.values.kunde);
      rec.due = { ut: dueStatus(rec.values.ut), reg: dueStatus(rec.values.reg) };
      rec.gender = genderOf(rec, genderCol >= 0 ? row[genderCol] : null);
      records.push(rec);
    }
    return { records, title, hasGenderColumn: genderCol >= 0 };
  }

  /* ---------- Innlasting ---------- */

  async function loadBuffer(buffer, sourceName) {
    const sheet = await XlsxReader.read(buffer);
    const { records, title, hasGenderColumn } = buildRecords(sheet);
    if (!records.length) throw new Error('Fant ingen rader med data i arket.');
    state.records = records;
    state.sheetName = sheet.sheetName;
    state.hasGenderColumn = hasGenderColumn;
    state.openId = null;
    state.filters.clear();

    const tenants = records.filter((r) => !r.vacant).length;
    const vacant = records.length - tenants;
    els.subtitle.textContent = title || 'Liste over leietakere';
    els.source.textContent = sourceName + ' · ' + tenants + ' ' + plural(tenants, 'leietaker', 'leietakere') +
      (vacant ? ' · ' + vacant + ' ' + plural(vacant, 'ledig bolig', 'ledige boliger') : '');
    els.vacantCount.textContent = vacant ? '(' + vacant + ')' : '';
    els.vacant.closest('label').hidden = !vacant;

    els.loader.hidden = true;
    document.body.classList.add('has-data');
    onSearch();
  }

  async function loadDefault() {
    try {
      const res = await fetch('Data.xlsx', { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      await loadBuffer(await res.arrayBuffer(), 'Data.xlsx');
    } catch (err) {
      showLoader('Velg Excel-fil',
        'Kunne ikke laste Data.xlsx automatisk (f.eks. når siden er åpnet direkte fra disken). ' +
        'Velg filen eller dra den hit – dataene leses bare i denne nettleseren.');
    }
  }

  async function loadFile(file) {
    if (!file) return;
    try {
      await loadBuffer(await file.arrayBuffer(), file.name);
      toast('Lastet inn: ' + file.name);
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

  /* ---------- Hjelpere for DOM ---------- */

  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  // Tekst med uthevede søketreff.
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
  };

  function iconSvg(name, cls) {
    const span = el('span', 'icon' + (cls ? ' ' + cls : ''));
    span.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICONS[name] + '</svg>';
    return span;
  }

  const GENDER_INFO = {
    K:   { icon: 'female',  cls: 'g-k',   text: 'Kvinne' },
    M:   { icon: 'male',    cls: 'g-m',   text: 'Mann' },
    org: { icon: 'org',     cls: 'g-org', text: 'Tjeneste / organisasjon' },
  };

  function genderTitle(rec) {
    const g = rec.gender;
    if (!g.value) return 'Kjønn ukjent – legg til kolonnen «Kjønn» (M/K) i Excel';
    const info = GENDER_INFO[g.value];
    if (g.value === 'org') return info.text;
    return info.text + (g.source === 'guess' ? ' (antatt ut fra fornavn)' : ' (fra Excel)');
  }

  function genderBadge(rec) {
    const info = GENDER_INFO[rec.gender.value] || { icon: 'unknown', cls: 'g-unknown' };
    const b = iconSvg(info.icon, 'gender ' + info.cls);
    b.title = genderTitle(rec);
    b.setAttribute('role', 'img');
    b.setAttribute('aria-label', genderTitle(rec));
    return b;
  }

  /* ---------- Liste og filtre ---------- */

  function baseRecords() {
    let list = state.records.filter((r) => state.showVacant || !r.vacant);
    if (state.tokens.length) list = list.filter((r) => recordMatches(r, state.tokens));
    return list;
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
      // Først kommende datoer (snarest først), så passerte (nyeste først), til slutt uten dato.
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
      const st = row.key === 'inn' ? null : rec.due[row.key];
      if (st && st !== 'ok') {
        line.title = dueSentence(row.key, st, v);
        line.append(duePill(row.key, st, v));
      } else if (rec.badDate[row.key]) {
        line.append(el('span', 'due due-bad', 'ugyldig'));
        line.title = row.title + ': ugyldig dato i Excel («' + txt + '»)';
      }
      box.append(line);
    }
    return box;
  }

  function buildFace(rec) {
    const face = el('button', 'face');
    face.type = 'button';
    face.setAttribute('aria-expanded', 'false');
    face.setAttribute('aria-controls', 'detail-' + rec.id);

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
    else top.append(genderBadge(rec));
    face.append(top);

    // Treffer søket et felt som ikke vises på kortet, vises feltet her.
    if (state.tokens.length) {
      const shown = shownFieldKeys().concat(DATE_ROWS.map((r) => r.key));
      const hidden = state.tokens.filter((t) => !shown.some((k) => fieldMatches(rec, k, t)));
      const used = new Set();
      for (const t of hidden) {
        if (used.size >= 2) break;
        const f = FIELDS.find((f) => !shown.includes(f.key) && !used.has(f.key) && fieldMatches(rec, f.key, t));
        if (!f) continue;
        used.add(f.key);
        const snip = el('span', 'snippet');
        snip.append(el('span', 'snippet-label', f.label + ': '));
        const txt = rec.text[f.key] && matchRanges(rec.text[f.key], state.tokens).length
          ? rec.text[f.key]
          : (rec.comments[f.key] ? rec.comments[f.key].text : rec.text[f.key]);
        snip.append(hl(txt || '', null));
        face.append(snip);
      }
    }

    face.append(buildDates(rec));
    face.addEventListener('click', () => toggle(rec.id));
    return face;
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
    // Telefonnummer (8–12 sifre) blir en «tel:»-lenke.
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
    d.id = 'detail-' + rec.id;
    d.setAttribute('role', 'region');

    // Toppfelt
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
    const close = el('button', 'close');
    close.type = 'button';
    close.setAttribute('aria-label', 'Lukk kortet');
    close.title = 'Lukk (Esc)';
    close.append(iconSvg('close'));
    close.addEventListener('click', () => toggle(rec.id));
    head.append(mono, titles, close);
    d.append(head);

    const alerts = buildAlerts(rec);
    if (alerts) d.append(alerts);

    // Viktige numre
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

    // Seksjoner
    const grid = el('div', 'sections');
    for (const s of SECTIONS) {
      const sec = el('section', 'section');
      sec.append(el('h3', null, s.title));
      const dl = el('dl');
      for (const k of s.fields) {
        const item = el('div', 'field');
        const dt = el('dt', null, FIELD[k].label);
        const dd = el('dd');
        const raw = rec.values[k];
        const txt = rec.text[k];
        if (!txt) {
          dd.append(el('span', 'is-missing', '—'));
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
          if (rec.badDate[k]) dd.append(el('span', 'due due-bad', 'ugyldig dato i Excel'));
        }
        const cm = rec.comments[k];
        if (cm) {
          const c = el('p', 'cell-comment');
          c.append(iconSvg('note'));
          const body = el('span');
          body.append(hl(cm.text, 'cell-comment-text'));
          if (cm.author) body.append(el('span', 'cell-comment-author', ' – ' + cm.author));
          c.append(body);
          dd.append(c);
        }
        item.append(dt, dd);
        dl.append(item);
      }
      sec.append(dl);
      grid.append(sec);
    }
    d.append(grid);

    d.append(el('p', 'detail-foot', 'Rad ' + rec.excelRow + ' i arket «' + state.sheetName + '»'));
    return d;
  }

  function buildTile(rec) {
    const li = el('li', 'tile');
    li.dataset.id = rec.id;
    li.style.viewTransitionName = 'tile-' + rec.id;
    if (rec.vacant) li.classList.add('is-vacant');
    if (rec.due.ut && rec.due.ut !== 'ok') li.classList.add('ut-' + rec.due.ut);
    if (state.openId === rec.id) {
      li.classList.add('is-open');
      li.append(buildDetail(rec));
    } else {
      li.append(buildFace(rec));
    }
    return li;
  }

  /* ---------- Oversikt over frister ---------- */

  function renderOverview(base) {
    els.overview.replaceChildren();
    for (const g of DUE_GROUPS) {
      const group = el('div', 'ov-group');
      group.setAttribute('role', 'group');
      group.setAttribute('aria-label', g.title);
      group.append(el('span', 'ov-title', g.title));
      for (const st of ['m1', 'm2', 'm3', 'past']) {
        const id = g.kind + ':' + st;
        const n = base.filter((r) => r.due[g.kind] === st).length;
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

  /* ---------- Tegning ---------- */

  function render() {
    const base = baseRecords();
    const list = sortList(base.filter(passesFilters));
    const frag = document.createDocumentFragment();
    for (const rec of list) frag.append(buildTile(rec));
    els.grid.replaceChildren(frag);

    els.grid.dataset.show = state.show.join(' ');
    els.grid.classList.toggle('only-initials', state.show.length === 1 && state.show[0] === 'init');

    renderOverview(base);

    const total = state.records.filter((r) => state.showVacant || !r.vacant).length;
    els.count.textContent = state.tokens.length || state.filters.size
      ? 'Viser ' + list.length + ' av ' + total
      : total + ' kort';

    const noResults = state.records.length > 0 && !list.length;
    els.empty.hidden = !noResults;
    if (noResults) {
      els.emptyText.textContent = state.query.trim()
        ? 'Ingen treff for «' + state.query.trim() + '».'
        : 'Ingen kort passer til filteret.';
    }
  }

  // Endringer i oppsettet med jevn animasjon (View Transitions API) der nettleseren støtter det.
  // `morphIds` — kort som endrer utseende (åpnes/lukkes) får en rask overtoning;
  // de andre glir bare over til sin nye plass.
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

  function toggle(id) {
    const opening = state.openId !== id;
    const prev = state.openId;
    const done = withTransition(() => {
      state.openId = opening ? id : null;
      for (const pid of [prev, id]) {
        if (pid == null) continue;
        const li = els.grid.querySelector('.tile[data-id="' + pid + '"]');
        if (li) li.replaceWith(buildTile(state.records[pid]));
      }
    }, [prev, id]);
    done.then(() => {
      const li = els.grid.querySelector('.tile[data-id="' + id + '"]');
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
        if (!next.length) { cb.checked = true; return; } // minst ett felt
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

  /* ---------- Melding ---------- */

  let toastTimer = null;
  function toast(msg) {
    els.toast.textContent = msg;
    els.toast.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove('is-visible'), 2200);
  }

  /* ---------- Hendelser ---------- */

  function onSearch() {
    state.query = els.search.value;
    const tokens = tokenize(state.query);
    // Først hele frasen (så «Kleivan 22» ikke treffer datoer med «2022»),
    // og hvis den ikke finnes noe sted — hvert ord for seg, i vilkårlig rekkefølge.
    const phrase = tokens.join(' ');
    const phraseHit = tokens.length > 1 && state.records.some((r) =>
      (state.showVacant || !r.vacant) && FIELDS.some((f) => fieldMatches(r, f.key, phrase)));
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
  document.addEventListener('click', (e) => {
    if (!els.showMenu.hidden && !e.target.closest('.show-picker')) setMenu(false);
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
    state.showVacant = els.vacant.checked;
    savePref('showVacant', state.showVacant);
    onSearch();
  });

  els.fileInput.addEventListener('change', () => {
    loadFile(els.fileInput.files[0]);
    els.fileInput.value = '';
  });
  els.pickBtn.addEventListener('click', () => els.fileInput.click());

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (!els.showMenu.hidden) { setMenu(false); els.showBtn.focus(); return; }
      if (state.openId != null) { toggle(state.openId); return; }
      if (document.activeElement === els.search && els.search.value) {
        els.search.value = ''; onSearch();
      }
      return;
    }
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);
    if (e.key === '/' && !typing && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      els.search.focus();
      els.search.select();
    }
  });

  // Dra og slipp en Excel-fil hvor som helst på siden.
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

  // Høyden på verktøylinjen — så et utvidet kort ikke havner under den.
  if ('ResizeObserver' in window) {
    new ResizeObserver(() => {
      document.documentElement.style.setProperty('--toolbar-h', els.toolbar.offsetHeight + 'px');
    }).observe(els.toolbar);
  }

  // Kortere tekst i søkefeltet på smale skjermer.
  const narrow = window.matchMedia('(max-width: 640px)');
  const setPlaceholder = () => {
    els.search.placeholder = narrow.matches
      ? 'Søk: navn, adresse, nummer…'
      : 'Søk: navn, initialer, adresse, nummer, telefon…';
  };
  setPlaceholder();
  if (narrow.addEventListener) narrow.addEventListener('change', setPlaceholder);

  buildShowMenu();
  loadDefault();
})();

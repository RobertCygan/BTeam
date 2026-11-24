/* Klienci — kafelki z danymi z pliku Excel (Data.xlsx). */
(function () {
  'use strict';

  /* ---------- Konfiguracja kolumn ----------
   * Kolumny są rozpoznawane po nagłówku; jeśli nagłówek się zmieni,
   * używana jest litera kolumny jako zapas.
   */
  const FIELDS = [
    { key: 'art',     col: 'A', headers: ['artnr'],                          label: 'ARTnr' },
    { key: 'abb',     col: 'B', headers: ['abb.nr.', 'abbnr'],               label: 'Abb.nr.' },
    { key: 'kunde',   col: 'C', headers: ['kundenr.', 'kundenr'],            label: 'Kundenr.' },
    { key: 'status',  col: 'D', headers: ['status'],                         label: 'Status', type: 'status' },
    { key: 'belop1',  col: 'E', headers: ['beløp1'],                         label: 'Kwota 1', type: 'money' },
    { key: 'belop2',  col: 'F', headers: ['beløp2'],                         label: 'Kwota 2', type: 'money' },
    { key: 'navn',    col: 'G', headers: ['navn'],                           label: 'Imię i nazwisko' },
    { key: 'init',    col: 'H', headers: ['kolumna1', 'initialer'],          label: 'Inicjały', type: 'initials' },
    { key: 'bosted',  col: 'I', headers: ['bosted'],                         label: 'Miejscowość' },
    { key: 'bolig',   col: 'J', headers: ['bolig/leilighet'],                label: 'Lokal / mieszkanie' },
    { key: 'adresse', col: 'K', headers: ['adresse'],                        label: 'Adres' },
    { key: 'inn',     col: 'L', headers: ['datoinnflytting'],                label: 'Data wprowadzenia', type: 'date' },
    { key: 'ut',      col: 'M', headers: ['datoutflytting'],                 label: 'Data wyprowadzenia', type: 'date' },
    { key: 'reg',     col: 'N', headers: ['nesteregulering'],                label: 'Następna regulacja', type: 'date' },
    { key: 'status2', col: 'O', headers: ['status2'],                        label: 'Status 2' },
    { key: 'komm',    col: 'P', headers: ['kommentar/telefomnr.', 'kommentar/telefonnr.'], label: 'Komentarz / telefon', type: 'note' },
  ];
  const FIELD = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

  // Co można pokazać na kafelku (kolejność na kafelku jest zawsze taka jak tutaj).
  const SHOW_OPTIONS = [
    { id: 'init',    label: 'Inicjały',            short: 'Inicjały' },
    { id: 'navn',    label: 'Imię i nazwisko',     short: 'Nazwisko' },
    { id: 'adresse', label: 'Adres',               short: 'Adres' },
    { id: 'nr',      label: 'Kundenr. i Abb.nr.',  short: 'Numery' },
  ];
  const DEFAULT_SHOW = ['adresse'];

  // Sekcje w rozwiniętym kafelku.
  const SECTIONS = [
    { title: 'Lokal',             fields: ['adresse', 'bosted', 'bolig'] },
    { title: 'Umowa i terminy',   fields: ['inn', 'ut', 'reg', 'status', 'status2'] },
    { title: 'Płatności',         fields: ['belop1', 'belop2'] },
    { title: 'Kontakt i uwagi',   fields: ['komm'] },
  ];

  /* ---------- Stan ---------- */

  const state = {
    records: [],
    headers: {},
    show: loadPref('show', DEFAULT_SHOW).filter((id) => SHOW_OPTIONS.some((o) => o.id === id)),
    sort: loadPref('sort', 'file'),
    showVacant: loadPref('showVacant', false),
    query: '',
    tokens: [],
    openId: null,
    sheetName: '',
  };
  if (!state.show.length) state.show = DEFAULT_SHOW.slice();

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

  /* ---------- Preferencje ---------- */

  function loadPref(key, fallback) {
    try {
      const raw = localStorage.getItem('klienci.' + key);
      return raw == null ? fallback : JSON.parse(raw);
    } catch (e) { return fallback; }
  }
  function savePref(key, value) {
    try { localStorage.setItem('klienci.' + key, JSON.stringify(value)); } catch (e) { /* bez zapisu */ }
  }

  /* ---------- Formatowanie ---------- */

  const moneyFmt = new Intl.NumberFormat('pl-PL', { style: 'currency', currency: 'NOK', minimumFractionDigits: 2 });
  const relFmt = new Intl.RelativeTimeFormat('pl', { numeric: 'auto' });
  const collator = new Intl.Collator('nb', { numeric: true, sensitivity: 'base' });

  const isEmpty = (v) => v == null || (typeof v === 'string' && v.trim() === '');

  function pad(n) { return String(n).padStart(2, '0'); }

  function formatDate(d) {
    return pad(d.getUTCDate()) + '.' + pad(d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
  }

  function relativeDate(d) {
    const now = new Date();
    const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    const days = Math.round((d.getTime() - today) / 86400000);
    const abs = Math.abs(days);
    if (abs < 1) return 'dzisiaj';
    if (abs < 45) return relFmt.format(days, 'day');
    const months = Math.round(days / 30.44);
    if (Math.abs(months) < 18) return relFmt.format(months, 'month');
    return relFmt.format(Math.round(days / 365.25), 'year');
  }

  // Tekst wartości do wyświetlenia (null = brak wartości).
  function display(key, v) {
    if (isEmpty(v)) return null;
    const f = FIELD[key];
    if (v instanceof Date) return isNaN(v) ? null : formatDate(v);
    if (f.type === 'money' && typeof v === 'number') return moneyFmt.format(v);
    if (f.type === 'status' && typeof v === 'boolean') return v ? 'Aktywny' : 'Nieaktywny';
    if (typeof v === 'boolean') return v ? 'Tak' : 'Nie';
    if (f.type === 'initials') return String(v).replace(/[\s.]+/g, '').toUpperCase();
    return String(v).trim();
  }

  /* ---------- Wyszukiwanie ----------
   * Porównanie bez wielkości liter i znaków diakrytycznych (ø→o, æ→ae, å→a, ł→l …),
   * a dla numerów także bez spacji (np. „45254373” znajdzie „452 54 373”).
   */
  const SPECIAL = { 'æ': 'ae', 'ø': 'o', 'å': 'a', 'ß': 'ss', 'ł': 'l', 'đ': 'd', 'œ': 'oe' };

  function foldChar(ch) {
    const lower = ch.toLowerCase();
    if (SPECIAL[lower]) return SPECIAL[lower];
    return lower.normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  // Zwraca tekst „złożony” i mapę pozycji do tekstu oryginalnego.
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

  function tokenize(query) {
    return fold(query).split(/\s+/).filter(Boolean);
  }

  // Przedziały [start, koniec) w tekście oryginalnym pasujące do któregokolwiek słowa.
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

  /* ---------- Dane ---------- */

  function headerKey(v) {
    return typeof v === 'string' ? v.toLowerCase().replace(/\s+/g, '') : '';
  }

  function buildRecords(sheet) {
    const rows = sheet.rows;
    let h = rows.findIndex((r, i) => i < 40 &&
      r.some((v) => headerKey(v) === 'navn') && r.some((v) => headerKey(v) === 'adresse'));
    if (h < 0) h = 2; // domyślnie wiersz 3
    const headerRow = rows[h] || [];

    const colOf = {};
    const headers = {};
    for (const f of FIELDS) {
      let idx = headerRow.findIndex((v) => f.headers.includes(headerKey(v)));
      if (idx < 0) idx = XlsxReader.colIndex(f.col);
      colOf[f.key] = idx;
      headers[f.key] = typeof headerRow[idx] === 'string' && headerRow[idx].trim() ? headerRow[idx].trim() : 'kolumna ' + f.col;
    }

    let title = '';
    for (let i = 0; i < h && !title; i++) {
      const cell = (rows[i] || []).find((v) => typeof v === 'string' && v.trim());
      if (cell) title = cell.trim();
    }

    const records = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || FIELDS.every((f) => isEmpty(row[colOf[f.key]]))) continue;
      const rec = { id: records.length, excelRow: r + 1, values: {}, text: {}, comments: {}, search: {} };
      for (const f of FIELDS) {
        let v = row[colOf[f.key]];
        if (typeof v === 'string') v = v.trim() || null;
        rec.values[f.key] = v ?? null;
        const txt = display(f.key, v);
        rec.text[f.key] = txt;
        const cm = sheet.comments.get(r + ',' + colOf[f.key]);
        if (cm) rec.comments[f.key] = cm;
        if (txt) {
          // Do wyszukiwania dodajemy też surową liczbę (np. kwota „13092”) i inicjały ze spacjami.
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
      records.push(rec);
    }
    return { records, headers, title };
  }

  /* ---------- Ładowanie pliku ---------- */

  async function loadBuffer(buffer, sourceName) {
    const sheet = await XlsxReader.read(buffer);
    const { records, headers, title } = buildRecords(sheet);
    if (!records.length) throw new Error('W arkuszu nie znaleziono żadnych wierszy z danymi.');
    state.records = records;
    state.headers = headers;
    state.sheetName = sheet.sheetName;
    state.openId = null;

    const clients = records.filter((r) => !r.vacant).length;
    const vacant = records.length - clients;
    els.subtitle.textContent = title || 'Lista klientów';
    els.source.textContent = sourceName + ' · ' + clients + ' ' + plural(clients, 'klient', 'klientów', 'klientów') +
      (vacant ? ' · ' + vacant + ' ' + plural(vacant, 'wolny lokal', 'wolne lokale', 'wolnych lokali') : '');
    els.vacantCount.textContent = vacant ? '(' + vacant + ')' : '';
    els.vacant.closest('label').hidden = !vacant;

    els.loader.hidden = true;
    document.body.classList.add('has-data');
    render();
  }

  function plural(n, one, few, many) {
    if (n === 1) return one;
    const d = n % 10, dd = n % 100;
    return d >= 2 && d <= 4 && !(dd >= 12 && dd <= 14) ? few : many;
  }

  async function loadDefault() {
    try {
      const res = await fetch('Data.xlsx', { cache: 'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      await loadBuffer(await res.arrayBuffer(), 'Data.xlsx');
    } catch (err) {
      showLoader('Wybierz plik Excel',
        'Nie udało się automatycznie wczytać pliku Data.xlsx (np. gdy strona jest otwarta bezpośrednio z dysku). ' +
        'Wybierz plik lub przeciągnij go tutaj — dane zostaną odczytane tylko w tej przeglądarce.');
    }
  }

  async function loadFile(file) {
    if (!file) return;
    try {
      await loadBuffer(await file.arrayBuffer(), file.name);
      toast('Wczytano: ' + file.name);
    } catch (err) {
      if (!state.records.length) showLoader('Nie udało się odczytać pliku', err.message || String(err));
      else toast('Nie udało się odczytać pliku: ' + (err.message || err));
    }
  }

  function showLoader(title, text) {
    els.loaderTitle.textContent = title;
    els.loaderText.textContent = text;
    els.pickBtn.hidden = false;
    els.loader.hidden = false;
    els.loader.classList.remove('is-loading');
  }

  /* ---------- Renderowanie ---------- */

  function el(tag, cls, text) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  // Tekst z podświetleniem dopasowań wyszukiwania.
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

  function visibleRecords() {
    let list = state.records.filter((r) => state.showVacant || !r.vacant);
    if (state.tokens.length) list = list.filter((r) => recordMatches(r, state.tokens));
    if (state.sort === 'alpha') {
      const key = sortKey();
      list = list.slice().sort((a, b) => {
        const x = a.text[key], y = b.text[key];
        if (!x) return y ? 1 : 0;
        if (!y) return -1;
        return collator.compare(x, y);
      });
    }
    return list;
  }

  function sortKey() {
    const primary = SHOW_OPTIONS.find((o) => state.show.includes(o.id)).id;
    return primary === 'nr' ? 'kunde' : primary;
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

  function buildFace(rec) {
    const face = el('button', 'face');
    face.type = 'button';
    face.setAttribute('aria-expanded', 'false');
    face.setAttribute('aria-controls', 'detail-' + rec.id);

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
        face.append(box);
        continue;
      }
      let txt = rec.text[o.id];
      if (!txt && rec.vacant && (o.id === 'navn' || o.id === 'init')) {
        face.append(el('span', 'f-' + o.id + role + ' is-missing', o.id === 'init' ? '—' : 'Wolny lokal'));
        continue;
      }
      face.append(txt ? hl(txt, 'f-' + o.id + role) : el('span', 'f-' + o.id + role + ' is-missing', '—'));
    }

    // Jeśli szukana fraza pasuje do pola, którego nie widać na kafelku — pokaż je.
    if (state.tokens.length) {
      const shown = shownFieldKeys();
      const hidden = state.tokens.filter((t) => !shown.some((k) => fieldMatches(rec, k, t)));
      const used = new Set();
      for (const t of hidden) {
        const f = FIELDS.find((f) => !shown.includes(f.key) && !used.has(f.key) && fieldMatches(rec, f.key, t));
        if (!f || used.size >= 2) continue;
        used.add(f.key);
        const snip = el('span', 'snippet');
        snip.append(el('span', 'snippet-label', f.label + ': '));
        const txt = rec.text[f.key] && (rec.search[f.key] && matchRanges(rec.text[f.key], state.tokens).length)
          ? rec.text[f.key]
          : (rec.comments[f.key] ? rec.comments[f.key].text : rec.text[f.key]);
        snip.append(hl(txt || '', null));
        face.append(snip);
      }
    }

    if (rec.vacant) face.append(el('span', 'badge', 'Wolny'));

    face.addEventListener('click', () => toggle(rec.id));
    return face;
  }

  function iconSvg(name) {
    const paths = {
      copy: '<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V6a2 2 0 0 1 2-2h9"/>',
      check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
      close: '<path d="M6 6l12 12M18 6L6 18"/>',
      phone: '<path d="M6.6 3h3l1.6 4.2-2 1.3a11 11 0 0 0 6.3 6.3l1.3-2 4.2 1.6v3a2 2 0 0 1-2.2 2A17 17 0 0 1 4.6 5.2 2 2 0 0 1 6.6 3z"/>',
      note: '<path d="M5 4h14v12l-4 4H5z"/><path d="M15 20v-4h4"/>',
    };
    const span = el('span', 'icon');
    span.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true">' + paths[name] + '</svg>';
    return span;
  }

  function copyButton(value, label) {
    const btn = el('button', 'copy');
    btn.type = 'button';
    btn.title = 'Kopiuj ' + label;
    btn.setAttribute('aria-label', 'Kopiuj ' + label);
    btn.append(iconSvg('copy'));
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      try {
        await navigator.clipboard.writeText(value);
        btn.replaceChildren(iconSvg('check'));
        btn.classList.add('is-done');
        toast('Skopiowano ' + label + ': ' + value);
        setTimeout(() => { btn.replaceChildren(iconSvg('copy')); btn.classList.remove('is-done'); }, 1400);
      } catch (err) {
        toast('Nie udało się skopiować do schowka');
      }
    });
    return btn;
  }

  function noteValue(text) {
    // Numer telefonu (8–12 cyfr) zamieniamy w link „tel:”.
    const m = /^\+?[\d\s]{8,}$/.test(text) ? text.replace(/\s+/g, '') : null;
    if (m && m.replace('+', '').length >= 8 && m.replace('+', '').length <= 12) {
      const a = el('a', 'tel');
      a.href = 'tel:' + m;
      a.append(iconSvg('phone'), hl(text, null));
      return a;
    }
    return hl(text, null);
  }

  function buildDetail(rec) {
    const d = el('div', 'detail');
    d.id = 'detail-' + rec.id;
    d.setAttribute('role', 'region');

    // Nagłówek
    const head = el('div', 'detail-head');
    const mono = el('div', 'monogram', rec.text.init || (rec.vacant ? '—' : '?'));
    mono.setAttribute('aria-hidden', 'true');
    const titles = el('div', 'detail-titles');
    const h2 = el('h2', 'detail-name');
    h2.tabIndex = -1;
    h2.append(rec.text.navn ? hl(rec.text.navn, null) : (rec.vacant ? 'Wolny lokal' : 'Bez nazwy'));
    titles.append(h2);
    const sub = el('p', 'detail-sub');
    const subParts = [rec.text.adresse, rec.text.bosted].filter(Boolean);
    subParts.forEach((p, i) => {
      if (i) sub.append(el('span', 'dot', ' · '));
      sub.append(hl(p, null));
    });
    if (subParts.length) titles.append(sub);
    const close = el('button', 'close');
    close.type = 'button';
    close.setAttribute('aria-label', 'Zwiń kafelek');
    close.title = 'Zwiń (Esc)';
    close.append(iconSvg('close'));
    close.addEventListener('click', () => toggle(rec.id));
    head.append(mono, titles, close);
    d.append(head);

    // Kluczowe numery
    const keys = el('div', 'keynums');
    for (const k of ['kunde', 'abb', 'art']) {
      const box = el('div', 'keynum');
      const lab = el('span', 'keynum-label', FIELD[k].label);
      box.append(lab);
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

    // Sekcje
    const grid = el('div', 'sections');
    for (const s of SECTIONS) {
      const sec = el('section', 'section');
      sec.append(el('h3', null, s.title));
      const dl = el('dl');
      for (const k of s.fields) {
        const item = el('div', 'field');
        const dt = el('dt');
        dt.append(el('span', null, FIELD[k].label));
        if (state.headers[k] && fold(state.headers[k]) !== fold(FIELD[k].label)) {
          dt.append(el('span', 'orig', state.headers[k]));
        }
        const dd = el('dd');
        const raw = rec.values[k];
        const txt = rec.text[k];
        if (!txt) {
          dd.append(el('span', 'is-missing', isEmpty(raw) ? '—' : String(raw)));
        } else if (FIELD[k].type === 'note') {
          dd.append(noteValue(txt));
        } else if (FIELD[k].type === 'status') {
          const pill = el('span', 'pill ' + (raw === true ? 'pill-on' : 'pill-off'));
          pill.append(hl(txt, null));
          dd.append(pill);
        } else {
          dd.append(hl(txt, null));
          if (raw instanceof Date && !isNaN(raw)) dd.append(el('span', 'rel', relativeDate(raw)));
        }
        const cm = rec.comments[k];
        if (cm) {
          const c = el('p', 'cell-comment');
          c.append(iconSvg('note'));
          const body = el('span');
          body.append(hl(cm.text, 'cell-comment-text'));
          if (cm.author) body.append(el('span', 'cell-comment-author', ' — ' + cm.author));
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

    d.append(el('p', 'detail-foot', 'Wiersz ' + rec.excelRow + ' w arkuszu „' + state.sheetName + '”'));
    return d;
  }

  function buildTile(rec) {
    const li = el('li', 'tile');
    li.dataset.id = rec.id;
    li.style.viewTransitionName = 'tile-' + rec.id;
    if (rec.vacant) li.classList.add('is-vacant');
    if (state.openId === rec.id) {
      li.classList.add('is-open');
      li.append(buildDetail(rec));
    } else {
      li.append(buildFace(rec));
    }
    return li;
  }

  function render() {
    const list = visibleRecords();
    const frag = document.createDocumentFragment();
    for (const rec of list) frag.append(buildTile(rec));
    els.grid.replaceChildren(frag);

    els.grid.dataset.show = state.show.join(' ');
    els.grid.classList.toggle('only-initials', state.show.length === 1 && state.show[0] === 'init');

    const total = state.records.filter((r) => state.showVacant || !r.vacant).length;
    els.count.textContent = state.tokens.length
      ? 'Znaleziono ' + list.length + ' z ' + total
      : total + ' ' + plural(total, 'kafelek', 'kafelki', 'kafelków');

    const noResults = state.records.length > 0 && !list.length;
    els.empty.hidden = !noResults;
    if (noResults) els.emptyText.textContent = 'Brak wyników dla „' + state.query.trim() + '”.';
  }

  // Zmiany układu z płynną animacją (View Transitions API), jeśli przeglądarka ją obsługuje.
  // `morphIds` — kafelki, które zmieniają wygląd (rozwijają się / zwijają): dostają szybkie
  // przenikanie treści; pozostałe tylko płynnie przesuwają się na nowe miejsce.
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

  /* ---------- Wybór pól na kafelkach ---------- */

  function buildShowMenu() {
    els.showMenu.replaceChildren();
    for (const o of SHOW_OPTIONS) {
      const label = el('label', 'menu-item');
      const cb = el('input');
      cb.type = 'checkbox';
      cb.value = o.id;
      cb.checked = state.show.includes(o.id);
      cb.addEventListener('change', () => {
        let next = SHOW_OPTIONS.map((x) => x.id).filter((id) =>
          id === o.id ? cb.checked : state.show.includes(id));
        if (!next.length) { cb.checked = true; return; } // przynajmniej jedno pole
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
    const hint = el('p', 'menu-hint', 'Można zaznaczyć kilka pól naraz.');
    els.showMenu.append(hint);
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

  /* ---------- Toast ---------- */

  let toastTimer = null;
  function toast(msg) {
    els.toast.textContent = msg;
    els.toast.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove('is-visible'), 2200);
  }

  /* ---------- Zdarzenia ---------- */

  function onSearch() {
    state.query = els.search.value;
    const tokens = tokenize(state.query);
    // Najpierw cała fraza (np. „Kleivan 22” nie złapie dat z „2022”),
    // a gdy nigdzie jej nie ma — każde słowo osobno, w dowolnej kolejności.
    const phrase = tokens.join(' ');
    const phraseHit = tokens.length > 1 && state.records.some((r) =>
      (state.showVacant || !r.vacant) && FIELDS.some((f) => fieldMatches(r, f.key, phrase)));
    state.tokens = phraseHit ? [phrase] : tokens;
    render();
  }

  els.search.addEventListener('input', onSearch);
  els.clearSearch.addEventListener('click', () => {
    els.search.value = '';
    onSearch();
    els.search.focus();
  });

  els.showBtn.addEventListener('click', () => setMenu(els.showMenu.hidden));
  document.addEventListener('click', (e) => {
    if (!els.showMenu.hidden && !e.target.closest('.show-picker')) setMenu(false);
  });

  els.sort.value = state.sort === 'alpha' ? 'alpha' : 'file';
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

  // Przeciągnij i upuść plik Excel w dowolne miejsce strony.
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

  // Wysokość paska narzędzi — żeby rozwinięty kafelek nie chował się pod nim.
  if ('ResizeObserver' in window) {
    new ResizeObserver(() => {
      document.documentElement.style.setProperty('--toolbar-h', els.toolbar.offsetHeight + 'px');
    }).observe(els.toolbar);
  }

  // Krótszy podpis pola wyszukiwania na wąskich ekranach.
  const narrow = window.matchMedia('(max-width: 640px)');
  const setPlaceholder = () => {
    els.search.placeholder = narrow.matches
      ? 'Szukaj: nazwisko, adres, numer…'
      : 'Szukaj: nazwisko, inicjały, adres, numer, telefon…';
  };
  setPlaceholder();
  if (narrow.addEventListener) narrow.addEventListener('change', setPlaceholder);

  buildShowMenu();
  loadDefault();
})();

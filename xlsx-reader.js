/*
 * Minimalny czytnik plików .xlsx — bez zewnętrznych bibliotek.
 *
 * Czyta pierwszy arkusz skoroszytu: wartości komórek (tekst, liczby, daty,
 * wartości logiczne) oraz komentarze do komórek. Korzysta wyłącznie z API
 * przeglądarki: DataView (ZIP), DecompressionStream (deflate) i DOMParser (XML).
 *
 * Wynik:
 *   {
 *     sheetName: 'Arkusz1',
 *     rows: [[v, v, ...], ...],          // rows[0] = wiersz 1 w Excelu, rows[r][0] = kolumna A
 *     comments: Map('r,c' -> { author, text })
 *   }
 * Daty są zwracane jako obiekty Date (północ UTC).
 */
(function (global) {
  'use strict';

  const utf8 = new TextDecoder('utf-8');

  /* ---------- ZIP ---------- */

  function openZip(buffer) {
    const view = new DataView(buffer);
    const bytes = new Uint8Array(buffer);

    // Koniec katalogu centralnego (EOCD) — szukamy od końca pliku.
    let eocd = -1;
    const minPos = Math.max(0, buffer.byteLength - 22 - 0xffff);
    for (let i = buffer.byteLength - 22; i >= minPos; i--) {
      if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('To nie jest plik .xlsx (brak struktury ZIP).');

    const count = view.getUint16(eocd + 10, true);
    let ptr = view.getUint32(eocd + 16, true);
    const entries = new Map();

    for (let n = 0; n < count; n++) {
      if (view.getUint32(ptr, true) !== 0x02014b50) throw new Error('Uszkodzony plik .xlsx.');
      const method = view.getUint16(ptr + 10, true);
      const compSize = view.getUint32(ptr + 20, true);
      const nameLen = view.getUint16(ptr + 28, true);
      const extraLen = view.getUint16(ptr + 30, true);
      const commentLen = view.getUint16(ptr + 32, true);
      const localOffset = view.getUint32(ptr + 42, true);
      const name = utf8.decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
      entries.set(name.toLowerCase(), { method, compSize, localOffset });
      ptr += 46 + nameLen + extraLen + commentLen;
    }

    async function read(name) {
      const e = entries.get(name.replace(/^\/+/, '').toLowerCase());
      if (!e) return null;
      const start = e.localOffset + 30 +
        view.getUint16(e.localOffset + 26, true) +
        view.getUint16(e.localOffset + 28, true);
      const data = bytes.subarray(start, start + e.compSize);
      if (e.method === 0) return utf8.decode(data);
      if (e.method !== 8) throw new Error('Nieobsługiwana metoda kompresji w pliku .xlsx.');
      if (typeof DecompressionStream === 'undefined') {
        throw new Error('Ta przeglądarka jest zbyt stara, aby odczytać plik .xlsx.');
      }
      const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      return new Response(stream).text();
    }

    return { read };
  }

  /* ---------- XML ---------- */

  const parser = new DOMParser();

  function parseXml(text) {
    const doc = parser.parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Błędny XML w pliku .xlsx.');
    return doc;
  }

  // Elementy wyszukujemy po nazwie lokalnej — niezależnie od prefiksu przestrzeni nazw.
  const all = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
  const children = (node, name) => Array.from(node.children).filter((c) => c.localName === name);
  const first = (node, name) => all(node, name)[0] || null;

  // Tekst elementu <si>/<is>/<text>: łączymy wszystkie <t>, pomijając fonetykę (<rPh>).
  function richText(node) {
    return all(node, 't')
      .filter((t) => !t.closest || !t.closest('rPh'))
      .map((t) => t.textContent)
      .join('');
  }

  async function readRels(zip, partPath) {
    const slash = partPath.lastIndexOf('/');
    const relsPath = partPath.slice(0, slash + 1) + '_rels/' + partPath.slice(slash + 1) + '.rels';
    const text = await zip.read(relsPath);
    const rels = new Map();
    if (!text) return rels;
    for (const r of all(parseXml(text), 'Relationship')) {
      rels.set(r.getAttribute('Id'), {
        type: r.getAttribute('Type') || '',
        target: resolvePath(partPath, r.getAttribute('Target') || ''),
      });
    }
    return rels;
  }

  function resolvePath(fromPart, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = fromPart.split('/').slice(0, -1);
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg && seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }

  /* ---------- Formaty liczb / daty ---------- */

  const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

  function isDateFormat(id, code) {
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    if (!code) return false;
    const stripped = code
      .split(';')[0]
      .replace(/"[^"]*"/g, '')     // teksty w cudzysłowie
      .replace(/\[[^\]]*\]/g, '')  // [Red], [$NOK], [h] …
      .replace(/\\./g, '')         // znaki ucieczki
      .replace(/[_*]./g, '');      // wypełnienia
    return /[dmyhs]/i.test(stripped) && !/general/i.test(stripped);
  }

  function serialToDate(serial, date1904) {
    let days = serial;
    if (date1904) days += 1462;
    if (days < 60) days += 1; // błąd roku przestępnego 1900 w Excelu
    const ms = Math.round((days - 25569) * 86400) * 1000;
    const d = new Date(ms);
    return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }

  function colIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/i.exec(ref || '');
    return m ? { c: colIndex(m[1]), r: parseInt(m[2], 10) - 1 } : null;
  }

  /* ---------- Skoroszyt ---------- */

  async function read(buffer) {
    const zip = openZip(buffer);

    const rootRels = await readRels(zip, '');
    let workbookPath = 'xl/workbook.xml';
    for (const rel of rootRels.values()) {
      if (rel.type.endsWith('/officeDocument')) workbookPath = rel.target;
    }
    const workbookXml = await zip.read(workbookPath);
    if (!workbookXml) throw new Error('W pliku nie znaleziono skoroszytu Excel.');
    const workbook = parseXml(workbookXml);
    const wbRels = await readRels(zip, workbookPath);

    const wbPr = first(workbook, 'workbookPr');
    const date1904 = !!wbPr && /^(1|true)$/i.test(wbPr.getAttribute('date1904') || '');

    const sheetEl = first(workbook, 'sheet');
    if (!sheetEl) throw new Error('Skoroszyt nie zawiera arkuszy.');
    const sheetName = sheetEl.getAttribute('name') || 'Arkusz';
    const rid = sheetEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
      sheetEl.getAttribute('r:id');
    const sheetPath = (wbRels.get(rid) || {}).target || 'xl/worksheets/sheet1.xml';

    // Teksty współdzielone
    let shared = [];
    let stylesPath = null;
    for (const rel of wbRels.values()) {
      if (rel.type.endsWith('/sharedStrings')) {
        const xml = await zip.read(rel.target);
        if (xml) shared = all(parseXml(xml), 'si').map(richText);
      } else if (rel.type.endsWith('/styles')) {
        stylesPath = rel.target;
      }
    }

    // Style → które komórki są datami
    const dateStyle = [];
    const stylesXml = stylesPath ? await zip.read(stylesPath) : null;
    if (stylesXml) {
      const styles = parseXml(stylesXml);
      const custom = new Map(all(styles, 'numFmt').map((f) =>
        [parseInt(f.getAttribute('numFmtId'), 10), f.getAttribute('formatCode')]));
      const cellXfs = first(styles, 'cellXfs');
      if (cellXfs) {
        for (const xf of children(cellXfs, 'xf')) {
          const id = parseInt(xf.getAttribute('numFmtId') || '0', 10);
          dateStyle.push(isDateFormat(id, custom.get(id)));
        }
      }
    }

    // Arkusz
    const sheetXml = await zip.read(sheetPath);
    if (!sheetXml) throw new Error('Nie można odczytać arkusza „' + sheetName + '”.');
    const sheet = parseXml(sheetXml);
    const rows = [];
    let lastRow = -1;
    for (const rowEl of all(sheet, 'row')) {
      const r = rowEl.hasAttribute('r') ? parseInt(rowEl.getAttribute('r'), 10) - 1 : lastRow + 1;
      lastRow = r;
      const row = rows[r] || (rows[r] = []);
      let lastCol = -1;
      for (const c of children(rowEl, 'c')) {
        const pos = splitRef(c.getAttribute('r'));
        const col = pos ? pos.c : lastCol + 1;
        lastCol = col;
        const t = c.getAttribute('t') || 'n';
        const vEl = children(c, 'v')[0];
        const v = vEl ? vEl.textContent : null;
        let value = null;
        if (t === 's') value = v == null ? null : shared[parseInt(v, 10)] ?? null;
        else if (t === 'inlineStr') value = children(c, 'is')[0] ? richText(children(c, 'is')[0]) : null;
        else if (t === 'str' || t === 'e') value = v;
        else if (t === 'b') value = v == null ? null : v === '1' || v === 'true';
        else if (t === 'd') value = v ? new Date(v) : null;
        else if (v != null && v !== '') {
          const num = Number(v);
          const s = parseInt(c.getAttribute('s') || '0', 10);
          value = dateStyle[s] && isFinite(num) ? serialToDate(num, date1904) : num;
        }
        row[col] = value;
      }
    }
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];

    // Komentarze do komórek
    const comments = new Map();
    const sheetRels = await readRels(zip, sheetPath);
    for (const rel of sheetRels.values()) {
      if (!rel.type.endsWith('/comments')) continue;
      const xml = await zip.read(rel.target);
      if (!xml) continue;
      const doc = parseXml(xml);
      const authors = all(doc, 'author').map((a) => a.textContent.trim());
      for (const cm of all(doc, 'comment')) {
        const pos = splitRef(cm.getAttribute('ref'));
        if (!pos) continue;
        const textEl = children(cm, 'text')[0];
        let text = textEl ? richText(textEl).trim() : '';
        let author = authors[parseInt(cm.getAttribute('authorId') || '0', 10)] || '';
        // Excel zaczyna treść komentarza od „Autor:\n” — oddzielamy to.
        const nl = text.indexOf('\n');
        if (nl > 0 && text.slice(0, nl).trim().endsWith(':')) {
          author = text.slice(0, nl).trim().slice(0, -1);
          text = text.slice(nl + 1).trim();
        }
        if (text) comments.set(pos.r + ',' + pos.c, { author, text });
      }
    }

    return { sheetName, rows, comments };
  }

  global.XlsxReader = { read, colIndex };
})(window);

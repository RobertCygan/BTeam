#!/usr/bin/env python3
"""
Bygger index.html — én selvstendig HTML-fil med alt innebygd (stiler, skript, logo og data).

Dataene hentes fra Excel-filen og krypteres (AES-256-GCM, nøkkel fra PBKDF2-SHA256 av
brukernavn og passord). Passordet lagres ikke noe sted; det oppgis når filen bygges:

    BOLIG_PASSORD='...' python3 tools/build.py            # eller: python3 tools/build.py (spør etter passord)
    python3 tools/build.py --data sti/til/Data.xlsx --bruker BoligTeam

Krever: pip install openpyxl cryptography
Salt og antall iterasjoner gjenbrukes fra en eksisterende index.html, slik at endringer som er lagret
i nettleseren, fortsatt kan åpnes etter en ny bygging (bruk --ny-salt for å lage nytt).
"""
import argparse
import base64
import datetime
import getpass
import hashlib
import json
import os
import re
import sys
from pathlib import Path

import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'src'
SCRIPTS = ['xlsx-reader.js', 'xlsx-writer.js', 'navn-kjonn.js', 'login.js', 'app.js']
ITERATIONS = 600_000

KEYS = ['art', 'abb', 'kunde', 'status', 'belop1', 'belop2', 'navn', 'init', 'bosted', 'bolig',
        'adresse', 'inn', 'ut', 'reg', 'status2', 'komm']
DATES = {'inn', 'ut', 'reg'}


def parse_text_date(s):
    m = re.match(r'^\s*(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})\s*$', s)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return datetime.date(y, mo, d).isoformat()
    except ValueError:
        return None


def read_excel(path):
    """Leser arket i samme format som appen bruker (se fromPlain i app.js)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows())
    hdr = next(i for i, r in enumerate(rows[:40])
               if any(isinstance(c.value, str) and c.value.strip().lower() == 'navn' for c in r))
    headers = [c.value for c in rows[hdr][:16]]
    title = next((c.value.strip() for r in rows[:hdr] for c in r
                  if isinstance(c.value, str) and c.value.strip()), '')
    records = []
    for r in rows[hdr + 1:]:
        cells = r[:16]
        if all(c.value is None or (isinstance(c.value, str) and not c.value.strip()) for c in cells):
            continue
        values = {}
        for k, c in zip(KEYS, cells):
            v = c.value
            if isinstance(v, str):
                v = v.strip() or None
            if isinstance(v, datetime.datetime):
                v = v.date().isoformat()
            elif k in DATES and isinstance(v, str):
                v = parse_text_date(v) or v
            elif isinstance(v, float) and v.is_integer():
                v = int(v)
            if k == 'art' and isinstance(v, (int, float)):
                v = str(int(v)).zfill(5)
            values[k] = v
        notes = []
        for c in r:
            if c.comment:
                text = c.comment.text.strip()
                author = ''
                first, _, rest = text.partition('\n')
                if rest and first.strip().endswith(':'):
                    author, text = first.strip()[:-1], rest.strip()
                notes.append(text + (f' ({author})' if author else ''))
        records.append({'uid': len(records) + 1, 'values': values, 'kjonn': None,
                        'merknad': '\n'.join(notes), 'changed': None})
    return {'version': 1,
            'meta': {'title': title, 'sheetName': ws.title, 'source': Path(path).name, 'headers': headers},
            'records': records}


def logo(n):
    g = f'bt-g{n}'
    return (
        '<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Bolig Team Steigen">'
        f'<defs><linearGradient id="{g}" x1="0" y1="0" x2="1" y2="1">'
        '<stop offset="0" stop-color="#2a8bad"/><stop offset="1" stop-color="#0f4d63"/></linearGradient></defs>'
        f'<rect width="48" height="48" rx="13" fill="url(#{g})"/>'
        # fjellene i Steigen
        '<path d="M3 33 14 19l6 7 9-12 16 19z" fill="#7cc5de" opacity=".5"/>'
        '<path d="M29 14l2.5 3-1.9-.8-1.6 1.1-1.5.1z" fill="#fff" opacity=".85"/>'
        # huset
        '<path d="M13 39V27.5l11-9 11 9V39z" fill="#fff"/>'
        '<path d="M10.5 28.8 24 17.5l13.5 11.3" fill="none" stroke="#fff" stroke-width="2.4" '
        'stroke-linecap="round" stroke-linejoin="round"/>'
        # vinduet = fire kort i appens farger
        '<rect x="19.4" y="27.6" width="4.2" height="4.2" rx="1" fill="#e0242f"/>'
        '<rect x="24.4" y="27.6" width="4.2" height="4.2" rx="1" fill="#ff8a1d"/>'
        '<rect x="19.4" y="32.6" width="4.2" height="4.2" rx="1" fill="#ffd43b"/>'
        '<rect x="24.4" y="32.6" width="4.2" height="4.2" rx="1" fill="#1f6f8b"/>'
        # havet
        '<path d="M5 43.2q4-2.2 8 0t8 0 8 0 8 0 6 0" fill="none" stroke="#a9dcee" stroke-width="1.6" '
        'stroke-linecap="round"/>'
        '</svg>'
    )


def existing_kdf(path):
    if not path.exists():
        return None
    m = re.search(r'<script type="application/json" id="payload">(.*?)</script>', path.read_text('utf-8'), re.S)
    if not m:
        return None
    try:
        p = json.loads(m.group(1))
        return base64.b64decode(p['salt']), int(p['iter'])
    except (ValueError, KeyError):
        return None


def encrypt(data, user, password, salt, iterations):
    key = hashlib.pbkdf2_hmac('sha256', (user.strip().lower() + '\n' + password).encode('utf-8'),
                              salt, iterations, dklen=32)
    iv = os.urandom(12)
    ct = AESGCM(key).encrypt(iv, json.dumps(data, ensure_ascii=False).encode('utf-8'), None)
    b = lambda x: base64.b64encode(x).decode('ascii')
    return {'v': 1, 'kdf': 'PBKDF2-SHA256', 'iter': iterations, 'salt': b(salt), 'iv': b(iv), 'ct': b(ct)}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--data', default=str(ROOT / 'Data.xlsx'), help='Excel-filen med leietakere')
    ap.add_argument('--bruker', default='BoligTeam', help='brukernavn (store/små bokstaver spiller ingen rolle)')
    ap.add_argument('--ut', default=str(ROOT / 'index.html'), help='filen som skal lages')
    ap.add_argument('--ny-salt', action='store_true', help='lag nytt salt (lagrede endringer i nettleseren kan da ikke åpnes)')
    args = ap.parse_args()

    password = os.environ.get('BOLIG_PASSORD') or getpass.getpass('Passord: ')
    if not password:
        sys.exit('Passord mangler.')

    out = Path(args.ut)
    kdf = None if args.ny_salt else existing_kdf(out)
    salt, iterations = kdf if kdf else (os.urandom(16), ITERATIONS)

    data = read_excel(args.data)
    payload = encrypt(data, args.bruker, password, salt, iterations)

    html = (SRC / 'template.html').read_text('utf-8')
    css = (SRC / 'style.css').read_text('utf-8')
    scripts = []
    for name in SCRIPTS:
        js = (SRC / name).read_text('utf-8')
        if re.search(r'</script|<!--', js, re.I):
            sys.exit(f'{name} inneholder «</script» eller «<!--», som ikke kan bygges inn.')
        scripts.append(f'<script>\n/* ===== {name} ===== */\n{js}</script>')

    n = 0

    def next_logo(_):
        nonlocal n
        n += 1
        return logo(n)

    html = re.sub(r'<!-- @LOGO -->', next_logo, html)
    html = html.replace('<!-- @STYLE -->', '<style>\n' + css + '</style>')
    html = html.replace('<!-- @PAYLOAD -->',
                        '<script type="application/json" id="payload">' + json.dumps(payload) + '</script>')
    html = html.replace('<!-- @SCRIPTS -->', '\n'.join(scripts))
    if '<!-- @' in html:
        sys.exit('Ukjent plassholder igjen i malen.')

    banner = ('<!--\n  Bolig Team Steigen — bygget ' + datetime.date.today().strftime('%d.%m.%Y') +
              ' fra src/ med tools/build.py. Ikke rediger denne filen direkte.\n'
              '  Dataene er kryptert og kan bare åpnes med riktig innlogging.\n-->\n')
    out.write_text(html.replace('<!doctype html>\n', '<!doctype html>\n' + banner, 1), 'utf-8')
    print(f'{out.name}: {len(data["records"])} boliger, {out.stat().st_size // 1024} KB'
          + (' (salt gjenbrukt)' if kdf else ' (nytt salt)'))


if __name__ == '__main__':
    main()

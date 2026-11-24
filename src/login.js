/*
 * Innlogging og kryptering.
 *
 * Dataene i index.html er kryptert (AES-256-GCM). Nøkkelen avledes av brukernavn og passord
 * med PBKDF2-SHA256, så uten riktig innlogging kan ikke dataene leses — heller ikke i kildekoden.
 * Endringer som lagres i nettleseren, krypteres med den samme nøkkelen.
 */
(function (global) {
  'use strict';

  const enc = new TextEncoder();
  const dec = new TextDecoder();
  const SESSION_KEY = 'boligteam.okt';
  let aesKey = null;

  const b64 = {
    to(buf) {
      const bytes = new Uint8Array(buf);
      let s = '';
      for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      return btoa(s);
    },
    from(str) {
      return Uint8Array.from(atob(str), (c) => c.charCodeAt(0));
    },
  };

  function readPayload() {
    const node = document.getElementById('payload');
    if (!node) return null;
    try { return JSON.parse(node.textContent); } catch (e) { return null; }
  }

  async function deriveKey(user, pass, p) {
    const base = await crypto.subtle.importKey(
      'raw', enc.encode(user.trim().toLowerCase() + '\n' + pass), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey(
      { name: 'PBKDF2', salt: b64.from(p.salt), iterations: p.iter, hash: 'SHA-256' },
      base, { name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
  }

  async function decryptWith(key, blob) {
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64.from(blob.iv) }, key, b64.from(blob.ct));
    return dec.decode(plain);
  }

  async function encrypt(text) {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, aesKey, enc.encode(text));
    return { iv: b64.to(iv), ct: b64.to(ct) };
  }

  const decrypt = (blob) => decryptWith(aesKey, blob);

  // Innloggingen huskes i denne fanen til den lukkes (sessionStorage).
  async function fromSession(p) {
    try {
      const raw = sessionStorage.getItem(SESSION_KEY);
      if (!raw) return null;
      const key = await crypto.subtle.importKey('raw', b64.from(raw), { name: 'AES-GCM' }, true, ['encrypt', 'decrypt']);
      return { key, text: await decryptWith(key, p) };
    } catch (e) {
      return null;
    }
  }

  async function remember(key) {
    try { sessionStorage.setItem(SESSION_KEY, b64.to(await crypto.subtle.exportKey('raw', key))); } catch (e) { /* ingen økt */ }
  }

  function unlock() {
    const screen = document.getElementById('login');
    const form = document.getElementById('loginForm');
    const user = document.getElementById('loginUser');
    const pass = document.getElementById('loginPass');
    const error = document.getElementById('loginError');
    const button = document.getElementById('loginBtn');
    const toggle = document.getElementById('pwToggle');
    const app = document.getElementById('app');

    const showError = (msg) => {
      error.textContent = msg;
      form.classList.remove('is-shaking');
      void form.offsetWidth; // start animasjonen på nytt
      form.classList.add('is-shaking');
    };

    toggle.addEventListener('click', () => {
      const show = pass.type === 'password';
      pass.type = show ? 'text' : 'password';
      toggle.setAttribute('aria-pressed', String(show));
      toggle.setAttribute('aria-label', show ? 'Skjul passord' : 'Vis passord');
      pass.focus();
    });

    return new Promise((resolve) => {
      const finish = (key, text) => {
        aesKey = key;
        const data = JSON.parse(text);
        document.body.classList.add('is-unlocked');
        app.hidden = false;
        screen.classList.add('is-leaving');
        setTimeout(() => { screen.hidden = true; }, 350);
        resolve(data);
      };

      const p = readPayload();
      if (!p) { showError('Fant ingen data i filen.'); button.disabled = true; return; }
      if (!global.crypto || !crypto.subtle) {
        showError('Nettleseren støtter ikke kryptering. Åpne filen i Chrome, Edge, Firefox eller Safari.');
        button.disabled = true;
        return;
      }

      fromSession(p).then((s) => {
        if (s) return finish(s.key, s.text);
        screen.classList.add('is-ready');
        user.focus();
      });

      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        error.textContent = '';
        if (!user.value.trim() || !pass.value) { showError('Fyll inn bruker og passord.'); return; }
        button.disabled = true;
        button.classList.add('is-busy');
        try {
          const key = await deriveKey(user.value, pass.value, p);
          const text = await decryptWith(key, p);
          await remember(key);
          finish(key, text);
        } catch (err) {
          showError('Feil bruker eller passord.');
          pass.select();
        } finally {
          button.disabled = false;
          button.classList.remove('is-busy');
        }
      });
    });
  }

  function logout() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) { /* ingenting */ }
    location.reload();
  }

  global.BoligAuth = { unlock, encrypt, decrypt, logout };
})(window);

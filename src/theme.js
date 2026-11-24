/*
 * Lyst / mørkt tema.
 * Uten valg følger appen innstillingen i operativsystemet; et valg huskes i nettleseren.
 * (Selve temaet settes allerede i <head>, så siden ikke blinker ved oppstart.)
 */
(function () {
  'use strict';

  const KEY = 'boligteam.tema';
  const root = document.documentElement;
  const media = window.matchMedia('(prefers-color-scheme: dark)');
  const current = () => root.getAttribute('data-theme') || (media.matches ? 'dark' : 'light');

  function update() {
    const dark = current() === 'dark';
    document.querySelectorAll('.theme-toggle').forEach((b) => {
      b.setAttribute('aria-pressed', String(dark));
      b.setAttribute('aria-label', dark ? 'Bytt til lyst tema' : 'Bytt til mørkt tema');
      b.title = dark ? 'Lyst tema' : 'Mørkt tema';
    });
  }

  document.querySelectorAll('.theme-toggle').forEach((b) => b.addEventListener('click', () => {
    const next = current() === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    try { localStorage.setItem(KEY, next); } catch (e) { /* ingen lagring */ }
    update();
  }));

  if (media.addEventListener) media.addEventListener('change', update);
  update();
})();

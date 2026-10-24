'use strict';
/* =====================================================================
   Online mode only (config.js sets CAL_API): password login, and keeping
   your own settings / icons / task ticks the same on every device.
   Your Google events themselves always come straight from Google.
   Loaded after app.js (uses S, render, save) and before sync.js.
   ===================================================================== */
(() => {
  const API = window.CAL_API || '';
  if (!API) return; // running on your computer: nothing to do

  const TOKEN_KEY = 'glasscal.token', REV_KEY = 'glasscal.cloudrev';
  // How each device looks (light/dark, colour theme, month style, which view is open) stays on that device,
  // so your phone being dark never changes your computer, and a second open copy can't undo your choice.
  const DEVICE_ONLY = ['theme', 'bg', 'look', 'view'];
  const get = k => { try { return localStorage.getItem(k) || ''; } catch { return ''; } };
  const put = (k, v) => { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); } catch { } };
  let token = get(TOKEN_KEY), pushed = '', timer = null, waiters = [];

  /* ---------- login screen ---------- */
  let ov = null;
  function showLogin(msg = '') {
    if (!ov) {
      ov = document.createElement('div');
      ov.className = 'cloud-login';
      ov.innerHTML = `<form class="glass"><h2>Calendar</h2>
        <label>Password<input type="password" id="clPass" autocomplete="current-password" autofocus></label>
        <div class="cl-msg" id="clMsg"></div><button class="btn primary" id="clGo">Open</button></form>`;
      document.body.append(ov);
      ov.querySelector('form').onsubmit = async e => {
        e.preventDefault();
        const go = ov.querySelector('#clGo'), m = ov.querySelector('#clMsg'), pw = ov.querySelector('#clPass');
        go.disabled = true; m.textContent = '';
        try {
          const r = await fetch(API + '/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ password: pw.value }) });
          const j = await r.json().catch(() => ({}));
          if (!r.ok) { m.textContent = j.error || 'Could not sign in'; go.disabled = false; pw.select(); return; }
          token = j.token; put(TOKEN_KEY, token); pw.value = '';
          ov.remove(); ov = null;
          const w = waiters; waiters = []; w.forEach(f => f());
        } catch { m.textContent = 'No connection — check your internet.'; go.disabled = false; }
      };
    }
    ov.querySelector('#clMsg').textContent = msg;
    setTimeout(() => ov?.querySelector('#clPass')?.focus(), 50);
  }
  const needLogin = (msg = 'Please sign in again.') => new Promise(res => { token = ''; put(TOKEN_KEY, ''); waiters.push(res); showLogin(msg); });

  /* ---------- your settings, the same on every device ---------- */
  // Everything except events that live in Google/Outlook (those are fetched fresh), but keep the ones you
  // personally decorated (icon, task tick) so that survives too.
  function payload() {
    const remote = new Set(S.cals.filter(c => c.source).map(c => c.id));
    return {
      groups: S.groups,
      cals: S.cals.map(c => c.source ? { ...c, source: { ...c.source, lastSync: undefined, error: undefined } } : c),
      events: S.events.filter(e => !remote.has(e.calId) || e.icon || e.task || (e.images && e.images.length) || (e.doneDates && e.doneDates.length)),
      settings: Object.fromEntries(Object.entries(S.settings).filter(([k]) => !DEVICE_ONLY.includes(k))),
    };
  }
  async function call(path, init = {}) {
    return fetch(API + path, { ...init, headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + token, ...(init.headers || {}) } });
  }
  async function pull() {
    const r = await call('/state');
    if (r.status === 401) throw new Error('login');
    if (!r.ok) throw new Error('state ' + r.status);
    const j = await r.json();
    const mine = +get(REV_KEY) || 0;
    if (j.data && j.updated_at > mine) {
      S.groups = j.data.groups || S.groups; S.cals = j.data.cals || S.cals; S.events = j.data.events || [];
      const mine = Object.fromEntries(DEVICE_ONLY.map(k => [k, S.settings[k]]));
      S.settings = { ...S.settings, ...(j.data.settings || {}), ...mine }; // this device keeps its own look
      window.ensureFun?.(); put(REV_KEY, String(j.updated_at)); save(); render();
    }
    pushed = JSON.stringify(payload());
  }
  async function push() {
    if (!token) return;
    const p = payload();
    // A brand-new device that hasn't connected Google yet only holds the sample data: don't overwrite the real saved settings with it.
    if (!p.cals.some(c => c.source) && !(+get(REV_KEY))) return;
    const str = JSON.stringify(p);
    if (str === pushed) return;
    const rev = Date.now();
    try {
      const r = await call('/state', { method: 'PUT', body: JSON.stringify({ data: p, updated_at: rev }) });
      if (r.ok) { pushed = str; put(REV_KEY, String(rev)); } else if (r.status === 401) needLogin();
    } catch { /* offline: the next change tries again */ }
  }

  window.CalCloud = {
    token: () => token,
    needLogin,
    changed() { clearTimeout(timer); timer = setTimeout(push, 4000); },
    async ready() {
      for (let i = 0; i < 3; i++) {
        if (!token) await new Promise(res => { waiters.push(res); showLogin(); });
        try { await pull(); return; }
        catch (e) { if (e.message === 'login') { await needLogin('Please sign in again.'); continue; } return; } // offline: carry on with what's on this device
      }
    },
  };
  addEventListener('visibilitychange', () => { if (document.hidden) push(); });
})();

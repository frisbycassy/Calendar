'use strict';
/* =====================================================================
   Live sync: Google Calendar + Outlook/Microsoft 365 (two-way) and
   .ics calendar links (read-only). Needs server.js (Start Calendar.bat).
   Loaded after app.js and uses its helpers and state (S, render, toast…).
   ===================================================================== */
const Sync = (() => {
  const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const PAST_DAYS = 120, FUTURE_DAYS = 400;
  const GAPI = 'https://www.googleapis.com/calendar/v3';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const NAMES = { google: 'Google', microsoft: 'Outlook', ics: 'calendar link' };
  let status = null, busy = false, timer = null, lastRun = 0;
  const tokens = {};

  /* ---------- plumbing ---------- */
  // Online (config.js sets CAL_API) the helper lives at your Supabase back-end and needs the password session.
  const API = window.CAL_API || '';
  async function server(p, opts = {}) {
    let url = p; const headers = { 'Content-Type': 'application/json', ...(opts.headers || {}) };
    if (API) { url = API + p.replace(/^\/api/, ''); const t = window.CalCloud?.token(); if (t) headers.Authorization = 'Bearer ' + t; }
    const r = await fetch(url, { ...opts, headers });
    const isJson = (r.headers.get('content-type') || '').includes('json');
    const body = isJson ? await r.json() : await r.text();
    if (r.status === 401 && isJson && body.error === 'login') {
      if (API) window.CalCloud.needLogin(); else location.href = '/login';
      throw new Error('Please sign in again');
    }
    if (!r.ok) throw new Error((isJson ? body.error : body) || r.statusText);
    return body;
  }
  async function refreshStatus() {
    if (location.protocol === 'file:') return status = null;
    try { status = await server('/api/status'); } catch { status = null; }
    return status;
  }
  async function token(acc) {
    const t = tokens[acc];
    if (t && t.exp > Date.now() + 60_000) return t.v;
    const j = await server('/api/token?id=' + encodeURIComponent(acc));
    tokens[acc] = { v: j.access_token, exp: j.expires_at };
    return j.access_token;
  }
  const sleep = ms => new Promise(res => setTimeout(res, ms));
  async function api(acc, method, url, body) {
    let r;
    for (let attempt = 0; ; attempt++) {
      const t = await token(acc);
      r = await fetch(url, {
        method, body: body ? JSON.stringify(body) : undefined,
        headers: { Authorization: 'Bearer ' + t, ...(body ? { 'Content-Type': 'application/json' } : {}) },
      });
      // Google/Microsoft say "slow down" with 429 (or 403 + rate/quota reason): wait and try again.
      if ((r.status === 429 || r.status === 403 || r.status === 503) && attempt < 5) {
        const why = await r.clone().json().then(j => JSON.stringify(j.error || '')).catch(() => '');
        if (r.status !== 403 || /rate ?limit|quota|userRateLimit/i.test(why)) {
          const wait = +(r.headers.get('Retry-After') || 0) * 1000 || 1000 * 2 ** attempt + Math.random() * 400;
          await sleep(wait);
          continue;
        }
      }
      break;
    }
    if (r.status === 401) delete tokens[acc];
    if (method === 'DELETE' && [204, 404, 410].includes(r.status)) return null;
    const j = r.status === 204 ? null : await r.json().catch(() => null);
    if (!r.ok) throw new Error(j?.error?.message || `${r.status} ${r.statusText}`);
    return j;
  }
  const isRemoteCal = c => !!(c?.source && c.source.type !== 'ics');
  const calById = id => S.cals.find(c => c.id === id);
  const stripHtml = s => s.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
  const baseEvent = (cal, remoteId) => ({
    id: `${cal.id}:${remoteId}`, calId: cal.id, repeat: 'none', until: '', days: [], exdates: [], doneDates: [], task: false, icon: '',
  });

  /* ---------- Google Calendar ---------- */
  const google = {
    async calendars(acc) {
      const out = []; let pt = '';
      do {
        const j = await api(acc, 'GET', `${GAPI}/users/me/calendarList?maxResults=250${pt ? '&pageToken=' + pt : ''}`);
        out.push(...(j.items || [])); pt = j.nextPageToken;
      } while (pt);
      return out.map(c => ({
        remoteId: c.id, name: c.summaryOverride || c.summary, color: c.backgroundColor || '#4285f4',
        readonly: !['owner', 'writer'].includes(c.accessRole), visible: c.selected !== false, primary: !!c.primary,
        defRem: (c.defaultReminders || []).find(r => r.method === 'popup')?.minutes ?? c.defaultReminders?.[0]?.minutes ?? null,
      }));
    },
    async events(acc, cal, from, to) {
      const out = []; let pt = '';
      do {
        const q = new URLSearchParams({ singleEvents: 'true', timeMin: from.toISOString(), timeMax: to.toISOString(), maxResults: '2500' });
        if (pt) q.set('pageToken', pt);
        const j = await api(acc, 'GET', `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events?${q}`);
        out.push(...(j.items || [])); pt = j.nextPageToken;
      } while (pt);
      return out.filter(g => g.status !== 'cancelled' && g.start).map(g => {
        const allDay = !!g.start.date;
        const rem = g.reminders?.useDefault ? cal.source.defRem ?? null : (g.reminders?.overrides || []).reduce((m, r) => m == null || r.minutes < m ? r.minutes : m, null);
        return {
          ...baseEvent(cal, g.id), title: g.summary || '(No title)', allDay,
          start: allDay ? g.start.date + 'T00:00' : toStr(new Date(g.start.dateTime)),
          end: allDay ? ymd(addDays(parse(g.end.date), -1)) + 'T00:00' : toStr(new Date(g.end.dateTime)),
          location: g.location || '', notes: stripHtml(g.description || ''), reminder: rem,
          remote: { id: g.id, series: g.recurringEventId || null, link: g.htmlLink || '' },
          attendees: (g.attendees || []).filter(a => !a.resource).map(a => ({
            email: a.email || '', name: a.displayName || '', status: a.responseStatus || 'needsAction', organizer: !!a.organizer, self: !!a.self,
          })),
          organizer: g.organizer && !g.organizer.self ? (g.organizer.displayName || g.organizer.email || '') : '',
          meet: g.hangoutLink || '',
          ...(/<img[^>]+src=["']?cid:/i.test(g.description || '') ? { wantImg: ((g.description || '').match(/src=["']?cid:([^"'>\s]+)/gi) || []).map(s => s.replace(/^.*cid:/i, '')) } : {}),
        };
      });
    },
    body(ev, isNew) {
      const s = parse(ev.start);
      const b = {
        summary: ev.title, location: ev.location || '', description: ev.notes || '',
        start: ev.allDay ? { date: ymd(s), dateTime: null, timeZone: null } : { dateTime: ev.start + ':00', timeZone: TZ, date: null },
        end: ev.allDay ? { date: ymd(addDays(parse(ev.end), 1)), dateTime: null, timeZone: null } : { dateTime: ev.end + ':00', timeZone: TZ, date: null },
        reminders: { useDefault: false, overrides: ev.reminder == null ? [] : [{ method: 'popup', minutes: Math.min(ev.reminder, 40320) }] },
      };
      // Guests are only sent when they changed, so replies already given are kept.
      if (isNew || ev._guestsChanged) b.attendees = (ev.attendees || []).map(a => ({ email: a.email, ...(a.name ? { displayName: a.name } : {}), responseStatus: a.status || 'needsAction' }));
      if (isNew && isRecurring(ev)) b.recurrence = ['RRULE:' + rruleOf(ev)];
      return b;
    },
    // sendUpdates: guests get an email only when "Email guests" was ticked for this save.
    create: (acc, cal, ev) => api(acc, 'POST', `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events?sendUpdates=${ev._notify ? 'all' : 'none'}`, google.body(ev, true)),
    update: (acc, cal, ev) => api(acc, 'PATCH', `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events/${encodeURIComponent(ev.remote.id)}?sendUpdates=${ev._notify ? 'all' : 'none'}`, google.body(ev, false)),
    remove: (acc, cal, ev) => api(acc, 'DELETE', `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events/${encodeURIComponent(ev.remote.id)}`),
  };

  /* ---------- Outlook / Microsoft Graph ---------- */
  const GRAPH_COLORS = { lightBlue: '#4c9be8', lightGreen: '#40c057', lightOrange: '#f76707', lightGray: '#868e96', lightYellow: '#fab005', lightTeal: '#12b886', lightPink: '#e8467c', lightBrown: '#a0522d', lightRed: '#e03131', maxColor: '#7950f2', auto: '#0f6cbd' };
  const DAYN = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  const microsoft = {
    async calendars(acc) {
      const out = []; let url = `${GRAPH}/me/calendars?$top=100`;
      while (url) { const j = await api(acc, 'GET', url); out.push(...j.value); url = j['@odata.nextLink']; }
      return out.map(c => ({
        remoteId: c.id, name: c.name, color: c.hexColor || GRAPH_COLORS[c.color] || '#0f6cbd',
        readonly: c.canEdit === false, visible: true, primary: !!c.isDefaultCalendar, defRem: null,
      }));
    },
    async events(acc, cal, from, to) {
      const sel = 'id,subject,start,end,isAllDay,location,bodyPreview,seriesMasterId,webLink,isReminderOn,reminderMinutesBeforeStart,isCancelled,attendees,organizer,isOrganizer,onlineMeeting';
      const out = [];
      let url = `${GRAPH}/me/calendars/${encodeURIComponent(cal.source.remoteId)}/calendarView?startDateTime=${from.toISOString()}&endDateTime=${to.toISOString()}&$top=250&$select=${sel}`;
      while (url) { const j = await api(acc, 'GET', url); out.push(...j.value); url = j['@odata.nextLink']; }
      // Graph returns times in UTC; all-day events are "floating" dates.
      const utc = s => new Date(s.slice(0, 19) + 'Z');
      return out.filter(e => !e.isCancelled).map(e => ({
        ...baseEvent(cal, e.id), title: e.subject || '(No title)', allDay: !!e.isAllDay,
        start: e.isAllDay ? e.start.dateTime.slice(0, 10) + 'T00:00' : toStr(utc(e.start.dateTime)),
        end: e.isAllDay ? ymd(addDays(parse(e.end.dateTime.slice(0, 10)), -1)) + 'T00:00' : toStr(utc(e.end.dateTime)),
        location: e.location?.displayName || '', notes: e.bodyPreview || '',
        reminder: e.isReminderOn ? e.reminderMinutesBeforeStart : null,
        remote: { id: e.id, series: e.seriesMasterId || null, link: e.webLink || '' },
        attendees: (e.attendees || []).filter(a => a.type !== 'resource').map(a => ({
          email: a.emailAddress?.address || '', name: a.emailAddress?.name || '',
          status: { accepted: 'accepted', declined: 'declined', tentativelyAccepted: 'tentative', organizer: 'accepted' }[a.status?.response] || 'needsAction',
        })),
        organizer: e.isOrganizer === false ? (e.organizer?.emailAddress?.name || e.organizer?.emailAddress?.address || '') : '',
        meet: e.onlineMeeting?.joinUrl || '',
      }));
    },
    body(ev, isNew) {
      const s = parse(ev.start), e = ev.allDay ? addDays(parse(ev.end), 1) : parse(ev.end);
      const dt = d => ev.allDay ? { dateTime: ymd(d) + 'T00:00:00', timeZone: 'UTC' } : { dateTime: d.toISOString().slice(0, 19), timeZone: 'UTC' };
      const b = {
        subject: ev.title, start: dt(s), end: dt(e), isAllDay: !!ev.allDay, location: { displayName: ev.location || '' },
        isReminderOn: ev.reminder != null, reminderMinutesBeforeStart: ev.reminder ?? 0,
      };
      if (isNew || ev._guestsChanged) b.attendees = (ev.attendees || []).map(a => ({ emailAddress: { address: a.email, name: a.name || a.email }, type: 'required' }));
      // Only send notes when they changed from the plain-text preview, so rich Outlook bodies aren't flattened.
      if (isNew || ev._notesChanged) b.body = { contentType: 'text', content: ev.notes || '' };
      if (isNew && isRecurring(ev)) {
        let pattern;
        const wd = ev.days?.length ? ev.days : [s.getDay()];
        switch (ev.repeat) {
          case 'daily': pattern = { type: 'daily', interval: 1 }; break;
          case 'weekdays': pattern = { type: 'weekly', interval: 1, daysOfWeek: DAYN.slice(1, 6), firstDayOfWeek: 'monday' }; break;
          case 'weekly': case 'biweekly': pattern = { type: 'weekly', interval: ev.repeat === 'weekly' ? 1 : 2, daysOfWeek: wd.map(d => DAYN[d]), firstDayOfWeek: 'monday' }; break;
          case 'monthly': pattern = { type: 'absoluteMonthly', interval: 1, dayOfMonth: s.getDate() }; break;
          case 'yearly': pattern = { type: 'absoluteYearly', interval: 1, dayOfMonth: s.getDate(), month: s.getMonth() + 1 }; break;
        }
        b.recurrence = { pattern, range: ev.until ? { type: 'endDate', startDate: ymd(s), endDate: ev.until } : { type: 'noEnd', startDate: ymd(s) } };
      }
      return b;
    },
    create: (acc, cal, ev) => api(acc, 'POST', `${GRAPH}/me/calendars/${encodeURIComponent(cal.source.remoteId)}/events`, microsoft.body(ev, true)),
    update: (acc, cal, ev) => api(acc, 'PATCH', `${GRAPH}/me/events/${encodeURIComponent(ev.remote.id)}`, microsoft.body(ev, false)),
    remove: (acc, cal, ev) => api(acc, 'DELETE', `${GRAPH}/me/events/${encodeURIComponent(ev.remote.id)}`),
  };
  const P = { google, microsoft };

  /* ---------- pulling ---------- */
  async function discover(acc) {
    let g = S.groups.find(x => x.account === acc.id);
    if (!g) { g = { id: 'g-' + uid(), name: acc.email, account: acc.id, provider: acc.provider, ignored: [] }; S.groups.push(g); }
    const list = await P[acc.provider].calendars(acc.id);
    for (const rc of list) {
      if ((g.ignored || []).includes(rc.remoteId)) continue;
      let cal = S.cals.find(c => c.source?.account === acc.id && c.source.remoteId === rc.remoteId);
      if (!cal) {
        cal = { id: 'c-' + uid(), name: rc.name, color: rc.color, icon: rc.primary ? '📅' : '', groupId: g.id, visible: rc.readonly ? rc.visible : true, // your own calendars always arrive switched on
          readonly: rc.readonly, source: { type: acc.provider, account: acc.id, remoteId: rc.remoteId } };
        S.cals.push(cal);
      }
      cal.readonly = rc.readonly;
      cal.source.defRem = rc.defRem;
    }
  }
  async function syncCal(cal) {
    const src = cal.source;
    const today = startOfDay(new Date());
    let evs;
    if (src.type === 'ics') {
      const parsed = parseICS(await server('/api/ics?url=' + encodeURIComponent(src.url)));
      src.feedName = parsed.name || '';
      evs = parsed.events.map(e => ({ ...e, id: `${cal.id}:${e.uid || e.id}`, calId: cal.id }));
    } else {
      evs = await P[src.type].events(src.account, cal, addDays(today, -PAST_DAYS), addDays(today, FUTURE_DAYS));
    }
    // Keep things that only live in this app: chosen icon, task tick-boxes.
    const mine = S.events.filter(e => e.calId === cal.id);
    const old = new Map(mine.map(e => [e.id, e]));
    // If Google gave an event a new id (e.g. after it became all-day), fall back to one with the same name.
    const byName = new Map(mine.map(e => [(e.title || '').trim().toLowerCase(), e]));
    for (const e of evs) {
      tidyEvent(e);
      const o = old.get(e.id) || byName.get((e.title || '').trim().toLowerCase());
      if (o) { e.icon = o.icon || ''; e.images = o.images || undefined; e.imgTried = o.imgTried; e.task = !!o.task; e.doneDates = o.doneDates || []; }
      else if (/^to[\s-]?do$/i.test(cal.name.trim())) e.task = true; // anything new in a "To Do" calendar gets a tick circle
    }
    S.events = S.events.filter(e => e.calId !== cal.id).concat(evs);
    src.lastSync = Date.now();
    delete src.error;
  }
  async function syncAll({ loud = false } = {}) {
    // Don't download while your changes are still being sent — that would briefly show the old version.
    // (The send queue refreshes the calendars itself when it's done.)
    if (busy || pending) return;
    await refreshStatus();
    const sources = S.cals.filter(c => c.source);
    if (!status) { if (loud) openDialog(); return; }
    if (!sources.length && !status.accounts.length) { if (loud) openDialog(); return; }
    busy = true; spin(true);
    const errors = [];
    for (const acc of status.accounts) {
      try { await discover(acc); } catch (e) { errors.push(`${acc.email}: ${e.message}`); }
    }
    for (const cal of S.cals.filter(c => c.source)) {
      if (cal.source.account && !status.accounts.some(a => a.id === cal.source.account)) { cal.source.error = 'Account disconnected'; continue; }
      try { await syncCal(cal); } catch (e) { cal.source.error = e.message; errors.push(`${cal.name}: ${e.message}`); }
    }
    busy = false; if (!pending) spin(false); lastRun = Date.now();
    // First time Google is connected: clear out the starter calendars right away (a backup is kept; Restore is offered).
    const removed = removeSamples();
    if (removed != null) {
      toast(`Removed the starter calendars (${removed} sample events). Your Google calendars are untouched.`, { icon: 'trash', ms: 12000,
        action: ['Restore', () => { try { Object.assign(S, JSON.parse(localStorage.getItem('glasscal.backup-before-removing-samples'))); S.settings.samplesRemoved = true; save(); render(); } catch { toast('Could not find the backup', { icon: 'alert' }); } }] });
    }
    save(); render();
    fetchInvitePictures();
    if ($('#dlgSync').open) drawDialog();
    if (errors.length) toast(`Sync problem — ${errors[0]}`, { icon: 'alert', ms: 9000 });
    else if (loud) toast('Everything is up to date', { icon: 'cloud' });
  }
  function spin(on) { $('#syncBtn').classList.toggle('spin', on); }

  /* ---------- pushing local edits ---------- */
  // before: copy of the event before the change (or null if new); after: the live event (or null if deleted)
  // Changes are sent one at a time (never a flood of requests), and the affected calendars
  // are refreshed once when the queue goes quiet.
  let queue = Promise.resolve(), pending = 0, resyncTimer = null;
  const dirty = new Set();
  function enqueue(job) {
    pending++; spin(true);
    queue = queue.then(job).catch(() => { }).finally(() => { if (--pending === 0 && !busy) spin(false); });
    return queue;
  }
  function scheduleResync(cals) {
    cals.forEach(c => dirty.add(c));
    clearTimeout(resyncTimer);
    resyncTimer = setTimeout(() => enqueue(async () => {
      const list = [...dirty]; dirty.clear();
      for (const c of list) { try { await syncCal(c); } catch { } }
      save(); render();
    }), 1200);
  }
  let lastErr = 0;
  const failToast = (who, e) => {
    if (Date.now() - lastErr < 4000) return; // one message, not a stack of them
    lastErr = Date.now();
    toast(`Couldn’t save to “${who}”: ${e.message}`, { icon: 'alert', ms: 9000 });
  };

  function pushSave(before, after) {
    const oc = before && calById(before.calId), nc = after && calById(after.calId);
    if (!isRemoteCal(oc) && !isRemoteCal(nc)) return;
    const who = (nc || oc).name;
    if (!status) { failToast(who, new Error('open the calendar with the desktop Calendar icon')); return; }
    return enqueue(() => doPush(before, after, oc, nc, who));
  }
  async function doPush(before, after, oc, nc, who) {
    const touched = [oc, nc].filter(isRemoteCal);
    try {
      if (isRemoteCal(oc) && before.remote?.id && (!after || after.calId !== before.calId)) await P[oc.source.type].remove(oc.source.account, oc, before);
      if (isRemoteCal(nc)) {
        after._guestsChanged = JSON.stringify((before?.attendees || []).map(a => a.email)) !== JSON.stringify((after.attendees || []).map(a => a.email));
        if (after.remote?.id && before && before.calId === after.calId) {
          after._notesChanged = before.notes !== after.notes;
          await P[nc.source.type].update(nc.source.account, nc, after);
          delete after._notesChanged;
        } else {
          const created = await P[nc.source.type].create(nc.source.account, nc, after);
          after.id = `${nc.id}:${created.id}`; // keeps the icon / task box after the refresh
          after.remote = { id: created.id };
        }
      }
    } catch (e) {
      failToast(who, e);
    } finally {
      if (after) { delete after._guestsChanged; delete after._notify; }
      scheduleResync(touched);
    }
  }

  /* "Save to all" / "Delete all" on a repeating Google/Outlook event: one request for the whole series.
     rep = one edited date of the series (already holding the new details), before = its old copy. */
  function pushSeries(before, rep) {
    const cal = calById(rep.calId);
    if (!isRemoteCal(cal) || !rep.remote?.series) return;
    return enqueue(async () => {
      const acc = cal.source.account, sid = rep.remote.series;
      try {
        rep._guestsChanged = JSON.stringify((before.attendees || []).map(a => a.email)) !== JSON.stringify((rep.attendees || []).map(a => a.email));
        rep._notesChanged = before.notes !== rep.notes;
        let dur = parse(rep.end) - parse(rep.start), span = dayDiff(parse(rep.end), parse(rep.start));
        // Safety: only change how long each date lasts if you actually changed it, and never to something silly.
        // (Otherwise a whole weekly series can end up as overlapping multi-week events.)
        const oldDur = parse(before.end) - parse(before.start), oldSpan = dayDiff(parse(before.end), parse(before.start));
        const keepLength = (rep.allDay ? span === oldSpan : dur === oldDur) || span > 13 || dur < 0 || dur > 14 * 864e5;
        if (keepLength) { dur = null; span = null; }
        if (cal.source.type === 'google') {
          const base = `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events/${encodeURIComponent(sid)}`;
          const m = await api(acc, 'GET', base);
          const body = google.body(rep, false);
          // Always keep the series' first date; apply all-day on/off, the new time and (if changed) the length.
          const wasAllDay = !!m.start?.date;
          const day0 = (m.start.date || m.start.dateTime).slice(0, 10);
          if (rep.allDay) {
            const len = span ?? (wasAllDay ? Math.max(0, dayDiff(parse(m.end.date.slice(0, 10)), parse(day0)) - 1) : 0);
            body.start = { date: day0, dateTime: null, timeZone: null };
            body.end = { date: ymd(addDays(parse(day0), len + 1)), dateTime: null, timeZone: null };
          } else {
            const s = parse(day0 + 'T' + rep.start.slice(11));
            const len = dur ?? (wasAllDay ? 3600e3 : new Date(m.end.dateTime) - new Date(m.start.dateTime));
            body.start = { dateTime: toStr(s) + ':00', timeZone: TZ, date: null };
            body.end = { dateTime: toStr(new Date(+s + len)) + ':00', timeZone: TZ, date: null };
          }
          await api(acc, 'PATCH', `${base}?sendUpdates=${rep._notify ? 'all' : 'none'}`, body);
        } else {
          const base = `${GRAPH}/me/events/${encodeURIComponent(sid)}`;
          const m = await api(acc, 'GET', `${base}?$select=start,end,isAllDay`);
          const body = microsoft.body(rep, false);
          const wasAllDay = !!m.isAllDay;
          const first = new Date(m.start.dateTime.slice(0, 19) + 'Z'); // series start (UTC)
          const day0 = wasAllDay ? m.start.dateTime.slice(0, 10) : ymd(first); // … as a local date
          body.isAllDay = !!rep.allDay;
          if (rep.allDay) {
            const len = span ?? (wasAllDay ? Math.max(0, dayDiff(parse(m.end.dateTime.slice(0, 10)), parse(day0)) - 1) : 0);
            body.start = { dateTime: day0 + 'T00:00:00', timeZone: 'UTC' };
            body.end = { dateTime: ymd(addDays(parse(day0), len + 1)) + 'T00:00:00', timeZone: 'UTC' };
          } else {
            const s = parse(day0 + 'T' + rep.start.slice(11));
            const len = dur ?? (wasAllDay ? 3600e3 : new Date(m.end.dateTime.slice(0, 19) + 'Z') - first);
            body.start = { dateTime: s.toISOString().slice(0, 19), timeZone: 'UTC' };
            body.end = { dateTime: new Date(+s + len).toISOString().slice(0, 19), timeZone: 'UTC' };
          }
          await api(acc, 'PATCH', base, body);
        }
      } catch (e) { failToast(cal.name, e); }
      finally { delete rep._guestsChanged; delete rep._notesChanged; delete rep._notify; scheduleResync([cal]); }
    });
  }
  function removeSeries(ev) {
    const cal = calById(ev.calId);
    if (!isRemoteCal(cal) || !ev.remote?.series) return;
    return enqueue(async () => {
      const acc = cal.source.account, sid = encodeURIComponent(ev.remote.series);
      try {
        if (cal.source.type === 'google') await api(acc, 'DELETE', `${GAPI}/calendars/${encodeURIComponent(cal.source.remoteId)}/events/${sid}`);
        else await api(acc, 'DELETE', `${GRAPH}/me/events/${sid}`);
      } catch (e) { failToast(cal.name, e); }
      finally { scheduleResync([cal]); }
    });
  }

  /* ---------- pictures from invitation emails (Gmail) ---------- */
  // Google keeps only a placeholder (cid:…) for a picture that came inside an invitation. The real picture is in the
  // invitation email, so we look it up in Gmail (read-only), shrink it and keep it on the event.
  const GMAIL = 'https://gmail.googleapis.com/gmail/v1/users/me';
  const b64ToBlob = (b64, type) => {
    const bin = atob(b64.replace(/-/g, '+').replace(/_/g, '/')), a = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
    return new Blob([a], { type });
  };
  function imageParts(part, out = []) {
    if (/^image\//.test(part.mimeType || '') && part.body && (part.body.attachmentId || part.body.data) && (part.body.size || 0) > 6000) out.push(part);
    (part.parts || []).forEach(p => imageParts(p, out));
    return out;
  }
  const cidOf = p => ((p.headers || []).find(h => /^content-id$/i.test(h.name))?.value || '').replace(/[<>]/g, '');
  async function pictureFromMail(acc, ev) {
    const title = (ev.title || '').replace(/["()]/g, ' ').trim();
    const q = encodeURIComponent(`subject:(${title}) newer_than:2y`);
    const list = await api(acc, 'GET', `${GMAIL}/messages?maxResults=6&q=${q}`);
    const pics = [];
    for (const m of list.messages || []) {
      const msg = await api(acc, 'GET', `${GMAIL}/messages/${m.id}?format=full`);
      let parts = imageParts(msg.payload);
      const wanted = parts.filter(p => (ev.wantImg || []).includes(cidOf(p)));
      if (wanted.length) parts = wanted;
      for (const p of parts) {
        let data = p.body.data;
        if (!data) data = (await api(acc, 'GET', `${GMAIL}/messages/${m.id}/attachments/${p.body.attachmentId}`)).data;
        const u = await shrinkImage(b64ToBlob(data, p.mimeType));
        if (u) pics.push(u);
        if (pics.length >= 6) break;
      }
      if (pics.length) break; // the invitation that holds the picture
    }
    return pics;
  }
  let hunting = false;
  async function fetchInvitePictures() {
    if (hunting) return; hunting = true;
    try {
      const todo = S.events.filter(e => e.wantImg && !(e.images || []).length && !e.imgTried).slice(0, 10);
      for (const ev of todo) {
        const cal = calById(ev.calId); if (!cal?.source?.account) continue;
        try {
          const pics = await pictureFromMail(cal.source.account, ev);
          if (pics.length) ev.images = pics;
          ev.imgTried = Date.now();
        } catch (e) {
          if (/scope|permission|insufficient|forbidden|not enabled|has not been used|403/i.test(e.message)) {
            if (!S.settings.gmailAsked) { S.settings.gmailAsked = Date.now(); toast('To show pictures from invitation emails, reconnect Google once (Settings → Sync).', { icon: 'cloud', ms: 12000 }); }
            break;
          }
        }
      }
      save(); render();
    } finally { hunting = false; }
  }

  /* ---------- put the built-in holidays into Google (two new calendars) ---------- */
  // Creates "South Africa holidays" and "Fun holidays" in Google and fills them with all-day events for the given
  // years. Safe to run again: anything already there is skipped. Afterwards the built-in copies are switched off.
  async function holidaysToGoogle(years = [2026, 2027, 2028], say = m => toast(m, { icon: 'cloud', ms: 5000 })) {
    await refreshStatus();
    const acc = (status?.accounts || []).find(a => a.provider === 'google');
    if (!acc) throw new Error('Connect your Google account first');
    const SETS = [
      { name: 'South Africa holidays', list: y => zaHolidays(y) },
      { name: 'Fun holidays', list: y => funHolidays(y) },
    ];
    busy = true; spin(true);
    try {
      const have = await google.calendars(acc.id);
      for (const set of SETS) {
        let rc = have.find(c => c.name === set.name && !c.readonly);
        if (!rc) {
          const made = await api(acc.id, 'POST', `${GAPI}/calendars`, { summary: set.name, timeZone: TZ });
          rc = { remoteId: made.id };
          say(`Created “${set.name}” in Google`);
        }
        const id = encodeURIComponent(rc.remoteId), base = `${GAPI}/calendars/${id}/events`;
        // what is already there (so running this twice never doubles anything)
        const exist = new Set(); let pt = '';
        do {
          const q = new URLSearchParams({ singleEvents: 'true', maxResults: '2500', timeMin: `${Math.min(...years)}-01-01T00:00:00Z`, timeMax: `${Math.max(...years) + 1}-01-02T00:00:00Z` });
          if (pt) q.set('pageToken', pt);
          const j = await api(acc.id, 'GET', `${base}?${q}`);
          (j.items || []).forEach(g => g.start?.date && exist.add(g.summary + '|' + g.start.date)); pt = j.nextPageToken;
        } while (pt);
        const todo = years.flatMap(y => set.list(y)).filter(e => !exist.has(e.title + '|' + e.start.slice(0, 10)));
        let n = 0;
        for (const e of todo) {
          const day = e.start.slice(0, 10);
          await api(acc.id, 'POST', `${base}?sendUpdates=none`, {
            summary: e.title, start: { date: day }, end: { date: ymd(addDays(parse(day), 1)) },
            transparency: 'transparent', reminders: { useDefault: false, overrides: [] },
          });
          if (++n % 20 === 0) say(`${set.name}: ${n} of ${todo.length} added…`);
          await sleep(150);
        }
        say(`${set.name}: ${todo.length ? todo.length + ' added' : 'already complete'}`);
      }
      // bring the new calendars in, put "Fun holidays" in its own folder, and retire the built-in copies
      await discover(acc);
      const fun = S.cals.find(c => c.source?.account === acc.id && c.name === 'Fun holidays');
      if (fun) { if (!S.groups.some(g => g.id === 'g-fun')) S.groups.push({ id: 'g-fun', name: 'Fun holidays' }); fun.groupId = 'g-fun'; }
      S.settings.holidaysInGoogle = true;
      S.cals = S.cals.filter(c => c.id !== 'za' && c.id !== 'fun');
      for (const c of S.cals.filter(c => c.source?.account === acc.id && SETS.some(s => s.name === c.name))) await syncCal(c);
      save(); render();
      say('Holidays are now in your Google Calendar ✓');
    } finally { busy = false; spin(false); }
  }

  /* ---------- accounts & feeds ---------- */
  async function connect(provider) {
    // Open the window straight away (so the browser allows it), then point it at the sign-in page.
    const w = window.open('about:blank', 'glasscal-auth', 'width=520,height=720');
    if (!w) return toast('Your browser blocked the sign-in window — allow pop-ups for this site', { icon: 'alert', ms: 8000 });
    try { w.document.write('<p style="font:16px system-ui;padding:24px">Opening sign-in…</p>'); } catch { }
    if (API) {
      try { w.location = (await server(`/api/oauth/${provider}/begin`, { method: 'POST' })).url; }
      catch (e) { w.close(); return toast('Couldn’t start sign-in: ' + e.message, { icon: 'alert', ms: 8000 }); }
    } else w.location = `/oauth/${provider}/start`;
    let done = false, bc = null;
    const finish = async d => {
      if (done || d?.type !== 'glasscal-auth') return; done = true;
      window.removeEventListener('message', on); try { bc?.close(); } catch { }
      if (!d.ok) return toast('Sign-in didn’t finish: ' + d.error, { icon: 'alert', ms: 8000 });
      toast(`Connected ${d.email} — fetching calendars…`, { icon: 'cloud' });
      await syncAll(); drawDialog();
    };
    const on = e => { if (e.origin === location.origin) finish(e.data); };
    window.addEventListener('message', on);
    // Some sign-in pages cut the link between windows, so also listen on a channel shared by this site's windows.
    try { bc = new BroadcastChannel('glasscal-auth'); bc.onmessage = e => finish(e.data); } catch { }
  }
  async function disconnect(accId) {
    const g = S.groups.find(x => x.account === accId);
    if (!confirm(`Disconnect ${g?.name || accId}? Its calendars disappear from here (nothing is deleted in your account).`)) return;
    try { await server('/api/disconnect', { method: 'POST', body: JSON.stringify({ id: accId }) }); } catch { }
    const ids = new Set(S.cals.filter(c => c.source?.account === accId).map(c => c.id));
    S.cals = S.cals.filter(c => !ids.has(c.id));
    S.events = S.events.filter(e => !ids.has(e.calId));
    S.groups = S.groups.filter(x => x.account !== accId);
    delete tokens[accId];
    await refreshStatus(); save(); render(); drawDialog();
  }
  async function addFeed(url, name) {
    if (!status) return toast('Open the calendar with Start Calendar.bat to add links', { icon: 'alert' });
    let g = S.groups.find(x => x.id === 'g-links');
    if (!g) { g = { id: 'g-links', name: 'Subscribed links' }; S.groups.push(g); }
    const cal = { id: 'c-' + uid(), name: name || 'Subscribed calendar', color: COLORS[(S.cals.length * 5) % COLORS.length], icon: '🔗', groupId: g.id, visible: true, readonly: true, source: { type: 'ics', url } };
    spin(true);
    try {
      await syncCal(cal);
      if (!name && cal.source.feedName) cal.name = cal.source.feedName;
      S.cals.push(cal);
      save(); render(); drawDialog();
      toast(`Subscribed to “${cal.name}”`, { icon: 'link' });
      return true;
    } catch (e) {
      S.events = S.events.filter(ev => ev.calId !== cal.id);
      toast('Couldn’t add that link: ' + e.message, { icon: 'alert', ms: 8000 });
    } finally { spin(false); }
  }
  /* Called when a synced calendar is deleted locally: stop syncing it (it stays in the account). */
  function forget(cal) {
    const g = S.groups.find(x => x.id === cal.groupId);
    if (g?.account) (g.ignored ||= []).push(cal.source.remoteId);
  }

  /* ---------- dialog ---------- */
  const ago = t => {
    if (!t) return 'never';
    const m = Math.round((Date.now() - t) / 6e4);
    return m < 1 ? 'just now' : m < 60 ? `${m} min ago` : m < 1440 ? `${Math.round(m / 60)} h ago` : `${Math.round(m / 1440)} days ago`;
  };
  const badge = p => `<span class="pbadge ${p}">${p === 'google' ? 'G' : p === 'microsoft' ? 'O' : svg('link')}</span>`;
  function accountRows(provider) {
    return (status?.accounts || []).filter(a => a.provider === provider).map(a => {
      const cals = S.cals.filter(c => c.source?.account === a.id);
      const last = Math.min(...cals.map(c => c.source.lastSync || 0));
      const err = cals.find(c => c.source.error)?.source.error;
      return `<div class="acct"><div><b>${esc(a.email)}</b><small>${cals.length} calendar${cals.length === 1 ? '' : 's'} · synced ${ago(isFinite(last) ? last : 0)}</small>${err ? `<small class="err">${svg('alert')}${esc(err)}</small>` : ''}</div>
        ${err ? `<button type="button" class="btn ghost" data-sync="connect" data-p="${provider}">${svg('refresh')}<span>Reconnect</span></button>` : ''}
        <button type="button" class="btn ghost" data-sync="disconnect" data-id="${esc(a.id)}">Disconnect</button></div>
      <div class="calrows">${cals.map(c => {
        const n = S.events.filter(e => e.calId === c.id).length;
        return `<div class="crow"><i style="background:${c.color}"></i><span class="cn">${c.icon || ''} ${esc(c.name)}</span>
          <small>${c.source.error ? `<b class="err">${esc(c.source.error)}</b>` : `${n} event${n === 1 ? '' : 's'} loaded`}</small>
          <button type="button" class="btn ghost sm" data-sync="toggle" data-id="${c.id}">${c.visible ? 'Showing' : 'Hidden · show it'}</button></div>`;
      }).join('')}</div>`;
    }).join('');
  }
  function providerCard(p) {
    const ready = status?.[p];
    const title = p === 'google' ? 'Google Calendar / Gmail' : 'Outlook / Microsoft 365';
    const steps = p === 'google' ? `
      <ol>
        <li>Open <a href="https://console.cloud.google.com/projectcreate" target="_blank" rel="noopener">Google Cloud Console</a> and create a project called <i>Glass Calendar</i>.</li>
        <li>Turn on the <a href="https://console.cloud.google.com/apis/library/calendar-json.googleapis.com" target="_blank" rel="noopener">Google Calendar API</a> (click <b>Enable</b>).</li>
        <li>Open <a href="https://console.cloud.google.com/auth/overview" target="_blank" rel="noopener">Google Auth Platform</a> → <b>Get started</b>: app name <i>Glass Calendar</i>, your email, audience <b>External</b>.
          Then under <b>Audience</b> add your Gmail address as a <b>test user</b>.</li>
        <li>Go to <a href="https://console.cloud.google.com/auth/clients" target="_blank" rel="noopener">Clients</a> → <b>Create client</b> → type <b>Web application</b>.
          Under <i>Authorised redirect URIs</i> add:<br><code>http://localhost:5173/oauth/google/callback</code></li>
        <li>Copy the <b>Client ID</b> and <b>Client secret</b> into the boxes below and press <b>Save</b>.</li>
      </ol>
      <p class="hint">While the app is in “Testing”, Google asks you to reconnect every 7 days. To stop that, press <b>Publish app</b> under Audience —
      Google will then show “Google hasn’t verified this app”; that’s expected for your own private app: choose <i>Advanced → Go to Glass Calendar</i>.</p>
      <div class="grid2">
        <label class="fld"><span>Client ID</span><input id="gId" placeholder="…apps.googleusercontent.com" autocomplete="off" value="${esc(status?.googleClientId || '')}"></label>
        <label class="fld"><span>Client secret</span><input id="gSecret" type="password" placeholder="GOCSPX-…" autocomplete="off"></label>
      </div>` : `
      <ol>
        <li>Open <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noopener">App registrations</a>
          (or <a href="https://portal.azure.com" target="_blank" rel="noopener">portal.azure.com</a> → <i>App registrations</i>) and sign in with your Outlook / Microsoft account.
          If it asks, create the free directory it offers.</li>
        <li><b>New registration</b>: name <i>Glass Calendar</i>; account types <b>“Accounts in any organizational directory and personal Microsoft accounts”</b>.</li>
        <li>Redirect URI: platform <b>Public client/native (mobile &amp; desktop)</b>, address:<br><code>http://localhost:5173/oauth/microsoft/callback</code> → <b>Register</b>.</li>
        <li>Copy the <b>Application (client) ID</b> from the Overview page into the box below and press <b>Save</b>.</li>
      </ol>
      <p class="hint">No secret is needed. Work or school accounts may need their IT admin to allow the app.</p>
      <label class="fld"><span>Application (client) ID</span><input id="msId" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" autocomplete="off" value="${esc(status?.msClientId || '')}"></label>`;
    return `<section class="sync-card">
      <header>${badge(p)}<div><b>${title}</b><small>${ready ? 'Two-way sync: changes here go to your account and back.' : 'One-time setup needed (about 10 minutes)'}</small></div>
        <button type="button" class="btn primary" data-sync="connect" data-p="${p}" ${ready ? '' : 'disabled'}>${svg('plus')}<span>Connect account</span></button></header>
      ${accountRows(p)}
      <details ${ready ? '' : 'open'}><summary>${ready ? 'Change app ID' : 'Set up (first time only)'}</summary>${steps}
        <div class="btnrow end"><button type="button" class="btn" data-sync="save" data-p="${p}">${svg('check')}<span>Save</span></button></div></details>
    </section>`;
  }
  function drawDialog() {
    const body = $('#syncBody');
    if (!status) {
      body.innerHTML = `<section class="sync-card warn"><header>${svg('alert')}<div><b>Start the calendar helper to use live sync</b>
        <small>Close this tab, then double-click <b>Start Calendar.bat</b> in your <code>F:\\Calendar</code> folder. The calendar opens at <code>http://localhost:5173</code>.</small></div></header>
        ${location.protocol === 'file:' ? `<p class="hint">You’re using the file version. Your events here are stored separately — before switching, open ${svg('settings')} Settings → <b>Back up</b>, then <b>Restore backup</b> in the new version.</p>` : ''}
        <div class="btnrow end"><button type="button" class="btn" data-sync="retry">${svg('refresh')}<span>Check again</span></button></div></section>`;
      return;
    }
    const feeds = S.cals.filter(c => c.source?.type === 'ics');
    body.innerHTML = providerCard('google') + providerCard('microsoft') + `
      <section class="sync-card"><header>${badge('ics')}<div><b>Subscribe to a calendar link</b><small>Read-only, refreshes automatically. No setup needed.</small></div></header>
        <p class="hint">Works with Google’s <i>Secret address in iCal format</i> (calendar settings → Integrate calendar), Outlook’s <i>published ICS link</i>
        (Settings → Calendar → Shared calendars → Publish), and school, sports or holiday calendars.</p>
        ${feeds.map(c => `<div class="acct"><div><b>${esc(c.name)}</b><small>synced ${ago(c.source.lastSync)}${c.source.error ? ` · <span class="err">${esc(c.source.error)}</span>` : ''}</small></div>
          <button type="button" class="btn ghost" data-sync="unfeed" data-id="${c.id}">Remove</button></div>`).join('')}
        <div class="feedrow"><input id="feedUrl" placeholder="https://… or webcal://… .ics" autocomplete="off"><input id="feedName" placeholder="Name (optional)" autocomplete="off">
        <button type="button" class="btn" data-sync="feed">${svg('link')}<span>Subscribe</span></button></div></section>
      <label class="fld inline"><span>Check for changes every</span><select id="syncEvery">
        ${[2, 5, 10, 15, 30, 60].map(m => `<option value="${m}" ${S.settings.syncEvery === m ? 'selected' : ''}>${m} minutes</option>`).join('')}</select></label>`;
  }
  async function openDialog() {
    await refreshStatus();
    drawDialog();
    if (!$('#dlgSync').open) $('#dlgSync').showModal();
  }
  async function onDialogClick(e) {
    const b = e.target.closest('[data-sync]'); if (!b) return;
    const act = b.dataset.sync;
    if (act === 'connect') return connect(b.dataset.p);
    if (act === 'disconnect') return disconnect(b.dataset.id);
    if (act === 'retry') { await refreshStatus(); drawDialog(); if (status) syncAll(); return; }
    if (act === 'save') {
      const body = b.dataset.p === 'google' ? { googleClientId: $('#gId').value, googleClientSecret: $('#gSecret').value } : { msClientId: $('#msId').value };
      if (b.dataset.p === 'google' && (!body.googleClientId.trim() || (!body.googleClientSecret.trim() && !status.google))) return toast('Paste both the Client ID and the Client secret', { icon: 'alert' });
      if (b.dataset.p === 'microsoft' && !body.msClientId.trim()) return toast('Paste the Application (client) ID', { icon: 'alert' });
      try { status = await server('/api/config', { method: 'POST', body: JSON.stringify(body) }); toast('Saved — now press Connect account', { icon: 'check' }); drawDialog(); }
      catch (err) { toast('Could not save: ' + err.message, { icon: 'alert' }); }
      return;
    }
    if (act === 'feed') {
      const url = $('#feedUrl').value.trim();
      if (!url) return toast('Paste a calendar link first', { icon: 'link' });
      b.disabled = true; await addFeed(url, $('#feedName').value.trim()); b.disabled = false;
      return;
    }
    if (act === 'toggle') {
      const c = calById(b.dataset.id); c.visible = !c.visible;
      save(); render(); drawDialog(); return;
    }
    if (act === 'unfeed') {
      const c = calById(b.dataset.id);
      S.cals = S.cals.filter(x => x !== c); S.events = S.events.filter(ev => ev.calId !== c.id);
      save(); render(); drawDialog();
    }
  }

  /* ---------- start ---------- */
  function schedule() {
    clearInterval(timer);
    timer = setInterval(() => { if (!document.hidden) syncAll(); }, (S.settings.syncEvery || 5) * 60_000);
  }
  async function init() {
    S.settings.syncEvery ||= 5;
    $('#dlgSync form').addEventListener('submit', e => e.preventDefault()); // Enter in a field shouldn't close it
    $('#syncBtn').onclick = () => (S.cals.some(c => c.source) && status) ? syncAll({ loud: true }) : openDialog();
    $('#setSync').onclick = () => { $('#dlgSettings').close(); openDialog(); };
    $('#syncBody').addEventListener('click', onDialogClick);
    $('#syncBody').addEventListener('change', e => { if (e.target.id === 'syncEvery') { S.settings.syncEvery = +e.target.value; save(); schedule(); } });
    document.addEventListener('visibilitychange', () => { if (!document.hidden && Date.now() - lastRun > 60_000) syncAll(); });
    await refreshStatus();
    $('#syncBtn').title = status ? 'Sync now' : 'Connect Google / Outlook';
    if (status && (status.accounts.length || S.cals.some(c => c.source))) syncAll();
    schedule();
  }

  return {
    init, syncAll, pushSave, pushSeries, removeSeries, openDialog, forget, isRemoteCal, holidaysToGoogle, fetchInvitePictures,
    get online() { return !!status; },
    providerName: cal => NAMES[cal?.source?.type] || '',
  };
})();
(window.CalCloud ? CalCloud.ready() : Promise.resolve()).then(() => Sync.init());

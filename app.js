'use strict';
/* =====================================================================
   Glass Calendar — a local, multi-calendar planner for Chrome.
   Data lives in localStorage under STORE_KEY. No server, no tracking.
   ===================================================================== */

const STORE_KEY = 'glasscal.v1';
const HOUR = 48; // px per hour in day/week views

/* ---------------- helpers ---------------- */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const pad = n => String(n).padStart(2, '0');
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const hm = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const toStr = d => `${ymd(d)}T${hm(d)}`;
const parse = s => {
  const [a, b = '00:00'] = s.split('T');
  const [y, m, d] = a.split('-').map(Number);
  const [h, mi] = b.split(':').map(Number);
  return new Date(y, m - 1, d, h || 0, mi || 0);
};
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const addMin = (d, n) => new Date(d.getTime() + n * 60000);
const startOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const dayDiff = (a, b) => Math.round((startOfDay(a) - startOfDay(b)) / 864e5);
const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-5);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (d, o) => d.toLocaleDateString(undefined, o);
const fmtTime = d => S.settings.h24 ? hm(d) : `${d.getHours() % 12 || 12}:${pad(d.getMinutes())}${d.getHours() < 12 ? 'am' : 'pm'}`;
const fmtHour = h => S.settings.h24 ? `${pad(h)}:00` : `${h % 12 || 12}${h < 12 ? 'am' : 'pm'}`;
const startOfWeek = d => { const x = startOfDay(d); return addDays(x, -((x.getDay() - S.settings.weekStart + 7) % 7)); };
const isoWeek = d => {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day);
  return Math.ceil(((t - Date.UTC(t.getUTCFullYear(), 0, 1)) / 864e5 + 1) / 7);
};
const isRecurring = ev => ev.repeat && ev.repeat !== 'none';
/* Hide "[In-person]" tags (added by Outlook/Google) from event names, and "In-person" as a location. */
const IN_PERSON = /\s*[[(]\s*in[\s-]?person\s*[\])]\s*/gi;
function tidyEvent(e) {
  if (e.title) e.title = e.title.replace(IN_PERSON, ' ').replace(/\s{2,}/g, ' ').trim() || e.title;
  if (/^\s*in[\s-]?person\s*$/i.test(e.location || '')) e.location = '';
  return e;
}
const clone = o => JSON.parse(JSON.stringify(o));

/* ---------------- icons (Lucide-style line icons) ---------------- */
const ICONS = {
  menu: '<path d="M4 6h16M4 12h16M4 18h16"/>',
  chevronLeft: '<path d="m15 18-6-6 6-6"/>',
  chevronRight: '<path d="m9 18 6-6-6-6"/>',
  chevronDown: '<path d="m6 9 6 6 6-6"/>',
  plus: '<path d="M5 12h14M12 5v14"/>',
  search: '<circle cx="11" cy="11" r="7.5"/><path d="m20.5 20.5-4.3-4.3"/>',
  calendar: '<rect x="3" y="4.5" width="18" height="17" rx="3"/><path d="M16 2.5v4M8 2.5v4M3 10h18"/>',
  today: '<rect x="3" y="4.5" width="18" height="17" rx="3"/><path d="M16 2.5v4M8 2.5v4M3 10h18"/><circle cx="12" cy="15.5" r="2" fill="currentColor"/>',
  calPlus: '<path d="M21 12V7.5a3 3 0 0 0-3-3H6a3 3 0 0 0-3 3V18.5a3 3 0 0 0 3 3h6"/><path d="M16 2.5v4M8 2.5v4M3 10h18M19 16v6M16 19h6"/>',
  folderPlus: '<path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.7-.9l-.8-1.2A2 2 0 0 0 7.9 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/><path d="M12 10v6M9 13h6"/>',
  sparkles: '<path d="M12 3l1.8 4.9L18.7 9.7l-4.9 1.8L12 16.4l-1.8-4.9L5.3 9.7l4.9-1.8Z"/><path d="M19 15v5M16.5 17.5h5M5 17v3M3.5 18.5h3"/>',
  sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/>',
  moon: '<path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>',
  settings: '<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/>',
  upload: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5M12 3v12"/>',
  download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5M12 15V3"/>',
  bell: '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
  trash: '<path d="M3 6h18M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2M10 11v6M14 11v6"/>',
  copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>',
  repeat: '<path d="m17 2 4 4-4 4"/><path d="M3 11v-1a4 4 0 0 1 4-4h14"/><path d="m7 22-4-4 4-4"/><path d="M21 13v1a4 4 0 0 1-4 4H3"/>',
  pin: '<path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"/><circle cx="12" cy="10" r="3"/>',
  clock: '<circle cx="12" cy="12" r="9.5"/><path d="M12 6.5V12l3.5 2"/>',
  check: '<path d="M20 6 9 17l-5-5"/>',
  checkCircle: '<circle cx="12" cy="12" r="9.5"/><path d="m8.5 12 2.5 2.5 4.5-5"/>',
  more: '<circle cx="12" cy="12" r="1.2"/><circle cx="19" cy="12" r="1.2"/><circle cx="5" cy="12" r="1.2"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
  eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  note: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6M16 13H8M16 17H8"/>',
  keyboard: '<rect x="2" y="5" width="20" height="14" rx="2.5"/><path d="M6 9h.01M10 9h.01M14 9h.01M18 9h.01M6 13h.01M18 13h.01M8 16h8M10 13h4"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
  smile: '<circle cx="12" cy="12" r="9.5"/><path d="M8 14s1.5 2 4 2 4-2 4-2M9 9h.01M15 9h.01"/>',
  cloud: '<path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/>',
  refresh: '<path d="M21 12a9 9 0 0 1-15.5 6.2L3 16"/><path d="M3 21v-5h5"/><path d="M3 12A9 9 0 0 1 18.5 5.8L21 8"/><path d="M21 3v5h-5"/>',
  link: '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
  alert: '<circle cx="12" cy="12" r="9.5"/><path d="M12 7.5v5M12 16.5h.01"/>',
  external: '<path d="M15 3h6v6M10 14 21 3M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
  users: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
  back: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  route: '<circle cx="6" cy="19" r="3"/><path d="M9 19h8.5a3.5 3.5 0 0 0 0-7h-11a3.5 3.5 0 0 1 0-7H15"/><circle cx="18" cy="5" r="3"/>',
  map: '<path d="M14.1 5.1 9.9 3 3 5.9v15.2l6.9-2.9 4.2 2.1 6.9-2.9V2.2Z"/><path d="M9.9 3v15.2M14.1 5.1v15.2"/>',
  vDay: '<rect x="4" y="3" width="16" height="18" rx="3"/><path d="M4 8.5h16M9 13h6"/>',
  vWork: '<rect x="3" y="4" width="18" height="16" rx="3"/><path d="M9 4v16M15 4v16"/>',
  vWeek: '<rect x="3" y="4" width="18" height="16" rx="3"/><path d="M7.5 4v16M12 4v16M16.5 4v16"/>',
  vMonth: '<rect x="3" y="4" width="18" height="16" rx="3"/><path d="M3 9.5h18M3 14.5h18M9 4v16M15 4v16"/>',
  vYear: '<rect x="3" y="3" width="7.5" height="7.5" rx="2"/><rect x="13.5" y="3" width="7.5" height="7.5" rx="2"/><rect x="3" y="13.5" width="7.5" height="7.5" rx="2"/><rect x="13.5" y="13.5" width="7.5" height="7.5" rx="2"/>',
  vAgenda: '<path d="M9 6h12M9 12h12M9 18h12"/><circle cx="4.5" cy="6" r="1.5"/><circle cx="4.5" cy="12" r="1.5"/><circle cx="4.5" cy="18" r="1.5"/>',
};
const svg = (n, cls = '') => `<svg class="i ${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[n] || ''}</svg>`;
const hydrateIcons = (root = document) => $$('[data-icon]', root).forEach(el => {
  el.insertAdjacentHTML('afterbegin', svg(el.dataset.icon));
  el.removeAttribute('data-icon');
});

/* Smart icons: first match wins, so specific words come before generic ones. */
const SMART = [
  [/birthday|b-?day|verjaar/i, '🎂'], [/halloween/i, '🎃'], [/christmas|xmas/i, '🎄'], [/graduat/i, '🎓'], [/back to school/i, '🎒'],
  [/payday|salary/i, '💰'], [/presentation|speech/i, '📊'], [/psych|therap|counsel/i, '🧠'], [/babysit|nanny/i, '👶'], [/anniversary|wedding/i, '💍'],
  [/video|film|movie|youtube|netflix/i, '🎬'],
  [/cricket/i, '🏏'], [/soccer|football/i, '⚽'], [/rugby/i, '🏉'], [/tennis|padel/i, '🎾'], [/netball|basketball/i, '🏀'],
  [/swim/i, '🏊'], [/\brun\b|marathon|parkrun|colou?r run/i, '🏃'], [/gym|workout|pilates|yoga|sport/i, '🏋️'],
  [/flight|airport|\bfly\b/i, '✈️'], [/vaccin|injection|\bjab\b/i, '💉'], [/braces|orthodont|dentist/i, '🦷'],
  [/doctor|\bdr\b|clinic|eye doc|physio|hospital|check-?up/i, '🩺'],
  [/hair|salon|barber|nails|botox|wax|brow|facial/i, '💇'], [/vet\b|groom|\bdog\b|\bcat\b|puppy/i, '🐾'],
  [/kumon|school|class|teacher|term|homework|lesson|exam|tutor|university|lecture/i, '🎓'],
  [/music|symphony|concert|piano|guitar|choir|band/i, '🎵'], [/holiday|vacation|\btrip\b|away with|getaway/i, '🧳'],
  [/pack/i, '🎒'], [/dinner|lunch|breakfast|brunch|cook|restaurant|braai/i, '🍽️'], [/coffee|tea\b/i, '☕'],
  [/party|social|celebrat|dreamland|show/i, '🎉'], [/meeting|\bcall\b|zoom|teams|interview|pta|committee/i, '👥'],
  [/e-?mail|mail\b/i, '✉️'], [/towel|bedding|laundry|wash|clean|chores?/i, '🧺'], [/plant|garden/i, '🪴'],
  [/\bpay\b|bill|invoice|budget|tax|bank|insurance/i, '💳'], [/shop|market|groceries|buy/i, '🛍️'],
  [/design|\bart\b|draw|paint|craft/i, '🎨'], [/au ?pair/i, '🍭'], [/babysit|nanny|kids|pick ?up|drop ?off|\bdrop\b/i, '🧒'],
  [/\bcar\b|service|licen[cs]e|tyres?/i, '🚗'], [/church|mass\b|sermon/i, '⛪'], [/read/i, '📖'],
  [/\bai\b|claude|computer|code|laptop/i, '💻'], [/club/i, '🎟️'], [/medic|pill|pharmac/i, '💊'], [/gift|present/i, '🎁'],
];
const EMOJIS = ['📅', '⭐', '❤️', '🎂', '🎉', '🎁', '💍', '👥', '📞', '✉️', '💻', '📋', '🎓', '📚', '📖', '🎨', '🎵', '🎬',
  '🏃', '🏋️', '⚽', '🏏', '🏉', '🎾', '🏊', '🩺', '💉', '🦷', '💊', '💇', '🍽️', '☕', '🛍️', '💳', '🏠', '🧺', '🪴', '🐾',
  '🚗', '✈️', '🧳', '🎒', '🧒', '🍭', '👶', '⛪', '✅', '⏰', '🌙', '☀️', '🏖️'];
const CAL_EMOJIS = ['', '📅', '💼', '🏠', '👪', '🧒', '🍭', '🎓', '🎂', '✅', '⚽', '❤️', '⭐', '🎵', '🩺', '✈️', '🏖️', '💻', '🛒', '🐾'];
const COLORS = ['#e8467c', '#d6336c', '#f76707', '#f59f00', '#fab005', '#40c057', '#12b886', '#15aabf', '#228be6', '#4c6ef5', '#7950f2', '#be4bdb', '#868e96', '#a0522d'];
const BGS = { blossom: 'Blossom', ocean: 'Ocean', aurora: 'Aurora', sunset: 'Sunset', graphite: 'Graphite', starry: 'Starry' };
const BG_PREVIEW = {
  blossom: 'linear-gradient(135deg,#ff9ac1,#c79bff,#ffc58f)', ocean: 'linear-gradient(135deg,#7cc4fa,#5fe0c9,#9aa5ff)',
  aurora: 'linear-gradient(135deg,#6be8b8,#a98bff,#6fd7fa)', sunset: 'linear-gradient(135deg,#ff9a6b,#ff7aa0,#ffd27a)',
  graphite: 'linear-gradient(135deg,#c6ccd8,#9aa4b8,#dfe3ea)',
  starry: 'url(bg-starry.webp) center/cover',
};

function eventIcon(ev, cal) {
  if (ev.icon === 'none') return ''; // chosen "None": no icon at all
  if (ev.icon) return ev.icon;
  if (S.settings.smartIcons) for (const [re, e] of SMART) if (re.test(ev.title)) return e;
  return cal?.icon || '';
}

/* ---------------- state ---------------- */
function seed() {
  const g1 = 'g-me';
  const cals = [
    { id: 'c-main', name: 'Calendar', color: '#e8467c', icon: '📅', groupId: g1, visible: true },
    { id: 'c-aupair', name: 'Au Pair', color: '#7950f2', icon: '🍭', groupId: g1, visible: true },
    { id: 'c-kumon', name: 'KUMON', color: '#15aabf', icon: '🎓', groupId: g1, visible: true },
    { id: 'c-bday', name: 'Birthdays', color: '#f59f00', icon: '🎂', groupId: g1, visible: true },
    { id: 'c-todo', name: 'To Do', color: '#f76707', icon: '✅', groupId: g1, visible: true },
    { id: 'za', name: 'South Africa holidays', color: '#868e96', icon: '🏖️', groupId: g1, visible: true, readonly: true },
  ];
  const E = (calId, title, start, end, extra = {}) => ({
    id: uid(), calId, title, start, end, allDay: false, repeat: 'none', until: '', days: [], exdates: [], doneDates: [],
    reminder: null, location: '', notes: '', task: false, icon: '', ...extra,
  });
  const A = (calId, title, day, extra = {}) => E(calId, title, day + 'T00:00', day + 'T00:00', { allDay: true, ...extra });
  const events = [
    E('c-main', 'Test Tuesday Video', '2026-08-31T19:00', '2026-08-31T20:00', { repeat: 'weekly' }),
    E('c-kumon', '# Myth Design', '2026-09-02T10:00', '2026-09-02T11:00', { repeat: 'weekly' }),
    E('c-kumon', 'Kumon Instruct', '2026-09-05T09:30', '2026-09-05T10:30', { repeat: 'weekly', reminder: 30 }),
    A('c-todo', 'Change Towel', '2026-09-06', { repeat: 'weekly', task: true }),
    E('c-aupair', 'Afternoon Au Pair', '2026-09-14T14:00', '2026-09-14T17:00', { repeat: 'weekdays' }),
    E('c-kumon', '100 Club', '2026-09-28T09:30', '2026-09-28T10:15', { repeat: 'weekly' }),
    E('c-kumon', 'SOM & POM', '2026-09-28T10:30', '2026-09-28T11:30', { repeat: 'weekly' }),
    E('c-main', 'Daffodil Groom', '2026-09-28T10:00', '2026-09-28T11:00'),
    A('c-bday', 'ABBA Birthday', '2026-09-05', { repeat: 'yearly' }),
    A('c-bday', 'Nikki Birthday', '2026-09-16', { repeat: 'yearly' }),
    A('c-bday', 'Aunty Bernie Birthday', '2026-09-29', { repeat: 'yearly' }),
    A('c-bday', 'Nikson Birthday', '2026-09-29', { repeat: 'yearly' }),
    A('c-bday', 'Pieter Birthday', '2026-09-29', { repeat: 'yearly' }),
    E('c-aupair', 'Matthew Cricket Clinic', '2026-09-29T09:00', '2026-09-29T12:00', { location: 'In-person' }),
    E('c-aupair', 'Matthew Cricket Clinic', '2026-09-30T09:00', '2026-09-30T12:00', { location: 'In-person' }),
    E('c-aupair', 'Read Matthew Classroom', '2026-09-29T11:30', '2026-09-29T12:00', { repeat: 'weekly', days: [2, 5] }),
    A('c-todo', 'Pack bags', '2026-09-30', { task: true }),
    E('c-main', 'Daffodil Vaccination', '2026-09-30T09:30', '2026-09-30T10:00', { reminder: 60 }),
    E('c-main', 'Wax face and Arms', '2026-09-30T19:00', '2026-09-30T20:00'),
    E('c-main', 'Babysit Thomas', '2026-10-01T18:00', '2026-10-01T21:00'),
    E('c-todo', 'Change Bedding', '2026-10-02T07:00', '2026-10-02T07:30', { task: true }),
    E('c-main', 'Hair Botox', '2026-10-02T10:00', '2026-10-02T11:30'),
    E('c-main', 'Josh drop at Airport', '2026-10-02T16:00', '2026-10-02T17:00'),
    E('c-main', 'Cassandra in PTA', '2026-10-02T16:30', '2026-10-04T20:00'),
    E('c-aupair', 'Kids Weekly Schedule', '2026-09-27T12:00', '2026-09-27T12:30', { repeat: 'weekly' }),
  ];
  return {
    groups: [{ id: g1, name: 'All calendars' }],
    cals, events,
    settings: { weekStart: 1, theme: 'dark', bg: 'blossom', look: 'planner', nightDefault: true, h24: false, time12: true, oneGroup: true, lolly: true, readTueFri: true, weekNums: true, smartIcons: true, defaultReminder: 15, defaultCal: 'c-main', workStart: 8, workEnd: 17, view: 'month' },
  };
}
function load() {
  try { const s = JSON.parse(localStorage.getItem(STORE_KEY)); if (s && s.cals) return s; } catch { }
  return null;
}
const S = load() || seed();
S.settings ||= {};
// One-time updates for calendars saved by older versions. These must run before the defaults
// below are filled in, otherwise the "already done" flags would look set.
// 1) A single "All calendars" group that also holds the holidays.
if (!S.settings.oneGroup) {
  const main = S.groups.find(g => g.id === 'g-me'), hol = S.groups.find(g => g.id === 'g-hol');
  if (main) {
    if (main.name === 'My calendars') main.name = 'All calendars';
    if (hol) { S.cals.forEach(c => { if (c.groupId === hol.id) c.groupId = main.id; }); S.groups = S.groups.filter(g => g !== hol); }
  }
  S.settings.oneGroup = true;
}
// 5) Once Google is connected, remove the sample calendars from the first version (a backup is kept first).
const SAMPLE_CALS = ['c-main', 'c-aupair', 'c-kumon', 'c-bday', 'c-todo'];
function removeSamples() {
  if (S.settings.samplesRemoved || !(S.cals || []).some(c => c.source?.type === 'google') || !S.cals.some(c => SAMPLE_CALS.includes(c.id))) return null;
  try { localStorage.setItem('glasscal.backup-before-removing-samples', JSON.stringify(S)); } catch { }
  const n = S.events.filter(e => SAMPLE_CALS.includes(e.calId)).length;
  S.cals = S.cals.filter(c => !SAMPLE_CALS.includes(c.id));
  S.events = S.events.filter(e => !SAMPLE_CALS.includes(e.calId));
  // New events now go to your main Google calendar.
  const primary = S.cals.find(c => c.source?.type === 'google' && /@/.test(c.source.remoteId) && !c.readonly) || S.cals.find(c => c.source?.type === 'google' && !c.readonly);
  if (primary) S.settings.defaultCal = primary.id;
  S.settings.samplesRemoved = true;
  return n;
}
{ const n = removeSamples(); if (n != null) window.__samplesRemovedNote = n; }
// 8) SOM & POM and 100 Club are monthly all-day tasks (tick circle).
if (!S.settings.monthlyTasks) {
  (S.events || []).forEach(e => { if (/^(som\s*&\s*pom|100 club)$/i.test((e.title || '').trim())) e.task = true; });
  S.settings.monthlyTasks = true;
}
// 7) Tidy "[In-person]" out of names already saved in this calendar (synced ones are tidied on every sync).
(S.events || []).forEach(e => tidyEvent(e));
// 6) 100 Club now lives in Google (To Do, all day on the 28th of every month): drop the old weekly sample copy.
if (!S.settings.club100) {
  S.events = (S.events || []).filter(e => !(/^100 club$/i.test((e.title || '').trim()) && !S.cals.find(c => c.id === e.calId)?.source));
  S.settings.club100 = true;
}
// 4) "Read Matthew Classroom" repeats every Tuesday and Friday.
if (!S.settings.readTueFri) {
  const master = (S.events || []).find(e => /^read matthew classroom$/i.test(e.title.trim()) && isRecurring(e));
  if (master) { master.repeat = 'weekly'; master.days = [2, 5]; }
  S.settings.readTueFri = true;
}
// 3) Au Pair uses the lollipop icon.
if (!S.settings.lolly) { S.cals.forEach(c => { if (/^au ?pair$/i.test(c.name.trim()) && c.icon === '🧒') c.icon = '🍭'; }); S.settings.lolly = true; }
// 2) 2:00pm-style times (Settings can turn 24-hour back on afterwards).
if (!S.settings.time12) { S.settings.h24 = false; S.settings.time12 = true; }
// 10) The default look is Blossom in night mode (set once for every device; you can still change it in Settings and it stays).
if (!S.settings.nightDefault) { S.settings.theme = 'dark'; S.settings.bg = 'blossom'; }
S.settings = { ...seed().settings, ...S.settings };
ensureFun();
save();
function save() {
  window.CalCloud?.changed();
  try { localStorage.setItem(STORE_KEY, JSON.stringify(S)); }
  catch { toast('Could not save — browser storage is full or blocked', { icon: 'x' }); }
}

let cur = startOfDay(new Date());
let miniCur = new Date(cur);
let view = S.settings.view;
let searchQ = '';
let gridScroll = null;
const occMap = new Map();

/* undo */
const undoStack = [];
function pushUndo() {
  undoStack.push(JSON.stringify({ events: S.events, cals: S.cals, groups: S.groups }));
  if (undoStack.length > 40) undoStack.shift();
}
/* Edits to Google/Outlook calendars are sent by sync.js. Undo can't reach those, so their toasts skip the Undo button. */
const isSynced = calId => typeof Sync !== 'undefined' && Sync.isRemoteCal(S.cals.find(c => c.id === calId));
/* Only these are stored in Google/Outlook. Icons and task tick-boxes live only in this calendar. */
const REMOTE_FIELDS = ['title', 'start', 'end', 'allDay', 'location', 'notes', 'reminder', 'calId', 'repeat', 'until'];
const guestKey = e => JSON.stringify((e.attendees || []).map(a => a.email));
const remoteChanged = (a, b) => !a || !b || !!b._notify || guestKey(a) !== guestKey(b) || JSON.stringify(a.days || []) !== JSON.stringify(b.days || [])
  || REMOTE_FIELDS.some(k => (a[k] ?? '') !== (b[k] ?? ''));
function pushRemote(before, after) {
  if (!((before && isSynced(before.calId)) || (after && isSynced(after.calId)))) return;
  if (before && after && !remoteChanged(before, after)) return; // nothing Google/Outlook stores has changed
  Sync.pushSave(before, after);
}
function undo() {
  const s = undoStack.pop();
  if (!s) return toast('Nothing to undo', { icon: 'repeat' });
  Object.assign(S, JSON.parse(s)); save(); render(); toast('Undone', { icon: 'repeat' });
}

/* ---------------- holidays (South Africa) ---------------- */
function easter(y) {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25),
    g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4,
    l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
    mo = Math.floor((h + l - 7 * m + 114) / 31), da = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(y, mo - 1, da);
}
const holCache = {};
function zaHolidays(y) {
  if (holCache[y]) return holCache[y];
  const E = easter(y);
  const list = [[new Date(y, 0, 1), "New Year's Day"], [new Date(y, 2, 21), 'Human Rights Day'], [addDays(E, -2), 'Good Friday'],
  [addDays(E, 1), 'Family Day'], [new Date(y, 3, 27), 'Freedom Day'], [new Date(y, 4, 1), "Workers' Day"],
  [new Date(y, 5, 16), 'Youth Day'], [new Date(y, 7, 9), "National Women's Day"], [new Date(y, 8, 24), 'Heritage Day'],
  [new Date(y, 11, 16), 'Day of Reconciliation'], [new Date(y, 11, 25), 'Christmas Day'], [new Date(y, 11, 26), 'Day of Goodwill']];
  const taken = new Set(list.map(([d]) => ymd(d)));
  const out = [];
  for (const [d, n] of list) {
    out.push([d, n]);
    if (d.getDay() === 0) { let o = addDays(d, 1); while (taken.has(ymd(o))) o = addDays(o, 1); taken.add(ymd(o)); out.push([o, n + ' (observed)']); }
  }
  return holCache[y] = out.map(([d, n]) => ({
    id: 'za-' + ymd(d), calId: 'za', title: 'Public Holiday: ' + n, start: ymd(d) + 'T00:00', end: ymd(d) + 'T00:00', allDay: true, repeat: 'none', icon: '🏖️', readonly: true, holiday: true,
  }));
}

/* ---------------- fun holidays (Halloween, Easter, Christmas, teacher days…) ---------------- */
const funCache = {};
function funHolidays(y) {
  if (funCache[y]) return funCache[y];
  const E = easter(y);
  // nth (1-5) weekday (0=Sun) of a month, or the last one with n = -1
  const nth = (m, wd, n) => {
    if (n > 0) { const f = new Date(y, m, 1); return new Date(y, m, 1 + ((wd - f.getDay() + 7) % 7) + (n - 1) * 7); }
    const l = new Date(y, m + 1, 0); return new Date(y, m, l.getDate() - ((l.getDay() - wd + 7) % 7));
  };
  const list = [
    [new Date(y, 1, 14), "Valentine's Day", '💝'], [addDays(E, -47), 'Shrove Tuesday (Pancake Day)', '🥞'],
    [addDays(E, -46), 'Ash Wednesday', '✝️'], [new Date(y, 2, 8), "International Women's Day", '💪'], [new Date(y, 2, 14), 'Pi Day', '🥧'],
    [new Date(y, 2, 17), "St Patrick's Day", '🍀'], [new Date(y, 2, 21), 'World Poetry Day', '📝'], [new Date(y, 3, 1), "April Fools' Day", '🤡'],
    [addDays(E, -7), 'Palm Sunday', '🌿'], [addDays(E, -3), 'Maundy Thursday', '🍞'], [addDays(E, -1), 'Holy Saturday', '🥚'], [E, 'Easter Sunday', '🐰'],
    [new Date(y, 3, 22), 'Earth Day', '🌍'], [new Date(y, 3, 23), 'World Book Day', '📚'], [new Date(y, 4, 4), 'Star Wars Day', '🚀'],
    [nth(4, 0, 2), "Mother's Day", '💐'], [new Date(y, 4, 25), 'Africa Day', '🌍'], [new Date(y, 5, 5), 'World Environment Day', '🌱'],
    [nth(5, 0, 3), "Father's Day", '👔'], [new Date(y, 6, 18), 'Mandela Day', '🤝'], [new Date(y, 6, 30), 'Friendship Day', '🫶'],
    [new Date(y, 8, 8), 'International Literacy Day', '📖'], [new Date(y, 8, 21), 'International Day of Peace', '🕊️'],
    [new Date(y, 9, 5), "World Teachers' Day", '🍎'],
    [new Date(y, 9, 31), 'Halloween', '🎃'], [new Date(y, 10, 5), 'Guy Fawkes Night', '🎆'], [new Date(y, 10, 20), "Universal Children's Day", '🧒'],
    [nth(10, 5, 4), 'Black Friday', '🛍️'], [new Date(y, 11, 24), 'Christmas Eve', '🎄'], [new Date(y, 11, 31), "New Year's Eve", '🥳'],
  ];
  return funCache[y] = list.map(([d, n, ic]) => ({
    id: 'fun-' + ymd(d) + '-' + n.length, calId: 'fun', title: n, start: ymd(d) + 'T00:00', end: ymd(d) + 'T00:00', allDay: true, repeat: 'none', icon: ic, readonly: true, fun: true,
  }));
}
/* The "Fun holidays" folder with its calendar (added once for calendars saved by older versions). */
function ensureFun() {
  if (S.settings.holidaysInGoogle) return; // they live in Google now
  if (!S.groups.some(g => g.id === 'g-fun')) S.groups.push({ id: 'g-fun', name: 'Fun holidays' });
  if (!S.cals.some(c => c.id === 'fun')) S.cals.push({ id: 'fun', name: 'Fun holidays', color: '#868e96', icon: '🎉', groupId: 'g-fun', visible: true, readonly: true });
  const f = S.cals.find(c => c.id === 'fun'); if (f) f.color = '#868e96'; // same colour as the public holidays
}
window.ensureFun = ensureFun;

/* ---------------- recurrence ---------------- */
const weekStartMon = d => addDays(startOfDay(d), -((d.getDay() + 6) % 7));
function occurrences(ev, from, to) {
  const s0 = parse(ev.start);
  let e0 = parse(ev.end); if (ev.allDay) e0 = addDays(e0, 1);
  const dur = Math.max(0, e0 - s0);
  const out = [], ex = ev.exdates || [];
  const until = ev.until ? addDays(parse(ev.until), 1) : null;
  const emit = s => {
    const e = new Date(+s + dur);
    if (s < to && (e > from || (+e === +s && s >= from)) && !ex.includes(ymd(s))) out.push({ start: s, end: e });
  };
  const r = ev.repeat || 'none';
  if (r === 'none') { emit(s0); return out; }
  const limit = until && until < to ? until : to;
  const at = (y, m, d) => new Date(y, m, d, s0.getHours(), s0.getMinutes());

  if (r === 'monthly' || r === 'yearly') {
    const step = r === 'monthly' ? 1 : 12;
    let i = (from.getFullYear() - s0.getFullYear()) * 12 + from.getMonth() - s0.getMonth() - 1 - Math.ceil(dur / (28 * 864e5));
    i = Math.max(0, Math.floor(i / step) * step);
    for (let g = 0; g < 600; g++, i += step) {
      const s = at(s0.getFullYear(), s0.getMonth() + i, s0.getDate());
      if (s >= limit) break;
      if (s.getDate() === s0.getDate()) emit(s);
    }
    return out;
  }
  let days = null, step = 1;
  if (r === 'weekdays') days = [1, 2, 3, 4, 5];
  else if (r === 'weekly' || r === 'biweekly') { if (ev.days?.length) days = ev.days; else step = r === 'weekly' ? 7 : 14; }
  const w0 = weekStartMon(s0);
  let i = Math.max(0, Math.floor((from - s0 - dur) / 864e5) - 1);
  if (!days) i = Math.floor(i / step) * step;
  for (let g = 0; g < 5000; g++, i += days ? 1 : step) {
    const s = at(s0.getFullYear(), s0.getMonth(), s0.getDate() + i);
    if (s >= limit) break;
    if (days) {
      if (!days.includes(s.getDay())) continue;
      if (r === 'biweekly' && Math.round((weekStartMon(s) - w0) / 6048e5) % 2) continue;
    }
    emit(s);
  }
  return out;
}

function matches(ev, q) {
  return `${ev.title} ${ev.location || ''} ${ev.notes || ''}`.toLowerCase().includes(q);
}
function getOcc(from, to, { all = false } = {}) {
  const vis = new Map(S.cals.filter(c => c.visible).map(c => [c.id, c]));
  let evs = S.events.filter(e => vis.has(e.calId));
  if (vis.has('za')) for (let y = from.getFullYear(); y <= to.getFullYear(); y++) evs = evs.concat(zaHolidays(y));
  if (vis.has('fun')) for (let y = from.getFullYear(); y <= to.getFullYear(); y++) evs = evs.concat(funHolidays(y));
  const q = all ? '' : searchQ.trim().toLowerCase();
  const out = [];
  for (const ev of evs) {
    if (q && !matches(ev, q)) continue;
    const cal = vis.get(ev.calId);
    for (const { start, end } of occurrences(ev, from, to)) {
      const o = { ev, cal, start, end, ro: !!(ev.readonly || cal.readonly), key: ev.id + '|' + ymd(start) };
      occMap.set(o.key, o); out.push(o);
    }
  }
  return out.sort((a, b) => (b.ev.allDay - a.ev.allDay) || (a.start - b.start) || ((b.end - b.start) - (a.end - a.start)));
}
const isDone = o => (o.ev.doneDates || []).includes(ymd(o.start));
const multiDay = o => o.ev.allDay || (o.end - o.start >= 864e5);

/* ---------------- render: shared bits ---------------- */
function taskBox(o) {
  if (!o.ev.task) return '';
  const d = isDone(o);
  return `<span class="tick ${d ? 'on' : ''}" role="checkbox" aria-checked="${d}" title="Mark done">${d ? svg('check') : ''}</span>`;
}
function chip(o, day) {
  const ic = eventIcon(o.ev, o.cal);
  const showTime = !o.ev.allDay && (!day || sameDayStr(o.start, day));
  return `<div class="chip ${o.ev.allDay || multiDay(o) ? 'allday' : ''} ${isDone(o) ? 'done' : ''} ${o.ro ? 'ro' : ''}" style="--c:${o.cal.color}" data-key="${esc(o.key)}" draggable="${!o.ro}" title="${esc(o.ev.title)}">
    ${showTime ? `<span class="t">${fmtTime(o.start)}:</span>` : ''}${ic ? `<span class="ic">${ic}</span>` : ''}<span class="ti">${esc(o.ev.title)}</span>${taskBox(o)}${pplIcon(o)}${isRecurring(o.ev) ? svg('repeat', 'rp') : ''}</div>`;
}
const sameDayStr = (d, k) => ymd(d) === k;
/* Small 'people' icon on events that have guests. */
const pplTag = ev => { const n = ev.attendees?.length || 0; return n ? `<span class="ppl-tag">${svg('users')}${n} ${n === 1 ? 'person' : 'people'}</span>` : ''; };
const pplIcon = o => o.ev.attendees?.length ? `<span class="ppl-ic" title="${o.ev.attendees.length} ${o.ev.attendees.length === 1 ? 'person' : 'people'}">${svg('users')}</span>` : '';
/* Compact time for planner stickers: "10am", "4:30pm" (or 24-hour). */
const tShort = d => S.settings.h24 ? hm(d) : `${d.getHours() % 12 || 12}${d.getMinutes() ? ':' + pad(d.getMinutes()) : ''}${d.getHours() < 12 ? 'am' : 'pm'}`;
/* Planner-look month entries: all-day notes become labels, timed events show icon + time + name. */
function sticker(o, day) {
  const ic = eventIcon(o.ev, o.cal);
  const common = `style="--c:${o.cal.color}" data-key="${esc(o.key)}" draggable="${!o.ro}" title="${esc(o.ev.title)}"`;
  const state = `${isDone(o) ? 'done' : ''} ${o.ro ? 'ro' : ''} ${o.end <= new Date() ? 'pastev' : ''}`;
  if (o.ev.allDay || multiDay(o)) {
    // No time: the tick box sits right after the icon, before the name.
    return `<div class="chip tag ${state}" ${common}>${taskBox(o)}${ic ? `<span class="ic">${ic}</span>` : ''}<span class="st-tx"><span class="ti">${esc(o.ev.title)}</span>${pplIcon(o)}</span></div>`;
  }
  const time = sameDayStr(o.start, day) ? `${fmtTime(o.start)} - ${fmtTime(o.end)}` : `until ${fmtTime(o.end)}`;
  // The tick box is part of the (never-wrapping) time, so it always sits between the time and the name.
  return `<div class="chip st ${state}" ${common}><span class="st-tx"><span class="t ${o.end <= new Date() ? 'tpast' : ''}"><span class="tt">${time}:</span>${taskBox(o)}</span> ${ic ? `<span class="ic">${ic}</span>` : ''}<span class="ti">${esc(o.ev.title)}</span>${pplIcon(o)}</span></div>`;
}
const plannerLook = () => S.settings.look === 'planner';
/* Birthdays: from a calendar called "Birthdays…" or with "birthday" in the title. Month view shows them as a 🧁. */
/* Holiday labels sit next to the date number in month view: public holidays (built-in or a subscribed
   "Holidays…" calendar) and all-day events called "School Holiday" (kept in your Au Pair calendar). */
const isSchoolHoliday = o => /^school holidays?$/i.test((o.ev.title || '').trim()) && (o.ev.allDay || o.end - o.start >= 864e5);
const isHoliday = o => !!o.ev.holiday || !!o.ev.fun || (/holiday/i.test(o.cal.name) && (!!o.cal.readonly || isOwnHolCal(o.cal))) || isSchoolHoliday(o);
const isOwnHolCal = c => /^(south africa|fun) holidays$/i.test((c.name || '').trim()); // the two holiday calendars kept in Google
const HOLIDAY_GREY = '#868e96';
/* A real day off: a built-in public holiday, or a holiday-calendar event marked "Public holiday" (not observances). */
const isDayOff = o => !!o.ev.holiday || (isHoliday(o) && /public holiday/i.test(`${o.ev.notes || ''} ${o.ev.title}`));
const isBirthday = o => /birthday|b-?day|verjaar/i.test(o.ev.title) || /birthday/i.test(o.cal.name);

function titleText() {
  if (searchQ && view === 'agenda') return `Search results`;
  if (view === 'month') return fmt(cur, { month: 'long', year: 'numeric' });
  if (view === 'year') return String(cur.getFullYear());
  if (view === 'day') return fmt(cur, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  if (view === 'agenda') return 'From ' + fmt(cur, { day: 'numeric', month: 'long', year: 'numeric' });
  const [a, n] = gridRange(); const b = addDays(a, n - 1);
  return a.getMonth() === b.getMonth()
    ? `${a.getDate()} – ${fmt(b, { day: 'numeric', month: 'long', year: 'numeric' })}`
    : `${fmt(a, { day: 'numeric', month: 'short' })} – ${fmt(b, { day: 'numeric', month: 'short', year: 'numeric' })}`;
}
function gridRange() {
  if (view === 'day') return [startOfDay(cur), 1];
  if (view === 'workweek') { const w = weekStartMon(cur); return [w, 5]; }
  return [startOfWeek(cur), 7];
}

/* Back button: every time the visible page (view + period) changes, remember where you were. */
const navHistory = [];
let lastNav = null, goingBack = false;
function trackNav() {
  const nav = { view, cur: ymd(cur), key: view + '|' + titleText() };
  if (lastNav && !goingBack && lastNav.key !== nav.key) {
    navHistory.push(lastNav);
    if (navHistory.length > 50) navHistory.shift();
  }
  lastNav = nav; goingBack = false;
  const b = $('#backBtn');
  b.disabled = !navHistory.length;
  if (navHistory.length) {
    const p = navHistory.at(-1);
    b.title = `Back to ${p.key.split('|')[1]} (${p.view}) — Backspace`;
  }
}
function goBack() {
  const p = navHistory.pop();
  if (!p) return;
  goingBack = true;
  view = p.view; S.settings.view = view; save();
  cur = parse(p.cur); miniCur = new Date(cur); gridScroll = null;
  render();
}

function render() {
  occMap.clear();
  applyTheme();
  trackNav();
  $$('#viewSeg button').forEach(b => b.classList.toggle('on', b.dataset.view === view));
  const v = $('#view');
  const banner = searchQ && view !== 'agenda'
    ? `<div class="search-banner">${svg('search')}Showing only events matching “${esc(searchQ)}”<button data-act="listResults">See all results</button><button data-act="clearSearch" style="margin-left:8px">Clear</button></div>` : '';
  let html = '';
  if (view === 'month') html = renderMonth();
  else if (view === 'year') html = renderYear();
  else if (view === 'agenda') html = renderAgenda();
  else if (view === 'week' || view === 'workweek' || view === 'day') html = renderWeekCols(); // day = one big box
  else html = renderGrid();
  // Big title with the previous / next arrows either side, at the top of every view.
  const head = `<div class="vhead"><button class="nav-arrow" data-nav="-1" title="Previous (←)">${svg('chevronLeft')}</button>`
    + `<h2 class="mtitle">${esc(titleText())}</h2><button class="nav-arrow" data-nav="1" title="Next (→)">${svg('chevronRight')}</button></div>`;
  v.innerHTML = head + banner + html;
  if (view === 'month' || view === 'week' || view === 'workweek' || view === 'day') fitMonth();
  renderSidebar();
  renderUpNext();
  updateBrand();
}

/* ---------------- month ---------------- */
function renderMonth() {
  const first = new Date(cur.getFullYear(), cur.getMonth(), 1);
  const gs = startOfWeek(first);
  const last = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);
  const rows = Math.ceil((dayDiff(last, gs) + 1) / 7);
  const ge = addDays(gs, rows * 7);
  const occ = getOcc(gs, ge);
  const byDay = {};
  for (const o of occ) {
    let d = startOfDay(o.start);
    do { if (d >= gs && d < ge) (byDay[ymd(d)] ||= []).push(o); d = addDays(d, 1); } while (d < o.end);
  }
  const today = ymd(new Date());
  const planner = plannerLook();
  const draw = planner ? sticker : chip;
  let h = '';
  h += `<div class="month ${planner ? 'planner' : ''} ${S.settings.weekNums ? '' : 'nowk'}" style="--rows:${rows}"><div class="wdc"></div>`;
  for (let i = 0; i < 7; i++) h += `<div class="wd"><span class="l">${fmt(addDays(gs, i), { weekday: 'long' })}</span><span class="s">${fmt(addDays(gs, i), { weekday: 'short' })}</span></div>`;
  for (let r = 0; r < rows; r++) {
    const ws = addDays(gs, r * 7);
    h += `<div class="wk" title="Week ${isoWeek(ws)}">${isoWeek(addDays(ws, 3))}</div>`;
    for (let c = 0; c < 7; c++) {
      // Days from the previous/next month are left out and run together as ONE plain block.
      if (addDays(ws, c).getMonth() !== cur.getMonth()) {
        let n = 1; while (c + n < 7 && addDays(ws, c + n).getMonth() !== cur.getMonth()) n++;
        h += `<div class="cell blank" style="grid-column: span ${n}"></div>`;
        c += n - 1; continue;
      }
      const d = addDays(ws, c);
      h += dayCell(d, byDay[ymd(d)] || [], today, draw);
    }
  }
  return h + '</div>';
}
/* One day box, used by the month view and the week view so they always look the same. */
function dayCell(d, list, today, draw) {
  const k = ymd(d);
  const bdays = list.filter(isBirthday);
  const hols = list.filter(o => !isBirthday(o) && isHoliday(o));
  // timed events first (in time order), all-day ones underneath
  const rest = list.filter(o => !isBirthday(o) && !isHoliday(o)).sort((a, b) => (multiDay(a) - multiDay(b)) || (a.start - b.start));
  const holLbl = hols.map(o => `<button class="hol-lbl" data-key="${esc(o.key)}" style="--c:${isSchoolHoliday(o) || isOwnHolCal(o.cal) ? HOLIDAY_GREY : o.cal.color}" title="${esc(o.ev.title)}">${esc(o.ev.title)}</button>`).join('');
  const cake = bdays.length ? `<button class="bday-btn" data-bday="${k}" title="${esc(bdays.map(o => o.ev.title).join(', '))}">🧁${bdays.length > 1 ? `<b>${bdays.length}</b>` : ''}</button>` : '';
  return `<div class="cell ${k === today ? 'today' : ''} ${k < today ? 'past' : ''} ${[0, 6].includes(d.getDay()) || hols.some(isDayOff) ? 'we' : ''} ${cake ? 'has-cake' : ''}" data-day="${k}">
    <div class="ch"><button class="dnum" data-goto="${k}" title="Open day">${d.getDate()}</button>${cake}${holLbl}
    <button class="cell-add" title="New event">${svg('plus')}</button></div>
    <div class="cl">${rest.map(o => draw(o, k)).join('')}</div></div>`;
}
/* ---------------- week: seven (or five) month-style day boxes ---------------- */
function renderWeekCols() {
  const [start, n] = gridRange();
  const end = addDays(start, n);
  const byDay = {};
  for (const o of getOcc(start, end)) {
    let d = startOfDay(o.start);
    do { if (d >= start && d < end) (byDay[ymd(d)] ||= []).push(o); d = addDays(d, 1); } while (d < o.end);
  }
  const today = ymd(new Date()), planner = plannerLook(), draw = planner ? sticker : chip;
  let h = `<div class="month weekcols ${planner ? 'planner' : ''} cols${n} ${S.settings.weekNums ? '' : 'nowk'}" style="--rows:1"><div class="wdc"></div>`;
  for (let i = 0; i < n; i++) h += `<div class="wd"><span class="l">${fmt(addDays(start, i), { weekday: 'long' })}</span><span class="s">${fmt(addDays(start, i), { weekday: 'short' })}</span></div>`;
  h += `<div class="wk" title="Week ${isoWeek(start)}">${isoWeek(addDays(start, 3))}</div>`;
  for (let i = 0; i < n; i++) { const d = addDays(start, i); h += dayCell(d, byDay[ymd(d)] || [], today, draw); }
  return h + '</div>';
}
/* Hide chips that don't fit and replace them with "+N more". */
function fitMonth() {
  $$('.month .cell').forEach(cell => {
    const cl = $('.cl', cell); if (!cl) return;
    const chips = [...cl.children];
    if (!chips.length) return;
    cl.style.cssText = 'display:flex;flex-direction:column;min-height:0;flex:1;overflow:hidden';
    if (plannerLook()) {
      // Planner look: shrink the text until every event fits in the day box.
      let fs = cell.closest('.cols1') ? 19 : cell.closest('.weekcols') ? 15 : 13; // the day box is the biggest, then week boxes
      cl.style.fontSize = fs + 'px';
      const overflows = () => cl.scrollHeight > cl.clientHeight + 1 || cl.scrollWidth > cl.clientWidth + 1;
      while (overflows() && fs > 7) { fs -= 0.5; cl.style.fontSize = fs + 'px'; }
      if (!overflows()) return;
    }
    cl.style.gap = '2px';
    // Entries have different heights (labels vs stickers), so add them up until the cell is full.
    const avail = cl.clientHeight + 2;
    let used = 0, fit = 0;
    for (const c of chips) { used += c.offsetHeight + 2; if (used > avail) break; fit++; }
    fit = Math.max(1, fit);
    if (chips.length <= fit) return;
    // "+N more" sits beside the date number, so every row of the cell can hold an event.
    chips.slice(fit).forEach(c => c.remove());
    $('.cell-add', cell).insertAdjacentHTML('beforebegin', `<button class="more-btn" data-day="${cell.dataset.day}" title="Show all">+${chips.length - fit}</button>`);
  });
}

/* ---------------- day / week grid ---------------- */
function layout(items) {
  items.sort((a, b) => a.s - b.s || b.e - a.e);
  let cluster = [], cols = [], end = -1;
  const flush = () => { cluster.forEach(it => it.n = cols.length); cluster = []; cols = []; };
  for (const it of items) {
    if (it.s >= end && cluster.length) flush();
    let c = cols.findIndex(e => e <= it.s);
    if (c < 0) { c = cols.length; cols.push(it.e); } else cols[c] = it.e;
    it.col = c; cluster.push(it); end = Math.max(end, it.e);
  }
  flush();
}
function evBlock(it) {
  const o = it.o, w = 100 / it.n, h = (it.e - it.s) * HOUR / 60;
  const short = h < 38, ic = eventIcon(o.ev, o.cal);
  return `<div class="ev ${o.ev.attendees?.length ? 'has-ppl' : ''} ${o.ro ? 'ro' : ''} ${short ? 'short' : ''} ${isDone(o) ? 'done' : ''} ${o.end <= new Date() ? 'past' : ''}" data-key="${esc(o.key)}"
    style="--c:${o.cal.color};top:${it.s * HOUR / 60}px;height:${h - 2}px;left:calc(${it.col * w}% + 2px);width:calc(${w}% - 4px)" title="${esc(o.ev.title)}">
    <div class="ev-t">${ic ? `<span class="ic">${ic}</span>` : ''}<b>${esc(o.ev.title)}</b>${taskBox(o)}${short ? `<span class="ev-m">${fmtTime(o.start)}</span>` : ''}</div>
    ${short ? '' : `<div class="ev-m">${fmtTime(o.start)} – ${fmtTime(o.end)}${o.ev.location ? ' · ' + esc(o.ev.location) : ''}</div>`}
    ${pplIcon(o)}${o.ro ? '' : '<div class="rsz"></div>'}</div>`;
}
function renderGrid() {
  const [start, n] = gridRange();
  const days = [...Array(n)].map((_, i) => addDays(start, i));
  const occ = getOcc(start, addDays(start, n));
  const today = ymd(new Date());
  const now = new Date();
  const one = n === 1; // Day view: the big heading already shows the date
  let head = one ? '' : `<div class="gh"><div></div>${days.map(d => `<button class="dh ${ymd(d) === today ? 'today' : ''}" data-goto="${ymd(d)}"><span class="dw">${fmt(d, { weekday: 'short' })}</span><span class="dn">${d.getDate()}</span></button>`).join('')}</div>`;
  let ad = `<div class="gad"><div class="lbl">${S.settings.weekNums ? 'W' + isoWeek(start) : 'all-day'}</div>${days.map(d => {
    const k = ymd(d), de = addDays(d, 1);
    const all = occ.filter(o => multiDay(o) && o.start < de && o.end > d);
    // birthdays collapse into one cupcake button, like the month view
    const bd = all.filter(isBirthday);
    const cake = !bd.length ? '' : `<button class="bday-chip" data-bday="${k}" title="${esc(bd.map(o => o.ev.title).join(', '))}">🧁 ${bd.length === 1 ? esc(bd[0].ev.title) : `${bd.length} birthdays`}</button>`;
    return `<div class="adc" data-day="${k}">${all.filter(o => !isBirthday(o)).map(o => chip(o, k)).join('')}${cake}</div>`;
  }).join('')}</div>`;
  let hours = '';
  for (let h = 0; h < 24; h++) {
    if (h) hours += `<div class="hr" style="top:${h * HOUR}px">${fmtHour(h)}</div>`;
    hours += `<div class="hr half" style="top:${(h + 0.5) * HOUR}px">${S.settings.h24 ? `${pad(h)}:30` : `${h % 12 || 12}:30`}</div>`;
  }
  const ws = S.settings.workStart * HOUR, we = S.settings.workEnd * HOUR;
  const cols = days.map((d, i) => {
    const k = ymd(d), de = addDays(d, 1);
    const items = occ.filter(o => !multiDay(o) && ((o.start < de && o.end > d) || (+o.start === +o.end && ymd(o.start) === k)))
      .map(o => { const s = Math.max(0, (o.start - d) / 6e4), e = Math.min(1440, (o.end - d) / 6e4); return { o, s, e: Math.max(e, s + 22) }; });
    layout(items);
    const nowLine = k === today ? `<div class="now" style="top:${(now - d) / 6e4 * HOUR / 60}px"></div>` : '';
    const off = [0, 6].includes(d.getDay()) || occ.some(o => isDayOff(o) && o.start < de && o.end > d);
    return `<div class="col ${k === today ? 'today' : ''} ${off ? 'we' : ''}" data-day="${k}" data-idx="${i}">${items.map(evBlock).join('')}${nowLine}</div>`;
  }).join('');
  return `<div class="gridv ${one ? 'one' : ''}" style="--h:${HOUR}px;--n:${n};--ws:${ws}px;--we:${we}px">${head}${ad}
    <div class="gb" id="gb"><div class="gi"><div class="hours">${hours}</div>${cols}</div></div></div>`;
}

function bindGrid() {
  const gb = $('#gb');
  const nowH = new Date().getHours();
  // Open at the start of the working day (or where you last scrolled), never at midnight.
  const target = gridScroll || Math.max(0, (Math.min(nowH, S.settings.workStart + 1) - 1) * HOUR);
  gb.scrollTop = target;
  requestAnimationFrame(() => { if (Math.abs(gb.scrollTop - target) > 2) gb.scrollTop = target; });
  gb.addEventListener('scroll', () => { if (gb.clientHeight) gridScroll = gb.scrollTop; }, { passive: true });
  gb.addEventListener('pointerdown', onGridDown);
}
const snapMin = (col, y, floor) => {
  const m = (y - col.getBoundingClientRect().top) / HOUR * 60;
  return clamp((floor ? Math.floor(m / 15) : Math.round(m / 15)) * 15, 0, 1440);
};
function onGridDown(e) {
  if (e.button !== 0 || e.target.closest('.tick')) return;
  const col = e.target.closest('.col'); if (!col) return;
  const evEl = e.target.closest('.ev');
  e.preventDefault();
  const x0 = e.clientX, y0 = e.clientY;
  let moved = false;
  const off = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); };
  let mv, up;

  if (evEl) {
    const o = occMap.get(evEl.dataset.key);
    if (o.ro) { openPeek(o, evEl); return; }
    const mode = e.target.classList.contains('rsz') ? 'resize' : 'move';
    const top0 = parseFloat(evEl.style.top), h0 = parseFloat(evEl.style.height) + 2;
    const idx0 = +col.dataset.idx;
    let dMin = 0, dDay = 0;
    const meta = $('.ev-m', evEl);
    mv = ev => {
      if (!moved && Math.hypot(ev.clientX - x0, ev.clientY - y0) < 5) return;
      if (!moved) { moved = true; evEl.classList.add('dragging'); }
      dMin = Math.round((ev.clientY - y0) / HOUR * 60 / 15) * 15;
      if (mode === 'move') {
        const c = document.elementFromPoint(ev.clientX, ev.clientY)?.closest('.col');
        if (c && c !== evEl.parentElement) c.appendChild(evEl);
        dDay = +evEl.parentElement.dataset.idx - idx0;
        evEl.style.top = top0 + dMin * HOUR / 60 + 'px';
        evEl.style.left = '2px'; evEl.style.width = 'calc(100% - 4px)';
        const ns = addMin(o.start, dMin);
        if (meta) meta.textContent = `${fmtTime(ns)} – ${fmtTime(new Date(+ns + (o.end - o.start)))}`;
      } else {
        dMin = Math.max(dMin, 15 - (o.end - o.start) / 6e4);
        evEl.style.height = h0 + dMin * HOUR / 60 - 2 + 'px';
        if (meta) meta.textContent = `${fmtTime(o.start)} – ${fmtTime(addMin(o.end, dMin))}`;
      }
    };
    up = () => {
      off();
      if (!moved) return openPeek(o, evEl);
      if (mode === 'move') { const ns = addMin(addDays(o.start, dDay), dMin); moveOcc(o, ns, new Date(+ns + (o.end - o.start))); }
      else moveOcc(o, o.start, addMin(o.end, dMin));
    };
  } else {
    const s = snapMin(col, y0, true);
    const ghost = document.createElement('div');
    ghost.className = 'drag-new';
    const draw = (a, b) => {
      ghost.style.top = a * HOUR / 60 + 'px'; ghost.style.height = Math.max(15, b - a) * HOUR / 60 + 'px';
      const d = parse(col.dataset.day);
      ghost.textContent = `${fmtTime(addMin(d, a))} – ${fmtTime(addMin(d, Math.max(b, a + 15)))}`;
    };
    let a = s, b = s + 60;
    mv = ev => {
      if (!moved && Math.abs(ev.clientY - y0) < 5) return;
      if (!moved) { moved = true; col.appendChild(ghost); }
      const m = snapMin(col, ev.clientY);
      a = Math.min(s, m); b = Math.max(s + 15, m);
      draw(a, b);
    };
    up = () => {
      off(); ghost.remove();
      const d = parse(col.dataset.day);
      openEditor(null, { start: addMin(d, a), end: addMin(d, moved ? b : a + 60) });
    };
  }
  window.addEventListener('pointermove', mv);
  window.addEventListener('pointerup', up);
}

/* Move/resize one occurrence. Recurring series get an exception + a standalone copy. */
function moveOcc(o, ns, ne) {
  pushUndo();
  const ev = o.ev;
  const startStr = ev.allDay ? ymd(ns) + 'T00:00' : toStr(ns);
  const endStr = ev.allDay ? ymd(addDays(ne, -1)) + 'T00:00' : toStr(ne);
  if (isRecurring(ev)) {
    (ev.exdates ||= []).push(ymd(o.start));
    const done = isDone(o) ? [ymd(ns)] : [];
    S.events.push({ ...clone(ev), id: uid(), uid: undefined, repeat: 'none', until: '', days: [], exdates: [], doneDates: done, start: startStr, end: endStr, seriesOf: ev.id });
    toast('Moved this occurrence only', { undo: true, icon: 'repeat' });
  } else {
    const before = clone(ev);
    ev.start = startStr; ev.end = endStr;
    toast(`Moved to ${fmt(ns, { weekday: 'short', day: 'numeric', month: 'short' })}${ev.allDay ? '' : ', ' + fmtTime(ns)}`, { undo: !isSynced(ev.calId) });
    pushRemote(before, ev);
  }
  save(); render();
}

/* ---------------- year ---------------- */
function renderYear() {
  const y = cur.getFullYear();
  const occ = getOcc(new Date(y, 0, 1), new Date(y + 1, 0, 1));
  const count = {}, hol = {}, off = {}, bd = {};
  for (const o of occ) {
    if (isBirthday(o)) (bd[ymd(o.start)] ||= []).push(o.ev.title);
    // holiday labels (public holidays, School Holiday) mark the day but don't count as events; real days off also get the weekend shading
    if (isHoliday(o)) { hol[ymd(o.start)] = o.ev.title; if (isDayOff(o)) off[ymd(o.start)] = true; continue; }
    let d = startOfDay(o.start);
    do { count[ymd(d)] = (count[ymd(d)] || 0) + 1; d = addDays(d, 1); } while (d < o.end);
  }
  const today = ymd(new Date());
  let h = '<div class="year">';
  for (let m = 0; m < 12; m++) {
    const first = new Date(y, m, 1), gs = startOfWeek(first), dim = new Date(y, m + 1, 0).getDate();
    h += `<div class="ym ${m === new Date().getMonth() && y === new Date().getFullYear() ? 'curm' : ''}"><h3><button data-month="${m}">${fmt(first, { month: 'long' })}</button></h3><div class="ym-g">`;
    for (let i = 0; i < 7; i++) h += `<div class="h">${fmt(addDays(gs, i), { weekday: 'narrow' })}</div>`;
    const lead = dayDiff(first, gs), trail = (7 - ((lead + dim) % 7)) % 7;
    if (lead) h += `<div class="ym-blank" style="grid-column: span ${lead}"></div>`; // days of the previous month run together as one block
    for (let d = 1; d <= dim; d++) {
      const dt = new Date(y, m, d), k = ymd(dt), c = count[k] || 0;
      const lvl = c >= 5 ? 'l3' : c >= 3 ? 'l2' : c >= 1 ? 'l1' : '';
      const tip = [hol[k], bd[k] ? '🧁 ' + bd[k].join(', ') : '', c ? `${c} event${c > 1 ? 's' : ''}` : ''].filter(Boolean).join(' · ');
      h += `<button class="ym-d ${lvl} ${k === today ? 'td' : ''} ${k < today ? 'past' : ''} ${hol[k] ? 'hd' : ''} ${[0, 6].includes(dt.getDay()) || off[k] ? 'we' : ''}" data-goto="${k}" title="${esc(tip)}"><span class="ym-n">${d}</span>${bd[k] ? '<span class="ym-cake">🧁</span>' : ''}</button>`;
    }
    if (trail) h += `<div class="ym-blank" style="grid-column: span ${trail}"></div>`; // … and so do the days of the next month
    h += '</div></div>';
  }
  return h + '</div>';
}

/* ---------------- agenda / search results ---------------- */
function renderAgenda() {
  const searching = !!searchQ.trim();
  const from = searching ? addDays(startOfDay(new Date()), -365) : startOfDay(cur);
  const to = searching ? addDays(from, 365 * 2 + 1) : addDays(from, 60);
  const occ = getOcc(from, to);
  const groups = new Map();
  for (const o of occ) {
    const d = o.start < from ? from : startOfDay(o.start);
    const k = ymd(d); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(o);
  }
  if (!groups.size) return `<div class="ag-empty">${svg(searching ? 'search' : 'sparkles')}<b>${searching ? 'No events match “' + esc(searchQ) + '”' : 'Nothing planned for the next 60 days'}</b><span>${searching ? 'Try another word.' : 'Enjoy the free time — or press N to add something.'}</span></div>`;
  const today = ymd(new Date());
  let h = searching ? `<div class="search-note">${svg('search')}${occ.length} result${occ.length === 1 ? '' : 's'} for “${esc(searchQ)}” <button class="btn ghost" data-act="clearSearch">${svg('x')}Clear</button></div>` : '';
  h += '<div class="agenda">';
  for (const [k, list] of [...groups].sort()) {
    const d = parse(k);
    h += `<section class="ag-day ${k === today ? 'today' : ''}"><div class="ag-date"><span class="n">${d.getDate()}</span><small>${fmt(d, { weekday: 'short' })}<br>${fmt(d, { month: 'short' })}${d.getFullYear() !== new Date().getFullYear() ? ' ' + d.getFullYear() : ''}</small></div><div class="ag-list">`;
    for (const o of list) {
      const time = o.ev.allDay ? 'All day' : multiDay(o) ? `${fmtTime(o.start)}<small>→ ${fmt(o.end, { day: 'numeric', month: 'short' })}</small>` : `${fmtTime(o.start)}<small>${fmtTime(o.end)}</small>`;
      h += `<button class="ag-item ${isDone(o) ? 'done' : ''}" data-key="${esc(o.key)}" style="--c:${o.cal.color}">
        <span class="ag-time">${time}</span><span class="bar"></span><span class="em">${eventIcon(o.ev, o.cal) || svg('calendar')}</span>
        <span class="ag-main"><b>${taskBox(o)} ${esc(o.ev.title)}</b><small>${o.cal.icon || ''} ${esc(o.cal.name)}${o.ev.location ? ` · ${svg('pin')}${esc(o.ev.location)}` : ''}${pplTag(o.ev)}</small></span>
        ${isRecurring(o.ev) ? svg('repeat') : '<span></span>'}</button>`;
    }
    h += '</div></section>';
  }
  return h + '</div>';
}

/* ---------------- sidebar ---------------- */
function renderMini() {
  const m = new Date(miniCur.getFullYear(), miniCur.getMonth(), 1);
  const gs = startOfWeek(m);
  const today = ymd(new Date());
  let selA = null, selN = 0;
  if (['day', 'week', 'workweek'].includes(view)) [selA, selN] = gridRange();
  let h = `<div class="mini-h"><b>${fmt(m, { month: 'long', year: 'numeric' })}</b>
    <button class="ib sm" data-mini="-1" title="Previous month">${svg('chevronLeft')}</button><button class="ib sm" data-mini="1" title="Next month">${svg('chevronRight')}</button></div>
    <div class="mini-g ${S.settings.weekNums ? '' : 'nowk'}"><div></div>`;
  for (let i = 0; i < 7; i++) h += `<div class="h">${fmt(addDays(gs, i), { weekday: 'narrow' })}</div>`;
  for (let r = 0; r < 6; r++) {
    h += `<div class="wk">${isoWeek(addDays(gs, r * 7 + 3))}</div>`;
    for (let c = 0; c < 7; c++) {
      const d = addDays(gs, r * 7 + c), k = ymd(d);
      const sel = selA && d >= selA && d < addDays(selA, selN);
      h += `<button class="d ${d.getMonth() !== m.getMonth() ? 'o' : ''} ${k === today ? 'td' : ''} ${sel ? 'sel' : ''} ${k === ymd(cur) ? 'cur' : ''}" data-pick="${k}">${d.getDate()}</button>`;
    }
  }
  $('#mini').innerHTML = h + '</div>';
}
function renderSidebar() {
  renderMini();
  $('#calList').innerHTML = S.groups.map(g => {
    const cals = S.cals.filter(c => c.groupId === g.id);
    const badge = g.provider ? `<span class="pbadge ${g.provider}">${g.provider === 'google' ? 'G' : 'O'}</span>` : '';
    return `<section class="grp ${g.collapsed ? 'collapsed' : ''}" data-g="${g.id}">
      <div class="grp-h" draggable="true" title="Drag to move this folder"><button class="grp-t" data-act="toggleGroup">${svg('chevronDown', 'chev')}${badge}<span>${esc(g.name)}</span></button>
        ${g.account ? `<button class="ib sm" data-act="syncSettings" title="Sync settings">${svg('cloud')}</button>` : `<button class="ib sm" data-act="addCal" title="Add a calendar to ${esc(g.name)}">${svg('plus')}</button>
        <button class="ib sm" data-act="editGroup" title="Rename or remove group">${svg('edit')}</button>`}</div>
      <div class="grp-b">${cals.map(c => {
        const src = c.source;
        const tip = src ? `Synced with ${src.type === 'ics' ? 'a calendar link' : src.type === 'google' ? 'Google' : 'Outlook'}${c.readonly ? ' (read-only)' : ''}${src.error ? ' — problem: ' + src.error : ''}` : c.name;
        return `<div class="cal ${c.visible ? 'on' : ''}" data-c="${c.id}" style="--c:${c.color}" draggable="true">
        <button class="sw" data-act="toggleCal" role="checkbox" aria-checked="${c.visible}" title="Show / hide">${svg('check')}</button>
        <span class="cal-ic">${c.icon || ''}</span><span class="cal-n" data-act="toggleCal" title="${esc(tip)}">${esc(c.name)}</span>
        ${src?.error ? `<span class="cal-err" title="${esc(src.error)}">${svg('alert')}</span>` : src ? `<span class="cal-sync" title="${esc(tip)}">${svg(src.type === 'ics' ? 'link' : 'cloud')}</span>` : ''}
        <button class="ib sm more" data-act="editCal" title="Edit calendar">${svg('more')}</button></div>`;
      }).join('') || '<div class="empty">No calendars yet</div>'}</div></section>`;
  }).join('');
}
function renderUpNext() {
  // Three parts, refreshed every 30 seconds by tick():
  //  Missed  – tasks from the last 7 days whose time has passed and that aren't ticked off
  //  To do   – today's all-day tasks not ticked off yet
  //  Coming up – today's timed events that haven't finished
  const now = new Date(), today0 = startOfDay(now), endOfToday = addDays(today0, 1);
  const occ = getOcc(addDays(today0, -7), endOfToday, { all: true });
  const isAllDayish = o => o.ev.allDay || multiDay(o);
  const missed = occ.filter(o => o.ev.task && !o.ro && !isDone(o) && o.end <= now).sort((a, b) => b.start - a.start);
  const todos = occ.filter(o => o.ev.task && !o.ro && !isDone(o) && isAllDayish(o) && o.start < endOfToday && o.end > now);
  const coming = occ.filter(o => !isAllDayish(o) && o.end > now && o.start < endOfToday && !(o.ev.task && isDone(o)))
    .sort((a, b) => a.start - b.start);
  const when = o => {
    if (o.start <= now) return `Now · until ${fmtTime(o.end)}`;
    const mins = Math.round((o.start - now) / 6e4);
    return mins < 60 ? `In ${mins} min · ${fmtTime(o.start)}` : `${fmtTime(o.start)} · in ${durText(o.start - now)}`;
  };
  const missedWhen = o => {
    const day = ymd(o.start) === ymd(now) ? 'today' : fmt(o.start, { weekday: 'short', day: 'numeric', month: 'short' });
    return `Missed · ${day}${isAllDayish(o) ? '' : ', ' + fmtTime(o.start)}`;
  };
  const item = (o, sub, cls = '') => `<button class="un ${cls}" data-key="${esc(o.key)}" style="--c:${o.cal.color}" title="${esc(o.ev.title)}">
    <span class="em">${eventIcon(o.ev, o.cal) || '📅'}</span><span class="tx"><b>${esc(o.ev.title)}</b>${sub || o.ev.attendees?.length ? `<small>${sub || ''}${pplTag(o.ev)}</small>` : ''}</span>${o.ev.task && !o.ro ? taskBox(o) : ''}</button>`;
  const section = (label, arr, fn, cls, max) => !arr.length ? '' :
    `<div class="un-sec ${cls}">${label}</div>` + arr.slice(0, max).map(o => item(o, fn(o), cls)).join('')
    + (arr.length > max ? `<div class="un-more">+${arr.length - max} more</div>` : '');
  const body = section('Missed', missed, missedWhen, 'missed', 4)
    + section('To do today', todos, () => '', 'todo', 5)
    + section('Coming up', coming, when, '', 4);
  $('#upnext').innerHTML = `<h4>${svg('clock')}Up next</h4>` + (body || '<div class="un-empty">Nothing else on today 🎉</div>');
}

/* ---------------- theme / brand ---------------- */
const mq = matchMedia('(prefers-color-scheme: dark)');
function applyTheme() {
  const t = S.settings.theme === 'auto' ? (mq.matches ? 'dark' : 'light') : S.settings.theme;
  const root = document.documentElement;
  root.dataset.theme = t; root.dataset.bg = S.settings.bg; root.dataset.look = S.settings.look;
  $('#themeBtn').innerHTML = svg(t === 'dark' ? 'sun' : 'moon');
}
function logoSVG(d) {
  return `<svg viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="lg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ff5c95"/><stop offset="1" stop-color="#8b5cf6"/></linearGradient></defs><rect x="2" y="3" width="28" height="27" rx="7" fill="#fff"/><path d="M2 10a7 7 0 0 1 7-7h14a7 7 0 0 1 7 7v1H2z" fill="url(#lg)"/><text x="16" y="26.5" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-weight="700" font-size="14" fill="#2a2140">${d.getDate()}</text></svg>`;
}
let brandDay = '';
function updateBrand() {
  const now = new Date(), k = ymd(now);
  if (k === brandDay) return;
  brandDay = k;
  $('#logo').innerHTML = logoSVG(now);
  $('#favicon').href = 'data:image/svg+xml,' + encodeURIComponent(logoSVG(now));
  document.title = `${fmt(now, { weekday: 'short', day: 'numeric', month: 'short' })} · Calendar`;
}

/* ---------------- navigation ---------------- */
function setView(v) {
  view = v; S.settings.view = v; save(); gridScroll = null; render();
}
function go(dir) {
  if (view === 'month') cur = new Date(cur.getFullYear(), cur.getMonth() + dir, 1);
  else if (view === 'year') cur = new Date(cur.getFullYear() + dir, cur.getMonth(), 1);
  else if (view === 'day') cur = addDays(cur, dir);
  else if (view === 'agenda') cur = addDays(cur, dir * 30);
  else cur = addDays(cur, dir * 7);
  miniCur = new Date(cur); render();
}
function goToday() { cur = startOfDay(new Date()); miniCur = new Date(cur); gridScroll = null; render(); }

/* ---------------- popover (+N more) ---------------- */
function showDayPop(k, anchor, onlyBirthdays) {
  const pop = $('#pop'), d = parse(k);
  let list = getOcc(d, addDays(d, 1));
  if (onlyBirthdays) list = list.filter(isBirthday);
  else if (view === 'month') list = list.filter(o => !isBirthday(o));
  const head = onlyBirthdays
    ? `🧁 ${list.length > 1 ? `${list.length} birthdays` : 'Birthday'} · ${fmt(d, { day: 'numeric', month: 'long' })}`
    : fmt(d, { weekday: 'long', day: 'numeric', month: 'long' });
  const rows = onlyBirthdays
    ? list.map(o => `<button class="bday-row" data-key="${esc(o.key)}" style="--c:${o.cal.color}"><span class="bd-ic">${o.ev.icon === 'none' ? '' : o.ev.icon || '🎂'}</span>
        <span><b>${esc(o.ev.title.replace(/['\u2019]?s?\s*(birthday|b-?day)\s*$/i, '').trim() || o.ev.title)}</b><small>${esc(o.ev.title)}${pplTag(o.ev)}</small></span></button>`).join('')
    : list.map(o => chip(o, k)).join('');
  pop.innerHTML = `<h4><span>${head}</span><button class="ib sm" data-act="closePop">${svg('x')}</button></h4>${rows}`;
  pop.hidden = false;
  const r = anchor.getBoundingClientRect();
  pop.style.left = clamp(r.left, 8, innerWidth - 272) + 'px';
  pop.style.top = clamp(r.top - 40, 8, innerHeight - pop.offsetHeight - 8) + 'px';
}
const hidePop = () => { $('#pop').hidden = true; };

/* ---------------- event overview (left-click) ---------------- */
let peekOcc = null, peekRect = null;
const hidePeek = () => { $('#peek').hidden = true; peekOcc = null; };
const isUrl = s => /^(https?:\/\/|www\.)\S+$/i.test(s.trim());
const isPlace = s => !!s && !isUrl(s) && !/https?:\/\//i.test(s) && /[a-z]/i.test(s) && !/^(in[- ]?person|online|virtual|tbc|tba|home|n\/?a|teams|zoom|google meet)$/i.test(s.trim());
const linkify = s => esc(s).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
function durText(ms) {
  const m = Math.round(ms / 6e4);
  if (m < 60) return `${m} min`;
  const h = Math.floor(m / 60), r = m % 60;
  return `${h} hr${h > 1 ? 's' : ''}${r ? ` ${r} min` : ''}`;
}
function remText(m) {
  if (m === 0) return 'At start time';
  for (const [n, u] of [[10080, 'week'], [1440, 'day'], [60, 'hour']]) if (m % n === 0) return `${m / n} ${u}${m > n ? 's' : ''} before`;
  return `${m} minutes before`;
}
function repeatText(ev) {
  let s = { daily: 'Every day', weekdays: 'Every weekday (Mon–Fri)', weekly: 'Every week', biweekly: 'Every 2 weeks', monthly: 'Every month', yearly: 'Every year' }[ev.repeat];
  if ((ev.repeat === 'weekly' || ev.repeat === 'biweekly') && ev.days?.length > 0) {
    const ws = startOfWeek(new Date());
    s += ' on ' + [...ev.days].sort((a, b) => ((a - S.settings.weekStart + 7) % 7) - ((b - S.settings.weekStart + 7) % 7))
      .map(d => fmt(addDays(ws, (d - S.settings.weekStart + 7) % 7), { weekday: 'short' })).join(', ');
  }
  if (ev.until) s += ` until ${fmt(parse(ev.until), { day: 'numeric', month: 'short', year: 'numeric' })}`;
  return s;
}
function whenBadge(o) {
  const now = new Date();
  if (o.start <= now && o.end > now) return ['Happening now', 'live'];
  if (o.end <= now) return ['Ended', 'past'];
  const d = dayDiff(o.start, now);
  if (d === 0) { const m = Math.round((o.start - now) / 6e4); return [m < 60 ? `Starts in ${m} min` : `Today, in ${durText(o.start - now)}`, 'soon']; }
  if (d === 1) return ['Tomorrow', ''];
  return [d < 14 ? `In ${d} days` : `In ${Math.round(d / 7)} weeks`, ''];
}
/* "People" section of the details card: who's invited and how they replied. */
const REPLY = { accepted: ['✓', 'Going'], declined: ['✗', "Can't go"], tentative: ['?', 'Maybe'], needsAction: ['⏳', 'Not replied'] };
function peopleBlock(ev) {
  const people = ev.attendees || [];
  if (!people.length && !ev.organizer && !ev.meet) return '';
  const count = k => people.filter(p => (p.status || 'needsAction') === k).length;
  const summary = [['accepted', 'going'], ['tentative', 'maybe'], ['declined', "can't go"], ['needsAction', 'not replied']]
    .map(([k, w]) => count(k) ? `${count(k)} ${w}` : '').filter(Boolean).join(' · ');
  const list = people.map(p => {
    const [ic, word] = REPLY[p.status] || REPLY.needsAction;
    return `<div class="pp"><span class="av" style="--h:${hueOf(p.email || p.name)}">${esc(initials(p))}</span>
      <span class="pn"><b>${esc(p.name || p.email)}</b>${p.self ? ' <small>(you)</small>' : ''}${p.organizer ? ' <small>organiser</small>' : ''}${p.name ? `<small>${esc(p.email)}</small>` : ''}</span>
      <span class="ps ${p.status || 'needsAction'}" title="${word}">${ic}</span></div>`;
  }).join('');
  return `<div class="pk-row">${svg('users')}<div class="pk-main">
      ${people.length ? `<b>${people.length} ${people.length === 1 ? 'person' : 'people'}</b>${summary ? `<span>${summary}</span>` : ''}` : ''}
      ${ev.organizer ? `<span>Invited by <b>${esc(ev.organizer)}</b></span>` : ''}
      ${ev.meet ? `<div class="pk-actions"><a class="btn primary sm" href="${esc(ev.meet)}" target="_blank" rel="noopener">${svg('external')}<span>Join video call</span></a></div>` : ''}
      ${list ? `<div class="ppl">${list}</div>` : ''}
    </div></div>`;
}
function openPeek(o, anchor) {
  if (!o) return;
  peekRect = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
  hidePop();
  peekOcc = o;
  const { ev, cal } = o;
  const ic = eventIcon(ev, cal) || '📅';
  const group = S.groups.find(g => g.id === cal.groupId);
  const long = d => fmt(d, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const short = d => fmt(d, { weekday: 'short', day: 'numeric', month: 'short' });
  let dateLine, timeLine;
  if (ev.allDay) {
    const last = addDays(o.end, -1), n = dayDiff(last, o.start) + 1;
    dateLine = n === 1 ? long(o.start) : `${short(o.start)} – ${short(last)} ${last.getFullYear()}`;
    timeLine = n === 1 ? 'All day' : `All day · ${n} days`;
  } else if (ymd(o.start) === ymd(addMin(o.end, -1))) {
    dateLine = long(o.start);
    timeLine = `${fmtTime(o.start)} – ${fmtTime(o.end)} · ${durText(o.end - o.start)}`;
  } else {
    dateLine = `${short(o.start)}, ${fmtTime(o.start)} → ${short(o.end)}, ${fmtTime(o.end)}`;
    timeLine = durText(o.end - o.start);
  }
  const [when, whenCls] = whenBadge(o);
  const loc = (ev.location || '').trim();
  const url = loc.match(/https?:\/\/\S+/)?.[0] || (/^www\./i.test(loc) ? 'https://' + loc : '');
  const q = encodeURIComponent(loc);
  const src = cal.source;
  const provider = src?.type === 'google' ? 'Google' : src?.type === 'microsoft' ? 'Outlook' : '';
  const done = isDone(o);
  const locBlock = !loc ? '' : `
    <div class="pk-row">${svg('pin')}<div class="pk-main">
      <b class="pk-loc">${esc(loc)}</b>
      <div class="pk-actions">
        ${url ? `<a class="btn primary sm" href="${esc(url)}" target="_blank" rel="noopener">${svg('external')}<span>Join</span></a>` : ''}
        ${isPlace(loc) ? `<a class="btn sm" href="https://www.google.com/maps/dir/?api=1&destination=${q}" target="_blank" rel="noopener">${svg('route')}<span>Directions</span></a>
        <a class="btn sm" href="https://www.google.com/maps/search/?api=1&query=${q}" target="_blank" rel="noopener">${svg('map')}<span>Open map</span></a>` : ''}
        <button class="btn ghost sm" data-pk="copyloc" title="Copy location">${svg('copy')}<span>Copy</span></button>
      </div>
      ${isPlace(loc) ? `<iframe class="pk-map" title="Map of ${esc(loc)}" src="https://maps.google.com/maps?q=${q}&z=15&output=embed"></iframe>` : ''}
    </div></div>`;
  $('#peek').innerHTML = `
    <header class="pk-head" style="--c:${cal.color}">
      <span class="pk-ic">${ic}</span>
      <div class="pk-title"><h3 class="${done ? 'done' : ''}">${esc(ev.title)}</h3>
        <small><i class="dot"></i>${esc(cal.name)}${group ? ` · ${esc(group.name)}` : ''}</small></div>
      ${o.ro ? '' : `<button class="ib sm" data-pk="edit" title="Edit (E)">${svg('edit')}</button>`}
      <button class="ib sm" data-pk="close" title="Close (Esc)">${svg('x')}</button>
    </header>
    <div class="pk-body">
      <div class="pk-row">${svg('clock')}<div class="pk-main"><b>${dateLine}</b><span>${timeLine}</span><em class="when ${whenCls}">${when}</em></div></div>
      ${isRecurring(ev) ? `<div class="pk-row">${svg('repeat')}<div class="pk-main"><span>${repeatText(ev)}</span></div></div>`
        : ev.remote?.series ? `<div class="pk-row">${svg('repeat')}<div class="pk-main"><span>Part of a repeating event</span></div></div>` : ''}
      ${ev.reminder != null && !ev.allDay ? `<div class="pk-row">${svg('bell')}<div class="pk-main"><span>Reminder ${remText(ev.reminder).toLowerCase()}</span></div></div>` : ''}
      ${locBlock}
      ${peopleBlock(ev)}
      ${ev.notes ? `<div class="pk-row">${svg('note')}<div class="pk-main pk-notes">${linkify(ev.notes)}</div></div>` : ''}
      ${ev.task && !o.ro ? `<div class="pk-row">${svg('checkCircle')}<div class="pk-main"><button class="btn sm ${done ? 'primary' : ''}" data-pk="done">${svg('check')}<span>${done ? 'Done — tap to undo' : 'Mark as done'}</span></button></div></div>` : ''}
      ${src ? `<div class="pk-row">${svg(src.type === 'ics' ? 'link' : 'cloud')}<div class="pk-main"><span>Synced with ${provider || 'a calendar link'}${o.ro ? ' · read-only' : ''}</span>
        ${ev.remote?.link ? `<a href="${esc(ev.remote.link)}" target="_blank" rel="noopener">Open in ${provider} ↗</a>` : ''}</div></div>` : ''}
    </div>
    ${o.ro ? '' : `<footer class="pk-foot">
      <button class="btn primary" data-pk="edit">${svg('edit')}<span>Edit</span></button>
      <button class="btn ghost" data-pk="dup">${svg('copy')}<span>Duplicate</span></button>
      <span class="grow"></span>
      <button class="btn danger" data-pk="del" title="Delete (Del)">${svg('trash')}<span>Delete</span></button>
    </footer>
    <div class="pk-confirm" hidden><span>Delete which?</span>
      <button class="btn sm" data-pk="delThis">Only this date</button><button class="btn sm danger" data-pk="delAll">Whole series</button>
      <button class="btn ghost sm" data-pk="delCancel">Cancel</button></div>`}
    <small class="pk-tip">Tip: right-click an event to edit it straight away</small>`;
  const p = $('#peek');
  p.hidden = false;
  placePeek();
  p.querySelector('.pk-map')?.addEventListener('load', placePeek, { once: true });
}
function placePeek() {
  const p = $('#peek'), r = peekRect, W = p.offsetWidth, H = p.offsetHeight, gap = 10;
  let x = r.right + gap, y = r.top - 12;
  if (x + W > innerWidth - 8) x = r.left - W - gap;
  if (x < 8) { x = clamp(r.left, 8, innerWidth - W - 8); y = r.bottom + gap; if (y + H > innerHeight - 8) y = r.top - H - gap; }
  p.style.left = x + 'px';
  p.style.top = clamp(y, 8, Math.max(8, innerHeight - H - 8)) + 'px';
}
function onPeekClick(e) {
  const b = e.target.closest('[data-pk]'); if (!b) return;
  const o = peekOcc; if (!o) return;
  switch (b.dataset.pk) {
    case 'close': hidePeek(); break;
    case 'edit': hidePeek(); openEditor(o); break;
    case 'dup': {
      hidePeek(); pushUndo();
      const copy = { ...clone(o.ev), id: uid(), uid: undefined, remote: undefined, seriesOf: undefined, title: o.ev.title + ' (copy)', repeat: 'none', until: '', days: [], exdates: [], doneDates: [],
        start: o.ev.allDay ? ymd(o.start) + 'T00:00' : toStr(o.start), end: o.ev.allDay ? ymd(addDays(o.end, -1)) + 'T00:00' : toStr(o.end) };
      S.events.push(copy); pushRemote(null, copy); save(); render();
      toast('Duplicated — right-click it to edit', { undo: !isSynced(copy.calId), icon: 'copy' });
      break;
    }
    case 'del':
      if (isRecurring(o.ev)) { $('.pk-confirm', $('#peek')).hidden = false; placePeek(); }
      else { hidePeek(); deleteOcc(o, true); }
      break;
    case 'delThis': hidePeek(); deleteOcc(o, false); break;
    case 'delAll': hidePeek(); deleteOcc(o, true); break;
    case 'delCancel': $('.pk-confirm', $('#peek')).hidden = true; break;
    case 'done': { toggleDone(o); const fresh = occMap.get(o.key); if (fresh) openPeek(fresh, peekRect); break; }
    case 'copyloc':
      navigator.clipboard?.writeText(o.ev.location).then(() => toast('Location copied', { icon: 'copy' }), () => toast('Could not copy', { icon: 'alert' }));
      break;
  }
}

/* ---------------- event editor ---------------- */
let editing = null, editIcon = '', editDays = [], editGuests = [];
/* Guests (people added to an event) in the editor. */
const EMAIL_RE = /^[^\s@<>,;]+@[^\s@<>,;]+\.[^\s@<>,;]+$/;
const guestLabel = a => a.name ? `${a.name}` : a.email;
function drawGuests() {
  $('#evGuestList').innerHTML = editGuests.map((a, i) => `<span class="guest" title="${esc(a.email)}"><span class="av" style="--h:${hueOf(a.email)}">${esc(initials(a))}</span>${esc(guestLabel(a))}${a.organizer ? ' <small>organiser</small>' : ''}
    <button type="button" data-rm="${i}" title="Remove">${svg('x')}</button></span>`).join('');
  const synced = isSynced($('#evCal').value);
  $('#evNotifyWrap').hidden = !(synced && editGuests.length);
}
function addGuestFromInput() {
  const inp = $('#evGuestIn');
  const parts = inp.value.split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
  const bad = [];
  for (const p of parts) {
    if (!EMAIL_RE.test(p)) { bad.push(p); continue; }
    if (!editGuests.some(a => a.email.toLowerCase() === p.toLowerCase())) editGuests.push({ email: p, name: knownName(p), status: 'needsAction' });
  }
  inp.value = bad.join(' ');
  if (bad.length) toast(`“${bad[0]}” isn't an email address`, { icon: 'alert' });
  drawGuests();
}
/* People you've invited before, for suggestions and names. */
function knownPeople() {
  const m = new Map();
  for (const e of S.events) for (const a of e.attendees || []) if (a.email && !a.self && !m.has(a.email.toLowerCase())) m.set(a.email.toLowerCase(), a);
  return [...m.values()];
}
const knownName = email => knownPeople().find(a => a.email.toLowerCase() === email.toLowerCase())?.name || '';
const initials = a => (a.name || a.email).replace(/@.*/, '').split(/[\s._-]+/).filter(Boolean).slice(0, 2).map(w => w[0].toUpperCase()).join('') || '?';
const hueOf = s => [...s].reduce((h, c) => (h * 31 + c.charCodeAt(0)) % 360, 7);
function newOnDay(k, allDay) {
  const d = parse(k);
  if (allDay) return openEditor(null, { start: d, end: addDays(d, 1), allDay: true });
  const now = new Date();
  const s = ymd(d) === ymd(now) ? addMin(startOfDay(now), (now.getHours() + 1) * 60) : addMin(d, 9 * 60);
  openEditor(null, { start: s, end: addMin(s, 60) });
}
function openEditor(o, preset = {}) {
  hidePop(); hidePeek();
  const ev = o?.ev;
  const isNew = !o;
  editing = { o, isNew };
  const allDay = o ? !!ev.allDay : !!preset.allDay;
  const s = o ? o.start : preset.start || addMin(startOfDay(new Date()), (new Date().getHours() + 1) * 60);
  let e = o ? o.end : preset.end || addMin(s, 60);
  if (allDay) e = addDays(e, -1);
  const writable = S.cals.filter(c => !c.readonly);
  const calId = ev?.calId || preset.calId || (writable.find(c => c.id === S.settings.defaultCal) || writable[0])?.id;
  const sel = $('#evCal');
  sel.innerHTML = (o?.ro ? [o.cal] : writable).map(c => `<option value="${c.id}">${c.icon || '●'}  ${esc(c.name)}  ·  ${esc(S.groups.find(g => g.id === c.groupId)?.name || '')}</option>`).join('');
  sel.value = calId;
  $('#evTitle').value = ev?.title || preset.title || '';
  $('#evSD').value = ymd(s); $('#evST').value = hm(s);
  $('#evED').value = ymd(e); $('#evET').value = hm(e);
  $('#evAll').checked = allDay;
  $('#evRep').value = ev?.repeat || 'none';
  editDays = ev?.days?.length ? [...ev.days] : [s.getDay()];
  $('#evUntil').value = ev?.until || '';
  const rem = ev ? ev.reminder : (allDay ? null : S.settings.defaultReminder);
  $('#evRem').value = rem == null || rem === '' ? '' : String(rem);
  $('#evLoc').value = ev?.location || '';
  $('#evNotes').value = ev?.notes || '';
  $('#evTask').checked = !!ev?.task;
  editIcon = ev?.icon || '';
  editGuests = clone(ev?.attendees || []);
  $('#evGuestIn').value = '';
  $('#evNotify').checked = false;
  $('#guestSuggest').innerHTML = knownPeople().map(a => `<option value="${esc(a.email)}">${esc(a.name)}</option>`).join('');
  drawGuests();
  // "Save to all": the repeating series this belongs to + other events with the same name.
  const rel = o && !o.ro ? relatedEvents(ev) : [];
  const series = !!(o && !o.ro && (isRecurring(ev) || rel.some(x => isRecurring(x))));
  editing.related = rel;
  const groupSize = rel.length + 1;
  const hasGroup = series || rel.length > 0;
  $('#evSaveAll').hidden = !hasGroup;
  const allLabel = !series ? `(${groupSize} events)` : null; // finished below once we know the series
  $('#evSaveAllN').textContent = allLabel || '';
  $('#evSaveLbl').textContent = hasGroup ? 'Save this one' : 'Save';
  $('#seriesRow').hidden = !hasGroup;
  const master = o && (isRecurring(ev) ? ev : rel.find(isRecurring));
  const seriesId = ev?.seriesOf || (ev && isRecurring(ev) ? ev.id : null);
  const extras = rel.filter(x => !isRecurring(x) && !(seriesId && x.seriesOf === seriesId)); // other one-off events with the same name
  const more = extras.length ? ` There ${extras.length === 1 ? 'is' : 'are'} also ${extras.length} other “${esc(ev.title)}” event${extras.length > 1 ? 's' : ''}.` : '';
  const day = o ? fmt(o.start, { day: 'numeric', month: 'short' }) : '';
  $('#seriesText').innerHTML = !hasGroup ? '' : master
    ? `Repeats ${esc(repeatText(master).toLowerCase())}.${ev !== master ? ' This date was changed separately from the series.' : ''}${more}
       <b>Save this one</b> changes only ${day}; <b>Save to all</b> updates the whole series${extras.length ? ' and the others' : ''} (each keeps its own date).`
    : `There ${rel.length === 1 ? 'is' : 'are'} ${rel.length} other “${esc(ev.title)}” event${rel.length > 1 ? 's' : ''}. <b>Save to all</b> copies your changes to every one — each keeps its own date.`;
  const finalLabel = allLabel || (extras.length ? `(whole series + ${extras.length} more)` : '(whole series)');
  $('#evSaveAllN').textContent = finalLabel;
  $('#evDelConfirm').hidden = true;
  $('#evDelAllN').textContent = finalLabel;
  $('#evDel').hidden = isNew || o.ro;
  $('#evDup').hidden = isNew || o.ro;
  $('#evSave').hidden = !!o?.ro;
  if (o?.ro) $('#evSaveAll').hidden = true;
  $('#roNote').hidden = !o?.ro;
  $$('#evForm input, #evForm select, #evForm textarea').forEach(el => el.disabled = !!o?.ro);
  // Events that came from Google / Outlook: show where they live; series rules are edited there.
  const srcCal = ev && S.cals.find(c => c.id === ev.calId);
  const rr = $('#remoteRow');
  rr.hidden = !srcCal?.source;
  if (srcCal?.source) {
    const where = srcCal.source.type === 'google' ? 'Google Calendar' : srcCal.source.type === 'microsoft' ? 'Outlook' : 'a subscribed link';
    const acct = S.groups.find(g => g.id === srcCal.groupId)?.account ? S.groups.find(g => g.id === srcCal.groupId).name : '';
    rr.innerHTML = `${svg('cloud')}<span>Synced with ${where}${acct ? ` · <b>${esc(acct)}</b>` : ''}${ev.remote?.series ? '<br><small>One date of a repeating event — changes here apply to this date only.</small>' : ''}</span>
      ${ev.remote?.link ? `<a class="btn ghost" href="${esc(ev.remote.link)}" target="_blank" rel="noopener">${svg('external')}<span>Open in ${srcCal.source.type === 'google' ? 'Google' : 'Outlook'}</span></a>` : ''}`;
    if (isSynced(ev.calId)) $('#evRep').disabled = true;
  }
  $('#evIconPop').hidden = true;
  syncEditor();
  $('#dlgEvent').showModal();
  if (!o?.ro) $('#evTitle').focus();
}
function syncEditor() {
  const all = $('#evAll').checked;
  $('#evST').hidden = $('#evET').hidden = all;
  const rep = $('#evRep').value;
  $('#untilWrap').hidden = rep === 'none';
  const showDays = rep === 'weekly' || rep === 'biweekly';
  const dc = $('#evDays'); dc.hidden = !showDays;
  if (showDays) {
    const order = [...Array(7)].map((_, i) => (S.settings.weekStart + i) % 7);
    const base = startOfWeek(new Date());
    dc.innerHTML = order.map((d, i) => `<button type="button" data-d="${d}" class="${editDays.includes(d) ? 'on' : ''}" title="${fmt(addDays(base, i), { weekday: 'long' })}">${fmt(addDays(base, i), { weekday: 'narrow' })}</button>`).join('');
  }
  const cal = S.cals.find(c => c.id === $('#evCal').value);
  const btn = $('#evIconBtn');
  btn.style.setProperty('--c', cal?.color || 'var(--accent)');
  const ic = editIcon === 'none' ? '' : editIcon || eventIcon({ title: $('#evTitle').value, icon: '' }, cal);
  btn.innerHTML = ic || svg('smile');
  btn.title = editIcon === 'none' ? 'No icon — click to change' : editIcon ? 'Icon (click to change)' : 'Automatic icon — click to choose your own';
}
function readForm() {
  const allDay = $('#evAll').checked;
  const sd = $('#evSD').value || ymd(new Date()), ed = $('#evED').value || sd;
  const st = $('#evST').value || '09:00', et = $('#evET').value || st;
  let start = allDay ? sd + 'T00:00' : `${sd}T${st}`;
  let end = allDay ? ed + 'T00:00' : `${ed}T${et}`;
  if (parse(end) < parse(start)) end = allDay ? start : toStr(addMin(parse(start), 60));
  const rep = $('#evRep').value;
  const rem = $('#evRem').value;
  return {
    title: $('#evTitle').value.trim() || '(No title)', calId: $('#evCal').value, allDay, start, end,
    repeat: rep, until: rep === 'none' ? '' : $('#evUntil').value,
    days: rep === 'weekly' || rep === 'biweekly' ? [...editDays].sort() : [],
    reminder: rem === '' ? null : +rem, location: $('#evLoc').value.trim(), notes: $('#evNotes').value,
    task: $('#evTask').checked, icon: editIcon,
    attendees: editGuests.map(a => ({ ...a })),
    ...($('#evNotify').checked && isSynced($('#evCal').value) ? { _notify: true } : {}),
  };
}
/* Events that "Save to all" / "Delete all" also touch: the repeating series this event came from
   (and its split-off dates) plus every other event with exactly the same name. */
function relatedEvents(ev) {
  const name = ev.title.trim().toLowerCase();
  const seriesId = ev.seriesOf || (isRecurring(ev) ? ev.id : null);
  const writable = id => !S.cals.find(c => c.id === id)?.readonly;
  return S.events.filter(x => x !== ev && writable(x.calId) && (
    (seriesId && (x.id === seriesId || x.seriesOf === seriesId)) || x.title.trim().toLowerCase() === name));
}
/* Copy the form's changes to the edited event and its group. Each event keeps its own date. */
function applyToAll(o, data, related) {
  const ev = o.ev;
  const st = data.allDay ? '00:00' : data.start.slice(11), et = data.allDay ? '00:00' : data.end.slice(11);
  const span = dayDiff(parse(data.end), parse(data.start));
  const { title, calId, icon, allDay, reminder, location, notes, task } = data;
  // The event being edited: take everything, including repeat settings. A date change on a series shifts the series.
  const before = clone(ev);
  if (isRecurring(ev)) {
    const shift = dayDiff(parse(data.start), o.start);
    const ns = addDays(startOfDay(parse(ev.start)), shift);
    Object.assign(ev, data, { start: ymd(ns) + 'T' + st, end: ymd(addDays(ns, span)) + 'T' + et });
    if (shift) ev.exdates = [];
  } else Object.assign(ev, data);
  // Dates of a repeating Google/Outlook event are sent as ONE update for the whole series.
  const series = new Map();
  let moveBlocked = false;
  const send = (b, x) => {
    // A whole repeating Google/Outlook series can't be moved to another calendar from here: keep it where it is.
    if (x.remote?.series && isSynced(b.calId) && b.calId !== x.calId) { x.calId = b.calId; moveBlocked = true; }
    if (x.remote?.series && isSynced(x.calId) && b.calId === x.calId) {
      if (remoteChanged(b, x) && !series.has(x.remote.series)) series.set(x.remote.series, [b, x]);
    } else pushRemote(b, x);
  };
  send(before, ev);
  // The rest: same details and times, but their own dates and repeat settings.
  for (const x of related) {
    const b = clone(x), d0 = startOfDay(parse(x.start));
    Object.assign(x, { title, calId, icon, allDay, reminder, location, notes, task,
      start: ymd(d0) + 'T' + st, end: ymd(addDays(d0, span)) + 'T' + et });
    send(b, x);
  }
  for (const [b, x] of series.values()) Sync.pushSeries(b, x);
  if (moveBlocked) setTimeout(() => toast('Repeating Google/Outlook events stay in their calendar — your other changes were saved. To move the series, use “Open in Google”.', { icon: 'alert', ms: 9000 }), 300);
  return related.length + 1;
}
function saveEditor(all = false) {
  const { o, isNew, related = [] } = editing;
  if (o?.ro) return;
  if ($('#evGuestIn').value.trim()) addGuestFromInput(); // an email typed but not yet added
  pushUndo();
  const data = readForm();
  const remote = isSynced(data.calId) || (o && isSynced(o.ev.calId)) || (all && related.some(x => isSynced(x.calId)));
  let msg = 'Saved';
  if (isNew) {
    const ev = { id: uid(), exdates: [], doneDates: [], ...data };
    S.events.push(ev);
    pushRemote(null, ev);
    msg = `Added “${data.title}”`;
  } else if (all) {
    const n = applyToAll(o, data, related);
    msg = `Saved to all ${n} “${data.title}” events${isRecurring(o.ev) || related.some(isRecurring) ? ' (and their repeats)' : ''}`;
  } else {
    const ev = o.ev;
    const before = clone(ev);
    if (isRecurring(ev)) {
      // Only this date: split it off from the series (it remembers which series it came from).
      (ev.exdates ||= []).push(ymd(o.start));
      const one = { id: uid(), exdates: [], doneDates: isDone(o) ? [data.start.slice(0, 10)] : [], ...data, repeat: 'none', until: '', days: [], seriesOf: ev.id };
      S.events.push(one);
      pushRemote(null, one);
      msg = `Saved ${fmt(o.start, { day: 'numeric', month: 'short' })} only`;
    } else {
      Object.assign(ev, data);
      pushRemote(before, ev);
    }
  }
  save(); $('#dlgEvent').close(); render();
  toast(msg, { undo: !remote });
}
/* Delete one occurrence (all=false, repeating events only) or the whole event/series. */
function deleteOcc(o, all) {
  const ev = o.ev;
  const onlyThis = isRecurring(ev) && !all;
  pushUndo();
  if (onlyThis) (ev.exdates ||= []).push(ymd(o.start));
  else { S.events = S.events.filter(x => x !== ev); pushRemote(clone(ev), null); }
  save(); render();
  toast(`Deleted “${ev.title}”${onlyThis ? ' on ' + fmt(o.start, { day: 'numeric', month: 'short' }) : ''}`, { undo: !isSynced(ev.calId), icon: 'trash' });
}
function deleteEditor(all) {
  const { o, related = [] } = editing;
  $('#dlgEvent').close();
  if (!all) return deleteOcc(o, false);
  pushUndo();
  const gone = new Set([o.ev, ...related]);
  S.events = S.events.filter(x => !gone.has(x));
  const seriesDone = new Set();
  for (const x of gone) {
    // a repeating Google/Outlook event is deleted once, as a whole series
    if (x.remote?.series && isSynced(x.calId)) { if (!seriesDone.has(x.remote.series)) { seriesDone.add(x.remote.series); Sync.removeSeries(x); } }
    else pushRemote(clone(x), null);
  }
  save(); render();
  toast(`Deleted all ${gone.size} “${o.ev.title}” events`, { undo: ![...gone].some(x => isSynced(x.calId)), icon: 'trash' });
}
function duplicateEditor() {
  pushUndo();
  const data = readForm();
  const copy = { id: uid(), exdates: [], doneDates: [], ...data, title: data.title + ' (copy)' };
  S.events.push(copy);
  pushRemote(null, copy);
  save(); $('#dlgEvent').close(); render();
  toast('Duplicated', { undo: !isSynced(copy.calId), icon: 'copy' });
}
function toggleDone(o) {
  if (!o || o.ro) return;
  pushUndo();
  const k = ymd(o.start), arr = (o.ev.doneDates ||= []);
  const i = arr.indexOf(k);
  if (i >= 0) arr.splice(i, 1); else arr.push(k);
  save(); render();
  if (i < 0) toast(`Done: ${o.ev.title} ✨`, { undo: true, icon: 'checkCircle' });
}

/* ---------------- calendars & groups ---------------- */
let editingCal = null, calColor = COLORS[0], calIcon = '';
function openCalDialog(cal, groupId) {
  editingCal = cal || null;
  calColor = cal?.color || COLORS[S.cals.length % COLORS.length];
  calIcon = cal?.icon ?? '📅';
  $('#calDlgTitle').textContent = cal ? 'Edit calendar' : 'New calendar';
  $('#calName').value = cal?.name || '';
  $('#calGroup').innerHTML = S.groups.map(g => `<option value="${g.id}">${esc(g.name)}</option>`).join('') + '<option value="__new">＋ New group…</option>';
  $('#calGroup').value = cal?.groupId || groupId || S.groups[0]?.id || '__new';
  $('#newGroupWrap').hidden = $('#calGroup').value !== '__new';
  $('#calNewGroup').value = '';
  $('#calDel').hidden = $('#calOnly').hidden = $('#calExport').hidden = !cal;
  drawCalPickers();
  $('#dlgCal').showModal();
  $('#calName').focus();
}
function drawCalPickers() {
  $('#calSw').innerHTML = COLORS.map(c => `<button type="button" style="--c:${c}" data-color="${c}" class="${c === calColor ? 'on' : ''}" title="${c}">${svg('check')}</button>`).join('')
    + `<input type="color" id="calCustom" value="${calColor}" title="Custom colour">`;
  $('#calIcons').innerHTML = CAL_EMOJIS.map(e => `<button type="button" data-emoji="${e}" class="${e === calIcon ? 'on' : ''}" title="${e ? '' : 'No icon'}">${e || '∅'}</button>`).join('');
}
function saveCal() {
  const name = $('#calName').value.trim();
  if (!name) { $('#calName').focus(); return false; }
  pushUndo();
  let gid = $('#calGroup').value;
  if (gid === '__new') {
    const gname = $('#calNewGroup').value.trim() || 'New group';
    gid = 'g-' + uid(); S.groups.push({ id: gid, name: gname });
  }
  if (editingCal) Object.assign(editingCal, { name, color: calColor, icon: calIcon, groupId: gid });
  else S.cals.push({ id: 'c-' + uid(), name, color: calColor, icon: calIcon, groupId: gid, visible: true });
  save(); render();
  toast(editingCal ? 'Calendar updated' : `Calendar “${name}” added`, { icon: 'calendar' });
  return true;
}
function deleteCal() {
  const c = editingCal;
  const n = S.events.filter(e => e.calId === c.id).length;
  if (c.source) {
    if (!confirm(`Stop showing “${c.name}” here? Nothing is deleted in your ${c.source.type === 'ics' ? 'subscription' : 'account'}.`)) return;
    Sync.forget(c);
  } else if (!confirm(`Delete “${c.name}”${n ? ` and its ${n} event${n > 1 ? 's' : ''}` : ''}?`)) return;
  pushUndo();
  S.cals = S.cals.filter(x => x !== c);
  S.events = S.events.filter(e => e.calId !== c.id);
  save(); $('#dlgCal').close(); render();
  toast(`Deleted “${c.name}”`, { undo: true, icon: 'trash' });
}
function editGroup(g) {
  const name = prompt('Rename group (leave empty to remove it — its calendars move to the first group):', g.name);
  if (name === null) return;
  pushUndo();
  if (name.trim()) g.name = name.trim();
  else {
    if (S.groups.length === 1) return toast('You need at least one group', { icon: 'x' });
    S.groups = S.groups.filter(x => x !== g);
    S.cals.forEach(c => { if (c.groupId === g.id) c.groupId = S.groups[0].id; });
  }
  save(); render();
}

/* ---------------- toasts ---------------- */
function toast(msg, opts = {}) {
  const el = document.createElement('div');
  el.className = 'toast glass';
  el.innerHTML = `${svg(opts.icon || 'check')}<span>${esc(msg)}</span>`;
  const kill = () => { el.classList.add('out'); setTimeout(() => el.remove(), 300); };
  if (opts.action) { const b = document.createElement('button'); b.textContent = opts.action[0]; b.onclick = () => { kill(); opts.action[1](); }; el.append(b); }
  if (opts.undo) { const b = document.createElement('button'); b.textContent = 'Undo'; b.onclick = () => { kill(); undo(); }; el.append(b); }
  $('#toasts').append(el);
  while ($('#toasts').children.length > 3) $('#toasts').firstChild.remove();
  setTimeout(kill, opts.ms || 4000);
}

/* ---------------- reminders ---------------- */
const FIRED_KEY = 'glasscal.fired';
let fired;
try { fired = new Set(JSON.parse(localStorage.getItem(FIRED_KEY)) || []); } catch { fired = new Set(); }
function checkReminders() {
  const now = new Date();
  for (const o of getOcc(addMin(now, -60), addDays(now, 8), { all: true })) {
    const r = o.ev.reminder;
    if (r == null || o.ro || o.ev.allDay) continue;
    const at = o.start - r * 6e4;
    if (now < at || now - at > 60 * 6e4 || now > o.end) continue;
    const k = `${o.key}|${r}`;
    if (fired.has(k)) continue;
    fired.add(k);
    const when = r === 0 ? 'Starting now' : o.start > now ? `Starts ${fmtTime(o.start)} (${Math.round((o.start - now) / 6e4)} min)` : `Started ${fmtTime(o.start)}`;
    const icon = eventIcon(o.ev, o.cal);
    toast(`${icon} ${o.ev.title} — ${when}`, { icon: 'bell', ms: 15000 });
    if ('Notification' in window && Notification.permission === 'granted') {
      try { new Notification(`${icon} ${o.ev.title}`, { body: `${when}${o.ev.location ? ' · ' + o.ev.location : ''}\n${o.cal.name}`, tag: k }); } catch { }
    }
  }
  const arr = [...fired].slice(-300);
  try { localStorage.setItem(FIRED_KEY, JSON.stringify(arr)); } catch { }
}
function updateBell() {
  const ok = 'Notification' in window && Notification.permission === 'granted';
  $('#bellBtn').classList.toggle('on', ok);
  $('#bellBtn').title = ok ? 'Desktop reminders are on' : 'Turn on desktop reminders';
}

/* ---------------- iCalendar import / export ---------------- */
const icsUnesc = s => s.replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1');
const icsEsc = s => String(s || '').replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/([,;])/g, '\\$1');
function icsDate(v, params) {
  v = v.trim();
  if (/VALUE=DATE(?!-)/.test(params) || /^\d{8}$/.test(v)) return { str: `${v.slice(0, 4)}-${v.slice(4, 6)}-${v.slice(6, 8)}T00:00`, allDay: true };
  const m = v.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/);
  if (!m) return null;
  const d = m[7] ? new Date(Date.UTC(+m[1], m[2] - 1, +m[3], +m[4], +m[5])) : new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5]);
  return { str: toStr(d), allDay: false };
}
function icsDuration(v) {
  const m = v.match(/P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?)?/);
  return m ? (+m[1] || 0) * 10080 + (+m[2] || 0) * 1440 + (+m[3] || 0) * 60 + (+m[4] || 0) : 60;
}
function parseICS(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\n[ \t]/g, '').split('\n');
  const raw = []; let c = null, inAlarm = false, calName = null;
  for (const line of lines) {
    if (line === 'BEGIN:VEVENT') { c = {}; continue; }
    if (line === 'END:VEVENT') { if (c) raw.push(c); c = null; continue; }
    if (line === 'BEGIN:VALARM') { inAlarm = true; continue; }
    if (line === 'END:VALARM') { inAlarm = false; continue; }
    const i = line.indexOf(':'); if (i < 0) continue;
    const [name, ...ps] = line.slice(0, i).split(';'); const params = ps.join(';').toUpperCase(); const val = line.slice(i + 1);
    if (!c) { if (name === 'X-WR-CALNAME') calName = icsUnesc(val); continue; }
    if (inAlarm) { if (name === 'TRIGGER' && c.reminder == null && val.startsWith('-')) c.reminder = icsDuration(val); continue; }
    switch (name) {
      case 'UID': c.uid = val; break;
      case 'SUMMARY': c.title = icsUnesc(val); break;
      case 'LOCATION': c.location = icsUnesc(val); break;
      case 'DESCRIPTION': c.notes = icsUnesc(val); break;
      case 'DTSTART': c.ds = icsDate(val, params); break;
      case 'DTEND': c.de = icsDate(val, params); break;
      case 'DURATION': c.dur = val; break;
      case 'RRULE': c.rrule = val; break;
      case 'EXDATE': (c.ex ||= []).push(...val.split(',').map(v => icsDate(v, params)?.str.slice(0, 10)).filter(Boolean)); break;
      case 'RECURRENCE-ID': c.recId = icsDate(val, params)?.str.slice(0, 10); break;
      case 'STATUS': c.status = val.trim(); break;
      case 'ATTENDEE': (c.att ||= []).push({
        email: val.replace(/^mailto:/i, '').trim(), name: (ps.join(';').match(/CN="?([^;":]*)"?/i) || [])[1] || '',
        status: { ACCEPTED: 'accepted', DECLINED: 'declined', TENTATIVE: 'tentative' }[(params.match(/PARTSTAT=([A-Z-]+)/) || [])[1]] || 'needsAction',
      }); break;
    }
  }
  const evs = [];
  for (const r of raw) {
    if (!r.ds) continue;
    const allDay = r.ds.allDay, s = parse(r.ds.str);
    let e = r.de ? parse(r.de.str) : r.dur ? addMin(s, icsDuration(r.dur)) : allDay ? addDays(s, 1) : addMin(s, 60);
    const end = allDay ? ymd(addDays(e, -1) < s ? s : addDays(e, -1)) + 'T00:00' : toStr(e);
    const ev = {
      id: uid(), uid: r.uid || uid(), title: r.title || '(No title)', start: r.ds.str, end, allDay,
      location: r.location || '', notes: r.notes || '', reminder: r.reminder ?? null, exdates: r.ex || [], doneDates: [],
      repeat: 'none', until: '', days: [], task: false, icon: '', attendees: r.att || [], _rec: r.recId, _cancel: r.status === 'CANCELLED',
    };
    if (r.rrule && !r.recId) Object.assign(ev, rruleToRepeat(r.rrule, s));
    evs.push(ev);
  }
  // Modified / cancelled single instances of a series become exceptions on the master.
  for (const ev of evs.filter(e => e._rec)) {
    const master = evs.find(x => x.uid === ev.uid && !x._rec);
    if (master) master.exdates.push(ev._rec);
    ev.uid = ev.uid + '#' + ev._rec;
  }
  return { name: calName, events: evs.filter(e => !e._cancel).map(({ _rec, _cancel, ...e }) => e) };
}
function rruleToRepeat(str, s) {
  const p = Object.fromEntries(str.split(';').map(x => x.split('=')));
  const iv = +(p.INTERVAL || 1);
  const byday = (p.BYDAY || '').split(',').filter(Boolean).map(d => ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'].indexOf(d.slice(-2)));
  let repeat = 'none', days = [];
  if (p.FREQ === 'DAILY') repeat = iv === 7 ? 'weekly' : iv === 14 ? 'biweekly' : 'daily';
  else if (p.FREQ === 'WEEKLY') {
    repeat = iv === 2 ? 'biweekly' : 'weekly';
    if (byday.length) days = byday;
    if (iv === 1 && byday.length === 5 && [1, 2, 3, 4, 5].every(d => byday.includes(d))) { repeat = 'weekdays'; days = []; }
  } else if (p.FREQ === 'MONTHLY') repeat = 'monthly';
  else if (p.FREQ === 'YEARLY') repeat = 'yearly';
  let until = '';
  if (p.UNTIL) until = icsDate(p.UNTIL, '')?.str.slice(0, 10) || '';
  else if (p.COUNT && repeat !== 'none') {
    // Walk the series to find the exact date of the last counted occurrence.
    const probe = { start: toStr(s), end: toStr(s), repeat, days };
    const list = occurrences(probe, s, new Date(s.getFullYear() + 30, 0, 1));
    const last = list[Math.min(+p.COUNT, list.length) - 1];
    if (last) until = ymd(last.start);
  }
  return { repeat, days, until };
}
/* iCalendar RRULE (without the "RRULE:" prefix) for a repeating event. */
function rruleOf(ev) {
  const DAY = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
  let r = { daily: 'FREQ=DAILY', weekdays: 'FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR', weekly: 'FREQ=WEEKLY', biweekly: 'FREQ=WEEKLY;INTERVAL=2', monthly: 'FREQ=MONTHLY', yearly: 'FREQ=YEARLY' }[ev.repeat];
  if ((ev.repeat === 'weekly' || ev.repeat === 'biweekly') && ev.days?.length) r += ';BYDAY=' + ev.days.map(d => DAY[d]).join(',');
  if (ev.until) r += `;UNTIL=${ev.until.replace(/-/g, '')}T235959`;
  return r;
}
function toICS(events, name) {
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Glass Calendar//EN', 'CALSCALE:GREGORIAN', `X-WR-CALNAME:${icsEsc(name)}`];
  const dt = (s, allDay) => allDay ? `;VALUE=DATE:${s.slice(0, 10).replace(/-/g, '')}` : `:${s.replace(/[-:]/g, '')}00`;
  const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15) + 'Z';
  for (const ev of events) {
    L.push('BEGIN:VEVENT', `UID:${ev.uid || ev.id + '@glasscal'}`, `DTSTAMP:${stamp}`, 'DTSTART' + dt(ev.start, ev.allDay),
      'DTEND' + (ev.allDay ? dt(ymd(addDays(parse(ev.end), 1)), true) : dt(ev.end)), `SUMMARY:${icsEsc(ev.title)}`);
    if (ev.location) L.push(`LOCATION:${icsEsc(ev.location)}`);
    if (ev.notes) L.push(`DESCRIPTION:${icsEsc(ev.notes)}`);
    if (isRecurring(ev)) {
      L.push('RRULE:' + rruleOf(ev));
      for (const x of ev.exdates || []) L.push(ev.allDay ? `EXDATE;VALUE=DATE:${x.replace(/-/g, '')}` : `EXDATE:${x.replace(/-/g, '')}T${ev.start.slice(11).replace(':', '')}00`);
    }
    for (const a of ev.attendees || []) {
      const ps = { accepted: 'ACCEPTED', declined: 'DECLINED', tentative: 'TENTATIVE' }[a.status] || 'NEEDS-ACTION';
      L.push(`ATTENDEE;${a.name ? `CN="${a.name.replace(/"/g, '')}";` : ''}PARTSTAT=${ps}:mailto:${a.email}`);
    }
    if (ev.reminder != null) L.push('BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:${icsEsc(ev.title)}`, `TRIGGER:-PT${ev.reminder}M`, 'END:VALARM');
    L.push('END:VEVENT');
  }
  L.push('END:VCALENDAR');
  return L.join('\r\n');
}
function download(name, text, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type }));
  a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
async function doImport() {
  const files = [...$('#importFile').files];
  if (!files.length) { toast('Choose at least one .ics file', { icon: 'upload' }); return false; }
  pushUndo();
  let added = 0, updated = 0;
  for (const f of files) {
    const { name, events } = parseICS(await f.text());
    let target = $('#importTarget').value;
    if (target === '__new') {
      const nm = name || f.name.replace(/\.ics$/i, '').replace(/[_-]+/g, ' ');
      let g = S.groups.find(g => g.name === 'Imported');
      if (!g) { g = { id: 'g-' + uid(), name: 'Imported' }; S.groups.push(g); }
      target = 'c-' + uid();
      S.cals.push({ id: target, name: nm, color: COLORS[(S.cals.length * 3) % COLORS.length], icon: '📅', groupId: g.id, visible: true });
    }
    for (const ev of events) {
      ev.calId = target;
      const i = S.events.findIndex(x => x.uid && x.uid === ev.uid && x.calId === target);
      if (i >= 0) { ev.id = S.events[i].id; ev.doneDates = S.events[i].doneDates || []; ev.icon = S.events[i].icon || ''; S.events[i] = ev; updated++; }
      else { S.events.push(ev); added++; }
    }
  }
  save(); render();
  toast(`Imported ${added} new${updated ? `, updated ${updated}` : ''} event${added === 1 ? '' : 's'}`, { undo: true, icon: 'upload', ms: 6000 });
  return true;
}

/* ---------------- settings ---------------- */
function openSettings() {
  const s = S.settings;
  $('#setBg').innerHTML = Object.entries(BGS).map(([k, n]) => `<button type="button" data-bg="${k}" class="${s.bg === k ? 'on' : ''}" style="background:${BG_PREVIEW[k]}">${n}</button>`).join('');
  $('#setLook').value = s.look;
  $('#setTheme').value = s.theme; $('#setWeek').value = s.weekStart;
  $('#setCal').innerHTML = S.cals.filter(c => !c.readonly).map(c => `<option value="${c.id}">${c.icon || ''} ${esc(c.name)}</option>`).join('');
  $('#setCal').value = s.defaultCal;
  $('#setRem').value = s.defaultReminder ?? '';
  $('#setWS').value = s.workStart; $('#setWE').value = s.workEnd;
  $('#setH24').checked = s.h24; $('#setWk').checked = s.weekNums; $('#setSmart').checked = s.smartIcons;
  $('#dlgSettings').showModal();
}
function readSettings() {
  const s = S.settings;
  s.look = $('#setLook').value;
  s.theme = $('#setTheme').value; s.weekStart = +$('#setWeek').value;
  s.defaultCal = $('#setCal').value; s.defaultReminder = $('#setRem').value === '' ? null : +$('#setRem').value;
  s.workStart = clamp(+$('#setWS').value || 0, 0, 23); s.workEnd = clamp(+$('#setWE').value || 24, s.workStart + 1, 24);
  s.h24 = $('#setH24').checked; s.weekNums = $('#setWk').checked; s.smartIcons = $('#setSmart').checked;
  save(); render();
}

/* ---------------- wiring ---------------- */
/* Time boxes in the event editor: show 12-hour times (2:00pm) whatever the computer's regional settings are.
   Each box still hands the rest of the app a plain "HH:MM" value. Accepts 2pm, 2:30 pm, 14:00, 1430. */
function parseTime(s) {
  const m = String(s || '').trim().toLowerCase().replace(/\./g, '').match(/^(\d{1,2})(?::?(\d{2}))?\s*(am|pm|a|p)?$/);
  if (!m) return null;
  let h = +m[1]; const mi = +(m[2] || 0), ap = m[3];
  if (mi > 59) return null;
  if (ap) { if (h < 1 || h > 12) return null; if (ap[0] === 'p' && h < 12) h += 12; if (ap[0] === 'a' && h === 12) h = 0; }
  else if (h > 23) return null;
  return `${pad(h)}:${pad(mi)}`;
}
function showTime(v) {
  if (!v) return '';
  const [h, m] = v.split(':').map(Number);
  return S.settings.h24 ? v : `${h % 12 || 12}:${pad(m)}${h < 12 ? 'am' : 'pm'}`;
}
function initTimeBoxes() {
  const nativeGet = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').get;
  const nativeSet = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  $$('.t12').forEach(el => {
    let v = '';
    Object.defineProperty(el, 'value', {
      get() { const p = parseTime(nativeGet.call(el)); if (p) v = p; else if (!nativeGet.call(el).trim()) v = ''; return v; }, // reads what you typed, even before you leave the box
      set(x) { v = x || ''; nativeSet.call(el, showTime(v)); },
    });
    el.addEventListener('change', () => { el.value = el.value; }); // tidy what you typed into "2:00pm" (or put the old time back if it wasn't a time)
  });
  $('#times12').innerHTML = [...Array(48)].map((_, i) => `<option value="${showTime(`${pad(Math.floor(i / 2))}:${i % 2 ? '30' : '00'}`)}">`).join('');
}

/* Date boxes in the event editor: show dates as "19 Oct 2026" whatever the computer's regional settings are.
   Each box still hands the rest of the app a plain "YYYY-MM-DD" value. Accepts 19 Oct 2026, 19/10/2026,
   2026-10-19, "19 oct", "oct 19", today, tomorrow. A small calendar button opens a picker. */
const MON3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function parseDateText(s, fallbackYear) {
  s = String(s || '').trim().toLowerCase().replace(/,/g, '');
  if (s === 'today') return ymd(new Date());
  if (s === 'tomorrow') return ymd(addDays(new Date(), 1));
  let y, m, d, x;
  const yr = v => (v == null || v === '' ? fallbackYear : (+v < 100 ? 2000 + +v : +v));
  if ((x = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/))) { y = +x[1]; m = +x[2]; d = +x[3]; }
  else if ((x = s.match(/^(\d{1,2})[-/.\s](\d{1,2})[-/.\s](\d{2,4})$/))) { d = +x[1]; m = +x[2]; y = yr(x[3]); }
  else if ((x = s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s*([a-z]{3,})\.?(?:\s+(\d{2,4}))?$/))) { d = +x[1]; m = MON3.findIndex(n => n.toLowerCase() === x[2].slice(0, 3)) + 1; y = yr(x[3]); }
  else if ((x = s.match(/^([a-z]{3,})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:\s+(\d{2,4}))?$/))) { m = MON3.findIndex(n => n.toLowerCase() === x[1].slice(0, 3)) + 1; d = +x[2]; y = yr(x[3]); }
  else return null;
  if (!m || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const t = new Date(y, m - 1, d);
  return t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === d ? ymd(t) : null;
}
const showDate = v => { if (!v) return ''; const [y, m, d] = v.split('-').map(Number); return `${d} ${MON3[m - 1]} ${y}`; };
function initDateBoxes() {
  const proto = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
  $$('.dfmt').forEach(el => {
    let v = '';
    Object.defineProperty(el, 'value', {
      get() { const raw = proto.get.call(el); const p = parseDateText(raw, v ? +v.slice(0, 4) : new Date().getFullYear()); if (p) v = p; else if (!raw.trim()) v = ''; return v; }, // reads what you typed, even before you leave the box
      set(x) { v = x || ''; proto.set.call(el, showDate(v)); },
    });
    el.addEventListener('change', () => { el.value = el.value; }); // tidy it into "19 Oct 2026" (or restore the old date if it wasn't a date)
    // small calendar button -> the browser's own date picker
    const pk = document.createElement('input'); pk.type = 'date'; pk.tabIndex = -1; pk.className = 'pk-hidden';
    const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'datepick'; btn.title = 'Pick a date'; btn.innerHTML = svg('calendar');
    btn.onclick = () => { pk.value = el.value || ymd(new Date()); try { pk.showPicker(); } catch { pk.focus(); pk.click(); } };
    pk.onchange = () => { if (pk.value) { el.value = pk.value; el.dispatchEvent(new Event('change', { bubbles: true })); } };
    el.after(btn, pk);
  });
}

function wire() {
  hydrateIcons();
  initTimeBoxes();
  initDateBoxes();
  $('#todayBtn').onclick = goToday;
  $('#backBtn').onclick = goBack;
  $('#newBtn').onclick = () => newOnDay(ymd(cur));
  $('#viewSeg').onclick = e => { const b = e.target.closest('[data-view]'); if (b) setView(b.dataset.view); };
  $('#themeBtn').onclick = () => { S.settings.theme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; save(); applyTheme(); };
  $('#settingsBtn').onclick = openSettings;
  $('#helpBtn').onclick = () => $('#dlgKeys').showModal();
  $('#menuBtn').onclick = () => {
    const app = $('#app');
    if (innerWidth <= 900) app.classList.toggle('side-open'); else app.classList.toggle('no-side');
    setTimeout(render, 260);
  };
  $('#bellBtn').onclick = async () => {
    if (!('Notification' in window)) return toast('This browser does not support notifications', { icon: 'bell' });
    if (Notification.permission === 'granted') return toast('Desktop reminders are on. Keep this tab open (pin it!) to receive them.', { icon: 'bell', ms: 6000 });
    const p = await Notification.requestPermission();
    updateBell();
    toast(p === 'granted' ? 'Desktop reminders turned on 🔔' : 'Reminders will show inside the calendar instead', { icon: 'bell' });
  };
  mq.addEventListener('change', () => { if (S.settings.theme === 'auto') applyTheme(); });

  // search
  let st;
  $('#searchInput').addEventListener('input', e => { clearTimeout(st); st = setTimeout(() => { searchQ = e.target.value; render(); }, 150); });
  $('#searchInput').addEventListener('keydown', e => { if (e.key === 'Enter' && searchQ) setView('agenda'); });

  // mini calendar + sidebar
  $('#sidebar').addEventListener('click', e => {
    const m = e.target.closest('[data-mini]');
    if (m) { miniCur = new Date(miniCur.getFullYear(), miniCur.getMonth() + +m.dataset.mini, 1); renderMini(); return; }
    const p = e.target.closest('[data-pick]');
    if (p) { cur = parse(p.dataset.pick); if (view === 'year' || view === 'agenda') view = 'day'; miniCur = new Date(cur); gridScroll = null; render(); $('#app').classList.remove('side-open'); return; }
    const unTick = e.target.closest('.upnext .tick');
    if (unTick) { toggleDone(occMap.get(unTick.closest('[data-key]').dataset.key)); return; }
    const un = e.target.closest('.un[data-key]');
    if (un) { const o = occMap.get(un.dataset.key); if (o) openPeek(o, un); return; }
    const act = e.target.closest('[data-act]'); if (!act) return;
    const g = S.groups.find(x => x.id === act.closest('[data-g]')?.dataset.g);
    const c = S.cals.find(x => x.id === act.closest('[data-c]')?.dataset.c);
    switch (act.dataset.act) {
      case 'toggleGroup': g.collapsed = !g.collapsed; save(); renderSidebar(); break;
      case 'addCal': openCalDialog(null, g.id); break;
      case 'editGroup': editGroup(g); break;
      case 'toggleCal': c.visible = !c.visible; save(); render(); break;
      case 'editCal': openCalDialog(c); break;
      case 'syncSettings': Sync.openDialog(); break;
    }
  });
  // Drag calendars to reorder them or move them into another group (drop on a row, or on a group's name).
  const side = $('#calList');
  let dragCal = null;
  const clearMarks = () => $$('.drop-before, .drop-after, .drop-into', side).forEach(x => x.classList.remove('drop-before', 'drop-after', 'drop-into'));
  const above = (row, y) => { const r = row.getBoundingClientRect(); return y < r.top + r.height / 2; };
  let dragGroup = null;
  side.addEventListener('dragstart', e => {
    const head = e.target.closest('.grp-h');
    if (head) { // dragging a whole folder
      const grp = head.closest('.grp');
      dragGroup = grp.dataset.g;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', dragGroup);
      requestAnimationFrame(() => grp.classList.add('dragging'));
      return;
    }
    const row = e.target.closest('.cal'); if (!row) return;
    dragCal = row.dataset.c;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', dragCal);
    requestAnimationFrame(() => row.classList.add('dragging'));
  });
  side.addEventListener('dragover', e => {
    if (dragGroup) {
      const grp = e.target.closest('.grp'); if (!grp) return;
      e.preventDefault(); clearMarks();
      if (grp.dataset.g !== dragGroup) grp.classList.add(above(grp, e.clientY) ? 'drop-before' : 'drop-after');
      return;
    }
    if (!dragCal) return;
    const row = e.target.closest('.cal'), grp = e.target.closest('.grp');
    if (!grp) return;
    e.preventDefault();
    clearMarks();
    if (row && row.dataset.c !== dragCal) row.classList.add(above(row, e.clientY) ? 'drop-before' : 'drop-after');
    else if (!row) grp.classList.add('drop-into');
  });
  side.addEventListener('dragleave', e => { if (!side.contains(e.relatedTarget)) clearMarks(); });
  side.addEventListener('drop', e => {
    if (dragGroup) {
      const target = e.target.closest('.grp'), moving = S.groups.find(g => g.id === dragGroup);
      dragGroup = null; clearMarks();
      if (!target || !moving || target.dataset.g === moving.id) return;
      e.preventDefault();
      pushUndo();
      S.groups = S.groups.filter(g => g !== moving);
      const i = S.groups.findIndex(g => g.id === target.dataset.g);
      S.groups.splice(above(target, e.clientY) ? i : i + 1, 0, moving);
      save(); render();
      return;
    }
    const row = e.target.closest('.cal'), grp = e.target.closest('.grp');
    const cal = S.cals.find(c => c.id === dragCal);
    dragCal = null; clearMarks();
    if (!cal || !grp || row?.dataset.c === cal.id) return;
    e.preventDefault();
    pushUndo();
    const fromGroup = cal.groupId;
    S.cals = S.cals.filter(c => c !== cal);
    cal.groupId = grp.dataset.g;
    if (row) {
      const i = S.cals.findIndex(c => c.id === row.dataset.c);
      S.cals.splice(above(row, e.clientY) ? i : i + 1, 0, cal);
    } else {
      let last = -1;
      S.cals.forEach((c, i) => { if (c.groupId === cal.groupId) last = i; });
      S.cals.splice(last + 1, 0, cal);
    }
    save(); render();
    if (fromGroup !== cal.groupId) toast(`Moved “${cal.name}” to ${S.groups.find(g => g.id === cal.groupId)?.name}`, { undo: true, icon: 'calendar' });
  });
  side.addEventListener('dragend', () => { dragCal = dragGroup = null; clearMarks(); $$('.dragging', side).forEach(x => x.classList.remove('dragging')); });

  $('#addCalBtn').onclick = () => openCalDialog();
  $('#addGroupBtn').onclick = () => {
    const name = prompt('Name for the new group or account (e.g. Work, Family, school email):');
    if (!name?.trim()) return;
    pushUndo(); S.groups.push({ id: 'g-' + uid(), name: name.trim() }); save(); render();
    toast(`Group “${name.trim()}” added — use its ＋ to add calendars`, { icon: 'folderPlus' });
  };
  $('#importBtn').onclick = openImport;

  // main view clicks
  const v = $('#view');
  v.addEventListener('click', e => {
    const nav = e.target.closest('[data-nav]');
    if (nav) { go(+nav.dataset.nav); return; }
    const tick = e.target.closest('.tick');
    if (tick) { e.stopPropagation(); toggleDone(occMap.get(tick.closest('[data-key]').dataset.key)); return; }
    const act = e.target.closest('[data-act]');
    if (act?.dataset.act === 'clearSearch') { searchQ = ''; $('#searchInput').value = ''; render(); return; }
    if (act?.dataset.act === 'listResults') { setView('agenda'); return; }
    const ch = e.target.closest('.chip, .ag-item, .hol-lbl');
    if (ch) { const o = occMap.get(ch.dataset.key); if (o) openPeek(o, ch); return; }
    const cake = e.target.closest('.bday-btn, .bday-chip');
    if (cake) { showDayPop(cake.dataset.bday, cake, true); return; }
    const more = e.target.closest('.more-btn');
    if (more) { showDayPop(more.dataset.day, more); return; }
    const gt = e.target.closest('[data-goto]');
    if (gt) { cur = parse(gt.dataset.goto); miniCur = new Date(cur); setView('day'); return; }
    const mo = e.target.closest('[data-month]');
    if (mo) { cur = new Date(cur.getFullYear(), +mo.dataset.month, 1); miniCur = new Date(cur); setView('month'); return; }
    const add = e.target.closest('.cell-add');
    if (add) { newOnDay(add.closest('[data-day]').dataset.day); return; }
  });
  v.addEventListener('dblclick', e => {
    if (e.target.closest('.chip, .more-btn, .dnum')) return;
    const cell = e.target.closest('.cell, .adc');
    if (cell) newOnDay(cell.dataset.day, cell.classList.contains('adc'));
  });
  // drag & drop between days (month + all-day row)
  let dragKey = null;
  v.addEventListener('dragstart', e => {
    const c = e.target.closest('.chip'); if (!c) return;
    dragKey = c.dataset.key; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', dragKey);
    hidePop();
  });
  v.addEventListener('dragover', e => {
    const cell = e.target.closest('.cell, .adc'); if (!cell || !dragKey) return;
    e.preventDefault();
    $$('.over', v).forEach(x => x !== cell && x.classList.remove('over'));
    cell.classList.add('over');
  });
  v.addEventListener('dragleave', e => { const cell = e.target.closest('.cell, .adc'); if (cell && !cell.contains(e.relatedTarget)) cell.classList.remove('over'); });
  v.addEventListener('drop', e => {
    const cell = e.target.closest('.cell, .adc'); if (!cell || !dragKey) return;
    e.preventDefault();
    const o = occMap.get(dragKey); dragKey = null;
    if (!o) return;
    const delta = dayDiff(parse(cell.dataset.day), o.start);
    if (delta) moveOcc(o, addDays(o.start, delta), addDays(o.end, delta)); else render();
  });
  v.addEventListener('dragend', () => { dragKey = null; $$('.over', v).forEach(x => x.classList.remove('over')); });

  // popover
  $('#pop').addEventListener('click', e => {
    if (e.target.closest('[data-act=closePop]')) return hidePop();
    const tick = e.target.closest('.tick');
    if (tick) { toggleDone(occMap.get(tick.closest('[data-key]').dataset.key)); hidePop(); return; }
    const ch = e.target.closest('.chip, .bday-row'); if (ch) openPeek(occMap.get(ch.dataset.key), ch);
  });
  document.addEventListener('pointerdown', e => {
    if (!$('#pop').hidden && !e.target.closest('#pop, .more-btn, .bday-btn, .bday-chip')) hidePop();
    if (!$('#peek').hidden && !e.target.closest('#peek')) hidePeek();
  });
  $('#peek').addEventListener('click', onPeekClick);
  // Right-click (or double-click) an event → straight to the editor.
  const editFrom = e => {
    const el = e.target.closest('#view [data-key], #pop [data-key], #upnext [data-key]');
    if (!el || e.target.closest('.tick')) return false;
    const o = occMap.get(el.dataset.key); if (!o) return false;
    e.preventDefault(); hidePop(); hidePeek(); openEditor(o);
    return true;
  };
  document.addEventListener('contextmenu', editFrom);
  document.addEventListener('dblclick', e => { if (!e.target.closest('dialog')) editFrom(e); });

  // dialogs: generic close buttons + click on backdrop closes
  $$('dialog').forEach(d => {
    // Chrome returns focus to the last focused input on close, which would swallow shortcut keys.
    d.addEventListener('close', () => setTimeout(() => document.activeElement?.matches('input') && document.activeElement.blur()));
    d.addEventListener('click', e => {
      if (e.target.closest('[data-close]')) { d.close(); return; }
      if (e.target === d) { const r = d.getBoundingClientRect(); if (e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) d.close(); }
    });
  });

  // event editor
  $('#evForm').onsubmit = e => { e.preventDefault(); saveEditor(); };
  $('#evDel').onclick = () => {
    const { o, related = [] } = editing;
    if (isRecurring(o.ev) || related.length) { $('#evDelConfirm').hidden = false; $('#evDelThis').focus(); }
    else deleteEditor(false);
  };
  $('#evDelThis').onclick = () => deleteEditor(false);
  $('#evDelAll').onclick = () => deleteEditor(true);
  $('#evDelCancel').onclick = () => { $('#evDelConfirm').hidden = true; };
  $('#evSaveAll').onclick = () => saveEditor(true);
  // guests
  $('#evGuestIn').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ',' || e.key === ';' || (e.key === 'Tab' && e.target.value.trim())) { e.preventDefault(); addGuestFromInput(); }
    if (e.key === 'Backspace' && !e.target.value && editGuests.length) { editGuests.pop(); drawGuests(); }
  });
  $('#evGuestIn').addEventListener('change', () => { if (EMAIL_RE.test($('#evGuestIn').value.trim())) addGuestFromInput(); });
  $('#evGuestList').addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (b) { editGuests.splice(+b.dataset.rm, 1); drawGuests(); } });
  $('#evCal').addEventListener('change', drawGuests);
  $('#evDup').onclick = duplicateEditor;
  ['#evAll', '#evRep', '#evCal'].forEach(s => $(s).addEventListener('change', syncEditor));
  $('#evTitle').addEventListener('input', () => { if (!editIcon) syncEditor(); });
  $('#evAll').addEventListener('change', () => { if (!$('#evAll').checked && $('#evRem').value === '') $('#evRem').value = S.settings.defaultReminder ?? ''; });
  let lastStart = null;
  const keepDuration = () => {
    const s = parse(`${$('#evSD').value}T${$('#evST').value || '00:00'}`);
    if (lastStart) {
      const e = parse(`${$('#evED').value}T${$('#evET').value || '00:00'}`);
      const ne = new Date(+e + (s - lastStart));
      $('#evED').value = ymd(ne); $('#evET').value = hm(ne);
    }
    lastStart = s;
  };
  ['#evSD', '#evST'].forEach(s => {
    $(s).addEventListener('focus', () => lastStart = parse(`${$('#evSD').value}T${$('#evST').value || '00:00'}`));
    $(s).addEventListener('change', keepDuration);
  });
  $('#evDays').addEventListener('click', e => {
    const b = e.target.closest('[data-d]'); if (!b) return;
    const d = +b.dataset.d;
    editDays = editDays.includes(d) ? editDays.filter(x => x !== d) : [...editDays, d];
    if (!editDays.length) editDays = [d];
    syncEditor();
  });
  $('#evIconBtn').onclick = () => {
    const p = $('#evIconPop');
    if (!p.hidden) { p.hidden = true; return; }
    p.innerHTML = `<button type="button" class="txt" data-em="">✨ Auto</button><button type="button" class="txt" data-em="none">🚫 None</button>` + EMOJIS.map(e => `<button type="button" data-em="${e}">${e}</button>`).join('');
    p.hidden = false;
  };
  $('#evIconPop').onclick = e => { const b = e.target.closest('[data-em]'); if (!b) return; editIcon = b.dataset.em; $('#evIconPop').hidden = true; syncEditor(); };

  // calendar dialog
  $('#calForm').onsubmit = e => { e.preventDefault(); if (saveCal()) $('#dlgCal').close(); };
  $('#calSw').addEventListener('click', e => { const b = e.target.closest('[data-color]'); if (b) { calColor = b.dataset.color; drawCalPickers(); } });
  $('#calSw').addEventListener('input', e => { if (e.target.id === 'calCustom') { calColor = e.target.value; $$('#calSw [data-color]').forEach(b => b.classList.remove('on')); } });
  $('#calIcons').addEventListener('click', e => { const b = e.target.closest('[data-emoji]'); if (b) { calIcon = b.dataset.emoji; drawCalPickers(); } });
  $('#calGroup').onchange = () => $('#newGroupWrap').hidden = $('#calGroup').value !== '__new';
  $('#calDel').onclick = deleteCal;
  $('#calOnly').onclick = () => { S.cals.forEach(c => c.visible = c === editingCal); save(); $('#dlgCal').close(); render(); toast(`Showing only “${editingCal.name}”`, { icon: 'eye' }); };
  $('#calExport').onclick = () => { const c = editingCal; download(`${c.name}.ics`, toICS(S.events.filter(e => e.calId === c.id), c.name), 'text/calendar'); };

  // import dialog
  $('#importForm').onsubmit = async e => { e.preventDefault(); if (await doImport()) $('#dlgImport').close(); };

  // settings dialog
  $('#setForm').addEventListener('change', e => { if (e.target.id !== 'bkFile') readSettings(); });
  $('#setBg').addEventListener('click', e => {
    const b = e.target.closest('[data-bg]'); if (!b) return;
    S.settings.bg = b.dataset.bg; save(); applyTheme();
    $$('#setBg button').forEach(x => x.classList.toggle('on', x === b));
  });
  $('#bkExport').onclick = () => download(`glass-calendar-backup-${ymd(new Date())}.json`, JSON.stringify(S, null, 1), 'application/json');
  $('#bkImport').onclick = () => $('#bkFile').click();
  $('#bkFile').onchange = async e => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const d = JSON.parse(await f.text());
      if (!d.cals || !d.events) throw 0;
      if (!confirm('Replace everything in this calendar with the backup?')) return;
      pushUndo(); Object.assign(S, d); S.settings = { ...seed().settings, ...d.settings };
      save(); render(); toast('Backup restored', { undo: true, icon: 'upload' });
    } catch { toast('That file is not a Glass Calendar backup', { icon: 'x' }); }
    e.target.value = '';
  };
  $('#icsExportAll').onclick = () => download('glass-calendar.ics', toICS(S.events, 'Glass Calendar'), 'text/calendar');
  $('#clearAll').onclick = () => {
    if (!confirm('Delete ALL events from every calendar? (Calendars are kept. You can Undo right after.)')) return;
    pushUndo(); S.events = []; save(); render(); toast('All events deleted', { undo: true, icon: 'trash' });
  };

  // keyboard
  document.addEventListener('keydown', e => {
    if (e.target.matches('input, textarea, select')) {
      if (e.key === 'Escape' && !$('dialog[open]')) { e.target.blur(); if (e.target.id === 'searchInput') { e.target.value = ''; searchQ = ''; render(); } }
      return;
    }
    if ($('dialog[open]')) return;
    if (peekOcc && !$('#peek').hidden) {
      if (e.key === 'Escape') return hidePeek();
      if (e.key === 'e' && !peekOcc.ro) { const o = peekOcc; hidePeek(); return openEditor(o); }
      if (e.key === 'Delete' && !peekOcc.ro) return $('[data-pk=del]', $('#peek'))?.click();
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); return; }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const k = e.key;
    const map = { d: 'day', x: 'workweek', w: 'week', m: 'month', y: 'year', a: 'agenda' };
    if (map[k.toLowerCase()]) return setView(map[k.toLowerCase()]);
    switch (k) {
      case 't': case 'T': goToday(); break;
      case 'Backspace': e.preventDefault(); goBack(); break;
      case 'ArrowLeft': case 'k': go(-1); break;
      case 'ArrowRight': case 'j': go(1); break;
      case 'n': case 'N': case 'c': e.preventDefault(); newOnDay(ymd(cur)); break;
      case '/': e.preventDefault(); $('#searchInput').focus(); break;
      case 'l': $('#themeBtn').click(); break;
      case '?': $('#dlgKeys').showModal(); break;
      case 'Escape': hidePop(); break;
    }
  });

  let rt;
  addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(render, 120); });
  // storage changes from another tab
  addEventListener('storage', e => { if (e.key === STORE_KEY && e.newValue) { Object.assign(S, JSON.parse(e.newValue)); render(); } });
}
function openImport() {
  $('#importTarget').innerHTML = '<option value="__new">＋ New calendar (named after the file)</option>' +
    S.cals.filter(c => !c.readonly && !c.source).map(c => `<option value="${c.id}">${c.icon || ''} ${esc(c.name)}</option>`).join('');
  $('#importFile').value = '';
  $('#dlgImport').showModal();
}

/* ---------------- tick ---------------- */
function tick() {
  const now = new Date();
  if (ymd(now) !== brandDay) { render(); return; }
  $$('.now').forEach(el => { const col = el.parentElement; el.style.top = (now - parse(col.dataset.day)) / 6e4 * HOUR / 60 + 'px'; });
  renderUpNext();
  checkReminders();
}

wire();
applyTheme();
render();
updateBell();
if (window.__samplesRemovedNote != null) {
  save();
  toast(`Removed the sample calendars (${window.__samplesRemovedNote} events). Your Google calendars are untouched.`, {
    icon: 'trash', ms: 15000,
    action: ['Restore', () => {
      try {
        const b = JSON.parse(localStorage.getItem('glasscal.backup-before-removing-samples'));
        Object.assign(S, b); S.settings.samplesRemoved = true; save(); render(); toast('Sample calendars restored', { icon: 'repeat' });
      } catch { toast('Could not find the backup', { icon: 'alert' }); }
    }],
  });
}
checkReminders();
setInterval(tick, 30000);

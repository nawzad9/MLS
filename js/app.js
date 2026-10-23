/**
 * MLS Course Registration — web app
 * University of Raparin — College of Science — Department of Medical Laboratory Science
 */
'use strict';

/* ═══ 1. Config ═══════════════════════════════════════════════════════════ */
const CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbxh7NpQ0YTCZkvdNOmxBtNvP51syCttL8s_jAXgi4IWxv2cYpnijilrHViJ6ar7_xfK/exec',   // Apps Script → Deploy → Web app → URL ending in /exec
  SHORT_NAME: 'MLS',
  DOCX_LIB: 'vendor/docx.iife.js',
  EXCEL_LIB: 'vendor/exceljs.min.js'
};

const TYPES = ['REGULAR', 'RETAKE', 'EXAM_RETAKE'];
const STATUSES = ['PENDING', 'APPROVED', 'REJECTED'];
const ISSUE_ORDER = ['UNIDENTIFIED', 'MULTI_SUBMIT', 'SAME_NAME', 'EMAIL_DOMAIN', 'CODE_MISMATCH',
                     'NO_SUBJECTS', 'OVER_ECTS', 'PREREQ', 'NOTE', 'NO_TEACHER'];

/* ═══ 2. Small helpers ════════════════════════════════════════════════════ */
const $ = sel => document.querySelector(sel);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => Number(v) || 0;
const numOr = (v, d) => (v === '' || v == null || !isFinite(Number(v)) ? d : Number(v));
const round2 = v => Math.round(v * 100) / 100;
const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const isStandalone = () => !!(navigator.standalone || (window.matchMedia && matchMedia('(display-mode: standalone)').matches));

function store(key, value) {
  try {
    if (value === undefined) return localStorage.getItem('mls_' + key);
    if (value === null) localStorage.removeItem('mls_' + key);
    else localStorage.setItem('mls_' + key, value);
  } catch (e) { return null; }
}

function t(key, ...args) {
  const v = I18N[S.lang][key] ?? I18N.en[key] ?? key;
  return typeof v === 'function' ? v(...args) : v;
}

const stageName = n => t('stage' + num(n));
const groupName = parallel => (parallel === 'Yes' ? t('Parallel') : t('Morning'));
const locale = () => (S.lang === 'ku' ? 'ckb-IQ' : 'en-GB');
function fmtDate(iso, withTime = true) {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleString(locale(), withTime ? { dateStyle: 'medium', timeStyle: 'short' } : { dateStyle: 'medium' });
}
const today = () => new Date().toISOString().slice(0, 10);
const sep = () => (S.lang === 'ku' ? '، ' : ', ');

/** Normalises Kurdish/Arabic spelling variants so the same name typed twice matches. */
const normName = n => String(n || '').replace(/[\s\u200c\u200f\u0640]/g, '').replace(/[يى]/g, 'ی').replace(/ك/g, 'ک').replace(/ة/g, 'ە');

/* ═══ 3. State ════════════════════════════════════════════════════════════ */
const S = {
  lang: store('lang') || 'ku',
  token: store('token'),
  user: JSON.parse(store('user') || 'null'),
  tab: '',
  data: null,        // head: raw data from the server
  m: null,           // head: derived model (see buildModel)
  tdata: null,       // teacher data
  semester: '',
  listFn: null,      // renders the #list part of the current view (for live filtering)
  openStudent: null,
  installEvent: null,
  updateWorker: null,
  reloadOnUpdate: false,
  saveState: 'idle',
  f: { q: '', stage: '', flagged: false, pending: true, subject: '', status: '', type: '', group: '', showChecked: false, sort: 'name', result: '', unlock: false }
};

/* ═══ 4. API ══════════════════════════════════════════════════════════════ */
async function api(action, payload = {}) {
  if (!CONFIG.API_URL || CONFIG.API_URL.startsWith('PASTE')) throw new Error('CONFIG');
  let res;
  try {
    const r = await fetch(CONFIG.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // avoids a CORS preflight
      body: JSON.stringify({ action, token: S.token, payload })
    });
    res = await r.json();
  } catch (e) {
    throw new Error('NETWORK');
  }
  if (res.renew && S.token) { S.token = res.renew; store('token', res.renew); }   // sessions renew themselves quietly
  if (!res.ok) {
    if (res.error === 'AUTH_REQUIRED' && action !== 'login') signOut(true);
    throw new Error(res.error);
  }
  return res.data;
}

function errText(e) {
  const code = String(e?.message || e);
  return I18N[S.lang]['E_' + code] ? t('E_' + code) : `${t('E_GENERIC')}: ${code}`;
}

let toastTimer;
function toast(msg, bad = false) {
  $('#toast').innerHTML = `<div class="${bad ? 'bad' : ''}" role="status">${esc(msg)}</div>`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { $('#toast').innerHTML = ''; }, 3400);
}

/* ═══ 5. Head model: joins, ECTS totals, flags and data issues ════════════ */
function buildModel(D) {
  const max = num(D.settings.MaxECTS);
  const subj = Object.fromEntries(D.subjects.map(s => [s.Code, s]));
  const scores = Object.fromEntries((D.scores || []).map(x => [x.RegID, x]));
  const users = Object.fromEntries(D.users.map(u => [u.UserID, u]));
  const stu = {};
  D.students.forEach(s => {
    stu[s.Code] = Object.assign(s, {
      regs: [], ects: 0,
      history: safeJSON(s.History, []),
      resolved: new Set(String(s.Resolved || '').split(',').filter(Boolean))
    });
  });
  D.regs.forEach(g => {
    const s = stu[g.Code];
    if (!s) return;
    s.regs.push(g);
    if (g.Status !== 'REJECTED') s.ects += num(subj[g.SubjectCode]?.ECTS);
  });

  const byName = {};
  Object.values(stu).forEach(s => (byName[normName(s.Name)] = byName[normName(s.Name)] || []).push(s));

  Object.values(stu).forEach(s => {
    s.regs.sort((a, b) => num(subj[a.SubjectCode]?.Stage) - num(subj[b.SubjectCode]?.Stage));
    const all = String(s.Flags || '').split(',').filter(Boolean);
    s.twins = (byName[normName(s.Name)] || []).filter(x => x !== s);
    if (s.twins.length) all.push('SAME_NAME');
    if (!s.regs.length) all.push('NO_SUBJECTS');
    if (max && s.ects > max) all.push('OVER_ECTS');
    if (/نەخێر|دڵنیا نیم/.test(s.PrereqAnswer || '')) all.push('PREREQ');
    if ((s.Note || '').trim()) all.push('NOTE');
    s.allFlags = all;
    s.flags = all.filter(f => !s.resolved.has(f));
  });

  const subjTeachers = {};
  D.assign.forEach(a => {
    const u = users[a.UserID];
    if (!u || u.Active === 'No') return;
    (subjTeachers[a.SubjectCode] = subjTeachers[a.SubjectCode] || []).push(u.Name + (a.Group !== 'All' ? ` (${t(a.Group)})` : ''));
  });

  // Data issues, grouped by kind
  const issues = Object.fromEntries(ISSUE_ORDER.map(k => [k, []]));
  Object.values(stu).sort(byStageName).forEach(s => {
    s.allFlags.forEach(f => { if (issues[f]) issues[f].push({ kind: f, s, done: s.resolved.has(f) }); });
  });
  (D.issues || []).forEach(row => issues.UNIDENTIFIED.push({ kind: 'UNIDENTIFIED', row, done: false }));
  const regCount = {};
  D.regs.forEach(g => { if (g.Status !== 'REJECTED') regCount[g.SubjectCode] = (regCount[g.SubjectCode] || 0) + 1; });
  D.subjects.filter(x => x.Active !== 'No' && regCount[x.Code] && !subjTeachers[x.Code])
    .forEach(x => issues.NO_TEACHER.push({ kind: 'NO_TEACHER', subject: x, count: regCount[x.Code], done: false }));
  const openIssues = ISSUE_ORDER.reduce((n, k) => n + issues[k].filter(i => !i.done).length, 0);

  return { max, subj, users, stu, scores, list: Object.values(stu), subjTeachers, issues, openIssues };
}

function safeJSON(s, fallback) { try { return JSON.parse(s) || fallback; } catch (e) { return fallback; } }
const byStageName = (a, b) => num(a.Stage) - num(b.Stage) || String(a.Name).localeCompare(String(b.Name), 'ckb');

/** Plain-text explanation of one issue (used on screen and in the Word report). */
function issueDetail(item) {
  const s = item.s;
  switch (item.kind) {
    case 'MULTI_SUBMIT': {
      const h = s.history;
      const changes = [];
      const distinct = key => [...new Set(h.map(x => String(x[key])))];
      if (distinct('n').length > 1) changes.push(distinct('n').join(' / '));
      if (distinct('e').length > 1) changes.push(distinct('e').join(' / '));
      if (distinct('s').length > 1) changes.push(distinct('s').map(stageName).join(' / '));
      return t('submittedTimes', h.length || s.Submissions) + ': ' + h.map(x => fmtDate(x.t)).join(sep()) +
        (changes.length ? ` — ${t('changedFields')} ${changes.join(' — ')}` : '');
    }
    case 'SAME_NAME': return `${t('sameAs')} ${s.twins.map(x => `${x.Code} (${stageName(x.Stage)})`).join(sep())}`;
    case 'EMAIL_DOMAIN': return s.Email;
    case 'CODE_MISMATCH': {
      const note = s.CodeNote || '';
      if (note.startsWith('OTHER_CODE:')) {
        const other = note.split(':')[1];
        return t('code_OTHER', other, S.m.stu[other]?.Name);
      }
      return t('code_' + (note || 'TEXT'));
    }
    case 'NO_SUBJECTS': return '—';
    case 'OVER_ECTS': return `${s.ects} / ${S.m.max} ECTS`;
    case 'PREREQ': return s.PrereqAnswer;
    case 'NOTE': return s.Note;
    default: return '';
  }
}

/* ═══ 6. Boot, session, language ══════════════════════════════════════════ */
function applyLang() {
  document.documentElement.lang = S.lang === 'ku' ? 'ckb' : 'en';
  document.documentElement.dir = S.lang === 'ku' ? 'rtl' : 'ltr';
  document.title = `${CONFIG.SHORT_NAME} — ${t('app')}`;
}

function toggleLang() {
  S.lang = S.lang === 'ku' ? 'en' : 'ku';
  store('lang', S.lang);
  applyLang();
  if (S.data) S.m = buildModel(S.data);   // issue texts are language-dependent
  render();
}

async function loadData() {
  $('#app').innerHTML = shell('<div class="loading" aria-busy="true">…</div>');
  try {
    if (S.user.role === 'HEAD') {
      S.data = await api('headData');
      S.m = buildModel(S.data);
      S.semester = S.data.settings.Semester;
    } else {
      S.tdata = await api('teacherData');
      S.semester = S.tdata.settings.Semester;
    }
    if (!S.tab) S.tab = S.user.role === 'HEAD' ? 'overview' : 'sheets';
    render();
    flushOutbox();
  } catch (e) {
    if (S.token) $('#app').innerHTML = shell(`<div class="empty">${esc(errText(e))}</div>`);
  }
}

async function reloadHead() {
  S.data = await api('headData');
  S.m = buildModel(S.data);
  S.semester = S.data.settings.Semester;
}

function signOut(expired = false) {
  if (S.token && !expired) api('logout').catch(() => {});
  Object.assign(S, { token: null, user: null, data: null, m: null, tdata: null, tab: '', listFn: null });
  store('token', null);
  store('user', null);
  renderLogin(expired ? t('E_AUTH_REQUIRED') : '');
}

async function doLogin(form) {
  const btn = form.querySelector('[type=submit]');
  btn.disabled = true;
  btn.textContent = t('signingIn');
  try {
    const d = await api('login', { username: form.username.value, password: form.password.value });
    if (store('outbox_user') !== d.user.id) { outbox.length = 0; saveOutbox(); }   // unsent marks belong to whoever typed them
    S.token = d.token;
    S.user = d.user;
    store('token', d.token);
    store('user', JSON.stringify(d.user));
    await loadData();
  } catch (e) {
    $('#loginErr').textContent = errText(e);
    btn.disabled = false;
    btn.textContent = t('signIn');
  }
}

/* ═══ 7. Page frame ═══════════════════════════════════════════════════════ */
function banners() {
  const out = [];
  if (!navigator.onLine) out.push(`<div class="banner off" role="status">${t('offline')}</div>`);
  if (S.updateWorker) out.push(`<div class="banner"><span>${t('updateReady')}</span><button class="btn sm primary" data-act="applyUpdate">${t('reload')}</button></div>`);
  if (isIOS() && !isStandalone() && !store('iosHint')) out.push(`<div class="banner"><span>${t('iosInstall')}</span><button class="btn sm" data-act="dismissIos">${t('close')}</button></div>`);
  return out.join('');
}

const installButton = () => (S.installEvent ? `<button class="btn ghost sm" data-act="install">${t('install')}</button>` : '');

/** Refreshes only the banners and install button, so typing in a form is never lost. */
function updateChrome() {
  const b = $('#banners'), i = $('#installSlot');
  if (b) b.innerHTML = banners();
  if (i) i.innerHTML = installButton();
}

function renderLogin(msg = '') {
  S.listFn = null;
  $('#app').innerHTML = `
  <div id="banners">${banners()}</div>
  <div class="login">
    <div class="login-card">
      <div class="tools" style="justify-content:space-between;margin-bottom:10px">
        <img src="icons/icon.svg" alt="" width="44" height="44">
        <span class="btn-row">
          <span id="installSlot">${installButton()}</span>
          <button class="btn ghost sm" data-act="lang">${t('lang')}</button>
        </span>
      </div>
      <span class="spec"><b>${CONFIG.SHORT_NAME}</b><i id="sem">${esc(S.semester)}</i></span>
      <h1>${t('app')}</h1>
      <p class="sub" style="margin:0">${t('dept')}<br>${t('college')}${sep()}${t('uni')}</p>
      <form id="loginForm" autocomplete="on">
        <label class="fld">${t('loginId')}<input type="text" name="username" autocomplete="username" inputmode="email" dir="ltr" autocapitalize="none" autocorrect="off" spellcheck="false" value="${esc(new URLSearchParams(location.search).get('u') || '')}" required></label>
        <label class="fld">${t('codeLabel')}<span class="pwrow"><input type="password" name="password" autocomplete="current-password" dir="ltr" autocapitalize="none" autocorrect="off" spellcheck="false" required>
          <button type="button" class="btn ghost sm" data-act="togglePw">${t('showPw')}</button></span></label>
        <p class="err" id="loginErr" aria-live="polite">${esc(msg)}</p>
        <button class="btn primary" type="submit">${t('signIn')}</button>
      </form>
    </div>
  </div>`;
  if (!S.semester) {
    api('info').then(d => { S.semester = d.semester; const el = $('#sem'); if (el) el.textContent = d.semester; }).catch(() => {});
  }
}

function tabs() {
  return S.user.role === 'HEAD'
    ? ['overview', 'review', 'issues', 'sheets', 'students', 'teachers', 'settings']
    : ['sheets'];
}

function shell(inner) {
  const pendingStudents = S.m ? S.m.list.filter(s => s.regs.some(g => g.Status === 'PENDING')).length : 0;
  const badge = k => (k === 'review' && pendingStudents ? `<span class="count">${pendingStudents}</span>`
                    : k === 'issues' && S.m?.openIssues ? `<span class="count">${S.m.openIssues}</span>` : '');
  return `
  <header class="top">
    <div class="top-in">
      <img src="icons/icon.svg" alt="" width="34" height="34">
      <div class="brand"><strong class="long">${t('dept')}</strong><strong class="short">${CONFIG.SHORT_NAME}</strong><span>${t('app')}${S.semester ? ' — ' + esc(S.semester) : ''}</span></div>
      <div class="tools">
        <span id="installSlot" class="hide-sm">${installButton()}</span>
        <button class="btn ghost sm" data-act="lang">${t('lang')}</button>
        <button class="btn ghost sm uname" data-act="pw" title="${t('changePw')}">${esc(S.user.name)}</button>
        <button class="btn sm" data-act="signout">${t('signOut')}</button>
      </div>
    </div>
    ${tabs().length > 1 ? `<nav class="nav" aria-label="Sections">
      ${tabs().map(k => `<button data-act="tab" data-tab="${k}" aria-current="${S.tab === k}">${t(k)}${badge(k)}</button>`).join('')}
    </nav>` : ''}
  </header>
  <div id="banners">${banners()}</div>
  <main id="main">${inner}</main>`;
}

const VIEWS = {
  overview: () => vOverview(), review: () => vReview(), issues: () => vIssues(), sheets: () => vSheets(),
  students: () => vStudents(), teachers: () => vTeachers(), settings: () => vSettings()
};

function render() {
  if (!S.token) return renderLogin();
  if (!S.m && !S.tdata) return loadData();
  S.listFn = null;
  const body = (VIEWS[S.tab] || VIEWS.overview)();
  $('#app').innerHTML = shell(body);
  const top = $('.top');
  if (top) document.documentElement.style.setProperty('--top-h', top.offsetHeight + 'px');
  refreshList();
  paintSaveState();
}

function refreshList() {
  const el = $('#list');
  if (el && S.listFn) el.innerHTML = S.listFn();
}

/* ═══ 8. Shared UI pieces ═════════════════════════════════════════════════ */
function subjectOf(code) {
  return S.m?.subj[code] || (S.tdata?.subjects || []).find(x => x.Code === code) || null;
}
const spec = code => { const s = subjectOf(code); return `<span class="spec"><b>${esc(code)}</b>${s ? `<i>${esc(s.ECTS)} ECTS</i>` : ''}</span>`; };
const subjName = code => subjectOf(code)?.Name || '';
const cap = type => `<span class="cap cap-${type}">${t(type)}</span>`;
const statusChip = st => `<span class="st st-${st}">${t(st)}</span>`;
const flagChips = flags => (flags.length ? `<div class="flags">${flags.map(f => `<span class="flag">${t(f)}</span>`).join('')}</div>` : '');
const legend = () => `<div class="legend">${TYPES.map(cap).join('')}</div>`;
const stageOptions = sel => `<option value="">${t('allStages')}</option>` +
  [1, 2, 3, 4].map(n => `<option value="${n}" ${String(sel) === String(n) ? 'selected' : ''}>${stageName(n)}</option>`).join('');
const studentLink = s => `<a href="#" data-act="student" data-code="${esc(s.Code)}">${esc(s.Name)}</a>`;
const matchQ = s => { const q = S.f.q.trim().toLowerCase(); return !q || String(s.Name || '').toLowerCase().includes(q) || String(s.Code).includes(q); };
const exportButtons = scope => `
  <button class="btn" data-act="export" data-scope="${scope}" data-format="csv">${t('exportCsv')}</button>
  <button class="btn" data-act="export" data-scope="${scope}" data-format="docx">${t('exportDocx')}</button>`;

/* ═══ 9. Views — head ═════════════════════════════════════════════════════ */
function vOverview() {
  const m = S.m, regs = S.data.regs;
  const active = regs.filter(g => g.Status !== 'REJECTED');
  const counts = {};
  active.forEach(g => {
    const c = counts[g.SubjectCode] = counts[g.SubjectCode] || { REGULAR: 0, RETAKE: 0, EXAM_RETAKE: 0, PENDING: 0, GRADED: 0 };
    c[g.Type]++;
    if (g.Status === 'PENDING') c.PENDING++;
    const sc = m.scores[g.RegID];
    if (sc && sc.A !== '' && sc.B !== '') c.GRADED++;
  });
  const widest = Math.max(1, ...Object.values(counts).map(c => c.REGULAR + c.RETAKE + c.EXAM_RETAKE));
  const byStage = {};
  S.data.subjects.filter(s => s.Active !== 'No').forEach(s => (byStage[s.Stage] = byStage[s.Stage] || []).push(s));
  const perStage = {};
  m.list.forEach(s => { perStage[s.Stage] = (perStage[s.Stage] || 0) + 1; });

  let html = `<h1>${t('overview')}</h1><p class="sub">${t('lastSync')}: ${fmtDate(S.data.settings.LastSync)}</p>
  <dl class="facts">
    <div><dt>${t('nStudents')}</dt><dd>${m.list.length}</dd></div>
    <div><dt>${t('nParallel')}</dt><dd>${m.list.filter(s => s.Parallel === 'Yes').length}</dd></div>
    <div><dt>${t('nRegs')}</dt><dd>${active.length}</dd></div>
    <div><dt>${t('nPending')}</dt><dd>${regs.filter(g => g.Status === 'PENDING').length}</dd></div>
    <div><dt>${t('nIssues')}</dt><dd>${m.openIssues}</dd></div>
  </dl>
  ${m.openIssues ? `<div class="cta-issues"><p>${t('issuesIntro')}</p><button class="btn primary" data-act="tab" data-tab="issues">${t('seeIssues', m.openIssues)}</button></div>` : ''}
  ${legend()}`;

  Object.keys(byStage).sort().forEach(stage => {
    html += `<h2>${stageName(stage)} <span class="muted" style="font-weight:400;font-size:14px">— ${perStage[stage] || 0} ${t('nStudents')}</span></h2>
    <div class="tbl-wrap"><table><thead><tr><th>${t('code')}</th><th>${t('subject')}</th><th>${t('perSubject')}</th>
      <th class="num">${t('total')}</th><th class="num">${t('PENDING')}</th><th class="num">${t('graded')}</th><th class="hide-sm">${t('teacher')}</th></tr></thead><tbody>`;
    byStage[stage].forEach(s => {
      const c = counts[s.Code] || { REGULAR: 0, RETAKE: 0, EXAM_RETAKE: 0, PENDING: 0, GRADED: 0 };
      const total = c.REGULAR + c.RETAKE + c.EXAM_RETAKE, again = c.RETAKE + c.EXAM_RETAKE;
      const w = n => (n / widest * 100).toFixed(1) + '%';
      html += `<tr class="click" data-act="openSheet" data-code="${esc(s.Code)}">
        <td>${spec(s.Code)}</td><td>${esc(s.Name)}</td>
        <td><div class="sbar" title="${t('REGULAR')} ${c.REGULAR} — ${t('RETAKE')} ${c.RETAKE} — ${t('EXAM_RETAKE')} ${c.EXAM_RETAKE}">
          <span style="width:${w(c.REGULAR)};background:var(--cap-reg)"></span><span style="width:${w(c.RETAKE)};background:var(--cap-ret)"></span><span style="width:${w(c.EXAM_RETAKE)};background:var(--cap-exam)"></span></div></td>
        <td class="num"><b>${total}</b>${again ? ` <span class="muted">(${again} ${t('retakes')})</span>` : ''}</td>
        <td class="num">${c.PENDING || '<span class="muted">0</span>'}</td>
        <td class="num">${total ? `${c.GRADED}<span class="muted"> / ${total}</span>` : '<span class="muted">—</span>'}</td>
        <td class="hide-sm">${esc((m.subjTeachers[s.Code] || []).join(sep())) || `<span class="muted">${t('teachersNone')}</span>`}</td></tr>`;
    });
    html += '</tbody></table></div>';
  });
  return html;
}

function vReview() {
  S.listFn = reviewList;
  return `<h1>${t('review')}</h1>
  <div class="bar">
    <input type="search" class="grow" data-f="q" value="${esc(S.f.q)}" placeholder="${t('search')}" aria-label="${t('search')}">
    <select data-f="stage" aria-label="${t('stage')}">${stageOptions(S.f.stage)}</select>
    <label class="check"><input type="checkbox" data-f="pending" ${S.f.pending ? 'checked' : ''}> ${t('pendingOnly')}</label>
    <label class="check"><input type="checkbox" data-f="flagged" ${S.f.flagged ? 'checked' : ''}> ${t('flaggedOnly')}</label>
  </div>
  ${legend()}
  <div id="list"></div>`;
}

function reviewList() {
  const max = S.m.max;
  const list = S.m.list.filter(s => matchQ(s) && (!S.f.stage || s.Stage === S.f.stage) &&
      (!S.f.pending || s.regs.some(g => g.Status === 'PENDING')) && (!S.f.flagged || s.flags.length))
    .sort((a, b) => (b.flags.length > 0) - (a.flags.length > 0) || byStageName(a, b));
  if (!list.length) return `<div class="empty">${t('noCases')}</div>`;

  return list.map(s => {
    const pending = s.regs.filter(g => g.Status === 'PENDING').map(g => g.RegID).join('|');
    const rows = s.regs.map(g => {
      const sub = S.m.subj[g.SubjectCode] || {};
      const fromOtherStage = sub.Stage && sub.Stage !== s.Stage ? ` <span class="tag">${stageName(sub.Stage)}</span>` : '';
      const btn = (st, cls, label) => `<button class="btn ${cls} sm" data-act="status" data-ids="${esc(g.RegID)}" data-status="${st}"${st === 'PENDING' ? ` title="${t('undo')}" aria-label="${t('undo')}"` : ''}>${label}</button>`;
      return `<tr>
        <td>${spec(g.SubjectCode)}</td>
        <td>${esc(sub.Name || '')}${fromOtherStage}${g.HeadNote ? `<div class="muted" style="font-size:13px">${esc(g.HeadNote)}</div>` : ''}</td>
        <td>${cap(g.Type)}</td><td>${statusChip(g.Status)}</td>
        <td><div class="acts">
          ${g.Status !== 'APPROVED' ? btn('APPROVED', 'ok', t('approve')) : ''}
          ${g.Status !== 'REJECTED' ? btn('REJECTED', 'no', t('reject')) : ''}
          ${g.Status !== 'PENDING' ? btn('PENDING', 'ghost', '↺') : ''}
        </div></td></tr>`;
    }).join('');
    const extra = [
      s.PrereqAnswer && `<div><b>${t('prereq')}:</b> ${esc(s.PrereqAnswer)}</div>`,
      s.RetakeDetails && `<div><b>${t('retakeDetails')}:</b> ${esc(s.RetakeDetails)}</div>`,
      s.Note && `<div><b>${t('note')}:</b> ${esc(s.Note)}</div>`
    ].filter(Boolean).join('');

    return `<article class="case ${s.flags.length ? 'flagged' : ''}">
      <header>
        <div><h3>${studentLink(s)}</h3>
          <p class="meta"><span class="ltr">${esc(s.Code)}</span><span>${stageName(s.Stage)}</span><span class="tag">${groupName(s.Parallel)}</span></p>
          ${flagChips(s.flags)}</div>
        <div class="ects ${max && s.ects > max ? 'over' : ''}">${s.ects} / ${max || '—'} ECTS</div>
      </header>
      ${rows ? `<div class="tbl-wrap flat"><table><tbody>${rows}</tbody></table></div>` : ''}
      ${extra ? `<div class="extra">${extra}</div>` : ''}
      ${pending ? `<footer>
        <button class="btn no" data-act="status" data-ids="${esc(pending)}" data-status="REJECTED">${t('rejectAll')}</button>
        <button class="btn primary" data-act="status" data-ids="${esc(pending)}" data-status="APPROVED">${t('approveAll')}</button></footer>` : ''}
    </article>`;
  }).join('');
}

function vIssues() {
  S.listFn = issuesList;
  return `<h1>${t('issues')}</h1>
  <p class="sub">${t('issuesIntro')}</p>
  <div class="bar">
    <label class="check"><input type="checkbox" data-f="showChecked" ${S.f.showChecked ? 'checked' : ''}> ${t('showChecked')}</label>
    <span class="grow"></span>
    ${exportButtons('issues')}
  </div>
  <div id="list"></div>`;
}

function visibleIssues(kind) {
  return S.m.issues[kind].filter(i => S.f.showChecked || !i.done);
}

function issuesList() {
  const nav = ISSUE_ORDER.map(k => {
    const n = S.m.issues[k].filter(i => !i.done).length;
    return `<a href="#iss-${k}" class="${n ? '' : 'zero'}">${t(k)} <b>${n}</b></a>`;
  }).join('');
  if (!S.m.openIssues && !S.f.showChecked) return `<nav class="issue-nav">${nav}</nav><div class="empty">${t('allClear')}</div>`;

  const sections = ISSUE_ORDER.map(kind => {
    const items = visibleIssues(kind);
    if (!items.length) return '';
    const open = S.m.issues[kind].filter(i => !i.done).length;
    const desc = kind === 'OVER_ECTS' ? t('OVER_ECTS_d', S.m.max) : t(kind + '_d');
    let head, rows;
    if (kind === 'UNIDENTIFIED') {
      head = `<th>${t('time')}</th><th>${t('name')}</th><th>${t('email')}</th>`;
      rows = items.map(i => `<tr><td>${fmtDate(i.row.Time)}</td><td>${esc(i.row.Name)}</td><td><span class="ltr">${esc(i.row.Email)}</span></td></tr>`);
    } else if (kind === 'NO_TEACHER') {
      head = `<th>${t('code')}</th><th>${t('subject')}</th><th>${t('stage')}</th><th class="num">${t('nStudents')}</th><th></th>`;
      rows = items.map(i => `<tr><td>${spec(i.subject.Code)}</td><td>${esc(i.subject.Name)}</td><td>${stageName(i.subject.Stage)}</td>
        <td class="num">${i.count}</td><td><button class="btn sm" data-act="tab" data-tab="teachers">${t('assignTeacher')}</button></td></tr>`);
    } else {
      head = `<th>${t('name')}</th><th>${t('code')}</th><th>${t('stage')}</th><th>${t('detail')}</th><th></th>`;
      rows = items.map(i => `<tr class="${i.done ? 'done' : ''}">
        <td>${studentLink(i.s)}</td><td><span class="ltr">${esc(i.s.Code)}</span></td><td>${stageName(i.s.Stage)}</td>
        <td class="diff">${esc(issueDetail(i))}</td>
        <td><button class="btn sm ${i.done ? 'ghost' : ''}" data-act="resolve" data-code="${esc(i.s.Code)}" data-flag="${kind}" data-done="${i.done}">
          ${i.done ? t('unmark') : t('markChecked')}</button></td></tr>`);
    }
    return `<section class="issue-sec" id="iss-${kind}">
      <h2>${t(kind)} <span class="n">${open}</span></h2><p>${esc(desc)}</p>
      <div class="tbl-wrap"><table><thead><tr>${head}</tr></thead><tbody>${rows.join('')}</tbody></table></div>
    </section>`;
  }).join('');
  return `<nav class="issue-nav">${nav}</nav>${sections}`;
}

function vStudents() {
  S.listFn = studentsList;
  return `<h1>${t('students')}</h1>
  <div class="bar">
    <input type="search" class="grow" data-f="q" value="${esc(S.f.q)}" placeholder="${t('search')}" aria-label="${t('search')}">
    <select data-f="stage" aria-label="${t('stage')}">${stageOptions(S.f.stage)}</select>
    <label class="check"><input type="checkbox" data-f="flagged" ${S.f.flagged ? 'checked' : ''}> ${t('flaggedOnly')}</label>
  </div>
  <div id="list"></div>`;
}

function studentsList() {
  const max = S.m.max;
  const list = S.m.list.filter(s => matchQ(s) && (!S.f.stage || s.Stage === S.f.stage) && (!S.f.flagged || s.flags.length)).sort(byStageName);
  if (!list.length) return `<div class="empty">${t('noRows')}</div>`;
  return `<div class="tbl-wrap"><table><thead><tr><th>${t('name')}</th><th>${t('code')}</th><th>${t('stage')}</th><th>${t('group')}</th>
    <th class="num">${t('ects')}</th><th class="hide-sm">${t('status')}</th></tr></thead><tbody>
  ${list.map(s => {
    const c = { PENDING: 0, APPROVED: 0, REJECTED: 0 };
    s.regs.forEach(g => { c[g.Status]++; });
    return `<tr class="click" data-act="student" data-code="${esc(s.Code)}">
      <td>${esc(s.Name)} ${s.flags.length ? `<span class="flag" title="${esc(s.flags.map(f => t(f)).join(sep()))}">${s.flags.length}</span>` : ''}</td>
      <td><span class="ltr">${esc(s.Code)}</span></td><td>${stageName(s.Stage)}</td><td>${groupName(s.Parallel)}</td>
      <td class="num ${max && s.ects > max ? 'ects over' : ''}">${s.ects}</td>
      <td class="hide-sm">${STATUSES.filter(k => c[k]).map(k => `${statusChip(k)} ${c[k]}`).join(' ')}</td></tr>`;
  }).join('')}</tbody></table></div>`;
}

function vTeachers() {
  const byUser = {};
  S.data.assign.forEach(a => (byUser[a.UserID] = byUser[a.UserID] || []).push(a));
  const users = S.data.users.slice().sort((a, b) => (a.Role === 'HEAD' ? -1 : 0) - (b.Role === 'HEAD' ? -1 : 0) || a.Name.localeCompare(b.Name, 'ckb'));
  const others = users.filter(u => u.Email && u.Active !== 'No' && u.UserID !== S.user.id);
  const fresh = others.filter(u => !u.LastInvite);
  const status = u => (u.LastInvite ? `<span class="st st-APPROVED">${t('invited', fmtDate(u.LastInvite, false))}</span>`
                     : u.HasCode === 'Yes' ? `<span class="st st-PENDING">${t('hasCode')}</span>` : `<span class="st st-REJECTED">${t('notInvited')}</span>`);
  return `<h1>${t('teachers')}</h1>
  <p class="sub">${t('inviteHint')}</p>
  <div class="bar btn-row">
    <button class="btn primary" data-act="editUser">${t('addTeacher')}</button>
    <button class="btn" data-act="inviteMany" data-which="new" ${fresh.length ? '' : 'disabled'}>${t('inviteNew', fresh.length)}</button>
    <button class="btn" data-act="inviteMany" data-which="all" ${others.length ? '' : 'disabled'}>${t('inviteAll', others.length)}</button>
  </div>
  <div class="tbl-wrap"><table><thead><tr><th>${t('name')}</th><th>${t('email')}</th><th>${t('role')}</th><th>${t('assigned')}</th><th>${t('invitation')}</th><th>${t('lastLogin')}</th><th></th></tr></thead><tbody>
  ${users.map(u => `<tr class="${u.Active === 'No' ? 'muted' : ''}" data-user="${esc(u.UserID)}">
    <td>${esc(u.Name)} ${u.UserID === S.user.id ? `<span class="tag">${t('you')}</span>` : ''}${u.Active === 'No' ? ` <span class="tag">${t('inactive')}</span>` : ''}</td>
    <td><span class="ltr">${esc(u.Email || u.Username)}</span></td><td>${t(u.Role)}</td>
    <td>${(byUser[u.UserID] || []).map(a => spec(a.SubjectCode) + (a.Group !== 'All' ? ` <span class="muted" style="font-size:12px">${t(a.Group)}</span>` : '')).join(' ') || '<span class="muted">—</span>'}</td>
    <td>${status(u)}</td>
    <td>${u.LastLogin ? fmtDate(u.LastLogin) : `<span class="muted">${t('never')}</span>`}</td>
    <td><div class="acts"><button class="btn sm" data-act="editUser" data-id="${esc(u.UserID)}">${t('edit')}</button>
      ${u.Email && u.UserID !== S.user.id && u.Active !== 'No' ? `<button class="btn sm primary" data-act="inviteOne" data-id="${esc(u.UserID)}">${u.LastInvite ? t('resend') : t('sendInvite')}</button>` : ''}</div></td></tr>`).join('')}
  </tbody></table></div>`;
}

function vSettings() {
  const s = S.data.settings;
  return `<h1>${t('settings')}</h1>
  <form id="setForm" class="tbl-wrap" style="padding:16px;display:grid;gap:12px;max-width:640px">
    <label class="fld">${t('semester')}<input type="text" name="Semester" value="${esc(s.Semester)}"></label>
    <div class="two">
      <label class="fld">${t('maxEcts')}<input type="number" name="MaxECTS" min="0" max="90" value="${esc(s.MaxECTS)}"></label>
      <label class="fld">${t('formSheet')}<input type="text" name="FormSheet" value="${esc(s.FormSheet)}" dir="ltr"></label>
      <label class="fld">${t('maxA')}<input type="number" name="MaxA" min="1" max="100" step="any" value="${esc(s.MaxA)}"></label>
      <label class="fld">${t('maxB')}<input type="number" name="MaxB" min="1" max="100" step="any" value="${esc(s.MaxB)}"></label>
      <label class="fld">${t('passMark')}<input type="number" name="PassMark" min="0" max="200" step="any" value="${esc(s.PassMark)}"></label>
      <label class="fld">${t('appUrl')}<input type="url" name="AppUrl" value="${esc(s.AppUrl)}" dir="ltr" placeholder="https://…"></label>
    </div>
    <div><button class="btn primary" type="submit">${t('save')}</button></div>
  </form>
  <h2>${t('sync')}</h2>
  <p class="sub">${t('syncHint')}<br>${t('lastSync')}: ${fmtDate(s.LastSync)}</p>
  <button class="btn" data-act="sync">${t('sync')}</button>
  <h2>${t('subjectsList')}</h2>
  <div class="tbl-wrap"><table><thead><tr><th>${t('code')}</th><th>${t('name')}</th><th>${t('stage')}</th><th>${t('ects')}</th><th>${t('prereqCode')}</th><th>${t('active')}</th><th></th></tr></thead><tbody>
  ${S.data.subjects.slice().sort((a, b) => num(a.Stage) - num(b.Stage)).map(x => `<tr>
    <td>${spec(x.Code)}</td>
    <td><input type="text" data-k="Name" value="${esc(x.Name)}" dir="ltr" style="width:100%;min-width:160px" aria-label="${t('name')}"></td>
    <td><select data-k="Stage" aria-label="${t('stage')}">${[0, 1, 2, 3, 4].map(n => `<option value="${n}" ${String(n) === x.Stage ? 'selected' : ''}>${n || '—'}</option>`).join('')}</select></td>
    <td><input type="number" data-k="ECTS" value="${esc(x.ECTS)}" min="0" max="30" style="width:70px" aria-label="ECTS"></td>
    <td><input type="text" data-k="Prerequisite" value="${esc(x.Prerequisite)}" dir="ltr" style="width:110px" aria-label="${t('prereqCode')}"></td>
    <td><input type="checkbox" data-k="Active" ${x.Active !== 'No' ? 'checked' : ''} aria-label="${t('active')}"></td>
    <td><button class="btn sm" data-act="saveSubject" data-code="${esc(x.Code)}">${t('save')}</button></td></tr>`).join('')}
  </tbody></table></div>`;
}

/* ═══ 10. Mark sheets: subject cards, the numbered list with 20 + 30 = 50, live saving ════ */
function limits() {
  const st = S.user.role === 'HEAD' ? S.data.settings : S.tdata.settings;
  return { maxA: numOr(st.MaxA, 20), maxB: numOr(st.MaxB, 30), pass: numOr(st.PassMark, 25) };
}

function sheetSubjects() {
  const list = S.user.role === 'HEAD' ? S.data.subjects.filter(s => s.Active !== 'No') : S.tdata.subjects;
  return list.slice().sort((a, b) => num(a.Stage) - num(b.Stage) || a.Code.localeCompare(b.Code));
}

function teacherNamesFor(code) {
  return S.user.role === 'HEAD' ? (S.m.subjTeachers[code] || []) : [S.user.name];
}

/** Rows of one subject with numeric scores, total and result. Rejected registrations never appear. */
function sheetRowsFor(code) {
  const lim = limits();
  const base = S.user.role === 'HEAD'
    ? S.data.regs.filter(g => g.SubjectCode === code && g.Status !== 'REJECTED').map(g => {
        const s = S.m.stu[g.Code] || {}, sc = S.m.scores[g.RegID] || {};
        return { RegID: g.RegID, Code: g.Code, Name: s.Name || g.Code, Stage: s.Stage, Parallel: s.Parallel, Type: g.Type, Status: g.Status, A: sc.A, B: sc.B };
      })
    : S.tdata.rows.filter(r => r.SubjectCode === code);
  return base.map(r => {
    const a = numOr(r.A, null), b = numOr(r.B, null);
    const total = a != null && b != null ? round2(a + b) : null;
    return Object.assign({}, r, { a, b, total, result: total == null ? '' : total >= lim.pass ? 'PASS' : 'FAIL' });
  });
}

function sortRows(rows, how) {
  const byName = (a, b) => String(a.Name).localeCompare(String(b.Name), 'ckb');
  const tot = r => (r.total == null ? -1 : r.total);
  const rank = (r, first) => (r.result === first ? 0 : r.result ? 1 : 2);
  const cmp = {
    high: (a, b) => tot(b) - tot(a) || byName(a, b),
    low: (a, b) => (a.total == null) - (b.total == null) || tot(a) - tot(b) || byName(a, b),
    pass: (a, b) => rank(a, 'PASS') - rank(b, 'PASS') || tot(b) - tot(a) || byName(a, b),
    fail: (a, b) => rank(a, 'FAIL') - rank(b, 'FAIL') || (a.result === 'FAIL' ? tot(a) - tot(b) : tot(b) - tot(a)) || byName(a, b),
    code: (a, b) => String(a.Code).localeCompare(String(b.Code))
  }[how] || byName;
  return rows.sort(cmp);
}

function visibleSheetRows(code, withSearch = true) {
  return sortRows(sheetRowsFor(code).filter(r =>
    (!withSearch || matchQ(r)) &&
    (!S.f.group || r.Parallel === S.f.group) &&
    (!S.f.result || (S.f.result === 'NONE' ? r.total == null : r.result === S.f.result))), S.f.sort);
}

function stats(rows) {
  const totals = rows.map(r => r.total).filter(v => v != null);
  const passed = rows.filter(r => r.result === 'PASS').length;
  return {
    count: rows.length, graded: totals.length, passed, failed: totals.length - passed,
    avg: totals.length ? round2(totals.reduce((a, b) => a + b, 0) / totals.length) : '',
    max: totals.length ? Math.max(...totals) : '', min: totals.length ? Math.min(...totals) : '',
    rate: totals.length ? passed / totals.length : ''
  };
}

function statsHtml(rows) {
  const st = stats(rows);
  return [
    [t('nStudents'), st.count, ''], [t('graded'), `<span dir="ltr">${st.graded}<small> / ${st.count}</small></span>`, ''],
    [t('average'), st.avg === '' ? '—' : st.avg, ''], [t('highest'), st.max === '' ? '—' : st.max, ''],
    [t('PASS'), st.passed, 'pass'], [t('FAIL'), st.failed, 'fail']
  ].map(([k, v, c]) => `<div class="${c}"><dt>${k}</dt><dd>${v}</dd></div>`).join('');
}

function helloCard(compact) {
  if (S.user.role === 'HEAD') return '';
  const clean = String(S.user.name).replace(/^[^\s]*\.\s*/, '').trim() || S.user.name;
  const subjects = S.tdata.subjects;
  return `<section class="hello${compact ? ' compact' : ''}">
    <div class="avatar" aria-hidden="true">${esc(clean.charAt(0))}</div>
    <div class="hello-text"><p>${t('welcome')}</p><h1>${esc(S.user.name)}</h1>
      ${subjects.length && !compact ? `<p class="muted">${t('subjectsYouTeach')}</p><div class="chips">${subjects.map(s => `<span class="chip" dir="ltr"><b>${esc(s.Code)}</b> ${esc(s.Name)}</span>`).join('')}</div>` : ''}
    </div></section>`;
}

function vSheets() {
  const head = S.user.role === 'HEAD';
  const subjects = sheetSubjects();
  if (!head && subjects.length === 1) S.f.subject = subjects[0].Code;
  if (S.f.subject && subjects.some(s => s.Code === S.f.subject)) return vSheetDetail(S.f.subject);
  S.f.subject = '';
  if (!subjects.length) return `${helloCard()}<div class="empty">${t('noSubjects')}</div>`;

  const card = s => {
    const rows = sheetRowsFor(s.Code), st = stats(rows);
    const pct = rows.length ? Math.round(st.graded / rows.length * 100) : 0;
    return `<button class="scard" data-act="openSheet" data-code="${esc(s.Code)}">
      <span class="scard-top">${spec(s.Code)}<span class="muted">${stageName(s.Stage)}</span></span>
      <h3 dir="ltr">${esc(s.Name)}</h3>
      ${head ? `<p>${esc(teacherNamesFor(s.Code).join(sep())) || `<span class="muted">${t('teachersNone')}</span>`}</p>` : ''}
      <div class="prog" role="img" aria-label="${pct}%"><span style="width:${pct}%"></span></div>
      <p><span dir="ltr"><b style="color:var(--ink)">${st.graded}</b> / ${rows.length}</span> ${t('graded')}${st.graded ? ` — ${t('average')} <b style="color:var(--ink)">${st.avg}</b>` : ''}</p>
    </button>`;
  };
  if (!head) return `${helloCard()}<div class="grid">${subjects.map(card).join('')}</div>`;

  const byStage = {};
  subjects.forEach(s => (byStage[s.Stage] = byStage[s.Stage] || []).push(s));
  return `<h1>${t('sheets')}</h1>
    <div class="bar btn-row"><span class="muted">${t('downloadAll')}</span>
      <button class="btn" data-act="export" data-scope="all" data-format="docx">${t('word')}</button>
      <button class="btn" data-act="export" data-scope="all" data-format="xlsx">${t('excel')}</button></div>
    ${Object.keys(byStage).sort().map(st => `<h2>${stageName(st)}</h2><div class="grid">${byStage[st].map(card).join('')}</div>`).join('')}`;
}

function vSheetDetail(code) {
  S.listFn = sheetList;
  const sub = subjectOf(code), head = S.user.role === 'HEAD';
  const rows = sheetRowsFor(code);
  const lim = limits(), editable = !head || S.f.unlock;
  const back = head || S.tdata.subjects.length > 1;
  const bothGroups = rows.some(r => r.Parallel === 'Yes') && rows.some(r => r.Parallel !== 'Yes');
  const opt = (v, label, cur) => `<option value="${v}" ${cur === v ? 'selected' : ''}>${label}</option>`;
  return `${helloCard(true)}
  <div class="sheet-head">
    ${back ? `<button class="btn ghost sm" data-act="closeSheet">${t('allSubjects')}</button>` : ''}
    <div class="sheet-title">${spec(code)}<h1 dir="ltr">${esc(sub.Name)}</h1></div>
    <p class="sub">${stageName(sub.Stage)} — ${esc(sub.ECTS)} ECTS — ${esc(S.semester)}${head ? ` — ${esc(teacherNamesFor(code).join(sep()) || t('teachersNone'))}` : ''}</p>
  </div>
  <dl class="facts" id="sheetStats">${statsHtml(rows)}</dl>
  <div class="bar sheet-bar">
    <input type="search" class="grow" data-f="q" value="${esc(S.f.q)}" placeholder="${t('search')}" aria-label="${t('search')}">
    <select data-f="sort" aria-label="${t('sortBy')}">${['name', 'high', 'low', 'pass', 'fail', 'code'].map(k => opt(k, t('sort_' + k), S.f.sort)).join('')}</select>
    <select data-f="result" aria-label="${t('result')}">${opt('', t('allResults'), S.f.result)}${opt('PASS', t('PASS'), S.f.result)}${opt('FAIL', t('FAIL'), S.f.result)}${opt('NONE', t('notGraded'), S.f.result)}</select>
    ${bothGroups ? `<select data-f="group" aria-label="${t('group')}">${opt('', t('All'), S.f.group)}${opt('No', t('Morning'), S.f.group)}${opt('Yes', t('Parallel'), S.f.group)}</select>` : ''}
  </div>
  <div class="bar sheet-actions">
    <span id="saveState" class="save-state" aria-live="polite"></span>
    <span class="grow"></span>
    ${head ? `<label class="check"><input type="checkbox" data-f="unlock" ${S.f.unlock ? 'checked' : ''}> ${t('editMarks')}</label>` : ''}
    <button class="btn" data-act="export" data-scope="sheet" data-format="docx">${t('word')}</button>
    <button class="btn" data-act="export" data-scope="sheet" data-format="xlsx">${t('excel')}</button>
  </div>
  ${editable ? `<p class="hint">${t('marksHint', lim)}</p>` : ''}
  <div id="list"></div>`;
}

const fmtScore = v => (v == null ? '' : String(v));
const resultBadge = r => (r.result ? `<span class="res ${r.result.toLowerCase()}">${t(r.result)}</span>` : `<span class="res none">${t('notGraded')}</span>`);

function markRow(r, i, lim, editable) {
  const tags = [
    r.Type !== 'REGULAR' ? `<span class="cap cap-${r.Type}">${t(r.Type + '_s')}</span>` : '',
    r.Parallel === 'Yes' ? `<span class="tag">${t('Parallel')}</span>` : '',
    r.Status === 'PENDING' ? `<span class="st st-PENDING">${t('PENDING')}</span>` : ''
  ].join(' ');
  const cell = (field, max, value) => {
    const label = field === 'A' ? t('scoreA') : t('scoreB');
    return `<label class="mcell"><span>${label} <i>/${max}</i></span><input class="sc" type="text" inputmode="decimal" autocomplete="off" enterkeyhint="next"
      data-reg="${esc(r.RegID)}" data-field="${field}" data-max="${max}" value="${fmtScore(value)}" aria-label="${esc(r.Name)} — ${label} /${max}"
      ${editable ? '' : 'readonly tabindex="-1"'}></label>`;
  };
  return `<li class="mrow ${r.result.toLowerCase()}" data-row="${esc(r.RegID)}">
    <span class="mno">${i + 1}</span>
    <div class="mwho"><b>${esc(r.Name)}</b><small><span class="ltr">${esc(r.Code)}</span> ${tags}</small></div>
    ${cell('A', lim.maxA, r.a)}${cell('B', lim.maxB, r.b)}
    <div class="mtot"><span>${t('total')} <i>/${lim.maxA + lim.maxB}</i></span><b>${r.total == null ? '—' : r.total}</b></div>
    <div class="mres">${resultBadge(r)}</div>
  </li>`;
}

function sheetList() {
  const code = S.f.subject, lim = limits();
  const editable = S.user.role !== 'HEAD' || S.f.unlock;
  if (!sheetRowsFor(code).length) return `<div class="empty">${t('noStudents')}</div>`;
  const rows = visibleSheetRows(code);
  if (!rows.length) return `<div class="empty">${t('noRows')}</div>`;
  return `<div class="mhead" aria-hidden="true"><span>#</span><span>${t('name')}</span><span>${t('scoreA')} <i>/${lim.maxA}</i></span>
      <span>${t('scoreB')} <i>/${lim.maxB}</i></span><span>${t('total')} <i>/${lim.maxA + lim.maxB}</i></span><span>${t('result')}</span></div>
    <ol class="marks ${editable ? '' : 'locked'}">${rows.map((r, i) => markRow(r, i, lim, editable)).join('')}</ol>`;
}

/* ── editing and saving ── */
const DIGITS = { '٠': 0, '١': 1, '٢': 2, '٣': 3, '٤': 4, '٥': 5, '٦': 6, '٧': 7, '٨': 8, '٩': 9, '۰': 0, '۱': 1, '۲': 2, '۳': 3, '۴': 4, '۵': 5, '۶': 6, '۷': 7, '۸': 8, '۹': 9 };
const latinDigits = s => String(s).replace(/[٠-٩۰-۹]/g, d => DIGITS[d]).replace(/[,٫]/g, '.');

function getScore(id, field) {
  if (S.user.role === 'HEAD') return (S.m.scores[id] || {})[field] || '';
  return (S.tdata.rows.find(r => r.RegID === id) || {})[field] || '';
}

function setLocalScore(id, field, value) {
  if (S.user.role !== 'HEAD') {
    const row = S.tdata.rows.find(r => r.RegID === id);
    if (row) row[field] = value;
    return;
  }
  let rec = S.m.scores[id];
  if (!rec) {
    const g = S.data.regs.find(x => x.RegID === id);
    rec = { RegID: id, Code: g.Code, SubjectCode: g.SubjectCode, A: '', B: '' };
    S.data.scores = S.data.scores || [];
    S.data.scores.push(rec);
    S.m.scores[id] = rec;
  }
  rec[field] = value;
}

function onScoreChange(el) {
  const raw = latinDigits(el.value.trim()), max = Number(el.dataset.max);
  let value = '';
  if (raw !== '') {
    const n = Number(raw);
    if (!isFinite(n) || n < 0 || n > max) {
      el.classList.add('bad');
      el.setAttribute('aria-invalid', 'true');
      toast(t('rangeMsg', max), true);
      return;
    }
    value = String(round2(n));
  }
  el.classList.remove('bad');
  el.removeAttribute('aria-invalid');
  el.value = value;
  const id = el.dataset.reg, field = el.dataset.field;
  if (getScore(id, field) === value) return;
  setLocalScore(id, field, value);
  queueScore(id, field, value);
  updateRowUI(id);
}

function updateRowUI(id) {
  const rows = sheetRowsFor(S.f.subject), r = rows.find(x => x.RegID === id);
  const li = document.querySelector(`li[data-row="${CSS.escape(id)}"]`);
  if (r && li) {
    li.className = 'mrow ' + r.result.toLowerCase();
    li.querySelector('.mtot b').textContent = r.total == null ? '—' : r.total;
    li.querySelector('.mres').innerHTML = resultBadge(r);
  }
  const st = $('#sheetStats');
  if (st) st.innerHTML = statsHtml(rows);
}

/* Marks wait in a small outbox (kept on the device) until the server confirms them,
   so nothing is lost on a weak connection or if the app is closed. */
function loadOutbox() { try { return JSON.parse(store('outbox') || '[]'); } catch (e) { return []; } }
const outbox = loadOutbox();
function saveOutbox() {
  store('outbox', outbox.length ? JSON.stringify(outbox) : null);
  if (S.user) store('outbox_user', S.user.id);
}
let flushing = false, flushTimer = null, retryTimer = null;

function queueScore(id, field, value) {
  const i = outbox.findIndex(x => x.RegID === id && x.Field === field);
  if (i > -1) outbox.splice(i, 1);
  outbox.push({ RegID: id, Field: field, Value: value });
  saveOutbox();
  setSave('saving');
  clearTimeout(flushTimer);
  flushTimer = setTimeout(flushOutbox, 400);
}

async function flushOutbox() {
  if (flushing || !S.token || !S.user) return;
  if (!outbox.length) { if (S.saveState === 'saving') setSave('saved'); return; }
  flushing = true;
  setSave('saving');
  const batch = outbox.slice(0, 300);
  let again = false;
  try {
    const r = await api('setScores', { items: batch });
    batch.forEach(b => {
      const i = outbox.findIndex(x => x.RegID === b.RegID && x.Field === b.Field && x.Value === b.Value);
      if (i > -1) outbox.splice(i, 1);
    });
    saveOutbox();
    if (r.errors.length) {
      toast(t('someNotSaved', r.errors.length), true);
      setSave('error');
      reloadMarks();
    } else {
      setSave(outbox.length ? 'saving' : 'saved');
    }
    again = outbox.length > 0;
  } catch (e) {
    if (e.message === 'NETWORK') {
      setSave('offline');
      clearTimeout(retryTimer);
      retryTimer = setTimeout(flushOutbox, 15000);
    } else if (e.message !== 'AUTH_REQUIRED') {
      setSave('error');
      toast(errText(e), true);
    }
  } finally {
    flushing = false;
    if (again) setTimeout(flushOutbox, 50);
  }
}

async function reloadMarks() {
  try {
    if (S.user.role === 'HEAD') await reloadHead(); else S.tdata = await api('teacherData');
    refreshList();
    const st = $('#sheetStats');
    if (st && S.f.subject) st.innerHTML = statsHtml(sheetRowsFor(S.f.subject));
  } catch (e) { /* the next save retries */ }
}

function setSave(state) { S.saveState = state; paintSaveState(); }
function paintSaveState() {
  const el = $('#saveState');
  if (!el) return;
  el.className = 'save-state ' + S.saveState;
  el.textContent = S.saveState === 'idle' ? '' : t('save_' + S.saveState);
}

/* ═══ 11. Dialogs ═════════════════════════════════════════════════════════ */
function showDialog(html) {
  const dlg = $('#dlg');
  dlg.innerHTML = html;
  if (!dlg.open) dlg.showModal();
}
const dialogHead = title => `<div class="dlg-h"><h2>${esc(title)}</h2><button type="button" class="btn ghost sm" data-act="closeDlg" aria-label="${t('close')}">✕</button></div>`;
const dialogFoot = () => `<div class="dlg-f"><button type="button" class="btn" data-act="closeDlg">${t('cancel')}</button><button class="btn primary" type="submit">${t('save')}</button></div>`;

function openStudent(code) {
  const s = S.m.stu[code];
  if (!s) return;
  S.openStudent = code;
  const taken = new Set(s.regs.map(g => g.SubjectCode));
  const available = S.data.subjects.filter(x => !taken.has(x.Code) && x.Active !== 'No');
  const issueLines = s.allFlags.map(f => `<li class="${s.resolved.has(f) ? 'muted' : ''}"><span class="flag">${t(f)}</span> ${esc(issueDetail({ kind: f, s }))}</li>`).join('');
  showDialog(`${dialogHead(s.Name)}
  <div class="dlg-b">
    ${issueLines ? `<ul class="issue-list">${issueLines}</ul>` : ''}
    <dl class="kv">
      <dt>${t('code')}</dt><dd><span class="ltr">${esc(s.Code)}</span></dd>
      <dt>${t('email')}</dt><dd><span class="ltr">${esc(s.Email)}</span></dd>
      <dt>${t('stage')}</dt><dd>${stageName(s.Stage)} — ${groupName(s.Parallel)}</dd>
      <dt>${t('ects')}</dt><dd>${s.ects} / ${S.m.max || '—'}</dd>
      <dt>${t('submissions')}</dt><dd>${esc(s.Submissions)} — ${t('lastSubmit')}: ${fmtDate(s.LastSubmit)}</dd>
      ${s.PrereqAnswer ? `<dt>${t('prereq')}</dt><dd>${esc(s.PrereqAnswer)}</dd>` : ''}
      ${s.RetakeDetails ? `<dt>${t('retakeDetails')}</dt><dd>${esc(s.RetakeDetails)}</dd>` : ''}
      ${s.Note ? `<dt>${t('note')}</dt><dd>${esc(s.Note)}</dd>` : ''}
    </dl>
    ${s.regs.length ? `<div class="tbl-wrap"><table><tbody>
      ${s.regs.map(g => `<tr><td>${spec(g.SubjectCode)}</td><td>${esc(subjName(g.SubjectCode))}</td><td>${cap(g.Type)}</td>
        <td><select data-change="regStatus" data-id="${esc(g.RegID)}" aria-label="${t('status')}">${STATUSES.map(x => `<option value="${x}" ${g.Status === x ? 'selected' : ''}>${t(x)}</option>`).join('')}</select></td>
        <td><button class="btn ghost sm" data-act="delReg" data-id="${esc(g.RegID)}" title="${t('delete')}" aria-label="${t('delete')}">🗑</button></td></tr>`).join('')}
    </tbody></table></div>` : ''}
    <div>
      <p class="sub" style="margin:0 0 6px">${t('addReg')}</p>
      <div class="bar" style="margin:0">
        <select id="arSub" class="grow" aria-label="${t('subject')}">${available.map(x => `<option value="${esc(x.Code)}">${esc(x.Code)} — ${esc(x.Name)} (${stageName(x.Stage)})</option>`).join('')}</select>
        <select id="arType" aria-label="${t('type')}">${TYPES.map(x => `<option value="${x}">${t(x)}</option>`).join('')}</select>
        <button class="btn primary" data-act="addReg" data-code="${esc(s.Code)}" ${available.length ? '' : 'disabled'}>${t('add')}</button>
      </div>
    </div>
  </div>`);
}

function openUser(id) {
  const u = id ? S.data.users.find(x => x.UserID === id) : { Name: '', Username: '', Email: '', Role: 'TEACHER', Active: 'Yes' };
  const mine = {};
  S.data.assign.filter(a => a.UserID === id).forEach(a => { mine[a.SubjectCode] = a.Group; });
  const byStage = {};
  S.data.subjects.forEach(s => (byStage[s.Stage] = byStage[s.Stage] || []).push(s));
  const status = u.LastInvite ? t('invited', fmtDate(u.LastInvite)) : u.HasCode === 'Yes' ? t('hasCode') : t('notInvited');
  showDialog(`<form id="userForm" data-id="${esc(id || '')}">
  ${dialogHead(id ? u.Name : t('addTeacher'))}
  <div class="dlg-b">
    <div class="two">
      <label class="fld">${t('name')}<input type="text" name="Name" value="${esc(u.Name)}" required></label>
      <label class="fld">${t('email')}<input type="email" name="Email" value="${esc(u.Email)}" dir="ltr" autocapitalize="none" placeholder="name@uor.edu.krd"></label>
      <label class="fld">${t('username')}<input type="text" name="Username" value="${esc(u.Username === u.Email ? '' : u.Username)}" dir="ltr" autocapitalize="none" pattern="[a-zA-Z0-9._@+\\-]{3,}" placeholder="${t('usernameAuto')}"></label>
      <label class="fld">${t('role')}<select name="Role">
        <option value="TEACHER" ${u.Role === 'TEACHER' ? 'selected' : ''}>${t('TEACHER')}</option>
        <option value="HEAD" ${u.Role === 'HEAD' ? 'selected' : ''}>${t('HEAD')}</option></select></label>
      <label class="check"><input type="checkbox" name="Active" ${u.Active !== 'No' ? 'checked' : ''}> ${t('active')}</label>
      ${id ? '' : `<label class="check"><input type="checkbox" name="sendNow" checked> ${t('sendNow')}</label>`}
    </div>
    ${id && id !== S.user.id ? `<div class="invite-box">
      <p><b>${t('invitation')}:</b> ${esc(status)}${u.LastLogin ? ` — ${t('lastLogin')}: ${fmtDate(u.LastLogin)}` : ''}</p>
      <div class="btn-row">
        ${u.Email ? `<button type="button" class="btn sm primary" data-act="inviteOne" data-id="${esc(id)}">${u.LastInvite ? t('resend') : t('sendInvite')}</button>` : ''}
        <button type="button" class="btn sm" data-act="showCode" data-id="${esc(id)}">${t('makeCode')}</button>
      </div>
      <p class="muted" style="margin:6px 0 0;font-size:13px">${t('newCodeNote')}</p>
    </div>` : ''}
    <p class="sub" style="margin:0">${t('assignHint')}</p>
    ${Object.keys(byStage).sort().map(stage => `<div class="asg-stage"><h3>${stageName(stage)}</h3>
      ${byStage[stage].map(s => `<div class="asg-row">
        <input type="checkbox" id="as_${esc(s.Code)}" data-asg="${esc(s.Code)}" ${mine[s.Code] ? 'checked' : ''}>
        <label class="grow" for="as_${esc(s.Code)}">${spec(s.Code)} ${esc(s.Name)}</label>
        <select data-asg-grp="${esc(s.Code)}" aria-label="${t('group')}">${['All', 'Morning', 'Parallel'].map(g => `<option value="${g}" ${(mine[s.Code] || 'All') === g ? 'selected' : ''}>${t(g)}</option>`).join('')}</select>
      </div>`).join('')}</div>`).join('')}
  </div>
  ${dialogFoot()}
  </form>`);
}

function openPassword() {
  showDialog(`<form id="pwForm">
  ${dialogHead(t('changePw'))}
  <div class="dlg-b">
    <label class="fld">${t('currentPw')}<input type="password" name="current" autocomplete="current-password" dir="ltr" required></label>
    <label class="fld">${t('newPw')}<input type="password" name="next" autocomplete="new-password" dir="ltr" minlength="6" required></label>
  </div>
  ${dialogFoot()}
  </form>`);
}

/* ═══ 12. Actions ═════════════════════════════════════════════════════════ */
async function setStatus(ids, status) {
  let note = null;
  if (status === 'REJECTED') {
    note = prompt(t('rejectNote'), '');
    if (note === null) { if (S.openStudent) openStudent(S.openStudent); return; }
  }
  try {
    const r = await api('setStatus', { ids, status, note });
    const set = new Set(ids);
    S.data.regs.forEach(g => { if (set.has(g.RegID)) Object.assign(g, { Status: status, UpdatedAt: r.at }, note !== null ? { HeadNote: note } : {}); });
    refreshModel();
    toast(t('saved'));
  } catch (e) { toast(errText(e), true); }
}

async function resolveIssue(code, flag, undo) {
  const s = S.m.stu[code];
  const codes = flag === 'SAME_NAME' ? [s, ...s.twins].map(x => x.Code) : [code];
  try {
    for (const c of codes) {
      const r = await api('resolveFlag', { Code: c, Flag: flag, Resolved: !undo });
      S.data.students.find(x => x.Code === c).Resolved = r.Resolved;
    }
    refreshModel();
    toast(t('saved'));
  } catch (e) { toast(errText(e), true); }
}

/** Rebuilds the model after a local change, keeping scroll position and the open dialog. */
function refreshModel() {
  const y = window.scrollY;
  S.m = buildModel(S.data);
  render();
  window.scrollTo(0, y);
  if ($('#dlg').open && S.openStudent) openStudent(S.openStudent);
}

async function saveUser(form) {
  const btn = form.querySelector('[type=submit]');
  btn.disabled = true;
  try {
    const email = form.Email.value.trim();
    const r = await api('saveUser', {
      UserID: form.dataset.id || undefined, Name: form.Name.value, Email: email,
      Username: form.Username.value.trim() || email, Role: form.Role.value, Active: form.Active.checked ? 'Yes' : 'No'
    });
    const items = [...form.querySelectorAll('[data-asg]:checked')].map(c => ({
      SubjectCode: c.dataset.asg, Group: form.querySelector(`[data-asg-grp="${c.dataset.asg}"]`).value
    }));
    await api('setAssignments', { UserID: r.UserID, items });
    const inviteNow = !!(form.sendNow && form.sendNow.checked && email);
    $('#dlg').close();
    await reloadHead();
    render();
    toast(t('saved'));
    if (inviteNow) await sendInvites([r.UserID]);
  } catch (e) {
    toast(errText(e), true);
    btn.disabled = false;
  }
}

/** Sends invitation e-mails in groups of 10 and reports what happened. */
async function sendInvites(ids, btn) {
  if (!ids.length) return;
  const appUrl = S.data.settings.AppUrl || (location.origin + location.pathname.replace(/index\.html$/, ''));
  const label = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = t('sending'); }
  let ok = 0;
  const failed = [];
  try {
    for (let i = 0; i < ids.length; i += 10) {
      const r = await api('sendInvites', { ids: ids.slice(i, i + 10), appUrl });
      r.results.forEach(x => (x.ok ? ok++ : failed.push(x)));
      if (btn) btn.textContent = t('sendingN', ok + failed.length, ids.length);
    }
  } catch (e) {
    toast(errText(e), true);
  }
  try { await reloadHead(); } catch (e) { /* keep the current view */ }
  if ($('#dlg').open) $('#dlg').close();
  render();
  if (failed.length) toast(t('sentSome', ok, ids.length) + ' ' + [...new Set(failed.map(f => errText(f.error)))].join(' '), true);
  else if (ok) toast(t('sentAll', ok));
  if (btn && document.body.contains(btn)) { btn.disabled = false; btn.textContent = label; }
}

/** Creates a code without e-mail and shows it once (for teachers who cannot receive the e-mail). */
async function showNewCode(id) {
  if (!confirm(t('confirmNewCode'))) return;
  try {
    const r = await api('resetCode', { UserID: id });
    const u = S.data.users.find(x => x.UserID === id) || {};
    const appUrl = S.data.settings.AppUrl || (location.origin + location.pathname.replace(/index\.html$/, ''));
    const text = `${appUrl}\n${t('loginId')}: ${r.username}\n${t('codeLabel')}: ${r.code}`;
    showDialog(`${dialogHead(t('newCodeTitle'))}
      <div class="dlg-b"><p style="margin:0"><b>${esc(u.Name)}</b></p>
        <dl class="kv"><dt>${t('loginId')}</dt><dd><span class="ltr">${esc(r.username)}</span></dd>
          <dt>${t('codeLabel')}</dt><dd><b class="ltr code-big">${esc(r.code)}</b></dd></dl>
        <p class="sub" style="margin:0">${t('codeOnce')}</p></div>
      <div class="dlg-f"><button type="button" class="btn" data-act="copyCode" data-text="${esc(text)}">${t('copy')}</button>
        <button type="button" class="btn primary" data-act="closeDlg">${t('close')}</button></div>`);
    await reloadHead();
    render();
  } catch (e) { toast(errText(e), true); }
}

/* ═══ 13. Export: CSV and Word ════════════════════════════════════════════ */
const libs = {};
/** Loads a vendor library once, only when the first file of that kind is made. */
function loadLib(src, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  libs[src] = libs[src] || new Promise((resolve, reject) => {
    const el = document.createElement('script');
    el.src = src;
    el.onload = () => (window[globalName] ? resolve(window[globalName]) : reject(new Error('EXPORT')));
    el.onerror = () => { delete libs[src]; reject(new Error('EXPORT')); };
    document.head.appendChild(el);
  });
  return libs[src];
}

function download(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

/** iPhone: hand the file to the share sheet (Save to Files, WhatsApp, Mail …) from a fresh tap. Elsewhere: download. */
function deliver(blob, filename) {
  if (isIOS() && navigator.canShare) {
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare({ files: [file] })) {
      S.pendingFile = file;
      showDialog(`${dialogHead(t('fileReady'))}
        <div class="dlg-b"><p class="filecard"><b class="ltr">${esc(filename)}</b><br><span class="muted">${Math.max(1, Math.round(blob.size / 1024))} KB</span></p>
        <p class="sub" style="margin:0">${t('iosSaveHint')}</p></div>
        <div class="dlg-f"><button type="button" class="btn primary" data-act="shareFile">${t('saveShare')}</button></div>`);
      return;
    }
  }
  download(blob, filename);
}

const letterhead = () => [t('uni'), t('college'), t('dept')];
const footerText = () => `${CONFIG.SHORT_NAME} — ${t('generated')} ${fmtDate(new Date().toISOString())}`;
const fileSafe = s => String(s || '').replace(/[^\w\u0600-\u06FF]+/g, '-').replace(/^-+|-+$/g, '');
const sortLabel = () => t('sort_' + S.f.sort);

/** Word + Excel content for one subject (what is on screen) or for every subject (head). */
function sheetExport(codes, all) {
  if (!codes.length || !codes[0]) throw new Error('CHOOSE');
  const lim = limits(), when = fmtDate(new Date().toISOString(), false);
  const sheets = [], sections = [], summary = [];
  const green = '2B7A4B', red = 'B0271F';
  const columns = [
    { label: '#', weight: 0.45, align: 'center' }, { label: t('name'), weight: 2.7 }, { label: t('code'), weight: 1.4 },
    { label: t('group'), weight: 1.05, align: 'center' }, { label: t('type'), weight: 1.15, align: 'center' },
    { label: `${t('scoreA')} (${lim.maxA})`, weight: 1.25, align: 'center' }, { label: `${t('scoreB')} (${lim.maxB})`, weight: 1.25, align: 'center' },
    { label: `${t('total')} (${lim.maxA + lim.maxB})`, weight: 1.2, align: 'center' }, { label: t('result'), weight: 1.1, align: 'center' }
  ];
  const filters = !all && (S.f.result || S.f.group)
    ? [S.f.result ? (S.f.result === 'NONE' ? t('notGraded') : t(S.f.result)) : '', S.f.group ? (S.f.group === 'Yes' ? t('Parallel') : t('Morning')) : ''].filter(Boolean).join(sep())
    : '';

  codes.forEach((code, k) => {
    const sub = subjectOf(code) || { Code: code, Name: '', Stage: '', ECTS: '' };
    const rows = all ? sortRows(sheetRowsFor(code), S.f.sort) : visibleSheetRows(code, false);
    const st = stats(rows), teachers = teacherNamesFor(code);
    const meta = [
      `${S.semester} — ${stageName(sub.Stage)} — ${sub.ECTS} ECTS`,
      `${t('teacher')}: ${teachers.join(sep()) || '—'}`,
      `${t('sortBy')}: ${sortLabel()}${filters ? ` — ${t('onlyShowing')}: ${filters}` : ''} — ${when}`
    ];
    const signatures = [{ label: t('sigTeacher'), name: teachers.join(sep()) }, { label: t('sigHead') }, { label: t('sigRegistrar') }];
    sheets.push({
      name: code, letterhead: letterhead(), title: `${code} — ${sub.Name}`, meta, limits: lim, signatures,
      rows: rows.map((r, i) => ({ n: i + 1, name: r.Name, code: r.Code, group: groupName(r.Parallel), type: t(r.Type + '_s'), A: r.a, B: r.b }))
    });
    sections.push({
      heading: all ? `${code} — ${sub.Name}` : undefined, note: all ? meta.slice(0, 2).join(' — ') : undefined, pageBreakBefore: all && k > 0,
      facts: [[t('nStudents'), st.count], [t('graded'), st.graded], [t('average'), st.avg === '' ? '—' : st.avg],
              [t('highest'), st.max === '' ? '—' : st.max], [t('PASS'), st.passed], [t('FAIL'), st.failed]],
      columns, signatures, emptyText: t('noStudents'),
      rows: rows.map((r, i) => [i + 1, r.Name, r.Code, groupName(r.Parallel), t(r.Type + '_s'), fmtScore(r.a), fmtScore(r.b),
        { t: r.total == null ? '—' : r.total, bold: true, color: r.result === 'PASS' ? green : r.result === 'FAIL' ? red : undefined },
        r.result ? { t: t(r.result), bold: true, color: r.result === 'PASS' ? green : red } : { t: t('notGraded') }])
    });
    summary.push([code, sub.Name, stageName(sub.Stage), teachers.join(sep()) || '—', st.count, st.graded,
                  st.avg === '' ? '' : st.avg, st.passed, st.failed, st.rate === '' ? '' : st.rate]);
  });

  const one = subjectOf(codes[0]) || {};
  const labels = {
    hash: '#', name: t('name'), code: t('code'), group: t('group'), type: t('type'), a: t('scoreA'), b: t('scoreB'),
    total: t('total'), result: t('result'), pass: t('PASS'), fail: t('FAIL'), invalid: t('invalidScore'), range: m => t('rangeMsg', m),
    stats: { students: t('nStudents'), graded: t('graded'), average: t('average'), highest: t('highest'), lowest: t('lowest'),
             passed: t('PASS'), failed: t('FAIL'), rate: t('passRate') }
  };
  return {
    name: all ? `${CONFIG.SHORT_NAME}_${t('allMarksFile')}_${today()}` : `${codes[0]}_${fileSafe(one.Name)}_${today()}`,
    doc: {
      rtl: S.lang === 'ku', letterhead: letterhead(), footer: footerText(), emptyText: t('noStudents'),
      title: all ? t('allMarksTitle') : `${codes[0]} — ${one.Name || ''}`,
      meta: all ? [`${S.semester} — ${when}`, `${t('sortBy')}: ${sortLabel()}`] : sheets[0].meta,
      sections
    },
    xlsx: {
      rtl: S.lang === 'ku', title: all ? t('allMarksTitle') : sheets[0].title, labels, sheets,
      summary: all ? {
        name: t('summaryName'), title: t('allMarksTitle'), letterhead: letterhead(), meta: [`${S.semester} — ${when}`],
        columns: [{ label: t('code'), width: 12, ltr: true }, { label: t('subject'), width: 28 }, { label: t('stage'), width: 14 },
                  { label: t('teacher'), width: 32 }, { label: t('nStudents'), width: 11, center: true }, { label: t('graded'), width: 11, center: true },
                  { label: t('average'), width: 11, center: true }, { label: t('PASS'), width: 11, center: true },
                  { label: t('FAIL'), width: 11, center: true }, { label: t('passRate'), width: 12, center: true, fmt: '0%' }],
        rows: summary
      } : undefined
    }
  };
}

function issuesExport() {
  const sections = [], csv = [[t('issues'), t('name'), t('code'), t('stage'), t('detail')]];
  ISSUE_ORDER.forEach(kind => {
    const items = visibleIssues(kind);
    if (!items.length) return;
    const desc = kind === 'OVER_ECTS' ? t('OVER_ECTS_d', S.m.max) : t(kind + '_d');
    if (kind === 'UNIDENTIFIED') {
      const rows = items.map(i => [fmtDate(i.row.Time), i.row.Name, i.row.Email]);
      rows.forEach(r => csv.push([t(kind), r[1], '', '', r[2]]));
      sections.push({ heading: `${t(kind)} (${rows.length})`, note: desc, rows,
        columns: [{ label: t('time'), weight: 2 }, { label: t('name'), weight: 2.5 }, { label: t('email'), weight: 3 }] });
    } else if (kind === 'NO_TEACHER') {
      const rows = items.map(i => [i.subject.Code, i.subject.Name, stageName(i.subject.Stage), i.count]);
      rows.forEach(r => csv.push([t(kind), r[1], r[0], r[2], r[3]]));
      sections.push({ heading: `${t(kind)} (${rows.length})`, note: desc, rows,
        columns: [{ label: t('code'), weight: 1.5 }, { label: t('subject'), weight: 3 }, { label: t('stage'), weight: 1.5 }, { label: t('nStudents'), weight: 1 }] });
    } else {
      const rows = items.map(i => [i.s.Name, i.s.Code, stageName(i.s.Stage), issueDetail(i) + (i.done ? ` (${t('checked')})` : '')]);
      rows.forEach(r => csv.push([t(kind), ...r]));
      sections.push({ heading: `${t(kind)} (${rows.length})`, note: desc, rows,
        columns: [{ label: t('name'), weight: 2.4 }, { label: t('code'), weight: 1.5 }, { label: t('stage'), weight: 1.4 }, { label: t('detail'), weight: 4 }] });
    }
  });
  return {
    name: `data-issues_${today()}`, csv,
    doc: { rtl: S.lang === 'ku', letterhead: letterhead(), title: t('issuesReport'),
           meta: [`${S.semester} — ${fmtDate(new Date().toISOString(), false)}`, `${S.m.openIssues} ${t('nIssues')}`],
           sections, emptyText: t('noIssues'), footer: footerText() }
  };
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function runExport(btn, scope, format) {
  let x;
  try {
    if (scope === 'issues') x = issuesExport();
    else if (scope === 'sheet') x = sheetExport([S.f.subject], false);
    else x = sheetExport(sheetSubjects().map(s => s.Code), true);
  } catch (e) { return toast(errText(e), true); }
  if (format === 'csv') return deliver(new Blob([Exporters.toCSV(x.csv)], { type: 'text/csv;charset=utf-8' }), x.name + '.csv');

  const label = btn.textContent;
  btn.disabled = true;
  btn.textContent = t('preparing');
  try {
    if (format === 'xlsx') {
      const lib = await loadLib(CONFIG.EXCEL_LIB, 'ExcelJS');
      const buffer = await Exporters.buildWorkbook(lib, x.xlsx).xlsx.writeBuffer();
      deliver(new Blob([buffer], { type: XLSX_MIME }), x.name + '.xlsx');
    } else {
      const lib = await loadLib(CONFIG.DOCX_LIB, 'docx');
      deliver(await lib.Packer.toBlob(Exporters.buildDocument(lib, x.doc)), x.name + '.docx');
    }
  } catch (e) {
    toast(t('E_EXPORT'), true);
  } finally {
    btn.disabled = false;
    btn.textContent = label;
  }
}

/* ═══ 14. Events ══════════════════════════════════════════════════════════ */
const ACTIONS = {
  lang: () => toggleLang(),
  signout: () => signOut(),
  pw: () => openPassword(),
  dismissIos: () => { store('iosHint', '1'); updateChrome(); },
  install: async () => { if (!S.installEvent) return; S.installEvent.prompt(); await S.installEvent.userChoice; S.installEvent = null; updateChrome(); },
  applyUpdate: () => { S.reloadOnUpdate = true; S.updateWorker?.postMessage('SKIP_WAITING'); },
  tab: el => { S.tab = el.dataset.tab; S.f.q = ''; render(); window.scrollTo(0, 0); },
  openSheet: el => {
    Object.assign(S.f, { subject: el.dataset.code, q: '', result: '', group: '' });
    S.tab = 'sheets';
    render();
    window.scrollTo(0, 0);
  },
  closeSheet: () => { S.f.subject = ''; S.f.q = ''; render(); window.scrollTo(0, 0); },
  togglePw: el => {
    const input = el.closest('.pwrow').querySelector('input');
    input.type = input.type === 'password' ? 'text' : 'password';
    el.textContent = input.type === 'password' ? t('showPw') : t('hidePw');
  },
  inviteOne: el => sendInvites([el.dataset.id], el),
  inviteMany: el => {
    const others = S.data.users.filter(u => u.Email && u.Active !== 'No' && u.UserID !== S.user.id);
    const list = el.dataset.which === 'new' ? others.filter(u => !u.LastInvite) : others;
    if (list.length && confirm(t('confirmInvite', list.length))) sendInvites(list.map(u => u.UserID), el);
  },
  showCode: el => showNewCode(el.dataset.id),
  copyCode: async el => {
    try { await navigator.clipboard.writeText(el.dataset.text); toast(t('copied')); }
    catch (e) { toast(t('copyFailed'), true); }
  },
  shareFile: async () => {
    const file = S.pendingFile;
    if (!file) return;
    try { await navigator.share({ files: [file], title: file.name }); $('#dlg').close(); }
    catch (e) { if (e.name !== 'AbortError') { download(file, file.name); $('#dlg').close(); } }
  },
  status: el => setStatus(el.dataset.ids.split('|').filter(Boolean), el.dataset.status),
  resolve: el => resolveIssue(el.dataset.code, el.dataset.flag, el.dataset.done === 'true'),
  student: el => openStudent(el.dataset.code),
  closeDlg: () => $('#dlg').close(),
  export: el => runExport(el, el.dataset.scope, el.dataset.format),
  print: () => window.print(),
  editUser: el => openUser(el.dataset.id),
  delReg: async el => {
    if (!confirm(t('confirmDelete'))) return;
    try {
      await api('deleteRegistration', { id: el.dataset.id });
      S.data.regs = S.data.regs.filter(g => g.RegID !== el.dataset.id);
      refreshModel();
      toast(t('saved'));
    } catch (e) { toast(errText(e), true); }
  },
  addReg: async el => {
    el.disabled = true;
    try {
      await api('addRegistration', { Code: el.dataset.code, SubjectCode: $('#arSub').value, Type: $('#arType').value });
      await reloadHead();
      refreshModel();
      toast(t('saved'));
    } catch (e) { toast(errText(e), true); el.disabled = false; }
  },
  sync: async el => {
    el.disabled = true;
    el.textContent = t('syncing');
    try {
      const r = await api('sync');
      await reloadHead();
      render();
      toast(t('syncDone', r));
    } catch (e) { toast(errText(e), true); el.disabled = false; el.textContent = t('sync'); }
  },
  saveSubject: async el => {
    const row = el.closest('tr');
    const v = k => row.querySelector(`[data-k="${k}"]`);
    el.disabled = true;
    try {
      await api('saveSubject', { Code: el.dataset.code, Name: v('Name').value, Stage: v('Stage').value, ECTS: v('ECTS').value,
                                 Prerequisite: v('Prerequisite').value, Active: v('Active').checked ? 'Yes' : 'No' });
      await reloadHead();
      toast(t('saved'));
    } catch (e) { toast(errText(e), true); }
    el.disabled = false;
  }
};

const FORMS = {
  loginForm: f => doLogin(f),
  userForm: f => saveUser(f),
  pwForm: async f => {
    try {
      const r = await api('changePassword', { current: f.current.value, next: f.next.value });
      S.token = r.token;                       // the old session ends when the code changes; keep this one signed in
      store('token', r.token);
      $('#dlg').close();
      toast(t('pwChanged'));
    } catch (e) { toast(errText(e), true); }
  },
  setForm: async f => {
    try {
      S.data.settings = await api('saveSettings', { Semester: f.Semester.value, MaxECTS: f.MaxECTS.value, FormSheet: f.FormSheet.value, MaxA: f.MaxA.value, MaxB: f.MaxB.value, PassMark: f.PassMark.value, AppUrl: f.AppUrl.value });
      S.semester = S.data.settings.Semester;
      refreshModel();
      toast(t('saved'));
    } catch (e) { toast(errText(e), true); }
  }
};

document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]');
  if (!el || !ACTIONS[el.dataset.act]) return;
  if (el.tagName === 'A') e.preventDefault();
  ACTIONS[el.dataset.act](el);
});

document.addEventListener('change', e => {
  const el = e.target;
  if (el.classList.contains('sc') && el.dataset.reg) { onScoreChange(el); return; }
  if (el.dataset.change === 'regStatus') { setStatus([el.dataset.id], el.value); return; }
  if (el.dataset.f && el.type !== 'search') {
    S.f[el.dataset.f] = el.type === 'checkbox' ? el.checked : el.value;
    if (el.dataset.f === 'unlock') render(); else refreshList();
  }
});

document.addEventListener('keydown', e => {
  if (e.key !== 'Enter' || !e.target.classList || !e.target.classList.contains('sc')) return;
  e.preventDefault();
  const all = [...document.querySelectorAll('input.sc:not([readonly])')], next = all[all.indexOf(e.target) + 1];
  if (next) { next.focus(); next.select(); } else e.target.blur();
});

document.addEventListener('input', e => {
  if (e.target.dataset.f === 'q') { S.f.q = e.target.value; refreshList(); }
});

document.addEventListener('submit', e => {
  e.preventDefault();
  const handler = FORMS[e.target.id];
  if (handler) handler(e.target);
});

$('#dlg').addEventListener('close', () => { S.openStudent = null; });

/* ═══ 15. PWA: service worker, install, updates, offline ══════════════════ */
function setupPWA() {
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});   // ask the phone to keep the saved sign-in
  window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); S.installEvent = e; updateChrome(); });
  window.addEventListener('appinstalled', () => { S.installEvent = null; updateChrome(); });
  window.addEventListener('online', () => { updateChrome(); flushOutbox(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flushOutbox(); });
  window.addEventListener('offline', updateChrome);

  if (!('serviceWorker' in navigator) || location.protocol === 'file:') return;
  navigator.serviceWorker.register('sw.js').then(reg => {
    const offer = worker => { S.updateWorker = worker; updateChrome(); };
    if (reg.waiting && navigator.serviceWorker.controller) offer(reg.waiting);
    reg.addEventListener('updatefound', () => {
      const w = reg.installing;
      w?.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) offer(w); });
    });
  }).catch(() => {});
  // Reload only when the user pressed "Reload" — not when the very first worker takes control.
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (S.reloadOnUpdate) { S.reloadOnUpdate = false; location.reload(); } });
}

/* ═══ 16. Start ═══════════════════════════════════════════════════════════ */
applyLang();
setupPWA();
if (S.token && S.user) loadData(); else renderLogin();

/* ===========================
   Alsami Workshop - Frontend
   Final app.js (stable v4)
   =========================== */

/* ---------- Config ---------- */
let API_BASE = (localStorage.getItem('api_base') || '').trim();
if (!API_BASE) {
  // على Render: استخدم مسار نسبي /api — محليًا: 127.0.0.1:9000/api
  const onRender = /alsami-workshop-backend\.onrender\.com$/i.test(location.hostname);
  API_BASE = onRender ? '/api' : 'http://127.0.0.1:9000/api';
}

/* ---------- Helpers ---------- */
const $  = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const mountRoot = $('#viewRoot');

function toast(msg, type = 'info') {
  const el = document.createElement('div');
  el.textContent = msg;
  el.className = `toast ${type === 'error' ? 'error' : ''}`;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2200);
}

function escapeHTML(s) {
  return String(s).replace(/[&<>"]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
}

function withTimeout(promise, ms = 12000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort('timeout'), ms);
  return promise(ctrl.signal).finally(() => clearTimeout(t));
}

async function fetchJSON(url, opts = {}, timeoutMs = 12000) {
  return withTimeout(async (signal) => {
    const r = await fetch(url, { ...opts, signal });
    if (!r.ok) throw new Error(await r.text().catch(()=>'HTTP '+r.status));
    return r.json();
  }, timeoutMs);
}

async function fetchBlob(url, opts = {}, timeoutMs = 20000) {
  return withTimeout(async (signal) => {
    const r = await fetch(url, { ...opts, signal });
    if (!r.ok) throw new Error(await r.text().catch(()=>'HTTP '+r.status));
    return r;
  }, timeoutMs);
}

/* ---------- Global Error Guards ---------- */
window.addEventListener('error', (e) => {
  console.error('GlobalError:', e.error || e.message);
  toast('حدث خطأ غير متوقع بالواجهة', 'error');
});
window.addEventListener('unhandledrejection', (e) => {
  console.error('UnhandledRejection:', e.reason);
  toast('تعذر إتمام العملية', 'error');
});

/* ---------- Online status ---------- */
function setOnline(ok) {
  $('#onlineDot')?.classList.toggle('online', !!ok);
  const t = $('#onlineText'); if (t) t.textContent = ok ? 'متصل' : 'غير متصل';
  const s = $('#srvStatus'); if (s) { s.textContent = ok ? 'جيد' : 'متعطّل'; s.className = `tag ${ok?'ok':''}`; }
}
async function ping() {
  try { const r = await fetch(`${API_BASE}/health`); setOnline(r.ok); }
  catch { setOnline(false); }
}

/* ---------- Router ---------- */
const routes = {
  home:           '#tpl-home',
  engines:        '#tpl-engines',

  'eng-supply':   '#tpl-eng-supply',
  'eng-issue':    '#tpl-eng-issue',
  'eng-rehab':    '#tpl-eng-rehab',
  'eng-check':    '#tpl-eng-check',
  'eng-upload':   '#tpl-eng-upload',
  'eng-lathe':    '#tpl-eng-lathe',
  'eng-pump':     '#tpl-eng-pump',
  'eng-elec':     '#tpl-eng-elec',

  // Generators + aliases to avoid "لا يوجد قالب للعرض gen-home"
  generators:     '#gen-home-tpl',
  'gen-home':     '#gen-home-tpl',
  genhome:        '#gen-home-tpl',

  'gen-supply':   '#gen-supply-tpl',
  'gen-issue':    '#gen-issue-tpl',
  'gen-inspect':  '#gen-inspect-tpl',

  reports:        '#tpl-reports',
};

let navLock = false;
function setActive(view) {
  if (navLock) return;
  navLock = true;
  try {
    // highlight sidebar
    $$('.nav-item').forEach(x => x.classList.toggle('active', x.dataset.view === view));
    show(view);
  } finally {
    navLock = false;
  }
}

function show(view) {
  const tplSel = routes[view];
  const tpl = $(tplSel);
  if (!tpl) {
    mountRoot.innerHTML = `<div class="card">⚠️ لا يوجد قالب للعرض: ${escapeHTML(view)}</div>`;
    return;
  }
  mountRoot.innerHTML = '';
  mountRoot.appendChild(tpl.content.cloneNode(true));

  // Common binds inside the loaded view
  bindView(view);
}

/* ---------- View Binders ---------- */
function bindView(view) {
  // Back buttons within any view:
  $$('[data-return]', mountRoot).forEach(btn => btn.addEventListener('click', () => setActive('home')));

  // Quick tiles navigation (engines dashboard)
  $$('.tile', mountRoot).forEach(tile => {
    const v = tile.dataset.view || tile.dataset.go;
    if (v) tile.addEventListener('click', () => setActive(v));
  });

  // Forms: generic binder
  $$('form.form', mountRoot).forEach(bindForm);

  // Home-specific
  if (view === 'home') {
    renderLast3();
    if (window.__lastSearchHtml) $('#searchOutput').innerHTML = window.__lastSearchHtml;
  }

  // Reports-specific
  if (view === 'reports') {
    $('#openExport')?.addEventListener('click', openExport);
    $('#runRepair')?.addEventListener('click', runRepair);
  }

  // Generators hub: ensure back buttons in generator templates
  if (view === 'gen-supply' || view === 'gen-issue' || view === 'gen-inspect') {
    $('.btn-back', mountRoot)?.addEventListener('click', () => setActive('generators'));
    bindGeneratorForms(view);
  }
}

/* ---------- Last 3 tables ---------- */
function tableHTML(items, key='serial', site='prevSite') {
  if (!items?.length) return '<div class="muted">لا توجد بيانات</div>';
  const rows = items.map(x =>
    `<tr><td>${escapeHTML(x[key]||'')}</td><td>${escapeHTML(x[site]||'')}</td></tr>`
  ).join('');
  return `<table class="table"><thead><tr><th>الرقم</th><th>الموقع السابق</th></tr></thead><tbody>${rows}</tbody></table>`;
}

async function renderLast3() {
  try {
    const [e, g] = await Promise.all([
      fetchJSON(`${API_BASE}/last3/engines`).catch(() => ({items:[]})),
      fetchJSON(`${API_BASE}/last3/generators`).catch(() => ({items:[]})),
    ]);
    $('#last3Engines')?.replaceChildren();
    $('#last3Engines')?.insertAdjacentHTML('afterbegin', tableHTML(e.items, 'serial', 'prevSite'));
    $('#last3Generators')?.replaceChildren();
    $('#last3Generators')?.insertAdjacentHTML('afterbegin', tableHTML(g.items, 'code', 'prevSite'));
  } catch {
    $('#last3Engines')?.insertAdjacentHTML('afterbegin','<div class="muted">خطأ في التحميل</div>');
    $('#last3Generators')?.insertAdjacentHTML('afterbegin','<div class="muted">خطأ في التحميل</div>');
  }
}

/* ---------- Search ---------- */
async function doSearch() {
  const key = ($('#searchInput')?.value || '').trim();
  if (!key) { toast('اكتب الرقم التسلسلي أو كود المولد','error'); return; }
  const btn = $('#btnSearch');
  const old = btn?.innerHTML;
  if (btn) btn.innerHTML = '⏳';
  try {
    const data = await fetchJSON(`${API_BASE}/search/${encodeURIComponent(key)}`);
    const html = buildSearchHtml(key, data);
    if ($('#searchOutput')) $('#searchOutput').innerHTML = html;
    window.__lastSearchHtml = html;
    setActive('home');
  } catch (e) {
    console.error(e);
    toast('حدث خطأ أثناء البحث','error');
  } finally {
    if (btn) btn.innerHTML = old;
  }
}

function buildSearchHtml(key, data) {
  const engines = data?.engines || {}, gens = data?.generators || {};
  const card = (title, arr) => `
    <div class="card">
      <div class="card-title">${title}</div>
      ${!arr?.length ? '<div class="muted">لا توجد سجلات</div>' :
        `<div class="mt">${arr.map(r => `
          <div class="row" style="flex-wrap:wrap;gap:14px;">
            ${Object.entries(r).map(([k,v]) => `<span class="tag"><b>${escapeHTML(k)}:</b> ${escapeHTML(v??'')}</span>`).join('')}
          </div>
        `).join('')}</div>`
      }
    </div>
  `;
  return `
    <div class="row" style="justify-content:space-between;align-items:center">
      <div>نتائج البحث عن: <span class="chip">${escapeHTML(key)}</span></div>
      <button class="btn-light" onclick="openExport()">تصدير النتائج</button>
    </div>
    <div class="grid-2 mt">
      <div>
        ${card('محركات - توريد', engines.supply)}
        ${card('محركات - صرف', engines.issue)}
        ${card('محركات - تأهيل', engines.rehab)}
        ${card('محركات - فحص', engines.check)}
        ${card('محركات - رفع', engines.upload)}
        ${card('محركات - مخرطة', engines.lathe)}
        ${card('محركات - بمبات ونوزلات', engines.pump)}
        ${card('محركات - كهرباء', engines.electrical)}
      </div>
      <div>
        ${card('مولدات - توريد', gens.supply)}
        ${card('مولدات - صرف', gens.issue)}
        ${card('مولدات - فحص/رفع', gens.inspect)}
      </div>
    </div>
  `;
}

/* ---------- Export (direct to /api/export/xlsx) ---------- */
function openExport(){ $('#exportModal')?.classList.remove('hidden'); }
function closeExport(){ $('#exportModal')?.classList.add('hidden'); }

async function doExport() {
  const type = ($$('input[name="exportType"]').find(i=>i.checked)?.value) || 'engines';
  const from = $('#fromDate')?.value || '';
  const to   = $('#toDate')?.value   || '';

  const btn = $('#doExport'); const old = btn?.innerHTML; if (btn) btn.innerHTML = '⏳';
  try {
    // نرسل مباشرة لنقطة الـ backend الحالية
    const payload = {
      filename: suggestName(type, from, to),
      scope: (type === 'both') ? 'both' : (type === 'generators' ? 'generators' : 'engines'),
      date_from: from || null,
      date_to: to || null
    };
    const res = await fetchBlob(`${API_BASE}/export/xlsx`, {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(payload)
    });
    const blob = await res.blob();
    const disp = res.headers.get('Content-Disposition') || '';
    const name = parseFileName(disp) || payload.filename || 'report.xlsx';

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);

    closeExport(); toast('تم التصدير بنجاح ✅');
  } catch (e) {
    console.error('Export error:', e);
    toast('فشل التصدير','error');
  } finally {
    if (btn) btn.innerHTML = old;
  }
}

function suggestName(type, from, to) {
  const now = new Date().toISOString().slice(0,10);
  const t = type==='engines'?'محركات': type==='generators'?'مولدات':'محركات-ومولدات';
  const range = (from||to) ? `_من-${from||'—'}_إلى-${to||'—'}` : '';
  return `تقرير_${t}${range}_${now}.xlsx`;
}
function parseFileName(d) {
  const m = /filename\*?=(?:UTF-8''|)([^;]+)/i.exec(d||'');
  if(!m) return '';
  try { return decodeURIComponent(m[1].replace(/(^"|"$)/g,'')); } catch { return m[1]; }
}

/* ---------- Admin repair (schema) ---------- */
async function runRepair() {
  try {
    const j = await fetchJSON(`${API_BASE}/admin/repair`, { method:'POST' });
    if (j?.ok) toast('تم الإصلاح بنجاح ✅'); else toast('فشل الإصلاح','error');
  } catch {
    toast('خطأ اتصال','error');
  }
}

/* ---------- IndexedDB (offline) ---------- */
const DB_NAME = 'alsami_offline_v1';
const STORES  = [
  // Engines
  'eng_supply','eng_issue','eng_rehab','eng_check','eng_upload','eng_lathe','eng_pump','eng_electrical',
  // Generators
  'gen_supply','gen_issue','gen_inspect'
];
let idb;

function idbOpen() {
  return new Promise((resolve,reject) => {
    const req = indexedDB.open(DB_NAME, 2);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      STORES.forEach(s => { if (!db.objectStoreNames.contains(s)) db.createObjectStore(s, { keyPath:'_id', autoIncrement:true }); });
      if (!db.objectStoreNames.contains('queue')) db.createObjectStore('queue', { keyPath:'_id', autoIncrement:true });
    };
    req.onsuccess = e => { idb = e.target.result; resolve(); };
    req.onerror   = e => reject(e);
  });
}
function idbAdd(store, payload) {
  return new Promise((resolve,reject) => {
    const tx = idb.transaction([store],'readwrite');
    tx.objectStore(store).add({ createdAt: Date.now(), ...payload });
    tx.oncomplete=()=>resolve(true);
    tx.onerror   =e=>reject(e);
  });
}
function idbGetAll(store) {
  return new Promise((resolve,reject) => {
    const tx  = idb.transaction([store],'readonly');
    const req = tx.objectStore(store).getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror   = e => reject(e);
  });
}
function idbClear(store) {
  return new Promise((resolve,reject) => {
    const tx  = idb.transaction([store],'readwrite');
    const req = tx.objectStore(store).clear();
    req.onsuccess = () => resolve(true);
    req.onerror   = e => reject(e);
  });
}

async function syncAll() {
  try {
    const items = [];
    for (const store of STORES) {
      const rows = await idbGetAll(store);
      rows.forEach(r => items.push({ store, payload: r }));
      if (rows.length) await idbClear(store);
    }
    if (!items.length) { toast('لا توجد بيانات لمزامنتها','error'); return; }
    const j = await fetchJSON(`${API_BASE}/sync/batch`, {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ items })
    });
    if (j?.ok) toast(`تمت المزامنة: ${j.saved} سجل ✅`);
    else toast('فشل المزامنة','error');
  } catch (e) {
    console.error('syncAll error:', e);
    toast('مشكلة اتصال أثناء المزامنة','error');
  }
}

/* ---------- Forms binder (generic) ---------- */
function bindForm(form) {
  const store = form.dataset.store; // مثال: eng_supply
  if (!store) return;

  // زر مزامنة عام داخل النموذج (اختياري)
  form.querySelector('[data-sync]')?.addEventListener('click', syncAll);

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(form).entries());

    // مفتاح أساسي حسب النوع
    const isGen = store.startsWith('gen_');
    const keyField = isGen ? 'code' : 'serial';
    const keyVal   = (data[keyField] || '').trim();
    if (!keyVal) { toast(`حقل ${isGen?'الكود':'الرقم التسلسلي'} مطلوب`,'error'); return; }

    try {
      await idbAdd(store, data);
      toast('تم الحفظ محليًا ✅');

      // مزامنة مباشرة + بحث فوري بعد الحفظ
      const payload = { items: [{ store, payload: data }], returnKey: keyVal };
      const j = await fetchJSON(`${API_BASE}/sync/batch`, {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify(payload)
      }).catch(()=>null);

      // نفّذ بحث أفوريًا حتى لو فشلت المزامنة (يظل محليًا)
      await quickSearchAndShow(keyVal);
      form.reset();
    } catch (err) {
      console.error(err);
      toast('فشل الحفظ','error');
    }
  });
}

async function quickSearchAndShow(key) {
  try {
    const data = await fetchJSON(`${API_BASE}/search/${encodeURIComponent(key)}`);
    const html = buildSearchHtml(key, data);
    if ($('#searchOutput')) $('#searchOutput').innerHTML = html;
    window.__lastSearchHtml = html;
    setActive('home');
  } catch (e) {
    console.warn('quickSearchAndShow:', e);
    // لا نكسر الواجهة لو البحث فشل (مثلاً السيرفر غير متاح)
  }
}

/* ---------- Generator forms (explicit bind) ---------- */
function bindGeneratorForms(view) {
  if (view === 'gen-supply') {
    const form = $('#gen-supply-form', mountRoot);
    if (!form) return;
    form.onsubmit = async (e) => {
      e.preventDefault();
      const data = {
        code:      $('#gen_code').value.trim(),
        gType:     $('#gen_gType').value.trim(),
        model:     $('#gen_model').value.trim(),
        prevSite:  $('#gen_prevSite').value.trim(),
        supDate:   $('#gen_supDate').value,
        supplier:  $('#gen_supplier').value.trim(),
        vendor:    $('#gen_vendor').value.trim(),
        notes:     $('#gen_notes').value.trim(),
      };
      if (!data.code) { toast('الكود مطلوب','error'); return; }
      try {
        await idbAdd('gen_supply', data);
        const payload = { items: [{ store:'gen_supply', payload: data }] };
        await fetchJSON(`${API_BASE}/sync/batch`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify(payload)
        }).catch(()=>null);
        await quickSearchAndShow(data.code);
        toast('تم الحفظ ✅');
      } catch (e) { console.error(e); toast('فشل الحفظ','error'); }
    };
  }

  if (view === 'gen-issue') {
    const form = $('#gen-issue-form', mountRoot);
    if (!form) return;
    form.onsubmit = async (e) => {
      e.preventDefault();
      const data = {
        code:      $('#gi_code').value.trim(),
        issueDate: $('#gi_issueDate').value,
        receiver:  $('#gi_receiver').value.trim(),
        requester: $('#gi_requester').value.trim(),
        currSite:  $('#gi_currSite').value.trim(),
        notes:     $('#gi_notes').value.trim(),
      };
      if (!data.code) { toast('الكود مطلوب','error'); return; }
      try {
        await idbAdd('gen_issue', data);
        const payload = { items: [{ store:'gen_issue', payload: data }] };
        await fetchJSON(`${API_BASE}/sync/batch`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify(payload)
        }).catch(()=>null);
        await quickSearchAndShow(data.code);
        toast('تم الحفظ ✅');
      } catch (e) { console.error(e); toast('فشل الحفظ','error'); }
    };
  }

  if (view === 'gen-inspect') {
    const form = $('#gen-inspect-form', mountRoot);
    if (!form) return;
    form.onsubmit = async (e) => {
      e.preventDefault();
      const data = {
        code:       $('#gins_code').value.trim(),
        inspector:  $('#gins_inspector').value.trim(),
        elecRehab:  $('#gins_elecRehab').value.trim(),
        rehabDate:  $('#gins_rehabDate').value,
        rehabUp:    $('#gins_rehabUp').value.trim(),
        checkUp:    $('#gins_checkUp').value.trim(),
        notes:      $('#gins_notes').value.trim(),
      };
      if (!data.code) { toast('الكود مطلوب','error'); return; }
      try {
        await idbAdd('gen_inspect', data);
        const payload = { items: [{ store:'gen_inspect', payload: data }] };
        await fetchJSON(`${API_BASE}/sync/batch`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify(payload)
        }).catch(()=>null);
        await quickSearchAndShow(data.code);
        toast('تم الحفظ ✅');
      } catch (e) { console.error(e); toast('فشل الحفظ','error'); }
    };
  }
}

/* ---------- Top-level event delegation (bind once) ---------- */
(function bindGlobalsOnce(){
  // Sidebar nav (single listener)
  $('.sidebar')?.addEventListener('click', (e) => {
    const btn = e.target.closest('.nav-item');
    if (!btn) return;
    const v = btn.dataset.view;
    if (v) setActive(v);
  });

  // Search bar buttons
  $('#btnSearch')?.addEventListener('click', doSearch);
  $('#searchInput')?.addEventListener('keydown', (e)=>{ if (e.key === 'Enter') doSearch(); });

  // Quick Export button in toolbar
  $('#btnQuickExport')?.addEventListener('click', openExport);

  // Export modal buttons
  $('#closeExport')?.addEventListener('click', closeExport);
  $('#cancelExport')?.addEventListener('click', closeExport);
  $('#doExport')?.addEventListener('click', doExport);

  // Settings modal
  $('#btnSettings')?.addEventListener('click', () => {
    $('#settingsModal')?.classList.remove('hidden');
    $('#apiBaseInput').value = API_BASE;
  });
  $('#closeSettings')?.addEventListener('click', () => $('#settingsModal')?.classList.add('hidden'));
  $('#cancelSettings')?.addEventListener('click', () => $('#settingsModal')?.classList.add('hidden'));
  $('#saveApiBase')?.addEventListener('click', () => {
    const v = ($('#apiBaseInput')?.value || '').trim();
    if (!v) { toast('أدخل عنوان API صحيح','error'); return; }
    API_BASE = v; localStorage.setItem('api_base', v);
    $('#settingsModal')?.classList.add('hidden');
    toast('تم الحفظ ✅'); ping(); renderLast3();
  });
})();

/* ---------- Init ---------- */
(async function init(){
  await idbOpen();
  await ping();
  setActive('home');
})();

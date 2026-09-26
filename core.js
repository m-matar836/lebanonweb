// ===================================================================
//   core.js - الكود المشترك بين كل صفحات الموقع (V43: تقسيم script.js لملفات أصغر لكل صفحة)
// ===================================================================

// ===================================================================
//                     DARK MODE
// ===================================================================
function initDarkMode() {
    const toggle = document.getElementById('darkModeToggle');
    if (!toggle) return;
    const saved = localStorage.getItem('theme');
    if (saved === 'dark' || (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        document.documentElement.setAttribute('data-theme', 'dark');
    }
    toggle.addEventListener('click', () => {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        document.documentElement.setAttribute('data-theme', isDark ? 'light' : 'dark');
        localStorage.setItem('theme', isDark ? 'light' : 'dark');
    });
}

// ===================================================================
//                     PWA INSTALL PROMPT
// ===================================================================
let deferredInstallPrompt = null;
function initPwaInstall() {
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredInstallPrompt = e;
        if (sessionStorage.getItem('installBannerShown') || !document.querySelector('nav.navbar')) return;
        sessionStorage.setItem('installBannerShown', '1');
        const banner = document.createElement('div');
        banner.id = 'installBanner';
        banner.innerHTML = `<span class="install-icon"><i class="fa-solid fa-mobile-screen-button"></i></span>
            <div style="flex:1;min-width:0;"><div style="font-weight:700;font-size:.95rem;">ثبّت لوحة التحكم على جهازك</div>
            <div class="small text-muted">افتحها بنقرة واحدة حتى من دون إنترنت.</div></div>
            <button id="installConfirmBtn" class="btn btn-sm btn-primary">تثبيت</button>
            <button id="installCloseBtn" class="btn btn-sm btn-outline-secondary" aria-label="إغلاق"><i class="fa-solid fa-xmark"></i></button>`;
        document.body.appendChild(banner);
        banner.querySelector('#installCloseBtn').addEventListener('click', () => banner.remove());
        banner.querySelector('#installConfirmBtn').addEventListener('click', async () => {
            if (!deferredInstallPrompt) { banner.remove(); return; }
            deferredInstallPrompt.prompt();
            await deferredInstallPrompt.userChoice.catch(() => {});
            deferredInstallPrompt = null;
            banner.remove();
        });
    });
}

// ===================================================================
//                     SESSION LOG (دفتر الجلسات)
// ===================================================================
function recordLoginSession(userName) {
    try {
        const sessions = JSON.parse(localStorage.getItem('loginSessions') || '[]');
        const now = new Date();
        const device = { ua: navigator.userAgent, t: now.getTime(), name: userName || '', time: now.toLocaleString('ar-EG') };
        const currentDevice = navigator.userAgent.replace(/\d+/g, '#');
        const isNewDevice = !sessions.some(s => s.ua && s.ua.replace(/\d+/g, '#') === currentDevice);
        sessions.unshift(device);
        localStorage.setItem('loginSessions', JSON.stringify(sessions.slice(0, 6)));
        return isNewDevice && sessions.length > 1;
    } catch (e) { return false; }
}

// ===================================================================
//                     SESSION TIMEOUT
// ===================================================================
const SESSION_TIMEOUT_MS = 8 * 60 * 60 * 1000;
const SESSION_WARNING_MS = 10 * 60 * 1000;
let sessionTimer = null;
let sessionWarningTimer = null;

function startSessionTimeout() {
    const userRaw = localStorage.getItem('currentUser') || sessionStorage.getItem('currentUser');
    if (!userRaw) return;
    const loginTime = Number(localStorage.getItem('loginTimestamp') || sessionStorage.getItem('loginTimestamp'));
    if (!loginTime) {
        const now = Date.now();
        if (localStorage.getItem('currentUser')) localStorage.setItem('loginTimestamp', now);
        else sessionStorage.setItem('loginTimestamp', now);
        startSessionTimeout();
        return;
    }
    const elapsed = Date.now() - loginTime;
    const remaining = SESSION_TIMEOUT_MS - elapsed;
    if (remaining <= 0) { forceLogout('انتهت صلاحية الجلسة. يرجى تسجيل الدخول مرة أخرى.'); return; }
    const warningAt = remaining - SESSION_WARNING_MS;
    if (warningAt > 0) {
        sessionWarningTimer = setTimeout(() => showSessionWarning(SESSION_WARNING_MS), warningAt);
    } else if (remaining > 0) {
        showSessionWarning(remaining);
    }
    sessionTimer = setTimeout(() => forceLogout('انتهت صلاحية الجلسة.'), remaining);
}

function showSessionWarning(durationMs) {
    const existing = document.querySelector('.session-timeout-banner');
    if (existing) return;
    const minutes = Math.ceil(durationMs / 60000);
    const banner = document.createElement('div');
    banner.className = 'session-timeout-banner';
    banner.innerHTML = `<i class="fa-solid fa-clock"></i> ستنتهي جلستك خلال ${minutes} دقيقة. <button id="extendSessionBtn">تمديد الجلسة</button>`;
    document.body.appendChild(banner);
    document.getElementById('extendSessionBtn').addEventListener('click', () => {
        banner.remove();
        clearTimeout(sessionTimer);
        clearTimeout(sessionWarningTimer);
        const now = Date.now();
        if (localStorage.getItem('currentUser')) localStorage.setItem('loginTimestamp', now);
        else sessionStorage.setItem('loginTimestamp', now);
        startSessionTimeout();
    });
}

function forceLogout(message) {
    clearTimeout(sessionTimer);
    clearTimeout(sessionWarningTimer);
    localStorage.removeItem('currentUser');
    sessionStorage.removeItem('currentUser');
    localStorage.removeItem('loginTimestamp');
    sessionStorage.removeItem('loginTimestamp');
    localStorage.removeItem('appDB');
    localStorage.removeItem('dbCacheTimestamp');
    localStorage.removeItem(FORM_STATE_KEY);
    sessionStorage.removeItem(EDIT_STATE_KEY);
    invalidateSmartCaches();
    // إعادة ضبط حالة الشاشات مثل زر الخروج حتى لا تبقى بيانات مستخدم سابق.
    Object.keys(viewMounted).forEach(k => delete viewMounted[k]);
    currentRoute = null;
    reportsMemoryByScope = {};
    memoryReportsCache = null;
    alert(message);
    navigateTo('login');
}

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyOFcr3Iu9f1wRdbhxBxaWxtOzO-7EQfFqkHeqjbkWZP0lIhh0ywlMbs_s6X6un0V9D/exec";
const CACHE_DURATION_MINUTES = 1440;
const FORM_STATE_KEY = 'reportFormLastState'; 
const EDIT_STATE_KEY = 'reportToEdit';
// V25: in-memory caches eliminate repeated localStorage JSON parsing during the same page session.
let memoryDbCache = null;
let memoryReportsCache = null;
let dbBgRefreshInProgress = false;

// ===================================================================
//      AJAX LAYER (jQuery) + SHARED REPORTS CACHE — تحسين السرعة
// ===================================================================
const REPORTS_CACHE_TTL_MINUTES = 3;
const REPORTS_CACHE_PREFIX = 'reportsCache_v2::';
let reportsMemoryByScope = {};
const reportsInflightMap = {};

function reportsScopeKey(params = {}) {
    const role = String(params.role || '').toLowerCase().replace(/[^a-z0-9]/gi, '') || 'default';
    const target = String(params.targetUserId || params.userId || 'all');
    return `${role}__${target}`;
}

// GET عبر jQuery (مع احتياطي fetch).
function apiGet(action, params = {}) {
    const query = new URLSearchParams();
    query.set('action', action);
    Object.entries(params).forEach(([k, v]) => { if (v !== undefined && v !== null && v !== '') query.set(k, v); });
    query.set('_', String(Date.now()));
    const url = `${SCRIPT_URL}?${query.toString()}`;
    if (!window.jQuery) {
        return fetch(url, { cache: 'no-store' }).then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); });
    }
    return $.ajax({ url, method: 'GET', dataType: 'json', cache: false, timeout: 120000 })
        .then(d => d, x => { throw new Error((x && (x.statusText || x.responseText)) || 'خطأ في الاتصال'); });
}

// POST عبر jQuery (مع احتياطي fetch).
function apiPost(action, payload = {}) {
    if (!window.jQuery) {
        return fetch(SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action, payload }) }).then(r => r.json());
    }
    return $.ajax({
        url: SCRIPT_URL, method: 'POST', dataType: 'json',
        contentType: 'text/plain;charset=utf-8', timeout: 120000,
        data: JSON.stringify({ action, payload })
    }).then(d => d, x => { throw new Error((x && (x.statusText || x.responseText)) || 'خطأ في الاتصال'); });
}

// جلب getReports مع منع التكرار أثناء وجود طلب جارٍ (dedupe).
function ajaxGetReports(params = {}) {
    const scope = reportsScopeKey(params);
    if (reportsInflightMap[scope]) return reportsInflightMap[scope];
    const p = apiGet('getReports', params);
    reportsInflightMap[scope] = p;
    p.then(() => { if (reportsInflightMap[scope] === p) delete reportsInflightMap[scope]; }, () => { if (reportsInflightMap[scope] === p) delete reportsInflightMap[scope]; });
    return p;
}

function saveReportsCacheEntry(scope, data, params = {}) {
    reportsMemoryByScope[scope] = data;
    try {
        localStorage.setItem(REPORTS_CACHE_PREFIX + scope, JSON.stringify(data));
        localStorage.setItem(REPORTS_CACHE_PREFIX + scope + '_ts', String(Date.now()));
    } catch (e) { /* امتلاء التخزين */ }
    // توافق مع الأجزاء القديمة التي تقرأ memoryReportsCache / reportsCache.
    if (!params.targetUserId || params.targetUserId === 'all') {
        memoryReportsCache = data;
        try { localStorage.setItem('reportsCache', JSON.stringify(data)); } catch (e) {}
    }
}

function invalidateReportsCache() {
    reportsMemoryByScope = {};
    memoryReportsCache = null;
    try {
        Object.keys(localStorage).forEach(k => { if (k.startsWith(REPORTS_CACHE_PREFIX)) localStorage.removeItem(k); });
        localStorage.removeItem('reportsCache');
    } catch (e) {}
}

// جلب التقارير بأسلوب «اعرض القديم فوراً + حدّث بالخلفية» (stale-while-revalidate):
// - كاش طازج  => يعيده فوراً دون أي شبكة.
// - كاش قديم  => يعيده فوراً ويجلب الأحدث بالخلفية ثم يرسل reportsCacheUpdated.
// - لا كاش    => ينتظر الشبكة.
// - أوفلاين   => يعيد أي كاش محفوظ مهما كان قديماً.
async function cachedReportsFetch(params = {}, opts = {}) {
    const force = !!opts.force;
    const ttlMs = (opts.ttlMinutes || REPORTS_CACHE_TTL_MINUTES) * 60000;
    const scope = reportsScopeKey(params);
    const storeKey = REPORTS_CACHE_PREFIX + scope;
    const tsKey = storeKey + '_ts';

    let cached = null, ts = 0;
    try {
        if (!reportsMemoryByScope[scope]) {
            const raw = localStorage.getItem(storeKey);
            if (raw && raw !== 'undefined') reportsMemoryByScope[scope] = JSON.parse(raw);
        }
        cached = reportsMemoryByScope[scope] || null;
        ts = Number(localStorage.getItem(tsKey) || 0);
    } catch (e) { cached = null; }

    const isFresh = !!cached && !!ts && (Date.now() - ts) < ttlMs;

    if (force) {
        try {
            const data = await ajaxGetReports(params);
            saveReportsCacheEntry(scope, data, params);
            return data;
        } catch (e) {
            return cached || [];
        }
    }

    if (isFresh || (!navigator.onLine && cached)) return cached || [];

    if (!cached) {
        try {
            const data = await ajaxGetReports(params);
            saveReportsCacheEntry(scope, data, params);
            return data;
        } catch (e) {
            return [];
        }
    }

    // كاش قديم: نعيده فوراً ونحدّثه بالخلفية.
    if (!reportsInflightMap[scope]) {
        ajaxGetReports(params).then(data => {
            saveReportsCacheEntry(scope, data, params);
            window.dispatchEvent(new CustomEvent('reportsCacheUpdated', { detail: { scope, data } }));
        }).catch(err => console.warn('تحديث خلفي للتقارير فشل (يستمر بالكاش):', err));
    }
    return cached;
}

let originalCreatedAt = null; 

// ===================================================================
//                     OFFLINE-FIRST STORAGE
// ===================================================================
const OFFLINE_DB_NAME = 'festivalOfflineDB';
const OFFLINE_DB_VERSION = 4;
const OFFLINE_QUEUE_STORE = 'pendingReports';
const OFFLINE_ATTENDANCE_STORE = 'pendingAttendance';
const OFFLINE_MOVEMENT_STORE = 'pendingMovements';

function openOfflineDB() {
    return new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) return reject(new Error('IndexedDB غير مدعوم'));
        const req = indexedDB.open(OFFLINE_DB_NAME, OFFLINE_DB_VERSION);
        req.onupgradeneeded = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains(OFFLINE_QUEUE_STORE)) {
                db.createObjectStore(OFFLINE_QUEUE_STORE, { keyPath: 'localId' });
            }
            if (!db.objectStoreNames.contains(OFFLINE_ATTENDANCE_STORE)) {
                db.createObjectStore(OFFLINE_ATTENDANCE_STORE, { keyPath: 'localId' });
            }
            if (!db.objectStoreNames.contains(OFFLINE_MOVEMENT_STORE)) {
                db.createObjectStore(OFFLINE_MOVEMENT_STORE, { keyPath: 'localId' });
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function queueReportOffline(reportData) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_QUEUE_STORE, 'readwrite');
        tx.objectStore(OFFLINE_QUEUE_STORE).put({
            localId: `${reportData.id}_${Date.now()}`,
            reportData,
            createdAt: Date.now()
        });
        tx.oncomplete = () => { db.close(); resolve(); registerBackgroundSync(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}

async function getPendingReports() {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_QUEUE_STORE, 'readonly');
        const req = tx.objectStore(OFFLINE_QUEUE_STORE).getAll();
        req.onsuccess = () => { db.close(); resolve(req.result || []); };
        req.onerror = () => { db.close(); reject(req.error); };
    });
}

async function removePendingReport(localId) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_QUEUE_STORE, 'readwrite');
        tx.objectStore(OFFLINE_QUEUE_STORE).delete(localId);
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}

async function syncPendingReports() {
    if (!navigator.onLine) return;
    let pending = [];
    try { pending = await getPendingReports(); } catch (e) { return; }
    let syncedAny = false;
    for (const item of pending) {
        try {
            const res = await fetch(SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'submitReport', payload: item.reportData })
            });
            const result = await res.json();
            if (result.status !== 'success') throw new Error(result.message || 'فشل المزامنة');
            await removePendingReport(item.localId);
            syncedAny = true;
        } catch (error) {
            console.warn('Offline sync stopped:', error);
            break;
        }
    }
    if (syncedAny) {
        // Syncing reports does not require rebuilding master data. Invalidate only
        // the local report list so the next history view gets fresh server data.
        invalidateReportsCache();
        window.dispatchEvent(new CustomEvent('reportsCacheInvalidated'));
    }
    updateOfflineStatus();
}

// ---- Attendance offline queue (same pattern as reports, separate store) ----
async function queueAttendanceOffline(payload) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_ATTENDANCE_STORE, 'readwrite');
        tx.objectStore(OFFLINE_ATTENDANCE_STORE).put({
            localId: `attendance_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
            payload,
            createdAt: Date.now()
        });
        tx.oncomplete = () => { db.close(); resolve(); registerBackgroundSync(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}

async function getPendingAttendance() {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_ATTENDANCE_STORE, 'readonly');
        const req = tx.objectStore(OFFLINE_ATTENDANCE_STORE).getAll();
        req.onsuccess = () => { db.close(); resolve(req.result || []); };
        req.onerror = () => { db.close(); reject(req.error); };
    });
}

async function removePendingAttendance(localId) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_ATTENDANCE_STORE, 'readwrite');
        tx.objectStore(OFFLINE_ATTENDANCE_STORE).delete(localId);
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}

async function syncPendingAttendance() {
    if (!navigator.onLine) return;
    let pending = [];
    try { pending = await getPendingAttendance(); } catch (e) { return; }
    let syncedAny = false;
    for (const item of pending) {
        try {
            const res = await fetch(SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'submitAttendance', payload: item.payload })
            });
            const result = await res.json();
            if (result.status !== 'success') throw new Error(result.message || 'فشل المزامنة');
            await removePendingAttendance(item.localId);
            syncedAny = true;
        } catch (error) {
            console.warn('Attendance offline sync stopped:', error);
            break;
        }
    }
    if (syncedAny) {
        localStorage.removeItem('attendanceCache');
        window.dispatchEvent(new CustomEvent('attendanceCacheInvalidated'));
    }
    updateOfflineStatus();
}


// ---- Materials movement offline queue ----
async function queueMovementOffline(payload) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_MOVEMENT_STORE, 'readwrite');
        tx.objectStore(OFFLINE_MOVEMENT_STORE).put({
            localId: `movement_${Date.now()}_${Math.random().toString(36).slice(2,8)}`,
            payload, createdAt: Date.now()
        });
        tx.oncomplete = () => { db.close(); resolve(); registerBackgroundSync(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}
async function getPendingMovements() {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_MOVEMENT_STORE, 'readonly');
        const req = tx.objectStore(OFFLINE_MOVEMENT_STORE).getAll();
        req.onsuccess = () => { db.close(); resolve(req.result || []); };
        req.onerror = () => { db.close(); reject(req.error); };
    });
}
async function removePendingMovement(localId) {
    const db = await openOfflineDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(OFFLINE_MOVEMENT_STORE, 'readwrite');
        tx.objectStore(OFFLINE_MOVEMENT_STORE).delete(localId);
        tx.oncomplete = () => { db.close(); resolve(); };
        tx.onerror = () => { db.close(); reject(tx.error); };
    });
}
async function syncPendingMovements() {
    if (!navigator.onLine) return;
    let pending=[]; try { pending=await getPendingMovements(); } catch(e) { return; }
    let syncedAny=false;
    for (const item of pending) {
        try {
            const res=await fetch(SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'addFestivalMovement',payload:item.payload})});
            const result=await res.json();
            if(!result || result.status!=='success') throw new Error(result?.message||'فشل مزامنة الحركة');
            await removePendingMovement(item.localId); syncedAny=true;
        } catch(e) { console.warn('Movement offline sync stopped:',e); break; }
    }
    if(syncedAny) window.dispatchEvent(new CustomEvent('movementCacheInvalidated'));
    updateOfflineStatus();
}

async function updateOfflineStatus() {
    const el = document.getElementById('offline-status');
    if (!el) return;
    let pendingCount = 0;
    try {
        const [reportsPending, attendancePending, movementsPending] = await Promise.all([getPendingReports(), getPendingAttendance(), getPendingMovements()]);
        pendingCount = reportsPending.length + attendancePending.length + movementsPending.length;
    } catch (e) {}
    const badge = document.getElementById('syncPendingBadge');
    if (badge) {
        if (pendingCount) {
            badge.textContent = pendingCount;
            badge.title = `${pendingCount} عنصر بانتظار المزامنة — اضغط للمزامنة الآن`;
            badge.classList.remove('d-none');
        } else {
            badge.classList.add('d-none');
        }
    }
    if (!navigator.onLine) {
        el.textContent = pendingCount ? `🔴 بدون إنترنت — ${pendingCount} عنصر بانتظار المزامنة` : '🔴 بدون إنترنت — العمل محفوظ محلياً';
        el.style.display = 'block';
        el.style.background = '#dc3545';
        el.style.color = '#fff';
    } else if (pendingCount) {
        el.textContent = `🟠 متصل — ${pendingCount} عنصر بانتظار المزامنة`;
        el.style.display = 'block';
        el.style.background = '#ffc107';
        el.style.color = '#000';
    } else {
        el.textContent = '🟢 متصل';
        el.style.display = 'block';
        el.style.background = '#198754';
        el.style.color = '#fff';
        setTimeout(() => { if (navigator.onLine) el.style.display = 'none'; }, 2500);
    }
}

// ===================================================================
//  نافذة تفاصيل المزامنة المعلّقة (#6) — قائمة العناصر + مزامنة يدوية
// ===================================================================
const SYNC_TYPE_LABELS = { reports: 'تقرير مبيعات', attendance: 'تسجيل دوام', movements: 'حركة سحب / مرتجع' };

async function openSyncDetails() {
    const modalEl = document.getElementById('syncDetailsModal');
    if (!modalEl) return;
    if (!window._syncModal) {
        window._syncModal = new bootstrap.Modal(modalEl);
        document.getElementById('syncNowBtn')?.addEventListener('click', async () => {
            const btn = document.getElementById('syncNowBtn');
            const orig = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-1"></i> جارٍ المزامنة...';
            try { await runPendingSyncIfNeeded(); } catch (e) {}
            btn.disabled = false;
            btn.innerHTML = orig;
            renderSyncList();
        });
    }
    renderSyncList();
    window._syncModal.show();
}

async function renderSyncList() {
    const container = document.getElementById('syncListContainer');
    const statusEl = document.getElementById('syncListStatus');
    if (!container) return;
    let items = [];
    try {
        const [reports, attendance, movements] = await Promise.all([
            getPendingReports(), getPendingAttendance(), getPendingMovements()
        ]);
        items = [
            ...reports.map(it => ({ led: it, store: 'reports', label: SYNC_TYPE_LABELS.reports, summary: (it.reportData && (it.reportData.campaign || it.reportData.market)) ? [it.reportData.campaign, it.reportData.market].filter(Boolean).join(' - ') : 'تقرير' })),
            ...attendance.map(it => ({ led: it, store: 'attendance', label: SYNC_TYPE_LABELS.attendance, summary: (it.payload && it.payload.status) || 'دوام' })),
            ...movements.map(it => ({ led: it, store: 'movements', label: SYNC_TYPE_LABELS.movements, summary: 'حركة مواد' }))
        ];
    } catch (e) { items = []; }
    items.sort((a, b) => (a.led.createdAt || 0) - (b.led.createdAt || 0));
    if (statusEl) {
        statusEl.textContent = navigator.onLine
            ? (items.length ? `${items.length} عنصر بانتظار الإرسال — اضغط «مزامنة الآن».` : 'كل العناصر متزامنة.')
            : `${items.length} عنصر محفوظ محلياً — ستُرسل تلقائياً عند عودة الإنترنت.`;
    }
    if (!items.length) {
        container.innerHTML = '<div class="text-center text-muted py-4"><i class="fa-solid fa-circle-check fs-3 d-block mb-2 text-success"></i>لا توجد عناصر بانتظار المزامنة</div>';
        return;
    }
    container.innerHTML = items.map(it => {
        const when = new Date(it.led.createdAt || Date.now()).toLocaleString('ar');
        return `<div class="list-group-item d-flex justify-content-between align-items-center"><div><strong>${escapeHtmlGlobal(it.label)}</strong><div class="small text-muted">${escapeHtmlGlobal(it.summary)}<br>${escapeHtmlGlobal(String(when))}</div></div><button type="button" class="btn btn-sm btn-outline-danger" data-remove-ledger="${escapeHtmlGlobal(String(it.led.localId))}" data-store="${it.store}" title="تجاهل هذا العنصر"><i class="fa-solid fa-xmark"></i></button></div>`;
    }).join('');
    container.querySelectorAll('[data-remove-ledger]').forEach(btn => {
        btn.addEventListener('click', async () => {
            const id = btn.dataset.removeLedger;
            const store = btn.dataset.store;
            try {
                if (store === 'reports') await removePendingReport(id);
                else if (store === 'attendance') await removePendingAttendance(id);
                else await removePendingMovement(id);
            } catch (e) {}
            renderSyncList();
            updateOfflineStatus();
        });
    });
}

window.addEventListener('online', () => { updateOfflineStatus(); syncPendingReports(); syncPendingAttendance(); syncPendingMovements(); });
window.addEventListener('offline', updateOfflineStatus);
navigator.serviceWorker?.addEventListener?.('message', (event) => {
    if (event.data && event.data.type === 'SYNC_PENDING') {
        syncPendingReports(); syncPendingAttendance(); syncPendingMovements();
    }
});
// V42: تسجيل مزامنة الخلفية عند وجود عناصر معلقة (يفضّل على الانتظار للـ setInterval).
async function registerBackgroundSync() {
    if (!navigator.serviceWorker?.ready || !('SyncManager' in window)) return;
    try {
        const reg = await navigator.serviceWorker.ready;
        await reg.sync.register('sync-pending');
    } catch (e) { /* SyncManager غير مدعوم — يبقى setInterval هو الاحتياط */ }
}
let pendingSyncTimer = null;
async function runPendingSyncIfNeeded() {
    if (!navigator.onLine) return;
    try {
        const [reportsPending, attendancePending, movementsPending] = await Promise.all([getPendingReports(), getPendingAttendance(), getPendingMovements()]);
        if (reportsPending.length) await syncPendingReports();
        if (attendancePending.length) await syncPendingAttendance();
        if (movementsPending.length) await syncPendingMovements();
    } catch (e) {
        console.warn('Pending sync check skipped:', e);
    }
}
pendingSyncTimer = setInterval(runPendingSyncIfNeeded, 300000);
document.addEventListener('DOMContentLoaded', () => {
    setupCacheRefreshButtons();
    setTimeout(() => { updateOfflineStatus(); syncPendingReports(); syncPendingAttendance(); syncPendingMovements(); }, 500);
});

// ===================================================================
//                     SPA ROUTER (شل واحد + راوتر تجزئة)
// ===================================================================
const SPA_ROUTES = {
    login: 'view-login',
    reports: 'view-reports',
    history: 'view-history',
    movement: 'view-movement',
    attendance: 'view-attendance',
    dashboard: 'view-dashboard'
};
const SPA_TITLES = {
    login: 'تسجيل الدخول - لوحة التحكم',
    reports: 'إدخال التقارير - لوحة التحكم',
    history: 'سجل التعديلات - لوحة التحكم',
    movement: 'سحب / مرتجع مواد - لوحة التحكم',
    attendance: 'الدوام - لوحة التحكم',
    dashboard: 'التحليلات - لوحة التحكم'
};
const viewActivators = {};
const viewMounted = {};
let currentRoute = null;

// تسجيل دالة تهيئة كل شاشة (تُنفَّذ مرة واحدة عند أول فتح للشاشة).
function registerView(route, activator) { viewActivators[route] = activator; }

function getStoredUser() {
    try {
        return JSON.parse(localStorage.getItem('currentUser')) || JSON.parse(sessionStorage.getItem('currentUser')) || null;
    } catch (e) { return null; }
}

// V46: حساب role="user" ومنصبه "مروج" مقيَّد بصفحة الدوام حصراً.
function isPromoterAccount(u) {
    const r = String(u?.role || '').trim().toLowerCase();
    const jp = String(u?.jobPosition || '').trim();
    return r === 'user' && jp === 'مروج';
}

function routeFromHash() {
    const h = (location.hash || '').replace(/^#\/?/, '');
    const name = h.split('?')[0];
    return SPA_ROUTES[name] ? name : 'login';
}

function queryFromHash() {
    const i = (location.hash || '').indexOf('?');
    return i < 0 ? new URLSearchParams() : new URLSearchParams((location.hash || '').slice(i + 1));
}

function navigateTo(route) {
    if (String(route).startsWith('#')) { location.hash = route; return; }
    location.hash = '#/' + route;
}

// الشاشة الرئيسية المناسبة للجلسة الحالية (المروج = الدوام، البقية = التقارير).
function navigateHome() {
    navigateTo(isPromoterAccount(getStoredUser()) ? 'attendance' : 'reports');
}

function hideAppSplash() {
    const splash = document.getElementById('app-splash');
    if (splash) splash.style.display = 'none';
}

function runActivator(route) {
    const fn = viewActivators[route];
    if (!fn) return;
    if (viewMounted[route]) {
        // إعادة فتح الشاشة: إشعار للشاشة لتعيد التحميل من الكاش/الخادم دون إعادة تثبيت.
        window.dispatchEvent(new CustomEvent('spaViewRevisited', { detail: { route } }));
        return;
    }
    viewMounted[route] = true;
    Promise.resolve().then(() => fn());
}

let shellControlsBound = false;
function bindShellUserControls() {
    const user = getStoredUser();
    if (!user) return;
    const welcomeEl = document.getElementById('welcomeMessage');
    if (welcomeEl) welcomeEl.textContent = `أهلاً بك، ${user.name}`;
    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn && logoutBtn.dataset.logoutBound !== '1') {
        logoutBtn.dataset.logoutBound = '1';
        logoutBtn.addEventListener('click', () => {
            localStorage.removeItem('currentUser');
            sessionStorage.removeItem('currentUser');
            localStorage.removeItem('loginTimestamp');
            sessionStorage.removeItem('loginTimestamp');
            localStorage.removeItem('appDB');
            localStorage.removeItem('dbCacheTimestamp');
            localStorage.removeItem(FORM_STATE_KEY);
            sessionStorage.removeItem(EDIT_STATE_KEY);
            invalidateSmartCaches();
            // إعادة ضبط حالة الشاشات حتى تُبنى من جديد ببيانات المستخدم التالي عند الدخول.
            Object.keys(viewMounted).forEach(k => delete viewMounted[k]);
            currentRoute = null;
            reportsMemoryByScope = {};
            memoryReportsCache = null;
            location.hash = '#/login';
            activateRoute();
        });
    }
    // V42/V50: التحليلات للمشرف والمدير والمدقق فقط.
    const role = String(user.role || '').trim().toLowerCase();
    const showDashboard = ['admin', 'manager', 'auditor'].includes(role);
    document.querySelectorAll('.nav-link-dashboard').forEach(el => { const it = el.closest('.nav-item'); if (it) it.style.display = showDashboard ? '' : 'none'; });
    // V46: المروج يرى رابط الدوام فقط.
    const isPromoter = isPromoterAccount(user);
    document.querySelectorAll('.nav-link-reports, .nav-link-history, .nav-link-movement, .nav-link-dashboard').forEach(el => { const it = el.closest('.nav-item'); if (it) it.style.display = isPromoter ? 'none' : ''; });
    shellControlsBound = true;
}

function activateRoute() {
    let route = routeFromHash();
    const user = getStoredUser();

    // حراسة الصلاحيات: بلا جلسة → شاشة الدخول، جلسة على الدخول → الرئيسية، المروج → الدوام.
    if (!user && route !== 'login') route = 'login';
    else if (user && route === 'login') { navigateHome(); return; }
    else if (user && isPromoterAccount(user) && route !== 'attendance') { navigateTo('attendance'); return; }

    currentRoute = route;
    const isLogin = route === 'login';
    hideAppSplash();
    document.body.classList.toggle('login-body', isLogin);

    const loginView = document.getElementById('view-login');
    const appShell = document.getElementById('app-shell');
    if (loginView) loginView.classList.toggle('d-none', !isLogin);
    if (appShell) appShell.classList.toggle('d-none', isLogin);

    Object.keys(SPA_ROUTES).forEach(r => {
        const sec = document.getElementById(SPA_ROUTES[r]);
        if (sec && r !== 'login') sec.classList.toggle('d-none', r !== route);
    });

    document.querySelectorAll('.navbar-nav .nav-link').forEach(link => {
        link.classList.toggle('active', link.dataset.route === route);
    });

    if (document.title !== SPA_TITLES[route]) document.title = SPA_TITLES[route] || 'لوحة التحكم';

    if (isLogin) { runActivator('login'); return; }

    bindShellUserControls();

    // أوامر التنقل بين الشاشات: تعديل تقرير (من السجل/QR) وفتح تقرير في السجل بعد الحفظ.
    if (route === 'reports') {
        const editId = sessionStorage.getItem('spaEditReportId');
        if (editId) { sessionStorage.removeItem('spaEditReportId'); window.__spaPendingEditId = editId; }
    }
    if (route === 'history') {
        const scrollId = sessionStorage.getItem('spaScrollReportId');
        if (scrollId) { sessionStorage.removeItem('spaScrollReportId'); window.__spaPendingScrollId = scrollId; }
    }

    runActivator(route);
    window.scrollTo(0, 0);
    // إغلاق قائمة الأزرار على الجوال بعد التنقل.
    try {
        const nav = document.getElementById('mainNav');
        if (nav) bootstrap.Collapse.getOrCreateInstance(nav)?.hide();
    } catch (e) {}
}

document.addEventListener('DOMContentLoaded', () => {
    initDarkMode();
    initPwaInstall();
    startSessionTimeout();
    activateRoute();
    window.addEventListener('hashchange', activateRoute);
});

// ===================================================================
//                      2. جلب البيانات من Google Sheet
// ===================================================================
async function getDbData() {
    if (memoryDbCache) return memoryDbCache;
    const DB_KEY = APP_DB_KEY;
    const TS_KEY = APP_DB_TS_KEY;
    const cachedDB = localStorage.getItem(DB_KEY);
    const cacheTimestamp = localStorage.getItem(TS_KEY);

    if (cachedDB) {
        const ageMinutes = cacheTimestamp ? (Date.now() - Number(cacheTimestamp)) / 60000 : Infinity;
        // استخدم الكاش مباشرة إذا كان Offline، حتى لو انتهت مدته.
        if (!navigator.onLine || (cacheTimestamp && ageMinutes < CACHE_DURATION_MINUTES)) {
            memoryDbCache = JSON.parse(cachedDB);
            return memoryDbCache;
        }
        // الكاش قديم: نعرضه فوراً ولا نعلّق الشاشة، ونحدّثه بالخلفية (stale-while-revalidate).
        if (!dbBgRefreshInProgress) {
            dbBgRefreshInProgress = true;
            apiGet('getInitialData', { v: APP_DB_VERSION }).then(fresh => {
                if (!fresh || fresh.status === 'error') return;
                memoryDbCache = fresh;
                localStorage.setItem(DB_KEY, JSON.stringify(fresh));
                localStorage.setItem(TS_KEY, String(Date.now()));
                window.dispatchEvent(new CustomEvent('dbCacheRefreshed', { detail: fresh }));
            }).catch(err => console.warn('تحديث خلفي لبيانات الأساسيات فشل (يستمر بالكاش):', err))
              .finally(() => { dbBgRefreshInProgress = false; });
        }
        memoryDbCache = JSON.parse(cachedDB);
        return memoryDbCache;
    }

    try {
        const dbData = await apiGet('getInitialData', { v: APP_DB_VERSION });
        if (dbData.status === 'error') throw new Error(dbData.message || 'API error');
        memoryDbCache = dbData;
        localStorage.setItem(DB_KEY, JSON.stringify(dbData));
        localStorage.setItem(TS_KEY, Date.now());
        return dbData;
    } catch (error) {
        if (cachedDB) {
            console.warn('Using cached DB because network request failed:', error);
            memoryDbCache = JSON.parse(cachedDB);
            return memoryDbCache;
        }
        throw error;
    }
}

// ===================================================================
//                 CACHE REFRESH / FAST DATA UPDATE
// ===================================================================
const APP_DB_VERSION = 'v42-employees-merged';
const APP_DB_KEY = `appDB_${APP_DB_VERSION}`;
// V39: unified browser cache helpers. Data is served instantly from memory/local
// storage/Service Worker, then refreshed in the background when online.
const SMART_CACHE_PREFIX = 'festivalSmartCache::';
function invalidateSmartCaches() {
    memoryDbCache = null;
    memoryReportsCache = null;
    try {
        Object.keys(localStorage).forEach(k => {
            if (k.startsWith(SMART_CACHE_PREFIX) || k.startsWith('attendanceCache::') || k === 'attendanceStatusCache') localStorage.removeItem(k);
        });
    } catch (e) {}
    try { caches?.keys?.().then(keys => keys.filter(k => k.includes('festival-app-v4')).forEach(k => caches.delete(k))).catch(()=>{}); } catch(e) {}
}

const APP_DB_TS_KEY = `dbCacheTimestamp_${APP_DB_VERSION}`;

let cacheRefreshInProgress = false;

async function refreshAppCache({ silent = false } = {}) {
    if (cacheRefreshInProgress) return { ok: false, busy: true };
    if (!navigator.onLine) { if (!silent) alert('لا يمكن تحديث البيانات بدون اتصال بالإنترنت.'); return { ok:false, offline:true }; }
    cacheRefreshInProgress = true;
    const buttons=document.querySelectorAll('[data-refresh-cache]');
    buttons.forEach(btn=>{btn.disabled=true;btn.dataset.originalHtml=btn.innerHTML;btn.innerHTML='<i class="fa-solid fa-spinner fa-spin me-1"></i>جاري تحديث كل البيانات...';});
    try {
        // Tell the Service Worker to remove stale API/static entries first.
        if (navigator.serviceWorker?.controller) navigator.serviceWorker.controller.postMessage({type:'CLEAR_APP_CACHE'});
        const currentUser=JSON.parse(localStorage.getItem('currentUser')||sessionStorage.getItem('currentUser')||'null');
        const requests=[apiGet('getInitialData',{forceRefresh:1,v:APP_DB_VERSION})];
        if(currentUser){
            const u=String(currentUser.id||''), r=String(currentUser.role||''), n=String(currentUser.name||'');
            requests.push(apiGet('getReports',{userId:u,role:r,userName:n,targetUserId:'all'}));
            requests.push(apiGet('getTeamOptions',{userId:u,role:r,userName:n}));
            requests.push(apiGet('getUserFestivalMovements',{userId:u,role:r,targetUserId:u}));
            requests.push(apiGet('getAttendance',{userId:u,role:r,userName:n,targetUserId:u}));
        }
        requests.push(apiGet('getStatusOptions'));
        const settled = await Promise.allSettled(requests);
        const first = settled[0];
        if (!first || first.status !== 'fulfilled' || !first.value || first.value.status === 'error') {
            // حالات الفشل الفرعية (تقارير/دوام/حركات...) لا تُفشل التحديث الأساسي.
            throw new Error((first && first.reason && (first.reason.message || String(first.reason))) || (first && first.value && first.value.message) || 'فشل جلب البيانات الأساسية');
        }
        const freshDB = first.value;
        memoryDbCache=freshDB; localStorage.setItem(APP_DB_KEY,JSON.stringify(freshDB)); localStorage.setItem(APP_DB_TS_KEY,String(Date.now()));
        const results = settled.map(s => s.status === 'fulfilled' ? s.value : null);
        if(currentUser && Array.isArray(results[1])) saveReportsCacheEntry(reportsScopeKey({role:currentUser.role,targetUserId:'all'}), results[1], {targetUserId:'all'});
        window.dispatchEvent(new CustomEvent('dbCacheRefreshed',{detail:freshDB}));
        window.dispatchEvent(new CustomEvent('reportsCacheInvalidated'));
        window.dispatchEvent(new CustomEvent('movementCacheInvalidated'));
        window.dispatchEvent(new CustomEvent('attendanceCacheInvalidated'));
        // حدث موحّد: كل الشاشات المفتوحة تعيد تحميل بياناتها بعد انتهاء التحديث.
        window.dispatchEvent(new CustomEvent('appDataRefreshed'));
        if(navigator.serviceWorker?.getRegistrations) navigator.serviceWorker.getRegistrations().then(regs=>Promise.all(regs.map(reg=>reg.update()))).catch(()=>{});
        buttons.forEach(btn=>{btn.classList.remove('btn-outline-primary');btn.classList.add('btn-outline-success');btn.innerHTML='<i class="fa-solid fa-check me-1"></i>تم تحديث كل البيانات';});
        setTimeout(()=>buttons.forEach(btn=>{btn.classList.remove('btn-outline-success');btn.classList.add('btn-outline-primary');btn.innerHTML=btn.dataset.originalHtml||'<i class="fa-solid fa-arrows-rotate me-1"></i>تحديث البيانات';btn.disabled=false;}),1800);
        return {ok:true,data:freshDB};
    } catch(error){
        console.error('Full cache refresh failed:',error);
        buttons.forEach(btn=>btn.innerHTML='<i class="fa-solid fa-triangle-exclamation me-1"></i>فشل التحديث');
        setTimeout(()=>buttons.forEach(btn=>{btn.innerHTML=btn.dataset.originalHtml||'<i class="fa-solid fa-arrows-rotate me-1"></i>تحديث البيانات';btn.disabled=false;}),2000);
        if(!silent) alert(error.name==='AbortError'?'انتهت مهلة الاتصال. حاول مرة أخرى.':`تعذر تحديث كل البيانات: ${error.message||error}`);
        return {ok:false,error};
    } finally { cacheRefreshInProgress=false; }
}

function setupCacheRefreshButtons() {
    document.querySelectorAll('[data-refresh-cache]').forEach(btn => {
        if (btn.dataset.refreshBound === '1') return;
        btn.dataset.refreshBound = '1';
        btn.addEventListener('click', async (event) => {
            event.preventDefault();
            await refreshAppCache();
            // مزامنة يدوية: إرسال أي تقارير/دوام/حركات معلّقة أولاً ثم تحديث العدّاد.
            await runPendingSyncIfNeeded();
        });
    });
    // شارة المزامنة المعلّقة = فتح نافذة التفاصيل مع زر «مزامنة الآن».
    document.querySelectorAll('#syncPendingBadge').forEach(badge => {
        if (badge.dataset.bound === '1') return;
        badge.dataset.bound = '1';
        badge.addEventListener('click', () => {
            openSyncDetails();
        });
    });
}

// ===================================================================
//   نظام الصلاحيات على الواجهة: admin (الكل) / manager (فريقه) / user (نفسه)
// ===================================================================
// يجلب أسماء الموظفين الذين يحق لصاحب الجلسة الحالية عرض بياناتهم:
// admin => كل الموظفين، manager => فريقه فقط، user => قائمة فارغة (لا تُعرض القائمة أصلاً).
async function fetchTeamOptions(currentUser) {
    const role = String(currentUser?.role || '').trim().toLowerCase();
    if (role !== 'admin' && role !== 'manager') return [];
    try {
        const result = await apiGet('getTeamOptions', {
            userId: String(currentUser.id || ''),
            role
        });
        if (!result || result.status !== 'success' || !Array.isArray(result.options)) return [];
        return result.options;
    } catch (e) {
        console.warn('تعذر تحميل قائمة الموظفين:', e);
        return [];
    }
}

// ---- Shared utility functions used by more than one page (history + dashboard) ----
function escapeHtmlGlobal(value) {
    return String(value ?? '').replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
    }[char]));
}

function reportsToCSV(reports) {
    if (!Array.isArray(reports) || !reports.length) return 'لا توجد بيانات';
    const escapeCsv = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const headers = ['رقم التقرير','التاريخ','الحملة','الحدث','المحافظة','المنطقة','المحل','المشرف','المنسق','تبعية الجرد','عدد الأيام','الوقت من','الوقت إلى','هاتف','المبيعات','الكمية','عدد المبيعات','المصاريف','عدد المصاريف','ملاحظات','أنشئ بواسطة','تاريخ الإنشاء'];
    const lines = [headers.join(',')];
    reports.forEach(r => {
        let salesTotal = 0, salesQty = 0, salesCount = 0;
        (r.sales || []).forEach(s => {
            salesTotal += (Number(s.price) || 0) * (Number(s.quantity) || 0);
            salesQty += Number(s.quantity) || 0;
            salesCount++;
        });
        let expenseTotal = 0, expenseCount = 0;
        (r.expenses || []).forEach(e => { expenseTotal += Number(e.quantity) || 0; expenseCount++; });
        lines.push([
            r.id, r.date, r.campaign, r.event, r.governorate, r.region, r.market,
            r.supervisor, r.coordinator, r.inventoryDependency, r.eventDays, r.timeFrom, r.timeTo,
            r.phoneNumber, salesTotal.toFixed(2), salesQty, salesCount,
            expenseTotal, expenseCount, r.notes, r.createdByName, r.createdAt
        ].map(escapeCsv).join(','));
    });
    return lines.join('\r\n');
}

function downloadTextFile(filename, content, mimeType) {
    const blob = new Blob([content], { type: (mimeType || 'text/plain') + ';charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

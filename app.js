/* ================= TRX MoneyMapper =================
   وب‌اپ مدیریت تراکنش بانکی + سینک دوطرفه با Google Sheets
   داده‌ها: شیت Google (Or در حالت میهمان: localStorage)
===================================================== */

/* ---------- تنظیمات ---------- */
const CFG_KEY = 'trx_cfg';
const LOCAL_KEY = 'trx_local';

const BANKS = [
  'ملت','ملی','صادرات','پاسارگاد','تجارت','سپه','رفاه','مسکن','کشاورزی','بلو',
  'صنعت و معدن','پارسیان','کارآفرین','انصار','سامان','سینا','شهر','دی',
  'خاورمیانه','آینده','مهر ایران','قوامین','توسعه صادرات'
];

const DEFAULT_CATEGORIES = [
  'کافه / قهوه','سوپرمارکت','میوه و تره‌بار','رستوران','ساندویچ / فست‌فود',
  'بنزین','تاکسی / اسنپ','خرید آنلاین','لباس و پوشاک','دارو / داروخانه',
  'قبض (برق/آب/گاز/موبایل)','اجاره','نان و لبنیات','خواروبار','تفریح و سرگرمی',
  'سلامت و ورزش','هدیه','ترمیم عیب وسایل','وام و قسط','نقد و واریز دستی'
];
let cats = load('trx_cats', []);

const SHEET_NAME = 'تراکنش‌ها';
const HEADERS = ['شناسه','تاریخ','نوع','مبلغ','بابت','بانک','کامنت','منبع','آخرین تغییر'];
const PULL_INTERVAL_MS = 15000;

let cfg = load(CFG_KEY, { scriptUrl:'', sheetId:'', gasToken:'' });
let txs = [];
let editingId = null;
let currentType = 'out';
let syncing = false;
let lastRemoteHash = '';

/* ---------- کمکی‌ها ---------- */
function load(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d } catch { return d } }
function save(k, v) { localStorage.setItem(k, JSON.stringify(v)) }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8) }
function fa(n) { return Number(n).toLocaleString('fa-IR') }
function esc(s) { const d = document.createElement('div'); d.textContent = s ?? ''; return d.innerHTML }
function nowISO() { return new Date().toISOString() }

function mapIn(row) {
  if (!row || row.length < 1) return null;
  const [id, date, type, amount, title, bank, note, source, updated] = row;
  if (!id || id === HEADERS[0]) return null;
  // گوگل گاهی تاریخ رو Date می‌کنه (مثل 2026-10-26T20:30:00Z) → به تاریخ محلی تبدیل کن
  let d = String(date || '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) { /* همین فرمت خوبه */ }
  else { const t = Date.parse(d); d = isNaN(t) ? '' : new Date(t).toLocaleDateString('en-CA') }
  return { id: String(id), date: /^\d{4}-\d{2}-\d{2}/.test(d) ? d : '', type: type === 'واریز' || type === 'in' ? 'in' : 'out',
    amount: parseNum(amount), title: title || '', bank: bank || '', note: note || '',
    source: source || '', updated: updated || '' };
}
function parseNum(v) { // "۱٬۲۵۰٬۰۰۰" or "1,250,000"
  const fa2en = { '۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9' };
  let s = String(v).replace(/[۰-۹]/g, c => fa2en[c]);
  return parseInt(s.replace(/[^\d-]/g, ''), 10) || 0;
}
function txToRow(t) {
  return [t.id, t.date, t.type === 'in' ? 'واریز' : 'برداشت', t.amount, t.title, t.bank, t.note, t.source, t.updated];
}

/* ---------- Apps Script API (روش اتصال) ---------- */
const isGas = () => !!cfg.scriptUrl;
const connected = () => isGas();

// POST ساده به Web App — بدون توکن و بدون CORS preflight
async function gasRequest(action, payload = {}) {
  const res = await fetch(cfg.scriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ token: cfg.gasToken || '', action, ...payload }),
  });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  const j = await res.json();
  if (!j.ok) throw new Error(j.error || 'خطای Apps Script');
  return j;
}

function extractRows(items) {
  return items.map(r => mapIn(r.values || r)).filter(Boolean);
}

async function pullFromSheet(silent = false) {
  if (!connected() || syncing) return;
  syncing = true;
  try {
    const res = await gasRequest('list');
    rowCache = new Map(res.rows.map(r => [r.values[0], r.rowNum]));
    const remote = extractRows(res.rows);
    const hash = JSON.stringify(remote.map(t => t.id + '|' + t.updated));
    if (hash !== lastRemoteHash) {
      lastRemoteHash = hash;
      // ادغام دوطرفه: جدیدتر برنده‌ست
      const localMap = idx(txs);
      const remoteIdx = idx(remote);
      const merged = [];
      for (const r of remote) {
        const l = localMap.get(r.id);
        if (l && l.updated > r.updated) merged.push(l); else merged.push(r);
      }
      // تراکنش‌های محلی که تازه‌ست و هنوز در شیت نیستن
      for (const l of txs) if (!remoteIdx.has(l.id)) merged.push(l);
      txs = merged;
      saveAll();
      render();
    }
    setSync(true, silent ? 'سینک خودکار' : 'متصل به Google Sheets');
    refreshBanks();
  } catch (e) { setSync(false, e.message) }
  finally { syncing = false }
}

let rowCache = new Map();  // id -> شماره ردیف در شیت

async function pushTx(t) {
  if (!connected()) return saveAll();
  try {
    const row = txToRow(t);
    const cached = rowCache.get(t.id);
    const res = cached
      ? await gasRequest('update', { rowNum: cached, row })
      : await gasRequest('append', { row });
    rowCache.set(t.id, res.rowNum || rowCache.size + 2);
    setSync(true);
    document.getElementById('lastSync').textContent = 'آخرین سینک: ' + new Date().toLocaleTimeString('fa-IR');
  } catch (e) { setSync(false, 'خطا در ارسال: ' + e.message) }
}

async function deleteTx(id) {
  if (!connected()) { txs = txs.filter(t => t.id !== id); saveAll(); render(); return }
  try {
    const rowNum = rowCache.get(id);
    if (rowNum) await gasRequest('delete', { rowNum });
    rowCache.delete(id);
    txs = txs.filter(t => t.id !== id); saveAll(); render();
    setSync(true);
  } catch (e) { setSync(false, e.message) }
}

function idx(arr) { return new Map(arr.map(t => [t.id, t])) }

// فیلتر بازه زمانی: today / month / range (از/تا)
function inPeriod(t, period) {
  const d = t.date || '';
  if (!d) return true;
  if (period === 'today') return d === new Date().toISOString().slice(0, 10);
  if (period === 'month') return d.slice(0, 7) === new Date().toISOString().slice(0, 7);
  if (period === 'range') {
    const from = document.getElementById('filterFrom').value;
    const to = document.getElementById('filterTo').value;
    if (from && d < from) return false;
    if (to && d > to) return false;
    return true;
  }
  return true;
}

function saveAll() { if (!connected()) save(LOCAL_KEY, txs) }

/* ---------- وضعیت سینک UI ---------- */
function setSync(ok, text) {
  const dot = document.getElementById('syncDot'), t = document.getElementById('syncText');
  dot.className = dot.className.replace(' live','') + (ok ? ' live' : '');
  if (text) t.textContent = text;
}

/* ---------- تراکنش‌ها ---------- */
function addTx(t) {
  txs.unshift(t);
  saveAll();
  pushTx(t);
  render();
}

function render() {
  const q = document.getElementById('q').value.trim();
  const fb = document.getElementById('filterBank').value;
  const ft = document.getElementById('filterType').value;
  const period = document.getElementById('filterPeriod').value;
  const list = txs
    .filter(t => !q || (t.title + ' ' + t.note + ' ' + t.bank).includes(q))
    .filter(t => !fb || t.bank === fb)
    .filter(t => !ft || t.type === ft)
    .filter(t => inPeriod(t, period))
    .sort((a, b) => (b.date || '').localeCompare(a.date || ''));

  const inc = txs.filter(t => t.type === 'in').reduce((s, t) => s + t.amount, 0);
  const exp = txs.filter(t => t.type === 'out').reduce((s, t) => s + t.amount, 0);
  const today = new Date().toISOString().slice(0, 10);
  const todayTxs = txs.filter(t => t.date === today);
  document.getElementById('stBalance').textContent = fa(inc - exp);
  document.getElementById('stIncome').textContent = fa(inc);
  document.getElementById('stExpense').textContent = fa(exp);
  document.getElementById('stCount').textContent = fa(txs.length);
  const tOut = todayTxs.filter(t => t.type === 'out').reduce((s, t) => s + t.amount, 0);
  const tIn = todayTxs.filter(t => t.type === 'in').reduce((s, t) => s + t.amount, 0);
  document.getElementById('stToday').textContent = `−${fa(tOut)} / +${fa(tIn)} (${fa(todayTxs.length)})`;

  const el = document.getElementById('txList');
  if (!list.length) { el.innerHTML = '<div class="empty">هنوز تراکنشی نیست — از فرم کنار اضافه کن 💸</div>'; return }
  el.innerHTML = list.map(t => `
    <div class="tx" data-id="${t.id}">
      <div class="badge ${t.type}">${t.type === 'in' ? '⬇️' : '⬆️'}</div>
      <div class="main">
        <div class="title">${esc(t.title) || '<span style="color:var(--muted)">بی‌عنوان</span>'}</div>
        <div class="sub">
          ${t.bank ? `<span class="bank">${esc(t.bank)}</span>` : ''}
          <span>${esc(t.date)}</span>
          ${t.source ? ' · ' + esc(t.source) : ''}
        </div>
        ${t.note ? `<div class="comment">💬 ${esc(t.note)}</div>` : ''}
      </div>
      <div class="side">
        <div class="amount ${t.type}">${t.type === 'in' ? '+' : '−'} ${fa(t.amount)}</div>
        <div class="actions">
          <button class="btn btn-sm" data-edit="${t.id}">ویرایش</button>
        </div>
      </div>
    </div>`).join('');
}

function openEdit(t) {
  editingId = t.id;
  document.getElementById('eAmount').value = t.amount;
  document.getElementById('eTitle').value = t.title;
  document.getElementById('eType').value = t.type;
  document.getElementById('eBank').value = t.bank;
  document.getElementById('eDate').value = t.date;
  document.getElementById('eSource').value = t.source;
  document.getElementById('eNote').value = t.note;
  editDlg.showModal();
}

function fillBankSelect(sel, current = '') {
  const used = [...new Set(txs.map(t => t.bank).filter(x => x))];
  const opts = ['<option value="">— انتخاب بانک —</option>',
    ...[...new Set([...BANKS, ...used])].map(b => `<option ${b === current ? 'selected':''}>${b}</option>`)
  ];
  sel.innerHTML = opts.join('');
}

/* ---------- دسته‌بندی‌های «بابت» ---------- */
function allCats() { return [...new Set([...DEFAULT_CATEGORIES, ...cats])] }
function fillCatList() {
  document.getElementById('catList').innerHTML = allCats().map(c => `<option value="${esc(c)}">`).join('');
}
function addCat() {
  const name = (prompt('نام دسته جدید رو بنویس (مثلاً «کافی‌شاپ»):') || '').trim();
  if (!name) return;
  if (!allCats().includes(name)) { cats.unshift(name); save('trx_cats', cats) }
  fillCatList();
  document.getElementById('fTitle').value = name;
}

function refreshBanks() {
  const cur = document.getElementById('filterBank').value;
  const used = [...new Set(txs.map(t => t.bank).filter(x => x))];
  document.getElementById('filterBank').innerHTML =
    '<option value="">همه بانک‌ها</option>' + used.map(b => `<option ${b === cur ? 'selected':''}>${b}</option>`).join('');
  fillBankSelect(document.getElementById('fBank'));
  fillBankSelect(document.getElementById('eBank'), document.getElementById('eBank').value);
}

/* ---------- Event wiring ---------- */
document.getElementById('typeOut').onclick = () => setType('out');
document.getElementById('typeIn').onclick = () => setType('in');
function setType(v) {
  currentType = v;
  document.getElementById('typeOut').className = 'btn ' + (v === 'out' ? 'btn-primary' : '');
  document.getElementById('typeIn').className = 'btn ' + (v === 'in' ? 'btn-primary' : '');
}

// ثبت فرم
document.getElementById('addBtn').onclick = () => {
  const amount = parseNum(document.getElementById('fAmount').value);
  const title = document.getElementById('fTitle').value.trim();
  if (!amount) return alert('مبلغ رو وارد کن!');
  addTx({ id: uid(), date: document.getElementById('fDate').value || new Date().toISOString().slice(0, 10),
    type: currentType, amount, title, bank: document.getElementById('fBank').value,
    source: '', note: document.getElementById('fNote').value.trim(), updated: nowISO() });
  ['fAmount','fTitle','fNote'].forEach(id => document.getElementById(id).value = '');
};

// کلیک روی لیست → ویرایش
document.getElementById('txList').addEventListener('click', e => {
  const btn = e.target.closest('[data-edit]');
  if (btn) { openEdit(txs.find(t => t.id === btn.dataset.edit)); return }
  const row = e.target.closest('.tx');
  if (row) openEdit(txs.find(t => t.id === row.dataset.id));
});

// ذخیره ویرایش
document.getElementById('editSave').onclick = () => {
  const t = txs.find(x => x.id === editingId);
  if (!t) return;
  Object.assign(t, {
    amount: parseNum(document.getElementById('eAmount').value),
    title: document.getElementById('eTitle').value.trim(),
    type: document.getElementById('eType').value,
    bank: document.getElementById('eBank').value,
    date: document.getElementById('eDate').value,
    source: document.getElementById('eSource').value.trim(),
    note: document.getElementById('eNote').value.trim(),
    updated: nowISO(),
  });
  saveAll(); pushTx(t); render(); editDlg.close();
};

// حذف
document.getElementById('editDelete').onclick = async () => {
  if (!confirm('این تراکنش حذف بشه؟')) return;
  await deleteTx(editingId);
  editDlg.close();
};

// جستجو و فیلتر
['q','filterBank','filterType','filterPeriod','filterFrom','filterTo'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'q' ? 'input' : 'change', () => {
    const p = document.getElementById('filterPeriod').value;
    const show = p === 'range';
    document.getElementById('filterFrom').style.display = show ? '' : 'none';
    document.getElementById('filterTo').style.display = show ? '' : 'none';
    render();
  });
});

// تنظیمات
document.getElementById('cfgSheet').value = cfg.sheetId;
document.getElementById('cfgScript').value = cfg.scriptUrl;
document.getElementById('cfgGasToken').value = cfg.gasToken || '';
document.getElementById('saveCfg').onclick = () => {
  cfg.sheetId = document.getElementById('cfgSheet').value.trim().match(/[-\w]{25,}/)?.[0] || cfg.sheetId;
  cfg.scriptUrl = document.getElementById('cfgScript').value.trim();
  cfg.gasToken = document.getElementById('cfgGasToken').value.trim();
  save(CFG_KEY, cfg);
  location.reload();
};

// اتصال از طریق Apps Script
async function connectGas() {
  setSync(false, 'در حال اتصال به Google Sheets...');
  try {
    rowCache = new Map();
    await pullFromSheet();
    setInterval(() => pullFromSheet(true), PULL_INTERVAL_MS);
    document.getElementById('guestBanner').style.display = 'none';
    document.getElementById('disconnectBtn').style.display = '';
  } catch (e) {
    setSync(false, 'اتصال ناموفق: ' + e.message);
  }
}

// قطع اتصال (پاک کردن تنظیمات)
document.getElementById('disconnectBtn').onclick = () => {
  if (!confirm('اتصال قطع بشه؟ (تنظیمات پاک می‌شود — تراکنش‌ها می‌مانند)')) return;
  cfg.scriptUrl = ''; cfg.gasToken = '';
  save(CFG_KEY, cfg);
  location.reload();
};

/* ---------- Init ---------- */
txs = load(LOCAL_KEY, []);
document.getElementById('fDate').value = new Date().toISOString().slice(0, 10);
fillBankSelect(document.getElementById('fBank'));
fillCatList();
document.getElementById('addCatBtn').onclick = addCat;
document.getElementById('addCatBtn2').onclick = addCat;
refreshBanks();
render();
setType('out');
setSync(!!cfg.scriptUrl, isGas() ? 'در حال اتصال به Google Sheets...' : 'حالت میهمان (localStorage)');
document.getElementById('guestBanner').hidden = isGas();
if (isGas()) connectGas();

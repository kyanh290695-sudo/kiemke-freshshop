/* Kiểm kê Freshshop — giao diện và luồng kiểm.
   Luồng: Chuyến kiểm mới (tải dữ liệu · chọn cửa hàng · người kiểm · ảnh check-in)
          → Kiểm hàng tồn kho / Kiểm tài sản / Biên bản kiểm tra (PDF). */
(function () {
  'use strict';
  const APP_VERSION = '1.10.1';
  const MASTER_URL = 'data/MASTERR.xlsx';

  const REASONS = {
    stock: ['Hao hụt', 'Hư hỏng', 'Chưa nhập sổ', 'Nhầm mã', 'Khác'],
    asset: ['Không tìm thấy', 'Hư hỏng', 'Đã điều chuyển', 'Khác']
  };
  const MODES = {
    asset: { title: 'Kiểm tài sản', unitWord: 'tài sản', okLabel: 'OK' },
    stock: { title: 'Kiểm hàng tồn kho', unitWord: 'mã hàng', okLabel: 'Đúng' }
  };
  const DENOMS = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000, 500];
  const CASH_ITEM = 8; // tiêu chí "Tiền mặt tại cửa hàng"
  const ICON = {
    asset: '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20h8M12 16v4"/><path d="M7 8h4"/></svg>',
    stock: '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8l9-5 9 5v8l-9 5-9-5z"/><path d="M3 8l9 5 9-5M12 13v8"/></svg>',
    doc: '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></svg>',
    back: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 18l-6-6 6-6"/></svg>',
    chev: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M9 6l6 6-6 6"/></svg>',
    search: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
    ok: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    bad: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    file: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
    cam: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/></svg>'
  };
  const STATUS = { todo: ['Chưa kiểm', 'p-todo', 'var(--todo)'], pend: ['Đang nhập', 'p-pend', 'var(--warn)'], ok: ['Khớp', 'p-ok', 'var(--ok)'],
    short: ['Thiếu', 'p-short', 'var(--short)'], over: ['Thừa', 'p-over', 'var(--over)'] };

  /* ---------- helpers ---------- */
  const nf = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 3 });
  const vnd = new Intl.NumberFormat('vi-VN');
  const fmt = n => nf.format(n);
  const sgn = n => (n > 0 ? '+' : '') + fmt(n);
  const money = n => vnd.format(n || 0) + ' đ';
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const plain = s => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
  const hm = t => t ? new Date(t).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) : '—';
  const dt = t => t ? new Date(t).toLocaleString('vi-VN') : '';
  const r3 = n => Math.round(n * 1000) / 1000;
  const decStr = n => n == null ? '' : String(n).replace('.', ',');
  function num(v) { v = String(v).trim().replace(',', '.'); if (v === '') return null; const n = Number(v); return isFinite(n) && n >= 0 ? n : NaN; }
  // Cộng dồn khi đếm: "5+4+3" hoặc "5 4 3" → 12 (dấu phẩy là phần thập phân)
  function sumExpr(v) {
    const parts = String(v == null ? '' : v).trim().split(/[+\s]+/).filter(Boolean);
    if (!parts.length) return null;
    let t = 0;
    for (const p of parts) { const n = Number(p.replace(',', '.')); if (!isFinite(n) || n < 0) return NaN; t += n; }
    return r3(t);
  }
  const isExpr = v => String(v || '').trim().replace(/\+$/, '').split(/[+\s]+/).filter(Boolean).length > 1;
  const sumLine = (v, unit) => isExpr(v) ? '= ' + fmt(sumExpr(v)) + ' ' + esc(unit) : '';
  const intOf = v => { const d = String(v || '').replace(/\D/g, ''); return d ? Number(d) : null; };
  const hasCF = it => it.cf != null && it.cf !== 1;

  /* ---------- state ---------- */
  let master = { stores: [], updatedAt: null };
  let visit = null;     // { status: 'setup'|'active', checker, storeOp, stock, dvt, asset: {fileName, parsed}, checkin:{data,at}, startedAt }
  const dvtMap = () => visit && visit.dvt ? visit.dvt.parsed.map : null;
  const S = { asset: null, stock: null };
  let report = null;
  const ui = { view: 'home', mode: null, filter: 'all', q: '', edit: null };
  let fileTarget = null;

  const sess = () => S[ui.mode];
  const toCheck = it => !it.zero || it.added;           // mã tồn cuối = 0 không cần kiểm
  const checkList = s => s.items.filter(toCheck);
  const storeByOp = op => master.stores.find(m => m.op === op) || null;
  const visitStore = () => visit && storeByOp(visit.storeOp);
  const byQty = s => s && s.by === 'qty';

  // Sổ sách / đơn vị theo cách kiểm: tồn kho theo số lượng (ĐVT) hoặc trọng lượng (ĐVT quy đổi)
  function bk(it, s) { s = s || sess(); return s && s.by ? (byQty(s) ? it.bookQty : it.book) : it.book; }
  function un(it, s) { s = s || sess(); if (!(s && s.by)) return it.unit; return byQty(s) ? (it.dvt || 'ĐV') : (it.u2 || 'Kg'); }
  function diffOf(it, s) { if (it.res === 'ok') return 0; if (it.res !== 'bad' || it.actual == null) return null; return r3(it.actual - bk(it, s)); }
  function st(it, s) { if (it.res == null) return 'todo'; const d = diffOf(it, s); return d == null ? 'pend' : d === 0 ? 'ok' : d < 0 ? 'short' : 'over'; }
  function pillHTML(it) {
    const s = st(it), d = diffOf(it);
    const txt = (s === 'short' || s === 'over') ? STATUS[s][0] + ' ' + sgn(d) + ' ' + un(it) : STATUS[s][0];
    return '<span class="pill ' + STATUS[s][1] + '">' + txt + '</span>';
  }

  let saveTimer = null;
  function save(now) {
    clearTimeout(saveTimer);
    const run = () => { ['asset', 'stock'].forEach(m => KKDB.set('sess-' + m, S[m]).catch(() => toast('Không lưu được dữ liệu trên máy. Kiểm tra bộ nhớ trống.'))); };
    if (now) run(); else saveTimer = setTimeout(run, 250);
  }
  const saveVisit = () => KKDB.set('visit', visit).catch(() => toast('Không lưu được dữ liệu trên máy.'));
  const saveReport = () => KKDB.set('report', report).catch(() => toast('Không lưu được biên bản trên máy.'));
  function logIt(act, code, detail) { sess().log.push({ t: Date.now(), act, code, detail }); }
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') save(true); });
  window.addEventListener('pagehide', () => save(true));

  /* ---------- render ---------- */
  const app = document.getElementById('app');
  const fileInput = document.getElementById('file');
  function go(view, mode) { ui.view = view; if (mode) ui.mode = mode; ui.edit = null; render(); window.scrollTo(0, 0); }
  function render() { app.innerHTML = VIEWS[ui.view](); BIND[ui.view] && BIND[ui.view](); }
  const rerender = () => { const y = window.scrollY; render(); window.scrollTo(0, y); };

  function topBar(title, sub, backTo) {
    return '<header class="top"><div class="top-row">' +
      (backTo ? '<button class="back" data-go="' + backTo + '" aria-label="Quay lại">' + ICON.back + '</button>' : '<img class="mark" src="icons/icon-192.png" alt="">') +
      '<div class="ttl"><b>' + esc(title) + '</b><small>' + sub + '</small></div></div>' + (ui.view === 'count' ? countHead() : '') + '</header>';
  }

  /* ---------- màn hình chính ---------- */
  function home() { return visit && visit.status === 'active' ? hub() : setupVisit(); }

  function dvtInfo() {
    const d = visit && visit.dvt;
    return d ? esc(d.fileName) + ' · ' + d.parsed.count + ' mã · OPER ' + esc(d.parsed.op || '—') + (d.parsed.date ? ' · ngày ' + esc(d.parsed.date) : '') : 'chưa có';
  }

  function setupVisit() {
    if (!visit) visit = { status: 'setup', checker: lastChecker(), storeOp: '', stock: null, dvt: null, asset: null, checkin: null };
    const V = visit, store = visitStore();
    const stockP = V.stock && V.stock.parsed, assetP = V.asset && V.asset.parsed;
    const assetMine = assetP && store ? assetP.items.filter(i => i.pca === store.pca).length : 0;
    const stockStore = stockP ? storeByOp(stockP.key) : null;
    const stockBad = stockP && store && stockP.key !== store.op;
    const dvtP = V.dvt && V.dvt.parsed;
    const dvtStore = dvtP && dvtP.op ? storeByOp(dvtP.op) : null;
    const dvtBad = dvtP && store && dvtP.op && dvtP.op !== store.op;
    const pairBad = stockP && dvtP && dvtP.op && stockP.key !== dvtP.op;
    const slot = (key, title, desc, body, btn) => '<div class="slot"><div class="slot-ic">' + ICON.file + '</div><div class="slot-body"><b>' + title + '</b>' +
      '<span class="hint">' + desc + '</span>' + (body || '') + '</div><button class="btn sm" data-pick="' + key + '">' + btn + '</button></div>';
    const stockBody = stockP ? '<span class="ok-line">' + esc(V.stock.fileName) + ' · ' + stockP.items.filter(i => !i.zero).length + ' mã cần kiểm · OPER ' + esc(stockP.key) +
      (stockP.meta.cutoff ? ' · chốt ' + esc(stockP.meta.cutoff) : '') + ' · ' + stockP.items.filter(i => i.zero).length + ' mã tồn 0 (không cần kiểm)' + '</span>' : '';
    const assetBody = assetP ? '<span class="ok-line">' + esc(V.asset.fileName) + ' · ' + assetP.items.length + ' tài sản · ' + Object.keys(assetP.keys).length + ' cửa hàng' +
      (store ? ' · <b>' + assetMine + ' của ' + esc(store.op) + '</b>' : '') + '</span>' : '';
    const missing = [];
    if (!stockP && !assetP) missing.push('tải file tồn kho hoặc tài sản');
    if (stockP && !dvtP) missing.push('tải file DVT (bắt buộc khi kiểm tồn kho)');
    if (dvtP && !stockP) missing.push('tải file STOCK đi kèm file DVT');
    if (!store) missing.push('chọn cửa hàng');
    if (!V.checker.trim()) missing.push('nhập tên người kiểm');
    if (!V.checkin) missing.push('chụp ảnh check-in');
    const blocking = stockBad || dvtBad || pairBad || (assetP && store && !assetMine && !stockP);
    return topBar('Kiểm kê Freshshop', 'Chuyến kiểm mới') + '<div class="body">' +
      (!master.stores.length ? '<div class="banner"><b>Chưa có danh sách cửa hàng</b><span>Mở app khi có mạng để tải danh sách cửa hàng.</span></div>' : '') +
      '<div class="eyebrow">1. Dữ liệu kiểm</div><div class="slots">' +
      slot('stock', 'Tồn kho', 'File STOCK.csv của cửa hàng', stockBody, stockP ? 'Đổi' : 'Chọn file') +
      slot('dvt', 'ĐVT tồn kho', 'File DVT.csv của cửa hàng · bắt buộc khi kiểm tồn kho', dvtP ? '<span class="ok-line">' + dvtInfo() + '</span>' : '', dvtP ? 'Đổi' : 'Chọn file') +
      slot('asset', 'Tài sản (tất cả cửa hàng)', 'File ASSET.xlsx, app tự lọc theo cửa hàng', assetBody, assetP ? 'Đổi' : 'Chọn file') +
      '</div><span class="hint">iPhone: bấm Chọn file → Duyệt → OneDrive. Android: có thể mở file trong app OneDrive, bấm Chia sẻ → "Kiểm kê".</span>' +
      '<div class="eyebrow" style="margin-top:6px">2. Cửa hàng</div>' +
      (store ? '<div class="store-pick"><div style="min-width:0"><b>' + esc(store.op + ' · ' + store.name) + '</b><span class="hint">' + esc([store.prov, store.area, 'PCA ' + store.pca].filter(Boolean).join(' · ')) +
        '</span>' + (stockP && !stockBad ? '<span class="match">✓ Khớp với file STOCK</span>' : '') + (dvtP && dvtP.op && !dvtBad ? '<span class="match">✓ Khớp với file DVT</span>' : '') + '</div><button class="btn sm" id="storeClear">Đổi</button></div>'
        : '<div class="search">' + ICON.search + '<input class="inp" id="storeQ" type="search" placeholder="Tìm tên hoặc mã cửa hàng, ví dụ: long khanh, FS176" autocomplete="off"></div><div class="store-res" id="storeRes"></div>' +
          ((stockStore || dvtStore) ? '<button class="btn" data-store="' + esc((stockStore || dvtStore).op) + '">Dùng cửa hàng trong file: ' + esc((stockStore || dvtStore).op + ' · ' + (stockStore || dvtStore).name) + '</button>' : '')) +
      (stockBad ? '<div class="err">File tồn kho thuộc ' + esc(stockP.key) + (stockStore ? ' (' + esc(stockStore.name) + ')' : '') + ', không phải ' + esc(store.op) + '. Chọn lại cửa hàng hoặc đổi file tồn kho.</div>' : '') +
      (stockP && !stockStore ? '<div class="err">Không tìm thấy OPER ' + esc(stockP.key) + ' trong danh sách cửa hàng.</div>' : '') +
      (pairBad ? '<div class="err">File STOCK (' + esc(stockP.key) + ') và file DVT (' + esc(dvtP.op) + ') không cùng cửa hàng.</div>' : '') +
      (dvtBad && !pairBad ? '<div class="err">File DVT thuộc ' + esc(dvtP.op) + ', không phải ' + esc(store.op) + '. Chọn lại cửa hàng hoặc đổi file DVT.</div>' : '') +
      (assetP && store && !assetMine ? '<div class="banner">File tài sản không có tài sản nào của ' + esc(store.op) + ' (PCA ' + esc(store.pca) + ').' + (stockP ? ' Chuyến này sẽ chỉ kiểm tồn kho.' : '') + '</div>' : '') +
      '<div class="eyebrow" style="margin-top:6px">3. Người kiểm</div>' +
      '<input class="inp" id="checker" autocomplete="name" placeholder="Họ và tên" value="' + esc(V.checker) + '">' +
      '<div class="eyebrow" style="margin-top:6px">4. Ảnh check-in khi đến cửa hàng</div>' +
      (V.checkin ? '<img class="photo" src="' + V.checkin.data + '" alt="Ảnh check-in"><button class="btn" id="checkinBtn"' + (store ? '' : ' disabled') + '>' + ICON.cam + 'Chụp lại</button>'
        : '<button class="btn' + (store ? ' pri' : '') + '" id="checkinBtn"' + (store ? '' : ' disabled') + '>' + ICON.cam + 'Chụp ảnh check-in</button><span class="hint">' +
          (store ? 'Chụp trước cửa hàng. App in giờ chụp, tên cửa hàng và toạ độ (nếu cho phép) lên ảnh.' : 'Chọn cửa hàng trước để app in đúng tên lên ảnh.') + '</span>') +
      (missing.length ? '<p class="hint" style="margin:0;color:var(--warn)">Cần ' + missing.join(', ') + '.</p>' : '') +
      '<button class="btn pri block" id="startVisit"' + (missing.length || blocking ? ' disabled' : '') + '>Bắt đầu kiểm</button>' +
'<a class="guide-link" href="huong-dan.html">📖 Hướng dẫn sử dụng</a>' +
      '<div class="hint" style="text-align:center">Phiên bản ' + APP_VERSION + ' · ' + master.stores.length + ' cửa hàng</div></div>';
  }

  function hub() {
    const store = visitStore() || {};
    const card = (m) => {
      const s = S[m];
      let small;
      if (!s) small = m === 'stock' ? (visit.dvt ? 'Chưa có file STOCK · bấm để chọn file' : 'Chưa có file DVT · bấm để chọn file') : 'Không có tài sản của cửa hàng · bấm để chọn file';
      else {
        const L = checkList(s), c = L.filter(i => i.res).length;
        small = c + '/' + L.length + ' ' + MODES[m].unitWord + (s.endedAt ? ' · đã kết thúc' : '');
      }
      const pct = s ? checkList(s).filter(i => i.res).length / Math.max(1, checkList(s).length) * 100 : 0;
      return '<button class="mcard" data-open="' + m + '"><span class="ic">' + ICON[m] + '</span><span style="min-width:0"><b>' + MODES[m].title + '</b><small>' + small + '</small>' +
        (s ? '<span class="bar" style="margin-top:6px"><i style="width:' + pct + '%"></i></span>' : '') + '</span>' + ICON.chev + '</button>';
    };
    const R = report;
    const rpSmall = R && R.confirmedAt ? 'Đã xác nhận ' + hm(R.confirmedAt) + ' · bấm để tải / gửi PDF' : 'Kiểm tra cửa hàng, tiền mặt, hình ảnh, ký tên';
    return topBar('Kiểm kê Freshshop', esc((store.op || '') + ' · ' + (store.name || '')), null) + '<div class="body">' +
      '<div class="visit-card">' + (visit.checkin ? '<img src="' + visit.checkin.data + '" alt="Ảnh check-in">' : '') +
      '<div style="min-width:0"><b>' + esc(store.name || '') + '</b><span class="hint">' + esc(store.op || '') + ' · ' + esc(visit.checker) + '</span>' +
      '<span class="hint">Check-in ' + (visit.checkin ? hm(visit.checkin.at) : '—') + ' · ' + new Date(visit.startedAt).toLocaleDateString('vi-VN') + '</span></div></div>' +
      '<div class="menu">' + card('stock') + card('asset') +
      '<button class="mcard" data-report><span class="ic">' + ICON.doc + '</span><span><b>Biên bản kiểm tra</b><small>' + rpSmall + '</small></span>' + ICON.chev + '</button></div>' +
      '<div class="foot-info"><span>ĐVT tồn kho: ' + dvtInfo() + '</span><button class="btn sm" data-pick="dvt">' + (visit.dvt ? 'Đổi' : 'Chọn file') + '</button></div>' +
      '<button class="btn" id="endVisit">Kết thúc chuyến kiểm · sang cửa hàng khác</button>' +
      '<p class="hint" style="margin:0">Mỗi lần bấm hoặc nhập số, app lưu ngay trên điện thoại. Tắt app hoặc mất mạng không làm mất số đã kiểm.</p>' +
'<a class="guide-link" href="huong-dan.html">📖 Hướng dẫn sử dụng</a>' +
      '<div class="hint" style="text-align:center">Phiên bản ' + APP_VERSION + '</div></div>';
  }

  /* ---------- màn hình kiểm ---------- */
  function counts() { const c = { all: 0, todo: 0, pend: 0, ok: 0, short: 0, over: 0 }; checkList(sess()).forEach(i => { c.all++; c[st(i)]++; }); return c; }
  const chipCount = (c, k) => k === 'todo' ? c.todo + c.pend : c[k];
  function countHead() {
    const c = counts(), s = sess();
    const chip = (k, label, color) => '<button class="chip" data-f="' + k + '" aria-pressed="' + (ui.filter === k) + '">' +
      (color ? '<span class="dot" style="background:' + color + '"></span>' : '') + label + ' <b data-c="' + k + '">' + chipCount(c, k) + '</b></button>';
    return '<div style="margin-top:12px" class="bar"><i id="prog" style="width:' + ((c.all - c.todo) / Math.max(1, c.all) * 100) + '%"></i></div>' +
      '<div class="stats">' + chip('all', 'Tất cả') + chip('todo', 'Chưa kiểm', 'var(--todo)') + chip('short', 'Thiếu', 'var(--short)') + chip('over', 'Thừa', 'var(--over)') + chip('ok', 'Khớp', 'var(--ok)') + '</div>' +
      '<div class="search">' + ICON.search + '<input class="inp" id="q" type="search" placeholder="Tìm mã hoặc tên" value="' + esc(ui.q) + '" autocomplete="off"></div>';
  }

  // Hệ số quy đổi viết dễ hiểu: "1 Khay = 0,3 Kg"
  const rate = it => '1 ' + esc(it.dvt || 'ĐV') + ' = ' + fmt(it.cf) + ' ' + esc(it.u2 || 'Kg');
  function bookLine(it) {
    if (ui.mode === 'asset') return '<span>Sổ sách <b>' + fmt(it.book) + ' ' + esc(it.unit) + '</b></span><span>' + esc(it.sub || '') + '</span>';
    const q = '<b>' + fmt(it.bookQty) + ' ' + esc(it.dvt || 'ĐV') + '</b>', w = '<b>' + fmt(it.book) + ' ' + esc(it.u2 || 'Kg') + '</b>';
    return byQty(sess())
      ? '<span>Sổ sách ' + q + (hasCF(it) ? ' · ' + w : '') + '</span>' + (hasCF(it) ? '<span>' + rate(it) + '</span>' : '')
      : '<span>Sổ sách ' + w + '</span>' + (hasCF(it) ? '<span>SL ' + q + '</span><span>' + rate(it) + '</span>' : '') +
        (!it.dvtKnown ? '<span class="unk">ĐVT chưa có</span>' : '');
  }

  function calcHTML(it) {
    const d = diffOf(it), u = un(it);
    if (d == null) return '<dl class="calc" style="background:var(--sunk)"><dt>Chênh lệch</dt><dd>Nhập số thực tế để tính</dd></dl>';
    const s = st(it), bg = s === 'ok' ? 'var(--ok-bg)' : s === 'short' ? 'var(--short-bg)' : 'var(--over-bg)', col = STATUS[s][2];
    let h = '<dl class="calc" style="background:' + bg + '"><dt>Sổ sách</dt><dd>' + fmt(bk(it)) + ' ' + esc(u) + '</dd><dt>Thực tế</dt><dd>' + fmt(it.actual) + ' ' + esc(u) + '</dd>' +
      '<dt style="color:' + col + ';font-weight:600">' + STATUS[s][0] + '</dt><dd class="big" style="color:' + col + '">' + sgn(d) + ' ' + esc(u) + '</dd>';
    if (ui.mode === 'stock' && hasCF(it) && d !== 0) {
      h += byQty(sess()) ? '<dt>Quy đổi</dt><dd style="color:' + col + '">≈ ' + sgn(r3(d * it.cf)) + ' ' + esc(it.u2 || 'Kg') + '</dd>'
        : '<dt>Quy đổi</dt><dd style="color:' + col + '">≈ ' + sgn(r3(d / it.cf)) + ' ' + esc(it.dvt || 'ĐV') + '</dd>';
    }
    return h + '</dl>';
  }

  function panelHTML(it, i) {
    const rs = REASONS[ui.mode].map(r => '<button class="rc" data-reason="' + esc(r) + '" aria-pressed="' + (it.reason === r) + '">' + esc(r) + '</button>').join('');
    const field = (attr, id, label, val, unit, hint, expr) => '<div class="field"><label for="' + id + '">' + label + '</label>' +
      '<div class="unit-row"><div class="unit"><input class="inp" id="' + id + '" ' + attr + ' data-unit="' + esc(unit) + '"' +
      ' inputmode="decimal" enterkeyhint="' + hint + '" autocomplete="off" value="' + esc(expr || decStr(val)) + '"><span>' + esc(unit) + '</span></div>' +
      '<button type="button" class="plus" data-plus aria-label="Cộng thêm">+</button></div>' +
      '<span class="sum-line" data-sumline>' + sumLine(expr, unit) + '</span></div>';
    let inputs;
    if (ui.mode === 'asset') inputs = '<div class="pair one">' + field('data-act', 'act-' + i, 'Số lượng thực tế', it.actual, it.unit, 'done', it.expr) + '</div>';
    else if (byQty(sess())) inputs = '<div class="pair one">' + field('data-act', 'act-' + i, 'Số lượng thực tế', it.actual, it.dvt || 'ĐV', 'done', it.expr) + '</div>' +
      '<span class="hint" style="margin-top:-6px">Đếm nhiều chỗ thì bấm <b>+</b> giữa các số, ví dụ 5 + 4 + 3, app tự cộng.</span>';
    else inputs = '<div class="pair' + (hasCF(it) ? '' : ' one') + '">' +
      (hasCF(it) ? field('data-aq', 'aq-' + i, 'Số lượng đếm', it.actQty, it.dvt || 'ĐV', 'next') : '') +
      field('data-act', 'act-' + i, 'Trọng lượng thực tế', it.actual, it.u2 || 'Kg', 'done') + '</div>' +
      (hasCF(it) ? '<span class="hint" style="margin-top:-6px">Nhập số lượng thì app tự quy đổi (' + rate(it) + '). Có cân thì sửa ô trọng lượng.</span>' : '');
    return '<div class="panel">' + inputs + '<div data-calc>' + calcHTML(it) + '</div>' +
      '<div class="field"><span class="lbl">Lý do</span><div class="reasons">' + rs + '</div></div>' +
      '<div class="field"><label for="note-' + i + '">Ghi chú</label><textarea class="inp" id="note-' + i + '" data-note rows="2" placeholder="' + (ui.mode === 'asset' ? 'Ví dụ: máy để ở kho sau, màn hình nứt' : 'Ví dụ: 2 gói rách bao bì, đã tách riêng') + '">' + esc(it.note) + '</textarea></div>' +
      '<button class="btn pri block" data-done>Xong</button></div>';
  }

  function rowHTML(it, i) {
    const s = st(it), ed = ui.edit === i, okL = MODES[ui.mode].okLabel;
    let recap = '';
    if (!ed && it.res === 'bad') recap = '<div class="recap">' + (it.actual != null ? '<span>Thực tế <b>' + fmt(it.actual) + ' ' + esc(un(it)) + '</b>' + (it.expr ? ' <span class="why">(' + esc(it.expr.replace(/\s+/g, '').replace(/\+$/, '').split('+').join(' + ')) + ')</span>' : '') + '</span>' : '') +
      (it.reason ? '<span class="why">' + esc(it.reason) + '</span>' : '<span class="why" style="color:var(--warn)">Chưa chọn lý do</span>') +
      (it.note ? '<span class="why">· ' + esc(it.note) + '</span>' : '') + '<button class="link" data-editrow>Sửa</button></div>';
    if (!ed && it.res === 'ok') recap = '<div class="recap"><span class="why">' + okL + ' lúc ' + hm(it.ts) + '</span></div>';
    return '<div class="row st-' + s + (ed ? ' editing' : '') + '" data-i="' + i + '"><span class="stripe"></span><div style="min-width:0">' +
      '<div class="r-head"><div><div class="code mono">' + esc(it.code) + (it.added ? ' <span class="added">Ngoài DS</span>' : '') + '</div>' +
      '<div class="name">' + esc(it.name) + '</div><div class="meta">' + bookLine(it) + '</div></div><span data-pill>' + pillHTML(it) + '</span></div>' +
      '<div class="act"><button class="ab ok" data-ok aria-pressed="' + (it.res === 'ok') + '">' + ICON.ok + okL + '</button>' +
      '<button class="ab bad" data-bad aria-pressed="' + (it.res === 'bad') + '">' + ICON.bad + 'Sai</button></div>' +
      recap + (ed ? panelHTML(it, i) : '') + '</div></div>';
  }

  function visible() {
    const q = plain(ui.q.trim());
    return sess().items.map((it, i) => [it, i]).filter(([it, i]) => toCheck(it) &&
      (ui.filter === 'all' || st(it) === ui.filter || (ui.filter === 'todo' && st(it) === 'pend') || ui.edit === i) &&
      (!q || plain(it.code + ' ' + it.name).includes(q)));
  }
  function listHTML() {
    const v = visible();
    if (v.length) return v.map(([it, i]) => rowHTML(it, i)).join('');
    if (ui.filter === 'todo' && !ui.q.trim()) return '<div class="empty"><b>Đã kiểm hết các mã.</b><br>Bấm "Tất cả" hoặc Thiếu / Thừa / Khớp để xem lại, hoặc bấm Kết thúc kiểm.</div>';
    return '<div class="empty">Không có mã nào khớp bộ lọc.<br>Thử xoá từ khoá hoặc chọn "Tất cả".</div>';
  }
  function count() {
    const s = sess();
    return topBar(MODES[ui.mode].title + ' · ' + s.store.op, esc(s.checker) + ' · bắt đầu ' + hm(s.startedAt) + ' · <span id="el"></span>', 'home') +
      '<div class="list" id="list">' + listHTML() + '</div>' +
      '<div class="foot"><button class="btn" id="add">+ Thêm mã</button><button class="btn pri" id="finish">Kết thúc kiểm</button></div>';
  }
  function refreshCounts() {
    const c = counts();
    Object.keys(c).forEach(k => { const el = app.querySelector('[data-c="' + k + '"]'); if (el) el.textContent = chipCount(c, k); });
    const p = document.getElementById('prog'); if (p) p.style.width = ((c.all - c.todo) / Math.max(1, c.all) * 100) + '%';
  }
  // Đang lọc "Chưa kiểm": mã vừa kiểm xong tự ẩn sau một nhịp ngắn
  function hideIfChecked(i) {
    if (ui.filter !== 'todo' || ui.edit === i) return;
    const it = sess().items[i];
    if (st(it) === 'todo' || st(it) === 'pend') return;
    const el = rowEl(i); if (!el) return;
    el.classList.add('leaving');
    setTimeout(() => {
      const cur = sess().items[i];
      if (ui.filter !== 'todo' || ui.edit === i || st(cur) === 'todo' || st(cur) === 'pend') return;
      const e2 = rowEl(i); if (e2) e2.remove();
      if (!document.querySelector('#list .row')) redrawList();
    }, 450);
  }
  const rowEl = i => app.querySelector('.row[data-i="' + i + '"]');
  function redrawRow(i) { const el = rowEl(i); if (!el) return; const t = document.createElement('div'); t.innerHTML = rowHTML(sess().items[i], i); el.replaceWith(t.firstChild); refreshCounts(); }
  function softRefresh(i) {
    const it = sess().items[i], el = rowEl(i); if (!el) return;
    el.className = 'row st-' + st(it) + (ui.edit === i ? ' editing' : '');
    el.querySelector('[data-pill]').innerHTML = pillHTML(it);
    const c = el.querySelector('[data-calc]'); if (c) c.innerHTML = calcHTML(it);
    refreshCounts();
  }
  function redrawList() { document.getElementById('list').innerHTML = listHTML(); }
  function nextTodo(after) {
    const v = visible(), k = v.findIndex(([, i]) => i === after);
    for (let j = k + 1; j < v.length; j++) if (st(v[j][0]) === 'todo') return v[j][1];
    return null;
  }

  function done() {
    const s = sess(), c = counts(), diffs = s.items.filter(i => i.res === 'bad' && (st(i) === 'short' || st(i) === 'over'));
    const dur = Math.round(((s.endedAt || Date.now()) - s.startedAt) / 60000);
    const box = k => '<div style="background:' + ({ ok: 'var(--ok-bg)', short: 'var(--short-bg)', over: 'var(--over-bg)', todo: 'var(--todo-bg)' })[k] + ';color:' + STATUS[k][2] + '"><b>' + chipCount(c, k) + '</b><span>' + STATUS[k][0] + '</span></div>';
    const list = diffs.length ? '<div class="dl">' + diffs.map(i => {
      const d = diffOf(i), k = st(i);
      return '<div><div style="display:flex;justify-content:space-between;gap:8px"><b style="min-width:0">' + esc(i.name) + '</b><span class="pill ' + STATUS[k][1] + '">' + sgn(d) + ' ' + esc(un(i)) + '</span></div>' +
        '<span class="hint"><span class="mono">' + esc(i.code) + '</span> · sổ ' + fmt(bk(i)) + ' → thực tế ' + fmt(i.actual) + ' ' + esc(un(i)) + '</span>' +
        '<span style="font-size:13px">' + esc(i.reason || 'Chưa chọn lý do') + (i.note ? ' · ' + esc(i.note) : '') + '</span></div>';
    }).join('') + '</div>' : '<div class="empty" style="padding:16px">Không có chênh lệch.</div>';
    return topBar('Kết quả · ' + s.store.op, MODES[ui.mode].title + ' · ' + esc(s.store.name), s.endedAt ? 'home' : 'count') + '<div class="body">' +
      '<div class="sum">' + box('ok') + box('short') + box('over') + box('todo') + '</div>' +
      '<dl class="kv"><dt>Người kiểm</dt><dd>' + esc(s.checker) + '</dd><dt>Cửa hàng</dt><dd>' + esc(s.store.op + ' · ' + s.store.name) + '</dd>' +
      '<dt>File</dt><dd>' + esc(s.fileName) + '</dd>' + 
      '<dt>Thời gian</dt><dd>' + dt(s.startedAt) + ' → ' + (s.endedAt ? hm(s.endedAt) : 'đang kiểm') + ' (' + dur + ' phút)</dd></dl>' +
      '<div class="eyebrow">Các mã có chênh lệch (' + diffs.length + ')</div>' + list +
      (s.endedAt
        ? '<button class="btn pri block" data-report>Lập biên bản PDF</button><div class="btns"><button class="btn" id="xlsx">Tải Excel</button><button class="btn" id="reopen">Mở lại để kiểm tiếp</button></div>'
        : '<button class="btn pri block" id="end">Kết thúc phiên kiểm</button><button class="btn" data-go="count">Quay lại kiểm tiếp</button>') +
      '</div>';
  }

  /* ---------- biên bản kiểm tra ---------- */
  function reportCtx() {
    const store = visitStore();
    if (!visit || visit.status !== 'active' || !store) return null;
    const same = ['stock', 'asset'].map(m => S[m]).filter(Boolean);
    const ends = same.map(s => s.endedAt).filter(Boolean);
    return { store, checker: visit.checker, stock: S.stock, asset: S.asset,
      start: visit.checkin ? visit.checkin.at : visit.startedAt, end: ends.length && ends.length === same.length ? Math.max(...ends) : null, date: Date.now() };
  }
  function newReport(op) {
    return { op, photo: visit && visit.checkin ? visit.checkin.data : null, photoAt: visit && visit.checkin ? visit.checkin.at : null,
      checklist: KKReport.CHECKLIST.map(() => ({ v: null, note: '' })), other: '', photos: [],
      cash: { revenue: null, transfer: null, notes: {} }, manager: '', sigManager: null, sigChecker: null, confirmedAt: null };
  }
  function ensureReport(ctx) {
    if (!report || report.op !== ctx.store.op) report = newReport(ctx.store.op);
    if (!report.photos) report.photos = [];
    if (!report.cash) report.cash = { revenue: null, transfer: null, notes: {} };
    if (report.notes) delete report.notes; // đã bỏ mục ghi chú báo cáo: không in ghi chú cũ vào PDF
    return report;
  }
  function unconfirm() { report.confirmedAt = null; pdfCache = null; }

  function cashCalc(c) {
    const cash = DENOMS.reduce((a, d) => a + d * (c.notes[d] || 0), 0);
    const total = cash + (c.transfer || 0);
    const has = c.revenue != null;
    return { cash, total, diff: has ? total - c.revenue : null };
  }
  function cashSummaryHTML(c) {
    const k = cashCalc(c);
    const d = k.diff, col = d == null ? 'var(--muted)' : d === 0 ? 'var(--ok)' : d < 0 ? 'var(--short)' : 'var(--over)';
    const lbl = d == null ? 'Nhập doanh thu để so sánh' : d === 0 ? 'Khớp' : d < 0 ? 'Thiếu' : 'Thừa';
    return '<dl class="calc" style="background:var(--sunk)"><dt>Tiền mặt đếm được</dt><dd>' + money(k.cash) + '</dd>' +
      '<dt>+ Đã chuyển khoản / nộp</dt><dd>' + money(c.transfer) + '</dd><dt><b>Tổng cộng</b></dt><dd>' + money(k.total) + '</dd>' +
      '<dt>Doanh thu cửa hàng</dt><dd>' + (c.revenue == null ? '—' : money(c.revenue)) + '</dd>' +
      '<dt style="color:' + col + ';font-weight:600">' + lbl + '</dt><dd class="big" style="color:' + col + '">' + (d == null ? '' : (d > 0 ? '+' : '') + money(d)) + '</dd></dl>';
  }
  function cashHTML(c) {
    const moneyInp = (id, val, label) => '<div class="field"><label for="' + id + '">' + label + '</label><div class="unit"><input class="inp" id="' + id + '" data-money="' + id + '" inputmode="numeric" autocomplete="off" value="' + (val == null ? '' : vnd.format(val)) + '"><span>đ</span></div></div>';
    return '<div class="cash"><b class="lbl">Tính tiền mặt</b>' + moneyInp('cRev', c.revenue, 'Doanh thu cửa hàng') + moneyInp('cTr', c.transfer, 'Tiền đã chuyển khoản / nộp công ty') +
      '<div class="denoms"><span class="hint">Mệnh giá</span><span class="hint">Số tờ</span><span class="hint" style="text-align:right">Thành tiền</span>' +
      DENOMS.map(d => '<span class="dn">' + vnd.format(d) + '</span><input class="inp" data-den="' + d + '" inputmode="numeric" autocomplete="off" placeholder="0" value="' + (c.notes[d] || '') + '">' +
        '<span class="dn-sum" data-dsum="' + d + '">' + (c.notes[d] ? vnd.format(d * c.notes[d]) : '') + '</span>').join('') + '</div>' +
      '<div id="cashSum">' + cashSummaryHTML(c) + '</div><span class="hint">Khi có doanh thu, app tự chọn Đạt (khớp) hoặc Không đạt (lệch). Bạn vẫn đổi được.</span></div>';
  }

  function reportView() {
    const ctx = reportCtx();
    if (!ctx) return topBar('Biên bản kiểm tra', '', 'home') + '<div class="body"><div class="banner">Hãy bắt đầu chuyến kiểm (tải dữ liệu, chọn cửa hàng, chụp check-in) trước khi lập biên bản.</div></div>';
    const R = ensureReport(ctx);
    const part = (m, s) => {
      if (!s) return '<div class="rp-part"><b>' + MODES[m].title + '</b><span class="hint">Không có dữ liệu, biên bản sẽ bỏ phần này.</span></div>';
      const L = checkList(s), c = L.filter(i => i.res).length;
      return '<div class="rp-part"><b>' + MODES[m].title + ' ✓</b><span class="hint">' + esc(s.fileName) + ' · đã kiểm ' + c + '/' + L.length + ' · ' + hm(s.startedAt) + ' → ' +
        (s.endedAt ? hm(s.endedAt) : '<span style="color:var(--warn)">chưa kết thúc</span>') + '</span></div>';
    };
    const answered = R.checklist.filter(c => c.v).length;
    const cl = KKReport.CHECKLIST.map(([label, hint], i) => {
      const c = R.checklist[i];
      return '<div class="cl-row" data-ci="' + i + '"><div style="min-width:0"><b>' + (i + 1) + '. ' + esc(label) + '</b><div class="hint">' + esc(hint) + '</div></div>' +
        (i === CASH_ITEM ? cashHTML(R.cash) : '') +
        '<div class="act"><button class="ab ok" data-cv="dat" aria-pressed="' + (c.v === 'dat') + '">' + ICON.ok + 'Đạt</button><button class="ab bad" data-cv="kd" aria-pressed="' + (c.v === 'kd') + '">' + ICON.bad + 'Không đạt</button></div>' +
        '<input class="inp cl-note" data-cn placeholder="' + (c.v === 'kd' ? 'Ghi chú lý do không đạt' : 'Ghi chú (nếu có)') + '" value="' + esc(c.note) + '" autocomplete="off"></div>';
    }).join('');
    const sig = (key, label, name) => '<div class="field"><span class="lbl">' + label + '</span><button class="sigbox" data-sign="' + key + '">' +
      (R[key] ? '<img src="' + R[key] + '" alt="Chữ ký ' + esc(name) + '">' : '<span>Bấm để ký</span>') + '</button><span class="hint">' + esc(name || '…') + (R[key] ? ' · bấm vào để ký lại' : '') + '</span></div>';
    const missing = [];
    if (answered < 10) missing.push('đánh giá đủ 10 tiêu chí (' + answered + '/10)');
    if (!R.manager.trim()) missing.push('nhập tên quản lý cửa hàng');
    if (!R.sigManager) missing.push('quản lý cửa hàng ký');
    if (!R.sigChecker) missing.push('người kiểm ký');
    return topBar('Biên bản kiểm tra', esc(ctx.store.op + ' · ' + ctx.store.name), 'home') + '<div class="body">' +
      '<div class="file"><div class="eyebrow">Nội dung biên bản</div>' + part('stock', ctx.stock) + part('asset', ctx.asset) +
      '<span class="hint">Người kiểm: <b>' + esc(ctx.checker) + '</b> · Ngày ' + new Date().toLocaleDateString('vi-VN') + '</span></div>' +
      '<div class="eyebrow">1. Ảnh check-in (trang Nội dung chính)</div>' +
      (R.photo ? '<img class="photo" src="' + R.photo + '" alt="Ảnh check-in"><button class="btn" id="takePhoto">Chụp lại</button>'
        : '<button class="btn pri" id="takePhoto">Chụp ảnh check-in</button>') +
      '<div class="eyebrow" style="margin-top:6px">2. Kiểm tra cửa hàng (' + answered + '/10)</div><div class="cl">' + cl + '</div>' +
      '<div class="field"><label for="other">Các vấn đề khác</label><textarea class="inp" id="other" rows="3" placeholder="Ví dụ: *Tủ đông số 2*: kêu to, đã báo bảo trì">' + esc(R.other) + '</textarea>' +
      '<span class="hint">Đặt chữ trong hai dấu * để in đậm, ví dụ *Tủ đông số 2*.</span><div class="bold-prev" id="otherPrev">' + boldPreview(R.other) + '</div></div>' +
      '<div class="eyebrow" style="margin-top:6px">3. Hình ảnh kiểm tra (' + R.photos.length + ' ảnh)</div>' +
      (R.photos.length ? '<div class="ev">' + R.photos.map((p, i) => '<div class="ev-card" data-pi="' + i + '"><img src="' + p.data + '" alt="Ảnh kiểm tra ' + (i + 1) + '">' +
        '<div class="ev-body"><div class="ev-top"><b>Ảnh ' + (i + 1) + '</b><span class="hint">' + hm(p.at) + '</span><button class="link" data-pdel>Xoá</button></div>' +
        '<textarea class="inp" data-pnote rows="3" placeholder="Lý do / mô tả, ví dụ: Bình chữa cháy hết hạn 06/2026">' + esc(p.note) + '</textarea></div></div>').join('') + '</div>' : '') +
      '<div class="btns"><button class="btn" id="evCam">' + ICON.cam + 'Chụp ảnh</button><button class="btn" id="evLib">Chọn từ thư viện</button></div>' +
      '<span class="hint">Mỗi ảnh là một trang trong PDF, kèm lý do bạn ghi. Không có ảnh thì PDF bỏ phần này.</span>' +
      '<div class="eyebrow" style="margin-top:6px">4. Ký xác nhận</div>' +
      '<div class="field"><label for="manager">Tên quản lý cửa hàng *</label><input class="inp" id="manager" autocomplete="off" placeholder="Họ và tên" value="' + esc(R.manager) + '"></div>' +
      sig('sigManager', 'Quản lý cửa hàng ký *', R.manager) + sig('sigChecker', 'Người kiểm (Kế toán) ký *', ctx.checker) +
      (missing.length ? '<p class="hint" style="margin:0;color:var(--warn)">Cần ' + missing.join(', ') + '.</p>' : '') +
      '<button class="btn pri block" id="confirm"' + (missing.length ? ' disabled' : '') + '>' + (R.confirmedAt ? 'Tạo lại PDF' : 'Xác nhận & tạo PDF') + '</button>' +
      (R.confirmedAt ? '<p class="hint" style="margin:0">Đã xác nhận lúc ' + dt(R.confirmedAt) + '. Giờ kết thúc trên biên bản = lúc tạo PDF; tạo lại sẽ cập nhật giờ. Mỗi lần tạo ra 2 file:</p>' +
        '<div class="pdf-list"><div><b>Biên Bản Kiểm Tra (Có đính kèm hình ảnh)</b><span class="hint">Bìa, nội dung chính, biên bản, hình ảnh · khổ ngang</span><button class="link" data-pdf="0">Tải</button></div>' +
        '<div><b>Biên Bản Kiểm Tra</b><span class="hint">Kiểm tra cửa hàng (A4 dọc) · Tài sản cố định (A4 dọc) · Hàng tồn kho (A4 ngang)</span><button class="link" data-pdf="1">Tải</button></div></div>' +
        '<div class="btns"><button class="btn" id="pdfSave">Tải cả 2 PDF</button><button class="btn pri" id="pdfShare">Gửi cả 2 PDF</button></div>' : '') +
      '</div>';
  }

  const boldPreview = t => /\*[^*\n]+\*/.test(t || '') ? esc(t).replace(/\*([^*\n]+)\*/g, '<b>$1</b>').replace(/\n/g, '<br>') : '';

  // Ảnh: thu nhỏ + in giờ, cửa hàng (và toạ độ nếu có) lên ảnh
  let lastPos = null;
  function getPos() {
    if (lastPos && Date.now() - lastPos.t < 120000) return Promise.resolve(lastPos.c);
    return new Promise(res => {
      if (!navigator.geolocation) return res(null);
      navigator.geolocation.getCurrentPosition(p => { lastPos = { t: Date.now(), c: p.coords }; res(p.coords); }, () => res(null), { enableHighAccuracy: true, timeout: 6000, maximumAge: 60000 });
    });
  }
  async function stampPhoto(file, store) {
    const posP = getPos();
    const url = URL.createObjectURL(file);
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('Không đọc được ảnh')); i.src = url; });
    const max = 1600, k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas'); c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
    const g = c.getContext('2d'); g.drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
    const pos = await posP, now = new Date();
    const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][now.getMonth()];
    const lines = [String(now.getDate()).padStart(2, '0') + ' ' + mon + ' ' + now.getFullYear() + ' at ' + now.toLocaleTimeString('en-GB'), store.name, store.op + (store.prov ? ' · ' + store.prov : '')];
    if (pos) lines.push(pos.latitude.toFixed(5) + ', ' + pos.longitude.toFixed(5));
    const fs = Math.round(c.width * 0.036);
    g.font = '600 ' + fs + 'px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif'; g.textAlign = 'right'; g.textBaseline = 'top';
    g.shadowColor = 'rgba(0,0,0,.75)'; g.shadowBlur = fs * 0.25; g.fillStyle = '#fff';
    lines.forEach((t, i) => g.fillText(t, c.width - fs * 0.6, fs * 0.6 + i * fs * 1.25));
    return { data: c.toDataURL('image/jpeg', 0.82), at: now.getTime() };
  }

  // Ô ký tên: vẽ bằng ngón tay / bút, cắt sát nét ký rồi lưu PNG nền trong suốt
  function openSignPad(title, onSave) {
    sheetOpen('<h3>' + esc(title) + '</h3><div class="pad"><canvas id="pad"></canvas><span class="pad-line"></span></div>' +
      '<p class="hint" style="margin:0">Ký trong khung. Xoay ngang điện thoại để ký rộng hơn.</p>' +
      '<div class="btns" style="grid-template-columns:1fr 1fr 1fr"><button class="btn" id="padClear">Xoá</button><button class="btn" data-close>Huỷ</button><button class="btn pri" id="padOk">Lưu chữ ký</button></div>', sc => {
      const cv = sc.querySelector('#pad'), dpr = window.devicePixelRatio || 1;
      const r0 = cv.getBoundingClientRect(); cv.width = r0.width * dpr; cv.height = r0.height * dpr;
      const g = cv.getContext('2d'); let drawing = false, last = null, inked = false, box = [1e9, 1e9, -1e9, -1e9];
      g.lineCap = 'round'; g.lineJoin = 'round'; g.strokeStyle = '#111'; g.lineWidth = 2.6 * dpr;
      const pt = e => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * dpr, (e.clientY - r.top) * dpr]; };
      const grow = ([x, y]) => { box = [Math.min(box[0], x), Math.min(box[1], y), Math.max(box[2], x), Math.max(box[3], y)]; };
      cv.addEventListener('pointerdown', e => { e.preventDefault(); cv.setPointerCapture(e.pointerId); drawing = true; last = pt(e); grow(last); g.beginPath(); g.arc(last[0], last[1], g.lineWidth / 2, 0, 7); g.fillStyle = '#111'; g.fill(); inked = true; });
      cv.addEventListener('pointermove', e => { if (!drawing) return; e.preventDefault(); const p = pt(e); g.beginPath(); g.moveTo(last[0], last[1]); g.lineTo(p[0], p[1]); g.stroke(); last = p; grow(p); });
      ['pointerup', 'pointercancel', 'pointerleave'].forEach(t => cv.addEventListener(t, () => { drawing = false; }));
      sc.querySelector('#padClear').onclick = () => { g.clearRect(0, 0, cv.width, cv.height); inked = false; box = [1e9, 1e9, -1e9, -1e9]; };
      sc.querySelector('#padOk').onclick = () => {
        if (!inked) return toast('Chưa có chữ ký. Ký vào khung rồi bấm Lưu.');
        const m = 6 * dpr, x = Math.max(0, box[0] - m), y = Math.max(0, box[1] - m), w = Math.min(cv.width, box[2] + m) - x, h = Math.min(cv.height, box[3] + m) - y;
        const out = document.createElement('canvas'); out.width = w; out.height = h; out.getContext('2d').drawImage(cv, x, y, w, h, 0, 0, w, h);
        onSave(out.toDataURL('image/png')); sc.remove();
      };
    });
  }

  let pdfCache = null;
  // Tạo 2 file: báo cáo kiểm kê (theme) + biên bản gửi Giám đốc chi nhánh (A4)
  async function makePdf() {
    const ctx = reportCtx(); ctx.report = report;
    const now = Date.now();
    ctx.start = (visit.checkin && visit.checkin.at) || visit.startedAt; ctx.end = now; ctx.date = now;
    const [b1, b2] = await Promise.all([KKReport.build(ctx), KKReport.buildForms(ctx)]);
    const d = new Date(), stamp = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
    pdfCache = [
      new File([b1], 'BienBanKiemTra_CoHinhAnh_' + ctx.store.op + '_' + stamp + '.pdf', { type: 'application/pdf' }),
      new File([b2], 'BienBanKiemTra_' + ctx.store.op + '_' + stamp + '.pdf', { type: 'application/pdf' })
    ];
    return pdfCache;
  }
  async function shareFiles(files, text) {
    if (navigator.canShare && navigator.canShare({ files })) {
      try { await navigator.share({ files, title: files.map(f => f.name).join(', '), text }); }
      catch (e) { if (e.name !== 'AbortError') { files.forEach(download); toast('Không mở được bảng chia sẻ. Đã tải file về máy.'); } }
    } else { files.forEach((f, i) => setTimeout(() => download(f), i * 700)); toast('Máy này không hỗ trợ chia sẻ file. Đã tải file về máy.'); }
  }
  async function shareFile(f, text) {
    if (navigator.canShare && navigator.canShare({ files: [f] })) {
      try { await navigator.share({ files: [f], title: f.name, text }); } catch (e) { if (e.name !== 'AbortError') { download(f); toast('Không mở được bảng chia sẻ. Đã tải file về máy.'); } }
    } else { download(f); toast('Máy này không hỗ trợ chia sẻ file. Đã tải file về máy.'); }
  }

  const VIEWS = { home, count, done, report: reportView };

  /* ---------- overlays ---------- */
  function toast(msg) {
    document.querySelectorAll('.toast').forEach(t => t.remove());
    const t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg;
    document.body.appendChild(t); setTimeout(() => t.remove(), 3200);
  }
  function sheetOpen(html, onBind) {
    const sc = document.createElement('div'); sc.className = 'scrim';
    sc.innerHTML = '<div class="sheet" role="dialog" aria-modal="true">' + html + '</div>';
    sc.addEventListener('click', e => { if (e.target === sc || e.target.closest('[data-close]')) sc.remove(); });
    document.body.appendChild(sc); onBind && onBind(sc);
    const f = sc.querySelector('input'); f && f.focus();
  }
  function confirmSheet(title, body, okText, onOk) {
    sheetOpen('<h3>' + esc(title) + '</h3><p style="margin:0">' + body + '</p><div class="btns"><button class="btn" data-close>Huỷ</button><button class="btn pri" id="cs-ok">' + esc(okText) + '</button></div>',
      sc => sc.querySelector('#cs-ok').onclick = () => { sc.remove(); onOk(); });
  }

  /* ---------- dữ liệu: file + phiên ---------- */
  const lastChecker = () => { try { return localStorage.getItem('kk-checker') || ''; } catch (e) { return ''; } };

  // Gắn ĐVT từ danh mục sản phẩm vào mã tồn kho
  function enrich(it) {
    const map = dvtMap(), p = map && map[it.code];
    it.dvtKnown = !!(p && p.dvt);
    it.dvt = p && p.dvt ? p.dvt : (it.cf == null || it.cf === 1 ? 'Kg' : 'ĐV');
    it.u2 = p && p.dvt2 ? p.dvt2 : 'Kg';
    if (p && p.cf && (it.cf == null)) it.cf = p.cf;
    return it;
  }
  const blankRes = () => ({ res: null, actual: null, actQty: null, reason: '', note: '', ts: null });
  function makeSession(mode, fileName, items, t) {
    const store = visitStore();
    return { checker: visit.checker, store, fileName, startedAt: t, endedAt: null, by: mode === 'stock' ? 'qty' : undefined,
      items: items.map(i => Object.assign(blankRes(), mode === 'stock' ? enrich(Object.assign({}, i)) : i)),
      log: [{ t, act: 'Bắt đầu', code: '', detail: 'Mở phiên kiểm ' + store.op + ' từ ' + fileName + ' (' + items.length + ' dòng)' }] };
  }
  function assetItemsFor(store) { return visit.asset ? visit.asset.parsed.items.filter(i => i.pca === store.pca) : []; }

  function pickFile(target) { fileTarget = target; fileInput.value = ''; fileInput.click(); }
  fileInput.addEventListener('change', () => { const f = fileInput.files && fileInput.files[0]; if (f) handleFile(f, fileTarget); });

  async function handleFile(file, target) {
    let parsed;
    try { parsed = KKParse.parseFile(await file.arrayBuffer(), file.name, XLSX); }
    catch (err) { return toast('Không đọc được ' + file.name + '. ' + err.message); }
    if (parsed.kind === 'master') return toast('Danh sách cửa hàng đã nằm sẵn trong app, không cần chọn file master.');
    const label = { stock: 'STOCK', dvt: 'DVT', asset: 'tài sản' };
    if (target && target !== parsed.kind && target !== 'any') toast('Đây là file ' + label[parsed.kind] + ', app đã đưa vào đúng mục.');
    if (parsed.kind === 'dvt' ? !parsed.count : !parsed.items.length) return toast('File ' + file.name + ' không có dòng dữ liệu nào.');
    if (!visit) visit = { status: 'setup', checker: lastChecker(), storeOp: '', stock: null, dvt: null, asset: null, checkin: null };
    const kind = parsed.kind;

    if (kind === 'dvt') {
      if (visit.status === 'active') {
        const store = visitStore();
        if (parsed.op && parsed.op !== store.op) return toast('File DVT này thuộc ' + parsed.op + ', không phải ' + store.op + '.');
        visit.dvt = { fileName: file.name, parsed }; saveVisit();
        if (S.stock) { S.stock.items.forEach(enrich); save(true); }
        toast('Đã nạp ĐVT của ' + parsed.count + ' mã.'); return rerender();
      }
      visit.dvt = { fileName: file.name, parsed };
      if (!visit.storeOp && parsed.op && storeByOp(parsed.op)) visit.storeOp = parsed.op;
      saveVisit(); ui.view = 'home'; return rerender();
    }

    if (visit.status === 'active') {
      // Đang trong chuyến kiểm: thêm/đổi dữ liệu cho cửa hàng hiện tại
      const store = visitStore();
      if (kind === 'stock' && parsed.key !== store.op) return toast('File tồn kho này thuộc ' + parsed.key + ', không phải ' + store.op + '.');
      if (kind === 'stock' && !visit.dvt) return toast('Chọn file DVT của ' + store.op + ' trước (bấm "Chọn file" ở dòng ĐVT tồn kho).');
      const items = kind === 'stock' ? parsed.items : parsed.items.filter(i => i.pca === store.pca);
      if (!items.length) return toast('File tài sản không có tài sản nào của ' + store.op + ' (PCA ' + store.pca + ').');
      const apply = () => {
        visit[kind] = { fileName: file.name, parsed }; saveVisit();
        S[kind] = makeSession(kind, file.name, items, Date.now()); save(true);
        toast('Đã nạp ' + items.length + ' ' + MODES[kind].unitWord + '.'); go('home');
      };
      if (S[kind] && S[kind].items.some(i => i.res)) confirmSheet('Thay file ' + (kind === 'stock' ? 'tồn kho' : 'tài sản') + '?', 'Số đã kiểm của phiên hiện tại sẽ bị xoá và kiểm lại từ đầu.', 'Thay file', apply);
      else apply();
      return;
    }
    visit[kind] = { fileName: file.name, parsed };
    if (kind === 'stock' && !visit.storeOp && storeByOp(parsed.key)) visit.storeOp = parsed.key;
    if (kind === 'asset' && !visit.storeOp && Object.keys(parsed.keys).length === 1) {
      const s = master.stores.find(m => m.pca === parsed.key); if (s) visit.storeOp = s.op;
    }
    saveVisit(); ui.view = 'home'; rerender();
  }

  async function checkSharedFile() {
    if (!/[?&]share=1/.test(location.search)) return;
    history.replaceState(null, '', location.pathname);
    try {
      const c = await caches.open('kk-shared'), r = await c.match('./shared-file');
      if (!r) return toast('Không nhận được file chia sẻ. Thử lại hoặc dùng nút Chọn file.');
      const name = decodeURIComponent(r.headers.get('x-file-name') || 'shared.xlsx'), blob = await r.blob();
      await c.delete('./shared-file');
      await handleFile(new File([blob], name), 'any');
    } catch (e) { toast('Không đọc được file chia sẻ.'); }
  }

  /* ---------- Excel tạm ---------- */
  async function checksum(s) {
    const canon = JSON.stringify({ store: s.store.op, checker: s.checker, start: s.startedAt, end: s.endedAt,
      items: s.items.map(i => [i.code, i.book, i.res, i.actual, i.reason || '', i.note || '', i.ts]), log: s.log.map(l => [l.t, l.act, l.code, l.detail]) });
    const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(canon));
    return Array.from(new Uint8Array(b)).map(x => x.toString(16).padStart(2, '0')).join('');
  }
  async function buildXlsx() {
    const s = sess(), stock = ui.mode === 'stock', hash = await checksum(s);
    const head = [['BIÊN BẢN ' + MODES[ui.mode].title.toUpperCase()], ['Cửa hàng', s.store.op + ' · ' + s.store.name], ['Người kiểm', s.checker],
      ['Bắt đầu', dt(s.startedAt)], ['Kết thúc', dt(s.endedAt)], ['File nguồn', s.fileName]].concat(stock ? [['Kiểm theo', byQty(s) ? 'Số lượng' : 'Trọng lượng']] : [], [[]]);
    const cols = ['STT', 'Mã', 'Tên'].concat(stock ? ['ĐVT', 'SL sổ', 'CF', 'ĐVT quy đổi', 'TL sổ', 'Sổ sách (theo cách kiểm)', 'Thực tế', 'Chênh lệch'] : ['ĐVT', 'Sổ sách', 'Thực tế', 'Chênh lệch'])
      .concat(['Kết quả', 'Lý do', 'Ghi chú', 'Giờ kiểm']);
    const rows = s.items.map((i, k) => {
      const d = diffOf(i), actual = i.res === 'ok' ? bk(i) : i.actual;
      const mid = stock ? [i.dvt, i.bookQty, i.cf, i.u2, i.book, bk(i), actual, d] : [i.unit, i.book, actual, d];
      return [k + 1, i.code, i.name].concat(mid, [STATUS[st(i)][0] + (i.added ? ' (ngoài DS)' : ''), i.reason || '', i.note || '', dt(i.ts)]);
    });
    const ws1 = XLSX.utils.aoa_to_sheet(head.concat([cols], rows));
    ws1['!cols'] = cols.map((c, k) => ({ wch: k === 2 ? 40 : k === 1 ? 18 : 12 }));
    const log = [['Thời gian', 'Thao tác', 'Mã', 'Chi tiết']].concat(s.log.map(l => [dt(l.t), l.act, l.code, l.detail]),
      [[], ['Checksum SHA-256', hash], ['Ghi chú', 'Checksum tính từ toàn bộ dữ liệu kiểm. Sửa số trong file sẽ không còn khớp checksum.']]);
    const ws2 = XLSX.utils.aoa_to_sheet(log); ws2['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 18 }, { wch: 60 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, 'Kiểm kê'); XLSX.utils.book_append_sheet(wb, ws2, 'Nhật ký');
    const d = new Date(s.endedAt || Date.now());
    const name = 'KiemKe_' + (stock ? 'TonKho' : 'TaiSan') + '_' + s.store.op + '_' + d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0') + '.xlsx';
    return new File([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  function download(file) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = file.name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  /* ---------- events ---------- */
  document.addEventListener('click', e => {
    const g = e.target.closest('[data-go]'); if (g && app.contains(g)) return go(g.dataset.go);
    const p = e.target.closest('[data-pick]'); if (p && app.contains(p)) return pickFile(p.dataset.pick);
    const r = e.target.closest('[data-report]'); if (r && app.contains(r)) return go('report');
  });

  const BIND = {
    home() {
      if (!visit || visit.status !== 'active') return bindSetup();
      app.querySelectorAll('[data-open]').forEach(b => b.onclick = () => {
        const m = b.dataset.open, s = S[m];
        if (!s) return pickFile(m === 'stock' && !visit.dvt ? 'dvt' : m);
        ui.filter = 'todo'; ui.q = '';
        if (s.endedAt) return go('done', m);
        go('count', m);
      });
      document.getElementById('endVisit').onclick = () => confirmSheet('Kết thúc chuyến kiểm?',
        'Toàn bộ số đã kiểm, ảnh và biên bản của ' + esc(visit.storeOp) + ' sẽ bị xoá khỏi máy để bắt đầu cửa hàng mới. Hãy chắc chắn đã gửi biên bản PDF.' + (report && report.confirmedAt ? '' : ' <b>Biên bản chưa được xác nhận.</b>'),
        'Kết thúc & xoá', async () => {
          const checker = visit.checker;
          visit = { status: 'setup', checker, storeOp: '', stock: null, dvt: null, asset: null, checkin: null };
          S.stock = null; S.asset = null; report = null; pdfCache = null;
          await Promise.all([KKDB.set('visit', visit), KKDB.del('sess-stock'), KKDB.del('sess-asset'), KKDB.del('report')]).catch(() => {});
          go('home');
        });
    },
    count() {
      const s = sess(), list = document.getElementById('list');
      const el = document.getElementById('el');
      const tick = () => { if (!document.body.contains(el)) return clearInterval(tm); const m = Math.floor((Date.now() - s.startedAt) / 60000); el.textContent = m >= 60 ? Math.floor(m / 60) + ' giờ ' + (m % 60) + ' phút' : m + ' phút'; };
      const tm = setInterval(tick, 15000); tick();
      app.querySelectorAll('[data-f]').forEach(b => b.onclick = () => { ui.filter = b.dataset.f; ui.edit = null; app.querySelectorAll('[data-f]').forEach(x => x.setAttribute('aria-pressed', x === b)); redrawList(); });
      const q = document.getElementById('q'); q.oninput = () => { ui.q = q.value; redrawList(); };
      const idxOf = t => { const r = t.closest('.row'); return r ? Number(r.dataset.i) : null; };
      const closeEdit = () => { const prev = ui.edit; ui.edit = null; if (prev != null) redrawRow(prev); };
      const openEdit = i => {
        const prev = ui.edit; ui.edit = i; if (prev != null && prev !== i) redrawRow(prev); redrawRow(i);
        const r = rowEl(i); r.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        const f = r.querySelector('[data-aq]') || r.querySelector('[data-act]'); f && f.focus({ preventScroll: true });
      };
      const finishEdit = i => {
        const it = s.items[i]; it.ts = Date.now();
        logIt('Sai', it.code, (it.actual == null ? 'chưa nhập số' : fmt(it.actual) + ' ' + un(it)) + (it.reason ? ' · ' + it.reason : '') + (it.note ? ' · ' + it.note : ''));
        save(); closeEdit(); hideIfChecked(i);
        if (!it.reason && it.actual != null && st(it) !== 'ok') toast('Đã lưu. Nên chọn lý do cho mã ' + it.code + '.');
      };
      list.addEventListener('click', e => {
        const i = idxOf(e.target); if (i == null) return; const it = s.items[i];
        if (e.target.closest('[data-ok]')) {
          if (it.res === 'ok') { it.res = null; it.ts = null; logIt('Bỏ ' + MODES[ui.mode].okLabel, it.code, ''); }
          else { Object.assign(it, { res: 'ok', actual: null, actQty: null, expr: null, exprQ: null, ts: Date.now() }); logIt(MODES[ui.mode].okLabel, it.code, ''); }
          save(); if (ui.edit === i) ui.edit = null; redrawRow(i);
          if (it.res === 'ok') {
            if (ui.filter === 'todo') hideIfChecked(i);
            else { const n = nextTodo(i); if (n != null) { const r = rowEl(n); r && r.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } }
          }
          return;
        }
        const pb = e.target.closest('[data-plus]');
        if (pb) {
          const inp = pb.parentNode.querySelector('input'), v = inp.value.trim().replace(/\++$/, '');
          if (v) inp.value = v + '+';
          inp.focus(); const L = inp.value.length; try { inp.setSelectionRange(L, L); } catch (err) {}
          return;
        }
        if (e.target.closest('[data-bad]')) { if (it.res !== 'bad') { it.res = 'bad'; it.ts = Date.now(); save(); } ui.edit === i ? closeEdit() : openEdit(i); return; }
        if (e.target.closest('[data-editrow]')) return openEdit(i);
        const rb = e.target.closest('[data-reason]');
        if (rb) {
          it.reason = it.reason === rb.dataset.reason ? '' : rb.dataset.reason; save();
          rowEl(i).querySelectorAll('[data-reason]').forEach(b => b.setAttribute('aria-pressed', b.dataset.reason === it.reason));
          if (it.reason === 'Khác') document.getElementById('note-' + i).focus();
          return;
        }
        if (e.target.closest('[data-done]')) finishEdit(i);
      });
      list.addEventListener('pointerdown', e => { if (e.target.closest('[data-plus]')) e.preventDefault(); });
      list.addEventListener('input', e => {
        const i = idxOf(e.target); if (i == null) return; const it = s.items[i];
        if (e.target.matches('[data-note]')) { it.note = e.target.value; save(); return; }
        const sl = e.target.closest('.field') && e.target.closest('.field').querySelector('[data-sumline]');
        if (sl && (e.target.matches('[data-act]') || e.target.matches('[data-aq]'))) sl.innerHTML = sumLine(e.target.value, e.target.dataset.unit || '');
        if (e.target.matches('[data-aq]')) {
          const n = sumExpr(e.target.value); if (Number.isNaN(n)) return; it.actQty = n; it.exprQ = isExpr(e.target.value) ? e.target.value : null;
          it.actual = n == null ? null : r3(n * it.cf); rowEl(i).querySelector('[data-act]').value = decStr(it.actual);
        }
        if (e.target.matches('[data-act]')) {
          const n = sumExpr(e.target.value); if (Number.isNaN(n)) return; it.actual = n; it.expr = isExpr(e.target.value) ? e.target.value : null;
          const qi = rowEl(i).querySelector('[data-aq]'); if (qi && n != null && it.actQty != null && r3(it.actQty * it.cf) !== n) { it.actQty = null; qi.value = ''; }
        }
        save(); softRefresh(i);
      });
      list.addEventListener('keydown', e => {
        if (e.key !== 'Enter' || e.target.tagName === 'TEXTAREA') return; const i = idxOf(e.target); if (i == null) return; e.preventDefault();
        if (e.target.matches('[data-aq]')) rowEl(i).querySelector('[data-act]').focus();
        else if (e.target.matches('[data-act]')) finishEdit(i);
      });
      document.getElementById('add').onclick = () => {
        const unitLabel = ui.mode === 'asset' ? 'Số lượng thực tế (PC)' : 'Số lượng thực tế (theo ĐVT)';
        sheetOpen('<h3>Thêm mã ngoài danh sách</h3><p class="hint" style="margin:0">Mã này không có trong sổ sách nên sẽ tính là Thừa.</p>' +
          '<div class="field"><label for="a-code">Mã *</label><input class="inp mono" id="a-code" autocapitalize="characters" autocomplete="off"></div>' +
          '<div class="field"><label for="a-name">Tên *</label><input class="inp" id="a-name" autocomplete="off"></div>' +
          '<div class="field"><label for="a-qty">' + unitLabel + ' *</label><div class="unit-row"><input class="inp" id="a-qty" inputmode="decimal" autocomplete="off"><button type="button" class="plus" id="a-plus" aria-label="Cộng thêm">+</button></div><span class="sum-line" id="a-sum"></span></div>' +
          '<div class="field"><label for="a-note">Ghi chú</label><input class="inp" id="a-note" placeholder="Ví dụ: hàng chưa có mã trên hệ thống" autocomplete="off"></div>' +
          '<p class="hint" id="a-err" style="margin:0;color:var(--short)"></p><div class="btns"><button class="btn" data-close>Huỷ</button><button class="btn pri" id="a-ok">Thêm</button></div>',
          sc => {
            const codeIn = sc.querySelector('#a-code'), nameIn = sc.querySelector('#a-name');
            const aq = sc.querySelector('#a-qty');
            aq.oninput = () => { sc.querySelector('#a-sum').innerHTML = sumLine(aq.value, ''); };
            sc.querySelector('#a-plus').onpointerdown = e => e.preventDefault();
            sc.querySelector('#a-plus').onclick = () => { const v = aq.value.trim().replace(/\++$/, ''); if (v) aq.value = v + '+'; aq.focus(); };
            codeIn.oninput = () => { const map = dvtMap(), p = map && map[codeIn.value.trim().toUpperCase()]; if (p && p.name && !nameIn.value) nameIn.value = p.name; };
            sc.querySelector('#a-ok').onclick = () => {
              const code = codeIn.value.trim().toUpperCase(), name = nameIn.value.trim(), qv = sumExpr(sc.querySelector('#a-qty').value);
              const err = sc.querySelector('#a-err');
              if (!code || !name || qv == null || Number.isNaN(qv)) { err.textContent = 'Nhập đủ mã, tên và số thực tế.'; return; }
              if (s.items.some(i => i.code === code)) { err.textContent = 'Mã ' + code + ' đã có trong danh sách. Tìm và kiểm trực tiếp ở dòng đó.'; return; }
              const base = { code, name, book: 0, bookQty: 0, cf: null, unit: ui.mode === 'asset' ? 'PC' : 'kg', sub: 'Ngoài danh sách', mv: {} };
              s.items.push(Object.assign(ui.mode === 'stock' ? enrich(base) : base, { res: 'bad', actual: qv, actQty: null, reason: 'Khác', note: sc.querySelector('#a-note').value.trim(), ts: Date.now(), added: true }));
              logIt('Thêm mã', code, name + ' · ' + fmt(qv)); save(true); sc.remove(); ui.filter = 'all'; ui.q = code; ui.edit = null; render(); toast('Đã thêm ' + code + '.');
            };
          });
      };
      document.getElementById('finish').onclick = () => {
        const c = counts(), left = c.todo + c.pend;
        if (!left) return go('done');
        confirmSheet('Còn ' + left + ' mã chưa kiểm xong', 'Bạn có thể quay lại kiểm tiếp, hoặc xem kết quả với những mã đã kiểm.', 'Xem kết quả', () => go('done'));
      };
    },
    done() {
      const s = sess(), end = document.getElementById('end');
      if (end) end.onclick = () => {
        const c = counts(), left = c.todo + c.pend;
        const doEnd = () => { s.endedAt = Date.now(); logIt('Kết thúc', '', 'Khớp ' + c.ok + ' · Thiếu ' + c.short + ' · Thừa ' + c.over + ' · Chưa kiểm ' + left); save(true); render(); };
        if (!left) return doEnd();
        confirmSheet('Kết thúc khi còn ' + left + ' mã chưa kiểm?', 'Những mã này sẽ ghi là "Chưa kiểm" trong biên bản. Sau khi kết thúc vẫn có thể mở lại để kiểm tiếp.', 'Kết thúc', doEnd);
      };
      const reopen = document.getElementById('reopen');
      if (reopen) reopen.onclick = () => { s.endedAt = null; logIt('Mở lại', '', ''); save(true); go('count'); };
      const xl = document.getElementById('xlsx');
      if (xl) xl.onclick = async () => { try { download(await buildXlsx()); } catch (e) { toast('Không tạo được file Excel: ' + e.message); } };
    },
    report() {
      if (!report || !reportCtx()) return;
      KKReport.preload();
      const R = report;
      document.getElementById('takePhoto').onclick = () => { photoTarget = 'checkin'; photoInput.value = ''; photoInput.click(); };
      document.getElementById('evCam').onclick = () => { photoTarget = 'evidence'; photoInput.value = ''; photoInput.click(); };
      document.getElementById('evLib').onclick = () => { photosInput.value = ''; photosInput.click(); };
      app.querySelectorAll('.ev-card').forEach(card => {
        const i = Number(card.dataset.pi), p = R.photos[i];
        card.querySelector('[data-pnote]').oninput = e => { p.note = e.target.value; unconfirm(); saveReport(); };
        card.querySelector('[data-pdel]').onclick = () => confirmSheet('Xoá ảnh ' + (i + 1) + '?', 'Ảnh và lý do đã ghi sẽ bị xoá khỏi biên bản.', 'Xoá ảnh',
          () => { R.photos.splice(i, 1); unconfirm(); saveReport(); rerender(); });
      });
      app.querySelectorAll('.cl-row').forEach(row => {
        const i = Number(row.dataset.ci), c = R.checklist[i];
        row.querySelectorAll('[data-cv]').forEach(b => b.onclick = () => {
          c.v = c.v === b.dataset.cv ? null : b.dataset.cv; c.auto = false; unconfirm(); saveReport(); rerender();
          if (c.v === 'kd') { const n = app.querySelector('.cl-row[data-ci="' + i + '"] [data-cn]'); n && n.focus(); }
        });
        row.querySelector('[data-cn]').oninput = e => { c.note = e.target.value; unconfirm(); saveReport(); };
      });
      // Tính tiền mặt
      const cash = R.cash, cRow = app.querySelector('.cl-row[data-ci="' + CASH_ITEM + '"]'), cItem = R.checklist[CASH_ITEM];
      const updateCash = () => {
        document.getElementById('cashSum').innerHTML = cashSummaryHTML(cash);
        const k = cashCalc(cash);
        if (k.diff != null) {
          if (cItem.auto !== false || cItem.v == null) { cItem.v = k.diff === 0 ? 'dat' : 'kd'; cItem.auto = true; }
          cRow.querySelectorAll('[data-cv]').forEach(b => b.setAttribute('aria-pressed', b.dataset.cv === cItem.v));
        }
        unconfirm(); saveReport();
      };
      cRow.querySelectorAll('[data-money]').forEach(inp => inp.oninput = () => {
        const v = intOf(inp.value); inp.value = v == null ? '' : vnd.format(v);
        if (inp.dataset.money === 'cRev') cash.revenue = v; else cash.transfer = v;
        updateCash();
      });
      cRow.querySelectorAll('[data-den]').forEach(inp => inp.oninput = () => {
        const d = Number(inp.dataset.den), v = intOf(inp.value); inp.value = v == null ? '' : String(v);
        if (v) cash.notes[d] = v; else delete cash.notes[d];
        cRow.querySelector('[data-dsum="' + d + '"]').textContent = v ? vnd.format(v * d) : '';
        updateCash();
      });
      document.getElementById('other').oninput = e => { R.other = e.target.value; document.getElementById('otherPrev').innerHTML = boldPreview(R.other); unconfirm(); saveReport(); };
      const mg = document.getElementById('manager');
      mg.oninput = () => { R.manager = mg.value; unconfirm(); saveReport(); };
      mg.onchange = rerender;
      app.querySelectorAll('[data-sign]').forEach(b => b.onclick = () => {
        const key = b.dataset.sign, ctx = reportCtx();
        openSignPad(key === 'sigManager' ? 'Quản lý cửa hàng ký: ' + (R.manager || '') : 'Người kiểm ký: ' + ctx.checker, data => { R[key] = data; unconfirm(); saveReport(); rerender(); });
      });
      const cf = document.getElementById('confirm');
      cf.onclick = async () => {
        cf.disabled = true; cf.textContent = 'Đang tạo PDF…';
        try { await makePdf(); R.confirmedAt = Date.now(); saveReport(); rerender(); toast('Đã tạo 2 file PDF: Biên Bản Kiểm Tra (Có đính kèm hình ảnh) và Biên Bản Kiểm Tra.'); }
        catch (e) { toast('Không tạo được PDF: ' + e.message); rerender(); }
      };
      const get = async () => pdfCache || makePdf();
      const ps = document.getElementById('pdfSave'); if (ps) ps.onclick = async () => { try { (await get()).forEach((f, i) => setTimeout(() => download(f), i * 700)); } catch (e) { toast('Không tạo được PDF: ' + e.message); } };
      const sh = document.getElementById('pdfShare'); if (sh) sh.onclick = async () => { try { await shareFiles(await get()); } catch (e) { toast('Không tạo được PDF: ' + e.message); } };
      app.querySelectorAll('[data-pdf]').forEach(b => b.onclick = async () => { try { download((await get())[Number(b.dataset.pdf)]); } catch (e) { toast('Không tạo được PDF: ' + e.message); } });
    }
  };

  function bindSetup() {
    const V = visit;
    const ch = document.getElementById('checker');
    ch.oninput = () => { V.checker = ch.value; try { localStorage.setItem('kk-checker', ch.value.trim()); } catch (e) {} saveVisit(); };
    ch.onchange = rerender;
    const q = document.getElementById('storeQ'), res = document.getElementById('storeRes');
    const pick = op => { V.storeOp = op; if (V.checkin) V.checkin = null; saveVisit(); rerender(); };
    if (q) {
      const show = () => {
        const t = plain(q.value.trim());
        if (!t) { res.innerHTML = ''; return; }
        const hits = master.stores.filter(m => plain(m.op + ' ' + m.name + ' ' + m.prov + ' ' + m.pca).includes(t)).slice(0, 8);
        res.innerHTML = hits.length ? hits.map(m => '<button class="store-hit" data-store="' + esc(m.op) + '"><b>' + esc(m.op) + '</b> ' + esc(m.name) + '<span class="hint">' + esc(m.prov + ' · ' + m.area) + '</span></button>').join('')
          : '<div class="hint" style="padding:8px 4px">Không tìm thấy cửa hàng.</div>';
        res.querySelectorAll('[data-store]').forEach(b => b.onclick = () => pick(b.dataset.store));
      };
      q.oninput = show;
    }
    app.querySelectorAll('[data-store]').forEach(b => b.onclick = () => pick(b.dataset.store));
    const clr = document.getElementById('storeClear');
    if (clr) clr.onclick = () => { V.storeOp = ''; V.checkin = null; saveVisit(); rerender(); setTimeout(() => { const s = document.getElementById('storeQ'); s && s.focus(); }, 0); };
    document.getElementById('checkinBtn').onclick = () => { photoTarget = 'visit'; photoInput.value = ''; photoInput.click(); };
    document.getElementById('startVisit').onclick = () => {
      const store = visitStore(); if (!store) return;
      const t = Date.now();
      V.status = 'active'; V.startedAt = t; V.checker = V.checker.trim();
      S.stock = V.stock ? makeSession('stock', V.stock.fileName, V.stock.parsed.items, t) : null;
      const assets = assetItemsFor(store);
      S.asset = assets.length ? makeSession('asset', V.asset.fileName, assets, t) : null;
      report = newReport(store.op); pdfCache = null;
      save(true); saveVisit(); saveReport();
      go('home');
    };
  }

  /* ---------- ảnh ---------- */
  const photoInput = document.getElementById('photo'), photosInput = document.getElementById('photos');
  let photoTarget = 'checkin';
  async function addPhotos(files, target) {
    if (!files.length) return;
    const store = visitStore(); if (!store) return;
    let added = 0;
    if (files.length > 1) toast('Đang xử lý ' + files.length + ' ảnh…');
    for (const f of files) {
      try {
        const p = await stampPhoto(f, store);
        if (target === 'visit') { visit.checkin = p; saveVisit(); }
        else if (target === 'checkin') { report.photo = p.data; report.photoAt = p.at; visit.checkin = p; saveVisit(); }
        else report.photos.push({ data: p.data, at: p.at, note: '' });
        added++;
      } catch (e) { toast('Không xử lý được ảnh ' + f.name + ': ' + e.message); }
    }
    if (!added) return;
    if (report && target !== 'visit') { unconfirm(); await saveReport(); }
    rerender();
    if (target === 'evidence') { const last = app.querySelector('.ev-card:last-child [data-pnote]'); if (last) { last.scrollIntoView({ block: 'center' }); last.focus({ preventScroll: true }); } }
  }
  photoInput.addEventListener('change', () => addPhotos(Array.from(photoInput.files || []).slice(0, 1), photoTarget));
  photosInput.addEventListener('change', () => addPhotos(Array.from(photosInput.files || []), 'evidence'));

  // Danh sách cửa hàng nằm sẵn trong app (data/MASTERR.xlsx). Sửa file đó rồi deploy lại là mọi máy tự cập nhật.
  async function loadBundledMaster() {
    try {
      const r = await fetch(MASTER_URL, { cache: 'no-cache' });
      if (!r.ok) throw new Error(r.status);
      const p = KKParse.parseFile(await r.arrayBuffer(), 'MASTERR.xlsx', XLSX);
      if (p.kind !== 'master' || !p.stores.length) throw new Error('bad master');
      master = { stores: p.stores, updatedAt: Date.now() };
      KKDB.set('master', master).catch(() => {});
    } catch (e) { /* mất mạng và chưa có cache: dùng bản đã lưu trong IndexedDB */ }
  }

  /* ---------- boot ---------- */
  async function boot() {
    try {
      const [m, a, k, v, r] = await Promise.all(['master', 'sess-asset', 'sess-stock', 'visit', 'report'].map(x => KKDB.get(x)));
      if (m && m.stores) master = m;
      S.asset = a || null; S.stock = k || null; visit = v || null; report = r || null;
      // Dữ liệu từ bản cũ (chưa có chuyến kiểm): bỏ để bắt đầu theo luồng mới
      if ((S.asset || S.stock) && !(visit && visit.status === 'active')) { S.asset = null; S.stock = null; report = null; }
    } catch (e) { toast('Trình duyệt chặn lưu trữ offline. Dữ liệu sẽ mất khi đóng app.'); }
    await loadBundledMaster();
    if (S.stock) S.stock.items.forEach(i => { if (i.zero == null) i.zero = !i.added && !i.book && !i.bookQty; });
    if (S.stock && S.stock.by !== 'qty') {
      S.stock.items.forEach(i => { if (i.res === 'bad' && i.actual != null) i.actual = i.actQty != null ? i.actQty : (hasCF(i) ? r3(i.actual / i.cf) : i.actual); });
      S.stock.by = 'qty'; save(true);
    }
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
    render();
    await checkSharedFile();
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  boot();
})();

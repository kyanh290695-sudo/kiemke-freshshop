/* Đọc file đầu vào: ASSET (xlsx), STOCK (csv TCVN3 hoặc xlsx), MASTER (xlsx).
   Dùng được cả trong trình duyệt (window.KKParse) lẫn Node (module.exports) để test. */
(function (root) {
  'use strict';

  // Bảng mã TCVN3 (ABC) -> Unicode, theo giá trị byte
  const TCVN3 = {
    0xA1:'Ă',0xA2:'Â',0xA3:'Ê',0xA4:'Ô',0xA5:'Ơ',0xA6:'Ư',0xA7:'Đ',0xA8:'ă',0xA9:'â',0xAA:'ê',0xAB:'ô',0xAC:'ơ',0xAD:'ư',0xAE:'đ',
    0xB5:'à',0xB6:'ả',0xB7:'ã',0xB8:'á',0xB9:'ạ',0xBB:'ằ',0xBC:'ẳ',0xBD:'ẵ',0xBE:'ắ',0xC6:'ặ',0xC7:'ầ',0xC8:'ẩ',0xC9:'ẫ',0xCA:'ấ',0xCB:'ậ',
    0xCC:'è',0xCE:'ẻ',0xCF:'ẽ',0xD0:'é',0xD1:'ẹ',0xD2:'ề',0xD3:'ể',0xD4:'ễ',0xD5:'ế',0xD6:'ệ',0xD7:'ì',0xD8:'ỉ',0xDC:'ĩ',0xDD:'í',0xDE:'ị',
    0xDF:'ò',0xE1:'ỏ',0xE2:'õ',0xE3:'ó',0xE4:'ọ',0xE5:'ồ',0xE6:'ổ',0xE7:'ỗ',0xE8:'ố',0xE9:'ộ',0xEA:'ờ',0xEB:'ở',0xEC:'ỡ',0xED:'ớ',0xEE:'ợ',
    0xEF:'ù',0xF1:'ủ',0xF2:'ũ',0xF3:'ú',0xF4:'ụ',0xF5:'ừ',0xF6:'ử',0xF7:'ữ',0xF8:'ứ',0xF9:'ự',0xFA:'ỳ',0xFB:'ỷ',0xFC:'ỹ',0xFD:'ý',0xFE:'ỵ'
  };

  function decodeText(buf) {
    const bytes = new Uint8Array(buf);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, ''); }
    catch (e) { /* không phải UTF-8 -> coi là TCVN3 */ }
    let out = '';
    for (let i = 0; i < bytes.length; i++) { const b = bytes[i]; out += TCVN3[b] || String.fromCharCode(b); }
    return out;
  }

  function parseCSV(text) {
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += c;
      } else if (c === '"') q = true;
      else if (c === ',') { row.push(cell); cell = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += c;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  const norm = v => String(v == null ? '' : v).trim().toLowerCase();
  const str = v => (v == null ? '' : String(v)).trim().replace(/\.0+$/, '');
  function num(v) {
    if (typeof v === 'number') return v;
    const s = String(v == null ? '' : v).trim().replace(/,/g, '');
    if (!s) return 0;
    const n = Number(s); return isFinite(n) ? n : 0;
  }
  const r3 = n => Math.round(n * 1000) / 1000;
  const r4 = n => Math.round(n * 10000) / 10000;

  function fmtDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      return String(v.getDate()).padStart(2, '0') + '/' + String(v.getMonth() + 1).padStart(2, '0') + '/' + v.getFullYear();
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) { // số serial Excel
      return fmtDate(new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000));
    }
    return str(v);
  }

  function rowsFromFile(buf, name, XLSX) {
    if (/\.(csv|txt)$/i.test(name || '')) return [parseCSV(decodeText(buf))];
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    return wb.SheetNames.map(n => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }));
  }

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const h = (rows[i] || []).map(norm);
      if (h.includes('rproduct_code')) return { kind: 'stock', i, h };
      if (h.includes('asset') && h.includes('description')) return { kind: 'asset', i, h };
      if (h.includes('operation_code') && h.includes('pca')) return { kind: 'master', i, h };
      if ((h.includes('product_code') || h.includes('mã sản phẩm') || h.includes('ma san pham')) && (h.includes('dvt') || h.includes('đvt'))) return { kind: 'dvt', i, h };
    }
    return null;
  }

  function topKey(counts) {
    let best = '', n = -1;
    Object.keys(counts).forEach(k => { if (counts[k] > n) { best = k; n = counts[k]; } });
    return best;
  }

  function buildStock(rows, hd) {
    const col = name => hd.h.indexOf(name.toLowerCase());
    const c = {
      op: col('rOperation'), opName: col('rOperation_name'), code: col('rProduct_code'), name: col('rProduct_name'),
      qty: col('rBal_qty'), wgt: col('rBal_wgt'), to: col('dfDate_to')
    };
    const pairs = ['rBal', 'rBF'].concat(Array.from({ length: 12 }, (_, k) => 'rTemp' + (k + 1)))
      .map(p => [col(p + '_qty'), col(p + '_wgt')]).filter(([a, b]) => a >= 0 && b >= 0);
    const keys = {}, items = [], seen = new Set();
    let opName = '', cutoff = '', skipped = 0;
    for (let r = hd.i + 1; r < rows.length; r++) {
      const row = rows[r] || [], code = str(row[c.code]);
      if (!code || seen.has(code)) continue;
      seen.add(code);
      const op = str(row[c.op]); if (op) keys[op] = (keys[op] || 0) + 1;
      if (!opName) opName = str(row[c.opName]);
      if (!cutoff && c.to >= 0) cutoff = fmtDate(row[c.to]);
      let cf = null;
      for (const [qi, wi] of pairs) { const q = num(row[qi]), w = num(row[wi]); if (q && w) { cf = r4(w / q); break; } }
      // Phát sinh theo số lượng (cho biên bản): Temp1 Purchase, Temp2 Produce, Temp3 Tran-In, Temp7 Usage, Temp8 Sale, Temp9 Trans-Out, Temp10 Sample&Dmg
      const q = name => { const i = col(name + '_qty'); return i >= 0 ? r3(num(row[i])) : 0; };
      const mv = { bf: q('rBF'), mua: q('rTemp1'), sx: q('rTemp2'), cv: q('rTemp3'), sd: q('rTemp7'), ban: q('rTemp8'), cr: q('rTemp9'), huy: q('rTemp10') };
      const book = r3(num(row[c.wgt])), bookQty = r3(num(row[c.qty]));
      if (!cf && !book && !bookQty) { skipped++; continue; } // CF, Q, W đều = 0: không phát sinh, không tồn → bỏ khỏi danh sách kiểm
      items.push({ code, name: str(row[c.name]) || code, book, bookQty, cf, unit: 'kg', mv, zero: !book && !bookQty });
    }
    return { kind: 'stock', items, keys, key: topKey(keys), meta: { opName, cutoff, skipped } };
  }

  function buildAsset(rows, hd) {
    const col = (...names) => { for (const n of names) { const i = hd.h.indexOf(n); if (i >= 0) return i; } return -1; };
    const c = {
      asset: col('asset'), sub: col('sub-number', 'subnumber'), cap: col('capitalized on'), desc: col('description'),
      pca: col('profit center'), qty: col('quantity'), acq: col('acquis.val.', 'acquis.val', 'acquisition value'), unit: col('base unit of measure', 'unit'), deact: col('deactivation on')
    };
    const keys = {}, items = [], seen = new Set();
    for (let r = hd.i + 1; r < rows.length; r++) {
      const row = rows[r] || [], a = str(row[c.asset]);
      if (!a) continue;
      const sub = c.sub >= 0 ? str(row[c.sub]) : '';
      const code = a + (sub && sub !== '0' ? '-' + sub : '');
      if (seen.has(code)) continue; seen.add(code);
      const pca = str(row[c.pca]); if (pca) keys[pca] = (keys[pca] || 0) + 1;
      const q = c.qty >= 0 && String(row[c.qty]).trim() !== '' ? num(row[c.qty]) : 1;
      const deact = c.deact >= 0 ? fmtDate(row[c.deact]) : '';
      const capDate = c.cap >= 0 ? fmtDate(row[c.cap]) : '';
      items.push({ code, name: str(row[c.desc]) || code, book: r3(q), unit: (c.unit >= 0 && str(row[c.unit])) || 'PC',
        sub: 'Ghi tăng ' + (capDate || '—') + (deact ? ' · Ngừng ' + deact : ''),
        capDate, acq: c.acq >= 0 ? num(row[c.acq]) : 0, pca });
    }
    return { kind: 'asset', items, keys, key: topKey(keys), meta: {} };
  }

  function buildMaster(rows, hd) {
    const col = n => hd.h.indexOf(n);
    const c = { op: col('operation_code'), name: col('operation_code_name'), prov: col('province'), area: col('area'), ba: col('ba'), pca: col('pca') };
    const stores = [];
    for (let r = hd.i + 1; r < rows.length; r++) {
      const row = rows[r] || [], op = str(row[c.op]);
      if (!op) continue;
      stores.push({ op, name: str(row[c.name]), prov: str(row[c.prov]), area: str(row[c.area]), ba: str(row[c.ba]), pca: str(row[c.pca]) });
    }
    return { kind: 'master', stores };
  }

  // Danh mục sản phẩm có tiêu đề: PRODUCT_CODE, DVT, DVT_QUY_DOI, CF, PRODUCT_NAME, NHOM
  function buildProduct(rows, hd) {
    const col = (...ns) => { for (const n of ns) { const i = hd.h.indexOf(n); if (i >= 0) return i; } return -1; };
    const c = { code: col('product_code', 'mã sản phẩm', 'ma san pham'), dvt: col('dvt', 'đvt'), dvt2: col('dvt_quy_doi', 'đvt quy đổi', 'dvt quy doi'),
      cf: col('cf'), name: col('product_name', 'tên sản phẩm', 'ten san pham'), group: col('nhom', 'nhóm', 'group') };
    const map = {};
    for (let r = hd.i + 1; r < rows.length; r++) {
      const row = rows[r] || [], code = str(row[c.code]);
      if (!code) continue;
      const cf = c.cf >= 0 ? num(row[c.cf]) : 0;
      map[code] = { dvt: str(row[c.dvt]), dvt2: c.dvt2 >= 0 ? str(row[c.dvt2]) : '', cf: cf > 0 ? cf : null,
        name: c.name >= 0 ? str(row[c.name]) : '', group: c.group >= 0 ? str(row[c.group]) : '' };
    }
    return { kind: 'dvt', map, count: Object.keys(map).length, op: '', date: '' };
  }
  // DVT.csv xuất từ hệ thống (không có tiêu đề): ..., [5] mã, [6] tên, [7] SL, [8] SL quy đổi, [9] ĐVT, [10] ĐVT quy đổi
  function buildProductCsv(rows) {
    const map = {}, ops = {};
    let date = '';
    rows.forEach(row => {
      if (!row || row.length < 11) return;
      const code = str(row[5]), dvt = str(row[9]);
      if (code && dvt) { const op = str(row[1]); if (op) ops[op] = (ops[op] || 0) + 1; if (!date) date = str(row[11]); }
      if (!code || !dvt) return;
      const q = num(row[7]), w = num(row[8]);
      const prev = map[code];
      const cf = q && w ? r4(w / q) : (prev ? prev.cf : null);
      map[code] = { dvt, dvt2: str(row[10]), cf, name: str(row[6]), group: '', qty: q, qty2: w };
    });
    return { kind: 'dvt', map, count: Object.keys(map).length, op: topKey(ops), ops, date };
  }

  function parseFile(buf, name, XLSX) {
    const sheets = rowsFromFile(buf, name, XLSX);
    for (const rows of sheets) {
      const hd = findHeader(rows);
      if (!hd) continue;
      if (hd.kind === 'stock') return buildStock(rows, hd);
      if (hd.kind === 'asset') return buildAsset(rows, hd);
      if (hd.kind === 'dvt') return buildProduct(rows, hd);
      return buildMaster(rows, hd);
    }
    if (/\.(csv|txt)$/i.test(name || '')) {
      const p = buildProductCsv(sheets[0]);
      if (p.count >= 3) return p;
    }
    throw new Error('Không nhận ra file. Cần file tồn kho STOCK (có cột rProduct_code), file ĐVT tồn kho (DVT.csv) hoặc file tài sản (có cột Asset, Description).');
  }

  const api = { parseFile, decodeText, parseCSV };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.KKParse = api;
})(typeof self !== 'undefined' ? self : this);

/* Tạo báo cáo kiểm kê (PDF 792 × 612 pt) theo theme "S1 Binh Duong - FS173 Di An".
   Toạ độ và cỡ chữ lấy theo file mẫu: nội dung nằm trong dải y 83 → 529.
   Trang: Bìa → Nội dung chính → 1. Kiểm tra cửa hàng và tồn kho (ghi chú + biên bản) → Biên bản kiểm kê kho
          → 2. Tài sản cố định (ghi chú + biên bản) → 3. Hình ảnh → Thank You.
   Các biên bản vẽ theo kích thước gốc rồi thu nhỏ bằng hệ số k để vừa khung. */
(function (root) {
  'use strict';
  const W = 792, H = 612;
  const CORAL = [254, 137, 118], RED = [255, 0, 0], INK = [30, 29, 29], GREY = [89, 89, 89], CYAN = [204, 255, 255], HL = [255, 255, 0];
  const COMPANY = 'Công Ty Cổ Phần Chăn Nuôi C.P. Việt Nam';
  // Cỡ chữ theo mẫu
  const FS = { cover: 52.8, section: 22, note: 15.4, store: 13.2, date: 11, band: 10, toc: 22, tocItem: 13.2, caption: 11, thanks: 60 };
  const NOTE_TOP = 209, NOTE_BOTTOM = 524, TITLE_Y = 180;

  const CHECKLIST = [
    ['Vệ sinh cửa hàng', 'Kiểm tra vệ sinh tại cửa hàng có sạch sẽ (nền nhà, tủ lạnh, cửa kính, counter, phòng pha lóc …)'],
    ['Trưng bày và Bảo quản sản phẩm', 'Sản phẩm có được trưng bày đầy đủ, đẹp mắt, có bảo quản theo đúng nhiệt độ yêu cầu của sản phẩm'],
    ['Hệ thống máy Pos, máy in', 'Kiểm tra hệ thống máy Pos, máy in có hoạt động, có in được bill bán hàng'],
    ['Trang thiết bị, dụng cụ', 'Trang thiết bị, dụng cụ có vệ sinh sạch sẽ, dao, thớt, tủ lạnh nhiệt độ có phù hợp…'],
    ['Hệ thống cân bán', 'Hệ thống cân có được kiểm định thường xuyên theo quy định, có bảo quản tốt không …'],
    ['Hệ thống đèn, điện', 'Dây điện gọn gàng, an toàn. Đèn, quạt hoạt động bình thường, lịch bảo trì máy móc thường xuyên'],
    ['Hệ thống phòng cháy chữa cháy', 'Bình chữa cháy có đúng quy định về PCCC: vị trí đặt, hạn sử dụng, có cập nhật phiếu kiểm tra bảo dưỡng định kỳ?'],
    ['Nhân viên tại cửa hàng', 'Đồng phục nhân viên, thái độ làm việc …'],
    ['Tiền mặt tại cửa hàng', 'Kiểm tra tiền mặt tại cửa hàng (bao gồm tiền lẻ thối lại và tiền bán hàng) khớp với báo cáo.'],
    ['Hệ thống Camera', 'Có hệ thống camera tại cửa hàng hay không? Nếu có thì có đang hoạt động không?']
  ];

  const FONTS = [
    ['Tinos-Regular.ttf', 'Tinos', 'normal'], ['Tinos-Bold.ttf', 'Tinos', 'bold'], ['Tinos-Italic.ttf', 'Tinos', 'italic'],
    ['Montserrat-Regular.ttf', 'Mont', 'normal'], ['Montserrat-SemiBold.ttf', 'Mont', 'bold']
  ];
  const IMAGES = { cp: 'report/logo-cp-circle.jpg', form: 'report/logo-cp.png', cover: 'report/cover-store.jpg' };

  /* ---------- tải thư viện + tài nguyên (đều đã được service worker cache) ---------- */
  function loadScript(src) {
    return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Không tải được ' + src)); document.head.appendChild(s); });
  }
  function b64(buf) { let s = ''; const b = new Uint8Array(buf); for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(s); }
  const blobToDataURL = blob => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(blob); });
  let resP = null;
  function resources() {
    if (resP) return resP;
    resP = (async () => {
      if (!root.jspdf) await loadScript('vendor/jspdf.umd.min.js');
      if (!root.jspdf.jsPDF.API.autoTable) await loadScript('vendor/jspdf.plugin.autotable.min.js');
      const fonts = await Promise.all(FONTS.map(async ([f]) => b64(await (await fetch('fonts/' + f)).arrayBuffer())));
      const img = {};
      for (const k of Object.keys(IMAGES)) img[k] = await blobToDataURL(await (await fetch(IMAGES[k])).blob());
      return { fonts, img };
    })();
    resP.catch(() => { resP = null; });
    return resP;
  }

  /* ---------- định dạng ---------- */
  const nQty = new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
  const nMoney = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
  const q = v => (v == null || v === '') ? '' : (Math.abs(v) < 1e-9 ? '-' : nQty.format(v));
  const qBlank = v => (!v) ? '' : nQty.format(v);
  const money = v => v ? nMoney.format(v) : '';
  const r3 = n => Math.round(n * 1000) / 1000;
  const T = s => String(s == null ? '' : s).normalize('NFC');
  const pad = n => String(n).padStart(2, '0');
  const dmy = t => { const d = new Date(t); return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear(); };
  const ngay = t => { const d = new Date(t); return 'Ngày ' + d.getDate() + ' Tháng ' + (d.getMonth() + 1) + ' Năm ' + d.getFullYear(); };
  const hhmm = t => { if (!t) return '……'; const d = new Date(t); return pad(d.getHours()) + 'H' + pad(d.getMinutes()); };

  /* ---------- khung trang theo theme ---------- */
  // Logo CP + vạch đỏ + FRESH SHOP (bìa thấp hơn các trang trong 10pt như mẫu)
  function brand(doc, R, cover) {
    const y = cover ? 103 : 93;
    doc.addImage(R.img.cp, 'JPEG', cover ? 21 : 18, y, 42, 42);
    doc.setDrawColor(...RED); doc.setLineWidth(1); doc.line(cover ? 74 : 71, y, cover ? 74 : 71, y + 41);
    doc.setTextColor(...RED); doc.setFont('Mont', 'bold'); doc.setFontSize(22);
    const x = cover ? 83 : 80; doc.text('FRESH', x, y + 29);
    const fw = doc.getTextWidth('FRESH ');
    doc.setFont('Mont', 'normal'); doc.text('SHOP', x + fw, y + 29);
    doc.setTextColor(0, 0, 0); doc.setDrawColor(0);
  }
  function title(doc, text) { doc.setFont('Tinos', 'bold'); doc.setFontSize(FS.section); doc.setTextColor(0, 0, 0); doc.text(T(text), 12, TITLE_Y); }
  function newPage(doc, R, t) { doc.addPage([W, H], 'l'); brand(doc, R); if (t) title(doc, t); }
  function fitImage(doc, data, fmt, x, y, bw, bh, valign) {
    const p = doc.getImageProperties(data), k = Math.min(bw / p.width, bh / p.height), w = p.width * k, h = p.height * k;
    const oy = valign === 'top' ? 0 : (bh - h) / 2;
    doc.addImage(data, fmt, x + (bw - w) / 2, y + oy, w, h);
    return { x: x + (bw - w) / 2, y: y + oy, w, h };
  }
  const pw = doc => doc.internal.pageSize.getWidth();
  function coralBar(doc, x, y, w, h) { doc.setFillColor(...CORAL); doc.rect(x, y, w, h, 'F'); }

  /* ---------- chữ nhiều kiểu (đậm / thường / tô vàng), tự xuống dòng ---------- */
  function richLines(doc, runs, width, fs) {
    const words = [];
    runs.forEach(r => T(r.text).split(/(\s+)/).forEach(w => { if (w) words.push({ t: w, bold: !!r.bold, hl: !!r.hl }); }));
    const lines = []; let cur = [], cw = 0;
    const wOf = w => { doc.setFont('Tinos', w.bold ? 'bold' : 'normal'); doc.setFontSize(fs); return doc.getTextWidth(w.t); };
    words.forEach(w => {
      const ww = wOf(w), space = /^\s+$/.test(w.t);
      if (space) { if (cur.length) { cur.push(Object.assign({ w: ww }, w)); cw += ww; } return; }
      if (cw + ww > width && cur.length) {
        while (cur.length && /^\s+$/.test(cur[cur.length - 1].t)) cur.pop();
        lines.push(cur); cur = []; cw = 0;
      }
      cur.push(Object.assign({ w: ww }, w)); cw += ww;
    });
    if (cur.length) lines.push(cur);
    return lines;
  }
  function drawRich(doc, lines, x, y, fs, lh) {
    lines.forEach((ln, i) => {
      let cx = x; const by = y + i * lh;
      // tô vàng liền một khối cho cụm chữ được đánh dấu
      let hs = null;
      ln.forEach(w => { if (w.hl) { if (hs == null) hs = cx; } else if (hs != null) { doc.setFillColor(...HL); doc.rect(hs, by - fs * 0.86, cx - hs, fs * 1.12, 'F'); hs = null; } cx += w.w; });
      if (hs != null) { doc.setFillColor(...HL); doc.rect(hs, by - fs * 0.86, cx - hs, fs * 1.12, 'F'); }
      cx = x;
      ln.forEach(w => { doc.setFont('Tinos', w.bold ? 'bold' : 'normal'); doc.setFontSize(fs); doc.setTextColor(...INK); doc.text(w.t, cx, by); cx += w.w; });
    });
    doc.setTextColor(0, 0, 0);
    return y + lines.length * lh;
  }
  // "Trưng bày: Thiếu nhãn" → phần trước dấu ":" in đậm (tô vàng ở mục Stock / Tài sản như mẫu)
  function labelRuns(text, hl) {
    const s = T(text).trim(), i = s.indexOf(':');
    if (i > 0 && i <= 60) return [{ text: s.slice(0, i + 1), bold: true, hl }, { text: s.slice(i + 1) }];
    return [{ text: s }];
  }
  const filled = items => (items || []).filter(it => (it.content || '').trim() || (it.action || '').trim());
  const hasNotes = secs => secs.some(s => filled(s.items).length);
  // Khối ghi chú: [{label, hl, items:[{content, action}]}]. draw=false để đo chiều cao.
  function notesBlock(doc, sections, x, y, width, fs, draw) {
    const lh = fs * 1.8; let cy = y;
    sections.forEach(sec => {
      const items = filled(sec.items);
      if (!items.length) return;
      if (draw) {
        doc.setFont('Tinos', 'bold'); doc.setFontSize(fs); doc.setTextColor(...INK);
        const lbl = T(sec.label + ':'); doc.text(lbl, x, cy);
        doc.setDrawColor(...INK); doc.setLineWidth(0.9); doc.line(x, cy + 1.6, x + doc.getTextWidth(lbl), cy + 1.6); doc.setDrawColor(0);
      }
      cy += lh;
      items.forEach(it => {
        if ((it.content || '').trim()) {
          const lines = richLines(doc, labelRuns(it.content, sec.hl), width - 20, fs);
          if (draw) { doc.setFont('Tinos', 'normal'); doc.setFontSize(fs); doc.setTextColor(...INK); doc.text('•', x, cy); drawRich(doc, lines, x + 20, cy, fs, lh); }
          cy += lines.length * lh;
        }
        if ((it.action || '').trim()) {
          const lines = richLines(doc, [{ text: '→ ' }, { text: 'HXL: ', bold: true }, { text: T(it.action).trim() }], width, fs);
          if (draw) drawRich(doc, lines, x, cy, fs, lh);
          cy += lines.length * lh;
        }
      });
    });
    return cy - y;
  }
  // Giữ cỡ 15,4pt như mẫu; chỉ thu nhỏ khi ghi chú quá dài so với khung
  function placeNotes(doc, sections, x, y, width, maxY) {
    let fs = FS.note;
    while (fs > 9 && y + notesBlock(doc, sections, x, y, width, fs, false) - fs * 1.8 > maxY) fs -= 0.4;
    notesBlock(doc, sections, x, y, width, fs, true);
  }

  /* ---------- khối ký tên (kích thước gốc × k) ---------- */
  function signBlock(doc, ctx, x, y, w, k, upper) {
    const cols = [x + w * 0.02, x + w * 0.36, x + w * 0.72], cw = w * 0.26;
    const labels = upper ? ['QUẢN LÝ CỬA HÀNG', 'KẾ TOÁN', 'GIÁM ĐỐC CHI NHÁNH'] : ['Quản lý cửa hàng', 'KẾ TOÁN', 'GIÁM ĐỐC CHI NHÁNH'];
    doc.setFont('Tinos', 'bold'); doc.setFontSize(11 * k);
    labels.forEach((l, i) => doc.text(T(l), cols[i] + cw / 2, y, { align: 'center' }));
    [ctx.report.sigManager, ctx.report.sigChecker].forEach((s, i) => { if (s) fitImage(doc, s, 'PNG', cols[i] + cw * 0.1, y + 8 * k, cw * 0.8, 62 * k); });
    doc.text(T(ctx.report.manager), cols[0] + cw / 2, y + 88 * k, { align: 'center' });
    doc.text(T(ctx.checker), cols[1] + cw / 2, y + 88 * k, { align: 'center' });
    doc.setFont('Tinos', 'normal');
    doc.text('.................................................', cols[2] + cw / 2, y + 88 * k, { align: 'center' });
    return y + 96 * k;
  }
  function checkbox(doc, x, y, s, on) {
    doc.setLineWidth(s * 0.075); doc.rect(x, y, s, s);
    if (on) { doc.setLineWidth(s * 0.16); doc.line(x + s * 0.19, y + s * 0.52, x + s * 0.44, y + s * 0.81); doc.line(x + s * 0.44, y + s * 0.81, x + s * 0.87, y + s * 0.19); }
  }

  /* ---------- biên bản kiểm tra cửa hàng (gốc rộng 840) ---------- */
  function storeForm(doc, R, ctx, x0, y0, w, pad) {
    const k = w / 840, rep = ctx.report, cp = (pad || 4) * k;
    let y = y0;
    doc.setDrawColor(0); doc.setLineWidth(1.2 * k);
    const headH = 74 * k;
    doc.rect(x0, y, w, headH);
    doc.addImage(R.img.form, 'PNG', x0 + 14 * k, y + 8 * k, 68 * k, 54 * k);
    doc.setFont('Tinos', 'italic'); doc.setFontSize(12 * k); doc.text('CÔNG TY CỔ PHẦN CHĂN NUÔI C.P VIỆT NAM', x0 + 110 * k, y + 22 * k);
    doc.setFont('Tinos', 'bold'); doc.setFontSize(19 * k); doc.text('BIÊN BẢN KIỂM TRA CỬA HÀNG', x0 + 110 * k, y + 56 * k);
    doc.setLineWidth(0.6 * k); doc.roundedRect(x0 + w - 220 * k, y + 34 * k, 210 * k, 32 * k, 4 * k, 4 * k);
    doc.setFont('Tinos', 'normal'); doc.setFontSize(10 * k);
    doc.text('Thời gian bắt đầu :  ' + hhmm(ctx.start), x0 + w - 212 * k, y + 47 * k);
    doc.text('Thời gian kết thúc:  ' + hhmm(ctx.end), x0 + w - 212 * k, y + 60 * k);
    y += headH;
    const info = [[120, 'Cửa Hàng', false], [210, ctx.store.name, true], [62, 'Quản Lý', false], [170, rep.manager, true], [60, 'Ngày', false], [218, dmy(ctx.date), true]];
    let x = x0; doc.setLineWidth(1.2 * k); doc.rect(x0, y, w, 24 * k); doc.setFontSize(11 * k);
    info.forEach(([cw0, t, red]) => {
      const cw = cw0 * k;
      doc.setFont('Tinos', red ? 'normal' : 'bold'); doc.setTextColor(...(red ? [200, 0, 0] : [0, 0, 0]));
      const tw = Math.min(doc.getTextWidth(T(t)), cw - 8 * k);
      doc.text(T(t), x + cw / 2, y + 16 * k, { align: 'center', maxWidth: cw - 8 * k });
      if (red && t) { doc.setDrawColor(200, 0, 0); doc.setLineWidth(0.6 * k); doc.line(x + cw / 2 - tw / 2, y + 18 * k, x + cw / 2 + tw / 2, y + 18 * k); doc.setDrawColor(0); }
      x += cw; if (x < x0 + w - 1) { doc.setLineWidth(1.2 * k); doc.line(x, y, x, y + 24 * k); }
    });
    doc.setTextColor(0, 0, 0);
    y += 28 * k;
    doc.autoTable({
      startY: y, margin: { left: x0, right: pw(doc) - x0 - w }, tableWidth: w, theme: 'grid',
      styles: { font: 'Tinos', fontSize: 10.5 * k, textColor: 0, lineColor: 0, lineWidth: 0.6 * k, cellPadding: cp, valign: 'middle' },
      headStyles: { fillColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
      head: [[{ content: 'STT', rowSpan: 2 }, { content: 'CHỈ TIÊU KIỂM TRA', rowSpan: 2 }, { content: 'ĐÁNH GIÁ', colSpan: 2 }, { content: 'GHI CHÚ', rowSpan: 2 }], ['Đạt', 'Không đạt']],
      body: CHECKLIST.map(([label], i) => [String(i + 1), T(label), '', '', T((rep.checklist[i] || {}).note || '')]),
      columnStyles: { 0: { cellWidth: 60 * k, halign: 'center' }, 1: { cellWidth: 330 * k, fontStyle: 'bold' }, 2: { cellWidth: 46 * k }, 3: { cellWidth: 62 * k }, 4: { cellWidth: w - 498 * k, fontSize: 9.5 * k } },
      didDrawCell: d => {
        if (d.section !== 'body' || (d.column.index !== 2 && d.column.index !== 3)) return;
        const v = (rep.checklist[d.row.index] || {}).v, s = 8 * k;
        checkbox(doc, d.cell.x + d.cell.width / 2 - s / 2, d.cell.y + d.cell.height / 2 - s / 2, s, d.column.index === 2 ? v === 'dat' : v === 'kd');
      }
    });
    y = doc.lastAutoTable.finalY;
    doc.setFont('Tinos', 'normal'); doc.setFontSize(10.5 * k);
    const other = boldLines(doc, rep.other, w - 16 * k, 10.5 * k);
    const lines = Math.max(pad ? 5 : 3, other.length), boxH = (22 + lines * 15) * k;
    doc.setLineWidth(1.2 * k); doc.rect(x0, y, w, boxH);
    doc.setFont('Tinos', 'bold'); doc.text('CÁC VẤN ĐỀ KHÁC', x0 + 4 * k, y + 14 * k);
    doc.setLineWidth(0.6 * k); doc.line(x0 + 4 * k, y + 16 * k, x0 + 4 * k + doc.getTextWidth('CÁC VẤN ĐỀ KHÁC'), y + 16 * k);
    doc.setFont('Tinos', 'normal');
    for (let i = 0; i < lines; i++) { const ly = y + (22 + (i + 1) * 15) * k; doc.setDrawColor(150); doc.line(x0, ly, x0 + w, ly); doc.setDrawColor(0); if (other[i] && other[i].length) drawRich(doc, [other[i]], x0 + 8 * k, ly - 4 * k, 10.5 * k, 0); }
    y += boxH + 10 * k;
    doc.setLineWidth(1.2 * k); doc.rect(x0, y, w, 128 * k);
    signBlock(doc, ctx, x0, y + 18 * k, w, k, false);
    return y + 128 * k;
  }
  // Chiều cao biên bản cửa hàng ở k = 1 (để căn giữa trong khung)
  // "*Tủ đông số 2*: kêu to" → phần trong dấu * in đậm; xuống dòng giữ nguyên. Dấu * lẻ thì in nguyên văn.
  function boldRuns(p) {
    return p.split(/(\*[^*\n]+\*)/).filter(Boolean).map(t => /^\*[^*\n]+\*$/.test(t) ? { text: t.slice(1, -1), bold: true } : { text: t });
  }
  function boldLines(doc, text, width, fs) {
    const out = [];
    T(text || '').split(/\r?\n/).forEach(p => { if (p.trim()) richLines(doc, boldRuns(p), width, fs).forEach(l => out.push(l)); else if (out.length) out.push([]); });
    while (out.length && !out[out.length - 1].length) out.pop();
    return out;
  }
  const STORE_FORM_H = rep => 74 + 28 + 36 + 10 * 26 + 22 + Math.max(3, Math.ceil(((rep.other || '').length || 1) / 120)) * 15 + 10 + 128;

  /* ---------- đầu biên bản kho / tài sản ---------- */
  function docHeader(doc, ctx, heading, y, left, right, sess, k) {
    doc.setFont('Tinos', 'bold'); doc.setFontSize(9.5 * k);
    doc.text(COMPANY, left, y); doc.text(T('Cửa Hàng ' + ctx.store.name), left, y + 12 * k);
    doc.setFontSize(17 * k); doc.text(heading, (left + right) / 2, y + 36 * k, { align: 'center' });
    doc.setFont('Tinos', 'italic'); doc.setFontSize(9.5 * k); doc.text(ngay(ctx.date), right, y + 48 * k, { align: 'right' });
    doc.setFont('Tinos', 'normal');
    doc.text('Thời gian bắt đầu: ' + hhmm(ctx.start) + '    Thời gian kết thúc: ' + hhmm(ctx.end), (left + right) / 2, y + 60 * k, { align: 'center' });
    return y + 66 * k;
  }

  /* ---------- biên bản kiểm kê hàng tồn kho (gốc rộng 1220) ---------- */
  function stockPages(doc, R, ctx, s, heading) {
    newPage(doc, R, heading);
    stockTable(doc, ctx, s, { left: 36, right: W - 36, top: 194, contTop: 200, bottom: 580, sigLimit: 604,
      onPage: () => { brand(doc, R); title(doc, heading); }, newPage: () => { newPage(doc, R, heading); return 210; } });
  }
  // o = { left, right, top, contTop, bottom, sigLimit, kf (cỡ chữ), onPage (trang nối tiếp), newPage () → y }
  function stockTable(doc, ctx, s, o) {
    const left = o.left, right = o.right, k = (right - left) / 1220, kf = o.kf || k;
    const y = docHeader(doc, ctx, 'BIÊN BẢN KIỂM KÊ HÀNG TỒN KHO', o.top, left, right, s, kf);
    const byQ = s.by === 'qty';
    const body = s.items.map((it, i) => {
      const mv = it.mv || {}, cfOk = it.cf && it.cf !== 1;
      let actual = null;
      if (it.res === 'ok') actual = it.bookQty;
      else if (it.res === 'bad' && it.actual != null) actual = byQ ? it.actual : (cfOk ? (it.actQty != null ? it.actQty : r3(it.actual / it.cf)) : it.actual);
      const diff = actual == null ? null : r3(actual - it.bookQty);
      const note = [it.res == null && !it.zero ? 'Chưa kiểm' : '', it.reason || '', it.note || '', it.added ? 'Ngoài danh sách' : ''].filter(Boolean).join(' · ');
      return [i + 1, it.code, T(it.name), T(it.dvt || ((it.cf == null || it.cf === 1) ? 'Kg' : 'ĐV')), q(mv.bf), qBlank(mv.sx), qBlank(mv.mua), qBlank(mv.cv), qBlank(mv.ban), qBlank(mv.sd), '', qBlank(mv.cr), qBlank(mv.huy),
        qBlank(it.bookQty), actual == null ? '' : (actual ? nQty.format(actual) : '-'), '', diff == null ? '' : (diff ? nQty.format(diff) : '-'), T(note)];
    });
    const cw = [22, 92, 190, 36, 42, 38, 38, 40, 38, 38, 38, 40, 34, 46, 46, 48, 46];
    const cs = {}; cw.forEach((v, i) => { cs[i] = { cellWidth: v * k }; });
    [4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 16].forEach(i => { cs[i].halign = 'right'; });
    cs[0].halign = 'center'; cs[17] = {};
    doc.autoTable({
      startY: y, margin: { left, right: pw(doc) - right, top: o.contTop, bottom: doc.internal.pageSize.getHeight() - o.bottom }, theme: 'grid',
      styles: { font: 'Tinos', fontSize: 7.8 * kf, textColor: 0, lineColor: 0, lineWidth: 0.4 * k, cellPadding: { top: 1.6 * kf, bottom: 1.6 * kf, left: 2 * k, right: 2 * k }, valign: 'middle', overflow: 'linebreak' },
      headStyles: { fillColor: CYAN, fontStyle: 'bold', halign: 'center', fontSize: 6.8 * kf },
      head: [
        [{ content: 'STT', rowSpan: 3 }, { content: 'MÃ SẢN PHẨM', rowSpan: 3 }, { content: 'TÊN SẢN PHẨM', rowSpan: 3 }, { content: 'DVT', rowSpan: 3 }, { content: 'THEO SỔ SÁCH', colSpan: 9 },
          { content: 'TỒN\nTHEO SỔ\nSÁCH', rowSpan: 3 }, { content: 'KIỂM KÊ\nTHỰC TẾ', rowSpan: 3 }, { content: 'SỐ LƯỢNG\nBÁN/NHẬP', rowSpan: 3 }, { content: 'CHÊNH\nLỆCH\n(TT/SS)', rowSpan: 3 }, { content: 'Ghi Chú', rowSpan: 3 }],
        [{ content: 'TỒN\nĐẦU', rowSpan: 2 }, { content: 'NHẬP', colSpan: 3 }, { content: 'XUẤT', colSpan: 5 }],
        ['SẢN XUẤT', 'MUA', 'Chuyển vào', 'BÁN', 'SỬ DỤNG', { content: 'DataPOS', styles: { textColor: [0, 0, 220] } }, { content: 'Chuyển ra', styles: { textColor: [0, 0, 220] } }, 'HỦY']
      ],
      body, columnStyles: cs,
      didParseCell: d => {
        if (d.section !== 'body') return;
        const it = s.items[d.row.index];
        if (d.column.index === 16 && d.cell.raw && d.cell.raw !== '-') d.cell.styles.textColor = String(d.cell.raw).startsWith('-') ? [194, 58, 43] : [37, 96, 168];
        if (it && it.res == null && d.column.index === 17) d.cell.styles.textColor = [140, 140, 140];
      },
      willDrawPage: d => { if (d.pageNumber > 1 && o.onPage) o.onPage(); }
    });
    let fy = doc.lastAutoTable.finalY + 14;
    if (fy + 100 * kf + 10 > o.sigLimit) fy = o.newPage();
    fy = signBlock(doc, ctx, left, fy, right - left, kf, true);
    doc.setFont('Tinos', 'normal'); doc.setFontSize(8.5 * kf);
    doc.text('Ghi chú: In báo cáo stock và báo cáo "Sale Report(List by Product)" trên máy Pos đính kèm biên bản kiểm tra', left + 20 * k, fy + 10 * kf);
  }

  /* ---------- biên bản tài sản cố định (gốc rộng 924) ---------- */
  // o = { left, right, top, contTop, bottom, sigLimit, kf, onPage, newPage () → y }
  function assetForm(doc, ctx, s, o) {
    const left = o.left, right = o.right, k = (right - left) / 924, kf = o.kf || k;
    const y = docHeader(doc, ctx, 'BIÊN BẢN KIỂM KÊ TÀI SẢN CỐ ĐỊNH', o.top, left, right, s, kf);
    let sb = 0, sa = 0, sr = 0, sd = 0;
    const body = s.items.map((it, i) => {
      const actual = it.res === 'ok' ? it.book : (it.res === 'bad' ? it.actual : null);
      const diff = actual == null ? null : r3(actual - it.book);
      sb += it.book || 0; sa += it.acq || 0; sr += actual || 0; if (diff) sd += diff;
      const note = [it.res == null ? 'Chưa kiểm' : '', it.reason || '', it.note || '', it.added ? 'Ngoài danh sách' : ''].filter(Boolean).join(' · ');
      return [i + 1, it.code, it.capDate || '', T(it.name), it.pca || ctx.store.pca || '', it.book, money(it.acq), actual == null ? '' : actual, diff == null ? '' : (diff ? diff : '-'), T(note)];
    });
    const cw = [32, 62, 66, 240, 74, 58, 78, 48, 56];
    const al = ['center', 'left', 'right', 'left', 'center', 'right', 'right', 'right', 'center'];
    const cs = {}; cw.forEach((v, i) => { cs[i] = { cellWidth: v * k, halign: al[i] }; }); cs[9] = {};
    doc.autoTable({
      startY: y, margin: { left, right: pw(doc) - right, top: o.contTop, bottom: doc.internal.pageSize.getHeight() - o.bottom }, theme: 'grid',
      styles: { font: 'Tinos', fontSize: 9 * kf, textColor: 0, lineColor: 0, lineWidth: 0.5 * k, cellPadding: { top: 1.8 * kf, bottom: 1.8 * kf, left: 3 * k, right: 3 * k }, valign: 'middle' },
      headStyles: { fillColor: [255, 255, 255], fontStyle: 'bold', halign: 'center', fontSize: 9 * kf },
      footStyles: { fillColor: [255, 255, 255], fontStyle: 'bold', halign: 'right' },
      head: [[{ content: 'STT', rowSpan: 2 }, { content: 'Mã\nTài Sản', rowSpan: 2 }, { content: 'Ngày Nhận', rowSpan: 2 }, { content: 'Tên Tài Sản', rowSpan: 2 }, { content: 'Mã\nBộ phận', rowSpan: 2 },
        { content: 'Số Lượng', colSpan: 4 }, { content: 'Ghi Chú', rowSpan: 2 }], ['Theo sổ sách', 'Nguyên Giá', 'Thực tế', 'Chênh Lệch\nTT/SC']],
      body,
      foot: [[{ content: 'TỔNG', colSpan: 5, styles: { halign: 'center' } }, sb, money(sa), sr, sd || '-', '']],
      showFoot: 'lastPage', columnStyles: cs,
      willDrawPage: d => { if (d.pageNumber > 1 && o.onPage) o.onPage(); }
    });
    let fy = doc.lastAutoTable.finalY + 14;
    if (fy + 92 * kf > o.sigLimit) fy = o.newPage();
    signBlock(doc, ctx, left, fy, right - left, kf, true);
  }

  /* ---------- các trang ---------- */
  // Khung ảnh bên phải (446→732, 185→427 như mẫu), căn giữa trong khung.
  // Bìa: ảnh mặt tiền Fresh Shop cố định · Nội dung chính: ảnh check-in của lần kiểm
  function photoPanel(doc, ctx, R) {
    const bx = 440, by = 160, bw = 300, bh = 290;
    if (R) fitImage(doc, R.img.cover, 'JPEG', bx, by, bw, bh);
    else if (ctx.report.photo) fitImage(doc, ctx.report.photo, 'JPEG', bx, by, bw, bh);
    else {
      doc.setDrawColor(210); doc.setLineDashPattern([5, 5], 0); doc.rect(bx, by + 30, bw, bh - 60); doc.setLineDashPattern([], 0); doc.setDrawColor(0);
      doc.setFont('Tinos', 'italic'); doc.setFontSize(FS.date); doc.setTextColor(...GREY); doc.text('Chưa có ảnh check-in', bx + bw / 2, by + bh / 2, { align: 'center' }); doc.setTextColor(0, 0, 0);
    }
  }
  function cover(doc, R, ctx) {
    brand(doc, R, true);
    doc.setFont('Tinos', 'bold'); doc.setFontSize(FS.cover); doc.setTextColor(0, 0, 0);
    doc.text('BÁO CÁO', 22, 233, { charSpace: 3.2 });
    doc.text('KIỂM KÊ', 22, 297, { charSpace: 3.2 });
    coralBar(doc, 21, 315, 223, 3);
    doc.setFontSize(FS.store); doc.setTextColor(...INK); doc.text(T(ctx.store.name), 22, 355);
    doc.setFont('Tinos', 'normal'); doc.setFontSize(FS.date); doc.setTextColor(...GREY); doc.text(dmy(ctx.date), 22, 378);
    doc.setTextColor(0, 0, 0);
    photoPanel(doc, ctx, R);
    coralBar(doc, 0, 496, 403, 20);
    doc.setFont('Tinos', 'normal'); doc.setFontSize(FS.band); doc.setTextColor(255, 255, 255);
    doc.text('Người kiểm tra:', 6, 509.5); doc.text(T(ctx.checker), 92, 509.5);
    doc.setTextColor(0, 0, 0);
  }
  function contents(doc, R, ctx, sections) {
    doc.addPage([W, H], 'l'); brand(doc, R, true);
    doc.setFont('Tinos', 'bold'); doc.setFontSize(FS.toc); doc.setTextColor(0, 0, 0);
    doc.text('NỘI DUNG CHÍNH', 30, 241, { charSpace: -0.8 });
    doc.setFont('Tinos', 'normal'); doc.setFontSize(FS.tocItem);
    sections.forEach((s, i) => doc.text(T(s), 42, 297 + i * 44));
    coralBar(doc, 40, 297 + (sections.length - 1) * 44 + 34, 80, 5);
    photoPanel(doc, ctx);
  }
  // Ghi chú bên trái, biên bản bên phải (375→774 như mẫu); không có ghi chú thì biên bản đặt giữa
  function storePage(doc, R, ctx, heading, notes) {
    newPage(doc, R, heading);
    if (hasNotes(notes)) {
      placeNotes(doc, notes, 20, NOTE_TOP, 332, NOTE_BOTTOM);
      const w = 399, h = STORE_FORM_H(ctx.report) * w / 840;
      storeForm(doc, R, ctx, 375, Math.max(112, (112 + 524 - h) / 2), w);
    } else {
      const w = 560, h = STORE_FORM_H(ctx.report) * w / 840;
      storeForm(doc, R, ctx, (W - w) / 2, Math.max(196, 196 + (524 - 196 - h) / 2), w);
    }
  }
  // Trang tài sản: chỉ có biên bản, không ghi chú
  function assetPages(doc, R, ctx, s, heading) {
    newPage(doc, R, heading);
    assetForm(doc, ctx, s, { left: 96, right: W - 96, top: 194, contTop: 200, bottom: 580, sigLimit: 604,
      onPage: () => { brand(doc, R); title(doc, heading); }, newPage: () => { newPage(doc, R, heading); return 210; } });
  }
  // 3 ảnh một trang, chú thích (lý do) dưới mỗi ảnh; không đánh số ảnh
  function photoPages(doc, R, ctx, heading) {
    const list = (ctx.report.photos || []).filter(p => p && p.data);
    for (let p0 = 0; p0 < list.length; p0 += 3) {
      newPage(doc, R, heading);
      const group = list.slice(p0, p0 + 3), bw = 220, gap = 26, total = group.length * bw + (group.length - 1) * gap;
      let x = (W - total) / 2;
      const notes = group.map(p => T(p.note || '').trim());
      const same = notes.length > 1 && notes.every(n => n && n === notes[0]);
      group.forEach((p, i) => {
        fitImage(doc, p.data, 'JPEG', x, 200, bw, 290);
        if (!same && notes[i]) {
          doc.setFont('Tinos', 'normal'); doc.setFontSize(FS.caption); doc.setTextColor(...INK);
          doc.text(doc.splitTextToSize(notes[i], bw).slice(0, 2), x + bw / 2, 508, { align: 'center', lineHeightFactor: 1.25 });
          doc.setTextColor(0, 0, 0);
        }
        x += bw + gap;
      });
      // cùng một lý do cho cả nhóm ảnh → một chú thích chung ở giữa như mẫu
      if (same) { doc.setFont('Tinos', 'normal'); doc.setFontSize(FS.caption); doc.setTextColor(...INK); doc.text(doc.splitTextToSize(notes[0], 600).slice(0, 2), W / 2, 508, { align: 'center', lineHeightFactor: 1.25 }); doc.setTextColor(0, 0, 0); }
    }
  }
  function thanks(doc, R) {
    doc.addPage([W, H], 'l'); brand(doc, R, true);
    doc.setFont('Tinos', 'bold'); doc.setFontSize(FS.thanks); doc.setTextColor(0, 0, 0);
    doc.text('Thank You', W / 2, 322, { align: 'center', charSpace: 2 });
    coralBar(doc, W / 2 - 60, 344, 120, 4);
  }

  /* ctx = { store, checker, date, start, end, report, asset (session|null), stock (session|null) }
     report.notes = { store: [{content, action}], stock: [...] } — mục trống không đưa vào PDF */
  async function build(ctx) {
    const R = await resources();
    const doc = new root.jspdf.jsPDF({ orientation: 'l', unit: 'pt', format: [W, H], compress: true });
    FONTS.forEach(([f, fam, style], i) => { doc.addFileToVFS(f, R.fonts[i]); doc.addFont(f, fam, style); });
    doc.setProperties({ title: 'Báo cáo kiểm kê ' + ctx.store.op, author: ctx.checker, creator: 'Kiểm kê Freshshop' });
    const notes = ctx.report.notes || {};
    const hasPhotos = (ctx.report.photos || []).some(p => p && p.data);
    const secs = ['Kiểm tra cửa hàng và tồn kho'];
    if (ctx.asset) secs.push('Tài sản cố định');
    if (hasPhotos) secs.push('Hình ảnh');
    const num = name => (secs.indexOf(name) + 1) + '. ' + name;

    cover(doc, R, ctx);
    contents(doc, R, ctx, secs.map((s, i) => (i + 1) + '. ' + s));
    const h1 = '1.  Kiểm tra cửa hàng và tồn kho';
    storePage(doc, R, ctx, h1, [{ label: 'Cửa hàng', hl: false, items: notes.store }, { label: 'Stock', hl: true, items: notes.stock }]);
    if (ctx.stock) stockPages(doc, R, ctx, ctx.stock, h1);
    if (ctx.asset) assetPages(doc, R, ctx, ctx.asset, num('Tài sản cố định'));
    if (hasPhotos) photoPages(doc, R, ctx, num('Hình ảnh'));
    thanks(doc, R);
    return doc.output('blob');
  }

  /* Biên bản gửi Giám đốc chi nhánh: 3 biên bản đúng khổ giấy in
     · Kiểm tra cửa hàng (A4 dọc) · Tài sản cố định (A4 dọc) · Hàng tồn kho (A4 ngang) */
  async function buildForms(ctx) {
    const R = await resources();
    const doc = new root.jspdf.jsPDF({ orientation: 'p', unit: 'pt', format: 'a4', compress: true });
    FONTS.forEach(([f, fam, style], i) => { doc.addFileToVFS(f, R.fonts[i]); doc.addFont(f, fam, style); });
    doc.setProperties({ title: 'Biên bản kiểm kê ' + ctx.store.op, author: ctx.checker, creator: 'Kiểm kê Freshshop' });
    const M = 32;
    // 1. Biên bản kiểm tra cửa hàng — A4 dọc (595 × 842)
    storeForm(doc, R, ctx, M, 48, 595.28 - 2 * M, 9);
    // 2. Tài sản cố định — A4 dọc; trang nối tiếp vẫn dọc
    if (ctx.asset) {
      doc.addPage('a4', 'p');
      assetForm(doc, ctx, ctx.asset, { left: M, right: 595.28 - M, top: 44, contTop: 36, bottom: 812, sigLimit: 822, kf: 0.8,
        newPage: () => { doc.addPage('a4', 'p'); return 60; } });
    }
    // 3. Hàng tồn kho — A4 ngang; ép các trang nối tiếp của bảng cũng ngang
    if (ctx.stock) {
      doc.addPage('a4', 'l');
      const orig = doc.addPage;
      doc.addPage = function (f, o) { return orig.call(this, f || 'a4', o || 'l'); };
      try {
        stockTable(doc, ctx, ctx.stock, { left: 24, right: 841.89 - 24, top: 36, contTop: 30, bottom: 568, sigLimit: 584, kf: 0.8,
          newPage: () => { doc.addPage('a4', 'l'); return 60; } });
      } finally { doc.addPage = orig; }
    }
    return doc.output('blob');
  }

  root.KKReport = { build, buildForms, CHECKLIST, preload: () => resources().catch(() => {}) };
})(self);

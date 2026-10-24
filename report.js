/* سيّارتي — PDF report layout (A4, right-to-left Arabic).
   Pure layout: app.js prepares a "model" of already-formatted strings, this file turns it into a PDF
   with pdf-lib. Arabic letters are shaped by fontkit (OpenType rules of the embedded Readex Pro font)
   and mixed Arabic / numbers / Latin text is ordered with the Unicode bidi algorithm (bidi-js).
   Works in the page (window.PDFLib, window.fontkit, window.bidi_js) and in Node for tests. */
(function (root) {
  'use strict';

  const A4 = [595.28, 841.89];
  const M = { x: 34, top: 34, bottom: 46 };
  const COLORS = {
    ink: [0x12, 0x26, 0x2e], text: [0x14, 0x24, 0x2a], muted: [0x62, 0x74, 0x7a], line: [0xd5, 0xde, 0xdc],
    amber: [0xe9, 0xa2, 0x3b], amberSoft: [0xfb, 0xeb, 0xd0], zebra: [0xf2, 0xf5, 0xf4], white: [0xff, 0xff, 0xff], cream: [0xf3, 0xee, 0xe4]
  };
  const ARABIC = /[؀-ۿݐ-ݿࡰ-ࣿﭐ-﷿ﹰ-﻿]/;

  /* text helpers (separate so the tests can check the visual order on its own) */
  function makeText(bidi, supported) {
  /* ---------- text ----------
     A paragraph's bidi levels are resolved once for the whole text, then it is broken into lines and
     each line is put in visual order (Unicode bidi rules L1–L2), so wrapping never flips brackets. */
  const BIDI_CTRL = /[\u200E\u200F\u2066-\u2069]/;
  const clean = s => [...String(s ?? '').replace(/[\u0000-\u001F\u007F]/g, ' ')].filter(ch => ch === ' ' || BIDI_CTRL.test(ch) || supported.has(ch.codePointAt(0))).join('');
  function para(raw) {
    const text = clean(raw);
    if (!text) return { text, levels: [], mir: new Map() };
    const emb = bidi.getEmbeddingLevels(text, 'rtl');
    return { text, levels: emb.levels, mir: bidi.getMirroredCharactersMap(text, emb.levels) };
  }
  /* visual runs of text[start, end); the font covers the BMP only, so indexes are UTF-16 units */
  function runsOf(p, start = 0, end = p.text.length) {
    let e = end;
    while (e > start && p.text[e - 1] === ' ') e--;            // L1: trailing spaces take the line's base level
    let s0 = start;
    while (s0 < e && p.text[s0] === ' ') s0++;
    const logical = [];
    for (let i = s0; i < e; i++) {
      const c = p.text[i];
      if (BIDI_CTRL.test(c)) continue;
      const ch = p.mir.get(i) || c;
      const lv = p.levels[i];
      const last = logical[logical.length - 1];
      if (last && last.level === lv) last.text += ch; else logical.push({ level: lv, text: ch });
    }
    if (!logical.length) return [];
    const maxLv = Math.max(...logical.map(r => r.level));
    const odd = logical.map(r => r.level).filter(l => l % 2 === 1);
    const minOdd = odd.length ? Math.min(...odd) : maxLv + 1;
    for (let lv = maxLv; lv >= minOdd; lv--) {                  // L2
      let i = 0;
      while (i < logical.length) {
        if (logical[i].level >= lv) {
          let j = i;
          while (j + 1 < logical.length && logical[j + 1].level >= lv) j++;
          const part = logical.slice(i, j + 1).reverse();
          logical.splice(i, part.length, ...part);
          i = j + 1;
        } else i++;
      }
    }
    const out = [];
    for (const r of logical) {
      // left-to-right run; Arabic-Indic digits are Arabic script, which the layout would reverse: pre-reverse them
      if (r.level % 2 === 0) { out.push(ARABIC.test(r.text) ? [...r.text].reverse().join('') : r.text); continue; }
      // right-to-left run: Arabic pieces are put in visual order by the font's layout, the rest by us
      const pieces = [];
      for (const ch of r.text) {
        const ar = ARABIC.test(ch) || (ch === ' ' && pieces.length > 0 && pieces[pieces.length - 1].ar);
        const last = pieces[pieces.length - 1];
        if (last && last.ar === ar) last.text += ch; else pieces.push({ ar, text: ch });
      }
      pieces.reverse();
      for (const pc of pieces) out.push(pc.ar ? pc.text : [...pc.text].reverse().join(''));
    }
    return out;
  }
    return { clean, para, runsOf };
  }

  async function build(model, libs) {
    const { PDFLib, fontkit, bidiFactory, fonts } = libs;
    const { PDFDocument, rgb } = PDFLib;
    const bidi = bidiFactory();
    const col = c => rgb(c[0] / 255, c[1] / 255, c[2] / 255);

    const doc = await PDFDocument.create();
    doc.registerFontkit(fontkit);
    doc.setTitle(model.fileTitle || model.title);
    doc.setAuthor('سيّارتي');
    doc.setCreator('سيّارتي');
    doc.setProducer('سيّارتي');
    doc.setLanguage('ar');
    const reg = await doc.embedFont(fonts.regular, { subset: true });
    const bold = await doc.embedFont(fonts.bold, { subset: true });
    const supported = new Set(reg.getCharacterSet());

    const { clean, para, runsOf } = makeText(bidi, supported);
    const runsWidth = (rs, font, size) => rs.reduce((a, t) => a + font.widthOfTextAtSize(t, size), 0);
    const width = (text, font, size) => runsWidth(runsOf(para(text)), font, size);
    function drawRuns(page, rs, x, y, size, font, color, align) {
      const w = runsWidth(rs, font, size);
      let cx = align === 'right' ? x - w : align === 'center' ? x - w / 2 : x;
      for (const t of rs) {
        if (t) page.drawText(t, { x: cx, y, size, font, color: col(color) });
        cx += font.widthOfTextAtSize(t, size);
      }
      return w;
    }
    /* one line of text; align: 'right' (x = right edge), 'left' (x = left edge), 'center' (x = centre) */
    const line = (page, text, x, y, size, font, color, align = 'right') => drawRuns(page, runsOf(para(text)), x, y, size, font, color, align);
    /* breaks text into lines that fit maxW (between words; inside a word only if it is too long) */
    function wrap(text, font, size, maxW) {
      const p = para(text);
      const W = (s, e) => runsWidth(runsOf(p, s, e), font, size);
      const words = [];
      const re = /[^ ]+/g; let m;
      while ((m = re.exec(p.text))) words.push([m.index, m.index + m[0].length]);
      const lines = [];
      let cur = null;
      for (let [ws, we] of words) {
        if (cur && W(cur[0], we) <= maxW) { cur[1] = we; continue; }
        if (cur) { lines.push(cur); cur = null; }
        while (W(ws, we) > maxW && we - ws > 1) {               // a word wider than the column
          let e = ws + 1;
          while (e < we && W(ws, e + 1) <= maxW) e++;
          lines.push([ws, e]); ws = e;
        }
        cur = [ws, we];
      }
      if (cur) lines.push(cur);
      if (!lines.length) lines.push([0, 0]);
      return lines.map(([s, e]) => runsOf(p, s, e));
    }

    /* ---------- pages ---------- */
    const W = A4[0], H = A4[1], CW = W - M.x * 2, R = W - M.x;
    let page, y;
    const pages = [];
    function newPage(withBand) {
      page = doc.addPage(A4); pages.push(page);
      if (withBand) {
        const bandH = 92;
        page.drawRectangle({ x: 0, y: H - bandH, width: W, height: bandH, color: col(COLORS.ink) });
        page.drawRectangle({ x: 0, y: H - bandH - 4, width: W, height: 4, color: col(model.color || COLORS.amber) });
        line(page, model.kicker, R, H - 28, 10, reg, COLORS.amberSoft);
        line(page, model.title, R, H - 58, 24, bold, COLORS.white);
        line(page, model.subtitle || '', R, H - 78, 10, reg, [0xa9, 0xbc, 0xc1]);
        line(page, model.dateLine, M.x, H - 28, 9, reg, [0xa9, 0xbc, 0xc1], 'left');
        y = H - bandH - 26;
      } else {
        page.drawRectangle({ x: 0, y: H - 6, width: W, height: 6, color: col(model.color || COLORS.amber) });
        line(page, model.title, R, H - M.top + 4, 10, bold, COLORS.muted);
        y = H - M.top - 18;
      }
    }
    const room = h => { if (y - h < M.bottom) { newPage(false); return true; } return false; };

    function heading(text, note) {
      room(96);                                    // keep a heading with the start of its table
      y -= 6;
      page.drawRectangle({ x: R - 4, y: y - 3, width: 4, height: 16, color: col(model.color || COLORS.amber) });
      line(page, text, R - 10, y, 14, bold, COLORS.ink);
      if (note) line(page, note, R - 16 - width(text, bold, 14), y + 1, 9, reg, COLORS.muted);
      y -= 14;
    }

    function tiles(items) {
      if (!items.length) return;
      const per = 3, gap = 8, tw = (CW - gap * (per - 1)) / per, th = 50;
      for (let i = 0; i < items.length; i += per) {
        room(th + gap);
        items.slice(i, i + per).forEach((it, k) => {
          const xR = R - k * (tw + gap);
          page.drawRectangle({ x: xR - tw, y: y - th, width: tw, height: th, color: col(COLORS.zebra), borderColor: col(COLORS.line), borderWidth: 0.5 });
          line(page, it.label, xR - 10, y - 17, 8.5, reg, COLORS.muted);
          const color = it.accent ? [0xb9, 0x45, 0x2f] : COLORS.ink;
          let size = 13;
          while (size > 9.5 && width(it.value, bold, size) > tw - 20) size -= 0.5;
          if (width(it.value, bold, size) <= tw - 20) line(page, it.value, xR - 10, y - 37, size, bold, color);
          else wrap(it.value, bold, 9.5, tw - 20).slice(0, 2).forEach((ln, j) => drawRuns(page, ln, xR - 10, y - 31 - j * 12, 9.5, bold, color, 'right'));
        });
        y -= th + gap;
      }
      y -= 4;
    }

    /* table: columns [{label, align:'right'|'center', weight}] — cells are strings */
    function table(t) {
      const size = t.size || 8.5, lh = size * 1.45, pad = 5;
      const cols = t.columns;
      // natural widths, then fit to the page width
      const nat = cols.map((c, i) => {
        let w = width(clean(c.label), bold, size);
        for (const r of t.rows.slice(0, 400)) w = Math.max(w, Math.min(width(clean(r[i]), reg, size), c.align === 'right' ? 220 : 140));
        if (t.totals && t.totals[i]) w = Math.max(w, width(clean(t.totals[i]), bold, size));
        return w + pad * 2;
      });
      let ws = nat.slice();
      const sum = ws.reduce((a, b) => a + b, 0);
      if (sum <= CW) {
        const flex = cols.map(c => (c.align === 'right' ? 2 : 1));
        const fs = flex.reduce((a, b) => a + b, 0);
        ws = ws.map((w, i) => w + (CW - sum) * flex[i] / fs);
      } else {
        const fixed = cols.map((c, i) => (c.align === 'right' ? 0 : nat[i]));
        const fixedSum = fixed.reduce((a, b) => a + b, 0);
        const textIdx = cols.map((c, i) => (c.align === 'right' ? i : -1)).filter(i => i >= 0);
        if (textIdx.length && fixedSum < CW * 0.8) {
          const left = CW - fixedSum, textNat = textIdx.reduce((a, i) => a + nat[i], 0);
          ws = cols.map((c, i) => (c.align === 'right' ? Math.max(40, left * nat[i] / textNat) : nat[i]));
        } else ws = nat.map(w => w * CW / sum);
        const s2 = ws.reduce((a, b) => a + b, 0); ws = ws.map(w => w * CW / s2);
      }
      const xs = []; let acc = R;
      for (const w of ws) { xs.push(acc); acc -= w; }   // right edge of each column
      const header = () => {
        const hl = cols.map((c, i) => wrap(c.label, bold, size, ws[i] - pad * 2));
        const hh = Math.max(...hl.map(l => l.length)) * lh + pad * 2;
        room(hh + lh * 2 + pad * 2);
        page.drawRectangle({ x: M.x, y: y - hh, width: CW, height: hh, color: col(COLORS.amber) });
        cols.forEach((c, i) => hl[i].forEach((ln, k) => cell(ln, i, y - pad - size - k * lh, bold, [0x1b, 0x13, 0x06], 'center')));
        y -= hh;
      };
      const cell = (rs, i, yy, font, color, align) => {
        if (align === 'center') drawRuns(page, rs, xs[i] - ws[i] / 2, yy, size, font, color, 'center');
        else drawRuns(page, rs, xs[i] - pad, yy, size, font, color, 'right');
      };
      const rowBox = (h, fill) => {
        page.drawRectangle({ x: M.x, y: y - h, width: CW, height: h, color: fill ? col(fill) : undefined, borderColor: col(COLORS.line), borderWidth: 0.4 });
        let x = R; for (let i = 0; i < ws.length - 1; i++) { x -= ws[i]; page.drawLine({ start: { x, y }, end: { x, y: y - h }, thickness: 0.4, color: col(COLORS.line) }); }
      };
      header();
      t.rows.forEach((r, ri) => {
        const ls = cols.map((c, i) => wrap(r[i], reg, size, ws[i] - pad * 2));
        const h = Math.max(...ls.map(l => l.length)) * lh + pad * 2;
        if (room(h)) header();
        rowBox(h, ri % 2 ? COLORS.zebra : null);
        cols.forEach((c, i) => ls[i].forEach((ln, k) => cell(ln, i, y - pad - size - k * lh, reg, i === 0 && t.index ? COLORS.muted : COLORS.text, c.align)));
        y -= h;
      });
      if (!t.rows.length) {
        room(lh + pad * 2);
        rowBox(lh + pad * 2, null);
        line(page, t.empty || '—', W / 2, y - pad - size, size, reg, COLORS.muted, 'center');
        y -= lh + pad * 2;
      }
      if (t.totals) {
        const ls = cols.map((c, i) => wrap(t.totals[i] || '', bold, size, ws[i] - pad * 2));
        const h = Math.max(...ls.map(l => l.length)) * lh + pad * 2;
        room(h);
        rowBox(h, COLORS.amberSoft);
        cols.forEach((c, i) => ls[i].forEach((ln, k) => cell(ln, i, y - pad - size - k * lh, bold, COLORS.ink, c.align)));
        y -= h;
      }
      y -= 14;
    }

    function keyValues(rows) {
      const size = 9.5, lh = size * 1.5, pad = 6, kw = CW * 0.38;
      for (let i = 0; i < rows.length; i++) {
        const [k, v] = rows[i];
        const vl = wrap(v, bold, size, CW - kw - pad * 2), kl = wrap(k, reg, size, kw - pad * 2);
        const h = Math.max(vl.length, kl.length) * lh + pad * 2;
        room(h);
        page.drawRectangle({ x: M.x, y: y - h, width: CW, height: h, color: i % 2 ? undefined : col(COLORS.zebra), borderColor: col(COLORS.line), borderWidth: 0.4 });
        kl.forEach((ln, j) => drawRuns(page, ln, R - pad, y - pad - size - j * lh, size, reg, COLORS.muted, 'right'));
        vl.forEach((ln, j) => drawRuns(page, ln, R - kw - pad, y - pad - size - j * lh, size, bold, COLORS.text, 'right'));
        y -= h;
      }
      y -= 14;
    }

    /* photos: two per row, each with its caption; the JPEG goes into the PDF as it is (full resolution,
       so it can be zoomed in a PDF viewer). A file that is not a valid JPEG gets a placeholder. */
    async function photos(block) {
      const gap = 10, cw = (CW - gap) / 2, size = 8, lh = 11, maxH = 300;
      const cells = [];
      for (const it of block.items || []) {
        let img = null;
        try { img = await doc.embedJpg(it.jpg); } catch (e) { img = null; }
        const k = img ? Math.min((cw - 12) / img.width, maxH / img.height) : 0;
        const w = img ? img.width * k : cw - 12, h = img ? img.height * k : 60;
        const cap = wrap(it.caption || '', reg, size, cw - 12).slice(0, 2);
        cells.push({ img, w, h, cap, ch: 8 + cap.length * lh + 6 + h + 8 });
      }
      if (!cells.length) return;
      room(Math.max(...cells.slice(0, 2).map(p => p.ch)) + gap + 34);   // the heading stays with its first photos
      heading(block.title, block.note);
      for (let i = 0; i < cells.length; i += 2) {
        const pair = cells.slice(i, i + 2);
        const rowH = Math.max(...pair.map(p => p.ch));
        room(rowH + gap);
        pair.forEach((p, k) => {
          const xR = R - k * (cw + gap), xL = xR - cw;
          page.drawRectangle({ x: xL, y: y - rowH, width: cw, height: rowH, color: col(COLORS.zebra), borderColor: col(COLORS.line), borderWidth: 0.5 });
          p.cap.forEach((ln, j) => drawRuns(page, ln, xR - 6, y - 6 - size - j * lh, size, j ? reg : bold, j ? COLORS.muted : COLORS.ink, 'right'));
          const top = y - 8 - p.cap.length * lh - 6;
          if (p.img) {
            page.drawImage(p.img, { x: xL + (cw - p.w) / 2, y: top - p.h, width: p.w, height: p.h });
            page.drawRectangle({ x: xL + (cw - p.w) / 2, y: top - p.h, width: p.w, height: p.h, borderColor: col(COLORS.line), borderWidth: 0.5 });
          } else line(page, 'تعذّر عرض هذه الصورة', xL + cw / 2, top - 34, 9, reg, COLORS.muted, 'center');
        });
        y -= rowH + gap;
      }
      y -= 4;
    }

    /* ---------- the report ---------- */
    newPage(true);
    if (model.notice) {
      const nl = wrap(model.notice, reg, 9, CW - 20);
      const h = nl.length * 13 + 12;
      page.drawRectangle({ x: M.x, y: y - h, width: CW, height: h, color: col(COLORS.amberSoft) });
      nl.forEach((ln, j) => drawRuns(page, ln, R - 10, y - 16 - j * 13, 9, reg, COLORS.ink, 'right'));
      y -= h + 12;
    }
    tiles(model.summary || []);
    for (const block of model.blocks || []) {
      if (block.type === 'kv') { heading(block.title, block.note); keyValues(block.rows); }
      else if (block.type === 'table') { heading(block.title, block.note); table(block); }
      else if (block.type === 'photos') await photos(block);
    }

    /* footer with page numbers */
    pages.forEach((p, i) => {
      p.drawLine({ start: { x: M.x, y: 30 }, end: { x: R, y: 30 }, thickness: 0.5, color: col(COLORS.line) });
      line(p, model.footer || '', R, 18, 8, reg, COLORS.muted);
      line(p, `صفحة ${i + 1} من ${pages.length}`, M.x, 18, 8, reg, COLORS.muted, 'left');
    });
    return doc.save();
  }

  root.SayyaratiReport = { build, _makeText: makeText };
})(typeof window !== 'undefined' ? window : globalThis);

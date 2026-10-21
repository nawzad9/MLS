/**
 * File builders: CSV, Word (.docx) and Excel (.xlsx).
 * Works in the browser (window.Exporters; libraries are loaded on demand) and in Node for tests.
 *
 *   buildDocument(docx, spec)      — Word: letterhead, title, facts strip, tables, signature lines, page numbers
 *   buildWorkbook(ExcelJS, spec)   — Excel: one sheet per subject with live formulas, filters and print setup
 */
(function (root) {
  'use strict';

  const ARABIC_SCRIPT = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;
  const A4 = { width: 11906, height: 16838, margin: 850 };
  const INK = '142A2F', MUTED = '5B6F74', TEAL = '1D6770', RULE = 'AFC0BD', ZEBRA = 'F1F6F5', BENCH = 'EEF3F2';

  /** rows: array of arrays. Adds a BOM so Excel opens Kurdish text correctly. */
  function toCSV(rows) {
    const cell = c => '"' + String(c == null ? '' : c).replace(/"/g, '""') + '"';
    return '﻿' + rows.map(r => r.map(cell).join(',')).join('\r\n');
  }

  /* ═══ Word ═══════════════════════════════════════════════════════════════════
   * spec = { rtl, letterhead: [line…], title, meta: [line…], footer, emptyText,
   *          sections: [{ heading?, note?, pageBreakBefore?, facts?: [[label, value]…],
   *                       columns: [{ label, weight?, align? }], rows: [[cell…]…], signatures?: [{ label, name? }] }] }
   * A cell is a string / number, or { t, color?, bold? }.
   */
  function buildDocument(docx, spec) {
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
            BorderStyle, ShadingType, Footer, PageNumber, VerticalAlign } = docx;
    const rtl = !!spec.rtl;
    const contentWidth = A4.width - 2 * A4.margin;
    const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const line = { style: BorderStyle.SINGLE, size: 4, color: RULE };
    const boxed = { top: line, bottom: line, left: line, right: line };

    const run = (text, o) => {
      o = o || {};
      const s = String(text == null ? '' : text);
      return new TextRun({ text: s, rightToLeft: ARABIC_SCRIPT.test(s), bold: o.bold, boldComplexScript: o.bold,
                           size: o.size, sizeComplexScript: o.size, color: o.color });
    };
    const para = (text, o) => {
      o = o || {};
      return new Paragraph({
        bidirectional: rtl, alignment: o.align, pageBreakBefore: o.pageBreakBefore, keepNext: o.keepNext,
        spacing: { before: o.before || 0, after: o.after == null ? 60 : o.after },
        border: o.rule ? { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 6 } } : undefined,
        children: [run(text, o)]
      });
    };
    const split = (weights, width) => {
      const total = weights.reduce((a, b) => a + b, 0);
      const w = weights.map(x => Math.floor(width * x / total));
      w[w.length - 1] += width - w.reduce((a, b) => a + b, 0);
      return w;
    };

    const dataTable = section => {
      const widths = split(section.columns.map(c => c.weight || 1), contentWidth);
      const cell = (value, i, kind) => {
        const c = value && typeof value === 'object' ? value : { t: value };
        const head = kind === 'head';
        return new TableCell({
          width: { size: widths[i], type: WidthType.DXA }, borders: boxed, verticalAlign: VerticalAlign.CENTER,
          shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: TEAL }
                 : kind === 'zebra' ? { type: ShadingType.CLEAR, color: 'auto', fill: ZEBRA } : undefined,
          margins: { top: 50, bottom: 50, left: 80, right: 80 },
          children: [new Paragraph({
            bidirectional: rtl, alignment: section.columns[i].align === 'center' ? AlignmentType.CENTER : undefined,
            children: [run(c.t, head ? { bold: true, color: 'FFFFFF', size: 18 } : { size: 19, color: c.color || INK, bold: c.bold })]
          })]
        });
      };
      return new Table({
        width: { size: contentWidth, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: rtl,
        rows: [
          new TableRow({ tableHeader: true, cantSplit: true, children: section.columns.map((c, i) => cell(c.label, i, 'head')) }),
          ...section.rows.map((r, n) => new TableRow({ cantSplit: true, children: r.map((v, i) => cell(v, i, n % 2 ? 'zebra' : 'body')) }))
        ]
      });
    };

    const factsTable = facts => {
      const widths = split(facts.map(() => 1), contentWidth);
      const p = (text, o) => new Paragraph({ bidirectional: rtl, alignment: AlignmentType.CENTER, children: [run(text, o)] });
      const cell = (child, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA }, borders: boxed, margins: { top: 60, bottom: 60, left: 90, right: 90 },
        shading: { type: ShadingType.CLEAR, color: 'auto', fill: BENCH }, children: [child]
      });
      return new Table({
        width: { size: contentWidth, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: rtl,
        rows: [
          new TableRow({ cantSplit: true, children: facts.map((f, i) => cell(p(f[0], { size: 16, color: MUTED }), i)) }),
          new TableRow({ cantSplit: true, children: facts.map((f, i) => cell(p(f[1], { size: 26, bold: true, color: TEAL }), i)) })
        ]
      });
    };

    /** Signature lines side by side, with a small gap between them. */
    const signatures = list => {
      const GAP = 400, n = list.length;
      const each = Math.floor((contentWidth - GAP * (n - 1)) / n);
      const widths = [];
      list.forEach((_, i) => { if (i) widths.push(GAP); widths.push(each); });
      widths[widths.length - 1] += contentWidth - widths.reduce((a, b) => a + b, 0);
      const noBorders = { top: none, bottom: none, left: none, right: none };
      const cells = [];
      list.forEach((s, i) => {
        if (i) cells.push(new TableCell({ width: { size: GAP, type: WidthType.DXA }, borders: noBorders, children: [new Paragraph('')] }));
        cells.push(new TableCell({
          width: { size: each, type: WidthType.DXA },
          borders: { top: { style: BorderStyle.SINGLE, size: 6, color: INK }, bottom: none, left: none, right: none },
          margins: { top: 60 },
          children: [
            new Paragraph({ bidirectional: rtl, alignment: AlignmentType.CENTER, children: [run(s.label, { size: 18, bold: true, color: INK })] }),
            new Paragraph({ bidirectional: rtl, alignment: AlignmentType.CENTER, children: [run(s.name || ' ', { size: 17, color: MUTED })] })
          ]
        }));
      });
      return new Table({ width: { size: contentWidth, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: rtl,
                         rows: [new TableRow({ cantSplit: true, children: cells })] });
    };

    const children = [];
    (spec.letterhead || []).forEach((l, i) =>
      children.push(para(l, { align: AlignmentType.CENTER, bold: i === 0, size: i === 0 ? 24 : 20, after: 0, color: INK })));
    children.push(para(spec.title, { align: AlignmentType.CENTER, bold: true, size: 30, before: 240, after: 60, color: TEAL }));
    (spec.meta || []).forEach(m => children.push(para(m, { align: AlignmentType.CENTER, size: 19, color: MUTED, after: 0 })));
    children.push(para('', { after: 120, rule: true }));

    spec.sections.forEach(section => {
      if (section.heading) children.push(para(section.heading, { bold: true, size: 24, before: 200, after: 40, color: INK, pageBreakBefore: section.pageBreakBefore, keepNext: true }));
      if (section.note) children.push(para(section.note, { size: 18, color: MUTED, after: 80, keepNext: true }));
      if (section.facts && section.facts.length) { children.push(factsTable(section.facts)); children.push(para('', { after: 100 })); }
      children.push(section.rows.length ? dataTable(section) : para(section.emptyText || spec.emptyText || '—', { size: 19, color: MUTED }));
      if (section.signatures && section.signatures.length) { children.push(para('', { before: 560, after: 0 })); children.push(signatures(section.signatures)); }
    });

    return new Document({
      creator: 'MLS Registration — University of Raparin', title: spec.title,
      styles: { default: { document: { run: { font: { ascii: 'Calibri', hAnsi: 'Calibri', cs: 'Arial' }, size: 21, sizeComplexScript: 21 } } } },
      sections: [{
        properties: { page: { size: { width: A4.width, height: A4.height }, margin: { top: A4.margin, bottom: A4.margin, left: A4.margin, right: A4.margin } } },
        footers: { default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.CENTER,   // left-to-right so "1 / 3" is not mirrored
          children: [run((spec.footer || '') + '   ', { size: 16, color: MUTED }),
                     new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], size: 16, color: MUTED })]
        })] }) },
        children: children
      }]
    });
  }

  /* ═══ Excel ══════════════════════════════════════════════════════════════════
   * spec = { rtl, title, labels, summary?, sheets: [{ name, letterhead, title, meta, limits: { maxA, maxB, pass },
   *          rows: [{ n, name, code, group, type, A, B }], signatures: [{ label, name }] }] }
   * Totals, results and statistics are real formulas, so they stay correct when a mark is edited in Excel.
   */
  function buildWorkbook(ExcelJS, spec) {
    const rtl = !!spec.rtl, L = spec.labels;
    const FONT = 'Arial';
    const argb = hex => 'FF' + hex;
    const thin = { style: 'thin', color: { argb: argb(RULE) } };
    const box = { top: thin, left: thin, bottom: thin, right: thin };
    const wb = new ExcelJS.Workbook();
    wb.creator = 'MLS Registration — University of Raparin';
    wb.title = spec.title;
    wb.created = new Date();
    const used = new Set();
    const sheetName = name => {
      const base = String(name).replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 28) || 'Sheet';
      let n = base, i = 2;
      while (used.has(n.toLowerCase())) n = base.slice(0, 26) + ' ' + i++;
      used.add(n.toLowerCase());
      return n;
    };
    const align = (h, extra) => Object.assign({ horizontal: h, vertical: 'middle', readingOrder: rtl ? 'rtl' : 'ltr' }, extra);
    const textAlign = rtl ? 'right' : 'left';

    const pageSetup = (ws, headerRow) => {
      ws.views = [{ rightToLeft: rtl, showGridLines: false, state: 'frozen', ySplit: headerRow }];
      ws.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
                       printTitlesRow: headerRow + ':' + headerRow, margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.7, header: 0.3, footer: 0.3 } };
      ws.headerFooter.oddFooter = '&C&"Arial"&8&P / &N';
    };
    const banner = (ws, ncols, o) => {
      let r = 1;
      const line = (text, font, height) => {
        ws.mergeCells(r, 1, r, ncols);
        const c = ws.getCell(r, 1);
        c.value = text; c.font = Object.assign({ name: FONT }, font); c.alignment = align('center');
        ws.getRow(r++).height = height;
      };
      (o.letterhead || []).forEach((l, i) => line(l, { size: i === 0 ? 13 : 11, bold: i === 0, color: { argb: argb(INK) } }, i === 0 ? 20 : 17));
      line(o.title, { size: 16, bold: true, color: { argb: argb(TEAL) } }, 30);
      (o.meta || []).forEach(m => line(m, { size: 10.5, color: { argb: argb(MUTED) } }, 17));
      return r + 1;                       // one blank row, then the table header
    };
    const headerCell = (cell, text) => {
      cell.value = text; cell.font = { name: FONT, size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(TEAL) } };
      cell.alignment = align('center', { wrapText: true }); cell.border = box;
    };
    const bodyCell = (cell, value, o) => {
      o = o || {};
      cell.value = value;
      cell.font = { name: FONT, size: 11, bold: o.bold, color: { argb: argb(o.color || INK) } };
      cell.alignment = align(o.center ? 'center' : (o.ltr ? 'left' : textAlign), { wrapText: !o.center });
      cell.border = box;
      if (o.zebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(ZEBRA) } };
      if (o.fmt) cell.numFmt = o.fmt;
    };

    if (spec.summary) {                    // overview sheet when several subjects are exported together
      const S = spec.summary, ws = wb.addWorksheet(sheetName(S.name));
      const n = S.columns.length;
      ws.columns = S.columns.map(c => ({ width: c.width || 14 }));
      const hr = banner(ws, n, S);
      S.columns.forEach((c, i) => headerCell(ws.getCell(hr, i + 1), c.label));
      ws.getRow(hr).height = 32;
      S.rows.forEach((row, k) => {
        row.forEach((v, i) => bodyCell(ws.getCell(hr + 1 + k, i + 1), v == null ? '' : v,
          { zebra: k % 2, center: S.columns[i].center, ltr: S.columns[i].ltr, fmt: S.columns[i].fmt }));
        ws.getRow(hr + 1 + k).height = 22;
      });
      if (S.rows.length) ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: hr + S.rows.length, column: n } };
      pageSetup(ws, hr);
    }

    spec.sheets.forEach(sh => {
      const ws = wb.addWorksheet(sheetName(sh.name));
      const N = 9, lim = sh.limits, pass = lim.pass;
      ws.columns = [5, 34, 13, 11, 13, 12, 12, 11, 12].map(width => ({ width }));
      const hr = banner(ws, N, sh);
      [L.hash, L.name, L.code, L.group, L.type, `${L.a}\n(${lim.maxA})`, `${L.b}\n(${lim.maxB})`, `${L.total}\n(${lim.maxA + lim.maxB})`, L.result]
        .forEach((c, i) => headerCell(ws.getCell(hr, i + 1), c));
      ws.getRow(hr).height = 34;

      const first = hr + 1, last = hr + sh.rows.length;
      sh.rows.forEach((s, k) => {
        const r = first + k, z = k % 2 === 1;
        const total = s.A != null && s.B != null ? Math.round((s.A + s.B) * 100) / 100 : null;
        bodyCell(ws.getCell(r, 1), s.n, { center: true, zebra: z, color: MUTED });
        bodyCell(ws.getCell(r, 2), s.name, { zebra: z, bold: true });
        bodyCell(ws.getCell(r, 3), s.code, { ltr: true, zebra: z });
        bodyCell(ws.getCell(r, 4), s.group, { center: true, zebra: z });
        bodyCell(ws.getCell(r, 5), s.type, { center: true, zebra: z });
        bodyCell(ws.getCell(r, 6), s.A == null ? null : s.A, { center: true, zebra: z, bold: true });
        bodyCell(ws.getCell(r, 7), s.B == null ? null : s.B, { center: true, zebra: z, bold: true });
        bodyCell(ws.getCell(r, 8), { formula: `IF(AND(ISNUMBER(F${r}),ISNUMBER(G${r})),F${r}+G${r},"")`, result: total == null ? '' : total },
                 { center: true, zebra: z, bold: true });
        bodyCell(ws.getCell(r, 9), { formula: `IF(ISNUMBER(H${r}),IF(H${r}>=${pass},"${L.pass}","${L.fail}"),"")`,
                 result: total == null ? '' : (total >= pass ? L.pass : L.fail) }, { center: true, zebra: z, bold: true });
        [[6, lim.maxA], [7, lim.maxB]].forEach(([c, max]) => {
          ws.getCell(r, c).dataValidation = { type: 'decimal', operator: 'between', allowBlank: true, showErrorMessage: true, errorStyle: 'stop',
                                              errorTitle: L.invalid, error: L.range(max), formulae: [0, max] };
        });
        ws.getRow(r).height = 22;
      });

      if (sh.rows.length) {
        ws.addConditionalFormatting({ ref: `H${first}:I${last}`, rules: [
          { type: 'expression', priority: 1, formulae: [`AND(ISNUMBER($H${first}),$H${first}<${pass})`], style: { font: { bold: true, color: { argb: 'FFB0271F' } } } },
          { type: 'expression', priority: 2, formulae: [`AND(ISNUMBER($H${first}),$H${first}>=${pass})`], style: { font: { bold: true, color: { argb: 'FF2B7A4B' } } } }
        ] });
        ws.autoFilter = { from: { row: hr, column: 1 }, to: { row: last, column: N } };
      }

      let r = last + 2;
      if (sh.rows.length) {                // statistics as formulas over the table
        const totals = sh.rows.map(s => (s.A != null && s.B != null ? s.A + s.B : null)).filter(v => v != null);
        const passed = totals.filter(v => v >= pass).length;
        const H = `H${first}:H${last}`, I = `I${first}:I${last}`;
        const r2 = v => Math.round(v * 100) / 100;
        [
          [L.stats.students, `COUNTA(B${first}:B${last})`, sh.rows.length],
          [L.stats.graded, `COUNT(${H})`, totals.length],
          [L.stats.average, `IF(COUNT(${H})=0,"",ROUND(AVERAGE(${H}),2))`, totals.length ? r2(totals.reduce((a, b) => a + b, 0) / totals.length) : ''],
          [L.stats.highest, `IF(COUNT(${H})=0,"",MAX(${H}))`, totals.length ? Math.max(...totals) : ''],
          [L.stats.lowest, `IF(COUNT(${H})=0,"",MIN(${H}))`, totals.length ? Math.min(...totals) : ''],
          [L.stats.passed, `COUNTIF(${I},"${L.pass}")`, passed],
          [L.stats.failed, `COUNTIF(${I},"${L.fail}")`, totals.length - passed],
          [L.stats.rate, `IF(COUNT(${H})=0,"",COUNTIF(${I},"${L.pass}")/COUNT(${H}))`, totals.length ? passed / totals.length : '', '0%']
        ].forEach(([label, formula, result, fmt]) => {
          ws.mergeCells(r, 2, r, 3);
          const a = ws.getCell(r, 2);
          a.value = label; a.font = { name: FONT, size: 10.5, color: { argb: argb(MUTED) } }; a.alignment = align(textAlign); a.border = box;
          ws.getCell(r, 3).border = box;
          const b = ws.getCell(r, 4);
          b.value = { formula, result }; b.font = { name: FONT, size: 11, bold: true, color: { argb: argb(TEAL) } }; b.alignment = align('center'); b.border = box;
          if (fmt) b.numFmt = fmt;
          ws.getRow(r++).height = 20;
        });
      }

      if (sh.signatures && sh.signatures.length) {
        r += 2;
        const span = Math.floor(N / sh.signatures.length);
        const top = { top: { style: 'thin', color: { argb: argb(INK) } } };
        sh.signatures.forEach((sg, i) => {
          const c1 = 1 + i * span, c2 = i === sh.signatures.length - 1 ? N : c1 + span - 1;
          ws.mergeCells(r, c1, r, c2);
          ws.mergeCells(r + 1, c1, r + 1, c2);
          for (let c = c1; c <= c2; c++) ws.getCell(r, c).border = top;
          const a = ws.getCell(r, c1);
          a.value = sg.label; a.font = { name: FONT, size: 10.5, bold: true, color: { argb: argb(INK) } }; a.alignment = align('center');
          const b = ws.getCell(r + 1, c1);
          b.value = sg.name || ''; b.font = { name: FONT, size: 10, color: { argb: argb(MUTED) } }; b.alignment = align('center');
        });
        ws.getRow(r).height = 20;
      }
      pageSetup(ws, hr);
    });
    return wb;
  }

  const api = { toCSV: toCSV, buildDocument: buildDocument, buildWorkbook: buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Exporters = api;
})(typeof self !== 'undefined' ? self : this);

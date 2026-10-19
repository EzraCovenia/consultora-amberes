/* ============================================================
   AMBERES INTELLIGENCE™ — Generador de informe PDF
   PDF nativo (jsPDF + autoTable): texto y tablas vectoriales, gráficos como imagen.
   Recorre los tableros ya renderizados en el DOM, de modo que el PDF refleja lo que ve el usuario.
   (No usa html2canvas: el DOM completo es demasiado alto y producía páginas en blanco.)
   ============================================================ */
(function () {
  'use strict';
  const NAVY = [16, 42, 54], GOLD = [200, 155, 60], TXT = [38, 48, 54], GRAY = [112, 124, 130], PAPER = [246, 244, 238];
  const SEV = { danger: [192, 57, 57], warning: [196, 135, 22], success: [46, 125, 87], info: [79, 124, 130] };
  const SEV_LABEL = { danger: 'CRITICO', warning: 'ATENCION', success: 'FAVORABLE', info: 'NOTA' };
  const clean = s => String(s == null ? '' : s)
    .replace(/[−–—]/g, '-').replace(/≈/g, '~').replace(/Δ/g, 'Var.').replace(/→/g, '->').replace(/…/g, '...')
    .replace(/[✓✔⚠]/g, '').replace(/[^\x09\x0A\x0D\x20-\x7E -ÿ]/g, '').replace(/\s+/g, ' ').trim();
  const visible = n => !(n.style && n.style.display === 'none') && !(n.hidden);
  const has = (n, c) => n.classList && n.classList.contains(c);
  const T = (n, sel) => { const e = n.querySelector(sel); return e ? clean(e.textContent) : ''; };

  function generate(opt) {
    const { jsPDF } = window.jspdf;
    if (!jsPDF || !(window.jspdf && (new jsPDF()).autoTable)) throw new Error('Librerías de PDF no disponibles');
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const PW = 297, PH = 210, MX = 12, TOP = 20, BOT = 14, W = PW - 2 * MX;
    let y = TOP, section = '';

    const newPage = () => { doc.addPage(); y = TOP; };
    const ensure = h => { if (y + h > PH - BOT) newPage(); };
    const setC = (c, kind) => (kind === 'fill' ? doc.setFillColor(c[0], c[1], c[2]) : kind === 'draw' ? doc.setDrawColor(c[0], c[1], c[2]) : doc.setTextColor(c[0], c[1], c[2]));

    function heading(text) {
      ensure(34); doc.setFont('helvetica', 'bold'); doc.setFontSize(11); setC(NAVY); doc.text(clean(text), MX, y + 4);
      setC(GOLD, 'draw'); doc.setLineWidth(0.5); doc.line(MX, y + 6, MX + 18, y + 6); y += 10;
    }
    function note(text) {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); setC(GRAY);
      const lines = doc.splitTextToSize(clean(text), W); ensure(lines.length * 3.4 + 2); doc.text(lines, MX, y + 3); y += lines.length * 3.4 + 3;
    }
    function cards(node, labelSel, valueSel, subSel) {
      const list = Array.from(node.children); if (!list.length) return;
      const per = 4, gap = 3, cw = (W - gap * (per - 1)) / per, ch = 22;
      for (let i = 0; i < list.length; i += per) {
        ensure(ch + 3);
        list.slice(i, i + per).forEach((c, j) => {
          const x = MX + j * (cw + gap);
          const tone = /danger/.test(c.className) ? SEV.danger : /warning/.test(c.className) ? SEV.warning : /success/.test(c.className) ? SEV.success : GOLD;
          setC(PAPER, 'fill'); doc.rect(x, y, cw, ch, 'F'); setC(tone, 'fill'); doc.rect(x, y, cw, 1, 'F');
          doc.setFont('helvetica', 'bold'); doc.setFontSize(6.5); setC(GRAY); doc.text(doc.splitTextToSize(T(c, labelSel).toUpperCase(), cw - 4).slice(0, 2), x + 2.5, y + 4.5);
          doc.setFontSize(13); setC(NAVY); doc.text(doc.splitTextToSize(T(c, valueSel), cw - 4)[0] || '', x + 2.5, y + 12);
          doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); setC(TXT); doc.text(doc.splitTextToSize(T(c, subSel), cw - 4).slice(0, 2), x + 2.5, y + 16.5);
        });
        y += ch + 3;
      }
    }
    function tile(canvas, title) {
      if (!canvas || !canvas.width || !canvas.height) return null;
      const w = canvas.width, h = canvas.height, tt = Math.round(h * 0.13);
      const c = document.createElement('canvas'); c.width = w; c.height = h + tt;
      const x = c.getContext('2d'); x.fillStyle = '#15313e'; x.fillRect(0, 0, c.width, c.height);
      x.fillStyle = '#f4efe5'; x.font = '600 ' + Math.round(tt * 0.42) + 'px Arial, sans-serif'; x.fillText(clean(title).toUpperCase(), Math.round(w * 0.03), Math.round(tt * 0.64));
      x.drawImage(canvas, 0, tt);
      return { url: c.toDataURL('image/jpeg', 0.88), ratio: c.height / c.width };
    }
    function charts(list) {
      const per = 2, gap = 4, cw = (W - gap) / per;
      const items = list.map(n => ({ t: tile(n.querySelector('canvas'), T(n, '.chart-title')) })).filter(i => i.t);
      for (let i = 0; i < items.length; i += per) {
        const row = items.slice(i, i + per), hgt = Math.min(Math.max.apply(null, row.map(r => r.t.ratio * cw)), 88);
        ensure(hgt + 4);
        row.forEach((r, j) => doc.addImage(r.t.url, 'JPEG', MX + j * (cw + gap), y, cw, Math.min(r.t.ratio * cw, hgt)));
        y += hgt + 4;
      }
    }
    function tableEl(t) {
      const cols = t.querySelectorAll('thead th').length; const fs = cols > 16 ? 5.2 : cols > 10 ? 6 : 7.2;
      doc.autoTable({
        html: t, startY: y, margin: { left: MX, right: MX, top: TOP, bottom: BOT }, theme: 'grid',
        styles: { fontSize: fs, cellPadding: cols > 10 ? 0.9 : 1.4, textColor: TXT, lineColor: [225, 222, 213], lineWidth: 0.1, overflow: 'linebreak' },
        headStyles: { fillColor: NAVY, textColor: [255, 255, 255], fontStyle: 'bold', halign: 'right' },
        alternateRowStyles: { fillColor: PAPER },
        didParseCell: d => {
          d.cell.text = d.cell.text.map(clean);
          if (d.column.index === 0) d.cell.styles.halign = 'left'; else d.cell.styles.halign = 'right';
          if (d.section === 'body') {
            const tr = d.row.raw && d.row.raw.classList ? d.row.raw : (d.row.raw && d.row.raw._element) || null;
            if (tr && (tr.classList.contains('tot') || tr.classList.contains('ma-total'))) { d.cell.styles.fontStyle = 'bold'; d.cell.styles.fillColor = [236, 228, 208]; }
            const td = d.cell.raw; if (td && td.classList && td.classList.contains('neg')) d.cell.styles.textColor = SEV.danger;
            if (td && td.classList && td.classList.contains('pos')) d.cell.styles.textColor = SEV.success;
          }
        }
      });
      y = doc.lastAutoTable.finalY + 4;
    }
    function insight(item) {
      const sev = (item.className.match(/\b(danger|warning|success|info)\b/) || [, 'info'])[1];
      const title = T(item, 'b'), body = (() => { const d = item.querySelector('div'); return d ? clean(d.textContent) : ''; })();
      const plain = !title && !body ? clean(item.textContent) : '';
      doc.setFontSize(7.6); doc.setFont('helvetica', 'normal');
      const lines = doc.splitTextToSize(body || plain, W - 8), hgt = (title ? 5 : 1) + lines.length * 3.5 + 3;
      ensure(hgt + 2);
      setC(PAPER, 'fill'); doc.rect(MX, y, W, hgt, 'F'); setC(SEV[sev], 'fill'); doc.rect(MX, y, 1.2, hgt, 'F');
      let ty = y + 4;
      if (title) {
        doc.setFont('helvetica', 'bold'); doc.setFontSize(6.5); setC(SEV[sev]); doc.text(SEV_LABEL[sev] || 'NOTA', MX + 4, ty);
        doc.setFontSize(8.2); setC(NAVY); doc.text(doc.splitTextToSize(title, W - 34)[0], MX + 24, ty); ty += 4.3;
      }
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.6); setC(TXT); doc.text(lines, MX + 4, ty);
      y += hgt + 2;
    }
    function emit(node) {
      if (!node || node.nodeType !== 1 || !visible(node)) return;
      const tag = node.tagName;
      if (/^(SCRIPT|STYLE|BUTTON|INPUT|SELECT|LABEL)$/.test(tag)) return;
      if (['selector-row', 'btn-row', 'sensitivity-panel', 'ma-panel', 'assump', 'ma-note', 'chart-title'].some(c => has(node, c))) return;
      if (has(node, 'metrics-grid')) return cards(node, '.metric-label', '.metric-value', '.metric-sub');
      if (has(node, 'ma-results-grid')) return cards(node, '.ma-result-label', '.ma-result-value', '.ma-result-sub');
      if (has(node, 'chart-grid')) return charts(Array.from(node.children).filter(c => has(c, 'chart-container')));
      if (has(node, 'chart-container')) return charts([node]);
      if (has(node, 'tbl-wrap') || has(node, 'ma-table-wrap')) { const t = node.querySelector('table'); if (t) tableEl(t); return; }
      if (has(node, 'section-h')) return heading(node.textContent);
      if (has(node, 'sub-note')) return note(node.textContent);
      if (has(node, 'insights-section')) {
        const ttl = node.querySelector('.insight-title'); if (ttl) heading(ttl.textContent);
        node.querySelectorAll('.insight-item').forEach(insight); y += 2; return;
      }
      if (has(node, 'insight-item')) return insight(node);
      if (has(node, 'empty-state')) return note(node.textContent);
      Array.from(node.children).forEach(emit);
    }

    /* ---- portada ---- */
    setC(NAVY, 'fill'); doc.rect(0, 0, PW, 70, 'F'); setC(GOLD, 'fill'); doc.rect(0, 70, PW, 1.4, 'F');
    let lx = MX;
    if (opt.logo) { try { setC([255, 255, 255], 'fill'); doc.roundedRect(MX, 14, 40, 40, 3, 3, 'F'); doc.addImage(opt.logo.url, 'PNG', MX + 2, 14 + (40 - 36 * opt.logo.ratio) / 2, 36, 36 * opt.logo.ratio); lx = MX + 48; } catch (e) { /* sin logo */ } }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setC(GOLD); doc.text('AMBERES CONSULTORA  |  FINANCIAL INTELLIGENCE', lx, 24);
    doc.setFontSize(26); setC([244, 239, 229]); doc.text('Informe financiero integrado', lx, 38);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(11); setC([185, 199, 201]); doc.text('P&L, flujo de caja, balance, evaluacion de proyectos y valoracion M&A', lx, 47);
    doc.setFontSize(9); doc.text('Generado el ' + clean(opt.date) + (opt.files && opt.files.length ? '  |  Archivos: ' + clean(opt.files.join(', ')) : ''), lx, 57);
    y = 84; heading('Origen de la informacion');
    doc.autoTable({ startY: y, margin: { left: MX, right: MX }, theme: 'grid', head: [['Modulo', 'Origen', 'Detalle']], body: (opt.sources || []).map(r => r.map(clean)),
      styles: { fontSize: 8, cellPadding: 1.8, textColor: TXT, lineColor: [225, 222, 213], lineWidth: 0.1 }, headStyles: { fillColor: NAVY, textColor: 255 }, alternateRowStyles: { fillColor: PAPER } });
    y = doc.lastAutoTable.finalY + 6;
    note('Cargado del archivo = datos que vienen de tu Excel. Generado automaticamente = estimado por el software con supuestos editables a partir de otros informes; no todas las cuentas se pueden traspasar de un informe a otro, por lo que deben revisarse.');
    y += 2; heading('Contenido');
    note((opt.panes || []).map(p => clean(p.title)).join('   |   '));

    /* ---- secciones ---- */
    (opt.panes || []).forEach(p => {
      const el = document.getElementById(p.id); if (!el || !el.textContent.trim()) return;
      newPage(); section = p.title;
      setC(NAVY, 'fill'); doc.rect(MX, y - 2, W, 9, 'F'); doc.setFont('helvetica', 'bold'); doc.setFontSize(12); setC([244, 239, 229]); doc.text(clean(p.title), MX + 3, y + 4.2);
      y += 13; Array.from(el.children).forEach(emit);
    });

    /* ---- encabezados y pies ---- */
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      if (i > 1) { setC(NAVY); doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.text('AMBERES INTELLIGENCE', MX, 11); setC(GOLD, 'draw'); doc.setLineWidth(0.4); doc.line(MX, 13.5, PW - MX, 13.5); }
      setC(GRAY); doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5);
      doc.text('Informe referencial generado en tu navegador; debe ser validado por un profesional antes de tomar decisiones de inversion, financiamiento o transacciones. Lo marcado como estimado proviene de supuestos del software.', MX, PH - 6, { maxWidth: W - 24 });
      doc.text(i + ' / ' + n, PW - MX, PH - 6, { align: 'right' });
    }
    return doc;
  }
  window.AmberesPDF = { generate, clean };
})();

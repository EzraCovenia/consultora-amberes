/* ============================================================
   AMBERES INTELLIGENCE™ — Interfaz
   Lee el libro (todas las pestañas), alimenta el motor (model.js / model-engine.js)
   y dibuja los tableros. Todo texto de usuario se inserta con textContent.
   ============================================================ */
(function () {
  'use strict';
  const AM = window.AmberesModel;
  const $ = id => document.getElementById(id);
  const fm = AM.fmt;
  const pc = (v, d) => (v == null || !isFinite(v)) ? '—' : (v * 100).toFixed(d == null ? 1 : d) + '%';
  const num = (v, d) => (v == null || !isFinite(v)) ? '—' : (Math.abs(v) < 0.5 && !d ? '0' : Number(v).toLocaleString('es-CL', { maximumFractionDigits: d == null ? 0 : d }));
  const COL = { navy: '#102a36', gold: '#c89b3c', goldL: '#e4c77a', petrol: '#4f7c82', cream: '#f4efe5', green: '#4ea878', red: '#d86767', amber: '#e4c77a', dim: '#b9c7c9', grid: 'rgba(255,255,255,.08)' };
  const SEV_LABEL = { danger: 'Crítico', warning: 'Atención', success: 'Favorable', info: 'Nota' };

  const store = AM.newStore();
  let Ain = {};            // supuestos fijados por el usuario
  let M = null;
  let maOver = {};         // ajustes manuales de M&A
  let charts = {};
  let tab = 'overview';
  let plEmpresa = null;

  /* ---------- utilidades de DOM seguras ---------- */
  function h(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k]; if (v == null || v === false) continue;
      if (k === 'class') e.className = v; else if (k === 'text') e.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v);
    }
    kids.flat().forEach(c => { if (c == null || c === false) return; e.appendChild(c.nodeType ? c : document.createTextNode(String(c))); });
    return e;
  }
  const clear = el => { while (el.firstChild) el.removeChild(el.firstChild); return el; };
  function chartDefaults() {
    try {
      if (!window.Chart) return;
      Chart.defaults.color = COL.dim; Chart.defaults.borderColor = COL.grid;
      Chart.defaults.font = Object.assign(Chart.defaults.font || {}, { family: 'Inter, sans-serif', size: 11 });
    } catch (e) { /* defaults opcionales */ }
  }
  function mk(canvasId, cfg) {
    if (!window.Chart) return null;
    const el = $(canvasId); if (!el) return null;
    if (charts[canvasId]) charts[canvasId].destroy();
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 350 }, plugins: {} }, cfg.options || {});
    cfg.options.plugins = Object.assign({ legend: { labels: { color: COL.dim, boxWidth: 12 } } }, cfg.options.plugins);
    charts[canvasId] = new Chart(el, cfg); return charts[canvasId];
  }
  function chartBox(id, title, hgt) { return h('div', { class: 'chart-container' }, h('div', { class: 'chart-title', text: title }), h('div', { style: 'height:' + (hgt || 300) + 'px' }, h('canvas', { id }))); }
  const yAxisM = { ticks: { color: COL.dim, callback: v => fm(v) }, grid: { color: COL.grid } };
  const xAxis = { ticks: { color: COL.dim }, grid: { display: false } };

  function metric(k) {
    return h('div', { class: 'metric-card ' + (k.cls || '') },
      h('div', { class: 'metric-label', text: k.label }),
      h('div', { class: 'metric-value' }, k.value, k.sem ? h('span', { class: 'semaforo ' + k.sem, title: k.sem }) : null),
      h('div', { class: 'metric-sub', text: k.sub || '' }));
  }
  function insightList(list, title) {
    const box = h('div', { class: 'insights-section rec-list' }, h('div', { class: 'insight-title', text: title || '💡 Insights y recomendaciones' }));
    if (!list || !list.length) box.appendChild(h('div', { class: 'insight-item info', text: 'Sin hallazgos relevantes con los datos actuales.' }));
    (list || []).forEach(i => box.appendChild(h('div', { class: 'insight-item ' + i.level }, h('span', { class: 'sev ' + i.level, text: SEV_LABEL[i.level] || 'Nota' }), h('b', { text: i.title }), h('div', { text: i.text }))));
    return box;
  }
  function table(head, rows, opt) {
    opt = opt || {};
    const t = h('table', null, h('thead', null, h('tr', null, head.map(x => h('th', { text: x })))),
      h('tbody', null, rows.map(r => {
        const cls = r.cls || ''; const cells = r.cells || r;
        return h('tr', { class: cls }, cells.map((c, i) => {
          const td = h('td');
          if (c && c.nodeType) td.appendChild(c);
          else { td.textContent = c == null ? '' : String(c); }
          if (r.neg && i > 0 && r.neg[i - 1]) td.className = 'neg';
          return td;
        }));
      })));
    return h('div', { class: 'tbl-wrap' }, t);
  }
  const tagEst = () => h('span', { class: 'tag-est', text: 'ESTIMADO' });
  const srcTag = (state) => h('span', { class: 'tag-src ' + (state === 'cargado' ? 'file' : ''), text: state === 'cargado' ? 'DEL ARCHIVO' : state === 'derivado' ? 'DERIVADO' : 'AUSENTE' });

  /* ---------- carga de archivos ---------- */
  function setMsg(t, err) { const m = $('loadMsg'); m.textContent = t || ''; m.className = 'load-msg' + (err ? ' err' : ''); }
  async function ingestBuffer(name, buf) {
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true });
    AM.ingest(store, name, AM.sheetsFromWorkbook(XLSX, wb));
  }
  async function loadFiles(files) {
    const list = Array.from(files || []); if (!list.length) return;
    $('processingOverlay').classList.add('active'); setMsg('');
    try {
      for (const f of list) {
        if (!/\.(xlsx|xls)$/i.test(f.name)) { setMsg('"' + f.name + '" no es un archivo Excel (.xlsx / .xls).', true); continue; }
        await ingestBuffer(f.name, await f.arrayBuffer());
      }
      onLoaded();
    } catch (e) { console.error(e); setMsg('No se pudo leer el archivo: ' + (e && e.message ? e.message : 'formato no válido') + '.', true); }
    finally { $('processingOverlay').classList.remove('active'); }
  }
  async function loadSample() {
    $('processingOverlay').classList.add('active'); setMsg('');
    try {
      const r = await fetch('./plantilla/Amberes_Modelo_Consolidado.xlsx'); if (!r.ok) throw new Error('HTTP ' + r.status);
      await ingestBuffer('Amberes_Modelo_Consolidado.xlsx', await r.arrayBuffer()); onLoaded();
    } catch (e) { setMsg('El ejemplo solo se puede cargar desde la página publicada (http/https). Descarga la plantilla y súbela con el botón de selección.', true); }
    finally { $('processingOverlay').classList.remove('active'); }
  }
  function onLoaded() {
    Ain = {}; maOver = {}; plEmpresa = null;
    rebuild(true);
    const unrec = store.inventory.filter(i => i.role === 'unknown' || i.status === 'no reconocida').length;
    setMsg(store.inventory.length + ' pestañas leídas · ' + store.inventory.filter(i => i.status === 'usada' || i.status === 'recalculada').length + ' utilizadas' + (unrec ? ' · ' + unrec + ' no reconocidas' : '') + '.');
  }

  /* ---------- reconstrucción ---------- */
  function rebuild(first) {
    try { M = AM.build(store, Ain); } catch (e) { console.error(e); setMsg('Error al construir el modelo: ' + e.message, true); return; }
    $('empty').style.display = M.empty ? '' : 'none';
    $('app').style.display = M.empty ? 'none' : '';
    $('sourceBar').style.display = M.empty ? 'none' : '';
    if (M.empty) { if (store.files.length) setMsg('No se encontró un P&L ni una evaluación con datos en el archivo. Revisa la pestaña "Modelo" o descarga la plantilla.', true); return; }
    renderSources(); if (first) renderAssump();
    renderAll();
  }
  function renderAll() {
    renderOverview(); renderPL(); renderCF(); renderBS(); renderMA(); renderModel();
  }

  function renderSources() {
    const bar = clear($('sourceBar')); const S = M.sources;
    const word = { cargado: 'cargado del archivo', derivado: 'generado automáticamente', ausente: 'sin datos' };
    [['P&L', S.pl, 'pl'], ['Evaluación de proyectos', S.proj, 'overview'], ['Flujo de caja', S.cf, 'cf'], ['Balance', S.bs, 'bs'], ['M&A', S.ma, 'ma']].forEach(([n, v, t]) =>
      bar.appendChild(h('button', { type: 'button', class: 'src-chip ' + v.state, title: 'Ir a ' + n, onclick: () => switchTab(t) }, h('i'), n + ': ' + word[v.state], v.note ? h('em', { text: '· ' + v.note }) : null)));
    bar.appendChild(h('span', { class: 'bar-actions' }, h('button', { class: 'btn ghost sm', type: 'button', text: '⬇ PDF', onclick: exportPdf }), h('button', { class: 'btn ghost sm', type: 'button', text: '⬇ Excel', onclick: exportXlsx })));
  }

  /* ---------- supuestos ---------- */
  function renderAssump() {
    const box = clear($('assump')); const A = M.Aeff;
    const det = h('details', null, h('summary', { text: '⚙️ Supuestos del modelo (editables) — se usan donde tu archivo no trae el dato' }));
    const g = h('div', { class: 'grid' });
    const field = (label, el) => h('div', null, h('label', { text: label }), el);
    const sel = (key, opts) => { const s = h('select', null, opts.map(([v, t]) => h('option', { value: v, text: t }))); s.value = Ain[key] || 'auto'; s.addEventListener('change', () => { Ain[key] = s.value === 'auto' ? undefined : s.value; if (key === 'empresa') plEmpresa = null; rebuild(); }); return s; };
    const inp = (key, ph, scale, step) => {
      const i = h('input', { type: 'number', step: step || 'any', placeholder: ph });
      if (Ain[key] != null) i.value = scale ? (Ain[key] * scale) : Ain[key];
      i.addEventListener('change', () => { const v = i.value === '' ? null : parseFloat(i.value); Ain[key] = v == null || isNaN(v) ? undefined : (scale ? v / scale : v); if (Ain[key] === undefined) delete Ain[key]; rebuild(); });
      return i;
    };
    g.appendChild(field('Tipo de empresa', sel('empresa', [['auto', 'Automático'], ['servicios', 'Servicios'], ['productiva', 'Productiva']])));
    g.appendChild(field('Escenario P&L base', sel('escenario', [['auto', 'Automático'], ['real', 'Real'], ['budget', 'Budget'], ['forecast', 'Forecast']])));
    g.appendChild(field('DSO · días de cobro', inp('dso', String(A.dso))));
    g.appendChild(field('DIO · días de inventario', inp('dio', String(A.dio))));
    g.appendChild(field('DPO · días de pago', inp('dpo', String(A.dpo))));
    g.appendChild(field('Capex (% de ingresos)', inp('capexPct', M.biz ? (M.biz.capexPct * 100).toFixed(1) : '3', 100)));
    g.appendChild(field('Caja inicial', inp('cash0', 'automática')));
    g.appendChild(field('Deuda inicial', inp('debt0', 'automática')));
    g.appendChild(field('Pago mensual de deuda', inp('repay', '0')));
    g.appendChild(field('Payout (% utilidad)', inp('payout', '0', 100)));
    g.appendChild(field('Costo de la deuda (%)', inp('kd', (A.kd * 100).toFixed(1), 100)));
    g.appendChild(field('Tasa de impuesto (%)', inp('tax', (A.tax * 100).toFixed(0), 100)));
    g.appendChild(field('Ingresos anuales (si solo hay balance o flujo)', inp('revAnnual', 'estimados')));
    g.appendChild(field('Margen EBITDA supuesto (%)', inp('ebitdaMargin', '15', 100)));
    g.appendChild(field('Costo de ventas (% ingresos)', inp('cogsPct', '60', 100)));
    g.appendChild(field('Vida útil activos (años)', inp('life', '10')));
    g.appendChild(field('Meses entre balances', inp('balMonths', '12')));
    g.appendChild(h('div', { class: 'hint', text: 'Deja un campo en blanco para usar el valor del archivo o la estimación del software. Los cambios recalculan flujo de caja, balance y M&A.' }));
    det.appendChild(g); box.appendChild(det);
  }

  /* ---------- 1. VISIÓN GENERAL ---------- */
  function renderOverview() {
    const root = clear($('tab-overview')); const O = M.overview, P = M.proj;
    root.appendChild(h('div', { class: 'metrics-grid' }, O.kpis.map(metric)));
    root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chOvCF', O.charts.cf.title || 'Flujo de caja'), chartBox('chOvEB', O.charts.eb.title || 'EBITDA y EBIT')));
    if (P && P.tornado) {
      root.appendChild(h('div', { class: 'section-h', text: 'Sensibilidad y escenarios del proyecto' }));
      root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chTornado', 'Tornado: impacto en el VAN (±10%)'), whatIfPanel(P)));
      if (P.scenarios) root.appendChild(table(['Escenario', 'Crec. anual', '% costos var.', 'WACC', 'Inversión', 'VAN', 'TIR anual', 'Payback'],
        P.scenarios.map(s => ({ cells: [s.label, pc(s.g), pc(s.cv), pc(s.wacc), fm(s.capex), fm(s.van), s.tirA == null ? 'n/d' : pc(s.tirA, 1), s.payback ? 'mes ' + s.payback : 'No recupera'], neg: [0, 0, 0, 0, s.van < 0, s.tirA != null && s.tirA < 0, !s.payback] }))));
      const notes = P.scenarios && P.scenarios.filter(s => s.note).map(s => s.label + ': ' + s.note);
      if (notes && notes.length) root.appendChild(h('div', { class: 'sub-note', text: notes.join(' · ') }));
    }
    root.appendChild(insightList(O.insights, '💡 Insights estratégicos'));
    // gráficos
    const cf = O.charts.cf;
    mk('chOvCF', { type: 'bar', data: { labels: cf.labels, datasets: [
      { type: 'bar', label: 'Flujo del período', data: cf.bars, backgroundColor: cf.bars.map(v => v >= 0 ? COL.petrol : COL.red), borderRadius: 3 },
      { type: 'line', label: 'Acumulado', data: cf.acc, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2, tension: .2 }] },
      options: { scales: { x: xAxis, y: yAxisM } } });
    const eb = O.charts.eb;
    mk('chOvEB', { type: 'line', data: { labels: eb.labels, datasets: [
      { label: 'EBITDA', data: eb.ebitda, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2, tension: .25 },
      { label: 'EBIT', data: eb.ebit, borderColor: COL.petrol, backgroundColor: COL.petrol, pointRadius: 0, borderWidth: 2, tension: .25 }] },
      options: { scales: { x: xAxis, y: yAxisM } } });
    if (P && P.tornado) {
      const t = P.tornado.slice().sort((a, b) => b.swing - a.swing);
      mk('chTornado', { type: 'bar', data: { labels: t.map(x => x.label), datasets: [
        { label: 'Variable −10%', data: t.map(x => x.down), backgroundColor: COL.red },
        { label: 'Variable +10%', data: t.map(x => x.up), backgroundColor: COL.green }] },
        options: { indexAxis: 'y', scales: { x: yAxisM, y: { ticks: { color: COL.dim }, grid: { display: false } } } } });
    }
  }
  function whatIfPanel(P) {
    const box = h('div', { class: 'sensitivity-panel', style: 'margin:0' }, h('div', { class: 'sensitivity-title', text: '🎛️ What-if del proyecto (motor mensual validado)' }));
    if (!P.canWhatIf) { box.appendChild(h('div', { class: 'insight-item info', text: 'El libro no trae los supuestos completos para simular. Carga la hoja "Supuestos".' })); return box; }
    const base = { g: P.sup.g, cv: P.sup.cv, wacc: P.sup.wacc, capex: P.sup.capex };
    const cur = Object.assign({}, base);
    const out = h('div', { class: 'sens-impact neutral', text: '' });
    const upd = () => {
      const r = AM.whatIf(P, cur); if (!r) return;
      const dv = r.van - P.van;
      out.className = 'sens-impact ' + (r.van >= 0 ? 'positive' : 'negative');
      out.textContent = 'VAN ' + fm(r.van) + ' (' + (dv >= 0 ? '+' : '') + fm(dv) + ' vs base) · TIR anual ' + (r.tirA == null ? 'n/d' : pc(r.tirA)) + ' · Payback ' + (r.payback ? 'mes ' + r.payback : 'no recupera en ' + P.H + ' meses');
    };
    const slider = (key, label, min, max, step, show) => {
      const val = h('b', { text: show(cur[key]) });
      const s = h('input', { type: 'range', min, max, step, value: cur[key] });
      s.addEventListener('input', () => { cur[key] = parseFloat(s.value); val.textContent = show(cur[key]); upd(); });
      return h('div', { class: 'ma-field' }, h('label', null, label, val), s);
    };
    box.appendChild(slider('g', 'Crecimiento anual de ingresos', -0.1, 0.4, 0.005, v => pc(v)));
    box.appendChild(slider('cv', '% costos variables', 0.1, 0.9, 0.005, v => pc(v)));
    box.appendChild(slider('wacc', 'WACC', 0.04, 0.25, 0.0025, v => pc(v, 2)));
    box.appendChild(slider('capex', 'Inversión en activos fijos', base.capex * 0.3, base.capex * 2, Math.max(1, base.capex / 100), v => fm(v)));
    box.appendChild(out); upd(); return box;
  }

  /* ---------- 2. P&L ---------- */
  function renderPL() {
    const root = clear($('tab-pl'));
    if (!M.hasPL) { renderPLEstimated(root); return; }
    const V = M.plView;
    const avail = ['servicios', 'productiva'].filter(t => ['real', 'budget', 'forecast'].some(e => M.pl[t][e] && M.pl[t][e].hasData));
    if (avail.length > 1) {
      const row = h('div', { class: 'selector-row' }, h('span', { class: 'selector-label', text: 'Tipo de empresa:' }));
      avail.forEach(t => row.appendChild(h('button', { class: 'sel-btn' + (V.tipo === t ? ' active' : ''), text: t === 'productiva' ? '🏭 Productiva' : '🏢 Servicios', onclick: () => { Ain.empresa = t; rebuild(); } })));
      root.appendChild(row);
    }
    root.appendChild(h('div', { class: 'metrics-grid' }, V.kpis.map(metric)));
    const sc = Object.keys(V.scen).filter(k => V.scen[k].pl && V.scen[k].pl.hasData);
    const labelOf = k => V.scen[k].label;
    root.appendChild(h('div', { class: 'section-h', text: 'Análisis horizontal · comparativo de escenarios' }));
    const head = ['Concepto'].concat(sc.map(labelOf), sc.map(k => '% ing. ' + labelOf(k)));
    if (sc.includes('real') && sc.includes('budget')) head.push('Var. Real vs Budget', '%');
    if (sc.includes('real') && sc.includes('forecast')) head.push('Var. Real vs Forecast', '%');
    const rows = V.table.map(t => {
      const cells = [t.label].concat(sc.map(k => fm(t.vals[k])), sc.map(k => pc(t.pct[k])));
      const neg = sc.map(() => 0).concat(sc.map(() => 0));
      const add = (v) => { cells.push(fm(v.abs), v.pct == null ? '—' : pc(v.pct)); const good = t.better === 'low' ? v.abs <= 0 : v.abs >= 0; neg.push(!good, !good); };
      if (sc.includes('real') && sc.includes('budget')) add(t.varRB);
      if (sc.includes('real') && sc.includes('forecast')) add(t.varRF || t.varFB);
      return { cells, cls: (t.kind === 'sub' || t.kind === 'total') ? 'tot' : '', neg };
    });
    root.appendChild(table(head, rows));
    root.appendChild(h('div', { class: 'section-h', text: 'Análisis vertical · composición sobre ingresos (%)' }));
    root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chVert', 'Estructura de costos y resultado sobre ingresos'), chartBox('chMargin', 'Margen mensual por escenario')));
    root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chHRev', 'Ingresos mensuales'), chartBox('chHEbit', 'EBIT mensual')));
    if (V.lineVar && V.lineVar.length) {
      root.appendChild(h('div', { class: 'section-h', text: 'Líneas con mayor variación (' + (V.lineVarNames || []).join(' vs ') + ')' }));
      root.appendChild(table(['Línea', V.lineVarNames[0] || 'A', V.lineVarNames[1] || 'B', 'Variación', '%'],
        V.lineVar.map(l => ({ cells: [l.label, fm(l.a), fm(l.b), fm(l.d), l.pct == null ? '—' : pc(l.pct)], neg: [0, 0, l.group === 'cost' ? l.d > 0 : l.d < 0, l.group === 'cost' ? l.d > 0 : l.d < 0] }))));
    }
    if (V.costLines && V.costLines.length) {
      root.appendChild(h('div', { class: 'section-h', text: 'Principales líneas de costo (peso sobre ingresos)' }));
      root.appendChild(table(['Línea', 'Total', '% ingresos'], V.costLines.slice(0, 8).map(l => [l.label, fm(l.total), pc(l.pct)])));
    }
    root.appendChild(insightList(V.insights, '💡 Insights del P&L'));
    const vt = V.vertical, ds = [['Costo de ventas', 'cogs', COL.red], ['Gastos operacionales', 'opex', COL.amber], ['Resultado no operacional', 'other', COL.petrol], ['Impuestos', 'tax', '#8a6d2b'], ['Utilidad neta', 'ni', COL.green]];
    mk('chVert', { type: 'bar', data: { labels: vt.labels, datasets: ds.filter(d => vt[d[1]].some(v => v)).map(d => ({ label: d[0], data: vt[d[1]], backgroundColor: d[2] })) },
      options: { scales: { x: Object.assign({ stacked: true }, xAxis), y: { stacked: true, ticks: { color: COL.dim, callback: v => v + '%' }, grid: { color: COL.grid }, max: 100 } } } });
    const H = V.horizontal, scCol = { real: COL.gold, budget: COL.petrol, forecast: COL.cream };
    const ln = (obj, pct) => sc.map(k => ({ label: labelOf(k), data: obj[k], borderColor: scCol[k], backgroundColor: scCol[k], pointRadius: 0, borderWidth: 2, tension: .25 }));
    mk('chMargin', { type: 'line', data: { labels: H.labels, datasets: ln(H.margin) }, options: { scales: { x: xAxis, y: { ticks: { color: COL.dim, callback: v => (v * (Math.abs(v) <= 1.5 ? 100 : 1)).toFixed(0) + '%' }, grid: { color: COL.grid } } } } });
    mk('chHRev', { type: 'line', data: { labels: H.labels, datasets: ln(H.rev) }, options: { scales: { x: xAxis, y: yAxisM } } });
    mk('chHEbit', { type: 'line', data: { labels: H.labels, datasets: ln(H.ebit) }, options: { scales: { x: xAxis, y: yAxisM } } });
  }
  function renderPLEstimated(root) {
    const b = M.biz;
    if (!b || !b.ser) { root.appendChild(emptyMsg('📈', 'Sin P&L en el archivo', 'Sube una pestaña de P&L (Estado de resultados) para ver análisis vertical y horizontal, comparativo Real vs Budget vs Forecast, alertas y recomendaciones.')); return; }
    const sr = b.ser, L = b.labels;
    root.appendChild(h('div', { class: 'section-h' }, 'Estado de resultados ', tagEst()));
    root.appendChild(h('div', { class: 'sub-note', text: 'No se cargó un P&L: este estado de resultados se reconstruyó desde tu ' + (M.synth ? M.synth.label.toLowerCase() : 'modelo') + ' con supuestos editables. Sube un P&L real para obtener el análisis vertical/horizontal y la comparación Real vs Budget vs Forecast.' }));
    if (M.synth) root.appendChild(insightList(M.synth.notes, 'ℹ️ Cómo se completó este informe'));
    const tot = a => sum(a), rev = tot(sr.rev);
    root.appendChild(h('div', { class: 'metrics-grid' }, [
      metric({ label: 'Ingresos (estimados)', value: fm(rev), sub: 'Base anual', cls: '' }),
      metric({ label: 'EBITDA', value: fm(tot(sr.ebit) + tot(sr.dep)), sub: 'Margen ' + pc(rev ? (tot(sr.ebit) + tot(sr.dep)) / rev : null), cls: '' }),
      metric({ label: 'EBIT', value: fm(tot(sr.ebit)), sub: 'Margen ' + pc(rev ? tot(sr.ebit) / rev : null), cls: '' }),
      metric({ label: 'Utilidad neta', value: fm(tot(sr.ni)), sub: 'Margen ' + pc(rev ? tot(sr.ni) / rev : null), cls: '' })]));
    const r = (n, a, c) => ({ cells: [n].concat(a.map(v => num(v)), [num(sum(a))]), cls: c || '', neg: a.map(v => v < 0).concat([sum(a) < 0]) });
    root.appendChild(table(['Concepto'].concat(L, ['Total']), [r('Ingresos', sr.rev, 'tot'), r('Costo de ventas', sr.cogs), r('Gastos operacionales', sr.opex.map((v, i) => v - sr.dep[i])), r('Depreciación', sr.dep), r('EBIT', sr.ebit, 'tot'), r('Gastos financieros', sr.finExp), r('Impuestos', sr.tax), r('Utilidad neta', sr.ni, 'tot')]));
    root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chEstRev', 'Ingresos y utilidad neta (estimados)'), chartBox('chEstEb', 'EBITDA y EBIT (estimados)')));
    mk('chEstRev', { type: 'bar', data: { labels: L, datasets: [{ label: 'Ingresos', data: sr.rev, backgroundColor: COL.petrol }, { type: 'line', label: 'Utilidad neta', data: sr.ni, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 2 }] }, options: { scales: { x: xAxis, y: yAxisM } } });
    mk('chEstEb', { type: 'line', data: { labels: L, datasets: [{ label: 'EBITDA', data: sr.ebit.map((e, i) => e + sr.dep[i]), borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2 }, { label: 'EBIT', data: sr.ebit, borderColor: COL.petrol, backgroundColor: COL.petrol, pointRadius: 0, borderWidth: 2 }] }, options: { scales: { x: xAxis, y: yAxisM } } });
  }
  function emptyMsg(icon, t, s) { return h('div', { class: 'empty-state' }, h('div', { class: 'empty-state-icon', text: icon }), h('div', { style: 'font-size:17px;margin-bottom:8px', text: t }), h('div', { style: 'font-size:13px;opacity:.7', text: s })); }

  /* ---------- 3. FLUJO DE CAJA ---------- */
  function renderCF() {
    const root = clear($('tab-cf')); const P = M.proj, b = M.biz;
    if (!P && !b) { root.appendChild(emptyMsg('💰', 'Sin datos para flujo de caja', 'Sube un P&L o una evaluación de proyectos.')); return; }
    if (b) {
      root.appendChild(h('div', { class: 'section-h' }, 'Flujo de caja operacional, de inversión y financiero ', M.sources.cf.state === 'derivado' ? tagEst() : srcTag('cargado')));
      if (M.synth) root.appendChild(insightList(M.synth.notes, 'ℹ️ Cómo se completó este informe'));
      else if (M.sources.cf.state === 'derivado') root.appendChild(h('div', { class: 'insight-item info', text: 'Generado automáticamente desde tu P&L: método indirecto (utilidad neta + depreciación − variación de capital de trabajo − capex ± deuda y dividendos).' }));
      if (M.synth && M.synth.hist) {
        root.appendChild(h('div', { class: 'section-h', text: 'Flujo histórico reconstruido por diferencias entre balances (' + M.synth.hist.periods.join(' → ') + ')' }));
        root.appendChild(table(['Concepto', 'Monto'], M.synth.hist.rows.map(r => ({ cells: [r[0], num(r[1])], cls: r[2] ? 'tot' : '', neg: [r[1] < 0] }))));
      }
      root.appendChild(h('div', { class: 'sub-note', text: 'Método indirecto construido desde: ' + (b.source || '') + '. Moneda: ' + (M.currency || 's/d') + '.' }));
      const R = b.ratios;
      root.appendChild(h('div', { class: 'metrics-grid' }, [
        metric({ label: 'Flujo operacional (CFO)', value: fm(sum(b.cfo)), sub: 'Conversión EBITDA→caja ' + pc(R.cfoToEbitda, 0), sem: sum(b.cfo) >= 0 ? 'green' : 'red', cls: sum(b.cfo) >= 0 ? 'success' : 'danger' }),
        metric({ label: 'Flujo de inversión (CFI)', value: fm(sum(b.cfi)), sub: 'Capex ' + pc(b.capexPct, 1) + ' de ingresos', cls: '' }),
        metric({ label: 'Flujo financiero (CFF)', value: fm(sum(b.cff)), sub: 'Deuda y dividendos', cls: '' }),
        metric({ label: 'Variación neta de caja', value: fm(sum(b.net)), sub: 'Caja final ' + fm(b.bs.cash[b.N]), sem: sum(b.net) >= 0 ? 'green' : 'red', cls: sum(b.net) >= 0 ? 'success' : 'danger' })]));
      root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chCFO', 'CFO · CFI · CFF mensual'), chartBox('chFCF', 'Flujo económico (FCFF) vs flujo al accionista (FCFE)')));
      const L = b.labels, r = (n, a, cls) => ({ cells: [n].concat(a.map(v => num(v))), cls: cls || '', neg: a.map(v => v < 0) });
      root.appendChild(table(['Concepto'].concat(L, ['Total']), [
        r('Utilidad neta', b.ser.ni).cells ? Object.assign(r('Utilidad neta', b.ser.ni), { cells: ['Utilidad neta'].concat(b.ser.ni.map(v => num(v)), [num(sum(b.ser.ni))]) }) : null,
        Object.assign(r('(+) Depreciación y amortización', b.ser.dep), { cells: ['(+) Depreciación y amortización'].concat(b.ser.dep.map(v => num(v)), [num(sum(b.ser.dep))]) }),
        Object.assign(r('(−) Δ Capital de trabajo', b.dNWC.map(v => -v)), { cells: ['(−) Δ Capital de trabajo'].concat(b.dNWC.map(v => num(-v)), [num(-sum(b.dNWC))]) }),
        Object.assign(r('FLUJO OPERACIONAL', b.cfo, 'tot'), { cells: ['FLUJO OPERACIONAL'].concat(b.cfo.map(v => num(v)), [num(sum(b.cfo))]) }),
        Object.assign(r('FLUJO DE INVERSIÓN (capex)', b.cfi, 'tot'), { cells: ['FLUJO DE INVERSIÓN (capex)'].concat(b.cfi.map(v => num(v)), [num(sum(b.cfi))]) }),
        Object.assign(r('FLUJO FINANCIERO', b.cff, 'tot'), { cells: ['FLUJO FINANCIERO'].concat(b.cff.map(v => num(v)), [num(sum(b.cff))]) }),
        Object.assign(r('VARIACIÓN NETA DE CAJA', b.net, 'tot'), { cells: ['VARIACIÓN NETA DE CAJA'].concat(b.net.map(v => num(v)), [num(sum(b.net))]) })
      ].filter(Boolean)));
      mk('chCFO', { type: 'bar', data: { labels: L, datasets: [
        { label: 'Operacional', data: b.cfo, backgroundColor: COL.green, stack: 's' }, { label: 'Inversión', data: b.cfi, backgroundColor: COL.amber, stack: 's' }, { label: 'Financiero', data: b.cff, backgroundColor: COL.petrol, stack: 's' },
        { type: 'line', label: 'Variación neta', data: b.net, borderColor: COL.cream, backgroundColor: COL.cream, pointRadius: 2, borderWidth: 2 }] }, options: { scales: { x: Object.assign({ stacked: true }, xAxis), y: Object.assign({ stacked: true }, yAxisM) } } });
      mk('chFCF', { type: 'line', data: { labels: L, datasets: [
        { label: 'FCFF (económico)', data: b.fcff, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2, tension: .25 },
        { label: 'FCFE (accionista)', data: b.fcfe, borderColor: COL.petrol, backgroundColor: COL.petrol, pointRadius: 0, borderWidth: 2, tension: .25 }] }, options: { scales: { x: xAxis, y: yAxisM } } });
    }
    if (P) {
      root.appendChild(h('div', { class: 'section-h' }, 'Flujo del proyecto: económico (FCL) y financiero (FCFE) ', srcTag('cargado')));
      root.appendChild(h('div', { class: 'sub-note', text: 'FCL = NOPAT + depreciación − capex − Δ capital de trabajo − pre-operativos + valor residual. Descuento mensual (1+WACC)^(1/12)−1.' }));
      root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chPFcl', 'FCL mensual y acumulado'), chartBox('chPFcfe', 'FCL vs FCFE acumulados')));
      const yrs = Math.ceil(P.H / 12), rowsY = [];
      for (let y = 0; y <= yrs; y++) {
        const a = y === 0 ? 0 : (y - 1) * 12 + 1, z = y === 0 ? 0 : Math.min(y * 12, P.H);
        const s = arr => { let t = 0; for (let i = a; i <= z; i++) t += (arr && arr[i]) || 0; return t; };
        rowsY.push({ cells: [y === 0 ? 'Año 0 (inversión)' : 'Año ' + y, num(s(P.fcl)), num(s(P.fcfe)), num(P.acc[z])], neg: [s(P.fcl) < 0, s(P.fcfe) < 0, P.acc[z] < 0] });
      }
      root.appendChild(table(['Período', 'FCL (económico)', 'FCFE (financiero)', 'FCL acumulado'], rowsY));
      const lab = P.fcl.map((_, i) => 'M' + i);
      mk('chPFcl', { type: 'bar', data: { labels: lab, datasets: [
        { type: 'bar', label: 'FCL', data: P.fcl, backgroundColor: P.fcl.map(v => v >= 0 ? COL.petrol : COL.red) },
        { type: 'line', label: 'Acumulado', data: P.acc, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2 }] }, options: { scales: { x: { ticks: { color: COL.dim, maxTicksLimit: 13 }, grid: { display: false } }, y: yAxisM } } });
      const accE = []; let c = 0; (P.fcfe || []).forEach((v, i) => { c += v; accE.push(c); });
      mk('chPFcfe', { type: 'line', data: { labels: lab, datasets: [
        { label: 'FCL acumulado', data: P.acc, borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 0, borderWidth: 2 },
        { label: 'FCFE acumulado', data: accE, borderColor: COL.petrol, backgroundColor: COL.petrol, pointRadius: 0, borderWidth: 2 }] }, options: { scales: { x: { ticks: { color: COL.dim, maxTicksLimit: 13 }, grid: { display: false } }, y: yAxisM } } });
    }
    root.appendChild(insightList(M.insights.cf, '💡 Alertas y recomendaciones de caja'));
  }
  const sum = a => (a || []).reduce((s, v) => s + (v || 0), 0);

  /* ---------- 4. BALANCE ---------- */
  function renderBS() {
    const root = clear($('tab-bs')); const b = M.biz;
    if (!b || !b.bs.assets.length) { root.appendChild(emptyMsg('🧾', 'Sin datos para balance', 'Sube un P&L o una evaluación de proyectos.')); return; }
    root.appendChild(h('div', { class: 'section-h' }, 'Balance proyectado ', M.sources.bs.state === 'derivado' ? tagEst() : srcTag('cargado')));
    if (M.synth && M.synth.kind === 'cf') root.appendChild(insightList(M.synth.notes, 'ℹ️ Cómo se completó este informe'));
    root.appendChild(h('div', { class: 'sub-note', text: (M.sources.bs.state === 'derivado' ? 'Saldo inicial estimado con DSO ' + M.Aeff.dso + ' · DIO ' + M.Aeff.dio + ' · DPO ' + M.Aeff.dpo + ' días. ' : 'Saldo inicial tomado de la hoja Balance. ') + 'El patrimonio se ajusta para que el balance cuadre; la fila de control debe ser cero.' }));
    const R = b.ratios;
    root.appendChild(h('div', { class: 'metrics-grid' }, [
      metric({ label: 'Razón corriente', value: R.currentRatio == null ? '—' : R.currentRatio.toFixed(2) + 'x', sub: 'Liquidez de corto plazo', sem: R.currentRatio >= 1.2 ? 'green' : R.currentRatio >= 1 ? 'yellow' : 'red', cls: R.currentRatio >= 1.2 ? 'success' : R.currentRatio >= 1 ? 'warning' : 'danger' }),
      metric({ label: 'Deuda neta', value: fm(R.netDebt), sub: R.ndEbitda == null ? '' : 'Deuda neta/EBITDA ' + R.ndEbitda.toFixed(2) + 'x', sem: R.ndEbitda == null ? '' : (R.ndEbitda < 2 ? 'green' : R.ndEbitda < 3.5 ? 'yellow' : 'red'), cls: R.ndEbitda != null && R.ndEbitda >= 3.5 ? 'danger' : '' }),
      metric({ label: 'Ciclo de conversión de caja', value: R.ccc == null ? '—' : Math.round(R.ccc) + ' días', sub: 'CT/ingresos ' + pc(R.nwcPctRev, 1), cls: '' }),
      metric({ label: 'Autonomía de caja', value: R.runway == null ? '—' : (R.runway > 60 ? '>60' : R.runway.toFixed(1)) + ' meses', sub: 'Caja / gasto mensual', sem: R.runway >= 6 ? 'green' : R.runway >= 3 ? 'yellow' : 'red', cls: R.runway >= 6 ? 'success' : R.runway >= 3 ? 'warning' : 'danger' }),
      metric({ label: 'ROE / ROA anualizado', value: (R.roe == null ? '—' : pc(R.roe, 1)) + ' / ' + (R.roa == null ? '—' : pc(R.roa, 1)), sub: 'Utilidad neta anualizada ' + fm(R.niAnn), cls: '' })]));
    root.appendChild(h('div', { class: 'chart-grid' }, chartBox('chBSa', 'Composición del activo'), chartBox('chBSp', 'Pasivo y patrimonio')));
    const hdr = ['Concepto', 'Inicial'].concat(b.labels), bs = b.bs;
    const row = (n, a, cls) => ({ cells: [n].concat(a.map(v => num(v))), cls: cls || '', neg: a.map(v => v < 0) });
    const rows = [row('Caja', bs.cash), row(b.nwcFixed ? 'Capital de trabajo neto' : 'Cuentas por cobrar', bs.ar)];
    if (!b.nwcFixed) rows.push(row('Inventarios', bs.inv));
    rows.push(row('Activo fijo neto', bs.ppe), row('TOTAL ACTIVOS', bs.assets, 'tot'));
    if (!b.nwcFixed) rows.push(row('Cuentas por pagar', bs.ap));
    rows.push(row('Deuda financiera', bs.debt), row('TOTAL PASIVOS', bs.liab, 'tot'), row('Capital', bs.capital), row('Resultados acumulados', bs.re), row('TOTAL PATRIMONIO', bs.equity, 'tot'));
    const ok = Math.abs(R.maxCheck || 0) < 1;
    const chk = row('Control (activos − pasivos − patrimonio)', bs.check, 'chk' + (ok ? '' : ' bad'));
    rows.push(chk);
    root.appendChild(table(hdr, rows));
    root.appendChild(h('div', { class: 'sub-note', text: ok ? '✔ El balance cuadra en todos los períodos.' : '⚠ Existen descuadres: revisa los supuestos.' }));
    const L = ['Inicial'].concat(b.labels);
    mk('chBSa', { type: 'bar', data: { labels: L, datasets: [{ label: 'Caja', data: bs.cash, backgroundColor: COL.green }, { label: 'Cuentas por cobrar', data: bs.ar, backgroundColor: COL.gold }, { label: 'Inventarios', data: bs.inv, backgroundColor: COL.amber }, { label: 'Activo fijo', data: bs.ppe, backgroundColor: COL.petrol }] }, options: { scales: { x: Object.assign({ stacked: true }, xAxis), y: Object.assign({ stacked: true }, yAxisM) } } });
    mk('chBSp', { type: 'bar', data: { labels: L, datasets: [{ label: 'Proveedores', data: bs.ap, backgroundColor: COL.amber }, { label: 'Deuda', data: bs.debt, backgroundColor: COL.red }, { label: 'Patrimonio', data: bs.equity, backgroundColor: COL.petrol }] }, options: { scales: { x: Object.assign({ stacked: true }, xAxis), y: Object.assign({ stacked: true }, yAxisM) } } });
    root.appendChild(insightList(M.insights.bs, '💡 Alertas y recomendaciones de balance'));
  }

  /* ---------- 5. M&A ---------- */
  const MA_FIELDS = [
    ['Operación (DCF)', [['ingBase', 'Ingresos base (año 0)', 'money'], ['growth', 'Crecimiento anual', 'pct', -0.1, 0.4, 0.005], ['margin', 'Margen EBITDA', 'pct', 0, 0.6, 0.005], ['da', 'D&A (% ingresos)', 'pct', 0, 0.15, 0.0025], ['capex', 'Capex (% ingresos)', 'pct', 0, 0.2, 0.0025], ['nwc', 'Capital de trabajo (% Δ ingresos)', 'pct', 0, 0.4, 0.005], ['tax', 'Impuesto', 'pct', 0, 0.45, 0.01], ['years', 'Años de proyección', 'int', 3, 10, 1]]],
    ['Costo de capital', [['termG', 'Crecimiento terminal (g)', 'pct', 0, 0.05, 0.0025], ['rf', 'Tasa libre de riesgo', 'pct', 0, 0.12, 0.0025], ['beta', 'Beta', 'x', 0.3, 2.5, 0.05], ['erp', 'Prima de riesgo de mercado', 'pct', 0.02, 0.12, 0.0025], ['country', 'Riesgo país / tamaño', 'pct', 0, 0.1, 0.0025], ['kd', 'Costo de deuda', 'pct', 0, 0.2, 0.0025], ['de', 'Deuda / (Deuda + Patrimonio)', 'pct', 0, 0.8, 0.01]]],
    ['Múltiplos, deuda y sinergias', [['mult', 'Múltiplo EV/EBITDA', 'x', 2, 20, 0.25], ['netDebt', 'Deuda financiera neta', 'money'], ['synCost', 'Sinergias de costos (anual)', 'money'], ['synRev', 'Sinergias de ingresos (anual)', 'money'], ['synYears', 'Años para capturar sinergias', 'int', 1, 8, 1], ['synOneoff', 'Costo de integración', 'money'], ['premium', 'Prima por control', 'pct', 0, 0.6, 0.01]]]
  ];
  function maParams() { return Object.assign({}, M.maParams, maOver); }
  function renderMA() {
    const root = clear($('tab-ma'));
    const left = h('div', { class: 'ma-panel' });
    const right = h('div', { id: 'maRight' });
    const fmtV = (kind, v) => kind === 'pct' ? pc(v, 2) : kind === 'x' ? v.toFixed(2) + (kind === 'x' ? '' : '') : kind === 'int' ? String(Math.round(v)) : fm(v);
    left.appendChild(h('div', { class: 'btn-row' }, h('button', { class: 'btn ghost', text: '↺ Restablecer desde el modelo', onclick: () => { maOver = {}; renderMA(); } })));
    MA_FIELDS.forEach(([title, fields]) => {
      left.appendChild(h('div', { class: 'ma-section-title', text: title }));
      fields.forEach(([k, label, kind, mn, mx, st]) => {
        const v = maParams()[k]; const src = M.maSrc[k] || '';
        const valEl = h('b', { text: fmtV(kind, v) });
        const lab = h('label', null, h('span', null, label, h('span', { class: 'tag-src ' + (/archivo|Archivo/.test(src) ? 'file' : ''), text: src.length > 24 ? src.slice(0, 22) + '…' : src })), valEl);
        let ctl;
        if (kind === 'money') {
          ctl = h('input', { type: 'number', value: Math.round(v), step: 'any' });
          ctl.addEventListener('change', () => { const n = parseFloat(ctl.value); if (!isNaN(n)) { maOver[k] = n; valEl.textContent = fm(n); drawMA(right); } });
        } else {
          ctl = h('input', { type: 'range', min: mn, max: mx, step: st, value: v });
          ctl.addEventListener('input', () => { maOver[k] = parseFloat(ctl.value); valEl.textContent = fmtV(kind, maOver[k]); drawMA(right); });
        }
        left.appendChild(h('div', { class: 'ma-field' }, lab, ctl));
      });
    });
    left.appendChild(h('div', { class: 'ma-note', text: 'Las etiquetas indican el origen de cada parámetro: del archivo, del modelo (P&L/proyecto) o supuesto de mercado por defecto.' }));
    root.appendChild(h('div', { class: 'ma-layout' }, left, right)); drawMA(right);
  }
  function drawMA(right) {
    const p = maParams(), c = AM.maCompute(p); clear(right);
    const card = (cls, l, v, s) => h('div', { class: 'ma-result-card ' + cls }, h('div', { class: 'ma-result-label', text: l }), h('div', { class: 'ma-result-value', text: v }), h('div', { class: 'ma-result-sub', text: s || '' }));
    right.appendChild(h('div', { class: 'ma-results-grid' },
      card('blue', 'WACC', pc(c.wacc.wacc, 2), 'Ke ' + pc(c.wacc.ke, 1) + ' · Kd neto ' + pc(c.wacc.kdAT, 1)),
      card('green', 'Valor empresa (DCF)', fm(c.dcf.ev), 'VP terminal ' + pc(c.dcf.ev ? c.dcf.pvTV / c.dcf.ev : 0, 0) + ' del EV'),
      card('amber', 'Valor patrimonio', fm(c.equity), 'Deuda neta ' + fm(p.netDebt)),
      card('blue', 'EV por múltiplos', fm(c.evMult), p.mult.toFixed(1) + 'x EBITDA año 1'),
      card('green', 'EV con sinergias', fm(c.evComb), 'VP sinergias ' + fm(c.npvSyn)),
      card('amber', 'Patrimonio con prima', fm(c.equityControl), 'Prima ' + pc(p.premium, 0))));
    right.appendChild(h('div', { class: 'chart-grid' }, chartBox('chFoot', 'Football field · rangos de valor (EV)', 260), chartBox('chDcf', 'Flujo libre proyectado', 260)));
    right.appendChild(h('div', { class: 'section-h', text: 'DCF · proyección y valor terminal' }));
    const yrs = c.rows.map(r => 'Año ' + r.y);
    const R = (n, k, cls) => ({ cells: [n].concat(c.rows.map(r => num(r[k]))), cls: cls || '', neg: c.rows.map(r => r[k] < 0) });
    right.appendChild(h('div', { class: 'ma-table-wrap' }, h('table', { class: 'ma-table' }, h('thead', null, h('tr', null, ['Concepto'].concat(yrs).map(x => h('th', { text: x })))),
      h('tbody', null, [R('Ingresos', 'rev'), R('EBITDA', 'ebitda'), R('D&A', 'da'), R('EBIT', 'ebit'), R('Impuestos', 'taxAmt'), R('NOPAT', 'nopat'), R('Capex', 'capex'), R('Δ Capital de trabajo', 'nwc'), R('Flujo libre', 'fcf', 'ma-total')].map(r => h('tr', { class: r.cls }, r.cells.map(x => h('td', { text: x }))))))));
    right.appendChild(h('div', { class: 'sub-note', style: 'margin-top:10px', text: 'VP flujos ' + fm(c.dcf.pvSum) + ' + VP valor terminal ' + fm(c.dcf.pvTV) + ' = EV ' + fm(c.dcf.ev) + ' (g efectivo ' + pc(c.dcf.g, 2) + ').' }));
    const G = AM.maGrid(p);
    right.appendChild(h('div', { class: 'section-h', text: 'Sensibilidad del EV: WACC × crecimiento terminal' }));
    right.appendChild(h('div', { class: 'ma-table-wrap' }, h('table', { class: 'ma-table' }, h('thead', null, h('tr', null, [h('th', { text: 'WACC \\ g' })].concat(G.gs.map(g => h('th', { text: pc(g, 2) }))))),
      h('tbody', null, G.waccs.map((w, i) => h('tr', null, [h('td', { text: pc(w, 2) })].concat(G.cells[i].map((v, j) => h('td', { text: fm(v), class: (i === 2 && j === 2) ? 'pos' : '' })))))))));
    if (p.synCost || p.synRev) {
      right.appendChild(h('div', { class: 'section-h', text: 'Sinergias (después de impuestos)' }));
      right.appendChild(table(['Año', 'Costos', 'Ingresos', 'Neto'], c.syn.map(s => [String(s.y), fm(s.cost), fm(s.rev), fm(s.afterTax)])));
    }
    right.appendChild(insightList(AM.maInsights(p, c, M.maSrc), '💡 Recomendaciones M&A'));
    mk('chFoot', { type: 'bar', data: { labels: c.football.map(f => f.label), datasets: [{ label: 'Rango EV', data: c.football.map(f => [f.lo, f.hi]), backgroundColor: COL.gold, borderRadius: 4 }] },
      options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: yAxisM, y: { ticks: { color: COL.dim }, grid: { display: false } } } } });
    mk('chDcf', { type: 'bar', data: { labels: yrs, datasets: [{ label: 'Flujo libre', data: c.rows.map(r => r.fcf), backgroundColor: c.rows.map(r => r.fcf >= 0 ? COL.petrol : COL.red) }, { type: 'line', label: 'EBITDA', data: c.rows.map(r => r.ebitda), borderColor: COL.gold, backgroundColor: COL.gold, pointRadius: 3 }] }, options: { scales: { x: xAxis, y: yAxisM } } });
  }

  /* ---------- 6. MODELO Y EXPORTAR ---------- */
  function renderModel() {
    const root = clear($('tab-model'));
    root.appendChild(h('div', { class: 'section-h', text: 'Pestañas identificadas' }));
    root.appendChild(table(['Archivo', 'Pestaña', 'Módulo', 'Estado', 'Detalle'], M.inventory.map(i => ({ cells: [i.file, i.sheet, AM.ROLE_LABEL[i.role] || i.role, h('span', { class: 'st-' + i.status, text: i.status }), i.detail || ''], cls: 'inv-row' }))));
    root.appendChild(h('div', { class: 'section-h', text: 'Trazabilidad de datos' }));
    const S = M.sources;
    root.appendChild(table(['Módulo', 'Origen', 'Detalle'], [['P&L', S.pl.state, S.pl.note], ['Evaluación de proyectos', S.proj.state, S.proj.note], ['Flujo de caja', S.cf.state, S.cf.note], ['Balance', S.bs.state, S.bs.note], ['M&A', S.ma.state, S.ma.note]]));
    root.appendChild(h('div', { class: 'section-h', text: 'Calidad del modelo' }));
    const q = h('div', { class: 'insights-section rec-list' });
    M.quality.forEach(i => q.appendChild(h('div', { class: 'insight-item ' + i.level }, h('span', { class: 'sev ' + i.level, text: SEV_LABEL[i.level] || 'Nota' }), h('b', { text: i.title }), h('div', { text: i.text }))));
    root.appendChild(q);
    root.appendChild(h('div', { class: 'section-h', text: 'Exportar' }));
    root.appendChild(h('div', { class: 'btn-row' }, h('button', { class: 'btn', text: '⬇ Exportar a Excel', onclick: exportXlsx }), h('button', { class: 'btn ghost', text: '⬇ Exportar a PDF', onclick: exportPdf })));
  }
  function exportXlsx() {
    try {
      const wb = XLSX.utils.book_new();
      AM.exportSheets(M).forEach(s => { const ws = XLSX.utils.aoa_to_sheet(s.aoa); ws['!cols'] = [{ wch: 38 }]; XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31)); });
      XLSX.writeFile(wb, 'Amberes_Intelligence_Reporte.xlsx');
    } catch (e) { setMsg('No se pudo exportar a Excel: ' + e.message, true); }
  }
  function logoData() {
    try {
      const im = $('logoImg'); if (!im || !im.naturalWidth) return null;
      const c = document.createElement('canvas'); c.width = im.naturalWidth; c.height = im.naturalHeight;
      c.getContext('2d').drawImage(im, 0, 0); return { url: c.toDataURL('image/png'), ratio: c.height / c.width };
    } catch (e) { return null; }   // imagen no accesible (p. ej. archivo local): el PDF sale sin logo
  }
  function exportPdf() {
    if (!window.jspdf || !window.AmberesPDF) { setMsg('Las librerías de PDF no se cargaron (¿sin conexión?). Recarga la página e inténtalo de nuevo.', true); return; }
    const prev = tab; $('processingOverlay').classList.add('active');
    const panes = Array.from(document.querySelectorAll('.dashboard-content'));
    panes.forEach(p => p.classList.add('active'));                 // todos visibles para que los gráficos tengan tamaño
    Object.values(charts).forEach(c => { try { c.resize(); c.update('none'); } catch (e) { /* gráfico ya destruido */ } });
    setTimeout(() => {
      try {
        const S = M.sources, st = { cargado: 'Cargado del archivo', derivado: 'Generado automáticamente', ausente: 'Sin datos' };
        const doc = window.AmberesPDF.generate({
          date: new Date().toLocaleDateString('es-CL'), files: store.files.slice(), logo: logoData(),
          sources: [['P&L', st[S.pl.state], S.pl.note], ['Evaluación de proyectos', st[S.proj.state], S.proj.note], ['Flujo de caja', st[S.cf.state], S.cf.note], ['Balance', st[S.bs.state], S.bs.note], ['M&A', st[S.ma.state], S.ma.note]],
          panes: [{ id: 'tab-overview', title: 'Visión General' }, { id: 'tab-pl', title: 'Análisis P&L' }, { id: 'tab-cf', title: 'Flujo de Caja' }, { id: 'tab-bs', title: 'Balance' }, { id: 'tab-ma', title: 'M&A / Valoración' }, { id: 'tab-model', title: 'Modelo y calidad de datos' }]
        });
        doc.save('Amberes_Intelligence_Informe.pdf'); setMsg('Informe PDF generado.');
      } catch (e) { console.error(e); setMsg('No se pudo generar el PDF: ' + e.message, true); }
      finally { $('processingOverlay').classList.remove('active'); switchTab(prev); }
    }, 450);
  }

  /* ---------- pestañas ---------- */
  function switchTab(t) {
    tab = t;
    document.querySelectorAll('#tabs .tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === t));
    document.querySelectorAll('.dashboard-content').forEach(p => p.classList.toggle('active', p.id === 'tab-' + t));
    Object.values(charts).forEach(c => { try { c.resize(); } catch (e) { /* canvas oculto */ } });
  }

  /* ---------- arranque ---------- */
  chartDefaults();
  document.querySelectorAll('#tabs .tab-btn').forEach(b => b.addEventListener('click', () => switchTab(b.dataset.tab)));
  const fi = $('fileInput'), dz = $('dropZone');
  $('btnPick').addEventListener('click', () => fi.click());
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  fi.addEventListener('change', () => { loadFiles(fi.files); fi.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('dragover'); }));
  dz.addEventListener('drop', e => loadFiles(e.dataTransfer.files));
  $('btnSample').addEventListener('click', loadSample);
  const logo = $('logoImg'); if (logo) logo.addEventListener('error', () => { logo.style.display = 'none'; });
  window.__AM_APP = { get M() { return M; }, store, loadFiles, rebuild };
})();

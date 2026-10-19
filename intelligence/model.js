/* ============================================================
   AMBERES INTELLIGENCE™ — Motor de modelo unificado
   Sin dependencias del DOM: se puede probar en Node.
   Flujo: ingest() → build() → model (con trazabilidad de fuentes)
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AmberesModel = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

const AM = {};
const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

/* ---------- utilidades ---------- */
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const isNum = v => typeof v === 'number' && isFinite(v);
const num = v => isNum(v) ? v : (typeof v === 'string' && v.trim() !== '' && isFinite(Number(v)) ? Number(v) : 0);
const sum = a => (a || []).reduce((x, y) => x + (isNum(y) ? y : 0), 0);
const avg = a => (a && a.length) ? sum(a) / a.length : 0;
const lastOf = a => (a && a.length) ? a[a.length - 1] : 0;
const zeros = n => Array.from({ length: n }, () => 0);
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
AM.util = { norm, num, sum, avg, isNum };

/* ---------- TIR por bisección (estable) ---------- */
function irrBisect(cf, lo, hi) {
  lo = lo == null ? -0.99 : lo; hi = hi == null ? 1 : hi;
  if (!cf || cf.length < 2 || !cf.some(v => v > 0) || !cf.some(v => v < 0)) return null;
  const f = r => { let s = 0; for (let t = 0; t < cf.length; t++) s += cf[t] / Math.pow(1 + r, t); return s; };
  let fl = f(lo), fh = f(hi);
  let guard = 0;
  while (fl * fh > 0 && guard++ < 6) { hi *= 2; fh = f(hi); }
  if (!isFinite(fl) || !isFinite(fh) || fl * fh > 0) return null;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2, fm = f(mid);
    if (Math.abs(fm) < 1e-7) return mid;
    if (fl * fm <= 0) { hi = mid; fh = fm; } else { lo = mid; fl = fm; }
  }
  return (lo + hi) / 2;
}
AM.irrBisect = irrBisect;
const monthlyRate = annual => Math.pow(1 + annual, 1 / 12) - 1;
const npvAt = (cf, rm) => cf.reduce((s, v, t) => s + v / Math.pow(1 + rm, t), 0);

/* ============================================================
   1. CLASIFICACIÓN DE PESTAÑAS (por nombre y por contenido)
   ============================================================ */
const ROLE_LABEL = {
  pl: 'P&L — Estado de Resultados', eval_panel: 'Evaluación · Panel de control', eval_supuestos: 'Evaluación · Supuestos',
  eval_ingresos: 'Evaluación · Ingresos', eval_costos: 'Evaluación · Costos', eval_er: 'Evaluación · Estado de resultados',
  eval_fin: 'Evaluación · Financiamiento', eval_cf: 'Evaluación · Flujo de caja', eval_resumen: 'Evaluación · Resumen',
  eval_esc: 'Evaluación · Escenarios', eval_val: 'Evaluación · Valoración', eval_ratios: 'Evaluación · Ratios',
  balance: 'Balance', cf_stmt: 'Flujo de caja (estado)', ma: 'M&A · Parámetros', info: 'Instrucciones', unknown: 'No reconocida'
};
AM.ROLE_LABEL = ROLE_LABEL;

function classifySheet(name, rows) {
  const n = norm(name);
  const colA = rows.map(r => norm(r && r[0]));
  const has = re => colA.some(l => re.test(l));
  const head = rows.slice(0, 8).map(r => (r || []).slice(0, 4).map(norm).join(' ')).join(' | ');
  const monthsHeader = /\b(enero|january|ene)\b/.test(head);
  const mesHeader = /\bmes\s*[01]\b/.test(head);

  if (/^(instrucciones|leeme|readme|ayuda|portada|control de cambios|notas)/.test(n)) return { role: 'info' };
  if (/m\s?&\s?a|m y a|fusion|adquisicion/.test(n)) return { role: 'ma' };
  if (/balance|situacion financiera/.test(n) && !/p&l/.test(n)) return { role: 'balance' };

  // P&L: por contenido (ingresos netos + utilidad) con encabezado de meses calendario
  if (has(/^ingresos netos/) && (has(/^utilidad neta/) || has(/^utilidad operacional/)) && (monthsHeader || !mesHeader)) {
    let empresa = /productiv/.test(n) ? 'productiva' : /servic/.test(n) ? 'servicios' : null;
    if (!empresa) empresa = has(/costo de servicios/) ? 'servicios' : (has(/costo de ventas|inventario/) ? 'productiva' : 'servicios');
    let esc = /\b(real|actual|reales)\b/.test(n) ? 'real' : /budget|presupuesto|ppto/.test(n) ? 'budget' : /forecast|proyecci|pronostico|estimado/.test(n) ? 'forecast' : null;
    if (!esc) { const m = /escenario:\s*(\w+)/.exec(rows.slice(0, 4).map(r => norm(r && r[0])).join(' ')); esc = m ? (/real/.test(m[1]) ? 'real' : /bud|pres/.test(m[1]) ? 'budget' : /fore|proy/.test(m[1]) ? 'forecast' : 'real') : 'real'; }
    return { role: 'pl', empresa, esc };
  }
  if (/^panel de control|^panel$/.test(n) || has(/informacion del proyecto/)) return { role: 'eval_panel' };
  if (/^supuestos/.test(n) || has(/inversion en activos fijos/)) return { role: 'eval_supuestos' };
  if (/flujo de caja/.test(n) && (mesHeader || has(/flujo de caja libre/))) return { role: 'eval_cf' };
  if (has(/flujo de caja libre/) && mesHeader) return { role: 'eval_cf' };
  if (has(/^flujo (neto )?(operacional|operativo|de operacion)|actividades de operacion|^flujo de caja operacional|^flujo de caja operativo/) && !has(/flujo de caja libre/)) return { role: 'cf_stmt' };
  if (/^(estado de )?flujo (de caja|de efectivo)/.test(n) && has(/utilidad neta|variacion neta|saldo (final|inicial) de (caja|efectivo)/)) return { role: 'cf_stmt' };
  if (/financiamiento/.test(n) || (has(/saldo inicial/) && has(/cuota/))) return { role: 'eval_fin' };
  if (/estado de resultados/.test(n) && mesHeader) return { role: 'eval_er' };
  if (/^(proyeccion de )?ingresos$/.test(n)) return { role: 'eval_ingresos' };
  if (/^(proyeccion de )?costos$/.test(n)) return { role: 'eval_costos' };
  if (/escenarios/.test(n) || has(/pesimista/)) return { role: 'eval_esc' };
  if (/^resumen/.test(n) || (has(/valor actual neto|^van\b/) && has(/tir/) && !has(/metodos? de valoracion/))) return { role: 'eval_resumen' };
  if (/valoracion/.test(n) || has(/metodos? de valoracion|multiplos comparables/)) return { role: 'eval_val' };
  if (/ratios/.test(n) || has(/ratios de liquidez/)) return { role: 'eval_ratios' };
  if (has(/^total activos?/) && has(/patrimonio/)) return { role: 'balance' };
  if (has(/premio por riesgo|beta/) && has(/wacc|crecimiento terminal|tasa libre/)) return { role: 'ma' };
  return { role: 'unknown' };
}
AM.classifySheet = classifySheet;

/* ============================================================
   2. PARSERS POR TIPO DE PESTAÑA
   ============================================================ */
const PL_RULES = [
  ['ingresosBrutos', l => l.startsWith('total ingresos brutos')],
  ['ingresosNetos', l => l === 'ingresos netos'],
  ['costoTotal', l => l === 'costo de ventas total' || l === 'costo de servicios total'],
  ['utilidadBruta', l => l === 'utilidad bruta'],
  ['margenBruto', l => l === 'margen bruto %'],
  ['gastosOp', l => l === 'total gastos operacionales'],
  ['ebit', l => l.startsWith('utilidad operacional')],
  ['margenOp', l => l === 'margen operacional %'],
  ['otrosTotal', l => l.startsWith('total otros ingresos')],
  ['ebt', l => l.startsWith('utilidad antes de impuesto')],
  ['impuesto', l => l.startsWith('(-) impuesto')],
  ['utilidadNeta', l => l === 'utilidad neta'],
  ['margenNeto', l => l === 'margen neto %'],
  ['ingFin', l => l.startsWith('(+) ingresos financieros')],
  ['gastosFin', l => l.startsWith('(-) gastos financieros')],
  ['invInicial', l => l === 'inventario inicial'],
  ['invFinal', l => l.startsWith('(-) inventario final')]
];

function parsePLSheet(rows) {
  const out = {};
  const dep = zeros(13);
  const lines = [];
  let depFound = false, taxRate = null, currency = null, taxLabelRate = null;
  const topText = rows.slice(0, 5).map(r => String((r && r[0]) || '')).join(' | ');
  const mc = /cifras en\s+([A-Za-z$€]+)/i.exec(topText); if (mc) currency = mc[1].toUpperCase();
  rows.forEach(r => {
    if (!r) return;
    const raw = String(r[0] == null ? '' : r[0]);
    const l = norm(raw);
    if (!l) return;
    const hasNums = r.slice(1, 14).some(isNum);
    const vals = r.slice(1, 14).map(num); while (vals.length < 13) vals.push(0);
    const isPct = /%\s*$/.test(raw.trim()) && !/impuesto/.test(l);
    if (!vals[12] && !isPct && vals.slice(0, 12).some(v => v)) vals[12] = sum(vals.slice(0, 12));
    if (l !== 'concepto' && !/^escenario|^cifras|^estado de resultados/.test(l) && r.slice(1, 14).some(v => v != null)) {
      lines.push({ label: raw.trim(), indent: (raw.match(/^\s*/) || [''])[0].length, vals, pct: isPct, key: l });
    } else if (l !== 'concepto' && !hasNums && !/^escenario|^cifras|^estado de resultados/.test(l)) {
      lines.push({ label: raw.trim(), section: true, vals: zeros(13), key: l });
    }
    for (const [key, test] of PL_RULES) {
      if (!out[key] && test(l)) {
        out[key] = vals;
        if (key === 'impuesto') { const m = /\((\d+(?:[.,]\d+)?)\s*%\)/.exec(raw); if (m) { taxRate = parseFloat(m[1].replace(',', '.')) / 100; taxLabelRate = taxRate; } }
        break;
      }
    }
    if (/depreciacion|amortizacion/.test(l) && !/total|%|acumulad/.test(l)) {
      depFound = true; for (let i = 0; i < 13; i++) dep[i] += vals[i];
    }
  });
  out.depreciacion = dep; out.depFound = depFound; out.lines = lines;
  out.taxRate = taxRate; out.taxLabelRate = taxLabelRate; out.currency = currency;
  // tasa efectiva = impuesto / EBT del año (permite detectar etiqueta ≠ cálculo)
  if (out.ebt && out.impuesto) { const e = sum(out.ebt.slice(0, 12)), t = Math.abs(sum(out.impuesto.slice(0, 12))); out.taxEffective = e > 0 ? t / e : null; }
  out.hasData = !!out.ingresosNetos && Math.abs(sum((out.ingresosNetos || []).slice(0, 12))) > 0;
  out.valid = !!out.ingresosNetos && !!(out.utilidadNeta || out.ebit);
  return out;
}

function pairsOf(rows) {
  const out = [];
  let section = '';
  rows.forEach(r => {
    if (!r || r[0] == null) return;
    const label = String(r[0]).trim(); if (!label) return;
    const k = norm(label);
    if (r[1] == null || r[1] === '') { section = k; return; }
    out.push({ k, label, v: r[1], section, row: r });
  });
  return out;
}
const pickPair = (pairs, pats, wantNum) => {
  for (const p of pairs) {
    if (pats.some(x => (x instanceof RegExp) ? x.test(p.k) : p.k.includes(x))) {
      if (wantNum === false) return p.v;
      if (isNum(p.v)) return p.v;
    }
  }
  return undefined;
};

function parseSeriesSheet(rows) {
  let h = rows.findIndex(r => r && (norm(r[0]) === 'concepto' || (r.slice(1, 4).some(c => /^mes\s*\d+/i.test(String(c || ''))))));
  if (h < 0) h = 2;
  const header = rows[h] || [];
  const m = /mes\s*(\d+)/i.exec(String(header[1] || ''));
  const start = m ? +m[1] : 1;
  const series = {};
  rows.slice(h + 1).forEach(r => {
    if (!r || r[0] == null) return;
    const label = String(r[0]).trim(); if (!label) return;
    const vals = r.slice(1);
    if (!vals.some(isNum)) return;
    series[norm(label)] = { label, vals: vals.map(v => isNum(v) ? v : 0) };
  });
  return { start, series };
}
// serie por número de mes (índice 0..H)
function byMonth(parsed, patterns, H) {
  if (!parsed) return null;
  const keys = Object.keys(parsed.series);
  const key = keys.find(k => patterns.some(p => (p instanceof RegExp) ? p.test(k) : (k === p || k.startsWith(p + ' ') || k.startsWith(p + '('))));
  if (!key) return null;
  const s = parsed.series[key].vals, out = zeros(H + 1);
  for (let m = 0; m <= H; m++) { const i = m - parsed.start; if (i >= 0 && i < s.length) out[m] = s[i]; }
  return out;
}

function parsePanel(rows) {
  const p = pairsOf(rows.concat([]));
  // el panel tiene pares "label: valor" incluso en secciones
  const all = [];
  rows.forEach(r => { if (r && r[0] != null && r[1] != null && r[1] !== '') all.push({ k: norm(r[0]), v: r[1] }); });
  const g = (pats, numeric) => { const f = all.find(x => pats.some(t => x.k.includes(t)) && (!numeric || isNum(x.v))); return f ? f.v : undefined; };
  const out = {
    nombre: g(['nombre del proyecto']), tipo: g(['tipo de operacion']), moneda: g(['moneda']), unidad: g(['unidad']),
    horizonte: g(['horizonte'], true), wacc: g(['tasa de descuento', 'wacc'], true), tax: g(['tasa impositiva', 'impuesto'], true)
  };
  if (out.wacc > 1) out.wacc /= 100; if (out.tax > 1) out.tax /= 100;
  return out;
}

function parseSupuestos(rows) {
  const pairs = pairsOf(rows);
  const P = (pats) => pickPair(pairs, pats);
  const fixedItems = pairs.filter(p => p.section.includes('costos fijos') && isNum(p.v));
  const s = {
    capex: P(['inversion en activos fijos']), nwc0: P(['capital de trabajo inicial']), preop: P(['gastos pre-operativos', 'gastos preoperativos']),
    ingMes: P(['ingresos mensuales']), g: P(['crecimiento anual ingresos']), cv: P(['% costos variables', 'costos variables sobre']),
    fixed: fixedItems.length ? sum(fixedItems.map(p => p.v)) : undefined, fixedItems: fixedItems.map(p => ({ label: p.label, v: p.v })),
    life: P(['vida util']), loan: P(['monto del prestamo']), rate: P(['tasa de interes anual']), term: P(['plazo del prestamo']),
    resid: P(['% valor residual', 'valor residual'])
  };
  return s;
}

function parseResumen(rows) {
  const pairs = []; rows.forEach(r => { if (r && r[0] != null && r[1] != null) pairs.push({ k: norm(r[0]), v: r[1] }); });
  const g = pats => { const f = pairs.find(p => pats.some(t => (t instanceof RegExp) ? t.test(p.k) : p.k.includes(t))); return f ? f.v : undefined; };
  const van = g([/^van\b/, 'valor actual neto']);
  const tirA = g(['tir anual']);
  const tirM = g([/^tir .*mensual/]);
  const pay = g(['periodo de recuperacion']);
  return { van: isNum(van) ? van : undefined, tirAnual: isNum(tirA) ? tirA : undefined, tirMensual: isNum(tirM) ? tirM : undefined, payback: isNum(pay) ? pay : null, paybackRaw: pay };
}

function parseEscenarios(rows) {
  const hi = rows.findIndex(r => r && r.some(c => /pesimista/.test(norm(c))));
  if (hi < 0) return null;
  const head = rows[hi].map(norm);
  const col = { pes: head.findIndex(x => /pesimista/.test(x)), base: head.findIndex(x => x === 'base'), opt: head.findIndex(x => /optimista/.test(x)) };
  const out = { pes: {}, base: {}, opt: {} };
  const map = [['g', /crecimiento anual/], ['cv', /costos variables/], ['wacc', /tasa de descuento|wacc/], ['capex', /inversion inicial|activos fijos/]];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const l = norm(r[0]); if (/resultados por escenario/.test(l)) break;
    const m = map.find(x => x[1].test(l)); if (!m) continue;
    ['pes', 'base', 'opt'].forEach(k => { const v = r[col[k]]; if (isNum(v)) out[k][m[0]] = v; });
  }
  return out;
}

function parseValoracion(rows) {
  const pairs = []; rows.forEach(r => { if (r && r[0] != null && r[1] != null) pairs.push({ k: norm(r[0]), v: r[1] }); });
  const g = pats => { const f = pairs.find(p => pats.some(t => p.k.includes(t)) && isNum(p.v)); return f ? f.v : undefined; };
  return {
    mult: g(['multiplo ev/ebitda', 'multiplo']), premium: g(['prima por control', 'prima de control']), synergies: g(['sinergias esperadas']),
    ebitdaProm: g(['ebitda promedio anual']), evDcf: g(['valor de la empresa (enterprise value)']), evMult: g(['valor de la empresa (por multiplo']),
    equityBook: g(['patrimonio neto'])
  };
}

const MA_KEYS = [
  ['ingBase', /ingresos base|ingresos anuales|ingresos \(ano 0\)/], ['growth', /crecimiento anual|crecimiento de ingresos|crecimiento ingresos/],
  ['margin', /margen ebitda/], ['da', /d&a|depreciacion/], ['capex', /capex/], ['nwc', /capital de trabajo/],
  ['tax', /impuesto/], ['years', /anos de proyeccion|horizonte/], ['termG', /crecimiento terminal/],
  ['rf', /libre de riesgo|\brf\b/], ['beta', /\bbeta\b/], ['erp', /premio por riesgo|\berp\b/], ['country', /riesgo pais|tamano/],
  ['kd', /costo de deuda|\bkd\b/], ['de', /deuda\s*\/|relacion deuda|estructura de capital/], ['mult', /multiplo/],
  ['netDebt', /deuda (financiera )?neta/], ['synCost', /sinergias de costos/], ['synRev', /sinergias de ingresos/],
  ['synYears', /anos para capturar|captura/], ['synOneoff', /costo de integracion|integracion/], ['premium', /prima/], ['waccOverride', /^wacc/]
];
const MA_PCT = new Set(['growth', 'margin', 'da', 'capex', 'nwc', 'tax', 'termG', 'rf', 'erp', 'country', 'kd', 'de', 'premium', 'waccOverride']);
function parseMA(rows) {
  const params = {};
  rows.forEach(r => {
    if (!r || r[0] == null || !isNum(r[1])) return;
    const l = norm(r[0]);
    const hit = MA_KEYS.find(([k, re]) => re.test(l) && params[k] === undefined);
    if (!hit) return;
    let v = r[1]; if (MA_PCT.has(hit[0]) && Math.abs(v) > 1) v = v / 100;
    params[hit[0]] = v;
  });
  return params;
}

const BS_KEYS = [
  ['cash', /^(caja|efectivo|banco|disponible)/], ['ar', /clientes|cuentas por cobrar|deudores/], ['inv', /inventario|existencias/],
  ['ppe', /activo fijo|activos fijos|propiedades|planta|inmovilizado|maquinaria|depreciacion acumulada/], ['ap', /proveedores|cuentas por pagar|acreedores/],
  ['debt', /deuda|prestamo|obligaciones financieras|credito bancario|leasing/], ['capital', /^capital (social|pagado|aportado)|^capital$/], ['re', /resultados acumulados|utilidades retenidas|resultado del (ejercicio|periodo)|utilidad del (ejercicio|periodo)/]
];
const BS_SUM = new Set(['cash', 'ar', 'inv', 'ppe', 'ap', 'debt', 're']);
// drivers de proyección (se ponen en la misma hoja Balance, sección "Supuestos de proyección")
const DRV_KEYS = [
  ['dso', /\bdso\b|dias de cobro|periodo medio de cobro/], ['dio', /\bdio\b|dias de inventario/], ['dpo', /\bdpo\b|dias de pago|dias de proveedores|periodo medio de pago/],
  ['capexPct', /capex.*(ingresos|ventas)|inversion.*(ingresos|ventas)/], ['repay', /pago.*(deuda|capital).*mensual|amortizacion mensual/],
  ['payout', /payout|reparto de dividendos|dividendos.*utilidad/], ['kd', /costo de la? deuda|tasa de interes de la deuda/], ['tax', /tasa de impuesto/]
];
const DRV_PCT = new Set(['capexPct', 'payout', 'kd', 'tax']);
// Devuelve el balance del último período (bal), el período anterior si existe (prev) y los supuestos de proyección (drv).
function parseBalance(rows) {
  const bal = {}, prev = {}, drv = {}; let labels = null, nPeriods = 1;
  rows.forEach(r => {
    if (!r || r[0] == null) return;
    const l = norm(r[0]); if (!l) return;
    const nums = r.slice(1).filter(isNum);
    if (nums.length >= 2 && nums.every(v => v >= 1990 && v <= 2100 && Math.floor(v) === v) && !BS_KEYS.some(([k, re]) => re.test(l))) { labels = nums.map(String); return; }
    if (!nums.length) { const st = r.slice(1).filter(x => typeof x === 'string' && x.trim() && !/^(opcional|defecto|en blanco)/i.test(x.trim())); if (st.length >= 2 && !labels && !/^total/.test(l)) labels = st.map(x => x.trim()); return; }
    if (/^total|^subtotal|^patrimonio$|^pasivos?( corrientes?| no corrientes?)?$|^activos?( corrientes?| no corrientes?)?$/.test(l)) return;
    const d = DRV_KEYS.find(([k, re]) => re.test(l) && drv[k] === undefined);
    if (d) { const v0 = nums[0]; drv[d[0]] = (DRV_PCT.has(d[0]) && Math.abs(v0) > 1) ? v0 / 100 : v0; return; }
    const hit = BS_KEYS.find(([k, re]) => re.test(l)); if (!hit) return;
    const k = hit[0]; let v = nums[nums.length - 1], pv = nums.length > 1 ? nums[nums.length - 2] : undefined;
    if (/depreciacion acumulada/.test(l)) { v = -Math.abs(v); if (pv !== undefined) pv = -Math.abs(pv); }
    if (nums.length > nPeriods) nPeriods = nums.length;
    if (BS_SUM.has(k)) { bal[k] = (bal[k] || 0) + v; if (pv !== undefined) prev[k] = (prev[k] || 0) + pv; }
    else if (bal[k] === undefined) { bal[k] = v; if (pv !== undefined) prev[k] = pv; }
  });
  return { bal, prev, drv, labels, nPeriods };
}

// Estado de flujo de caja genérico (períodos en columnas)
const CF_KEYS = [
  ['cash0', /^saldo inicial.*(caja|efectivo)|^caja inicial|^efectivo inicial/], ['cashEnd', /^saldo final.*(caja|efectivo)|^caja final|^efectivo final/],
  ['cfo', /^flujo (neto )?(operacional|operativo|de operacion|de caja operacional|de caja operativo)|^(total )?(flujo|efectivo) (neto )?(de|por) actividades de operacion|^flujo de actividades de operacion/],
  ['cfi', /^flujo (neto )?(de inversion|de caja de inversion)|^(total )?(flujo|efectivo) (neto )?(de|por) actividades de inversion|^flujo de actividades de inversion/],
  ['cff', /^flujo (neto )?(financiero|de financiamiento|de caja financiero)|^(total )?(flujo|efectivo) (neto )?(de|por) actividades de financiamiento|^flujo de actividades de financiamiento/],
  ['net', /^variacion neta|^flujo neto (del periodo|de caja|total)|^aumento.*neto|^(\(?[+=]\)? ?)?flujo de caja neto|^flujo neto$/],
  ['ni', /^utilidad neta|^resultado neto|^resultado del (periodo|ejercicio)|^utilidad del (periodo|ejercicio)/], ['dep', /depreciacion|amortizacion de intangibles|^d&a/],
  ['dwc', /(variacion|cambio|aumento|disminucion|incremento).*(capital de trabajo|cuentas por cobrar|clientes|inventario|existencias|proveedores|cuentas por pagar)|^\(?[+-]?\)? ?(cuentas por cobrar|inventarios?|proveedores|cuentas por pagar)/],
  ['capex', /capex|compra.*(activo|propiedad|maquinaria|equipo)|adquisicion.*(activo|propiedad|equipo)|inversion en (activo|propiedad|maquinaria)/],
  ['borrow', /(obtencion|nuevo|nueva|desembolso|ingreso).*(deuda|prestamo|credito|financiamiento)|prestamos obtenidos|aumento de capital|aporte/], ['repay', /(pago|amortizacion|abono|servicio).*(deuda|prestamo|capital|credito)|prestamos pagados/],
  ['div', /dividendo|retiro de socios|retiros/], ['int', /interes|gastos? financieros?/], ['tax', /impuesto/], ['rev', /^ingresos|^cobros.*clientes|^ventas/]
];
function parseCFStatement(rows) {
  const out = {}, lines = []; let labels = null, maxCols = 0;
  rows.forEach(r => {
    if (!r || r[0] == null) return;
    const l = norm(r[0]); if (!l) return;
    const nums = r.slice(1).filter(isNum);
    if (!nums.length) { const st = r.slice(1).filter(x => typeof x === 'string' && x.trim()); if (st.length >= 2 && !labels) labels = st.map(x => x.trim()); return; }
    if (nums.length >= 2 && nums.every(v => v >= 1990 && v <= 2100 && Math.floor(v) === v) && !CF_KEYS.some(([k, re]) => re.test(l))) { labels = nums.map(String); return; }
    const hit = CF_KEYS.find(([k, re]) => re.test(l)); if (!hit) return;
    const k = hit[0]; maxCols = Math.max(maxCols, nums.length);
    const arr = nums.slice();
    if (['dwc', 'capex', 'borrow', 'repay', 'div', 'int', 'tax', 'dep', 'rev'].includes(k)) { out[k] = out[k] ? out[k].map((v, i) => v + (arr[i] || 0)).concat(arr.slice(out[k].length)) : arr; }
    else if (!out[k]) out[k] = arr;
    lines.push({ key: k, label: String(r[0]).trim() });
  });
  return { data: out, lines, labels, nCols: maxCols, hasData: Object.keys(out).length >= 2 && (out.cfo || out.ni || out.net || out.cff || out.capex) };
}

/* ============================================================
   3. INGESTA — acumula pestañas de uno o varios archivos
   ============================================================ */
AM.newStore = function () {
  return { pl: { servicios: {}, productiva: {} }, ev: {}, balance: null, balancePrev: null, balLabels: null, balPeriods: 0, cfs: null, drivers: null, ma: null, inventory: [], files: [] };
};

// sheets: [{name, rows, uncalc}]
AM.ingest = function (store, fileName, sheets) {
  store.inventory = store.inventory.filter(i => i.file !== fileName);
  store.files = store.files.filter(f => f !== fileName); store.files.push(fileName);
  sheets.forEach(sh => {
    const rows = sh.rows || [];
    const cls = classifySheet(sh.name, rows);
    const e = { file: fileName, sheet: sh.name, role: cls.role, status: 'usada', detail: '' };
    try {
      switch (cls.role) {
        case 'pl': {
          const p = parsePLSheet(rows);
          if (!p.valid) { e.status = 'ignorada'; e.detail = 'No se encontraron filas de ingresos/utilidad'; break; }
          store.pl[cls.empresa][cls.esc] = p;
          e.detail = `${cls.empresa} · ${cls.esc}` + (p.hasData ? '' : ' · sin datos (todo en cero)');
          if (!p.hasData) e.status = 'vacía';
          break;
        }
        case 'eval_panel': store.ev.panel = parsePanel(rows); break;
        case 'eval_supuestos': store.ev.supuestos = parseSupuestos(rows); break;
        case 'eval_ingresos': store.ev.ingresos = parseSeriesSheet(rows); break;
        case 'eval_costos': store.ev.costos = parseSeriesSheet(rows); break;
        case 'eval_er': store.ev.er = parseSeriesSheet(rows); break;
        case 'eval_fin': store.ev.fin = parseSeriesSheet(rows); break;
        case 'eval_cf': store.ev.cf = parseSeriesSheet(rows); break;
        case 'eval_resumen': store.ev.resumen = parseResumen(rows); break;
        case 'eval_esc': store.ev.esc = parseEscenarios(rows); if (!store.ev.esc) { e.status = 'ignorada'; e.detail = 'Sin columnas Pesimista/Base/Optimista'; } break;
        case 'eval_val': store.ev.val = parseValoracion(rows); break;
        case 'eval_ratios': e.detail = 'Se recalculan en la app'; e.status = 'recalculada'; break;
        case 'balance': {
          const pb = parseBalance(rows), bal = pb.bal, drv = pb.drv;
          store.balance = Object.keys(bal).length ? bal : null; store.drivers = Object.keys(drv).length ? drv : null;
          store.balancePrev = store.balance && Object.keys(pb.prev).length ? pb.prev : null; store.balLabels = pb.labels; store.balPeriods = pb.nPeriods;
          if (!store.balance && !store.drivers) { e.status = 'vacía'; e.detail = 'Plantilla sin completar (opcional): el balance se estima desde el P&L'; }
          else e.detail = [store.balance ? (pb.nPeriods > 1 ? pb.nPeriods + ' períodos de balance' : 'saldos de balance') : '', store.drivers ? 'supuestos de proyección' : ''].filter(Boolean).join(' + ');
          break;
        }
        case 'cf_stmt': {
          const c = parseCFStatement(rows);
          if (!c.hasData) { e.status = 'vacía'; e.detail = 'Sin cifras de flujo reconocibles'; break; }
          store.cfs = c; e.detail = Object.keys(c.data).length + ' conceptos · ' + c.nCols + ' período(s)';
          break;
        }
        case 'ma': {
          const m = parseMA(rows); store.ma = Object.keys(m).length ? m : null;
          if (!store.ma) { e.status = 'vacía'; e.detail = 'Plantilla sin completar (opcional): M&A se autocompleta desde el modelo'; } else e.detail = Object.keys(m).length + ' parámetros';
          break;
        }
        case 'info': e.status = 'informativa'; break;
        default: e.status = 'no reconocida'; e.detail = 'Nombre/estructura no coincide con ningún módulo';
      }
    } catch (err) { e.status = 'error'; e.detail = String(err && err.message || err); }
    if (sh.uncalc) { e.detail += (e.detail ? ' · ' : '') + `${sh.uncalc} fórmulas sin valor calculado`; store.uncalc = (store.uncalc || 0) + sh.uncalc; }
    store.inventory.push(e);
  });
  return store;
};

// Convierte un libro de SheetJS en [{name, rows, uncalc}] (rows = matriz de valores calculados)
AM.sheetsFromWorkbook = function (XLSX, wb) {
  return wb.SheetNames.map(n => {
    const ws = wb.Sheets[n]; let uncalc = 0;
    for (const a in ws) { if (a[0] === '!') continue; const c = ws[a]; if (c && c.f && c.v === undefined) uncalc++; }
    return { name: n, rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }), uncalc };
  });
};

AM.hasPL = store => ['servicios', 'productiva'].some(t => ['real', 'budget', 'forecast'].some(e => store.pl[t][e] && store.pl[t][e].hasData));

/* (continúa en model-engine: proyecto, negocio derivado, M&A, insights) */
AM._internal = { parsePLSheet, parseSeriesSheet, byMonth, irrBisect, monthlyRate, npvAt, pairsOf, pickPair, MESES, norm, num, sum, avg, lastOf, zeros, clamp, isNum };
return AM;
});

/* ============================================================
   AMBERES INTELLIGENCE™ — Motor de cálculo
   · Réplica del modelo de evaluación de proyectos (valida contra el Excel)
   · Derivación de flujo de caja (op/inv/fin) y balance desde el P&L
   · M&A: WACC, DCF, múltiplos, sinergias, sensibilidad
   · Alertas, semáforos y recomendaciones
   ============================================================ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./model.js'));
  else factory(root.AmberesModel);
})(typeof self !== 'undefined' ? self : this, function (AM) {
'use strict';
const U = AM._internal;
const { sum, avg, lastOf, zeros, clamp, isNum, irrBisect, monthlyRate, npvAt, byMonth, MESES } = U;
const maxAbs = a => a.reduce((m, v) => Math.max(m, Math.abs(v)), 0);

/* ============================================================
   1. MOTOR DE PROYECTO (réplica del Excel de evaluación)
   ============================================================ */
AM.projectEngine = function (s) {
  const H = Math.round(s.horizon || 60);
  const lifeM = Math.max(1, s.life || 10) * 12;
  const depM = (s.capex || 0) / lifeM;
  const r = (s.rate || 0) / 12, n = Math.round(s.term || 0), P = s.loan || 0;
  const pay = (P > 0 && n > 0) ? (r > 0 ? P * r / (1 - Math.pow(1 + r, -n)) : P / n) : 0;
  const o = { H, rev: zeros(H + 1), cvCost: zeros(H + 1), fixed: zeros(H + 1), ebitda: zeros(H + 1), dep: zeros(H + 1), ebit: zeros(H + 1),
    interest: zeros(H + 1), amort: zeros(H + 1), saldo: zeros(H + 1), ebt: zeros(H + 1), taxEbit: zeros(H + 1), taxEbt: zeros(H + 1),
    nopat: zeros(H + 1), ni: zeros(H + 1), resid: zeros(H + 1), fcl: zeros(H + 1), fcfe: zeros(H + 1) };
  let bal = P;
  o.saldo[0] = P;
  o.fcl[0] = -((s.capex || 0) + (s.nwc0 || 0) + (s.preop || 0));
  o.fcfe[0] = o.fcl[0] + P;
  for (let t = 1; t <= H; t++) {
    const rev = (s.ingMes || 0) * Math.pow(1 + (s.g || 0), Math.floor((t - 1) / 12));
    o.rev[t] = rev; o.cvCost[t] = rev * (s.cv || 0); o.fixed[t] = s.fixed || 0;
    o.ebitda[t] = rev - o.cvCost[t] - o.fixed[t];
    o.dep[t] = t <= lifeM ? depM : 0;
    o.ebit[t] = o.ebitda[t] - o.dep[t];
    if (t <= n) { o.interest[t] = bal * r; o.amort[t] = pay - o.interest[t]; bal -= o.amort[t]; }
    o.saldo[t] = Math.max(0, bal);
    o.ebt[t] = o.ebit[t] - o.interest[t];
    o.taxEbit[t] = Math.max(0, o.ebit[t]) * (s.tax || 0);
    o.taxEbt[t] = Math.max(0, o.ebt[t]) * (s.tax || 0);
    o.nopat[t] = o.ebit[t] - o.taxEbit[t];
    o.ni[t] = o.ebt[t] - o.taxEbt[t];
    o.resid[t] = (t === H) ? (s.resid || 0) * (s.capex || 0) : 0;
    o.fcl[t] = o.nopat[t] + o.dep[t] + o.resid[t];
    o.fcfe[t] = o.ni[t] + o.dep[t] - o.amort[t] + o.resid[t];
  }
  const rm = monthlyRate(s.wacc != null ? s.wacc : 0.12);
  o.van = npvAt(o.fcl, rm);
  o.tirM = irrBisect(o.fcl);
  o.tirA = o.tirM == null ? null : Math.pow(1 + o.tirM, 12) - 1;
  let acc = 0; o.payback = null; o.acc = o.fcl.map(v => (acc += v));
  for (let t = 0; t <= H; t++) if (o.acc[t] >= 0 && t > 0) { o.payback = t; break; }
  o.tirFinM = irrBisect(o.fcfe); o.tirFinA = o.tirFinM == null ? null : Math.pow(1 + o.tirFinM, 12) - 1;
  return o;
};
const engineComplete = s => !!s && ['capex', 'ingMes', 'g', 'cv', 'fixed'].every(k => isNum(s[k]));

AM.tornado = function (sup, pct) {
  pct = pct || 0.10;
  const base = AM.projectEngine(sup).van;
  const drivers = [
    ['Ingresos', k => ({ ingMes: sup.ingMes * k })], ['% Costos variables', k => ({ cv: sup.cv * k })],
    ['Costos fijos', k => ({ fixed: sup.fixed * k })], ['Inversión (CAPEX)', k => ({ capex: sup.capex * k })],
    ['WACC', k => ({ wacc: sup.wacc * k })], ['Tasa de impuesto', k => ({ tax: sup.tax * k })]
  ];
  return drivers.map(([label, f]) => {
    const lo = AM.projectEngine(Object.assign({}, sup, f(1 - pct))).van;
    const hi = AM.projectEngine(Object.assign({}, sup, f(1 + pct))).van;
    return { label, down: lo - base, up: hi - base, swing: Math.abs(hi - lo) };
  }).sort((a, b) => b.swing - a.swing);
};

/* ============================================================
   2. SUPUESTOS POR DEFECTO (todos editables en la app)
   ============================================================ */
AM.defaultAssumptions = function () {
  return { empresa: 'auto', escenario: 'auto', tax: null, dso: 30, dio: null, dpo: null, capexPct: null, cash0: null, ppe0: null,
    debt0: null, repay: 0, payout: 0, kd: 0.075,
    revAnnual: null, ebitdaMargin: 0.15, cogsPct: 0.6, life: 10, balMonths: 12 };
};

/* ============================================================
   3. NEGOCIO DERIVADO: flujo de caja + balance (3 estados cuadrados)
   ============================================================ */
function deriveBusiness(ser, A, open, nwcFixed) {
  const N = ser.rev.length;
  open = open || {};
  const tax = A.tax;
  const cashCosts = ser.cogs.map((c, i) => c + ser.opex[i] - ser.dep[i]);
  let AR, INV, AP;
  const ovr = ser.ovr || null;
  if (ovr && ovr.dWC) { let a = open.ar != null ? open.ar : 0; AR = ovr.dWC.map(w => (a += w)); INV = zeros(N); AP = zeros(N); }
  else if (nwcFixed != null) { AR = zeros(N).map(() => nwcFixed); INV = zeros(N); AP = zeros(N); }
  else {
    AR = ser.rev.map(r => Math.max(0, r) * A.dso / 30);
    INV = (ser.invFinal && ser.invFinal.some(v => v)) ? ser.invFinal.slice(0, N) : ser.cogs.map(c => Math.max(0, c) * A.dio / 30);
    AP = cashCosts.map(c => Math.max(0, c) * A.dpo / 30);
  }
  const AR0 = open.ar != null ? open.ar : AR[0];
  const INV0 = open.inv != null ? open.inv : INV[0];
  const AP0 = open.ap != null ? open.ap : AP[0];
  const PPE0 = open.ppe != null ? open.ppe : (A.ppe0 || 0);
  const cash0 = open.cash != null ? open.cash : (A.cash0 || 0);
  const notes = [];
  let debt0 = open.debt != null ? open.debt : (A.debt0 != null ? A.debt0 : null);
  if (debt0 == null) {
    const annFin = sum(ser.finExp) * 12 / N;
    if (annFin > 0 && A.kd > 0) { debt0 = annFin / A.kd; notes.push({ key: 'debt0', label: 'Deuda inicial', value: debt0, source: 'implícita', detail: `gastos financieros anualizados ÷ Kd ${(A.kd * 100).toFixed(1)}%` }); }
    else debt0 = 0;
  }
  if (ovr && ovr.repay) { let mn = 0, c = debt0; for (let t = 0; t < N; t++) { c += (ovr.borrow ? ovr.borrow[t] : 0) - ovr.repay[t]; mn = Math.min(mn, c); } if (mn < 0) { debt0 -= mn; notes.push({ key: 'debt0', label: 'Deuda inicial', value: debt0, source: 'implícita', detail: 'mínimo necesario para sostener los pagos de deuda del flujo' }); } }
  const capexPct = A.capexPct != null ? A.capexPct : (sum(ser.rev) > 0 ? sum(ser.dep) / sum(ser.rev) : 0);
  const capexArr = ser.capex || ser.rev.map(r => capexPct * r);

  const eq0 = cash0 + AR0 + INV0 + PPE0 - AP0 - debt0;
  let capital0, re0, adj0 = 0;
  if (open.capital != null || open.re != null) { capital0 = open.capital || 0; re0 = open.re || 0; adj0 = eq0 - capital0 - re0; }
  else { capital0 = eq0; re0 = 0; }

  const bs = { cash: [cash0], ar: [AR0], inv: [INV0], ppe: [PPE0], ap: [AP0], debt: [debt0], re: [re0] };
  const cfo = [], cfi = [], cff = [], net = [], fcff = [], fcfe = [], div = [], repayA = [], dNWC = [], taxEbitA = [];
  for (let t = 0; t < N; t++) {
    const dAR = AR[t] - bs.ar[t], dINV = INV[t] - bs.inv[t], dAP = AP[t] - bs.ap[t];
    const wc = (ovr && ovr.dWC) ? ovr.dWC[t] : dAR + dINV - dAP;
    const ni = ser.ni[t], dep = ser.dep[t];
    const c_o = ni + dep - wc;
    const capex = capexArr[t] || 0;
    const rp = (ovr && ovr.repay) ? ovr.repay[t] : Math.min(A.repay || 0, bs.debt[t]);
    const dv = (ovr && ovr.div) ? ovr.div[t] : (A.payout || 0) * Math.max(ni, 0);
    const nd = (ovr && ovr.borrow) ? ovr.borrow[t] : 0;
    const c_i = -capex, c_f = -rp - dv + nd;
    cfo.push(c_o); cfi.push(c_i); cff.push(c_f); net.push(c_o + c_i + c_f);
    div.push(dv); repayA.push(rp); dNWC.push(wc);
    const tEbit = Math.max(0, ser.ebit[t]) * tax; taxEbitA.push(tEbit);
    fcff.push(ser.ebit[t] - tEbit + dep - wc - capex);
    fcfe.push(c_o + c_i - rp + nd);
    bs.cash.push(bs.cash[t] + c_o + c_i + c_f);
    bs.ar.push(AR[t]); bs.inv.push(INV[t]); bs.ap.push(AP[t]);
    bs.ppe.push(bs.ppe[t] + capex - dep);
    bs.debt.push(bs.debt[t] - rp + nd);
    bs.re.push(bs.re[t] + ni - dv);
  }
  bs.assets = bs.cash.map((c, i) => c + bs.ar[i] + bs.inv[i] + bs.ppe[i]);
  bs.liab = bs.ap.map((a, i) => a + bs.debt[i]);
  bs.capital = bs.cash.map(() => capital0);
  bs.adj = bs.cash.map(() => adj0);
  bs.equity = bs.re.map((r, i) => capital0 + adj0 + r);
  bs.check = bs.assets.map((a, i) => a - bs.liab[i] - bs.equity[i]);
  bs.nwc = bs.ar.map((a, i) => a + bs.inv[i] - bs.ap[i]);

  // ratios con anualización de los últimos ≤12 meses
  const w = Math.min(12, N), ann = arr => sum(arr.slice(N - w)) * 12 / w;
  const ebitda = ser.ebit.map((e, i) => e + ser.dep[i]);
  const ebitdaAnn = ann(ebitda), niAnn = ann(ser.ni), finAnn = ann(ser.finExp), revAnn = ann(ser.rev);
  const endI = N, avgEq = avg(bs.equity.slice(Math.max(0, N - w), N + 1)), avgAs = avg(bs.assets.slice(Math.max(0, N - w), N + 1));
  const cashCostM = avg(cashCosts.slice(Math.max(0, N - w)));
  const ratios = {
    currentRatio: bs.ap[endI] > 0 ? (bs.cash[endI] + bs.ar[endI] + bs.inv[endI]) / bs.ap[endI] : null,
    netDebt: bs.debt[endI] - bs.cash[endI],
    ndEbitda: ebitdaAnn > 0 ? (bs.debt[endI] - bs.cash[endI]) / ebitdaAnn : null,
    coverage: finAnn > 0 ? ebitdaAnn / finAnn : null,
    runway: cashCostM > 0 ? bs.cash[endI] / cashCostM : null,
    roe: (open.capital != null || open.re != null || open.cash != null) && avgEq > 0 ? niAnn / avgEq : null, roa: (open.capital != null || open.re != null || open.cash != null) && avgAs > 0 ? niAnn / avgAs : null,
    gearing: (bs.debt[endI] + bs.equity[endI]) > 0 ? bs.debt[endI] / (bs.debt[endI] + bs.equity[endI]) : null,
    nwcPctRev: revAnn > 0 ? bs.nwc[endI] / revAnn : null,
    ccc: nwcFixed != null ? null : (A.dso + A.dio - A.dpo),
    ebitdaAnn, niAnn, revAnn, finAnn,
    cfoToEbitda: sum(ebitda) !== 0 ? sum(cfo) / sum(ebitda) : null,
    maxCheck: maxAbs(bs.check)
  };
  return { N, cfo, cfi, cff, net, fcff, fcfe, div, repay: repayA, dNWC, taxEbit: taxEbitA, bs, ratios, notes, capexPct, debt0, cash0, eq0, ebitda };
}

/* ============================================================
   4. M&A
   ============================================================ */
AM.maDefaults = function () {
  return { ingBase: 1e9, growth: 0.05, margin: 0.20, da: 0.03, capex: 0.04, nwc: 0.08, tax: 0.27, years: 5, termG: 0.025,
    rf: 0.05, beta: 1.10, erp: 0.065, country: 0.015, kd: 0.075, de: 0.35, mult: 7.5, netDebt: 0,
    synCost: 0, synRev: 0, synYears: 3, synOneoff: 0, premium: 0.20 };
};
AM.calcWACC = function (p) {
  const ke = p.rf + p.beta * p.erp + p.country, kdAT = p.kd * (1 - p.tax), we = 1 - p.de;
  return { wacc: we * ke + p.de * kdAT, ke, kdAT, we, wd: p.de };
};
function projectMA(p) {
  const rows = []; let prev = p.ingBase;
  for (let y = 1; y <= p.years; y++) {
    const rev = p.ingBase * Math.pow(1 + p.growth, y), ebitda = rev * p.margin, da = rev * p.da, ebit = ebitda - da;
    const taxAmt = Math.max(ebit, 0) * p.tax, nopat = ebit - taxAmt, capex = rev * p.capex, nwc = (rev - prev) * p.nwc;
    rows.push({ y, rev, ebitda, da, ebit, taxAmt, nopat, capex, nwc, fcf: nopat + da - capex - nwc }); prev = rev;
  }
  return rows;
}
function dcfEV(rows, wacc, termG) {
  let pv = 0; rows.forEach(r => { pv += r.fcf / Math.pow(1 + wacc, r.y); });
  const g = Math.min(termG, wacc - 0.005);
  const tv = (lastOf(rows).fcf * (1 + g)) / (wacc - g);
  const pvTV = tv / Math.pow(1 + wacc, rows.length);
  return { pvSum: pv, tv, pvTV, ev: pv + pvTV, g };
}
AM.maCompute = function (p) {
  const rows = projectMA(p), w = AM.calcWACC(p), dcf = dcfEV(rows, w.wacc, p.termG);
  const equity = dcf.ev - p.netDebt;
  const evMult = rows[0].ebitda * p.mult;
  let pvSyn = 0; const syn = [];
  for (let y = 1; y <= p.years; y++) {
    const frac = Math.min(1, y / p.synYears), pre = (p.synCost + p.synRev) * frac, at = pre * (1 - p.tax);
    pvSyn += at / Math.pow(1 + w.wacc, y); syn.push({ y, cost: p.synCost * frac, rev: p.synRev * frac, afterTax: at });
  }
  const lastSyn = syn.length ? lastOf(syn).afterTax : 0, gS = Math.min(p.termG, w.wacc - 0.005);
  const pvTVs = (lastSyn * (1 + gS)) / (w.wacc - gS) / Math.pow(1 + w.wacc, p.years);
  const npvSyn = pvSyn + pvTVs - p.synOneoff;
  const evComb = dcf.ev + npvSyn;
  const equityControl = equity * (1 + p.premium);
  const lo = dcfEV(rows, Math.max(0.01, w.wacc - 0.005), p.termG).ev, hi = dcfEV(rows, w.wacc + 0.005, p.termG).ev;
  return { rows, wacc: w, dcf, equity, evMult, syn, npvSyn, evComb, equityControl, evControl: equityControl + p.netDebt,
    football: [
      { label: 'DCF (WACC ±50pb)', lo: Math.min(lo, hi), hi: Math.max(lo, hi) },
      { label: 'Múltiplos comparables (±1,0x)', lo: rows[0].ebitda * Math.max(1, p.mult - 1), hi: rows[0].ebitda * (p.mult + 1) },
      { label: 'DCF + sinergias (±4%)', lo: evComb * 0.96, hi: evComb * 1.04 }
    ] };
};
AM.maGrid = function (p) {
  const rows = projectMA(p), w0 = AM.calcWACC(p).wacc;
  const waccs = [-0.02, -0.01, 0, 0.01, 0.02].map(d => Math.max(0.02, w0 + d));
  const gs = [-0.01, -0.005, 0, 0.005, 0.01].map(d => Math.max(0, p.termG + d));
  return { waccs, gs, cells: waccs.map(w => gs.map(g => dcfEV(rows, w, g).ev)) };
};
AM.maInsights = function (p, c, src) {
  const out = [], f = AM.fmt;
  const w = c.wacc.wacc;
  const spread = p.growth - w;
  const tvShare = c.dcf.ev ? c.dcf.pvTV / c.dcf.ev : 0;
  out.push(c.equity > 0
    ? { level: 'success', title: `Valor patrimonial DCF positivo (${f(c.equity)})`, text: `EV ${f(c.dcf.ev)} menos deuda neta ${f(p.netDebt)}. Rango de referencia por múltiplos: ${f(c.football[1].lo)} – ${f(c.football[1].hi)}.` }
    : { level: 'danger', title: 'Valor patrimonial DCF negativo', text: `El EV (${f(c.dcf.ev)}) no cubre la deuda neta (${f(p.netDebt)}). Revisa margen, crecimiento y capex; en una compra, negocia el precio a la baja o estructura earn-out.` });
  out.push(tvShare > 0.75
    ? { level: 'warning', title: `El valor terminal pesa ${(tvShare * 100).toFixed(0)}% del EV`, text: 'Valoración muy sensible a g y WACC. Recomendación: respaldar con comparables y mostrar la tabla de sensibilidad al comité.' }
    : { level: 'success', title: `Valor terminal razonable (${(tvShare * 100).toFixed(0)}% del EV)`, text: 'La valoración depende en mayor medida de los flujos explícitos.' });
  const gap = c.evMult && c.dcf.ev ? (c.evMult - c.dcf.ev) / c.dcf.ev : 0;
  out.push(Math.abs(gap) > 0.35
    ? { level: 'warning', title: `DCF y múltiplos difieren ${(gap * 100).toFixed(0)}%`, text: 'Diferencias sobre ±35% suelen indicar un múltiplo desalineado o supuestos de margen/capex poco realistas. Triangula antes de fijar un precio.' }
    : { level: 'success', title: 'DCF y múltiplos consistentes', text: `Brecha de ${(gap * 100).toFixed(0)}% entre ambos métodos.` });
  out.push(p.termG > w - 0.02
    ? { level: 'danger', title: 'g terminal demasiado cercano al WACC', text: 'El valor terminal se dispara cuando g se acerca al WACC. Mantén g ≤ crecimiento de largo plazo de la economía (2–3%).' }
    : { level: 'info', title: `WACC ${(w * 100).toFixed(2)}% · Ke ${(c.wacc.ke * 100).toFixed(1)}% · Kd(AT) ${(c.wacc.kdAT * 100).toFixed(1)}%`, text: 'Si el beta o el riesgo país no provienen de una fuente de mercado, trátalos como supuestos y sensibilízalos.' });
  if (p.synCost + p.synRev > 0) out.push(c.npvSyn > 0
    ? { level: 'success', title: `Sinergias crean ${f(c.npvSyn)} de valor neto`, text: `EV combinado ${f(c.evComb)}. Paga a lo sumo una fracción de las sinergias al vendedor (prima de control ${(p.premium * 100).toFixed(0)}% equivale a ${f(c.equity * p.premium)}).` }
    : { level: 'warning', title: 'Sinergias no cubren el costo de integración', text: 'Revisa el ramp-up o el costo de integración antes de pagar prima por control.' });
  if (src) { const sup = Object.keys(src).filter(k => src[k] === 'Supuesto'); if (sup.length) out.push({ level: 'info', title: 'Parámetros por defecto en uso', text: `Valores de mercado/estructura sin dato en tu archivo: ${sup.map(k => MA_LABELS[k] || k).join(', ')}. Reemplázalos por datos reales (en la pestaña M&A del archivo o con los controles).` }); }
  return out;
};
const MA_LABELS = { ingBase: 'ingresos base', growth: 'crecimiento', margin: 'margen EBITDA', da: 'D&A', capex: 'capex', nwc: 'capital de trabajo', tax: 'impuesto', years: 'años de proyección', termG: 'g terminal', rf: 'tasa libre de riesgo', beta: 'beta', erp: 'prima de riesgo de mercado', country: 'riesgo país', kd: 'costo de deuda', de: 'estructura de capital', mult: 'múltiplo EV/EBITDA', netDebt: 'deuda neta', synCost: 'sinergias de costos', synRev: 'sinergias de ingresos', synYears: 'años de sinergias', synOneoff: 'costo de integración', premium: 'prima de control' };
AM.MA_LABELS = MA_LABELS;

/* ============================================================
   5. FORMATO
   ============================================================ */
AM.fmt = function (v) {
  if (v === undefined || v === null || !isFinite(v)) return '—';
  const a = Math.abs(v), sg = v < 0 ? '-' : '';
  if (a >= 1e9) return sg + '$' + (a / 1e9).toFixed(2) + 'B';
  if (a >= 1e6) return sg + '$' + (a / 1e6).toFixed(2) + 'M';
  if (a >= 1e3) return sg + '$' + (a / 1e3).toFixed(1) + 'K';
  return sg + '$' + a.toFixed(0);
};
const fmt = AM.fmt;
const pc = (v, d) => (v == null || !isFinite(v)) ? '—' : (v * 100).toFixed(d == null ? 1 : d) + '%';

/* ============================================================
   6. CONSTRUCCIÓN DEL MODELO UNIFICADO
   ============================================================ */
function pickPL(store, A) {
  const esc3 = ['real', 'budget', 'forecast'];
  const withData = t => esc3.filter(e => store.pl[t][e] && store.pl[t][e].hasData);
  let tipo = A.empresa !== 'auto' && withData(A.empresa).length ? A.empresa : null;
  if (!tipo) {
    const score = t => withData(t).reduce((s, e) => s + Math.abs(sum(store.pl[t][e].ingresosNetos.slice(0, 12))), 0);
    tipo = score('servicios') >= score('productiva') ? 'servicios' : 'productiva';
    if (!withData(tipo).length) tipo = null;
  }
  if (!tipo) return null;
  const avail = withData(tipo);
  const esc = (A.escenario !== 'auto' && avail.includes(A.escenario)) ? A.escenario : (avail.includes('real') ? 'real' : avail[0]);
  return { tipo, esc, pl: store.pl[tipo][esc], avail };
}

function serFromPL(pl) {
  const take = k => (pl[k] || zeros(13)).slice(0, 12);
  const rev = take('ingresosNetos'), ebitRaw = pl.ebit ? take('ebit') : null;
  const cogs = take('costoTotal'), opex = take('gastosOp');
  const ebit = ebitRaw || rev.map((r, i) => r - cogs[i] - opex[i]);
  const finExp = take('gastosFin').map(Math.abs), finInc = take('ingFin');
  const dep = take('depreciacion');
  const tax = take('impuesto').map(Math.abs);
  const ni = pl.utilidadNeta ? take('utilidadNeta') : ebit.map((e, i) => e + finInc[i] - finExp[i] - tax[i]);
  return { labels: MESES.slice(), rev, cogs, opex, dep, ebit, finExp, finInc, tax, ni, invFinal: pl.invFinal ? take('invFinal') : null };
}


/* ============================================================
   3b. ESTADOS SINTÉTICOS: reconstruir un P&L estimado a partir de un Balance o de un Flujo de caja
   ============================================================ */
const norm12 = (arr) => {
  if (!arr || !arr.length) return null;
  if (arr.length >= 12) return arr.slice(0, 12);
  const v = arr[arr.length - 1] / 12; return Array.from({ length: 12 }, () => v);   // cifra anual → se reparte en 12 meses
};
function synthFromBalance(store, A) {
  const bal = store.balance, prev = store.balancePrev, notes = [];
  if (!bal) return null;
  const t = clamp(A.tax != null ? A.tax : 0.27, 0, 0.5), life = Math.max(1, A.life || 10), pm = Math.max(1, A.balMonths || 12);
  const Ab = {};
  // 1) ingresos anuales: dato del usuario o implícitos del saldo de clientes (DSO)
  let revA = isNum(A.revAnnual) && A.revAnnual > 0 ? A.revAnnual : null, revSrc = 'ingresado por ti';
  if (!revA && bal.ar > 0 && A.dso > 0) { revA = bal.ar * 365 / A.dso; revSrc = `clientes ÷ DSO ${A.dso} días`; }
  if (!revA && bal.ap > 0 && A.dpo > 0) { revA = bal.ap * 365 / A.dpo / (1 - A.cogsPct * 0.5); revSrc = 'proveedores ÷ DPO (muy aproximado)'; }
  const hasRev = !!revA; if (!revA) revA = 0;
  notes.push({ level: hasRev ? 'info' : 'warning', title: hasRev ? 'Ingresos anuales estimados' : 'No hay datos para estimar ingresos', text: hasRev ? `Ingresos ≈ ${fmt(revA)} (${revSrc}). Un balance no contiene ventas: ingresa "Ingresos anuales" en Supuestos para afinar el P&L, el flujo y M&A.` : 'Un balance no informa ventas. Ingresa "Ingresos anuales" en Supuestos para estimar P&L, flujo de caja y valoración.' });
  // 2) depreciación y resultado
  const dep12 = Math.max(0, bal.ppe || 0) / life;                   // anual
  let margin = A.ebitdaMargin, hist = null;
  const finA = Math.max(0, bal.debt || 0) * A.kd;
  if (prev) {
    const dRE = ((bal.re || 0) - (prev.re || 0)), niA = dRE * 12 / pm;
    if (hasRev && niA > 0) {
      const ebt = niA / (1 - t), ebitda = ebt + finA + dep12; margin = clamp(ebitda / revA, 0.02, 0.6);
      notes.push({ level: 'info', title: 'Margen EBITDA calibrado con tu balance', text: `La variación de resultados acumulados (${fmt(dRE)} en ${pm} meses) implica un margen EBITDA ≈ ${(margin * 100).toFixed(1)}%. Supone columnas de balance consecutivas y sin dividendos.` });
    }
    // flujo histórico por diferencias (método indirecto sobre dos balances)
    const d = k => (bal[k] || 0) - (prev[k] || 0), depP = Math.max(0, prev.ppe || 0) / life * pm / 12;
    const ni = d('re'), cfoH = ni + depP - d('ar') - d('inv') + d('ap'), capexH = d('ppe') + depP, cfiH = -capexH;
    const cffH = d('debt') + ((bal.capital || 0) - (prev.capital || 0)), netH = cfoH + cfiH + cffH, dCash = d('cash');
    hist = { periods: store.balLabels && store.balLabels.length >= 2 ? store.balLabels.slice(-2) : ['Anterior', 'Actual'], months: pm,
      rows: [['Resultado del período (Δ resultados acumulados)', ni], ['(+) Depreciación estimada', depP], ['(−) Δ Clientes', -d('ar')], ['(−) Δ Inventarios', -d('inv')], ['(+) Δ Proveedores', d('ap')],
        ['FLUJO OPERACIONAL', cfoH, 1], ['(−) Capex implícito (Δ activo fijo + depreciación)', cfiH], ['FLUJO DE INVERSIÓN', cfiH, 1], ['(+) Δ Deuda financiera', d('debt')], ['(+) Aportes de capital', (bal.capital || 0) - (prev.capital || 0)],
        ['FLUJO FINANCIERO', cffH, 1], ['VARIACIÓN NETA ESTIMADA', netH, 1], ['Variación real de caja (balance)', dCash, 1], ['No explicado (dividendos, otros)', dCash - netH, 1]] };
  }
  // 3) serie mensual plana (sin estacionalidad)
  const rev = zeros(12).map(() => revA / 12), cogs = rev.map(r => r * A.cogsPct), dep = zeros(12).map(() => dep12 / 12);
  const ebitda = rev.map(r => r * margin), ebit = ebitda.map((e, i) => e - dep[i]);
  const finExp = zeros(12).map(() => finA / 12), tax = ebit.map((e, i) => Math.max(0, e - finExp[i]) * t);
  const ni = ebit.map((e, i) => e - finExp[i] - tax[i]), opex = rev.map((r, i) => r - cogs[i] - ebit[i]);
  const ser = { labels: MESES.slice(), rev, cogs, opex, dep, ebit, finExp, finInc: zeros(12), tax, ni, invFinal: null };
  // 4) calibrar DSO/DIO/DPO para que el primer mes no genere un salto de capital de trabajo
  if (hasRev) {
    const cashC = (cogs[0] + opex[0] - dep[0]);
    if (bal.ar > 0) Ab.dso = bal.ar / rev[0] * 30; if (bal.inv > 0 && cogs[0] > 0) Ab.dio = bal.inv / cogs[0] * 30; if (bal.ap > 0 && cashC > 0) Ab.dpo = bal.ap / cashC * 30;
  }
  if (prev && (prev.debt || 0) > (bal.debt || 0)) { Ab.repay = ((prev.debt || 0) - (bal.debt || 0)) / pm; notes.push({ level: 'info', title: 'Amortización de deuda implícita', text: `La deuda bajó ${fmt((prev.debt || 0) - (bal.debt || 0))} entre períodos: se proyecta el mismo ritmo de pago (${fmt(Ab.repay)} al mes).` }); }
  notes.push({ level: 'info', title: 'P&L y flujo proyectados desde tu balance', text: 'Se proyectan 12 meses planos a partir del último balance (sin estacionalidad), con costo de ventas ' + Math.round(A.cogsPct * 100) + '% y margen EBITDA ' + (margin * 100).toFixed(1) + '%. Sube un P&L para reemplazar esta estimación.' });
  return { kind: 'balance', label: 'Balance', ser, open: Object.assign({}, bal), Ab, notes, hist, openFromFile: true, hasRev };
}
function synthFromCF(store, A) {
  const c = store.cfs; if (!c || !c.hasData) return null; const D = c.data, notes = [];
  const t = clamp(A.tax != null ? A.tax : 0.27, 0, 0.5);
  const annual = c.nCols < 12;
  if (annual) notes.push({ level: 'info', title: 'Flujo con períodos anuales', text: `El archivo trae ${c.nCols} período(s); se usa el último y se reparte en 12 meses iguales (sin estacionalidad).` });
  const z = zeros(12), g = k => norm12(D[k]);
  const dep = (g('dep') || z).map(Math.abs);
  let ni = g('ni'), cfo = g('cfo'), cfi = g('cfi'), cff = g('cff'), net = g('net');
  let dwcCash = g('dwc');                                          // efecto en caja (negativo = consume caja)
  if (!dwcCash && cfo && ni) dwcCash = cfo.map((v, i) => v - ni[i] - dep[i]);
  if (!dwcCash) dwcCash = z;
  if (!ni && cfo) ni = cfo.map((v, i) => v - dep[i] - dwcCash[i]);
  if (!ni) ni = z;
  const wc = dwcCash.map(v => -v);                                 // aumento de capital de trabajo
  let capex = g('capex'); capex = capex ? capex.map(Math.abs) : (cfi ? cfi.map(v => Math.max(0, -v)) : z);
  let borrow = g('borrow'), repay = g('repay'), div = g('div');
  const hasFinDetail = !!(borrow || repay || div);
  if (!hasFinDetail && cff) { borrow = cff.map(v => Math.max(0, v)); repay = cff.map(v => Math.max(0, -v)); div = z; notes.push({ level: 'info', title: 'Financiamiento sin detalle', text: 'Solo se informó el flujo financiero total: se interpretó como deuda neta (positivo = nueva deuda, negativo = pago).' }); }
  borrow = (borrow || z).map(Math.abs); repay = (repay || z).map(Math.abs); div = (div || z).map(Math.abs);
  const finExp = (g('int') || z).map(Math.abs);
  let taxA = g('tax'); taxA = taxA ? taxA.map(Math.abs) : ni.map(v => v > 0 ? v * t / (1 - t) : 0);
  const ebit = ni.map((v, i) => v + taxA[i] + finExp[i]);
  const ebitda = ebit.map((e, i) => e + dep[i]);
  let rev = g('rev'), revNote = null;
  if (!rev) {
    if (isNum(A.revAnnual) && A.revAnnual > 0) { rev = z.map(() => A.revAnnual / 12); revNote = 'ingresada por ti'; }
    else { rev = ebitda.map(e => Math.max(e, 0) / A.ebitdaMargin); revNote = `EBITDA ÷ margen supuesto ${(A.ebitdaMargin * 100).toFixed(0)}%`; }
    notes.push({ level: 'warning', title: 'Ingresos estimados', text: `Un flujo de caja no informa ventas completas: se estimaron (${revNote}). Ingresa "Ingresos anuales" en Supuestos o sube el P&L para reemplazar esta estimación.` });
  }
  const cogs = rev.map(r => r * A.cogsPct), opex = rev.map((r, i) => r - cogs[i] - ebit[i]);
  // caja inicial
  let cash0 = null, cashSrc = null;
  if (D.cash0 && D.cash0.length) { cash0 = D.cash0[0]; cashSrc = 'saldo inicial del archivo'; }
  else if (D.cashEnd && D.cashEnd.length) { const ce = norm12(D.cashEnd); const nt = net || ni.map((v, i) => v + dep[i] + dwcCash[i] - capex[i] + borrow[i] - repay[i] - div[i]); cash0 = ce[0] - nt[0]; cashSrc = 'saldo final de caja del archivo'; }
  else if (A.cash0 != null) { cash0 = A.cash0; cashSrc = 'supuesto'; }
  else { cash0 = 0; notes.push({ level: 'info', title: 'Caja inicial supuesta en 0', text: 'El flujo no informa el saldo de caja. Ingrésalo en Supuestos para que el balance parta del saldo real.' }); }
  const ser = { labels: MESES.slice(), rev, cogs, opex, dep, ebit, finExp, finInc: z.slice(), tax: taxA, ni, invFinal: null, capex,
    ovr: { dWC: wc, repay, div, borrow } };
  notes.push({ level: 'info', title: 'P&L y balance reconstruidos desde tu flujo de caja', text: 'Utilidad e impuestos se tomaron del flujo cuando existen; EBIT = utilidad + impuestos + intereses. El balance parte de caja inicial y se ajusta para cuadrar; activos fijos y deuda iniciales son supuestos (editables).' });
  const open = { cash: cash0, ar: 0, inv: 0, ap: 0 };
  if (A.ppe0 != null) open.ppe = A.ppe0; else open.ppe = 0;
  return { kind: 'cf', label: 'Flujo de caja', ser, open, Ab: {}, notes, hist: null, openFromFile: !!cashSrc && cashSrc !== 'supuesto', hasRev: true };
}

function buildProject(ev, A) {
  if (!ev || (!ev.cf && !ev.supuestos)) return null;
  const panel = ev.panel || {}, s = ev.supuestos || {};
  const tax = isNum(panel.tax) ? panel.tax : 0.27, wacc = isNum(panel.wacc) ? panel.wacc : 0.12;
  const sup = Object.assign({}, s, { tax, wacc });
  let H = isNum(panel.horizonte) ? Math.round(panel.horizonte) : 60;
  if (ev.cf) { const len = Math.max(...Object.values(ev.cf.series).map(x => x.vals.length)) + ev.cf.start - 1; if (len > 0) H = Math.min(H, len); }
  sup.horizon = H;
  const canEngine = engineComplete(sup);
  const eng = canEngine ? AM.projectEngine(sup) : null;
  const P = { H, wacc, tax, name: panel.nombre || 'Proyecto', moneda: panel.moneda || '', unidad: panel.unidad || '', sup, canWhatIf: canEngine, engine: eng, source: 'motor' };
  const L = ev.cf ? {
    fcl: byMonth(ev.cf, ['flujo de caja libre'], H), ebitCF: byMonth(ev.cf, ['ebit'], H), taxEbit: byMonth(ev.cf, ['impuestos sobre ebit'], H),
    dep: byMonth(ev.cf, ['(+) depreciacion'], H), capex: byMonth(ev.cf, ['(-) capex'], H), nwc: byMonth(ev.cf, ['(-) cambio en capital'], H),
    preop: byMonth(ev.cf, ['(-) gastos pre-operativos'], H), resid: byMonth(ev.cf, ['(+) valor residual'], H)
  } : {};
  const E = ev.er ? { rev: byMonth(ev.er, ['ingresos por ventas', 'ingresos'], H), ebitda: byMonth(ev.er, ['ebitda'], H), ebit: byMonth(ev.er, ['ebit'], H),
    ni: byMonth(ev.er, ['utilidad neta'], H), cv: byMonth(ev.er, ['costos variables'], H), fixed: byMonth(ev.er, ['costos fijos'], H), dep: byMonth(ev.er, ['depreciacion'], H),
    fin: byMonth(ev.er, ['gastos financieros'], H), taxEbt: byMonth(ev.er, ['impuestos'], H) } : {};
  const F = ev.fin ? { interest: byMonth(ev.fin, ['intereses'], H), amort: byMonth(ev.fin, ['amortizacion de capital'], H), saldo: byMonth(ev.fin, ['saldo final'], H) } : {};

  const has = a => a && a.some(v => v);
  // El FCL se recompone desde sus componentes (NOPAT + D&A − CAPEX − ΔCT − pre-operativos + residual)
  // en lugar de confiar en la fila "Flujo de caja libre" del archivo, que puede estar mal sumada.
  L.nopat = ev.cf ? byMonth(ev.cf, ['nopat'], H) : null;
  const compsOK = has(L.nopat) && has(L.dep) && L.capex && L.nwc && L.preop && L.resid;
  P.fclFile = has(L.fcl) ? L.fcl : null;
  if (compsOK) {
    P.fcl = L.nopat.map((v, t) => v + L.dep[t] + L.capex[t] + L.nwc[t] + L.preop[t] + L.resid[t]);
    P.source = 'archivo';
    if (P.fclFile) {
      const bad = P.fcl.filter((v, t) => Math.abs(v - P.fclFile[t]) > 1).length;
      P.fclError = bad > 0 ? { months: bad, total: H + 1, fileVan: npvAt(P.fclFile, monthlyRate(wacc)) } : null;
    }
  } else if (has(L.fcl)) { P.fcl = L.fcl; P.source = 'archivo'; }
  else { P.fcl = eng ? eng.fcl : zeros(H + 1); }
  P.rev = has(E.rev) ? E.rev : (eng ? eng.rev : zeros(H + 1));
  P.ebitda = has(E.ebitda) ? E.ebitda : (eng ? eng.ebitda : zeros(H + 1));
  P.ebit = has(E.ebit) ? E.ebit : (eng ? eng.ebit : zeros(H + 1));
  P.ni = has(E.ni) ? E.ni : (eng ? eng.ni : zeros(H + 1));
  // flujo financiero (al accionista): FCL − impuestos de más/menos por escudo fiscal − intereses − amortización + desembolso del préstamo
  const loan = isNum(s.loan) ? s.loan : 0;
  if (has(L.fcl) && has(F.amort) && has(E.taxEbt) && has(L.taxEbit)) {
    P.fcfe = P.fcl.map((v, t) => t === 0 ? v + loan : v + ((L.taxEbit[t] || 0) - (E.taxEbt[t] || 0)) - (F.interest[t] || 0) - (F.amort[t] || 0));
    P.fcfeSource = 'archivo';
  } else if (eng) { P.fcfe = eng.fcfe; P.fcfeSource = 'motor'; }
  else { P.fcfe = null; P.fcfeSource = null; }
  const rm = monthlyRate(wacc);
  P.van = npvAt(P.fcl, rm);
  P.tirM = irrBisect(P.fcl); P.tirA = P.tirM == null ? null : Math.pow(1 + P.tirM, 12) - 1;
  let a = 0; P.acc = P.fcl.map(v => (a += v));
  P.payback = null; for (let t = 1; t <= H; t++) if (P.acc[t] >= 0) { P.payback = t; break; }
  P.inversion = (sup.capex || 0) + (sup.nwc0 || 0) + (sup.preop || 0);
  const f1 = (arr) => arr.slice(1, H + 1);
  P.avgEbitda = avg(f1(P.ebitda)); P.avgEbit = avg(f1(P.ebit)); P.avgRev = avg(f1(P.rev));
  if (canEngine) {
    const cm = 1 - sup.cv;
    P.beOp = cm > 0 ? { rev: sup.fixed / cm, revAcc: (sup.fixed + (sup.capex || 0) / (Math.max(1, sup.life || 10) * 12)) / cm } : null;
    if (P.beOp) P.beOp.safety = P.rev[1] ? (P.rev[1] - P.beOp.rev) / P.rev[1] : null;
  }
  // conciliación con el Excel
  P.reconcile = {};
  if (ev.resumen && isNum(ev.resumen.van)) {
    P.reconcile.vanFile = ev.resumen.van; P.reconcile.vanApp = P.van;
    P.reconcile.diffPct = Math.abs(ev.resumen.van) > 1 ? Math.abs(P.van - ev.resumen.van) / Math.abs(ev.resumen.van) : Math.abs(P.van - ev.resumen.van);
    if (P.fclFile) P.reconcile.vanFromFileFcl = npvAt(P.fclFile, rm);
  }
  if (isNum(ev.resumen && ev.resumen.tirAnual)) { P.reconcile.tirFile = ev.resumen.tirAnual; P.reconcile.tirApp = P.tirA; }
  if (eng && compsOK) P.reconcile.engineFclMaxDiff = maxAbs(eng.fcl.map((v, i) => v - P.fcl[i]));
  // escenarios y tornado (requieren motor completo)
  if (canEngine) {
    const E3 = ev.esc;
    if (E3) {
      P.scenarios = [['pes', 'Pesimista'], ['base', 'Base'], ['opt', 'Optimista']].map(([k, label]) => {
        const o = E3[k] || {}; const inp = Object.assign({}, sup); let note = null;
        if (isNum(o.g)) inp.g = o.g; if (isNum(o.cv)) inp.cv = o.cv; if (isNum(o.wacc)) inp.wacc = o.wacc;
        if (isNum(o.capex)) {
          // Salvaguarda: una inversión de escenario <50% o >200% de la base casi siempre es un error de unidad/cero
          if (sup.capex > 0 && (o.capex < sup.capex * 0.5 || o.capex > sup.capex * 2)) note = `Inversión declarada ${fmt(o.capex)} inconsistente con la base ${fmt(sup.capex)}: se usó la base`;
          else inp.capex = o.capex;
        }
        const r = AM.projectEngine(inp);
        return { key: k, label, g: inp.g, cv: inp.cv, wacc: inp.wacc, capex: inp.capex, van: r.van, tirA: r.tirA, payback: r.payback, note };
      });
    } else {
      P.scenarios = [['Pesimista', { g: sup.g * 0.5, cv: Math.min(0.95, sup.cv * 1.1), wacc: wacc + 0.02 }], ['Base', {}], ['Optimista', { g: sup.g * 1.5, cv: sup.cv * 0.9, wacc: Math.max(0.01, wacc - 0.02) }]].map(([label, o]) => {
        const inp = Object.assign({}, sup, o), r = AM.projectEngine(inp);
        return { key: label, label, g: inp.g, cv: inp.cv, wacc: inp.wacc, capex: inp.capex, van: r.van, tirA: r.tirA, payback: r.payback, generated: true };
      });
    }
    P.tornado = AM.tornado(sup, 0.10);
  }
  return P;
}
AM.whatIf = function (P, over) {
  if (!P || !P.canWhatIf) return null;
  const r = AM.projectEngine(Object.assign({}, P.sup, over));
  return { van: r.van, tirA: r.tirA, payback: r.payback, ni1: sum(r.ni.slice(1, 13)), niAvg: avg(r.ni.slice(1, r.H + 1)) };
};

/* ---------- construcción principal ---------- */
AM.build = function (store, Ain) {
  const A = Object.assign(AM.defaultAssumptions(), store.drivers || {}, Ain || {});
  const M = { A, inventory: store.inventory, files: store.files, quality: [], insights: {}, sources: {}, driversFile: store.drivers || null, balanceFile: store.balance || null };
  const plSel = AM.hasPL(store) ? pickPL(store, A) : null;
  const proj = buildProject(store.ev, A);
  M.pl = store.pl; M.plSel = plSel; M.proj = proj;
  M.hasPL = !!plSel; M.hasProj = !!proj;
  const synth = (!plSel && !proj) ? (synthFromCF(store, A) || synthFromBalance(store, A)) : null;
  M.synth = synth ? { kind: synth.kind, label: synth.label, notes: synth.notes, hist: synth.hist } : null;
  M.empty = !plSel && !proj && !synth;
  if (M.empty) return M;

  M.currency = (plSel && plSel.pl.currency) || (proj && proj.moneda) || '';
  A.tax = A.tax != null ? A.tax : ((plSel && plSel.pl.taxRate) || (proj && proj.tax) || 0.27);
  if (synth && synth.Ab) synth.Ab.tax = A.tax;
  const tipo = plSel ? plSel.tipo : 'servicios';
  const Aeff = Object.assign({}, A);
  Aeff.dio = A.dio != null ? A.dio : (tipo === 'productiva' ? 60 : 0);
  Aeff.dpo = A.dpo != null ? A.dpo : (tipo === 'productiva' ? 45 : 15);
  M.Aeff = Aeff;

  // ---- negocio: del P&L (preferente) o del proyecto
  const bal = store.balance || null;
  let ser = null, open = null, nwcFixed = null, bizSource = null;
  if (plSel) {
    ser = serFromPL(plSel.pl); bizSource = 'P&L'; open = bal ? Object.assign({}, bal) : null;
  } else if (synth) {
    ser = synth.ser; bizSource = synth.label + ' (estimado)'; open = synth.open;
  } else if (proj) {
    const H = Math.min(proj.H, 60), sl = a => (a || zeros(H + 1)).slice(1, H + 1);
    const eng = proj.engine;
    const useFile = !!store.ev.er;
    const g = (arrE, arrEng) => useFile && arrE && arrE.some(v => v) ? sl(arrE) : sl(arrEng);
    const E = store.ev.er ? {
      rev: byMonth(store.ev.er, ['ingresos por ventas', 'ingresos'], H), cv: byMonth(store.ev.er, ['costos variables'], H), fx: byMonth(store.ev.er, ['costos fijos'], H),
      dep: byMonth(store.ev.er, ['depreciacion'], H), ebit: byMonth(store.ev.er, ['ebit'], H), fin: byMonth(store.ev.er, ['gastos financieros'], H),
      tax: byMonth(store.ev.er, ['impuestos'], H), ni: byMonth(store.ev.er, ['utilidad neta'], H) } : {};
    const rev = g(E.rev, eng && eng.rev), cogs = g(E.cv, eng && eng.cvCost), fx = g(E.fx, eng && eng.fixed), dep = g(E.dep, eng && eng.dep);
    const ebit = g(E.ebit, eng && eng.ebit), fin = g(E.fin, eng && eng.interest), tx = g(E.tax, eng && eng.taxEbt), ni = g(E.ni, eng && eng.ni);
    ser = { labels: Array.from({ length: H }, (_, i) => 'M' + (i + 1)), rev, cogs, opex: fx.map((v, i) => v + dep[i]), dep, ebit, finExp: fin.map(Math.abs), finInc: zeros(H), tax: tx, ni, invFinal: null, capex: zeros(H) };
    const s = proj.sup;
    nwcFixed = 0; // el modelo del archivo mantiene el capital de trabajo inicial constante: se trata como caja inicial
    open = { cash: s.nwc0 || 0, ar: 0, inv: 0, ap: 0, ppe: s.capex || 0, debt: s.loan || 0, capital: (s.capex || 0) + (s.nwc0 || 0) + (s.preop || 0) - (s.loan || 0), re: -(s.preop || 0) };
    bizSource = 'Evaluación de proyectos';
  }
  if (ser) {
    const Ab = Object.assign({}, Aeff);
    if (synth) Object.assign(Ab, synth.Ab, { tax: Aeff.tax });
    else if (!plSel) { Ab.capexPct = 0; Ab.debt0 = null; }
    const b = deriveBusiness(ser, Ab, open, nwcFixed);
    b.source = bizSource; b.ser = ser; b.nwcFixed = nwcFixed != null; b.labels = ser.labels;
    b.openFromFile = !!(open && plSel) || !!(synth && synth.openFromFile);
    M.biz = b;
    if (b.bs.assets.length) {
      const ebitdaArr = ser.ebit.map((e, i) => e + ser.dep[i]);
      const cmRatio = sum(ser.rev) > 0 ? 1 - sum(ser.cogs) / sum(ser.rev) : null;
      const fixedCash = sum(ser.opex) - sum(ser.dep);
      b.be = (cmRatio > 0 && fixedCash > 0) ? { op: (fixedCash / ser.rev.length) / cmRatio, acc: (sum(ser.opex) / ser.rev.length) / cmRatio, cm: cmRatio, revAvg: avg(ser.rev) } : null;
      if (b.be) b.be.safety = b.be.revAvg ? (b.be.revAvg - b.be.op) / b.be.revAvg : null;
      b.ebitdaArr = ebitdaArr;
    }
  } else M.biz = null;

  // ---- fuentes (trazabilidad)
  const S = M.sources;
  S.pl = M.hasPL ? { state: 'cargado', note: `${plSel.tipo} · ${plSel.esc}` } : (synth ? { state: 'derivado', note: 'estimado desde ' + synth.label.toLowerCase() } : { state: 'ausente', note: '' });
  S.proj = M.hasProj ? { state: 'cargado', note: proj.source === 'archivo' ? 'flujo del archivo' : 'recalculado' } : { state: 'ausente', note: '' };
  S.cf = M.hasProj ? { state: 'cargado', note: 'proyecto (económico y financiero)' } : (synth && synth.kind === 'cf' ? { state: 'cargado', note: 'estado de flujo del archivo' } : (M.biz ? { state: 'derivado', note: synth ? (synth.hist ? 'variaciones del balance + proyección' : 'proyección desde balance') : 'desde P&L' } : { state: 'ausente', note: '' }));
  S.bs = synth && synth.kind === 'balance' ? { state: 'cargado', note: 'balance del archivo' + (synth.hist ? ' (' + (store.balPeriods) + ' períodos)' : '') } : store.balance && M.hasPL ? { state: 'cargado', note: 'saldo inicial del archivo + proyección' } : (M.biz ? { state: 'derivado', note: M.hasPL ? 'desde P&L con supuestos' : (synth ? 'reconstruido desde ' + synth.label.toLowerCase() : 'desde proyecto') } : { state: 'ausente', note: '' });
  S.ma = store.ma ? { state: 'cargado', note: 'parámetros del archivo' } : { state: 'derivado', note: 'autocompletado desde el modelo' };

  // ---- M&A autocompletado
  Object.assign(M, AM.deriveMA(M, store));
  M.maCalc = AM.maCompute(M.maParams);

  // ---- vistas
  M.plView = AM.buildPL(M);
  M.overview = buildOverview(M);
  M.insights.cf = cfInsights(M);
  M.insights.bs = bsInsights(M);
  M.quality = qualityChecks(M, store);
  return M;
};

/* ---------- M&A derivado ---------- */
AM.deriveMA = function (M, store) {
  const p = AM.maDefaults(), src = {};
  Object.keys(p).forEach(k => src[k] = 'Supuesto');
  const b = M.biz, A = M.Aeff;
  const ev0 = store.ev || {};
  if (ev0.supuestos && isNum(ev0.supuestos.g)) { p.growth = ev0.supuestos.g; src.growth = 'Evaluación'; }
  if (b) {
    const n = Math.min(12, b.ser.rev.length);
    const act = M.hasPL || !!M.synth;
    const fromP = M.hasPL ? 'P&L' : (M.synth ? M.synth.label + ' (estimado)' : 'Evaluación');
    // P&L: el año real es el año 0. Proyecto: el año 1 del modelo es el primer año proyectado.
    const rev = sum(b.ser.rev.slice(0, n)) * 12 / n / (act ? 1 : (1 + p.growth));
    if (rev > 0) {
      p.ingBase = rev; src.ingBase = fromP;
      const hz = act ? n : b.ser.rev.length;
      const eb = sum(b.ebitda.slice(0, hz)) / sum(b.ser.rev.slice(0, hz)); p.margin = clamp(eb, 0.01, 0.9); src.margin = fromP + (act ? '' : ' (prom. horizonte)');
      p.da = clamp(sum(b.ser.dep.slice(0, n)) / sum(b.ser.rev.slice(0, n)), 0, 0.3); src.da = fromP;
      p.capex = clamp(act ? b.capexPct : (M.proj ? Math.max(0.02, p.da) : 0.04), 0, 0.3); src.capex = act ? 'Supuesto del modelo' : 'Supuesto';
      const nw = b.bs.nwc[b.bs.nwc.length - 1]; if (act && isFinite(nw) && nw > 0) { p.nwc = clamp(nw / (rev), 0, 0.5); src.nwc = fromP + ' (CT/ingresos)'; }
      p.tax = clamp(A.tax, 0, 0.5); src.tax = M.hasPL ? 'P&L' : (M.synth ? 'Supuesto' : 'Evaluación');
      p.netDebt = b.debt0 - b.cash0; src.netDebt = 'Saldo inicial del modelo';
      p.kd = A.kd; src.kd = 'Supuesto del modelo';
    }
  }
  const ev = store.ev || {};
  if (ev.val) {
    if (isNum(ev.val.mult)) { p.mult = ev.val.mult; src.mult = 'Evaluación'; }
    if (isNum(ev.val.premium)) { p.premium = ev.val.premium; src.premium = 'Evaluación'; }
    if (isNum(ev.val.synergies)) { p.synCost = ev.val.synergies; src.synCost = 'Evaluación'; }
  }
  p.termG = Math.min(0.025, Math.max(0, p.growth));
  if (store.ma) Object.keys(store.ma).forEach(k => { if (k in p && isNum(store.ma[k])) { p[k] = store.ma[k]; src[k] = 'Archivo M&A'; } });
  p.growth = clamp(p.growth, -0.2, 0.5);
  return { maParams: p, maSrc: src, maWaccOverride: store.ma && isNum(store.ma.waccOverride) ? store.ma.waccOverride : null };
};


/* ============================================================
   6b. VISTA P&L — Real vs Budget vs Forecast, análisis vertical y horizontal
   ============================================================ */
const ESC_LABEL = { real: 'Real', budget: 'Budget', forecast: 'Forecast' };
function plAgg(pl) {
  const t = k => (pl[k] || zeros(13)).slice(0, 12);
  const rev = t('ingresosNetos'), cogs = t('costoTotal'), opex = t('gastosOp'), dep = t('depreciacion');
  const gross = pl.utilidadBruta ? t('utilidadBruta') : rev.map((r, i) => r - cogs[i]);
  const ebit = pl.ebit ? t('ebit') : gross.map((g, i) => g - opex[i]);
  const ebitda = ebit.map((v, i) => v + dep[i]);
  const other = pl.otrosTotal ? t('otrosTotal') : zeros(12);
  const ebt = pl.ebt ? t('ebt') : ebit.map((v, i) => v + other[i]);
  const tax = t('impuesto').map(Math.abs);
  const ni = pl.utilidadNeta ? t('utilidadNeta') : ebt.map((v, i) => v - tax[i]);
  const S = a => sum(a);
  return { rev, cogs, gross, opex, dep, ebit, ebitda, other, ebt, tax, ni,
    tot: { rev: S(rev), cogs: S(cogs), gross: S(gross), opex: S(opex), dep: S(dep), ebit: S(ebit), ebitda: S(ebitda), other: S(other), ebt: S(ebt), tax: S(tax), ni: S(ni) } };
}
// asigna cada línea a su grupo (ingresos / costo / gasto op. / otros) según el orden del P&L
function plLeafLines(pl) {
  const out = []; let g = null;
  (pl.lines || []).forEach(L => {
    const k = L.key;
    if (L.section) {
      if (/^ingresos (por|operacionales)/.test(k)) g = 'ing';
      else if (/^costo de (servicios|ventas)$/.test(k)) g = 'cost';
      else if (/^gastos operacionales$/.test(k)) g = 'opex';
      else if (/^otros ingresos y gastos/.test(k)) g = 'other';
      return;
    }
    if (/^total ingresos brutos/.test(k)) { g = 'ing-post'; return; }
    if (/^costo de (servicios|ventas) total/.test(k)) { g = null; return; }
    if (/^total gastos operacionales/.test(k)) { g = null; return; }
    if (/^total otros/.test(k)) { g = null; return; }
    if (L.pct || !g || g === 'ing-post') return;
    out.push({ label: L.label.trim(), key: k.trim(), group: g, vals: L.vals });
  });
  return out;
}
const semV = (v, tol) => v >= 0 ? 'green' : v >= -tol ? 'yellow' : 'red';      // más es mejor
const semC = (v, tol) => v <= 0 ? 'green' : v <= tol ? 'yellow' : 'red';        // más es peor
const varPct = (a, b) => (b && Math.abs(b) > 1e-9) ? (a - b) / Math.abs(b) : null;

AM.buildPL = function (M) {
  if (!M.plSel) return null;
  const tipo = M.plSel.tipo, avail = M.plSel.avail;
  const V = { tipo, avail, esc: M.plSel.esc, scen: {}, insights: [], kpis: [] };
  avail.forEach(k => { const pl = M.pl[tipo][k]; V.scen[k] = { key: k, label: ESC_LABEL[k], pl, agg: plAgg(pl), lines: plLeafLines(pl) }; });
  const R = V.scen.real, B = V.scen.budget, F = V.scen.forecast;
  const base = R || B || F;               // escenario de referencia para estructura
  V.base = base.key;
  const pctRev = (agg, x) => agg.tot.rev ? x / agg.tot.rev : null;

  // ---- tabla comparativa
  const rows = [
    ['Ingresos netos', 'rev', 'sum', 'high'], ['Costo de ventas / servicios', 'cogs', 'cost', 'low'], ['Utilidad bruta', 'gross', 'sub', 'high'],
    ['Gastos operacionales', 'opex', 'cost', 'low'], ['EBITDA', 'ebitda', 'sub', 'high'], ['Depreciación y amortización', 'dep', 'cost', 'low'],
    ['EBIT (resultado operacional)', 'ebit', 'sub', 'high'], ['Resultado no operacional', 'other', 'sum', 'high'], ['Utilidad antes de impuestos', 'ebt', 'sub', 'high'],
    ['Impuesto a la renta', 'tax', 'cost', 'low'], ['Utilidad neta', 'ni', 'total', 'high']
  ];
  V.table = rows.map(([label, k, kind, better]) => {
    const row = { label, k, kind, better, vals: {}, pct: {} };
    avail.forEach(s => { const a = V.scen[s].agg; row.vals[s] = a.tot[k]; row.pct[s] = pctRev(a, a.tot[k]); });
    row.varRB = (R && B) ? { abs: R.agg.tot[k] - B.agg.tot[k], pct: varPct(R.agg.tot[k], B.agg.tot[k]) } : null;
    row.varFB = (F && B) ? { abs: F.agg.tot[k] - B.agg.tot[k], pct: varPct(F.agg.tot[k], B.agg.tot[k]) } : null;
    row.varRF = (R && F) ? { abs: R.agg.tot[k] - F.agg.tot[k], pct: varPct(R.agg.tot[k], F.agg.tot[k]) } : null;
    return row;
  });
  const T = k => V.table.find(r => r.k === k);

  // ---- KPIs con semáforo vs Budget
  const mk = (label, k, kind) => {
    const row = T(k), a = base.agg.tot, v = a[k], mg = a.rev ? v / a.rev : null;
    let sub = [], s = 'yellow';
    const cmp = (row.varRB && R) ? row.varRB : (row.varFB ? row.varFB : null);
    if (R && B && row.varRB) { sub.push(`vs Budget ${row.varRB.pct == null ? 'n/d' : (row.varRB.pct >= 0 ? '+' : '') + (row.varRB.pct * 100).toFixed(1) + '%'}`); s = row.better === 'high' ? semV(row.varRB.pct == null ? 0 : row.varRB.pct, 0.05) : semC(row.varRB.pct == null ? 0 : row.varRB.pct, 0.05); }
    if (F && B && !R && row.varFB) { sub.push(`Forecast vs Budget ${row.varFB.pct == null ? 'n/d' : (row.varFB.pct >= 0 ? '+' : '') + (row.varFB.pct * 100).toFixed(1) + '%'}`); s = semV(row.varFB.pct == null ? 0 : row.varFB.pct, 0.05); }
    if (R && F && row.varRF) sub.push(`vs Forecast ${row.varRF.pct == null ? 'n/d' : (row.varRF.pct >= 0 ? '+' : '') + (row.varRF.pct * 100).toFixed(1) + '%'}`);
    if (k !== 'rev' && mg != null) sub.unshift(`Margen ${(mg * 100).toFixed(1)}%`);
    if (!(R && B) && !(F && B)) s = v > 0 ? 'green' : v < 0 ? 'red' : 'yellow';
    return kpi(label, fmt(v), sub.join(' · ') || ESC_LABEL[base.key], s);
  };
  V.kpis = [mk('Ingresos netos', 'rev'), mk('Utilidad bruta', 'gross'), mk('EBITDA', 'ebitda'), mk('EBIT', 'ebit'), mk('Utilidad neta', 'ni')];

  // ---- análisis vertical: estructura como % de ingresos por escenario
  V.vertical = { labels: avail.map(s => ESC_LABEL[s]), cogs: [], opex: [], tax: [], ni: [], other: [] };
  avail.forEach(s => {
    const a = V.scen[s].agg.tot, r = a.rev || 1;
    const nonop = -a.other; // gasto neto no operacional (positivo = gasto)
    V.vertical.cogs.push(a.cogs / r * 100); V.vertical.opex.push(a.opex / r * 100);
    V.vertical.other.push(Math.max(0, nonop) / r * 100); V.vertical.tax.push(a.tax / r * 100); V.vertical.ni.push(a.ni / r * 100);
  });
  // líneas con mayor peso (escenario base)
  const topLines = base.lines.map(L => ({ label: L.label, group: L.group, total: sum(L.vals.slice(0, 12)) })).filter(L => Math.abs(L.total) > 0);
  const revB = base.agg.tot.rev || 1;
  V.revLines = topLines.filter(L => L.group === 'ing').map(L => Object.assign(L, { pct: L.total / revB })).sort((a, b) => b.total - a.total);
  V.costLines = topLines.filter(L => L.group === 'cost' || L.group === 'opex').map(L => Object.assign(L, { pct: L.total / revB })).sort((a, b) => b.total - a.total);

  // ---- análisis horizontal: series mensuales por escenario y variación mes a mes
  const ser = (fn) => { const o = {}; avail.forEach(s => o[s] = fn(V.scen[s].agg)); return o; };
  V.horizontal = { labels: MESES.slice(), rev: ser(a => a.rev), ebit: ser(a => a.ebit), ni: ser(a => a.ni), margin: ser(a => a.rev.map((r, i) => r ? a.ebit[i] / r * 100 : 0)) };
  V.mom = {}; avail.forEach(s => { const r = V.scen[s].agg.rev; V.mom[s] = r.map((v, i) => i === 0 ? null : (r[i - 1] ? (v - r[i - 1]) / Math.abs(r[i - 1]) : null)); });
  // variación por línea Real vs Budget (o Forecast vs Budget)
  const [x, y] = (R && B) ? [R, B] : (F && B) ? [F, B] : [null, null];
  V.lineVar = [];
  if (x) {
    const mapY = {}; y.lines.forEach(L => { mapY[L.key] = L; });
    const seen = new Set();
    x.lines.forEach(L => { seen.add(L.key); const m = mapY[L.key]; const a = sum(L.vals.slice(0, 12)), b = m ? sum(m.vals.slice(0, 12)) : 0; if (a || b) V.lineVar.push({ label: L.label, group: L.group, a, b, d: a - b, pct: varPct(a, b) }); });
    y.lines.forEach(L => { if (!seen.has(L.key)) { const b = sum(L.vals.slice(0, 12)); if (b) V.lineVar.push({ label: L.label, group: L.group, a: 0, b, d: -b, pct: -1 }); } });
    V.lineVar.sort((p, q) => Math.abs(q.d) - Math.abs(p.d));
    V.lineVarNames = [x.label, y.label];
  }

  // ---- insights, alertas y recomendaciones
  const I = V.insights, pcs = (v, d) => (v == null || !isFinite(v)) ? 'n/d' : (v >= 0 ? '+' : '') + (v * 100).toFixed(d == null ? 1 : d) + '%';
  if (R && B) {
    const r = T('rev').varRB, ni = T('ni').varRB, eb = T('ebit').varRB, cg = T('cogs').varRB;
    I.push({ level: r.pct == null ? 'info' : r.pct >= 0 ? 'success' : r.pct > -0.05 ? 'warning' : 'danger', title: `Ingresos Real ${pcs(r.pct)} vs Budget (${fmt(r.abs)})`, text: r.pct >= 0 ? 'Superas el presupuesto de ventas. Recomendación: validar si el sobrecumplimiento es recurrente antes de elevar el Budget o comprometer capacidad adicional.' : 'Ventas bajo presupuesto. Recomendación: revisar pipeline y precio/mix; activar plan comercial correctivo y actualizar el forecast.' });
    const mgR = R.agg.tot.rev ? R.agg.tot.gross / R.agg.tot.rev : 0, mgB = B.agg.tot.rev ? B.agg.tot.gross / B.agg.tot.rev : 0, dpp = (mgR - mgB) * 100;
    I.push({ level: dpp >= 0 ? 'success' : dpp > -2 ? 'warning' : 'danger', title: `Margen bruto ${(mgR * 100).toFixed(1)}% vs ${(mgB * 100).toFixed(1)}% del Budget (${dpp >= 0 ? '+' : ''}${dpp.toFixed(1)} pp)`, text: dpp >= 0 ? 'El margen se expande: el crecimiento viene con mejor mix o eficiencia.' : 'El margen se deteriora respecto al plan: el costo directo crece más que los ingresos. Recomendación: revisar productividad del personal directo, subcontrataciones y precios.' });
    if (ni.pct != null) I.push({ level: ni.pct >= 0 ? 'success' : ni.pct > -0.1 ? 'warning' : 'danger', title: `Utilidad neta ${pcs(ni.pct)} vs Budget`, text: `EBIT ${pcs(eb.pct)} vs Budget; costos ${pcs(cg.pct)}. ${ni.pct >= 0 ? '' : 'Recomendación: identificar las líneas de mayor desvío (tabla de variaciones) y asignar responsables.'}` });
    const un = V.lineVar.filter(L => (L.group === 'cost' || L.group === 'opex') && L.d > 0).slice(0, 2);
    if (un.length) I.push({ level: 'warning', title: `Mayores sobrecostos vs Budget: ${un.map(L => L.label).join(' y ')}`, text: un.map(L => `${L.label}: ${fmt(L.d)} sobre presupuesto (${L.b ? pcs(L.pct, 0) : 'sin presupuesto'})`).join(' · ') + '. Recomendación: confirmar si son recurrentes y si el precio los traspasa al cliente.' });
  }
  if (F && B) {
    const r = T('rev').varFB, e2 = T('ebit').varFB;
    I.push({ level: r.pct == null ? 'info' : r.pct >= 0 ? 'success' : 'warning', title: `Forecast ${pcs(r.pct)} en ingresos y ${pcs(e2.pct)} en EBIT vs Budget`, text: 'El forecast es la mejor estimación de cierre. ' + (r.pct >= 0 ? 'Proyecta superar el presupuesto.' : 'Anticipa brecha frente al presupuesto: define acciones antes del cierre.') });
  }
  if (R && F) { const r = T('rev').varRF; I.push({ level: Math.abs(r.pct || 0) < 0.1 ? 'success' : 'warning', title: `Real vs Forecast: ${pcs(r.pct)} en ingresos`, text: Math.abs(r.pct || 0) < 0.1 ? 'El forecast es consistente con el desempeño real.' : 'Diferencia sobre ±10%: el forecast no anticipa bien el desempeño; recalíbralo con el último cierre mensual.' }); }
  // estructura
  if (V.revLines.length) { const t = V.revLines[0]; if (t.pct > 0.7 && V.revLines.length > 0) I.push({ level: 'warning', title: `Concentración de ingresos: "${t.label}" = ${(t.pct * 100).toFixed(0)}%`, text: 'Dependencia alta de una sola línea de negocio. Recomendación: diversificar o asegurar contratos de largo plazo en esa línea.' }); }
  if (V.costLines.length) { const t = V.costLines[0]; I.push({ level: t.pct > 0.5 ? 'warning' : 'info', title: `Mayor partida de costo: "${t.label}" = ${(t.pct * 100).toFixed(1)}% de los ingresos`, text: t.pct > 0.5 ? 'Una sola partida absorbe más de la mitad de los ingresos: pequeñas variaciones de productividad mueven el resultado. Recomendación: medir utilización y margen por proyecto.' : 'Principal driver de costo del escenario base (análisis vertical).' }); }
  const mgOp = base.agg.tot.rev ? base.agg.tot.ebit / base.agg.tot.rev : 0;
  I.push({ level: mgOp > 0.15 ? 'success' : mgOp > 0.05 ? 'warning' : 'danger', title: `Margen operacional ${(mgOp * 100).toFixed(1)}% (${ESC_LABEL[base.key]})`, text: mgOp > 0.15 ? 'Estructura rentable.' : mgOp > 0.05 ? 'Margen ajustado: sin colchón ante shocks de costo.' : 'Margen insuficiente para sostener inversión y deuda.' });
  if (base.agg.tot.opex === 0) I.push({ level: 'warning', title: 'El P&L no informa gastos operacionales', text: 'Ventas, marketing y administración están en cero: el EBIT podría estar sobrestimado. Completa estas filas para un análisis de eficiencia confiable.' });
  // impuesto: etiqueta vs cálculo
  avail.forEach(s => { const pb = V.scen[s].pl; if (pb.taxLabelRate != null && pb.taxEffective != null && Math.abs(pb.taxLabelRate - pb.taxEffective) > 0.005) I.push({ level: 'warning', title: `Impuesto (${ESC_LABEL[s]}): etiqueta ${(pb.taxLabelRate * 100).toFixed(0)}% pero el cálculo equivale a ${(pb.taxEffective * 100).toFixed(1)}%`, text: 'La fila de impuesto está rotulada con una tasa distinta a la que realmente se aplica. Alinea etiqueta y fórmula (régimen general 27% / Pro Pyme 25%) para no distorsionar la utilidad neta.' }); });
  // horizontal: estacionalidad / caídas
  const m = V.mom[base.key] || [], drops = m.map((v, i) => ({ v, i })).filter(o => o.v != null && o.v < -0.1);
  if (drops.length) I.push({ level: 'warning', title: `Caídas mensuales de ingresos >10% en ${drops.map(o => MESES[o.i]).join(', ')} (${ESC_LABEL[base.key]})`, text: 'Revisa estacionalidad, días hábiles y contratos con entrega puntual; planifica caja para esos meses.' });
  else if (base.agg.rev.every(v => Math.abs(v - base.agg.rev[0]) < 1)) I.push({ level: 'info', title: `Ingresos planos durante 12 meses (${ESC_LABEL[base.key]})`, text: 'Sin crecimiento ni estacionalidad mes a mes. Si es un promedio cargado como mensual, el análisis horizontal no revelará tendencias: sube los datos mensuales reales.' });
  return V;
};

/* ============================================================
   7. VISTAS E INSIGHTS
   ============================================================ */
const sem = (v, g, y) => v >= g ? 'green' : v >= y ? 'yellow' : 'red';
const lvlCls = { green: 'success', yellow: 'warning', red: 'danger' };
function kpi(label, value, sub, s, tone) { return { label, value, sub, sem: s, cls: tone || lvlCls[s] || '' }; }

function buildOverview(M) {
  const O = { kpis: [], charts: {}, insights: [], mode: M.hasProj ? 'proyecto' : 'negocio' };
  const b = M.biz, P = M.proj, mc = M.maCalc, wacc = P ? P.wacc : mc.wacc.wacc;
  const waccTxt = P ? 'cargado del archivo' : 'estimado (CAPM + estructura objetivo)';
  if (P) {
    const van = P.van, tir = P.tirA;
    const vanNote = P.fclError ? ` · el Excel reporta ${fmt(P.reconcile.vanFile != null ? P.reconcile.vanFile : P.fclError.fileVan)} (FCL mal sumado)` : '';
    O.kpis.push(kpi('VAN', fmt(van), (van > 0 ? '✓ Crea valor' : '⚠ Destruye valor al WACC') + vanNote, van > 0 ? 'green' : 'red'));
    O.kpis.push(kpi('TIR anual', tir == null ? 'n/d' : pc(tir, 2), tir == null ? 'Flujos sin cambio de signo' : (tir > wacc ? `✓ Supera WACC (${pc(wacc, 1)})` : `⚠ Bajo WACC (${pc(wacc, 1)})`), tir == null ? 'yellow' : (tir > wacc ? 'green' : 'red')));
    O.kpis.push(kpi('WACC', pc(wacc, 2), 'Tasa de descuento ' + waccTxt, 'yellow', ''));
    O.kpis.push(kpi('EBITDA (prom. mensual)', fmt(P.avgEbitda), `Margen: ${pc(P.avgRev ? P.avgEbitda / P.avgRev : 0, 1)} · horizonte ${P.H} m`, P.avgEbitda > 0 ? 'green' : 'red'));
    O.kpis.push(kpi('EBIT (prom. mensual)', fmt(P.avgEbit), `Margen: ${pc(P.avgRev ? P.avgEbit / P.avgRev : 0, 1)}`, P.avgEbit > 0 ? 'green' : 'red'));
    const be = P.beOp;
    O.kpis.push(kpi('Punto de equilibrio operativo', be ? fmt(be.rev) + ' / mes' : 'n/d', be ? `Ingresos mensuales necesarios (EBITDA = 0) · margen de seguridad ${pc(be.safety, 0)}` : 'Requiere hoja Supuestos', be ? (be.safety > 0.15 ? 'green' : be.safety > 0 ? 'yellow' : 'red') : 'yellow'));
    O.kpis.push(kpi('Recuperación de la inversión', P.payback ? `Mes ${P.payback}` : `No en ${P.H} meses`, P.payback ? `Inversión total ${fmt(P.inversion)}` : `Acumulado al mes ${P.H}: ${fmt(lastOf(P.acc))}`, P.payback ? (P.payback <= 36 ? 'green' : 'yellow') : 'red'));
    O.charts.cf = { labels: Array.from({ length: P.H }, (_, i) => 'M' + (i + 1)), bars: P.fcl.slice(1), acc: P.acc.slice(1), title: `Flujo de caja libre y acumulado — ${P.H} meses` };
    O.charts.eb = { labels: O.charts.cf.labels.slice(0, 24), ebitda: P.ebitda.slice(1, 25), ebit: P.ebit.slice(1, 25), title: 'EBITDA y EBIT mensuales — primeros 24 meses' };
  } else if (b) {
    const s = b.ser, n = s.rev.length, rev = sum(s.rev), eb = sum(b.ebitda), ebit = sum(s.ebit), ni = sum(s.ni);
    const est = M.synth ? ' · estimado' : '';
    O.kpis.push(kpi('Ingresos netos', fmt(rev), M.plSel ? `${M.plSel.tipo} · ${M.plSel.esc}` : 'estimado desde ' + M.synth.label.toLowerCase(), M.synth ? 'yellow' : 'green', ''));
    O.kpis.push(kpi('EBITDA', fmt(eb), `Margen: ${pc(rev ? eb / rev : 0, 1)}${(!M.plSel || M.plSel.pl.depFound) ? est : ' · P&L sin D&A'}`, eb > 0 ? (eb / rev > 0.15 ? 'green' : 'yellow') : 'red'));
    O.kpis.push(kpi('EBIT', fmt(ebit), `Margen: ${pc(rev ? ebit / rev : 0, 1)}`, ebit > 0 ? 'green' : 'red'));
    O.kpis.push(kpi('Utilidad neta', fmt(ni), `Margen: ${pc(rev ? ni / rev : 0, 1)}`, ni > 0 ? 'green' : 'red'));
    O.kpis.push(kpi('WACC', pc(wacc, 2), 'Estimado (CAPM + estructura objetivo)', 'yellow', ''));
    O.kpis.push(kpi('Valor empresa (DCF)', fmt(mc.dcf.ev), `Equity: ${fmt(mc.equity)} · autogenerado en M&A`, mc.dcf.ev > 0 ? 'green' : 'red'));
    const be = b.be;
    O.kpis.push(kpi('Punto de equilibrio operativo', be ? fmt(be.op) + ' / mes' : 'n/d', be ? `Ingresos mensuales necesarios · margen de seguridad ${pc(be.safety, 0)}` : 'Sin margen de contribución', be ? (be.safety > 0.2 ? 'green' : be.safety > 0 ? 'yellow' : 'red') : 'yellow'));
    const acc = []; let a = 0; b.fcff.forEach(v => acc.push(a += v));
    O.charts.cf = { labels: s.labels, bars: b.fcff, acc, title: 'Flujo de caja libre (económico) y acumulado — año' };
    O.charts.eb = { labels: s.labels, ebitda: b.ebitda, ebit: s.ebit, title: 'EBITDA y EBIT mensuales' };
  }
  // insights
  const I = O.insights;
  if (P) {
    if (P.fclError) I.push({ level: 'danger', title: 'Tu Excel sobrestima el flujo de caja libre', text: `La hoja "Flujo de Caja" suma EBIT e impuestos además del NOPAT. VAN del archivo ${fmt(P.reconcile.vanFile != null ? P.reconcile.vanFile : P.fclError.fileVan)} vs VAN corregido ${fmt(P.van)}. Todas las cifras de esta app usan el flujo corregido. Detalle y fórmula sugerida en la pestaña Modelo.` });
    I.push(P.van > 0 ? { level: 'success', title: `Proyecto viable: VAN ${fmt(P.van)}`, text: 'Crea valor al WACC del archivo. Recomendación: validar riesgos operativos y la robustez del VAN con el tornado de sensibilidad.' }
      : { level: 'danger', title: `VAN negativo (${fmt(P.van)}) en el horizonte de ${P.H} meses`, text: 'En el escenario base el proyecto no remunera el costo de capital. Recomendación: renegociar inversión, mejorar margen de contribución o evaluar un horizonte mayor con valor terminal.' });
    if (P.tirA != null) I.push(P.tirA > P.wacc ? { level: 'success', title: `TIR ${pc(P.tirA, 1)} supera al WACC (${pc(P.wacc, 1)})`, text: `Spread de ${pc(P.tirA - P.wacc, 1)}.` } : { level: 'warning', title: `TIR ${pc(P.tirA, 1)} bajo el WACC (${pc(P.wacc, 1)})`, text: 'La rentabilidad no compensa el riesgo. Necesitas ≈ ' + pc(P.wacc - P.tirA, 1) + ' más de retorno anual.' });
    I.push(P.payback ? { level: P.payback <= 36 ? 'success' : 'warning', title: `Recuperación de la inversión en el mes ${P.payback}`, text: P.payback <= 36 ? 'Plazo de recuperación razonable para una PYME.' : 'Recuperación tardía: aumenta la exposición a riesgo de mercado. Evalúa reducir capex inicial o escalonar la inversión.' }
      : { level: 'danger', title: `La inversión no se recupera dentro de ${P.H} meses`, text: `Flujo acumulado al cierre: ${fmt(lastOf(P.acc))}. Recomendación: no aprobar sin mejorar ingresos/margen o ampliar horizonte con justificación.` });
    if (P.beOp) I.push(P.beOp.safety > 0.15 ? { level: 'success', title: `Margen de seguridad ${pc(P.beOp.safety, 0)} sobre el punto de equilibrio`, text: `Se requieren ${fmt(P.beOp.rev)} mensuales para cubrir costos fijos; el mes 1 proyecta ${fmt(P.rev[1])}.` }
      : { level: P.beOp.safety > 0 ? 'warning' : 'danger', title: P.beOp.safety > 0 ? 'Margen de seguridad estrecho' : 'Ingresos iniciales bajo el punto de equilibrio', text: `Se requieren ${fmt(P.beOp.rev)} mensuales (EBITDA = 0) vs ${fmt(P.rev[1])} proyectados en el mes 1. Recomendación: acelerar ramp-up comercial o reducir costos fijos (${fmt(P.sup.fixed)}/mes).` });
    if (P.tornado) { const t = P.tornado[0]; I.push({ level: 'info', title: `Variable más sensible: ${t.label}`, text: `±10% mueve el VAN entre ${fmt(t.down)} y ${fmt(t.up)} respecto al base. Concentra la validación comercial en este supuesto.` }); }
  } else if (b) {
    const rev = sum(b.ser.rev), eb = sum(b.ebitda), mg = rev ? eb / rev : 0;
    I.push(mg > 0.2 ? { level: 'success', title: `Margen EBITDA sólido (${pc(mg, 1)})`, text: 'Estructura de costos saludable para sostener inversión y servicio de deuda.' } : mg > 0.08 ? { level: 'warning', title: `Margen EBITDA moderado (${pc(mg, 1)})`, text: 'Espacio para eficiencia: revisa costo de ventas y gastos operacionales.' } : { level: 'danger', title: `Margen EBITDA débil (${pc(mg, 1)})`, text: 'Prioriza pricing y control de costos antes de comprometer capex o deuda.' });
    if (b.be) I.push(b.be.safety > 0.2 ? { level: 'success', title: `Margen de seguridad ${pc(b.be.safety, 0)}`, text: `Equilibrio operativo en ${fmt(b.be.op)} mensuales vs ${fmt(b.be.revAvg)} promedio actual (supuesto: costo de ventas variable, gastos operacionales fijos).` } : { level: b.be.safety > 0 ? 'warning' : 'danger', title: 'Margen de seguridad reducido', text: `Equilibrio en ${fmt(b.be.op)} mensuales vs ${fmt(b.be.revAvg)} actuales. Reduce costos fijos o aumenta el margen de contribución.` });
    if (M.plSel && (!M.plSel.pl.depFound || sum(b.ser.dep) === 0)) I.push({ level: 'warning', title: 'El P&L no registra depreciación ni amortización', text: 'EBITDA = EBIT y el flujo no incorpora reposición de activos. Recomendación: completar D&A y capex para un flujo de caja y valoración realistas.' });
    I.push({ level: 'info', title: `Valor empresa DCF estimado: ${fmt(mc.dcf.ev)}`, text: 'Calculado en la pestaña M&A con parámetros de mercado por defecto. Edítalos para refinar la valoración.' });
    if (M.synth) M.synth.notes.forEach(n => I.push(n));
    else I.push({ level: 'info', title: 'Tienes solo P&L: flujo de caja y balance fueron generados', text: 'Revisa los supuestos (DSO/DIO/DPO, capex, caja inicial) en la pestaña Modelo y súbelos en el archivo unificado para mayor precisión. VAN/TIR requieren la hoja de Evaluación de proyectos.' });
  }
  return O;
}

function cfInsights(M) {
  const I = [], P = M.proj, b = M.biz;
  if (P) {
    I.push(P.fcl[0] < 0 ? { level: 'info', title: `Desembolso inicial de ${fmt(-P.fcl[0])}`, text: `Incluye CAPEX ${fmt(P.sup.capex || 0)}, capital de trabajo ${fmt(P.sup.nwc0 || 0)} y gastos pre-operativos ${fmt(P.sup.preop || 0)}.` } : { level: 'info', title: 'Sin desembolso inicial', text: '' });
    const neg = P.fcl.slice(1).filter(v => v < 0).length;
    if (neg > 0) I.push({ level: neg > P.H * 0.25 ? 'danger' : 'warning', title: `${neg} meses con flujo operativo negativo`, text: 'Requieren financiamiento de corto plazo o caja inicial. Recomendación: definir una línea de capital de trabajo equivalente al máximo déficit acumulado.' });
    const minAcc = Math.min(...P.acc); I.push({ level: P.payback ? 'success' : 'danger', title: `Exposición máxima de caja: ${fmt(minAcc)}`, text: P.payback ? `Se recupera en el mes ${P.payback}.` : 'No se recupera en el horizonte.' });
    if (P.fcfe) {
      const tirF = irrBisect(P.fcfe); const tirFA = tirF == null ? null : Math.pow(1 + tirF, 12) - 1;
      I.push({ level: 'info', title: 'Flujo financiero (al accionista) vs económico', text: `Desembolso propio inicial ${fmt(-P.fcfe[0])} (el resto lo financia la deuda). ${tirFA != null ? 'TIR del accionista ' + pc(tirFA, 1) + ' vs TIR del proyecto ' + (P.tirA != null ? pc(P.tirA, 1) : 'n/d') + (P.tirA != null && tirFA > P.tirA ? ' — el apalancamiento agrega retorno.' : ' — el apalancamiento no mejora el retorno.') : ''}` });
    }
    if (P.sup.nwc0 && P.resid && !(P.sup && 0)) {/* sin acción */}
  }
  if (b) {
    const R = b.ratios, cfoS = sum(b.cfo), cfiS = sum(b.cfi), cffS = sum(b.cff);
    I.push(cfoS >= 0 ? { level: 'success', title: `Caja operacional positiva (${fmt(cfoS)})`, text: R.cfoToEbitda != null ? `Conversión de EBITDA a caja: ${pc(R.cfoToEbitda, 0)}.` + (R.cfoToEbitda < 0.7 ? ' Conversión baja: el capital de trabajo está absorbiendo caja; revisa cobranza e inventarios.' : '') : '' } : { level: 'danger', title: `Caja operacional negativa (${fmt(cfoS)})`, text: 'La operación consume caja. Recomendación: acelerar cobranza, renegociar plazos con proveedores y revisar la estructura de costos.' });
    I.push({ level: 'info', title: `Inversión ${fmt(cfiS)} · Financiamiento ${fmt(cffS)}`, text: b.capexPct > 0 ? `Capex supuesto ${pc(b.capexPct, 1)} de ingresos.` : 'Capex supuesto en 0%: el P&L no informa depreciación. Define capex de reposición para un flujo más realista.' });
    if (R.runway != null) I.push(R.runway >= 3 ? { level: 'success', title: `Caja final cubre ${R.runway.toFixed(1)} meses de costos`, text: 'Colchón de liquidez adecuado (≥ 3 meses).' } : R.runway >= 1 ? { level: 'warning', title: `Caja final cubre ${R.runway.toFixed(1)} meses de costos`, text: 'Colchón ajustado. Recomendación: mantener ≥ 3 meses o contar con línea comprometida.' } : { level: 'danger', title: `Caja final cubre solo ${Math.max(0, R.runway).toFixed(1)} meses de costos`, text: 'Riesgo de liquidez. Recomendación: refinanciar o aportar capital; prioriza cobranza.' });
    const minCash = Math.min(...b.bs.cash); if (minCash < 0) I.push({ level: 'danger', title: `Déficit de caja proyectado (mínimo ${fmt(minCash)})`, text: 'El modelo requiere financiamiento adicional o una caja inicial mayor (ingrésala en Supuestos).' });
  }
  return I;
}

function bsInsights(M) {
  const I = [], b = M.biz; if (!b) return I;
  const R = b.ratios, bs = b.bs, e = bs.assets.length - 1;
  I.push(R.maxCheck < 1 ? { level: 'success', title: 'Balance cuadrado en todos los períodos', text: 'Activos = Pasivos + Patrimonio. Caja del balance = caja del estado de flujo de efectivo.' } : { level: 'danger', title: `Descuadre máximo de ${fmt(R.maxCheck)}`, text: 'Revisa el balance inicial cargado: no cumple Activos = Pasivos + Patrimonio.' });
  if (R.currentRatio != null) I.push(R.currentRatio >= 1.5 ? { level: 'success', title: `Liquidez corriente ${R.currentRatio.toFixed(2)}x`, text: 'Sobre 1,5x: holgura para cubrir obligaciones de corto plazo.' } : R.currentRatio >= 1 ? { level: 'warning', title: `Liquidez corriente ${R.currentRatio.toFixed(2)}x`, text: 'Justa. Recomendación: elevar caja o extender plazos con proveedores.' } : { level: 'danger', title: `Liquidez corriente ${R.currentRatio.toFixed(2)}x`, text: 'Bajo 1x: los pasivos de corto plazo superan los activos corrientes. Acción inmediata de capital de trabajo.' });
  if (R.ndEbitda != null) I.push(R.ndEbitda <= 0 ? { level: 'success', title: 'Posición de caja neta', text: 'Caja ≥ deuda: capacidad para financiar crecimiento o adquisiciones sin deuda adicional.' } : R.ndEbitda < 2 ? { level: 'success', title: `Deuda neta / EBITDA ${R.ndEbitda.toFixed(1)}x`, text: 'Endeudamiento conservador (< 2x).' } : R.ndEbitda < 3.5 ? { level: 'warning', title: `Deuda neta / EBITDA ${R.ndEbitda.toFixed(1)}x`, text: 'Moderado-alto. Evita nueva deuda sin mejorar el EBITDA.' } : { level: 'danger', title: `Deuda neta / EBITDA ${R.ndEbitda.toFixed(1)}x`, text: 'Sobre 3,5x: riesgo de covenants. Prioriza desapalancamiento.' });
  if (R.coverage != null) I.push(R.coverage >= 4 ? { level: 'success', title: `Cobertura de intereses ${R.coverage.toFixed(1)}x`, text: 'Holgada.' } : R.coverage >= 2 ? { level: 'warning', title: `Cobertura de intereses ${R.coverage.toFixed(1)}x`, text: 'Aceptable pero sensible a caídas de EBITDA.' } : { level: 'danger', title: `Cobertura de intereses ${R.coverage.toFixed(1)}x`, text: 'Bajo 2x: los intereses consumen la mayor parte del EBITDA.' });
  if (R.nwcPctRev != null && !b.nwcFixed) I.push({ level: R.nwcPctRev > 0.25 ? 'warning' : 'info', title: `Capital de trabajo = ${pc(R.nwcPctRev, 1)} de los ingresos anuales`, text: R.ccc != null ? `Ciclo de conversión de caja supuesto: ${R.ccc} días (DSO ${M.Aeff.dso} + DIO ${M.Aeff.dio} − DPO ${M.Aeff.dpo}). Cada día de DSO que reduzcas libera ≈ ${fmt(R.revAnn / 365)}.` : '' });
  if (R.roe != null) I.push({ level: R.roe > 0.15 ? 'success' : R.roe > 0.08 ? 'warning' : 'info', title: `ROE ${pc(R.roe, 1)} · ROA ${pc(R.roa, 1)}`, text: 'Calculados sobre patrimonio y activos promedio del período proyectado.' });
  if (!b.openFromFile && M.hasPL) I.push({ level: 'info', title: 'Balance inicial supuesto', text: 'No se cargó la hoja "Balance". Se asumió caja, activo fijo y deuda iniciales según el panel de supuestos. Súbela para reemplazar estos saldos por los reales.' });
  return I;
}

function qualityChecks(M, store) {
  const Q = [], P = M.proj;
  if (store.uncalc) Q.push({ level: 'warning', title: `${store.uncalc} fórmulas sin valor calculado`, text: 'El archivo fue generado sin recalcular. Ábrelo en Excel, guárdalo y vuelve a subirlo para que la app lea los resultados.' });
  store.inventory.filter(i => i.role === 'pl' && i.status === 'vacía').forEach(i => Q.push({ level: 'info', title: `Pestaña "${i.sheet}" vacía`, text: 'Reconocida pero sin datos; se omitió del análisis.' }));
  store.inventory.filter(i => i.status === 'no reconocida').forEach(i => Q.push({ level: 'warning', title: `Pestaña "${i.sheet}" no reconocida`, text: 'Renómbrala según la plantilla (P&L …, Supuestos, Flujo de Caja, Balance, M&A) para que se incluya.' }));
  if (P) {
    const r = P.reconcile;
    if (P.fclError) {
      Q.unshift({ level: 'danger', title: 'Error de fórmula en la hoja "Flujo de Caja" de tu Excel', text: `La fila "FLUJO DE CAJA LIBRE" no equivale a NOPAT + D&A − CAPEX − ΔCT − pre-operativos + valor residual (difiere en ${P.fclError.months} de ${P.fclError.total} meses): suma también las filas EBIT e Impuestos sobre EBIT, sobrestimando el flujo. VAN del archivo ${fmt(r.vanFile != null ? r.vanFile : P.fclError.fileVan)} vs VAN corregido ${fmt(P.van)}. Corrección sugerida en cada mes: =C6+SUM(C7:C11) en lugar de =SUM(C4:C11). La app ya usa el flujo corregido.` });
    } else if (r.vanFile != null) Q.push(r.diffPct < 0.005 ? { level: 'success', title: 'VAN conciliado con tu Excel', text: `VAN de la app ${fmt(r.vanApp)} = VAN del archivo ${fmt(r.vanFile)} (horizonte ${P.H} meses, tasa mensual equivalente al WACC anual).` } : { level: 'warning', title: `VAN difiere del Excel (${(r.diffPct * 100).toFixed(1)}%)`, text: `App ${fmt(r.vanApp)} vs archivo ${fmt(r.vanFile)}. Revisa horizonte, WACC y periodicidad de descuento.` });
    if (!P.fclError && r.tirFile != null && r.tirApp != null && Math.abs(r.tirFile - r.tirApp) > 0.005) Q.push({ level: 'warning', title: 'TIR difiere del Excel', text: `App ${pc(r.tirApp, 2)} vs archivo ${pc(r.tirFile, 2)}.` });
    if (r.engineFclMaxDiff != null) Q.push(r.engineFclMaxDiff < 1 ? { level: 'success', title: 'Motor de sensibilidad validado', text: 'La re-simulación reproduce tu flujo de caja (corregido) mes a mes; los sliders, escenarios y el tornado son confiables.' } : { level: 'warning', title: 'La re-simulación no replica exactamente tu flujo', text: `Diferencia máxima ${fmt(r.engineFclMaxDiff)}. Los escenarios/sensibilidades son aproximados porque tu hoja usa una lógica distinta a la estándar.` });
    if (P.sup && P.sup.nwc0 && P.fcl[P.H] && !(store.ev.cf && store.ev.cf.series['(-) cambio en capital de trabajo'] && byMonth(store.ev.cf, ['(-) cambio en capital'], P.H)[P.H])) Q.push({ level: 'info', title: 'El capital de trabajo inicial no se recupera al final del horizonte', text: `El flujo del mes ${P.H} incluye valor residual de activos fijos pero no la liberación de ${fmt(P.sup.nwc0)} de capital de trabajo. Si corresponde, el VAN real sería mayor.` });
    const e = store.ev.esc;
    (P.scenarios || []).filter(sc => sc.note).forEach(sc => Q.push({ level: 'warning', title: `Escenario ${sc.label}: inversión inconsistente`, text: `${sc.note}. Un escenario ${sc.label.toLowerCase()} con una inversión de otro orden de magnitud suele ser un cero faltante o un error de unidad; corrígelo en la hoja Escenarios (fila "Inversión Inicial en Activos Fijos").` }));
    const v = store.ev.val;
    if (v && isNum(v.ebitdaProm) && P.engine) {
      const mine = avg(P.engine.ebitda.slice(1, 61)) * 12;
      if (mine > 0 && (v.ebitdaProm / mine > 3 || v.ebitdaProm / mine < 1 / 3)) Q.push({ level: 'warning', title: 'Valoración del Excel inconsistente', text: `La hoja Valoración usa un EBITDA promedio anual de ${fmt(v.ebitdaProm)}, pero el modelo implica ≈ ${fmt(mine)} por año. El EV por múltiplos del archivo (${fmt(v.evMult)}) no es comparable con su DCF (${fmt(v.evDcf)}). La app recalcula la valoración en la pestaña M&A.` });
    }
    if (isNum(P.sup.life) && P.sup.life * 12 < P.H) Q.push({ level: 'info', title: 'Vida útil menor al horizonte', text: 'La depreciación termina antes del fin del horizonte; verifica reinversión.' });
  }
  if (M.hasPL) {
    const pl = M.plSel.pl;
    if (!pl.depFound || sum(pl.depreciacion.slice(0, 12)) === 0) Q.push({ level: 'warning', title: 'P&L sin depreciación/amortización', text: 'Todas las filas de D&A están en cero: EBITDA = EBIT y el capex derivado es 0.' });
    if (pl.gastosOp && sum(pl.gastosOp.slice(0, 12)) === 0) Q.push({ level: 'warning', title: 'Gastos operacionales en cero', text: 'El P&L no informa gastos de ventas ni administración; el margen operacional podría estar sobreestimado.' });
  }
  if (M.biz && M.biz.ratios.maxCheck >= 1) Q.push({ level: 'danger', title: 'Balance inicial no cuadra', text: `Descuadre de ${fmt(M.biz.ratios.maxCheck)}; se muestra como "ajuste de balance inicial" en el patrimonio.` });
  return Q;
}

/* ============================================================
   8. EXPORTACIÓN DEL MODELO COMPLETO (hojas con valores)
   ============================================================ */
AM.exportSheets = function (M) {
  const sheets = [];
  const b = M.biz, P = M.proj;
  const rows0 = [['AMBERES INTELLIGENCE™ — Modelo exportado'], ['Fecha', new Date().toISOString().slice(0, 10)], [], ['Módulo', 'Estado', 'Detalle']];
  Object.entries({ 'P&L': M.sources.pl, 'Evaluación de proyectos': M.sources.proj, 'Flujo de caja': M.sources.cf, 'Balance': M.sources.bs, 'M&A': M.sources.ma }).forEach(([k, v]) => rows0.push([k, v.state, v.note]));
  rows0.push([], ['KPIs de la Visión General']);
  (M.overview ? M.overview.kpis : []).forEach(k => rows0.push([k.label, k.value, k.sub]));
  sheets.push({ name: 'Resumen', aoa: rows0 });
  if (b) {
    const L = b.labels, hdr = ['Concepto'].concat(L), row = (n, a) => [n].concat(a.map(v => Math.round(v * 100) / 100));
    sheets.push({ name: 'Flujo de Caja (negocio)', aoa: [['Estado de flujo de efectivo — método indirecto'], hdr,
      row('Utilidad neta', b.ser.ni), row('(+) Depreciación y amortización', b.ser.dep), row('(−) Δ Capital de trabajo', b.dNWC.map(v => -v)),
      row('FLUJO OPERACIONAL', b.cfo), row('(−) Capex', b.cfi), row('FLUJO DE INVERSIÓN', b.cfi), row('(−) Amortización de deuda', b.repay.map(v => -v)), row('(−) Dividendos', b.div.map(v => -v)),
      row('FLUJO FINANCIERO', b.cff), row('VARIACIÓN NETA DE CAJA', b.net), [], row('Flujo económico (FCFF)', b.fcff), row('Flujo al accionista (FCFE)', b.fcfe)] });
    const h2 = ['Concepto', 'Inicial'].concat(L), r2 = (n, a) => [n].concat(a.map(v => Math.round(v * 100) / 100));
    sheets.push({ name: 'Balance (negocio)', aoa: [['Balance proyectado'], h2, r2('Caja', b.bs.cash), r2(b.nwcFixed ? 'Capital de trabajo neto' : 'Cuentas por cobrar', b.bs.ar), ...(b.nwcFixed ? [] : [r2('Inventarios', b.bs.inv)]), r2('Activo fijo neto', b.bs.ppe), r2('TOTAL ACTIVOS', b.bs.assets),
      ...(b.nwcFixed ? [] : [r2('Cuentas por pagar', b.bs.ap)]), r2('Deuda financiera', b.bs.debt), r2('TOTAL PASIVOS', b.bs.liab), r2('Capital', b.bs.capital), r2('Resultados acumulados', b.bs.re), r2('Ajuste balance inicial', b.bs.adj), r2('TOTAL PATRIMONIO', b.bs.equity), r2('Cuadre (A − P − Pat.)', b.bs.check)] });
  }
  if (P) {
    const aoa = [['Proyecto', P.name], ['VAN', P.van], ['TIR anual', P.tirA], ['WACC', P.wacc], ['Payback (mes)', P.payback || 'No recupera'], [], ['Mes', 'Flujo libre', 'Acumulado', 'Flujo financiero']];
    P.fcl.forEach((v, t) => aoa.push([t, Math.round(v), Math.round(P.acc[t]), P.fcfe ? Math.round(P.fcfe[t]) : '']));
    sheets.push({ name: 'Proyecto', aoa });
    if (P.scenarios) sheets.push({ name: 'Escenarios', aoa: [['Escenario', 'Crecimiento', '% Costos var.', 'WACC', 'Inversión', 'VAN', 'TIR anual', 'Payback'], ...P.scenarios.map(s => [s.label, s.g, s.cv, s.wacc, s.capex, Math.round(s.van), s.tirA, s.payback || 'No recupera'])] });
  }
  const mc = M.maCalc, p = M.maParams;
  sheets.push({ name: 'M&A (DCF)', aoa: [['Parámetro', 'Valor', 'Fuente'], ...Object.keys(p).map(k => [AM.MA_LABELS[k] || k, p[k], M.maSrc[k]]), [], ['Resultado', 'Valor'], ['WACC', mc.wacc.wacc], ['EV (DCF)', Math.round(mc.dcf.ev)], ['Equity (DCF)', Math.round(mc.equity)], ['EV (múltiplos)', Math.round(mc.evMult)], ['Equity con prima de control', Math.round(mc.equityControl)], [], ['Año', 'Ingresos', 'EBITDA', 'FCF'], ...mc.rows.map(r => [r.y, Math.round(r.rev), Math.round(r.ebitda), Math.round(r.fcf)])] });
  sheets.push({ name: 'Supuestos del modelo', aoa: [['Supuesto', 'Valor'], ['DSO (días)', M.Aeff.dso], ['DIO (días)', M.Aeff.dio], ['DPO (días)', M.Aeff.dpo], ['Capex % ingresos', b ? b.capexPct : ''], ['Caja inicial', b ? b.cash0 : ''], ['Deuda inicial', b ? b.debt0 : ''], ['Pago de deuda mensual', M.Aeff.repay], ['Payout dividendos', M.Aeff.payout], ['Tasa de impuesto', M.Aeff.tax]] });
  sheets.push({ name: 'Inventario de pestañas', aoa: [['Archivo', 'Pestaña', 'Módulo', 'Estado', 'Detalle'], ...M.inventory.map(i => [i.file, i.sheet, AM.ROLE_LABEL[i.role] || i.role, i.status, i.detail])] });
  if (M.quality.length) sheets.push({ name: 'Calidad del modelo', aoa: [['Nivel', 'Hallazgo', 'Detalle'], ...M.quality.map(q => [q.level, q.title, q.text])] });
  return sheets;
};

AM._engine = { deriveBusiness, buildProject, pickPL, serFromPL, projectMA, dcfEV };
return AM;
});

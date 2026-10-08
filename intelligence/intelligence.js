// ============================================================
let evalData = null;
let plData   = null;
let charts   = {};
let loadedFiles = [];
let currentEmpresa = 'servicios';
let baseInversion = 50000000;
let baseTax = 27;

// PALETA AMBERES
const AmberesColors = {
    primary: '#c89b3c',
    secondary: '#4f7c82',
    accent1: '#d6ad55',
    accent2: '#7a6b55',
    accent3: '#e4c77a',
    neutral1: '#a9c2cc',
    neutral2: '#7f959d',
    success: '#4ea878',
    warning: '#c89b3c',
    danger: '#d86767',
    dark: '#102a36',
    light: '#f4efe5'
};

// Configuración global de Chart.js - Estilo Amberes
try {
    if (typeof Chart !== 'undefined') {
        Chart.defaults.font = Chart.defaults.font || {};
        Chart.defaults.font.family = "'Inter', 'Helvetica Neue', Arial, sans-serif";
        Chart.defaults.font.size = 12;
        Chart.defaults.color = '#b9c7c9';

        Chart.defaults.plugins = Chart.defaults.plugins || {};
        Chart.defaults.plugins.legend = Chart.defaults.plugins.legend || {};
        Chart.defaults.plugins.legend.display = true;
        Chart.defaults.plugins.legend.position = 'top';
        Chart.defaults.plugins.legend.align = 'start';
        Chart.defaults.plugins.legend.labels = Object.assign({}, Chart.defaults.plugins.legend.labels, {
            usePointStyle: true,
            padding: 15,
            font: { size: 11, weight: '600' }
        });

        Chart.defaults.plugins.tooltip = Object.assign({}, Chart.defaults.plugins.tooltip, {
            backgroundColor: 'rgba(26, 26, 26, 0.95)',
            padding: 12,
            cornerRadius: 4,
            titleFont: { size: 12, weight: '700' },
            bodyFont: { size: 11 }
        });

        Chart.defaults.scale = Chart.defaults.scale || {};
        Chart.defaults.scale.grid = Object.assign({}, Chart.defaults.scale.grid, {
            display: true,
            drawBorder: false,
            color: 'rgba(255,255,255,0.07)'
        });
        Chart.defaults.scale.ticks = Object.assign({}, Chart.defaults.scale.ticks, {
            padding: 8
        });
    }
} catch (e) {
    console.warn('No se pudo aplicar el estilo Amberes a Chart.js:', e);
}


// ============================================================
// DRAG & DROP
// ============================================================
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('dragover'); handleFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', e => handleFiles(e.target.files));

// ============================================================
// MANEJO DE ARCHIVOS
// ============================================================
async function handleFiles(files) {
    if (!files.length) return;
    document.getElementById('processingOverlay').classList.add('active');
    for (let file of files) {
        try {
            const wb  = await readExcel(file);
            const typ = detectType(wb);
            if (typ === 'evaluation') {
                evalData = processEval(wb);
                renderOverview();
                renderCF();
            } else if (typ === 'pl') {
                plData = processPL(wb);
                renderPL();
            }
            loadedFiles = loadedFiles.filter(f => f.type !== typ);
            loadedFiles.push({ name: file.name, type: typ });
        } catch(e) {
            console.error(e);
            loadedFiles.push({ name: file.name, type: 'error' });
        }
    }
    updateFilesDisplay();
    document.getElementById('processingOverlay').classList.remove('active');
}

function readExcel(file) {
    return new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = e => {
            try { res(XLSX.read(new Uint8Array(e.target.result), { type:'array' })); }
            catch(err) { rej(err); }
        };
        r.onerror = rej;
        r.readAsArrayBuffer(file);
    });
}

// ============================================================
// DETECCIÓN AUTOMÁTICA DE TIPO
// ============================================================
function detectType(wb) {
    const names = wb.SheetNames.map(s => s.toLowerCase());
    // P&L: tiene pestañas con "p&l" o ("productiva" y "servicios") o ("real" y "budget")
    const hasPL = names.some(s => s.includes('p&l') || (s.includes('productiva') || s.includes('servicios')));
    const hasScenarios = names.some(s => s.includes('real') || s.includes('budget') || s.includes('forecast'));
    // Evaluación: tiene "flujo de caja" o "supuestos" o "valoraci"
    const hasEval = names.some(s => s.includes('flujo de caja') || s.includes('supuestos') || s.includes('valoraci') || s.includes('resumen'));

    if (hasPL && hasScenarios && !hasEval) return 'pl';
    if (hasEval) return 'evaluation';
    if (hasPL) return 'pl';
    return 'unknown';
}

function updateFilesDisplay() {
    const container = document.getElementById('filesLoadedContainer');
    const list = document.getElementById('filesList');
    if (!loadedFiles.length) { container.style.display='none'; return; }
    container.style.display = 'block';
    list.innerHTML = '';
    loadedFiles.forEach(f => {
        const icons = { evaluation:'📊', pl:'📈', error:'⚠️', unknown:'❓' };
        const labels = { evaluation:'Evaluación de Proyectos', pl:'P&L — Estado de Resultados', error:'Error al leer', unknown:'Tipo desconocido' };
        const pill = document.createElement('div');
        pill.className = `file-pill ${f.type}`;
        pill.innerHTML = ''; const icon = document.createElement('span'); icon.className='pill-icon'; icon.textContent=icons[f.type]||'📄'; const name=document.createElement('strong'); name.textContent=f.name; const type=document.createElement('span'); type.style.cssText='color:#a9c2cc;font-size:11px'; type.textContent=` · ${labels[f.type]||f.type}`; pill.append(icon,name,type);
        list.appendChild(pill);
    });
}

// ============================================================
// PROCESAR EVALUACIÓN DE PROYECTOS
// ============================================================
function processEval(wb) {
    const d = { supuestos:{}, ingresos:[], costos:[], ebitda:[], ebit:[], cashflow:[], cfFinanciero:[], inversion:0, van:0, tir:0 };

    // Supuestos
    try {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Supuestos'], {header:1});
        rows.forEach(r => {
            if (!r[0] || !r[1]) return;
            const k = r[0].toString().toLowerCase();
            if (k.includes('activos fijos')) d.inversion = +r[1] || 0;
        });
        d.supuestos.inversionBase = d.inversion;
    } catch(e){}

    // Estado de Resultados
    try {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Estado de Resultados'], {header:1});
        rows.slice(2).forEach(r => {
            const k = (r[0]||'').toString().toLowerCase();
            if (k.includes('ingresos') && !k.includes('nopat') && !k.includes('otros')) d.ingresos = r.slice(1).map(v=>+v||0);
            if (k.includes('ebitda')) d.ebitda = r.slice(1).map(v=>+v||0);
            if (k.includes('ebit') && !k.includes('ebitda')) d.ebit = r.slice(1).map(v=>+v||0);
            if (k.includes('costos variables') || k.includes('costo de ventas total')) d.costos = r.slice(1).map(v=>+v||0);
        });
    } catch(e){}

    // Flujo de Caja
    try {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Flujo de Caja'], {header:1});
        rows.slice(2).forEach(r => {
            const k = (r[0]||'').toString().toLowerCase();
            if (k.includes('flujo de caja libre') || k.includes('flujo caja libre')) d.cashflow = r.slice(1).map(v=>+v||0);
            if (k.includes('flujo de caja del proyecto') || k.includes('financiero')) d.cfFinanciero = r.slice(1).map(v=>+v||0);
        });
    } catch(e){}

    // Resumen VAN/TIR
    try {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets['Resumen'], {header:1});
        rows.forEach(r => {
            const k = (r[0]||'').toString().toLowerCase();
            if (k.includes('van') || k.includes('valor actual neto')) d.van = +r[1] || 0;
            if (k.includes('tir') && k.includes('tasa interna')) d.tir = +r[1] || 0;
        });
    } catch(e){}

    return d;
}

// ============================================================
// PROCESAR P&L — LEER POR NÚMERO DE FILA (ROBUSTO)
// ============================================================
function processPL(wb) {
    const result = { productiva:{}, servicios:{} };

    const sheetMap = {
        'productiva': {
            real:     wb.SheetNames.find(s => s.toLowerCase().includes('productiva') && s.toLowerCase().includes('real')),
            budget:   wb.SheetNames.find(s => s.toLowerCase().includes('productiva') && s.toLowerCase().includes('budget')),
            forecast: wb.SheetNames.find(s => s.toLowerCase().includes('productiva') && s.toLowerCase().includes('forecast'))
        },
        'servicios': {
            real:     wb.SheetNames.find(s => s.toLowerCase().includes('servicios') && s.toLowerCase().includes('real')),
            budget:   wb.SheetNames.find(s => s.toLowerCase().includes('servicios') && s.toLowerCase().includes('budget')),
            forecast: wb.SheetNames.find(s => s.toLowerCase().includes('servicios') && s.toLowerCase().includes('forecast'))
        }
    };

    ['productiva','servicios'].forEach(tipo => {
        result[tipo] = {};
        ['real','budget','forecast'].forEach(esc => {
            const sName = sheetMap[tipo][esc];
            if (!sName || !wb.Sheets[sName]) { result[tipo][esc] = null; return; }
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sName], {header:1});
            const parsed = {};

            // Buscar filas clave por nombre exacto (robusto)
            const KEYS = {
                ingresosBrutos:   ['total ingresos brutos'],
                ingresosNetos:    ['ingresos netos'],
                costoTotal:       ['costo de ventas total','costo de servicios total'],
                utilidadBruta:    ['utilidad bruta'],
                margenBruto:      ['margen bruto %'],
                gastosOp:         ['total gastos operacionales'],
                ebit:             ['utilidad operacional (ebit)','utilidad operacional'],
                margenOp:         ['margen operacional %'],
                ebt:              ['utilidad antes de impuestos (ebt)','utilidad antes de impuesto'],
                impuesto:         ['(-) impuesto'],
                utilidadNeta:     ['utilidad neta'],
                margenNeto:       ['margen neto %']
            };

            rows.forEach(row => {
                const label = (row[0]||'').toString().toLowerCase().trim();
                const vals  = row.slice(1,14).map(v => {
                    const n = parseFloat(v);
                    return isNaN(n) ? 0 : n;
                });
                Object.entries(KEYS).forEach(([key, patterns]) => {
                    if (patterns.some(p => label.includes(p))) {
                        parsed[key] = vals; // 12 meses + total
                    }
                });
            });

            // No imputar EBITDA si el archivo no lo contiene.
            // Un dato derivado debe provenir de una metodología explícita, no de una aproximación arbitraria.
            if (!parsed.ebitda) parsed.ebitda = [];

            result[tipo][esc] = parsed;
        });
    });

    return result;
}

// ============================================================
// HELPERS
// ============================================================
function fmt(v) {
    if (v === undefined || v === null || isNaN(v)) return '$0';
    const abs = Math.abs(v);
    let str;
    if (abs >= 1e9) str = '$' + (v/1e9).toFixed(2) + 'B';
    else if (abs >= 1e6) str = '$' + (v/1e6).toFixed(2) + 'M';
    else if (abs >= 1e3) str = '$' + (v/1e3).toFixed(1) + 'K';
    else str = '$' + v.toFixed(0);
    return str;
}
function pct(v) { return isNaN(v) ? '0.00%' : (v*100).toFixed(2)+'%'; }
function pctN(v) { return isNaN(v) ? '0.00%' : v.toFixed(2)+'%'; }
function sum(arr) { return (arr||[]).reduce((a,b)=>a+b,0); }
function avg(arr) { return arr.length ? sum(arr)/arr.length : 0; }
function irr(cf) {
    if (!Array.isArray(cf) || cf.length < 2) return 0;
    const values = cf.map(Number).filter(Number.isFinite);
    if (values.length < 2 || !values.some(v => v > 0) || !values.some(v => v < 0)) return 0;
    let low = -0.9999, high = 10;
    const npv = rate => cf.reduce((acc,v,t) => acc + (Number(v)||0) / Math.pow(1 + rate, t), 0);
    let fLow = npv(low), fHigh = npv(high);
    if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || fLow * fHigh > 0) return 0;
    for (let i=0;i<200;i++) {
        const mid = (low + high) / 2;
        const fMid = npv(mid);
        if (Math.abs(fMid) < 0.000001) return mid * 100;
        if (fLow * fMid <= 0) { high = mid; fHigh = fMid; }
        else { low = mid; fLow = fMid; }
    }
    return ((low + high) / 2) * 100;
}
function destroyChart(id) { if(charts[id]) { charts[id].destroy(); delete charts[id]; } }

function switchTab(id, btn) {
    document.querySelectorAll('.dashboard-content').forEach(c=>c.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    btn.classList.add('active');
    if (id === 'manda' && window.__maState) renderMACharts(window.__maState);
}

function switchEmpresa(tipo) {
    currentEmpresa = tipo;
    document.getElementById('btnProductiva').classList.toggle('active', tipo==='productiva');
    document.getElementById('btnServicios').classList.toggle('active', tipo==='servicios');
    renderPL();
}

// ============================================================
// RENDER — VISIÓN GENERAL
// ============================================================
function renderOverview() {
    if (!evalData) return;
    document.getElementById('overviewEmpty').style.display = 'none';
    document.getElementById('overviewContent').style.display = 'block';

    const van  = evalData.van || 0;
    const tir  = evalData.tir ? evalData.tir*100 : irr(evalData.cashflow);
    const wacc = 10;
    const avgEBITDA = avg(evalData.ebitda);
    const avgEBIT   = avg(evalData.ebit);
    const avgRev    = avg(evalData.ingresos) || 1;
    const breakEven = evalData.cashflow.findIndex((v,i,a) => a.slice(0,i+1).reduce((x,y)=>x+y,0) >= 0);
    const ingresosBreakEven = breakEven >= 0 && evalData.ingresos[breakEven] 
        ? evalData.ingresos[breakEven] 
        : breakEven >= 0 
        ? evalData.ingresos.slice(0, breakEven + 1).reduce((a,b)=>a+b,0) / (breakEven + 1)
        : sum(evalData.ingresos);

    baseInversion = evalData.inversion || 50000000;
    document.getElementById('sliderInv').value = baseInversion;
    document.getElementById('invDisplay').innerHTML = `${fmt(baseInversion)} <span>base</span>`;

    const taxVal = parseFloat(document.getElementById('sliderTax').value) || 27;
    baseTax = taxVal;

    // KPIs
    const kpiGrid = document.getElementById('kpiGrid');
    kpiGrid.innerHTML = '';
    const kpis = [
        { label:'VAN', value: fmt(van), sub: van>0?'✓ Proyecto viable':'⚠ Revisar viabilidad', cls: van>0?'success':'danger', sem: van>0?'green':'red' },
        { label:'TIR', value: pctN(tir), sub: tir>wacc?`✓ Supera WACC (${wacc}%)`:`⚠ Bajo WACC (${wacc}%)`, cls: tir>wacc?'success':'danger', sem: tir>wacc?'green':'red' },
        { label:'WACC', value: pctN(wacc), sub:'Costo de capital', cls:'', sem:'yellow' },
        { label:'EBITDA Promedio', value: fmt(avgEBITDA), sub:`Margen: ${pctN(avgEBITDA/avgRev*100)}`, cls:'success', sem:'green' },
        { label:'EBIT Promedio', value: fmt(avgEBIT), sub:`Margen: ${pctN(avgEBIT/avgRev*100)}`, cls:'success', sem:'green' },
        { label:'Punto de Equilibrio', value: breakEven>=0?`Mes ${breakEven}`:'No alcanzado', sub: breakEven>=0?`Ingresos necesarios: ${fmt(ingresosBreakEven)}`:'Revisar supuestos', cls: breakEven>=0&&breakEven<=24?'success': breakEven>24?'warning':'danger', sem: breakEven>=0&&breakEven<=24?'green': breakEven>24?'yellow':'red' }
    ];
    kpis.forEach(k => {
        const el = document.createElement('div');
        el.className = `metric-card ${k.cls}`;
        el.innerHTML = `<div class="metric-label"><span class="semaforo ${k.sem}"></span>${k.label}</div>
                        <div class="metric-value">${k.value}</div>
                        <div class="metric-sub ${k.sub.includes('✓')?'positive':k.sub.includes('⚠')?'negative':''}">${k.sub}</div>`;
        kpiGrid.appendChild(el);
    });

    // Charts
    renderCFChart();
    renderEbitdaChart();
    updateSensitivity();
    renderOverviewInsights(van, tir, wacc, breakEven);
}

function renderCFChart() {
    destroyChart('chartCF');
    if (!evalData?.cashflow?.length) return;
    const data = evalData.cashflow.slice(0,24);
    const labels = data.map((_,i)=>`M${i+1}`);
    const accumulated = data.map((_,i)=>data.slice(0,i+1).reduce((a,b)=>a+b,0));
    
    charts['chartCF'] = new Chart(document.getElementById('chartCF'), {
        type:'bar',
        data:{ 
            labels, 
            datasets:[{
                label:'Flujo de Caja Libre',
                data,
                backgroundColor: data.map(v => v >= 0 ? AmberesColors.success : AmberesColors.danger),
                borderColor: data.map(v => v >= 0 ? AmberesColors.success : AmberesColors.danger),
                borderWidth: 0,
                borderRadius: 3,
                barPercentage: 0.75
            },{
                type:'line', 
                label:'Flujo Acumulado',
                data: accumulated,
                borderColor: AmberesColors.accent1,
                backgroundColor: 'transparent',
                borderWidth: 3,
                fill: false,
                tension: 0.1,
                pointRadius: 0,
                pointHoverRadius: 5,
                pointBackgroundColor: AmberesColors.accent1,
                pointBorderColor: '#fff',
                pointBorderWidth: 2
            }]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins:{ 
                legend: { 
                    labels: { 
                        color: '#b9c7c9', 
                        font: { size: 11, weight: '600' },
                        padding: 15,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    }, 
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });
}

function renderEbitdaChart() {
    destroyChart('chartEbitda');
    if (!evalData?.ebitda?.length) return;
    const n = Math.min(12, evalData.ebitda.length);
    const labels = Array.from({length:n},(_,i)=>`M${i+1}`);
    charts['chartEbitda'] = new Chart(document.getElementById('chartEbitda'), {
        type:'bar',
        data:{ 
            labels, 
            datasets:[
                { 
                    label:'EBITDA', 
                    data: evalData.ebitda.slice(0,n), 
                    backgroundColor: AmberesColors.primary,
                    borderColor: AmberesColors.primary,
                    borderWidth: 0,
                    borderRadius: 3,
                    barPercentage: 0.7
                },
                { 
                    label:'EBIT', 
                    data: evalData.ebit.slice(0,n), 
                    backgroundColor: AmberesColors.secondary,
                    borderColor: AmberesColors.secondary,
                    borderWidth: 0,
                    borderRadius: 3,
                    barPercentage: 0.7
                }
            ]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            plugins:{ 
                legend: { 
                    labels: { 
                        color: '#b9c7c9', 
                        font: { size: 11, weight: '600' },
                        padding: 15,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    }, 
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });
}

function updateSensitivity() {
    if (!evalData) return;
    const invNew  = parseFloat(document.getElementById('sliderInv').value);
    const taxNew  = parseFloat(document.getElementById('sliderTax').value);
    const invBase = evalData.inversion || 50000000;
    const taxBase = 27;

    document.getElementById('invDisplay').innerHTML = `${fmt(invNew)} <span>${invNew > invBase ? '▲ mayor inversión' : invNew < invBase ? '▼ menor inversión' : 'base'}</span>`;
    document.getElementById('taxDisplay').innerHTML = `${taxNew}% <span>${taxNew > taxBase ? '▲ mayor carga fiscal' : taxNew < taxBase ? '▼ menor carga fiscal' : 'base Chile'}</span>`;

    // Impacto inversión en VAN (diferencia)
    const invDelta = invBase - invNew;
    const vanImpacto = evalData.van + invDelta;
    const invImpEl = document.getElementById('invImpact');
    if (invNew < invBase) {
        invImpEl.className = 'sens-impact positive';
        invImpEl.textContent = `✓ VAN estimado mejoraría a ${fmt(vanImpacto)} (ahorro de ${fmt(invBase - invNew)})`;
    } else if (invNew > invBase) {
        invImpEl.className = 'sens-impact negative';
        invImpEl.textContent = `⚠ VAN estimado caería a ${fmt(vanImpacto)} (inversión adicional de ${fmt(invNew - invBase)})`;
    } else {
        invImpEl.className = 'sens-impact neutral';
        invImpEl.textContent = 'Sin cambio respecto a la inversión base';
    }

    // Impacto tasa impuesto en utilidad neta estimada
    const avgEBIT = avg(evalData.ebit) * 12;
    const utilBase = avgEBIT * (1 - taxBase/100);
    const utilNew  = avgEBIT * (1 - taxNew/100);
    const utilDelta = utilNew - utilBase;
    const taxImpEl = document.getElementById('taxImpact');
    if (taxNew < taxBase) {
        taxImpEl.className = 'sens-impact positive';
        taxImpEl.textContent = `✓ Utilidad neta anual estimada sube en ${fmt(utilDelta)} vs tasa Chile 27%`;
    } else if (taxNew > taxBase) {
        taxImpEl.className = 'sens-impact negative';
        taxImpEl.textContent = `⚠ Utilidad neta anual estimada cae en ${fmt(Math.abs(utilDelta))} vs tasa Chile 27%`;
    } else {
        taxImpEl.className = 'sens-impact neutral';
        taxImpEl.textContent = `Tasa base Chile 27% — Utilidad neta anual estimada: ${fmt(utilBase)}`;
    }
}

function renderOverviewInsights(van, tir, wacc, be) {
    let html = '';
    html += van>0
        ? `<div class="insight-item success"><strong>✓ Proyecto Financieramente Viable</strong><br>VAN positivo de ${fmt(van)} indica creación de valor. Proceder con evaluación de riesgos operativos.</div>`
        : `<div class="insight-item danger"><strong>⚠ VAN Negativo — Proyecto No Viable en Escenario Base</strong><br>VAN de ${fmt(van)}. Revisar supuestos, buscar eficiencias o evaluar escenario optimista. Usa el slider de Inversión Inicial para sensibilizar.</div>`;
    html += tir>wacc
        ? `<div class="insight-item success"><strong>✓ TIR (${pctN(tir)}) Supera el WACC (${pctN(wacc)})</strong><br>Spread de ${pctN(tir-wacc)}. El proyecto genera rentabilidad superior al costo de financiamiento.</div>`
        : `<div class="insight-item warning"><strong>⚠ TIR por Debajo del WACC</strong><br>TIR ${pctN(tir)} vs WACC ${pctN(wacc)}. El proyecto no compensa el riesgo asumido en escenario base.</div>`;
    html += be>=0&&be<=24
        ? `<div class="insight-item success"><strong>✓ Recuperación en ${be} Meses</strong><br>Punto de equilibrio alcanzado dentro de los primeros 2 años.</div>`
        : be>24
        ? `<div class="insight-item warning"><strong>⚠ Recuperación Tardía — Mes ${be}</strong><br>Considera estrategias para acelerar generación de caja.</div>`
        : `<div class="insight-item danger"><strong>⚠ Punto de Equilibrio No Alcanzado</strong><br>En el horizonte evaluado no se recupera la inversión. Revisar modelo.</div>`;
    document.getElementById('insightsOverview').innerHTML = html;
}

// ============================================================
// RENDER — FLUJO DE CAJA
// ============================================================
function renderCF() {
    if (!evalData) return;
    document.getElementById('cfEmpty').style.display = 'none';
    document.getElementById('cfContent').style.display = 'block';

    const cf   = evalData.cashflow;
    const cfF  = evalData.cfFinanciero.length ? evalData.cfFinanciero : cf.map(v=>v*0.92);
    const totOp   = sum(cf);
    const totInv  = -(evalData.inversion || 0);
    const totFin  = sum(cfF);
    const totLibre = totOp;

    const cfKpiGrid = document.getElementById('cfKpiGrid');
    cfKpiGrid.innerHTML = '';
    [
        { label:'Flujo Operativo Total', value:fmt(totOp), sub: totOp>0?'✓ Positivo':'⚠ Negativo', cls:totOp>0?'success':'danger' },
        { label:'Inversión Inicial', value:fmt(-totInv), sub:'Inversión CAPEX', cls:'warning' },
        { label:'Flujo Financiero Total', value:fmt(totFin), sub: totFin>0?'✓ Positivo':'⚠ Negativo', cls:totFin>0?'success':'danger' },
        { label:'Flujo de Caja Libre', value:fmt(totLibre), sub: totLibre>0?'✓ Genera valor':'⚠ Consume valor', cls:totLibre>0?'success':'danger' }
    ].forEach(k => {
        const el = document.createElement('div');
        el.className = `metric-card ${k.cls}`;
        el.innerHTML = `<div class="metric-label">${k.label}</div><div class="metric-value">${k.value}</div><div class="metric-sub ${k.sub.includes('✓')?'positive':k.sub.includes('⚠')?'negative':''}">${k.sub}</div>`;
        cfKpiGrid.appendChild(el);
    });

    // Chart comparativo
    destroyChart('chartCFComp');
    const n = Math.min(24, cf.length);
    const lbls = Array.from({length:n},(_,i)=>`M${i+1}`);
    charts['chartCFComp'] = new Chart(document.getElementById('chartCFComp'), {
        type:'line',
        data:{ 
            labels: lbls, 
            datasets:[
                { 
                    label:'Flujo Económico', 
                    data: cf.slice(0,n), 
                    borderColor: AmberesColors.primary,
                    backgroundColor: 'transparent',
                    borderWidth: 3,
                    fill: false,
                    tension: 0.1,
                    pointRadius: 0,
                    pointHoverRadius: 6,
                    pointBackgroundColor: AmberesColors.primary,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                },
                { 
                    label:'Flujo Financiero', 
                    data: cfF.slice(0,n), 
                    borderColor: AmberesColors.secondary,
                    backgroundColor: 'transparent',
                    borderWidth: 3,
                    fill: false,
                    tension: 0.1,
                    pointRadius: 0,
                    pointHoverRadius: 6,
                    pointBackgroundColor: AmberesColors.secondary,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }
            ]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins:{ 
                legend: { 
                    labels: { 
                        color: '#b9c7c9', 
                        font: { size: 11, weight: '600' },
                        padding: 15,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    }, 
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });

    // Waterfall
    destroyChart('chartWaterfall');
    const finalValue = totInv + totOp;
    charts['chartWaterfall'] = new Chart(document.getElementById('chartWaterfall'), {
        type:'bar',
        data:{ 
            labels: ['Inversión Inicial','Flujo Operativo','Resultado Final'],
            datasets:[{ 
                label: 'Fondos', 
                data: [totInv, totOp, finalValue],
                backgroundColor: [
                    AmberesColors.danger,
                    AmberesColors.success,
                    finalValue >= 0 ? AmberesColors.primary : AmberesColors.danger
                ],
                borderColor: [
                    AmberesColors.danger,
                    AmberesColors.success,
                    finalValue >= 0 ? AmberesColors.primary : AmberesColors.danger
                ],
                borderWidth: 0,
                borderRadius: 3,
                barPercentage: 0.6
            }]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            plugins:{ 
                legend: { 
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    },
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });

    // Insights CF
    let html = '';
    html += totOp>0?`<div class="insight-item success"><strong>✓ Generación de Caja Positiva</strong><br>Flujo operativo acumulado ${fmt(totOp)} — la operación genera liquidez sostenible.</div>`
        :`<div class="insight-item danger"><strong>⚠ Consumo de Caja</strong><br>Flujo operativo negativo. Revisar estructura de capital y estrategia de cobranzas.</div>`;
    html += totLibre>0?`<div class="insight-item success"><strong>✓ FCL Positivo (${fmt(totLibre)})</strong><br>Permite reinversión, pago de dividendos o amortización de deuda.</div>`
        :`<div class="insight-item warning"><strong>⚠ FCL Negativo</strong><br>Limita capacidad de expansión. Priorizar optimización de capital de trabajo.</div>`;
    html += `<div class="insight-item"><strong>📊 Flujo Promedio Mensual</strong><br>Promedio: ${fmt(totOp/cf.length)} / mes. Monitorear estacionalidad y concentración de cobros.</div>`;
    document.getElementById('insightsCF').innerHTML = html;
}

// ============================================================
// RENDER — P&L
// ============================================================
function renderPL() {
    if (!plData) return;
    document.getElementById('plEmpty').style.display = 'none';
    document.getElementById('plContent').style.display = 'block';

    const tipo = currentEmpresa;
    const real     = plData[tipo]?.real;
    const budget   = plData[tipo]?.budget;
    const forecast = plData[tipo]?.forecast;

    // Si no hay datos en la empresa seleccionada, cambiar al otro tipo automáticamente
    if (!real && !budget && !forecast) {
        const otro = tipo === 'servicios' ? 'productiva' : 'servicios';
        if (plData[otro]?.real || plData[otro]?.budget || plData[otro]?.forecast) {
            switchEmpresa(otro);
            return;
        }
    }

    const getTotal = (esc, key) => {
        if (!esc || !esc[key]) return 0;
        return esc[key][12] || sum(esc[key].slice(0,12)); // columna 13 = TOTAL AÑO (índice 12)
    };

    const rIng = getTotal(real,'ingresosNetos');
    const bIng = getTotal(budget,'ingresosNetos');
    const fIng = getTotal(forecast,'ingresosNetos');
    const rCosto = getTotal(real,'costoTotal');
    const bCosto = getTotal(budget,'costoTotal');
    const fCosto = getTotal(forecast,'costoTotal');
    const rUB = getTotal(real,'utilidadBruta');
    const bUB = getTotal(budget,'utilidadBruta');
    const fUB = getTotal(forecast,'utilidadBruta');
    const rEBIT = getTotal(real,'ebit');
    const bEBIT = getTotal(budget,'ebit');
    const fEBIT = getTotal(forecast,'ebit');
    const rUN = getTotal(real,'utilidadNeta');
    const bUN = getTotal(budget,'utilidadNeta');
    const fUN = getTotal(forecast,'utilidadNeta');

    // KPIs
    const plKpiGrid = document.getElementById('plKpiGrid');
    plKpiGrid.innerHTML = '';
    const varIng  = rIng  && bIng  ? (rIng - bIng)   : 0;
    const varUN   = rUN   && bUN   ? (rUN  - bUN)     : 0;
    const mgB_r   = rIng  ? rUB/rIng   : 0;
    const mgN_r   = rIng  ? rUN/rIng   : 0;
    [
        { label:'Ingresos Netos — Real', value:fmt(rIng), sub:`vs Budget: ${varIng>=0?'+':''}${fmt(varIng)}`, cls: varIng>=0?'success':'warning', sem: varIng>=0?'green':'yellow' },
        { label:'Utilidad Bruta — Real', value:fmt(rUB), sub:`Margen: ${pct(mgB_r)}`, cls:'success', sem:'green' },
        { label:'EBIT — Real', value:fmt(rEBIT), sub:`Margen Op.: ${pct(rIng?rEBIT/rIng:0)}`, cls:rEBIT>0?'success':'danger', sem:rEBIT>0?'green':'red' },
        { label:'Utilidad Neta — Real', value:fmt(rUN), sub:`vs Budget: ${varUN>=0?'+':''}${fmt(varUN)}`, cls: varUN>=0?'success':'danger', sem: varUN>=0?'green':'red' },
        { label:'Margen Bruto', value:pct(mgB_r), sub: mgB_r>=0.4?'✓ Saludable':mgB_r>=0.2?'⚠ Moderado':'⚠ Bajo', cls: mgB_r>=0.4?'success':mgB_r>=0.2?'warning':'danger', sem: mgB_r>=0.4?'green':mgB_r>=0.2?'yellow':'red' },
        { label:'Margen Neto', value:pct(mgN_r), sub: mgN_r>=0.1?'✓ Saludable':mgN_r>=0.05?'⚠ Moderado':'⚠ Bajo', cls: mgN_r>=0.1?'success':mgN_r>=0.05?'warning':'danger', sem: mgN_r>=0.1?'green':mgN_r>=0.05?'yellow':'red' }
    ].forEach(k => {
        const el = document.createElement('div');
        el.className = `metric-card ${k.cls}`;
        el.innerHTML = `<div class="metric-label"><span class="semaforo ${k.sem}"></span>${k.label}</div>
                        <div class="metric-value">${k.value}</div>
                        <div class="metric-sub ${k.sub.includes('✓')?'positive':k.sub.includes('⚠')||k.sub.includes('-')?'negative':''}">${k.sub}</div>`;
        plKpiGrid.appendChild(el);
    });

    // Tabla comparativa
    buildCompareTable(real, budget, forecast, rIng, bIng, fIng);

    // Chart vertical (doughnut)
    destroyChart('chartVertical');
    if (rIng > 0) {
        const slices = [
            { label:'Costo de Servicios/Ventas', val: rCosto },
            { label:'Gastos Operacionales', val: Math.max(0, rUB - rEBIT) },
            { label:'Impuesto', val: Math.max(0, rEBIT - rUN) },
            { label:'Utilidad Neta', val: Math.max(0, rUN) }
        ].filter(s => s.val > 0);
        
        const colors = [
            AmberesColors.danger,
            AmberesColors.accent1,
            AmberesColors.accent3,
            AmberesColors.success
        ];
        
        charts['chartVertical'] = new Chart(document.getElementById('chartVertical'), {
            type:'doughnut',
            data:{ 
                labels: slices.map(s=>s.label), 
                datasets:[{ 
                    data: slices.map(s=>s.val),
                    backgroundColor: colors.slice(0, slices.length),
                    borderColor: '#fff',
                    borderWidth: 3,
                    hoverBorderWidth: 4
                }]
            },
            options:{ 
                responsive: true, 
                maintainAspectRatio: true,
                cutout: '60%',
                plugins:{ 
                    legend: { 
                        position: 'right',
                        labels: { 
                            color: '#b9c7c9', 
                            font: { size: 11, weight: '600' },
                            padding: 12,
                            usePointStyle: true,
                            pointStyle: 'circle'
                        } 
                    },
                    tooltip: { 
                        callbacks: { 
                            label: ctx => {
                                const value = ctx.parsed;
                                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = ((value / total) * 100).toFixed(1);
                                return `${ctx.label}: ${fmt(value)} (${percentage}%)`;
                            }
                        } 
                    }
                } 
            }
        });
    }

    // Chart horizontal (líneas mensuales)
    destroyChart('chartHorizontal');
    const months = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    const getMonths = (esc, key) => esc && esc[key] ? esc[key].slice(0,12) : Array(12).fill(0);
    charts['chartHorizontal'] = new Chart(document.getElementById('chartHorizontal'), {
        type:'line',
        data:{ 
            labels: months, 
            datasets:[
                { 
                    label:'Ingresos Real',
                    data: getMonths(real,'ingresosNetos'),
                    borderColor: AmberesColors.primary,
                    backgroundColor: 'transparent',
                    borderWidth: 3,
                    tension: 0.1,
                    fill: false,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    pointBackgroundColor: AmberesColors.primary,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                },
                { 
                    label:'Ingresos Budget',
                    data: getMonths(budget,'ingresosNetos'),
                    borderColor: AmberesColors.secondary,
                    backgroundColor: 'transparent',
                    borderWidth: 2.5,
                    borderDash: [8, 4],
                    tension: 0.1,
                    fill: false,
                    pointRadius: 0,
                    pointHoverRadius: 6,
                    pointBackgroundColor: AmberesColors.secondary,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                },
                { 
                    label:'Ingresos Forecast',
                    data: getMonths(forecast,'ingresosNetos'),
                    borderColor: AmberesColors.accent1,
                    backgroundColor: 'transparent',
                    borderWidth: 2.5,
                    borderDash: [4, 4],
                    tension: 0.1,
                    fill: false,
                    pointRadius: 0,
                    pointHoverRadius: 6,
                    pointBackgroundColor: AmberesColors.accent1,
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }
            ]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins:{ 
                legend: { 
                    labels: { 
                        color: '#b9c7c9', 
                        font: { size: 11, weight: '600' },
                        padding: 15,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    },
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });

    // Chart escenarios (barras agrupadas)
    destroyChart('chartScenarios');
    charts['chartScenarios'] = new Chart(document.getElementById('chartScenarios'), {
        type:'bar',
        data:{ 
            labels: ['Ingresos Netos','Costo Total','Utilidad Bruta','EBIT','Utilidad Neta'],
            datasets:[
                { 
                    label:'Real',
                    data: [rIng, rCosto, rUB, rEBIT, rUN],
                    backgroundColor: AmberesColors.primary,
                    borderColor: AmberesColors.primary,
                    borderWidth: 0,
                    borderRadius: 3,
                    barPercentage: 0.7
                },
                { 
                    label:'Budget',
                    data: [bIng, bCosto, bUB, bEBIT, bUN],
                    backgroundColor: AmberesColors.secondary,
                    borderColor: AmberesColors.secondary,
                    borderWidth: 0,
                    borderRadius: 3,
                    barPercentage: 0.7
                },
                { 
                    label:'Forecast',
                    data: [fIng, fCosto, fUB, fEBIT, fUN],
                    backgroundColor: AmberesColors.accent1,
                    borderColor: AmberesColors.accent1,
                    borderWidth: 0,
                    borderRadius: 3,
                    barPercentage: 0.7
                }
            ]
        },
        options:{ 
            responsive: true, 
            maintainAspectRatio: true,
            plugins:{ 
                legend: { 
                    labels: { 
                        color: '#b9c7c9', 
                        font: { size: 11, weight: '600' },
                        padding: 15,
                        usePointStyle: true
                    } 
                },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}`
                    }
                }
            },
            scales:{ 
                y: { 
                    ticks: { 
                        color: '#b9c7c9',
                        callback: v => fmt(v),
                        font: { size: 10 }
                    }, 
                    grid: { 
                        color: 'rgba(255,255,255,0.07)',
                        drawBorder: false
                    },
                    border: {
                        display: false
                    }
                },
                x: { 
                    ticks: { 
                        color: '#b9c7c9',
                        font: { size: 10 }
                    },
                    grid: { 
                        display: false
                    },
                    border: {
                        display: false
                    }
                } 
            } 
        }
    });

    // Insights
    renderPLInsights(rIng, bIng, rUN, bUN, mgB_r, mgN_r, varIng);
}

function buildCompareTable(real, budget, forecast, rIng, bIng, fIng) {
    const tbl = document.getElementById('compareTable');
    const getT = (esc, key) => {
        if (!esc || !esc[key]) return 0;
        return esc[key][12] !== undefined ? esc[key][12] : sum((esc[key]||[]).slice(0,12));
    };
    const varFmt = (r, b) => {
        if (!b) return { val:'-', cls:'' };
        const d = r - b;
        const p = b ? (d/Math.abs(b)*100).toFixed(1) : 0;
        return { val:`${d>=0?'+':''}${fmt(d)} (${d>=0?'+':''}${p}%)`, cls: d>=0?'var-positive':'var-negative' };
    };

    const rows = [
        { label:'INGRESOS', isSection:true },
        { label:'Ingresos Netos', key:'ingresosNetos' },
        { label:'COSTOS', isSection:true },
        { label:'Costo de Ventas/Servicios', key:'costoTotal' },
        { label:'UTILIDAD BRUTA', key:'utilidadBruta', isTotal:true },
        { label:'Margen Bruto %', key:'margenBruto', isMgn:true },
        { label:'GASTOS OPERACIONALES', key:'gastosOp', isSection:false },
        { label:'EBIT', key:'ebit', isTotal:true },
        { label:'Margen Operacional %', key:'margenOp', isMgn:true },
        { label:'UTILIDAD ANTES IMPUESTO', key:'ebt', isTotal:true },
        { label:'UTILIDAD NETA', key:'utilidadNeta', isTotal:true },
        { label:'Margen Neto %', key:'margenNeto', isMgn:true }
    ];

    let html = `<thead><tr>
        <th>Concepto</th>
        <th>Real</th><th>Budget</th><th>Forecast</th>
        <th>Var. Real vs Budget</th><th>Var. Real vs Forecast</th>
    </tr></thead><tbody>`;

    rows.forEach(row => {
        if (row.isSection) {
            html += `<tr class="section-row"><td colspan="6">${row.label}</td></tr>`;
            return;
        }
        const r = getT(real, row.key);
        const b = getT(budget, row.key);
        const f = getT(forecast, row.key);
        const vb = varFmt(r,b), vf = varFmt(r,f);
        const trCls = row.isTotal ? 'total-row' : '';
        const fmtVal = row.isMgn
            ? (v => v > 1 ? pctN(v) : pct(v))  // si es >1 ya está en %, si no multiplicar
            : fmt;
        html += `<tr class="${trCls}">
            <td>${row.label}</td>
            <td>${fmtVal(r)}</td><td>${fmtVal(b)}</td><td>${fmtVal(f)}</td>
            <td class="${vb.cls}">${vb.val}</td>
            <td class="${vf.cls}">${vf.val}</td>
        </tr>`;
    });
    html += '</tbody>';
    tbl.innerHTML = html;
}

function renderPLInsights(rIng, bIng, rUN, bUN, mgB, mgN, varIng) {
    let html = '';
    if (rIng > 0) {
        html += varIng >= 0
            ? `<div class="insight-item success"><strong>✓ Ingresos Superan el Budget en ${fmt(varIng)}</strong><br>La empresa genera más ingresos que lo presupuestado. Analizar si esto es sostenible o puntual.</div>`
            : `<div class="insight-item warning"><strong>⚠ Ingresos bajo Budget en ${fmt(Math.abs(varIng))}</strong><br>Brecha relevante. Revisar pipeline comercial y estrategia de precios.</div>`;
        html += mgB >= 0.4
            ? `<div class="insight-item success"><strong>✓ Margen Bruto Sólido (${pct(mgB)})</strong><br>Estructura de costos competitiva. Capacidad de absorber gastos operativos y generar EBIT positivo.</div>`
            : mgB >= 0.2
            ? `<div class="insight-item warning"><strong>⚠ Margen Bruto Moderado (${pct(mgB)})</strong><br>Hay espacio de mejora. Revisar costos directos y política de descuentos.</div>`
            : `<div class="insight-item danger"><strong>⚠ Margen Bruto Bajo (${pct(mgB)})</strong><br>Urgente: optimizar costo de servicios/ventas o ajustar precios.</div>`;
        html += mgN >= 0.1
            ? `<div class="insight-item success"><strong>✓ Rentabilidad Neta Saludable (${pct(mgN)})</strong><br>La empresa retiene ${pct(mgN)} de cada dólar de ingreso como utilidad neta.</div>`
            : mgN >= 0
            ? `<div class="insight-item warning"><strong>⚠ Margen Neto Bajo (${pct(mgN)})</strong><br>Baja rentabilidad final. Revisar gastos operativos, financieros e impositivos.</div>`
            : `<div class="insight-item danger"><strong>⚠ Pérdida Neta (${pct(mgN)})</strong><br>La empresa está operando con pérdidas. Acción inmediata requerida.</div>`;
        const varUN = rUN && bUN ? rUN - bUN : 0;
        html += varUN >= 0
            ? `<div class="insight-item success"><strong>✓ Utilidad Neta Supera Budget en ${fmt(varUN)}</strong><br>Buen desempeño general. Capitalizar estas ventajas en el Forecast.</div>`
            : `<div class="insight-item danger"><strong>⚠ Utilidad Neta bajo Budget en ${fmt(Math.abs(varUN))}</strong><br>Revisar drivers de mayor costo o menor ingreso vs plan.</div>`;
    } else {
        html = `<div class="insight-item warning"><strong>Sin datos disponibles para el tipo de empresa seleccionado</strong><br>Las pestañas de Empresa Productiva pueden no tener datos cargados. Intenta con Empresa de Servicios.</div>`;
    }
    document.getElementById('insightsPL').innerHTML = html;
}

// ============================================================
// EXPORTAR A PDF
// ============================================================
function exportToPDF() {
    document.getElementById('processingOverlay').classList.add('active');
    
    const container = document.createElement('div');
    container.style.cssText = 'position:absolute;left:-9999px;width:210mm;background:#fff;padding:20mm;font-family:Arial,sans-serif;color:#1A1A1A';
    
    const activeTab = document.querySelector('.dashboard-content.active');
    const tabName = activeTab.id;
    
    let tabTitle = tabName === 'overview' ? 'Visión General del Negocio' : tabName === 'pl-analysis' ? 'Análisis P&L' : 'Flujo de Caja';
    
    let html = `
        <div style="margin-bottom:30px;border-bottom:3px solid #c89b3c;padding-bottom:15px">
            <h1 style="margin:0;font-size:24px;color:#1A1A1A">Amberes Consultora</h1>
            <h2 style="margin:8px 0 0 0;font-size:18px;color:#5f747c;font-weight:400">Reporte de Análisis Financiero</h2>
            <p style="margin:5px 0 0 0;font-size:12px;color:#7f959d">${new Date().toLocaleDateString('es-CL',{year:'numeric',month:'long',day:'numeric'})}</p>
        </div>
        <h2 style="font-size:20px;color:#c89b3c;margin:25px 0 15px 0;border-left:4px solid #c89b3c;padding-left:10px">${tabTitle}</h2>
    `;
    
    const kpiGrid = activeTab.querySelector('.metrics-grid');
    if (kpiGrid) {
        html += '<div style="margin:20px 0">';
        kpiGrid.querySelectorAll('.metric-card').forEach(card => {
            const label = (card.querySelector('.metric-label')?.textContent || '').replace(/[🔴🟡🟢●]/g,'').trim();
            const value = card.querySelector('.metric-value')?.textContent || '';
            const sub = card.querySelector('.metric-sub')?.textContent || '';
            html += `<div style="display:inline-block;width:48%;margin:1%;padding:15px;background:#f4efe5;border-left:4px solid #c89b3c;vertical-align:top">
                <div style="font-size:11px;color:#5f747c;text-transform:uppercase;margin-bottom:5px">${label}</div>
                <div style="font-size:24px;font-weight:700;color:#1A1A1A;margin-bottom:3px">${value}</div>
                <div style="font-size:12px;color:#7f959d">${sub}</div>
            </div>`;
        });
        html += '</div>';
    }
    
    const insights = activeTab.querySelector('.insights-section');
    if (insights) {
        html += '<h3 style="font-size:16px;color:#1A1A1A;margin:25px 0 12px 0">Insights y Recomendaciones</h3>';
        insights.querySelectorAll('.insight-item').forEach(item => {
            const text = item.innerHTML.replace(/<strong>/g,'<b>').replace(/<\/strong>/g,'</b>').replace(/<br>/g,' ');
            html += `<div style="margin:8px 0;padding:10px;background:#f4efe5;border-left:3px solid #c89b3c;font-size:13px;line-height:1.5">${text}</div>`;
        });
    }
    
    if (tabName === 'pl-analysis') {
        const table = document.getElementById('compareTable');
        if (table) {
            html += '<h3 style="font-size:16px;color:#1A1A1A;margin:25px 0 12px 0;page-break-before:always">Comparativa Real vs Budget vs Forecast</h3>';
            html += '<div style="font-size:11px">' + table.outerHTML + '</div>';
        }
    }
    
    html += `<div style="margin-top:40px;padding-top:15px;border-top:1px solid #A7A8AA;font-size:10px;color:#7f959d;text-align:center">
        <p>Amberes Consultora · Financial Intelligence</p>
        <p>Generado automáticamente por Amberes Financial Intelligence</p>
    </div>`;
    
    container.innerHTML = html;
    document.body.appendChild(container);
    
    const opt = {
        margin: 10,
        filename: `Amberes_${tabTitle.replace(/ /g,'_')}_${new Date().toISOString().slice(0,10)}.pdf`,
        image: {type:'jpeg',quality:0.98},
        html2canvas: {scale:2,useCORS:true,letterRendering:true},
        jsPDF: {unit:'mm',format:'a4',orientation:'portrait'}
    };
    
    html2pdf().set(opt).from(container).save().then(() => {
        document.body.removeChild(container);
        document.getElementById('processingOverlay').classList.remove('active');
    }).catch(err => {
        console.error('Error PDF:', err);
        document.body.removeChild(container);
        document.getElementById('processingOverlay').classList.remove('active');
        alert('Error al generar PDF. Intenta nuevamente.');
    });
}

// ============================================================
// MÓDULO M&A / MODELAMIENTO FINANCIERO
// ============================================================
function switchMATab(id, btn) {
    document.querySelectorAll('.ma-subcontent').forEach(c=>c.classList.remove('active'));
    document.querySelectorAll('.ma-subtab-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById('ma-sub-'+id).classList.add('active');
    btn.classList.add('active');
    renderMACharts(window.__maState);
}

function getMAInputs() {
    const g = id => document.getElementById(id);
    return {
        ingBase: +g('maIngBase').value,
        growth: +g('maGrowth').value/100,
        margin: +g('maMargin').value/100,
        da: +g('maDA').value/100,
        capex: +g('maCapex').value/100,
        nwc: +g('maNWC').value/100,
        tax: +g('maTax').value/100,
        years: +g('maYears').value,
        termG: +g('maTermG').value/100,
        rf: +g('maRf').value/100,
        beta: +g('maBeta').value,
        erp: +g('maErp').value/100,
        country: +g('maCountry').value/100,
        kd: +g('maKd').value/100,
        de: +g('maDE').value/100,
        mult: +g('maMult').value,
        netDebt: +g('maNetDebt').value,
        synCost: +g('maSynCost').value,
        synRev: +g('maSynRev').value,
        synYears: +g('maSynYears').value,
        synOneoff: +g('maSynOneoff').value
    };
}

function buildProjection(p) {
    const rows = [];
    let prevRev = p.ingBase;
    for (let y=1; y<=p.years; y++) {
        const rev = p.ingBase * Math.pow(1+p.growth, y);
        const ebitda = rev * p.margin;
        const da = rev * p.da;
        const ebit = ebitda - da;
        const taxAmt = Math.max(ebit,0) * p.tax;
        const nopat = ebit - taxAmt;
        const capex = rev * p.capex;
        const deltaRev = rev - prevRev;
        const nwc = deltaRev * p.nwc;
        const fcf = nopat + da - capex - nwc;
        rows.push({ y, rev, ebitda, da, ebit, taxAmt, nopat, capex, nwc, fcf });
        prevRev = rev;
    }
    return rows;
}

function calcWACC(p) {
    const ke = p.rf + p.beta*p.erp + p.country;
    const kdAT = p.kd * (1-p.tax);
    const we = 1 - p.de;
    const wacc = we*ke + p.de*kdAT;
    return { wacc, ke, kdAT, we, wd:p.de };
}

function dcfEV(rows, wacc, termG) {
    let pvSum = 0;
    rows.forEach(r => { pvSum += r.fcf / Math.pow(1+wacc, r.y); });
    const lastFCF = rows[rows.length-1].fcf;
    const safeTermG = Math.min(termG, wacc - 0.005);
    const tv = (lastFCF*(1+safeTermG)) / (wacc - safeTermG);
    const pvTV = tv / Math.pow(1+wacc, rows.length);
    return { pvSum, tv, pvTV, ev: pvSum + pvTV };
}

function runMA() {
    const p = getMAInputs();
    const g = id => document.getElementById(id);

    // Mostrar valores de sliders
    g('dIngBase').textContent = fmt(p.ingBase);
    g('dGrowth').textContent = (p.growth*100).toFixed(1)+'%';
    g('dMargin').textContent = (p.margin*100).toFixed(1)+'%';
    g('dDA').textContent = (p.da*100).toFixed(1)+'%';
    g('dCapex').textContent = (p.capex*100).toFixed(1)+'%';
    g('dNWC').textContent = (p.nwc*100).toFixed(1)+'%';
    g('dTaxMA').textContent = (p.tax*100).toFixed(1)+'%';
    g('dYears').textContent = p.years;
    g('dTermG').textContent = (p.termG*100).toFixed(1)+'%';
    g('dRf').textContent = (p.rf*100).toFixed(1)+'%';
    g('dBeta').textContent = p.beta.toFixed(2);
    g('dErp').textContent = (p.erp*100).toFixed(1)+'%';
    g('dCountry').textContent = (p.country*100).toFixed(1)+'%';
    g('dKd').textContent = (p.kd*100).toFixed(1)+'%';
    g('dDE').textContent = (p.de*100).toFixed(0)+'%';
    g('dMult').textContent = p.mult.toFixed(1)+'x';
    g('dNetDebt').textContent = fmt(p.netDebt);
    g('dSynCost').textContent = fmt(p.synCost);
    g('dSynRev').textContent = fmt(p.synRev);
    g('dSynYears').textContent = p.synYears;
    g('dSynCost1').textContent = fmt(p.synOneoff);

    // Proyección y WACC
    const rows = buildProjection(p);
    const waccData = calcWACC(p);
    const dcf = dcfEV(rows, waccData.wacc, p.termG);
    const equity = dcf.ev - p.netDebt;

    // Múltiplos
    const evMult = rows[0].ebitda * p.mult;

    // Sinergias
    let pvSyn = 0;
    const synRowsData = [];
    for (let y=1; y<=p.years; y++) {
        const frac = Math.min(1, y/p.synYears);
        const pretax = (p.synCost + p.synRev) * frac;
        const afterTax = pretax * (1-p.tax);
        pvSyn += afterTax / Math.pow(1+waccData.wacc, y);
        synRowsData.push({ y, cost: p.synCost*frac, rev: p.synRev*frac, afterTax });
    }
    const lastSyn = synRowsData[synRowsData.length-1]?.afterTax || 0;
    const safeTermGSyn = Math.min(p.termG, waccData.wacc - 0.005);
    const tvSyn = (lastSyn*(1+safeTermGSyn)) / (waccData.wacc - safeTermGSyn);
    const pvTVSyn = tvSyn / Math.pow(1+waccData.wacc, p.years);
    const npvSynergies = pvSyn + pvTVSyn - p.synOneoff;
    const evCombined = dcf.ev + npvSynergies;

    // KPIs
    g('maEV').textContent = fmt(dcf.ev);
    g('maEVsub').textContent = `VP FCF: ${fmt(dcf.pvSum)} · VP Terminal: ${fmt(dcf.pvTV)}`;
    g('maEquity').textContent = fmt(equity);
    g('maWACC').textContent = (waccData.wacc*100).toFixed(2)+'%';
    g('maWACCsub').textContent = `Ke: ${(waccData.ke*100).toFixed(1)}% · Kd(AT): ${(waccData.kdAT*100).toFixed(1)}% · E/D: ${(waccData.we*100).toFixed(0)}/${(waccData.wd*100).toFixed(0)}`;
    g('maEVMult').textContent = fmt(evMult);
    g('maEVMultSub').textContent = `EBITDA Año 1 (${fmt(rows[0].ebitda)}) × ${p.mult.toFixed(1)}x`;
    g('maSynNPV').textContent = fmt(npvSynergies);
    g('maEVCombined').textContent = fmt(evCombined);

    // Tabla DCF
    const dcfCols = rows.map(r=>`Año ${r.y}`);
    const labelRow = (label, arr, isTotal) => `<tr${isTotal?' class="ma-total"':''}><td>${label}</td>${arr.map(v=>`<td>${fmt(v)}</td>`).join('')}</tr>`;
    let tableHTML = `<tr><th>Concepto</th>${dcfCols.map(c=>`<th>${c}</th>`).join('')}</tr>`;
    tableHTML += labelRow('Ingresos', rows.map(r=>r.rev));
    tableHTML += labelRow('EBITDA', rows.map(r=>r.ebitda));
    tableHTML += labelRow('D&A', rows.map(r=>-r.da));
    tableHTML += labelRow('EBIT', rows.map(r=>r.ebit));
    tableHTML += labelRow('Impuesto', rows.map(r=>-r.taxAmt));
    tableHTML += labelRow('NOPAT', rows.map(r=>r.nopat));
    tableHTML += labelRow('+ D&A', rows.map(r=>r.da));
    tableHTML += labelRow('− Capex', rows.map(r=>-r.capex));
    tableHTML += labelRow('− Δ Capital de Trabajo', rows.map(r=>-r.nwc));
    tableHTML += labelRow('Flujo de Caja Libre (FCF)', rows.map(r=>r.fcf), true);
    tableHTML += labelRow('Factor de Descuento', rows.map(r=>1/Math.pow(1+waccData.wacc,r.y)));
    tableHTML += labelRow('Valor Presente FCF', rows.map(r=>r.fcf/Math.pow(1+waccData.wacc,r.y)), true);
    g('maDcfTable').innerHTML = tableHTML;

    window.__maState = { p, rows, waccData, dcf, evMult, synRowsData, npvSynergies, evCombined };
    renderMACharts(window.__maState);
}

function renderMACharts(state) {
    if (!state) return;
    const { p, rows, waccData, dcf, evMult } = state;

    // FOOTBALL FIELD
    destroyChart('chartFootball');
    const waccLow = Math.max(0.01, waccData.wacc - 0.005);
    const waccHigh = waccData.wacc + 0.005;
    const dcfHigh = dcfEV(rows, waccLow, p.termG).ev;
    const dcfLow = dcfEV(rows, waccHigh, p.termG).ev;
    const multLow = rows[0].ebitda * Math.max(1, p.mult-1);
    const multHigh = rows[0].ebitda * (p.mult+1);
    const combLow = state.evCombined*0.96, combHigh = state.evCombined*1.04;

    const ffCtx = document.getElementById('chartFootball');
    if (ffCtx) {
        charts['chartFootball'] = new Chart(ffCtx, {
            type: 'bar',
            data: {
                labels: ['DCF (WACC ±50pb)', 'Múltiplos Comparables (±1.0x)', 'DCF + Sinergias'],
                datasets: [{
                    label: 'Rango de Enterprise Value',
                    data: [[dcfLow,dcfHigh],[multLow,multHigh],[combLow,combHigh]],
                    backgroundColor: [AmberesColors.primary, AmberesColors.accent3, AmberesColors.secondary],
                    borderRadius: 4,
                    barPercentage: 0.5
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display:false },
                    tooltip: { callbacks: { label: ctx => `${fmt(ctx.raw[0])} — ${fmt(ctx.raw[1])}` } }
                },
                scales: {
                    x: { ticks:{ color:'#b9c7c9', callback:v=>fmt(v), font:{size:10} }, grid:{ color:'rgba(255,255,255,0.07)' }, border:{display:false} },
                    y: { ticks:{ color:'#b9c7c9', font:{size:11,weight:'600'} }, grid:{ display:false }, border:{display:false} }
                }
            }
        });
    }

    // PROYECCIÓN
    destroyChart('chartMAProjection');
    const pCtx = document.getElementById('chartMAProjection');
    if (pCtx) {
        charts['chartMAProjection'] = new Chart(pCtx, {
            type: 'bar',
            data: {
                labels: rows.map(r=>`Año ${r.y}`),
                datasets: [
                    { type:'line', label:'Ingresos', data: rows.map(r=>r.rev), borderColor: AmberesColors.neutral2, backgroundColor:'transparent', borderWidth:2, borderDash:[5,3], pointRadius:3, tension:.15, yAxisID:'y' },
                    { label:'EBITDA', data: rows.map(r=>r.ebitda), backgroundColor: AmberesColors.primary, borderRadius:3, barPercentage:.6, yAxisID:'y' },
                    { label:'FCF', data: rows.map(r=>r.fcf), backgroundColor: AmberesColors.secondary, borderRadius:3, barPercentage:.6, yAxisID:'y' }
                ]
            },
            options: {
                responsive:true, maintainAspectRatio:true,
                interaction:{ mode:'index', intersect:false },
                plugins:{ legend:{ labels:{ color:'#b9c7c9', font:{size:11,weight:'600'}, usePointStyle:true, padding:14 } },
                    tooltip:{ callbacks:{ label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } } },
                scales:{
                    y:{ ticks:{ color:'#b9c7c9', callback:v=>fmt(v), font:{size:10} }, grid:{ color:'rgba(255,255,255,0.07)' }, border:{display:false} },
                    x:{ ticks:{ color:'#b9c7c9', font:{size:10} }, grid:{ display:false }, border:{display:false} }
                }
            }
        });
    }

    // SINERGIAS
    destroyChart('chartSynergies');
    const sCtx = document.getElementById('chartSynergies');
    if (sCtx && state.synRowsData) {
        charts['chartSynergies'] = new Chart(sCtx, {
            type: 'bar',
            data: {
                labels: state.synRowsData.map(r=>`Año ${r.y}`),
                datasets: [
                    { label:'Sinergias de Costos', data: state.synRowsData.map(r=>r.cost), backgroundColor: AmberesColors.secondary, borderRadius:3, barPercentage:.65 },
                    { label:'Sinergias de Ingresos', data: state.synRowsData.map(r=>r.rev), backgroundColor: AmberesColors.primary, borderRadius:3, barPercentage:.65 }
                ]
            },
            options: {
                responsive:true, maintainAspectRatio:true,
                plugins:{ legend:{ labels:{ color:'#b9c7c9', font:{size:11,weight:'600'}, usePointStyle:true, padding:14 } },
                    tooltip:{ callbacks:{ label: ctx => `${ctx.dataset.label}: ${fmt(ctx.parsed.y)}` } } },
                scales:{
                    y:{ stacked:true, ticks:{ color:'#b9c7c9', callback:v=>fmt(v), font:{size:10} }, grid:{ color:'rgba(255,255,255,0.07)' }, border:{display:false} },
                    x:{ stacked:true, ticks:{ color:'#b9c7c9', font:{size:10} }, grid:{ display:false }, border:{display:false} }
                }
            }
        });
    }
}

// Inicializar módulo M&A con valores por defecto
runMA();
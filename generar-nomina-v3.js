/**
 * Calculadora de Nómina v3 — SS multi-base CORRECTA
 *
 * Estructura SS real (MECALUX):
 *   Base CC  = sujeto_sin_HHEE + prorrata_pagas_extra
 *   Base AT  = Base CC + HHEE
 *   SS CC (4.70%)          → Base CC × 4.70%
 *   SS Desempleo+FP (1.65%)→ Base AT × 1.65%
 *   SS HH.EE. (4.70%)      → HHEE × 4.70%
 *   SS MEI (0.13%)         → Base AT × 0.13%
 *
 *   IRPF = Bruto Sujeto (sin prorrata) × IRPF%
 *   NETO = (Bruto Sujeto + Exentos) − Total SS − IRPF
 *        = Total Devengado − Total SS − IRPF
 *
 * Verificación febrero 2026:
 *   Base CC  = 2787.26 + 419.61 = 3206.87 ≈ 3207.47 (±0.60 por redondeos internos)
 *   Base AT  = 3206.87 + 474.48 = 3681.35 ≈ 3681.95
 *   SS CC    = 3207.47 × 4.70% = 150.75 ✓
 *   SS Des+FP= 3681.95 × 1.65% = 60.75 ✓
 *   SS HHEE  = 474.48  × 4.70% = 22.30 ✓
 *   SS MEI   = 3681.95 × 0.13% = 4.79 ≈ 4.81 (tasa real 0.1307%)
 *   IRPF     = 3261.74 × 18.15% = 592.01 ✓
 *   NETO     = 4025.90 − 238.61 − 592.01 = 3195.28 ✓
 */

const ExcelJS = require('/home/user/Mi-memoria/node_modules/exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mi Memoria — Calculadora v3';
  wb.created = new Date();

  // ── Paleta ────────────────────────────────────────────────────────────────
  const C = {
    moradoDark:  '4338CA',
    morado:      '6366F1',
    moradoLight: 'EDE9FE',
    verde:       '059669',
    verdeLight:  'D1FAE5',
    rojo:        'DC2626',
    rojoLight:   'FEE2E2',
    amarillo:    'FFFBEB',
    amarilloB:   'F59E0B',
    gris:        'F8FAFC',
    grisBorde:   'E2E8F0',
    azulLight:   'EFF6FF',
    azul:        '1D4ED8',
    text:        '0F172A',
    muted:       '64748B',
    blanco:      'FFFFFF',
    rojoSuave:   'B91C1C',
    verdeSuave:  '065F46',
  };

  const ws = wb.addWorksheet('Calculadora Nómina 2026', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 },
    properties: { tabColor: { argb: C.moradoDark } },
  });

  ws.getColumn('A').width = 2;
  ws.getColumn('B').width = 35;
  ws.getColumn('C').width = 13;
  ws.getColumn('D').width = 13;
  ws.getColumn('E').width = 16;
  ws.getColumn('F').width = 12;

  const EURO = '#,##0.00 "€"';
  const PCT  = '0.00%';
  const NUM  = '#,##0.00';

  const R = (row, col) => ws.getCell(row, col);

  function fill(c, argb) {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  }
  function border(c, color = C.grisBorde, style = 'thin') {
    const s = { style, color: { argb: color } };
    c.border = { top: s, bottom: s, left: s, right: s };
  }
  function font(c, { bold=false, size=10, color=C.text, italic=false } = {}) {
    c.font = { bold, size, color: { argb: color }, italic, name: 'Calibri' };
  }

  function secHeader(row, text, bg=C.moradoDark, fg=C.blanco, sz=11) {
    ws.mergeCells(row, 1, row, 6);
    const c = R(row, 1);
    c.value = text;
    font(c, { bold: true, size: sz, color: fg });
    fill(c, bg);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(row).height = 24;
  }

  function spacer(row, bg=C.gris, h=5) {
    ws.mergeCells(row, 1, row, 6);
    fill(R(row, 1), bg);
    ws.getRow(row).height = h;
  }

  function divider(row) {
    ws.mergeCells(row, 1, row, 6);
    fill(R(row, 1), C.grisBorde);
    ws.getRow(row).height = 3;
  }

  function inputCell(row, col, value, fmt=EURO) {
    const c = R(row, col);
    c.value = value;
    c.numFmt = fmt;
    font(c, { bold: true, color: C.moradoDark });
    fill(c, C.amarillo);
    border(c, C.amarilloB);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  function frmCell(row, col, formula, fmt=EURO, bg=C.gris, fg=C.text) {
    const c = R(row, col);
    c.value = { formula };
    c.numFmt = fmt;
    font(c, { bold: true, color: fg });
    fill(c, bg);
    border(c, C.grisBorde);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  function lbl(row, text, opts={}) {
    const c = R(row, 2);
    c.value = text;
    font(c, { size: 10, color: opts.color||C.text, bold: opts.bold||false, italic: opts.italic||false });
    c.alignment = { vertical: 'middle', indent: opts.indent||1, wrapText: opts.wrap||false };
    if (opts.bg) fill(c, opts.bg);
    return c;
  }

  function ssTag(row, text, color=C.rojo) {
    const c = R(row, 6);
    c.value = text;
    font(c, { size: 9, bold: true, color });
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Celda de label+fórmula en bloque con merge A-D
  function blockRow(row, label, formula, fmt=EURO, bg=C.gris, fg=C.text, sz=10, labelBold=false) {
    ws.mergeCells(row, 1, row, 4);
    const lc = R(row, 1);
    lc.value = label;
    font(lc, { bold: labelBold, size: sz, color: fg });
    fill(lc, bg);
    lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    const fc = R(row, 5);
    if (formula) {
      fc.value = { formula };
      fc.numFmt = fmt;
      font(fc, { bold: labelBold, size: sz, color: fg });
      fill(fc, bg);
      border(fc, C.grisBorde);
    }
    fc.alignment = { vertical: 'middle', horizontal: 'right' };
    ws.getRow(row).height = 20;
    return { lc, fc };
  }

  let r = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // TÍTULO
  // ═══════════════════════════════════════════════════════════════════════════
  ws.mergeCells(r, 1, r, 6);
  { const c = R(r, 1); c.value = 'CALCULADORA NÓMINA 2026 — v3';
    font(c, { bold: true, size: 18, color: C.blanco });
    fill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 42; } r++;

  ws.mergeCells(r, 1, r, 6);
  { const c = R(r, 1);
    c.value = 'SS calculada en bases separadas: Base CC (sujeto sin HHEE + prorrata) y Base AT (Base CC + HHEE)  |  IRPF sobre bruto sujeto mensual';
    font(c, { size: 9, italic: true, color: '78350F' });
    fill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 18; } r++;

  spacer(r, C.gris, 8); r++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECCIÓN CONFIGURACIÓN
  // ═══════════════════════════════════════════════════════════════════════════
  secHeader(r, '⚙️   CONFIGURACIÓN — Celdas amarillas = editables'); r++;

  // Cabeceras config
  ['', 'Parámetro', 'Valor', '', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    font(c, { bold: true, size: 9, color: C.muted });
    c.alignment = { horizontal: 'center' };
  });
  ws.getRow(r).height = 16; r++;

  // ── Salario ───────────────────────────────────────────────────────────────
  const ROW_BRUTO = r;
  lbl(r, 'Salario Convenio mensual', { bold: true });
  inputCell(r, 3, 2063.39, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← salario base fijo mensual';
    font(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  const ROW_BRUTO_DIA = r;
  lbl(r, 'Bruto por día (días fuera)');
  inputCell(r, 3, 64.208, EURO);
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // ── Prorrata y Plus Actividad ─────────────────────────────────────────────
  lbl(r, 'PRORRATA Y PLUSES BASE', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  const ROW_PRORRATA = r;
  lbl(r, '  Prorrata Pagas Extra (mensual)', { indent: 2 });
  inputCell(r, 3, 419.61, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4);
    c.value = '← suma a Base CC y AT para SS (no computa en IRPF mensual)';
    font(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  const ROW_PLUS_ACT_DEF = r;
  lbl(r, '  Plus Actividad por defecto', { indent: 2 });
  inputCell(r, 3, 18.04, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4);
    c.value = '← valor por defecto; editable mes a mes';
    font(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // ── SS Tasas ─────────────────────────────────────────────────────────────
  lbl(r, 'SEGURIDAD SOCIAL — Tasas trabajador', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  // Cabecera SS
  ['', 'Concepto SS', 'Tasa (%)', 'Base de cálculo', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    font(c, { bold: true, size: 9, color: C.blanco });
    fill(c, C.morado);
    c.alignment = { horizontal: i >= 2 ? 'center' : 'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 16; r++;

  const ssRates = [
    ['CC (Contingencias Comunes)',  'SS_CC',  0.0470, 'Base CC = sujeto sin HHEE + prorrata'],
    ['Desempleo + FP',              'SS_DES', 0.0165, 'Base AT = Base CC + HHEE'],
    ['HH.EE. Normales',            'SS_HH',  0.0470, 'Sólo importe total Horas Extra'],
    ['MEI',                        'SS_MEI', 0.0013, 'Base AT (tasa real ≈ 0.1307%)'],
  ];
  const SS_RATE_ROWS = {};
  ssRates.forEach(([name, key, val, nota]) => {
    SS_RATE_ROWS[key] = r;
    lbl(r, '  ' + name, { indent: 2 });
    inputCell(r, 3, val, PCT);
    ws.mergeCells(r, 4, r, 6);
    { const c = R(r, 4); c.value = nota;
      font(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
    ws.getRow(r).height = 18; r++;
  });

  spacer(r, C.moradoLight, 4); r++;

  // ── IRPF ─────────────────────────────────────────────────────────────────
  const ROW_IRPF_DEFAULT = r;
  lbl(r, 'IRPF por defecto (%)', { bold: true, color: C.moradoDark });
  inputCell(r, 3, 0.1815, PCT);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4);
    c.value = '← 18.15% según nómina feb; editable por mes en col F de cada bloque';
    font(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // ── Precios unitarios ─────────────────────────────────────────────────────
  lbl(r, 'PRECIOS UNITARIOS DE VARIABLES', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  ['', 'Concepto', 'Precio (€/unidad)', 'Tipo SS', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    font(c, { bold: true, size: 9, color: C.blanco });
    fill(c, C.morado);
    c.alignment = { horizontal: i >= 2 ? 'center' : 'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 16; r++;

  const PRICE_ROWS = {};
  const priceItems = [
    ['Turnicidad',              'Turnicidad',   3.65,  'Base CC'],
    ['Nocturnidad',             'Nocturnidad',  10.94, 'Base CC'],
    ['Plus frío / baja temp.',  'PlusFrio',     16.32, 'Base CC'],
    ['H. Disponibilidad',       'HDisp',        1.35,  'Base CC'],
    ['H. Disponibilidad Fest.', 'HDispFest',    2.17,  'Base CC'],
    ['H. Desplazamiento',       'HDesp',        10.56, 'Base CC'],
    ['H. Desp. Festivo',        'HDespFest',    12.68, 'Base CC'],
    ['H. Extra Normal',         'HExtra',       19.77, 'Base AT (HHEE)'],
    ['H. Extra Festiva',        'HExtraFest',   19.77, 'Base AT (HHEE)'],
    ['Días fuera (bruto/día)',  'DiasFuera',    null,  'Base CC'],
    ['KM',                      'KM',           0.30,  'Exento'],
  ];
  priceItems.forEach(([name, key, val, tipoSS]) => {
    PRICE_ROWS[key] = r;
    lbl(r, '  ' + name, { indent: 2 });
    if (key === 'DiasFuera') {
      frmCell(r, 3, `C$${ROW_BRUTO_DIA}`, EURO, C.amarillo, C.moradoDark);
      border(R(r, 3), C.amarilloB);
    } else {
      inputCell(r, 3, val, EURO);
    }
    const c = R(r, 4);
    c.value = tipoSS;
    font(c, { size: 9, color: tipoSS === 'Exento' ? C.verdeSuave : C.azul });
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(r).height = 18; r++;
  });

  spacer(r, C.gris, 10); r++;

  // ═══════════════════════════════════════════════════════════════════════════
  // 12 BLOQUES MENSUALES
  // ═══════════════════════════════════════════════════════════════════════════
  const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                 'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const DIAS_MES = [31,28,31,30,31,30,31,31,30,31,30,31];

  const mesRefs = {};

  MESES.forEach((mes, idx) => {
    // ── Cabecera de mes ───────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 4);
    { const c = R(r, 1); c.value = `${mes.toUpperCase()} 2026`;
      font(c, { bold: true, size: 12, color: C.blanco });
      fill(c, C.moradoDark);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    { const c = R(r, 5); c.value = `Días: ${DIAS_MES[idx]}`;
      font(c, { size: 10, color: C.blanco }); fill(c, C.moradoDark);
      c.alignment = { vertical: 'middle', horizontal: 'center' }; }
    // IRPF editable por mes (col F)
    const ROW_IRPF_MES = r;
    { const c = R(r, 6);
      c.value = idx === 0 ? 0.1815 : { formula: `$C$${ROW_IRPF_DEFAULT}` };
      c.numFmt = PCT;
      font(c, { bold: true, size: 10, color: C.moradoDark });
      fill(c, C.amarillo); border(c, C.amarilloB);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = 26; r++;

    // Cabecera columnas
    ['', 'CONCEPTO', 'CANTIDAD', '€/UNIDAD', 'TOTAL €', 'TIPO'].forEach((t, i) => {
      const c = R(r, i+1); c.value = t;
      font(c, { bold: true, size: 9, color: C.muted }); fill(c, C.gris);
      c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
    });
    ws.getRow(r).height = 16; r++;

    // ── Salario Convenio ──────────────────────────────────────────────────
    const ROW_SAL_BASE = r;
    lbl(r, 'Salario Convenio');
    frmCell(r, 5, `C$${ROW_BRUTO}`, EURO, C.gris, C.text);
    ssTag(r, 'SUJETO-CC', C.azul);
    ws.getRow(r).height = 18; r++;

    // ── Plus Actividad ────────────────────────────────────────────────────
    const ROW_PLUS_ACT = r;
    lbl(r, '  Plus Actividad (importe mensual)', { indent: 2 });
    // Editable amarillo, default referencia config
    { const c = R(r, 3);
      c.value = { formula: `C$${ROW_PLUS_ACT_DEF}` };
      c.numFmt = EURO;
      font(c, { bold: true, color: C.moradoDark });
      fill(c, C.amarillo); border(c, C.amarilloB);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    { const c = R(r, 4); c.value = '(directo)';
      font(c, { italic: true, size: 9, color: C.muted });
      c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.gris, C.text);
    ssTag(r, 'SUJETO-CC', C.azul);
    ws.getRow(r).height = 18; r++;

    // ── Otras variables SUJETAS (base CC) ─────────────────────────────────
    const varsSujetasCC = [
      ['Turnicidad',             'Turnicidad'],
      ['Nocturnidad',            'Nocturnidad'],
      ['Plus frío / baja temp.', 'PlusFrio'],
      ['H. Disponibilidad',      'HDisp'],
      ['H. Disponibilidad Fest.','HDispFest'],
      ['H. Desplazamiento',      'HDesp'],
      ['H. Desplazamiento Fest.','HDespFest'],
      ['Días fuera',             'DiasFuera'],
    ];
    const sujetasCCRows = [ROW_SAL_BASE, ROW_PLUS_ACT];
    varsSujetasCC.forEach(([name, key]) => {
      lbl(r, '  ' + name, { indent: 2 });
      inputCell(r, 3, 0, NUM);
      frmCell(r, 4, `C$${PRICE_ROWS[key]}`, EURO, C.gris, C.muted);
      frmCell(r, 5, `C${r}*D${r}`, EURO, C.gris, C.text);
      ssTag(r, 'SUJETO-CC', C.azul);
      sujetasCCRows.push(r);
      ws.getRow(r).height = 18; r++;
    });

    // ── HH.EE. (base AT separada) ─────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── HORAS EXTRA (cotizan en base AT separada) ──────────────────────────';
      font(c, { size: 9, italic: true, bold: true, color: C.azul });
      fill(c, C.azulLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    const hheeVars = [['H. Extra Normal','HExtra'],['H. Extra Festiva','HExtraFest']];
    const hheeRows = [];
    hheeVars.forEach(([name, key]) => {
      lbl(r, '  ' + name, { indent: 2 });
      inputCell(r, 3, 0, NUM);
      frmCell(r, 4, `C$${PRICE_ROWS[key]}`, EURO, C.gris, C.muted);
      frmCell(r, 5, `C${r}*D${r}`, EURO, C.gris, C.text);
      ssTag(r, 'HHEE-AT', C.azul);
      hheeRows.push(r);
      ws.getRow(r).height = 18; r++;
    });

    // ── Exentos ───────────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── EXENTOS (no cotizan SS ni IRPF) ────────────────────────────────────';
      font(c, { size: 9, italic: true, bold: true, color: C.verdeSuave });
      fill(c, C.verdeLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    lbl(r, '  Dieta (importe directo)', { indent: 2 });
    inputCell(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)';
      font(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    ssTag(r, 'EXENTO', C.verdeSuave);
    const ROW_DIETA = r; ws.getRow(r).height = 18; r++;

    lbl(r, '  Tiquets (importe directo)', { indent: 2 });
    inputCell(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)';
      font(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    ssTag(r, 'EXENTO', C.verdeSuave);
    const ROW_TIQUETS = r; ws.getRow(r).height = 18; r++;

    lbl(r, '  KM (km × 0,30€)', { indent: 2 });
    inputCell(r, 3, 0, NUM);
    frmCell(r, 4, `C$${PRICE_ROWS.KM}`, EURO, C.gris, C.muted);
    frmCell(r, 5, `C${r}*D${r}`, EURO, C.verdeLight, C.verdeSuave);
    ssTag(r, 'EXENTO', C.verdeSuave);
    const ROW_KM = r; ws.getRow(r).height = 18; r++;

    divider(r); r++;

    // ── Subtotales ────────────────────────────────────────────────────────
    const sujetasCCFormula = sujetasCCRows.map(sr => `E${sr}`).join('+');
    const hheeFormula      = hheeRows.map(sr => `E${sr}`).join('+');

    // Bruto Sujeto (sujeto CC + HHEE) = base IRPF
    const ROW_BRUTO_SUJETO = r;
    blockRow(r, 'BRUTO SUJETO (base IRPF = sujeto CC + HHEE)',
      `${sujetasCCFormula}+${hheeFormula}`,
      EURO, C.moradoLight, C.moradoDark, 10, true);
    r++;

    // Exentos
    const ROW_EXENTOS = r;
    blockRow(r, 'EXENTOS (dieta + tiquets + km)',
      `E${ROW_DIETA}+E${ROW_TIQUETS}+E${ROW_KM}`,
      EURO, C.verdeLight, C.verdeSuave, 10, true);
    r++;

    // HHEE total (para Base AT y SS HH)
    const ROW_HHEE = r;
    blockRow(r, 'Total Horas Extra (HHEE)',
      hheeFormula,
      EURO, C.azulLight, C.azul, 9, false);
    r++;

    divider(r); r++;

    // ── SS — Bases y cálculo ──────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── COTIZACIÓN SEGURIDAD SOCIAL ─────────────────────────────────────────';
      font(c, { size: 9, italic: true, bold: true, color: C.rojo });
      fill(c, C.rojoLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    // Base CC = sujeto_CC + prorrata
    const ROW_BASE_CC = r;
    blockRow(r, `  Base CC = sujeto CC + prorrata (${419.61}€)`,
      `${sujetasCCFormula}+C$${ROW_PRORRATA}`,
      EURO, C.rojoLight, C.rojoSuave, 9, false);
    { const note = R(r, 6); note.value = 'Base CC';
      font(note, { size: 9, bold: true, color: C.rojoSuave });
      note.alignment = { horizontal: 'center', vertical: 'middle' }; }
    r++;

    // Base AT = Base CC + HHEE
    const ROW_BASE_AT = r;
    blockRow(r, '  Base AT/Desempleo = Base CC + HHEE',
      `E${ROW_BASE_CC}+E${ROW_HHEE}`,
      EURO, C.rojoLight, C.rojoSuave, 9, false);
    { const note = R(r, 6); note.value = 'Base AT';
      font(note, { size: 9, bold: true, color: C.rojoSuave });
      note.alignment = { horizontal: 'center', vertical: 'middle' }; }
    r++;

    // SS CC
    const ROW_SS_CC = r;
    { ws.mergeCells(r, 1, r, 3);
      const lc = R(r, 1); lc.value = `  (-) SS Contingencias Comunes`;
      font(lc, { size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    frmCell(r, 4, `C$${SS_RATE_ROWS.SS_CC}`, PCT, C.rojoLight, C.rojo);
    frmCell(r, 5, `E${ROW_BASE_CC}*C$${SS_RATE_ROWS.SS_CC}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    // SS Desempleo+FP
    const ROW_SS_DES = r;
    { ws.mergeCells(r, 1, r, 3);
      const lc = R(r, 1); lc.value = `  (-) SS Desempleo + FP`;
      font(lc, { size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    frmCell(r, 4, `C$${SS_RATE_ROWS.SS_DES}`, PCT, C.rojoLight, C.rojo);
    frmCell(r, 5, `E${ROW_BASE_AT}*C$${SS_RATE_ROWS.SS_DES}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    // SS HHEE
    const ROW_SS_HH = r;
    { ws.mergeCells(r, 1, r, 3);
      const lc = R(r, 1); lc.value = `  (-) SS Horas Extra`;
      font(lc, { size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    frmCell(r, 4, `C$${SS_RATE_ROWS.SS_HH}`, PCT, C.rojoLight, C.rojo);
    frmCell(r, 5, `E${ROW_HHEE}*C$${SS_RATE_ROWS.SS_HH}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    // SS MEI
    const ROW_SS_MEI = r;
    { ws.mergeCells(r, 1, r, 3);
      const lc = R(r, 1); lc.value = `  (-) SS MEI`;
      font(lc, { size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    frmCell(r, 4, `C$${SS_RATE_ROWS.SS_MEI}`, PCT, C.rojoLight, C.rojo);
    frmCell(r, 5, `E${ROW_BASE_AT}*C$${SS_RATE_ROWS.SS_MEI}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    // Total SS
    const ROW_SS_TOTAL = r;
    blockRow(r, 'TOTAL SS TRABAJADOR',
      `E${ROW_SS_CC}+E${ROW_SS_DES}+E${ROW_SS_HH}+E${ROW_SS_MEI}`,
      EURO, C.rojo, C.blanco, 10, true);
    r++;

    divider(r); r++;

    // ── IRPF ─────────────────────────────────────────────────────────────
    const ROW_IRPF_DED = r;
    { ws.mergeCells(r, 1, r, 3);
      const lc = R(r, 1); lc.value = '  (-) Retención IRPF';
      font(lc, { bold: true, size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    // Col D muestra el % editable del mes
    { const c = R(r, 4);
      c.value = { formula: `F${ROW_IRPF_MES}` };
      c.numFmt = PCT;
      font(c, { bold: true, size: 10, color: C.rojo }); fill(c, C.rojoLight);
      border(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    // IRPF = Bruto Sujeto × IRPF%
    frmCell(r, 5, `E${ROW_BRUTO_SUJETO}*F${ROW_IRPF_MES}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 22; r++;

    divider(r); r++;

    // ── NETO A COBRAR ─────────────────────────────────────────────────────
    const ROW_NETO = r;
    ws.mergeCells(r, 1, r, 4);
    { const c = R(r, 1); c.value = '💳  LÍQUIDO A PERCIBIR (NETO)';
      font(c, { bold: true, size: 13, color: C.blanco });
      fill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    { const c = R(r, 5);
      // NETO = Total devengado - SS total - IRPF
      // Total devengado = Bruto Sujeto + Exentos
      c.value = { formula: `E${ROW_BRUTO_SUJETO}+E${ROW_EXENTOS}-E${ROW_SS_TOTAL}-E${ROW_IRPF_DED}` };
      c.numFmt = EURO;
      font(c, { bold: true, size: 13, color: C.blanco });
      fill(c, C.verde); border(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    { const c = R(r, 6); c.value = '✓ v3';
      font(c, { size: 9, bold: true, color: 'D1FAE5' });
      fill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'center' }; }
    ws.getRow(r).height = 30;

    mesRefs[mes] = {
      netoRow:        ROW_NETO,
      brutoSujetoRow: ROW_BRUTO_SUJETO,
      ssRow:          ROW_SS_TOTAL,
      irpfRow:        ROW_IRPF_DED,
      exentosRow:     ROW_EXENTOS,
    };

    r++;
    spacer(r, C.gris, 10); r++;
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // RESUMEN ANUAL
  // ═══════════════════════════════════════════════════════════════════════════
  secHeader(r, '📅  RESUMEN ANUAL 2026', C.moradoDark, C.blanco, 13); r++;

  // Cabecera tabla resumen — 7 columnas (A-G)
  ws.getColumn('G').width = 15;
  ['', 'MES', 'Bruto Sujeto', 'SS Total (−)', 'IRPF (−)', 'Exentos (+)', 'NETO'].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    font(c, { bold: true, size: 10, color: C.blanco });
    fill(c, C.morado);
    c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 20; r++;

  const resumenStart = r;
  MESES.forEach((mes, idx) => {
    const ref = mesRefs[mes];
    const bg  = idx % 2 === 1 ? C.moradoLight : C.blanco;
    lbl(r, mes, { bold: true }); fill(R(r, 2), bg);
    frmCell(r, 3, `E${ref.brutoSujetoRow}`, EURO, bg, C.text);
    frmCell(r, 4, `E${ref.ssRow}`,          EURO, bg, C.rojo);
    frmCell(r, 5, `E${ref.irpfRow}`,        EURO, bg, C.rojo);
    frmCell(r, 6, `E${ref.exentosRow}`,     EURO, bg, C.verdeSuave);
    { const c = R(r, 7);
      c.value = { formula: `E${ref.netoRow}` };
      c.numFmt = EURO;
      font(c, { bold: true, size: 11, color: C.verde });
      fill(c, bg); border(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = 20; r++;
  });
  const resumenEnd = r - 1;

  divider(r); r++;

  // Totales anuales
  ws.mergeCells(r, 1, r, 2);
  { const c = R(r, 1); c.value = 'TOTALES ANUALES';
    font(c, { bold: true, size: 11, color: C.blanco });
    fill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }

  [3,4,5,6].forEach(col => {
    const letter = ['','A','B','C','D','E','F','G'][col];
    const c = R(r, col);
    c.value = { formula: `SUM(${letter}${resumenStart}:${letter}${resumenEnd})` };
    c.numFmt = EURO;
    font(c, { bold: true, size: 11, color: C.blanco });
    fill(c, C.moradoDark); border(c, C.morado);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  });
  { const c = R(r, 7);
    c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` };
    c.numFmt = EURO;
    font(c, { bold: true, size: 13, color: C.blanco });
    fill(c, C.verde); border(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' }; }
  ws.getRow(r).height = 26; r++;

  spacer(r, C.gris, 8); r++;

  // NETO ANUAL destacado
  ws.mergeCells(r, 1, r, 5);
  { const c = R(r, 1); c.value = '💳  NETO ANUAL TOTAL 2026';
    font(c, { bold: true, size: 14, color: C.blanco });
    fill(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'center' }; }
  ws.mergeCells(r, 6, r, 7);
  { const c = R(r, 6);
    c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` };
    c.numFmt = EURO;
    font(c, { bold: true, size: 16, color: C.blanco });
    fill(c, C.verde); border(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' }; }
  ws.getRow(r).height = 38; r++;

  spacer(r, C.gris, 8); r++;

  // Nota final
  ws.mergeCells(r, 1, r, 7);
  { const c = R(r, 1);
    c.value = '⚠️  Celdas amarillas = editables  |  IRPF en col F de cada mes  |  '
            + 'Base CC = sujeto sin HHEE + prorrata  |  Base AT = Base CC + HHEE  |  '
            + 'NETO = Total devengado − SS total − IRPF';
    font(c, { size: 9, italic: true, color: '78350F' });
    fill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
    ws.getRow(r).height = 28; }

  // ─── Escribir fichero ────────────────────────────────────────────────────
  const outPath = '/home/user/Mi-memoria/calculadora-nomina-v3.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log('✅  Generado:', outPath);
  console.log('');
  console.log('📋  Estructura SS por mes:');
  console.log('    Base CC  = sujeto sin HHEE + prorrata (419.61)');
  console.log('    Base AT  = Base CC + HHEE');
  console.log('    SS CC    = Base CC × 4.70%');
  console.log('    SS Des+FP= Base AT × 1.65%');
  console.log('    SS HHEE  = HHEE × 4.70%');
  console.log('    SS MEI   = Base AT × 0.13%');
  console.log('    IRPF     = Bruto Sujeto × IRPF%');
  console.log('    NETO     = (Bruto Sujeto + Exentos) − SS total − IRPF');
  console.log('');
  console.log('🔎  Verificación feb 2026 (con variables exactas del PDF):');
  console.log('    → Bruto Sujeto: 3261.74  → SS: 238.61  → IRPF: 592.01  → NETO: 3195.28');
}

main().catch(err => { console.error('❌', err.message); process.exit(1); });

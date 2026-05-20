/**
 * Calculadora de Nómina v3 — SS multi-base + Artículo 7p IRPF
 *
 * SS real (MECALUX):
 *   Base CC  = sujeto_sin_HHEE + prorrata_pagas_extra (puede ser 0)
 *   Base AT  = Base CC + HHEE
 *   SS CC (4.70%)          → Base CC × 4.70%
 *   SS Desempleo+FP (1.65%)→ Base AT × 1.65%
 *   SS HH.EE. (4.70%)      → HHEE × 4.70%
 *   SS MEI (0.13%)         → Base AT × 0.13%
 *
 *   IRPF = Bruto Sujeto × IRPF%
 *   NETO = (Bruto Sujeto + Exentos) − Total SS − IRPF
 *
 * Artículo 7p LIRPF:
 *   Los rendimientos del trabajo percibidos por trabajos realizados en el
 *   extranjero están exentos de IRPF hasta 60.100€/año si:
 *     - El trabajo se realiza físicamente fuera de España
 *     - El país destino tiene IRPF (no paraíso fiscal)
 *   Renta 7p = proporción salario días fuera + plus desplazamiento
 *            + plus desplazamiento festivo + pago días fuera (bruto/día)
 *   El IRPF retenido sobre renta 7p puede reclamarse en la Declaración de la Renta
 */

const ExcelJS = require('/home/user/Mi-memoria/node_modules/exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mi Memoria — Calculadora v3';
  wb.created = new Date();

  const C = {
    moradoDark:  '4338CA', morado: '6366F1', moradoLight: 'EDE9FE',
    verde:       '059669', verdeLight: 'D1FAE5',
    rojo:        'DC2626', rojoLight: 'FEE2E2',
    amarillo:    'FFFBEB', amarilloB: 'F59E0B',
    naranja:     'FFF7ED', naranjaB: 'EA580C', naranjaDark: '7C2D12',
    gris:        'F8FAFC', grisBorde: 'E2E8F0',
    azulLight:   'EFF6FF', azul: '1D4ED8',
    text:        '0F172A', muted: '64748B', blanco: 'FFFFFF',
    rojoSuave:   'B91C1C', verdeSuave: '065F46',
  };

  const ws = wb.addWorksheet('Calculadora Nómina 2026', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 },
    properties: { tabColor: { argb: C.moradoDark } },
  });

  ws.getColumn('A').width = 2;   // margen
  ws.getColumn('B').width = 35;  // concepto
  ws.getColumn('C').width = 13;  // cantidad
  ws.getColumn('D').width = 13;  // precio
  ws.getColumn('E').width = 16;  // total
  ws.getColumn('F').width = 12;  // tipo / IRPF

  const EURO = '#,##0.00 "€"';
  const PCT  = '0.00%';
  const NUM  = '#,##0.00';

  const R = (row, col) => ws.getCell(row, col);

  function fill(c, argb) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }
  function brd(c, color = C.grisBorde, style = 'thin') {
    const s = { style, color: { argb: color } };
    c.border = { top: s, bottom: s, left: s, right: s };
  }
  function fnt(c, { bold=false, size=10, color=C.text, italic=false } = {}) {
    c.font = { bold, size, color: { argb: color }, italic, name: 'Calibri' };
  }

  function secHeader(row, text, bg=C.moradoDark, fg=C.blanco, sz=11) {
    ws.mergeCells(row, 1, row, 6);
    const c = R(row, 1);
    c.value = text; fnt(c, { bold: true, size: sz, color: fg });
    fill(c, bg); c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(row).height = 24;
  }
  function spacer(row, bg=C.gris, h=5) {
    ws.mergeCells(row, 1, row, 6); fill(R(row, 1), bg); ws.getRow(row).height = h;
  }
  function divider(row) {
    ws.mergeCells(row, 1, row, 6); fill(R(row, 1), C.grisBorde); ws.getRow(row).height = 3;
  }

  function inputCell(row, col, value, fmt=EURO) {
    const c = R(row, col);
    c.value = value; c.numFmt = fmt;
    fnt(c, { bold: true, color: C.moradoDark });
    fill(c, C.amarillo); brd(c, C.amarilloB);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }
  function frmCell(row, col, formula, fmt=EURO, bg=C.gris, fg=C.text) {
    const c = R(row, col);
    c.value = { formula }; c.numFmt = fmt;
    fnt(c, { bold: true, color: fg });
    fill(c, bg); brd(c, C.grisBorde);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }
  function lbl(row, text, opts={}) {
    const c = R(row, 2);
    c.value = text;
    fnt(c, { size: 10, color: opts.color||C.text, bold: opts.bold||false, italic: opts.italic||false });
    c.alignment = { vertical: 'middle', indent: opts.indent||1 };
    if (opts.bg) fill(c, opts.bg);
    return c;
  }
  function blockRow(row, label, formula, fmt=EURO, bg=C.gris, fg=C.text, sz=10, bold=false) {
    ws.mergeCells(row, 1, row, 4);
    const lc = R(row, 1); lc.value = label;
    fnt(lc, { bold, size: sz, color: fg }); fill(lc, bg);
    lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    if (formula) {
      const fc = R(row, 5);
      fc.value = { formula }; fc.numFmt = fmt;
      fnt(fc, { bold, size: sz, color: fg });
      fill(fc, bg); brd(fc, C.grisBorde);
      fc.alignment = { vertical: 'middle', horizontal: 'right' };
    }
    ws.getRow(row).height = 20;
  }
  // Bloque 3 col: merge A-C label | D fórmula/tasa | E importe
  function ssRow(row, label, rateFormula, amtFormula, bg=C.rojoLight, fg=C.rojo) {
    ws.mergeCells(row, 1, row, 3);
    const lc = R(row, 1); lc.value = label;
    fnt(lc, { size: 10, color: fg }); fill(lc, bg);
    lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 };
    frmCell(row, 4, rateFormula, PCT, bg, fg);
    frmCell(row, 5, amtFormula,  EURO, bg, fg);
    ws.getRow(row).height = 20;
  }
  // Bloque naranja para Art. 7p
  function artRow(row, label, formula, fmt=EURO) {
    ws.mergeCells(row, 1, row, 4);
    const lc = R(row, 1); lc.value = label;
    fnt(lc, { size: 9, color: C.naranjaDark }); fill(lc, C.naranja);
    lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 };
    if (formula) {
      const fc = R(row, 5);
      fc.value = { formula }; fc.numFmt = fmt;
      fnt(fc, { size: 9, color: C.naranjaDark });
      fill(fc, C.naranja); brd(fc, C.grisBorde);
      fc.alignment = { vertical: 'middle', horizontal: 'right' };
    }
    ws.getRow(row).height = 18;
  }

  let r = 1;

  // ══════════════════════════════════════════════════════════════════════════
  // TÍTULO
  // ══════════════════════════════════════════════════════════════════════════
  ws.mergeCells(r, 1, r, 6);
  { const c = R(r, 1); c.value = 'CALCULADORA NÓMINA 2026 — v3';
    fnt(c, { bold: true, size: 18, color: C.blanco });
    fill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 42; } r++;

  ws.mergeCells(r, 1, r, 6);
  { const c = R(r, 1);
    c.value = 'SS en bases separadas (Base CC y Base AT)  |  Artículo 7p: renta trabajos en extranjero potencialmente exenta de IRPF hasta 60.100€/año';
    fnt(c, { size: 9, italic: true, color: '78350F' });
    fill(c, C.amarillo); c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 18; } r++;
  spacer(r, C.gris, 8); r++;

  // ══════════════════════════════════════════════════════════════════════════
  // CONFIGURACIÓN
  // ══════════════════════════════════════════════════════════════════════════
  secHeader(r, '⚙️   CONFIGURACIÓN — Celdas amarillas = editables'); r++;
  ['', 'Parámetro', 'Valor', '', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    fnt(c, { bold: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' };
  });
  ws.getRow(r).height = 16; r++;

  // Salario
  const ROW_BRUTO = r;
  lbl(r, 'Salario Convenio mensual', { bold: true });
  inputCell(r, 3, 2063.39, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← salario base fijo mensual';
    fnt(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  const ROW_BRUTO_DIA = r;
  lbl(r, 'Bruto por día (días fuera)');
  inputCell(r, 3, 64.208, EURO);
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // Prorrata + Plus Actividad
  lbl(r, 'PRORRATA Y PLUSES BASE', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  const ROW_PRORRATA = r;
  lbl(r, '  Prorrata Pagas Extra en base SS', { indent: 2 });
  inputCell(r, 3, 0, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4);
    c.value = '← solo suma a Base CC/AT para SS; NO se cobra mensualmente. Si aparece en tu nómina, ponla (ej: 419,61€)';
    fnt(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1, wrapText: true }; }
  ws.getRow(r).height = 28; r++;

  const ROW_PLUS_ACT_DEF = r;
  lbl(r, '  Plus Actividad por defecto', { indent: 2 });
  inputCell(r, 3, 18.04, EURO);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← editable mes a mes en cada bloque';
    fnt(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // SS Tasas
  lbl(r, 'SEGURIDAD SOCIAL — Tasas trabajador', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;
  ['', 'Concepto SS', 'Tasa (%)', 'Base de cálculo', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    fnt(c, { bold: true, size: 9, color: C.blanco }); fill(c, C.morado);
    c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 16; r++;

  const ssRates = [
    ['CC (Contingencias Comunes)',  'SS_CC',  0.0470, 'Base CC = sujeto sin HHEE + prorrata'],
    ['Desempleo + FP',              'SS_DES', 0.0165, 'Base AT = Base CC + HHEE'],
    ['HH.EE. Normales',            'SS_HH',  0.0470, 'Solo importe total Horas Extra'],
    ['MEI',                        'SS_MEI', 0.0013, 'Base AT (tasa exacta ≈ 0.1307%)'],
  ];
  const SS_RATE_ROWS = {};
  ssRates.forEach(([name, key, val, nota]) => {
    SS_RATE_ROWS[key] = r;
    lbl(r, '  ' + name, { indent: 2 });
    inputCell(r, 3, val, PCT);
    ws.mergeCells(r, 4, r, 6);
    { const c = R(r, 4); c.value = nota;
      fnt(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
    ws.getRow(r).height = 18; r++;
  });

  spacer(r, C.moradoLight, 4); r++;

  // IRPF
  const ROW_IRPF_DEFAULT = r;
  lbl(r, 'IRPF por defecto (%)', { bold: true, color: C.moradoDark });
  inputCell(r, 3, 0.1815, PCT);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← 18.15% según nómina feb. Editable en col F de cada mes';
    fnt(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  spacer(r, C.moradoLight, 4); r++;

  // Precios unitarios
  lbl(r, 'PRECIOS UNITARIOS DE VARIABLES', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;
  ['', 'Concepto', 'Precio (€/unidad)', 'Tipo SS', '', ''].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    fnt(c, { bold: true, size: 9, color: C.blanco }); fill(c, C.morado);
    c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 16; r++;

  const PRICE_ROWS = {};
  [
    ['Turnicidad',              'Turnicidad',   3.65,  'Base CC'],
    ['Nocturnidad',             'Nocturnidad',  10.94, 'Base CC'],
    ['Plus frío / baja temp.',  'PlusFrio',     16.32, 'Base CC'],
    ['H. Disponibilidad',       'HDisp',        1.35,  'Base CC'],
    ['H. Disponibilidad Fest.', 'HDispFest',    2.17,  'Base CC'],
    ['H. Desplazamiento',       'HDesp',        10.56, 'Base CC + Art.7p'],
    ['H. Desp. Festivo',        'HDespFest',    12.68, 'Base CC + Art.7p'],
    ['H. Extra Normal',         'HExtra',       19.77, 'Base AT (HHEE)'],
    ['H. Extra Festiva',        'HExtraFest',   19.77, 'Base AT (HHEE)'],
    ['Días fuera (bruto/día)',  'DiasFuera',    null,  'Base CC + Art.7p'],
    ['KM',                      'KM',           0.30,  'Exento'],
  ].forEach(([name, key, val, tipoSS]) => {
    PRICE_ROWS[key] = r;
    lbl(r, '  ' + name, { indent: 2 });
    if (key === 'DiasFuera') {
      frmCell(r, 3, `C$${ROW_BRUTO_DIA}`, EURO, C.amarillo, C.moradoDark);
      brd(R(r, 3), C.amarilloB);
    } else {
      inputCell(r, 3, val, EURO);
    }
    const c = R(r, 4); c.value = tipoSS;
    fnt(c, { size: 9, color: tipoSS.includes('Art.7p') ? C.naranjaDark : tipoSS==='Exento' ? C.verdeSuave : C.azul });
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(r).height = 18; r++;
  });

  spacer(r, C.gris, 10); r++;

  // ══════════════════════════════════════════════════════════════════════════
  // 12 BLOQUES MENSUALES
  // ══════════════════════════════════════════════════════════════════════════
  const MESES    = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const DIAS_MES = [31,28,31,30,31,30,31,31,30,31,30,31];

  const mesRefs = {};

  MESES.forEach((mes, idx) => {
    const diasN = DIAS_MES[idx];

    // ── Cabecera de mes ─────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 4);
    { const c = R(r, 1); c.value = `${mes.toUpperCase()} 2026`;
      fnt(c, { bold: true, size: 12, color: C.blanco });
      fill(c, C.moradoDark); c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    { const c = R(r, 5); c.value = `Días: ${diasN}`;
      fnt(c, { size: 10, color: C.blanco }); fill(c, C.moradoDark);
      c.alignment = { vertical: 'middle', horizontal: 'center' }; }
    const ROW_IRPF_MES = r;
    { const c = R(r, 6);
      c.value = idx === 0 ? 0.1815 : { formula: `$C$${ROW_IRPF_DEFAULT}` };
      c.numFmt = PCT; fnt(c, { bold: true, size: 10, color: C.moradoDark });
      fill(c, C.amarillo); brd(c, C.amarilloB);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = 26; r++;

    // Cabecera columnas
    ['', 'CONCEPTO', 'CANTIDAD', '€/UNIDAD', 'TOTAL €', 'TIPO'].forEach((t, i) => {
      const c = R(r, i+1); c.value = t;
      fnt(c, { bold: true, size: 9, color: C.muted }); fill(c, C.gris);
      c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
    });
    ws.getRow(r).height = 16; r++;

    // ── Salario Convenio ────────────────────────────────────────────────────
    const ROW_SAL_BASE = r;
    lbl(r, 'Salario Convenio');
    frmCell(r, 5, `C$${ROW_BRUTO}`, EURO, C.gris, C.text);
    { const c = R(r, 6); c.value = 'SUJETO-CC';
      fnt(c, { size: 9, bold: true, color: C.azul }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    ws.getRow(r).height = 18; r++;

    // ── Plus Actividad (sujeto directo) ─────────────────────────────────────
    const ROW_PLUS_ACT = r;
    lbl(r, '  Plus Actividad', { indent: 2 });
    { const c = R(r, 3); c.value = { formula: `C$${ROW_PLUS_ACT_DEF}` }; c.numFmt = EURO;
      fnt(c, { bold: true, color: C.moradoDark }); fill(c, C.amarillo); brd(c, C.amarilloB);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    { const c = R(r, 4); c.value = '(directo)';
      fnt(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.gris, C.text);
    { const c = R(r, 6); c.value = 'SUJETO-CC';
      fnt(c, { size: 9, bold: true, color: C.azul }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    ws.getRow(r).height = 18; r++;

    // ── Variables sujetas Base CC ────────────────────────────────────────────
    const varsSujetasCC = [
      ['Turnicidad',             'Turnicidad',  false],
      ['Nocturnidad',            'Nocturnidad', false],
      ['Plus frío / baja temp.', 'PlusFrio',    false],
      ['H. Disponibilidad',      'HDisp',       false],
      ['H. Disponibilidad Fest.','HDispFest',   false],
      ['H. Desplazamiento',      'HDesp',       true ],  // Art. 7p
      ['H. Desp. Festivo',       'HDespFest',   true ],  // Art. 7p
      ['Días fuera',             'DiasFuera',   true ],  // Art. 7p
    ];

    const sujetasCCRows = [ROW_SAL_BASE, ROW_PLUS_ACT];
    let ROW_H_DESP_VAR     = null;
    let ROW_H_DESP_FEST_VAR = null;
    let ROW_DIAS_FUERA_VAR  = null;

    varsSujetasCC.forEach(([name, key, is7p]) => {
      const thisRow = r;
      lbl(r, '  ' + name, { indent: 2 });
      inputCell(r, 3, 0, NUM);
      frmCell(r, 4, `C$${PRICE_ROWS[key]}`, EURO, C.gris, C.muted);
      frmCell(r, 5, `C${r}*D${r}`, EURO, C.gris, C.text);
      { const c = R(r, 6); c.value = is7p ? 'CC+7p' : 'SUJETO-CC';
        fnt(c, { size: 9, bold: true, color: is7p ? C.naranjaDark : C.azul });
        c.alignment = { horizontal: 'center', vertical: 'middle' }; }
      sujetasCCRows.push(thisRow);
      if (key === 'HDesp')     ROW_H_DESP_VAR     = thisRow;
      if (key === 'HDespFest') ROW_H_DESP_FEST_VAR = thisRow;
      if (key === 'DiasFuera') ROW_DIAS_FUERA_VAR  = thisRow;
      ws.getRow(r).height = 18; r++;
    });

    // ── HH.EE. (Base AT) ────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── HORAS EXTRA (cotizan en base AT separada) ───────────────────────────';
      fnt(c, { size: 9, italic: true, bold: true, color: C.azul }); fill(c, C.azulLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    const hheeRows = [];
    [['H. Extra Normal','HExtra'],['H. Extra Festiva','HExtraFest']].forEach(([name, key]) => {
      hheeRows.push(r);
      lbl(r, '  ' + name, { indent: 2 });
      inputCell(r, 3, 0, NUM);
      frmCell(r, 4, `C$${PRICE_ROWS[key]}`, EURO, C.gris, C.muted);
      frmCell(r, 5, `C${r}*D${r}`, EURO, C.gris, C.text);
      { const c = R(r, 6); c.value = 'HHEE-AT';
        fnt(c, { size: 9, bold: true, color: C.azul }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
      ws.getRow(r).height = 18; r++;
    });

    // ── Exentos ─────────────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── EXENTOS (no cotizan SS ni IRPF) ─────────────────────────────────────';
      fnt(c, { size: 9, italic: true, bold: true, color: C.verdeSuave }); fill(c, C.verdeLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    const ROW_DIETA = r;
    lbl(r, '  Dieta (importe directo)', { indent: 2 });
    inputCell(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)';
      fnt(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    { const c = R(r, 6); c.value = 'EXENTO';
      fnt(c, { size: 9, bold: true, color: C.verdeSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    ws.getRow(r).height = 18; r++;

    const ROW_TIQUETS = r;
    lbl(r, '  Tiquets (importe directo)', { indent: 2 });
    inputCell(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)';
      fnt(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    { const c = R(r, 6); c.value = 'EXENTO';
      fnt(c, { size: 9, bold: true, color: C.verdeSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    ws.getRow(r).height = 18; r++;

    const ROW_KM = r;
    lbl(r, '  KM (km × 0,30€)', { indent: 2 });
    inputCell(r, 3, 0, NUM);
    frmCell(r, 4, `C$${PRICE_ROWS.KM}`, EURO, C.gris, C.muted);
    frmCell(r, 5, `C${r}*D${r}`, EURO, C.verdeLight, C.verdeSuave);
    { const c = R(r, 6); c.value = 'EXENTO';
      fnt(c, { size: 9, bold: true, color: C.verdeSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    ws.getRow(r).height = 18; r++;

    divider(r); r++;

    // ── Subtotales ──────────────────────────────────────────────────────────
    const sujetasCCFrm = sujetasCCRows.map(sr => `E${sr}`).join('+');
    const hheeFrm      = hheeRows.map(sr => `E${sr}`).join('+');

    const ROW_BRUTO_SUJETO = r;
    blockRow(r, 'BRUTO SUJETO (base IRPF = sujeto CC + HHEE)',
      `${sujetasCCFrm}+${hheeFrm}`, EURO, C.moradoLight, C.moradoDark, 10, true); r++;

    const ROW_EXENTOS = r;
    blockRow(r, 'EXENTOS (dieta + tiquets + km)',
      `E${ROW_DIETA}+E${ROW_TIQUETS}+E${ROW_KM}`, EURO, C.verdeLight, C.verdeSuave, 10, true); r++;

    const ROW_HHEE = r;
    blockRow(r, 'Total Horas Extra (HHEE)',
      hheeFrm, EURO, C.azulLight, C.azul, 9, false); r++;

    divider(r); r++;

    // ── SS ─────────────────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1); c.value = '── COTIZACIÓN SEGURIDAD SOCIAL ───────────────────────────────────────────';
      fnt(c, { size: 9, italic: true, bold: true, color: C.rojo }); fill(c, C.rojoLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 16; r++;

    const ROW_BASE_CC = r;
    blockRow(r, `  Base CC = sujeto CC + prorrata ($C$${ROW_PRORRATA})`,
      `${sujetasCCFrm}+C$${ROW_PRORRATA}`, EURO, C.rojoLight, C.rojoSuave, 9);
    { const c = R(r, 6); c.value = 'Base CC';
      fnt(c, { size: 9, bold: true, color: C.rojoSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    r++;

    const ROW_BASE_AT = r;
    blockRow(r, '  Base AT/Desempleo = Base CC + HHEE',
      `E${ROW_BASE_CC}+E${ROW_HHEE}`, EURO, C.rojoLight, C.rojoSuave, 9);
    { const c = R(r, 6); c.value = 'Base AT';
      fnt(c, { size: 9, bold: true, color: C.rojoSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
    r++;

    const ROW_SS_CC  = r; ssRow(r, '  (-) SS Contingencias Comunes', `C$${SS_RATE_ROWS.SS_CC}`,  `E${ROW_BASE_CC}*C$${SS_RATE_ROWS.SS_CC}`);  r++;
    const ROW_SS_DES = r; ssRow(r, '  (-) SS Desempleo + FP',        `C$${SS_RATE_ROWS.SS_DES}`, `E${ROW_BASE_AT}*C$${SS_RATE_ROWS.SS_DES}`); r++;
    const ROW_SS_HH  = r; ssRow(r, '  (-) SS Horas Extra',           `C$${SS_RATE_ROWS.SS_HH}`,  `E${ROW_HHEE}*C$${SS_RATE_ROWS.SS_HH}`);    r++;
    const ROW_SS_MEI = r; ssRow(r, '  (-) SS MEI',                   `C$${SS_RATE_ROWS.SS_MEI}`, `E${ROW_BASE_AT}*C$${SS_RATE_ROWS.SS_MEI}`); r++;

    const ROW_SS_TOTAL = r;
    blockRow(r, 'TOTAL SS TRABAJADOR',
      `E${ROW_SS_CC}+E${ROW_SS_DES}+E${ROW_SS_HH}+E${ROW_SS_MEI}`,
      EURO, C.rojo, C.blanco, 10, true); r++;

    divider(r); r++;

    // ── IRPF ────────────────────────────────────────────────────────────────
    const ROW_IRPF_DED = r;
    ws.mergeCells(r, 1, r, 3);
    { const lc = R(r, 1); lc.value = '  (-) Retención IRPF';
      fnt(lc, { bold: true, size: 10, color: C.rojo }); fill(lc, C.rojoLight);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    { const c = R(r, 4); c.value = { formula: `F${ROW_IRPF_MES}` }; c.numFmt = PCT;
      fnt(c, { bold: true, size: 10, color: C.rojo }); fill(c, C.rojoLight); brd(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    frmCell(r, 5, `E${ROW_BRUTO_SUJETO}*F${ROW_IRPF_MES}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 22; r++;

    divider(r); r++;

    // ── NETO ────────────────────────────────────────────────────────────────
    const ROW_NETO = r;
    ws.mergeCells(r, 1, r, 4);
    { const c = R(r, 1); c.value = '💳  LÍQUIDO A PERCIBIR (NETO)';
      fnt(c, { bold: true, size: 13, color: C.blanco }); fill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    { const c = R(r, 5);
      c.value = { formula: `E${ROW_BRUTO_SUJETO}+E${ROW_EXENTOS}-E${ROW_SS_TOTAL}-E${ROW_IRPF_DED}` };
      c.numFmt = EURO; fnt(c, { bold: true, size: 13, color: C.blanco });
      fill(c, C.verde); brd(c, C.verde); c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    { const c = R(r, 6); c.value = '✓ v3';
      fnt(c, { size: 9, bold: true, color: 'D1FAE5' }); fill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'center' }; }
    ws.getRow(r).height = 30; r++;

    // ── Art. 7p ─────────────────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1);
      c.value = '🌍  ARTÍCULO 7p — Renta por trabajos en el extranjero (potencialmente exenta de IRPF) ─────────';
      fnt(c, { size: 9, italic: false, bold: true, color: C.naranjaDark }); fill(c, C.naranja);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
    ws.getRow(r).height = 18; r++;

    // Fórmula 7p: prop. salario + plus desp + plus desp fest + pago días fuera
    // La proporción de salario por días fuera = (bruto/días_mes) × qty_días_fuera
    const propSalFrm = `(C$${ROW_BRUTO}/${diasN})*C${ROW_DIAS_FUERA_VAR}`;
    const desp7pFrm  = `E${ROW_H_DESP_VAR}`;
    const despF7pFrm = `E${ROW_H_DESP_FEST_VAR}`;
    const diasFPagoFrm = `E${ROW_DIAS_FUERA_VAR}`;

    artRow(r, '  Proporción salario días fuera España', propSalFrm);   r++;
    artRow(r, '  Plus Desplazamiento',                  desp7pFrm);    r++;
    artRow(r, '  Plus Desplazamiento Festivo',           despF7pFrm);   r++;
    artRow(r, '  Pago extra días fuera (bruto/día)',     diasFPagoFrm); r++;

    const ROW_7P_TOTAL = r;
    ws.mergeCells(r, 1, r, 4);
    { const lc = R(r, 1); lc.value = '  RENTA POTENCIALMENTE EXENTA (Art. 7p)';
      fnt(lc, { bold: true, size: 10, color: C.naranjaDark }); fill(lc, C.naranja);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 }; }
    { const fc = R(r, 5);
      fc.value = { formula: `${propSalFrm}+${desp7pFrm}+${despF7pFrm}+${diasFPagoFrm}` };
      fc.numFmt = EURO; fnt(fc, { bold: true, size: 10, color: C.naranjaDark });
      fill(fc, C.naranja); brd(fc, C.grisBorde);
      fc.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = 22; r++;

    const ROW_7P_IRPF = r;
    ws.mergeCells(r, 1, r, 3);
    { const lc = R(r, 1); lc.value = '  IRPF retenido sobre renta 7p (a reclamar en Renta)';
      fnt(lc, { bold: true, size: 10, color: C.naranjaDark }); fill(lc, C.naranja);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 3 }; }
    { const c = R(r, 4); c.value = { formula: `F${ROW_IRPF_MES}` }; c.numFmt = PCT;
      fnt(c, { bold: true, color: C.naranjaDark }); fill(c, C.naranja); brd(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    { const fc = R(r, 5);
      fc.value = { formula: `E${ROW_7P_TOTAL}*F${ROW_IRPF_MES}` }; fc.numFmt = EURO;
      fnt(fc, { bold: true, size: 10, color: C.naranjaDark });
      fill(fc, C.naranja); brd(fc, C.grisBorde);
      fc.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = 22; r++;

    // Nota art. 7p
    ws.mergeCells(r, 1, r, 6);
    { const c = R(r, 1);
      c.value = 'ℹ️  Art. 7p LIRPF: si los días fuera son en otro país con IRPF equivalente, esa renta puede estar exenta (límite 60.100€/año). '
              + 'El IRPF aquí retenido se recupera en la Declaración de la Renta. Consulta con gestor.';
      fnt(c, { size: 8, italic: true, color: C.naranjaDark }); fill(c, C.naranja);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true }; }
    ws.getRow(r).height = 30; r++;

    mesRefs[mes] = {
      netoRow:        ROW_NETO,
      brutoSujetoRow: ROW_BRUTO_SUJETO,
      ssRow:          ROW_SS_TOTAL,
      irpfRow:        ROW_IRPF_DED,
      exentosRow:     ROW_EXENTOS,
      art7pRow:       ROW_7P_TOTAL,
      art7pIrpfRow:   ROW_7P_IRPF,
    };

    spacer(r, C.gris, 10); r++;
  });

  // ══════════════════════════════════════════════════════════════════════════
  // RESUMEN ANUAL
  // ══════════════════════════════════════════════════════════════════════════
  ws.getColumn('G').width = 15;
  ws.getColumn('H').width = 18;
  ws.getColumn('I').width = 18;

  secHeader(r, '📅  RESUMEN ANUAL 2026', C.moradoDark, C.blanco, 13); r++;

  // Cabecera 9 cols
  ['','MES','Bruto Sujeto','SS Total (−)','IRPF (−)','Exentos (+)','NETO','Renta 7p','IRPF recup. 7p'].forEach((t, i) => {
    const c = R(r, i+1); c.value = t;
    fnt(c, { bold: true, size: 9, color: C.blanco });
    fill(c, i >= 7 ? C.naranjaB : C.morado);
    c.alignment = { horizontal: i>=2?'center':'left', vertical: 'middle', indent: i===1?1:0 };
  });
  ws.getRow(r).height = 20; r++;

  const resumenStart = r;
  MESES.forEach((mes, idx) => {
    const ref = mesRefs[mes];
    const bg  = idx % 2 === 1 ? C.moradoLight : C.blanco;
    const bgN = idx % 2 === 1 ? 'FFF0E6' : C.naranja;

    lbl(r, mes, { bold: true }); R(r, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    frmCell(r, 3, `E${ref.brutoSujetoRow}`, EURO, bg, C.text);
    frmCell(r, 4, `E${ref.ssRow}`,          EURO, bg, C.rojo);
    frmCell(r, 5, `E${ref.irpfRow}`,        EURO, bg, C.rojo);
    frmCell(r, 6, `E${ref.exentosRow}`,     EURO, bg, C.verdeSuave);
    { const c = R(r, 7); c.value = { formula: `E${ref.netoRow}` }; c.numFmt = EURO;
      fnt(c, { bold: true, size: 11, color: C.verde }); fill(c, bg); brd(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' }; }
    frmCell(r, 8, `E${ref.art7pRow}`,       EURO, bgN, C.naranjaDark);
    frmCell(r, 9, `E${ref.art7pIrpfRow}`,   EURO, bgN, C.naranjaDark);
    ws.getRow(r).height = 20; r++;
  });
  const resumenEnd = r - 1;

  divider(r); r++;

  // Totales anuales
  ws.mergeCells(r, 1, r, 2);
  { const c = R(r, 1); c.value = 'TOTALES ANUALES';
    fnt(c, { bold: true, size: 11, color: C.blanco }); fill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }; }
  [3,4,5,6].forEach(col => {
    const letter = ['','A','B','C','D','E','F','G','H','I'][col];
    const c = R(r, col);
    c.value = { formula: `SUM(${letter}${resumenStart}:${letter}${resumenEnd})` };
    c.numFmt = EURO; fnt(c, { bold: true, color: C.blanco });
    fill(c, C.moradoDark); brd(c, C.morado);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  });
  { const c = R(r, 7); c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` }; c.numFmt = EURO;
    fnt(c, { bold: true, size: 13, color: C.blanco }); fill(c, C.verde); brd(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' }; }
  [8,9].forEach(col => {
    const letter = ['','A','B','C','D','E','F','G','H','I'][col];
    const c = R(r, col); c.value = { formula: `SUM(${letter}${resumenStart}:${letter}${resumenEnd})` };
    c.numFmt = EURO; fnt(c, { bold: true, size: 11, color: C.blanco });
    fill(c, C.naranjaB); brd(c, C.naranjaB);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  });
  ws.getRow(r).height = 26; r++;

  spacer(r, C.gris, 8); r++;

  // Bloques destacados anuales
  ws.mergeCells(r, 1, r, 5);
  { const c = R(r, 1); c.value = '💳  NETO ANUAL TOTAL 2026';
    fnt(c, { bold: true, size: 14, color: C.blanco }); fill(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'center' }; }
  ws.mergeCells(r, 6, r, 9);
  { const c = R(r, 6);
    c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` }; c.numFmt = EURO;
    fnt(c, { bold: true, size: 16, color: C.blanco }); fill(c, C.verde); brd(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' }; }
  ws.getRow(r).height = 38; r++;

  spacer(r, C.gris, 6); r++;

  // Art. 7p anual destacado
  ws.mergeCells(r, 1, r, 9);
  { const c = R(r, 1);
    c.value = '🌍  ARTÍCULO 7p — Resumen anual';
    fnt(c, { bold: true, size: 12, color: C.blanco }); fill(c, C.naranjaB);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 28; } r++;

  [
    ['Renta exenta acumulada (Art. 7p)',  `SUM(H${resumenStart}:H${resumenEnd})`],
    ['Límite legal Art. 7p',             '60100'],
    ['Renta exenta restante hasta límite', `60100-SUM(H${resumenStart}:H${resumenEnd})`],
    ['IRPF retenido total a recuperar',  `SUM(I${resumenStart}:I${resumenEnd})`],
  ].forEach(([label, frm]) => {
    ws.mergeCells(r, 1, r, 6);
    { const lc = R(r, 1); lc.value = '  ' + label;
      fnt(lc, { bold: label.includes('recuperar'), size: 11, color: C.naranjaDark }); fill(lc, C.naranja);
      lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 }; }
    ws.mergeCells(r, 7, r, 9);
    { const fc = R(r, 7);
      fc.value = frm === '60100' ? 60100 : { formula: frm }; fc.numFmt = EURO;
      const isRec = label.includes('recuperar');
      fnt(fc, { bold: true, size: isRec ? 14 : 11, color: C.naranjaDark }); fill(fc, C.naranja);
      brd(fc, C.grisBorde); fc.alignment = { vertical: 'middle', horizontal: 'right' }; }
    ws.getRow(r).height = isRec => isRec ? 32 : 24;
    ws.getRow(r).height = label.includes('recuperar') ? 32 : 22; r++;
  });

  spacer(r, C.gris, 8); r++;

  ws.mergeCells(r, 1, r, 9);
  { const c = R(r, 1);
    c.value = '⚠️  Art. 7p: requiere que el trabajo se realice físicamente en el extranjero, el país destino tenga IRPF equivalente, y la empresa beneficiaria sea la entidad para la que trabajas. '
            + 'Consulta con tu gestor antes de aplicarlo en la Declaración de la Renta.  |  '
            + 'Prorrata pagas extra: ponla a 0 si no aparece en tu nómina, o al valor que muestre (ej: 419,61€) si aparece en la base de cotización.';
    fnt(c, { size: 9, italic: true, color: '78350F' }); fill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
    ws.getRow(r).height = 36; }

  // ── Escribir fichero ────────────────────────────────────────────────────
  const outPath = '/home/user/Mi-memoria/calculadora-nomina-v3.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log('✅  Generado:', outPath);
  console.log('');
  console.log('Cambios respecto a v2:');
  console.log('  - Prorrata pagas extra: 0 por defecto (solo base SS si aparece en nómina)');
  console.log('  - SS en 4 líneas: Base CC (CC 4.70%) + Base AT (Des+FP 1.65%, HHEE 4.70%, MEI 0.13%)');
  console.log('  - Art. 7p por mes: prop. salario + desplazamiento + días fuera');
  console.log('  - Resumen anual Art. 7p: total exento, límite 60.100€, IRPF a recuperar');
}

main().catch(err => { console.error('❌', err.message); process.exit(1); });

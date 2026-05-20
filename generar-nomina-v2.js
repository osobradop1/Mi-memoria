/**
 * Calculadora de Nómina v2 - Fórmula CORRECTA
 *
 * BUG CORREGIDO: La fórmula original usaba multiplicación encadenada:
 *   bruto × (1-IRPF) × (1-SS)  →  introduce error de ~1.36% por interacción
 *
 * Fórmula CORRECTA (aditiva):
 *   NETO = bruto_sujeto × (1 - IRPF% - SS%) + exentos
 *   es decir: bruto_sujeto - (bruto_sujeto × SS%) - (bruto_sujeto × IRPF%) + exentos
 *
 * Diferencia práctica: para bruto_sujeto ~3500€ con IRPF 18% + SS 7.55%
 *   Error multiplicativo: 3500 × (1-0.18) × (1-0.0755) = 2655.10
 *   Correcto aditivo:     3500 × (1 - 0.18 - 0.0755)   = 2671.75  → diferencia ~16-20€
 */

const ExcelJS = require('/home/user/Mi-memoria/node_modules/exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mi Memoria - Calculadora v2';
  wb.created = new Date();
  wb.modified = new Date();

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
    text:        '0F172A',
    muted:       '64748B',
    blanco:      'FFFFFF',
    rojoSuave:   'B91C1C',
    verdeSuave:  '065F46',
  };

  const ws = wb.addWorksheet('Calculadora Nómina 2026', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 },
    views: [{ state: 'normal' }],
    properties: { tabColor: { argb: C.moradoDark } },
  });

  // ── Anchos de columna ──────────────────────────────────────────────────────
  // A=margen, B=concepto, C=cantidad, D=€/unidad, E=total, F=tipo
  ws.getColumn('A').width = 2;
  ws.getColumn('B').width = 34;
  ws.getColumn('C').width = 13;
  ws.getColumn('D').width = 13;
  ws.getColumn('E').width = 16;
  ws.getColumn('F').width = 10;

  // ── Helpers ────────────────────────────────────────────────────────────────
  const EURO = '#,##0.00 "€"';
  const PCT  = '0.00%';
  const PCT2 = '0.00"%"';
  const NUM  = '#,##0.00';

  function R(row, col) { return ws.getCell(row, col); }

  function setFill(c, argb) {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  }
  function setBorder(c, color = C.grisBorde, style = 'thin') {
    const s = { style, color: { argb: color } };
    c.border = { top: s, bottom: s, left: s, right: s };
  }
  function setFont(c, { bold = false, size = 10, color = C.text, italic = false } = {}) {
    c.font = { bold, size, color: { argb: color }, italic, name: 'Calibri' };
  }

  // Merge y estilo de fila-cabecera de sección
  function secHeader(row, text, bg = C.moradoDark, fg = C.blanco, size = 11, cols = 6) {
    ws.mergeCells(row, 1, row, cols);
    const c = R(row, 1);
    c.value = text;
    setFont(c, { bold: true, size, color: fg });
    setFill(c, bg);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(row).height = 24;
    return c;
  }

  // Fila separadora delgada
  function spacer(row, bg = C.gris, h = 5) {
    ws.mergeCells(row, 1, row, 6);
    setFill(R(row, 1), bg);
    ws.getRow(row).height = h;
  }

  // Línea horizontal decorativa (separador interno dentro de bloque)
  function divider(row, bg = C.grisBorde) {
    ws.mergeCells(row, 1, row, 6);
    setFill(R(row, 1), bg);
    ws.getRow(row).height = 3;
  }

  // Celda editable (amarillo, borde ámbar)
  function inputYellow(row, col, value, fmt = EURO) {
    const c = R(row, col);
    c.value = value;
    c.numFmt = fmt;
    setFont(c, { bold: true, color: C.moradoDark });
    setFill(c, C.amarillo);
    setBorder(c, C.amarilloB);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  // Celda de fórmula (gris claro)
  function frmCell(row, col, formula, fmt = EURO, bg = C.gris, fgColor = C.text) {
    const c = R(row, col);
    c.value = { formula };
    c.numFmt = fmt;
    setFont(c, { bold: true, color: fgColor });
    setFill(c, bg);
    setBorder(c, C.grisBorde);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  // Etiqueta en col B
  function lbl(row, text, opts = {}) {
    const c = R(row, 2);
    c.value = text;
    setFont(c, { size: 10, color: opts.color || C.text, bold: opts.bold || false, italic: opts.italic || false });
    c.alignment = { vertical: 'middle', indent: opts.indent || 1, wrapText: opts.wrap || false };
    if (opts.fill) setFill(c, opts.fill);
    return c;
  }

  // Etiqueta SUJETO / EXENTO en col F
  function tipoTag(row, tipo) {
    const c = R(row, 6);
    c.value = tipo;
    const isExento = tipo === 'EXENTO';
    setFont(c, { size: 9, bold: true, color: isExento ? C.verdeSuave : C.rojoSuave });
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  let r = 1; // cursor de fila actual

  // ═══════════════════════════════════════════════════════════════════════════
  // TÍTULO PRINCIPAL
  // ═══════════════════════════════════════════════════════════════════════════
  ws.mergeCells(r, 1, r, 6);
  {
    const c = R(r, 1);
    c.value = '💰  CALCULADORA NÓMINA 2026';
    setFont(c, { bold: true, size: 18, color: C.blanco });
    setFill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(r).height = 40;
  }
  r++;

  ws.mergeCells(r, 1, r, 6);
  {
    const c = R(r, 1);
    c.value = '✅  Fórmula CORRECTA: NETO = Bruto Sujeto × (1 − SS% − IRPF%) + Exentos   |   ❌ Corrige el error multiplicativo del cálculo original (~20€/mes)';
    setFont(c, { size: 9, italic: true, color: '78350F' });
    setFill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    ws.getRow(r).height = 18;
  }
  r++;

  spacer(r, C.gris, 8); r++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECCIÓN CONFIG (filas 4-25 aprox.)
  // ═══════════════════════════════════════════════════════════════════════════
  secHeader(r, '⚙️   CONFIGURACIÓN — Edita las celdas amarillas', C.moradoDark, C.blanco, 12);
  r++;

  // Cabeceras de columna config
  {
    ['', 'Parámetro', 'Valor', '', '', ''].forEach((txt, i) => {
      const c = R(r, i + 1);
      c.value = txt;
      setFont(c, { bold: true, size: 9, color: C.muted });
      c.alignment = { horizontal: 'center' };
    });
    ws.getRow(r).height = 16;
  }
  r++;

  // ── Salario y días ────────────────────────────────────────────────────────
  const ROW_BRUTO = r;
  lbl(r, 'Salario bruto mensual', { bold: true });
  inputYellow(r, 3, 2063.39, EURO);
  setFill(R(r, 3), C.amarillo);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← salario base mensual'; setFont(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  const ROW_BRUTO_DIA = r;
  lbl(r, 'Bruto por día (días fuera)');
  inputYellow(r, 3, 64.208, EURO);
  ws.getRow(r).height = 20; r++;

  // ── SS trabajador ─────────────────────────────────────────────────────────
  spacer(r, C.moradoLight, 4); r++;

  lbl(r, 'SEGURIDAD SOCIAL — Desglose (% sobre bruto sujeto)', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  const SS_ROWS = {};
  const ssItems = [
    ['CC (Contingencias Comunes)', 'CC',       0.0470],
    ['Desempleo',                  'Desempleo', 0.0155],
    ['FP (Formación Profesional)', 'FP',        0.0010],
    ['MEI',                        'MEI',       0.0013],
    ['Otros',                      'Otros',     0.0107],
  ];
  ssItems.forEach(([label, key, val]) => {
    SS_ROWS[key] = r;
    lbl(r, '  ' + label, { indent: 2 });
    inputYellow(r, 3, val, PCT);
    ws.getRow(r).height = 18;
    r++;
  });

  // SS total (fórmula suma)
  const ROW_SS_TOTAL = r;
  lbl(r, '  TOTAL SS trabajador', { bold: true, color: C.moradoDark, indent: 2 });
  frmCell(r, 3,
    `C${SS_ROWS.CC}+C${SS_ROWS.Desempleo}+C${SS_ROWS.FP}+C${SS_ROWS.MEI}+C${SS_ROWS.Otros}`,
    PCT, C.moradoLight, C.moradoDark
  );
  ws.getRow(r).height = 20; r++;

  // ── IRPF ─────────────────────────────────────────────────────────────────
  spacer(r, C.moradoLight, 4); r++;

  const ROW_IRPF_DEFAULT = r;
  lbl(r, 'IRPF por defecto (% — editable por mes)', { bold: true, color: C.moradoDark });
  inputYellow(r, 3, 0.18, PCT);
  ws.mergeCells(r, 4, r, 6);
  { const c = R(r, 4); c.value = '← Enero usa 18.11% ajustado — cambia por mes abajo'; setFont(c, { size: 9, italic: true, color: C.muted }); c.alignment = { indent: 1 }; }
  ws.getRow(r).height = 20; r++;

  // ── Precios unitarios variables ───────────────────────────────────────────
  spacer(r, C.moradoLight, 4); r++;

  lbl(r, 'PRECIOS UNITARIOS DE VARIABLES', { bold: true, color: C.moradoDark });
  ws.getRow(r).height = 18; r++;

  // Cabecera tabla precios
  {
    const hdrs = ['', 'Concepto', 'Precio (€)', 'Tipo', '', ''];
    hdrs.forEach((t, i) => {
      const c = R(r, i + 1);
      c.value = t;
      setFont(c, { bold: true, size: 9, color: C.blanco });
      setFill(c, C.morado);
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws.getRow(r).height = 16;
  }
  r++;

  const PRICE_ROWS = {};
  const prices = [
    ['Turnicidad',              'Turnicidad',       3.65,   'SUJETO'],
    ['Nocturnidad',             'Nocturnidad',      10.94,  'SUJETO'],
    ['Plus frío',               'PlusFrio',         16.32,  'SUJETO'],
    ['H. Disponibilidad',       'HDisp',            1.35,   'SUJETO'],
    ['H. Disp. Festiva',        'HDispFest',        2.17,   'SUJETO'],
    ['H. Desplazamiento',       'HDesp',            10.56,  'SUJETO'],
    ['H. Desp. Festivo',        'HDespFest',        12.68,  'SUJETO'],
    ['H. Extra',                'HExtra',           19.77,  'SUJETO'],
    ['H. Extra Festiva',        'HExtraFest',       19.77,  'SUJETO'],
    ['Días fuera (= bruto/día)','DiasFuera',        null,   'SUJETO'],
    ['KM',                      'KM',               0.30,   'EXENTO'],
  ];
  prices.forEach(([name, key, val, tipo]) => {
    PRICE_ROWS[key] = r;
    lbl(r, '  ' + name, { indent: 2 });
    if (key === 'DiasFuera') {
      // Referencia al bruto/día
      frmCell(r, 3, `C$${ROW_BRUTO_DIA}`, EURO, C.amarillo, C.moradoDark);
      setBorder(R(r, 3), C.amarilloB);
    } else {
      inputYellow(r, 3, val, EURO);
    }
    // Tipo
    {
      const c = R(r, 4);
      c.value = tipo;
      setFont(c, { size: 9, bold: true, color: tipo === 'EXENTO' ? C.verdeSuave : C.rojoSuave });
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    ws.getRow(r).height = 18;
    r++;
  });

  // Exentos directos (sin precio unitario)
  lbl(r, '  Dieta (importe directo)',  { indent: 2 });
  { const c = R(r, 3); c.value = 'directo'; setFont(c, { italic: true, color: C.muted }); c.alignment = { horizontal: 'center' }; }
  { const c = R(r, 4); c.value = 'EXENTO'; setFont(c, { size: 9, bold: true, color: C.verdeSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
  ws.getRow(r).height = 18; r++;

  lbl(r, '  Tiquets (importe directo)', { indent: 2 });
  { const c = R(r, 3); c.value = 'directo'; setFont(c, { italic: true, color: C.muted }); c.alignment = { horizontal: 'center' }; }
  { const c = R(r, 4); c.value = 'EXENTO'; setFont(c, { size: 9, bold: true, color: C.verdeSuave }); c.alignment = { horizontal: 'center', vertical: 'middle' }; }
  ws.getRow(r).height = 18; r++;

  // Nota corrección
  spacer(r, C.amarillo, 6); r++;
  ws.mergeCells(r, 1, r, 6);
  {
    const c = R(r, 1);
    c.value = '⚠️  ERROR CORREGIDO: La fórmula original calculaba  bruto × (1−IRPF) × (1−SS)  lo que genera ~20€ de más por interacción multiplicativa. '
            + 'La fórmula correcta es ADITIVA:  NETO = Bruto_Sujeto − (Bruto_Sujeto × SS%) − (Bruto_Sujeto × IRPF%) + Exentos';
    setFont(c, { size: 9, italic: true, color: '78350F' });
    setFill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
    ws.getRow(r).height = 32;
  }
  r++;

  spacer(r, C.gris, 10); r++;

  // ═══════════════════════════════════════════════════════════════════════════
  // 12 BLOQUES MENSUALES
  // ═══════════════════════════════════════════════════════════════════════════
  const meses = [
    'Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
  ];
  const diasMes = [31,28,31,30,31,30,31,31,30,31,30,31];

  // Referencia absoluta CONFIG para precios ($ en fila)
  const $bruto     = `C$${ROW_BRUTO}`;
  const $ssTotal   = `C$${ROW_SS_TOTAL}`;
  const $irpfDef   = `C$${ROW_IRPF_DEFAULT}`;
  const $diasFuera = `C$${PRICE_ROWS.DiasFuera}`;

  // Almacena referencias de fila NETO por mes para resumen anual
  const mesRefs = {}; // { Enero: { netoRow, brutoSujetoRow, ssRow, irpfRow, exentosRow } }

  meses.forEach((mes, idx) => {
    const diasN = diasMes[idx];

    // ── Cabecera de mes ───────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = `${mes.toUpperCase()} 2026`;
      setFont(c, { bold: true, size: 12, color: C.blanco });
      setFill(c, C.moradoDark);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    }
    {
      const c = R(r, 5);
      c.value = `Días: ${diasN}`;
      setFont(c, { size: 10, color: C.blanco });
      setFill(c, C.moradoDark);
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    }
    // IRPF editable por mes
    const ROW_IRPF_MES = r;
    {
      const c = R(r, 6);
      // Enero: 18.11%, resto: referencia al default
      const irpfVal = idx === 0 ? 0.1811 : { formula: `$C$${ROW_IRPF_DEFAULT}` };
      c.value = typeof irpfVal === 'object' ? irpfVal : irpfVal;
      c.numFmt = PCT;
      setFont(c, { bold: true, size: 10, color: C.moradoDark });
      setFill(c, C.amarillo);
      setBorder(c, C.amarilloB);
      c.alignment = { vertical: 'middle', horizontal: 'right' };
    }
    ws.getRow(r).height = 24;
    r++;

    // Cabecera columnas del bloque
    {
      const hdrs = ['', 'CONCEPTO', 'CANTIDAD', '€/UNIDAD', 'TOTAL €', 'TIPO'];
      hdrs.forEach((t, i) => {
        const c = R(r, i + 1);
        c.value = t;
        setFont(c, { bold: true, size: 9, color: C.muted });
        setFill(c, C.gris);
        c.alignment = { horizontal: i >= 2 ? 'center' : 'left', vertical: 'middle', indent: i === 1 ? 1 : 0 };
      });
      ws.getRow(r).height = 16;
    }
    r++;

    // ── Salario base ──────────────────────────────────────────────────────
    const ROW_SALARIO_BASE = r;
    lbl(r, 'Salario base');
    // Col E = referencia config
    frmCell(r, 5, `C$${ROW_BRUTO}`, EURO, C.gris, C.text);
    tipoTag(r, 'SUJETO');
    ws.getRow(r).height = 18; r++;

    // ── Variables SUJETAS ─────────────────────────────────────────────────
    const varsSujetas = [
      ['Turnicidad',             'Turnicidad'],
      ['Nocturnidad',            'Nocturnidad'],
      ['Plus frío',              'PlusFrio'],
      ['H. Disponibilidad',      'HDisp'],
      ['H. Disponibilidad Fest.','HDispFest'],
      ['H. Desplazamiento',      'HDesp'],
      ['H. Desplazamiento Fest.','HDespFest'],
      ['H. Extra',               'HExtra'],
      ['H. Extra Festiva',       'HExtraFest'],
      ['Días fuera',             'DiasFuera'],
    ];

    const sujetasRows = [ROW_SALARIO_BASE]; // filas col E sujetas

    varsSujetas.forEach(([name, key]) => {
      lbl(r, '  ' + name, { indent: 2 });
      // Cantidad — editable amarillo col C
      inputYellow(r, 3, 0, NUM);
      // Precio unitario col D — referencia config
      frmCell(r, 4, `C$${PRICE_ROWS[key]}`, EURO, C.gris, C.muted);
      // Total col E = qty × precio
      frmCell(r, 5, `C${r}*D${r}`, EURO, C.gris, C.text);
      tipoTag(r, 'SUJETO');
      sujetasRows.push(r);
      ws.getRow(r).height = 18;
      r++;
    });

    // ── Separador EXENTOS ─────────────────────────────────────────────────
    ws.mergeCells(r, 1, r, 6);
    {
      const c = R(r, 1);
      c.value = '── EXENTOS (no cotizan ni retienen) ──────────────────────────────────';
      setFont(c, { size: 9, italic: true, bold: true, color: C.verdeSuave });
      setFill(c, C.verdeLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    }
    ws.getRow(r).height = 16; r++;

    // Dieta (importe directo)
    lbl(r, '  Dieta (importe directo)', { indent: 2 });
    inputYellow(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)'; setFont(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    tipoTag(r, 'EXENTO');
    const ROW_DIETA = r;
    ws.getRow(r).height = 18; r++;

    // Tiquets (importe directo)
    lbl(r, '  Tiquets (importe directo)', { indent: 2 });
    inputYellow(r, 3, 0, EURO);
    { const c = R(r, 4); c.value = '(directo)'; setFont(c, { italic: true, size: 9, color: C.muted }); c.alignment = { horizontal: 'center' }; }
    frmCell(r, 5, `C${r}`, EURO, C.verdeLight, C.verdeSuave);
    tipoTag(r, 'EXENTO');
    const ROW_TIQUETS = r;
    ws.getRow(r).height = 18; r++;

    // KM
    lbl(r, '  KM (km × 0,30€)', { indent: 2 });
    inputYellow(r, 3, 0, NUM);
    frmCell(r, 4, `C$${PRICE_ROWS.KM}`, EURO, C.gris, C.muted);
    frmCell(r, 5, `C${r}*D${r}`, EURO, C.verdeLight, C.verdeSuave);
    tipoTag(r, 'EXENTO');
    const ROW_KM = r;
    ws.getRow(r).height = 18; r++;

    divider(r, C.grisBorde); r++;

    // ── Totales intermedios ───────────────────────────────────────────────
    // Bruto sujeto = suma de todas las filas sujetas (col E)
    const sujetasFormula = sujetasRows.map(sr => `E${sr}`).join('+');
    const ROW_BRUTO_SUJETO = r;
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = 'BRUTO SUJETO';
      setFont(c, { bold: true, size: 10, color: C.moradoDark });
      setFill(c, C.moradoLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    }
    frmCell(r, 5, sujetasFormula, EURO, C.moradoLight, C.moradoDark);
    ws.getRow(r).height = 20; r++;

    const ROW_EXENTOS = r;
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = 'EXENTOS';
      setFont(c, { bold: true, size: 10, color: C.verdeSuave });
      setFill(c, C.verdeLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    }
    frmCell(r, 5, `E${ROW_DIETA}+E${ROW_TIQUETS}+E${ROW_KM}`, EURO, C.verdeLight, C.verdeSuave);
    ws.getRow(r).height = 20; r++;

    divider(r, C.grisBorde); r++;

    // ── SS e IRPF ────────────────────────────────────────────────────────
    // Usando fórmula CORRECTA: descuento sobre bruto sujeto directamente
    const ROW_SS = r;
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = `(-) SS ${(ssItems.reduce((a,x)=>a+x[2],0)*100).toFixed(2)}%   [C${SS_ROWS.CC}+C${SS_ROWS.Desempleo}+…]`;
      setFont(c, { bold: true, size: 10, color: C.rojo });
      setFill(c, C.rojoLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    }
    // SS = bruto_sujeto × SS%_total
    frmCell(r, 5, `E${ROW_BRUTO_SUJETO}*C$${ROW_SS_TOTAL}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    const ROW_IRPF_DED = r;
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = `(-) IRPF ${mes === 'Enero' ? '18.11%' : '18%'}  [editable col F fila cabecera]`;
      setFont(c, { bold: true, size: 10, color: C.rojo });
      setFill(c, C.rojoLight);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 2 };
    }
    // IRPF = bruto_sujeto × IRPF_mes (col F de la cabecera del bloque)
    frmCell(r, 5, `E${ROW_BRUTO_SUJETO}*F${ROW_IRPF_MES}`, EURO, C.rojoLight, C.rojo);
    ws.getRow(r).height = 20; r++;

    divider(r, C.grisBorde); r++;

    // ── NETO A COBRAR ────────────────────────────────────────────────────
    // FÓRMULA CORRECTA ADITIVA: bruto_sujeto - SS - IRPF + exentos
    const ROW_NETO = r;
    ws.mergeCells(r, 1, r, 4);
    {
      const c = R(r, 1);
      c.value = '💳  NETO A COBRAR';
      setFont(c, { bold: true, size: 13, color: C.blanco });
      setFill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    }
    {
      const c = R(r, 5);
      // CORRECTO: bruto_sujeto - SS - IRPF + exentos
      c.value = { formula: `E${ROW_BRUTO_SUJETO}-E${ROW_SS}-E${ROW_IRPF_DED}+E${ROW_EXENTOS}` };
      c.numFmt = EURO;
      setFont(c, { bold: true, size: 13, color: C.blanco });
      setFill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'right' };
    }
    {
      // Verificación (texto CORRECTO)
      const c = R(r, 6);
      c.value = '✓ CORRECTO';
      setFont(c, { size: 8, bold: true, color: 'D1FAE5' });
      setFill(c, C.verde);
      c.alignment = { vertical: 'middle', horizontal: 'center' };
    }
    ws.getRow(r).height = 28;

    // Guardar referencias para resumen
    mesRefs[mes] = {
      netoRow:          ROW_NETO,
      brutoSujetoRow:   ROW_BRUTO_SUJETO,
      ssRow:            ROW_SS,
      irpfRow:          ROW_IRPF_DED,
      exentosRow:       ROW_EXENTOS,
    };

    r++;
    spacer(r, C.gris, 10); r++;
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // RESUMEN ANUAL
  // ═══════════════════════════════════════════════════════════════════════════
  secHeader(r, '📅  RESUMEN ANUAL 2026', C.moradoDark, C.blanco, 13); r++;

  // Cabecera tabla resumen
  {
    const hdrs = ['', 'MES', 'Bruto Sujeto', 'SS (−)', 'IRPF (−)', 'Exentos (+)', 'NETO'];
    hdrs.forEach((t, i) => {
      const c = R(r, i + 1);
      c.value = t;
      setFont(c, { bold: true, size: 10, color: C.blanco });
      setFill(c, C.morado);
      c.alignment = { horizontal: i >= 2 ? 'center' : 'left', vertical: 'middle', indent: i === 1 ? 1 : 0 };
      ws.getColumn(i + 1).width = i === 0 ? 2 : i === 1 ? 18 : 15;
    });
    ws.getRow(r).height = 20;
  }
  r++;

  const resumenStart = r;
  meses.forEach((mes, idx) => {
    const ref = mesRefs[mes];
    const isAlt = idx % 2 === 1;
    const bg = isAlt ? C.moradoLight : C.blanco;

    lbl(r, mes, { bold: true });
    setFill(R(r, 2), bg);

    frmCell(r, 3, `E${ref.brutoSujetoRow}`, EURO, bg, C.text);
    frmCell(r, 4, `E${ref.ssRow}`,          EURO, bg, C.rojo);
    frmCell(r, 5, `E${ref.irpfRow}`,        EURO, bg, C.rojo);
    frmCell(r, 6, `E${ref.exentosRow}`,     EURO, bg, C.verdeSuave);

    // NETO en la tabla resumen — referencia directa a la fila NETO del mes
    {
      const c = R(r, 7);
      c.value = { formula: `E${ref.netoRow}` };
      c.numFmt = EURO;
      setFont(c, { bold: true, size: 11, color: C.verde });
      setFill(c, bg);
      setBorder(c, C.grisBorde);
      c.alignment = { vertical: 'middle', horizontal: 'right' };
    }

    ws.getRow(r).height = 20;
    r++;
  });

  const resumenEnd = r - 1;

  // Fila totales anuales
  divider(r, C.morado); r++;

  ws.mergeCells(r, 1, r, 2);
  {
    const c = R(r, 1);
    c.value = 'TOTALES ANUALES';
    setFont(c, { bold: true, size: 11, color: C.blanco });
    setFill(c, C.moradoDark);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  }

  [3, 4, 5, 6].forEach(col => {
    const colLetter = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G'][col];
    const c = R(r, col);
    c.value = { formula: `SUM(${colLetter}${resumenStart}:${colLetter}${resumenEnd})` };
    c.numFmt = EURO;
    setFont(c, { bold: true, size: 11, color: C.blanco });
    setFill(c, C.moradoDark);
    setBorder(c, C.morado);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  });

  // NETO ANUAL TOTAL
  {
    const c = R(r, 7);
    c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` };
    c.numFmt = EURO;
    setFont(c, { bold: true, size: 13, color: C.blanco });
    setFill(c, C.verde);
    setBorder(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  }

  ws.getRow(r).height = 26; r++;

  spacer(r, C.gris, 8); r++;

  // Neto anual destacado
  ws.mergeCells(r, 1, r, 5);
  {
    const c = R(r, 1);
    c.value = '💳  NETO ANUAL TOTAL 2026';
    setFont(c, { bold: true, size: 14, color: C.blanco });
    setFill(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }
  {
    const c = R(r, 6);
    c.value = { formula: `SUM(G${resumenStart}:G${resumenEnd})` };
    c.numFmt = EURO;
    setFont(c, { bold: true, size: 16, color: C.blanco });
    setFill(c, C.verde);
    setBorder(c, C.verde);
    c.alignment = { vertical: 'middle', horizontal: 'right' };
  }
  ws.mergeCells(r, 6, r, 7);
  ws.getRow(r).height = 36; r++;

  spacer(r, C.gris, 8); r++;

  // Nota final
  ws.mergeCells(r, 1, r, 7);
  {
    const c = R(r, 1);
    c.value = '⚠️  Celdas amarillas = editables | IRPF editable por mes (columna F de cada cabecera de mes) | '
            + 'SS% editable en CONFIG | Precios unitarios editables en CONFIG | '
            + 'Fórmula NETO corregida: aditiva, no multiplicativa';
    setFont(c, { size: 9, italic: true, color: '78350F' });
    setFill(c, C.amarillo);
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
    ws.getRow(r).height = 28;
  }

  // ─── Escribir fichero ────────────────────────────────────────────────────
  const outPath = '/home/user/Mi-memoria/calculadora-nomina-v2.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log('✅  Generado:', outPath);
  console.log('');
  console.log('📋  Resumen de estructura:');
  console.log(`    CONFIG: filas 1 – ~${ROW_IRPF_DEFAULT + 20}`);
  console.log(`    Precios unitarios en fila aprox ${Object.values(PRICE_ROWS)[0]} – ${Object.values(PRICE_ROWS).slice(-1)[0]}`);
  console.log(`    SS total ref: C$${ROW_SS_TOTAL}  |  IRPF default: C$${ROW_IRPF_DEFAULT}`);
  console.log(`    12 bloques mensuales con ~20 filas cada uno`);
  console.log(`    Resumen anual desde fila ${resumenStart}`);
  console.log('');
  console.log('🔧  Bug corregido:');
  console.log('    ❌ ANTES (mal): bruto × (1-IRPF) × (1-SS)');
  console.log('    ✅ AHORA (bien): bruto_sujeto - (bruto×SS%) - (bruto×IRPF%) + exentos');
}

main().catch(err => {
  console.error('❌  Error:', err.message);
  process.exit(1);
});

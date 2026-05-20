const ExcelJS = require('exceljs');

async function generarNomina() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Mi Memoria';
  wb.created = new Date();

  // ── Paleta ──────────────────────────────────────────────────────────────
  const MORADO      = '6366F1';
  const MORADO_DARK = '4338CA';
  const MORADO_LIGHT= 'EDE9FE';
  const VERDE       = '059669';
  const VERDE_LIGHT = 'D1FAE5';
  const ROJO        = 'DC2626';
  const ROJO_LIGHT  = 'FEE2E2';
  const GRIS        = 'F8FAFC';
  const GRIS_BORDE  = 'E2E8F0';
  const TEXT        = '0F172A';
  const MUTED       = '64748B';
  const AMARILLO    = 'FFFBEB';
  const AMARILLO_B  = 'F59E0B';

  const ws = wb.addWorksheet('Calculadora Nómina', {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 },
    views: [{ state: 'frozen', ySplit: 3 }]
  });

  // ── Anchos ──────────────────────────────────────────────────────────────
  ws.columns = [
    { key: 'A', width: 2  },
    { key: 'B', width: 32 },
    { key: 'C', width: 16 },
    { key: 'D', width: 16 },
    { key: 'E', width: 2  },
  ];

  // ── Helpers ─────────────────────────────────────────────────────────────
  const euro = { numFmt: '#,##0.00 "€"' };
  const pct  = { numFmt: '0.00"%"' };

  function cell(row, col) { return ws.getCell(row, col); }

  function header(row, text, bg = MORADO, fg = 'FFFFFF', size = 11) {
    ws.mergeCells(row, 1, row, 5);
    const c = cell(row, 1);
    c.value = text;
    c.font  = { bold: true, size, color: { argb: fg }, name: 'Calibri' };
    c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(row).height = 22;
  }

  function label(row, txt, muted = false) {
    const c = cell(row, 2);
    c.value = txt;
    c.font  = { size: 10, color: { argb: muted ? MUTED : TEXT }, name: 'Calibri' };
    c.alignment = { vertical: 'middle', indent: 1 };
  }

  function inputCell(row, value, fmt = euro, bg = 'FFFFFF') {
    const c = cell(row, 3);
    c.value = value;
    c.numFmt = fmt.numFmt;
    c.font   = { size: 10, bold: true, color: { argb: MORADO_DARK }, name: 'Calibri' };
    c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c.border = {
      top:    { style: 'thin', color: { argb: MORADO } },
      bottom: { style: 'thin', color: { argb: MORADO } },
      left:   { style: 'thin', color: { argb: MORADO } },
      right:  { style: 'thin', color: { argb: MORADO } },
    };
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  function formulaCell(row, formula, fmt = euro, bg = GRIS) {
    const c = cell(row, 4);
    c.value  = { formula };
    c.numFmt = fmt.numFmt;
    c.font   = { size: 10, bold: true, color: { argb: TEXT }, name: 'Calibri' };
    c.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    c.border = {
      top:    { style: 'thin', color: { argb: GRIS_BORDE } },
      bottom: { style: 'thin', color: { argb: GRIS_BORDE } },
      left:   { style: 'thin', color: { argb: GRIS_BORDE } },
      right:  { style: 'thin', color: { argb: GRIS_BORDE } },
    };
    c.alignment = { vertical: 'middle', horizontal: 'right' };
    return c;
  }

  function totalRow(row, text, formula, bgTxt, bgVal, fgTxt = 'FFFFFF', fgVal = 'FFFFFF') {
    ws.mergeCells(row, 1, row, 1);
    const lc = cell(row, 2);
    lc.value = text;
    lc.font  = { bold: true, size: 11, color: { argb: fgTxt }, name: 'Calibri' };
    lc.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgTxt } };
    lc.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };

    const vc = cell(row, 4);
    vc.value  = { formula };
    vc.numFmt = euro.numFmt;
    vc.font   = { bold: true, size: 12, color: { argb: fgVal }, name: 'Calibri' };
    vc.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgVal } };
    vc.alignment = { vertical: 'middle', horizontal: 'right' };
    ws.getRow(row).height = 24;
  }

  function rowHeight(row, h = 20) { ws.getRow(row).height = h; }

  function colLabel(row) {
    cell(row, 3).value = 'Dato / mes';
    cell(row, 4).value = 'Calculado';
    [3, 4].forEach(c => {
      const cl = cell(row, c);
      cl.font  = { bold: true, size: 9, color: { argb: MUTED }, name: 'Calibri' };
      cl.alignment = { horizontal: 'right' };
    });
  }

  function spacer(row) {
    ws.mergeCells(row, 1, row, 5);
    cell(row, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRIS } };
    ws.getRow(row).height = 6;
  }

  // ── TÍTULO ───────────────────────────────────────────────────────────────
  ws.mergeCells(1, 1, 2, 5);
  const titulo = cell(1, 1);
  titulo.value = '💰  Calculadora de Nómina';
  titulo.font  = { bold: true, size: 16, color: { argb: 'FFFFFF' }, name: 'Calibri' };
  titulo.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: MORADO } };
  titulo.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 32;

  spacer(3);

  // ── SECCIÓN 1: DATOS DEL EMPLEADO ────────────────────────────────────────
  let r = 4;
  header(r, '👤  Datos del empleado', MORADO_DARK);
  colLabel(r); r++;

  const rows1 = [
    ['Nombre',                         'Tu nombre'],
    ['Categoría profesional',          'Ej: Técnico Grado Medio'],
    ['Convenio colectivo',             'Ej: Convenio Metal'],
    ['Tipo de contrato',               'Indefinido / Temporal'],
    ['Grupo de cotización SS',         '1 (Ing.) – 11 (Peones)'],
  ];
  rows1.forEach(([lbl, eg]) => {
    label(r, lbl);
    const c = cell(r, 3);
    c.value = eg;
    c.font  = { size: 10, italic: true, color: { argb: MUTED }, name: 'Calibri' };
    c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMARILLO } };
    c.border = { top:{style:'thin',color:{argb:AMARILLO_B}}, bottom:{style:'thin',color:{argb:AMARILLO_B}}, left:{style:'thin',color:{argb:AMARILLO_B}}, right:{style:'thin',color:{argb:AMARILLO_B}} };
    c.alignment = { vertical:'middle', horizontal:'center' };
    rowHeight(r); r++;
  });

  spacer(r); r++;

  // ── SECCIÓN 2: DEVENGOS ───────────────────────────────────────────────────
  header(r, '📥  Devengos (ingresos)', MORADO);
  colLabel(r); r++;

  label(r, 'Salario base');                  inputCell(r, 1500.00); rowHeight(r); r++;
  label(r, 'Complemento antigüedad');        inputCell(r, 0);       rowHeight(r); r++;
  label(r, 'Plus de convenio');              inputCell(r, 0);       rowHeight(r); r++;
  label(r, 'Complemento productividad');     inputCell(r, 0);       rowHeight(r); r++;
  label(r, 'Plus transporte');               inputCell(r, 0);       rowHeight(r); r++;
  label(r, 'Dietas / gastos empresa');       inputCell(r, 0);       rowHeight(r); r++;
  label(r, 'Otros complementos');            inputCell(r, 0);       rowHeight(r); r++;

  // Horas extra
  label(r, '── Horas extra ──', true);       rowHeight(r); r++;
  label(r, '  Horas extra ordinarias');      inputCell(r, 0);
  formulaCell(r, `C${r}*C${r+2}`);          rowHeight(r); r++;
  label(r, '  Precio hora extra ordinaria'); inputCell(r, 15.00);   rowHeight(r); r++;
  label(r, '  Horas extra festivas');        inputCell(r, 0);
  formulaCell(r, `C${r}*C${r+2}`);          rowHeight(r); r++;
  label(r, '  Precio hora extra festiva');   inputCell(r, 22.50);   rowHeight(r); r++;

  // Calcular fila de inicio de devengos
  // Devengos empiezan en fila 6 (salario base)
  const devStart  = 6;   // row del salario base
  const devEnd    = r-1; // última fila de devengos

  const totalDevengadoRow = r;
  totalRow(r,
    '  TOTAL DEVENGADO BRUTO / MES',
    `SUM(C${devStart}:C${devEnd})+C${devStart+8}+C${devStart+10}`,
    MORADO, MORADO
  );
  r++;

  spacer(r); r++;

  // ── SECCIÓN 3: PAGAS EXTRA ────────────────────────────────────────────────
  header(r, '🎁  Pagas extra y bonus', '7C3AED');
  colLabel(r); r++;

  label(r, 'Número de pagas extra al año'); inputCell(r, 2);         rowHeight(r); r++;
  label(r, 'Importe paga extra (bruto)');   inputCell(r, 0,euro,'FEF3C7');
  const pagaExtraRef = r;
  cell(r,3).font = {size:10,bold:true,color:{argb:'92400E'},name:'Calibri'};
  cell(r,3).border = {top:{style:'thin',color:{argb:AMARILLO_B}},bottom:{style:'thin',color:{argb:AMARILLO_B}},left:{style:'thin',color:{argb:AMARILLO_B}},right:{style:'thin',color:{argb:AMARILLO_B}}};
  formulaCell(r, `IF(C${r}=0,D${totalDevengadoRow}/12*C${r-1},C${r})`, euro, AMARILLO);
  const nota = cell(r, 4);
  nota.note = { texts: [{ text: 'Si dejas 0, calcula la paga prorrateada automáticamente' }] };
  rowHeight(r); r++;

  label(r, 'Bonus / variable anual');       inputCell(r, 0);
  formulaCell(r, `C${r}/12`);
  rowHeight(r); r++;

  const bonusRow = r - 1;
  const pagaExtraProrrateada = totalDevengadoRow + 3; // row de paga extra

  spacer(r); r++;

  // ── SECCIÓN 4: IRPF Y SEGURIDAD SOCIAL ──────────────────────────────────
  header(r, '📤  Deducciones', ROJO);
  colLabel(r); r++;

  const deducStart = r;

  // IRPF
  label(r, '── IRPF ──', true); rowHeight(r); r++;
  label(r, '  % retención IRPF (mensual)');
  inputCell(r, 15, pct);
  const irpfRow = r;
  formulaCell(r, `D${totalDevengadoRow}*C${r}/100`, euro, ROJO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:ROJO},name:'Calibri'};
  rowHeight(r); r++;

  // SS Trabajador
  label(r, '── Seguridad Social (trabajador) ──', true); rowHeight(r); r++;
  label(r, '  Contingencias comunes (4,70%)');
  cell(r,3).value = 4.70; cell(r,3).numFmt = '0.00"%"';
  cell(r,3).font = {size:10,color:{argb:MUTED},name:'Calibri'};
  cell(r,3).alignment = {horizontal:'right'};
  formulaCell(r, `D${totalDevengadoRow}*C${r}/100`, euro, ROJO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:ROJO},name:'Calibri'};
  const ssCC = r; rowHeight(r); r++;

  label(r, '  Desempleo (1,55%)');
  cell(r,3).value = 1.55; cell(r,3).numFmt = '0.00"%"';
  cell(r,3).font = {size:10,color:{argb:MUTED},name:'Calibri'};
  cell(r,3).alignment = {horizontal:'right'};
  formulaCell(r, `D${totalDevengadoRow}*C${r}/100`, euro, ROJO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:ROJO},name:'Calibri'};
  const ssDesempleo = r; rowHeight(r); r++;

  label(r, '  Formación profesional (0,10%)');
  cell(r,3).value = 0.10; cell(r,3).numFmt = '0.00"%"';
  cell(r,3).font = {size:10,color:{argb:MUTED},name:'Calibri'};
  cell(r,3).alignment = {horizontal:'right'};
  formulaCell(r, `D${totalDevengadoRow}*C${r}/100`, euro, ROJO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:ROJO},name:'Calibri'};
  const ssFP = r; rowHeight(r); r++;

  label(r, '  MEI (0,13%)');
  cell(r,3).value = 0.13; cell(r,3).numFmt = '0.00"%"';
  cell(r,3).font = {size:10,color:{argb:MUTED},name:'Calibri'};
  cell(r,3).alignment = {horizontal:'right'};
  formulaCell(r, `D${totalDevengadoRow}*C${r}/100`, euro, ROJO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:ROJO},name:'Calibri'};
  const ssMEI = r; rowHeight(r); r++;

  const totalDeduccionesRow = r;
  totalRow(r,
    '  TOTAL DEDUCCIONES / MES',
    `D${irpfRow}+D${ssCC}+D${ssDesempleo}+D${ssFP}+D${ssMEI}`,
    ROJO, ROJO
  );
  r++;

  spacer(r); r++;

  // ── SECCIÓN 5: RESUMEN ────────────────────────────────────────────────────
  header(r, '💡  Resumen mensual', VERDE, 'FFFFFF', 12);
  r++;

  const resStart = r;

  label(r, 'Salario bruto mensual');
  formulaCell(r, `D${totalDevengadoRow}`, euro, VERDE_LIGHT);
  cell(r,4).font = {size:11,bold:true,color:{argb:VERDE},name:'Calibri'};
  rowHeight(r); r++;

  label(r, '  − IRPF');
  formulaCell(r, `-D${irpfRow}`, euro, VERDE_LIGHT);
  cell(r,4).font = {size:10,bold:false,color:{argb:ROJO},name:'Calibri'};
  rowHeight(r); r++;

  label(r, '  − Seguridad Social (trabajador)');
  formulaCell(r, `-(D${ssCC}+D${ssDesempleo}+D${ssFP}+D${ssMEI})`, euro, VERDE_LIGHT);
  cell(r,4).font = {size:10,bold:false,color:{argb:ROJO},name:'Calibri'};
  rowHeight(r); r++;

  // NETO MENSUAL
  ws.mergeCells(r, 1, r, 2);
  const netoLbl = cell(r, 1);
  netoLbl.value = '💳  NETO A COBRAR / MES';
  netoLbl.font  = { bold: true, size: 13, color: { argb: 'FFFFFF' }, name: 'Calibri' };
  netoLbl.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE } };
  netoLbl.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  const netoVal = cell(r, 4);
  netoVal.value  = { formula: `D${totalDevengadoRow}-D${totalDeduccionesRow}` };
  netoVal.numFmt = euro.numFmt;
  netoVal.font   = { bold: true, size: 14, color: { argb: 'FFFFFF' }, name: 'Calibri' };
  netoVal.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERDE } };
  netoVal.alignment = { vertical: 'middle', horizontal: 'right' };
  ws.getRow(r).height = 28;
  const netoRow = r; r++;

  spacer(r); r++;

  // Paga extra neta
  label(r, 'Paga extra neta estimada');
  formulaCell(r, `D${pagaExtraRef}*(1-C${irpfRow}/100-(C${ssCC}+C${ssDesempleo}+C${ssFP}+C${ssMEI})/100)`, euro, MORADO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:MORADO_DARK},name:'Calibri'};
  rowHeight(r); r++;

  label(r, 'Bonus / variable neto mensual');
  formulaCell(r, `(C${bonusRow}/12)*(1-C${irpfRow}/100-(C${ssCC}+C${ssDesempleo}+C${ssFP}+C${ssMEI})/100)`, euro, MORADO_LIGHT);
  cell(r,4).font = {size:10,bold:true,color:{argb:MORADO_DARK},name:'Calibri'};
  rowHeight(r); r++;

  spacer(r); r++;

  // Anual
  header(r, '📅  Proyección anual', '0EA5E9', 'FFFFFF');
  r++;

  label(r, 'Bruto anual (12 mensualidades)');
  formulaCell(r, `D${totalDevengadoRow}*12`, euro, 'E0F2FE');
  cell(r,4).font = {size:10,bold:true,color:{argb:'0369A1'},name:'Calibri'};
  rowHeight(r); r++;

  label(r, '+ Pagas extra brutas');
  formulaCell(r, `D${pagaExtraRef}*C${pagaExtraRef-1}`, euro, 'E0F2FE');
  cell(r,4).font = {size:10,bold:true,color:{argb:'0369A1'},name:'Calibri'};
  rowHeight(r); r++;

  label(r, '+ Bonus anual bruto');
  formulaCell(r, `C${bonusRow}`, euro, 'E0F2FE');
  cell(r,4).font = {size:10,bold:true,color:{argb:'0369A1'},name:'Calibri'};
  rowHeight(r); r++;

  const brutoAnualRow = r - 3;
  const pagaAnualRow  = r - 2;
  const bonusAnualRow = r - 1;

  label(r, 'BRUTO ANUAL TOTAL');
  formulaCell(r, `D${brutoAnualRow}+D${pagaAnualRow}+D${bonusAnualRow}`, euro, '0EA5E9');
  cell(r,4).font = {size:11,bold:true,color:{argb:'FFFFFF'},name:'Calibri'};
  cell(r,4).fill = {type:'pattern',pattern:'solid',fgColor:{argb:'0EA5E9'}};
  rowHeight(r,22); r++;

  label(r, 'NETO ANUAL ESTIMADO');
  formulaCell(r, `D${netoRow}*12+D${r-5}*(1-C${irpfRow}/100-(C${ssCC}+C${ssDesempleo}+C${ssFP}+C${ssMEI})/100)+C${bonusRow}*(1-C${irpfRow}/100-(C${ssCC}+C${ssDesempleo}+C${ssFP}+C${ssMEI})/100)`, euro, VERDE);
  cell(r,4).font = {size:11,bold:true,color:{argb:'FFFFFF'},name:'Calibri'};
  cell(r,4).fill = {type:'pattern',pattern:'solid',fgColor:{argb:VERDE}};
  rowHeight(r,22); r++;

  spacer(r); r++;

  // Coste empresa
  header(r, '🏢  Coste empresa (referencia)', '94A3B8', 'FFFFFF');
  r++;
  label(r, '  SS empresa ~30,48% (indefinido)');
  formulaCell(r, `D${totalDevengadoRow}*0.3048`, euro, GRIS);
  rowHeight(r); r++;
  label(r, '  COSTE TOTAL EMPRESA / MES');
  formulaCell(r, `D${totalDevengadoRow}+D${r-1}`, euro, GRIS);
  cell(r,4).font = {size:10,bold:true,color:{argb:TEXT},name:'Calibri'};
  rowHeight(r); r++;

  spacer(r); r++;

  // Nota final
  ws.mergeCells(r, 1, r, 5);
  const nota2 = cell(r, 1);
  nota2.value = '⚠️  Las celdas en amarillo son editables. Los % de SS son los de 2025/2026. El IRPF varía según situación personal — consulta con tu empresa o gestor.';
  nota2.font  = { size: 9, italic: true, color: { argb: '78350F' }, name: 'Calibri' };
  nota2.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBEB' } };
  nota2.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(r).height = 30;

  // Proteger celdas de fórmula (opcional, deja las amarillas editables)
  // ws.protect('', { selectLockedCells: false });

  await wb.xlsx.writeFile('/home/user/Mi-memoria/calculadora-nomina.xlsx');
  console.log('✅ Generado: calculadora-nomina.xlsx');
}

generarNomina().catch(console.error);

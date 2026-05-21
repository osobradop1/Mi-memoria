const ExcelJS = require('/home/user/Mi-memoria/node_modules/exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();

  const C = {
    moradoDark:'4338CA', morado:'6366F1', moradoLight:'EDE9FE',
    verde:'059669', verdeLight:'D1FAE5',
    rojo:'DC2626', rojoLight:'FEE2E2',
    am:'FFFBEB', amB:'F59E0B', amDark:'92400E',
    nar:'FFF7ED', narB:'EA580C', narDark:'7C2D12',
    gris:'F8FAFC', borde:'E2E8F0',
    text:'0F172A', muted:'64748B', blanco:'FFFFFF',
    rs:'B91C1C', vs:'065F46',
    peLight:'FEFCE8', peDark:'713F12',
  };

  const EURO='#,##0.00 "€"', PCT='0.00%', NUM='#,##0.00';
  const MONTHS=['Enero','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const DIAS=[31,28,31,30,31,30,31,31,30,31,30,31];
  const CL = n => 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[n-1];

  // Column indices
  const CB=2, CC=3, CJ=4, CD=15, CT=16; // B=label, C=tag, D..O=months, P=total

  // Main sheet row numbers
  const R = {
    TITLE:1, DIV0:2,
    HEADER:3, IRPF:4, DIV1:5,
    SEC_DEV:6,
    SAL:7, ACT:8, TUR:9, NOC:10, FRI:11, DIS:12, DIF:13,
    DES:14, DESF:15, DIAF:16, HEX:17, HEXF:18, DIV2:19,
    SEC_EX:20,
    DIETA:21, TIQ:22, KM:23, DIV3:24,
    SEC_PE:25,
    PE:26, DIV4:27,
    SEC_TOT:28,
    BS:29, HHEE:30, EX:31, DIV5:32,
    SEC_SS:33,
    BCC:34, BAT:35, SCC:36, SDES:37, SHH:38, SMEI:39, STOT:40, DIV6:41,
    SEC_PE2:42,
    PESS:43, PEIRPF:44, PENET:45, DIV7:46,
    IRPFD:47, NETO:48, NETOT:49, SP1:50,
    SEC7P:51,
    D7P:52, PROP7P:53, TOT7P:54, IRPF7P:55, SP2:56,
    RES7P:57, SUM7P:58, LIM7P:59, REST7P:60, IRPF7PTOT:61,
    SP3:62, NOTE:63,
  };

  // Config sheet row numbers
  const K = {
    BRUTO:5, BDIA:6, PRO:7, ACTD:8,
    SSCC:12, SSDES:13, SSHH:14, SSMEI:15,
    IRPF:18,
    PTUR:22, PNOC:23, PFRI:24, PDIS:25, PDIF:26,
    PDES:27, PDESF:28, PHEX:29, PHEXF:30, PDIAF:31, PKM:32,
  };
  const CF = row => `Configuración!$C$${row}`;

  // ── Helpers ──────────────────────────────────────────────────────────────
  function fillCell(c, argb) { c.fill={type:'pattern',pattern:'solid',fgColor:{argb}}; }
  function brd(c, col=C.borde) { const s={style:'thin',color:{argb:col}}; c.border={top:s,bottom:s,left:s,right:s}; }
  function fnt(c,{b=false,sz=10,col=C.text,it=false}={}) { c.font={bold:b,size:sz,color:{argb:col},italic:it,name:'Calibri'}; }
  function aln(c,{h='right',v='middle',ind=0,wrap=false}={}) { c.alignment={horizontal:h,vertical:v,indent:ind,wrapText:wrap}; }

  function inputC(c, val, fmt=EURO) {
    c.value=val; c.numFmt=fmt;
    fnt(c,{b:true,col:C.moradoDark}); fillCell(c,C.am); brd(c,C.amB); aln(c);
  }
  function frmC(c, formula, fmt=EURO, bg=C.gris, fg=C.text, bold=false) {
    c.value={formula}; c.numFmt=fmt;
    fnt(c,{b:bold,col:fg}); fillCell(c,bg); brd(c); aln(c);
  }
  function lblC(c, text, {b=false,sz=10,col=C.text,it=false,bg=null,h='left',ind=1}={}) {
    c.value=text; fnt(c,{b,sz,col,it}); aln(c,{h,v:'middle',ind});
    if(bg) fillCell(c,bg);
  }

  // Merge A-P and style as section header
  function secRow(ws, row, text, bg, fg=C.blanco, sz=10) {
    ws.mergeCells(row,1,row,CT);
    const c=ws.getCell(row,1); c.value=text;
    fnt(c,{b:true,sz,col:fg}); fillCell(c,bg); aln(c,{h:'left',v:'middle',ind:1});
    ws.getRow(row).height=20;
  }
  function divRow(ws, row, bg=C.borde, h=3) {
    ws.mergeCells(row,1,row,CT); fillCell(ws.getCell(row,1),bg); ws.getRow(row).height=h;
  }
  function sumFrm(row) { return `SUM(${CL(CJ)}${row}:${CL(CD)}${row})`; }

  // ══════════════════════════════════════════════════════════════════════════
  // SHEET 1: Configuración
  // ══════════════════════════════════════════════════════════════════════════
  const wc = wb.addWorksheet('Configuración', {properties:{tabColor:{argb:C.morado}}});
  wc.getColumn(1).width=2; wc.getColumn(2).width=32;
  wc.getColumn(3).width=14; wc.getColumn(4).width=44;

  function cfgSec(row, text) {
    wc.mergeCells(row,1,row,4);
    const c=wc.getCell(row,1); c.value=text;
    fnt(c,{b:true,sz:10,col:C.blanco}); fillCell(c,C.moradoDark); aln(c,{h:'left',v:'middle',ind:1});
    wc.getRow(row).height=20;
  }
  function cfgRow(row, label, value, fmt, note) {
    lblC(wc.getCell(row,2), label, {ind:1});
    const vc=wc.getCell(row,3);
    vc.value=value; vc.numFmt=fmt||EURO;
    fnt(vc,{b:true,col:C.moradoDark}); fillCell(vc,C.am); brd(vc,C.amB); aln(vc);
    if(note){ const nc=wc.getCell(row,4); nc.value=note; fnt(nc,{sz:9,it:true,col:C.muted}); aln(nc,{h:'left',ind:1}); }
    wc.getRow(row).height=20;
  }

  wc.mergeCells(1,1,1,4);
  { const c=wc.getCell(1,1); c.value='⚙️  CONFIGURACIÓN — Nómina 2026';
    fnt(c,{b:true,sz:14,col:C.blanco}); fillCell(c,C.moradoDark); aln(c,{h:'center',v:'middle'}); wc.getRow(1).height=32; }
  wc.mergeCells(2,1,2,4); fillCell(wc.getCell(2,1),C.gris); wc.getRow(2).height=6;

  cfgSec(3, '💰  Salario y bases');
  [2,3,4].forEach(c=>{ const cc=wc.getCell(4,c); cc.value=['','Parámetro','Valor','Nota'][c];
    fnt(cc,{b:true,sz:9,col:C.muted}); });
  wc.getRow(4).height=14;
  cfgRow(K.BRUTO,   'Salario Convenio mensual',       2063.39, EURO, '← salario fijo mensual');
  cfgRow(K.BDIA,    'Bruto por día (días fuera)',      64.208,  EURO, '← precio diario al estar fuera');
  cfgRow(K.PRO,     'Prorrata Pagas Extra (base SS)',  419.61,  EURO, '← aparece en la base CC de tu nómina (Base CC 3.207,47 = sujeto CC + 419,61). La paga extra se cobra por separado en jun/dic, pero este importe suma a la base de cotización cada mes.');
  cfgRow(K.ACTD,    'Plus Actividad (por defecto)',    18.04,   EURO, '← valor que aparece normalmente; editable mes a mes en la hoja principal');
  wc.getRow(K.PRO).height=28;

  wc.mergeCells(9,1,9,4); fillCell(wc.getCell(9,1),C.gris); wc.getRow(9).height=6;
  cfgSec(10, '🔴  SS — Tasas trabajador');
  [2,3,4].forEach(c=>{ const cc=wc.getCell(11,c); cc.value=['','Concepto','Tasa','Base de cálculo'][c];
    fnt(cc,{b:true,sz:9,col:C.muted}); }); wc.getRow(11).height=14;
  cfgRow(K.SSCC,  'CC Contingencias Comunes', 0.047,  PCT, 'Base CC = sujeto sin HHEE + prorrata');
  cfgRow(K.SSDES, 'Desempleo + FP',           0.0165, PCT, 'Base AT = Base CC + HHEE');
  cfgRow(K.SSHH,  'HH.EE. Normales',          0.047,  PCT, 'Solo total horas extra');
  cfgRow(K.SSMEI, 'MEI',                       0.0013, PCT, 'Base AT (tasa real ≈ 0.1307%)');

  wc.mergeCells(16,1,16,4); fillCell(wc.getCell(16,1),C.gris); wc.getRow(16).height=6;
  cfgSec(17, '📊  IRPF');
  cfgRow(K.IRPF,  'IRPF por defecto (%)',      0.1815, PCT, '← editable mes a mes en la hoja principal (fila IRPF %)');

  wc.mergeCells(19,1,19,4); fillCell(wc.getCell(19,1),C.gris); wc.getRow(19).height=6;
  cfgSec(20, '💶  Precios unitarios de variables (€/unidad o €/hora)');
  [2,3,4].forEach(c=>{ const cc=wc.getCell(21,c); cc.value=['','Concepto','€/unidad','Tipo SS'][c];
    fnt(cc,{b:true,sz:9,col:C.muted}); }); wc.getRow(21).height=14;
  [
    [K.PTUR,  'Turnicidad',             3.65,  'Base CC'],
    [K.PNOC,  'Nocturnidad',            10.94, 'Base CC'],
    [K.PFRI,  'Plus frío / baja temp.', 16.32, 'Base CC'],
    [K.PDIS,  'H. Disponibilidad',      1.35,  'Base CC'],
    [K.PDIF,  'H. Disp. Festiva',       2.17,  'Base CC'],
    [K.PDES,  'H. Desplazamiento',      10.56, 'Base CC + Art.7p'],
    [K.PDESF, 'H. Desp. Festivo',       12.68, 'Base CC + Art.7p'],
    [K.PHEX,  'H. Extra Normal',        19.77, 'HHEE + Art.7p'],
    [K.PHEXF, 'H. Extra Festiva',       19.77, 'HHEE'],
    [K.PDIAF, 'Días fuera (precio/día ref.)', null,  'Ref. informativa ← = C6 (no suma a devengos; solo para referencia del valor diario)'],
    [K.PKM,   'KM',                     0.30,  'Exento'],
  ].forEach(([row, name, val, tipo]) => {
    lblC(wc.getCell(row,2), name, {ind:1});
    const vc=wc.getCell(row,3);
    if(val===null){ vc.value={formula:`C${K.BDIA}`}; }
    else { vc.value=val; }
    vc.numFmt=EURO; fnt(vc,{b:true,col:C.moradoDark}); fillCell(vc,C.am); brd(vc,C.amB); aln(vc);
    const tc=wc.getCell(row,4); tc.value=tipo; fnt(tc,{sz:9,col:C.muted}); aln(tc,{h:'left',ind:1});
    wc.getRow(row).height=18;
  });

  wc.mergeCells(33,1,33,4); fillCell(wc.getCell(33,1),C.gris); wc.getRow(33).height=6;
  wc.mergeCells(34,1,34,4);
  { const c=wc.getCell(34,1);
    c.value='🌍 Art.7p: Para calcular los importes de Turnicidad/Nocturnidad/Plus frío/etc. en €: multiplica el precio unitario × horas/unidades del mes y ponlo directamente en la celda amarilla de la hoja principal.';
    fnt(c,{sz:9,it:true,col:C.amDark}); fillCell(c,C.am); aln(c,{h:'left',v:'middle',ind:1,wrap:true});
    wc.getRow(34).height=36; }

  // ══════════════════════════════════════════════════════════════════════════
  // SHEET 2: Nómina 2026 (horizontal grid)
  // ══════════════════════════════════════════════════════════════════════════
  const ws = wb.addWorksheet('Nómina 2026', {
    pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1},
    properties:{tabColor:{argb:C.moradoDark}},
  });

  ws.getColumn(1).width=2; ws.getColumn(CB).width=26; ws.getColumn(CC).width=10;
  for(let c=CJ;c<=CD;c++) ws.getColumn(c).width=10;
  ws.getColumn(CT).width=13;

  // ── TITLE ──────────────────────────────────────────────────────────────
  ws.mergeCells(R.TITLE,1,R.TITLE,CT);
  { const c=ws.getCell(R.TITLE,1); c.value='💰  CALCULADORA NÓMINA 2026';
    fnt(c,{b:true,sz:16,col:C.blanco}); fillCell(c,C.moradoDark); aln(c,{h:'center',v:'middle'});
    ws.getRow(R.TITLE).height=36; }
  divRow(ws, R.DIV0, C.moradoLight, 4);

  // ── MONTH HEADERS ──────────────────────────────────────────────────────
  ws.getRow(R.HEADER).height=22;
  ws.getRow(R.IRPF).height=20;
  // col B: label
  { const c=ws.getCell(R.HEADER,CB); c.value='MES';
    fnt(c,{b:true,sz:10,col:C.blanco}); fillCell(c,C.moradoDark); aln(c,{h:'center',v:'middle'}); }
  { const c=ws.getCell(R.IRPF,CB); c.value='IRPF %';
    fnt(c,{b:true,sz:10,col:C.moradoDark}); fillCell(c,C.am); aln(c,{h:'left',v:'middle',ind:1}); }
  // col C: type col header
  { const c=ws.getCell(R.HEADER,CC); c.value='Tipo';
    fnt(c,{b:true,sz:8,col:C.muted}); fillCell(c,C.gris); aln(c,{h:'center',v:'middle'}); }
  { const c=ws.getCell(R.IRPF,CC); c.value='← editable';
    fnt(c,{sz:8,it:true,col:C.muted}); aln(c,{h:'center',v:'middle'}); }

  for(let col=CJ; col<=CD; col++) {
    const idx=col-CJ;
    const isPE = idx===5 || idx===11; // Jun, Dic
    const mName = MONTHS[idx] + (isPE?' ⭐':'');
    const hc=ws.getCell(R.HEADER,col); hc.value=mName;
    fnt(hc,{b:true,sz:9,col:C.blanco}); fillCell(hc, isPE ? C.peDark : C.moradoDark); aln(hc,{h:'center',v:'middle'});
    // IRPF row
    const ic=ws.getCell(R.IRPF,col);
    // Enero=18.11%, Febrero=18.15% (datos reales nómina), resto=config
    ic.value = (idx===0) ? 0.1811 : (idx===1) ? 0.1815 : {formula:`${CF(K.IRPF)}`};
    ic.numFmt=PCT; fnt(ic,{b:true,col:C.moradoDark}); fillCell(ic,C.am); brd(ic,C.amB); aln(ic);
  }
  // Total col headers
  { const c=ws.getCell(R.HEADER,CT); c.value='TOTAL AÑO';
    fnt(c,{b:true,sz:9,col:C.blanco}); fillCell(c,C.morado); aln(c,{h:'center',v:'middle'}); }
  { const c=ws.getCell(R.IRPF,CT); c.value='↑ edita cada mes';
    fnt(c,{sz:8,it:true,col:C.muted}); aln(c,{h:'center',v:'middle'}); }

  divRow(ws, R.DIV1, C.borde, 3);

  // ── Helper: fill one data row ────────────────────────────────────────────
  // type: 'input'|'formula'|'cfgref'
  // opts: { fmt, bg, fg, bold, tag, totalFrm }
  function dataRow(row, label, tag, type, monthFn, opts={}) {
    const {fmt=EURO, bg=C.gris, fg=C.text, bold=false, tagColor=C.muted, totalFrm=null, totalFmt=EURO} = opts;
    ws.getRow(row).height = opts.h || 18;

    // Label
    const lc=ws.getCell(row,CB); lc.value=label;
    fnt(lc,{sz:9,col:fg,b:bold}); if(opts.lbg) fillCell(lc,opts.lbg); else fillCell(lc,bg);
    aln(lc,{h:'left',v:'middle',ind:1});

    // Tag col
    if(tag) {
      const tc=ws.getCell(row,CC); tc.value=tag;
      fnt(tc,{sz:8,col:tagColor}); fillCell(tc,bg); aln(tc,{h:'center',v:'middle'});
    }

    // Month cols
    for(let col=CJ; col<=CD; col++) {
      const idx=col-CJ;
      const c=ws.getCell(row,col);
      if(type==='input') {
        const def = typeof monthFn==='function' ? monthFn(idx,col) : monthFn;
        inputC(c, def, fmt);
      } else if(type==='formula') {
        const frm = typeof monthFn==='function' ? monthFn(idx,col) : monthFn;
        frmC(c, frm, fmt, bg, fg, bold);
      }
    }

    // Total col — siempre en € (para qty rows totalFrm ya convierte qty×precio a €)
    const tc=ws.getCell(row,CT);
    const tfrm = totalFrm || (type==='input'||type==='formula' ? sumFrm(row) : null);
    if(tfrm) {
      tc.value={formula:tfrm}; tc.numFmt=totalFmt;
      fnt(tc,{b:bold,col:bold?fg:C.muted}); fillCell(tc, bold?bg:C.gris); brd(tc); aln(tc);
    }
  }

  // Shorthand for column letter of each month
  const cl = col => CL(col);

  // ── DEVENGOS ────────────────────────────────────────────────────────────
  secRow(ws, R.SEC_DEV, '📋  DEVENGOS — Celdas amarillas editables. Variables (horas/días/km): introduce CANTIDADES. Salario, Actividad, Dieta y Paga Extra: introduce €.', C.morado);

  dataRow(R.SAL,  'Salario Convenio',          'SUJETO-CC',  'formula',
    (i,col) => CF(K.BRUTO), {fg:C.text, bg:C.gris});
  dataRow(R.ACT,  'Plus Actividad €',          'SUJETO-CC',  'input',
    (i) => i===0 ? {formula:CF(K.ACTD)} : {formula:CF(K.ACTD)});
  // Variables → el usuario introduce CANTIDAD (horas/días/km), la hoja calcula € = cantidad × precio config
  // Datos pre-cargados desde Calculo_sueldo_mes.xlsx (Ene–May; Jun–Dic pendientes de rellenar)
  const TUR_D  = [0,    0,     0,     0,    8,    0,0,0,0,0,0,0];
  const NOC_D  = [0,    0,     0,     0,    3,    0,0,0,0,0,0,0];
  const FRI_D  = [6,    16,    12,    0,    14,   0,0,0,0,0,0,0];
  const DES_D  = [9.23, 22.48, 24.2,  8.48, 27.98,0,0,0,0,0,0,0];
  const DESF_D = [1,    17,    1,     1.5,  5.5,  0,0,0,0,0,0,0];
  const DIAF_D = [0,    6,     0,     0,    0,    0,0,0,0,0,0,0];
  const HEX_D  = [2.5,  0,     21.67, 3,    44.5, 0,0,0,0,0,0,0];
  const HEXF_D = [8,    24,    6.67,  10,   44.5, 0,0,0,0,0,0,0];
  const DIETA_D= [286.56,764.16,692.52,286.56,1098.48,0,0,0,0,0,0,0];

  dataRow(R.TUR,  'Turnicidad (horas)',            'SUJETO-CC',  'input', (i)=>TUR_D[i]||0,  {fmt:NUM, totalFrm:`${sumFrm(R.TUR)}*${CF(K.PTUR)}`});
  dataRow(R.NOC,  'Nocturnidad (horas)',           'SUJETO-CC',  'input', (i)=>NOC_D[i]||0,  {fmt:NUM, totalFrm:`${sumFrm(R.NOC)}*${CF(K.PNOC)}`});
  dataRow(R.FRI,  'Plus frío / baja temp. (días)', 'SUJETO-CC',  'input', (i)=>FRI_D[i]||0,  {fmt:NUM, totalFrm:`${sumFrm(R.FRI)}*${CF(K.PFRI)}`});
  dataRow(R.DIS,  'H. Disponibilidad (horas)',     'SUJETO-CC',  'input', 0,                 {fmt:NUM, totalFrm:`${sumFrm(R.DIS)}*${CF(K.PDIS)}`});
  dataRow(R.DIF,  'H. Disp. Festiva (horas)',      'SUJETO-CC',  'input', 0,                 {fmt:NUM, totalFrm:`${sumFrm(R.DIF)}*${CF(K.PDIF)}`});
  dataRow(R.DES,  'H. Desplazamiento (horas)',     'CC + 7p',    'input', (i)=>DES_D[i]||0,  {fmt:NUM, tagColor:C.narDark, totalFrm:`${sumFrm(R.DES)}*${CF(K.PDES)}`});
  dataRow(R.DESF, 'H. Desp. Festivo (horas)',      'CC + 7p',    'input', (i)=>DESF_D[i]||0, {fmt:NUM, tagColor:C.narDark, totalFrm:`${sumFrm(R.DESF)}*${CF(K.PDESF)}`});
  dataRow(R.DIAF, 'Días fuera España (días)',       'contador 7p','input', (i)=>DIAF_D[i]||0, {fmt:NUM, tagColor:C.narDark, totalFrm:sumFrm(R.DIAF), totalFmt:NUM});
  dataRow(R.HEX,  'H. Extra Normal (horas)',       'HHEE + 7p',  'input', (i)=>HEX_D[i]||0,  {fmt:NUM, tagColor:C.morado,  totalFrm:`${sumFrm(R.HEX)}*${CF(K.PHEX)}`});
  dataRow(R.HEXF, 'H. Extra Festiva (horas)',      'HHEE',       'input', (i)=>HEXF_D[i]||0,  {fmt:NUM, tagColor:C.morado,  totalFrm:`${sumFrm(R.HEXF)}*${CF(K.PHEXF)}`});
  divRow(ws, R.DIV2, C.borde, 2);

  // ── EXENTOS ─────────────────────────────────────────────────────────────
  secRow(ws, R.SEC_EX, '🟢  EXENTOS — No cotizan SS ni retienen IRPF', C.verde);
  // Feb: Dieta 764.16€ (exento, importe directo). KM: introduce km, calcula € = km × 0,30
  dataRow(R.DIETA, 'Dieta € (importe directo)',  'EXENTO', 'input', (i)=>DIETA_D[i]||0, {bg:C.verdeLight, fg:C.vs, tagColor:C.vs});
  dataRow(R.TIQ,   'Tiquets € (importe directo)','EXENTO', 'input', 0,                  {bg:C.verdeLight, fg:C.vs, tagColor:C.vs});
  dataRow(R.KM,    'KM (kilómetros)',             'EXENTO', 'input', 0,                  {fmt:NUM, bg:C.verdeLight, fg:C.vs, tagColor:C.vs, totalFrm:`${sumFrm(R.KM)}*${CF(K.PKM)}`});
  divRow(ws, R.DIV3, C.borde, 2);

  // ── PAGA EXTRA ──────────────────────────────────────────────────────────
  secRow(ws, R.SEC_PE, '⭐  PAGA EXTRA — Normalmente junio y diciembre. Pon el importe bruto en esos meses.', C.peDark);
  dataRow(R.PE, 'Paga Extra € (importe bruto)', 'PAGA EXTRA', 'input', 0,
    {bg:C.peLight, fg:C.peDark, tagColor:C.peDark, bold:false, h:20});
  divRow(ws, R.DIV4, C.borde, 4);

  // ── TOTALES ─────────────────────────────────────────────────────────────
  secRow(ws, R.SEC_TOT, '🔢  TOTALES CALCULADOS  ·  Bruto Sujeto = devengos del mes sin paga extra  ·  Paga extra ⭐ aparece en su sección propia y se suma en el RESUMEN ANUAL al pie de la hoja', C.moradoDark);
  // BS, HHEE y Exentos: las qty rows se multiplican por su precio de config
  dataRow(R.BS,   'Bruto Sujeto (base IRPF)',    'BASE IRPF',  'formula',
    (i,col) => `${cl(col)}${R.SAL}+${cl(col)}${R.ACT}`
      +`+${cl(col)}${R.TUR}*${CF(K.PTUR)}+${cl(col)}${R.NOC}*${CF(K.PNOC)}`
      +`+${cl(col)}${R.FRI}*${CF(K.PFRI)}+${cl(col)}${R.DIS}*${CF(K.PDIS)}`
      +`+${cl(col)}${R.DIF}*${CF(K.PDIF)}+${cl(col)}${R.DES}*${CF(K.PDES)}`
      +`+${cl(col)}${R.DESF}*${CF(K.PDESF)}`
      +`+${cl(col)}${R.HEX}*${CF(K.PHEX)}+${cl(col)}${R.HEXF}*${CF(K.PHEXF)}`,
    {bg:C.moradoLight, fg:C.moradoDark, bold:true});
  dataRow(R.HHEE, 'Total Horas Extra (HHEE)',    'SS sep.',    'formula',
    (i,col) => `${cl(col)}${R.HEX}*${CF(K.PHEX)}+${cl(col)}${R.HEXF}*${CF(K.PHEXF)}`,
    {fg:C.morado});
  dataRow(R.EX,   'Total Exentos',               'EXENTO',    'formula',
    (i,col) => `${cl(col)}${R.DIETA}+${cl(col)}${R.TIQ}+${cl(col)}${R.KM}*${CF(K.PKM)}`,
    {bg:C.verdeLight, fg:C.vs, tagColor:C.vs});
  divRow(ws, R.DIV5, C.borde, 3);

  // ── SS ──────────────────────────────────────────────────────────────────
  secRow(ws, R.SEC_SS, '🔴  COTIZACIÓN SEGURIDAD SOCIAL', C.rojo);
  dataRow(R.BCC,  'Base CC = sujeto CC + prorrata', '',       'formula',
    (i,col) => `${cl(col)}${R.BS}-${cl(col)}${R.HHEE}+${CF(K.PRO)}`,
    {bg:C.rojoLight, fg:C.rs});
  dataRow(R.BAT,  'Base AT = Base CC + HHEE',       '',       'formula',
    (i,col) => `${cl(col)}${R.BCC}+${cl(col)}${R.HHEE}`,
    {bg:C.rojoLight, fg:C.rs});
  dataRow(R.SCC,  '(-) SS CC 4.70%',               '4.70%',  'formula',
    (i,col) => `${cl(col)}${R.BCC}*${CF(K.SSCC)}`, {bg:C.rojoLight, fg:C.rojo, tagColor:C.rs});
  dataRow(R.SDES, '(-) SS Desempleo+FP 1.65%',     '1.65%',  'formula',
    (i,col) => `${cl(col)}${R.BAT}*${CF(K.SSDES)}`, {bg:C.rojoLight, fg:C.rojo, tagColor:C.rs});
  dataRow(R.SHH,  '(-) SS HH.EE. 4.70%',           '4.70%',  'formula',
    (i,col) => `${cl(col)}${R.HHEE}*${CF(K.SSHH)}`, {bg:C.rojoLight, fg:C.rojo, tagColor:C.rs});
  dataRow(R.SMEI, '(-) SS MEI 0.13%',              '0.13%',  'formula',
    (i,col) => `${cl(col)}${R.BAT}*${CF(K.SSMEI)}`, {bg:C.rojoLight, fg:C.rojo, tagColor:C.rs});
  dataRow(R.STOT, 'TOTAL SS',                       '',       'formula',
    (i,col) => `${cl(col)}${R.SCC}+${cl(col)}${R.SDES}+${cl(col)}${R.SHH}+${cl(col)}${R.SMEI}`,
    {bg:C.rojo, fg:C.blanco, bold:true, h:22});
  divRow(ws, R.DIV6, C.borde, 3);

  // ── SS + IRPF PAGA EXTRA ─────────────────────────────────────────────────
  secRow(ws, R.SEC_PE2, '⭐  SS + IRPF — PAGA EXTRA', C.peDark);
  dataRow(R.PESS,  'SS Paga Extra',    '', 'formula',
    (i,col) => `${cl(col)}${R.PE}*(${CF(K.SSCC)}+${CF(K.SSDES)}+${CF(K.SSMEI)})`,
    {bg:C.peLight, fg:C.peDark});
  dataRow(R.PEIRPF,'IRPF Paga Extra', '', 'formula',
    (i,col) => `${cl(col)}${R.PE}*${cl(col)}${R.IRPF}`,
    {bg:C.peLight, fg:C.peDark});
  dataRow(R.PENET, 'NETO Paga Extra', '', 'formula',
    (i,col) => `${cl(col)}${R.PE}-${cl(col)}${R.PESS}-${cl(col)}${R.PEIRPF}`,
    {bg:C.peLight, fg:C.peDark, bold:true, h:20});
  divRow(ws, R.DIV7, C.borde, 4);

  // ── IRPF + NETO ──────────────────────────────────────────────────────────
  dataRow(R.IRPFD, '(-) IRPF retenido mensual', '', 'formula',
    (i,col) => `${cl(col)}${R.BS}*${cl(col)}${R.IRPF}`,
    {bg:C.rojoLight, fg:C.rojo, h:20});
  dataRow(R.NETO,  'NETO MENSUAL', '💳', 'formula',
    (i,col) => `${cl(col)}${R.BS}+${cl(col)}${R.EX}-${cl(col)}${R.STOT}-${cl(col)}${R.IRPFD}`,
    {bg:C.verde, fg:C.blanco, bold:true, h:26, tagColor:C.blanco});
  dataRow(R.NETOT, 'NETO TOTAL (mensual + paga extra)', '💳+⭐', 'formula',
    (i,col) => `${cl(col)}${R.NETO}+${cl(col)}${R.PENET}`,
    {bg:C.verde, fg:C.blanco, bold:true, h:28, tagColor:C.blanco});

  divRow(ws, R.SP1, C.gris, 8);

  // ── ART. 7p ─────────────────────────────────────────────────────────────
  secRow(ws, R.SEC7P, '🌍  ARTÍCULO 7p — Renta trabajos en el extranjero (potencialmente exenta de IRPF hasta 60.100€/año)', C.narB);
  // D7P auto desde DIAF (días fuera): no hace falta introducir dos veces
  dataRow(R.D7P,   'Días fuera España (= fila Días fuera)',  'auto 7p', 'formula',
    (i,col) => `${cl(col)}${R.DIAF}`,
    {fmt:NUM, bg:C.nar, fg:C.narDark, tagColor:C.narDark});
  dataRow(R.PROP7P,'Proporción salario días fuera',  '7p', 'formula',
    (i,col) => `(${CF(K.BRUTO)}/${DIAS[i]})*${cl(col)}${R.D7P}`,
    {bg:C.nar, fg:C.narDark, tagColor:C.narDark});
  // TOT7P usa qty×precio para DES, DESF, DIAF
  dataRow(R.TOT7P, 'Renta 7p (salario prop. días fuera + desplazamiento)', '7p TOTAL', 'formula',
    (i,col) => `${cl(col)}${R.PROP7P}+${cl(col)}${R.DES}*${CF(K.PDES)}+${cl(col)}${R.DESF}*${CF(K.PDESF)}`,
    {bg:C.nar, fg:C.narDark, bold:true, h:20, tagColor:C.narDark});
  dataRow(R.IRPF7P,'IRPF retenido sobre renta 7p (a reclamar)', '→ Renta', 'formula',
    (i,col) => `${cl(col)}${R.TOT7P}*${cl(col)}${R.IRPF}`,
    {bg:C.narB, fg:C.blanco, bold:true, h:22, tagColor:C.blanco});

  divRow(ws, R.SP2, C.gris, 8);

  // ── RESUMEN ANUAL ART. 7p ────────────────────────────────────────────────
  secRow(ws, R.RES7P, '📊  RESUMEN ANUAL ARTÍCULO 7p', C.narB);

  function sumBlock(row, label, formula, bg, fg, bold=false, h=22) {
    ws.mergeCells(row,1,row,CT-3);
    const lc=ws.getCell(row,1); lc.value=label;
    fnt(lc,{b:bold,sz:10,col:fg}); fillCell(lc,bg); aln(lc,{h:'left',v:'middle',ind:2});
    ws.mergeCells(row,CT-2,row,CT);
    const vc=ws.getCell(row,CT-2);
    vc.value={formula}; vc.numFmt=EURO; fnt(vc,{b:bold,sz:bold?13:11,col:fg});
    fillCell(vc,bg); brd(vc); aln(vc,{h:'right',v:'middle'});
    ws.getRow(row).height=h;
  }

  sumBlock(R.SUM7P,     'Renta exenta acumulada anual (Art. 7p)',       sumFrm(R.TOT7P),   C.nar,    C.narDark);
  sumBlock(R.LIM7P,     'Límite legal Art. 7p',                         '60100',           C.nar,    C.narDark);
  { ws.getCell(R.LIM7P, CT-2).value=60100; ws.getCell(R.LIM7P,CT-2).numFmt=EURO; }
  sumBlock(R.REST7P,    'Renta exenta restante hasta límite',           `60100-${sumFrm(R.TOT7P)}`, C.nar, C.narDark);
  sumBlock(R.IRPF7PTOT, 'IRPF retenido total a recuperar en la Renta', sumFrm(R.IRPF7P),  C.narB,   C.blanco, true, 30);

  divRow(ws, R.SP3, C.gris, 8);

  // ── NOTA FINAL ───────────────────────────────────────────────────────────
  ws.mergeCells(R.NOTE,1,R.NOTE,CT);
  { const c=ws.getCell(R.NOTE,1);
    c.value='⚠️  Celdas amarillas = editables  |  IRPF editable por mes (fila "IRPF %")  |  '
           +'Días fuera España = introduce NÚMERO DE DÍAS (solo cuenta para Art.7p; NO suma al bruto mensual, la paga por estar fuera va en Dieta/Desplazamiento)  |  '
           +'Paga extra ⭐ = ponla en junio/diciembre; el "Bruto Sujeto" del mes NO la incluye pero el BRUTO ANUAL del resumen inferior SÍ la suma  |  '
           +'Art.7p: requiere trabajo físico fuera de España; consulta con gestor';
    fnt(c,{sz:9,it:true,col:C.amDark}); fillCell(c,C.am); aln(c,{h:'left',v:'middle',ind:1,wrap:true});
    ws.getRow(R.NOTE).height=32; }

  // ── RESUMEN BRUTO / NETO ANUAL ───────────────────────────────────────────
  // 4 cifras: bruto total, bruto fijo, neto total, neto fijo
  const RS = R.NOTE + 2;

  function sumRow(row, label, formula, bg, fg, bold=false) {
    ws.mergeCells(row, 1, row, CT-1);
    const lc=ws.getCell(row,1); lc.value=label;
    fnt(lc,{b:bold,sz:10,col:fg}); fillCell(lc,bg); aln(lc,{h:'left',v:'middle',ind:2});
    const vc=ws.getCell(row,CT); vc.value={formula}; vc.numFmt=EURO;
    fnt(vc,{b:bold,sz:bold?14:11,col:fg}); fillCell(vc,bg); brd(vc); aln(vc,{h:'right',v:'middle'});
    ws.getRow(row).height = bold ? 30 : 22;
  }

  divRow(ws, RS-1, C.gris, 10);

  ws.mergeCells(RS,1,RS,CT);
  { const c=ws.getCell(RS,1);
    c.value='📊  RESUMEN ANUAL — Bruto y Neto total (con variables) vs Solo salario fijo (sin variables)';
    fnt(c,{b:true,sz:11,col:C.blanco}); fillCell(c,C.moradoDark); aln(c,{h:'left',v:'middle',ind:1});
    ws.getRow(RS).height=24; }

  // Bruto CON variables = todo lo devengado en el año (salario + variables + exentos + pagas extra)
  const bCV = `${sumFrm(R.BS)}+${sumFrm(R.EX)}+${sumFrm(R.PE)}`;
  // Bruto SIN variables = SOLO salario convenio × 12 (sin pagas extra, sin extras)
  const bSV = `${CF(K.BRUTO)}*12`;
  // Neto CON variables = suma de NETO TOTAL de todos los meses (incluye pagas extra)
  const nCV = sumFrm(R.NETOT);
  // Neto SIN variables = solo salario fijo neto (sin pagas extra, sin variables)
  // SS_fijo: sin HHEE → Base CC = Base AT = salario×12 + prorrata×12
  const ss12 = `(${CF(K.BRUTO)}*12+${CF(K.PRO)}*12)*(${CF(K.SSCC)}+${CF(K.SSDES)}+${CF(K.SSMEI)})`;
  const nSV  = `${CF(K.BRUTO)}*12-${ss12}-${CF(K.BRUTO)}*12*${CF(K.IRPF)}`;

  sumRow(RS+1, '💰  BRUTO ANUAL con variables  (salario + variables + exentos + pagas extra)', bCV, C.moradoLight, C.moradoDark, true);
  sumRow(RS+2, '     Bruto anual sin variables  (solo salario convenio × 12)',                 bSV, C.gris,        C.muted);
  sumRow(RS+3, '     Diferencia por variables',  `(${bCV})-(${bSV})`,                              C.gris,        C.text);

  divRow(ws, RS+4, C.borde, 3);

  sumRow(RS+5, '💳  NETO ANUAL con variables  (líquido total: salario + variables + pagas extra)', nCV, C.verdeLight, C.vs, true);
  sumRow(RS+6, '     Neto anual sin variables  (solo salario fijo neto × 12)',                     nSV, C.gris,       C.muted);
  sumRow(RS+7, '     Diferencia por variables',  `(${nCV})-(${nSV})`,                                  C.gris,       C.text);

  divRow(ws, RS+8, C.gris, 8);

  // Make Nómina 2026 the active sheet
  wb.views = [{ activeTab: 1 }];

  const outPath='/home/user/Mi-memoria/calculadora-nomina-v3.xlsx';
  await wb.xlsx.writeFile(outPath);
  console.log('✅ Generado:', outPath);
  console.log('   Layout: horizontal (meses en columnas D-O)');
  console.log('   Filas totales: ~63 (vs ~720 del layout vertical)');
  console.log('   Hojas: Nómina 2026 + Configuración');
  console.log('   Novedades: paga extra ⭐ (jun/dic), Art.7p, SS multi-base');
}

main().catch(e => { console.error('❌', e.message); process.exit(1); });

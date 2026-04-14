// ============================================================
//  IV Care Report — Google Apps Script
//  ACWA · Ar Rass 2
//  → Summary sheet + individual sheet per module
//  → Each module sheet: full data table + IV curve chart
// ============================================================
//  SETUP:
//  1. New Google Sheet → Extensions → Apps Script → paste → Save
//  2. Deploy → New Deployment → Web App
//     Execute as: Me  |  Who has access: Anyone
//  3. Copy Web App URL → paste into PWA ⚙️ Settings
// ============================================================

const C = {
  darkBlue:  '#1F4E79', midBlue:   '#2E75B6',
  greenHdr:  '#375623', greenSub:  '#4E6B3A',
  brownHdr:  '#7B3F00', brownSub:  '#8B4513',
  purpleHdr: '#4A235A', purpleSub: '#6C3483',
  paramBlue: '#BDD7EE', dataLight: '#DEEAF1',
  deltaLav:  '#F5E6FA', assessBg:  '#EBF3FB',
  rowGreen:  '#E2EFDA', rowRed:    '#FCE8E6',
  rowAmber:  '#FFF3CD', rowBlue:   '#E8F0FE',
  white:     '#FFFFFF', black:     '#000000',
  red:       '#C00000', dkGreen:   '#006400',
  blueFont:  '#1F4E79',
};

// ── ENTRY POINTS ───────────────────────────────────────────
function doPost(e) {
  try {
    processModule(JSON.parse(e.postData.contents));
    return resp({status:'ok'});
  } catch(err) {
    return resp({status:'error', msg: err.message});
  }
}
function doGet(e) {
  // If ?action=list → return all modules as JSON for PWA preview
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') {
    try {
      return resp({ status:'ok', modules: getModuleCards() });
    } catch(err) {
      return resp({ status:'error', msg: err.message });
    }
  }
  return resp({status:'ok', msg:'IV Report endpoint live'});
}

// Read Summary sheet rows and return card data
function getModuleCards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Summary');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];

  const data = sheet.getRange(4, 1, lastRow - 3, 21).getValues();
  return data
    .filter(r => r[0] !== '' && r[0] != null)
    .map(r => ({
      num:          r[0],
      type:         r[1],
      serialNumber: r[2],
      model:        r[3],
      mvps:         r[4],
      inverter:     r[5],
      dcb:          r[6],
      string:       r[7],
      ratedPmax:    r[8],
      t1Perf:       r[9],
      t1PmaxMeas:   r[10],
      t1PmaxPred:   r[11],
      t1PmaxSTC:    r[12],
      t2Perf:       r[13],
      t2PmaxMeas:   r[14],
      t2PmaxPred:   r[15],
      t2PmaxSTC:    r[16],
      bifacialW:    r[17],
      bifacialPct:  r[18],
      alerts:       r[19],
      assessment:   r[20],
    }));
}
function resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── PROCESS MODULE ─────────────────────────────────────────
function processModule(m) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheets().filter(s => /^M\d+$/.test(s.getName()));
  const num = existing.length + 1;
  const name = 'M' + String(num).padStart(2,'0');

  ensureSummary(ss);
  appendSummaryRow(ss, m, num);
  buildModuleSheet(ss, m, num, name);
}

// ── SUMMARY ────────────────────────────────────────────────
function ensureSummary(ss) {
  if (ss.getSheetByName('Summary')) return;
  const s = ss.getSheets()[0].getName() === 'Sheet1'
    ? (ss.getSheets()[0].setName('Summary'), ss.getSheets()[0])
    : ss.insertSheet('Summary', 0);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  hdr(s,'A1:U1', 'IV CURVE TEST REPORT — SOLAR MODULE PERFORMANCE SUMMARY', C.darkBlue, 14, true);
  s.setRowHeight(1, 38);

  hdr(s,'A2:U2', `Test Date: ${today}  |  Location: GMT+3  |  Device: Fluke Solmetric PV Analyzer 5.1  |  Total Modules: 0`, C.midBlue, 10, false);
  s.setRowHeight(2, 22);

  const hdrs = ['No.','Status','Serial Number','Model','MVPS','Inv','DCB','String',
    'Rated\nPmax (W)',
    'T1\nPerf (%)','T1 Pmax\nMeas (W)','T1 Pmax\nPred (W)','T1 Pmax\nSTC (W)',
    'T2\nPerf (%)','T2 Pmax\nMeas (W)','T2 Pmax\nPred (W)','T2 Pmax\nSTC (W)',
    'Bifacial\nGain (W)','Bifacial\n(%)','Alerts','Comments'];
  s.getRange(3,1,1,hdrs.length).setValues([hdrs])
    .setBackground(C.darkBlue).setFontColor(C.white).setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  s.setRowHeight(3, 40);

  [5,10,26,20,9,6,8,7,8, 8,11,11,11, 8,11,11,11, 10,8,26,52]
    .forEach((w,i) => s.setColumnWidth(i+1, w*7));

  s.setFrozenRows(3);
  s.setFrozenColumns(3);
}

function appendSummaryRow(ss, m, num) {
  const s = ss.getSheetByName('Summary');
  const t1 = m.t1||{}, t2 = m.t2||{}, np = m.nameplate||{};
  const mtype = m.type||'normal';

  let dW='', dP='';
  try {
    const a=parseFloat(t1.pmaxSTC), b=parseFloat(t2.pmaxSTC);
    if(!isNaN(a)&&!isNaN(b)&&a&&b){ dW=rnd(a-b,2); dP=rnd((a-b)/a*100,1); }
  } catch(e){}

  const alerts = [t1.alerts,t2.alerts].filter(Boolean).join(' | ')||'None';
  const statuses = {damaged:'🔴 DAMAGED', spare:'Spare', normal:'Normal'};

  const row = [
    num, statuses[mtype]||'Normal',
    m.serialNumber||'', m.model||'', m.mvps||'', m.inverter||'', m.dcb||'', m.string||'',
    np.pmax||'',
    t1.performance||'', t1.pmaxMeasured||'', t1.pmaxPredicted||'', t1.pmaxSTC||'',
    t2.performance!=null?t2.performance:'N/A', t2.pmaxMeasured!=null?t2.pmaxMeasured:'N/A',
    t2.pmaxPredicted!=null?t2.pmaxPredicted:'N/A', t2.pmaxSTC!=null?t2.pmaxSTC:'N/A',
    dW, dP, alerts, m.assessment||''
  ];

  const lr = s.getLastRow()+1;
  const rng = s.getRange(lr, 1, 1, row.length);
  rng.setValues([row]).setFontSize(8)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  s.setRowHeight(lr, 54);

  let bg = C.rowGreen;
  if (mtype==='damaged') bg=C.rowRed;
  else if (mtype==='spare') bg=C.rowBlue;
  else if (alerts!=='None') bg=C.rowAmber;
  rng.setBackground(bg);
  s.getRange(lr,2).setFontWeight('bold');
  perfColor(s.getRange(lr,10), t1.performance);
  perfColor(s.getRange(lr,14), t2.performance);

  // update total count in subtitle
  s.getRange('A2').setValue(
    `Test Date: ${Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy')}  |  ` +
    `Location: GMT+3  |  Device: Fluke Solmetric PV Analyzer 5.1  |  Total Modules: ${num}`
  );
}

// ── MODULE SHEET ───────────────────────────────────────────
function buildModuleSheet(ss, m, num, name) {
  const s = ss.insertSheet(name);
  const t1=m.t1||{}, t2=m.t2||{}, np=m.nameplate||{};
  const mtype=m.type||'normal';
  const stat={damaged:'DAMAGED',spare:'Spare',normal:'Normal'}[mtype]||'Normal';
  const sn=m.serialNumber||'SN N/A';

  // Title
  hdr(s,'A1:K1',`Module ${num} — ${stat}  |  ${sn}`, C.darkBlue, 12, true);
  s.setRowHeight(1, 28);

  // Module info section
  hdr(s,'A2:K2','Module Information', C.midBlue, 10, true);
  s.setRowHeight(2, 18);

  s.getRange(3,1,1,11).setValues([['Serial Number','Model','MVPS','DCB','String','Inverter','Rated Pmax','Voc (nom)','Vmp (nom)','Isc (nom)','Imp (nom)']])
    .setBackground(C.midBlue).setFontColor(C.white).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(3, 20);

  s.getRange(4,1,1,11).setValues([[
    sn, m.model||'', m.mvps||'', m.dcb||'', m.string||'', m.inverter||'',
    np.pmax?np.pmax+'W':'', np.voc?np.voc+'V':'', np.vmp?np.vmp+'V':'',
    np.isc?np.isc+'A':'', np.imp?np.imp+'A':''
  ]]).setBackground(C.white).setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(4, 26);

  // IV Results section
  hdr(s,'A5:K5','IV Test Results', C.midBlue, 10, true);
  s.setRowHeight(5, 18);

  // Group headers row 6
  s.getRange('A6:A7').merge().setValue('Parameter')
    .setBackground(C.darkBlue).setFontColor(C.white).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.getRange('B6:E6').merge().setValue('Without Cover (Test 1)')
    .setBackground(C.greenHdr).setFontColor(C.white).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.getRange('F6:I6').merge().setValue('With Backside Covered (Test 2)')
    .setBackground(C.brownHdr).setFontColor(C.white).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.getRange('J6:K6').merge().setValue('Delta (T1 − T2)')
    .setBackground(C.purpleHdr).setFontColor(C.white).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(6, 18);

  // Sub-headers row 7
  [['B',C.greenSub,'Measured'],['C',C.greenSub,'Predicted'],['D',C.greenSub,'STC'],['E',C.greenSub,'Unit'],
   ['F',C.brownSub,'Measured'],['G',C.brownSub,'Predicted'],['H',C.brownSub,'STC'],['I',C.brownSub,'Unit'],
   ['J',C.purpleSub,'Meas Δ'],['K',C.purpleSub,'STC Δ']
  ].forEach(([col,bg,val]) => {
    s.getRange(col+'7').setValue(val).setBackground(bg).setFontColor(C.white)
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  s.setRowHeight(7, 18);

  // Data rows 8-18
  const sf = v => (v!=null&&v!=='') ? v : '—';
  const dl = (a,b) => {
    try{const x=parseFloat(a),y=parseFloat(b); if(!isNaN(x)&&!isNaN(y)) return rnd(x-y,3);}catch(e){}
    return '—';
  };

  const rows = [
    ['Performance (%)', sf(t1.performance),'—','—','%',           sf(t2.performance),'—','—','%',           dl(t1.performance,t2.performance),'—',                        true],
    ['Fill Factor',     sf(t1.fillFactor),'—','—','',              sf(t2.fillFactor),'—','—','',              dl(t1.fillFactor,t2.fillFactor),'—',                          false],
    ['Pmax (W)',        sf(t1.pmaxMeasured),sf(t1.pmaxPredicted),sf(t1.pmaxSTC),'W',   sf(t2.pmaxMeasured),sf(t2.pmaxPredicted),sf(t2.pmaxSTC),'W',   dl(t1.pmaxMeasured,t2.pmaxMeasured),dl(t1.pmaxSTC,t2.pmaxSTC),true],
    ['Irr (W/m²)',      sf(t1.irradiance),'—',1000,'W/m²',        sf(t2.irradiance),'—',1000,'W/m²',        dl(t1.irradiance,t2.irradiance),'—',                          false],
    ['Isc (A)',         sf(t1.iscMeasured),sf(t1.iscPredicted),sf(t1.iscSTC),'A',       sf(t2.iscMeasured),sf(t2.iscPredicted),sf(t2.iscSTC),'A',       dl(t1.iscMeasured,t2.iscMeasured),dl(t1.iscSTC,t2.iscSTC),  true],
    ['Cell Temp (°C)',  sf(t1.cellTemp),'—',25,'°C',              sf(t2.cellTemp),'—',25,'°C',              dl(t1.cellTemp,t2.cellTemp),'—',                              false],
    ['Voc (V)',         sf(t1.vocMeasured),sf(t1.vocPredicted),sf(t1.vocSTC),'V',       sf(t2.vocMeasured),sf(t2.vocPredicted),sf(t2.vocSTC),'V',       dl(t1.vocMeasured,t2.vocMeasured),dl(t1.vocSTC,t2.vocSTC),  true],
    ['Imp (A)',         sf(t1.impMeasured),sf(t1.impPredicted),sf(t1.impSTC),'A',       sf(t2.impMeasured),sf(t2.impPredicted),sf(t2.impSTC),'A',       dl(t1.impMeasured,t2.impMeasured),dl(t1.impSTC,t2.impSTC),  false],
    ['Vmp (V)',         sf(t1.vmpMeasured),sf(t1.vmpPredicted),sf(t1.vmpSTC),'V',       sf(t2.vmpMeasured),sf(t2.vmpPredicted),sf(t2.vmpSTC),'V',       dl(t1.vmpMeasured,t2.vmpMeasured),dl(t1.vmpSTC,t2.vmpSTC),  true],
    ['Current Ratio',   sf(t1.currentRatioMeas),sf(t1.currentRatioPred),sf(t1.currentRatioSTC),'', sf(t2.currentRatioMeas),sf(t2.currentRatioPred),sf(t2.currentRatioSTC),'', dl(t1.currentRatioMeas,t2.currentRatioMeas),'—',false],
    ['Voltage Ratio',   sf(t1.voltageRatioMeas),sf(t1.voltageRatioPred),sf(t1.voltageRatioSTC),'', sf(t2.voltageRatioMeas),sf(t2.voltageRatioPred),sf(t2.voltageRatioSTC),'', dl(t1.voltageRatioMeas,t2.voltageRatioMeas),'—',true],
  ];

  rows.forEach((rd,ri) => {
    const r=ri+8;
    const [param,t1m,t1p,t1s,u1,t2m,t2p,t2s,u2,dm,ds,alt]=rd;
    const af=alt?C.dataLight:C.white;

    s.getRange(r,1).setValue(param).setBackground(C.paramBlue)
      .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
    s.getRange(r,2,1,4).setValues([[t1m,t1p,t1s,u1]])
      .setBackground(af).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    s.getRange(r,6,1,4).setValues([[t2m,t2p,t2s,u2]])
      .setBackground(af).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');

    [dm,ds].forEach((dv,di) => {
      const cell=s.getRange(r,10+di);
      cell.setValue(dv).setBackground(C.deltaLav).setFontWeight('bold')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
      try{
        const n=parseFloat(dv);
        if(!isNaN(n)) cell.setFontColor(n>0?C.red:n<0?C.dkGreen:C.black);
      }catch(e){}
    });
    s.setRowHeight(r,20);
  });

  // Assessment
  s.getRange(19,1,2,11).merge()
    .setValue('📋 ASSESSMENT: '+(m.assessment||'No assessment recorded.'))
    .setBackground(C.assessBg).setFontColor(C.blueFont)
    .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
  s.setRowHeight(19,50); s.setRowHeight(20,16);

  // Column widths
  [140,84,84,84,50,84,84,84,50,84,84].forEach((w,i)=>s.setColumnWidth(i+1,w));

  // IV Curve section
  buildIVCurve(s, m, 22);
}

// ── IV CURVE ───────────────────────────────────────────────
function buildIVCurve(s, m, startRow) {
  const t1=m.t1||{}, t2=m.t2||{};

  hdr(s,`A${startRow}:K${startRow}`,'IV Curve — Current vs Voltage', C.midBlue, 10, true);
  s.setRowHeight(startRow,20);

  // Sub-headers
  const hr=startRow+1;
  s.getRange(hr,1,1,6).setValues([['V — Test 1 (V)','I — Test 1 (A)','V — Test 2 (V)','I — Test 2 (A)','V — MPP T1','V — MPP T2']])
    .setBackground(C.darkBlue).setFontColor(C.white).setFontWeight('bold').setFontSize(8)
    .setHorizontalAlignment('center');
  s.setRowHeight(hr,18);

  // Generate points
  const p1=genIV(t1), p2=genIV(t2);
  const maxN=Math.max(p1.length,p2.length);
  if(maxN===0) return;

  const dataRows=[];
  for(let i=0;i<maxN;i++){
    const a=p1[i]||[null,null];
    const b=p2[i]||[null,null];
    dataRows.push([a[0],a[1],b[0],b[1],'','']);
  }

  // Mark MPP point on each curve
  const mpp1=getMPP(p1); const mpp2=getMPP(p2);
  if(mpp1>=0) dataRows[mpp1][4]=p1[mpp1][0];
  if(mpp2>=0) dataRows[mpp2][5]=p2[mpp2][0];

  const dr=s.getRange(hr+1,1,dataRows.length,6);
  dr.setValues(dataRows).setFontSize(7).setHorizontalAlignment('center');

  // Build chart
  const cb=s.newChart()
    .setChartType(Charts.ChartType.LINE)
    .setPosition(startRow, 7, 5, 5)
    .setOption('title','I-V Curve')
    .setOption('titleTextStyle',{color:C.darkBlue,fontSize:11,bold:true})
    .setOption('backgroundColor',{fill:'#F8FBFF'})
    .setOption('hAxis',{
      title:'Voltage (V)',
      titleTextStyle:{color:C.darkBlue,bold:true,fontSize:10},
      gridlines:{color:'#DCE8F5'},
      minValue:0,
      textStyle:{color:'#333333',fontSize:8}
    })
    .setOption('vAxis',{
      title:'Current (A)',
      titleTextStyle:{color:C.darkBlue,bold:true,fontSize:10},
      gridlines:{color:'#DCE8F5'},
      minValue:0,
      textStyle:{color:'#333333',fontSize:8}
    })
    .setOption('series',{
      0:{color:'#2E75B6',lineWidth:2.5,pointSize:0,labelInLegend:'Test 1 — Without Cover'},
      1:{color:'#C0392B',lineWidth:2.5,pointSize:0,lineDashStyle:[5,3],labelInLegend:'Test 2 — Backside Covered'},
      2:{color:'#2E75B6',lineWidth:0,pointSize:8,pointShape:'circle',labelInLegend:'MPP — Test 1'},
      3:{color:'#C0392B',lineWidth:0,pointSize:8,pointShape:'circle',labelInLegend:'MPP — Test 2'},
    })
    .setOption('legend',{position:'bottom',textStyle:{color:'#333333',fontSize:8}})
    .setOption('chartArea',{left:55,top:35,right:15,bottom:55,width:'82%',height:'73%'})
    .setOption('width',500)
    .setOption('height',300)
    .setOption('curveType','function')
    .setOption('interpolateNulls',true)
    .addRange(s.getRange(hr+1,1,dataRows.length,1)) // V1
    .addRange(s.getRange(hr+1,2,dataRows.length,1)) // I1
    .addRange(s.getRange(hr+1,3,dataRows.length,1)) // V2 — NOTE: Google Sheets LINE chart uses column 1 as X axis
    .addRange(s.getRange(hr+1,4,dataRows.length,1)); // I2

  // Use scatter chart for proper X-Y plotting
  const scatterCb = s.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .setPosition(startRow, 7, 5, 5)
    .setOption('title','I-V Curve')
    .setOption('titleTextStyle',{color:C.darkBlue,fontSize:11,bold:true})
    .setOption('backgroundColor',{fill:'#F8FBFF'})
    .setOption('hAxis',{
      title:'Voltage (V)',
      titleTextStyle:{color:C.darkBlue,bold:true,fontSize:10},
      gridlines:{color:'#DCE8F5'},
      minValue:0,textStyle:{color:'#333333',fontSize:8}
    })
    .setOption('vAxis',{
      title:'Current (A)',
      titleTextStyle:{color:C.darkBlue,bold:true,fontSize:10},
      gridlines:{color:'#DCE8F5'},
      minValue:0,textStyle:{color:'#333333',fontSize:8}
    })
    .setOption('series',{
      0:{color:'#2E75B6',pointSize:2,lineWidth:2,labelInLegend:'Test 1 — Without Cover'},
      1:{color:'#C0392B',pointSize:2,lineWidth:2,labelInLegend:'Test 2 — Backside Covered'},
    })
    .setOption('legend',{position:'bottom',textStyle:{color:'#333333',fontSize:8}})
    .setOption('chartArea',{left:55,top:35,right:15,bottom:55,width:'82%',height:'73%'})
    .setOption('width',500).setOption('height',300)
    .addRange(s.getRange(hr+1,1,dataRows.length,2))  // T1: V,I
    .addRange(s.getRange(hr+1,3,dataRows.length,2)); // T2: V,I

  s.insertChart(scatterCb.build());
}

// ── IV MATH ────────────────────────────────────────────────
function genIV(t) {
  const Isc=parseFloat(t.iscMeasured), Voc=parseFloat(t.vocMeasured);
  const Imp=parseFloat(t.impMeasured), Vmp=parseFloat(t.vmpMeasured);
  if(isNaN(Isc)||isNaN(Voc)||Isc<=0||Voc<=0) return [];

  const N=40;
  let C1=0.0001, C2=0.10;
  if(!isNaN(Imp)&&!isNaN(Vmp)&&Imp>0&&Vmp>0){
    C2=solveC2(Isc,Voc,Imp,Vmp);
    C1=1.0/(Math.exp(1.0/C2)-1.0);
  }

  const pts=[];
  for(let i=0;i<=N;i++){
    const V=Voc*i/N;
    let I=Isc*(1-C1*(Math.exp(V/(C2*Voc))-1));
    if(I<0) I=0;
    pts.push([rnd(V,2), rnd(I,3)]);
  }
  return pts;
}

function getMPP(pts) {
  let maxP=-1, idx=-1;
  pts.forEach((p,i) => { const pw=p[0]*p[1]; if(pw>maxP){maxP=pw;idx=i;} });
  return idx;
}

function solveC2(Isc,Voc,Imp,Vmp) {
  let lo=0.005, hi=0.5;
  for(let i=0;i<60;i++){
    const mid=(lo+hi)/2;
    const c1=1.0/(Math.exp(1.0/mid)-1.0);
    const Ical=Isc*(1-c1*(Math.exp(Vmp/(mid*Voc))-1));
    if(Ical>Imp) lo=mid; else hi=mid;
  }
  return (lo+hi)/2;
}

// ── HELPERS ────────────────────────────────────────────────
function hdr(s, range, val, bg, size, bold) {
  s.getRange(range).merge().setValue(val)
    .setBackground(bg).setFontColor(C.white)
    .setFontWeight(bold?'bold':'normal').setFontSize(size)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function rnd(v,d) { return Math.round(v*Math.pow(10,d))/Math.pow(10,d); }

function perfColor(cell,val) {
  try {
    const p=parseFloat(val);
    if(isNaN(p)) return;
    if(p<90||p>110) cell.setFontColor(C.red).setFontWeight('bold');
    else if(p>=95) cell.setFontColor(C.dkGreen).setFontWeight('bold');
  } catch(e){}
}

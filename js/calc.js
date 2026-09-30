/* ตรรกะคำนวณ — ตรงกับสูตรในไฟล์ TIMESHEET.xlsx (ไม่แตะ DOM / ไม่ผูกกับฐานข้อมูล) */

export const MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
export const WD = ['อา','จ','อ','พ','พฤ','ศ','ส'];
export const NONCHARGE = [
  ['OFFICE', "1. OFFICE (FIRM'S WORK)"], ['ลาป่วย','2. ลาป่วย (SICK)'], ['ลาพักร้อน','3. ลาพักร้อน (HOLIDAY)'],
  ['อบรม','4. อบรม (TRAINING)'], ['ลา OT สะสม','5. ลา OT สะสม (OT ACCUMULATED)'], ['ลากิจ','6. ลากิจ (PERSONAL LEAVE)']
];
export const DAYOFF_LABELS = ['หยุดวันเสาร์','หยุดวันอาทิตย์','หยุดวันนักขัตฤกษ์'];
export const OT_TYPES = ['O.T.', 'Extra Hours'];

export const DEFAULT_SETTINGS = {
  company: 'Dr.Virach & Associates',
  stdHours: 8,
  lunchHours: 1,
  breaks: ['12:00-13:00', '20:00-21:00'],
  clients: Array.from({length: 200}, (_, i) => 'A' + (i + 1)),
  periods: ['69YE', '68YE', 'Q1/69', 'Q2/69', 'Q3/69'],
  holidays: [
    ['2026-01-01','วันขึ้นปีใหม่'],['2026-01-02','วันหยุดพิเศษ (มติ ครม.)'],['2026-03-03','วันมาฆบูชา'],
    ['2026-04-06','วันจักรี'],['2026-04-13','วันสงกรานต์'],['2026-04-14','วันสงกรานต์'],['2026-04-15','วันสงกรานต์'],
    ['2026-05-01','วันแรงงานแห่งชาติ'],['2026-05-04','วันฉัตรมงคล'],['2026-05-31','วันวิสาขบูชา'],
    ['2026-06-01','วันหยุดชดเชยวันวิสาขบูชา'],['2026-06-03','วันเฉลิมพระชนมพรรษา สมเด็จพระราชินี'],
    ['2026-07-28','วันเฉลิมพระชนมพรรษา ร.10'],['2026-07-29','วันอาสาฬหบูชา'],['2026-07-30','วันเข้าพรรษา'],
    ['2026-08-12','วันแม่แห่งชาติ'],['2026-10-13','วันนวมินทรมหาราช'],['2026-10-23','วันปิยมหาราช'],
    ['2026-12-05','วันพ่อแห่งชาติ'],['2026-12-07','วันหยุดชดเชยวันพ่อแห่งชาติ'],['2026-12-10','วันรัฐธรรมนูญ'],['2026-12-31','วันสิ้นปี']
  ].map(([date, name]) => ({date, name}))
};

export const clone = o => JSON.parse(JSON.stringify(o));
export const mergeSettings = s => Object.assign(clone(DEFAULT_SETTINGS), s || {});

/* ---------- เวลา ---------- */
export function parseTime(str){
  if (str == null) return null;
  str = String(str).trim(); if (!str) return null;
  let h, m, mt;
  if ((mt = str.match(/^(\d{1,2})[.:,\s](\d{1,2})$/))){ h=+mt[1]; m=+mt[2]; if (mt[2].length===1) m=+mt[2]*10; }
  else if ((mt = str.match(/^(\d{3,4})$/))){ const v=mt[1].padStart(4,'0'); h=+v.slice(0,2); m=+v.slice(2); }
  else if ((mt = str.match(/^(\d{1,2})$/))){ h=+mt[1]; m=0; }
  else return NaN;
  if (m>59 || h>24 || (h===24 && m>0)) return NaN;
  return h*60+m;
}
export const fmtTime = mins => (mins==null || isNaN(mins)) ? '' : String(Math.floor(mins/60)).padStart(2,'0')+':'+String(mins%60).padStart(2,'0');
export const f2 = x => (x==null || x==='' || isNaN(x)) ? '' : (Math.round(x*100)/100).toFixed(2);

export function breakList(S){
  return (S.breaks||[]).map(b => { const m=b.split('-').map(x=>parseTime(x.trim())); return (m.length===2 && m.every(x=>x!=null && !isNaN(x)) && m[1]>m[0]) ? m : null; }).filter(Boolean);
}
export function rowMins(S, r){
  const a=parseTime(r.in), b=parseTime(r.out);
  if (a==null || b==null || isNaN(a) || isNaN(b) || b<=a) return null;
  let t=b-a;
  for (const [s,e] of breakList(S)) t -= Math.max(0, Math.min(b,e)-Math.max(a,s));
  return Math.max(0,t);
}

/* ---------- ปฏิทิน (ปี = พ.ศ.) ---------- */
export const ceYear = be => be - 543;
export const daysInMonth = (be, m) => new Date(ceYear(be), m, 0).getDate();
export const ymd = (y,m,d) => y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');
export function dayKind(S, be, m, d){
  const dow = new Date(ceYear(be), m-1, d).getDay();
  const h = (S.holidays||[]).find(x => x.date === ymd(ceYear(be), m, d));
  if (h) return {kind:'hol', tag:'นักขัตฤกษ์', title:h.name, off:true, dow, label:'หยุดวันนักขัตฤกษ์'};
  if (dow===6) return {kind:'sat', tag:'เสาร์', off:true, dow, label:'หยุดวันเสาร์'};
  if (dow===0) return {kind:'sat', tag:'อาทิตย์', off:true, dow, label:'หยุดวันอาทิตย์'};
  return {kind:'', off:false, dow};
}

/* ---------- แถวข้อมูล ---------- */
export const blankRow = () => ({in:'', out:'', client:'', service:'', note:'', ot:''});
export const isBlank = r => !(r.in || r.out || r.client || r.service || r.note || r.ot);
/** entries(flat [{d,...}]) -> {1:[rows],...} มีอย่างน้อย 1 แถวต่อวัน */
export function toRowsByDay(entries, n){
  const rows = {};
  (entries||[]).forEach(e => { const d=+e.d; if (d>=1 && d<=n) { const {d:_d, ...rest} = e; (rows[d] = rows[d] || []).push(Object.assign(blankRow(), rest)); } });
  for (let d=1; d<=n; d++) if (!rows[d] || !rows[d].length) rows[d] = [blankRow()];
  return rows;
}
export function toEntries(rowsByDay){
  const out = [];
  Object.keys(rowsByDay).map(Number).sort((a,b)=>a-b).forEach(d => rowsByDay[d].forEach(r => { if (!isBlank(r)) out.push(Object.assign({d}, r)); }));
  return out;
}

/* ---------- สรุปรายเดือน (TIME SHEET) ---------- */
export function summarize(S, be, m, rowsByDay){
  const n = daysInMonth(be, m), std = +S.stdHours || 8;
  const kinds = []; for (let d=1; d<=n; d++) kinds[d] = dayKind(S, be, m, d);
  const allRows = [];
  for (let d=1; d<=n; d++) for (const r of (rowsByDay[d]||[])) allRows.push({d, r, m:rowMins(S, r)});
  const dayTot = Array(n+1).fill(0);
  allRows.forEach(x => { if (x.m!=null) dayTot[x.d] += x.m/60; });
  const sumFor = code => { const v=Array(n+1).fill(0); allRows.forEach(x => { if (x.m!=null && x.r.client===code) v[x.d] += x.m/60; }); return v; };
  const used = new Set(allRows.filter(x => x.r.client).map(x => x.r.client));
  const clientRows = S.clients.filter(c => used.has(c)).map((c,i) => {
    const first = allRows.find(x => x.r.client===c && x.r.service);
    const v = sumFor(c);
    return {no:i+1, code:c, period:first ? first.r.service : '', v, total:v.slice(1).reduce((a,b)=>a+b,0)};
  });
  const colSum = rows => { const v=Array(n+1).fill(0); rows.forEach(r => { for (let d=1; d<=n; d++) v[d]+=r.v[d]; }); return v; };
  const A = colSum(clientRows);
  const nonRows = NONCHARGE.map(([code,label]) => { const v=sumFor(code); return {label, v, total:v.slice(1).reduce((a,b)=>a+b,0)}; });
  const B = colSum(nonRows);
  const otFor = type => { const v=Array(n+1).fill(null);
    for (let d=1; d<=n; d++){
      if (!allRows.some(x => x.d===d && x.r.ot===type)) continue;
      v[d] = kinds[d].off ? dayTot[d] : Math.max(0, dayTot[d]-std);
    } return v; };
  const C = otFor('O.T.'), D = otFor('Extra Hours');
  const sum = a => a.slice(1).reduce((x,y) => x+(y||0), 0);
  const Sd = Array(n+1).fill(null);
  for (let d=1; d<=n; d++){ const ab=A[d]+B[d]; Sd[d] = ab===0 ? null : ab-(C[d]||0)-(D[d]||0); }
  const tA=sum(A), tB=sum(B), tC=sum(C), tD=sum(D);
  return {n, kinds, clientRows, A, nonRows, B, C, D, S:Sd, tA, tB, tC, tD, tS:tA+tB-tC-tD, allRows, dayTot, std};
}

/* ---------- O.T. รายแถว (TIME REPORT O.T.) ---------- */
export function otRows(S, be, m, rowsByDay, type){
  const stdM = (+S.stdHours||8)*60, lunchM = (+S.lunchHours||0)*60, out = [];
  const n = daysInMonth(be, m);
  for (let d=1; d<=n; d++){
    const off = dayKind(S, be, m, d).off; let cum = 0;
    for (const r of (rowsByDay[d]||[])){
      const mins = rowMins(S, r); if (mins==null) continue;
      const before = cum; cum += mins;
      if (type && r.ot !== type) continue;
      if (!r.ot) continue;
      const a = parseTime(r.in), e = parseTime(r.out);
      let om, start;
      if (off){ om = mins; start = a; }
      else {
        om = Math.min(mins, Math.max(0, before+mins-stdM));
        start = before>=stdM ? a : a + (stdM-before) + (before===0 ? lunchM : 0);
        if (start > e) start = e;
      }
      out.push({d, start, end:e, client:r.client, service:r.service, note:r.note, mins:om, type:r.ot, workMins:mins});
    }
  }
  return out;
}

/* ---------- ข้อมูลสำหรับรายงานวิเคราะห์ (รายจ็อบ/รายปี ฯลฯ) ---------- */
export function categoryOf(client){
  if (!client) return 'ไม่ระบุ';
  if (NONCHARGE.some(x => x[0]===client)) return 'Non-charge / ลา';
  if (DAYOFF_LABELS.includes(client)) return 'วันหยุด';
  return 'เรียกเก็บ (Charge)';
}
/** doc: {uid,name,year,month,entries[]} -> facts[] (1 รายการต่อแถวที่มีชั่วโมง) */
export function factsOf(S, doc){
  const n = daysInMonth(doc.year, doc.month);
  const rows = toRowsByDay(doc.entries, n);
  const ot = otRows(S, doc.year, doc.month, rows, null);
  const facts = [];
  for (let d=1; d<=n; d++) rows[d].forEach((r, i) => {
    const mins = rowMins(S, r); if (mins==null) return;
    facts.push({uid:doc.uid, name:doc.name||doc.uid, year:doc.year, month:doc.month, day:d,
      client:r.client||'(ไม่ระบุ)', service:r.service||'(ไม่ระบุ)', cat:categoryOf(r.client), mins, ot:r.ot||'', otMins:0, _k:d+'#'+i});
  });
  // ใส่ otMins ให้แถวที่ตรงกัน (วัน + ลำดับ)
  const perDay = {};
  ot.forEach(o => { (perDay[o.d] = perDay[o.d] || []).push(o); });
  for (let d=1; d<=n; d++){
    const withOt = facts.filter(f => f.day===d && f.ot);
    const list = perDay[d] || [];
    withOt.forEach((f, k) => { f.otMins = list[k] ? list[k].mins : 0; });
  }
  return facts;
}

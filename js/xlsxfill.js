/* เติมข้อมูลลงในไฟล์ต้นแบบ template.xlsx (TIMESHEET เดิม) โดยแก้เฉพาะช่องกรอกข้อมูล
   ส่วนหน้าตา สูตร สี เส้นขอบ แผ่นงาน validation ฯลฯ คงเดิมทุกอย่าง — Excel จะคำนวณใหม่เองเมื่อเปิดไฟล์ */
import { MONTHS, DAYOFF_LABELS, parseTime, dayKind, daysInMonth, ceYear, toRowsByDay } from './calc.js';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const REPORT_FIRST_ROW = 6, REPORT_LAST_ROW = 37;          // สูตร SUMIFS ใน TIME SHEET อ่านช่วง 6..37
export const REPORT_CAPACITY = REPORT_LAST_ROW - REPORT_FIRST_ROW + 1; // 32 แถว

const colNum = s => s.split('').reduce((a,c) => a*26 + c.charCodeAt(0)-64, 0);
const splitRef = ref => { const m = ref.match(/^([A-Z]+)(\d+)$/); return {col:m[1], num:colNum(m[1]), row:+m[2]}; };

class Sheet {
  constructor(xml){
    this.doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (this.doc.getElementsByTagName('parsererror').length) throw new Error('อ่านไฟล์ต้นแบบไม่ได้ (XML)');
    this.sd = this.doc.getElementsByTagName('sheetData')[0];
    this.rows = new Map();
    for (const r of this.sd.children) this.rows.set(+r.getAttribute('r'), r);
  }
  row(n, create){
    let r = this.rows.get(n);
    if (!r && create){
      r = this.doc.createElementNS(NS, 'row'); r.setAttribute('r', n);
      let before = null; for (const [k,v] of this.rows) if (k > n && (!before || k < +before.getAttribute('r'))) before = v;
      this.sd.insertBefore(r, before); this.rows.set(n, r);
    }
    return r;
  }
  cell(ref, create){
    const {num, row} = splitRef(ref);
    const r = this.row(row, create); if (!r) return null;
    let before = null;
    for (const c of r.children){
      const cn = splitRef(c.getAttribute('r')).num;
      if (cn === num) return c;
      if (cn > num){ before = c; break; }
    }
    if (!create) return null;
    const c = this.doc.createElementNS(NS, 'c'); c.setAttribute('r', ref);
    r.insertBefore(c, before); r.removeAttribute('spans');
    return c;
  }
  _reset(c){
    const f = c.getElementsByTagName('f')[0];
    if (f && f.getAttribute('t') === 'shared' && f.textContent) throw new Error('cell ' + c.getAttribute('r') + ' เป็นสูตรหลักของ shared formula');
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
  }
  /** v: number | string | null(ล้างค่า)  style: ใช้ s ของช่องนี้ ถ้าไม่มีให้ยืมจาก styleFrom */
  set(ref, v, styleFrom){
    const empty = (v === null || v === undefined || v === '');
    const c = this.cell(ref, !empty); if (!c) return;
    this._reset(c);
    if (!c.getAttribute('s') && styleFrom){ const o = this.cell(styleFrom, false); if (o && o.getAttribute('s')) c.setAttribute('s', o.getAttribute('s')); }
    if (empty) return;
    if (typeof v === 'number'){ const e = this.doc.createElementNS(NS, 'v'); e.textContent = String(v); c.appendChild(e); }
    else {
      c.setAttribute('t', 'inlineStr');
      const is = this.doc.createElementNS(NS, 'is'), t = this.doc.createElementNS(NS, 't');
      t.textContent = String(v); if (/^\s|\s$/.test(String(v))) t.setAttribute('xml:space', 'preserve');
      is.appendChild(t); c.appendChild(is);
    }
  }
  /** ล้างค่าเซลล์ที่มีอยู่จริงในคอลัมน์ col ช่วงแถว a..b */
  clearRange(col, a, b){ for (let r=a; r<=b; r++){ const c = this.cell(col+r, false); if (c) this._reset(c); } }
  has(ref){ return !!this.cell(ref, false); }
  serialize(){ return XML_HEAD + new XMLSerializer().serializeToString(this.doc.documentElement); }
}

const timeNum = str => {           // "08:30" -> 8.3 (รูปแบบ ชม.นาที ที่ไฟล์เดิมใช้), 00:00 ขาเข้า -> 0.01
  const m = parseTime(str); if (m == null || isNaN(m)) return null;
  return m === 0 ? 0.01 : Number((Math.floor(m/60) + (m%60)/100).toFixed(2));
};
const serial = iso => { const [y,m,d] = iso.split('-').map(Number); return Math.round((Date.UTC(y, m-1, d) - Date.UTC(1899, 11, 30)) / 86400000); };

/** doc: {name, division, year(พ.ศ.), month, entries[], approvedBy?}  -> {blob(Uint8Array), warnings[]} */
export async function fillTemplate(templateBuf, S, doc){
  const JSZip = window.JSZip; if (!JSZip) throw new Error('ไม่พบไลบรารี JSZip');
  const zip = await JSZip.loadAsync(templateBuf);
  const warnings = [];

  // --- ทำให้ Excel คำนวณใหม่เมื่อเปิด + ตัด calcChain (เพราะเราเปลี่ยนสูตรบางเซลล์เป็นค่า) ---
  let wb = await zip.file('xl/workbook.xml').async('string');
  wb = /<calcPr[^>]*\/>/.test(wb)
    ? wb.replace(/<calcPr([^>]*?)\/>/, (m, a) => '<calcPr' + a.replace(/\sfullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>')
    : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.file('xl/workbook.xml', wb);
  if (zip.file('xl/calcChain.xml')){
    zip.remove('xl/calcChain.xml');
    let rel = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    rel = rel.replace(/<Relationship [^>]*calcChain[^>]*\/>/, ''); zip.file('xl/_rels/workbook.xml.rels', rel);
    let ct = await zip.file('[Content_Types].xml').async('string');
    ct = ct.replace(/<Override [^>]*calcChain[^>]*\/>/, ''); zip.file('[Content_Types].xml', ct);
  }

  const load = async p => new Sheet(await zip.file(p).async('string'));
  const ts = await load('xl/worksheets/sheet1.xml');   // TIME SHEET
  const tr = await load('xl/worksheets/sheet2.xml');   // TIME REPORT
  const ot = await load('xl/worksheets/sheet3.xml');   // TIME REPORT O.T.
  const hd = await load('xl/worksheets/sheet4.xml');   // HOLIDAYS

  // --- หัวรายงาน ---
  ts.set('C3', MONTHS[doc.month-1]); ts.set('F3', +doc.year);
  ts.set('Z3', doc.name || ''); ts.set('AH3', doc.division || '');
  ts.set('AO1', +S.stdHours || 8);
  ot.set('T1', +S.stdHours || 8); ot.set('T2', +S.lunchHours || 0);
  tr.set('O2', doc.approvedBy || null);

  // --- ตารางบันทึกเวลา: 1 แถว/รายการ (วันที่ไม่มีรายการยังคงมี 1 แถวเหมือนไฟล์เดิม) ---
  const n = daysInMonth(doc.year, doc.month), rows = toRowsByDay(doc.entries, n);
  let items = [];
  for (let d=1; d<=n; d++){
    const k = dayKind(S, doc.year, doc.month, d);
    rows[d].forEach((r, i) => items.push({d, r, empty: !(r.in||r.out||r.client||r.service||r.note||r.ot), label: (i===0 && k.off) ? k.label : ''}));
  }
  if (items.length > REPORT_CAPACITY){                       // ตัดแถววันว่างออกก่อน
    let over = items.length - REPORT_CAPACITY;
    items = items.filter(it => { if (over > 0 && it.empty && !it.label) { over--; return false; } return true; });
    // ยังเกิน -> ตัดแถววันหยุดว่าง
    let over2 = items.length - REPORT_CAPACITY;
    items = items.filter(it => { if (over2 > 0 && it.empty) { over2--; return false; } return true; });
  }
  if (items.length > REPORT_CAPACITY){
    warnings.push(`เดือนนี้มี ${items.length} รายการ แต่ฟอร์ม Excel รองรับสูงสุด ${REPORT_CAPACITY} แถว — ${items.length-REPORT_CAPACITY} รายการสุดท้ายไม่ถูกใส่ในไฟล์ Excel (ข้อมูลในระบบยังครบ)`);
    items = items.slice(0, REPORT_CAPACITY);
  }
  for (let i=0; i<REPORT_CAPACITY; i++){
    const row = REPORT_FIRST_ROW + i, it = items[i];
    if (!it){ ['A','B','C','E','H','I','Q'].forEach(c => tr.set(c+row, null)); continue; }
    tr.set('A'+row, it.d);
    tr.set('B'+row, timeNum(it.r.in));
    tr.set('C'+row, it.r.out && parseTime(it.r.out) === 1440 ? 24 : timeNum(it.r.out));
    tr.set('E'+row, it.r.client || it.label || null);
    tr.set('H'+row, it.r.service || null);
    tr.set('I'+row, it.r.note || null);
    tr.set('Q'+row, it.r.ot || null);
  }

  // --- HOLIDAYS: วันหยุด / รายชื่อลูกค้า / Period จากการตั้งค่า ---
  const hols = [...S.holidays].sort((a,b) => a.date.localeCompare(b.date));
  if (hols.length > 23 && hd.has('A28')){                    // ย้ายหมายเหตุออกจากช่วงรายการวันหยุด
    const note = hd.cell('A28', false); const txt = note.textContent;
    hd.set('A28', null);
    if (txt) hd.set('A2', txt.trim() || null);
  }
  for (let i=0; i<195; i++){
    const row = 5 + i, h = hols[i];
    if (h){ hd.set('A'+row, serial(h.date), 'A5'); hd.set('B'+row, h.name, 'B5'); if (!hd.has('C'+row)) hd.set('C'+row, null); }
    else { hd.set('A'+row, null); if (i >= 22) hd.set('B'+row, null); else hd.set('B'+row, null); }
  }
  S.clients.slice(0, 489).forEach((c, i) => hd.set('H'+(12+i), c));
  hd.clearRange('H', 12 + Math.min(S.clients.length, 489), 500);
  S.periods.slice(0, 190).forEach((p, i) => hd.set('I'+(3+i), p));
  hd.clearRange('I', 3 + Math.min(S.periods.length, 190), 200);

  zip.file('xl/worksheets/sheet1.xml', ts.serialize());
  zip.file('xl/worksheets/sheet2.xml', tr.serialize());
  zip.file('xl/worksheets/sheet3.xml', ot.serialize());
  zip.file('xl/worksheets/sheet4.xml', hd.serialize());

  const out = await zip.generateAsync({type:'uint8array', compression:'DEFLATE', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  return {data: out, warnings};
}

export const xlsxName = doc => `TIMESHEET_${(doc.name||'noname').replace(/[\\/:*?"<>|\s]+/g,'_')}_${doc.year}-${String(doc.month).padStart(2,'0')}.xlsx`;

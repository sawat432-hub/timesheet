import * as C from './calc.js';
import { createBackend } from './backend.js';
import { fillTemplate, xlsxName, REPORT_CAPACITY } from './xlsxfill.js';

const $ = (s, r=document) => r.querySelector(s);
const $$ = (s, r=document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const lsGet = (k, d) => { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch(e){ return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} };

/* ================= state ================= */
let be, user = null, S = C.mergeSettings();
const now = new Date();
const last = lsGet('ts_last', {});
let view = {uid:'', name:'', division:'', year: last.year || now.getFullYear()+543, month: last.month || now.getMonth()+1};
let rows = {};                                   // {day:[row]}
let meta = {status:'draft'};                     // สถานะเอกสารเดือนนี้
let docExists = false;
let activeTab = 'report';
let users = [];                                  // แอดมิน: รายชื่อพนักงาน
const isAdmin = () => user && user.role === 'admin';
const isSelf = () => user && view.uid === user.uid;
const readOnly = () => meta.status === 'approved' && !isAdmin();
const n = () => C.daysInMonth(view.year, view.month);

function toast(msg, ms=2600){ const t=$('#toast'); t.textContent=msg; t.classList.add('on'); clearTimeout(t._t); t._t=setTimeout(()=>t.classList.remove('on'), ms); }
function download(name, data, type){ const a=document.createElement('a'); a.href=URL.createObjectURL(data instanceof Blob ? data : new Blob([data],{type})); a.download=name; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();}, 800); }
const titleLine = t => `${S.company} — ${t} · ${view.name||''} · ${view.division||''} · ${C.MONTHS[view.month-1]} ${view.year}`;

/* ================= เริ่มระบบ / ล็อกอิน ================= */
(async function boot(){
  $('#mSel').innerHTML = C.MONTHS.map((m,i) => `<option value="${i+1}">${m}</option>`).join('');
  try { be = await createBackend(); }
  catch(e){ $('#login').style.display='flex'; $('#lgErr').textContent = 'เริ่มระบบไม่สำเร็จ: ' + e.message; return; }
  $('#login').style.display = 'flex';
  { const dl = C.mergeSettings().divisions; let last = ''; try { last = sessionStorage.getItem('ts_pickdiv') || ''; } catch(e){}
    $('#lgDiv').innerHTML = '<option value="">— เลือกสายงาน —</option>' + dl.map(x => `<option value="${esc(x.name)}">${esc(x.code + ' ' + x.name)}</option>`).join('');
    $('#lgDiv').value = last;
    $('#lgDiv').onchange = () => { try { sessionStorage.setItem('ts_pickdiv', $('#lgDiv').value); } catch(e){} }; }
  $('#lgFirebase').style.display = be.mode==='firebase' ? '' : 'none';
  $('#lgDemo').style.display = be.mode==='demo' ? '' : 'none';
  $('#modebar').style.display = be.mode==='demo' ? '' : 'none';
  $('#btnGoogle').onclick = async () => { $('#lgErr').textContent=''; try { await be.signIn(); } catch(e){ $('#lgErr').textContent = 'เข้าสู่ระบบไม่สำเร็จ: ' + (e.code || e.message); } };
  $('#btnDemo').onclick = () => be.signIn($('#dmEmail').value, $('#dmName').value);
  $('#btnOut').onclick = async () => { await flushSave(); await be.signOut(); };
  be.onAuth(async (u, err) => {
    if (err) $('#lgErr').textContent = err.message || String(err);
    if (!u){ user = null; $('#app').style.display='none'; $('#login').style.display='flex'; return; }
    try { await enter(u); } catch(e){ console.error(e); $('#lgErr').textContent = 'โหลดข้อมูลไม่สำเร็จ: ' + (e.code || e.message); $('#app').style.display='none'; $('#login').style.display='flex'; }
  });
})();

async function enter(u){
  user = u;
  try { const cfg = await be.getConfig(); S = C.mergeSettings(cfg); } catch(e){ S = C.mergeSettings(); }
  view.uid = u.uid; view.name = u.name || u.email; view.division = u.division || '';
  $('#login').style.display = 'none'; $('#app').style.display = 'block';
  if (!C.findDiv(S, u.division)) {
    const pick = C.findDiv(S, ($('#lgDiv') && $('#lgDiv').value) || '');
    if (pick) { try { await be.saveProfile(u.uid, {name:u.name || u.email, division:pick.name}); u.division = pick.name; } catch(e){ console.error(e); } }
    if (!C.findDiv(S, u.division)) await askProfile(u);
  }
  view.name = u.name || u.email; view.division = u.division || '';
  $('#hCompany').textContent = S.company;
  $('#uName').textContent = u.name || u.email; $('#uRole').textContent = isAdmin() ? 'แอดมิน' : 'พนักงาน';
  $('#uPhoto').src = u.photo || ''; $('#uPhoto').style.display = u.photo ? '' : 'none';
  $$('.adminonly').forEach(b => b.style.display = isAdmin() ? '' : 'none');
  $('#fEmp').style.display = isAdmin() ? '' : 'none';
  $('#rEmpF').style.display = isAdmin() ? '' : 'none';
  $('#mSel').value = view.month; $('#yIn').value = view.year;
  $('#bApprove').style.display = isAdmin() ? '' : 'none';
  if (isAdmin()) { await refreshUsers(); }
  showTab('report');
  initReportControls();
  await loadView();
}

async function refreshUsers(){
  try { users = await be.listUsers(); } catch(e){ users = []; }
  users.sort((a,b) => (a.name||'').localeCompare(b.name||'', 'th'));
  $('#empSel').innerHTML = users.map(x => `<option value="${esc(x.uid)}">${esc(x.name||x.email)}${x.uid===user.uid?' (ฉัน)':''}</option>`).join('');
  $('#empSel').value = view.uid;
}

/* ================= โหลด / บันทึก ================= */
async function loadView(){
  await flushSave();
  const d = await be.getSheet(view.uid, view.year, view.month);
  docExists = !!d;
  rows = C.toRowsByDay(d ? d.entries : [], n());
  meta = {status: d?.status || 'draft', approvedBy: d?.approvedBy || '', approvedAt: d?.approvedAt || '', submittedAt: d?.submittedAt || ''};
  if (!isSelf()){ const p = users.find(x => x.uid===view.uid); if (p){ view.name = d?.name || p.name || p.email; view.division = d?.division ?? p.division ?? ''; } }
  else { view.name = user.name || user.email; view.division = user.division || ''; }
  $('#pName').value = view.name; fillDivSel(view.division);
  $('#pName').disabled = $('#pDiv').disabled = !isSelf();
  $('#hCompany').textContent = S.company;
  applyState(); renderReport(); renderOthers();
}
let saveTimer = null, saving = null;
function setSaveState(t, cls){ const el=$('#saveState'); el.textContent=t; el.className='savestate '+(cls||''); }
function scheduleSave(){ setSaveState('กำลังบันทึก…'); clearTimeout(saveTimer); saveTimer = setTimeout(flushSave, 700); }
function currentDoc(){
  return {uid:view.uid, name:view.name, division:view.division, year:view.year, month:view.month, entries:C.toEntries(rows),
    status:meta.status, approvedBy:meta.approvedBy||'', approvedAt:meta.approvedAt||'', submittedAt:meta.submittedAt||''};
}
async function flushSave(){
  if (!saveTimer && !saving) return;
  clearTimeout(saveTimer); saveTimer = null;
  const doc = currentDoc();
  if (!docExists && !doc.entries.length) { setSaveState(''); return; }
  saving = be.saveSheet(doc);
  try { await saving; docExists = true; setSaveState('✓ บันทึกแล้ว', 'ok'); }
  catch(e){ console.error(e); setSaveState('บันทึกไม่สำเร็จ: ' + (e.code||e.message), 'err'); toast('บันทึกไม่สำเร็จ — ตรวจสอบอินเทอร์เน็ต/สิทธิ์การใช้งาน', 4000); }
  saving = null;
}
async function saveNow(){ saveTimer = saveTimer || 1; await flushSave(); }

/* ================= สถานะ / ปุ่ม ================= */
const ST = {draft:'ร่าง', submitted:'ส่งแล้ว', approved:'อนุมัติแล้ว', none:'ยังไม่มี'};
function applyState(){
  const chip = $('#stChip'); chip.className = 'chip ' + meta.status; chip.textContent = ST[meta.status] + (meta.status==='approved' && meta.approvedBy ? ' โดย ' + meta.approvedBy : '');
  $('#roBanner').style.display = readOnly() ? '' : 'none';
  $('#tReport').classList.toggle('ro', readOnly());
  $('#bClear').style.display = readOnly() ? 'none' : '';
  const sb = $('#bSubmit'); sb.style.display = (isSelf() || isAdmin()) && meta.status!=='approved' ? '' : 'none';
  sb.textContent = meta.status==='submitted' ? 'ยกเลิกการส่ง' : 'ส่งรายงานเดือนนี้';
  const ab = $('#bApprove'); ab.textContent = meta.status==='approved' ? 'ยกเลิกการอนุมัติ' : 'อนุมัติ';
  ab.style.display = isAdmin() ? '' : 'none';
  ab.disabled = !docExists && !C.toEntries(rows).length;
}
$('#bSubmit').onclick = async () => {
  meta.status = meta.status==='submitted' ? 'draft' : 'submitted';
  meta.submittedAt = meta.status==='submitted' ? new Date().toISOString() : '';
  docExists = docExists || true; await saveNow(); applyState(); toast(meta.status==='submitted' ? 'ส่งรายงานแล้ว' : 'ยกเลิกการส่งแล้ว');
};
$('#bApprove').onclick = async () => {
  if (meta.status==='approved'){ meta.status='submitted'; meta.approvedBy=''; meta.approvedAt=''; }
  else { meta.status='approved'; meta.approvedBy=user.name||user.email; meta.approvedAt=new Date().toISOString(); }
  await saveNow(); applyState(); renderReport(); toast(meta.status==='approved' ? 'อนุมัติแล้ว' : 'ยกเลิกการอนุมัติแล้ว');
};

/* ================= หัวเรื่อง: เดือน/ปี/ชื่อ ================= */
async function onPeriodChange(){
  await flushSave();
  view.month = +$('#mSel').value; view.year = Math.min(2700, Math.max(2500, +$('#yIn').value || view.year));
  lsSet('ts_last', {year:view.year, month:view.month}); $('#yIn').value = view.year;
  await loadView();
}
$('#mSel').onchange = $('#yIn').onchange = onPeriodChange;
$('#empSel').onchange = async () => { await flushSave(); view.uid = $('#empSel').value; await loadView(); };
function divOptions(cur){
  const d = C.findDiv(S, cur);
  return `<option value="">— เลือกสายงาน —</option>` + S.divisions.map(x => `<option value="${esc(x.name)}">${esc(x.code + ' ' + x.name)}</option>`).join('') + (cur && !d ? `<option value="${esc(cur)}">${esc(cur)}</option>` : '');
}
function fillDivSel(cur){ const el = $('#pDiv'), d = C.findDiv(S, cur); el.innerHTML = divOptions(cur); el.value = d ? d.name : (cur || ''); }
/* หน้าต่างยืนยันข้อมูลพนักงาน (แสดงเมื่อยังไม่ได้เลือกสายงาน) */
function askProfile(u){
  return new Promise(resolve => {
    $('#pmEmail').value = u.email || ''; $('#pmName').value = u.name || ''; $('#pmDiv').innerHTML = divOptions(''); $('#pmDiv').value = ''; $('#pmErr').textContent = '';
    $('#profModal').style.display = 'flex';
    $('#pmOk').onclick = async () => {
      const name = $('#pmName').value.trim(), div = $('#pmDiv').value;
      if (!name) return $('#pmErr').textContent = 'กรุณากรอกชื่อ-นามสกุล';
      if (!div) return $('#pmErr').textContent = 'กรุณาเลือกสายงาน';
      try { await be.saveProfile(u.uid, {name, division:div}); } catch(e){ return $('#pmErr').textContent = 'บันทึกไม่สำเร็จ: ' + (e.code || e.message); }
      u.name = name; u.division = div; $('#profModal').style.display = 'none'; resolve();
    };
  });
}
$('#pName').onchange = $('#pDiv').onchange = async () => {
  if (!isSelf()) return;
  view.name = $('#pName').value.trim() || view.name; view.division = $('#pDiv').value;
  user.name = view.name; user.division = view.division; $('#uName').textContent = user.name;
  try { await be.saveProfile(user.uid, {name:view.name, division:view.division}); } catch(e){ toast('บันทึกโปรไฟล์ไม่สำเร็จ'); }
  if (docExists) scheduleSave(); renderOthers(); renderReport();
};

/* ================= TAB ================= */
function showTab(t){
  activeTab = t;
  $$('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab===t));
  $$('.panel').forEach(p => p.classList.toggle('active', p.id==='p-'+t));
  if (t==='sheet') renderSheet(); if (t==='ot') renderOt(); if (t==='settings') fillSettings(); if (t==='admin') renderAdmin(); if (t==='reports') syncReportControls();
}
$$('.tab').forEach(b => b.onclick = () => showTab(b.dataset.tab));
function renderOthers(){ if (activeTab==='sheet') renderSheet(); if (activeTab==='ot') renderOt(); }

/* ================= บันทึกเวลา ================= */
function optHTML(opts){ return opts.map(o => `<option>${esc(o)}</option>`).join(''); }
function clientOptions(){
  return `<option value=""></option><optgroup label="งานลูกค้า / จ็อบ">${optHTML(C.clientsFor(S, view.division))}</optgroup><optgroup label="งานที่ไม่เรียกเก็บ / ลา">${optHTML(C.NONCHARGE.map(x=>x[0]))}</optgroup><optgroup label="วันหยุด">${optHTML(C.DAYOFF_LABELS)}</optgroup>`;
}
function setSel(sel, v){
  if (v && ![...sel.options].some(o => o.value===v)){ const o=document.createElement('option'); o.textContent=v; sel.appendChild(o); }
  sel.value = v || '';
}
function renderReport(){
  const cOpt = clientOptions(), sOpt = `<option value=""></option>${optHTML(S.periods)}`, oOpt = `<option value=""></option>${optHTML(C.OT_TYPES)}`;
  let html = '';
  for (let d=1; d<=n(); d++){
    const k = C.dayKind(S, view.year, view.month, d);
    rows[d].forEach((r,i) => {
      html += `<tr class="${k.kind}" data-d="${d}" data-i="${i}">
        <td class="${i?'cont':'day'}">${i ? '↳' : `<span class="dn">${d}</span> <span class="wd">${C.WD[k.dow]}</span>${k.tag?`<span class="tag">${esc(k.tag)}</span>`:''}${k.title?`<span class="dt">${esc(k.title)}</span>`:''}`}</td>
        <td><input class="t" data-f="in" inputmode="decimal" placeholder="08:30" value="${esc(r.in)}"></td>
        <td><input class="t" data-f="out" inputmode="decimal" placeholder="17:30" value="${esc(r.out)}"></td>
        <td><select class="cl" data-f="client">${cOpt}</select></td>
        <td><select data-f="service">${sOpt}</select></td>
        <td><input data-f="note" style="width:170px" value="${esc(r.note)}"></td>
        <td class="num h"></td><td class="num m"></td>
        <td><select data-f="ot">${oOpt}</select></td>
        <td class="act">${i ? '<button class="ic rm" data-a="rm" title="ลบแถวนี้">×</button>' : '<button class="ic" data-a="add" title="เพิ่มรายการในวันนี้">＋</button>'}</td></tr>`;
    });
  }
  $('#rbody').innerHTML = html;
  $$('#rbody tr').forEach(tr => { const r = rows[tr.dataset.d][tr.dataset.i]; setSel($('[data-f=client]',tr), r.client); setSel($('[data-f=service]',tr), r.service); setSel($('[data-f=ot]',tr), r.ot); });
  const ro = readOnly(); $$('#rbody input,#rbody select,#rbody button').forEach(e => e.disabled = ro);
  refreshComputed();
}
function refreshComputed(){
  let tot = 0, cnt = 0;
  $$('#rbody tr').forEach(tr => {
    const r = rows[tr.dataset.d][tr.dataset.i], m = C.rowMins(S, r), bad = !!(r.in && r.out && m==null);
    $('.h',tr).textContent = m==null ? '' : Math.floor(m/60); $('.m',tr).textContent = m==null ? '' : m%60;
    $$('input.t',tr).forEach(inp => inp.classList.toggle('bad', isNaN(C.parseTime(inp.value)) || (bad && inp.dataset.f==='out')));
    $('input[data-f=out]',tr).title = bad ? 'เวลาออกต้องมากกว่าเวลาเข้า' : '';
    if (m!=null) tot += m; if (!C.isBlank(r)) cnt++;
  });
  $('#totH').textContent = Math.floor(tot/60); $('#totM').textContent = tot%60;
  const rc = $('#rowCount'); rc.textContent = `${cnt} รายการ`; rc.style.color = cnt > REPORT_CAPACITY ? 'var(--bad)' : '';
  if (cnt > REPORT_CAPACITY) rc.textContent += ` (เกิน ${REPORT_CAPACITY} — Excel จะแสดงไม่ครบ)`;
}
$('#rbody').addEventListener('change', e => {
  const el = e.target, tr = el.closest('tr'); if (!tr || !el.dataset.f || readOnly()) return;
  const r = rows[tr.dataset.d][tr.dataset.i]; let v = el.value;
  if (el.classList.contains('t')){ const p = C.parseTime(v); if (!v.trim()) v = ''; else if (!isNaN(p)){ v = C.fmtTime(p); el.value = v; } }
  r[el.dataset.f] = v; scheduleSave(); refreshComputed(); renderOthers();
});
$('#rbody').addEventListener('click', e => {
  const b = e.target.closest('button[data-a]'); if (!b || readOnly()) return;
  const tr = b.closest('tr'), d = tr.dataset.d, i = +tr.dataset.i;
  if (b.dataset.a==='add'){ rows[d].push(C.blankRow()); renderReport(); const rs=$$(`#rbody tr[data-d="${d}"]`); $('input',rs[rs.length-1]).focus(); }
  else { rows[d].splice(i,1); scheduleSave(); renderReport(); renderOthers(); }
});
$('#rbody').addEventListener('keydown', e => {
  if (e.key==='Enter' && e.target.matches('input.t')){ e.preventDefault(); e.target.blur(); const tr=e.target.closest('tr'); const nx = e.target.dataset.f==='in' ? $('input[data-f=out]',tr) : $('select[data-f=client]',tr); nx && nx.focus(); }
});
$('#bClear').onclick = async () => {
  if (!confirm(`ล้างข้อมูลทั้งหมดของ ${C.MONTHS[view.month-1]} ${view.year} (${view.name}) ?`)) return;
  rows = C.toRowsByDay([], n()); meta = {status:'draft'}; docExists = docExists; await saveNow(); applyState(); renderReport(); renderOthers();
};

/* ================= สรุปรายเดือน ================= */
function renderSheet(){
  const c = C.summarize(S, view.year, view.month, rows), nn = c.n;
  const cls = d => c.kinds[d].kind==='hol' ? 'hd' : c.kinds[d].off ? 'off' : '';
  const cell = (v,d) => { const s=C.f2(v), z=!s || +s===0; return `<td class="${cls(d)}${z?' z':''}">${z?'·':s}</td>`; };
  let h = `<table class="sum"><thead><tr><th class="lbl">DAY</th><th>Period</th>`;
  for (let d=1; d<=nn; d++) h += `<th class="${cls(d)}" title="${esc(c.kinds[d].title||'')}">${d}</th>`;
  h += `<th>TOTAL</th></tr></thead><tbody><tr class="sec"><td class="lbl" style="background:#eef1f5">CLIENT NAME (เรียกเก็บ)</td><td colspan="${nn+2}"></td></tr>`;
  if (!c.clientRows.length) h += `<tr><td class="lbl" style="color:var(--muted)">— ยังไม่มีข้อมูลลูกค้า —</td><td colspan="${nn+2}"></td></tr>`;
  const tot = v => `<td class="tot">${C.f2(v)}</td>`;
  c.clientRows.forEach(r => { h += `<tr><td class="lbl">${r.no}. ${esc(r.code)}</td><td class="per">${esc(r.period)}</td>`; for (let d=1; d<=nn; d++) h += cell(r.v[d],d); h += tot(r.total)+'</tr>'; });
  h += `<tr class="total"><td class="lbl">TOTAL CHARGE-( A )</td><td></td>`; for (let d=1; d<=nn; d++) h += `<td>${c.A[d]?C.f2(c.A[d]):''}</td>`; h += tot(c.tA)+'</tr>';
  h += `<tr class="sec"><td class="lbl" style="background:#eef1f5">NON CHARGE</td><td colspan="${nn+2}"></td></tr>`;
  c.nonRows.forEach(r => { h += `<tr><td class="lbl">${esc(r.label)}</td><td></td>`; for (let d=1; d<=nn; d++) h += cell(r.v[d],d); h += tot(r.total)+'</tr>'; });
  h += `<tr class="total"><td class="lbl">TOTAL NON CHARGE-( B )</td><td></td>`; for (let d=1; d<=nn; d++) h += `<td>${c.B[d]?C.f2(c.B[d]):''}</td>`; h += tot(c.tB)+'</tr>';
  h += `<tr><td class="lbl">OVER TIME-( C )</td><td></td>`; for (let d=1; d<=nn; d++) h += cell(c.C[d],d); h += tot(c.tC)+'</tr>';
  h += `<tr><td class="lbl">EXTRA HOURS-( D )</td><td></td>`; for (let d=1; d<=nn; d++) h += cell(c.D[d],d); h += tot(c.tD)+'</tr>';
  h += `<tr class="total"><td class="lbl">STANDARDS (A+B)-( C )-( D )</td><td></td>`; for (let d=1; d<=nn; d++) h += `<td>${c.S[d]==null?'':C.f2(c.S[d])}</td>`; h += tot(c.tS)+'</tr></tbody></table>';
  $('#sheetWrap').innerHTML = h; $('#ptSheet').textContent = titleLine('MONTHLY TIME SHEET');
}

/* ================= O.T. ================= */
function renderOt(){
  const type = $('#otFilter').value, list = C.otRows(S, view.year, view.month, rows, type);
  let tb = '', tot = 0;
  list.forEach(r => { tot += r.mins; tb += `<tr><td>${r.d}</td><td>${C.fmtTime(r.start)}</td><td>${C.fmtTime(r.end)}</td><td>${esc(r.client)}</td><td>${esc(r.service)}</td><td class="num">${Math.floor(r.mins/60)}</td><td class="num">${r.mins%60}</td><td>${esc(r.type)}</td></tr>`; });
  if (!list.length) tb = `<tr><td colspan="8" style="color:var(--muted);padding:16px">ไม่มีรายการ ${esc(type)} ในเดือนนี้ (เลือกช่อง O.T./Extra ในหน้าบันทึกเวลา)</td></tr>`;
  $('#tOt tbody').innerHTML = tb;
  const st = 'style="font-weight:700;background:var(--brand-soft)"';
  $('#tOt tfoot').innerHTML = `<tr><td colspan="5" ${st} align="right">รวม</td><td class="num" ${st}>${Math.floor(tot/60)}</td><td class="num" ${st}>${tot%60}</td><td ${st}></td></tr>`;
  $('#otStd').textContent = S.stdHours; $('#ptOt').textContent = titleLine('TIME REPORT O.T.');
}
$('#otFilter').onchange = renderOt;

/* ================= ส่งออก Excel (ฟอร์มเดิม) ================= */
let templateBuf = null;
async function getTemplate(){ if (!templateBuf){ const r = await fetch('template.xlsx'); if (!r.ok) throw new Error('ไม่พบ template.xlsx'); templateBuf = await r.arrayBuffer(); } return templateBuf; }
async function buildXlsx(doc){
  const {data, warnings} = await fillTemplate(await getTemplate(), S, doc);
  return {blob:new Blob([data], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), warnings, data};
}
async function exportCurrent(){
  try {
    await flushSave();
    const doc = currentDoc(); doc.entries = C.toEntries(rows);
    const {blob, warnings} = await buildXlsx(doc);
    download(xlsxName(doc), blob);
    toast(warnings.length ? warnings[0] : 'ดาวน์โหลดแล้ว — เปิดใน Excel แล้วกด “Enable Editing” เพื่อให้สูตรคำนวณ', warnings.length ? 6000 : 4000);
  } catch(e){ console.error(e); toast('สร้างไฟล์ Excel ไม่สำเร็จ: ' + e.message, 5000); }
}
$('#bXlsx').onclick = $('#bXlsx2').onclick = exportCurrent;
[['#bPrint',()=>$('#ptReport').textContent=titleLine('TIME REPORT')],['#bPrint2',()=>{}],['#bPrint3',()=>{}]].forEach(([s,f]) => $(s).onclick = () => { f(); window.print(); });

/* ================= พนักงาน / อนุมัติ (แอดมิน) ================= */
async function renderAdmin(){
  $('#admTitle').textContent = `ภาพรวม ${C.MONTHS[view.month-1]} ${view.year}`;
  await refreshUsers();
  let docs = []; try { docs = await be.listSheets({year:view.year}); } catch(e){ toast('โหลดข้อมูลไม่สำเร็จ: ' + (e.code||e.message)); }
  const byUid = new Map(docs.filter(d => d.month===view.month).map(d => [d.uid, d]));
  const tb = $('#tAdm tbody');
  tb.innerHTML = users.map(u => {
    const d = byUid.get(u.uid), st = d ? (d.status||'draft') : 'none';
    const hrs = d ? C.payDays(S, d).reduce((a,f) => a+f.work, 0)/60 : 0;
    const upd = d?.updatedAt ? new Date(d.updatedAt).toLocaleString('th-TH', {dateStyle:'short', timeStyle:'short'}) : '';
    return `<tr data-uid="${esc(u.uid)}"><td class="l">${esc(u.name||'')}<div class="hint">${esc(C.divLabel(S, u.division))}</div></td><td class="l">${esc(u.email||'')}</td>
      <td><select data-a="role" ${u.uid===user.uid?'disabled title="ไม่สามารถเปลี่ยนสิทธิ์ของตัวเอง"':''}><option value="user"${u.role!=='admin'?' selected':''}>พนักงาน</option><option value="admin"${u.role==='admin'?' selected':''}>แอดมิน</option></select></td>
      <td><span class="chip ${st}">${ST[st]}</span></td><td class="num">${d?C.f2(hrs):''}</td><td>${esc(upd)}</td>
      <td class="act"><button class="btn sm" data-a="open">เปิดดู/แก้ไข</button> ${d?'<button class="btn sm" data-a="xlsx">Excel</button> <button class="btn sm" data-a="appr">'+(st==='approved'?'ยกเลิกอนุมัติ':'อนุมัติ')+'</button>':''}</td></tr>`;
  }).join('') || '<tr><td colspan="7" style="padding:18px;color:var(--muted)">ยังไม่มีพนักงาน</td></tr>';
  tb._docs = byUid;
}
$('#tAdm').addEventListener('click', async e => {
  const b = e.target.closest('button[data-a]'); if (!b) return;
  const uid = b.closest('tr').dataset.uid, d = $('#tAdm tbody')._docs.get(uid), a = b.dataset.a;
  if (a==='open'){ view.uid = uid; $('#empSel').value = uid; showTab('report'); await loadView(); }
  if (a==='xlsx' && d){ try { const {blob, warnings} = await buildXlsx(d); download(xlsxName(d), blob); if (warnings.length) toast(warnings[0], 5000); } catch(err){ toast('ผิดพลาด: ' + err.message); } }
  if (a==='appr' && d){
    const ap = d.status==='approved';
    const nd = {...d, status: ap ? 'submitted' : 'approved', approvedBy: ap ? '' : (user.name||user.email), approvedAt: ap ? '' : new Date().toISOString()};
    try { await be.saveSheet(nd); toast(ap ? 'ยกเลิกการอนุมัติแล้ว' : 'อนุมัติแล้ว'); renderAdmin(); } catch(err){ toast('ผิดพลาด: ' + (err.code||err.message)); }
  }
});
$('#tAdm').addEventListener('change', async e => {
  if (e.target.dataset.a!=='role') return;
  const uid = e.target.closest('tr').dataset.uid;
  try { await be.setRole(uid, e.target.value); toast('เปลี่ยนสิทธิ์แล้ว'); } catch(err){ toast('เปลี่ยนสิทธิ์ไม่สำเร็จ: ' + (err.code||err.message)); renderAdmin(); }
});
$('#admRefresh').onclick = renderAdmin;

/* ----- สำรอง / นำเข้าข้อมูล (JSON) ----- */
$('#admBackup').onclick = async () => {
  try {
    const data = await be.exportAll();
    const pack = {app:'timesheet-backup', version:1, exportedAt:new Date().toISOString(), exportedBy:user.email || user.name, ...data};
    const day = new Date().toISOString().slice(0,10);
    download(`timesheet-backup_${day}.json`, JSON.stringify(pack, null, 1), 'application/json');
    toast(`สำรองข้อมูลแล้ว: พนักงาน ${data.users.length} คน · ใบลงเวลา ${data.sheets.length} เดือน`, 4000);
  } catch(e){ toast('สำรองข้อมูลไม่สำเร็จ: ' + (e.code||e.message), 5000); }
};
function cleanBackup(obj){
  if (!obj || obj.app !== 'timesheet-backup' || !Array.isArray(obj.users) || !Array.isArray(obj.sheets)) throw new Error('ไม่ใช่ไฟล์สำรองของระบบนี้');
  const users = obj.users.filter(u => u && typeof u.uid === 'string' && u.uid && !u.uid.includes('/')).map(u => ({...u, role: u.role === 'admin' ? 'admin' : 'user'}));
  const sheets = obj.sheets.filter(d => d && typeof d.uid === 'string' && d.uid && !d.uid.includes('/') && Number.isInteger(d.year) && Number.isInteger(d.month) && d.month >= 1 && d.month <= 12 && Array.isArray(d.entries));
  if (users.length !== obj.users.length || sheets.length !== obj.sheets.length) throw new Error(`ไฟล์มีรายการที่รูปแบบไม่ถูกต้อง (พนักงาน ${obj.users.length-users.length} · ใบลงเวลา ${obj.sheets.length-sheets.length} รายการ) จึงไม่นำเข้า`);
  const config = obj.config && typeof obj.config === 'object' ? obj.config : null;
  return {users, sheets, config};
}
$('#admRestore').onclick = () => $('#admFile').click();
$('#admFile').onchange = async e => {
  const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try {
    const pack = cleanBackup(JSON.parse(await f.text()));
    const when = pack.exportedAt || '';
    if (!confirm(`นำเข้าข้อมูลจากไฟล์ ${f.name}\nพนักงาน ${pack.users.length} คน · ใบลงเวลา ${pack.sheets.length} เดือน${pack.config ? ' · การตั้งค่า' : ''}\n\nรายการที่ซ้ำในระบบจะถูกเขียนทับด้วยข้อมูลในไฟล์ (รายการที่ไม่มีในไฟล์จะไม่ถูกลบ)\nยืนยันนำเข้า?`)) return;
    const n = await be.importAll(pack);
    toast(`นำเข้าแล้ว ${n} รายการ`, 4000);
    const cfg = await be.getConfig().catch(() => null); S = C.mergeSettings(cfg); $('#hCompany').textContent = S.company;
    repDocs = null; await loadView(); renderAdmin();
  } catch(err){ toast('นำเข้าไม่สำเร็จ: ' + (err.code || err.message), 6000); }
};
$('#admZip').onclick = async () => {
  try {
    const docs = (await be.listSheets({year:view.year})).filter(d => d.month===view.month);
    if (!docs.length) return toast('เดือนนี้ยังไม่มีข้อมูล');
    const zip = new window.JSZip(); const used = new Set();
    for (const d of docs){ const {data} = await buildXlsx(d); let nm = xlsxName(d); while (used.has(nm)) nm = nm.replace('.xlsx','_2.xlsx'); used.add(nm); zip.file(nm, data); }
    download(`TIMESHEET_${view.year}-${String(view.month).padStart(2,'0')}_ทุกคน.zip`, await zip.generateAsync({type:'blob'}));
    toast(`ดาวน์โหลด ${docs.length} ไฟล์แล้ว`);
  } catch(e){ toast('ผิดพลาด: ' + e.message, 5000); }
};

/* ================= ตั้งค่า (แอดมิน) ================= */
let curDiv = '';
const linesOf = id => [...new Set($(id).value.split('\n').map(s => s.trim()).filter(Boolean))];
const listFor = key => key ? S.divisions.find(d => d.name === key).clients : S.clients;
function commitClients(){ const l = linesOf('#sClients'); if (curDiv) S.divisions.find(d => d.name === curDiv).clients = l; else S.clients = l; }
$('#sDivSel').onchange = () => { commitClients(); curDiv = $('#sDivSel').value; $('#sClients').value = listFor(curDiv).join('\n'); };
function fillSettings(){
  $('#sCompany').value = S.company; $('#sStd').value = S.stdHours; $('#sLunch').value = S.lunchHours;
  $('#sBreaks').value = S.breaks.join('\n'); $('#sPeriods').value = S.periods.join('\n');
  $('#sDivSel').innerHTML = `<option value="">ทั่วไป (ไม่ระบุสายงาน / สายที่ยังไม่มีรายชื่อ)</option>` + S.divisions.map(d => `<option value="${esc(d.name)}">${esc(d.code + ' ' + d.name)}</option>`).join('');
  curDiv = ''; $('#sDivSel').value = ''; $('#sClients').value = S.clients.join('\n'); renderHol();
}
function renderHol(){
  const hs = [...S.holidays].sort((a,b) => a.date.localeCompare(b.date));
  $('#tHol tbody').innerHTML = hs.map(h => `<tr><td><input type="date" value="${esc(h.date)}" data-o="${esc(h.date)}" data-hf="date"></td><td><input value="${esc(h.name)}" data-o="${esc(h.date)}" data-hf="name" style="width:100%;min-width:220px;text-align:left"></td><td><button class="ic rm" data-del="${esc(h.date)}">×</button></td></tr>`).join('');
}
$('#tHol').addEventListener('click', e => { const b = e.target.closest('[data-del]'); if (!b) return; S.holidays = S.holidays.filter(h => h.date !== b.dataset.del); renderHol(); });
$('#tHol').addEventListener('change', e => { const el = e.target; if (!el.dataset.hf) return; const h = S.holidays.find(x => x.date === el.dataset.o); if (h){ h[el.dataset.hf] = el.value; renderHol(); } });
$('#bAddHol').onclick = () => { let k = C.ymd(C.ceYear(view.year), view.month, 1), d = 1; while (S.holidays.some(h => h.date===k)) k = C.ymd(C.ceYear(view.year), view.month, ++d); S.holidays.push({date:k, name:'วันหยุด'}); renderHol(); };
const lines = id => [...new Set($(id).value.split('\n').map(s => s.trim()).filter(Boolean))];
$('#bSaveSet').onclick = async () => {
  S.company = $('#sCompany').value.trim() || C.DEFAULT_SETTINGS.company; S.stdHours = +$('#sStd').value || 8; S.lunchHours = +$('#sLunch').value || 0;
  commitClients();
  if (C.allClients(S).length > 489) return toast('รายชื่อลูกค้ารวมทุกสายเกิน 489 รายการ (ข้อจำกัดของฟอร์ม Excel)', 4000);
  S.breaks = lines('#sBreaks'); S.periods = lines('#sPeriods').slice(0, 190); S.holidays = S.holidays.filter(h => h.date);
  try { await be.saveConfig(S); const el = $('#savedSet'); el.classList.add('on'); setTimeout(() => el.classList.remove('on'), 1600); $('#hCompany').textContent = S.company; renderReport(); renderOthers(); }
  catch(e){ toast('บันทึกไม่สำเร็จ: ' + (e.code||e.message), 4000); }
};

/* ================= รายงานค่าแรง / ค่าล่วงเวลา ================= */
const PAY_COLS = [['work','ชม.ทำงานรวม'], ['reg','ค่าแรงปกติ (วันปกติ ≤ 8 ชม.)'], ['ot1','ล่วงเวลา 1 เท่า (วันหยุด 8 ชม.แรก)'], ['ot15','ล่วงเวลา 1.5 เท่า (วันปกติ เกิน 8 ชม.)'], ['ot3','ล่วงเวลา 3 เท่า (วันหยุด เกิน 8 ชม.)']];
/* หัวข้อ "แสดงแยกตาม": แต่ละหัวข้อ = แถวพนักงาน (หรือลูกค้า) × คอลัมน์เดือน; "พนักงาน" = รายงานค่าแรง/ล่วงเวลา; "ทุกหัวข้อ" = ทุกหัวข้อเป็นคอลัมน์ */
const ITEMS = {
  employee:{label:'พนักงาน'}, client:{label:'รายชื่อลูกค้า'},
  office:{label:'OFFICE', f:s => s.nonRows[0].total}, sick:{label:'ลาป่วย', f:s => s.nonRows[1].total}, vacation:{label:'ลาพักร้อน', f:s => s.nonRows[2].total}, training:{label:'อบรม', f:s => s.nonRows[3].total},
  otacc:{label:'ลา OT สะสม', f:s => s.nonRows[4].total}, personal:{label:'ลากิจ', f:s => s.nonRows[5].total},
  ot:{label:'OVER TIME', f:s => s.tC}, extra:{label:'EXTRA HOUR', f:s => s.tD}, charge:{label:'TOTAL CHARGE', f:s => s.tA},
  nonch:{label:'TOTAL NON CHARGE', f:s => s.tB}, standard:{label:'STANDARD', f:s => s.tS}, all:{label:'ทุกหัวข้อ'}
};
const ALL_KEYS = ['office','sick','vacation','training','otacc','personal','ot','extra','charge','nonch','standard'];
let repDocs = null, repYear = null, lastRep = null;
function initReportControls(){
  const mo = C.MONTHS.map((m,i) => `<option value="${i+1}">${m}</option>`).join('');
  $('#rM1').innerHTML = $('#rM2').innerHTML = mo; $('#rM1').value = 1; $('#rM2').value = 12;
  $('#rGrp').innerHTML = Object.entries(ITEMS).map(([k,v]) => `<option value="${k}">${v.label}</option>`).join(''); $('#rGrp').value = 'employee'; $('#rYear').value = view.year;
  repDocs = null; lastRep = null; $('#repWrap').innerHTML = '<div style="padding:24px;color:var(--muted)">เลือกเงื่อนไขแล้วกด “สร้างรายงาน”</div>';
  $('#rEmp').innerHTML = isAdmin() ? '<option value="">ทุกคน</option>' : `<option>${esc(user.name||user.email)}</option>`;
}
function syncReportControls(){ if (!repDocs) $('#rYear').value = view.year; }
async function loadReportData(force){
  const y = +$('#rYear').value || view.year;
  if (force || !repDocs || repYear !== y){ repDocs = await be.listSheets({year:y, uid: isAdmin() ? undefined : user.uid}); repYear = y; }
  if (isAdmin()){
    const names = [...new Set(repDocs.map(d => d.name || d.uid))].sort((a,b) => a.localeCompare(b,'th'));
    const keep = $('#rEmp').value; $('#rEmp').innerHTML = `<option value="">ทุกคน</option>` + names.map(x => `<option>${esc(x)}</option>`).join(''); $('#rEmp').value = names.includes(keep) ? keep : '';
  }
}
async function runReport(){
  try { await loadReportData(false); } catch(e){ return toast('โหลดข้อมูลไม่สำเร็จ: ' + (e.code||e.message), 4000); }
  const m1 = +$('#rM1').value, m2 = +$('#rM2').value, emp = $('#rEmp').value, key = $('#rGrp').value, it = ITEMS[key];
  const docs = repDocs.filter(d => d.month>=m1 && d.month<=m2 && (!emp || (d.name||d.uid)===emp));
  const cmpS = (a,b) => String(a).localeCompare(String(b),'th');
  const months = []; for (let m=m1; m<=m2; m++) months.push(m);
  let head, cols, rows = [];            // rows: {label:[...], v:[ชั่วโมง...]}
  const sums = new Map(); docs.forEach(d => sums.set(d, C.summarize(S, d.year, d.month, C.toRowsByDay(d.entries, C.daysInMonth(d.year, d.month)))));
  if (key === 'employee'){
    head = ['พนักงาน']; cols = PAY_COLS.map(c => c[1]); const map = new Map();
    docs.flatMap(d => C.payDays(S, d)).forEach(d => { const n = d.name; let r = map.get(n); if (!r){ r = {label:[n], v:PAY_COLS.map(() => 0)}; map.set(n, r); } PAY_COLS.forEach(([f], i) => r.v[i] += d[f]/60); });
    rows = [...map.values()].sort((a,b) => cmpS(a.label[0], b.label[0]));
  } else if (key === 'client'){
    head = ['รายชื่อลูกค้า']; cols = [...months.map(m => C.MONTHS[m-1]), 'รวม']; const map = new Map();
    docs.forEach(d => sums.get(d).clientRows.forEach(c => { let r = map.get(c.code); if (!r){ r = {label:[c.code], v:months.map(() => 0)}; map.set(c.code, r); } r.v[months.indexOf(d.month)] += c.total; }));
    rows = [...map.values()].sort((a,b) => cmpS(a.label[0], b.label[0])).map(r => ({...r, v:[...r.v, r.v.reduce((x,y) => x+y, 0)]}));
  } else if (key === 'all'){
    head = ['พนักงาน']; cols = ALL_KEYS.map(k => ITEMS[k].label); const map = new Map();
    docs.forEach(d => { const n = d.name || d.uid; let r = map.get(n); if (!r){ r = {label:[n], v:ALL_KEYS.map(() => 0)}; map.set(n, r); } ALL_KEYS.forEach((k, i) => r.v[i] += ITEMS[k].f(sums.get(d)) || 0); });
    rows = [...map.values()].sort((a,b) => cmpS(a.label[0], b.label[0]));
  } else {
    head = ['พนักงาน']; cols = [...months.map(m => C.MONTHS[m-1]), 'รวม']; const map = new Map();
    docs.forEach(d => { const n = d.name || d.uid; let r = map.get(n); if (!r){ r = {label:[n], v:months.map(() => 0)}; map.set(n, r); } r.v[months.indexOf(d.month)] += it.f(sums.get(d)) || 0; });
    rows = [...map.values()].sort((a,b) => cmpS(a.label[0], b.label[0])).map(r => ({...r, v:[...r.v, r.v.reduce((x,y) => x+y, 0)]}));
  }
  const tot = cols.map((_, i) => rows.reduce((a, r) => a + r.v[i], 0));
  const hrs = v => Math.abs(v) > 1e-9 ? C.f2(v) : '·';
  let h = `<table class="piv"><thead><tr>${head.map(x => `<th class="rl">${x}</th>`).join('')}${cols.map(c => `<th>${c}</th>`).join('')}</tr></thead><tbody>`;
  rows.forEach(r => { h += `<tr>${r.label.map(x => `<td class="rl">${esc(x)}</td>`).join('')}${r.v.map(v => `<td class="n${Math.abs(v) > 1e-9 ? '' : ' z'}">${hrs(v)}</td>`).join('')}</tr>`; });
  if (!rows.length) h += `<tr><td class="rl" colspan="${head.length+cols.length}" style="color:var(--muted);padding:18px">ไม่พบข้อมูลตามเงื่อนไขนี้</td></tr>`;
  h += `</tbody><tfoot><tr><td class="rl" colspan="${head.length}" style="background:var(--brand-soft)">รวมทั้งหมด (ชม.)</td>${tot.map(v => `<td class="n">${C.f2(v)}</td>`).join('')}</tr></tfoot></table>`;
  $('#repWrap').innerHTML = h;
  const scope = `ปี ${$('#rYear').value} เดือน ${C.MONTHS[m1-1]}–${C.MONTHS[m2-1]}` + (emp ? ` · ${emp}` : '');
  const title = key === 'employee' ? 'รายงานค่าแรง / ค่าล่วงเวลา — แยกตามพนักงาน' : `รายงานวิเคราะห์ — ${it.label}`;
  lastRep = {head, cols, rows, tot, scope, title};
  $('#ptRep').textContent = `${S.company} — ${title} · ${scope}`;
  $('#repNote').textContent = `${title} · ${scope} · หน่วย: ชั่วโมง · ` + (key === 'employee' ? 'ไม่นับรายการลา · ' : 'ตัวเลขเดียวกับหน้า “สรุปรายเดือน” · ') + `จากข้อมูล ${docs.length} เอกสารรายเดือน`;
}
$('#rGo').onclick = () => { repDocs = null; runReport(); };
$('#rPrint').onclick = () => window.print();
$('#rXlsx').onclick = async () => {
  const r = lastRep; if (!r) return toast('กดสร้างรายงานก่อน');
  if (!window.ExcelJS) return toast('ไม่พบไลบรารีสร้าง Excel');
  try {
    const r2 = v => Math.round(v/60*100)/100, NUM = '#,##0.00;-#,##0.00;"–"';
    const wb = new window.ExcelJS.Workbook(); wb.creator = S.company; wb.created = new Date();
    const ws = wb.addWorksheet('REPORT', {views:[{state:'frozen', ySplit:5, showGridLines:false}]});
    const nl = r.head.length, nc = nl + r.cols.length, FONT = {name:'Calibri', size:11};
    // หัวรายงาน
    [[S.company, 16, true], [r.title, 12, true], [r.scope + ' · หน่วย: ชั่วโมง', 10, false]].forEach(([t, sz, b], i) => {
      ws.mergeCells(i+1, 1, i+1, nc); const c = ws.getCell(i+1, 1); c.value = t; c.font = {name:'Calibri', size:sz, bold:b, color:{argb: i===2 ? 'FF595959' : 'FF1F3864'}}; c.alignment = {vertical:'middle', horizontal:'left'};
    });
    ws.getRow(1).height = 24; ws.getRow(2).height = 20;
    // ตาราง
    const names = [...r.head, ...r.cols];
    const rows = r.rows.length ? r.rows.map(x => [...x.label, ...x.v.map(v => Math.round(v*100)/100)]) : [['ไม่พบข้อมูล', ...Array(nc-1).fill('').map((_, i) => i < nl-1 ? '' : 0)]];
    ws.addTable({
      name:'tblReport', ref:'A5', headerRow:true, totalsRow:true,
      style:{theme:'TableStyleMedium2', showRowStripes:true},
      columns: names.map((n, i) => i === 0 ? {name:n, totalsRowLabel:'รวมทั้งหมด (ชม.)'} : i < nl ? {name:n, totalsRowLabel:' '} : {name:n, totalsRowFunction:'sum'}),
      rows
    });
    const first = 5, last = 5 + rows.length + 1;   // ส่วนหัว + ข้อมูล + แถวรวม
    for (let rr = first; rr <= last; rr++){
      const row = ws.getRow(rr);
      for (let cc = 1; cc <= nc; cc++){
        const c = row.getCell(cc); c.font = {...FONT, bold: rr===first || rr===last};
        const hdr = rr === first, tot = rr === last, line = {style:'thin', color:{argb:'FFBDD7EE'}};
        c.border = {top:line, bottom:line, left:line, right:line};
        c.fill = {type:'pattern', pattern:'solid', fgColor:{argb: hdr ? 'FF1F4E79' : tot ? 'FFDDEBF7' : (rr - first) % 2 === 0 ? 'FFEAF1FA' : 'FFFFFFFF'}};
        if (hdr) c.font = {...FONT, bold:true, color:{argb:'FFFFFFFF'}};
        if (tot) c.border = {...c.border, top:{style:'medium', color:{argb:'FF1F4E79'}}};
        if (rr === first) c.alignment = {vertical:'middle', horizontal: cc <= nl ? 'left' : 'center', wrapText:true};
        else if (cc <= nl) c.alignment = {vertical:'middle', horizontal:'left'};
        else { c.numFmt = NUM; c.alignment = {vertical:'middle', horizontal:'right'}; }
      }
      row.height = rr === first ? (r.cols.length <= 6 ? 48 : 24) : 20;
    }
    // ผลรวมแถวสุดท้าย (ใส่ค่าคำนวณไว้ด้วย เผื่อโปรแกรมที่ไม่คำนวณสูตรทันที)
    r.cols.forEach((_, i) => { const c = ws.getCell(last, nl + 1 + i); c.value = {formula:`SUBTOTAL(109,${ws.getColumn(nl+1+i).letter}${first+1}:${ws.getColumn(nl+1+i).letter}${last-1})`, result:Math.round(r.tot[i]*100)/100}; });
    // ความกว้างคอลัมน์
    ws.columns.forEach((col, i) => { col.width = i < nl ? 30 : 15; });
    // ตั้งค่าหน้ากระดาษ A4 พอดีความกว้าง
    ws.pageSetup = {paperSize:9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight:0, horizontalCentered:true,
      margins:{left:0.4, right:0.4, top:0.6, bottom:0.6, header:0.3, footer:0.3}, printTitlesRow:'5:5', printArea:`A1:${ws.getColumn(nc).letter}${last}`};
    ws.headerFooter = {oddFooter:'&L&"Calibri,Regular"&9' + S.company.replace(/&/g, '&&') + '&R&"Calibri,Regular"&9หน้า &P / &N'};
    const buf = await wb.xlsx.writeBuffer();
    download(`REPORT_${$('#rGrp').value.toUpperCase()}_${$('#rYear').value}.xlsx`, new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  } catch(e){ toast('สร้างไฟล์ Excel ไม่สำเร็จ: ' + e.message, 5000); }
};

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
  $('#pName').value = view.name; $('#pDiv').value = view.division;
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
$('#pName').onchange = $('#pDiv').onchange = async () => {
  if (!isSelf()) return;
  view.name = $('#pName').value.trim() || view.name; view.division = $('#pDiv').value.trim();
  user.name = view.name; user.division = view.division; $('#uName').textContent = user.name;
  try { await be.saveProfile(user.uid, {name:view.name, division:view.division}); } catch(e){ toast('บันทึกโปรไฟล์ไม่สำเร็จ'); }
  if (docExists) scheduleSave(); renderOthers();
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
  return `<option value=""></option><optgroup label="งานลูกค้า / จ็อบ">${optHTML(S.clients)}</optgroup><optgroup label="งานที่ไม่เรียกเก็บ / ลา">${optHTML(C.NONCHARGE.map(x=>x[0]))}</optgroup><optgroup label="วันหยุด">${optHTML(C.DAYOFF_LABELS)}</optgroup>`;
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
    const hrs = d ? C.factsOf(S, d).reduce((a,f) => a+f.mins, 0)/60 : 0;
    const upd = d?.updatedAt ? new Date(d.updatedAt).toLocaleString('th-TH', {dateStyle:'short', timeStyle:'short'}) : '';
    return `<tr data-uid="${esc(u.uid)}"><td class="l">${esc(u.name||'')}<div class="hint">${esc(u.division||'')}</div></td><td class="l">${esc(u.email||'')}</td>
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
function fillSettings(){
  $('#sCompany').value = S.company; $('#sStd').value = S.stdHours; $('#sLunch').value = S.lunchHours;
  $('#sBreaks').value = S.breaks.join('\n'); $('#sClients').value = S.clients.join('\n'); $('#sPeriods').value = S.periods.join('\n'); renderHol();
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
  S.breaks = lines('#sBreaks'); S.clients = lines('#sClients').slice(0, 489); S.periods = lines('#sPeriods').slice(0, 190); S.holidays = S.holidays.filter(h => h.date);
  try { await be.saveConfig(S); const el = $('#savedSet'); el.classList.add('on'); setTimeout(() => el.classList.remove('on'), 1600); $('#hCompany').textContent = S.company; renderReport(); renderOthers(); }
  catch(e){ toast('บันทึกไม่สำเร็จ: ' + (e.code||e.message), 4000); }
};

/* ================= รายงานวิเคราะห์ (รายจ็อบ / รายปี / รายพนักงาน …) ================= */
const DIMS = {
  job:{label:'จ็อบ / ลูกค้า', key:f => f.client}, employee:{label:'พนักงาน', key:f => f.name}, service:{label:'Service / Period', key:f => f.service},
  month:{label:'เดือน', key:f => f.month}, category:{label:'ประเภทงาน', key:f => f.cat}, none:{label:'รวม', key:() => 'รวม'}
};
const MEAS = {
  charge:{label:'ชั่วโมงเรียกเก็บ (Charge)', val:f => f.cat.startsWith('เรียกเก็บ') ? f.mins : 0},
  noncharge:{label:'ชั่วโมง Non-charge / ลา', val:f => f.cat.startsWith('Non-charge') ? f.mins : 0},
  total:{label:'ชั่วโมงทั้งหมด', val:f => f.mins},
  ot:{label:'ชั่วโมง O.T.', val:f => f.ot==='O.T.' ? f.otMins : 0},
  extra:{label:'ชั่วโมง Extra Hours', val:f => f.ot==='Extra Hours' ? f.otMins : 0}
};
const PRESETS = [
  ['รายจ็อบ × รายเดือน', {meas:'charge', row:'job', col:'month'}],
  ['รายปี: พนักงาน × เดือน', {meas:'total', row:'employee', col:'month'}],
  ['พนักงาน × จ็อบ', {meas:'charge', row:'employee', col:'job'}],
  ['จ็อบ × พนักงาน', {meas:'charge', row:'job', col:'employee'}],
  ['จ็อบ × Period', {meas:'charge', row:'job', col:'service'}],
  ['สรุปวันลา / Non-charge', {meas:'noncharge', row:'employee', col:'job'}],
  ['O.T. พนักงาน × เดือน', {meas:'ot', row:'employee', col:'month'}],
  ['ภาพรวมรายเดือน', {meas:'total', row:'month', col:'category'}]
];
let repDocs = null, repYear = null, repFacts = [], lastPivot = null;
function initReportControls(){
  const dimOpts = keys => keys.map(k => `<option value="${k}">${DIMS[k].label}</option>`).join('');
  $('#rRow').innerHTML = dimOpts(['job','employee','service','month','category']);
  $('#rCol').innerHTML = dimOpts(['month','job','employee','service','category','none']);
  const mo = C.MONTHS.map((m,i) => `<option value="${i+1}">${m}</option>`).join('');
  $('#rM1').innerHTML = $('#rM2').innerHTML = mo; $('#rM1').value = 1; $('#rM2').value = 12;
  $('#rRow').value = 'job'; $('#rCol').value = 'month'; $('#rMeas').value = 'charge';
  $('#presets').innerHTML = PRESETS.map((p,i) => `<button data-i="${i}">${esc(p[0])}</button>`).join('');
  $('#rYear').value = view.year;
  repDocs = null; lastPivot = null; $('#repWrap').innerHTML = '<div style="padding:24px;color:var(--muted)">เลือกเงื่อนไขแล้วกด “สร้างรายงาน”</div>';
  $('#rEmp').innerHTML = isAdmin() ? '<option value="">ทุกคน</option>' : `<option>${esc(user.name||user.email)}</option>`;
  $('#rJob').innerHTML = $('#rSvc').innerHTML = '<option value="">ทั้งหมด</option>';
}
function syncReportControls(){ if (!repDocs) $('#rYear').value = view.year; }
$('#presets').onclick = e => {
  const b = e.target.closest('button'); if (!b) return; const p = PRESETS[+b.dataset.i][1];
  $('#rMeas').value = p.meas; $('#rRow').value = p.row; $('#rCol').value = p.col;
  $$('#presets button').forEach(x => x.classList.toggle('on', x===b)); runReport();
};
async function loadReportData(force){
  const y = +$('#rYear').value || view.year;
  if (force || !repDocs || repYear !== y){
    repDocs = await be.listSheets({year:y, uid: isAdmin() ? undefined : user.uid}); repYear = y;
  }
  repFacts = repDocs.flatMap(d => C.factsOf(S, d));
  if (isAdmin()){
    const names = [...new Set(repFacts.map(f => f.name))].sort((a,b) => a.localeCompare(b,'th'));
    const keep = $('#rEmp').value; $('#rEmp').innerHTML = `<option value="">ทุกคน</option>` + names.map(x => `<option>${esc(x)}</option>`).join(''); $('#rEmp').value = names.includes(keep) ? keep : '';
  }
  const fill = (id, arr, keep) => { $(id).innerHTML = `<option value="">ทั้งหมด</option>` + arr.map(x => `<option>${esc(x)}</option>`).join(''); $(id).value = arr.includes(keep) ? keep : ''; };
  fill('#rJob', [...new Set(repFacts.map(f => f.client))].sort((a,b) => a.localeCompare(b,'th',{numeric:true})), $('#rJob').value);
  fill('#rSvc', [...new Set(repFacts.map(f => f.service))].sort(), $('#rSvc').value);
}
function pivot(facts, meas, rowDim, colDim){
  const val = MEAS[meas].val, rk = DIMS[rowDim].key, ck = DIMS[colDim].key;
  const cells = new Map(), rt = new Map(), ct = new Map(); let grand = 0;
  facts.forEach(f => {
    const v = val(f); if (!v) return; const r = rk(f), c = ck(f);
    if (!cells.has(r)) cells.set(r, new Map());
    cells.get(r).set(c, (cells.get(r).get(c)||0) + v); rt.set(r, (rt.get(r)||0)+v); ct.set(c, (ct.get(c)||0)+v); grand += v;
  });
  const order = (m, dim) => [...m.keys()].sort((a,b) => dim==='month' ? a-b : (m.get(b)-m.get(a)) || String(a).localeCompare(String(b),'th',{numeric:true}));
  return {rows: order(rt,rowDim), cols: order(ct,colDim), cells, rt, ct, grand, rowDim, colDim, meas};
}
async function runReport(){
  try { await loadReportData(false); } catch(e){ return toast('โหลดข้อมูลไม่สำเร็จ: ' + (e.code||e.message), 4000); }
  const m1 = +$('#rM1').value, m2 = +$('#rM2').value, emp = $('#rEmp').value, job = $('#rJob').value, svc = $('#rSvc').value;
  const facts = repFacts.filter(f => f.month>=m1 && f.month<=m2 && (!emp || f.name===emp) && (!job || f.client===job) && (!svc || f.service===svc));
  const meas = $('#rMeas').value, rowDim = $('#rRow').value, colDim = $('#rCol').value;
  const p = lastPivot = pivot(facts, meas, rowDim, colDim);
  const lbl = (dim, k) => dim==='month' ? C.MONTHS[k-1] : k;
  const hrs = v => C.f2(v/60);
  let h = `<table class="piv"><thead><tr><th class="rl">${esc(DIMS[rowDim].label)}</th>`;
  if (colDim!=='none') p.cols.forEach(c => h += `<th>${esc(lbl(colDim,c))}</th>`);
  h += `<th>รวม (ชม.)</th></tr></thead><tbody>`;
  const max = Math.max(1, ...p.rows.map(r => p.rt.get(r)));
  p.rows.forEach(r => {
    h += `<tr><td class="rl">${esc(lbl(rowDim,r))}</td>`;
    if (colDim!=='none') p.cols.forEach(c => { const v = p.cells.get(r).get(c); h += v ? `<td class="n">${hrs(v)}</td>` : `<td class="n z">·</td>`; });
    h += `<td class="tt"><div class="bar" style="width:${Math.round(p.rt.get(r)/max*100)}%"></div><span>${hrs(p.rt.get(r))}</span></td></tr>`;
  });
  if (!p.rows.length) h += `<tr><td class="rl" colspan="${p.cols.length+2}" style="color:var(--muted);padding:18px">ไม่พบข้อมูลตามเงื่อนไขนี้</td></tr>`;
  h += `</tbody><tfoot><tr><td class="rl" style="background:var(--brand-soft)">รวมทั้งหมด</td>`;
  if (colDim!=='none') p.cols.forEach(c => h += `<td class="n">${hrs(p.ct.get(c))}</td>`);
  h += `<td class="n">${hrs(p.grand)}</td></tr></tfoot></table>`;
  $('#repWrap').innerHTML = h;
  const scope = `ปี ${$('#rYear').value} เดือน ${C.MONTHS[m1-1]}–${C.MONTHS[m2-1]}` + (emp ? ` · ${emp}` : '') + (job ? ` · จ็อบ ${job}` : '') + (svc ? ` · ${svc}` : '');
  p.title = `${MEAS[meas].label} — ${DIMS[rowDim].label} × ${DIMS[colDim].label}`; p.scope = scope;
  $('#ptRep').textContent = `${S.company} — ${p.title} · ${scope}`;
  $('#repNote').textContent = `${p.title} · ${scope} · หน่วย: ชั่วโมง · จากข้อมูล ${repDocs.length} เอกสารรายเดือน`;
}
$('#rGo').onclick = () => { $$('#presets button').forEach(x => x.classList.remove('on')); repDocs = null; runReport(); };
$('#rPrint').onclick = () => window.print();
$('#rXlsx').onclick = () => {
  const p = lastPivot; if (!p) return toast('กดสร้างรายงานก่อน');
  const lbl = (dim, k) => dim==='month' ? C.MONTHS[k-1] : k, r2 = v => Math.round(v/60*100)/100;
  const aoa = [[S.company], [p.title], [p.scope], [], [DIMS[p.rowDim].label, ...(p.colDim!=='none' ? p.cols.map(c => lbl(p.colDim,c)) : []), 'รวม (ชม.)']];
  p.rows.forEach(r => aoa.push([lbl(p.rowDim,r), ...(p.colDim!=='none' ? p.cols.map(c => r2(p.cells.get(r).get(c)||0)||'') : []), r2(p.rt.get(r))]));
  aoa.push(['รวมทั้งหมด', ...(p.colDim!=='none' ? p.cols.map(c => r2(p.ct.get(c))) : []), r2(p.grand)]);
  const X = window.XLSX, ws = X.utils.aoa_to_sheet(aoa); ws['!cols'] = [{wch:28}, ...Array(p.cols.length+1).fill({wch:12})];
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'REPORT');
  X.writeFile(wb, `REPORT_${p.meas}_${p.rowDim}x${p.colDim}_${$('#rYear').value}.xlsx`);
};

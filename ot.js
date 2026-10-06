/* ============================================================
 * ระบบแจ้งลา + โอที — ot.js — v1.0
 * ระบบโอที (OT)
 *  - พนักงาน: แจ้งโอทีเอง (เลือกวัน + ช่องชั่วโมง) → สลิปโอที → ประวัติโอที / ยกเลิก
 *  - ผู้ดูแล: ลงโอทีให้ทีเดียวหลายคน (เลือกทุกคน/ติ๊กออก) / ยืนยัน-ไม่ยืนยัน / ลบ / ส่งออก
 * ใช้ helper จาก app.js และ admin.js ($, esc, apiGet, apiPost, toast, uiDialog, T, TH, tr ฯลฯ)
 * ============================================================ */
window.APP_VER = window.APP_VER || {};
window.APP_VER.ot = '1.0';

/* ---------- ตัวช่วยร่วม ---------- */
function otSort(arr){ return arr.slice().sort((a,b)=>OT_SLOTS.indexOf(a)-OT_SLOTS.indexOf(b)); }
/* 17:00-18:00, 18:00-19:00 → 17:00–19:00 */
function otCompact(v){
  const arr=otSort((Array.isArray(v)?v:String(v||'').split(',')).map(s=>String(s).trim()).filter(s=>OT_SLOTS.indexOf(s)>=0));
  const out=[]; let st=null, en=null;
  arr.forEach(s=>{ const p=s.split('-'); if(en===p[0]) en=p[1]; else { if(st!==null) out.push(st+'–'+en); st=p[0]; en=p[1]; } });
  if(st!==null) out.push(st+'–'+en);
  return out.join(', ');
}
/* ข้อมูลของวัน (คิดในเครื่องก่อน แล้วค่อยแทนด้วยคำตอบจากระบบ) */
function otLocalInfo(ds){
  const working=isWorkDay(ds);
  return { date:ds, working_day:working, allowed: working ? OT_SLOTS.filter(s=>TIME_SLOTS.indexOf(s)<0) : OT_SLOTS.slice(), taken:[] };
}
function otDaysAgo(n){ const d=new Date(); d.setDate(d.getDate()-n); return ymd(d); }
/* วันตั้งต้นของฟอร์มโอที: ก่อนตี 5 = คนที่เพิ่งเลิกโอทีข้ามคืน → ตั้งเป็น "เมื่อวาน" ให้ */
function otDefaultDate(){ return (new Date().getHours()<5) ? otDaysAgo(1) : todayStr(); }
/* จัดกลุ่มช่องเวลา: วันทำงาน = หลังเลิกงาน / พักเที่ยง / ก่อนเข้างาน • วันหยุด = กลางวัน / หลังเลิกงาน / ก่อนเข้างาน */
function otGroups(info){
  const i=s=>OT_SLOTS.indexOf(s), ok=s=>info.allowed.indexOf(s)>=0;
  const early=OT_SLOTS.filter(s=>i(s)<3 && ok(s)), eve=OT_SLOTS.filter(s=>i(s)>=12 && ok(s));
  if(info.working_day) return [['ot_g_evening',eve],['ot_g_lunch',OT_SLOTS.filter(s=>s==='12:00-13:00' && ok(s))],['ot_g_early',early]].filter(g=>g[1].length);
  return [['ot_g_day',OT_SLOTS.filter(s=>i(s)>=3 && i(s)<12 && ok(s))],['ot_g_evening',eve],['ot_g_early',early]].filter(g=>g[1].length);
}
function otTimelineHTML(info, sel, bilingual){
  const taken=new Set(info.taken||[]);
  return otGroups(info).map(g=>{
    const label=bilingual?TH(g[0]):esc(tr(g[0],null,'th'));
    return '<div class="otGroup">'+label+'</div>'+g[1].map(s=>{
      const tm=s.replace('-','–');
      if(taken.has(s)) return '<div class="slot taken"><div class="tk">✓</div><div class="tm">'+tm+'</div><span class="lz">'+esc(bilingual?T('ot_taken'):tr('ot_taken',null,'th'))+'</span></div>';
      const on=sel.has(s);
      return '<button type="button" class="slot'+(on?' on':'')+'" data-ot="'+s+'" aria-pressed="'+on+'"><div class="tk">'+(on?'✓':'')+'</div><div class="tm">'+tm+'</div></button>';
    }).join('');
  }).join('');
}
function otStatusBadge(st, approver, lang){
  const cls={'รอยืนยัน':'ap-wait','ยืนยัน':'ap-ok','ไม่ยืนยัน':'ap-no','ยกเลิก':'ap-no'}[st]; if(!cls) return '';
  const icon={'ap-wait':'⏳','ap-ok':'✅','ap-no':'❌'}[cls];
  const txt=lang?esc(tr(OT_STATUS_KEY[st],null,lang)):TH(OT_STATUS_KEY[st]);
  return '<div class="apBadge '+cls+'">'+icon+' <span>'+txt+'</span></div>'+((approver && st!=='รอยืนยัน')?'<span class="histFiled"> · '+esc(approver)+'</span>':'');
}
function otDayTag(dayType, lang){
  const off=(dayType==='วันหยุด'), key=OT_DAY_KEY[dayType]||'ot_day_work';
  return '<span class="tag '+(off?'off':'ot')+'">'+(lang?esc(tr(key,null,lang)):TH(key))+'</span>';
}

/* ============================================================
 *  พนักงาน: แจ้งโอที
 * ============================================================ */
let otSel=new Set(), otInfo=null, otSeq=0, otPending=null, otSubmitting=false;

function otReset(){ otSel.clear(); otInfo=null; otPending=null; $('otDate').value=otDefaultDate(); $('otNote').value=''; }
function otOpen(){
  const d=$('otDate'); d.max=todayStr();
  if(!d.value || d.value>todayStr()) d.value=otDefaultDate();
  if(!d.min) d.min=otDaysAgo(7);
  otLoadDay();
}
async function otLoadDay(){
  const ds=$('otDate').value; if(!ds || !selectedEmp){ otRender(); return; }
  if(!otInfo || otInfo.date!==ds){ otInfo=otLocalInfo(ds); otSel.clear(); }
  otRender();
  const seq=++otSeq, forId=selectedEmp.emp_id;
  try{
    const r=await apiGet('otDay',{date:ds, emp_id:forId});
    if(seq!==otSeq || !selectedEmp || selectedEmp.emp_id!==forId || $('otDate').value!==ds) return;   // เปลี่ยนวัน/เปลี่ยนคนไปแล้ว
    if(r.ok){
      otInfo={date:ds, working_day:r.working_day, allowed:r.allowed||[], taken:r.taken||[]};
      if(r.min_date) $('otDate').min=r.min_date;
      [...otSel].forEach(s=>{ if(otInfo.allowed.indexOf(s)<0 || otInfo.taken.indexOf(s)>=0) otSel.delete(s); });
      otRender();
    }
  }catch(e){}
}
function otRender(){
  const ds=$('otDate').value;
  $('otDateEcho').textContent=ds?fmtDateLong(ds):'';
  if(!ds || !otInfo){ $('otDayNote').innerHTML=''; $('otTimeline').innerHTML=''; otUpdateTotal(); return; }
  $('otDayNote').innerHTML=otInfo.working_day ? '<div class="note">🕗 '+TH('ot_workday_note')+'</div>' : '<div class="note note-ot">🌤 '+TH('ot_holiday_note')+'</div>';
  $('otTimeline').innerHTML=otTimelineHTML(otInfo, otSel, true);
  $('otTimeline').querySelectorAll('[data-ot]').forEach(b=>{ b.onclick=()=>{ const v=b.dataset.ot; if(otSel.has(v)) otSel.delete(v); else otSel.add(v); otRender(); }; });
  otUpdateTotal();
}
function otUpdateTotal(){
  const n=otSel.size, el=$('otHourTotal');
  el.textContent=T('ot_total',{n:n}); el.className='hourTotal'+(n===0?' zero':' ot');
  const parts=[T('ot_word')], ds=$('otDate').value;
  if(ds) parts.push(fmtDMY(ds));
  if(n>0){ parts.push(otCompact([...otSel])); parts.push(n+' '+T('u_hours')); }
  const sm=$('otSubmitSummary'); sm.textContent=parts.join(' • '); sm.classList.toggle('l-my',LANG==='my');
}
function otOnLang(){ if(!$('paneOT').classList.contains('hidden')) otRender(); }

function otAskConfirm(){
  if(otSubmitting || submitting || !selectedEmp) return;
  const ds=$('otDate').value;
  if(!ds){ toast(T('e_pick_date'),true); return; }
  if(ds>todayStr()){ toast(T('E_OT_FUTURE'),true); return; }
  if(otSel.size===0){ toast(T('e_ot_pick_slot'),true); return; }
  const slots=otSort([...otSel]), info=otInfo||otLocalInfo(ds);
  otPending={emp_id:selectedEmp.emp_id, ot_date:ds, slots:slots, note:$('otNote').value.trim(), device:deviceInfo()};
  pendingKind='ot';
  $('csTitle').setAttribute('data-i18n','ot_confirm_title'); $('csTitle').innerHTML=TH('ot_confirm_title');
  $('confirmSummary').innerHTML=sumRow('f_name',esc(nameFor(selectedEmp)))+
    sumRow('ot_f_date',esc(fmtDateLong(ds)))+
    sumRow('f_time',esc(otCompact(slots))+'<small>'+esc(slots.length+' '+T('u_hours'))+'</small>')+
    sumRow('ot_f_daytype',TH(info.working_day?'ot_day_work':'ot_day_off'));
  $('confirmSubmitModal').classList.remove('hidden');
}
async function otDoSubmit(){
  const p=otPending, who=selectedEmp; if(!p || !who || otSubmitting) return;
  otSubmitting=true;
  const btn=$('otSubmitBtn'); btn.disabled=true; const old=btn.innerHTML; btn.innerHTML=TH('sending');
  try{
    let res;
    try{ res=await apiPost('submitOT', p); }catch(e){ throw new Error(T('e_net')); }
    if(!res.ok){ if(res.code==='OT_DUP' || res.code==='OT_WORKTIME'){ otInfo=null; otLoadDay(); } throw new Error(errText(res)); }
    selectedEmp=who; historyCache=null; lastKind='ot';
    setSuccessTexts('ot_success_title', res.status==='รอยืนยัน'?'ot_success_wait':'success_sub', 'ot_again');
    showScreen('screenSuccess');
    showSlip('successSlip', otSlipData(who, {ot_id:res.ot_id, slip_code:res.slip&&res.slip.code, ot_date:res.ot_date, range:res.range, hours:res.hours, day_type:res.day_type, status:res.status, note:p.note, filed_at:res.filed_at}));
    otSel.clear(); otInfo=null; $('otNote').value='';
  }catch(err){ toast(err.message,true); }
  otSubmitting=false; btn.disabled=false; btn.innerHTML=old; otPending=null; pendingKind='leave';
}
function otSlipData(e, o){
  return { kind:'ot', id:o.ot_id, code:o.slip_code||'', emp_id:e.emp_id, name_th:e.name_th||e.name_mm, name_mm:(e.nationality==='MM'?e.name_mm:''),
    company:'MAIN', ot_date:o.ot_date, range:o.range, hours:Number(o.hours)||0, day_type:o.day_type, status:o.status, note:o.note||'',
    filed_at:o.filed_at||'', state:(o.status==='ยกเลิก'?'cancelled':'active') };
}

/* ---------- ประวัติ: สลับ การลา / โอที ---------- */
function histSwitchHTML(res){
  const nL=(res.history||[]).length, nO=((res.ot&&res.ot.list)||[]).length;
  return '<div class="histSwitch"><button type="button" data-hv="leave" class="'+(histView==='leave'?'active':'')+'">'+TH('hist_leave')+'<i>'+nL+'</i></button>'+
    '<button type="button" data-hv="ot" class="'+(histView==='ot'?'active':'')+'">'+TH('hist_ot')+'<i>'+nO+'</i></button></div>';
}
function bindHistSwitch(){ $('historyContainer').querySelectorAll('[data-hv]').forEach(b=>{ b.onclick=()=>{ histView=b.dataset.hv; renderHistory(); }; }); }
function otHistoryHTML(res){
  const ot=res.ot||{list:[],totals:{}}, t=ot.totals||{}, list=ot.list||[];
  const card=(icon,key,val,cls)=>'<div class="useCard'+(cls?' '+cls:'')+'"><div class="ui">'+icon+'</div><div class="ut">'+TH(key)+'</div><div class="uv">'+(val||0)+' <span>'+esc(T('u_hours'))+'</span></div></div>';
  let html='<div class="useHead">'+TH('ot_cycle_head')+(ot.cycle?' <span style="font-weight:500">('+esc(fmtDMY(ot.cycle.from).slice(0,5))+' – '+esc(fmtDMY(ot.cycle.to).slice(0,5))+')</span>':'')+'</div>'+
    '<div class="useGrid">'+card('🌙','ot_day_work',t.work,'ot')+card('🌤','ot_day_off',t.off,'ot')+card('⏳','ot_u_wait',t.pending,'')+'</div>';
  if(!list.length) return html+'<div class="empEmpty">'+TH('ot_hist_empty')+'</div>';
  return html+list.map(x=>{
    const dead=(x.status==='ไม่ยืนยัน'||x.status==='ยกเลิก');
    return '<div class="histItem ot'+(x.status==='รอยืนยัน'?' wait':'')+(dead?' dead':'')+'">'+
      '<div class="histTop"><span class="histDate">'+esc(fmtDateLong(x.ot_date))+'</span>'+otDayTag(x.day_type)+'</div>'+
      '<div class="histMeta"><b>⏱ '+esc(x.range)+'</b><span>'+esc((Number(x.hours)||0)+' '+T('u_hours'))+'</span></div>'+
      (x.note?'<div class="histReason">"'+esc(x.note)+'"</div>':'')+
      otStatusBadge(x.status, x.approver)+
      '<div class="histFiled">'+esc(T('f_filed'))+' '+esc(x.filed_at)+(x.source==='ผู้ดูแลลงให้'?' · '+esc(T('ot_by_admin')):'')+'</div>'+
      '<div class="histBtns"><button class="miniBtn" type="button" data-otslip="'+esc(x.ot_id)+'">🧾 '+TH('view_slip')+'</button>'+
        (x.can_cancel?'<button class="miniBtn del" type="button" data-otcancel="'+esc(x.ot_id)+'">✕ '+TH('ot_cancel')+'</button>':'')+'</div></div>';
  }).join('');
}
function otBindHistory(){
  const box=$('historyContainer'), find=id=>((historyCache&&historyCache.ot&&historyCache.ot.list)||[]).find(r=>r.ot_id===id);
  box.querySelectorAll('[data-otslip]').forEach(b=>{ b.onclick=()=>{ const x=find(b.dataset.otslip); if(!x) return; $('slipModal').classList.remove('hidden'); showSlip('modalSlip', otSlipData(selectedEmp, x)); }; });
  box.querySelectorAll('[data-otcancel]').forEach(b=>{ b.onclick=async()=>{
    const x=find(b.dataset.otcancel); if(!x) return;
    const c=await uiDialog({titleHTML:TH('ot_cancel_title'), bodyHTML:'<b>'+esc(fmtDateLong(x.ot_date))+'</b><br>'+esc(x.range)+' · '+esc(x.hours+' '+T('u_hours')), okHTML:TH('cancel_confirm'), cancelHTML:TH('cancel_back'), danger:true});
    if(!c.ok) return;
    try{
      const r=await apiPost('cancelOT',{emp_id:selectedEmp.emp_id, ot_id:x.ot_id});
      if(r.ok){ toast(T('ot_cancelled_ok')); loadHistory(); } else { toast((r.code==='NOT_FOUND'||r.code==='NOT_YOURS')?T('e_generic'):errText(r),true); loadHistory(); }
    }catch(e){ toast(T('e_net'),true); }
  }; });
}

/* ---------- หน้าตรวจสลิปโอที ---------- */
function otVerifyHTML(res, id){
  const head='<h2 class="vTitle">'+TB('v_title_ot')+'</h2>';
  if(!res.found) return head+'<div class="vResult warn"><div class="vIcon">!</div><div>'+TB('v_gone_ot')+'</div></div><div class="vNo">'+esc(id)+'</div>';
  const items=(res.items||[]).map(it=>{
    const st=it.status, dead=(st==='ไม่ยืนยัน'||st==='ยกเลิก');
    return '<div class="vItem'+(dead?' cancelled':'')+'">'+
      '<div class="vDate">'+esc(fmtDateLong(it.ot_date,'th'))+'</div>'+
      '<div class="vMeta">โอที <span class="l-my">အိုတီ</span> · '+esc(it.range)+' · '+esc(String(it.hours))+' ชม. · '+esc(tr(OT_DAY_KEY[it.day_type]||'ot_day_work',null,'th'))+'</div>'+
      '<div class="vState '+(dead?'c':(st==='รอยืนยัน'?'w':'a'))+'">'+TB(OT_STATUS_KEY[st]||'ot_st_wait')+'</div>'+
      ((it.approver && st==='ยืนยัน')?'<div class="vMeta" style="margin-top:4px">ยืนยันโดย '+esc(it.approver)+'</div>':'')+'</div>';
  }).join('');
  return head+'<div class="vResult good"><div class="vIcon">✓</div><div>'+TB('v_real')+'</div></div>'+
    '<div class="vWho"><b>'+esc(res.name)+'</b><span>'+esc(res.emp_id)+' · '+esc(res.company)+'</span></div>'+items+
    '<div class="vNo">'+esc(id)+' · '+esc(tr('v_checked_at',{t:res.checked_at},'th'))+'</div>';
}

/* ============================================================
 *  ผู้ดูแล: ลงโอทีให้หลายคน + รายการโอที
 * ============================================================ */
let aotSel=new Set(), aotPicked=new Set(), aotInfo=null, aotTakenAll={}, aotSeq=0;
let adminOTs=[], otPeople=[], otTotals={}, otRange=null, otStFilter='ALL', otListSeq=0;

function aotOpen(){
  const d=$('aotDate'); d.max=todayStr(); d.min=otDaysAgo(62);
  if(!d.value) d.value=otDefaultDate();
  if(!$('otMonth').value) $('otMonth').value=currentCycleMonth(adminCutoff);
  aotLoadDay(); loadOTAdmin();
}
async function aotLoadDay(){
  const ds=$('aotDate').value; if(!ds) return;
  if(!aotInfo || aotInfo.date!==ds){ aotInfo=otLocalInfo(ds); aotSel.clear(); aotTakenAll={}; }
  aotRender();
  const seq=++aotSeq;
  try{
    const r=await apiGet('otDay',{date:ds, all:'1', pin:adminPin});
    if(seq!==aotSeq || $('aotDate').value!==ds) return;
    if(r.ok){
      aotInfo={date:ds, working_day:r.working_day, allowed:r.allowed||[], taken:[]};
      aotTakenAll=r.taken_all||{};
      [...aotSel].forEach(s=>{ if(aotInfo.allowed.indexOf(s)<0) aotSel.delete(s); });
      aotRender();
    }
  }catch(e){}
}
/* คนนี้ลงโอทีช่วงที่เลือกไว้แล้วหรือยัง */
function aotClash(id){ const t=aotTakenAll[id]; return !!t && [...aotSel].some(s=>t.indexOf(s)>=0); }
function aotRender(){
  if(!aotInfo) return;
  $('aotDayNote').innerHTML='<div class="note'+(aotInfo.working_day?'':' note-ot')+'" style="margin-top:0">'+
    esc(fmtDateLong(aotInfo.date,'th'))+' — '+esc(tr(aotInfo.working_day?'ot_workday_note':'ot_holiday_note',null,'th'))+'</div>';
  $('aotTimeline').innerHTML=otTimelineHTML(aotInfo, aotSel, false);
  $('aotTimeline').querySelectorAll('[data-ot]').forEach(b=>{ b.onclick=()=>{ const v=b.dataset.ot; if(aotSel.has(v)) aotSel.delete(v); else aotSel.add(v); aotRender(); }; });
  const n=aotSel.size, el=$('aotHourTotal');
  el.textContent=n?(otCompact([...aotSel])+' · '+n+' ชม.'):'ยังไม่ได้เลือกเวลา'; el.className='hourTotal'+(n===0?' zero':' ot');
  aotRenderList();
}
function aotSorted(){
  const key=e=>(e.nationality==='MM'?'1':'0')+String(e.emp_id).replace(/\d+/g,n=>n.padStart(6,'0'));
  return adminEmployees.slice().sort((a,b)=>key(a)<key(b)?-1:1);
}
function aotRenderList(){
  const box=$('aotList');
  if(!adminEmployees.length){ box.innerHTML='<button class="empEmpty retry" type="button" id="aotReload" style="margin:0; width:100%">กำลังโหลดรายชื่อ… (แตะเพื่อโหลดใหม่)</button>'; $('aotReload').onclick=()=>loadAdminEmployees(); $('aotCount').textContent=''; return; }
  [...aotPicked].forEach(id=>{ if(aotClash(id)) aotPicked.delete(id); });      // คนที่ลงช่วงนี้ไว้แล้ว เลือกซ้ำไม่ได้
  const q=$('aotSearch').value.trim().toLowerCase();
  let list=aotSorted();
  if(q) list=list.filter(e=>(e.name_th||'').toLowerCase().includes(q)||(e.name_mm||'').toLowerCase().includes(q)||String(e.emp_id||'').toLowerCase().includes(q));
  box.innerHTML=list.length?list.map(e=>{
    const id=esc(e.emp_id), dis=aotClash(e.emp_id), sel=aotPicked.has(e.emp_id), had=aotTakenAll[e.emp_id];
    return '<label class="pickRow'+(sel?' sel':'')+(dis?' dis':'')+'"><input type="checkbox" data-pick="'+id+'"'+(sel?' checked':'')+(dis?' disabled':'')+'>'+
      '<b>'+natFlag(e.nationality)+' '+esc(e.name_th||e.name_mm)+'</b><small>'+id+(had?(' · ลงแล้ว '+esc(otCompact(had))):'')+'</small></label>';
  }).join(''):'<div class="empEmpty" style="margin:0">ไม่พบพนักงาน</div>';
  box.querySelectorAll('[data-pick]').forEach(c=>{ c.onchange=()=>{ if(c.checked) aotPicked.add(c.dataset.pick); else aotPicked.delete(c.dataset.pick); c.parentNode.classList.toggle('sel',c.checked); aotCount(); }; });
  aotCount();
}
function aotCount(){
  const n=aotPicked.size, h=aotSel.size;
  $('aotCount').textContent='เลือก '+n+' คน'+(n&&h?(' · รวม '+(n*h)+' ชม.'):'');
  $('aotSave').textContent=n?('บันทึกโอที '+n+' คน'):'บันทึกโอที';
}
function aotPick(mode){
  if(mode==='none'){ aotPicked.clear(); aotRenderList(); return; }
  aotPicked.clear();
  adminEmployees.forEach(e=>{
    if(aotClash(e.emp_id)) return;
    if(mode==='all' || (mode==='TH' && e.nationality!=='MM') || (mode==='MM' && e.nationality==='MM')) aotPicked.add(e.emp_id);
  });
  aotRenderList();
  toast('เลือกแล้ว '+aotPicked.size+' คน — ติ๊กออกคนที่ไม่ได้ทำได้เลย');
}
async function aotSave(){
  const ds=$('aotDate').value;
  if(!ds){ toast('เลือกวันที่ก่อน',true); return; }
  if(!aotSel.size){ toast('เลือกช่วงเวลาที่ทำโอทีก่อน',true); return; }
  if(!aotPicked.size){ toast('เลือกพนักงานอย่างน้อย 1 คน',true); return; }
  const slots=otSort([...aotSel]), ids=aotSorted().filter(e=>aotPicked.has(e.emp_id)).map(e=>e.emp_id);
  const names=aotSorted().filter(e=>aotPicked.has(e.emp_id)).map(e=>esc(e.name_th||e.name_mm));
  const c=await uiDialog({title:'ลงโอที '+ids.length+' คน ?',
    bodyHTML:'<b>'+esc(fmtDateLong(ds,'th'))+'</b> · '+esc(tr(aotInfo&&!aotInfo.working_day?'ot_day_off':'ot_day_work',null,'th'))+'<br><b>'+esc(otCompact(slots))+'</b> ('+slots.length+' ชม./คน · รวม '+(slots.length*ids.length)+' ชม.)<br><br><span style="font-size:12.5px">'+names.slice(0,12).join(', ')+(names.length>12?' … และอีก '+(names.length-12)+' คน':'')+'</span>',
    okText:'บันทึก', cancelText:'ยกเลิก'});
  if(!c.ok) return;
  const btn=$('aotSave'); btn.disabled=true; btn.textContent='กำลังบันทึก...';
  try{
    const r=await apiPost('adminAddOT',{pin:adminPin, ot_date:ds, slots:slots, emp_ids:ids, note:$('aotNote').value.trim()});
    if(!r.ok) throw new Error(r.error||'บันทึกไม่สำเร็จ');
    toast('ลงโอทีแล้ว '+r.added+' คน ✓');
    if(r.skipped && r.skipped.length) uiDialog({title:'ข้าม '+r.skipped.length+' คน', bodyHTML:r.skipped.map(x=>esc((x.name||x.emp_id)+' — '+x.reason)).join('<br>'), alert:true});
    aotPicked.clear(); aotSel.clear(); $('aotNote').value=''; aotInfo=null;
    aotLoadDay(); loadOTAdmin();
  }catch(err){ toast(err.message||'เชื่อมต่อไม่ได้',true); }
  btn.disabled=false; aotCount();
}

/* ---------- รายการโอที (ผู้ดูแล) ---------- */
async function loadOTAdmin(){
  $('otContent').innerHTML='<div class="loading"><div class="spinner"></div>กำลังโหลด...</div>';
  const seq=++otListSeq;
  try{
    const res=await apiGet('otAdmin',{pin:adminPin, month:$('otMonth').value, cycle:$('otCycle').checked?'1':''});
    if(seq!==otListSeq) return;                        // มีคำขอใหม่กว่าแล้ว (เปลี่ยนเดือนเร็ว ๆ)
    if(!res.ok) throw new Error(res.error||'load failed');
    adminOTs=res.ots||[]; otPeople=res.people||[]; otTotals=res.totals||{}; otRange=res.range||null;
    renderOTAdmin();
  }catch(err){ if(seq!==otListSeq) return; $('otSummary').innerHTML=''; $('otStFilter').innerHTML=''; $('otContent').innerHTML='<div class="empEmpty">⚠️ '+esc(err.message||'โหลดไม่สำเร็จ')+'</div>'; }
}
function otStMatch(f, x){ return f==='ALL' || (f==='WAIT'?x.status==='รอยืนยัน':(f==='OK'?x.status==='ยืนยัน':(x.status==='ไม่ยืนยัน'||x.status==='ยกเลิก'))); }
function visibleOTs(){
  const q=$('otSearch').value.trim().toLowerCase();
  let list=adminOTs.filter(x=>otStMatch(otStFilter,x));
  if(q) list=list.filter(x=>(x.name||'').toLowerCase().includes(q)||String(x.emp_id||'').toLowerCase().includes(q));
  return list;
}
function renderOTAdmin(){
  const t=otTotals||{}, waitN=adminOTs.filter(x=>x.status==='รอยืนยัน').length;
  let sum=(otRange?'<div class="rangeNote">📆 '+esc(otRange.label)+' · '+esc(fmtDMY(otRange.from))+' – '+esc(fmtDMY(otRange.to))+'</div>':'')+
    '<div class="statGrid">'+
      '<div class="statCard ot"><div class="sv">'+(t.total||0)+'</div><div class="sl">รวมโอที (ชม.)</div></div>'+
      '<div class="statCard ot"><div class="sv">'+(t.work||0)+'</div><div class="sl">โอทีวันทำงาน (ชม.)</div></div>'+
      '<div class="statCard ot"><div class="sv">'+(t.off||0)+'</div><div class="sl">ทำงานวันหยุด (ชม.)</div></div>'+
      (t.pending?'<div class="statCard"><div class="sv" style="color:var(--amber)">'+t.pending+'</div><div class="sl">รอยืนยัน (ชม.) ยังไม่นับ</div></div>':'')+
    '</div>';
  if(waitN) sum+='<button class="btn btn-primary btn-block btn-ot" id="otConfirmAll" type="button" style="margin-bottom:12px">✅ ยืนยันทั้งหมดที่รอ ('+waitN+' รายการ)</button>';
  const ppl=otPeople.filter(p=>p.total||p.pending);
  if(ppl.length) sum+='<details class="moreTools"'+(ppl.length<=12?' open':'')+'><summary>ยอดรายคน ('+ppl.length+' คน)</summary><div style="overflow-x:auto; margin-top:8px"><table class="otSumTable"><thead><tr><th>พนักงาน</th><th>วันทำงาน</th><th>วันหยุด</th><th>รวม</th><th>รอ</th></tr></thead><tbody>'+
    ppl.map(p=>'<tr><td>'+natFlag(p.nationality)+' '+esc(p.name)+' <small style="color:var(--muted)">'+esc(p.emp_id)+'</small></td><td>'+p.work+'</td><td>'+p.off+'</td><td class="t">'+p.total+'</td><td>'+(p.pending||'')+'</td></tr>').join('')+
    '</tbody><tfoot><tr><td>รวม</td><td>'+(t.work||0)+'</td><td>'+(t.off||0)+'</td><td class="t">'+(t.total||0)+'</td><td>'+(t.pending||'')+'</td></tr></tfoot></table></div></details>';
  $('otSummary').innerHTML=sum;
  if($('otConfirmAll')) $('otConfirmAll').onclick=confirmAllOT;
  const cnt={ALL:adminOTs.length, WAIT:waitN, OK:adminOTs.filter(x=>x.status==='ยืนยัน').length, NO:adminOTs.filter(x=>x.status==='ไม่ยืนยัน'||x.status==='ยกเลิก').length};
  $('otStFilter').innerHTML=adminOTs.length?[['ALL','ทั้งหมด'],['WAIT','รอยืนยัน'],['OK','ยืนยันแล้ว'],['NO','ไม่นับ']].map(d=>'<button type="button" class="fChip'+(otStFilter===d[0]?' active':'')+'" data-f="'+d[0]+'">'+d[1]+'<i>'+cnt[d[0]]+'</i></button>').join(''):'';
  $('otStFilter').querySelectorAll('.fChip').forEach(b=>{ b.onclick=()=>{ otStFilter=b.dataset.f; renderOTAdmin(); }; });
  const list=visibleOTs();
  if(!list.length){ $('otContent').innerHTML='<div class="empEmpty">ไม่มีรายการโอทีในช่วงนี้</div>'; return; }
  $('otContent').innerHTML='<div class="desc" style="margin-bottom:10px">ทั้งหมด '+list.length+' รายการ</div>'+list.map(x=>{
    const id=esc(x.ot_id), dead=(x.status==='ไม่ยืนยัน'||x.status==='ยกเลิก');
    const btns=(x.status==='รอยืนยัน'?'<button class="miniBtn" style="background:var(--green-bg);color:var(--green)" type="button" data-otst="ยืนยัน" data-id="'+id+'">✅ ยืนยัน</button><button class="miniBtn hide" type="button" data-otst="ไม่ยืนยัน" data-id="'+id+'">❌ ไม่ได้ทำ</button>':
      (x.status==='ยืนยัน'?'<button class="miniBtn hide" type="button" data-otst="ไม่ยืนยัน" data-id="'+id+'">เปลี่ยนเป็นไม่ยืนยัน</button>':
      (x.status==='ไม่ยืนยัน'?'<button class="miniBtn" type="button" data-otst="ยืนยัน" data-id="'+id+'">เปลี่ยนเป็นยืนยัน</button>':'')))+
      '<button class="miniBtn del" type="button" data-otdel="'+id+'">ลบ</button>';
    return '<div class="histItem ot'+(x.status==='รอยืนยัน'?' wait':'')+(dead?' dead':'')+'">'+
      '<div class="histTop"><span class="histDate">'+esc(fmtDateLong(x.ot_date,'th'))+'</span>'+otDayTag(x.day_type,'th')+'</div>'+
      '<div class="histMeta"><span><b>'+natFlag(x.nationality)+' '+esc(x.name)+'</b> ('+esc(x.emp_id)+')</span></div>'+
      '<div class="histMeta" style="margin-top:3px"><b>⏱ '+esc(x.range)+'</b><span>'+x.hours+' ชม.</span><span>'+esc(x.source)+'</span></div>'+
      (x.note?'<div class="histReason">"'+esc(x.note)+'"</div>':'')+
      otStatusBadge(x.status, x.approver, 'th')+
      '<div class="histFiled">บันทึกเมื่อ '+esc(x.filed_at)+'</div>'+
      '<div style="display:flex;gap:8px;margin-top:10px;flex-wrap:wrap">'+btns+'</div></div>';
  }).join('');
  $('otContent').querySelectorAll('[data-otst]').forEach(b=>{ b.onclick=()=>{ const x=adminOTs.find(o=>o.ot_id===b.dataset.id); setOTStatus([b.dataset.id], b.dataset.otst, x?x.status:''); }; });
  $('otContent').querySelectorAll('[data-otdel]').forEach(b=>{ b.onclick=()=>deleteOTAdmin(b.dataset.otdel); });
}
/* expect = สถานะที่หน้าจอเห็นอยู่ — ถ้ามีคนเปลี่ยนไปก่อนแล้ว (เช่น พนักงานยกเลิก / หัวหน้ากดใน Telegram) ระบบจะข้ามรายการนั้น */
async function setOTStatus(ids, status, expect){
  try{
    const r=await apiPost('updateOTStatus',{pin:adminPin, ot_ids:ids, status:status, expect:expect||''});
    if(!r.ok) throw new Error(r.error||'ผิดพลาด');
    if(r.changed) toast((status==='ยืนยัน'?'ยืนยันแล้ว ':'เปลี่ยนสถานะแล้ว ')+r.changed+' รายการ ✓');
    if(r.skipped && r.skipped.length){
      const nm=id=>{ const x=adminOTs.find(o=>o.ot_id===id); return x?(x.name+' '+fmtDMY(x.ot_date)+' '+x.range):id; };
      uiDialog({title:'ไม่ได้เปลี่ยน '+r.skipped.length+' รายการ', bodyHTML:r.skipped.slice(0,12).map(x=>esc(nm(x.ot_id)+' — '+x.reason)).join('<br>'), alert:true});
    }
    loadOTAdmin();
  }catch(err){ toast(err.message||'เชื่อมต่อไม่ได้',true); }
}
async function confirmAllOT(){
  const ids=adminOTs.filter(x=>x.status==='รอยืนยัน').map(x=>x.ot_id); if(!ids.length) return;
  const c=await uiDialog({title:'ยืนยันโอทีทั้งหมด '+ids.length+' รายการ?', body:'ทุกรายการที่ "รอยืนยัน" ในช่วงที่แสดงอยู่ จะถูกนับเป็นโอทีจริง', okText:'ยืนยันทั้งหมด', cancelText:'ยกเลิก'});
  if(c.ok) setOTStatus(ids,'ยืนยัน','รอยืนยัน');
}
async function deleteOTAdmin(id){
  const x=adminOTs.find(o=>o.ot_id===id);
  const c=await uiDialog({title:'ลบรายการโอทีนี้?', body:(x?(x.name+' · '+fmtDMY(x.ot_date)+' · '+x.range+'\n\n'):'')+'ลบถาวร ใช้เมื่อลงผิดเท่านั้น', okText:'ลบถาวร', cancelText:'ไม่ลบ', danger:true});
  if(!c.ok) return;
  try{
    const r=await apiPost('deleteOT',{pin:adminPin, ot_id:id});
    if(r.ok){ toast('ลบแล้ว'); loadOTAdmin(); aotInfo=null; aotLoadDay(); } else toast(r.error||'ผิดพลาด',true);
  }catch(err){ toast('เชื่อมต่อไม่ได้',true); }
}
function exportOTCSV(){
  const list=visibleOTs();
  if(!list.length){ toast('ไม่มีข้อมูลให้ส่งออก',true); return; }
  const head=['วันที่ทำโอที','รหัสพนักงาน','ชื่อ','สัญชาติ','เวลา','ชั่วโมง','ประเภทวัน','สถานะ','ผู้ยืนยัน','ที่มา','หมายเหตุ','บันทึกเมื่อ'];
  const rows=list.slice().sort((a,b)=>a.ot_date<b.ot_date?-1:(a.ot_date>b.ot_date?1:(a.emp_id<b.emp_id?-1:1))).map(x=>[fmtDMY(x.ot_date),x.emp_id,x.name,x.nationality,x.range,x.hours,x.day_type,x.status,x.approver||'',x.source||'',x.note||'',x.filed_at]);
  const escCsv=v=>{ v=String(v==null?'':v); if(/^[=+\-@]/.test(v)) v="'"+v; return /[",\n]/.test(v)?('"'+v.replace(/"/g,'""')+'"'):v; };
  const csv=[head].concat(rows).map(r=>r.map(escCsv).join(',')).join('\r\n');
  const blob=new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url; a.download=APP_TAG+'_OT_'+($('otMonth').value||'all')+($('otCycle').checked?'_cycle':'')+'.csv';
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),4000);
  toast('ดาวน์โหลดไฟล์แล้ว ✓');
}
function saveOtSettings(){ return withBtn('saveOtSetBtn','บันทึก...',async()=>{
  const d=$('setOtBack').value.trim();
  if(!(/^\d+$/.test(d) && +d>=0 && +d<=62)) throw new Error('จำนวนวันต้องเป็นเลข 0–62');
  await setOne('ot_back_days',d); await setOne('ot_need_confirm',$('setOtConfirm').checked?'1':'0');
  toast('บันทึกการตั้งค่าโอทีแล้ว ✓');
}); }

/* ---------- ผูกปุ่ม ---------- */
$('stOT').onclick=()=>switchSubTab('ot');
$('otDate').onchange=()=>{ otInfo=null; otLoadDay(); };
$('otSubmitBtn').onclick=otAskConfirm;
$('atOT').onclick=()=>switchAdminTab('ot');
$('aotDate').onchange=()=>{ aotInfo=null; aotLoadDay(); };
$('aotAll').onclick=()=>aotPick('all');
$('aotTH').onclick=()=>aotPick('TH');
$('aotMM').onclick=()=>aotPick('MM');
$('aotNone').onclick=()=>aotPick('none');
$('aotSearch').oninput=debounce(aotRenderList,160);
$('aotSave').onclick=aotSave;
$('otMonth').onchange=loadOTAdmin;
$('otCycle').onchange=loadOTAdmin;
$('otSearch').oninput=debounce(renderOTAdmin,160);
$('otExportBtn').onclick=exportOTCSV;
$('saveOtSetBtn').onclick=saveOtSettings;

console.log('[leave] ot.js v' + window.APP_VER.ot);

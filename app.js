/* ============================================================
 * ระบบแจ้งลา + โอที — app.js — v1.0
 * ฝั่งพนักงาน: เลือกคน / ยืนยัน / แจ้งลา / สลิป / ประวัติ / ตรวจสลิป  (ส่วนโอทีอยู่ใน ot.js)
 * (มี helper ที่ admin.js ใช้ร่วม: $, esc, apiGet, apiPost, toast, uiDialog,
 *  avatarHTML, avatarColor, initials, imgErr, fmtDMY, debounce)
 * ============================================================ */
window.APP_VER = window.APP_VER || {};
window.APP_VER.app = '1.0';

/* ---------- State ---------- */
let allEmployees = [], currentCo = 'MAIN', currentNat = 'TH', selectedEmp = null;
let selectedSlots = new Set(), selectedType = 'ลากิจ', override = false;
let dayMode = 'single';          // 'single' | 'range'
let holidayMap = {};             // 'yyyy-MM-dd' -> ชื่อวันหยุด
let workOverrideMap = {};        // 'yyyy-MM-dd' -> วันทำงานพิเศษ (สลับวันหยุด)
let pendingPayload = null, submitting = false;
let pendingKind = 'leave';       // 'leave' | 'ot' — ป๊อปอัพยืนยันใช้ร่วมกัน
let lastKind = 'leave';          // สิ่งที่เพิ่งส่งสำเร็จ (ปุ่ม "แจ้งอีกครั้ง" จะกลับไปแท็บเดิม)
let histView = 'leave';          // หน้าประวัติ: 'leave' | 'ot'
let historyCache = null;         // ผลล่าสุดของหน้าประวัติ (ใช้ render ใหม่ตอนสลับภาษา)
let currentSlip = null;          // {data, url} สลิปที่กำลังแสดง

/* ---------- Helpers (ใช้ร่วมกับ admin.js) ---------- */
const $ = id => document.getElementById(id);
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function todayStr(){const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
function avatarColor(name){let h=0;for(let i=0;i<(name||'').length;i++)h=name.charCodeAt(i)+((h<<5)-h);return 'hsl('+(Math.abs(h)%360)+',48%,46%)';}
function initials(name){const n=(name||'?').trim();return Array.from(n)[0]||'?';}
function toast(msg,isErr){const t=$('toast');t.textContent=msg;t.className='toast show'+(isErr?' err':'');clearTimeout(t._t);t._t=setTimeout(()=>{t.className='toast';},isErr?4200:2600);}
function debounce(fn,ms){let t;return function(){const a=arguments,c=this;clearTimeout(t);t=setTimeout(()=>fn.apply(c,a),ms);};}
function sizedUrl(u,size){return (u&&size)?u.replace(/=w\d+$/,'=w'+size):u;}
function fmtDMY(s){ const m=/^(\d{4})-(\d{2})-(\d{2})/.exec(s||''); return m?(m[3]+'/'+m[2]+'/'+m[1]):(s||''); }
function parseYMD(s){ const m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s||''); return m?new Date(+m[1],+m[2]-1,+m[3]):null; }
function ymd(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
/* วันที่แบบอ่านง่าย: "อังคาร 06/10/2026" (ตามภาษาที่เลือก) */
function fmtDateLong(s, lang){ const d=parseYMD(s); if(!d) return s||''; return WD[lang||LANG][d.getDay()]+' '+fmtDMY(s); }
function sortSlots(arr){ return arr.slice().sort((a,b)=>TIME_SLOTS.indexOf(a)-TIME_SLOTS.indexOf(b)); }
/* ย่อช่วงเวลาที่ติดกัน: 08:00-09:00, 09:00-10:00 → 08:00–10:00 */
function compactSlots(v){
  const arr=sortSlots((Array.isArray(v)?v:String(v||'').split(',')).map(s=>String(s).trim()).filter(s=>TIME_SLOTS.indexOf(s)>=0));
  if(!arr.length) return String(Array.isArray(v)?v.join(', '):(v||''));
  const out=[]; let st=null, en=null;
  arr.forEach(s=>{ const p=s.split('-'); if(en===p[0]) en=p[1]; else { if(st) out.push(st+'–'+en); st=p[0]; en=p[1]; } });
  if(st) out.push(st+'–'+en);
  return out.join(', ');
}
/* ข้อความช่วงเวลา: ครบ 8 ชม. = "เต็มวัน" */
function slotsText(slots, hours, lang){ return (Number(hours)>=8) ? tr('full_day',null,lang||LANG) : compactSlots(slots); }
function deviceInfo(){
  const ua=navigator.userAgent||'';
  let os='อื่นๆ';
  if(/iPhone|iPad|iPod/i.test(ua))os='iPhone/iPad'; else if(/Android/i.test(ua))os='Android';
  else if(/Windows/i.test(ua))os='Windows'; else if(/Mac/i.test(ua))os='Mac'; else if(/Linux/i.test(ua))os='Linux';
  let br='อื่นๆ';
  if(/Line\//i.test(ua))br='LINE'; else if(/Edg/i.test(ua))br='Edge'; else if(/CriOS|Chrome/i.test(ua))br='Chrome';
  else if(/FxiOS|Firefox/i.test(ua))br='Firefox'; else if(/Safari/i.test(ua))br='Safari';
  return os+' · '+br;
}

/* ---------- API ---------- */
const API = (function(){
  // ผู้ดูแลตั้งค่าทดสอบได้ (เฉพาะ URL ของ Apps Script): localStorage.setItem('lv_api','https://script.google.com/macros/s/.../exec')
  try{ const o=localStorage.getItem('lv_api'); if(o && /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(o)) return o; }catch(e){}
  return (typeof API_URL==='string') ? API_URL.trim() : '';
})();
/* อ่านข้อมูล (GET): Apps Script ตอบพลาดเป็นครั้งคราวตอนเพิ่งตื่น → ลองใหม่เองอีก 2 ครั้งก่อนจะแจ้งว่าโหลดไม่ได้ */
async function apiGet(action,params){
  const q=new URLSearchParams(Object.assign({action:action},params||{})).toString();
  let lastErr;
  for(let i=0;i<3;i++){
    try{ const r=await fetch(API+'?'+q); return await r.json(); }
    catch(e){ lastErr=e; if(i<2) await new Promise(res=>setTimeout(res,700*(i+1))); }
  }
  throw lastErr;
}
async function apiPost(action,data){
  const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(Object.assign({action:action},data||{}))});
  return r.json();
}

/* ---------- กล่องถาม/ยืนยัน (แทน alert/confirm/prompt) → Promise<{ok,value,date}> ---------- */
function uiDialog(o){
  o=o||{};
  return new Promise(resolve=>{
    $('dlgTitle').innerHTML=o.titleHTML||esc(o.title||'');
    $('dlgBody').innerHTML=o.bodyHTML||esc(o.body||'').replace(/\n/g,'<br>');
    const inp=$('dlgInput'), dt=$('dlgDate');
    inp.classList.toggle('hidden',!o.input); inp.value=o.value||''; inp.placeholder=o.placeholder||'';
    dt.classList.toggle('hidden',!o.date); dt.value=o.dateValue||'';
    const okB=$('dlgOk'), cB=$('dlgCancel');
    okB.innerHTML=o.okHTML||esc(o.okText||T('ok')); cB.innerHTML=o.cancelHTML||esc(o.cancelText||T('cancel'));
    okB.className='btn '+(o.danger?'btn-danger':'btn-primary');
    cB.classList.toggle('hidden',!!o.alert);
    $('dlgBtns').classList.toggle('one',!!o.alert);
    const m=$('dlgModal'); m.classList.remove('hidden');
    if(o.input) setTimeout(()=>inp.focus(),80);
    const done=(ok)=>{ m.classList.add('hidden'); okB.onclick=cB.onclick=m.onclick=null; resolve({ok:ok,value:inp.value.trim(),date:dt.value}); };
    okB.onclick=()=>{
      if(o.input && o.required && !inp.value.trim()){ toast(o.requiredMsg||T('e_generic'),true); inp.focus(); return; }
      done(true);
    };
    cB.onclick=()=>done(false);
    m.onclick=e=>{ if(e.target===m && !o.input) done(false); };
  });
}

/* รูปพนักงาน + fallback เป็นอักษรย่อ (ถ้าโหลดรูปไม่ได้) */
function imgErr(img){
  const name=img.getAttribute('data-nm')||'';
  const d=document.createElement('div');
  d.className=img.className;
  d.style.background=avatarColor(name);
  d.textContent=initials(name);
  if(img.id) d.id=img.id;
  img.replaceWith(d);
}
function avatarHTML(cls,name,photoUrl,id,size){
  const idAttr=id?(' id="'+id+'"'):'';
  const url=sizedUrl(photoUrl,size);
  if(url) return '<img class="'+cls+'"'+idAttr+' src="'+esc(url)+'" alt="" loading="lazy" decoding="async" referrerpolicy="no-referrer" data-nm="'+esc(name)+'" onerror="imgErr(this)">';
  return '<div class="'+cls+'"'+idAttr+' style="background:'+avatarColor(name)+'">'+esc(initials(name))+'</div>';
}

function coOf(e){ return e.company==='MAIN'?'MAIN':'ALT'; }
function nameFor(e){return e.nationality==='MM'?(e.name_mm||e.name_th):(e.name_th||e.name_mm);}
function subFor(e){return (e.nationality==='MM' && e.name_mm && e.name_th)?e.name_th:'';}

/* ---------- โหลด + แสดงรายชื่อ (แสดงของที่จำไว้ทันที แล้วค่อยอัปเดตจากระบบ) ---------- */
function readCache(k){ try{ return JSON.parse(localStorage.getItem(k)||'null'); }catch(e){ return null; } }
function writeCache(k,v){ try{ localStorage.setItem(k,JSON.stringify(v)); }catch(e){} }

async function loadEmployees(showToast){
  if(!API){      // ยังไม่ได้ใส่ URL หลังบ้านใน config.js
    $('empContainer').innerHTML='<div class="empEmpty">⚙️ ยังไม่ได้ตั้งค่าระบบ<br><small>ใส่ URL ของ Apps Script (ลงท้าย /exec) ในไฟล์ <b>config.js</b> บรรทัด <b>API_URL</b> แล้วเปิดหน้านี้ใหม่</small></div>';
    return;
  }
  const cached=readCache('lv_emp');
  if(cached && Array.isArray(cached.list) && cached.list.length && !allEmployees.length){ allEmployees=cached.list; afterEmployees(); }
  else if(!allEmployees.length) $('empContainer').innerHTML='<div class="loading"><div class="spinner"></div>'+TB('loading_names')+'</div>';
  loadHolidays();
  try{
    const res=await apiGet('employees');
    if(!res.ok) throw new Error(res.error||'load failed');
    allEmployees=res.employees||[];
    writeCache('lv_emp',{t:Date.now(),list:allEmployees});
    afterEmployees();
    if(showToast) toast(T('refreshed'));
  }catch(err){
    if(allEmployees.length){ if(showToast) toast(T('e_net'),true); return; }
    $('empContainer').innerHTML='<button class="empEmpty retry" id="retryLoad" type="button">⚠️ '+TB('load_failed')+'</button>';
    $('retryLoad').onclick=()=>loadEmployees(true);
  }
}
function afterEmployees(){
  const nB=allEmployees.filter(e=>coOf(e)==='ALT').length, nK=allEmployees.length-nB;
  const both=(nB>0 && nK>0);
  $('coTabs').classList.toggle('hidden',!both);
  if(!both) currentCo = nK>0 ? 'MAIN' : 'ALT';      // มีบริษัทเดียว → ไม่ต้องให้เลือก
  $('cntALT').textContent=tr('n_people',{n:nB},'th');
  $('cntMAIN').textContent=tr('n_people',{n:nK},'th');
  setTabs();
  renderEmployees();
}
function setTabs(){
  $('coALT').classList.toggle('active',currentCo==='ALT');
  $('coMAIN').classList.toggle('active',currentCo==='MAIN');
  $('tabTH').classList.toggle('active',currentNat==='TH');
  $('tabMM').classList.toggle('active',currentNat==='MM');
  const inCo=allEmployees.filter(e=>coOf(e)===currentCo);
  $('cntTH').textContent=tr('n_people',{n:inCo.filter(e=>e.nationality!=='MM').length},'th');
  $('cntMM').textContent=tr('n_people',{n:inCo.filter(e=>e.nationality==='MM').length},'my');
}
function renderEmployees(){
  if(!allEmployees.length) return;
  const q=$('searchInput').value.trim().toLowerCase();
  let list;
  if(q){   // ค้นหา = หาทุกกลุ่ม ไม่ต้องสลับแท็บเอง
    list=allEmployees.filter(e=>(e.name_th||'').toLowerCase().includes(q)||(e.name_mm||'').toLowerCase().includes(q)||String(e.emp_id||'').toLowerCase().includes(q));
  } else {
    list=allEmployees.filter(e=>coOf(e)===currentCo && (currentNat==='MM'?e.nationality==='MM':e.nationality!=='MM'));
  }
  if(list.length===0){$('empContainer').innerHTML='<div class="empEmpty">'+TB('not_found')+(q?' "'+esc(q)+'"':'')+'</div>';return;}
  $('empContainer').innerHTML='<div class="empGrid">'+list.map(e=>{
    const nm=nameFor(e), sub=subFor(e);
    return '<button class="empCard" type="button" data-id="'+esc(e.emp_id)+'">'+avatarHTML('avatar',nm,e.photo_url,'',160)+
      '<div class="nm'+(e.nationality==='MM'&&e.name_mm?' l-my':'')+'">'+esc(nm)+(sub?'<small>'+esc(sub)+'</small>':'')+'</div>'+
      (q?'<span class="idTag">'+esc(e.emp_id)+'</span>':'')+'</button>';
  }).join('')+'</div>';
  $('empContainer').querySelectorAll('.empCard').forEach(c=>{c.onclick=()=>openConfirm(c.dataset.id);});
}

/* ---------- ยืนยันตัวตน (กันกดผิด) ---------- */
function openConfirm(empId){
  const e=allEmployees.find(x=>x.emp_id===empId); if(!e) return;
  selectedEmp=e;
  const nm=nameFor(e), sub=subFor(e);
  $('cmAvatar').outerHTML=avatarHTML('bigAvatar',nm,e.photo_url,'cmAvatar',300);
  $('cmName').innerHTML=esc(nm)+(sub?'<small>'+esc(sub)+'</small>':'')+'<small class="idLine">'+esc(e.emp_id)+'</small>';
  $('confirmModal').classList.remove('hidden');
}

/* ---------- ไปหน้าแจ้งลา ---------- */
function showScreen(id){
  ['screenVerify','screenSelect','screenForm','screenSuccess'].forEach(s=>$(s).classList.toggle('hidden',s!==id));
  window.scrollTo(0,0);
}
function gotoForm(){
  $('confirmModal').classList.add('hidden');
  const e=selectedEmp; if(!e) return;
  if((e.nationality==='MM')!==(LANG==='my')) setLang(e.nationality==='MM'?'my':'th');   // ภาษาตามคนที่เลือก
  showScreen('screenForm');
  const nm=nameFor(e);
  $('pbAvatar').outerHTML=avatarHTML('pa',nm,e.photo_url,'pbAvatar',120);
  $('pbName').textContent=nm;
  $('pbSub').textContent=e.emp_id+(subFor(e)?' · '+subFor(e):'');
  historyCache=null; histView='leave';
  resetForm(); if(typeof otReset==='function') otReset();
  switchSubTab('leave');
}
function resetForm(){
  selectedSlots.clear(); selectedType='ลากิจ'; override=false; pendingPayload=null;
  $('leaveDate').value=todayStr(); $('reasonInput').value='';
  $('dateFrom').value=todayStr(); $('dateTo').value=todayStr();
  renderTypeChips();
  markQuick(null);
  dayMode='single';
  applyTypeMode();
  renderTimeline(); updateHourTotal(); updateStatus();
}
function renderTypeChips(){
  const mm=selectedEmp && selectedEmp.nationality==='MM';   // พม่า: ยังไม่เปิดใช้ลาพักร้อน
  $('leaveTypeChips').querySelectorAll('.chip').forEach(c=>{
    c.classList.toggle('active',c.dataset.type===selectedType);
    const lock=(c.dataset.type==='ลาพักร้อน' && mm);
    c.classList.toggle('disabled',lock);
    c.disabled=lock;
  });
}

/* ---------- ไทม์ไลน์เวลา ---------- */
const AM_SLOTS=TIME_SLOTS.slice(0,4);   // 08:00-12:00
const PM_SLOTS=TIME_SLOTS.slice(4,8);   // 13:00-17:00
function markQuick(which){ ['btnFull','btnAM','btnPM'].forEach(id=>$(id).classList.toggle('active', id===which)); }
function setQuick(slots,which){
  const same=(sortSlots([...selectedSlots]).join()===slots.join());
  selectedSlots.clear();
  if(!same) slots.forEach(s=>selectedSlots.add(s));     // กดซ้ำ = ยกเลิกที่เลือก
  syncQuickHighlight(); renderTimeline(); updateHourTotal();
}
function syncQuickHighlight(){
  const arr=sortSlots([...selectedSlots]).join(',');
  if(arr===TIME_SLOTS.join(',')) markQuick('btnFull');
  else if(arr===AM_SLOTS.join(',')) markQuick('btnAM');
  else if(arr===PM_SLOTS.join(',')) markQuick('btnPM');
  else markQuick(null);
}
function renderTimeline(){
  $('timeline').innerHTML=DAY_BLOCKS.map(b=>{
    if(b.lunch) return '<div class="slot lunch"><div class="tk">🍴</div><div class="tm">'+b.label+'</div><span class="lz">'+esc(T('lunch'))+'</span></div>';
    const on=selectedSlots.has(b.v);
    return '<button type="button" class="slot'+(on?' on':'')+'" data-slot="'+b.v+'" aria-pressed="'+on+'"><div class="tk">'+(on?'✓':'')+'</div><div class="tm">'+b.v.replace('-','–')+'</div></button>';
  }).join('');
  $('timeline').querySelectorAll('.slot[data-slot]').forEach(s=>{
    s.onclick=()=>{const v=s.dataset.slot; if(selectedSlots.has(v))selectedSlots.delete(v); else selectedSlots.add(v); renderTimeline(); updateHourTotal(); syncQuickHighlight();};
  });
}

/* ---------- ประเภทลา → ปรับโหมด (พักร้อน = ช่วงวัน เต็มวันเสมอ) ---------- */
function applyTypeMode(){
  const isVaca = (selectedType==='ลาพักร้อน');
  $('timeBlock').classList.toggle('hidden', isVaca);
  $('dayModeToggle').classList.toggle('hidden', isVaca);
  $('vacaNote').classList.toggle('hidden', !isVaca);
  $('reasonNum').textContent = isVaca ? '3' : '4';
  if(isVaca){
    selectedSlots.clear(); TIME_SLOTS.forEach(s=>selectedSlots.add(s)); // นับเต็มวัน
    setDayMode('range');
  } else setDayMode(dayMode);
}
function updateHourTotal(){
  const n=selectedSlots.size, el=$('hourTotal');
  el.textContent=n>=8?T('total_full'):T('total_hours',{n:n});
  el.className='hourTotal'+(n===0?' zero':'');
  updateSubmitSummary();
}

/* ---------- วันหยุดบริษัท + วันทำงานพิเศษ ---------- */
async function loadHolidays(){
  const c=readCache('lv_hol'); if(c){ holidayMap=c.h||{}; workOverrideMap=c.w||{}; renderCalNotice(); }
  try{ const r=await apiGet('holidays'); if(r.ok){
    holidayMap={}; (r.holidays||[]).forEach(h=>{holidayMap[h.date]=h.name;});
    workOverrideMap={}; (r.work_overrides||[]).forEach(w=>{workOverrideMap[w.date]=w.name||'วันทำงานพิเศษ';});
    writeCache('lv_hol',{h:holidayMap,w:workOverrideMap});
    if(!$('screenForm').classList.contains('hidden')) updateStatus();
  } }catch(e){}
  renderCalNotice();
}
/* ประกาศหน้าแรก: วันหยุดพิเศษ / วันอาทิตย์ที่ต้องมาทำงาน ใน 14 วันข้างหน้า (สองภาษา) */
function renderCalNotice(){
  const box=$('calNotice'); if(!box) return;
  const t=todayStr(), end=new Date(); end.setDate(end.getDate()+14); const lim=ymd(end), rows=[];
  Object.keys(workOverrideMap).forEach(d=>{ if(d>=t && d<=lim) rows.push({d:d, work:true}); });
  Object.keys(holidayMap).forEach(d=>{ if(d>=t && d<=lim && !workOverrideMap[d]) rows.push({d:d, work:false}); });
  rows.sort((a,b)=>a.d<b.d?-1:1);
  box.classList.toggle('hidden', rows.length===0);
  box.innerHTML=rows.map(r=>{
    const dt=parseYMD(r.d), dm=fmtDMY(r.d).slice(0,5), key=r.work?'cal_work':'cal_off';
    return '<div class="calRow '+(r.work?'w':'h')+'"><span class="ci">'+(r.work?'🟢':'🔴')+'</span><div><b>'+esc(tr(key,{d:WD.th[dt.getDay()]+' '+dm},'th'))+'</b>'+
      '<span class="l-my">'+esc(tr(key,{d:WD.my[dt.getDay()]+' '+dm},'my'))+'</span></div></div>';
  }).join('');
}
/* วันนี้ต้องมาทำงานไหม (ตรงกับ isWorkingDay_ ฝั่งหลังบ้าน) */
function isWorkDay(ds){ if(workOverrideMap[ds]) return true; if(holidayMap[ds]) return false; const d=parseYMD(ds); return d && d.getDay()!==0; }

/* ---------- โหมดวันเดียว/หลายวัน ---------- */
function setDayMode(m){
  dayMode=m;
  $('modeSingle').classList.toggle('active',m==='single');
  $('modeRange').classList.toggle('active',m==='range');
  $('singleDateWrap').classList.toggle('hidden',m!=='single');
  $('rangeDateWrap').classList.toggle('hidden',m!=='range');
  updateStatus();
}

/* ---------- สถานะการแจ้ง + วันอาทิตย์/วันหยุด ---------- */
function filingKey(ds){ const t=todayStr(); return ds>t?'st_adv':(ds===t?'st_same':'st_back'); }
function statusBadge(ds){
  const k=filingKey(ds), cls={st_adv:'st-adv',st_same:'st-same',st_back:'st-back'}[k];
  return '<div class="statusBadge '+cls+'"><span class="st-dot"></span><span>'+TH(k)+'</span></div>';
}
function overrideBox(key,p){
  return '<div class="sundayBox"><div class="sw-txt">'+TH(key,p)+'</div><label class="switch"><input type="checkbox" id="daySwitch"'+(override?' checked':'')+'><span class="slider"></span></label></div>';
}
function updateStatus(){
  const sw=$('statusWrap'), suw=$('sundayWrap');
  if(dayMode==='range'){
    override=false; suw.innerHTML='';
    const a=$('dateFrom').value, b=$('dateTo').value;
    if(a && b && b>=a){
      let n=0, m=0; const d=parseYMD(a), end=parseYMD(b); let guard=0;
      while(d<=end && guard++<400){ if(isWorkDay(ymd(d))) n++; else m++; d.setDate(d.getDate()+1); }
      sw.innerHTML=statusBadge(a)+'<div class="rangeCount">'+TH('range_preview',{n:n,m:m})+'</div>';
    } else sw.innerHTML='';
    updateSubmitSummary(); return;
  }
  const d=$('leaveDate').value;
  $('dateEcho').textContent=d?fmtDateLong(d):'';
  if(!d){sw.innerHTML='';suw.innerHTML='';updateSubmitSummary();return;}
  sw.innerHTML=statusBadge(d);
  const dt=parseYMD(d), hol=holidayMap[d], isWork=!!workOverrideMap[d];
  const needSwitch = !isWork && (hol || dt.getDay()===0);
  if(!needSwitch) override=false;
  if(isWork){
    suw.innerHTML='<div class="note note-green">🔁 '+TH('workday_box')+'</div>';
  } else if(hol){
    suw.innerHTML=overrideBox('holiday_box',{name:hol});
  } else if(dt.getDay()===0){
    suw.innerHTML=overrideBox('sunday_box');
  } else suw.innerHTML='';
  if($('daySwitch')) $('daySwitch').onchange=ev=>{override=ev.target.checked;};
  updateSubmitSummary();
}
/* บรรทัดสรุปเหนือปุ่มส่ง: ประเภท • วันที่ • ชั่วโมง */
function updateSubmitSummary(){
  const el=$('submitSummary'); if(!el) return;
  const n=selectedSlots.size, parts=[typeLabel(selectedType)];
  if(dayMode==='range'){ const a=$('dateFrom').value, b=$('dateTo').value; if(a&&b) parts.push(a===b?fmtDMY(a):T('range_fmt',{a:fmtDMY(a).slice(0,5),b:fmtDMY(b)})); }
  else if($('leaveDate').value) parts.push(fmtDMY($('leaveDate').value));
  if(n>0) parts.push(n>=8?T('full_day'):(n+' '+T('u_hours')));
  el.textContent=parts.join(' • ');
  el.classList.toggle('l-my',LANG==='my');
}

/* ---------- แท็บย่อย ---------- */
function switchSubTab(which){
  $('stLeave').classList.toggle('active',which==='leave');
  $('stOT').classList.toggle('active',which==='ot');
  $('stHistory').classList.toggle('active',which==='history');
  $('paneLeave').classList.toggle('hidden',which!=='leave');
  $('paneOT').classList.toggle('hidden',which!=='ot');
  $('paneHistory').classList.toggle('hidden',which!=='history');
  if(which==='history') loadHistory();
  if(which==='ot' && typeof otOpen==='function') otOpen();
  window.scrollTo(0,0);
}

/* ---------- ส่งใบลา (มีป๊อปอัพยืนยันก่อน) ---------- */
function sumRow(key,valHTML){ return '<div class="row"><span class="k">'+TH(key)+'</span><span class="v">'+valHTML+'</span></div>'; }
function buildLeavePayload(){
  if(selectedSlots.size===0){toast(T('e_pick_slot'),true);return null;}
  const slots=sortSlots([...selectedSlots]);
  const reason=$('reasonInput').value.trim();
  const nm=nameFor(selectedEmp);
  const typeHTML='<span'+(LANG==='my'?' class="l-my"':'')+'>'+esc(typeLabel(selectedType))+'</span>';
  if(dayMode==='range'){
    const from=$('dateFrom').value, to=$('dateTo').value;
    if(!from||!to){toast(T('e_pick_range'),true);return null;}
    if(to<from){toast(T('E_END_BEFORE_START'),true);return null;}
    const sm=sumRow('f_name',esc(nm))+
      sumRow('f_type',typeHTML)+
      sumRow('f_range',esc(T('range_fmt',{a:fmtDMY(from),b:fmtDMY(to)})))+
      sumRow('f_time_each',esc(slotsText(slots,slots.length))+'<small>'+esc(slots.length+' '+T('u_hours'))+'</small>');
    return {payload:{emp_id:selectedEmp.emp_id, date_from:from, date_to:to, slots:slots, leave_type:selectedType, reason:reason, device:deviceInfo()}, summaryHTML:sm};
  }
  const date=$('leaveDate').value;
  if(!date){toast(T('e_pick_date'),true);return null;}
  const dt=parseYMD(date);
  if(!workOverrideMap[date]){   // วันทำงานพิเศษ (สลับวันหยุด) = ส่งได้เลย ไม่ต้องเปิดสวิตช์
    if(holidayMap[date] && !override){toast(T('E_HOLIDAY',{name:holidayMap[date]}),true);return null;}
    if(dt.getDay()===0 && !override){toast(T('E_SUNDAY'),true);return null;}
  }
  const sm=sumRow('f_name',esc(nm))+
    sumRow('f_type',typeHTML)+
    sumRow('f_date',esc(fmtDateLong(date)))+
    sumRow('f_time',esc(slotsText(slots,slots.length))+'<small>'+esc(slots.length+' '+T('u_hours'))+'</small>')+
    sumRow('f_status',TH(filingKey(date)));
  return {payload:{emp_id:selectedEmp.emp_id, leave_date:date, slots:slots, leave_type:selectedType, reason:reason, override:override?'1':'', device:deviceInfo()}, summaryHTML:sm};
}
function askConfirmSubmit(){
  if(submitting || (typeof otSubmitting!=='undefined' && otSubmitting)) return;
  const b=buildLeavePayload(); if(!b)return;
  pendingPayload=b.payload; pendingKind='leave';
  $('csTitle').setAttribute('data-i18n','confirm_title'); $('csTitle').innerHTML=TH('confirm_title');
  $('confirmSummary').innerHTML=b.summaryHTML;
  $('confirmSubmitModal').classList.remove('hidden');
}
async function doSubmit(){
  $('confirmSubmitModal').classList.add('hidden');
  if(pendingKind==='ot' && typeof otDoSubmit==='function') return otDoSubmit();
  const p=pendingPayload, who=selectedEmp; if(!p||!who||submitting)return;
  submitting=true;
  const btn=$('submitBtn'); btn.disabled=true; const old=btn.innerHTML; btn.innerHTML=TH('sending');
  try{
    let res;
    try{ res=await apiPost('submitLeave', p); }catch(e){ throw new Error(T('e_net')); }
    if(!res.ok) throw new Error(errText(res));
    showSuccess(res,p,who);
  }catch(err){toast(err.message,true);}
  submitting=false; btn.disabled=false; btn.innerHTML=old; pendingPayload=null;
}

/* ---------- หน้าสำเร็จ + สลิป ---------- */
function slipFromSubmit(res,p,e){
  const multi=!!res.multi;
  const dates=multi?(res.created||[]).map(c=>c.date):[p.leave_date];
  const slots=p.slots||[];
  return {
    id:(res.slip&&res.slip.id)||res.leave_id||'', code:(res.slip&&res.slip.code)||'',
    emp_id:e.emp_id, name_th:e.name_th||e.name_mm, name_mm:(e.nationality==='MM'?e.name_mm:''),
    company:res.company||coOf(e), leave_type:p.leave_type, dates:dates, slots:slots, hours_each:slots.length,
    filing_status:multi?((res.created&&res.created[0]&&res.created[0].filing_status)||''):res.filing_status,
    approval:res.pending_approval?'รออนุมัติ':'', filed_at:res.filed_at||'', skipped:(res.skipped||[]).length, state:'active'
  };
}
async function showSlip(holderId, data){
  const holder=$(holderId);
  holder.innerHTML='<div class="loading"><div class="spinner"></div>'+TH('slip_making')+'</div>';
  try{
    const url=await renderSlip(data);
    currentSlip={data:data,url:url};
    holder.innerHTML='<img class="slipImg" alt="'+esc(T('slip_title'))+'" src="'+url+'">';
  }catch(err){
    currentSlip=null;
    holder.innerHTML='<div class="empEmpty">⚠️ '+esc(T('e_generic'))+'</div>';
  }
}
function showSuccess(res,p,who){
  selectedEmp=who;                    // เผื่อกด "เปลี่ยนคน" ระหว่างรอ — ใบลาส่งสำเร็จแล้ว ต้องโชว์สลิปของคนที่ส่ง
  historyCache=null; lastKind='leave';
  setSuccessTexts('success_title','success_sub','again');
  showScreen('screenSuccess');
  showSlip('successSlip', slipFromSubmit(res,p,who));
}
function setSuccessTexts(titleKey, subKey, againKey){
  [['successTitle',titleKey],['successSub',subKey],['againBtn',againKey]].forEach(x=>{ $(x[0]).setAttribute('data-i18n',x[1]); $(x[0]).innerHTML=TH(x[1]); });
}
function slipFileName(d){ return (d.kind==='ot'?'ot-slip_':'leave-slip_')+String(d.emp_id).replace(/[^\w-]/g,'')+'_'+(d.kind==='ot'?(d.ot_date||''):((d.dates&&d.dates[0])||''))+'.png'; }
function dataUrlToBlob(u){ const p=u.split(','), bin=atob(p[1]), a=new Uint8Array(bin.length); for(let i=0;i<bin.length;i++)a[i]=bin.charCodeAt(i); return new Blob([a],{type:'image/png'}); }
async function shareSlip(){
  if(!currentSlip) return;
  const file=new File([dataUrlToBlob(currentSlip.url)], slipFileName(currentSlip.data), {type:'image/png'});
  try{
    if(navigator.canShare && navigator.canShare({files:[file]})){ await navigator.share({files:[file], title:T('slip_title')}); return; }
  }catch(e){ if(e && e.name==='AbortError') return; }
  if(inAppBrowser()){ uiDialog({titleHTML:TH('share_slip'), bodyHTML:TH('longpress_hint'), alert:true}); return; }
  saveSlip();
}
function inAppBrowser(){ return /Line\/|FBAN|FBAV|FB_IAB|Instagram|MicroMessenger|Telegram/i.test(navigator.userAgent||''); }
function isIOS(){ return /iPhone|iPad|iPod/i.test(navigator.userAgent||'') || (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1); }
async function saveSlip(){
  if(!currentSlip) return;
  // เบราว์เซอร์ในแอป (LINE ฯลฯ) ดาวน์โหลดไฟล์ไม่ได้ → ให้กดค้างที่รูปแทน
  if(inAppBrowser()){ uiDialog({titleHTML:TH('save_slip'), bodyHTML:TH('longpress_hint'), alert:true}); return; }
  if(isIOS()){        // iPhone: เปิดหน้าแชร์ของเครื่อง → เลือก "บันทึกรูปภาพ" เข้าอัลบั้มได้เลย
    const file=new File([dataUrlToBlob(currentSlip.url)], slipFileName(currentSlip.data), {type:'image/png'});
    try{ if(navigator.canShare && navigator.canShare({files:[file]})){ await navigator.share({files:[file]}); return; } }
    catch(e){ if(e && e.name==='AbortError') return; }
    uiDialog({titleHTML:TH('save_slip'), bodyHTML:TH('longpress_hint'), alert:true}); return;
  }
  const a=document.createElement('a');
  a.href=URL.createObjectURL(dataUrlToBlob(currentSlip.url)); a.download=slipFileName(currentSlip.data);
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href),4000);
  toast(T('slip_saved'));
}

/* ---------- ประวัติของฉัน (+ สลิป + ยกเลิกใบลา) ---------- */
async function loadHistory(){
  if(!historyCache) $('historyContainer').innerHTML='<div class="loading"><div class="spinner"></div>'+TH('loading')+'</div>';
  else renderHistory();
  const forId=selectedEmp.emp_id;
  try{
    const res=await apiGet('history',{emp_id:forId});
    if(!selectedEmp || selectedEmp.emp_id!==forId) return;      // เปลี่ยนคนไปแล้วระหว่างรอ → ทิ้งผลนี้
    if(!res.ok) throw new Error(errText(res));
    historyCache=res; renderHistory();
  }catch(err){ if(selectedEmp && selectedEmp.emp_id===forId && !historyCache) $('historyContainer').innerHTML='<div class="empEmpty">⚠️ '+esc(err.message||T('e_net'))+'</div>'; }
}
function usageHTML(u){
  if(!u) return '';
  const H=u.hours_per_day||8, d=h=>Math.round(h/H*100)/100;
  const card=(icon,key,hours,quota)=>{
    let extra='';
    if(quota>0){ const left=Math.round((quota-hours/H)*100)/100; extra='<small class="'+(left<0?'over':'')+'">'+TH(left<0?'quota_over':'quota_left',{n:Math.abs(left)})+'</small>'; }
    return '<div class="useCard"><div class="ui">'+icon+'</div><div class="ut">'+TH(key)+'</div><div class="uv">'+d(hours)+' <span>'+esc(T('u_days'))+'</span></div>'+extra+'</div>';
  };
  return '<div class="useHead">'+TH('usage_year')+'</div><div class="useGrid">'+
    card('🤒','type_sick',u.sick_hours||0,u.sick_quota_days||0)+
    card('📝','type_personal',u.personal_hours||0,u.personal_quota_days||0)+
    card('🌴','type_vacation',u.vacation_hours||0,0)+'</div>';
}
function approvalBadge(v){
  const cls={'รออนุมัติ':'ap-wait','อนุมัติ':'ap-ok','ไม่อนุมัติ':'ap-no'}[v];
  if(!cls) return '';
  const icon={'ap-wait':'⏳','ap-ok':'✅','ap-no':'❌'}[cls];
  return '<div class="apBadge '+cls+'">'+icon+' <span>'+TH(APPROVAL_KEY[v])+'</span></div>';
}
function renderHistory(){
  const res=historyCache; if(!res) return;
  const h=res.history||[];
  const sw=(typeof histSwitchHTML==='function')?histSwitchHTML(res):'';
  if(histView==='ot' && typeof otHistoryHTML==='function'){ $('historyContainer').innerHTML=sw+otHistoryHTML(res); bindHistSwitch(); otBindHistory(); return; }
  let html=sw+usageHTML(res.usage);
  if(h.length===0){ $('historyContainer').innerHTML=html+'<div class="empEmpty">'+TH('history_empty')+'</div>'; if(sw) bindHistSwitch(); return; }
  const t=todayStr();
  const item=x=>{
    const fs=String(x.filing_status||'');
    const c=fs.indexOf('ย้อนหลัง')>=0?'back':(fs.indexOf('กระทันหัน')>=0?'same':'adv');
    const stKey=STATUS_KEY[fs];
    return '<div class="histItem '+c+(x.approval_status==='ไม่อนุมัติ'?' rejected':'')+'">'+
      '<div class="histTop"><span class="histDate">'+esc(fmtDateLong(x.leave_date))+'</span><span class="tag '+c+'">'+(stKey?TH(stKey):esc(fs))+'</span></div>'+
      '<div class="histMeta"><b>'+(TYPE_KEY[x.leave_type]?TH(TYPE_KEY[x.leave_type]):esc(x.leave_type))+'</b><span>⏰ '+esc(slotsText(x.slots,x.hours))+'</span><span>'+esc((Number(x.hours)||0)+' '+T('u_hours'))+'</span></div>'+
      (x.reason?'<div class="histReason">"'+esc(x.reason)+'"</div>':'')+
      approvalBadge(x.approval_status)+
      '<div class="histFiled">'+esc(T('f_filed'))+' '+esc(x.filed_at)+'</div>'+
      '<div class="histBtns">'+
        '<button class="miniBtn" type="button" data-slip="'+esc(x.leave_id)+'">🧾 '+TH('view_slip')+'</button>'+
        (x.can_cancel?'<button class="miniBtn del" type="button" data-cancel="'+esc(x.leave_id)+'">✕ '+TH('cancel_leave')+'</button>':'')+
      '</div></div>';
  };
  const up=h.filter(x=>x.leave_date>=t).reverse(), past=h.filter(x=>x.leave_date<t);
  if(up.length) html+='<div class="histGroup">'+TH('upcoming')+'</div>'+up.map(item).join('');
  if(past.length) html+='<div class="histGroup">'+TH('past')+'</div>'+past.map(item).join('');
  $('historyContainer').innerHTML=html;
  if(sw) bindHistSwitch();
  $('historyContainer').querySelectorAll('[data-cancel]').forEach(b=>{b.onclick=()=>cancelMyLeave(b.dataset.cancel);});
  $('historyContainer').querySelectorAll('[data-slip]').forEach(b=>{b.onclick=()=>openHistorySlip(b.dataset.slip);});
}
function openHistorySlip(leaveId){
  const x=(historyCache&&historyCache.history||[]).find(r=>r.leave_id===leaveId); if(!x) return;
  const e=selectedEmp;
  const slots=String(x.slots||'').split(',').map(s=>s.trim()).filter(Boolean);
  $('slipModal').classList.remove('hidden');
  showSlip('modalSlip',{ id:x.leave_id, code:x.slip_code||'', emp_id:e.emp_id, name_th:e.name_th||e.name_mm, name_mm:(e.nationality==='MM'?e.name_mm:''),
    company:x.company||coOf(e), leave_type:x.leave_type, dates:[x.leave_date], slots:slots, hours_each:Number(x.hours)||slots.length,
    filing_status:x.filing_status, approval:x.approval_status||'', filed_at:x.filed_at, skipped:0, state:'active' });
}
async function cancelMyLeave(leaveId){
  const x=(historyCache&&historyCache.history||[]).find(r=>r.leave_id===leaveId);
  const r0=await uiDialog({ titleHTML:TH('cancel_title'),
    bodyHTML:x?('<b>'+esc(fmtDateLong(x.leave_date))+'</b><br>'+esc(typeLabel(x.leave_type))+' · '+esc(slotsText(x.slots,x.hours))):'',
    input:true, required:true, requiredMsg:T('e_cancel_reason'), placeholder:T('cancel_reason_ph'),
    okHTML:TH('cancel_confirm'), cancelHTML:TH('cancel_back'), danger:true });
  if(!r0.ok) return;
  try{
    const r=await apiPost('cancelLeave',{emp_id:selectedEmp.emp_id, leave_id:leaveId, reason_cancel:r0.value});
    if(r.ok){toast(T('cancelled_ok'));loadHistory();}
    else toast(errText(r),true);
  }catch(err){toast(T('e_net'),true);}
}

/* ---------- ตรวจสลิป (เปิดจาก QR: ?slip=…&c=…) ---------- */
async function runVerify(id, code){
  showScreen('screenVerify');
  const box=$('verifyCard');
  box.innerHTML='<h2 class="vTitle">'+TB('v_title')+'</h2><div class="loading"><div class="spinner"></div>'+TB('v_checking')+'</div>';
  let res;
  try{ res=await apiGet('verify',{id:id,c:code}); }catch(e){ res={ok:false,code:'NET'}; }
  const head='<h2 class="vTitle">'+TB('v_title')+'</h2>';
  if(!res.ok && res.code!=='BAD_CODE' && res.code!=='MISSING'){     // ระบบขัดข้อง ≠ สลิปปลอม
    box.innerHTML=head+'<button class="empEmpty retry" id="verifyRetry" type="button">⚠️ '+TB('load_failed')+'</button>';
    $('verifyRetry').onclick=()=>runVerify(id,code);
    return;
  }
  if(!res.ok){
    box.innerHTML=head+'<div class="vResult bad"><div class="vIcon">✕</div><div>'+TB('v_fake')+'</div></div><div class="vNo">'+esc(id)+'</div>';
    return;
  }
  if(res.kind==='ot' && typeof otVerifyHTML==='function'){ box.innerHTML=otVerifyHTML(res,id); return; }
  if(!res.found){
    box.innerHTML=head+'<div class="vResult warn"><div class="vIcon">!</div><div>'+TB('v_gone')+'</div></div><div class="vNo">'+esc(id)+'</div>';
    return;
  }
  const items=(res.items||[]).map(it=>{
    const cancelled=(it.state==='cancelled');
    const ap=it.approval_status;
    return '<div class="vItem'+(cancelled?' cancelled':'')+'">'+
      '<div class="vDate">'+esc(fmtDateLong(it.leave_date,'th'))+'</div>'+
      '<div class="vMeta">'+esc(typeLabel(it.leave_type,'th'))+' <span class="l-my">'+esc(typeLabel(it.leave_type,'my'))+'</span> · '+esc(slotsText(it.slots,it.hours,'th'))+' · '+esc(String(it.hours))+' ชม.</div>'+
      '<div class="vState '+(cancelled?'c':(ap==='ไม่อนุมัติ'?'c':(ap==='รออนุมัติ'?'w':'a')))+'">'+
        (cancelled?TB('v_cancelled'):(ap?TB(APPROVAL_KEY[ap]):TB('v_active')))+'</div></div>';
  }).join('');
  box.innerHTML=head+
    '<div class="vResult good"><div class="vIcon">✓</div><div>'+TB('v_real')+'</div></div>'+
    '<div class="vWho"><b>'+esc(res.name)+'</b><span>'+esc(res.emp_id)+' · '+esc(res.company)+'</span></div>'+
    items+
    '<div class="vNo">'+esc(id)+' · '+esc(tr('v_checked_at',{t:res.checked_at},'th'))+'</div>';
}

/* ---------- สลับภาษา: วาดส่วนที่สร้างด้วยโค้ดใหม่ ---------- */
function onLangChange(){
  $('langTH').classList.toggle('active',LANG==='th');
  $('langMY').classList.toggle('active',LANG==='my');
  if(allEmployees.length){ setTabs(); renderEmployees(); }
  if(!$('screenForm').classList.contains('hidden')){
    renderTimeline(); updateHourTotal(); updateStatus();
    if(typeof otOnLang==='function') otOnLang();
    if(!$('paneHistory').classList.contains('hidden')) renderHistory();
  }
}

/* ---------- ผูกปุ่มฝั่งพนักงาน ---------- */
function pickNat(nat){ currentNat=nat; setLang(nat==='MM'?'my':'th'); setTabs(); renderEmployees(); }
$('tabTH').onclick=()=>pickNat('TH');
$('tabMM').onclick=()=>pickNat('MM');
$('coALT').onclick=()=>{currentCo='ALT';setTabs();renderEmployees();};
$('coMAIN').onclick=()=>{currentCo='MAIN';setTabs();renderEmployees();};
$('langTH').onclick=()=>setLang('th');
$('langMY').onclick=()=>setLang('my');
$('searchInput').oninput=debounce(renderEmployees,160);
$('refreshBtn').onclick=()=>{ loadEmployees(true); };
$('cmCancel').onclick=()=>$('confirmModal').classList.add('hidden');
$('cmConfirm').onclick=gotoForm;
$('confirmModal').onclick=e=>{if(e.target===$('confirmModal'))$('confirmModal').classList.add('hidden');};
$('changePersonBtn').onclick=()=>{ selectedEmp=null; showScreen('screenSelect'); };
$('stLeave').onclick=()=>switchSubTab('leave');
$('stHistory').onclick=()=>switchSubTab('history');
$('leaveDate').onchange=updateStatus;
$('dateFrom').onchange=()=>{ if($('dateTo').value<$('dateFrom').value) $('dateTo').value=$('dateFrom').value; updateStatus(); };
$('dateTo').onchange=updateStatus;
$('modeSingle').onclick=()=>setDayMode('single');
$('modeRange').onclick=()=>setDayMode('range');
$('leaveTypeChips').querySelectorAll('.chip').forEach(c=>c.onclick=()=>{
  if(c.classList.contains('disabled'))return;
  const wasVaca=(selectedType==='ลาพักร้อน');
  selectedType=c.dataset.type;
  if(wasVaca && selectedType!=='ลาพักร้อน'){ selectedSlots.clear(); markQuick(null); dayMode='single'; renderTimeline(); }
  renderTypeChips(); applyTypeMode(); updateHourTotal();
});
$('btnFull').onclick=()=>setQuick(TIME_SLOTS,'btnFull');
$('btnAM').onclick=()=>setQuick(AM_SLOTS,'btnAM');
$('btnPM').onclick=()=>setQuick(PM_SLOTS,'btnPM');
$('submitBtn').onclick=askConfirmSubmit;
$('csCancel').onclick=()=>{$('confirmSubmitModal').classList.add('hidden');pendingPayload=null;};
$('csConfirm').onclick=doSubmit;
$('confirmSubmitModal').onclick=e=>{if(e.target===$('confirmSubmitModal')){$('confirmSubmitModal').classList.add('hidden');pendingPayload=null;}};
$('againBtn').onclick=()=>{ const k=lastKind; gotoForm(); if(k==='ot') switchSubTab('ot'); };
$('doneBtn').onclick=()=>{selectedEmp=null;showScreen('screenSelect');loadEmployees();};
$('slipSaveBtn').onclick=saveSlip; $('slipShareBtn').onclick=shareSlip;
$('mSlipSave').onclick=saveSlip; $('mSlipShare').onclick=shareSlip;
$('mSlipClose').onclick=()=>$('slipModal').classList.add('hidden');
$('slipModal').onclick=e=>{if(e.target===$('slipModal'))$('slipModal').classList.add('hidden');};
$('verifyHome').onclick=()=>{ try{ history.replaceState(null,'',location.pathname); }catch(e){} showScreen('screenSelect'); };

/* ---------- เริ่ม ---------- */
STR.range_fmt     = {th:'{a} ถึง {b}', my:'{a} မှ {b} အထိ'};
STR.range_preview = {th:'ลาได้ {n} วันทำงาน (ข้ามวันหยุด {m} วัน)', my:'အလုပ်ရက် {n} ရက် ခွင့်ယူမည် (ပိတ်ရက် {m} ရက် ကျော်)'};
(function init(){
  try{ const last=readCache('lv_last_tab'); if(last){ if(last.co) currentCo=last.co; if(last.nat) currentNat=last.nat; } }catch(e){}
  window.addEventListener('pagehide',()=>writeCache('lv_last_tab',{co:currentCo,nat:currentNat}));
  applyI18n(); onLangChange();
  $('leaveDate').value=todayStr();
  const qs=new URLSearchParams(location.search), slipId=qs.get('slip');
  if(slipId) runVerify(slipId, qs.get('c')||'');
  loadEmployees();
})();
console.log('[leave] app.js v' + window.APP_VER.app);

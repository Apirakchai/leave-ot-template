/* ============================================================
 * ระบบแจ้งลา + โอที — admin.js — v1.0  (ส่วนโอทีของผู้ดูแลอยู่ใน ot.js)
 * ฝั่งผู้ดูแล: PIN / พนักงาน (เพิ่ม-แก้-ลบ-ซ่อน-นำเข้า-ย้ายบริษัท) / รูป /
 *             สรุปการลา / ใบลา / วันหยุด / รอบเงินเดือน / Telegram
 * ใช้ helper จาก app.js ($, esc, apiGet, apiPost, toast, uiDialog, avatarHTML ฯลฯ)
 * ============================================================ */
window.APP_VER = window.APP_VER || {};
window.APP_VER.admin = '1.0';

let adminPin='', adminEmployees=[], editingEmpId=null, editingPhotoId='';
let adminLeaves=[], leaveStatuses=['ลาล่วงหน้า','ลากระทันหัน (วันเดียวกัน)','ลาย้อนหลัง'];
let empCoFilter='ALL', dashCoFilter='ALL', leavesCoFilter='ALL';
let selectMode=false, selectedIds=new Set(), importRows=[];
let leavesMix=false, leavesRange=null;

function coPill(co){ return co==='MAIN'?'<span class="pill main">MAIN</span>':'<span class="pill">ALT</span>'; }
function exPill(x){ return x?'<span class="pill ex">อดีต ALT</span>':''; }
function natFlag(n){ return n==='MM'?'🇲🇲':'🇹🇭'; }
function hasMix(list){ return false; }      // MAIN: ระบบบริษัทเดียว ไม่มีป้าย/ตัวกรองบริษัท
function thisMonth(){ const d=new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
/* รอบเงินเดือนปัจจุบันจบเดือนไหน (ตัดรอบวันที่ 25: วันที่ 26 ขึ้นไป = รอบเดือนถัดไป) */
function currentCycleMonth(cut){ const d=new Date(); if(d.getDate()>(cut||25)) d.setMonth(d.getMonth()+1,1); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
let adminCutoff=25;

/* ปุ่มกรองบริษัท: ทั้งหมด / ALT / MAIN / อดีต ALT (โชว์เมื่อมีมากกว่า 1 กลุ่ม) */
function renderCoChips(boxId, current, items, onPick, force){
  const box=$(boxId);
  if(!force && !hasMix(items)){ box.innerHTML=''; return; }
  const cnt={ALL:items.length, ALT:items.filter(x=>x.company!=='MAIN').length, MAIN:items.filter(x=>x.company==='MAIN').length, EX:items.filter(x=>x.ex_alt).length};
  const defs=[['ALL','ทั้งหมด'],['ALT','ALT'],['MAIN','MAIN'],['EX','อดีต ALT']];
  box.innerHTML=defs.map(d=>'<button type="button" class="fChip'+(current===d[0]?' active':'')+'" data-f="'+d[0]+'">'+d[1]+'<i>'+cnt[d[0]]+'</i></button>').join('');
  box.querySelectorAll('.fChip').forEach(b=>{ b.onclick=()=>onPick(b.dataset.f); });
}
function coMatch(f, x){ return f==='ALL' || (f==='EX'?!!x.ex_alt:(f==='MAIN'?x.company==='MAIN':x.company!=='MAIN')); }

/* ---------- เข้าสู่ระบบด้วย PIN ---------- */
function openPinModal(){$('pinInput').value='';$('pinModal').classList.remove('hidden');setTimeout(()=>$('pinInput').focus(),100);}
async function submitPin(){
  const pin=$('pinInput').value.trim(); if(!pin)return;
  const btn=$('pinSubmit'); btn.disabled=true; btn.textContent='ตรวจสอบ...';
  try{
    const res=await apiGet('checkpin',{pin:pin});   // ตรวจ PIN แบบเบา → เปิดหน้าแอดมินทันที แล้วค่อยโหลดรายชื่อ
    if(res.ok){adminPin=pin;$('pinModal').classList.add('hidden');openAdmin();loadAdminEmployees();}
    else toast('PIN ไม่ถูกต้อง',true);
  }catch(err){toast('เชื่อมต่อไม่ได้ ลองใหม่อีกครั้ง',true);}
  btn.disabled=false; btn.textContent='เข้าสู่ระบบ';
}
function openAdmin(){
  $('adminScreen').classList.remove('hidden');
  const v=window.APP_VER||{};
  $('verFoot').textContent='เวอร์ชันหน้าเว็บ '+(v.app||'?')+' (config '+(v.config||'?')+')';
  $('adminEmpList').innerHTML='<div class="loading"><div class="spinner"></div>กำลังโหลดรายชื่อ...</div>';
  setSelectMode(false);
  switchAdminTab('emp');
}
function closeAdmin(){ $('adminScreen').classList.add('hidden'); loadEmployees(); }
async function loadAdminEmployees(){
  try{
    const res=await apiGet('employees',{admin:'1',pin:adminPin});
    if(res.ok){adminEmployees=res.employees||[];renderAdminEmp(); if(typeof aotRenderList==='function' && !$('adminOTPane').classList.contains('hidden')) aotRenderList();}
    else{adminEmployees=[];$('adminEmpList').innerHTML='<div class="empEmpty">⚠️ '+esc(res.error||'โหลดรายชื่อไม่สำเร็จ')+'</div>';}
  }catch(err){$('adminEmpList').innerHTML='<div class="empEmpty">⚠️ โหลดรายชื่อไม่สำเร็จ</div>';}
}
function switchAdminTab(t){
  $('atEmp').classList.toggle('active',t==='emp');
  $('atDash').classList.toggle('active',t==='dash');
  $('atLeaves').classList.toggle('active',t==='leaves');
  $('atOT').classList.toggle('active',t==='ot');
  $('atSet').classList.toggle('active',t==='set');
  $('adminEmpPane').classList.toggle('hidden',t!=='emp');
  $('adminDashPane').classList.toggle('hidden',t!=='dash');
  $('adminLeavesPane').classList.toggle('hidden',t!=='leaves');
  $('adminOTPane').classList.toggle('hidden',t!=='ot');
  $('adminSetPane').classList.toggle('hidden',t!=='set');
  if(t==='ot' && typeof aotOpen==='function') aotOpen();
  if(t==='dash'){if(!$('dashMonth').value)$('dashMonth').value=currentCycleMonth(adminCutoff);loadDashboard();}
  if(t==='leaves'){if(!$('leavesMonth').value)$('leavesMonth').value=currentCycleMonth(adminCutoff);loadLeaves();}
  if(t==='set'){loadCalendarAdmin();loadAdminSettings();}
}

/* ---------- รายชื่อพนักงาน (แอดมิน) ---------- */
function empSortKey(e){ return (e.company==='MAIN'?'1':'0')+(e.nationality==='MM'?'1':'0')+String(e.emp_id).replace(/\d+/g,n=>n.padStart(6,'0')); }
function visibleAdminEmp(){
  const q=$('adminSearch').value.trim().toLowerCase();
  let list=adminEmployees.filter(e=>coMatch(empCoFilter,e));
  if(q) list=list.filter(e=>(e.name_th||'').toLowerCase().includes(q)||(e.name_mm||'').toLowerCase().includes(q)||String(e.emp_id||'').toLowerCase().includes(q)||(e.note||'').toLowerCase().includes(q));
  return list.sort((a,b)=>empSortKey(a)<empSortKey(b)?-1:1);
}
function renderAdminEmp(){
  if(empCoFilter!=='ALL' && !hasMix(adminEmployees)) empCoFilter='ALL';      // ไม่มีหลายกลุ่มแล้ว → เลิกกรอง (กันรายชื่อว่างเปล่า)
  renderCoChips('empCoFilter', empCoFilter, adminEmployees, f=>{empCoFilter=f;renderAdminEmp();});
  const list=visibleAdminEmp();
  if(list.length===0){$('adminEmpList').innerHTML='<div class="empEmpty">ไม่พบพนักงาน</div>';updateMoveBar();return;}
  const mix=hasMix(adminEmployees);
  $('adminEmpList').innerHTML=list.map(e=>{
    const nm=e.name_th||e.name_mm, id=esc(e.emp_id), sel=selectedIds.has(e.emp_id);
    const tags=(mix?coPill(e.company):'')+exPill(e.ex_alt)+(e.ex_alt&&e.moved_at?'<span>ย้าย '+esc(fmtDMY(e.moved_at))+'</span>':'')+(e.move_pending?'<span class="pill main">จะย้ายไป MAIN '+esc(fmtDMY(e.moved_at))+'</span>':'')+(e.visible?'':'<span class="pill off">ซ่อน</span>');
    return '<div class="adminRow'+(sel?' sel':'')+'" data-row="'+id+'">'+
      (selectMode?'<input type="checkbox" data-pick="'+id+'"'+(sel?' checked':'')+' aria-label="เลือก '+esc(nm)+'">':'')+
      avatarHTML('ra',nm,e.photo_url,'',120)+
      '<div class="rn"><b>'+esc(nm)+'</b><small><span>'+natFlag(e.nationality)+' '+id+(e.note?' · '+esc(e.note):'')+'</span>'+tags+'</small></div>'+
      (selectMode?'':'<div class="rowBtns">'+
        '<button class="miniBtn" type="button" data-act="edit" data-id="'+id+'">แก้</button>'+
        '<button class="miniBtn hide" type="button" data-act="hide" data-id="'+id+'">'+(e.visible?'ซ่อน':'แสดง')+'</button>'+
        '<button class="miniBtn del" type="button" data-act="del" data-id="'+id+'">ลบ</button></div>')+
    '</div>';
  }).join('');
  const box=$('adminEmpList');
  box.querySelectorAll('[data-act]').forEach(b=>{b.onclick=()=>adminAction(b.dataset.act,b.dataset.id);});
  if(selectMode) box.querySelectorAll('.adminRow').forEach(r=>{ r.onclick=ev=>{
    const id=r.dataset.row; if(ev.target.tagName!=='INPUT'){ if(selectedIds.has(id))selectedIds.delete(id); else selectedIds.add(id); }
    else { if(ev.target.checked)selectedIds.add(id); else selectedIds.delete(id); }
    renderAdminEmp(); }; });
  updateMoveBar();
}
async function adminAction(act,id){
  const e=adminEmployees.find(x=>x.emp_id===id); if(!e) return;
  if(act==='edit'){openEmpModal(e);return;}
  if(act==='hide'){
    try{
      const r=await apiPost('toggleVisible',{pin:adminPin,emp_id:id});
      if(r.ok){e.visible=r.visible;renderAdminEmp();toast(r.visible?'แสดงชื่อแล้ว':'ซ่อนชื่อแล้ว');}else toast(r.error||'ผิดพลาด',true);
    }catch(err){toast('เชื่อมต่อไม่ได้',true);}
    return;
  }
  if(act==='del'){
    const c=await uiDialog({title:'ลบ "'+(e.name_th||e.name_mm)+'" ออกจากระบบ?', body:'ประวัติการลายังเก็บไว้ แต่จะไม่แสดงในรายชื่อและแจ้งลาไม่ได้อีก', okText:'ลบ', cancelText:'ไม่ลบ', danger:true});
    if(!c.ok) return;
    try{
      const r=await apiPost('deleteEmployee',{pin:adminPin,emp_id:id});
      if(r.ok){adminEmployees=adminEmployees.filter(x=>x.emp_id!==id);renderAdminEmp();toast('ลบแล้ว');}else toast(r.error||'ผิดพลาด',true);
    }catch(err){toast('เชื่อมต่อไม่ได้',true);}
  }
}

/* ---------- ย้ายพนักงานระหว่างบริษัท ---------- */
function setSelectMode(on){
  selectMode=on; selectedIds.clear();
  $('moveBar').classList.toggle('hidden',!on);
  $('selectModeBtn').classList.toggle('btn-primary',on); $('selectModeBtn').classList.toggle('btn-ghost',!on);
  if(adminEmployees.length) renderAdminEmp();
}
function updateMoveBar(){
  $('moveCount').textContent='เลือก '+selectedIds.size+' คน';
  const picked=adminEmployees.filter(e=>selectedIds.has(e.emp_id));
  $('moveToMainBtn').disabled=!picked.some(e=>e.company!=='MAIN' && !e.move_pending);
  $('moveToAltBtn').disabled=!picked.some(e=>e.ex_alt || e.move_pending);   // ยกเลิกการย้ายได้เฉพาะคนที่เคยย้ายมาจาก ALT
}
async function doTransfer(to){
  const picked=adminEmployees.filter(e=>selectedIds.has(e.emp_id) && (to==='MAIN'?(e.company!=='MAIN' && !e.move_pending):(e.ex_alt||e.move_pending)));
  if(!picked.length){toast('ยังไม่ได้เลือกพนักงาน',true);return;}
  const names=picked.slice(0,8).map(e=>esc(e.name_th||e.name_mm)).join(', ')+(picked.length>8?' …':'');
  let c;
  if(to==='MAIN'){
    c=await uiDialog({title:'ย้าย '+picked.length+' คน ไปเป็นพนักงาน MAIN', bodyHTML:names+'<br><br><b>เริ่มเป็นพนักงาน MAIN ตั้งแต่วันที่</b><br><span style="font-size:12.5px">ใบลาตั้งแต่วันนี้นับเป็นของ MAIN ก่อนหน้านั้นยังเป็นของ ALT รหัสพนักงานและประวัติลาเดิมอยู่ครบ</span>',
      date:true, dateValue:todayStr(), okText:'ย้ายไป MAIN', cancelText:'ยกเลิก'});
    if(!c.ok) return;
    if(!c.date){toast('เลือกวันที่มีผลก่อน',true);return;}
  } else {
    c=await uiDialog({title:'ยกเลิกการย้าย '+picked.length+' คน (กลับเป็นพนักงาน ALT)?', bodyHTML:names+'<br><br>ใช้เมื่อย้ายผิดคน/ผิดวัน — ใบลาทั้งหมดของคนเหล่านี้จะกลับไปนับเป็นของ ALT', okText:'ยกเลิกการย้าย', cancelText:'ไม่ทำ'});
    if(!c.ok) return;
  }
  toast('กำลังย้าย...');
  try{
    const r=await apiPost('transferEmployees',{pin:adminPin,emp_ids:picked.map(e=>e.emp_id),to:to,effective:c.date||''});
    if(!r.ok) throw new Error(r.error||'ย้ายไม่สำเร็จ');
    toast('ย้ายแล้ว '+r.moved+' คน ✓');
    setSelectMode(false); empCoFilter='ALL';
    await loadAdminEmployees();
  }catch(err){toast(err.message||'เชื่อมต่อไม่ได้',true);}
}

/* ---------- เพิ่ม/แก้ พนักงาน ---------- */
function openEmpModal(e){
  editingEmpId=e?e.emp_id:null; editingPhotoId=e?(e.photo_id||''):'';
  $('empModalTitle').textContent=e?('แก้ไขพนักงาน · '+e.emp_id):'เพิ่มพนักงานใหม่';
  $('empNameTh').value=e?(e.name_th||''):'';
  $('empNameMm').value=e?(e.name_mm||''):'';
  $('empNat').value=e?e.nationality:'TH';
  $('empCo').value='MAIN';
  $('empNote').value=e?(e.note||''):'';
  $('empIdInput').value=e?e.emp_id:'';                  // MAIN: แก้รหัสได้ (ระบบเปลี่ยนในใบลา/โอทีเก่าให้ครบ)
  $('empIdInput').placeholder=e?'':'เว้นว่าง = ระบบตั้งให้';
  $('empIdHint').classList.toggle('hidden',!e);
  const nm=e?(e.name_th||e.name_mm):'?', prev=$('empPrev');
  if(e&&e.photo_url){prev.textContent='';prev.style.background="url('"+e.photo_url+"') center/cover";}
  else{prev.style.background='var(--bg2)';prev.textContent=e?initials(nm):'?';}
  $('empPhoto').value='';
  $('empModal').classList.remove('hidden');
}
async function saveEmp(){
  const name_th=$('empNameTh').value.trim(), name_mm=$('empNameMm').value.trim();
  if(!name_th && !name_mm){toast('กรุณากรอกชื่อ',true);return;}
  const btn=$('empSave'); btn.disabled=true; btn.textContent='บันทึก...';
  try{
    const data={pin:adminPin,name_th:name_th,name_mm:name_mm,nationality:$('empNat').value,company:$('empCo').value,note:$('empNote').value.trim(),photo_id:editingPhotoId};
    let r;
    if(editingEmpId){
      const newId=$('empIdInput').value.trim();
      if(!newId) throw new Error('กรุณากรอกรหัสพนักงาน');
      if(newId!==editingEmpId){
        const c=await uiDialog({title:'เปลี่ยนรหัส '+editingEmpId+' → '+newId+' ?', body:'ใบลาและโอทีเก่าของคนนี้จะเปลี่ยนเป็นรหัสใหม่ทั้งหมด', okText:'เปลี่ยนรหัส', cancelText:'ยกเลิก'});
        if(!c.ok){ btn.disabled=false; btn.textContent='บันทึก'; return; }
        const cr=await apiPost('changeEmpId',{pin:adminPin,emp_id:editingEmpId,new_id:newId});
        if(!cr.ok) throw new Error(cr.error||'เปลี่ยนรหัสไม่สำเร็จ');
        editingEmpId=newId;
      }
      data.emp_id=editingEmpId;r=await apiPost('updateEmployee',data);
    }
    else { const id=$('empIdInput').value.trim(); if(id) data.emp_id=id; r=await apiPost('addEmployee',data); }
    if(!r.ok) throw new Error(r.error||'บันทึกไม่สำเร็จ');
    $('empModal').classList.add('hidden'); toast(editingEmpId?'บันทึกแล้ว ✓':('เพิ่มแล้ว · รหัส '+r.emp_id));
    await loadAdminEmployees();
  }catch(err){toast(err.message||'เชื่อมต่อไม่ได้',true);}
  btn.disabled=false; btn.textContent='บันทึก';
}

/* ---------- นำเข้ารายชื่อ (วางจาก Excel) ---------- */
function parseNat(s, def){ s=String(s||'').trim().toLowerCase(); if(/^(mm|my|พม่า|เมียนมา|myan|burm)/.test(s)) return 'MM'; if(/^(th|ไทย|thai)/.test(s)) return 'TH'; return def; }
function looksLikeId(s){ return /^[A-Za-z0-9][A-Za-z0-9\-_.]{0,19}$/.test(s) && /\d/.test(s); }
function parseImport(){
  const def=$('impNat').value, have=new Set(adminEmployees.map(e=>String(e.emp_id).toLowerCase())), seen=new Set();
  importRows=$('impText').value.split(/\r?\n/).map(l=>l.replace(/\s+$/,'')).filter(l=>l.trim()).map(line=>{
    let c=line.split('\t').map(x=>x.trim());
    if(c.length===1) c=line.split(/\s{2,}|,/).map(x=>x.trim());
    let row;
    if(c.length===1) row={emp_id:'',name_th:c[0],name_mm:'',nationality:def,note:''};
    else if(c.length===2 && looksLikeId(c[0])) row={emp_id:c[0],name_th:c[1],name_mm:'',nationality:def,note:''};
    else if(c.length===2) row={emp_id:'',name_th:c[0],name_mm:c[1],nationality:def,note:''};
    else row={emp_id:c[0]||'',name_th:c[1]||'',name_mm:c[2]||'',nationality:parseNat(c[3],def),note:c[4]||''};
    if(/^(รหัส|รหัสพนักงาน|id|emp[_ ]?id|code|no\.?|ลำดับ)$/i.test(c[0]||'') && /ชื่อ|name/i.test(c.slice(1).join(' '))) return null;   // บรรทัดหัวตาราง
    row.err='';
    if(!row.name_th && !row.name_mm) row.err='ไม่มีชื่อ';
    else if(row.emp_id && !/^[A-Za-z0-9][A-Za-z0-9\-_.]{0,19}$/.test(row.emp_id)) row.err='รหัสไม่ถูกรูปแบบ';
    else if(row.emp_id && (have.has(row.emp_id.toLowerCase())||seen.has(row.emp_id.toLowerCase()))) row.err='รหัสซ้ำ';
    if(row.emp_id) seen.add(row.emp_id.toLowerCase());
    return row;
  }).filter(Boolean);
  const good=importRows.filter(r=>!r.err).length;
  $('impPreview').innerHTML=importRows.length?('<table><tr><th>รหัส</th><th>ชื่อ</th><th>ชื่อพม่า/อังกฤษ</th><th>สัญชาติ</th><th>ตำแหน่ง</th></tr>'+
    importRows.map(r=>'<tr'+(r.err?' class="bad"':'')+'><td>'+esc(r.emp_id||'(อัตโนมัติ)')+'</td><td>'+esc(r.name_th)+(r.err?' — '+r.err:'')+'</td><td>'+esc(r.name_mm)+'</td><td>'+natFlag(r.nationality)+'</td><td>'+esc(r.note)+'</td></tr>').join('')+'</table>'):'';
  $('impSave').disabled=!good; $('impSave').textContent=good?('นำเข้า '+good+' คน'):'นำเข้า';
}
async function saveImport(){
  const rows=importRows.filter(r=>!r.err).map(r=>({emp_id:r.emp_id,name_th:r.name_th,name_mm:r.name_mm,nationality:r.nationality,note:r.note}));
  if(!rows.length) return;
  const btn=$('impSave'); btn.disabled=true; btn.textContent='กำลังนำเข้า...';
  try{
    const r=await apiPost('importEmployees',{pin:adminPin,company:$('impCo').value,rows:rows});
    if(!r.ok) throw new Error(r.error||'นำเข้าไม่สำเร็จ');
    $('importModal').classList.add('hidden');
    toast('นำเข้าแล้ว '+r.added+' คน'+(r.skipped&&r.skipped.length?(' (ข้าม '+r.skipped.length+')'):'')+' ✓');
    await loadAdminEmployees();
  }catch(err){toast(err.message||'เชื่อมต่อไม่ได้',true); btn.disabled=false; btn.textContent='นำเข้า';}
}

/* ---------- รูปพนักงาน: ย่อขนาดในเครื่องก่อนส่ง ---------- */
function resizePhoto(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{const img=new Image();img.onload=()=>{
      const max=600;let w=img.width,h=img.height;
      if(w>h){if(w>max){h=h*max/w;w=max;}}else{if(h>max){w=w*max/h;h=max;}}
      const cv=document.createElement('canvas');cv.width=w;cv.height=h;
      cv.getContext('2d').drawImage(img,0,0,w,h);
      resolve(cv.toDataURL('image/jpeg',0.85));
    };img.onerror=reject;img.src=e.target.result;};
    reader.onerror=reject; reader.readAsDataURL(file);
  });
}
async function handlePhotoUpload(){
  const f=$('empPhoto').files[0]; if(!f)return;
  try{
    const b64=await resizePhoto(f);
    $('empPrev').textContent=''; $('empPrev').style.background="url('"+b64+"') center/cover";
    toast('กำลังอัปโหลดรูป...');
    const r=await apiPost('uploadPhoto',{pin:adminPin,base64:b64,mime:'image/jpeg',filename:(editingEmpId||'emp_'+Date.now())+'.jpg'});
    if(r.ok){editingPhotoId=r.photo_id;toast('อัปโหลดรูปแล้ว ✓ กดบันทึกเพื่อใช้รูปนี้');}
    else toast(r.error||'อัปโหลดไม่สำเร็จ',true);
  }catch(err){toast('อัปโหลดไม่สำเร็จ',true);}
}
async function handleBatchPhotos(){
  const files=Array.from($('batchPhotos').files||[]);
  if(files.length===0)return;
  const log=$('batchLog'); log.innerHTML='';
  const byId={}; adminEmployees.forEach(e=>{byId[String(e.emp_id).trim().toLowerCase()]=e;});
  let ok=0;
  function line(txt,color){const d=document.createElement('div');d.style.cssText='font-size:13px;padding:4px 0;color:'+(color||'var(--text)');d.innerHTML=txt;log.appendChild(d);log.scrollTop=log.scrollHeight;}
  for(const f of files){
    const base=f.name.replace(/\.[^.]+$/,'').trim();
    const emp=byId[base.toLowerCase()];
    if(!emp){ line('⚠️ '+esc(f.name)+' — ไม่พบรหัส "'+esc(base)+'"','var(--amber)'); continue; }
    try{
      const b64=await resizePhoto(f);
      const up=await apiPost('setEmpPhoto',{pin:adminPin,emp_id:emp.emp_id,base64:b64,mime:'image/jpeg',filename:base+'.jpg'});
      if(!up.ok) throw new Error(up.error||'อัปโหลดไม่สำเร็จ');
      emp.photo_id=up.photo_id; if(up.photo_url)emp.photo_url=up.photo_url;
      ok++; line('✅ '+esc(emp.name_th||emp.name_mm)+' ('+esc(emp.emp_id)+')','var(--green)');
    }catch(err){ line('❌ '+esc(f.name)+' — '+esc(err.message),'var(--red)'); }
  }
  line('<b>เสร็จ — สำเร็จ '+ok+'/'+files.length+' รูป</b>');
  toast('อัปโหลดเสร็จ '+ok+'/'+files.length);
  renderAdminEmp();
}
async function syncPhotosFromDrive(){
  const c=await uiDialog({title:'ซิงค์รูปจากโฟลเดอร์ Drive?', body:'ต้องวางรูปในโฟลเดอร์ Leave_Photos และตั้งชื่อไฟล์เป็นรหัสพนักงาน (เช่น K001.jpg) ก่อน', okText:'ซิงค์', cancelText:'ยกเลิก'});
  if(!c.ok) return;
  const btn=$('syncDriveBtn'); btn.disabled=true; const old=btn.textContent; btn.textContent='⏳ กำลังซิงค์...';
  try{
    const r=await apiPost('syncPhotos',{pin:adminPin});
    if(!r.ok) throw new Error(r.error||'ซิงค์ไม่สำเร็จ');
    await loadAdminEmployees();
    let msg='อัปเดตรูป '+r.matched+' คน';
    if(r.unmatched_files && r.unmatched_files.length) msg+='\n\nไฟล์ที่ชื่อไม่ตรงรหัสพนักงาน ('+r.unmatched_files.length+'):\n'+r.unmatched_files.slice(0,15).join(', ')+(r.unmatched_files.length>15?' ...':'');
    uiDialog({title:'✅ ซิงค์รูปเสร็จ', body:msg, alert:true});
  }catch(err){toast(err.message||'เชื่อมต่อไม่ได้',true);}
  btn.disabled=false; btn.textContent=old;
}

/* ---------- สรุปการลา ---------- */
function hd(h){ return Math.round(h/8*100)/100; }
async function loadDashboard(){
  $('dashContent').innerHTML='<div class="loading"><div class="spinner"></div>กำลังโหลด...</div>';
  try{
    const res=await apiGet('dashboard',{pin:adminPin,month:$('dashMonth').value,cycle:$('dashCycle').checked?'1':'',company:dashCoFilter==='ALL'?'':dashCoFilter});
    if(!res.ok) throw new Error(res.error||'load failed');
    const s=res.summary, mix=res.mix&&res.mix.show;
    const box=$('dashCoFilter');
    if(mix){
      const defs=[['ALL','ทั้งหมด'],['ALT','ALT'],['MAIN','MAIN'],['EX','อดีต ALT']];
      box.innerHTML=defs.map(d=>'<button type="button" class="fChip'+(dashCoFilter===d[0]?' active':'')+'" data-f="'+d[0]+'">'+d[1]+'</button>').join('');
      box.querySelectorAll('.fChip').forEach(b=>{ b.onclick=()=>{dashCoFilter=b.dataset.f;loadDashboard();}; });
    } else box.innerHTML='';
    let html=(res.range?'<div class="rangeNote">📆 '+esc(res.range.label)+' · '+esc(fmtDMY(res.range.from))+' – '+esc(fmtDMY(res.range.to))+'</div>':'')+
      '<div class="statGrid">'+
      '<div class="statCard"><div class="sv">'+s.people_count+'</div><div class="sl">คนที่ลา</div></div>'+
      '<div class="statCard"><div class="sv">'+s.total_days+'</div><div class="sl">รวม (วัน)</div></div>'+
      '<div class="statCard"><div class="sv">'+s.total_hours+'</div><div class="sl">รวม (ชั่วโมง)</div></div>'+
      '<div class="statCard"><div class="sv">'+s.total_requests+'</div><div class="sl">ใบลา</div></div>'+
      (s.pending_requests?'<div class="statCard"><div class="sv" style="color:var(--amber)">'+s.pending_requests+'</div><div class="sl">พักร้อนรออนุมัติ (ยังไม่นับ)</div></div>':'')+
    '</div>';
    const types=Object.entries(res.by_type||{});
    if(types.length){const mx=Math.max.apply(null,types.map(t=>t[1]))||1;
      html+='<div class="fblock"><div class="flabel">แยกตามประเภทการลา</div>'+types.map(t=>'<div class="bar"><div class="bl">'+esc(t[0])+'</div><div class="bt"><div class="bf" style="width:'+(t[1]/mx*100)+'%"></div></div><div class="bv">'+hd(t[1])+' วัน</div></div>').join('')+'</div>';}
    const st=res.by_status||{}, sv=Object.values(st), smx=(sv.length?Math.max.apply(null,sv):0)||1;
    const sb=(lab,key,col)=>'<div class="bar"><div class="bl">'+lab+'</div><div class="bt"><div class="bf" style="width:'+((st[key]||0)/smx*100)+'%;background:var(--'+col+')"></div></div><div class="bv">'+(st[key]||0)+' ใบ</div></div>';
    html+='<div class="fblock"><div class="flabel">แยกตามการแจ้ง</div>'+sb('🟢 ล่วงหน้า','ลาล่วงหน้า','green')+sb('🟡 กระทันหัน','ลากระทันหัน (วันเดียวกัน)','amber')+sb('🔴 ย้อนหลัง','ลาย้อนหลัง','red')+'</div>';
    if(mix && dashCoFilter==='ALL'){ const bc=res.by_company||{};
      html+='<div class="fblock"><div class="flabel">แยกตามบริษัท</div><div class="statGrid" style="margin:0"><div class="statCard"><div class="sv">'+hd(bc.ALT||0)+'</div><div class="sl">ALT (วัน)</div></div><div class="statCard"><div class="sv" style="color:var(--main)">'+hd(bc.MAIN||0)+'</div><div class="sl">MAIN (วัน)</div></div></div></div>'; }
    if(res.people.length){
      html+='<div class="fblock"><div class="flabel">รายคน (เรียงจากลามากสุด)</div>'+res.people.map(p=>{
        const parts=[]; if(p.sick)parts.push('ป่วย '+hd(p.sick)); if(p.personal)parts.push('กิจ '+hd(p.personal)); if(p.vacation)parts.push('พักร้อน '+hd(p.vacation));
        return '<div class="adminRow"><div class="rn"><b>'+natFlag(p.nationality)+' '+esc(p.name)+'</b><small><span>'+esc(p.emp_id)+' · '+parts.join(' · ')+'</span>'+(mix?coPill(p.company):'')+exPill(p.ex_alt)+'</small></div><div style="text-align:right;flex-shrink:0"><div style="font-weight:800;color:var(--brand)">'+p.days+' วัน</div><small style="color:var(--muted)">'+p.hours+' ชม. · '+p.count+' ใบ</small></div></div>';
      }).join('')+'</div>';
    }else{html+='<div class="empEmpty">ไม่มีการลาในช่วงนี้</div>';}
    $('dashContent').innerHTML=html;
  }catch(err){$('dashContent').innerHTML='<div class="empEmpty">⚠️ '+esc(err.message||'โหลดไม่สำเร็จ')+'</div>';}
}

/* ---------- จัดการใบลา (แก้สถานะ / ลบ / ส่งออก) ---------- */
async function loadLeaves(){
  $('leavesContent').innerHTML='<div class="loading"><div class="spinner"></div>กำลังโหลด...</div>';
  try{
    const res=await apiGet('leaves',{admin:'1',pin:adminPin,month:$('leavesMonth').value,cycle:$('leavesCycle').checked?'1':'',company:leavesCoFilter==='MAIN'?'MAIN':''});   // รอบของ MAIN อาจตัดคนละวันกับ ALT
    if(!res.ok) throw new Error(res.error||'load failed');
    adminLeaves=res.leaves||[]; leavesMix=!!(res.mix&&res.mix.show); leavesRange=res.range||null;
    if(!leavesMix) leavesCoFilter='ALL';
    if(res.statuses) leaveStatuses=res.statuses;
    renderLeaves();
  }catch(err){$('leavesContent').innerHTML='<div class="empEmpty">⚠️ '+esc(err.message||'โหลดไม่สำเร็จ')+'</div>';}
}
function statusClass(s){s=String(s||'');return s.indexOf('ย้อนหลัง')>=0?'back':(s.indexOf('กระทันหัน')>=0?'same':'adv');}
function visibleLeaves(){
  const q=$('leavesSearch').value.trim().toLowerCase();
  let list=adminLeaves.filter(x=>coMatch(leavesCoFilter,x));
  if(q) list=list.filter(x=>(x.name||'').toLowerCase().includes(q)||String(x.emp_id||'').toLowerCase().includes(q));
  return list;
}
function renderLeaves(){
  renderCoChips('leavesCoFilter', leavesCoFilter, adminLeaves, f=>{ const reload=(f==='MAIN')!==(leavesCoFilter==='MAIN'); leavesCoFilter=f; if(reload) loadLeaves(); else renderLeaves(); }, leavesMix);
  const list=visibleLeaves(), mix=leavesMix||hasMix(adminLeaves);
  const rangeNote=leavesRange?'<div class="rangeNote">📆 '+esc(leavesRange.label)+' · '+esc(fmtDMY(leavesRange.from))+' – '+esc(fmtDMY(leavesRange.to))+'</div>':'';
  if(list.length===0){$('leavesContent').innerHTML=rangeNote+'<div class="empEmpty">ไม่มีใบลาในช่วงนี้</div>';return;}
  $('leavesContent').innerHTML=rangeNote+'<div class="desc" style="margin-bottom:10px">ทั้งหมด '+list.length+' ใบ</div>'+
  list.map(x=>{
    const c=statusClass(x.filing_status), id=esc(x.leave_id);
    const opts=leaveStatuses.map(s=>'<option value="'+esc(s)+'"'+(s===x.filing_status?' selected':'')+'>'+esc(s)+'</option>').join('');
    return '<div class="histItem '+c+(x.approval_status==='ไม่อนุมัติ'?' rejected':'')+'">'+
      '<div class="histTop"><span class="histDate">'+esc(fmtDateLong(x.leave_date,'th'))+'</span><span class="tag '+c+'">'+esc(x.filing_status)+'</span></div>'+
      '<div class="histMeta"><span><b>'+natFlag(x.nationality)+' '+esc(x.name)+'</b> ('+esc(x.emp_id)+')</span>'+(mix?coPill(x.company):'')+exPill(x.ex_alt)+'</div>'+
      '<div class="histMeta" style="margin-top:3px"><b>'+esc(x.leave_type)+'</b><span>⏰ '+esc(slotsText(x.slots,x.hours,'th'))+'</span><span>'+x.hours+' ชม.</span></div>'+
      (x.reason?'<div class="histReason">"'+esc(x.reason)+'"</div>':'')+
      approvalBadge(x.approval_status)+(x.approver?'<span class="histFiled"> โดย '+esc(x.approver)+'</span>':'')+
      '<div class="histFiled">แจ้งเมื่อ '+esc(x.filed_at)+'</div>'+
      '<div style="display:flex;gap:8px;align-items:center;margin-top:10px;flex-wrap:wrap">'+
        '<select class="leaveStatusSel" data-id="'+id+'" aria-label="เปลี่ยนสถานะการแจ้ง" style="flex:1;min-width:160px;padding:9px;border:2px solid var(--line);border-radius:9px;background:#fff;min-height:42px">'+opts+'</select>'+
        '<button class="miniBtn del" type="button" data-del="'+id+'">ลบใบนี้</button>'+
      '</div>'+
    '</div>';
  }).join('');
  $('leavesContent').querySelectorAll('.leaveStatusSel').forEach(sel=>{sel.onchange=()=>changeLeaveStatus(sel.dataset.id,sel.value);});
  $('leavesContent').querySelectorAll('[data-del]').forEach(b=>{b.onclick=()=>deleteLeaveAdmin(b.dataset.del);});
}
async function changeLeaveStatus(leaveId,status){
  try{
    const r=await apiPost('updateLeaveStatus',{pin:adminPin,leave_id:leaveId,status:status});
    if(r.ok){ const x=adminLeaves.find(l=>l.leave_id===leaveId); if(x)x.filing_status=status; renderLeaves(); toast('เปลี่ยนสถานะแล้ว ✓'); }
    else { toast(r.error||'ผิดพลาด',true); renderLeaves(); }
  }catch(err){toast('เชื่อมต่อไม่ได้',true); renderLeaves();}
}
async function deleteLeaveAdmin(leaveId){
  const x=adminLeaves.find(l=>l.leave_id===leaveId);
  const c=await uiDialog({title:'ลบใบลานี้ออกจากระบบ?', body:(x?(x.name+' · '+fmtDMY(x.leave_date)+' · '+x.leave_type+'\n\n'):'')+'ลบถาวร ใช้เมื่อกรอกผิดเท่านั้น', okText:'ลบถาวร', cancelText:'ไม่ลบ', danger:true});
  if(!c.ok) return;
  try{
    const r=await apiPost('deleteLeave',{pin:adminPin,leave_id:leaveId});
    if(r.ok){adminLeaves=adminLeaves.filter(l=>l.leave_id!==leaveId);renderLeaves();toast('ลบใบลาแล้ว');}
    else toast(r.error||'ผิดพลาด',true);
  }catch(err){toast('เชื่อมต่อไม่ได้',true);}
}
function exportLeavesCSV(){
  const list=visibleLeaves();
  if(!list.length){toast('ไม่มีข้อมูลให้ส่งออก',true);return;}
  const head=['วันที่ลา','รหัสพนักงาน','ชื่อ','สัญชาติ','ประเภทการลา','ช่วงเวลา','ชั่วโมง','เต็มวัน','สถานะการแจ้ง','ผลอนุมัติ','เหตุผล','แจ้งเมื่อ'];
  const rows=list.slice().sort((a,b)=>a.leave_date<b.leave_date?-1:1).map(x=>[fmtDMY(x.leave_date),x.emp_id,x.name,x.nationality,x.leave_type,x.slots,x.hours,(x.is_full_day?'เต็มวัน':''),x.filing_status,x.approval_status||'',(x.reason||''),x.filed_at]);
  // กัน Excel ตีความข้อความที่ขึ้นต้นด้วย = + - @ เป็นสูตร
  const escCsv=v=>{v=String(v==null?'':v); if(/^[=+\-@]/.test(v)) v="'"+v; return /[",\n]/.test(v)?('"'+v.replace(/"/g,'""')+'"'):v;};
  const csv=[head].concat(rows).map(r=>r.map(escCsv).join(',')).join('\r\n');
  const blob=new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8;'}); // BOM ให้ไทยไม่เพี้ยนใน Excel
  const url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url; a.download=APP_TAG+'_leaves_'+($('leavesMonth').value||'all')+($('leavesCycle').checked?'_cycle':'')+'.csv';
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),4000);
  toast('ดาวน์โหลดไฟล์แล้ว ✓');
}

/* ---------- ปฏิทินบริษัท: วันหยุด + วันทำงานพิเศษ ---------- */
async function loadCalendarAdmin(){
  $('holList').innerHTML='<div class="desc">กำลังโหลด...</div>'; $('workList').innerHTML='';
  try{
    const r=await apiGet('holidays');
    const row=(d,name,attr)=>'<div class="adminRow"><div class="rn"><b>'+esc(fmtDateLong(d,'th'))+'</b><small><span>'+esc(name)+'</span></small></div><button class="miniBtn del" type="button" '+attr+'="'+esc(d)+'">ลบ</button></div>';
    const hs=r.holidays||[], ws=r.work_overrides||[];
    $('holList').innerHTML=hs.length?hs.map(h=>row(h.date,h.name,'data-delhol')).join(''):'<div class="desc" style="margin:0">ยังไม่มีวันหยุด</div>';
    $('workList').innerHTML=ws.length?ws.map(w=>row(w.date,'🔁 '+w.name,'data-delwork')).join(''):'<div class="desc" style="margin:0">ยังไม่มีวันทำงานพิเศษ</div>';
    $('holList').querySelectorAll('[data-delhol]').forEach(b=>{b.onclick=()=>delCalendar('deleteHoliday','วันหยุด',b.dataset.delhol);});
    $('workList').querySelectorAll('[data-delwork]').forEach(b=>{b.onclick=()=>delCalendar('deleteWorkday','วันทำงานพิเศษ',b.dataset.delwork);});
  }catch(err){$('holList').innerHTML='<div class="empEmpty">⚠️ โหลดไม่สำเร็จ</div>';}
}
async function addCalendar(action,dateId,nameId,btnId,defName,label){
  const date=$(dateId).value, name=$(nameId).value.trim();
  if(!date){toast('เลือกวันที่ก่อน',true);return;}
  const btn=$(btnId), old=btn.textContent; btn.disabled=true; btn.textContent='บันทึก...';
  try{
    const r=await apiPost(action,{pin:adminPin,date:date,name:name||defName});
    if(r.ok){toast('เพิ่ม'+label+'แล้ว ✓');$(dateId).value='';$(nameId).value='';loadCalendarAdmin();loadHolidays();}
    else toast(r.error||'ผิดพลาด',true);
  }catch(err){toast('เชื่อมต่อไม่ได้',true);}
  btn.disabled=false; btn.textContent=old;
}
/* สลับวันหยุด: วันหยุด → มาทำงาน คู่กับ วันทำงาน → หยุดแทน */
function swapEcho(){
  const a=$('swWork').value, b=$('swOff').value, parts=[];
  if(a) parts.push('🟢 '+fmtDateLong(a,'th')+' มาทำงาน'+(isWorkDay(a)?' ⚠️ (วันนี้เป็นวันทำงานอยู่แล้ว)':''));
  if(b) parts.push('🔴 '+fmtDateLong(b,'th')+' หยุดแทน'+(!isWorkDay(b)?' ⚠️ (วันนี้เป็นวันหยุดอยู่แล้ว)':''));
  $('swEcho').innerHTML=parts.map(esc).join('<br>');
}
function saveSwap(){ return withBtn('swSaveBtn','กำลังบันทึก...',async()=>{
  const a=$('swWork').value, b=$('swOff').value;
  if(!a || !b) throw new Error('เลือกวันให้ครบทั้งสองช่อง');
  const c=await uiDialog({title:'สลับวันหยุด?', bodyHTML:'🟢 <b>'+esc(fmtDateLong(a,'th'))+'</b> — มาทำงาน<br>🔴 <b>'+esc(fmtDateLong(b,'th'))+'</b> — หยุดแทน', okText:'สลับ', cancelText:'ยกเลิก'});
  if(!c.ok) return;
  const r=await apiPost('calSwap',{pin:adminPin, work_date:a, off_date:b});
  if(!r.ok) throw new Error(r.error||'สลับไม่สำเร็จ');
  $('swWork').value=''; $('swOff').value=''; $('swEcho').innerHTML='';
  toast('สลับวันหยุดแล้ว ✓'); loadCalendarAdmin(); loadHolidays();
  if(r.leaves_on_off) uiDialog({title:'มีคนแจ้งลาไว้ในวันที่หยุดแทน', body:'วันที่ '+fmtDMY(b)+' มีคนแจ้งลาไว้แล้ว '+r.leaves_on_off+' คน\nใบลาเดิมยังอยู่ ถ้าไม่ต้องการให้นับ ลบได้ที่แท็บ "ใบลา"', alert:true});
}); }
async function delCalendar(action,label,date){
  const c=await uiDialog({title:'ลบ'+label+' '+fmtDMY(date)+'?', body:'ถ้าเป็นรายการที่มาจากการสลับวันหยุด อีกวันที่จับคู่กันจะถูกลบด้วย', okText:'ลบ', cancelText:'ไม่ลบ', danger:true});
  if(!c.ok) return;
  try{
    const r=await apiPost(action,{pin:adminPin,date:date});
    if(r.ok){toast('ลบแล้ว');loadCalendarAdmin();loadHolidays();}else toast(r.error||'ผิดพลาด',true);
  }catch(err){toast('เชื่อมต่อไม่ได้',true);}
}

/* ---------- ตั้งค่า: รอบเงินเดือน / Telegram / PIN ---------- */
async function loadAdminSettings(){
  try{
    const r=await apiGet('adminSettings',{pin:adminPin});
    if(!r.ok) return;
    adminCutoff=Number(r.cutoff_day)||25;
    $('setCutoff').value=r.cutoff_day||25;
    $('setCutoffMain').value=r.cutoff_day_main||'';
    $('setTgChat').value=r.telegram_chat_id||'';
    $('setTgToken').value='';
    $('tgTokenState').className='pill '+(r.telegram_token_set?'ok':'off');
    $('tgTokenState').textContent=r.telegram_token_set?'ตั้งไว้แล้ว':'ยังไม่ได้ตั้ง';
    $('verFoot').textContent='เวอร์ชันหน้าเว็บ '+(window.APP_VER.app||'?')+' · หลังบ้าน '+(r.ver||'?')+' · config '+(window.APP_VER.config||'?');
    $('pinWarn').classList.toggle('hidden',!r.pin_is_default);
    if($('setOtBack')){ $('setOtBack').value=(r.ot_back_days===undefined?7:r.ot_back_days); $('setOtConfirm').checked=(r.ot_need_confirm!==false); }
  }catch(e){}
}
async function setOne(key,value){ const r=await apiPost('setSetting',{pin:adminPin,key:key,value:value}); if(!r.ok) throw new Error(r.error||'บันทึกไม่สำเร็จ'); }
async function withBtn(btnId, busy, fn){
  const btn=$(btnId), old=btn.textContent; btn.disabled=true; btn.textContent=busy;
  try{ await fn(); }catch(err){ toast(err.message||'เชื่อมต่อไม่ได้',true); }
  btn.disabled=false; btn.textContent=old;
}
function saveCutoff(){ return withBtn('saveCutoffBtn','บันทึก...',async()=>{
  const a=$('setCutoff').value.trim();
  if(!(+a>=1 && +a<=28)) throw new Error('วันตัดรอบต้องเป็นเลข 1–28');
  await setOne('cutoff_day',a);
  adminCutoff=+a; toast('บันทึกรอบเงินเดือนแล้ว ✓');
}); }
function saveTelegram(){ return withBtn('saveTgBtn','บันทึก...',async()=>{
  const tok=$('setTgToken').value.trim();
  if(tok) await setOne('telegram_bot_token',tok);          // เว้นว่าง = ไม่แตะ token เดิม
  await setOne('telegram_chat_id',$('setTgChat').value.trim());
  toast('บันทึก Telegram แล้ว ✓'); loadAdminSettings();
}); }
function testTelegram(){ return withBtn('testTgBtn','กำลังส่ง...',async()=>{
  const r=await apiPost('testTelegram',{pin:adminPin});
  toast(r.ok?'ส่งทดสอบสำเร็จ ✓ ดูในกลุ่ม Telegram':('ส่งไม่สำเร็จ: '+(r.error||'')),!r.ok);
}); }
function savePin(){ return withBtn('savePinBtn','บันทึก...',async()=>{
  const p=$('setPin').value.trim();
  if(p.length<4) throw new Error('PIN ต้องยาวอย่างน้อย 4 ตัว');
  const c=await uiDialog({title:'เปลี่ยน PIN ผู้ดูแล?', body:'ครั้งต่อไปต้องใช้ PIN ใหม่ในการเข้า อย่าลืมจดไว้', okText:'เปลี่ยน', cancelText:'ยกเลิก'});
  if(!c.ok) return;
  await setOne('admin_pin',p); adminPin=p; $('setPin').value=''; toast('เปลี่ยน PIN แล้ว ✓');
}); }

/* ---------- ผูกปุ่มฝั่งแอดมิน ---------- */
$('adminLink').onclick=openPinModal;
$('secretDot').onclick=openPinModal;
$('pinCancel').onclick=()=>$('pinModal').classList.add('hidden');
$('pinSubmit').onclick=submitPin;
$('pinInput').onkeydown=e=>{if(e.key==='Enter')submitPin();};
$('adminClose').onclick=closeAdmin;
$('atEmp').onclick=()=>switchAdminTab('emp');
$('atDash').onclick=()=>switchAdminTab('dash');
$('atLeaves').onclick=()=>switchAdminTab('leaves');
$('atSet').onclick=()=>switchAdminTab('set');
$('adminSearch').oninput=debounce(renderAdminEmp,160);
$('addEmpBtn').onclick=()=>openEmpModal(null);
$('importEmpBtn').onclick=()=>{ $('impText').value=''; importRows=[]; $('impPreview').innerHTML=''; $('impSave').disabled=true; $('impSave').textContent='นำเข้า'; $('importModal').classList.remove('hidden'); };
$('impText').oninput=debounce(parseImport,200);
$('impNat').onchange=parseImport;
$('impCancel').onclick=()=>$('importModal').classList.add('hidden');
$('impSave').onclick=saveImport;
$('selectModeBtn').onclick=()=>setSelectMode(!selectMode);
$('moveCancelBtn').onclick=()=>setSelectMode(false);
$('moveAllBtn').onclick=()=>{ visibleAdminEmp().forEach(e=>selectedIds.add(e.emp_id)); renderAdminEmp(); };
$('moveToMainBtn').onclick=()=>doTransfer('MAIN');
$('moveToAltBtn').onclick=()=>doTransfer('ALT');
$('batchPhotoBtn').onclick=()=>{$('batchLog').innerHTML='';$('batchPhotos').value='';$('batchModal').classList.remove('hidden');};
$('syncDriveBtn').onclick=syncPhotosFromDrive;
$('batchClose').onclick=()=>$('batchModal').classList.add('hidden');
$('batchPhotos').onchange=handleBatchPhotos;
$('empCancel').onclick=()=>$('empModal').classList.add('hidden');
$('empSave').onclick=saveEmp;
$('empPhoto').onchange=handlePhotoUpload;
$('dashMonth').onchange=loadDashboard;
$('dashCycle').onchange=loadDashboard;
$('leavesMonth').onchange=loadLeaves;
$('leavesCycle').onchange=loadLeaves;
$('leavesSearch').oninput=debounce(renderLeaves,160);
$('exportBtn').onclick=exportLeavesCSV;
$('addHolBtn').onclick=()=>addCalendar('addHoliday','holDate','holName','addHolBtn','วันหยุด','วันหยุด');
$('addWorkBtn').onclick=()=>addCalendar('addWorkday','workDate','workName','addWorkBtn','วันทำงานพิเศษ','วันทำงานพิเศษ');
$('saveCutoffBtn').onclick=saveCutoff;
$('swSaveBtn').onclick=saveSwap;
$('swWork').onchange=swapEcho; $('swOff').onchange=swapEcho;
$('saveTgBtn').onclick=saveTelegram;
$('testTgBtn').onclick=testTelegram;
$('savePinBtn').onclick=savePin;

/* แตะโลโก้ 5 ครั้งเร็วๆ เพื่อเข้าแอดมิน (สำรองนอกจากลิงก์ท้ายหน้า) */
let logoTaps=0,logoTimer=null;
$('hdrLogo').onclick=()=>{logoTaps++;clearTimeout(logoTimer);logoTimer=setTimeout(()=>{logoTaps=0;},1200);if(logoTaps>=5){logoTaps=0;openPinModal();}};

console.log('[leave] admin.js v' + window.APP_VER.admin);

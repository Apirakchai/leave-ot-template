/* ============================================================
 * ระบบแจ้งลา + โอที — slip.js — v1.0
 * วาด "สลิปใบลา" เป็นรูป PNG (เหมือนสลิปโอนเงิน) พร้อม QR ให้สแกนตรวจกับระบบ
 * renderSlip(data) → Promise<dataURL>
 * data = { id, code, emp_id, name_th, name_mm, company, leave_type, dates[], slots[], hours_each,
 *          filing_status, approval, filed_at, skipped, state }
 * สลิปโอที: data.kind='ot' + { ot_date, range, hours, day_type, status, note }
 * สลิปเขียนสองภาษาเสมอ (ไทย + พม่า) เพื่อให้พนักงานและหัวหน้าอ่านได้ทั้งคู่
 * ============================================================ */
window.APP_VER = window.APP_VER || {};
window.APP_VER.slip = '1.0';

const SLIP_FONT = '"Noto Sans Thai Looped","Noto Sans Myanmar",system-ui,sans-serif';
let _slipLogo = null;

function slipLoadLogo(){
  if(_slipLogo) return _slipLogo;
  _slipLogo = new Promise(res=>{ const im=new Image(); im.onload=()=>res(im); im.onerror=()=>res(null); im.src=(typeof APP_LOGO!=='undefined'?APP_LOGO:'logo.png'); });
  return _slipLogo;
}
function slipFontsReady(){
  if(!document.fonts || !document.fonts.load) return Promise.resolve();
  const want=['400 30px "Noto Sans Thai Looped"','600 30px "Noto Sans Thai Looped"','700 30px "Noto Sans Thai Looped"','800 30px "Noto Sans Thai Looped"',
              '400 30px "Noto Sans Myanmar"','600 30px "Noto Sans Myanmar"','700 30px "Noto Sans Myanmar"'];
  const all=Promise.all(want.map(f=>document.fonts.load(f, f.indexOf('Myanmar')>=0?'ခွင့်':'ลา A1').catch(()=>null)));
  return Promise.race([all, new Promise(r=>setTimeout(r,2500))]);   // ฟอนต์โหลดช้า → วาดด้วยฟอนต์เครื่องไปก่อน ไม่ค้าง
}
/* ตัดบรรทัดให้พอดีความกว้าง (ภาษาไทย/พม่าไม่มีเว้นวรรค → ตัดตามคำด้วย Intl.Segmenter ถ้ามี) */
function slipWrap(ctx, text, maxW){
  text=String(text==null?'':text);
  if(ctx.measureText(text).width<=maxW) return [text];
  let parts;
  try{ parts=Array.from(new Intl.Segmenter('th',{granularity:'word'}).segment(text), s=>s.segment); }
  catch(e){ parts=text.split(/(\s+)/); }
  const lines=[]; let cur='';
  const push=()=>{ if(cur.trim()) lines.push(cur.trim()); cur=''; };
  parts.forEach(p=>{
    if(ctx.measureText(cur+p).width<=maxW){ cur+=p; return; }
    push();
    if(ctx.measureText(p).width<=maxW){ cur=p; return; }
    Array.from(p).forEach(ch=>{ if(ctx.measureText(cur+ch).width>maxW) push(); cur+=ch; });   // คำเดียวยาวเกิน → ตัดทีละตัว
  });
  push();
  return lines.length?lines:[text];
}
function slipRound(ctx,x,y,w,h,r){ ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); ctx.closePath(); }

function slipStatus(d){
  if(d.kind==='ot'){
    if(d.status==='ยืนยัน') return {th:'ยืนยันแล้ว', my:'အတည်ပြုပြီး', bg:'#E2F5EB', fg:'#157347'};
    if(d.status==='ไม่ยืนยัน') return {th:'ไม่ยืนยัน', my:'အတည်မပြုပါ', bg:'#FCE6E6', fg:'#B3261E'};
    if(d.status==='ยกเลิก') return {th:'ยกเลิกแล้ว', my:'ပယ်ဖျက်ပြီး', bg:'#FCE6E6', fg:'#B3261E'};
    return {th:'รอหัวหน้ายืนยัน', my:'အတည်ပြုရန် စောင့်နေသည်', bg:'#FFF1D6', fg:'#8A5A00'};
  }
  if(d.state==='cancelled') return {th:'ยกเลิกแล้ว', my:'ပယ်ဖျက်ပြီး', bg:'#FCE6E6', fg:'#B3261E'};
  if(d.approval==='รออนุมัติ') return {th:'รออนุมัติ', my:'ခွင့်ပြုချက် စောင့်နေသည်', bg:'#FFF1D6', fg:'#8A5A00'};
  if(d.approval==='ไม่อนุมัติ') return {th:'ไม่อนุมัติ', my:'ခွင့်မပြုပါ', bg:'#FCE6E6', fg:'#B3261E'};
  if(d.approval==='อนุมัติ') return {th:'อนุมัติแล้ว', my:'ခွင့်ပြုပြီး', bg:'#E2F5EB', fg:'#157347'};
  return {th:'ส่งใบลาแล้ว', my:'ခွင့်တင်ပြီး', bg:'#E2F5EB', fg:'#157347'};
}

async function renderSlip(d){
  const [logo]=await Promise.all([slipLoadLogo(), slipFontsReady()]);
  const W=1080, PAD=64, INNER=W-PAD*2, HEAD=240;
  const co=(typeof COMPANY_NAMES!=='undefined' && COMPANY_NAMES[d.company]) || {th:d.company||'', short:d.company||''};
  const dates=(d.dates||[]).slice().sort();
  const multi=dates.length>1;
  const hours=Number(d.hours_each)||((d.slots||[]).length);
  const slotTxt=(typeof compactSlots==='function')?compactSlots(d.slots||[]):(d.slots||[]).join(', ');
  const typeTh=typeLabel(d.leave_type,'th'), typeMy=typeLabel(d.leave_type,'my');
  const st=slipStatus(d);
  const isOT=(d.kind==='ot'), C1=isOT?'#3E7F14':'#1F4E79', C2=isOT?'#2C5F0C':'#163A5C', INK='#12283D';
  const fsKey=STATUS_KEY[d.filing_status], fsCol={st_adv:'#157347',st_same:'#A85D00',st_back:'#C62828'}[fsKey]||'#12283D';

  // ---- วัดความสูงก่อน (วาด 2 รอบ: รอบแรกวัด รอบสองวาดจริง) ----
  const measure=document.createElement('canvas').getContext('2d');
  function layout(ctx, draw){
    // y = ขอบบนของบรรทัดถัดไป (วางตัวอักษรที่ baseline = y + 1.12×ขนาด เผื่อสระบน/ล่างของไทย-พม่า)
    let y=HEAD+44;
    const put=(line,size,color)=>{ if(draw){ ctx.fillStyle=color; ctx.textAlign='left'; ctx.fillText(line, PAD, y+Math.round(size*1.12)); } y+=Math.round(size*1.62); };
    const label=(th,my)=>{ ctx.font='500 27px '+SLIP_FONT; put(th+'  ·  '+my, 27, '#5A6B7D'); y-=6; };
    const value=(txt,size,weight,color)=>{ ctx.font=(weight||700)+' '+size+'px '+SLIP_FONT; slipWrap(ctx,txt,INNER).forEach(line=>put(line,size,color||'#12283D')); };
    // รายการสั้น ๆ หลายตัว (เช่น วันที่) → ตัดบรรทัดที่ตัวคั่นเท่านั้น
    const tokens=(arr,size,weight,color)=>{ ctx.font=(weight||600)+' '+size+'px '+SLIP_FONT; let cur='';
      arr.forEach(t=>{ const nx=cur?cur+',  '+t:t; if(cur && ctx.measureText(nx).width>INNER){ put(cur+',',size,color); cur=t; } else cur=nx; });
      if(cur) put(cur,size,color); };
    const gap=n=>{ y+=n; };
    const rule=()=>{ y+=14; if(draw){ ctx.strokeStyle='#D6E0EA'; ctx.lineWidth=2; ctx.setLineDash([10,10]); ctx.beginPath(); ctx.moveTo(PAD,y); ctx.lineTo(W-PAD,y); ctx.stroke(); ctx.setLineDash([]); } y+=30; };

    // ผู้ลา
    value(d.name_th||'', 54, 800, '#12283D');
    if(d.name_mm) value(d.name_mm, 34, 600, '#3B4F63');
    value((d.emp_id||'')+(co.short?('   ·   '+co.short):''), 30, 600, '#5A6B7D');
    rule();

    if(isOT){
      const od=d.ot_date||'', odd=parseYMD(od), off=(d.day_type==='วันหยุด');
      label('รายการ','အမျိုးအစား');
      value('ทำโอที (OT)', 48, 800, C1);
      value('အိုတီ (အချိန်ပို)', 32, 600, C1);
      gap(16);
      label('วันที่ทำโอที','အိုတီဆင်းသည့်ရက်');
      value(fmtDMY(od), 62, 800, INK);
      if(odd) value(WD.th[odd.getDay()]+'  ·  '+WD.my[odd.getDay()], 32, 600, '#3B4F63');
      gap(16);
      label('เวลา','အချိန်');
      value(String(d.range||''), 44, 800, INK);
      value('รวม '+d.hours+' ชั่วโมง  ·  စုစုပေါင်း '+d.hours+' နာရီ', 34, 700, C1);
      gap(16);
      label('ประเภทวัน','ရက်အမျိုးအစား');
      value(off?'วันหยุด / วันอาทิตย์  ·  ပိတ်ရက်':'วันทำงานปกติ  ·  အလုပ်ရက်', 34, 700, off?'#9A4A00':INK);
      if(d.note){ gap(10); label('หมายเหตุ','မှတ်ချက်'); value(String(d.note).slice(0,160), 30, 500, '#3B4F63'); }
      rule();
      return y;
    }

    // ประเภท
    label('ประเภทการลา','ခွင့်အမျိုးအစား');
    value(typeTh, 48, 800, C1);
    if(typeMy && typeMy!==typeTh) value(typeMy, 32, 600, C1);
    gap(16);

    // วันที่
    label(multi?('วันที่ลา '+dates.length+' วัน'):'วันที่ลา', multi?('ခွင့်ရက် '+dates.length+' ရက်'):'ခွင့်ရက်');
    if(!multi){
      const one=dates[0]||'', dd=parseYMD(one);
      value(fmtDMY(one), 62, 800, '#12283D');
      if(dd) value(WD.th[dd.getDay()]+'  ·  '+WD.my[dd.getDay()], 32, 600, '#3B4F63');
    } else {
      value(fmtDMY(dates[0])+'  –  '+fmtDMY(dates[dates.length-1]), 46, 800, '#12283D');
      tokens(dates.map(x=>fmtDMY(x).slice(0,5)), 30, 600, '#3B4F63');
    }
    gap(16);

    // เวลา
    label(multi?'เวลาที่ลา (ต่อวัน)':'เวลาที่ลา', multi?'အချိန် (တစ်ရက်လျှင်)':'အချိန်');
    value(hours>=8 ? 'เต็มวัน  ·  တစ်နေကုန်  (8 ชม.)' : (slotTxt+'   ('+hours+' ชม. · '+hours+' နာရီ)'), 40, 700, '#12283D');
    gap(16);

    // สถานะการแจ้ง
    if(fsKey){
      label('สถานะการแจ้ง','အခြေအနေ');
      value(tr(fsKey,null,'th')+'  ·  '+tr(fsKey,null,'my'), 36, 700, fsCol);
      gap(10);
    }
    if(d.skipped>0){ value('ข้าม '+d.skipped+' วัน (วันหยุด/ลาซ้ำ)  ·  '+d.skipped+' ရက် ကျော်ထားသည်', 27, 500, '#8A5A00'); }
    rule();
    return y;
  }
  const bodyEnd=layout(measure,false);
  const FOOT=300, H=bodyEnd+FOOT;

  const cv=document.createElement('canvas'); cv.width=W; cv.height=H;
  const ctx=cv.getContext('2d');
  ctx.textBaseline='alphabetic';

  // ---- พื้น + หัวสลิป ----
  ctx.fillStyle='#FFFFFF'; ctx.fillRect(0,0,W,H);
  const g=ctx.createLinearGradient(0,0,W,HEAD); g.addColorStop(0,C1); g.addColorStop(1,C2);
  ctx.fillStyle=g; ctx.fillRect(0,0,W,HEAD);
  // รอยปรุ (ฟันปลา) ใต้หัวสลิป
  ctx.fillStyle='#FFFFFF';
  for(let x=18;x<W;x+=36){ ctx.beginPath(); ctx.arc(x,HEAD,11,0,Math.PI*2); ctx.fill(); }

  // ป้ายสถานะ (มุมขวาบนในแถบหัว)
  ctx.font='800 30px '+SLIP_FONT; const wTh=ctx.measureText(st.th).width;
  ctx.font='600 22px '+SLIP_FONT; const wMy=ctx.measureText(st.my).width;
  const pw=Math.max(wTh,wMy)+52, ph=96, px=W-PAD-pw, py=52;
  ctx.fillStyle=st.bg; slipRound(ctx,px,py,pw,ph,22); ctx.fill();
  ctx.fillStyle=st.fg; ctx.textAlign='center';
  ctx.font='800 30px '+SLIP_FONT; ctx.fillText(st.th, px+pw/2, py+42);
  ctx.font='600 22px '+SLIP_FONT; ctx.fillText(st.my, px+pw/2, py+78);

  let tx=PAD;
  if(logo){   // โลโก้สีบนแผ่นขาวมุมมน
    const lh=96, lw=Math.round(logo.width*lh/logo.height), bw=lw+28, bh=lh+24;
    ctx.fillStyle='#FFFFFF'; slipRound(ctx, PAD, 38, bw, bh, 22); ctx.fill();
    ctx.drawImage(logo, PAD+14, 50, lw, lh); tx=PAD+bw+26;
  }
  ctx.fillStyle='#FFFFFF'; ctx.textAlign='left';
  ctx.font='800 56px '+SLIP_FONT; ctx.fillText(isOT?'ใบแจ้งโอที':'ใบแจ้งลา', tx, 100);
  ctx.font='600 30px '+SLIP_FONT; ctx.globalAlpha=.92; ctx.fillText(isOT?'အိုတီစလစ်':'ခွင့်စလစ်', tx, 150);
  ctx.font='500 24px '+SLIP_FONT; ctx.globalAlpha=.85;
  ctx.fillText(slipWrap(ctx, co.th||'', W-PAD*2)[0]||'', PAD, 206);
  ctx.globalAlpha=1;

  // ---- เนื้อหา ----
  layout(ctx,true);

  // ---- ท้ายสลิป: QR + เลขที่ ----
  const fy=bodyEnd, QS=212;
  const url=((typeof SITE_URL!=='undefined' && SITE_URL)?SITE_URL.replace(/\/+$/,''):location.origin)+'/?slip='+encodeURIComponent(d.id||'')+'&c='+encodeURIComponent(d.code||'');
  try{
    const qr=qrcode(0,'M'); qr.addData(url); qr.make();
    const n=qr.getModuleCount(), cell=Math.floor(QS/(n+2)), size=cell*n, ox=PAD+Math.floor((QS-size)/2), oy=fy+Math.floor((QS-size)/2);
    ctx.fillStyle='#FFFFFF'; ctx.fillRect(PAD,fy,QS,QS);
    ctx.fillStyle='#12283D';
    for(let r=0;r<n;r++) for(let c=0;c<n;c++) if(qr.isDark(r,c)) ctx.fillRect(ox+c*cell, oy+r*cell, cell, cell);
  }catch(e){}
  const rx=PAD+QS+34; ctx.textAlign='left';
  ctx.fillStyle='#12283D'; ctx.font='700 30px '+SLIP_FONT; ctx.fillText('สแกนเพื่อตรวจสอบสลิป', rx, fy+40);
  ctx.fillStyle='#3B4F63'; ctx.font='500 25px '+SLIP_FONT; ctx.fillText('စစ်ဆေးရန် စကင်န်ဖတ်ပါ', rx, fy+80);
  ctx.fillStyle='#5A6B7D'; ctx.font='500 24px '+SLIP_FONT;
  ctx.fillText('เลขที่ · အမှတ်  '+(d.id||'-'), rx, fy+128);
  ctx.fillText('รหัสตรวจ  '+(d.code||'-'), rx, fy+164);
  ctx.fillText('แจ้งเมื่อ · တင်ချိန်  '+String(d.filed_at||'').slice(0,16), rx, fy+200);
  ctx.fillStyle='#8C9BAB'; ctx.font='500 22px '+SLIP_FONT; ctx.textAlign='center';
  ctx.fillText(((typeof SITE_URL!=='undefined' && SITE_URL)?SITE_URL.replace(/^https?:\/\//,'').replace(/\/+$/,''):location.host)+'  ·  ระบบแจ้งลา · โอที', W/2, H-30);

  // ยกเลิกแล้ว → คาดทับ
  if(d.state==='cancelled'){
    ctx.save(); ctx.translate(W/2,H/2); ctx.rotate(-Math.PI/9);
    ctx.strokeStyle='rgba(179,38,30,.55)'; ctx.lineWidth=8; ctx.fillStyle='rgba(179,38,30,.55)';
    ctx.font='800 110px '+SLIP_FONT; ctx.textAlign='center';
    const w=ctx.measureText('ยกเลิกแล้ว').width+80; slipRound(ctx,-w/2,-110,w,160,24); ctx.stroke();
    ctx.fillText('ยกเลิกแล้ว',0,10); ctx.restore();
  }
  return cv.toDataURL('image/png');
}
console.log('[leave] slip.js v' + window.APP_VER.slip);

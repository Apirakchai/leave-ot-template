/*************************************************************
 * ระบบแจ้งลา + โอที พนักงาน — หลังบ้าน (Google Apps Script) — v1.0
 * ใช้คู่กับหน้าเว็บใน repo เดียวกัน (index.html, app.js, ...) • ข้อมูลทั้งหมดอยู่ใน Google Sheet ของคุณเอง
 *
 * ติดตั้งครั้งแรก (รายละเอียดอยู่ใน README.md):
 *  1) สร้าง Google Sheet เปล่า → เมนู ส่วนขยาย > Apps Script → วางโค้ดนี้ทั้งหมด → แก้บล็อก "ตั้งค่าของบริษัท" ด้านล่าง → บันทึก
 *  2) เลือกฟังก์ชัน setup → Run (อนุญาตสิทธิ์ครั้งแรก) → ระบบสร้างชีตทั้งหมดให้เอง
 *  3) Deploy > New deployment > Web app : Execute as = Me, Who has access = Anyone → ได้ URL ลงท้าย /exec
 *  4) ใส่ URL นั้นใน config.js (API_URL) ของหน้าเว็บ แล้วเอาหน้าเว็บขึ้นโฮสต์
 *  5) Telegram (ถ้าใช้): ใส่ token ในหน้าผู้ดูแล → รันฟังก์ชัน connectTelegramNow
 *
 * อัปเดตโค้ดภายหลัง: วางทับ → Deploy > Manage deployments > ✏️ > New version (URL เดิม)
 *************************************************************/

/* ============================================================
 *  ⚙️ ตั้งค่าของบริษัท — แก้เฉพาะบล็อกนี้ก่อนรัน setup ครั้งแรก
 *  (ชื่อบริษัท/ที่อยู่ แก้ทีหลังได้ในชีต Settings • PIN แก้ได้ในหน้าผู้ดูแล → ตั้งค่า)
 * ============================================================ */
var APP = {
  NAME:  'บริษัทของฉัน',                    // ชื่อสั้น — ใช้ในข้อความ Telegram / ป้ายในชีต
  TITLE: 'MY COMPANY',                      // ชื่อบนหัวชีต "สรุป" และ "อ่านก่อน"
  FULL:  'บริษัท ตัวอย่าง จำกัด',            // ชื่อเต็ม (ค่าเริ่มต้นของ company_name ในชีต Settings)
  ADDR:  '',                                // ที่อยู่บริษัท (ค่าเริ่มต้นของ company_address)
  SITE:  '',                                // ที่อยู่หน้าเว็บ เช่น https://leave.example.com (ใช้ทำลิงก์ตรวจสลิป) — ใส่ทีหลังในชีต Settings (site_url) ก็ได้
  MENU:  'ระบบลา',                          // ชื่อเมนูที่โผล่บนแถบเมนูของ Google Sheet
  CO: 'MAIN', MULTI: false                  // ค่าภายในของระบบ — ไม่ต้องแก้
};

/* ====== ค่าคงที่ ====== */
var TZ = 'Asia/Bangkok';
var SHEETS = {
  EMP: 'Employees',
  LEAVE: 'Leaves',
  PAY_TH: 'Payroll_TH',
  PAY_MM: 'Payroll_MM',
  HOLIDAY: 'Holidays',
  WORKDAY: 'WorkExtra',
  CANCEL: 'CancelLog',
  DAILY_TH: 'รายวัน_TH',
  DAILY_MM: 'รายวัน_MM',
  DDATA: '_DailyData',
  DMETA: '_DailyMeta',
  SET: 'Settings'
};
var DAILY_TABLE_START = 8;  // แถวแรกของตารางรายวัน
var DEFAULT_PIN = '1234';        // PIN เริ่มต้นของหน้าผู้ดูแล — เปลี่ยนทันทีหลังติดตั้ง (หน้าผู้ดูแล → ตั้งค่า)                 // PIN แอดมินเริ่มต้น (เปลี่ยนได้ในชีท Settings)
var PHOTO_FOLDER_NAME = 'Leave_Photos';

/* URL ของ Web app (/exec) — ค่าสำรองสำหรับตั้ง Telegram webhook (ปุ่มอนุมัติลาพักร้อน)
 * ต้องตรงกับ API_URL ใน config.js — ระบบจะจำค่าล่าสุดลงชีต Settings (web_app_url) ให้อัตโนมัติเมื่อมีคนเปิดหน้าเว็บ */
var WEBAPP_EXEC_URL = '';   // เว้นว่างไว้: ระบบจำ URL /exec เองลงชีต Settings (web_app_url) เมื่อมีคนเปิดหน้าเว็บครั้งแรก

/* ประเภทการลา (ใช้แยกคอลัมน์ใน Payroll) */
var LEAVE_TYPES = ['ลาป่วย','ลากิจ','ลาพักร้อน'];

/* ช่วงเวลาทำงาน (นาฬิกา 24 ชม.) — 12:00-13:00 คือพักเที่ยง ไม่ให้ลา */
var TIME_SLOTS = [
  '08:00-09:00','09:00-10:00','10:00-11:00','11:00-12:00',
  '13:00-14:00','14:00-15:00','15:00-16:00','16:00-17:00'
];
var WORK_HOURS_PER_DAY = TIME_SLOTS.length; // = 8 ชั่วโมง/วัน

/* ====== Helper พื้นฐาน ====== */
function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name){ return ss_().getSheetByName(name); }

function json_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayStr_(){
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}
function nowStr_(){
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

/* ===== รอบตัดเงินเดือน (แยกตามบริษัทได้: cutoff_day_main ว่าง = ใช้ cutoff_day เดียวกับ ALT) ===== */
function getCutoffDay_(company){
  var c=NaN;
  if(company==='MAIN') c=parseInt(getSetting_('cutoff_day_main'),10);
  if(!(c>=1 && c<=28)) c=parseInt(getSetting_('cutoff_day'),10);
  return (c>=1 && c<=28)?c:25;
}
function thaiMonthAbbr_(m){ return ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][m]; }
function cycleForMonth_(d, company){
  var co=getCutoffDay_(company), y=d.getFullYear(), m=d.getMonth();
  var end=new Date(y,m,co), start=new Date(y,m-1,co+1);
  var fmt=function(x){ return x.getDate()+' '+thaiMonthAbbr_(x.getMonth()); };
  return { start:start, end:end, startStr:dateToStr_(start), endStr:dateToStr_(end),
    key:Utilities.formatDate(end,TZ,'yyyy-MM'),
    label:'รอบ '+thaiMonthAbbr_(m)+' ('+fmt(start)+'–'+fmt(end)+')' };
}

function getSetting_(key){
  var v=settingsMap_()[key];
  return (v===undefined||v===null)?'':v;
}
function setSetting_(key, value){
  var sh = sheet_(SHEETS.SET);
  var vals = sh.getDataRange().getValues();
  _setCache=null;
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim() === key){
      sh.getRange(i+1,2).setNumberFormat('@').setValue(String(value));   // เก็บเป็นข้อความ (PIN 0123 ต้องไม่กลายเป็น 123)
      return;
    }
  }
  var nr=sh.getLastRow()+1;
  sh.getRange(nr,2).setNumberFormat('@');
  sh.getRange(nr,1,1,2).setValues([[key, String(value)]]);
}

function checkPin_(pin){
  var real = getSetting_('admin_pin') || DEFAULT_PIN;
  return String(pin) === String(real);
}

/* สร้างรหัสพนักงานอัตโนมัติ (เมื่อไม่ได้กรอกรหัสเอง): คนไทย K001, K002 … / พม่า KM-01, KM-02 … */
function genEmpId_(nat, company){
  var main=(company==='MAIN'), mm=(nat==='MM');
  var prefix = mm ? 'KM-' : 'K';
  var pad = mm ? 2 : 3;
  var re = new RegExp('^'+prefix.replace('-','\\-')+'(\\d+)$','i');
  var vals = sheet_(SHEETS.EMP).getDataRange().getValues();
  var max = 0;
  for(var i=1;i<vals.length;i++){
    var m = re.exec(String(vals[i][0]).trim());
    if(m){ var n=parseInt(m[1],10); if(!isNaN(n) && n>max) max=n; }
  }
  var num = String(max+1);
  while(num.length<pad) num = '0'+num;
  return prefix+num;
}

/* สร้าง URL รูปจาก Drive file id (ใช้ CDN ของ Google เสถียรกว่า uc?export) */
function photoUrl_(fileId){
  if(!fileId) return '';
  return 'https://lh3.googleusercontent.com/d/' + fileId + '=w400';
}

/* ====== ติดตั้ง/อัปเดตโครงสร้าง (รันซ้ำได้ ไม่ลบข้อมูล) ====== */
function setupCore_(){
  var _lk=LockService.getScriptLock(); try{ _lk.tryLock(25000); }catch(e){} // กันชนตอนรัน setup พร้อมมีคนแจ้งลา
  var ss = ss_();
  var folder = getOrCreatePhotoFolder_();

  // 1) Settings — สร้างถ้ายังไม่มี / เติม key ที่ขาด (ไม่ทับค่าที่ตั้งไว้แล้ว เช่น PIN/Token)
  var st = sheet_(SHEETS.SET);
  if(!st){
    st = ss.insertSheet(SHEETS.SET);
    st.getRange(1,1,1,2).setValues([['key','value']]).setFontWeight('bold');
    st.setColumnWidth(1,200); st.setColumnWidth(2,420);
  }
  ensureSetting_('admin_pin', DEFAULT_PIN);
  ensureSetting_('photo_folder_id', folder.getId());
  ensureSetting_('company_name', APP.FULL);
  ensureSetting_('company_address', APP.ADDR);
  ensureSetting_('telegram_bot_token', '');
  ensureSetting_('telegram_chat_id', '');
  ensureSetting_('web_app_url', '');
  ensureSetting_('cutoff_day', '25');           // วันตัดรอบเงินเดือน (รอบ = 26 เดือนก่อน–25 เดือนนี้)
  ensureSetting_('summary_eve_hour', '17');
  ensureSetting_('summary_eve_min', '0');
  ensureSetting_('summary_morning_hour', '7');
  ensureSetting_('summary_morning_min', '30');
  ensureSetting_('site_url', APP.SITE);   // ใช้ทำลิงก์ตรวจสลิป
  ensureSetting_('tg_webhook_url', '');         // ว่าง = Telegram ยิงเข้า Apps Script ตรง ๆ / ใส่ https://<เว็บของคุณ>/tg = ผ่านตัวส่งต่อ (เร็ว+ปลอดภัยกว่า)
  ensureSetting_('tg_webhook_secret', '');
  slipSecret_();                                // สร้างรหัสลับสำหรับตรวจสลิป (ครั้งแรกครั้งเดียว)

  // 2) Employees — หัวตารางภาษาไทย + คำอธิบาย (ไม่ล้างข้อมูล)
  var emp = sheet_(SHEETS.EMP) || ss.insertSheet(SHEETS.EMP);
  ensureHeader_(emp,
    ['รหัสพนักงาน','ชื่อ (ไทย)','ชื่อ (พม่า/อังกฤษ)','สัญชาติ','รหัสรูป','แสดงชื่อ','ยังทำงานอยู่','ตำแหน่ง/หมายเหตุ','เพิ่มเมื่อ','บริษัท','(ระบบใช้เอง)','(ระบบใช้เอง)'],
    ['รหัสพนักงาน (ห้ามซ้ำ) — ถ้าจะเปลี่ยนรหัส ให้เปลี่ยนจากหน้าผู้ดูแลในเว็บ ระบบจะแก้ใบลาเก่าให้ครบ','ชื่อที่แสดงในหน้าเว็บและใบลา','ชื่ออักษรพม่าหรืออังกฤษ (ถ้ามี)','TH = คนไทย, MM = พม่า','รหัสไฟล์รูปใน Google Drive — ระบบใส่ให้เอง','TRUE = โชว์ในหน้าเลือกชื่อ, FALSE = ซ่อน','TRUE = ยังทำงาน, FALSE = ลาออก/ลบแล้ว (เก็บประวัติลาไว้)','ตำแหน่งหรือโน้ต','วันเวลาที่เพิ่มเข้าระบบ','MAIN','ช่องสำรอง ไม่ต้องกรอก','ช่องสำรอง ไม่ต้องกรอก']);
  emp.setFrozenRows(1);
  emp.setColumnWidths(1,1,100); emp.setColumnWidths(2,2,180); emp.setColumnWidths(10,3,120);
  emp.getRange('L2:L').setNumberFormat('dd/mm/yyyy');
  emp.getRange('A2:A').setNumberFormat('@');     // รหัสพนักงานเป็นข้อความเสมอ (0012 ต้องไม่กลายเป็น 12)

  // 3) Leaves — หัวตารางภาษาไทย + สีสถานะ (ไม่ล้างข้อมูล)
  var lv = sheet_(SHEETS.LEAVE) || ss.insertSheet(SHEETS.LEAVE);
  ensureHeader_(lv,
    ['แจ้งเมื่อ','เลขใบลา','รหัสพนักงาน','ชื่อ','สัญชาติ','วันที่ลา','ช่วงเวลาที่ลา','จำนวนชั่วโมง','เต็มวัน','ประเภทการลา','สถานะการแจ้ง','เหตุผล','ผลอนุมัติ','ผู้อนุมัติ','อนุมัติเมื่อ','รหัสกลุ่ม','อุปกรณ์','บริษัท'],
    ['วันเวลาที่กดส่งใบลา','เลขประจำใบลา (ขึ้นต้น L) — ใช้ตรวจสลิป','รหัสพนักงานที่ลา','ชื่อพนักงาน ณ ตอนแจ้ง','TH = ไทย, MM = พม่า','วันที่ขอลา (1 แถว = 1 วัน)','ช่วงเวลาที่ลา แยกเป็นชั่วโมง','จำนวนชั่วโมงที่ลาในวันนั้น (เต็มวัน = 8)','ขึ้นคำว่า เต็มวัน เมื่อลาครบ 8 ชม.','ลาป่วย / ลากิจ / ลาพักร้อน','ลาล่วงหน้า = แจ้งก่อนวันลา • ลากระทันหัน = แจ้งวันเดียวกับที่ลา • ลาย้อนหลัง = แจ้งหลังวันลา','เหตุผลที่พนักงานกรอก','ใช้กับลาพักร้อนเท่านั้น: รออนุมัติ / อนุมัติ / ไม่อนุมัติ (ลาป่วย-ลากิจเว้นว่าง = นับเลย)','ชื่อคนที่กดอนุมัติใน Telegram','วันเวลาที่กดอนุมัติ','ใบลาที่ส่งพร้อมกันหลายวันจะมีรหัสกลุ่มเดียวกัน (ขึ้นต้น G)','เครื่อง/เบราว์เซอร์ที่ใช้แจ้ง','บริษัท (MAIN) — ระบบใส่ให้เอง']);
  lv.setFrozenRows(1);
  lv.setColumnWidths(1,1,150); lv.setColumnWidths(6,1,110); lv.setColumnWidths(7,1,220); lv.setColumnWidths(12,1,260); lv.setColumnWidths(18,1,80);
  lv.getRange('F2:F').setNumberFormat('dd/mm/yyyy');
  lv.getRange('B2:C').setNumberFormat('@');
  applyLeaveFormatting_(lv);
  backfillLeaveCompany_();

  // 4) Holidays / WorkExtra
  var hd = sheet_(SHEETS.HOLIDAY) || ss.insertSheet(SHEETS.HOLIDAY);
  ensureHeader_(hd, ['วันหยุด (วันที่)','ชื่อวันหยุด','เพิ่มเมื่อ'], ['วันที่บริษัทหยุด — พนักงานลาทับไม่ได้','เช่น สงกรานต์','']);
  hd.setFrozenRows(1);
  hd.setColumnWidths(1,1,130); hd.setColumnWidths(2,1,260);
  hd.getRange('A2:A').setNumberFormat('dd/mm/yyyy');

  var wk = sheet_(SHEETS.WORKDAY) || ss.insertSheet(SHEETS.WORKDAY);
  ensureHeader_(wk, ['วันทำงานพิเศษ (วันที่)','หมายเหตุ','เพิ่มเมื่อ'], ['วันอาทิตย์/วันหยุดที่ต้องมาทำงาน (สลับวันหยุด) — ลาได้ตามปกติ','','']);
  wk.setFrozenRows(1);
  wk.setColumnWidths(1,1,160); wk.setColumnWidths(2,1,260);
  wk.getRange('A2:A').setNumberFormat('dd/mm/yyyy');

  // 4.1) CancelLog — ประวัติการยกเลิกใบลา
  var cl = sheet_(SHEETS.CANCEL) || ss.insertSheet(SHEETS.CANCEL);
  ensureHeader_(cl, ['ยกเลิกวันที่','ยกเลิกเวลา','เดือนของวันลา','รหัสพนักงาน','ชื่อ','สัญชาติ','วันที่ลา','ช่วงเวลา','ประเภทการลา','ชม.','สถานะเดิม','เหตุผลยกเลิก','บริษัท','เลขใบลา','รหัสกลุ่ม']);
  cl.setFrozenRows(1);
  cl.setColumnWidths(1,1,110); cl.setColumnWidths(2,1,90); cl.setColumnWidths(3,1,110);
  cl.setColumnWidths(4,1,100); cl.setColumnWidths(5,1,160); cl.setColumnWidths(7,1,110);
  cl.setColumnWidths(8,1,200); cl.setColumnWidths(12,1,260);
  cl.getRange('A2:A').setNumberFormat('dd/mm/yyyy'); cl.getRange('G2:G').setNumberFormat('dd/mm/yyyy'); cl.getRange('D2:D').setNumberFormat('@');

  // 5) ชีตเก็บถาวร (_DailyData / _DailyMeta) — ต้องมีก่อน Payroll (สูตรอ้างถึง)
  var dd = sheet_(SHEETS.DDATA) || ss.insertSheet(SHEETS.DDATA);
  ensureHeader_(dd, ['รหัส','ชื่อ','สัญชาติ','รอบ(เดือนจบ)','วันที่ลา','ประเภท','ช่วงเวลา','ชม.','หักเงิน','ค่าแรงต่อวัน','แก้เอง','บันทึกเมื่อ','บริษัท']);
  dd.setFrozenRows(1);
  dd.getRange('A2:A').setNumberFormat('@'); dd.getRange('D2:D').setNumberFormat('@');
  dedupeDailyData_();
  var dm = sheet_(SHEETS.DMETA) || ss.insertSheet(SHEETS.DMETA);
  ensureHeader_(dm, ['รหัส','ค่าแรงต่อวัน','โควต้าป่วยต่อปี','โควต้ากิจต่อปี','อัปเดตเมื่อ']);
  dm.setFrozenRows(1);
  dm.getRange('A2:A').setNumberFormat('@');

  // 6) Payroll (ไทย/พม่า)
  syncPayrollRosters_();

  // 7) หน้าสรุปรายวัน + ซ่อนชีตเก็บถาวร
  buildDailySheet_(SHEETS.DAILY_TH, 'TH');
  buildDailySheet_(SHEETS.DAILY_MM, 'MM');
  try{ dd.hideSheet(); dm.hideSheet(); }catch(e){}

  return 'Setup/อัปเดตเสร็จแล้ว ✔ (ข้อมูลเดิมไม่ถูกลบ)';
}

function ensureSetting_(key, def){
  var sh = sheet_(SHEETS.SET);
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++) if(String(vals[i][0]).trim()===key) return;
  sh.appendRow([key, def]);
  _setCache=null;
}

/* ใส่/อัปเดตหัวตาราง (แถว 1) โดยไม่แตะข้อมูลด้านล่าง + คำอธิบายแต่ละช่อง (เอาเมาส์ชี้ที่หัวคอลัมน์) */
function ensureHeader_(sh, headers, notes){
  var rng=sh.getRange(1,1,1,headers.length);
  rng.setValues([headers]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff')
     .setVerticalAlignment('middle').setWrap(true);
  if(notes && notes.length===headers.length) rng.setNotes([notes]);
}

function getOrCreatePhotoFolder_(){
  var existingId = getSetting_('photo_folder_id');
  if(existingId){
    try { return DriveApp.getFolderById(existingId); } catch(e){}
  }
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if(it.hasNext()) return it.next();
  var f = DriveApp.createFolder(PHOTO_FOLDER_NAME);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

/* สีสถานะใน Leaves: ล่วงหน้า=เขียว / กระทันหัน=เหลือง / ย้อนหลัง=แดง */
function applyLeaveFormatting_(lv){
  var rng = lv.getRange('K2:K1000'); // คอลัมน์ filing_status
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ย้อนหลัง').setBackground('#F8D7DA').setFontColor('#B02A37').setBold(true)
      .setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('กระทันหัน').setBackground('#FFF3CD').setFontColor('#997404')
      .setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('ล่วงหน้า').setBackground('#D1E7DD').setFontColor('#0F5132')
      .setRanges([rng]).build()
  ];
  lv.setConditionalFormatRules(rules);
}

/* รีเฟรชชีต Payroll จาก Employees (เฉพาะคนที่ยังทำงาน) แยก 2 ชีต: คนไทย / พม่า */
function syncPayrollRosters_(){
  var vals = sheet_(SHEETS.EMP).getDataRange().getValues();
  var R={TH:[],MM:[]};
  for(var i=1;i<vals.length;i++){
    var e=empFromRow_(vals[i], i+1);
    if(!e.emp_id || !e.active) continue;
    R[e.nationality].push({id:e.emp_id, name:e.name_th||e.name_mm, note:''});
  }
  rebuildPayroll_(SHEETS.PAY_TH, 'คนไทย', R.TH, 'MAIN');
  rebuildPayroll_(SHEETS.PAY_MM, 'พม่า', R.MM, 'MAIN');
  try{ buildSummarySheet_(); }catch(e){}
  try{ refreshDailyDropdown_(SHEETS.DAILY_TH,'TH'); refreshDailyDropdown_(SHEETS.DAILY_MM,'MM'); }catch(e){}
}

function rebuildPayrollCore_(name, label, rows, company){
  var ss = ss_();
  var sh = sheet_(name) || ss.insertSheet(name);
  company = normCompany_(company);
  var CO = getCutoffDay_(company);
  var L = SHEETS.LEAVE, DD = "'"+SHEETS.DDATA+"'", DM = "'"+SHEETS.DMETA+"'";
  var coCrit = '"<>-"';      // ระบบบริษัทเดียว: นับทุกใบลา
  var coOp = '<>"-"';

  // 1) ย้ายค่าแรงเดิม (โครงเก่า) -> _DailyMeta — อ่าน meta "ครั้งเดียว" แล้วเติมแบบ batch
  var idSet={}; rows.forEach(function(r){ idSet[r.id]=true; });
  var oldLast = sh.getLastRow();
  var oldWageById={};
  if(oldLast >= 3){
    var old = sh.getRange(1,1,oldLast,3).getValues();
    for(var i=2;i<old.length;i++){
      var oid=String(old[i][0]).trim();
      if(idSet[oid] && old[i][2]!=='' && old[i][2]!=null) oldWageById[oid]=old[i][2];
    }
  }
  var dmSheet=sheet_(SHEETS.DMETA);
  if(dmSheet){
    var dmVals=dmSheet.getDataRange().getValues(), metaRow={}, metaWage={};
    for(var im=1;im<dmVals.length;im++){ var mid=String(dmVals[im][0]).trim(); if(mid){ metaRow[mid]=im+1; metaWage[mid]=dmVals[im][1]; } }
    var addMeta=[];
    Object.keys(oldWageById).forEach(function(oid){
      var mw=metaWage[oid];
      if(mw===''||mw===undefined||mw===null){
        if(metaRow[oid]) dmSheet.getRange(metaRow[oid],2).setValue(oldWageById[oid]);
        else addMeta.push([oid, oldWageById[oid], '', '', nowStr_()]);
      }
    });
    if(addMeta.length) dmSheet.getRange(dmSheet.getLastRow()+1,1,addMeta.length,5).setValues(addMeta);
  }

  // 2) ล้างโครงเก่าที่ค้าง: เซลล์ผสาน/โน้ต/สี/validation (สาเหตุที่หัวคอลัมน์ลาป่วย-ลากิจ-ลาพักร้อนเคยหาย)
  var b1=sh.getRange('B1').getValue();
  var maxR=sh.getMaxRows(), maxC=Math.max(13, sh.getMaxColumns());
  try{ sh.getRange(1,1,maxR,maxC).breakApart(); }catch(e){}     // ทั้งชีต (ถ้าเลือกไม่ครบเซลล์ที่ผสาน จะ error)
  sh.getRange(1,1,2,maxC).clearContent().clearFormat().clearNote().clearDataValidations();
  if(maxR>=3){ var clr=sh.getRange(3,1,maxR-2,maxC); clr.clearContent(); clr.clearFormat(); clr.clearNote(); clr.clearDataValidations(); }
  try{ sh.clearConditionalFormatRules(); }catch(e){}

  // 3) แถว 1: รอบเงินเดือน + ช่วงรอบ
  sh.getRange('A1').setValue('รอบเงินเดือน ('+label+') :').setFontWeight('bold');
  if(!(b1 instanceof Date)){ b1=new Date(); b1.setDate(1); }
  sh.getRange('B1').setValue(b1).setNumberFormat('mmmm yyyy').setBackground('#FFF3CD').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('B1').setNote('พิมพ์วันที่ 1 ของเดือนที่รอบจบ เพื่อเปลี่ยนรอบ\nเช่น 1/10/2026 = รอบ '+(CO+1)+' ก.ย. – '+CO+' ต.ค. (ตัดรอบทุกวันที่ '+CO+')');
  sh.getRange('C1').setFormula('="รอบ "&TEXT($B$1,"mmm yyyy")&" ("&TEXT(DATE(YEAR($B$1),MONTH($B$1)-1,'+(CO+1)+'),"d/m/yyyy")&" – "&TEXT(DATE(YEAR($B$1),MONTH($B$1),'+CO+'),"d/m/yyyy")&")"')
    .setFontColor('#1F4E79').setFontWeight('bold');
  sh.getRange('F1').setValue('ⓘ หน้านี้สรุปยอดให้ดูอย่างเดียว • ค่าแรง + การติ๊กหัก/ไม่หัก ทำที่หน้า "รายวัน"').setFontColor('#5E768A');

  // 4) แถว 2: หัวตาราง 12 คอลัมน์ + คำอธิบาย
  var head=['รหัสพนักงาน','ชื่อพนักงาน','ค่าแรง/วัน (บาท)','ค่าแรง/ชม. (บาท)','ลาป่วย (ชม.)','ลากิจ (ชม.)','ลาพักร้อน (ชม.)','รวมลา (ชม.)','รวมลา (วัน)','ชม.ที่หักเงินจริง','ยอดหักรวม (บาท)','หมายเหตุ'];
  var notes=['','','ดึงจากหน้า "รายวัน" (ช่องค่าแรง/วัน)','ค่าแรง/วัน ÷ '+WORK_HOURS_PER_DAY,
    'ชั่วโมงลาป่วยในรอบนี้ (เฉพาะ '+company+')','ชั่วโมงลากิจในรอบนี้ (เฉพาะ '+company+')','ชั่วโมงลาพักร้อนที่อนุมัติแล้วในรอบนี้ (เฉพาะ '+company+')',
    'ลาป่วย + ลากิจ + ลาพักร้อน','รวมลา (ชม.) ÷ '+WORK_HOURS_PER_DAY,'ชั่วโมงที่ติ๊ก "หักเงิน" ในหน้ารายวัน (รอบเดียวกัน)','ชม.ที่หักเงินจริง × ค่าแรง/ชม.','หมายเหตุ'];
  sh.getRange(2,1,1,head.length).setValues([head]).setNotes([notes])
    .setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(2,44);
  sh.setFrozenRows(2);
  sh.setColumnWidths(1,1,100); sh.setColumnWidths(2,1,190); sh.setColumnWidths(3,2,100);
  sh.setColumnWidths(5,3,95); sh.setColumnWidths(8,2,90); sh.setColumnWidths(10,1,110);
  sh.setColumnWidths(11,1,130); sh.setColumnWidths(12,1,200);

  if(rows.length === 0){
    sh.getRange('A3').setValue('— ยังไม่มีพนักงาน '+label+' —').setFontColor('#999999');
    return;
  }

  // 5) เขียนข้อมูล (เริ่มแถว 3)
  function sumByType(rn, type){
    var lo='DATE(YEAR($B$1),MONTH($B$1)-1,'+(CO+1)+')';
    var hi='DATE(YEAR($B$1),MONTH($B$1),'+CO+')';
    return '=IFERROR(SUMIFS('+L+'!$H:$H,'+L+'!$C:$C,$A'+rn+','+
      L+'!$F:$F,">="&'+lo+','+L+'!$F:$F,"<="&'+hi+','+L+'!$J:$J,"'+type+'",'+
      L+'!$M:$M,"<>รออนุมัติ",'+L+'!$M:$M,"<>ไม่อนุมัติ",'+L+'!$R:$R,'+coCrit+'),0)';
  }
  var cycKey='TEXT($B$1,"yyyy-mm")';
  var out=[];
  for(var r=0;r<rows.length;r++){
    var rn=r+3, id=rows[r].id;
    out.push([
      id, rows[r].name,
      '=IFERROR(VLOOKUP($A'+rn+','+DM+'!$A:$B,2,FALSE),"")',                  // C ค่าแรง/วัน (จาก _DailyMeta)
      '=IF($C'+rn+'="","",$C'+rn+'/'+WORK_HOURS_PER_DAY+')',                   // D ค่าแรง/ชม.
      sumByType(rn,'ลาป่วย'), sumByType(rn,'ลากิจ'), sumByType(rn,'ลาพักร้อน'), // E,F,G
      '=$E'+rn+'+$F'+rn+'+$G'+rn,                                             // H รวมลา (ชม.)
      '=$H'+rn+'/'+WORK_HOURS_PER_DAY,                                        // I รวมลา (วัน)
      '=IFERROR(SUMPRODUCT(('+DD+'!$A$2:$A=$A'+rn+')*(TEXT('+DD+'!$D$2:$D,"yyyy-mm")='+cycKey+')*('+DD+'!$I$2:$I=TRUE)*('+DD+'!$M$2:$M'+coOp+')*('+DD+'!$H$2:$H)),0)', // J ชม.หักจริง (เทียบคีย์รอบแบบข้อความ ไม่ขึ้นกับว่าชีตเก็บเป็นวันที่หรือข้อความ)
      '=IF($D'+rn+'="",0,ROUND($J'+rn+'*$D'+rn+',2))',                         // K ยอดหัก
      rows[r].note||''
    ]);
  }
  var n=out.length;
  sh.getRange(3,1,n,1).setNumberFormat('@');
  sh.getRange(3,1,n,12).setValues(out);
  sh.getRange(3,3,n,2).setNumberFormat('#,##0.00');
  sh.getRange(3,5,n,4).setNumberFormat('#,##0.##');
  sh.getRange(3,9,n,1).setNumberFormat('#,##0.00');
  sh.getRange(3,10,n,1).setNumberFormat('#,##0.##').setBackground('#F2F8FD');
  sh.getRange(3,11,n,1).setNumberFormat('#,##0.00').setBackground('#E2F5EB').setFontWeight('bold');
  sh.getRange(3,12,n,1).setFontColor('#5E768A');
  // แถวรวม
  var tr=n+3;
  sh.getRange(tr,2).setValue('รวมทั้งหมด').setFontWeight('bold').setHorizontalAlignment('right');
  var sums=[]; ['E','F','G','H','I','J','K'].forEach(function(c){ sums.push('=SUM('+c+'3:'+c+(n+2)+')'); });
  sh.getRange(tr,5,1,7).setFormulas([sums]).setFontWeight('bold').setBackground('#EEF4F9');
  sh.getRange(tr,5,1,4).setNumberFormat('#,##0.##'); sh.getRange(tr,9,1,1).setNumberFormat('#,##0.00');
  sh.getRange(tr,10,1,1).setNumberFormat('#,##0.##'); sh.getRange(tr,11,1,1).setNumberFormat('#,##0.00');
}

/* ============================================================
 *  หน้าสรุปรายวัน (X+)  — รายวัน_TH / รายวัน_MM
 *  เลือกพนักงานจากดรอปดาวน์ -> ตารางทั้งรอบ + ติ๊กหักรายวัน
 *  จำค่าถาวรใน _DailyData (ติ๊ก) + _DailyMeta (ค่าแรง/โควต้า)
 * ============================================================ */
var DAILY_THAI_WD = ['อา.','จ.','อ.','พ.','พฤ.','ศ.','ส.'];

/* รายชื่อสำหรับดรอปดาวน์หน้ารายวัน (ทุกบริษัท ตามสัญชาติ) — คน MAIN มีป้ายต่อท้าย */
function dailyEmpList_(nat){
  var vals=sheet_(SHEETS.EMP).getDataRange().getValues(), out=[];
  for(var i=1;i<vals.length;i++){
    var e=empFromRow_(vals[i], i+1);
    if(!e.emp_id || !e.active || e.nationality!==nat) continue;
    out.push(e.emp_id+' — '+(e.name_th||e.name_mm)+'');
  }
  return out;
}
function dailyParseEmpId_(sel){ return String(sel||'').split('—')[0].trim(); }

/* ---- เก็บ/อ่าน meta (ค่าแรง + โควต้า) รายคน ---- */
function getDailyMeta_(empId){
  var sh=sheet_(SHEETS.DMETA); if(!sh) return {wage:'',qSick:'',qPersonal:''};
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++) if(String(vals[i][0]).trim()===empId)
    return {wage:vals[i][1], qSick:vals[i][2], qPersonal:vals[i][3]};
  return {wage:'',qSick:'',qPersonal:''};
}
function saveDailyMeta_(name, nat){
  var sh=sheet_(name), empId=dailyParseEmpId_(sh.getRange('B1').getValue());
  if(!empId) return;
  var wage=sh.getRange('B2').getValue(), qS=sh.getRange('B3').getValue(), qP=sh.getRange('B4').getValue();
  var meta=sheet_(SHEETS.DMETA), vals=meta.getDataRange().getValues(), at=nowStr_();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim()===empId){
      meta.getRange(i+1,2,1,4).setValues([[wage,qS,qP,at]]); return;
    }
  }
  meta.appendRow([empId,wage,qS,qP,at]);
}
/* เซ็ตเฉพาะค่าแรง (ใช้ย้ายค่าแรงเดิมจาก Payroll มาเก็บที่ meta) */
function setDailyMetaWage_(empId, wage){
  var meta=sheet_(SHEETS.DMETA); if(!meta) return;
  var vals=meta.getDataRange().getValues();
  for(var i=1;i<vals.length;i++) if(String(vals[i][0]).trim()===empId){ meta.getRange(i+1,2).setValue(wage); return; }
  meta.appendRow([empId,wage,'','',nowStr_()]);
}

/* อ่านการติ๊กหักของพนักงาน 1 คน ในช่วงวันที่ (จับคู่ด้วยวันลา ไม่ใช้คีย์รอบ) */
function getDailyTicks_(empId, fromStr, toStr){
  var sh=sheet_(SHEETS.DDATA), map={}; if(!sh) return map;
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim()!==empId) continue;
    var ds=dateKeyOf_(vals[i][4]);
    if(ds<fromStr || ds>toStr) continue;
    map[ds]={tick:(vals[i][8]===true), override:(vals[i][10]===true)};
  }
  return map;
}
function saveDailyTick_(name, nat, row){
  var sh=sheet_(name), empId=dailyParseEmpId_(sh.getRange('B1').getValue());
  if(!empId) return;
  var emp=findEmployee_(empId), nm=emp?(emp.name_th||emp.name_mm):'';
  var ds=dateKeyOf_(sh.getRange(row,1).getValue());
  if(!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return;
  var type=sh.getRange(row,3).getValue(), slots=sh.getRange(row,4).getValue(), hours=sh.getRange(row,5).getValue();
  if(!type) return;   // ไม่ใช่แถววันลา
  var tick=sh.getRange(row,6).getValue()===true;
  var wage=sh.getRange('B2').getValue();
  var co=companyOn_(emp, ds);
  var dd=sheet_(SHEETS.DDATA), vals=dd.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim()===empId && dateKeyOf_(vals[i][4])===ds){     // 1 คน 1 วัน = 1 แถว
      if(String(vals[i][12]||'').trim()) co=normCompany_(vals[i][12]);
      dd.getRange(i+1,1,1,13).setValues([[empId,nm,nat,cycleKeyForDate_(ds,co),ds,type,slots,hours,tick,wage,true,nowStr_(),co]]); return; // แก้เอง=true
    }
  }
  dd.appendRow([empId,nm,nat,cycleKeyForDate_(ds,co),ds,type,slots,hours,tick,wage,true,nowStr_(),co]);
}

/* ---- สร้างโครงหน้ารายวัน (คงค่าที่กรอกไว้) ---- */
function buildDailySheetCore_(name, nat){
  var ss=ss_(), sh=sheet_(name)||ss.insertSheet(name);
  sh.getRange('A1').setValue('เลือกพนักงาน:').setFontWeight('bold');
  var list=dailyEmpList_(nat);
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(list.length?list:['(ยังไม่มีพนักงาน)'],true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule).setBackground('#E7F1FA').setFontWeight('bold');
  sh.getRange('E1').setValue('รอบ (เดือนจบ):').setFontWeight('bold').setHorizontalAlignment('right');
  if(!(sh.getRange('F1').getValue() instanceof Date)){ var t=new Date(); t.setDate(1); sh.getRange('F1').setValue(t).setNumberFormat('mmmm yyyy').setBackground('#FFF3CD').setFontWeight('bold'); }
  sh.getRange('A2').setValue('ค่าแรง/วัน (บาท):').setFontWeight('bold');
  sh.getRange('B2').setBackground('#FFFBEA');
  sh.getRange('C2').setValue('ราคา/ชม.:').setHorizontalAlignment('right');
  sh.getRange('D2').setFormula('=IF($B$2="","",$B$2/'+WORK_HOURS_PER_DAY+')').setNumberFormat('#,##0.00').setFontWeight('bold');
  sh.getRange('A3').setValue('โควต้าลาป่วย/ปี (วัน):').setFontWeight('bold'); sh.getRange('B3').setBackground('#FFFBEA');
  sh.getRange('C3').setValue('ใช้ไป (วัน):').setHorizontalAlignment('right');
  sh.getRange('E3').setValue('เหลือ (วัน):').setHorizontalAlignment('right');
  sh.getRange('F3').setFormula('=IF($B$3="","",$B$3-$D$3)').setNumberFormat('#,##0.##').setFontWeight('bold');
  sh.getRange('A4').setValue('โควต้าลากิจ/ปี (วัน):').setFontWeight('bold'); sh.getRange('B4').setBackground('#FFFBEA');
  sh.getRange('C4').setValue('ใช้ไป (วัน):').setHorizontalAlignment('right');
  sh.getRange('E4').setValue('เหลือ (วัน):').setHorizontalAlignment('right');
  sh.getRange('F4').setFormula('=IF($B$4="","",$B$4-$D$4)').setNumberFormat('#,##0.##').setFontWeight('bold');
  sh.getRange('A5').setValue('ⓘ ติ๊กช่อง "หักเงิน?" เองรายวัน — ค่าที่ติ๊ก/ค่าแรง/โควต้า ระบบจำให้อัตโนมัติเมื่อแก้ • ลาพักร้อนที่ไม่อนุมัติจะไม่แสดง')
    .setFontColor('#5E768A').setFontSize(10);
  // หัวตาราง (แถว 7)
  var head=['วันที่','วัน','ประเภทลา','ช่วงเวลา','ชม.','หักเงิน?','หมายเหตุ'];
  sh.getRange(7,1,1,7).setValues([head]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#fff');
  sh.setFrozenRows(7);
  sh.setColumnWidths(1,1,100); sh.setColumnWidths(2,1,46); sh.setColumnWidths(3,1,110);
  sh.setColumnWidths(4,1,180); sh.setColumnWidths(5,1,60); sh.setColumnWidths(6,1,80); sh.setColumnWidths(7,1,220);
  rebuildDailyView_(name, nat);
}
function refreshDailyDropdown_(name, nat){
  var sh=sheet_(name); if(!sh) return;
  var list=dailyEmpList_(nat);
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(list.length?list:['(ยังไม่มีพนักงาน)'],true).setAllowInvalid(true).build();
  sh.getRange('B1').setDataValidation(rule);
}

/* ---- คำนวณการลา+ติ๊กหักอัตโนมัติตามโควต้า ของพนักงาน 1 คน 1 รอบ ----
 * - เกินโควต้า(ปี) = ติ๊กหักอัตโนมัติ / ในโควต้า = ไม่หัก
 * - วันที่ผู้ใช้ "แก้เอง" จะคงค่าไว้ (ไม่ทับ)
 * - วันเดียวมีหลายใบลา (เช่น เช้า 1 ชม. + บ่าย 1 ชม.) จะรวมชั่วโมงให้ครบ (เดิมนับแค่ใบสุดท้าย)
 * - บันทึกลง _DailyData (ให้ Payroll ดึงยอดไปแสดง)
 */
function computeDailyForEmp_(empId, nat, cycleDate){
  var emp=findEmployee_(empId), nm=emp?(emp.name_th||emp.name_mm):'';
  var cyc=cycleForMonth_(cycleDate, emp?emp.company:'ALT');
  var meta=getDailyMeta_(empId);
  var qSick=Number(meta.qSick)||0, qPers=Number(meta.qPersonal)||0, wage=meta.wage;
  var existing=getDailyTicks_(empId, cyc.startStr, cyc.endStr);  // ds -> {tick, override}
  var holMap=holidayMap_(), workSet=workExtraSet_();

  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var yr=cyc.end.getFullYear(), yLo=yr+'-01-01', yHi=yr+'-12-31';
  var dayMap={}, usedSick=0, usedPersonal=0, priorSick=0, priorPers=0;
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][2]).trim()!==empId) continue;
    var ap=String(vals[i][12]); if(ap==='รออนุมัติ'||ap==='ไม่อนุมัติ') continue;
    var ds=(vals[i][5] instanceof Date)?Utilities.formatDate(vals[i][5],TZ,'yyyy-MM-dd'):String(vals[i][5]).trim();
    var type=String(vals[i][9]), hrs=Number(vals[i][7])||0;
    if(ds>=yLo && ds<=yHi){ if(type==='ลาป่วย')usedSick+=hrs; else if(type==='ลากิจ')usedPersonal+=hrs; }
    if(ds>=yLo && ds<cyc.startStr){ if(type==='ลาป่วย')priorSick+=hrs; else if(type==='ลากิจ')priorPers+=hrs; }
    if(ds>=cyc.startStr && ds<=cyc.endStr){
      var co=String(vals[i][17]===undefined?'':vals[i][17]).trim();
      if(!dayMap[ds]) dayMap[ds]=[];
      dayMap[ds].push({type:type, slots:String(vals[i][6]||''), hours:hrs, company:(co?normCompany_(co):companyOn_(emp,ds))});
    }
  }

  var rows=[], cur=new Date(cyc.start), runS=priorSick, runP=priorPers;
  while(cur<=cyc.end){
    var d2=dateToStr_(cur), wd=DAILY_THAI_WD[cur.getDay()], items=dayMap[d2];
    if(items){
      var auto=false, notes=[], types=[], slotSet={}, hoursSum=0, coDay=items[0].company;
      for(var k=0;k<items.length;k++){
        var it=items[k];
        hoursSum+=it.hours;
        if(types.indexOf(it.type)<0) types.push(it.type);
        it.slots.split(',').forEach(function(s){ s=s.trim(); if(s) slotSet[s]=true; });
        var nt='';
        if(it.type==='ลาป่วย'){ runS+=it.hours; if(qSick>0 && runS>qSick*WORK_HOURS_PER_DAY){auto=true; nt='เกินโควต้าป่วย → หัก';} else nt='ในโควต้าป่วย'; }
        else if(it.type==='ลากิจ'){ runP+=it.hours; if(qPers>0 && runP>qPers*WORK_HOURS_PER_DAY){auto=true; nt='เกินโควต้ากิจ → หัก';} else nt='ในโควต้ากิจ'; }
        else if(it.type==='ลาพักร้อน'){ nt='พักร้อน (อนุมัติแล้ว)'; }
        if(nt && notes.indexOf(nt)<0) notes.push(nt);
      }
      var slotList=Object.keys(slotSet).sort(function(a,b){ return TIME_SLOTS.indexOf(a)-TIME_SLOTS.indexOf(b); }).join(', ');
      var ex=existing[d2], ovr=!!(ex&&ex.override), tick=ovr?ex.tick:auto;
      var note=notes.join(' • ');
      if(items.length>1) note+=' • รวม '+items.length+' ใบ';
      if(ovr) note+=' • แก้เอง';
      if(companyMix_().show) note+=' • '+coDay;
      rows.push({ds:d2,wd:wd,type:types.join(' + '),slots:slotList,hours:hoursSum,tick:tick,override:ovr,note:note,isLeave:true,company:coDay});
    } else {
      var dnote = workSet[d2] ? '(ทำงานพิเศษ)' : (holMap[d2] ? holMap[d2] : (cur.getDay()===0?'อาทิตย์':''));
      rows.push({ds:d2,wd:wd,type:'',slots:'',hours:'',tick:'',override:false,note:dnote,isLeave:false});
    }
    cur.setDate(cur.getDate()+1);
  }

  // ลบ _DailyData ของ emp+cycle เดิมทั้งหมด แล้วเขียนใหม่ (คง tick/override ที่อ่านไว้แล้ว)
  var dd=sheet_(SHEETS.DDATA), dv=dd.getDataRange().getValues(), del=[];
  for(var j=1;j<dv.length;j++){ var jd=dateKeyOf_(dv[j][4]); if(String(dv[j][0]).trim()===empId && jd>=cyc.startStr && jd<=cyc.endStr) del.push(j+1); }
  del.sort(function(a,b){return b-a;}).forEach(function(rn){ dd.deleteRow(rn); });
  var toAppend=[];
  // คีย์รอบของแต่ละวัน คิดตามวันตัดรอบของ "บริษัท ณ วันลา" → ตรงกับสูตรในชีต Payroll ของบริษัทนั้นเสมอ (แม้ ALT/MAIN ตัดรอบไม่ตรงกัน)
  rows.forEach(function(r){ if(r.isLeave) toAppend.push([empId,nm,nat,cycleKeyForDate_(r.ds,r.company),r.ds,r.type,r.slots,r.hours,(r.tick===true),wage,(r.override===true),nowStr_(),r.company]); });
  if(toAppend.length) dd.getRange(dd.getLastRow()+1,1,toAppend.length,13).setValues(toAppend);

  return {rows:rows, usedSick:usedSick, usedPersonal:usedPersonal, cyc:cyc};
}

/* ---- เติมตารางรายวันของคนที่เลือก (เรียก compute แล้ว render) ---- */
function rebuildDailyViewCore_(name, nat){
  var sh=sheet_(name); if(!sh) return;
  var empId=dailyParseEmpId_(sh.getRange('B1').getValue());
  var mv=sh.getRange('F1').getValue();
  if(!(mv instanceof Date)){ var t=new Date(); t.setDate(1); sh.getRange('F1').setValue(t).setNumberFormat('mmmm yyyy'); mv=t; }
  var empObj=empId?findEmployee_(empId):null;
  var cyc=cycleForMonth_(mv, empObj?empObj.company:'ALT');
  sh.getRange('G1').setValue(cyc.label+(empObj && companyMix_().show?('  •  '+coTag_(empObj)):'')).setFontColor('#1F4E79').setFontWeight('bold');

  sh.getRange(DAILY_TABLE_START,1,45,7).clearContent();
  sh.getRange(DAILY_TABLE_START,6,45,1).clearDataValidations();
  if(!empId || !empObj){ sh.getRange(DAILY_TABLE_START,1).setValue('— เลือกพนักงานด้านบน —').setFontColor('#999'); return; }

  var meta=getDailyMeta_(empId);
  sh.getRange('B2').setValue(meta.wage); sh.getRange('B3').setValue(meta.qSick); sh.getRange('B4').setValue(meta.qPersonal);

  var R=computeDailyForEmp_(empId, nat, mv);
  sh.getRange('D3').setValue(R.usedSick/WORK_HOURS_PER_DAY).setNumberFormat('#,##0.##');
  sh.getRange('D4').setValue(R.usedPersonal/WORK_HOURS_PER_DAY).setNumberFormat('#,##0.##');

  var tableRows=R.rows.map(function(r){ return [parseDateStr_(r.ds),r.wd,r.type,r.slots,r.hours,(r.isLeave?(r.tick===true):''),r.note]; });
  var n=tableRows.length;
  sh.getRange(DAILY_TABLE_START,1,n,7).setValues(tableRows).setFontColor('#000000');
  sh.getRange(DAILY_TABLE_START,1,n,1).setNumberFormat('dd/mm/yyyy');
  var cb=SpreadsheetApp.newDataValidation().requireCheckbox().build();
  R.rows.forEach(function(r,k){ if(r.isLeave) sh.getRange(DAILY_TABLE_START+k,6).setDataValidation(cb).setHorizontalAlignment('center'); });

  var s=DAILY_TABLE_START, eRow=s+n-1, sumRow=s+n+1;
  sh.getRange(sumRow,3).setValue('รวมชั่วโมงลา (รอบนี้):').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(sumRow,5).setFormula('=SUM(E'+s+':E'+eRow+')').setNumberFormat('#,##0.##').setFontWeight('bold');
  sh.getRange(sumRow+1,3).setValue('รวมชั่วโมงที่หักจริง:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(sumRow+1,5).setFormula('=SUMIF(F'+s+':F'+eRow+',TRUE,E'+s+':E'+eRow+')').setNumberFormat('#,##0.##').setFontWeight('bold').setBackground('#F2F8FD');
  sh.getRange(sumRow+2,3).setValue('ยอดหักรวม (บาท):').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(sumRow+2,5).setFormula('=ROUND(E'+(sumRow+1)+'*IF($B$2="",0,$B$2/'+WORK_HOURS_PER_DAY+'),2)').setNumberFormat('#,##0.00').setFontWeight('bold').setBackground('#E2F5EB');
}

/* ปุ่ม: คำนวณหักตามโควต้าให้ทุกคน (ทั้งไทย+พม่า) ตามรอบที่เลือกในแต่ละหน้า */
function computeAllDaily_(){
  withLock_(function(){
    var ss=ss_();
    [{name:SHEETS.DAILY_TH,nat:'TH'},{name:SHEETS.DAILY_MM,nat:'MM'}].forEach(function(o){
      var sh=sheet_(o.name); if(!sh) return;
      var mv=sh.getRange('F1').getValue(); if(!(mv instanceof Date)){ mv=new Date(); mv.setDate(1); }
      dailyEmpList_(o.nat).forEach(function(s){ var id=dailyParseEmpId_(s); if(id) computeDailyForEmp_(id, o.nat, mv); });
    });
    rebuildDailyView_(SHEETS.DAILY_TH,'TH'); rebuildDailyView_(SHEETS.DAILY_MM,'MM');
    ss.toast('คำนวณหักตามโควต้าให้ทุกคนแล้ว — ดูยอดได้ที่ Payroll','สำเร็จ',5);
  });
}

function refreshDailyBoth_(){ buildDailySheet_(SHEETS.DAILY_TH,'TH'); buildDailySheet_(SHEETS.DAILY_MM,'MM'); }

/* ---- onEdit: เปลี่ยนคน/เดือน -> rebuild • แก้ค่าแรง/โควต้า -> เซฟ • ติ๊ก -> เซฟ ---- */
function onEdit(e){
  try{
    if(!e||!e.range) return;
    var sh=e.range.getSheet(), name=sh.getName();
    if(name!==SHEETS.DAILY_TH && name!==SHEETS.DAILY_MM) return;
    var nat=(name===SHEETS.DAILY_MM)?'MM':'TH';
    var r=e.range.getRow(), c=e.range.getColumn();
    withLock_(function(){
      if(r===1 && (c===2 || c===6)){ rebuildDailyView_(name, nat); return; }   // เปลี่ยนคน(B1)/เดือน(F1)
      if(c===2 && (r===2||r===3||r===4)){ saveDailyMeta_(name, nat); return; } // ค่าแรง/โควต้า
      if(c===6 && r>=DAILY_TABLE_START){ saveDailyTick_(name, nat, r); }       // ติ๊กหักรายวัน
    });
  }catch(err){}
}

function doGet(e){
  var p = (e && e.parameter) || {};
  var action = p.action || 'ping';
  captureWebAppUrl_();   // จำ URL /exec ลงชีต Settings อัตโนมัติ (ทำครั้งเดียวตอนแรก)
  try{
    if(action === 'ping')       return json_({ok:true, time:nowStr_(), ver:BACKEND_VER});
    if(action === 'config')     return json_(getConfig_());
    if(action === 'checkpin')   return json_({ok: checkPin_(p.pin)});
    if(action === 'employees')  return json_(getEmployees_(p));
    if(action === 'history')    return json_(getHistory_(p));
    if(action === 'dashboard')  return json_(getDashboard_(p));
    if(action === 'leaves')     return json_(getLeavesAdmin_(p));
    if(action === 'holidays')   return json_(getHolidays_());
    if(action === 'verify')     return json_(verifySlip_(p));
    if(action === 'adminSettings') return json_(getAdminSettings_(p));
    if(action === 'otAdmin')    return json_(getOTAdmin_(p));
    if(action === 'otDay')      return json_(getOTDay_(p));
    return json_({ok:false, error:'unknown action: '+action});
  }catch(err){
    return json_({ok:false, error:String(err)});
  }
}

function doPost(e){
  var data = {};
  try { data = JSON.parse(e.postData.contents); } catch(err){ data = {}; }
  // ---- Telegram webhook (ปุ่มอนุมัติ + คำสั่งบอท) ----
  if(data && (data.callback_query || data.message || data.update_id)){
    try{
      if(!tgSecretOk_(e)) return json_({ok:true});
      if(tgSeenUpdate_(data.update_id)) return json_({ok:true});      // Telegram ยิงซ้ำ → ไม่ตอบซ้ำ
      if(data.callback_query) return handleTgCallback_(data.callback_query);
      if(data.message) return handleTgMessage_(data.message);
    }catch(err){}
    return json_({ok:true});
  }
  var action = data.action || '';
  try{
    return withLock_(function(){
    if(action === 'submitLeave')    return json_(submitLeave_(data));
    if(action === 'cancelLeave')    return json_(cancelLeave_(data));
    if(action === 'submitOT')       return json_(submitOT_(data));
    if(action === 'cancelOT')       return json_(cancelOT_(data));
    // ---- ต้องใส่ PIN แอดมิน ----
    if(['addEmployee','updateEmployee','deleteEmployee','toggleVisible','uploadPhoto','setEmpPhoto','syncPhotos','setSetting','testTelegram','updateLeaveStatus','deleteLeave','addHoliday','deleteHoliday','addWorkday','deleteWorkday','importEmployees','transferEmployees','changeEmpId','adminAddOT','updateOTStatus','deleteOT','calSwap']
        .indexOf(action) >= 0){
      if(!checkPin_(data.pin)) return json_({ok:false, code:'PIN', error:'PIN ไม่ถูกต้อง'});
    }
    if(action === 'addEmployee')    return json_(addEmployee_(data));
    if(action === 'updateEmployee') return json_(updateEmployee_(data));
    if(action === 'deleteEmployee') return json_(deleteEmployee_(data));
    if(action === 'toggleVisible')  return json_(toggleVisible_(data));
    if(action === 'uploadPhoto')    return json_(uploadPhoto_(data));
    if(action === 'setEmpPhoto')    return json_(setEmpPhoto_(data));
    if(action === 'syncPhotos')     return json_(syncPhotosFromDrive_(data));
    if(action === 'setSetting')     return json_(setSettingApi_(data));
    if(action === 'updateLeaveStatus') return json_(updateLeaveStatus_(data));
    if(action === 'deleteLeave')    return json_(deleteLeave_(data));
    if(action === 'addHoliday')     return json_(calAddHolidayApi_(data));
    if(action === 'calSwap')        return json_(calSwapApi_(data));
    if(action === 'deleteHoliday')  return json_(calRemoveApi_('H', data));
    if(action === 'addWorkday')     return json_(calAddWorkApi_(data));
    if(action === 'deleteWorkday')  return json_(calRemoveApi_('W', data));
    if(action === 'importEmployees')   return json_(importEmployees_(data));
    if(action === 'transferEmployees') return json_(transferEmployees_(data));
    if(action === 'changeEmpId')    return json_(changeEmpId_(data));
    if(action === 'adminAddOT')     return json_(adminAddOT_(data));
    if(action === 'updateOTStatus') return json_(updateOTStatus_(data));
    if(action === 'deleteOT')       return json_(deleteOT_(data));
    if(action === 'testTelegram')   return json_(tgSend_('✅ <b>ทดสอบ Telegram จากหน้าแอดมิน</b>\nเชื่อมต่อกลุ่มนี้เรียบร้อยแล้ว'));
    return json_({ok:false, error:'unknown action: '+action});
    }, true);
  }catch(err){
    if(String(err).indexOf('BUSY')>=0) return json_({ok:false, code:'BUSY', error:'ระบบกำลังมีคนใช้งานพร้อมกันหลายคน กรุณาลองใหม่อีกครั้ง'});
    return json_({ok:false, error:String(err)});
  }
}

/* ====== ฟังก์ชันสำหรับหน้าเว็บ ====== */

function getConfigCore_(){
  return {
    ok:true, ver:BACKEND_VER,
    company_name: getSetting_('company_name'),
    company_address: getSetting_('company_address'),
    time_slots: TIME_SLOTS,
    work_hours_per_day: WORK_HOURS_PER_DAY,
    cutoff_day: getCutoffDay_('ALT'), cutoff_day_main: getCutoffDay_('MAIN'),
    today: todayStr_()
  };
}

function getEmployees_(p){
  var wantAdmin = !!(p && p.admin === '1');
  var admin = wantAdmin && checkPin_(p.pin);
  if(wantAdmin && !admin) return {ok:false, code:'PIN', error:'PIN ไม่ถูกต้อง'};
  var vals = sheet_(SHEETS.EMP).getDataRange().getValues();
  var list = [], today = todayStr_();
  for(var i=1;i<vals.length;i++){
    var e = empFromRow_(vals[i], i+1);
    if(!e.emp_id || !e.active) continue;        // คนที่ลบ/ลาออก ไม่โชว์
    if(!admin && !e.visible) continue;           // หน้าเลือกชื่อ: เฉพาะที่เปิดให้แสดง
    var coNow = companyOn_(e, today);            // ย้ายล่วงหน้า (วันที่มีผลยังไม่ถึง) = ยังอยู่บริษัทเดิม
    list.push({
      emp_id: e.emp_id, name_th: e.name_th, name_mm: e.name_mm, nationality: e.nationality,
      photo_url: photoUrl_(e.photo_id), company: coNow, ex_alt: (e.ex_alt && coNow==='MAIN'),
      move_pending: admin ? (e.company!==coNow) : undefined,
      photo_id: admin ? e.photo_id : undefined,
      visible: admin ? e.visible : undefined,
      note: admin ? e.note : undefined,
      moved_at: admin ? e.moved_at : undefined
    });
  }
  return {ok:true, ver:BACKEND_VER, employees:list};
}

/* คำนวณสถานะการกรอก เทียบวันลา vs วันนี้ */
function computeFilingStatus_(leaveDateStr){
  var today = todayStr_();
  if(leaveDateStr > today)  return 'ลาล่วงหน้า';
  if(leaveDateStr === today) return 'ลากระทันหัน (วันเดียวกัน)';
  return 'ลาย้อนหลัง';
}

/* ---------- ตัวช่วยเรื่องวันที่/วันหยุด/ลาซ้ำ ---------- */
function parseDateStr_(s){ var p=s.split('-'); return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2])); }
function dateToStr_(dt){ return Utilities.formatDate(dt, TZ, 'yyyy-MM-dd'); }
/* แปลง yyyy-MM-dd -> dd/mm/yyyy (สำหรับแสดงผล) */
function fmtDMY_(s){ var m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m?(m[3]+'/'+m[2]+'/'+m[1]):String(s||''); }

/* ล็อกกันเขียนชนกัน (เขียนทีละคน) — ใช้เฉพาะ entry point ชั้นนอก ห้ามเรียกซ้อน
 * strict=true (คำขอจากหน้าเว็บ): ถ้ารอคิว 25 วิแล้วยังไม่ได้ล็อก จะไม่ทำต่อ (กันลบ/เขียนผิดแถว) แล้วให้ผู้ใช้ลองใหม่ */
function withLock_(fn, strict){
  var lock=LockService.getScriptLock(), got=false;
  try{ got=lock.tryLock(25000); }catch(e){}
  if(!got && strict) throw new Error('BUSY');
  try{ return fn(); }
  finally{ if(got){ try{ lock.releaseLock(); }catch(e){} } }
}

function isHoliday_(dateStr){
  var sh = sheet_(SHEETS.HOLIDAY);
  if(!sh) return null;
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    var d=vals[i][0];
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(ds===dateStr) return vals[i][1]||'วันหยุด';
  }
  return null;
}
/* วันทำงานพิเศษ (สลับวันหยุด) — วันที่อยู่ในชีต WorkExtra = ต้องมาทำงาน (ชนะวันหยุด/อาทิตย์) */
function isWorkExtra_(dateStr){
  var sh = sheet_(SHEETS.WORKDAY);
  if(!sh) return false;
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    var d=vals[i][0];
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(ds===dateStr) return true;
  }
  return false;
}
/* ★ สมองกลางตัวเดียว: วันนี้เป็นวันทำงานไหม (ใช้ร่วมทั้งการลา + แจ้งเตือนสรุป)
 *   ลำดับ: วันทำงานพิเศษ(ชนะทุกอย่าง) → วันหยุดบริษัท(หยุด) → อาทิตย์(หยุด) → ที่เหลือ(ทำงาน) */
function isWorkingDay_(dateStr){
  if(isWorkExtra_(dateStr)) return true;
  if(isHoliday_(dateStr)) return false;
  return parseDateStr_(dateStr).getDay()!==0;
}
/* อ่านทั้งชีตทีเดียว (ใช้ในลูปหลายวัน เพื่อความเร็ว) */
function holidayMap_(){
  var sh=sheet_(SHEETS.HOLIDAY), m={}; if(!sh) return m;
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){ var d=vals[i][0]; if(!d) continue;
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(/^\d{4}-\d{2}-\d{2}$/.test(ds)) m[ds]=vals[i][1]||'วันหยุด'; }
  return m;
}
function workExtraSet_(){
  var sh=sheet_(SHEETS.WORKDAY), m={}; if(!sh) return m;
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){ var d=vals[i][0]; if(!d) continue;
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(/^\d{4}-\d{2}-\d{2}$/.test(ds)) m[ds]=true; }
  return m;
}
/* ช่วงเวลาที่พนักงานคนนี้ลาไปแล้วในวันนั้น (กันลาซ้ำ) — ใบที่ "ไม่อนุมัติ" ไม่นับ */
function getTakenSlots_(empId, dateStr){
  var vals = sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var taken = {};
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][2]).trim()!==empId) continue;
    if(String(vals[i][12])==='ไม่อนุมัติ') continue;
    var d=vals[i][5];
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(ds!==dateStr) continue;
    String(vals[i][6]||'').split(',').forEach(function(s){ s=s.trim(); if(s) taken[s]=true; });
  }
  return taken;
}
/* เขียนใบลา 1 แถว คืนผลลัพธ์ */
function writeLeaveRow_(emp, dateStr, slots, leaveType, reason, groupId, device){
  var hours=slots.length;
  var isFull=(hours>=WORK_HOURS_PER_DAY);
  var status=computeFilingStatus_(dateStr);
  var leaveId='L'+Utilities.formatDate(new Date(),TZ,'yyMMddHHmmss')+Math.floor(Math.random()*900+100);
  var name=emp.name_th||emp.name_mm;
  var approval = (leaveType==='ลาพักร้อน') ? 'รออนุมัติ' : '';  // พักร้อนต้องอนุมัติ
  var company = companyOn_(emp, dateStr);
  var filedAt = nowStr_();
  var lv=sheet_(SHEETS.LEAVE);
  lv.appendRow([filedAt, leaveId, emp.emp_id, name, emp.nationality, parseDateStr_(dateStr),
    slots.join(', '), hours, isFull?'เต็มวัน':'', leaveType, status, safeText_(reason||''),
    approval, '', '', groupId||leaveId, safeText_(device||''), company]);
  lv.getRange(lv.getLastRow(),6).setNumberFormat('dd/mm/yyyy');
  return {leave_id:leaveId, hours:hours, is_full_day:isFull, filing_status:status, name:name, date:dateStr,
          approval:approval, company:company, filed_at:filedAt, slip_code:slipCode_(leaveId)};
}

/* วันเริ่มรอบปัจจุบัน (yyyy-MM-dd) — ใช้ล็อกห้ามลาย้อนหลังเกินรอบ */
function currentCycleStartStr_(company){
  var now=new Date(), co=getCutoffDay_(company);
  var dayNow=parseInt(Utilities.formatDate(now,TZ,'d'),10);
  var y=parseInt(Utilities.formatDate(now,TZ,'yyyy'),10), m=parseInt(Utilities.formatDate(now,TZ,'M'),10)-1;
  var em=(dayNow<=co)?new Date(y,m,1):new Date(y,m+1,1);
  return cycleForMonth_(em, company).startStr;
}

function submitLeave_(d){
  var empId = String(d.emp_id||'').trim();
  if(!empId) return err_('NO_EMP','ไม่พบรหัสพนักงาน');
  var emp = findEmployee_(empId);
  if(!emp || !emp.active) return err_('NO_EMP','ไม่พบพนักงานคนนี้');

  var slots = (d.slots||[]).filter(function(s,idx,arr){ return TIME_SLOTS.indexOf(s)>=0 && arr.indexOf(s)===idx; });
  if(slots.length === 0) return err_('NO_SLOT','กรุณาเลือกช่วงเวลาที่ลาอย่างน้อย 1 ช่วง');
  slots.sort(function(a,b){ return TIME_SLOTS.indexOf(a)-TIME_SLOTS.indexOf(b); });
  var leaveType = (LEAVE_TYPES.indexOf(d.leave_type)>=0) ? d.leave_type : 'ลากิจ';
  var reason = String(d.reason||'').slice(0,500);
  var override = (d.override===true || d.override==='1');
  var device = String(d.device||'').slice(0,80);
  var isVaca = (leaveType==='ลาพักร้อน');
  if(isVaca) slots = TIME_SLOTS.slice();   // พักร้อนนับเต็มวันเสมอ
  var groupId = 'G'+Utilities.formatDate(new Date(),TZ,'yyMMddHHmmss')+Math.floor(Math.random()*900+100);
  var minDate = currentCycleStartStr_(companyOn_(emp, todayStr_()));  // ห้ามลาย้อนหลังเกินรอบเงินเดือนปัจจุบัน
  var slip = {id:groupId, code:slipCode_(groupId)};

  // ===== โหมดลาหลายวัน (date_from .. date_to) =====
  if(d.date_from && d.date_to){
    var from=String(d.date_from).trim(), to=String(d.date_to).trim();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return err_('BAD_RANGE','ช่วงวันไม่ถูกต้อง');
    var dFrom=parseDateStr_(from), dTo=parseDateStr_(to);
    if(dTo<dFrom) return err_('END_BEFORE_START','วันสิ้นสุดต้องไม่ก่อนวันเริ่ม');
    if(from < minDate) return err_('TOO_OLD','ลาย้อนหลังเกินรอบเงินเดือนปัจจุบันไม่ได้ (ย้อนได้ถึง '+fmtDMY_(minDate)+')',{min_date:minDate});
    if(Math.round((dTo-dFrom)/86400000) > 61) return err_('TOO_LONG','ลาได้ครั้งละไม่เกิน 62 วัน');
    // ⚡ preload ครั้งเดียว
    var holMap=holidayMap_(), workSet=workExtraSet_();
    var takenByDate={};
    (function(){
      var lvAll=sheet_(SHEETS.LEAVE).getDataRange().getValues();
      for(var ti=1;ti<lvAll.length;ti++){
        if(String(lvAll[ti][2]).trim()!==empId) continue;
        if(String(lvAll[ti][12])==='ไม่อนุมัติ') continue;   // ใบที่ไม่อนุมัติ ไม่กันช่วงเวลา
        var td=lvAll[ti][5], tds=(td instanceof Date)?Utilities.formatDate(td,TZ,'yyyy-MM-dd'):String(td).trim();
        if(!takenByDate[tds]) takenByDate[tds]={};
        String(lvAll[ti][6]||'').split(',').forEach(function(s){ s=s.trim(); if(s) takenByDate[tds][s]=true; });
      }
    })();
    var created=[], skipped=[], cur=new Date(dFrom), guard=0;
    while(cur<=dTo && guard<70){
      guard++;
      var ds=dateToStr_(cur), why=null, whyCode='';
      if(workSet[ds]){ /* วันทำงานพิเศษ (สลับวันหยุด) — ลาได้ปกติ */ }
      else if(holMap[ds]){ why='วันหยุด: '+holMap[ds]; whyCode='HOLIDAY'; }
      else if(cur.getDay()===0){ why='วันอาทิตย์'; whyCode='SUNDAY'; }
      if(!why){
        var taken=takenByDate[ds]||{};
        if(slots.filter(function(s){return taken[s];}).length>0){ why='ลาซ้ำ'; whyCode='DUP'; }
      }
      if(why) skipped.push({date:ds, reason:why, code:whyCode});
      else created.push(writeLeaveRow_(emp, ds, slots, leaveType, reason, groupId, device));
      cur.setDate(cur.getDate()+1);
    }
    if(created.length===0) return err_('NONE_CREATED','ไม่มีวันที่ลาได้ในช่วงนี้ (วันหยุด/อาทิตย์/ลาซ้ำทั้งหมด)',{skipped:skipped});
    var base={ok:true, multi:true, created_count:created.length, created:created, skipped:skipped, hours_each:slots.length,
              slip:slip, filed_at:created[0].filed_at, company:created[0].company, name:created[0].name, leave_type:leaveType};
    if(isVaca){
      notifyVacationApproval_(emp, created, groupId, reason);
      base.pending_approval=true;
      return base;
    }
    notifyTelegramRange_(emp, created, skipped, slots, leaveType, reason);
    return base;
  }

  // ===== โหมดวันเดียว =====
  var leaveDate = String(d.leave_date||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(leaveDate)) return err_('BAD_DATE','วันที่ลาไม่ถูกต้อง');
  var dObj = parseDateStr_(leaveDate);
  if(leaveDate < minDate) return err_('TOO_OLD','ลาย้อนหลังเกินรอบเงินเดือนปัจจุบันไม่ได้ (ย้อนได้ถึง '+fmtDMY_(minDate)+')',{min_date:minDate});
  if(!isWorkExtra_(leaveDate)){   // วันทำงานพิเศษ (สลับวันหยุด) = ลาได้เลย ไม่ต้องเปิดสวิตช์
    var hname = isHoliday_(leaveDate);
    if(hname && !override) return err_('HOLIDAY','วันนี้เป็นวันหยุด ('+hname+') ไม่ต้องลา',{holiday:hname});
    if(dObj.getDay()===0 && !override) return err_('SUNDAY','วันอาทิตย์เป็นวันหยุด (ถ้าต้องมาทำงาน เปิดสวิตช์ก่อนส่ง)');
  }
  var taken1 = getTakenSlots_(empId, leaveDate);
  var conf = slots.filter(function(s){ return taken1[s]; });
  if(conf.length>0) return err_('DUP','ลาซ้ำ! ช่วงเวลานี้ลาไปแล้ว: '+conf.join(', '),{slots:conf});

  var res = writeLeaveRow_(emp, leaveDate, slots, leaveType, reason, groupId, device);
  var out = {ok:true, leave_id:res.leave_id, hours:res.hours, is_full_day:res.is_full_day, filing_status:res.filing_status, name:res.name,
             slip:slip, filed_at:res.filed_at, company:res.company, leave_date:leaveDate, slots:slots, leave_type:leaveType};
  if(isVaca){
    notifyVacationApproval_(emp, [res], groupId, reason);
    out.pending_approval=true;
    return out;
  }
  notifyTelegram_(emp, leaveDate, slots, res.hours, leaveType, res.filing_status, reason);
  return out;
}

function findEmployee_(empId){
  empId=String(empId||'').trim();
  if(!empId) return null;
  var vals = sheet_(SHEETS.EMP).getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]).trim() === empId) return empFromRow_(vals[i], i+1);
  }
  return null;
}

function getHistoryCore_(p){
  var empId = String(p.emp_id||'').trim();
  if(!empId) return err_('NO_EMP','ไม่พบรหัสพนักงาน');
  var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  var yr = nowStr.slice(0,4);
  var vals = sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var out = [], used={sick:0, personal:0, vacation:0, pending:0};
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][2]).trim() !== empId) continue;
    var ds = fmtDate_(vals[i][5]), ap=String(vals[i][12]||''), type=String(vals[i][9]||''), hrs=Number(vals[i][7])||0;
    if(ds.slice(0,4)===yr){
      if(ap==='รออนุมัติ') used.pending+=hrs;
      else if(ap!=='ไม่อนุมัติ'){
        if(type==='ลาป่วย') used.sick+=hrs; else if(type==='ลากิจ') used.personal+=hrs; else if(type==='ลาพักร้อน') used.vacation+=hrs;
      }
    }
    var lid=String(vals[i][1]);
    out.push({
      leave_id: lid,
      filed_at: fmtCell_(vals[i][0]),
      leave_date: ds,
      slots: vals[i][6],
      hours: vals[i][7],
      is_full_day: vals[i][8],
      leave_type: type,
      filing_status: vals[i][10],
      reason: vals[i][11],
      approval_status: ap,
      company: normCompany_(vals[i][17]),
      slip_code: slipCode_(lid),
      can_cancel: (nowStr < leaveStartStr_(ds, vals[i][6]))   // ยกเลิกได้เฉพาะก่อนถึงเวลาเริ่มลา
    });
  }
  out.sort(function(a,b){ return a.leave_date<b.leave_date?1:(a.leave_date>b.leave_date?-1:(a.filed_at<b.filed_at?1:-1)); }); // วันลาล่าสุดอยู่บน
  var meta=getDailyMeta_(empId);
  return {ok:true, history:out, year:yr,
    usage:{ sick_hours:used.sick, personal_hours:used.personal, vacation_hours:used.vacation, pending_hours:used.pending,
            sick_quota_days:Number(meta.qSick)||0, personal_quota_days:Number(meta.qPersonal)||0, hours_per_day:WORK_HOURS_PER_DAY }};
}

/* เวลาเริ่มลา (yyyy-MM-dd HH:mm) จากวันที่ + ช่วงเวลาแรกสุด */
function leaveStartStr_(leaveDateStr, slotsStr){
  var first = String(slotsStr||'').split(',')[0].trim();      // "08:00-09:00"
  var hhmm = first ? first.split('-')[0].trim() : '00:00';     // "08:00"
  if(!/^\d{2}:\d{2}$/.test(hhmm)) hhmm = '00:00';
  return leaveDateStr + ' ' + hhmm;
}

/* พนักงานยกเลิกใบลาของตัวเอง — ได้เฉพาะ "ก่อนถึงเวลาเริ่มลา" + ต้องมีเหตุผล */
function cancelLeave_(d){
  var empId = String(d.emp_id||'').trim();
  var leaveId = String(d.leave_id||'').trim();
  var reason = String(d.reason_cancel||'').trim().slice(0,500);
  if(!empId || !leaveId) return err_('MISSING','ข้อมูลไม่ครบ');
  if(!reason) return err_('NO_REASON','กรุณาใส่เหตุผลการยกเลิก');

  var sh = sheet_(SHEETS.LEAVE);
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][1]).trim() !== leaveId) continue;
    if(String(vals[i][2]).trim() !== empId) return err_('NOT_YOURS','ใบลานี้ไม่ใช่ของคุณ');

    var ds = (vals[i][5] instanceof Date) ? Utilities.formatDate(vals[i][5],TZ,'yyyy-MM-dd') : String(vals[i][5]).trim();
    var slots = vals[i][6], leaveType = vals[i][9], hours = vals[i][7], status = vals[i][10], name = vals[i][3], nat = vals[i][4];
    var company = normCompany_(vals[i][17]), groupId = String(vals[i][15]||'');

    // เช็คเวลา: ห้ามยกเลิกถ้าเลยเวลาเริ่มลาแล้ว (กันโกง)
    var nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    var startStr = leaveStartStr_(ds, slots);
    if(nowStr >= startStr){
      return err_('TOO_LATE','เลยเวลาเริ่มลาแล้ว ('+startStr+') — ยกเลิกไม่ได้',{start:startStr});
    }

    // บันทึกลงประวัติยกเลิก (ก่อนลบ)
    var nowD = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var nowT = Utilities.formatDate(new Date(), TZ, 'HH:mm:ss');
    var clog = sheet_(SHEETS.CANCEL);
    clog.appendRow([parseDateStr_(nowD), nowT, "'"+ds.slice(0,7), empId, name, nat, parseDateStr_(ds), slots, leaveType, hours, status, safeText_(reason), company, leaveId, groupId]);
    var lr = clog.getLastRow();
    clog.getRange(lr,1).setNumberFormat('dd/mm/yyyy'); clog.getRange(lr,7).setNumberFormat('dd/mm/yyyy');

    sh.deleteRow(i+1);

    // แจ้ง Telegram
    notifyCancelTelegram_(empId, name, nat, ds, slots, leaveType, hours, nowD+' '+nowT, reason);
    return {ok:true};
  }
  return err_('NOT_FOUND','ไม่พบใบลานี้');
}

/* แจ้ง Telegram เมื่อมีการยกเลิกใบลา */
function notifyCancelTelegram_(empId, name, nat, leaveDate, slots, leaveType, hours, cancelAt, reason){
  var token=getSetting_('telegram_bot_token'), chatId=getSetting_('telegram_chat_id');
  if(!token||!chatId) return;
  var nm = name;
  var emp = findEmployee_(empId);
  if(emp && emp.name_mm && nat==='MM') nm = (emp.name_th||name)+' / '+emp.name_mm;
  var text =
    '🚫 <b>ยกเลิกใบลา</b>\n' +
    '👤 ' + tgEscape_(nm) + ' (' + tgEscape_(empId) + ')\n' +
    tgCoLine_(emp, leaveDate) +
    '📅 วันลาที่ยกเลิก: ' + tgEscape_(fmtDMY_(leaveDate)) + '\n' +
    '⏰ ช่วงเวลา: ' + tgEscape_(compactSlots_(slots)) + '  (' + hours + ' ชม.)\n' +
    '📝 ประเภท: ' + tgEscape_(leaveType) + '\n' +
    '🕒 ยกเลิกเมื่อ: ' + tgEscape_(cancelAt) + '\n' +
    '💬 เหตุผลยกเลิก: ' + tgReasonValue_(reason);
  if(emp && emp.photo_id){ var r=tgSendPhoto_(emp.photo_id,text); if(!r.ok) tgSend_(text); }
  else tgSend_(text);
}

/* ---------- วันหยุดบริษัท ---------- */
function getHolidays_(){
  var sh = sheet_(SHEETS.HOLIDAY);
  if(!sh) return {ok:true, holidays:[]};
  var vals = sh.getDataRange().getValues();
  var out = [];
  for(var i=1;i<vals.length;i++){
    var d=vals[i][0]; if(!d) continue;
    var ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(ds)) continue;
    out.push({date:ds, name:vals[i][1]||'วันหยุด'});
  }
  out.sort(function(a,b){ return a.date<b.date?-1:1; });
  // วันทำงานพิเศษ (สลับวันหยุด)
  var wsh = sheet_(SHEETS.WORKDAY), wout = [];
  if(wsh){
    var wv = wsh.getDataRange().getValues();
    for(var j=1;j<wv.length;j++){
      var wd=wv[j][0]; if(!wd) continue;
      var wds=(wd instanceof Date)?Utilities.formatDate(wd,TZ,'yyyy-MM-dd'):String(wd).trim();
      if(!/^\d{4}-\d{2}-\d{2}$/.test(wds)) continue;
      wout.push({date:wds, name:wv[j][1]||'วันทำงานพิเศษ'});
    }
    wout.sort(function(a,b){ return a.date<b.date?-1:1; });
  }
  return {ok:true, holidays:out, work_overrides:wout};
}
function addWorkOverride_(d){
  var ds=String(d.date||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return {ok:false, error:'วันที่ไม่ถูกต้อง'};
  var sh=sheet_(SHEETS.WORKDAY);
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(dateKeyOf_(vals[i][0])===ds) return {ok:false, error:'มีวันทำงานพิเศษนี้อยู่แล้ว'};
  }
  sh.appendRow([parseDateStr_(ds), safeText_(d.name||'วันทำงานพิเศษ'), nowStr_()]);
  sh.getRange(sh.getLastRow(),1).setNumberFormat('dd/mm/yyyy');
  return {ok:true};
}
function deleteWorkOverride_(d){
  var ds=String(d.date||'').trim();
  var sh=sheet_(SHEETS.WORKDAY);
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    var dd=vals[i][0]; var dds=(dd instanceof Date)?Utilities.formatDate(dd,TZ,'yyyy-MM-dd'):String(dd).trim();
    if(dds===ds){ sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:'ไม่พบวันทำงานพิเศษนี้'};
}
function addHoliday_(d){
  var ds=String(d.date||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return {ok:false, error:'วันที่ไม่ถูกต้อง'};
  var sh=sheet_(SHEETS.HOLIDAY);
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(dateKeyOf_(vals[i][0])===ds) return {ok:false, error:'มีวันหยุดนี้อยู่แล้ว'};
  }
  sh.appendRow([parseDateStr_(ds), safeText_(d.name||'วันหยุด'), nowStr_()]);
  sh.getRange(sh.getLastRow(),1).setNumberFormat('dd/mm/yyyy');
  return {ok:true};
}
function deleteHoliday_(d){
  var ds=String(d.date||'').trim();
  var sh=sheet_(SHEETS.HOLIDAY);
  var vals=sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    var dd=vals[i][0]; var dds=(dd instanceof Date)?Utilities.formatDate(dd,TZ,'yyyy-MM-dd'):String(dd).trim();
    if(dds===ds){ sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:'ไม่พบวันหยุดนี้'};
}

/* Dashboard สรุปลา (เฉพาะแอดมิน) — ไม่นับใบที่ "ไม่อนุมัติ" */
function getDashboard_(p){
  if(!(p && checkPin_(p.pin))) return err_('PIN','PIN ไม่ถูกต้อง');
  var rg=adminRange_(p), filter=String(p.company||'').toUpperCase();
  var emps=empMap_();
  var vals = sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var map = {}, totalHours = 0, totalCount = 0, pending = 0, byType = {}, byStatus = {}, byCompany={ALT:0,MAIN:0};

  for(var i=1;i<vals.length;i++){
    var d = vals[i][5];
    if(!(d instanceof Date)) continue;
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if(ds<rg.from || ds>rg.to) continue;
    var ap=String(vals[i][12]||'');
    if(ap==='ไม่อนุมัติ') continue;
    var empId = String(vals[i][2]).trim(), e=emps[empId];
    var company = String(vals[i][17]||'').trim() ? normCompany_(vals[i][17]) : companyOn_(e, ds);
    var ex = !!(e && e.ex_alt);
    if(!coFilterMatch_(filter, company, ex)) continue;
    if(ap==='รออนุมัติ'){ pending++; continue; }

    var hours = Number(vals[i][7])||0, type = vals[i][9]||'อื่นๆ', status = vals[i][10]||'';
    if(!map[empId]) map[empId] = {emp_id:empId, name:vals[i][3], nationality:vals[i][4], company:(e?e.company:company), ex_alt:ex, hours:0, days:0, count:0, sick:0, personal:0, vacation:0};
    var a=map[empId];
    a.hours += hours; a.count += 1;
    if(type==='ลาป่วย') a.sick+=hours; else if(type==='ลากิจ') a.personal+=hours; else if(type==='ลาพักร้อน') a.vacation+=hours;
    totalHours += hours; totalCount += 1;
    byType[type] = (byType[type]||0)+hours;
    byStatus[status] = (byStatus[status]||0)+1;
    byCompany[company] = (byCompany[company]||0)+hours;
  }

  var people = Object.keys(map).map(function(k){ return map[k]; });
  people.sort(function(a,b){ return b.hours - a.hours; });
  people.forEach(function(x){ x.days = Math.round(x.hours/WORK_HOURS_PER_DAY*100)/100; });

  return {
    ok:true, month:rg.month, range:rg,
    summary:{ total_hours:totalHours, total_requests:totalCount, pending_requests:pending,
              total_days: Math.round(totalHours/WORK_HOURS_PER_DAY*100)/100,
              people_count: people.length },
    by_type: byType, by_status: byStatus, by_company: byCompany, people: people, mix: companyMix_()
  };
}

/* ====== แอดมิน: จัดการใบลา (แก้สถานะ / ลบ) ====== */

var LEAVE_STATUSES = ['ลาล่วงหน้า','ลากระทันหัน (วันเดียวกัน)','ลาย้อนหลัง'];

/* ดึงรายการใบลา (เฉพาะแอดมิน) */
function getLeavesAdmin_(p){
  if(!(p && p.admin === '1' && checkPin_(p.pin))) return err_('PIN','PIN ไม่ถูกต้อง');
  var useRange = /^\d{4}-\d{2}$/.test(String(p.month||'').trim());
  var rg = useRange ? adminRange_(p) : null;
  var emps=empMap_();
  var vals = sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var out = [];
  for(var i=1;i<vals.length;i++){
    var d = vals[i][5], ds = fmtDate_(d);
    if(rg){
      if(!(d instanceof Date)) continue;
      if(ds<rg.from || ds>rg.to) continue;
    }
    var empId=String(vals[i][2]).trim(), e=emps[empId];
    out.push({
      leave_id: vals[i][1], emp_id: empId, name: vals[i][3], nationality: vals[i][4],
      leave_date: ds, slots: vals[i][6], hours: vals[i][7], is_full_day: vals[i][8],
      leave_type: vals[i][9], filing_status: vals[i][10], reason: vals[i][11],
      approval_status: String(vals[i][12]||''), approver: String(vals[i][13]||''),
      company: (String(vals[i][17]||'').trim() ? normCompany_(vals[i][17]) : companyOn_(e, ds)),
      ex_alt: !!(e && e.ex_alt),
      filed_at: fmtCell_(vals[i][0])
    });
  }
  out.sort(function(a,b){ return a.leave_date<b.leave_date?1:(a.leave_date>b.leave_date?-1:(a.filed_at<b.filed_at?1:-1)); });
  return {ok:true, leaves:out, statuses:LEAVE_STATUSES, range:rg, mix:companyMix_()};
}

/* แก้สถานะการกรอกของใบลา (filing_status) */
function updateLeaveStatus_(d){
  var leaveId = String(d.leave_id||'').trim();
  var status = String(d.status||'').trim();
  if(!leaveId) return {ok:false, error:'ไม่พบ leave_id'};
  if(LEAVE_STATUSES.indexOf(status) < 0) return {ok:false, error:'สถานะไม่ถูกต้อง'};
  var sh = sheet_(SHEETS.LEAVE);
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][1]).trim() === leaveId){
      sh.getRange(i+1, 11).setValue(status); // คอลัมน์ filing_status (สีจะเปลี่ยนเองตามเงื่อนไข)
      return {ok:true, filing_status:status};
    }
  }
  return {ok:false, error:'ไม่พบใบลานี้'};
}

/* ลบใบลา (กรณีกรอกผิด) */
function deleteLeave_(d){
  var leaveId = String(d.leave_id||'').trim();
  if(!leaveId) return {ok:false, error:'ไม่พบ leave_id'};
  var sh = sheet_(SHEETS.LEAVE);
  var vals = sh.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][1]).trim() === leaveId){
      sh.deleteRow(i+1);
      return {ok:true};
    }
  }
  return {ok:false, error:'ไม่พบใบลานี้'};
}

/* ====== แอดมิน: จัดการพนักงาน ====== */

function addEmployee_(d){
  var nat = (d.nationality === 'MM') ? 'MM' : 'TH';
  var company = normCompany_(d.company);
  var id = String(d.emp_id||'').trim();
  if(id){
    if(!validEmpId_(id)) return err_('BAD_ID','รหัสพนักงานใช้ได้เฉพาะ A-Z 0-9 - _ . (ไม่เกิน 20 ตัว)');
    if(empIdExists_(id)) return err_('DUP_ID','รหัส '+id+' มีอยู่แล้ว');
  } else id = genEmpId_(nat, company);
  if(!String(d.name_th||'').trim() && !String(d.name_mm||'').trim()) return err_('NO_NAME','กรุณากรอกชื่อ');
  var sh = sheet_(SHEETS.EMP), nr = sh.getLastRow()+1;
  sh.getRange(nr,1).setNumberFormat('@');       // รหัสเป็นข้อความเสมอ
  sh.getRange(nr,1,1,12).setValues([[
    id, safeText_(String(d.name_th||'').trim()), safeText_(String(d.name_mm||'').trim()), nat, String(d.photo_id||''),
    true, true, safeText_(d.note||''), nowStr_(), company, false, ''
  ]]);
  _coMix=null;
  syncPayrollRosters_();
  return {ok:true, emp_id:id};
}

function updateEmployee_(d){
  var emp = findEmployee_(d.emp_id);
  if(!emp) return err_('NO_EMP','ไม่พบพนักงาน');
  var sh = sheet_(SHEETS.EMP);
  var row = emp.row;
  var needSync = false; // รีเฟรช payroll เฉพาะเมื่อชื่อ/สัญชาติ/บริษัทเปลี่ยน
  if(d.name_th !== undefined && String(d.name_th) !== String(emp.name_th||'')){ sh.getRange(row,2).setValue(safeText_(d.name_th)); needSync=true; }
  if(d.name_mm !== undefined && String(d.name_mm) !== String(emp.name_mm||'')){ sh.getRange(row,3).setValue(safeText_(d.name_mm)); needSync=true; }
  if(d.nationality !== undefined){ var nat=d.nationality==='MM'?'MM':'TH'; if(nat!==String(emp.nationality)){ sh.getRange(row,4).setValue(nat); needSync=true; } }
  if(d.photo_id !== undefined && d.photo_id !== '') sh.getRange(row,5).setValue(d.photo_id);
  if(d.note !== undefined) sh.getRange(row,8).setValue(safeText_(d.note));
  if(d.company !== undefined && d.company !== ''){
    var co=normCompany_(d.company);
    if(co!==emp.company){
      // แก้บริษัทตรง ๆ ได้เฉพาะคนที่ยังไม่มีใบลา (กรอกผิดตอนเพิ่ม) — คนที่มีใบลาแล้วต้องใช้ "ย้ายบริษัท" เพื่อไม่ให้ประวัติ/ยอดเงินเดือนเก่าเปลี่ยน
      if(empHasLeaves_(emp.emp_id)) return err_('USE_TRANSFER','พนักงานคนนี้มีใบลาแล้ว — เปลี่ยนบริษัทให้ใช้ปุ่ม "ย้ายบริษัท" (ระบุวันที่มีผล)');
      sh.getRange(row,10,1,3).setValues([[co, false, '']]);
      needSync=true; _coMix=null;
    }
  }
  if(needSync) syncPayrollRosters_();
  return {ok:true};
}

/* "ลบ" = ตั้ง active=false (เก็บประวัติลาไว้) */
function deleteEmployee_(d){
  var emp = findEmployee_(d.emp_id);
  if(!emp) return {ok:false, error:'ไม่พบพนักงาน'};
  sheet_(SHEETS.EMP).getRange(emp.row,7).setValue(false);
  syncPayrollRosters_();
  return {ok:true};
}

function toggleVisible_(d){
  var emp = findEmployee_(d.emp_id);
  if(!emp) return {ok:false, error:'ไม่พบพนักงาน'};
  var sh = sheet_(SHEETS.EMP);
  var cur = String(sh.getRange(emp.row,6).getValue()).toUpperCase() !== 'FALSE';
  sh.getRange(emp.row,6).setValue(!cur);
  return {ok:true, visible: !cur};
}

/* ซิงค์รูปจากโฟลเดอร์ Drive: จับคู่ชื่อไฟล์ (= รหัสพนักงาน) กับพนักงาน แล้วเซ็ต photo_id ทีเดียว */
function syncPhotosFromDrive_(d){
  var folder = getOrCreatePhotoFolder_();
  // 1) สแกนไฟล์รูปในโฟลเดอร์ -> map ชื่อไฟล์(ตัดนามสกุล ตัวเล็ก) -> file
  var it = folder.getFiles(), byName = {};
  while(it.hasNext()){
    var f = it.next();
    if(String(f.getMimeType()).indexOf('image/')!==0) continue;
    var base = f.getName().replace(/\.[^.]+$/,'').trim().toLowerCase();
    if(base) byName[base] = f;
  }
  // 2) วนพนักงาน จับคู่ตามรหัส
  var sh = sheet_(SHEETS.EMP);
  var vals = sh.getDataRange().getValues();
  var empIds = {}, matched = [], col = [];
  for(var i=1;i<vals.length;i++){
    var id = String(vals[i][0]).trim();
    var cur = vals[i][4];
    if(id) empIds[id.toLowerCase()] = true;
    var f2 = id ? byName[id.toLowerCase()] : null;
    if(f2){
      try{ f2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
      col.push([f2.getId()]);
      matched.push(id);
    } else {
      col.push([cur]); // คงรูปเดิม
    }
  }
  if(col.length) sh.getRange(2,5,col.length,1).setValues(col); // เขียนคอลัมน์ photo_id ทีเดียว
  // 3) หาไฟล์ที่ชื่อไม่ตรงรหัสใคร (เผื่อตั้งชื่อผิด)
  var unmatched = [];
  Object.keys(byName).forEach(function(k){ if(!empIds[k]) unmatched.push(byName[k].getName()); });
  return {ok:true, matched:matched.length, unmatched_files:unmatched, folder_url:folder.getUrl()};
}

/* อัปโหลดรูป (base64) -> เก็บใน Drive -> คืน file_id */
function uploadPhoto_(d){
  if(!d.base64) return {ok:false, error:'ไม่มีไฟล์รูป'};
  var folder = getOrCreatePhotoFolder_();
  var b64 = d.base64.indexOf(',') >= 0 ? d.base64.split(',')[1] : d.base64;
  var mime = d.mime || 'image/jpeg';
  var name = (d.filename || ('photo_'+Date.now())) ;
  var bytes = Utilities.base64Decode(b64);
  var blob = Utilities.newBlob(bytes, mime, name);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok:true, photo_id:file.getId(), photo_url:photoUrl_(file.getId())};
}

/* อัปโหลดรูป + เซ็ตให้พนักงานในครั้งเดียว (ใช้กับอัปหลายคน เพื่อลดรอบเรียก) */
function setEmpPhoto_(d){
  var emp = findEmployee_(d.emp_id);
  if(!emp) return {ok:false, error:'ไม่พบพนักงาน'};
  var up = uploadPhoto_(d);
  if(!up.ok) return up;
  sheet_(SHEETS.EMP).getRange(emp.row,5).setValue(up.photo_id); // ไม่ rebuild payroll (แค่รูป)
  return {ok:true, photo_id:up.photo_id, photo_url:up.photo_url};
}

function setSettingApi_(d){
  var key=String(d.key||'');
  if(SETTING_KEYS_ADMIN.indexOf(key)<0) return err_('BAD_KEY','ตั้งค่านี้จากหน้าเว็บไม่ได้');
  var value=String(d.value===undefined||d.value===null?'':d.value).trim();
  if(key==='cutoff_day' || key==='cutoff_day_main'){
    if(value!=='' || key==='cutoff_day'){
      var n=parseInt(value,10);
      if(!(n>=1 && n<=28)) return err_('BAD_VALUE','วันตัดรอบต้องเป็นเลข 1–28');
      value=String(n);
    }
  }
  if(key==='admin_pin' && value.length<4) return err_('BAD_VALUE','PIN ต้องยาวอย่างน้อย 4 ตัว');
  setSetting_(key, value);
  if(key==='cutoff_day' || key==='cutoff_day_main') syncPayrollRosters_();   // สูตรในชีต Payroll ฝังวันตัดรอบไว้
  return {ok:true};
}

/* ====== Telegram ======
 * แจ้งเตือนเข้ากลุ่ม Telegram ทุกครั้งที่มีคนแจ้งลา
 * ต้องตั้งค่า 2 ค่าในชีท Settings: telegram_bot_token, telegram_chat_id
 * (ดูวิธีหา chat id ได้จากฟังก์ชัน getTelegramChatId ด้านล่าง)
 * ถ้ายังไม่ได้ตั้งค่า จะข้ามไปเฉยๆ ไม่ error
 */

// escape อักขระพิเศษของ HTML กันข้อความพังเวลาเหตุผล/ชื่อมี < > &
function tgEscape_(s){
  return String(s===undefined||s===null?'':s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

/* ---------- แปลเหตุผลภาษาพม่า → ไทย (สำหรับข้อความ Telegram) ----------
 * ตรวจว่ามีอักษรพม่า (Unicode U+1000–U+109F) ปนอยู่ไหม ถ้ามีจึงแปล
 * ใช้ LanguageApp (Google Translate ในตัว Apps Script) — ฟรี ไม่ต้องใช้ API key
 * ถ้าแปลไม่ได้/ผิดพลาด คืนข้อความเดิม (ไม่ทำให้ส่งใบลาล้มเหลว) */
function maybeTranslateReason_(reason){
  var orig = String(reason===undefined||reason===null?'':reason);
  if(!orig) return {text:'', translated:false, original:''};
  if(!/[\u1000-\u109F]/.test(orig)) return {text:orig, translated:false, original:orig}; // ไม่มีอักษรพม่า
  try{
    var th = LanguageApp.translate(orig, 'my', 'th');
    if(th && String(th).trim() && String(th).trim()!==orig.trim()){
      return {text:String(th), translated:true, original:orig};
    }
  }catch(e){}
  return {text:orig, translated:false, original:orig};
}
/* คืน "ค่า" ของเหตุผล (แปลถ้าเป็นพม่า) — ผู้เรียกเติม label/ขึ้นบรรทัดเอง */
function tgReasonValue_(reason){
  var t = maybeTranslateReason_(reason);
  if(t.translated) return tgEscape_(t.text)+'  (🇲🇲 แปลจากพม่า)\n   ต้นฉบับ: '+tgEscape_(t.original);
  return tgEscape_(t.original);
}

/* ส่งข้อความพร้อมปุ่ม อนุมัติ/ไม่อนุมัติ (สำหรับลาพักร้อน) */
function tgSendButtons_(text, groupId, chatId){
  var token=getSetting_('telegram_bot_token');
  chatId = chatId || getSetting_('telegram_chat_id');
  if(!token||!chatId) return {ok:false};
  var kb={inline_keyboard:[[
    {text:'✅ อนุมัติ', callback_data:'apv:'+groupId},
    {text:'❌ ไม่อนุมัติ', callback_data:'rej:'+groupId}
  ]]};
  return tgSendTo_(chatId, text, kb);
}
function tgAnswerCallback_(cbId, text){
  var token=getSetting_('telegram_bot_token'); if(!token) return;
  try{ UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/answerCallbackQuery',{
    method:'post', contentType:'application/json',
    payload:JSON.stringify({callback_query_id:cbId, text:text||'', show_alert:false}),
    muteHttpExceptions:true }); }catch(e){}
}
function tgEditMessage_(chatId, msgId, html){
  var token=getSetting_('telegram_bot_token'); if(!token||!chatId||!msgId) return;
  try{ UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/editMessageText',{
    method:'post', contentType:'application/json',
    payload:JSON.stringify({chat_id:chatId, message_id:msgId, text:html, parse_mode:'HTML'}),
    muteHttpExceptions:true }); }catch(e){}
}

/* แจ้งคำขอลาพักร้อน + ปุ่มอนุมัติ */
function notifyVacationApproval_(emp, created, groupId, reason){
  var token=getSetting_('telegram_bot_token'), chatId=getSetting_('telegram_chat_id');
  if(!token||!chatId) return;
  tgSendButtons_(vacationRequestText_(emp, created.map(function(c){return c.date;}), reason), groupId);
}

/* รับการกดปุ่มจาก Telegram (webhook) */
function handleTgCallback_(cq){
  var data=String(cq.data||''), parts=data.split(':'), act=parts[0], gid=parts.slice(1).join(':');
  var clicker=((cq.from&&cq.from.first_name)||'')+((cq.from&&cq.from.username)?(' @'+cq.from.username):'');
  if(!clicker.trim()) clicker='ผู้ดูแล';
  clicker=safeText_(clicker).slice(0,80);
  var chat=cq.message&&cq.message.chat, chatId=chat&&chat.id;
  var msgId=cq.message&&cq.message.message_id;
  if(!tgAllowed_(chat, cq.from)){ tgAnswerCallback_(cq.id,'ใช้ได้เฉพาะสมาชิกกลุ่ม'); return json_({ok:true}); }

  if(act==='q' || act==='p'){          // ปุ่มเมนู: q = ส่งข้อความใหม่ · p = เปลี่ยนข้อความเดิม (แผงปุ่ม)
    var inPlace=(act==='p') && ['today','tomorrow','week','month','ot','cal'].indexOf(gid)>=0;
    tgAnswerCallback_(cq.id, inPlace?'อัปเดตแล้ว ✓':'');
    tgRunQuery_(gid, '', chatId, {editMsgId:(act==='p'?msgId:null), from:cq.from});
    return json_({ok:true});
  }
  if(act==='m'){                       // สลับชุดปุ่ม: หลัก ⇄ เพิ่มเติม
    tgAnswerCallback_(cq.id,'');
    tgApi_('editMessageReplyMarkup',{chat_id:chatId, message_id:msgId, reply_markup:(parts[1]==='more'?tgMoreKb_(parts[2]):tgMenuKb_(parts[2]))});
    return json_({ok:true});
  }
  if(act==='w'){                       // เลือกพนักงานจากผลค้นหา
    tgAnswerCallback_(cq.id,'');
    tgWhoReply_(chatId, gid, msgId);
    return json_({ok:true});
  }
  if(act==='oa' || act==='or') return tgOtCallback_(cq, act, gid, clicker, chatId, msgId);   // ยืนยันโอที
  if(act==='c') return tgCalCallback_(cq, parts, clicker, chatId, msgId);
  if(act==='e') return tgEmpCallback_(cq, parts, chatId, msgId);                     // ปฏิทิน: สลับวันหยุด ฯลฯ
  if(act!=='apv' && act!=='rej'){ tgAnswerCallback_(cq.id,''); return json_({ok:true}); }

  var lock=LockService.getScriptLock(), got=false;
  try{ got=lock.tryLock(15000); }catch(e){}
  try{
    var sh=sheet_(SHEETS.LEAVE), vals=sh.getDataRange().getValues();
    var rows=[], curStatus='', curApprover='', who='', dates=[];
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][15]).trim()===gid && String(vals[i][9])==='ลาพักร้อน'){
        rows.push(i+1); curStatus=String(vals[i][12]); curApprover=String(vals[i][13]);
        who=String(vals[i][3]); dates.push(fmtDMY_(fmtDate_(vals[i][5])));
      }
    }
    if(rows.length===0){ tgAnswerCallback_(cq.id,'ไม่พบคำขอนี้ (อาจถูกยกเลิกแล้ว)'); return json_({ok:true}); }
    if(curStatus!=='รออนุมัติ'){
      tgAnswerCallback_(cq.id,'ดำเนินการไปแล้ว: '+curStatus+(curApprover?(' โดย '+curApprover):''));
      return json_({ok:true});
    }
    var newStatus=(act==='apv')?'อนุมัติ':'ไม่อนุมัติ';
    var at=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm');
    rows.forEach(function(r){ sh.getRange(r,13,1,3).setValues([[newStatus, clicker, at]]); });
    var icon=(act==='apv')?'✅':'❌';
    var origText=tgEscape_((cq.message&&cq.message.text)||'');
    tgEditMessage_(chatId, msgId, origText+'\n\n'+icon+' <b>'+newStatus+'</b> โดย '+tgEscape_(clicker)+'\n🕒 '+at);
    tgAnswerCallback_(cq.id, icon+' '+newStatus+'แล้ว');
    // ถ้ากดจากแชตส่วนตัว (เมนู "รออนุมัติ") แจ้งกลุ่มด้วย ให้ทุกคนเห็นผล
    var groupChat=String(getSetting_('telegram_chat_id')||'').trim();
    if(groupChat && String(chatId)!==groupChat){
      tgSend_(icon+' <b>'+newStatus+'ลาพักร้อน</b>\n👤 '+tgEscape_(who)+'\n📅 '+tgEscape_(dates.join(', '))+'\nโดย '+tgEscape_(clicker));
    }
  } finally { if(got){ try{ lock.releaseLock(); }catch(e){} } }
  return json_({ok:true});
}

/* ตั้ง/ลบ webhook (รันจากเมนูหรือเปิด URL) */
/* คืน URL /exec ของ web app: ชีต Settings (web_app_url) > ค่าคงที่ WEBAPP_EXEC_URL > getUrl() */
function webAppUrl_(){
  var u=String(getSetting_('web_app_url')||'').trim();
  if(/\/exec$/.test(u)) return u;
  if(/\/exec$/.test(WEBAPP_EXEC_URL)) return WEBAPP_EXEC_URL;
  try{ var g=String(ScriptApp.getService().getUrl()||'').trim(); if(/\/exec$/.test(g)) return g; }catch(e){}
  return '';
}
/* จำ URL /exec ลงชีต Settings อัตโนมัติ (เรียกจาก doGet) — เขียนครั้งเดียวตอนยังว่าง */
function captureWebAppUrl_(){
  try{
    if(String(getSetting_('web_app_url')||'').trim()) return;
    var u=String(ScriptApp.getService().getUrl()||'').trim();
    if(/\/exec$/.test(u)) setSetting_('web_app_url', u);
  }catch(e){}
}
/* ตั้ง webhook: ถ้า Settings มี tg_webhook_url (เช่น https://<เว็บของคุณ>/tg) จะยิงผ่านตัวส่งต่อ + รหัสลับ */
function setTelegramWebhook(){
  var token=getSetting_('telegram_bot_token');
  if(!token) return 'ยังไม่มี telegram_bot_token ในชีท Settings';
  var relay=String(getSetting_('tg_webhook_url')||'').trim();
  var url=relay||webAppUrl_();
  if(!url) return 'ไม่พบ URL /exec — วาง API_URL (/exec) ในชีท Settings แถว web_app_url';
  var sec=String(getSetting_('tg_webhook_secret')||'').trim();
  if(!sec){ sec=Utilities.getUuid().replace(/-/g,''); setSetting_('tg_webhook_secret', sec); }
  var payload={allowed_updates:['message','callback_query']};
  if(relay){ payload.url=relay; payload.secret_token=sec; }      // ผ่านตัวส่งต่อ: Telegram แนบรหัสลับในหัวคำขอ ตัวส่งต่อใส่ ?tgs= ให้
  else payload.url=url+'?tgs='+sec;                              // ยิงตรง: รหัสลับอยู่ใน URL ของ webhook (มีแค่ Telegram ที่รู้)
  var r=tgApi_('setWebhook', payload);
  tgSetCommands_();
  return 'URL: '+(relay||url)+(relay?' (ผ่านตัวส่งต่อ)':' (ยิงตรง + รหัสลับ)')+'\n'+JSON.stringify(r);
}
/* ★ แก้ปุ่มอนุมัติไม่ทำงาน: เลือกฟังก์ชันนี้แล้วกด Run ได้เลยจาก Apps Script (ไม่ต้อง deploy ใหม่) */
function setWebhookNow(){ menuSetWebhook_(); }
/* เมนู: ตั้ง webhook (ทำให้ปุ่มอนุมัติ + คำสั่งบอทใน Telegram ทำงาน) */
function menuSetWebhook_(){
  var r=setTelegramWebhook();
  _alert_('ตั้ง Webhook + เมนูคำสั่งบอท:\n\n'+r+'\n\nถ้าขึ้น "ok":true = สำเร็จ → ลองพิมพ์ /menu ในกลุ่ม Telegram');
}
/* เมนู: เช็คสถานะ webhook ปัจจุบัน */
function menuCheckWebhook_(){
  var token=getSetting_('telegram_bot_token');
  if(!token){ _alert_('ยังไม่มี telegram_bot_token ในชีท Settings'); return; }
  try{
    var res=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/getWebhookInfo', {muteHttpExceptions:true});
    _alert_('สถานะ Webhook ปัจจุบัน:\n\n'+res.getContentText()+'\n\n(ดู "url" ว่าตรงกับ /exec ปัจจุบันไหม · "last_error_message" มี error ไหม)');
  }catch(e){ _alert_('ผิดพลาด: '+e); }
}

// ส่งข้อความเข้า Telegram (คืน response object) — ใช้ภายใน
function tgSend_(text){
  var token = getSetting_('telegram_bot_token');
  var chatId = getSetting_('telegram_chat_id');
  if(!token || !chatId) return {ok:false, error:'ยังไม่ได้ตั้งค่า token หรือ chat_id ในชีท Settings'};
  try{
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendMessage', {
      method:'post',
      contentType:'application/json',
      payload: JSON.stringify({
        chat_id: chatId, text: text, parse_mode:'HTML', disable_web_page_preview:true
      }),
      muteHttpExceptions:true
    });
    var body = JSON.parse(res.getContentText());
    return body.ok ? {ok:true} : {ok:false, error: body.description || 'ส่งไม่สำเร็จ'};
  }catch(e){
    return {ok:false, error:String(e)};
  }
}

// ส่งรูป + คำบรรยาย (caption) เข้า Telegram โดยดึงไฟล์จาก Drive
function tgSendPhoto_(fileId, caption){
  var token = getSetting_('telegram_bot_token');
  var chatId = getSetting_('telegram_chat_id');
  if(!token || !chatId || !fileId) return {ok:false, error:'no token/chat/file'};
  try{
    var blob = DriveApp.getFileById(fileId).getBlob();
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/sendPhoto', {
      method:'post',
      payload:{ chat_id:chatId, caption:caption, parse_mode:'HTML', photo:blob },
      muteHttpExceptions:true
    });
    var body = JSON.parse(res.getContentText());
    return body.ok ? {ok:true} : {ok:false, error: body.description || 'ส่งรูปไม่สำเร็จ'};
  }catch(e){
    return {ok:false, error:String(e)};
  }
}

function notifyTelegram_(emp, leaveDate, slots, hours, leaveType, status, reason){
  var token = getSetting_('telegram_bot_token');
  var chatId = getSetting_('telegram_chat_id');
  if(!token || !chatId) return; // ยังไม่เปิดใช้

  var statusIcon = status.indexOf('ย้อนหลัง')>=0 ? '🔴' :
                   (status.indexOf('กระทันหัน')>=0 ? '🟡' : '🟢');
  var name = emp.name_th || emp.name_mm;
  if(emp.name_mm && emp.nationality==='MM') name = (emp.name_th||'') + ' / ' + emp.name_mm;

  // เน้นพิเศษถ้าเป็นลาย้อนหลัง/กระทันหัน (แอดมินต้องจับตา)
  var alertHead = '';
  if(status.indexOf('ย้อนหลัง')>=0) alertHead = '⚠️⚠️ <b>ลาย้อนหลัง — ต้องตรวจสอบ</b> ⚠️⚠️\n';
  else if(status.indexOf('กระทันหัน')>=0) alertHead = '⚠️ <b>ลากระทันหัน (แจ้งวันลา)</b>\n';

  var text = alertHead +
    '📋 <b>แจ้งลาใหม่</b>\n' +
    '👤 ' + tgEscape_(name) + ' (' + tgEscape_(emp.emp_id) + ')\n' +
    tgCoLine_(emp, leaveDate) +
    '📅 วันลา: ' + tgEscape_(fmtDMY_(leaveDate)) + '\n' +
    '⏰ เวลา: ' + tgEscape_(compactSlots_(slots)) + '  (' + hours + ' ชม.)\n' +
    '📝 ประเภท: ' + tgEscape_(leaveType) + '\n' +
    statusIcon + ' สถานะ: <b>' + tgEscape_(status) + '</b>' +
    (reason ? ('\n💬 เหตุผล: ' + tgReasonValue_(reason)) : '');

  // ถ้ามีรูปพนักงาน ส่งรูปพร้อมคำบรรยาย; ถ้าไม่มี/ส่งรูปไม่ได้ ส่งข้อความแทน
  if(emp.photo_id){
    var r = tgSendPhoto_(emp.photo_id, text);
    if(!r.ok) tgSend_(text);
  } else {
    tgSend_(text);
  }
}

/* แจ้ง Telegram สรุปการลาหลายวัน (ส่งครั้งเดียว) */
function notifyTelegramRange_(emp, created, skipped, slots, leaveType, reason){
  var token=getSetting_('telegram_bot_token'), chatId=getSetting_('telegram_chat_id');
  if(!token||!chatId) return;
  var name=emp.name_th||emp.name_mm;
  if(emp.name_mm && emp.nationality==='MM') name=(emp.name_th||'')+' / '+emp.name_mm;
  var dates=created.map(function(c){return fmtDMY_(c.date);});
  var anyBack=created.some(function(c){return c.filing_status.indexOf('ย้อนหลัง')>=0;});
  var head = anyBack ? '⚠️ <b>ลาหลายวัน (มีลาย้อนหลัง)</b>\n' : '📋 <b>แจ้งลาหลายวัน</b>\n';
  var text = head +
    '👤 '+tgEscape_(name)+' ('+tgEscape_(emp.emp_id)+')\n'+
    tgCoLine_(emp, created[0].date)+
    '📅 ลา '+created.length+' วัน: '+tgEscape_(dates.join(', '))+'\n'+
    '⏰ เวลา/วัน: '+tgEscape_(compactSlots_(slots))+'  ('+slots.length+' ชม./วัน)\n'+
    '📝 ประเภท: '+tgEscape_(leaveType)+
    (reason?('\n💬 เหตุผล: '+tgReasonValue_(reason)):'')+
    (skipped.length?('\nℹ️ ข้าม '+skipped.length+' วัน (วันหยุด/อาทิตย์/ลาซ้ำ)'):'');
  if(emp.photo_id){ var r=tgSendPhoto_(emp.photo_id,text); if(!r.ok) tgSend_(text); }
  else tgSend_(text);
}

/* ============================================================
 *  แจ้งเตือน "สรุปคนลา" เข้า Telegram อัตโนมัติ (time-based trigger)
 *   - รอบเย็น (summaryEve_)    : สรุปคนลา "พรุ่งนี้"  → เตรียมกำลังคนล่วงหน้า
 *   - รอบเช้า (summaryMorning_): สรุปคนลา "วันนี้"   → roll-call ก่อนเริ่มงาน
 *   - ไม่ส่งในวันที่ไม่ใช่วันทำงาน (อาทิตย์/วันหยุด ที่ไม่ได้ตั้งเป็นวันทำงานพิเศษ)
 * ============================================================ */
var SUMMARY_WD_FULL = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];

function summaryEve_(){
  var t=new Date(); t.setDate(t.getDate()+1);
  var ds=Utilities.formatDate(t,TZ,'yyyy-MM-dd'), chatId=String(getSetting_('telegram_chat_id')||'').trim();
  var note=''; try{ note=calTomorrowNote_(ds); }catch(e){}
  if(note && chatId && getSetting_('telegram_bot_token')) tgSendTo_(chatId, note, tgBottomKb_());
  if(!isWorkingDay_(ds)) return;            // พรุ่งนี้ไม่ทำงาน → ไม่ส่งสรุปคนลา
  sendDailyLeaveSummary_(ds, false);
}
/* trigger handler: สรุป "วันนี้" (รอบเช้า) */
function summaryMorning_(){
  var ds=todayStr_();
  if(!isWorkingDay_(ds)) return;            // วันนี้ไม่ทำงาน → ไม่ส่ง
  sendDailyLeaveSummary_(ds, true);
}

function sendDailyLeaveSummary_(dateStr, isToday){
  var token=getSetting_('telegram_bot_token'), chatId=getSetting_('telegram_chat_id');
  if(!token||!chatId) return {ok:false, error:'ยังไม่ได้ตั้ง Telegram'};
  var r=tgSendTo_(chatId, buildLeaveSummaryText_(dateStr, isToday)+otSummaryBlock_(isToday), tgBottomKb_());
  return (r && r.ok) ? {ok:true} : {ok:false, error:(r && r.description) || 'ส่งไม่สำเร็จ'};
}

/* ป้ายบอกช่วงเวลาที่ลา (เต็มวัน/ครึ่งเช้า/ครึ่งบ่าย/ชม.) จากข้อมูลที่รวมแล้ว */
function summaryDurationLabel_(p){
  if(p.hours>=WORK_HOURS_PER_DAY) return 'เต็มวัน';
  var am=TIME_SLOTS.slice(0,4).slice().sort().join(',');
  var pm=TIME_SLOTS.slice(4,8).slice().sort().join(',');
  var k=Object.keys(p.slots).sort().join(',');
  if(k===am) return 'ครึ่งเช้า';
  if(k===pm) return 'ครึ่งบ่าย';
  return p.hours+' ชม.';
}
function summaryFlag_(nat){ return nat==='MM'?'🇲🇲':'🇹🇭'; }

/* สร้างข้อความสรุปคนลาของวัน dateStr (อ่าน Leaves ครั้งเดียว) — anyDay=true ใช้หัวข้อกลาง ๆ (คำสั่ง /date) */
function buildLeaveSummaryText_(dateStr, isToday, anyDay){
  var wd=SUMMARY_WD_FULL[parseDateStr_(dateStr).getDay()];
  var title = anyDay && !isToday ? '🗓 <b>สรุปคนลา</b>' : (isToday?'☀️ <b>สรุปคนลาวันนี้</b>':'🌙 <b>สรุปเตรียมงานพรุ่งนี้</b>');
  var head=title+'\n📅 '+wd+' '+fmtDMY_(dateStr)+(isWorkingDay_(dateStr)?'':'  (วันหยุด)');
  var mix=companyMix_(), emps=mix.show?empMap_():{};

  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues();
  var TYPES=['ลาป่วย','ลากิจ','ลาพักร้อน'];
  var groups={}; TYPES.forEach(function(t){ groups[t]={}; });
  var pending={};   // ลาพักร้อนที่ยังรออนุมัติ
  for(var i=1;i<vals.length;i++){
    var d=vals[i][5], ds=(d instanceof Date)?Utilities.formatDate(d,TZ,'yyyy-MM-dd'):String(d).trim();
    if(ds!==dateStr) continue;
    var ap=String(vals[i][12]||'');
    if(ap==='ไม่อนุมัติ') continue;                     // ไม่อนุมัติ = ตัดออก
    var empId=String(vals[i][2]).trim(), nm=vals[i][3]||empId, nat=vals[i][4]||'TH';
    var type=String(vals[i][9]||''), hrs=Number(vals[i][7])||0;
    var bucket=(type==='ลาพักร้อน'&&ap==='รออนุมัติ')?pending:(groups[type]||(groups[type]={}));
    if(!bucket[empId]){
      var e=emps[empId], co=String(vals[i][17]||'').trim()?normCompany_(vals[i][17]):companyOn_(e, ds);
      bucket[empId]={name:nm,nat:nat,hours:0,slots:{},co:co,ex:!!(e&&e.ex_alt&&co==='MAIN')};
    }
    bucket[empId].hours+=hrs;
    String(vals[i][6]||'').split(',').forEach(function(s){ s=s.trim(); if(s) bucket[empId].slots[s]=true; });
  }

  // นับคนยืนยันลา (ป่วย/กิจ/พักร้อนที่อนุมัติ) แยกไทย/พม่า
  var confirmed={}, thN=0, mmN=0, coN={ALT:0,MAIN:0};
  TYPES.forEach(function(t){ Object.keys(groups[t]).forEach(function(id){ confirmed[id]=groups[t][id]; }); });
  Object.keys(confirmed).forEach(function(id){ if(confirmed[id].nat==='MM') mmN++; else thN++; coN[confirmed[id].co]++; });
  var total=Object.keys(confirmed).length, pendN=Object.keys(pending).length;

  if(total===0 && pendN===0){
    return head+'\n\n✅ '+(anyDay&&!isToday?'':(isToday?'วันนี้':'พรุ่งนี้'))+'ไม่มีคนลา — ทำงานเต็มกำลัง';
  }
  var line=function(p){ return '\n• '+summaryFlag_(p.nat)+' '+tgEscape_(p.name)+' — '+summaryDurationLabel_(p)+(mix.show?('  <i>['+p.co+(p.ex?' · อดีต ALT':'')+']</i>'):''); };
  var out=head+'\n\n👥 ลาทั้งหมด <b>'+total+'</b> คน'+(total?(' (🇹🇭 '+thN+' · 🇲🇲 '+mmN+')'):'');
  if(mix.show && total) out+='\n🏢 ALT '+coN.ALT+' · MAIN '+coN.MAIN;
  var icons={'ลาป่วย':'🤒','ลากิจ':'👔','ลาพักร้อน':'🏖️'};
  TYPES.forEach(function(t){
    var ids=Object.keys(groups[t]); if(ids.length===0) return;
    out+='\n\n'+icons[t]+' <b>'+t+(t==='ลาพักร้อน'?' (อนุมัติแล้ว)':'')+'</b> ('+ids.length+')';
    ids.map(function(id){ return groups[t][id]; })
       .sort(function(a,b){ return a.name<b.name?-1:1; })
       .forEach(function(p){ out+=line(p); });
  });
  if(pendN>0){
    out+='\n\n⏳ <b>รออนุมัติ–พักร้อน</b> ('+pendN+')  <i>(ยังไม่ยืนยัน)</i>';
    Object.keys(pending).map(function(id){ return pending[id]; })
      .sort(function(a,b){ return a.name<b.name?-1:1; })
      .forEach(function(p){ out+=line(p); });
  }
  return out;
}

/* ---- ตั้ง/ลบ trigger แจ้งเตือนสรุป (รันจากเมนู) ---- */
function pad2_(n){ n=String(n); return n.length<2?('0'+n):n; }
function installSummaryTriggers_(){
  // ลบ trigger เดิมของ 2 ฟังก์ชันนี้ก่อน กันซ้ำซ้อน
  ScriptApp.getProjectTriggers().forEach(function(tr){
    var fn=tr.getHandlerFunction();
    if(fn==='summaryEve_'||fn==='summaryMorning_') ScriptApp.deleteTrigger(tr);
  });
  var eh=parseInt(getSetting_('summary_eve_hour'),10);     if(isNaN(eh)||eh<0||eh>23) eh=17;
  var em=parseInt(getSetting_('summary_eve_min'),10);      if(isNaN(em)||em<0||em>59) em=0;
  var mh=parseInt(getSetting_('summary_morning_hour'),10); if(isNaN(mh)||mh<0||mh>23) mh=7;
  var mm=parseInt(getSetting_('summary_morning_min'),10);  if(isNaN(mm)||mm<0||mm>59) mm=30;
  ScriptApp.newTrigger('summaryEve_').timeBased().everyDays(1).atHour(eh).nearMinute(em).inTimezone(TZ).create();
  ScriptApp.newTrigger('summaryMorning_').timeBased().everyDays(1).atHour(mh).nearMinute(mm).inTimezone(TZ).create();
  return {eh:eh,em:em,mh:mh,mm:mm};
}
function menuInstallSummary_(){
  try{
    var r=installSummaryTriggers_();
    _alert_('✅ ตั้งแจ้งเตือนสรุปลาแล้ว\n\n'+
      '• รอบเย็น (สรุป "พรุ่งนี้") ~ '+pad2_(r.eh)+':'+pad2_(r.em)+' น.\n'+
      '• รอบเช้า (สรุป "วันนี้")  ~ '+pad2_(r.mh)+':'+pad2_(r.mm)+' น.\n\n'+
      'ℹ️ เวลาอาจคลาดเคลื่อน ±15 นาที (ระบบ Google)\n'+
      'ℹ️ ไม่ส่งในวันอาทิตย์/วันหยุด ที่ไม่ได้ตั้งเป็น "วันทำงานพิเศษ"\n\n'+
      'ปรับเวลาได้ในชีท Settings (summary_eve_hour / summary_eve_min / summary_morning_hour / summary_morning_min) แล้วกดเมนูนี้ซ้ำ');
  }catch(e){ _alert_('❌ ตั้งไม่สำเร็จ: '+e+'\n\n(หากเป็นครั้งแรก ให้กดอนุญาตสิทธิ์ แล้วลองใหม่)'); }
}
function menuRemoveSummary_(){
  var n=0;
  ScriptApp.getProjectTriggers().forEach(function(tr){
    var fn=tr.getHandlerFunction();
    if(fn==='summaryEve_'||fn==='summaryMorning_'){ ScriptApp.deleteTrigger(tr); n++; }
  });
  _alert_(n ? ('🔕 ปิดแจ้งเตือนสรุปลาแล้ว ('+n+' ตัว)') : 'ไม่พบแจ้งเตือนสรุปลาที่ตั้งไว้');
}
function menuTestSummaryTomorrow_(){
  var t=new Date(); t.setDate(t.getDate()+1);
  var ds=Utilities.formatDate(t,TZ,'yyyy-MM-dd');
  var work=isWorkingDay_(ds), r=sendDailyLeaveSummary_(ds, false);
  _alert_((work?'':'⚠️ พรุ่งนี้เป็นวันหยุด/ไม่ทำงาน — ของจริงจะไม่ส่ง (นี่เป็นตัวอย่าง)\n\n')+
          (r&&r.ok?'✅ ส่งตัวอย่างสรุป "พรุ่งนี้" เข้ากลุ่มแล้ว':'❌ ส่งไม่สำเร็จ: '+((r&&r.error)||'')));
}
function menuTestSummaryToday_(){
  var ds=todayStr_();
  var work=isWorkingDay_(ds), r=sendDailyLeaveSummary_(ds, true);
  _alert_((work?'':'⚠️ วันนี้เป็นวันหยุด/ไม่ทำงาน — ของจริงจะไม่ส่ง (นี่เป็นตัวอย่าง)\n\n')+
          (r&&r.ok?'✅ ส่งตัวอย่างสรุป "วันนี้" เข้ากลุ่มแล้ว':'❌ ส่งไม่สำเร็จ: '+((r&&r.error)||'')));
}

/* -------- ตัวช่วยตั้งค่า Telegram (รันจากเมนูในชีทได้) -------- */

// 1) หลังสร้างบอทแล้ว ให้เพิ่มบอทเข้ากลุ่ม แล้วพิมพ์อะไรก็ได้ในกลุ่ม 1 ข้อความ
//    จากนั้นรันฟังก์ชันนี้ -> มันจะแสดง chat id ของกลุ่มให้ (ดูใน popup / Logs)
function getTelegramChatId(){
  var token = getSetting_('telegram_bot_token');
  if(!token){
    _alert_('กรุณาใส่ telegram_bot_token ในชีท Settings ก่อน');
    return;
  }
  try{
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/getUpdates', {muteHttpExceptions:true});
    var body = JSON.parse(res.getContentText());
    if(!body.ok){ _alert_('Token ผิดหรือมีปัญหา: ' + (body.description||'')); return; }
    if(!body.result || body.result.length===0){
      _alert_('ยังไม่เจอข้อความ\n\nวิธีทำ:\n1) เพิ่มบอทเข้ากลุ่ม\n2) พิมพ์ข้อความอะไรก็ได้ในกลุ่ม 1 ครั้ง\n3) รันฟังก์ชันนี้อีกที');
      return;
    }
    var lines = [];
    var seen = {};
    body.result.forEach(function(u){
      var chat = (u.message && u.message.chat) || (u.channel_post && u.channel_post.chat);
      if(chat && !seen[chat.id]){
        seen[chat.id] = true;
        lines.push('• ' + (chat.title || chat.first_name || chat.type) + '  →  chat_id = ' + chat.id);
      }
    });
    var msg = 'พบแชตเหล่านี้:\n\n' + lines.join('\n') +
              '\n\nก๊อปปี้ chat_id ของกลุ่มที่ต้องการ ไปวางในชีท Settings แถว telegram_chat_id';
    Logger.log(msg);
    _alert_(msg);
  }catch(e){
    _alert_('ผิดพลาด: ' + e);
  }
}

// 2) ทดสอบส่งข้อความเข้ากลุ่ม (รันจากเมนู หรือเรียกผ่าน API ปุ่มในหน้าแอดมิน)
function testTelegram(){
  var r = tgSend_('✅ <b>ทดสอบ Telegram สำเร็จ</b>\nระบบแจ้งลาเชื่อมต่อกลุ่มนี้เรียบร้อยแล้ว');
  _alert_(r.ok ? '✅ ส่งทดสอบสำเร็จ! ดูข้อความในกลุ่ม Telegram ได้เลย'
              : '❌ ส่งไม่สำเร็จ: ' + r.error);
  return r;
}

function _alert_(msg){
  try { SpreadsheetApp.getUi().alert(msg); } catch(e){ Logger.log(msg); }
}

/* ====== ฟอร์แมตค่า ====== */
function fmtCell_(v){
  if(v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  return String(v||'');
}
function fmtDate_(v){
  if(v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v||'');
}

/* เมนู wrapper ที่ล็อกก่อนเรียก (กันชนกับการเขียนจากเว็บ) */
function menuSyncPayroll_(){ withLock_(syncPayrollRosters_); }
function menuRefreshDaily_(){ withLock_(refreshDailyBoth_); }

function onOpen(){
  var ui=SpreadsheetApp.getUi();
  ui.createMenu(APP.MENU)
    .addItem('▶ ติดตั้ง/อัปเดตโครงสร้าง (setup)','menuSetup_')
    .addItem('🔄 รีเฟรชรายชื่อ Payroll','menuSyncPayroll_')
    .addItem('📅 รีเฟรชหน้าสรุปรายวัน','menuRefreshDaily_')
    .addItem('🧮 คำนวณหักตามโควต้า (ทุกคน)','computeAllDaily_')
    .addItem('📤 ส่งออกไฟล์สรุปการลา + โอที (รอบที่เลือกในชีต สรุป)','menuExportMAIN_')
    .addItem('🗂 จัดโฟลเดอร์ใน Drive (ไฟล์ชีต + รูปพนักงาน)','menuOrganizeDrive_')
    .addSeparator()
    .addSubMenu(ui.createMenu('💬 Telegram')
      .addItem('★ เชื่อม Telegram ทั้งหมดในครั้งเดียว','menuConnectTelegram_')
      .addItem('📌 ส่งแผงปุ่มเมนูเข้ากลุ่ม (ปักหมุด)','menuPostPanel_')
      .addItem('✅ ทดสอบส่งข้อความ','testTelegram')
      .addItem('🔍 เช็คสถานะ Webhook','menuCheckWebhook_')
      .addSeparator()
      .addItem('🔔 เปิด/อัปเดตสรุปรายวัน (เย็น+เช้า)','menuInstallSummary_')
      .addItem('⏸ ปิดสรุปรายวัน','menuRemoveSummary_')
      .addItem('🧪 ทดสอบส่งสรุป "พรุ่งนี้"','menuTestSummaryTomorrow_')
      .addItem('🧪 ทดสอบส่งสรุป "วันนี้"','menuTestSummaryToday_'))
    .addToUi();
}

/* ====== ระบบนี้มีบริษัทเดียว (MAIN) — ทุกคน/ทุกใบลาเป็นของ MAIN ====== */
function normCompany_(c){ return 'MAIN'; }

function isTrue_(v){ return v===true || String(v).toUpperCase()==='TRUE'; }

function empFromRow_(row, rowNum){
  var moved=row[11];
  var movedStr=(moved instanceof Date)?Utilities.formatDate(moved,TZ,'yyyy-MM-dd'):String(moved===undefined||moved===null?'':moved).trim();
  return {
    emp_id:String(row[0]).trim(), name_th:row[1], name_mm:row[2], nationality:(row[3]==='MM'?'MM':'TH'),
    photo_id:row[4], visible:String(row[5]).toUpperCase()!=='FALSE', active:String(row[6]).toUpperCase()!=='FALSE',
    note:row[7], company:normCompany_(row[9]), ex_alt:isTrue_(row[10]), moved_at:movedStr, row:rowNum
  };
}

function companyOn_(emp, dateStr){ return 'MAIN'; }

/* อ่านพนักงานทั้งหมดเป็น map (id -> emp) ครั้งเดียว */
function empMap_(){
  var vals=sheet_(SHEETS.EMP).getDataRange().getValues(), m={};
  for(var i=1;i<vals.length;i++){ var e=empFromRow_(vals[i],i+1); if(e.emp_id) m[e.emp_id]=e; }
  return m;
}

/* กลุ่มพนักงานที่มีอยู่ตอนนี้ — ใช้ตัดสินว่าต้องโชว์ป้ายบริษัทไหม (โชว์เฉพาะช่วงที่ยังปนกันอยู่) */
var _coMix=null;

function companyMix_(){
  if(_coMix) return _coMix;
  var m=empMap_(), k=0;
  Object.keys(m).forEach(function(id){ if(m[id].active) k++; });
  _coMix={alt:0, main:k, ex:0, show:false};
  return _coMix;
}

function coTag_(emp, dateStr){
  var co=companyOn_(emp, dateStr||todayStr_());
  return co+((emp && emp.ex_alt && co==='MAIN')?' · อดีต ALT':'');
}

/* บรรทัดบริษัทในข้อความ Telegram (ว่างถ้าทุกคนอยู่กลุ่มเดียวกันหมด) */
function tgCoLine_(emp, dateStr){
  if(!emp || !companyMix_().show) return '';
  return '🏢 '+tgEscape_(coTag_(emp, dateStr))+'\n';
}

/* ข้อความที่ผู้ใช้พิมพ์เข้ามา ห้ามกลายเป็นสูตรในชีต (ขึ้นต้นด้วย = + - @ → เติม ' นำหน้าให้เป็นข้อความ)
 * กันคนพิมพ์สูตรในช่องเหตุผลเพื่อดึงข้อมูลชีตอื่น (เช่น PIN/Token ใน Settings) ออกมาทางหน้าประวัติ */
function safeText_(v){
  var s=String(v===undefined||v===null?'':v);
  return /^\s*[=+\-@]/.test(s) ? ("'"+s) : s;
}

/* ====== Settings (อ่านชีตครั้งเดียวต่อการทำงาน 1 รอบ — เร็วขึ้นมาก) ====== */
var _setCache=null;

function settingsMap_(){
  if(_setCache) return _setCache;
  var m={}, sh=sheet_(SHEETS.SET);
  if(sh){ var v=sh.getDataRange().getValues(); for(var i=1;i<v.length;i++){ var k=String(v[i][0]).trim(); if(k) m[k]=String(v[i][1]); } }
  _setCache=m; return m;
}

/* คีย์รอบใน _DailyData: ชีตอาจแปลง "2026-10" เป็นวันที่เอง → อ่านกลับให้เป็น yyyy-MM เสมอ */
function cycKeyOf_(v){ return (v instanceof Date) ? Utilities.formatDate(v,TZ,'yyyy-MM') : String(v===undefined||v===null?'':v).trim(); }

function dateKeyOf_(v){ return (v instanceof Date) ? Utilities.formatDate(v,TZ,'yyyy-MM-dd') : String(v===undefined||v===null?'':v).trim(); }

/* รอบเงินเดือนที่วันนั้นตกอยู่ (คีย์ = เดือนที่รอบจบ) ตามวันตัดรอบของบริษัท */
function cycleKeyForDate_(dateStr, company){
  var d=parseDateStr_(dateStr), co=getCutoffDay_(company);
  var e=(d.getDate()<=co) ? new Date(d.getFullYear(), d.getMonth(), 1) : new Date(d.getFullYear(), d.getMonth()+1, 1);
  return Utilities.formatDate(e,TZ,'yyyy-MM');
}

/* _DailyData: ลบแถวซ้ำ (พนักงาน+วันลาเดียวกัน) เก็บแถวล่าสุด — ของเก่าซ้ำได้เพราะคีย์รอบถูกแปลงเป็นวันที่ */
function dedupeDailyData_(){
  var dd=sheet_(SHEETS.DDATA); if(!dd) return 0;
  var last=dd.getLastRow(); if(last<3) return 0;
  var vals=dd.getRange(2,1,last-1,13).getValues(), seen={}, keep=[];
  for(var i=vals.length-1;i>=0;i--){
    var id=String(vals[i][0]).trim(); if(!id) continue;
    var k=id+'|'+dateKeyOf_(vals[i][4]);
    if(seen[k]) continue;
    seen[k]=true;
    vals[i][3]=cycKeyOf_(vals[i][3]); vals[i][4]=dateKeyOf_(vals[i][4]);
    keep.unshift(vals[i]);
  }
  var removed=vals.length-keep.length;
  if(removed>0){
    dd.getRange(2,1,last-1,13).clearContent();
    if(keep.length) dd.getRange(2,1,keep.length,13).setValues(keep);
  }
  return removed;
}

/* เติมบริษัทให้ใบลาเก่าที่ยังว่าง (ถือเป็น ALT) — เขียนทีเดียวทั้งคอลัมน์ */
function backfillLeaveCompany_(){
  var lv=sheet_(SHEETS.LEAVE), last=lv.getLastRow();
  if(last<2) return 0;
  var ids=lv.getRange(2,2,last-1,1).getValues(), col=lv.getRange(2,18,last-1,1).getValues(), n=0;
  for(var i=0;i<col.length;i++){
    if(String(ids[i][0]).trim() && !String(col[i][0]).trim()){ col[i][0]='MAIN'; n++; }
  }
  if(n) lv.getRange(2,18,last-1,1).setValues(col);
  return n;
}

var BACKEND_VER = '1.0';

/* ผลลัพธ์ error แบบมีรหัส (ให้หน้าเว็บแปลเป็นภาษาพม่าได้) */
function err_(code, msg, extra){
  var o={ok:false, code:code, error:msg};
  if(extra) Object.keys(extra).forEach(function(k){ o[k]=extra[k]; });
  return o;
}

/* ====== สลิปใบลา: รหัสตรวจสอบ + หน้าตรวจ ====== */
function slipSecret_(){
  var s=getSetting_('slip_secret');
  if(!s){ s=Utilities.getUuid().replace(/-/g,''); setSetting_('slip_secret', s); }
  return s;
}

function slipCode_(id){
  var raw=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, slipSecret_()+'|'+String(id), Utilities.Charset.UTF_8);
  var hex='';
  for(var i=0;i<4;i++){ var b=raw[i]; if(b<0) b+=256; hex+=(b<16?'0':'')+b.toString(16); }
  return hex.toUpperCase();
}

/* ตรวจสลิป: id = เลขใบลา (L…) หรือรหัสกลุ่ม (G…), c = รหัสตรวจสอบบนสลิป */
function verifySlipCore_(p){
  var id=String(p.id||'').trim(), code=String(p.c||'').trim().toUpperCase();
  if(!id) return err_('MISSING','ไม่พบเลขสลิป');
  if(!code || code!==slipCode_(id)) return err_('BAD_CODE','รหัสตรวจสอบไม่ถูกต้อง — สลิปนี้ไม่ได้ออกจากระบบ');
  var items=[], who=null;
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][1]).trim()!==id && String(vals[i][15]).trim()!==id) continue;
    if(!who) who={emp_id:String(vals[i][2]).trim(), name:vals[i][3], company:normCompany_(vals[i][17])};
    items.push({state:'active', leave_id:String(vals[i][1]), leave_date:fmtDate_(vals[i][5]), slots:vals[i][6], hours:vals[i][7],
      leave_type:vals[i][9], filing_status:vals[i][10], approval_status:String(vals[i][12]||''), filed_at:fmtCell_(vals[i][0])});
  }
  var cl=sheet_(SHEETS.CANCEL);
  if(cl){
    var cv=cl.getDataRange().getValues();
    for(var j=1;j<cv.length;j++){
      if(String(cv[j][13]||'').trim()!==id && String(cv[j][14]||'').trim()!==id) continue;
      if(!who) who={emp_id:String(cv[j][3]).trim(), name:cv[j][4], company:normCompany_(cv[j][12])};
      items.push({state:'cancelled', leave_id:String(cv[j][13]||''), leave_date:fmtDate_(cv[j][6]), slots:cv[j][7], hours:cv[j][9],
        leave_type:cv[j][8], filing_status:cv[j][10], approval_status:'', cancelled_at:fmtDate_(cv[j][0])+' '+String(cv[j][1]||'').slice(0,5)});
    }
  }
  items.sort(function(a,b){ return a.leave_date<b.leave_date?-1:1; });
  if(items.length===0) return {ok:true, found:false, checked_at:nowStr_()};
  return {ok:true, found:true, emp_id:who.emp_id, name:who.name, company:who.company, items:items, checked_at:nowStr_()};
}

/* ช่วงวันที่ใช้กรองในหน้าแอดมิน: ตามเดือนปฏิทิน หรือ ตามรอบเงินเดือน (cycle=1) */
function adminRange_(p){
  var month = String((p&&p.month)||'').trim();
  if(!/^\d{4}-\d{2}$/.test(month)) month = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  var y=parseInt(month.slice(0,4),10), m=parseInt(month.slice(5,7),10)-1;
  if(p && p.cycle==='1'){
    var cyc=cycleForMonth_(new Date(y,m,1), normCompany_(p.company));
    return {month:month, from:cyc.startStr, to:cyc.endStr, label:cyc.label, cycle:true};
  }
  var last=new Date(y,m+1,0);
  return {month:month, from:month+'-01', to:dateToStr_(last), label:thaiMonthAbbr_(m)+' '+y, cycle:false};
}

/* กรองบริษัท: '' = ทั้งหมด, ALT, MAIN, EX = เฉพาะอดีต ALT */
function coFilterMatch_(filter, company, exAlt){
  if(!filter || filter==='ALL') return true;
  if(filter==='EX') return !!exAlt;
  return company===filter;
}

/* ====== แอดมิน: จัดการพนักงาน ====== */
function validEmpId_(id){ return /^[A-Za-z0-9][A-Za-z0-9\-_.]{0,19}$/.test(id); }

function empIdExists_(id){
  var vals=sheet_(SHEETS.EMP).getDataRange().getValues(), low=String(id).toLowerCase();
  for(var i=1;i<vals.length;i++) if(String(vals[i][0]).trim().toLowerCase()===low) return true;
  return false;
}

function empHasLeaves_(empId){
  var lv=sheet_(SHEETS.LEAVE), last=lv.getLastRow();
  if(last<2) return false;
  var ids=lv.getRange(2,3,last-1,1).getValues();
  for(var i=0;i<ids.length;i++) if(String(ids[i][0]).trim()===empId) return true;
  return false;
}

/* ปรับคอลัมน์บริษัทในใบลาของหลายคนทีเดียว (อ่าน 1 ครั้ง เขียน 1 ครั้ง)
 * plan = { empId: {effective:'yyyy-MM-dd'|null, company:'MAIN'|'ALT'} } — มีวันที่: ก่อนวันนั้น = ALT, ตั้งแต่วันนั้น = company */
function setLeaveCompanyBatch_(plan){
  var lv=sheet_(SHEETS.LEAVE), last=lv.getLastRow(), counts={};
  if(last<2) return counts;
  var data=lv.getRange(2,1,last-1,18).getValues(), col=[], changed=false;
  for(var i=0;i<data.length;i++){
    var cur=data[i][17], id=String(data[i][2]).trim(), p=plan[id];
    if(p){
      var ds=dateKeyOf_(data[i][5]);
      var want = p.effective ? (ds<p.effective?'ALT':p.company) : p.company;
      if(String(cur).trim()!==want){ cur=want; changed=true; counts[id]=(counts[id]||0)+1; }
    }
    col.push([cur]);
  }
  if(changed) lv.getRange(2,18,col.length,1).setValues(col);
  return counts;
}

function setLeaveCompanyForEmp_(empId, effective, company){
  var plan={}; plan[empId]={effective:effective, company:company};
  return setLeaveCompanyBatch_(plan)[empId]||0;
}

/* นำเข้าพนักงานหลายคนทีเดียว (วางรายชื่อจาก Excel ในหน้าแอดมิน) */
function importEmployees_(d){
  var rows = d.rows || [];
  if(!rows.length) return err_('MISSING','ไม่มีรายชื่อให้นำเข้า');
  if(rows.length>300) return err_('TOO_MANY','นำเข้าได้ครั้งละไม่เกิน 300 คน');
  var company = normCompany_(d.company);
  var sh = sheet_(SHEETS.EMP);
  var vals = sh.getDataRange().getValues(), have={};
  for(var i=1;i<vals.length;i++){ var k=String(vals[i][0]).trim().toLowerCase(); if(k) have[k]=true; }
  var add=[], skipped=[], at=nowStr_();
  // ตัวนับรหัสอัตโนมัติ (กรณีไม่ได้ส่งรหัสมา)
  var seq={};
  function nextId(nat){
    var main=(company==='MAIN'), mm=(nat==='MM');
    var prefix = mm ? 'KM-' : 'K', pad = mm?2:3;
    if(seq[prefix]===undefined){
      var re=new RegExp('^'+prefix.replace('-','\\-')+'(\\d+)$','i'), max=0;
      Object.keys(have).forEach(function(k){ var m=re.exec(k); if(m){ var n=parseInt(m[1],10); if(n>max) max=n; } });
      seq[prefix]=max;
    }
    var id;
    do{ seq[prefix]++; var num=String(seq[prefix]); while(num.length<pad) num='0'+num; id=prefix+num; } while(have[id.toLowerCase()]);
    return id;
  }
  rows.forEach(function(r, idx){
    var nameTh=String(r.name_th||'').trim(), nameMm=String(r.name_mm||'').trim();
    if(!nameTh && !nameMm){ skipped.push({line:idx+1, emp_id:String(r.emp_id||''), reason:'ไม่มีชื่อ'}); return; }
    var nat=(String(r.nationality||'').toUpperCase()==='MM')?'MM':'TH';
    var id=String(r.emp_id||'').trim();
    if(id){
      if(!validEmpId_(id)){ skipped.push({line:idx+1, emp_id:id, reason:'รหัสไม่ถูกรูปแบบ'}); return; }
      if(have[id.toLowerCase()]){ skipped.push({line:idx+1, emp_id:id, reason:'รหัสซ้ำ'}); return; }
    } else id=nextId(nat);
    have[id.toLowerCase()]=true;
    add.push([id, safeText_(nameTh), safeText_(nameMm), nat, '', true, true, safeText_(String(r.note||'').trim()), at, company, false, '']);
  });
  if(add.length){
    var r0=sh.getLastRow()+1;
    sh.getRange(r0,1,add.length,1).setNumberFormat('@');     // รหัสเป็นข้อความเสมอ (0012 / 10-12 ต้องไม่ถูกแปลง)
    sh.getRange(r0,1,add.length,12).setValues(add);
    _coMix=null;
    syncPayrollRosters_();
  }
  return {ok:true, added:add.length, added_ids:add.map(function(a){return a[0];}), skipped:skipped};
}

function transferEmployees_(d){ return err_('NA','ระบบนี้มีบริษัทเดียว ไม่มีการย้ายบริษัท'); }

/* ค่าตั้งค่าที่หน้าแอดมินอ่าน/แก้ได้ (ไม่ส่ง token กลับไป — บอกแค่ว่าตั้งแล้วหรือยัง) */
var SETTING_KEYS_ADMIN = ['telegram_bot_token','telegram_chat_id','cutoff_day','admin_pin','ot_back_days','ot_need_confirm',
  'summary_eve_hour','summary_eve_min','summary_morning_hour','summary_morning_min'];

function getAdminSettingsCore_(p){
  if(!(p && checkPin_(p.pin))) return err_('PIN','PIN ไม่ถูกต้อง');
  return {ok:true, ver:BACKEND_VER,
    telegram_token_set: !!getSetting_('telegram_bot_token'),
    pin_is_default: checkPin_(DEFAULT_PIN),
    telegram_chat_id: getSetting_('telegram_chat_id'),
    cutoff_day: getCutoffDay_('ALT'),
    cutoff_day_main: getSetting_('cutoff_day_main'),
    summary: {eh:getSetting_('summary_eve_hour'), em:getSetting_('summary_eve_min'), mh:getSetting_('summary_morning_hour'), mm:getSetting_('summary_morning_min')},
    mix: companyMix_()};
}

/* ย่อช่วงเวลาที่ติดกันในข้อความ Telegram: 08:00-09:00, 09:00-10:00 → 08:00–10:00 (ครบวัน = "เต็มวัน 08:00–17:00") */
function compactSlots_(v){
  var arr=(v instanceof Array ? v : String(v||'').split(',')).map(function(s){ return String(s).trim(); })
    .filter(function(s){ return TIME_SLOTS.indexOf(s)>=0; })
    .sort(function(a,b){ return TIME_SLOTS.indexOf(a)-TIME_SLOTS.indexOf(b); });
  if(!arr.length) return String(v instanceof Array ? v.join(', ') : (v||''));
  if(arr.length>=WORK_HOURS_PER_DAY) return 'เต็มวัน 08:00–17:00';
  var out=[], st=null, en=null;
  arr.forEach(function(s){ var p=s.split('-'); if(en===p[0]) en=p[1]; else { if(st) out.push(st+'–'+en); st=p[0]; en=p[1]; } });
  if(st) out.push(st+'–'+en);
  return out.join(', ');
}

/* ====== Telegram ====== */
function tgApi_(method, payload){
  var token=getSetting_('telegram_bot_token');
  if(!token) return {ok:false, description:'no token'};
  try{
    var res=UrlFetchApp.fetch('https://api.telegram.org/bot'+token+'/'+method,{
      method:'post', contentType:'application/json', payload:JSON.stringify(payload||{}), muteHttpExceptions:true });
    return JSON.parse(res.getContentText());
  }catch(e){ return {ok:false, description:String(e)}; }
}

function tgSendTo_(chatId, text, keyboard){
  var p={chat_id:chatId, text:tgClip_(text), parse_mode:'HTML', disable_web_page_preview:true};
  if(keyboard) p.reply_markup=keyboard;
  return tgApi_('sendMessage', p);
}

/* ตรวจว่าคำขอมาจาก Telegram จริง: ต้องมีรหัสลับ (?tgs=…) ตรงกับที่ตั้งไว้ตอนผูก webhook
 * ยังไม่เคยตั้งรหัสลับ = ไม่รับคำสั่งจาก Telegram เลย (กันคนอื่นยิงคำขอปลอมมากดอนุมัติลาพักร้อน) */
function tgSecretOk_(e){
  var sec=String(getSetting_('tg_webhook_secret')||'').trim();
  if(!sec) return false;
  var got=String((e && e.parameter && e.parameter.tgs) || '');
  return got===sec;
}

/* Telegram ส่งอัปเดตเดิมซ้ำได้ (ถ้าคิดว่าเราไม่ตอบ) → จำ update_id ไว้ 10 นาที ไม่ทำซ้ำ */
function tgSeenUpdate_(updateId){
  if(updateId===undefined || updateId===null) return false;
  try{
    var cache=CacheService.getScriptCache(), k='tgu_'+updateId;
    if(cache.get(k)) return true;
    cache.put(k,'1',600);
  }catch(e){}
  return false;
}

/* ใครสั่งบอทได้: ข้อความในกลุ่มที่ตั้งไว้ หรือแชตส่วนตัวของ "สมาชิกกลุ่ม" เท่านั้น */
function tgAllowed_(chat, from){
  var gid=String(getSetting_('telegram_chat_id')||'').trim();
  if(!chat || !gid) return false;
  if(String(chat.id)===gid) return true;
  if(chat.type==='private' && from && String(chat.id)===String(from.id)){
    var cache=CacheService.getScriptCache(), k='tgm_'+from.id, c=cache.get(k);
    if(c!==null && c!==undefined) return c==='1';
    var r=tgApi_('getChatMember',{chat_id:gid, user_id:from.id});
    var ok=!!(r && r.ok && r.result && ['creator','administrator','member'].indexOf(r.result.status)>=0);
    cache.put(k, ok?'1':'0', ok?3600:300);
    return ok;
  }
  return false;
}

function vacationRequestText_(emp, dateList, reason){
  var name=emp.name_th||emp.name_mm;
  if(emp.name_mm && emp.nationality==='MM') name=(emp.name_th||'')+' / '+emp.name_mm;
  var dates=dateList.map(function(c){return fmtDMY_(c);});
  return '🌴 <b>คำขอลาพักร้อน (รออนุมัติ)</b>\n'+
    '👤 '+tgEscape_(name)+' ('+tgEscape_(emp.emp_id)+')\n'+
    tgCoLine_(emp, dateList[0])+
    '📅 วันที่ลา ('+dateList.length+' วัน): '+tgEscape_(dates.join(', '))+'\n'+
    (reason?('💬 เหตุผล: '+tgReasonValue_(reason)+'\n'):'')+
    '\nกดปุ่มด้านล่างเพื่ออนุมัติ/ไม่อนุมัติ';
}

function tgMenuKb_(mode){
  mode=(mode==='q')?'q':'p';
  return {inline_keyboard:[
    [{text:TG_BTN.today, callback_data:mode+':today'},{text:TG_BTN.tomorrow, callback_data:mode+':tomorrow'}],
    [{text:TG_BTN.week, callback_data:mode+':week'},{text:TG_BTN.pending, callback_data:mode+':pending'}],
    [{text:'➕ เพิ่มเติม', callback_data:'m:more:'+mode}]
  ]};
}

function tgHelpText_(chatId){
  if(tgIsPrivate_(chatId))
    return '📋 <b>ระบบลาหยุด '+APP.NAME+'</b>\nกดปุ่มด้านล่างจอได้เลย 👇\n\nหรือพิมพ์มาได้เลย:\n• ชื่อหรือรหัสพนักงาน → ดูวันลาของคนนั้น\n• วันที่ เช่น <code>15/10</code> → ดูว่าวันนั้นใครลา';
  return '📋 <b>ระบบลาหยุด '+APP.NAME+'</b>\nกดปุ่มด้านล่างจอได้เลย 👇\n<i>ถ้าไม่เห็นปุ่ม: กดไอคอนแป้นสี่เหลี่ยม ⌨️ ข้างช่องพิมพ์</i>';
}

/* ข้อความที่พิมพ์เข้ามา: คำสั่ง / คำตอบของคำถามที่บอทถาม (ค้นหา, ดูวันอื่น, ปฏิทิน) / แชตส่วนตัว */
function handleTgMessage_(msg){
  var text=String(msg.text||'').trim(), chat=msg.chat;
  if(!text || !chat) return json_({ok:true});
  var priv=(chat.type==='private'), isCmd=(text.charAt(0)==='/'), ask='', cal='', rtx='';
  var rt=msg.reply_to_message;
  if(!isCmd && rt && rt.from && rt.from.is_bot){            // คำตอบของคำถามที่บอทถาม
    rtx=String(rt.text||'');
    if(rtx.indexOf(TG_ASK.find)===0) ask='who'; else if(rtx.indexOf(TG_ASK.pick)===0) ask='date';
    else Object.keys(CAL_ASK).forEach(function(k){ if(rtx.indexOf(CAL_ASK[k])===0) cal=k; });
  }
  var key='';                                                       // ปุ่มล่างจอ (ข้อความตรงกับป้ายปุ่มเป๊ะ)
  var nk=function(s){ return String(s).replace(/[\uFE0E\uFE0F]/g,'').replace(/\s+/g,' ').trim(); };
  if(!isCmd){ var nt=nk(text); Object.keys(TG_BTN).forEach(function(k){ if(nk(TG_BTN[k])===nt) key=k; }); }
  if(key){ ask=''; cal=''; }
  if(!isCmd && !ask && !cal && !key && !priv) return json_({ok:true});      // ในกลุ่ม: ตอบเฉพาะคำสั่ง / ปุ่ม / คำตอบที่บอทถาม
  if(!tgAllowed_(chat, msg.from)){
    if(priv && msg.from && String(chat.id)===String(msg.from.id)) tgSendTo_(chat.id,'⛔ บอทนี้ใช้ได้เฉพาะสมาชิกกลุ่มลาหยุดของบริษัท');
    return json_({ok:true});
  }
  text=text.slice(0,120);
  var opt={from:msg.from};
  if(isCmd){
    var parts=text.split(/\s+/), cmd=parts[0].split('@')[0].toLowerCase().replace('/',''), arg=parts.slice(1).join(' ').trim();
    if(cmd==='start'||cmd==='menu'||cmd==='help') tgSendTo_(chat.id, tgHelpText_(chat.id), tgKb_(chat.id));
    else tgRunQuery_(cmd, arg, chat.id, opt);
    return json_({ok:true});
  }
  if(cal){ tgCalTyped_(cal, text, chat.id, rtx); return json_({ok:true}); }
  if(ask){ tgRunQuery_(ask, text, chat.id, opt); return json_({ok:true}); }
  if(key){ tgRunQuery_(key, '', chat.id, opt); return json_({ok:true}); }
  // แชตส่วนตัว: พิมพ์วันที่ / ชื่อหรือรหัสพนักงาน ได้เลย
  if(tgParseDate_(text)) tgRunQuery_('date', text, chat.id, opt);
  else tgRunQuery_('who', text, chat.id, opt);
  return json_({ok:true});
}

/* opt.editMsgId = ข้อความแผงปุ่มที่ให้เปลี่ยนเนื้อหา (ไม่มี = ส่งข้อความใหม่) · opt.from = คนที่กด/พิมพ์ */
function tgRunQuery_(cmd, arg, chatId, opt){
  opt=opt||{};
  var ed=opt.editMsgId||null, t;
  if(cmd==='today'){ tgShow_(chatId, buildLeaveSummaryText_(todayStr_(), true), ed); return; }
  if(cmd==='tomorrow'){ t=new Date(); t.setDate(t.getDate()+1); tgShow_(chatId, buildLeaveSummaryText_(Utilities.formatDate(t,TZ,'yyyy-MM-dd'), false), ed); return; }
  if(cmd==='week'){ tgShow_(chatId, tgWeekText_(), ed); return; }
  if(cmd==='pending'){ tgSendPending_(chatId, ed); return; }
  if(cmd==='month'){ tgShow_(chatId, tgMonthText_(), ed); return; }
  if(cmd==='ot'){ tgShow_(chatId, tgOtText_(), ed); return; }
  if(cmd==='cal'){ if(ed) tgShow_(chatId, tgCalText_(), ed, tgCalKb_()); else tgSendTo_(chatId, tgCalText_(), tgCalKb_()); return; }
  if(cmd==='more'){ tgSendTo_(chatId,'➕ <b>เพิ่มเติม</b> — เลือกรายการ 👇', tgMoreKb_('p')); return; }
  if(cmd==='date'||cmd==='pick'){
    if(!String(arg||'').trim()){ tgAsk_(chatId,'pick',opt.from); return; }
    var ds=tgParseDate_(arg);
    if(!ds){ tgSendTo_(chatId,'❓ อ่านวันที่ไม่ออก — พิมพ์แบบ วัน/เดือน เช่น <code>15/10</code>', tgKb_(chatId)); return; }
    tgShow_(chatId, buildLeaveSummaryText_(ds, ds===todayStr_(), true), ed); return;
  }
  if(cmd==='who'||cmd==='find'||cmd==='whohelp'){
    if(!String(arg||'').trim()){ var eh=tgEmpHome_(); tgSendTo_(chatId, eh.text, eh.kb); return; }
    tgWhoReply_(chatId, arg, null); return;
  }
  tgSendTo_(chatId, tgHelpText_(chatId), tgKb_(chatId));
}

function tgParseDate_(s){
  var m=/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/.exec(String(s||'').trim());
  if(!m) return '';
  var y=m[3]?parseInt(m[3],10):parseInt(todayStr_().slice(0,4),10);
  if(y<100) y += (y>=43 ? 2500 : 2000);     // 26 → 2026, 69 → 2569 (พ.ศ.)
  if(y>2400) y-=543;                         // รับปี พ.ศ. ด้วย
  var mo=parseInt(m[2],10), d=parseInt(m[1],10);
  if(mo<1||mo>12||d<1||d>31) return '';
  return y+'-'+pad2_(mo)+'-'+pad2_(d);
}

function tgWeekText_(){
  var start=parseDateStr_(todayStr_()), days=[], idx={};
  for(var k=0;k<7;k++){ var d=new Date(start); d.setDate(d.getDate()+k); var ds=dateToStr_(d); idx[ds]=days.length; days.push({ds:ds, d:d, list:[], pend:[]}); }
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), seen={};
  for(var i=1;i<vals.length;i++){
    var dv=vals[i][5], s=(dv instanceof Date)?Utilities.formatDate(dv,TZ,'yyyy-MM-dd'):String(dv).trim();
    if(idx[s]===undefined) continue;
    var ap=String(vals[i][12]||''); if(ap==='ไม่อนุมัติ') continue;
    var key=s+'|'+vals[i][2]+'|'+ap; if(seen[key]) continue; seen[key]=true;
    var label=summaryFlag_(vals[i][4])+' '+tgEscape_(vals[i][3])+' ('+tgEscape_(String(vals[i][9]).replace('ลา',''))+')';
    if(ap==='รออนุมัติ') days[idx[s]].pend.push(label+' ⏳'); else days[idx[s]].list.push(label);
  }
  var out='🗓 <b>คนลา 7 วันข้างหน้า</b>';
  days.forEach(function(x){
    var wd=SUMMARY_WD_FULL[x.d.getDay()], all=x.list.concat(x.pend);
    var off=!isWorkingDay_(x.ds);
    out+='\n\n<b>'+wd+' '+fmtDMY_(x.ds).slice(0,5)+'</b>'+(off?' (วันหยุด)':'')+' — '+(all.length?(all.length+' คน\n'+all.map(function(n){return '• '+n;}).join('\n')):'ไม่มีคนลา');
  });
  return out;
}

/* ปุ่ม "รออนุมัติ" = ลาพักร้อนรออนุมัติ + โอทีรอยืนยัน */
function tgSendPending_(chatId, editMsgId){
  var list=tgPendingGroups_(), ots=otAll_().filter(function(o){ return o.status===OT_ST.WAIT; });
  if(!list.length && !ots.length){ tgShow_(chatId,'✅ ไม่มีรายการรออนุมัติ\n(ลาพักร้อน 0 • โอที 0)', editMsgId); return; }
  tgShow_(chatId,'⏳ <b>รออนุมัติ</b>\n🏖️ ลาพักร้อน '+list.length+' คำขอ • ⏱ โอที '+ots.length+' รายการ\nกดปุ่มได้ที่ข้อความท้ายแชต 👇'+((list.length>8||ots.length>10)?'\n<i>(แสดงบางส่วน กดปุ่มนี้อีกครั้งหลังจัดการชุดแรกเสร็จ)</i>':''), editMsgId);
  list.slice(0,8).forEach(function(g){            // Telegram จำกัด ~20 ข้อความ/นาที ต่อกลุ่ม → ส่งรวมไม่เกิน 19
    var emp=findEmployee_(g.emp); if(!emp) return;
    g.dates.sort();
    tgSendButtons_(vacationRequestText_(emp, g.dates, g.reason), g.gid, chatId);
  });
  var emps=empMap_();
  ots.sort(function(a,b){ return a.ot_date<b.ot_date?-1:1; }).slice(0,10).forEach(function(o){
    var e=emps[o.emp_id]||{emp_id:o.emp_id, name_th:o.name, nationality:o.nationality};
    tgSendTo_(chatId, otCardText_(e, o, o.note), otButtons_(o.ot_id));
  });
}

function tgMonthTextCore_(){
  var now=new Date(), co=getCutoffDay_('ALT');
  var dayNow=parseInt(Utilities.formatDate(now,TZ,'d'),10), y=parseInt(Utilities.formatDate(now,TZ,'yyyy'),10), m=parseInt(Utilities.formatDate(now,TZ,'M'),10)-1;
  var cyc=cycleForMonth_((dayNow<=co)?new Date(y,m,1):new Date(y,m+1,1), 'ALT');
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), by={}, T={'ลาป่วย':0,'ลากิจ':0,'ลาพักร้อน':0}, total=0, pend=0, coH={ALT:0,MAIN:0};
  for(var i=1;i<vals.length;i++){
    var dv=vals[i][5]; if(!(dv instanceof Date)) continue;
    var ds=Utilities.formatDate(dv,TZ,'yyyy-MM-dd'); if(ds<cyc.startStr||ds>cyc.endStr) continue;
    var ap=String(vals[i][12]||''); if(ap==='ไม่อนุมัติ') continue; if(ap==='รออนุมัติ'){ pend++; continue; }
    var h=Number(vals[i][7])||0, id=String(vals[i][2]).trim(), type=String(vals[i][9]);
    if(!by[id]) by[id]={name:vals[i][3], nat:vals[i][4], h:0};
    by[id].h+=h; total+=h; if(T[type]!==undefined) T[type]+=h; coH[normCompany_(vals[i][17])]+=h;
  }
  var people=Object.keys(by).map(function(k){return by[k];}).sort(function(a,b){return b.h-a.h;});
  var dfmt=function(h){ return (Math.round(h/WORK_HOURS_PER_DAY*100)/100)+' วัน'; };
  var out='📊 <b>สรุปการลา '+cyc.label+'</b>\n\n👥 คนที่ลา <b>'+people.length+'</b> คน • รวม <b>'+dfmt(total)+'</b> ('+total+' ชม.)'+
    '\n🤒 ลาป่วย '+dfmt(T['ลาป่วย'])+'\n👔 ลากิจ '+dfmt(T['ลากิจ'])+'\n🏖️ ลาพักร้อน '+dfmt(T['ลาพักร้อน'])+
    (pend?('\n⏳ รออนุมัติ '+pend+' วัน'):'');
  if(companyMix_().show) out+='\n🏢 ALT '+dfmt(coH.ALT)+' • MAIN '+dfmt(coH.MAIN);
  if(getCutoffDay_('MAIN')!==co) out+='\n<i>(ช่วงวันตามรอบของ ALT — MAIN ตัดรอบวันที่ '+getCutoffDay_('MAIN')+')</i>';
  if(people.length){
    out+='\n\n<b>ลามากสุด</b>';
    people.slice(0,10).forEach(function(p,k){ out+='\n'+(k+1)+'. '+summaryFlag_(p.nat)+' '+tgEscape_(p.name)+' — '+dfmt(p.h); });
  }
  return out;
}

/* เดิม: คืนข้อความล้วน (คงไว้ให้ของเก่าเรียกได้) */
function tgWhoText_(arg){
  var hits=tgFindEmps_(arg);
  if(!String(arg||'').trim()) return '🔎 พิมพ์ชื่อหรือรหัสพนักงาน';
  if(!hits.length) return '🔎 ไม่พบพนักงานที่ตรงกับ "'+tgEscape_(arg)+'"';
  if(hits.length>3) return '🔎 พบ '+hits.length+' คน พิมพ์ให้เจาะจงขึ้น:\n'+hits.slice(0,15).map(function(e){ return '• '+tgEscape_(e.emp_id)+' '+tgEscape_(e.name_th||e.name_mm); }).join('\n');
  return hits.map(function(e){ return tgEmpCard_(e); }).join('\n\n———\n\n');
}

/* ====== ส่งออกไฟล์สรุปการลา + โอที ของรอบที่เลือก (ไฟล์ Google Sheet แยก ไม่มีข้อมูลค่าแรง — ส่งให้ฝ่ายบัญชีได้) ====== */
function exportCompanyReport_(company){
  company=normCompany_(company);
  var ps=sheet_(SHEETS.PAY_TH), mv=ps?ps.getRange('B1').getValue():null;
  if(!(mv instanceof Date)){ mv=new Date(); mv.setDate(1); }
  var cyc=cycleForMonth_(mv, company), emps=empMap_();
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), per={}, detail=[];
  for(var i=1;i<vals.length;i++){
    var d=vals[i][5]; if(!(d instanceof Date)) continue;
    var ds=Utilities.formatDate(d,TZ,'yyyy-MM-dd'); if(ds<cyc.startStr||ds>cyc.endStr) continue;
    var id=String(vals[i][2]).trim(), e=emps[id];
    var co=String(vals[i][17]||'').trim()?normCompany_(vals[i][17]):companyOn_(e, ds);
    if(co!==company) continue;
    var ap=String(vals[i][12]||''); if(ap==='ไม่อนุมัติ') continue;
    var type=String(vals[i][9]), h=Number(vals[i][7])||0;
    if(!per[id]) per[id]={id:id, name:vals[i][3], nat:vals[i][4], ex:!!(e&&e.ex_alt), sick:0, pers:0, vaca:0, pend:0};
    if(ap==='รออนุมัติ') per[id].pend+=h;
    else if(type==='ลาป่วย') per[id].sick+=h; else if(type==='ลากิจ') per[id].pers+=h; else if(type==='ลาพักร้อน') per[id].vaca+=h;
    detail.push([parseDateStr_(ds), id, vals[i][3], vals[i][4], type, vals[i][6], h, vals[i][8], vals[i][10], ap, safeText_(vals[i][11]), fmtCell_(vals[i][0])]);
  }
  // ใส่ทุกคนของบริษัทนี้ (แม้ไม่ลา) ให้ครบรายชื่อ
  Object.keys(emps).forEach(function(id){ var e=emps[id]; if(e.active && e.company===company && !per[id]) per[id]={id:id,name:e.name_th||e.name_mm,nat:e.nationality,ex:e.ex_alt,sick:0,pers:0,vaca:0,pend:0}; });
  var H=WORK_HOURS_PER_DAY;
  var sum=Object.keys(per).map(function(k){ return per[k]; }).sort(function(a,b){ return a.nat===b.nat?(a.id<b.id?-1:1):(a.nat==='TH'?-1:1); })
    .map(function(p){ var t=p.sick+p.pers+p.vaca; return [p.id,p.name,(p.nat==='MM'?'พม่า':'ไทย'),(p.ex?'อดีต ALT':''),p.sick,p.pers,p.vaca,t,Math.round(t/H*100)/100,p.pend]; });
  detail.sort(function(a,b){ return a[0]-b[0]; });

  var nss=SpreadsheetApp.create('สรุปการลา '+coLabel_(company)+' '+cyc.label+' (ส่งออก '+fmtDMY_(todayStr_())+')');
  var s1=nss.getSheets()[0]; s1.setName('สรุปรายคน');
  s1.getRange(1,1).setValue('สรุปการลา '+coLabel_(company)+' — '+cyc.label+' ('+fmtDMY_(cyc.startStr)+' – '+fmtDMY_(cyc.endStr)+')').setFontWeight('bold').setFontSize(13);
  var h1=['รหัสพนักงาน','ชื่อ','สัญชาติ','หมายเหตุ','ลาป่วย (ชม.)','ลากิจ (ชม.)','ลาพักร้อน (ชม.)','รวมลา (ชม.)','รวมลา (วัน)','พักร้อนรออนุมัติ (ชม.)'];
  s1.getRange(3,1,1,h1.length).setValues([h1]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff').setWrap(true);
  if(sum.length){ s1.getRange(4,1,sum.length,1).setNumberFormat('@'); s1.getRange(4,1,sum.length,h1.length).setValues(sum); }
  s1.setFrozenRows(3); s1.setColumnWidth(2,200);
  var s2=nss.insertSheet('รายการลา');
  var h2=['วันที่ลา','รหัสพนักงาน','ชื่อ','สัญชาติ','ประเภทการลา','ช่วงเวลา','ชั่วโมง','เต็มวัน','สถานะการแจ้ง','ผลอนุมัติ','เหตุผล','แจ้งเมื่อ'];
  s2.getRange(1,1,1,h2.length).setValues([h2]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff');
  if(detail.length){ s2.getRange(2,2,detail.length,1).setNumberFormat('@'); s2.getRange(2,1,detail.length,h2.length).setValues(detail); s2.getRange(2,1,detail.length,1).setNumberFormat('dd/mm/yyyy'); }
  s2.setFrozenRows(1); s2.setColumnWidth(3,180); s2.setColumnWidth(6,220); s2.setColumnWidth(11,240);
  exportOtInto_(nss, cyc, company);
  return {url:nss.getUrl(), name:nss.getName(), people:sum.length, rows:detail.length, label:cyc.label};
}

function menuExportMAIN_(){ menuExport_('MAIN'); }

function menuExportALT_(){ menuExport_('ALT'); }

function menuExport_(co){
  try{
    var r=exportCompanyReport_('MAIN');
    _alert_('✅ สร้างไฟล์แล้ว: '+r.name+'\n\n'+r.people+' คน • '+r.rows+' รายการลา\n\nเปิดไฟล์: '+r.url+'\n\n(รอบที่ส่งออก = เดือนที่เลือกไว้ในช่อง B1 ของชีต Payroll_TH)\nไฟล์นี้ไม่มีข้อมูลค่าแรง แชร์ให้ฝ่ายบัญชีได้เลย');
  }catch(e){ _alert_('❌ ส่งออกไม่สำเร็จ: '+e); }
}

function menuSetup_(){ var r=setup(); _alert_(r); }

/* ==================================================================
 *  ปุ่มเมนู Telegram (แผงปุ่มปักหมุด / ปุ่มใต้สรุปรายวัน / ค้นหาด้วยปุ่ม)
 * ================================================================== */
/* ตัดข้อความยาวเกินที่ Telegram รับได้ (ตัดที่ขึ้นบรรทัด ไม่ผ่ากลางแท็ก) */
function tgClip_(text){
  text=String(text);
  if(text.length>3900){ var cut=text.lastIndexOf('\n',3900); text=text.slice(0,cut>0?cut:3900)+'\n…'; }
  return text;
}

/* ปุ่มติดล่างจอ (ทั้งในกลุ่มและแชตส่วนตัว) — 6 ปุ่มที่ใช้บ่อย + "เพิ่มเติม" */
var TG_BTN = {today:'☀️ วันนี้', tomorrow:'🌙 พรุ่งนี้', week:'🗓 7 วัน', pending:'⏳ รออนุมัติ', ot:'⏱ โอที', more:'➕ เพิ่มเติม',
  month:'📊 สรุปเดือน', find:'🔎 ค้นหา', cal:'🔁 วันหยุด'};

var TG_ASK = {find:'🔎 ค้นหาพนักงาน', pick:'📅 ดูวันอื่น'};

function tgIsPrivate_(chatId){ return Number(chatId)>0; }

function tgMoreKb_(mode){
  mode=(mode==='q')?'q':'p';
  return {inline_keyboard:[
    [{text:'📊 สรุปเดือนนี้', callback_data:mode+':month'},{text:'⏱ โอที', callback_data:mode+':ot'}],
    [{text:'🔎 ค้นหาพนักงาน', callback_data:mode+':find'},{text:'📅 ดูวันอื่น', callback_data:mode+':pick'}],
    [{text:'🔁 วันหยุด · สลับวัน', callback_data:mode+':cal'}],
    [{text:'⬅️ กลับ', callback_data:'m:main:'+mode}]
  ]};
}

function tgBottomKb_(){
  return {keyboard:[
    [{text:TG_BTN.today},{text:TG_BTN.tomorrow},{text:TG_BTN.week}],
    [{text:TG_BTN.pending},{text:TG_BTN.ot},{text:TG_BTN.more}]
  ], resize_keyboard:true, is_persistent:true, input_field_placeholder:'กดปุ่มด้านล่าง'};
}

/* ข้อความใหม่ทุกข้อความของบอทแนบปุ่มล่างจอไปด้วย (ปุ่มจึงค้างอยู่ตลอด และคนที่เพิ่งเข้ากลุ่มก็ได้ปุ่ม) */
function tgKb_(chatId){ return tgBottomKb_(); }

/* แสดงคำตอบ: ถ้ากดมาจากแผงปุ่ม (editMsgId) → เปลี่ยนข้อความเดิม · ไม่งั้นส่งข้อความใหม่ */
function tgShow_(chatId, text, editMsgId, kb){
  if(editMsgId){
    var stamp='\n\n<i>🕒 อัปเดต '+Utilities.formatDate(new Date(),TZ,'dd/MM HH:mm')+' น.</i>';
    var r=tgApi_('editMessageText',{chat_id:chatId, message_id:editMsgId, text:tgClip_(text)+stamp, parse_mode:'HTML',
      disable_web_page_preview:true, reply_markup:(kb||tgMenuKb_('p'))});
    if(r && (r.ok || /not modified/i.test(String(r.description||'')))) return r;
    // แก้ข้อความเดิมไม่ได้ (ถูกลบ/เก่าเกิน) → ส่งใหม่แทน
  }
  return tgSendTo_(chatId, text, kb||tgKb_(chatId));
}

/* บอทถามกลับ (ค้นหาพนักงาน / ดูวันอื่น) — ในกลุ่มใช้ "ตอบกลับข้อความนี้" บอทจึงเห็นคำตอบแม้เปิดโหมดความเป็นส่วนตัว */
function tgAsk_(chatId, kind, from){
  kind=(kind==='pick')?'pick':'find';
  var priv=tgIsPrivate_(chatId);
  var who=(!priv && from && from.id) ? ('<a href="tg://user?id='+from.id+'">'+tgEscape_(String(from.first_name||'คุณ').slice(0,40))+'</a> ') : '';
  var body=(kind==='pick') ? 'พิมพ์วันที่ที่อยากดู เช่น <code>15/10</code>' : 'พิมพ์ชื่อหรือรหัสพนักงาน เช่น <code>K001</code> หรือ <code>สมชาย</code>';
  var text='<b>'+TG_ASK[kind]+'</b>\n'+who+body;
  if(priv) return tgSendTo_(chatId, text, tgBottomKb_());
  return tgSendTo_(chatId, text+'\n<i>(พิมพ์ตอบกลับข้อความนี้)</i>',
    {force_reply:true, selective:!!who, input_field_placeholder:(kind==='pick'?'เช่น 15/10':'ชื่อ หรือ รหัสพนักงาน')});
}

function tgPendingGroups_(){
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), groups={}, order=[];
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][9])!=='ลาพักร้อน' || String(vals[i][12])!=='รออนุมัติ') continue;
    var gid=String(vals[i][15]).trim(); if(!gid) continue;
    if(!groups[gid]){ groups[gid]={gid:gid, emp:String(vals[i][2]).trim(), dates:[], reason:vals[i][11]}; order.push(groups[gid]); }
    groups[gid].dates.push(fmtDate_(vals[i][5]));
  }
  return order;
}

/* หาพนักงานจากชื่อ/รหัส — รหัสตรงเป๊ะ = คืนคนเดียว */
function tgFindEmps_(arg){
  var q=String(arg||'').trim().toLowerCase(), emps=empMap_(), hits=[], exact=null;
  if(!q) return [];
  Object.keys(emps).forEach(function(id){ var e=emps[id]; if(!e.active) return;
    if(id.toLowerCase()===q){ exact=e; return; }
    if(id.toLowerCase().indexOf(q)>=0 || String(e.name_th||'').toLowerCase().indexOf(q)>=0 || String(e.name_mm||'').toLowerCase().indexOf(q)>=0) hits.push(e); });
  return exact ? [exact] : hits;
}

/* ข้อมูลการลาของพนักงาน 1 คน (ปีนี้ใช้ไปเท่าไร / กำลังจะลา / ลาล่าสุด) */
function tgEmpCardCore_(e){
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), today=todayStr_(), yr=today.slice(0,4);
  var U={'ลาป่วย':0,'ลากิจ':0,'ลาพักร้อน':0}, up=[], past=[];
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][2]).trim()!==e.emp_id) continue;
    var ap=String(vals[i][12]||''); if(ap==='ไม่อนุมัติ') continue;
    var ds=fmtDate_(vals[i][5]), type=String(vals[i][9]), h=Number(vals[i][7])||0;
    if(ap!=='รออนุมัติ' && ds.slice(0,4)===yr && U[type]!==undefined) U[type]+=h;
    var line=fmtDMY_(ds)+' '+tgEscape_(type)+' '+(h>=WORK_HOURS_PER_DAY?'เต็มวัน':(h+' ชม.'))+(ap==='รออนุมัติ'?' ⏳':'');
    if(ds>=today) up.push({ds:ds,t:line}); else past.push({ds:ds,t:line});
  }
  up.sort(function(a,b){return a.ds<b.ds?-1:1;}); past.sort(function(a,b){return a.ds<b.ds?1:-1;});
  var meta=getDailyMeta_(e.emp_id), dd=function(h){ return Math.round(h/WORK_HOURS_PER_DAY*100)/100; };
  var s=summaryFlag_(e.nationality)+' <b>'+tgEscape_(e.name_th||e.name_mm)+'</b> ('+tgEscape_(e.emp_id)+')'+
    (companyMix_().show?('\n🏢 '+tgEscape_(coTag_(e))):'')+
    '\n\n<b>ปี '+yr+' ใช้ไป</b>'+
    '\n🤒 ลาป่วย '+dd(U['ลาป่วย'])+' วัน'+(Number(meta.qSick)>0?(' / โควต้า '+meta.qSick):'')+
    '\n👔 ลากิจ '+dd(U['ลากิจ'])+' วัน'+(Number(meta.qPersonal)>0?(' / โควต้า '+meta.qPersonal):'')+
    '\n🏖️ ลาพักร้อน '+dd(U['ลาพักร้อน'])+' วัน';
  s+='\n\n<b>กำลังจะลา</b>\n'+(up.length?up.slice(0,6).map(function(x){return '• '+x.t;}).join('\n'):'– ไม่มี –');
  s+='\n\n<b>ลาล่าสุด</b>\n'+(past.length?past.slice(0,5).map(function(x){return '• '+x.t;}).join('\n'):'– ไม่มี –');
  return s;
}

/* ตอบการค้นหาพนักงาน: เจอคนเดียว = แสดงข้อมูลเลย · เจอหลายคน = ขึ้นปุ่มชื่อให้กดเลือก */
function tgWhoReply_(chatId, arg, editMsgId){
  var hits=tgFindEmps_(arg);
  if(!hits.length){ tgShow_(chatId,'🔎 ไม่พบพนักงานที่ตรงกับ "'+tgEscape_(String(arg).slice(0,60))+'"\nลองพิมพ์สั้นลง หรือใช้รหัสพนักงาน', editMsgId); return; }
  if(hits.length===1){ tgShow_(chatId, tgEmpCard_(hits[0]), editMsgId); return; }
  hits.sort(function(a,b){ return String(a.emp_id)<String(b.emp_id)?-1:1; });
  var rows=[];
  hits.slice(0,8).forEach(function(e){
    var data='w:'+e.emp_id;
    if(unescape(encodeURIComponent(data)).length>64) return;      // Telegram จำกัด 64 ไบต์
    rows.push([{text:summaryFlag_(e.nationality)+' '+String(e.name_th||e.name_mm||'').slice(0,40)+' ('+e.emp_id+')', callback_data:data}]);
  });
  var text='🔎 พบ <b>'+hits.length+'</b> คน — กดเลือกได้เลย 👇'+(hits.length>8?'\n<i>(แสดง 8 คนแรก ถ้าไม่เจอให้พิมพ์ให้เจาะจงขึ้น)</i>':'');
  if(editMsgId){
    var r=tgApi_('editMessageText',{chat_id:chatId, message_id:editMsgId, text:text, parse_mode:'HTML', reply_markup:{inline_keyboard:rows}});
    if(r && r.ok) return;
  }
  tgSendTo_(chatId, text, {inline_keyboard:rows});
}

/* เมนูคำสั่ง (ปุ่ม "/" หรือ "Menu" ข้างช่องพิมพ์) — ใส่เฉพาะที่ใช้บ่อย คำสั่งอื่นยังพิมพ์ใช้ได้เหมือนเดิม */
function tgSetCommands_(){
  return tgApi_('setMyCommands',{commands:[
    {command:'menu',description:'📋 เปิดปุ่มเมนู'},
    {command:'today',description:'☀️ วันนี้ใครลา'},
    {command:'tomorrow',description:'🌙 พรุ่งนี้ใครลา'},
    {command:'pending',description:'⏳ รออนุมัติ (ลาพักร้อน + โอที)'},
    {command:'ot',description:'⏱ โอที'},
    {command:'cal',description:'🔁 วันหยุด · สลับวัน'}]});
}

/* ส่ง: (1) ข้อความที่ทำให้ปุ่มล่างจอขึ้น (2) แผงปุ่มปักหมุด (แบบกดแล้วเปลี่ยนข้อความเดิม) */
function tgPostPanel_(){
  var token=getSetting_('telegram_bot_token'), chatId=String(getSetting_('telegram_chat_id')||'').trim();
  if(!token||!chatId) return '❌ ยังไม่ได้ตั้งค่า Telegram (token / chat id) ในชีต Settings';
  tgSetCommands_();
  var k=tgSendTo_(chatId, '⌨️ <b>ปุ่มเมนูอยู่ด้านล่างจอแล้ว</b>\nกดได้เลย ไม่ต้องพิมพ์คำสั่ง\n<i>ถ้าไม่เห็นปุ่ม: กดไอคอนแป้นสี่เหลี่ยม ⌨️ ข้างช่องพิมพ์ หรือพิมพ์ /menu</i>', tgBottomKb_());
  if(!k || !k.ok) return '❌ ส่งปุ่มเข้ากลุ่มไม่สำเร็จ: '+((k && k.description) || 'ไม่ทราบสาเหตุ');
  var old=String(getSetting_('tg_panel_msg_id')||'').trim();
  var r=tgSendTo_(chatId, '📋 <b>ระบบลาหยุด '+APP.NAME+'</b> — แผงปุ่ม (ปักหมุด)\nกดแล้วข้อความนี้จะเปลี่ยนเป็นคำตอบ ไม่ส่งข้อความใหม่', tgMenuKb_('p'));
  if(!r || !r.ok) return '✅ ส่งปุ่มล่างจอเข้ากลุ่มแล้ว\n⚠️ ส่งแผงปุ่มปักหมุดไม่สำเร็จ: '+((r && r.description) || '');
  var mid=r.result && r.result.message_id;
  if(old && String(old)!==String(mid)){
    tgApi_('unpinChatMessage',{chat_id:chatId, message_id:Number(old)});
    tgApi_('deleteMessage',{chat_id:chatId, message_id:Number(old)});
  }
  if(mid) setSetting_('tg_panel_msg_id', String(mid));
  var p=mid ? tgApi_('pinChatMessage',{chat_id:chatId, message_id:mid, disable_notification:true}) : null;
  return '✅ ส่งปุ่มล่างจอ + แผงปุ่มเมนูเข้ากลุ่ม Telegram แล้ว\n'+
    ((p && p.ok) ? '📌 ปักหมุดแผงปุ่มไว้บนสุดของกลุ่มให้แล้ว' : '⚠️ ยังปักหมุดไม่ได้ (บอทไม่ใช่แอดมินกลุ่ม) → กดค้างที่ข้อความแผงปุ่ม แล้วเลือก "ปักหมุด" เองได้');
}

function sendMenuNow(){ return say_('ส่งปุ่มเมนู', tgPostPanel_()); }

function menuPostPanel_(){ _alert_(tgPostPanel_()); }


function coLabel_(co){ return APP.NAME; }

/* ==================================================================
 *  MAIN — ติดตั้ง / คู่มือในชีต / เปลี่ยนรหัสพนักงาน / เชื่อม Telegram
 * ================================================================== */

/* รายชื่อตั้งต้น (ใส่ให้ครั้งเดียวตอนชีต Employees ยังว่าง) : [รหัส, ชื่อไทย, ชื่อพม่า/อังกฤษ, สัญชาติ, รหัสรูป] */
var SEED_STAFF = [];       // เว้นว่าง = เริ่มจากรายชื่อเปล่า (เพิ่ม/นำเข้าพนักงานจากหน้าผู้ดูแล) • ถ้าจะใส่ตั้งต้น: ['รหัส','ชื่อไทย','ชื่อพม่า/อังกฤษ','TH หรือ MM','']

var GUIDE_SHEET = 'อ่านก่อน';

/* คำอธิบายของแต่ละค่าในชีต Settings (คอลัมน์ C) */
var SETTING_DESC = {
  admin_pin:'รหัส PIN เข้าหน้าผู้ดูแลของเว็บ (เปลี่ยนได้ที่นี่ หรือในหน้าผู้ดูแล → ตั้งค่า)',
  photo_folder_id:'โฟลเดอร์ใน Google Drive ที่เก็บรูปพนักงาน — ระบบใส่ให้เอง ห้ามแก้',
  company_name:'ชื่อบริษัท',
  company_address:'ที่อยู่บริษัท',
  telegram_bot_token:'รหัสของบอท Telegram (ได้จาก @BotFather) — ห้ามให้คนอื่นเห็น',
  telegram_chat_id:'เลขประจำกลุ่ม Telegram ที่ให้บอทแจ้งเตือน — ระบบหาให้เองตอนกด "เชื่อม Telegram"',
  web_app_url:'ที่อยู่ของระบบหลังบ้าน (ลงท้าย /exec) — ระบบจำให้เอง ห้ามแก้',
  cutoff_day:'วันตัดรอบเงินเดือน เช่น 25 = รอบวันที่ 26 เดือนก่อน ถึง 25 เดือนนี้ (แก้ได้ 1–28)',
  summary_eve_hour:'สรุป "พรุ่งนี้ใครลา" ส่งกี่โมง (ชั่วโมง 0–23)',
  summary_eve_min:'สรุป "พรุ่งนี้ใครลา" ส่งนาทีที่เท่าไร',
  summary_morning_hour:'สรุป "วันนี้ใครลา" ส่งกี่โมง (ชั่วโมง 0–23)',
  summary_morning_min:'สรุป "วันนี้ใครลา" ส่งนาทีที่เท่าไร',
  site_url:'ที่อยู่หน้าเว็บ — ใช้ทำลิงก์ตรวจสลิป ห้ามแก้ถ้าไม่ได้ย้ายเว็บ',
  tg_webhook_url:'ที่อยู่ตัวส่งต่อ Telegram ของเว็บคุณ เช่น https://<เว็บของคุณ>/tg — แนะนำให้ใส่ (ถ้าเว้นว่าง Telegram จะยิงเข้า Apps Script ตรง ๆ ซึ่งมักช้าหรือค้าง) ใส่แล้วรัน connectTelegramNow',
  tg_webhook_secret:'รหัสลับระหว่าง Telegram กับระบบ — ระบบสร้างเอง ห้ามแก้ ห้ามลบ',
  tg_panel_msg_id:'เลขข้อความแผงปุ่มที่ปักหมุดในกลุ่ม — ระบบใส่เอง',
  ot_back_days:'พนักงานแจ้งโอทีย้อนหลังได้กี่วัน (0 = แจ้งได้เฉพาะวันที่ทำ, แก้ได้ 0–62)',
  ot_need_confirm:'1 = โอทีที่พนักงานแจ้งเองต้องมีคนกดยืนยันก่อนจึงนับ (แนะนำ) • 0 = นับเลยไม่ต้องยืนยัน',
  drive_folder_id:'โฟลเดอร์หลักของระบบใน Google Drive (เก็บไฟล์ชีตนี้ + โฟลเดอร์รูปพนักงาน) — ระบบใส่ให้เอง ห้ามแก้',
  slip_secret:'รหัสลับสำหรับตรวจสลิปใบลาว่าเป็นของจริง — ระบบสร้างเอง ห้ามแก้ ห้ามลบ (ถ้าแก้ สลิปเก่าจะตรวจไม่ผ่านทั้งหมด)'
};

/* ★ ติดตั้ง/อัปเดต: รันซ้ำได้ ข้อมูลเดิมไม่ถูกลบ */
function setup(){
  try{ ss_().setSpreadsheetTimeZone(TZ); }catch(e){}      // วันที่ในชีตต้องเป็นเวลาไทย
  _styleErr=[];
  var msg=setupCore_();
  var seeded=seedStaff_();
  if(seeded) syncPayrollRosters_();
  otFormatSheet_(otSheet_());
  ensureSetting_('ot_back_days', '7');
  ensureSetting_('ot_need_confirm', '1');
  polishSheets_();
  styleAllSheets_();
  var stz=''; try{ stz=Session.getScriptTimeZone(); }catch(e){}
  if(stz && stz!==TZ) msg+='\n\n⚠️ เขตเวลาของ Apps Script เป็น "'+stz+'" — ต้องแก้เป็นเวลาไทย: กดรูปเฟือง (การตั้งค่าโปรเจ็กต์) ด้านซ้าย → เขตเวลา → เลือก (GMT+07:00) Bangkok แล้วรัน setup อีกครั้ง';
  if(_styleErr.length) msg+='\n\n⚠️ ตกแต่งชีตบางส่วนไม่สำเร็จ (ระบบใช้งานได้ปกติ) — ส่งข้อความนี้ให้ผู้ดูแลระบบ:\n'+_styleErr.join('\n');
  msg+=(seeded?('\nใส่รายชื่อพนักงานตั้งต้นให้แล้ว '+seeded+' คน'):'');
  try{ Logger.log(msg); }catch(e){}
  try{ ss_().toast(_styleErr.length?'เสร็จ แต่ตกแต่งชีตบางส่วนไม่สำเร็จ — ดูบันทึกการทำงาน':'จัดหน้าตาชีตเรียบร้อย', 'ตั้งค่าเสร็จ', 8); }catch(e){}
  return msg;
}

function seedStaff_(){
  var sh=sheet_(SHEETS.EMP);
  if(sh.getLastRow()>1) return 0;                 // มีพนักงานอยู่แล้ว → ไม่ใส่ซ้ำ
  var at=nowStr_();
  var rows=SEED_STAFF.map(function(s){ return [s[0], s[1], s[2], s[3], s[4], true, true, '', at, 'MAIN', '', '']; });
  if(!rows.length) return 0;
  sh.getRange(2,1,rows.length,1).setNumberFormat('@');
  sh.getRange(2,1,rows.length,12).setValues(rows);
  _coMix=null;
  return rows.length;
}

/* จัดชีตให้อ่านง่าย: คำอธิบาย Settings, หน้า "อ่านก่อน", สีแท็บ, ลำดับชีต, ซ่อนช่องที่ระบบใช้เอง */
function polishSheets_(){
  var ss=ss_();
  // Settings: คอลัมน์ C = คำอธิบาย
  var st=sheet_(SHEETS.SET);
  if(st){
    st.getRange(1,1,1,3).setValues([['ชื่อค่า (ห้ามแก้)','ค่า','คำอธิบาย']]).setFontWeight('bold').setBackground('#1F4E79').setFontColor('#ffffff');
    var v=st.getDataRange().getValues(), d=[];
    for(var i=1;i<v.length;i++) d.push([SETTING_DESC[String(v[i][0]).trim()] || (v[i].length>2?v[i][2]:'') || '']);      // ค่าที่ไม่มีคำอธิบายของระบบ → คงโน้ตที่ผู้ใช้พิมพ์ไว้
    if(d.length) st.getRange(2,3,d.length,1).setValues(d).setWrap(true).setFontColor('#555555');
    st.setColumnWidth(3,520); st.setFrozenRows(1);
  }
  // Employees: คอลัมน์ J–L เป็นช่องที่ระบบใช้เอง → ซ่อน
  if(!APP.MULTI){ try{ sheet_(SHEETS.EMP).hideColumns(10,3); }catch(e){} }
  buildGuideSheet_();
  // สีแท็บ + ลำดับชีต (หน้าที่คนใช้อยู่ซ้าย / ชีตที่ระบบใช้เองอยู่ขวา)
  sumSheet_();
  var order=[[GUIDE_SHEET,'#8CC63F'],[SUM_SHEET,'#F2B705'],[SHEETS.EMP,'#1F4E79'],[SHEETS.LEAVE,'#1F4E79'],[OT_SHEET,'#5A9E1B'],[SHEETS.PAY_TH,'#1D9BD7'],[SHEETS.PAY_MM,'#1D9BD7'],[SHEETS.PAY_MAIN_TH,'#1D9BD7'],[SHEETS.PAY_MAIN_MM,'#1D9BD7'],
    [SHEETS.DAILY_TH,'#1D9BD7'],[SHEETS.DAILY_MM,'#1D9BD7'],[SHEETS.HOLIDAY,'#F59E0B'],[SHEETS.WORKDAY,'#F59E0B'],[SHEETS.CANCEL,'#9CA3AF'],[SHEETS.TRANSFER,'#9CA3AF'],[SHEETS.SET,'#9CA3AF']].filter(function(o){ return !!o[0]; });
  order.forEach(function(o, idx){
    var sh=sheet_(o[0]); if(!sh) return;
    try{ sh.setTabColor(o[1]); }catch(e){}
    var hid=false; try{ hid=sh.isSheetHidden(); }catch(e){}
    if(hid) return;                                  // ชีตที่ซ่อนอยู่ไม่ต้องย้าย (การเลือกชีตจะทำให้มันโผล่)
    try{ ss.setActiveSheet(sh); ss.moveActiveSheet(idx+1); }catch(e){}
  });
  // ลบชีตเปล่าที่ติดมากับไฟล์ใหม่
  ['Sheet1','ชีต1','แผ่น1','แผ่นงาน1'].forEach(function(n){
    try{ var s=ss.getSheetByName(n); if(!APP.MULTI && s && s.getLastRow()===0 && !s.getCharts().length && ss.getSheets().length>1) ss.deleteSheet(s); }catch(e){}
  });
  try{ ss.setActiveSheet(sheet_(GUIDE_SHEET)); }catch(e){}
}

function buildGuideSheet_(){
  var ss=ss_(), sh=sheet_(GUIDE_SHEET);
  if(sh && sh.getLastRow()>0 && String(sh.getRange(1,1).getValue()).indexOf('ระบบแจ้งลาพนักงาน')!==0){ shelveSheet_(sh); sh=null; }   // ชีตชื่อนี้ของผู้ใช้เอง → เปลี่ยนชื่อเก็บไว้
  if(!sh) sh=ss.insertSheet(GUIDE_SHEET);
  sh.clear();
  try{ sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart(); }catch(e){}
  var site=String(getSetting_('site_url')||APP.SITE);
  var co=getCutoffDay_(APP.CO);
  var NAVY='#1F4E79', GREEN='#8CC63F', SOFT='#EDF3F8';
  var rows=[], fmt=[];          // fmt: [row, 'title'|'h'|'th'|'warn']
  function add(a,b,c,kind){ rows.push([a||'',b||'',c||'']); if(kind) fmt.push([rows.length,kind]); }
  add('ระบบแจ้งลาพนักงาน — '+APP.TITLE,'','','title');
  add('หน้าเว็บสำหรับพนักงาน: '+site+'    •    หน้าผู้ดูแล: เปิดเว็บเดียวกัน เลื่อนลงล่างสุด กด "สำหรับผู้ดูแล" แล้วใส่ PIN','','');
  add('','','');
  add('ระบบทำงานอย่างไร','','','h');
  add('1','พนักงานเปิดเว็บ เลือกชื่อตัวเอง เลือกประเภทลา วันที่ และช่วงเวลา แล้วกดส่ง','พนักงานได้ "สลิปใบลา" เก็บเป็นหลักฐาน (มี QR ไว้ตรวจว่าเป็นของจริง)');
  add('2','ใบลาเข้ามาที่ชีต Leaves ทันที 1 แถว = 1 วันที่ลา','ไม่ต้องพิมพ์เองในชีต');
  add('3','บอท Telegram แจ้งเข้ากลุ่มทุกครั้งที่มีคนลา และส่งสรุป "พรุ่งนี้ใครลา" ตอนเย็น / "วันนี้ใครลา" ตอนเช้า','ลาพักร้อนต้องมีคนกด ✅ อนุมัติ ในกลุ่มก่อน จึงจะนับ');
  add('4','โอที: พนักงานกด "แจ้งโอที" ในเว็บ เลือกวันและช่วงเวลา → เข้าชีต OT สถานะ "รอยืนยัน" → หัวหน้ากด ✅ ยืนยัน ใน Telegram จึงจะนับ','หรือผู้ดูแลลงโอทีให้ทีเดียวหลายคนจากหน้าผู้ดูแล → โอที (นับเลย ไม่ต้องยืนยัน)');
  add('5','ชีต Payroll_TH / Payroll_MM รวมชั่วโมงลา และชั่วโมงโอที ของแต่ละคนตามรอบเงินเดือนให้อัตโนมัติ','ดูรวมทุกคนได้ที่ชีต สรุป (เปลี่ยนรอบที่ช่องสีเหลือง)');
  add('','','');
  add('แต่ละชีตคืออะไร','','','h');
  add('ชีต','ใช้ทำอะไร','แก้เองได้ไหม','th');
  add('สรุป','หน้าหลักสำหรับทำเงินเดือน: ทุกคนในหน้าเดียว — ชั่วโมงลา วันลา ยอดหัก และโอที ของรอบที่เลือก พร้อมตัวเลขรวมด้านบน','แก้ได้เฉพาะช่องสีเหลือง (เลือกรอบเงินเดือน) — ชีต Payroll ทุกชีตเปลี่ยนตามช่องนี้');
  add('Employees','รายชื่อพนักงานทั้งหมด (รหัส ชื่อ สัญชาติ รูป)','ควรเพิ่ม/แก้/ลบ จากหน้าผู้ดูแลในเว็บ — ไม่ควรพิมพ์ในชีตเอง');
  add('Leaves','ใบลาทุกใบ: ใครลา วันไหน กี่ชั่วโมง ประเภทอะไร แจ้งล่วงหน้าหรือย้อนหลัง','ดูได้ กรองได้ — ถ้าจะลบใบลาให้ลบจากหน้าผู้ดูแล');
  add('OT','รายการโอทีทุกรายการ: ใครทำ วันไหน กี่โมงถึงกี่โมง กี่ชั่วโมง เป็นวันทำงานหรือวันหยุด และยืนยันแล้วหรือยัง','ดูได้ กรองได้ — เพิ่ม/ยืนยัน/ลบ ให้ทำจากหน้าผู้ดูแล → โอที');
  add('Payroll_TH','สรุปของ "คนไทย" ตามรอบเงินเดือน: ลาป่วย / ลากิจ / ลาพักร้อน / รวมลา / ชั่วโมงที่ต้องหักเงิน และ โอทีวันทำงาน / ทำงานวันหยุด / รวมโอที','ดูอย่างเดียว (เป็นสูตรทั้งหน้า) — เปลี่ยนรอบที่ชีต สรุป');
  add('Payroll_MM','เหมือน Payroll_TH แต่เป็น "พนักงานพม่า"','ดูอย่างเดียว');
  add('รายวัน_TH / รายวัน_MM','ดูรายละเอียดการลาทีละคนทีละวัน กรอกค่าแรงต่อวัน โควต้าลาป่วย/ลากิจต่อปี และติ๊กว่าวันไหนหักเงิน','แก้ได้ตามช่องที่ระบุในหน้า (เลือกชื่อ ค่าแรง โควต้า ช่องติ๊กหักเงิน)');
  add('Holidays','วันหยุดของบริษัท (พนักงานจะกดลาวันนี้ไม่ได้) — วันอาทิตย์เป็นวันหยุดอยู่แล้วไม่ต้องใส่','ควรเพิ่ม/ลบ จากหน้าผู้ดูแล → วันหยุด');
  add('WorkExtra','วันอาทิตย์/วันหยุดที่ต้องมาทำงาน (สลับวันหยุด) — วันเหล่านี้กดลาได้ตามปกติ','ควรเพิ่ม/ลบ จากหน้าผู้ดูแล → วันหยุด');
  add('CancelLog','ประวัติใบลาที่พนักงานกดยกเลิก (เก็บไว้ตรวจย้อนหลัง)','ดูอย่างเดียว');
  add('Settings','ค่าตั้งของระบบ เช่น PIN, วันตัดรอบ, Telegram (มีคำอธิบายทุกบรรทัดในคอลัมน์ C)','แก้เฉพาะบรรทัดที่คำอธิบายบอกว่าแก้ได้');
  add('','','');
  add('คำที่เจอในชีต Leaves','','','h');
  add('คำ','ความหมาย','','th');
  add('ลาล่วงหน้า (สีเขียว)','แจ้งก่อนวันที่ลา','');
  add('ลากระทันหัน (สีเหลือง)','แจ้งในวันเดียวกับวันที่ลา','');
  add('ลาย้อนหลัง (สีแดง)','แจ้งหลังจากวันที่ลาผ่านไปแล้ว','');
  add('ผลอนุมัติ','ใช้กับลาพักร้อนเท่านั้น: รออนุมัติ → อนุมัติ / ไม่อนุมัติ  (ลาป่วย-ลากิจ เว้นว่าง = นับเลย)','ใบที่ "รออนุมัติ" หรือ "ไม่อนุมัติ" จะไม่ถูกนับในชีต Payroll');
  add('เต็มวัน','ลาครบ 8 ชั่วโมง (08:00–12:00 และ 13:00–17:00)','ลาเป็นรายชั่วโมงได้ ระบบนับชั่วโมงให้');
  add('','','');
  add('คำที่เจอในชีต OT','','','h');
  add('คำ','ความหมาย','','th');
  add('รอยืนยัน (สีเหลือง)','พนักงานแจ้งเอง ยังไม่มีหัวหน้ายืนยัน','ยังไม่ถูกนับในชีต Payroll');
  add('ยืนยัน (สีเขียว)','นับเป็นโอทีจริง','หัวหน้ากดยืนยันใน Telegram หรือผู้ดูแลลงให้');
  add('ไม่ยืนยัน (สีแดง) / ยกเลิก (สีเทา)','หัวหน้ากดว่าไม่ได้ทำ / พนักงานยกเลิกเอง','ไม่ถูกนับ แต่เก็บแถวไว้ให้ตรวจย้อนหลังได้');
  add('วันทำงาน','โอทีนอกเวลา 08:00–17:00 ของวันทำงานปกติ','แยกคอลัมน์ในชีต Payroll เพราะปกติจ่ายคนละอัตรากับวันหยุด');
  add('วันหยุด','ทำงานในวันอาทิตย์ หรือวันหยุดบริษัท (ชีต Holidays)','ถ้าวันอาทิตย์นั้นตั้งเป็น "วันทำงานพิเศษ" จะนับเป็นวันทำงาน');
  add('','','');
  add('รอบเงินเดือน','','','h');
  add('ตอนนี้ตัดรอบวันที่ '+co+' ของทุกเดือน','เช่น รอบเดือนตุลาคม = วันที่ '+(co+1)+' ก.ย. ถึง '+co+' ต.ค.','เปลี่ยนวันตัดรอบได้ที่หน้าผู้ดูแล → ตั้งค่า (หรือชีต Settings บรรทัด cutoff_day)');
  add('','','');
  add('งานที่ทำบ่อย','','','h');
  add('เพิ่มพนักงานใหม่','หน้าผู้ดูแล → พนักงาน → "+ เพิ่ม"  หรือ "📋 นำเข้ารายชื่อ" (คัดลอกจาก Excel มาวางได้ทีละหลายคน)','ถ้าไม่ใส่รหัส ระบบออกรหัสให้เอง');
  add('เปลี่ยนรหัสพนักงาน','หน้าผู้ดูแล → พนักงาน → แก้ไข → แก้ช่องรหัส แล้วบันทึก','ระบบเปลี่ยนรหัสในใบลาเก่าทุกใบให้เอง — ห้ามแก้รหัสในชีตด้วยมือ');
  add('พนักงานลาออก','หน้าผู้ดูแล → พนักงาน → ลบ','ชื่อหายจากหน้าเว็บ แต่ประวัติลายังอยู่ในชีตครบ');
  add('ลงโอทีให้พนักงานทีเดียวหลายคน','หน้าผู้ดูแล → โอที → เลือกวันที่ ช่วงเวลา กด "เลือกทุกคน" แล้วติ๊กออกคนที่ไม่ได้ทำ → บันทึก','นับเป็นโอทีทันที');
  add('ส่งสรุปการลา + โอที ให้ฝ่ายบัญชี','เมนูด้านบนของชีต "'+APP.MENU+'" → 📤 ส่งออกไฟล์สรุปการลา','ได้ไฟล์ Google Sheets แยกต่างหาก (สรุปการลา + สรุปโอที) ไม่มีข้อมูลค่าแรง');
  add('สลับวันหยุด / เพิ่มวันทำงาน / เพิ่มวันหยุด','Telegram: ปุ่ม ➕ เพิ่มเติม → 🔁 วันหยุด · สลับวัน  หรือ หน้าผู้ดูแล → ตั้งค่า → สลับวันหยุด','ทุกคนในกลุ่ม Telegram ทำได้ • พนักงานเห็นประกาศในหน้าแรกของเว็บ');
  add('เปลี่ยน PIN ผู้ดูแล','หน้าผู้ดูแล → ตั้งค่า → เปลี่ยน PIN','');
  if(typeof guideExtra_==='function') guideExtra_(add);
  add('','','');
  add('ข้อห้าม','','','h');
  add('⚠️','ห้ามเปลี่ยนชื่อชีต ห้ามลบชีต ห้ามแทรก/ลบ/สลับคอลัมน์ ในชีต Employees, Leaves, OT, Settings และชีตที่ซ่อนอยู่','ถ้าทำ ระบบจะอ่านข้อมูลผิดช่อง','warn');
  add('⚠️','ห้ามแก้ช่องที่เป็นสูตรในชีต Payroll','ถ้าเผลอแก้ ให้กดเมนู "'+APP.MENU+'" → 🔄 รีเฟรชรายชื่อ Payroll','warn');
  add('⚠️','ห้ามแก้หรือลบบรรทัด slip_secret และ tg_webhook_secret ในชีต Settings','','warn');

  sh.getRange(1,1,rows.length,3).setValues(rows).setWrap(true).setVerticalAlignment('top').setFontSize(11);
  sh.setColumnWidth(1,230); sh.setColumnWidth(2,520); sh.setColumnWidth(3,430);
  fmt.forEach(function(f){
    var r=sh.getRange(f[0],1,1,3);
    if(f[1]==='title') r.setFontSize(16).setFontWeight('bold').setFontColor(NAVY);
    else if(f[1]==='h') r.setFontWeight('bold').setFontSize(13).setBackground(NAVY).setFontColor('#ffffff');
    else if(f[1]==='th') r.setFontWeight('bold').setBackground(SOFT).setFontColor(NAVY);
    else if(f[1]==='warn') r.setBackground('#FFF4E5');
  });
  try{ sh.setHiddenGridlines(true); }catch(e){}
  try{ sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(pr){ pr.remove(); }); }catch(e){}
  try{ sh.getRange(1,1,rows.length,3).protect().setWarningOnly(true).setDescription('หน้าคู่มือ — ระบบสร้างใหม่ทุกครั้งที่กด setup'); }catch(e){}
}

/* ---------- เปลี่ยนรหัสพนักงาน (เปลี่ยนในใบลาเก่า/ประวัติยกเลิก/ข้อมูลรายวัน ให้ครบทุกที่) ---------- */
function replaceIdInCol_(sheetName, col, oldId, newId){
  var sh=sheet_(sheetName); if(!sh) return 0;
  var last=sh.getLastRow(); if(last<2) return 0;
  var rng=sh.getRange(2,col,last-1,1), v=rng.getValues(), n=0;
  for(var i=0;i<v.length;i++){ if(String(v[i][0]).trim()===oldId){ v[i][0]=newId; n++; } }
  if(n){ rng.setNumberFormat('@'); rng.setValues(v); }
  return n;
}

function changeEmpId_(d){
  var oldId=String(d.emp_id||'').trim(), newId=String(d.new_id||'').trim();
  var emp=findEmployee_(oldId);
  if(!emp) return err_('NO_EMP','ไม่พบพนักงาน');
  if(!validEmpId_(newId)) return err_('BAD_ID','รหัสพนักงานใช้ได้เฉพาะ A-Z 0-9 - _ . (ไม่เกิน 20 ตัว)');
  if(newId===oldId) return {ok:true, emp_id:newId, rows:0};
  if(newId.toLowerCase()!==oldId.toLowerCase() && empIdExists_(newId)) return err_('DUP_ID','รหัส '+newId+' มีอยู่แล้ว');
  var n=0;
  n+=replaceIdInCol_(SHEETS.LEAVE, 3, oldId, newId);
  n+=replaceIdInCol_(SHEETS.CANCEL, 4, oldId, newId);
  n+=replaceIdInCol_(OT_SHEET, 3, oldId, newId);
  n+=replaceIdInCol_(SHEETS.DDATA, 1, oldId, newId);
  n+=replaceIdInCol_(SHEETS.DMETA, 1, oldId, newId);
  if(SHEETS.TRANSFER) n+=replaceIdInCol_(SHEETS.TRANSFER, 2, oldId, newId);
  sheet_(SHEETS.EMP).getRange(emp.row,1).setNumberFormat('@').setValue(newId);
  _coMix=null;
  syncPayrollRosters_();
  // หน้ารายวัน: ถ้ากำลังเลือกคนนี้อยู่ ให้ชื่อในช่องเลือกเป็นรหัสใหม่
  [SHEETS.DAILY_TH, SHEETS.DAILY_MM].forEach(function(nm){
    try{ var ds=sheet_(nm); if(ds && String(ds.getRange('B1').getValue()).indexOf(oldId+' — ')===0) ds.getRange('B1').setValue(newId+' — '+(emp.name_th||emp.name_mm)); }catch(e){}
  });
  return {ok:true, emp_id:newId, old_id:oldId, rows:n};
}

/* ---------- ★ เชื่อม Telegram ทั้งหมดในครั้งเดียว ----------
 * ก่อนรัน: (1) ใส่ token ของบอทแล้ว (หน้าผู้ดูแล → ตั้งค่า หรือชีต Settings)  (2) เพิ่มบอทเข้ากลุ่ม แล้วพิมพ์ /start ในกลุ่ม 1 ครั้ง
 * ทำให้: หาเลขกลุ่ม → ผูก webhook (ปุ่มอนุมัติ/ปุ่มเมนู) → เปิดสรุปรายวันเย็น+เช้า → ส่งแผงปุ่มเข้ากลุ่ม */
/* รันจากหน้า Apps Script: ผลขึ้นใน "บันทึกการดำเนินการ" ด้านล่าง (ไม่ต้องสลับไปกดตกลงในหน้าชีต) */
function say_(title, msg){ Logger.log(msg); try{ ss_().toast(String(msg).slice(0,300), title, 20); }catch(e){} return msg; }

function connectTelegramNow(){ return say_('เชื่อม Telegram', connectTelegram_()); }

function menuConnectTelegram_(){ _alert_(connectTelegram_()); }

function connectTelegram_(){
  var token=String(getSetting_('telegram_bot_token')||'').trim();
  if(!token) return '❌ ยังไม่ได้ใส่ token ของบอท\n→ เปิดหน้าเว็บ → สำหรับผู้ดูแล → ตั้งค่า → Telegram → วาง token แล้วกดบันทึก จากนั้นรันฟังก์ชันนี้ใหม่';
  var me=tgApi_('getMe',{});
  if(!me || !me.ok) return '❌ token ของบอทไม่ถูกต้อง ('+((me&&me.description)||'')+')\n→ ก๊อปปี้ token จาก @BotFather มาวางใหม่ให้ครบทั้งบรรทัด';
  var out=['🤖 บอท: @'+(me.result&&me.result.username)];
  var chatId=String(getSetting_('telegram_chat_id')||'').trim();
  if(!chatId){
    tgApi_('deleteWebhook',{});                         // ต้องไม่มี webhook ก่อน จึงจะอ่านข้อความล่าสุดได้
    var up=tgApi_('getUpdates',{timeout:0});
    var found=null;
    ((up&&up.result)||[]).forEach(function(u){
      var c=(u.message&&u.message.chat) || (u.my_chat_member&&u.my_chat_member.chat) || (u.channel_post&&u.channel_post.chat);
      if(c && (c.type==='group'||c.type==='supergroup')) found=c;      // เอากลุ่มล่าสุด
    });
    if(!found) return out.join('\n')+'\n\n❌ ยังไม่เจอกลุ่ม Telegram\n→ 1) เพิ่มบอทตัวนี้เข้ากลุ่ม  2) พิมพ์ /start ในกลุ่ม 1 ครั้ง  3) รันฟังก์ชันนี้ใหม่';
    chatId=String(found.id);
    setSetting_('telegram_chat_id', chatId);
    out.push('👥 กลุ่ม: '+(found.title||chatId));
  } else out.push('👥 กลุ่ม: ตั้งไว้แล้ว ('+chatId+')');
  if(!webAppUrl_()) return out.join('\n')+'\n\n❌ ระบบยังไม่รู้ที่อยู่ของตัวเอง (web_app_url)\n→ เปิดหน้าเว็บแจ้งลา 1 ครั้งให้รายชื่อพนักงานขึ้น แล้วรันฟังก์ชันนี้ใหม่';
  var w=setTelegramWebhook();
  out.push(/"ok":true/.test(w) ? '🔗 ผูกปุ่มอนุมัติ/ปุ่มเมนู: สำเร็จ' : ('⚠️ ผูก webhook ไม่สำเร็จ: '+w));
  try{ var t=installSummaryTriggers_(); out.push('🔔 สรุปรายวัน: เย็น ~'+pad2_(t.eh)+':'+pad2_(t.em)+' น. / เช้า ~'+pad2_(t.mh)+':'+pad2_(t.mm)+' น.'); }
  catch(e){ out.push('⚠️ เปิดสรุปรายวันไม่สำเร็จ: '+e); }
  out.push(tgPostPanel_());
  return out.join('\n');
}

/* ---------- ★ จัดโฟลเดอร์ใน Google Drive ให้อยู่ที่เดียว: โฟลเดอร์หลัก 1 อัน มีไฟล์ชีตนี้ + โฟลเดอร์รูปพนักงาน (แชร์ทั้งโฟลเดอร์ได้) — รันซ้ำได้ */
var DRIVE_FOLDER = 'ระบบลา + โอที';

function organizeDriveNow(){ return say_('จัดโฟลเดอร์', organizeDrive_()); }

function menuOrganizeDrive_(){ _alert_(organizeDrive_()); }

function driveInFolder_(item, folder){
  var it=item.getParents(), fid=folder.getId();
  while(it.hasNext()){ if(it.next().getId()===fid) return true; }
  return false;
}

function organizeDrive_(){
  var out=[], ss=ss_();
  // 1) โฟลเดอร์หลัก
  var parent=null, pid=String(getSetting_('drive_folder_id')||'').trim();
  if(pid){ try{ parent=DriveApp.getFolderById(pid); if(parent.isTrashed()) parent=null; }catch(e){ parent=null; } }
  if(!parent){ var it=DriveApp.getFoldersByName(DRIVE_FOLDER); while(it.hasNext() && !parent){ var f=it.next(); if(!f.isTrashed()) parent=f; } }
  if(!parent){ parent=DriveApp.createFolder(DRIVE_FOLDER); out.push('📁 สร้างโฟลเดอร์ "'+DRIVE_FOLDER+'" แล้ว'); }
  else out.push('📁 ใช้โฟลเดอร์ "'+parent.getName()+'" ที่มีอยู่');
  setSetting_('drive_folder_id', parent.getId());
  // 2) ไฟล์ชีต
  var file=DriveApp.getFileById(ss.getId());
  if(driveInFolder_(file, parent)) out.push('📄 ไฟล์ชีตอยู่ในโฟลเดอร์แล้ว');
  else { file.moveTo(parent); out.push('📄 ย้ายไฟล์ชีต "'+ss.getName()+'" เข้าโฟลเดอร์แล้ว'); }
  // 3) โฟลเดอร์รูป
  var pf=getOrCreatePhotoFolder_();
  if(driveInFolder_(pf, parent)) out.push('🖼 โฟลเดอร์รูป "'+pf.getName()+'" อยู่ในโฟลเดอร์แล้ว');
  else { pf.moveTo(parent); out.push('🖼 ย้ายโฟลเดอร์รูป "'+pf.getName()+'" เข้าโฟลเดอร์แล้ว'); }
  try{ pf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
  setSetting_('photo_folder_id', pf.getId());
  // 4) รูปพนักงาน: ไฟล์ที่ยังอยู่นอกโฟลเดอร์รูปของระบบนี้ → ก๊อบปี้เข้ามา แล้วชี้ไปที่ไฟล์ใหม่
  var sh=sheet_(SHEETS.EMP), vals=sh.getDataRange().getValues(), copied=0, kept=0, none=[], fail=[];
  for(var i=1;i<vals.length;i++){
    var id=String(vals[i][0]).trim(), ph=String(vals[i][4]||'').trim();
    if(!id || String(vals[i][6]).toUpperCase()==='FALSE') continue;
    if(!ph){ none.push(id); continue; }
    try{
      var src=DriveApp.getFileById(ph);
      if(driveInFolder_(src, pf)){ kept++; continue; }
      var ext=(/\.[A-Za-z0-9]{2,5}$/.exec(String(src.getName()||''))||['.jpg'])[0];
      var cp=src.makeCopy(id+ext, pf);
      try{ cp.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(e){}
      sh.getRange(i+1,5).setValue(cp.getId());
      copied++;
    }catch(e){ fail.push(id); }
  }
  out.push('👤 รูปพนักงาน: ก๊อบปี้เข้าโฟลเดอร์รูป '+copied+' คน'+(kept?(' • อยู่ในโฟลเดอร์อยู่แล้ว '+kept+' คน'):''));
  if(none.length) out.push('   ยังไม่มีรูป: '+none.join(', ')+' (อัปโหลดได้ที่หน้าผู้ดูแล → พนักงาน → แก้)');
  if(fail.length) out.push('   ⚠️ ก๊อบปี้ไม่ได้: '+fail.join(', ')+' (ไฟล์รูปเดิมอาจถูกลบ → อัปโหลดใหม่ที่หน้าผู้ดูแล)');
  out.push('\n🔗 โฟลเดอร์: '+parent.getUrl());
  return out.join('\n');
}

/* ==================================================================
 *  MAIN — ระบบโอที (OT)
 *  ชีต OT : 1 แถว = 1 คน 1 วัน • นับเงินเฉพาะแถวที่สถานะ "ยืนยัน"
 *  A แจ้งเมื่อ | B เลขโอที | C รหัสพนักงาน | D ชื่อ | E สัญชาติ | F วันที่ทำโอที | G ช่วงเวลา (รายชั่วโมง) | H เวลา (ย่อ)
 *  I ชั่วโมง | J ประเภทวัน | K สถานะ | L ผู้ยืนยัน | M ยืนยันเมื่อ | N ที่มา | O หมายเหตุ | P รหัสกลุ่ม | Q อุปกรณ์
 * ================================================================== */
var OT_SHEET = 'OT';

/* ช่องเวลาโอที (ชั่วโมงละช่อง 05:00–24:00) */
var OT_SLOTS = ['05:00-06:00','06:00-07:00','07:00-08:00','08:00-09:00','09:00-10:00','10:00-11:00','11:00-12:00','12:00-13:00',
  '13:00-14:00','14:00-15:00','15:00-16:00','16:00-17:00','17:00-18:00','18:00-19:00','19:00-20:00','20:00-21:00','21:00-22:00','22:00-23:00','23:00-24:00'];

var OT_ST = {WAIT:'รอยืนยัน', OK:'ยืนยัน', NO:'ไม่ยืนยัน', CANCEL:'ยกเลิก'};

var OT_DAY = {WORK:'วันทำงาน', OFF:'วันหยุด'};

var OT_HEAD = ['แจ้งเมื่อ','เลขโอที','รหัสพนักงาน','ชื่อ','สัญชาติ','วันที่ทำโอที','ช่วงเวลา (รายชั่วโมง)','เวลา','จำนวนชั่วโมง','ประเภทวัน','สถานะ','ผู้ยืนยัน','ยืนยันเมื่อ','ที่มา','หมายเหตุ','รหัสกลุ่ม','อุปกรณ์','บริษัท'];

var OT_NOTES = ['วันเวลาที่บันทึกรายการนี้','เลขประจำรายการโอที (ขึ้นต้น OT) — ใช้ตรวจสลิป','รหัสพนักงานที่ทำโอที','ชื่อพนักงาน ณ ตอนบันทึก','TH = ไทย, MM = พม่า',
  'วันที่ทำโอที (1 แถว = 1 คน 1 วัน)','ช่วงเวลาที่ทำ แยกเป็นชั่วโมง','ช่วงเวลาแบบย่อ เช่น 17:00–19:00','จำนวนชั่วโมงโอทีของแถวนี้',
  'วันทำงาน = โอทีนอกเวลา 08:00–17:00 ของวันทำงานปกติ • วันหยุด = ทำงานในวันอาทิตย์/วันหยุดบริษัท (ปกติจ่ายคนละอัตรา)',
  'รอยืนยัน = พนักงานแจ้งเอง ยังไม่มีหัวหน้ายืนยัน (ยังไม่นับ) • ยืนยัน = นับเป็นโอทีจริง • ไม่ยืนยัน = หัวหน้ากดว่าไม่ได้ทำ • ยกเลิก = พนักงานยกเลิกเอง',
  'ชื่อคนที่กดยืนยันใน Telegram หรือ "ผู้ดูแล" ถ้าลงจากหน้าผู้ดูแล','วันเวลาที่ยืนยัน','พนักงานแจ้งเอง หรือ ผู้ดูแลลงให้','หมายเหตุ (ถ้ามี)','รายการที่ผู้ดูแลลงพร้อมกันหลายคนจะมีรหัสกลุ่มเดียวกัน','เครื่อง/เบราว์เซอร์ที่ใช้แจ้ง',
  'บริษัทของพนักงาน ณ วันที่ทำโอที — ระบบใส่ให้เอง'];

/* ชีต OT นี้เป็นของระบบไหม: หัวตารางช่อง A1/B1/C1 ตรงสักช่อง หรือแถวแรกมีเลขโอที (กันเผลอแก้หัวตารางช่องเดียวแล้วระบบเข้าใจผิด) */
function otIsOurs_(sh){
  var h=sh.getRange(1,1,2,3).getValues();
  for(var c=0;c<3;c++) if(String(h[0][c]).trim()===OT_HEAD[c]) return true;
  return /^OT\d/.test(String(h[1][1]).trim());
}

function otSheet_(){
  var sh=sheet_(OT_SHEET);
  if(sh && sh.getLastRow()>0 && !otIsOurs_(sh)){ shelveSheet_(sh); sh=null; }   // มีชีตชื่อ OT ของผู้ใช้อยู่ก่อน → เปลี่ยนชื่อเก็บไว้ ไม่แตะข้อมูล
  if(sh) return sh;
  sh=ss_().insertSheet(OT_SHEET);
  otFormatSheet_(sh);
  return sh;
}

function otFormatSheet_(sh){
  ensureHeader_(sh, OT_HEAD, OT_NOTES);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1,1,150); sh.setColumnWidths(4,1,170); sh.setColumnWidths(6,1,110); sh.setColumnWidths(7,1,240); sh.setColumnWidths(8,1,130); sh.setColumnWidths(15,1,220);
  sh.getRange('F2:F').setNumberFormat('dd/mm/yyyy');
  sh.getRange('B2:C').setNumberFormat('@');
  sh.getRange('G2:H').setNumberFormat('@');
  try{
    var rng=sh.getRange('K2:K');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(OT_ST.OK).setBackground('#D1E7DD').setFontColor('#0F5132').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(OT_ST.WAIT).setBackground('#FFF3CD').setFontColor('#997404').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(OT_ST.NO).setBackground('#F8D7DA').setFontColor('#B02A37').setRanges([rng]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(OT_ST.CANCEL).setBackground('#E5E7EB').setFontColor('#6B7280').setRanges([rng]).build()
    ]);
  }catch(e){}
}

function otBackDays_(){ var n=parseInt(getSetting_('ot_back_days'),10); return (n>=0 && n<=62)?n:7; }

function otNeedConfirm_(){ return String(getSetting_('ot_need_confirm')).trim()!=='0'; }

/* ช่องเวลาที่ลงโอทีได้ของวันนั้น: วันทำงาน = นอกเวลางานปกติ • วันหยุด = ทุกช่อง */
function otAllowedSlots_(dateStr){
  if(!isWorkingDay_(dateStr)) return OT_SLOTS.slice();
  return OT_SLOTS.filter(function(s){ return TIME_SLOTS.indexOf(s)<0; });
}

function otCleanSlots_(slots){
  var a=(slots instanceof Array ? slots : String(slots||'').split(',')).map(function(s){ return String(s).trim(); });
  return a.filter(function(s,i,arr){ return OT_SLOTS.indexOf(s)>=0 && arr.indexOf(s)===i; })
          .sort(function(x,y){ return OT_SLOTS.indexOf(x)-OT_SLOTS.indexOf(y); });
}

/* 17:00-18:00, 18:00-19:00 → 17:00–19:00 */
function otCompact_(slots){
  var arr=otCleanSlots_(slots), out=[], st=null, en=null;
  arr.forEach(function(s){
    var p=s.split('-');
    if(en===p[0]) en=p[1]; else { if(st!==null) out.push(st+'–'+en); st=p[0]; en=p[1]; }
  });
  if(st!==null) out.push(st+'–'+en);
  return out.join(', ');
}

function otAll_(){
  var sh=sheet_(OT_SHEET); if(!sh) return [];
  var last=sh.getLastRow(); if(last<2) return [];
  var v=sh.getRange(2,1,last-1,18).getValues(), out=[];
  for(var i=0;i<v.length;i++){
    var id=String(v[i][1]).trim(); if(!id) continue;
    out.push({row:i+2, filed_at:fmtCell_(v[i][0]), ot_id:id, emp_id:String(v[i][2]).trim(), name:v[i][3], nationality:v[i][4],
      ot_date:fmtDate_(v[i][5]), slots:String(v[i][6]||''), range:String(v[i][7]||''), hours:Number(v[i][8])||0, day_type:String(v[i][9]||''),
      status:String(v[i][10]||''), approver:String(v[i][11]||''), approved_at:fmtCell_(v[i][12]), source:String(v[i][13]||''), note:String(v[i][14]||''), group_id:String(v[i][15]||''), company:String(v[i][17]||'').trim()});
  }
  return out;
}

function otLive_(o){ return o.status!==OT_ST.NO && o.status!==OT_ST.CANCEL; }

/* ช่องที่คนนี้ลงโอทีไว้แล้วในวันนั้น (ไม่นับรายการที่ไม่ยืนยัน/ยกเลิก) */
function otTakenMap_(all, dateStr){
  var m={};
  all.forEach(function(o){
    if(o.ot_date!==dateStr || !otLive_(o)) return;
    if(!m[o.emp_id]) m[o.emp_id]={};
    o.slots.split(',').forEach(function(s){ s=s.trim(); if(s) m[o.emp_id][s]=true; });
  });
  return m;
}

function otNewId_(prefix){ return prefix+Utilities.formatDate(new Date(),TZ,'yyMMddHHmmss')+Math.floor(Math.random()*900+100); }

function otRowValues_(emp, id, dateStr, slots, status, approver, source, note, groupId, device, at){
  return [at, id, emp.emp_id, emp.name_th||emp.name_mm, emp.nationality, parseDateStr_(dateStr), slots.join(', '), otCompact_(slots), slots.length,
    (isWorkingDay_(dateStr)?OT_DAY.WORK:OT_DAY.OFF), status, approver, (status===OT_ST.OK?at:''), source, safeText_(note||''), groupId, safeText_(device||''), companyOn_(emp, dateStr)];
}

function otAppend_(rows){
  var sh=otSheet_(), r0=sh.getLastRow()+1, need=r0+rows.length-1-sh.getMaxRows();
  if(need>0) sh.insertRowsAfter(sh.getMaxRows(), need+200);      // ชีตเต็ม → เพิ่มแถวก่อน (กันเขียนเกินขอบชีต)
  sh.getRange(r0,2,rows.length,2).setNumberFormat('@');
  sh.getRange(r0,6,rows.length,1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(r0,7,rows.length,2).setNumberFormat('@');
  sh.getRange(r0,1,rows.length,18).setValues(rows);
}

/* วันที่ต้องเป็นวันที่จริง (2026-02-31 ไม่ผ่าน) */
function otValidDate_(ds){ return /^\d{4}-\d{2}-\d{2}$/.test(ds) && dateToStr_(parseDateStr_(ds))===ds; }

/* ---------- พนักงานแจ้งโอทีเอง ---------- */
function submitOT_(d){
  var emp=findEmployee_(d.emp_id);
  if(!emp || !emp.active) return err_('NO_EMP','ไม่พบพนักงานคนนี้');
  var ds=String(d.ot_date||'').trim();
  if(!otValidDate_(ds)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง');
  var today=todayStr_();
  if(ds>today) return err_('OT_FUTURE','แจ้งโอทีล่วงหน้าไม่ได้ — แจ้งได้ในวันที่ทำหรือหลังจากทำแล้ว');
  var back=new Date(parseDateStr_(today)); back.setDate(back.getDate()-otBackDays_());
  var minDate=dateToStr_(back);
  if(ds<minDate) return err_('OT_TOO_OLD','แจ้งโอทีย้อนหลังได้ไม่เกิน '+otBackDays_()+' วัน (ย้อนได้ถึง '+fmtDMY_(minDate)+')',{min_date:minDate, days:otBackDays_()});
  var slots=otCleanSlots_(d.slots||[]);
  if(!slots.length) return err_('NO_SLOT','กรุณาเลือกช่วงเวลาที่ทำโอทีอย่างน้อย 1 ชั่วโมง');
  var allowed=otAllowedSlots_(ds), bad=slots.filter(function(s){ return allowed.indexOf(s)<0; });
  if(bad.length) return err_('OT_WORKTIME','ช่วงเวลานี้เป็นเวลางานปกติ ลงเป็นโอทีไม่ได้: '+otCompact_(bad),{slots:bad});
  var taken=otTakenMap_(otAll_(), ds)[emp.emp_id]||{}, dup=slots.filter(function(s){ return taken[s]; });
  if(dup.length) return err_('OT_DUP','ช่วงเวลานี้ลงโอทีไปแล้ว: '+otCompact_(dup),{slots:dup});

  var need=otNeedConfirm_(), status=need?OT_ST.WAIT:OT_ST.OK, at=nowStr_(), id=otNewId_('OT');
  otAppend_([otRowValues_(emp, id, ds, slots, status, (need?'':'อัตโนมัติ'), 'พนักงานแจ้งเอง', String(d.note||'').slice(0,300), id, String(d.device||'').slice(0,80), at)]);
  var res={ok:true, ot_id:id, slip:{id:id, code:slipCode_(id)}, emp_id:emp.emp_id, name:emp.name_th||emp.name_mm, ot_date:ds, slots:slots,
    range:otCompact_(slots), hours:slots.length, day_type:(isWorkingDay_(ds)?OT_DAY.WORK:OT_DAY.OFF), status:status, filed_at:at};
  notifyOT_(emp, res, String(d.note||''));
  return res;
}

/* พนักงานยกเลิกรายการที่ยังรอยืนยัน */
function cancelOT_(d){
  var empId=String(d.emp_id||'').trim(), id=String(d.ot_id||'').trim();
  if(!empId || !id) return err_('MISSING','ข้อมูลไม่ครบ');
  var o=otAll_().filter(function(x){ return x.ot_id===id; })[0];
  if(!o) return err_('NOT_FOUND','ไม่พบรายการโอทีนี้');
  if(o.emp_id!==empId) return err_('NOT_YOURS','รายการนี้ไม่ใช่ของคุณ');
  if(o.status!==OT_ST.WAIT) return err_('OT_LOCKED','รายการนี้ '+o.status+' แล้ว ยกเลิกเองไม่ได้ — ติดต่อผู้ดูแล',{status:o.status});
  otSheet_().getRange(o.row,11,1,3).setValues([[OT_ST.CANCEL, 'พนักงานยกเลิกเอง', nowStr_()]]);
  tgSend_('🗑 <b>ยกเลิกรายการโอที</b>\n👤 '+tgEscape_(o.name)+' ('+tgEscape_(o.emp_id)+')\n📅 '+fmtDMY_(o.ot_date)+' · 🕔 '+tgEscape_(o.range)+' ('+o.hours+' ชม.)');
  return {ok:true};
}

/* ---------- ผู้ดูแลลงโอทีให้หลายคนพร้อมกัน ---------- */
function adminAddOT_(d){
  var ds=String(d.ot_date||'').trim();
  if(!otValidDate_(ds)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง');
  var today=todayStr_();
  if(ds>today) return err_('OT_FUTURE','ลงโอทีล่วงหน้าไม่ได้ — ลงได้ในวันที่ทำหรือหลังจากทำแล้ว');
  var lim=new Date(parseDateStr_(today)); lim.setDate(lim.getDate()-62);
  if(ds<dateToStr_(lim)) return err_('OT_TOO_OLD','ลงย้อนหลังได้ไม่เกิน 62 วัน');
  var slots=otCleanSlots_(d.slots||[]);
  if(!slots.length) return err_('NO_SLOT','กรุณาเลือกช่วงเวลาอย่างน้อย 1 ชั่วโมง');
  var allowed=otAllowedSlots_(ds), bad=slots.filter(function(s){ return allowed.indexOf(s)<0; });
  if(bad.length) return err_('OT_WORKTIME','ช่วงเวลานี้เป็นเวลางานปกติ ลงเป็นโอทีไม่ได้: '+otCompact_(bad),{slots:bad});
  var ids=(d.emp_ids||[]).map(function(x){ return String(x).trim(); }).filter(function(x,i,a){ return x && a.indexOf(x)===i; });
  if(!ids.length) return err_('NO_EMP','กรุณาเลือกพนักงานอย่างน้อย 1 คน');
  if(ids.length>300) return err_('TOO_MANY','ลงได้ครั้งละไม่เกิน 300 คน');
  var emps=empMap_(), taken=otTakenMap_(otAll_(), ds), at=nowStr_(), gid=otNewId_('OG'), rows=[], added=[], skipped=[];
  var note=String(d.note||'').slice(0,300);
  ids.forEach(function(id, k){
    var e=emps[id];
    if(!e || !e.active){ skipped.push({emp_id:id, name:'', reason:'ไม่พบพนักงาน', code:'NO_EMP'}); return; }
    var t=taken[id]||{}, dup=slots.filter(function(s){ return t[s]; });
    if(dup.length){ skipped.push({emp_id:id, name:e.name_th||e.name_mm, reason:'ลงไว้แล้ว '+otCompact_(dup), code:'DUP'}); return; }
    var oid=gid.replace('OG','OT')+'-'+(k+1);
    rows.push(otRowValues_(e, oid, ds, slots, OT_ST.OK, 'ผู้ดูแล', 'ผู้ดูแลลงให้', note, gid, '', at));
    added.push({emp_id:id, name:e.name_th||e.name_mm, ot_id:oid});
  });
  if(rows.length){
    otAppend_(rows);
    var names=added.map(function(a){ return tgEscape_(a.name); });
    tgSend_('⏱ <b>ลงโอที '+added.length+' คน</b> (ผู้ดูแลลงให้)\n📅 '+SUMMARY_WD_FULL[parseDateStr_(ds).getDay()]+' '+fmtDMY_(ds)+' · '+(isWorkingDay_(ds)?OT_DAY.WORK:OT_DAY.OFF)+
      '\n🕔 '+tgEscape_(otCompact_(slots))+' ('+slots.length+' ชม./คน • รวม '+(slots.length*added.length)+' ชม.)'+(note?('\n📝 '+tgEscape_(note)):'')+
      '\n👥 '+(names.length>40?(names.slice(0,40).join(', ')+' … และอีก '+(names.length-40)+' คน'):names.join(', ')));
  }
  return {ok:true, added:added.length, added_list:added, skipped:skipped, hours_each:slots.length, range:otCompact_(slots)};
}

function otSetStatus_(o, status, who){
  otSheet_().getRange(o.row,11,1,3).setValues([[status, safeText_(who||'ผู้ดูแล'), (status===OT_ST.WAIT?'':nowStr_())]]);
}

/* เปลี่ยนสถานะ (ผู้ดูแล) — d.expect = สถานะที่หน้าจอเห็นอยู่ (ถ้ามีคนเปลี่ยนไปก่อนแล้ว จะข้าม ไม่ทับ)
 * รายการที่พนักงานยกเลิกเองแล้ว เปลี่ยนกลับไม่ได้ • จะกลับมา "นับ" ได้ต่อเมื่อช่วงเวลาไม่ชนกับรายการอื่นของคนเดิม */
function updateOTStatus_(d){
  var st=String(d.status||'').trim(), expect=String(d.expect||'').trim();
  if([OT_ST.OK, OT_ST.NO, OT_ST.WAIT].indexOf(st)<0) return err_('BAD_STATUS','สถานะไม่ถูกต้อง');
  var ids=(d.ot_ids instanceof Array ? d.ot_ids : [d.ot_id]).map(function(x){ return String(x||'').trim(); }).filter(function(x){ return x; });
  if(!ids.length) return err_('MISSING','ข้อมูลไม่ครบ');
  var all=otAll_(), n=0, skipped=[];
  ids.forEach(function(id){
    var o=all.filter(function(x){ return x.ot_id===id; })[0];
    if(!o){ skipped.push({ot_id:id, reason:'ไม่พบรายการ'}); return; }
    if(o.status===st) return;
    if(o.status===OT_ST.CANCEL){ skipped.push({ot_id:id, reason:'พนักงานยกเลิกรายการนี้แล้ว'}); return; }
    if(expect && o.status!==expect){ skipped.push({ot_id:id, reason:'สถานะเปลี่ยนเป็น "'+o.status+'" ไปแล้ว'}); return; }
    if(!otLive_(o) && st!==OT_ST.NO){            // จะกลับมานับ → ต้องไม่ชนกับรายการอื่นที่ยังนับอยู่
      var mine=o.slots.split(',').map(function(x){ return x.trim(); });
      var clash=all.some(function(x){ return x.ot_id!==o.ot_id && x.emp_id===o.emp_id && x.ot_date===o.ot_date && otLive_(x) &&
        x.slots.split(',').some(function(sl){ return mine.indexOf(sl.trim())>=0; }); });
      if(clash){ skipped.push({ot_id:id, reason:'ช่วงเวลาชนกับรายการโอทีอื่นของคนเดียวกัน'}); return; }
    }
    otSetStatus_(o, st, 'ผู้ดูแล'); o.status=st; n++;
  });
  if(!n && ids.length===1 && skipped.length && skipped[0].reason==='ไม่พบรายการ') return err_('NOT_FOUND','ไม่พบรายการโอทีนี้');
  return {ok:true, changed:n, skipped:skipped};
}

function deleteOT_(d){
  var id=String(d.ot_id||'').trim();
  var o=otAll_().filter(function(x){ return x.ot_id===id; })[0];
  if(!o) return err_('NOT_FOUND','ไม่พบรายการโอทีนี้');
  otSheet_().deleteRow(o.row);
  return {ok:true};
}

/* รวมชั่วโมงโอทีของรายการชุดหนึ่ง */
function otTotals_(list){
  var t={work:0, off:0, total:0, pending:0, pending_count:0};
  list.forEach(function(o){
    if(o.status===OT_ST.OK){ if(o.day_type===OT_DAY.OFF) t.off+=o.hours; else t.work+=o.hours; t.total+=o.hours; }
    else if(o.status===OT_ST.WAIT){ t.pending+=o.hours; t.pending_count++; }
  });
  return t;
}

function otCurrentCycle_(){
  var now=new Date(), co=getCutoffDay_(APP.CO);
  var day=parseInt(Utilities.formatDate(now,TZ,'d'),10), y=parseInt(Utilities.formatDate(now,TZ,'yyyy'),10), m=parseInt(Utilities.formatDate(now,TZ,'M'),10)-1;
  return cycleForMonth_((day<=co)?new Date(y,m,1):new Date(y,m+1,1), APP.CO);
}

/* หน้าผู้ดูแล: รายการโอทีของเดือน/รอบ + ยอดรวมรายคน */
function getOTAdmin_(p){
  if(!(p && checkPin_(p.pin))) return err_('PIN','PIN ไม่ถูกต้อง');
  var rg=adminRange_(p);
  var list=otAll_().filter(function(o){ return o.ot_date>=rg.from && o.ot_date<=rg.to; });
  list.sort(function(a,b){ return a.ot_date<b.ot_date?1:(a.ot_date>b.ot_date?-1:(a.filed_at<b.filed_at?1:-1)); });
  var by={};
  list.forEach(function(o){
    if(!by[o.emp_id]) by[o.emp_id]={emp_id:o.emp_id, name:o.name, nationality:o.nationality, list:[]};
    by[o.emp_id].list.push(o);
  });
  var people=Object.keys(by).map(function(k){ var t=otTotals_(by[k].list); return {emp_id:k, name:by[k].name, nationality:by[k].nationality, work:t.work, off:t.off, total:t.total, pending:t.pending}; })
    .sort(function(a,b){ return b.total-a.total || (a.emp_id<b.emp_id?-1:1); });
  return {ok:true, ots:list, people:people, totals:otTotals_(list), range:rg, slots:OT_SLOTS, work_slots:TIME_SLOTS};
}

/* โอทีของพนักงาน 1 คน (แสดงในหน้าประวัติ) */
function otForEmp_(empId){
  var cyc=otCurrentCycle_(), mine=otAll_().filter(function(o){ return o.emp_id===empId; });
  mine.sort(function(a,b){ return a.ot_date<b.ot_date?1:(a.ot_date>b.ot_date?-1:(a.filed_at<b.filed_at?1:-1)); });
  var inCyc=mine.filter(function(o){ return o.ot_date>=cyc.startStr && o.ot_date<=cyc.endStr; });
  return {list:mine.slice(0,60).map(function(o){ return {ot_id:o.ot_id, filed_at:o.filed_at, ot_date:o.ot_date, slots:o.slots, range:o.range, hours:o.hours, day_type:o.day_type,
            status:o.status, approver:o.approver, source:o.source, note:o.note, company:(o.company?normCompany_(o.company):''), slip_code:slipCode_(o.ot_id), can_cancel:(o.status===OT_ST.WAIT)}; }),
          cycle:{label:cyc.label, from:cyc.startStr, to:cyc.endStr}, totals:otTotals_(inCyc)};
}

/* ค่าที่หน้าเว็บต้องใช้ตอนเปิดฟอร์มโอที: ช่องเวลาที่ลงได้ + ช่องที่ลงไปแล้วของวันนั้น */
function getOTDay_(p){
  var ds=String(p.date||'').trim();
  if(!otValidDate_(ds)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง');
  var back=new Date(parseDateStr_(todayStr_())); back.setDate(back.getDate()-otBackDays_());
  var out={ok:true, date:ds, today:todayStr_(), min_date:dateToStr_(back), working_day:isWorkingDay_(ds), allowed:otAllowedSlots_(ds), slots:OT_SLOTS, need_confirm:otNeedConfirm_()};
  var empId=String(p.emp_id||'').trim();
  if(empId){ var t=otTakenMap_(otAll_(), ds)[empId]||{}; out.taken=Object.keys(t); }
  if(p.all==='1' && checkPin_(p.pin)){ var tm=otTakenMap_(otAll_(), ds), o2={}; Object.keys(tm).forEach(function(k){ o2[k]=Object.keys(tm[k]); }); out.taken_all=o2; }
  return out;
}

/* ---------- Telegram ---------- */
function otCardText_(emp, o, note, head){
  var name=emp.name_th||emp.name_mm;
  if(emp.name_mm && emp.nationality==='MM' && emp.name_th) name=emp.name_th+' / '+emp.name_mm;
  return (head||'⏱ <b>แจ้งโอที (รอยืนยัน)</b>')+'\n👤 '+tgEscape_(name)+' ('+tgEscape_(emp.emp_id)+')\n📅 '+SUMMARY_WD_FULL[parseDateStr_(o.ot_date).getDay()]+' '+fmtDMY_(o.ot_date)+' · '+tgEscape_(o.day_type)+
    '\n🕔 '+tgEscape_(o.range)+' ('+o.hours+' ชม.)'+(note?('\n📝 '+tgReasonValue_(note)):'');
}

function otButtons_(id){ return {inline_keyboard:[[{text:'✅ ยืนยัน', callback_data:'oa:'+id},{text:'❌ ไม่ได้ทำ', callback_data:'or:'+id}]]}; }

function notifyOT_(emp, o, note){
  var chatId=String(getSetting_('telegram_chat_id')||'').trim();
  if(!getSetting_('telegram_bot_token') || !chatId) return;
  if(o.status===OT_ST.WAIT) tgSendTo_(chatId, otCardText_(emp, o, note)+'\n\nกดปุ่มด้านล่างเพื่อยืนยันว่าทำจริง', otButtons_(o.ot_id));
  else tgSendTo_(chatId, otCardText_(emp, o, note, '⏱ <b>แจ้งโอที</b>'));
}

function tgOtCallback_(cq, act, id, clicker, chatId, msgId){
  var lock=LockService.getScriptLock(), got=false;
  try{ got=lock.tryLock(15000); }catch(e){}
  if(!got){ tgAnswerCallback_(cq.id,'ระบบกำลังยุ่ง กดอีกครั้ง'); return json_({ok:true}); }   // ไม่ได้ล็อก = ไม่เขียน (กันเขียนผิดแถว)
  try{
    var o=otAll_().filter(function(x){ return x.ot_id===id; })[0];
    if(!o){ tgAnswerCallback_(cq.id,'ไม่พบรายการนี้ (อาจถูกลบแล้ว)'); return json_({ok:true}); }
    if(o.status!==OT_ST.WAIT){ tgAnswerCallback_(cq.id,'ดำเนินการไปแล้ว: '+o.status+(o.approver?(' โดย '+o.approver):'')); return json_({ok:true}); }
    var st=(act==='oa')?OT_ST.OK:OT_ST.NO, icon=(act==='oa')?'✅':'❌', at=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm');
    otSetStatus_(o, st, clicker);
    var orig=tgEscape_((cq.message&&cq.message.text)||'').replace(/\n\nกดปุ่มด้านล่าง[^\n]*$/,'');
    tgEditMessage_(chatId, msgId, orig+'\n\n'+icon+' <b>'+st+'</b> โดย '+tgEscape_(clicker)+'\n🕒 '+at);
    tgAnswerCallback_(cq.id, icon+' '+st+'แล้ว');
    var groupChat=String(getSetting_('telegram_chat_id')||'').trim();
    if(groupChat && String(chatId)!==groupChat) tgSend_(icon+' <b>'+st+'โอที</b>\n👤 '+tgEscape_(o.name)+'\n📅 '+fmtDMY_(o.ot_date)+' · 🕔 '+tgEscape_(o.range)+'\nโดย '+tgEscape_(clicker));
  } finally { if(got){ try{ lock.releaseLock(); }catch(e){} } }
  return json_({ok:true});
}

/* บรรทัดโอทีในสรุปรายวัน: รอบเช้า = โอทีเมื่อวาน • รอบเย็น = โอทีวันนี้ที่ลงแล้ว */
function otSummaryBlock_(isToday){
  var all=otAll_(), ds=todayStr_(), label='วันนี้ (ที่ลงแล้ว)';
  if(isToday){ var y=new Date(parseDateStr_(ds)); y.setDate(y.getDate()-1); ds=dateToStr_(y); label='เมื่อวาน'; }
  var day=all.filter(function(o){ return o.ot_date===ds && o.status===OT_ST.OK; });
  var wait=all.filter(function(o){ return o.status===OT_ST.WAIT; });
  if(!day.length && !wait.length) return '';
  var out='\n\n⏱ <b>โอที'+label+'</b> ';
  if(day.length){
    var people={}; day.forEach(function(o){ people[o.emp_id]=(people[o.emp_id]||0)+o.hours; });
    out+=Object.keys(people).length+' คน • รวม '+otTotals_(day).total+' ชม.';
  } else out+='– ไม่มี –';
  if(wait.length) out+='\n⏳ โอทีรอยืนยัน '+wait.length+' รายการ (กดปุ่ม "รออนุมัติ" เพื่อดู)';
  return out;
}

/* ปุ่ม ⏱ โอที : โอทีวันนี้/เมื่อวาน + ยอดรอบเงินเดือนนี้ */
function tgOtText_(){
  var all=otAll_(), today=todayStr_(), y=new Date(parseDateStr_(today)); y.setDate(y.getDate()-1);
  var yd=dateToStr_(y), cyc=otCurrentCycle_();
  function dayBlock(ds, title){
    var l=all.filter(function(o){ return o.ot_date===ds && otLive_(o); });
    if(!l.length) return '\n\n<b>'+title+' '+fmtDMY_(ds).slice(0,5)+'</b> — ไม่มี';
    var s='\n\n<b>'+title+' '+fmtDMY_(ds).slice(0,5)+'</b> — '+l.length+' รายการ • '+l.reduce(function(a,o){ return a+o.hours; },0)+' ชม.';
    l.slice(0,40).forEach(function(o){ s+='\n• '+summaryFlag_(o.nationality)+' '+tgEscape_(o.name)+' — '+tgEscape_(o.range)+' ('+o.hours+' ชม.)'+(o.status===OT_ST.WAIT?' ⏳':''); });
    if(l.length>40) s+='\n… และอีก '+(l.length-40)+' รายการ';
    return s;
  }
  var inCyc=all.filter(function(o){ return o.ot_date>=cyc.startStr && o.ot_date<=cyc.endStr; }), t=otTotals_(inCyc);
  var by={}; inCyc.forEach(function(o){ if(o.status!==OT_ST.OK) return; if(!by[o.emp_id]) by[o.emp_id]={name:o.name, nat:o.nationality, h:0}; by[o.emp_id].h+=o.hours; });
  var top=Object.keys(by).map(function(k){ return by[k]; }).sort(function(a,b){ return b.h-a.h; });
  var out='⏱ <b>โอที</b>'+dayBlock(today,'วันนี้')+dayBlock(yd,'เมื่อวาน')+
    '\n\n📊 <b>'+cyc.label+'</b>\nโอทีวันทำงาน <b>'+t.work+'</b> ชม. • ทำงานวันหยุด <b>'+t.off+'</b> ชม. • รวม <b>'+t.total+'</b> ชม. ('+top.length+' คน)'+
    (t.pending_count?('\n⏳ รอยืนยัน '+t.pending_count+' รายการ ('+t.pending+' ชม.)'):'');
  if(top.length){ out+='\n\n<b>ทำโอทีมากสุด</b>'; top.slice(0,10).forEach(function(p,k){ out+='\n'+(k+1)+'. '+summaryFlag_(p.nat)+' '+tgEscape_(p.name)+' — '+p.h+' ชม.'; }); }
  return out;
}

/* เงื่อนไขบริษัทในสูตรโอที: ระบบบริษัทเดียว = นับทุกแถว • หลายกลุ่ม = แยกตามคอลัมน์ R ของชีต OT (ว่าง = บริษัทหลัก) */
function otCoCrit_(company){ return APP.MULTI ? ((company==='MAIN')?'"MAIN"':'"<>MAIN"') : '"<>-"'; }

/* บริษัทของรายการโอที (แถวเก่าที่ช่องบริษัทว่าง → ดูจากพนักงาน ณ วันนั้น) */
function otCompanyOf_(o, emps){ return o.company ? normCompany_(o.company) : companyOn_(emps[o.emp_id], o.ot_date); }

/* ---------- ชีต Payroll: เพิ่มคอลัมน์โอที (M–P) ---------- */
function rebuildPayroll_(name, label, rows, company){
  otSheet_();                                    // ต้องมีชีต OT ก่อน สูตรจึงอ้างถึงได้
  sumSheet_();                                   // ช่องเลือกรอบอยู่ที่ชีต "สรุป"
  rebuildPayrollCore_(name, label, rows, company);
  var sh=sheet_(name), CO=getCutoffDay_(company), n=rows.length, coCrit=otCoCrit_(company);
  var head=['โอทีวันทำงาน (ชม.)','ทำงานวันหยุด (ชม.)','รวมโอที (ชม.)','โอทีรอยืนยัน (ชม.)'];
  var notes=['โอทีนอกเวลา 08:00–17:00 ของวันทำงานปกติ ในรอบนี้ (เฉพาะที่ "ยืนยัน" แล้ว)','ชั่วโมงทำงานในวันอาทิตย์/วันหยุดบริษัท ในรอบนี้ (เฉพาะที่ "ยืนยัน" แล้ว)',
    'โอทีวันทำงาน + ทำงานวันหยุด','พนักงานแจ้งเองแต่ยังไม่มีคนยืนยัน — ยังไม่นับ (ให้หัวหน้ากดยืนยันใน Telegram หรือหน้าผู้ดูแล)'];
  sh.getRange(2,13,1,4).setValues([head]).setNotes([notes])
    .setFontWeight('bold').setBackground('#5A9E1B').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setColumnWidths(13,4,105);
  if(!n){ payrollPolish_(name, 0, label); return; }
  var lo='DATE(YEAR($B$1),MONTH($B$1)-1,'+(CO+1)+')', hi='DATE(YEAR($B$1),MONTH($B$1),'+CO+')', O="'"+OT_SHEET+"'";
  function f(rn, dayType, status){
    return '=IFERROR(SUMIFS('+O+'!$I:$I,'+O+'!$C:$C,$A'+rn+','+O+'!$F:$F,">="&'+lo+','+O+'!$F:$F,"<="&'+hi+
      (dayType?(','+O+'!$J:$J,"'+dayType+'"'):'')+','+O+'!$K:$K,"'+status+'",'+O+'!$R:$R,'+coCrit+'),0)';
  }
  var out=[];
  for(var r=0;r<n;r++){ var rn=r+3; out.push([f(rn,OT_DAY.WORK,OT_ST.OK), f(rn,OT_DAY.OFF,OT_ST.OK), '=$M'+rn+'+$N'+rn, f(rn,'',OT_ST.WAIT)]); }
  sh.getRange(3,13,n,4).setValues(out).setNumberFormat('#,##0.##');
  sh.getRange(3,13,n,3).setBackground('#F1F8E9');
  var tr=n+3, sums=['M','N','O','P'].map(function(c){ return '=SUM('+c+'3:'+c+(n+2)+')'; });
  sh.getRange(tr,13,1,4).setFormulas([sums]).setFontWeight('bold').setBackground('#EEF4F9').setNumberFormat('#,##0.##');
  payrollPolish_(name, n, label);
}

/* ---------- ไฟล์ส่งออก: เพิ่มชีต "สรุปโอที" + "รายการโอที" ---------- */
function exportOtInto_(nss, cyc, company){
  var emps=empMap_(), per={};
  company=normCompany_(company);
  var list=otAll_().filter(function(o){ return o.ot_date>=cyc.startStr && o.ot_date<=cyc.endStr && o.status!==OT_ST.CANCEL && otCompanyOf_(o, emps)===company; });
  Object.keys(emps).forEach(function(id){ var e=emps[id]; if(e.active && e.company===company) per[id]={id:id, name:e.name_th||e.name_mm, nat:e.nationality, l:[]}; });
  list.forEach(function(o){ if(!per[o.emp_id]) per[o.emp_id]={id:o.emp_id, name:o.name, nat:o.nationality, l:[]}; per[o.emp_id].l.push(o); });
  var sum=Object.keys(per).map(function(k){ return per[k]; }).sort(function(a,b){ return a.nat===b.nat?(a.id<b.id?-1:1):(a.nat==='TH'?-1:1); })
    .map(function(p){ var t=otTotals_(p.l); return [p.id, safeText_(p.name), (p.nat==='MM'?'พม่า':'ไทย'), t.work, t.off, t.total, t.pending]; });
  var s=nss.insertSheet('สรุปโอที');
  s.getRange(1,1).setValue('สรุปโอที '+coLabel_(company)+' — '+cyc.label+' ('+fmtDMY_(cyc.startStr)+' – '+fmtDMY_(cyc.endStr)+')').setFontWeight('bold').setFontSize(13);
  var h=['รหัสพนักงาน','ชื่อ','สัญชาติ','โอทีวันทำงาน (ชม.)','ทำงานวันหยุด (ชม.)','รวมโอที (ชม.)','รอยืนยัน (ชม.) — ยังไม่นับ'];
  s.getRange(3,1,1,h.length).setValues([h]).setFontWeight('bold').setBackground('#5A9E1B').setFontColor('#ffffff').setWrap(true);
  if(sum.length){ s.getRange(4,1,sum.length,1).setNumberFormat('@'); s.getRange(4,1,sum.length,h.length).setValues(sum); }
  s.setFrozenRows(3); s.setColumnWidth(2,200);
  var d=nss.insertSheet('รายการโอที');
  var h2=['วันที่ทำโอที','รหัสพนักงาน','ชื่อ','สัญชาติ','เวลา','ชั่วโมง','ประเภทวัน','สถานะ','ผู้ยืนยัน','ที่มา','หมายเหตุ','แจ้งเมื่อ'];
  d.getRange(1,1,1,h2.length).setValues([h2]).setFontWeight('bold').setBackground('#5A9E1B').setFontColor('#ffffff');
  var rows=list.slice().sort(function(a,b){ return a.ot_date<b.ot_date?-1:(a.ot_date>b.ot_date?1:(a.emp_id<b.emp_id?-1:1)); })
    .map(function(o){ return [parseDateStr_(o.ot_date), o.emp_id, safeText_(o.name), o.nationality, o.range, o.hours, o.day_type, o.status, safeText_(o.approver), o.source, safeText_(o.note), o.filed_at]; });
  if(rows.length){ d.getRange(2,2,rows.length,1).setNumberFormat('@'); d.getRange(2,5,rows.length,1).setNumberFormat('@'); d.getRange(2,1,rows.length,h2.length).setValues(rows); d.getRange(2,1,rows.length,1).setNumberFormat('dd/mm/yyyy'); }
  d.setFrozenRows(1); d.setColumnWidth(3,180); d.setColumnWidth(5,150);
  return {ot_rows:rows.length};
}

/* ---------- ต่อเข้ากับของเดิม ---------- */
function getHistory_(p){
  var r=getHistoryCore_(p);
  if(r && r.ok) r.ot=otForEmp_(String(p.emp_id||'').trim());
  return r;
}

function verifySlip_(p){
  var r=verifySlipAny_(p);
  if(r && r.ok && r.found && r.company) r.company_name=coLabel_(normCompany_(r.company));     // ชื่อกลุ่มที่แสดง (รหัสในข้อมูลอาจต่างจากชื่อ)
  return r;
}

function verifySlipAny_(p){
  var id=String(p.id||'').trim();
  if(id.indexOf('OT')!==0) return verifySlipCore_(p);
  var code=String(p.c||'').trim().toUpperCase();
  if(!code || code!==slipCode_(id)) return err_('BAD_CODE','รหัสตรวจสอบไม่ถูกต้อง — สลิปนี้ไม่ได้ออกจากระบบ');
  var o=otAll_().filter(function(x){ return x.ot_id===id; })[0];
  if(!o) return {ok:true, found:false, kind:'ot', checked_at:nowStr_()};
  return {ok:true, found:true, kind:'ot', emp_id:o.emp_id, name:o.name, company:otCompanyOf_(o, empMap_()), checked_at:nowStr_(),
    items:[{ot_id:o.ot_id, ot_date:o.ot_date, range:o.range, slots:o.slots, hours:o.hours, day_type:o.day_type, status:o.status, approver:o.approver, filed_at:o.filed_at}]};
}

function getConfig_(){
  var c=getConfigCore_();
  c.ot={slots:OT_SLOTS, back_days:otBackDays_(), need_confirm:otNeedConfirm_()};
  if(APP.MULTI) c.company2_name=coLabel_('MAIN');
  return c;
}

function tgMonthText_(){
  var cyc=otCurrentCycle_(), t=otTotals_(otAll_().filter(function(o){ return o.ot_date>=cyc.startStr && o.ot_date<=cyc.endStr; }));
  return tgMonthTextCore_()+'\n\n⏱ <b>โอทีรอบนี้</b>\nวันทำงาน '+t.work+' ชม. • วันหยุด '+t.off+' ชม. • รวม <b>'+t.total+'</b> ชม.'+(t.pending_count?('\n⏳ รอยืนยัน '+t.pending_count+' รายการ'):'');
}

function tgEmpCard_(e){
  var cyc=otCurrentCycle_(), mine=otAll_().filter(function(o){ return o.emp_id===e.emp_id && o.ot_date>=cyc.startStr && o.ot_date<=cyc.endStr; }), t=otTotals_(mine);
  return tgEmpCardCore_(e)+'\n\n<b>⏱ โอที '+cyc.label+'</b>\nวันทำงาน '+t.work+' ชม. • วันหยุด '+t.off+' ชม. • รวม '+t.total+' ชม.'+(t.pending_count?(' • รอยืนยัน '+t.pending+' ชม.'):'');
}

function getAdminSettings_(p){
  var r=getAdminSettingsCore_(p);
  if(r && r.ok){ r.ot_back_days=otBackDays_(); r.ot_need_confirm=otNeedConfirm_(); r.pin_default=DEFAULT_PIN; if(APP.MULTI) r.company2_name=coLabel_('MAIN'); }
  return r;
}

/* ==================================================================
 *  MAIN — ปฏิทินบริษัท: สลับวันหยุด / เพิ่มวันทำงาน / เพิ่มวันหยุด
 *  ทำได้จากหน้าผู้ดูแลในเว็บ และจาก Telegram (ทุกคนในกลุ่ม)
 *  เก็บข้อมูลที่เดิม: ชีต WorkExtra (วันหยุดที่ต้องมาทำงาน) + ชีต Holidays (วันที่หยุด)
 *  การสลับ = เพิ่ม 2 รายการที่จับคู่กัน โดยเขียน "(สลับกับ dd/mm/yyyy)" ไว้ในชื่อ → ลบรายการใดรายการหนึ่ง อีกรายการหายด้วย
 * ================================================================== */
var CAL_SWAP_RE = /\(สลับกับ (\d{2})\/(\d{2})\/(\d{4})\)/;

function calPairOf_(name){ var m=CAL_SWAP_RE.exec(String(name||'')); return m ? (m[3]+'-'+m[2]+'-'+m[1]) : ''; }

function calLabel_(ds){ return SUMMARY_WD_FULL[parseDateStr_(ds).getDay()]+' '+fmtDMY_(ds); }

function calShort_(ds){ return DAILY_THAI_WD[parseDateStr_(ds).getDay()]+' '+fmtDMY_(ds).slice(0,5); }

function calC_(ds){ return String(ds).replace(/-/g,''); }

function calD_(c){ c=String(c||''); return /^\d{8}$/.test(c) ? (c.slice(0,4)+'-'+c.slice(4,6)+'-'+c.slice(6,8)) : ''; }

function calShift_(ds, n){ var d=parseDateStr_(ds); d.setDate(d.getDate()+n); return dateToStr_(d); }

/* วันที่ใช้ได้: วันที่จริง และอยู่ในช่วง ย้อนหลัง 31 วัน ถึง ล่วงหน้า 400 วัน */
function calDateOk_(ds){
  if(!otValidDate_(ds)) return false;
  var t=todayStr_();
  return ds>=calShift_(t,-31) && ds<=calShift_(t,400);
}

/* จำนวนคนที่แจ้งลาไว้ในวันนั้น (ไม่นับใบที่ไม่อนุมัติ) */
function calLeavesOn_(ds){
  var vals=sheet_(SHEETS.LEAVE).getDataRange().getValues(), seen={}, n=0;
  for(var i=1;i<vals.length;i++){
    if(fmtDate_(vals[i][5])!==ds || String(vals[i][12])==='ไม่อนุมัติ') continue;
    var id=String(vals[i][2]).trim(); if(!seen[id]){ seen[id]=true; n++; }
  }
  return n;
}

/* ปฏิทินเปลี่ยน → ปรับ "ประเภทวัน" ของรายการโอทีวันนั้นให้ตรง (วันทำงาน/วันหยุด) */
function otRetypeDate_(ds){
  var sh=sheet_(OT_SHEET); if(!sh) return 0;
  var last=sh.getLastRow(); if(last<2) return 0;
  var v=sh.getRange(2,6,last-1,5).getValues(), want=(isWorkingDay_(ds)?OT_DAY.WORK:OT_DAY.OFF), n=0;
  for(var i=0;i<v.length;i++){
    if(fmtDate_(v[i][0])===ds && String(v[i][4])!==want){ sh.getRange(i+2,10).setValue(want); n++; }
  }
  return n;
}

function calNameOf_(kind, ds){
  var h=getHolidays_(), list=(kind==='W' ? (h.work_overrides||[]) : (h.holidays||[]));
  for(var i=0;i<list.length;i++) if(list[i].date===ds) return String(list[i].name||'');
  return null;
}

/* ds เป็น "วันทำงานพิเศษ" ที่เป็นครึ่งหนึ่งของคู่สลับหรือไม่ → คืนวันที่คู่ (dd/mm/yyyy→yyyy-mm-dd) หรือ '' */
function calPairedWork_(ds){ var nm=calNameOf_('W',ds); return nm===null?'':(calPairOf_(nm)||''); }

/* อ่านปฏิทินครั้งเดียวต่อการทำงาน (ใช้เฉพาะตอนวาดปุ่มเลือกวัน) */
var _calDay=null;

function calIsWork_(ds){
  try{
    if(!_calDay){ _calDay={h:{}, w:{}};
      [[SHEETS.HOLIDAY,'h'],[SHEETS.WORKDAY,'w']].forEach(function(p){ var sh=sheet_(p[0]); if(!sh) return; var v=sh.getDataRange().getValues();
        for(var i=1;i<v.length;i++){ var k=dateKeyOf_(v[i][0]); if(k) _calDay[p[1]][k]=true; } }); }
    return _calDay.w[ds] ? true : (_calDay.h[ds] ? false : parseDateStr_(ds).getDay()!==0);
  }catch(e){ return isWorkingDay_(ds); }
}

/* สลับวันหยุด: workDs (ตอนนี้เป็นวันหยุด) → มาทำงาน • offDs (ตอนนี้เป็นวันทำงาน) → หยุดแทน */
function calSwap_(workDs, offDs){
  if(!calDateOk_(workDs) || !calDateOk_(offDs)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง (เลือกได้ย้อนหลัง 31 วัน ถึงล่วงหน้า 400 วัน)');
  if(workDs===offDs) return err_('CAL_SAME','สองวันต้องไม่ใช่วันเดียวกัน');
  if(isWorkingDay_(workDs)) return err_('CAL_IS_WORK', calLabel_(workDs)+' เป็นวันทำงานอยู่แล้ว — ช่องแรกต้องเป็นวันหยุด (วันอาทิตย์/วันหยุดบริษัท)');
  if(!isWorkingDay_(offDs)) return err_('CAL_IS_OFF', calLabel_(offDs)+' เป็นวันหยุดอยู่แล้ว — ช่องที่สองต้องเป็นวันทำงาน');
  var pw=calPairedWork_(offDs); if(pw) return err_('CAL_PAIRED', calLabel_(offDs)+' เป็นวันทำงานจากการสลับกับ '+fmtDMY_(pw)+' — ยกเลิกรายการสลับเดิมก่อน');
  var r=addWorkOverride_({date:workDs, name:'มาทำงานแทนวันหยุด (สลับกับ '+fmtDMY_(offDs)+')'});
  if(!r.ok) return r;
  if(isWorkExtra_(offDs)) deleteWorkOverride_({date:offDs});
  if(isWorkingDay_(offDs)){
    r=addHoliday_({date:offDs, name:'หยุดชดเชย (สลับกับ '+fmtDMY_(workDs)+')'});
    if(!r.ok){ deleteWorkOverride_({date:workDs}); return r; }
  }
  otRetypeDate_(workDs); otRetypeDate_(offDs);
  return {ok:true, work_date:workDs, off_date:offDs, leaves_on_off:calLeavesOn_(offDs)};
}

/* วันหยุด → ให้เป็นวันทำงาน (ไม่มีวันหยุดแทน) */
function calAddWork_(ds, note){
  if(!calDateOk_(ds)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง');
  if(isWorkingDay_(ds)) return err_('CAL_IS_WORK', calLabel_(ds)+' เป็นวันทำงานอยู่แล้ว');
  var r=addWorkOverride_({date:ds, name:String(note||'วันทำงานพิเศษ').slice(0,80)});
  if(!r.ok) return r;
  otRetypeDate_(ds);
  return {ok:true, date:ds};
}

/* วันทำงาน → ให้เป็นวันหยุด */
function calAddHoliday_(ds, name){
  if(!calDateOk_(ds)) return err_('BAD_DATE','วันที่ไม่ถูกต้อง');
  if(!isWorkingDay_(ds)) return err_('CAL_IS_OFF', calLabel_(ds)+' เป็นวันหยุดอยู่แล้ว');
  var pw=calPairedWork_(ds); if(pw) return err_('CAL_PAIRED', calLabel_(ds)+' เป็นวันทำงานจากการสลับกับ '+fmtDMY_(pw)+' — ยกเลิกรายการสลับเดิมก่อน');
  if(isWorkExtra_(ds)) deleteWorkOverride_({date:ds});
  if(isWorkingDay_(ds)){ var r=addHoliday_({date:ds, name:String(name||'วันหยุดบริษัท').slice(0,80)}); if(!r.ok) return r; }
  otRetypeDate_(ds);
  return {ok:true, date:ds, leaves:calLeavesOn_(ds)};
}

/* ลบรายการ (kind: 'H' = วันหยุด, 'W' = วันทำงานพิเศษ) — ถ้าเป็นคู่สลับ ลบอีกฝั่งด้วย */
function calRemove_(kind, ds){
  kind=(kind==='W')?'W':'H';
  var name=calNameOf_(kind, ds);
  if(name===null) return err_('NOT_FOUND','ไม่พบรายการนี้ (อาจถูกลบไปแล้ว)');
  var r=(kind==='W') ? deleteWorkOverride_({date:ds}) : deleteHoliday_({date:ds});
  if(!r.ok) return r;
  var removed=[{kind:kind, date:ds}], pair=calPairOf_(name), other=(kind==='W'?'H':'W');
  if(pair){
    var pn=calNameOf_(other, pair);
    if(pn!==null && calPairOf_(pn)===ds){
      if(other==='W') deleteWorkOverride_({date:pair}); else deleteHoliday_({date:pair});
      removed.push({kind:other, date:pair});
      otRetypeDate_(pair);
    }
  }
  otRetypeDate_(ds);
  return {ok:true, removed:removed};
}

/* รายการพิเศษตั้งแต่วันที่ from เป็นต้นไป เรียงตามวัน */
function calEntries_(from){
  var h=getHolidays_(), out=[];
  (h.holidays||[]).forEach(function(x){ if(x.date>=from) out.push({kind:'H', date:x.date, name:String(x.name||'')}); });
  (h.work_overrides||[]).forEach(function(x){ if(x.date>=from) out.push({kind:'W', date:x.date, name:String(x.name||'')}); });
  out.sort(function(a,b){ return a.date<b.date?-1:(a.date>b.date?1:0); });
  return out;
}

function calEntryLine_(e){
  return (e.kind==='W' ? '🟢 ' : '🔴 ')+'<b>'+tgEscape_(calShort_(e.date))+'</b> — '+(e.kind==='W'?'มาทำงาน':'หยุด')+(e.name?(' · '+tgEscape_(e.name)):'');
}

/* ---------- ต่อกับหน้าผู้ดูแลในเว็บ ---------- */
function calAnnounce_(text){ try{ if(getSetting_('telegram_bot_token') && getSetting_('telegram_chat_id')) tgSend_(text); }catch(e){} }

function calSwapApi_(d){
  var r=calSwap_(String(d.work_date||'').trim(), String(d.off_date||'').trim());
  if(r.ok) calAnnounce_(calSwapDoneText_(r, 'ผู้ดูแล (หน้าเว็บ)'));
  return r;
}

function calAddHolidayApi_(d){
  var ds0=String(d.date||'').trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(ds0) && isWorkExtra_(ds0)) return err_('CAL_PAIRED', calLabel_(ds0)+' ถูกตั้งเป็น "วันทำงานพิเศษ" อยู่ — ลบรายการวันทำงานพิเศษนั้นก่อน (ถ้าเป็นคู่สลับ ระบบจะลบทั้งคู่)');
  var r=addHoliday_(d);
  if(r.ok){ otRetypeDate_(String(d.date||'').trim()); calAnnounce_('🔴 <b>เพิ่มวันหยุด</b>\n'+tgEscape_(calLabel_(String(d.date).trim()))+' · '+tgEscape_(d.name||'วันหยุด')+'\nโดย ผู้ดูแล (หน้าเว็บ)'); }
  return r;
}

function calAddWorkApi_(d){
  var r=addWorkOverride_(d);
  if(r.ok){ otRetypeDate_(String(d.date||'').trim()); calAnnounce_('🟢 <b>เพิ่มวันทำงานพิเศษ</b>\n'+tgEscape_(calLabel_(String(d.date).trim()))+' · '+tgEscape_(d.name||'วันทำงานพิเศษ')+'\nโดย ผู้ดูแล (หน้าเว็บ)'); }
  return r;
}

function calRemoveApi_(kind, d){
  var r=calRemove_(kind, String(d.date||'').trim());
  if(r.ok) calAnnounce_('🗑 <b>ยกเลิกรายการวันหยุด/วันทำงาน</b>\n'+r.removed.map(function(x){ return tgEscape_(calLabel_(x.date))+' ('+(x.kind==='W'?'มาทำงาน':'หยุด')+')'; }).join('\n')+'\nโดย ผู้ดูแล (หน้าเว็บ)');
  return r;
}

/* ---------- Telegram ---------- */
function calSwapDoneText_(r, who){
  return '🔁 <b>สลับวันหยุดแล้ว</b>\n🟢 '+tgEscape_(calLabel_(r.work_date))+' — <b>มาทำงาน</b>\n🔴 '+tgEscape_(calLabel_(r.off_date))+' — <b>หยุดแทน</b>'+
    (r.leaves_on_off?('\n⚠️ วันที่หยุดแทนมีคนแจ้งลาไว้ '+r.leaves_on_off+' คน (ใบลาเดิมยังอยู่ ถ้าไม่ต้องการให้นับ ลบได้ที่หน้าผู้ดูแล → ใบลา)'):'')+
    '\nโดย '+tgEscape_(who)+' · 🕒 '+Utilities.formatDate(new Date(),TZ,'dd/MM HH:mm');
}

function tgCalText_(){
  var list=calEntries_(todayStr_());
  var out='🔁 <b>วันหยุด · วันทำงาน</b>\nวันหยุดปกติ: ทุกวันอาทิตย์\n\n<b>รายการพิเศษที่กำลังจะถึง</b>\n';
  out+=list.length ? list.slice(0,14).map(calEntryLine_).join('\n')+(list.length>14?('\n… และอีก '+(list.length-14)+' รายการ'):'') : '– ไม่มี –';
  return out+'\n\nกดปุ่มด้านล่างเพื่อเปลี่ยน 👇';
}

function tgCalKb_(){
  var rows=[[{text:'🔁 สลับวันหยุด', callback_data:'c:sw'}],
    [{text:'➕ เพิ่มวันทำงาน', callback_data:'c:w'},{text:'🏖 เพิ่มวันหยุด', callback_data:'c:h'}]];
  if(calEntries_(calShift_(todayStr_(),-7)).length) rows.push([{text:'🗑 ยกเลิกรายการ', callback_data:'c:del'}]);
  rows.push([{text:'⬅️ เมนูหลัก', callback_data:'m:main:p'}]);
  return {inline_keyboard:rows};
}

/* ปุ่มวันที่ แถวละ 2 ปุ่ม */
function calDateRows_(dates, prefix){
  var rows=[], row=[];
  dates.forEach(function(ds){ row.push({text:calShort_(ds), callback_data:prefix+calC_(ds)}); if(row.length===2){ rows.push(row); row=[]; } });
  if(row.length) rows.push(row);
  return rows;
}

/* วันหยุดที่กำลังจะถึง (วันอาทิตย์/วันหยุดบริษัท ที่ยังไม่ได้ตั้งให้มาทำงาน) */
function calNextOff_(n){
  var out=[], t=todayStr_();
  for(var k=0;k<70 && out.length<n;k++){ var ds=calShift_(t,k); if(!calIsWork_(ds)) out.push(ds); }
  return out;
}

/* วันทำงานรอบ ๆ วันที่ center (ไม่ย้อนเกินวันนี้) */
function calWorkAround_(center, n){
  var out=[], t=todayStr_();
  for(var k=-6;k<=8 && out.length<n;k++){ var ds=calShift_(center,k); if(ds>=t && ds!==center && calIsWork_(ds)) out.push(ds); }
  return out;
}

function calNextWork_(n){
  var out=[], t=todayStr_();
  for(var k=0;k<40 && out.length<n;k++){ var ds=calShift_(t,k); if(calIsWork_(ds)) out.push(ds); }
  return out;
}

/* หน้าจอแต่ละขั้น → {text, kb} */
function calView_(step, a, b){
  _calDay=null;                                   // อ่านปฏิทินใหม่ทุกครั้งที่วาดหน้าจอ
  var X=[{text:'✖️ ยกเลิก', callback_data:'c:home'}];
  if(step==='sw') return {text:'🔁 <b>สลับวันหยุด (1/2)</b>\nเลือก <b>วันหยุดที่จะให้มาทำงาน</b>',
    kb:{inline_keyboard:calDateRows_(calNextOff_(8),'c:sw1:').concat([[{text:'📝 พิมพ์วันอื่น', callback_data:'c:ask:sw1'}],X])}};
  if(step==='sw1') return {text:'🔁 <b>สลับวันหยุด (2/2)</b>\n🟢 มาทำงานวัน: <b>'+tgEscape_(calLabel_(a))+'</b>\n\nเลือก <b>วันที่จะหยุดแทน</b>',
    kb:{inline_keyboard:calDateRows_(calWorkAround_(a,12),'c:sw2:'+calC_(a)+':').concat([[{text:'📝 พิมพ์วันอื่น', callback_data:'c:ask:sw2:'+calC_(a)}],[{text:'⬅️ ย้อนกลับ', callback_data:'c:sw'}],X])}};
  if(step==='sw2'){ var n=calLeavesOn_(b);
    return {text:'🔁 <b>ยืนยันสลับวันหยุด?</b>\n🟢 '+tgEscape_(calLabel_(a))+' — <b>มาทำงาน</b>\n🔴 '+tgEscape_(calLabel_(b))+' — <b>หยุดแทน</b>'+(n?('\n\n⚠️ วันที่จะหยุดแทนมีคนแจ้งลาไว้แล้ว '+n+' คน'):''),
      kb:{inline_keyboard:[[{text:'✅ ยืนยันสลับ', callback_data:'c:swok:'+calC_(a)+':'+calC_(b)}],[{text:'⬅️ เลือกวันใหม่', callback_data:'c:sw1:'+calC_(a)}],X]}}; }
  if(step==='w') return {text:'➕ <b>เพิ่มวันทำงาน</b>\nเลือกวันหยุดที่จะให้ <b>มาทำงาน</b> (ไม่มีวันหยุดแทน)',
    kb:{inline_keyboard:calDateRows_(calNextOff_(8),'c:w1:').concat([[{text:'📝 พิมพ์วันอื่น', callback_data:'c:ask:calw'}],X])}};
  if(step==='w1') return {text:'➕ <b>ยืนยันเพิ่มวันทำงาน?</b>\n🟢 '+tgEscape_(calLabel_(a))+' — <b>มาทำงาน</b>',
    kb:{inline_keyboard:[[{text:'✅ ยืนยัน', callback_data:'c:wok:'+calC_(a)}],X]}};
  if(step==='h') return {text:'🏖 <b>เพิ่มวันหยุด</b>\nเลือกวันทำงานที่จะให้ <b>หยุด</b>',
    kb:{inline_keyboard:calDateRows_(calNextWork_(12),'c:h1:').concat([[{text:'📝 พิมพ์วันอื่น / ใส่ชื่อวันหยุด', callback_data:'c:ask:calh'}],X])}};
  if(step==='h1'){ var m=calLeavesOn_(a), nm=calHolName_(a);
    return {text:'🏖 <b>ยืนยันเพิ่มวันหยุด?</b>\n🔴 '+tgEscape_(calLabel_(a))+' — <b>หยุด</b>'+(nm?(' · '+tgEscape_(nm)):'')+(m?('\n\n⚠️ วันนี้มีคนแจ้งลาไว้แล้ว '+m+' คน'):''),
      kb:{inline_keyboard:[[{text:'✅ ยืนยัน', callback_data:'c:hok:'+calC_(a)}],X]}}; }
  if(step==='del'){ var list=calEntries_(calShift_(todayStr_(),-7)).slice(0,10);
    return {text:'🗑 <b>ยกเลิกรายการ</b>\nเลือกรายการที่จะยกเลิก'+(list.length?'':'\n– ไม่มีรายการ –'),
      kb:{inline_keyboard:list.map(function(e){ return [{text:(e.kind==='W'?'🟢 ':'🔴 ')+calShort_(e.date)+' '+(e.kind==='W'?'มาทำงาน':'หยุด'), callback_data:'c:d1:'+e.kind+':'+calC_(e.date)}]; }).concat([X])}}; }
  if(step==='d1'){ var nm2=calNameOf_(a,b), pair=calPairOf_(nm2);
    return {text:'🗑 <b>ยืนยันยกเลิกรายการนี้?</b>\n'+calEntryLine_({kind:a, date:b, name:nm2||''})+(pair?('\n\nเป็นรายการสลับวัน → จะยกเลิกคู่ของมัน (<b>'+tgEscape_(calShort_(pair))+'</b>) ด้วย'):''),
      kb:{inline_keyboard:[[{text:'✅ ยืนยันยกเลิก', callback_data:'c:dok:'+a+':'+calC_(b)}],X]}}; }
  return {text:tgCalText_(), kb:tgCalKb_()};
}

/* ชื่อวันหยุดที่พิมพ์มาทาง Telegram (จำไว้ชั่วคราว 15 นาที ระหว่างรอกดยืนยัน) */
function calHolName_(ds, set){
  try{
    var c=CacheService.getScriptCache(), k='calhn_'+ds;
    if(set!==undefined){ c.put(k, String(set).slice(0,80), 900); return String(set); }
    return c.get(k)||'';
  }catch(e){ return ''; }
}

function calEdit_(chatId, msgId, v){
  var r=tgApi_('editMessageText',{chat_id:chatId, message_id:msgId, text:tgClip_(v.text), parse_mode:'HTML', disable_web_page_preview:true, reply_markup:v.kb});
  if(r && (r.ok || /not modified/i.test(String(r.description||'')))) return r;
  return tgSendTo_(chatId, v.text, v.kb);
}

var CAL_ASK = {sw1:'🔁 สลับวันหยุด (1/2)', sw2:'🔁 สลับวันหยุด (2/2)', calw:'➕ เพิ่มวันทำงาน', calh:'🏖 เพิ่มวันหยุด'};

function calAsk_(chatId, kind, from, d1){
  var who=(!tgIsPrivate_(chatId) && from && from.id) ? ('<a href="tg://user?id='+from.id+'">'+tgEscape_(String(from.first_name||'คุณ').slice(0,40))+'</a> ') : '';
  var body={sw1:'พิมพ์ <b>วันหยุดที่จะให้มาทำงาน</b> เช่น <code>11/10</code>',
    sw2:'มาทำงานวัน: '+fmtDMY_(d1||'')+'\nพิมพ์ <b>วันที่จะหยุดแทน</b> เช่น <code>13/10</code>',
    calw:'พิมพ์ <b>วันหยุดที่จะให้มาทำงาน</b> เช่น <code>11/10</code>',
    calh:'พิมพ์ <b>วันที่จะหยุด</b> ตามด้วยชื่อวันหยุด (ถ้ามี) เช่น <code>13/10 ปิดซ่อมเครื่อง</code>'}[kind];
  return tgSendTo_(chatId, '<b>'+CAL_ASK[kind]+'</b>\n'+who+body+'\n<i>(พิมพ์ตอบกลับข้อความนี้)</i>',
    {force_reply:true, selective:!!who, input_field_placeholder:'เช่น 11/10'});
}

/* คำตอบที่พิมพ์กลับมา → ส่งการ์ดยืนยัน (ข้อความใหม่) */
function tgCalTyped_(kind, text, chatId, promptText){
  var parts=String(text||'').trim().split(/\s+/), ds=tgParseDate_(parts[0]), rest=parts.slice(1).join(' ').trim(), v=null, bad='';
  if(ds && parts[0].split(/[\/\-.]/).length===2 && ds<calShift_(todayStr_(),-31)) ds=(parseInt(ds.slice(0,4),10)+1)+ds.slice(4);   // พิมพ์ 3/1 ตอนธันวาคม = ปีหน้า
  if(!ds || !calDateOk_(ds)) bad='❓ อ่านวันที่ไม่ออก — พิมพ์แบบ วัน/เดือน เช่น <code>11/10</code>';
  else if(kind==='sw1'){ if(isWorkingDay_(ds)) bad='⚠️ '+tgEscape_(calLabel_(ds))+' เป็นวันทำงานอยู่แล้ว — ต้องเลือกวันหยุด (วันอาทิตย์/วันหยุดบริษัท)'; else v=calView_('sw1', ds); }
  else if(kind==='sw2'){
    var m=/มาทำงานวัน: (\d{2})\/(\d{2})\/(\d{4})/.exec(String(promptText||'')), d1=m?(m[3]+'-'+m[2]+'-'+m[1]):'';
    if(!d1 || !calDateOk_(d1)) bad='❓ ไม่พบวันที่ตั้งต้น — กดปุ่ม 🔁 สลับวันหยุด เริ่มใหม่อีกครั้ง';
    else if(!isWorkingDay_(ds)) bad='⚠️ '+tgEscape_(calLabel_(ds))+' เป็นวันหยุดอยู่แล้ว — ต้องเลือกวันทำงาน';
    else if(ds===d1) bad='⚠️ สองวันต้องไม่ใช่วันเดียวกัน';
    else v=calView_('sw2', d1, ds);
  }
  else if(kind==='calw'){ if(isWorkingDay_(ds)) bad='⚠️ '+tgEscape_(calLabel_(ds))+' เป็นวันทำงานอยู่แล้ว'; else v=calView_('w1', ds); }
  else if(kind==='calh'){ if(!isWorkingDay_(ds)) bad='⚠️ '+tgEscape_(calLabel_(ds))+' เป็นวันหยุดอยู่แล้ว'; else { calHolName_(ds, safeText_(rest)); v=calView_('h1', ds); } }
  if(bad){ tgSendTo_(chatId, bad, {inline_keyboard:[[{text:'🔁 วันหยุด · สลับวัน', callback_data:'p:cal'}]]}); return; }
  tgSendTo_(chatId, v.text, v.kb);
}

function tgCalCallback_(cq, parts, clicker, chatId, msgId){
  var sub=parts[1]||'home', a=parts[2]||'', b=parts[3]||'';
  var show=function(step, x, y){ tgAnswerCallback_(cq.id,''); calEdit_(chatId, msgId, calView_(step, x, y)); return json_({ok:true}); };
  var fail=function(msg){ tgAnswerCallback_(cq.id, String(msg).slice(0,190)); calEdit_(chatId, msgId, calView_('home')); return json_({ok:true}); };
  if(sub==='ask'){ tgAnswerCallback_(cq.id,''); if(CAL_ASK[a]) calAsk_(chatId, a, cq.from, calD_(b)); return json_({ok:true}); }
  if(sub==='sw' || sub==='w' || sub==='h' || sub==='del' || sub==='home') return show(sub);
  if(sub==='sw1' || sub==='w1' || sub==='h1'){ var d=calD_(a); if(!calDateOk_(d)) return fail('วันที่ไม่ถูกต้อง'); return show(sub, d); }
  if(sub==='sw2'){ var d1=calD_(a), d2=calD_(b); if(!calDateOk_(d1)||!calDateOk_(d2)) return fail('วันที่ไม่ถูกต้อง'); return show('sw2', d1, d2); }
  if(sub==='d1'){ var dd=calD_(b); if(calNameOf_(a,dd)===null) return fail('ไม่พบรายการนี้ (อาจถูกลบไปแล้ว)'); return show('d1', (a==='W'?'W':'H'), dd); }
  if(['swok','wok','hok','dok'].indexOf(sub)<0) return show('home');

  // ----- ลงมือเปลี่ยนปฏิทิน (ต้องได้ล็อกก่อน) -----
  var lock=LockService.getScriptLock(), got=false;
  try{ got=lock.tryLock(15000); }catch(e){}
  if(!got){ tgAnswerCallback_(cq.id,'ระบบกำลังยุ่ง กดอีกครั้ง'); return json_({ok:true}); }
  try{
    var r, text, at=Utilities.formatDate(new Date(),TZ,'dd/MM HH:mm');
    if(sub==='swok'){ r=calSwap_(calD_(a), calD_(b)); if(r.ok) text=calSwapDoneText_(r, clicker); }
    else if(sub==='wok'){ r=calAddWork_(calD_(a)); if(r.ok) text='🟢 <b>เพิ่มวันทำงานแล้ว</b>\n'+tgEscape_(calLabel_(r.date))+' — <b>มาทำงาน</b>\nโดย '+tgEscape_(clicker)+' · 🕒 '+at; }
    else if(sub==='hok'){ var nm=calHolName_(calD_(a)); r=calAddHoliday_(calD_(a), nm);
      if(r.ok) text='🔴 <b>เพิ่มวันหยุดแล้ว</b>\n'+tgEscape_(calLabel_(r.date))+' — <b>หยุด</b>'+(nm?(' · '+tgEscape_(nm)):'')+(r.leaves?('\n⚠️ วันนี้มีคนแจ้งลาไว้ '+r.leaves+' คน (ใบลาเดิมยังอยู่)'):'')+'\nโดย '+tgEscape_(clicker)+' · 🕒 '+at; }
    else { r=calRemove_(a, calD_(b));
      if(r.ok) text='🗑 <b>ยกเลิกรายการแล้ว</b>\n'+r.removed.map(function(x){ return tgEscape_(calLabel_(x.date))+' ('+(x.kind==='W'?'มาทำงาน':'หยุด')+') → กลับเป็นปกติ'; }).join('\n')+'\nโดย '+tgEscape_(clicker)+' · 🕒 '+at; }
    if(!r.ok) return fail(r.error||'ทำรายการไม่สำเร็จ');
    tgAnswerCallback_(cq.id,'บันทึกแล้ว ✓');
    var groupChat=String(getSetting_('telegram_chat_id')||'').trim();
    if(groupChat) tgSendTo_(groupChat, text, tgBottomKb_());             // ประกาศเป็นข้อความใหม่ในกลุ่มเสมอ (ทุกคนได้แจ้งเตือน)
    calEdit_(chatId, msgId, (String(chatId)===groupChat) ? calView_('home') : {text:text, kb:{inline_keyboard:[[{text:'📅 ดูรายการทั้งหมด', callback_data:'p:cal'}]]}});
  } catch(e){ try{ tgAnswerCallback_(cq.id,'เกิดข้อผิดพลาด ลองใหม่อีกครั้ง'); }catch(e2){} }
  finally { try{ lock.releaseLock(); }catch(e){} }
  return json_({ok:true});
}

/* รอบเย็น: ถ้าพรุ่งนี้เป็นวันพิเศษ (มาทำงานวันหยุด / หยุดพิเศษ) ให้บอทเตือนกลุ่มก่อน */
function calTomorrowNote_(ds){
  var w=calNameOf_('W', ds), h=calNameOf_('H', ds);
  if(w!==null && isWorkingDay_(ds)) return '📣 <b>พรุ่งนี้ '+tgEscape_(calLabel_(ds))+' เป็นวันทำงาน</b>\n'+tgEscape_(w);
  if(h!==null && !isWorkingDay_(ds) && parseDateStr_(ds).getDay()!==0) return '📣 <b>พรุ่งนี้ '+tgEscape_(calLabel_(ds))+' หยุด</b>\n'+tgEscape_(h);
  return '';
}

/* ==================================================================
 *  ดูข้อมูลพนักงานจาก Telegram: เลือกจากรายชื่อ (ปุ่ม แบ่งหน้า) หรือ พิมพ์ชื่อ/รหัส
 *  callback: e:home · e:l:<TH|MM>:<หน้า> · e:s:<รหัส> · e:ask · e:n
 * ================================================================== */
var TG_EMP_PER_PAGE = 10;

function tgEmpByNat_(nat){
  var m=empMap_(), out=[];
  Object.keys(m).forEach(function(id){ var e=m[id]; if(e.active && e.nationality===nat) out.push(e); });
  var key=function(e){ return String(e.emp_id).replace(/\d+/g,function(n){ return ('000000'+n).slice(-6); }); };
  out.sort(function(a,b){ return key(a)<key(b)?-1:1; });
  return out;
}

function tgEmpHome_(){
  var th=tgEmpByNat_('TH').length, mm=tgEmpByNat_('MM').length;
  return {text:'🔎 <b>ดูข้อมูลพนักงาน</b>\nเลือกจากรายชื่อ หรือพิมพ์ชื่อ/รหัสเอง 👇',
    kb:{inline_keyboard:[[{text:'🇹🇭 คนไทย ('+th+')', callback_data:'e:l:TH:0'},{text:'🇲🇲 พม่า ('+mm+')', callback_data:'e:l:MM:0'}],
      [{text:'⌨️ พิมพ์ชื่อหรือรหัส', callback_data:'e:ask'}]]}};
}

function tgEmpList_(nat, page){
  nat=(nat==='MM')?'MM':'TH';
  var all=tgEmpByNat_(nat), pages=Math.max(1, Math.ceil(all.length/TG_EMP_PER_PAGE));
  page=Math.min(Math.max(parseInt(page,10)||0,0), pages-1);
  var part=all.slice(page*TG_EMP_PER_PAGE, (page+1)*TG_EMP_PER_PAGE), rows=[], row=[];
  part.forEach(function(e){
    var data='e:s:'+e.emp_id; if(unescape(encodeURIComponent(data)).length>64) return;
    row.push({text:String(e.name_th||e.name_mm||'').slice(0,22)+' · '+e.emp_id, callback_data:data});
    if(row.length===2){ rows.push(row); row=[]; }
  });
  if(row.length) rows.push(row);
  if(pages>1) rows.push([{text:(page>0?'◀ ก่อนหน้า':' '), callback_data:(page>0?('e:l:'+nat+':'+(page-1)):'e:n')},
    {text:'หน้า '+(page+1)+'/'+pages, callback_data:'e:n'},
    {text:(page<pages-1?'ถัดไป ▶':' '), callback_data:(page<pages-1?('e:l:'+nat+':'+(page+1)):'e:n')}]);
  rows.push([{text:(nat==='TH'?'🇲🇲 ดูพม่า':'🇹🇭 ดูคนไทย'), callback_data:'e:l:'+(nat==='TH'?'MM':'TH')+':0'},{text:'⌨️ พิมพ์ค้นหา', callback_data:'e:ask'}]);
  return {text:(nat==='TH'?'🇹🇭 <b>พนักงานคนไทย</b>':'🇲🇲 <b>พนักงานพม่า</b>')+' — '+all.length+' คน\n'+(all.length?'กดชื่อเพื่อดูข้อมูลการลา + โอที 👇':'<i>ยังไม่มีพนักงาน</i>'),
    kb:{inline_keyboard:rows}};
}

function tgEmpCallback_(cq, parts, chatId, msgId){
  var sub=parts[1]||'home';
  try{
    if(sub==='n'){ tgAnswerCallback_(cq.id,''); return json_({ok:true}); }
    if(sub==='ask'){ tgAnswerCallback_(cq.id,''); tgAsk_(chatId,'find',cq.from); return json_({ok:true}); }
    tgAnswerCallback_(cq.id,'');
    if(sub==='l'){ calEdit_(chatId, msgId, tgEmpList_(parts[2], parts[3])); return json_({ok:true}); }
    if(sub==='s'){
      var id=parts.slice(2).join(':'), e=empMap_()[id];
      if(!e){ calEdit_(chatId, msgId, {text:'ไม่พบพนักงานรหัส '+tgEscape_(id)+' (อาจถูกลบหรือเปลี่ยนรหัสแล้ว)', kb:tgEmpHome_().kb}); return json_({ok:true}); }
      var nat=(e.nationality==='MM')?'MM':'TH', all=tgEmpByNat_(nat), idx=0;
      for(var i=0;i<all.length;i++) if(all[i].emp_id===id) idx=i;
      calEdit_(chatId, msgId, {text:tgEmpCard_(e), kb:{inline_keyboard:[[{text:'◀ กลับไปรายชื่อ', callback_data:'e:l:'+nat+':'+Math.floor(idx/TG_EMP_PER_PAGE)},{text:'🔎 ค้นหาใหม่', callback_data:'e:home'}]]}});
      return json_({ok:true});
    }
    calEdit_(chatId, msgId, tgEmpHome_());
  }catch(err){ try{ tgAnswerCallback_(cq.id,'เกิดข้อผิดพลาด ลองใหม่อีกครั้ง'); }catch(e2){} }
  return json_({ok:true});
}

/* ==================================================================
 *  MAIN — หน้าตาชีตแบบโปรแกรม: ชีต "สรุป" (แดชบอร์ดรอบเงินเดือน) + จัดรูปแบบทุกชีตให้เหมือนกัน
 *  ทุกขั้นห่อด้วย try: ถ้าตกแต่งส่วนไหนไม่ได้ ระบบหลักยังทำงานปกติ และ setup จะรายงานว่าส่วนไหนพลาด
 * ================================================================== */
var SUM_SHEET = 'สรุป';

var UI_C = {TITLE:'#5A9E1B', NAVY:'#1F4E79', NAVY2:'#163A5C', GREEN:'#5A9E1B', SKY:'#1D9BD7', AMBER:'#B26A00', SOFT:'#EDF3F8', BAND:'#F6F9FC', LINE:'#D6E0EA', INK:'#12283D', MUTED:'#5A6B7D', YELLOW:'#FFF3CD'};

var UI_FONT = 'Sarabun';

var _styleErr = [];

function styleTry_(label, fn){ try{ fn(); }catch(e){ _styleErr.push(label+': '+String(e).slice(0,120)); } }

/* เดือนของรอบเงินเดือนปัจจุบัน (วันที่ 1 ของเดือนที่รอบจบ) */
function sumDefaultMonth_(){ var c=otCurrentCycle_(); return new Date(c.end.getFullYear(), c.end.getMonth(), 1); }

/* ชีตสรุปต้องมีก่อน (ชีต Payroll อ้างช่องเลือกรอบจากที่นี่) */
function sumSheet_(){
  var sh=sheet_(SUM_SHEET);
  if(sh && sh.getLastRow()>0 && String(sh.getRange('A3').getValue()).trim()!=='รอบเงินเดือน ▸'){ shelveSheet_(sh); sh=null; }   // ชีตชื่อนี้ของผู้ใช้เอง → เปลี่ยนชื่อเก็บไว้
  if(!sh){ sh=ss_().insertSheet(SUM_SHEET); sh.getRange('A3').setValue('รอบเงินเดือน ▸'); }
  var v=sh.getRange('C3').getValue();
  if(!(v instanceof Date)) sh.getRange('C3').setValue(sumDefaultMonth_()).setNumberFormat('mmmm yyyy');
  return sh;
}

/* รายชื่อในชีตสรุป: [{emp_id, name, tag (ข้อความคอลัมน์ C), pay (ชื่อชีต Payroll ที่ดึงตัวเลข)}] — ไทยก่อน แล้วพม่า เรียงตามรหัส */
function sumRoster_(){
  var out=[], vals=sheet_(SHEETS.EMP).getDataRange().getValues();
  for(var i=1;i<vals.length;i++){ var e=empFromRow_(vals[i], i+1); if(!e.emp_id || !e.active) continue;
    out.push({emp_id:e.emp_id, name:e.name_th||e.name_mm, nat:e.nationality, grp:0, tag:(e.nationality==='MM'?'พม่า':'ไทย'), pay:(e.nationality==='MM'?SHEETS.PAY_MM:SHEETS.PAY_TH)}); }
  return sumSort_(out);
}

function sumSort_(list){
  var key=function(e){ return String(e.grp||0)+(e.nat==='MM'?'1':'0')+String(e.emp_id).replace(/\d+/g,function(n){ return ('000000'+n).slice(-6); }); };
  return list.sort(function(a,b){ return key(a)<key(b)?-1:1; });
}

/* ชีตชื่อชนกับของผู้ใช้ → เปลี่ยนชื่อชีตเดิมเก็บไว้ (ไม่ลบ ไม่แก้ข้อมูล) แล้วระบบค่อยสร้างชีตของตัวเอง */
function shelveSheet_(sh){
  var base=sh.getName()+' (เดิม)', nm=base, k=2;
  while(ss_().getSheetByName(nm)) nm=base+' '+(k++);
  sh.setName(nm);
}

/* สร้าง/วาดชีต "สรุป" ใหม่ทั้งหน้า (ค่าในช่องเลือกรอบ C3 คงเดิม) */
function buildSummarySheet_(){
  var sh=sumSheet_(), month=sh.getRange('C3').getValue(), CO=getCutoffDay_(APP.CO);
  var emps=sumRoster_();
  var n=emps.length, R0=9, NC=13, last=R0+Math.max(n,1)-1, tot=last+1;

  styleTry_('สรุป: ล้างหน้า', function(){
    try{ sh.getBandings().forEach(function(b){ b.remove(); }); }catch(e){}
    try{ sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); }); }catch(e){}
    var mr=sh.getMaxRows(), mc=sh.getMaxColumns();
    try{ sh.getRange(1,1,mr,mc).breakApart(); }catch(e){}
    sh.getRange(1,1,mr,mc).clearContent().clearFormat().clearNote().clearDataValidations();
    try{ sh.clearConditionalFormatRules(); }catch(e){}
  });
  // ---- หัวหน้า + ช่องเลือกรอบ ----
  sh.getRange('A1').setValue(APP.TITLE);
  sh.getRange('C1').setValue('สรุปการลา + โอที  ประจำรอบเงินเดือน');
  sh.getRange('A3').setValue('รอบเงินเดือน ▸');
  sh.getRange('C3').setValue(month).setNumberFormat('mmmm yyyy')
    .setNote('เปลี่ยนรอบ: พิมพ์วันที่ 1 ของเดือนที่รอบจบ เช่น 1/11/2026 = รอบ '+(CO+1)+' ต.ค. – '+CO+' พ.ย.\nทุกชีต (สรุป, Payroll_TH, Payroll_MM) จะเปลี่ยนตามช่องนี้');
  sh.getRange('E3').setFormula('="ช่วงวันที่  "&TEXT(DATE(YEAR($C$3),MONTH($C$3)-1,'+(CO+1)+'),"d/m/yyyy")&"  –  "&TEXT(DATE(YEAR($C$3),MONTH($C$3),'+CO+'),"d/m/yyyy")&"     •  เปลี่ยนรอบได้ที่ช่องสีเหลือง"');
  // ---- การ์ดตัวเลข ----
  var cards=[['A','B','พนักงานทั้งหมด (คน)','=COUNTUNIQUE($A$'+R0+':$A$'+last+')'],
             ['C','E','คนที่ลาในรอบนี้ (คน)','=IFERROR(COUNTUNIQUE(FILTER($A$'+R0+':$A$'+last+',$G$'+R0+':$G$'+last+'>0)),0)'],
             ['F','G','รวมวันลา (วัน)','=SUM($G$'+R0+':$G$'+last+')'],
             ['H','I','ยอดหักรวม (บาท)','=SUM($I$'+R0+':$I$'+last+')'],
             ['J','K','รวมโอที (ชม.)','=SUM($L$'+R0+':$L$'+last+')'],
             ['L','M','โอทีรอยืนยัน (ชม.)','=SUM($M$'+R0+':$M$'+last+')']];
  cards.forEach(function(c){ sh.getRange(c[0]+'5').setValue(c[2]); sh.getRange(c[0]+'6').setFormula(n?c[3]:'=0'); });
  // ---- หัวตาราง 2 ชั้น ----
  sh.getRange(7,1,1,NC).setValues([['พนักงาน','','','การลา (ชั่วโมง)','','','','หักเงิน','','โอที (ชั่วโมง)','','','']]);
  sh.getRange(8,1,1,NC).setValues([['รหัส','ชื่อ','สัญชาติ','ลาป่วย','ลากิจ','ลาพักร้อน','รวมลา (วัน)','ชม.ที่หักเงิน','ยอดหัก (บาท)','โอทีวันทำงาน','ทำงานวันหยุด','รวมโอที','รอยืนยัน']])
    .setNotes([['รหัสพนักงาน','ชื่อพนักงาน','ไทย / พม่า','ชั่วโมงลาป่วยในรอบนี้','ชั่วโมงลากิจในรอบนี้','ชั่วโมงลาพักร้อนที่อนุมัติแล้ว','ลาป่วย + ลากิจ + ลาพักร้อน คิดเป็นวัน (8 ชม. = 1 วัน)',
      'ชั่วโมงที่ติ๊ก "หักเงิน" ในชีตรายวัน','ชม.ที่หักเงิน × ค่าแรงต่อชั่วโมง (กรอกค่าแรงในชีตรายวัน)','โอทีนอกเวลา 08:00–17:00 ของวันทำงาน (ยืนยันแล้ว)','ชั่วโมงทำงานในวันอาทิตย์/วันหยุด (ยืนยันแล้ว)','โอทีวันทำงาน + ทำงานวันหยุด','พนักงานแจ้งเองแต่ยังไม่มีคนยืนยัน — ยังไม่นับ']]);
  // ---- ข้อมูลรายคน (ดึงจากชีต Payroll ตามรหัส) ----
  if(n){
    var rows=emps.map(function(e, k){
      var r=R0+k, P="'"+e.pay+"'!$A:$P";
      var f=function(col){ return '=IFERROR(VLOOKUP($A'+r+','+P+','+col+',FALSE),0)'; };
      return [e.emp_id, e.name, e.tag, f(5), f(6), f(7), f(9), f(10), f(11), f(13), f(14), f(15), f(16)];
    });
    sh.getRange(R0,1,n,1).setNumberFormat('@');
    sh.getRange(R0,1,n,NC).setValues(rows);
    sh.getRange(tot,1,1,3).setValues([['','รวมทั้งหมด','']]);
    sh.getRange(tot,4,1,10).setFormulas([['D','E','F','G','H','I','J','K','L','M'].map(function(c){ return '=SUM('+c+R0+':'+c+last+')'; })]);
  } else sh.getRange(R0,1).setValue('— ยังไม่มีพนักงาน —');

  // ---- ตกแต่ง ----
  styleTry_('สรุป: รูปแบบ', function(){
    var C=UI_C, all=sh.getRange(1,1,Math.max(tot,R0)+1,NC);
    all.setFontFamily(UI_FONT).setFontSize(11).setFontColor(C.INK).setVerticalAlignment('middle');
    sh.getRange('A1:B1').merge().setBackground(C.TITLE).setFontColor('#ffffff').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('C1:M1').merge().setBackground(C.NAVY).setFontColor('#ffffff').setFontSize(17).setFontWeight('bold').setHorizontalAlignment('left');
    sh.setRowHeight(1,46); sh.setRowHeight(2,8); sh.setRowHeight(3,34); sh.setRowHeight(4,10); sh.setRowHeight(5,22); sh.setRowHeight(6,44); sh.setRowHeight(7,26); sh.setRowHeight(8,38);
    sh.getRange('A3:B3').merge().setFontWeight('bold').setHorizontalAlignment('right').setFontSize(12);
    sh.getRange('C3:D3').merge().setBackground(C.YELLOW).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(13)
      .setBorder(true,true,true,true,false,false,'#E0B43C',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange('E3:M3').merge().setFontColor(C.MUTED).setHorizontalAlignment('left');
    [['A','B',C.NAVY],['C','E',C.NAVY],['F','G',C.SKY],['H','I',C.AMBER],['J','K',C.GREEN],['L','M',C.AMBER]].forEach(function(c){
      sh.getRange(c[0]+'5:'+c[1]+'5').merge().setBackground(C.SOFT).setFontColor(C.MUTED).setFontSize(10).setHorizontalAlignment('center');
      sh.getRange(c[0]+'6:'+c[1]+'6').merge().setBackground(C.SOFT).setFontColor(c[2]).setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('#,##0');
      sh.getRange(c[0]+'5:'+c[1]+'6').setBorder(true,true,true,true,false,false,'#ffffff',SpreadsheetApp.BorderStyle.SOLID_THICK);
    });
    sh.getRange('F6').setNumberFormat('#,##0.0');
    sh.getRange('A7:B7').merge(); sh.getRange('A7:C7').setBackground(C.NAVY2); sh.getRange('D7:G7').merge().setBackground(C.NAVY);
    sh.getRange('H7:I7').merge().setBackground(C.AMBER); sh.getRange('J7:M7').merge().setBackground(C.GREEN);
    sh.getRange(7,1,1,NC).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
    sh.getRange(8,1,1,NC).setBackground(C.SOFT).setFontColor(C.INK).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true)
      .setBorder(false,false,true,false,false,false,C.NAVY,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    if(n){
      var bg=[]; for(var k=0;k<n;k++){ var line=[]; for(var c=0;c<NC;c++) line.push(k%2===1?C.BAND:'#ffffff'); bg.push(line); }
      sh.getRange(R0,1,n,NC).setBackgrounds(bg);
      sh.getRange(R0,1,n,1).setHorizontalAlignment('center').setFontColor(C.MUTED);
      sh.getRange(R0,2,n,1).setFontWeight('bold');
      sh.getRange(R0,3,n,1).setHorizontalAlignment('center');
      sh.getRange(R0,4,n+1,10).setHorizontalAlignment('right').setNumberFormat('#,##0;-#,##0;"–"');
      sh.getRange(R0,7,n+1,1).setNumberFormat('#,##0.00;-#,##0.00;"–"');
      sh.getRange(R0,9,n+1,1).setNumberFormat('#,##0.00;-#,##0.00;"–"');
      sh.getRange(R0,12,n,1).setFontWeight('bold').setFontColor(C.GREEN);
      sh.getRange(R0,1,n,NC).setBorder(null,null,true,null,null,true,C.LINE,SpreadsheetApp.BorderStyle.SOLID);
      sh.getRange(tot,1,1,NC).setBackground(C.SOFT).setFontWeight('bold').setBorder(true,false,true,false,false,false,C.NAVY,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sh.getRange(tot,2).setHorizontalAlignment('right');
      var amber=SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#B26A00').setBold(true).setRanges([sh.getRange(R0,13,n,1)]).build();
      var red=SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#C62828').setRanges([sh.getRange(R0,8,n,2)]).build();
      sh.setConditionalFormatRules([amber, red]);
    }
    [90,210,(APP.MULTI?170:70),80,80,90,95,100,110,105,105,90,90].forEach(function(w, i){ sh.setColumnWidth(i+1, w); });
    sh.setFrozenRows(8); sh.setFrozenColumns(2);
    sh.setHiddenGridlines(true);
    if(sh.getMaxColumns()>NC) sh.hideColumns(NC+1, sh.getMaxColumns()-NC);
  });
  styleTry_('สรุป: ป้องกันแก้ผิดช่อง', function(){
    var p=sh.protect().setWarningOnly(true).setDescription('หน้าสรุป — แก้ได้เฉพาะช่องสีเหลือง (เลือกรอบ)');
    try{ p.setUnprotectedRanges([sh.getRange('C3:D3')]); }catch(e){}
  });
}

/* ชีต Payroll: ช่อง B1 ตามช่องเลือกรอบของชีต "สรุป" + ตกแต่ง */
function payrollPolish_(name, n, label){
  var sh=sheet_(name); if(!sh) return;
  sh.getRange('B1').setFormula("='"+SUM_SHEET+"'!$C$3").setNumberFormat('mmmm yyyy')
    .setNote('รอบเงินเดือนของหน้านี้ตามช่องสีเหลืองในชีต "'+SUM_SHEET+'" — เปลี่ยนรอบที่ชีตนั้นที่เดียว');
  styleTry_(name+': รูปแบบ', function(){
    var C=UI_C, NC=16;
    sh.getRange(1,1,Math.max(n,1)+3,NC).setFontFamily(UI_FONT);
    sh.getRange('A1').setValue(label||'รอบ ▸').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(9).setWrap(true).setFontColor(C.NAVY);
    sh.getRange('F1').setValue('ⓘ เปลี่ยนรอบที่ชีต "'+SUM_SHEET+'" (ช่องสีเหลือง)  •  ค่าแรง และการติ๊กหัก/ไม่หัก ทำที่ชีต "รายวัน"').setFontColor(C.MUTED);
    sh.setRowHeight(1,30); sh.setRowHeight(2,46);
    sh.getRange('B1').setBackground(C.SOFT);
    if(n){ var b9=[], b1=[]; for(var k=0;k<n;k++){ var col=(k%2===1?C.BAND:'#ffffff'), ln=[]; for(var c=0;c<9;c++) ln.push(col); b9.push(ln); b1.push([col]); }
      sh.getRange(3,1,n,9).setBackgrounds(b9); sh.getRange(3,12,n,1).setBackgrounds(b1); }
    if(n){ sh.getRange(3,1,n,NC).setBorder(null,null,true,null,null,true,C.LINE,SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('middle');
      sh.getRange(3,1,n,1).setHorizontalAlignment('center'); }
    sh.setFrozenColumns(2);
    sh.setHiddenGridlines(true);
    if(sh.getMaxColumns()>NC) sh.hideColumns(NC+1, sh.getMaxColumns()-NC);
  });
  styleTry_(name+': ป้องกันแก้สูตร', function(){
    try{ sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); }); }catch(e){}
    sh.protect().setWarningOnly(true).setDescription('หน้านี้เป็นสูตรทั้งหน้า — เปลี่ยนรอบที่ชีต "'+SUM_SHEET+'"');
  });
}

/* ชีตข้อมูล (ตารางยาว): หัวตารางสูง + แถบสีสลับ + ตัวกรอง + ซ่อนช่องเทคนิค + ซ่อนคอลัมน์ว่างด้านขวา */
function dataSheetPolish_(name, ncols, hideCols, centerCols){
  var sh=sheet_(name); if(!sh) return;
  styleTry_(name+': รูปแบบ', function(){
    var C=UI_C, mr=sh.getMaxRows();
    ncols=Math.max(ncols, sh.getLastColumn());        // ผู้ใช้เพิ่มคอลัมน์เองทางขวา → นับรวมด้วย (ไม่ซ่อน และตัวกรองครอบถึง)
    sh.getRange(1,1,mr,ncols).setFontFamily(UI_FONT).setVerticalAlignment('middle');
    sh.getRange(1,1,1,ncols).setHorizontalAlignment('center').setFontSize(11);
    sh.setRowHeight(1,40);
    try{ sh.getBandings().forEach(function(b){ b.remove(); }); }catch(e){}
    sh.getRange(1,1,mr,ncols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
      .setHeaderRowColor(C.NAVY).setFirstRowColor('#FFFFFF').setSecondRowColor(C.BAND);
    (centerCols||[]).forEach(function(c){ if(mr>1) sh.getRange(2,c,mr-1,1).setHorizontalAlignment('center'); });
    if(sh.getMaxColumns()>ncols) sh.hideColumns(ncols+1, sh.getMaxColumns()-ncols);
    (hideCols||[]).forEach(function(c){ sh.hideColumns(c, 1); });
  });
  styleTry_(name+': ตัวกรอง', function(){ if(!sh.getFilter()) sh.getRange(1,1,sh.getMaxRows(),Math.max(ncols, sh.getLastColumn())).createFilter(); });
}

/* จัดรูปแบบทั้งไฟล์ (เรียกจาก setup) */
function styleAllSheets_(){
  buildSummarySheet_();
  // Employees: ซ่อนรหัสรูป (E) + ช่องระบบ (J–L)
  var mix=companyMix_().show || APP.MULTI;        // ระบบหลายกลุ่ม: คอลัมน์บริษัทเป็นข้อมูลจริง ไม่ซ่อน
  dataSheetPolish_(SHEETS.EMP, 12, (mix?[5]:[5,10,11,12]), [1,4,6,7]);
  // Leaves: ซ่อน รหัสกลุ่ม (P) / อุปกรณ์ (Q) / บริษัท (R)
  dataSheetPolish_(SHEETS.LEAVE, 18, (mix?[16,17]:[16,17,18]), [2,3,5,6,8,9,10,11,13]);
  // OT: ซ่อน รหัสกลุ่ม (P) / อุปกรณ์ (Q)
  dataSheetPolish_(OT_SHEET, 18, (mix?[16,17]:[16,17,18]), [2,3,5,6,8,9,10,11]);
  dataSheetPolish_(SHEETS.HOLIDAY, 3, [], [1]);
  dataSheetPolish_(SHEETS.WORKDAY, 3, [], [1]);
  if(SHEETS.TRANSFER) dataSheetPolish_(SHEETS.TRANSFER, 8, [], [1,2,4,5,6,7,8]);
  dataSheetPolish_(SHEETS.CANCEL, 15, (mix?[15]:[13,15]), [1,2,3,4,6,7,10,11]);
  styleTry_('Settings: รูปแบบ', function(){
    var st=sheet_(SHEETS.SET); if(!st) return;
    st.getRange(1,1,st.getMaxRows(),3).setFontFamily(UI_FONT).setVerticalAlignment('middle');
    st.setColumnWidth(1,210); st.setColumnWidth(2,380); st.setRowHeight(1,34);
    if(st.getMaxColumns()>3) st.hideColumns(4, st.getMaxColumns()-3);
  });
  styleTry_('อ่านก่อน: รูปแบบ', function(){
    var gs=sheet_(GUIDE_SHEET); if(!gs) return;
    gs.getRange(1,1,gs.getMaxRows(),3).setFontFamily(UI_FONT);
    gs.getRange('A1:C1').merge().setBackground(UI_C.NAVY).setFontColor('#ffffff').setFontSize(17).setVerticalAlignment('middle');
    gs.getRange('A2:C2').merge().setWrap(true).setFontColor(UI_C.MUTED).setBackground(UI_C.SOFT).setVerticalAlignment('middle');
    gs.setRowHeight(1,46); gs.setRowHeight(2,34);
    var gv=gs.getRange(1,1,gs.getLastRow(),1).getValues(), al=gv.map(function(r){ var v=String(r[0]); return [(/^\d+$/.test(v)||v==='⚠️')?'center':'left']; });
    gs.getRange(1,1,al.length,1).setHorizontalAlignments(al);
    gs.setFrozenRows(1);
    if(gs.getMaxColumns()>3) gs.hideColumns(4, gs.getMaxColumns()-3);
  });
}

/* ==================================================================
 *  ชีต "รายวัน_TH / รายวัน_MM" — หน้าตาแบบโปรแกรม (ตำแหน่งช่องเดิมทุกช่อง: B1 เลือกคน, F1 รอบ, B2–B4 ค่าแรง/โควต้า, คอลัมน์ F ติ๊กหัก)
 * ================================================================== */
var DAILY_STYLE_ROWS = 50;      // จำนวนแถวใต้หัวตารางที่จัดรูปแบบทุกครั้ง (วันในรอบ ≤ 31 + แถวสรุป)

function dailyIsDate_(v){ return Object.prototype.toString.call(v)==='[object Date]'; }

/* ส่วนที่อยู่กับที่: ความกว้างคอลัมน์ แถบหัว ช่องกรอกสีเหลือง หัวตาราง (เรียกตอน setup / รีเฟรชหน้ารายวัน) */
function dailyFrame_(sh){
  if(!sh) return;

  var C=UI_C, S=DAILY_TABLE_START, NR=DAILY_STYLE_ROWS, B=SpreadsheetApp.BorderStyle;
  sh.getRange(1,1,sh.getMaxRows(),7).setFontFamily(UI_FONT).setVerticalAlignment('middle');
  [150,70,112,200,70,110,250].forEach(function(w, i){ sh.setColumnWidth(i+1, w); });
  try{ sh.getRange(1,1,6,7).breakApart(); }catch(e){}
  sh.getRange('B1:D1').merge();                               // ช่องเลือกพนักงานกว้างพอเห็นชื่อเต็ม
  sh.getRange('A5:G5').merge();
  sh.setRowHeight(1,38); sh.setRowHeight(2,28); sh.setRowHeight(3,28); sh.setRowHeight(4,28); sh.setRowHeight(5,24); sh.setRowHeight(6,8); sh.setRowHeight(7,32);
  // แถว 1: เลือกพนักงาน + รอบ
  sh.getRange('A1:G1').setBackground(C.SOFT);
  sh.getRange('A1').setValue('พนักงาน ▸').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right').setFontColor(C.INK);
  sh.getRange('B1:D1').setBackground('#ffffff').setFontWeight('bold').setFontSize(12).setFontColor(C.INK).setHorizontalAlignment('left')
    .setBorder(true,true,true,true,false,false,C.NAVY,B.SOLID_MEDIUM);
  sh.getRange('E1').setValue('รอบ ▸').setFontWeight('bold').setHorizontalAlignment('right').setFontColor(C.INK);
  sh.getRange('F1').setBackground(C.YELLOW).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('mmm yyyy')
    .setBorder(true,true,true,true,false,false,'#E0B43C',B.SOLID_MEDIUM)
    .setNote('รอบของหน้านี้ (ดูทีละคน): พิมพ์วันที่ 1 ของเดือนที่รอบจบ เช่น 1/11/2026\nหน้านี้เลือกรอบแยกจากชีต "'+SUM_SHEET+'" ได้ — เวลาทำเงินเดือนให้ตั้งเป็นรอบเดียวกัน');
  sh.getRange('G1').setFontColor(C.NAVY).setFontWeight('bold').setHorizontalAlignment('left');
  // แถว 2–4: ค่าแรง / โควต้า (ช่องสีเหลือง = กรอกได้)
  sh.getRange('A2:A4').setFontWeight('bold').setHorizontalAlignment('right').setFontColor(C.INK);
  sh.getRange('B2:B4').setBackground(C.YELLOW).setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true,true,true,true,false,true,'#E0B43C',B.SOLID);
  sh.getRange('C2:C4').setFontColor(C.MUTED).setHorizontalAlignment('right').setFontWeight('normal');
  sh.getRange('E2:E4').setFontColor(C.MUTED).setHorizontalAlignment('right').setFontWeight('normal');
  sh.getRange('D2:D4').setFontWeight('bold').setHorizontalAlignment('left').setFontColor(C.INK);
  sh.getRange('F3:F4').setFontWeight('bold').setHorizontalAlignment('center').setFontColor(C.GREEN);
  sh.getRange('A5').setFontColor(C.MUTED).setFontSize(9).setHorizontalAlignment('left').setWrap(false);
  // หัวตาราง
  sh.getRange(S-1,1,1,7).setBackground(C.NAVY).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(11);
  // ตาราง: การจัดแนวของแต่ละคอลัมน์
  sh.getRange(S,1,NR,2).setHorizontalAlignment('center');
  sh.getRange(S,4,NR,1).setHorizontalAlignment('left');
  sh.getRange(S,5,NR,2).setHorizontalAlignment('center');
  sh.getRange(S,7,NR,1).setHorizontalAlignment('left');
  try{ sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#C62828').setBold(true).setRanges([sh.getRange('F3:F4')]).build()]); }catch(e){}
  try{ sh.setHiddenGridlines(true); }catch(e){}
  try{ var used=Math.max(7, sh.getLastColumn()); if(sh.getMaxColumns()>used) sh.hideColumns(used+1, sh.getMaxColumns()-used); }catch(e){}
  try{
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
    sh.protect().setWarningOnly(true).setDescription('หน้ารายวัน — แก้ได้เฉพาะช่องสีเหลือง (เลือกคน / รอบ / ค่าแรง / โควต้า) และช่องติ๊ก "หักเงิน?"')
      .setUnprotectedRanges([sh.getRange('B1:D1'), sh.getRange('F1'), sh.getRange('B2:B4'), sh.getRange(S,6,NR,1)]);
  }catch(e){}
  dailyStyleRows_(sh);
}

/* ส่วนที่เปลี่ยนตามคน/รอบ: สีของแต่ละแถว (เรียกทุกครั้งที่ตารางถูกสร้างใหม่ — ใช้คำสั่งน้อย ให้ทันตอนเปลี่ยนชื่อในดรอปดาวน์) */
function dailyStyleRows_(sh){
  var C=UI_C, S=DAILY_TABLE_START, NR=DAILY_STYLE_ROWS;
  var v=sh.getRange(S,1,NR,7).getValues(), bg=[], fc=[], fw=[], nDays=0;
  var TINT={'ลาป่วย':'#FDECEC', 'ลากิจ':'#FFF4DB', 'ลาพักร้อน':'#E6F2FD'}, TCOL={'ลาป่วย':'#B3261E', 'ลากิจ':'#8A5A00', 'ลาพักร้อน':'#0B5CAD'};
  for(var k=0;k<NR;k++){
    var r=v[k], b='#ffffff', f=C.INK, w='normal', rb=[], rf=[], rw=[], c;
    var isDay=dailyIsDate_(r[0]), type=String(r[2]||'').trim(), note=String(r[6]||'');
    if(isDay){
      nDays=k+1;
      if(type){ b=TINT[type]||C.SOFT; }
      else if(String(r[1]).trim()===DAILY_THAI_WD[0] || /หยุด|อาทิตย์/.test(note)){ b='#F3F4F6'; f='#9AA3AF'; }
    } else if(/^(รวม|ยอดหัก)/.test(type)){ b=C.SOFT; w='bold'; }
    else if(String(r[0]).indexOf('—')===0){ f='#9AA3AF'; }
    for(c=0;c<7;c++){ rb.push(b); rf.push(f); rw.push(w); }
    if(isDay && type){ rf[2]=TCOL[type]||C.NAVY; rw[2]='bold'; rw[4]='bold'; }
    bg.push(rb); fc.push(rf); fw.push(rw);
  }
  var rng=sh.getRange(S,1,NR,7);
  rng.setBackgrounds(bg); rng.setFontColors(fc); rng.setFontWeights(fw);
  rng.setBorder(false,false,false,false,false,false);
  if(nDays) sh.getRange(S,1,nDays,7).setBorder(null,null,true,null,null,true,C.LINE,SpreadsheetApp.BorderStyle.SOLID);
}

function rebuildDailyView_(name, nat){
  rebuildDailyViewCore_(name, nat);
  try{ var sh=sheet_(name); if(sh) dailyStyleRows_(sh); }catch(e){}
}

function buildDailySheet_(name, nat){
  buildDailySheetCore_(name, nat);
  styleTry_(name+': รูปแบบ', function(){ dailyFrame_(sheet_(name)); });
}

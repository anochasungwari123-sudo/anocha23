/** MindScope | Google Apps Script backend
 *  Deploy as a Google Apps Script web app, access restricted to your Workspace domain.
 *  Set Script Properties: ALLOWED_EMAILS (comma-separated), optionally SPREADSHEET_ID.
 *  Runs as the user accessing the app; each user must have Sheet access.
 *  For personal Gmail or configurations where Session.getActiveUser() is blank,
 *  the system FAILS CLOSED. Do not remove authorization to make it work.
 */
const TAB = 'Records';
const HEADERS = ['id','createdAt','visitDate','hn','fullName','age','citizenId','assessor','q2','q9','q8','q2Score','q9Score','q8Score','createdBy'];
const WEIGHTS_8Q = [1,2,6,8,8,9,4,10,4];
function doGet(e){
  try {
    authorized_();
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('MindScope | ระบบคัดกรองโรคซึมเศร้า')
      .addMetaTag('viewport','width=device-width, initial-scale=1');
  } catch(err){
    return HtmlService.createHtmlOutput('<h2>Access denied</h2><p>ไม่สามารถยืนยันสิทธิ์ผู้ใช้งานได้ กรุณาติดต่อผู้ดูแลระบบ</p>').setTitle('Access denied');
  }
}
function authorized_(){
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const prop = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  const allow = prop.split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
  if (!email || !allow.length || !allow.includes(email)) throw new Error('ไม่มีสิทธิ์เข้าใช้ระบบ กรุณาตรวจสอบ ALLOWED_EMAILS และบัญชี Google');
  return email;
}
function sheet_(){
  const p=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss=p?SpreadsheetApp.openById(p):SpreadsheetApp.getActiveSpreadsheet();
  if(!ss)throw new Error('ไม่พบ Google Sheets กรุณาผูก Apps Script กับชีต หรือตั้งค่า SPREADSHEET_ID');
  const sh=ss.getSheetByName(TAB);
  if(!sh)throw new Error('ยังไม่พบชีต Records กรุณาเรียก installSystem ก่อน');
  return sh;
}
function installSystem(){
  authorized_();
  const p=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss=p?SpreadsheetApp.openById(p):SpreadsheetApp.getActiveSpreadsheet();
  if(!ss)throw new Error('ไม่พบ Google Sheet');
  let sh=ss.getSheetByName(TAB);
  if(!sh) sh=ss.insertSheet(TAB);
  if(sh.getLastRow()===0){
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,HEADERS.length).setBackground('#172035').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange('C:C').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
  } else {
    const current=sh.getRange(1,1,1,HEADERS.length).getValues()[0];
    if(JSON.stringify(current)!==JSON.stringify(HEADERS))throw new Error('Records มีหัวตารางไม่ตรงกับระบบ ห้ามเปลี่ยนลำดับคอลัมน์');
  }
  return 'ติดตั้ง Records สำเร็จ';
}
function parseAnswers_(a,count,label,allowNull){
  if(a===null && allowNull)return null;
  if(!Array.isArray(a)||a.length!==count)throw new Error(label+' ไม่ถูกต้อง');
  return a.map(v=>{
    if(v===null||v===undefined||v==='')return null;
    if(typeof v!=='number'||!Number.isInteger(v))throw new Error(label+' มีคำตอบผิดรูปแบบ');
    return v;
  });
}
function allAnswered_(a,skipIndex){return a && a.every((x,i)=>(skipIndex===i)?true:x!==null)}
function summarize_(row){
  const [id,createdAt,visitDate,hn,fullName,age,citizenId,assessor,q2,q9,q8,q2Score,q9Score,q8Score]=row;
  const n=v=>v===''||v===null?null:Number(v);
  return {id:String(id),createdAt:String(createdAt),visitDate:String(visitDate),hn:String(hn),fullName:String(fullName),age:n(age),assessor:String(assessor||''),q2Score:n(q2Score),q9Score:n(q9Score),q8Score:n(q8Score)};
}
function listRecords(){
  authorized_();
  const sh=sheet_();
  const last=sh.getLastRow();if(last<2)return [];
  const rows=sh.getRange(2,1,last-1,HEADERS.length).getDisplayValues();
  return rows.map(summarize_).sort((a,b)=>String(b.visitDate).localeCompare(String(a.visitDate))||String(b.createdAt).localeCompare(String(a.createdAt)));
}
function getRecord(id){
  authorized_();
  if(typeof id!=='string'||!id||id.length>100)throw new Error('ID ไม่ถูกต้อง');
  const sh=sheet_(),last=sh.getLastRow();
  if(last<2)throw new Error('ไม่พบรายการ');
  const vals=sh.getRange(2,1,last-1,HEADERS.length).getDisplayValues();
  const row=vals.find(r=>r[0]===id);
  if(!row)throw new Error('ไม่พบรายการ');
  const out=summarize_(row);
  for(const [i,key] of [[8,'q2'],[9,'q9'],[10,'q8']]){
    try{out[key]=row[i]?JSON.parse(row[i]):null}catch(e){out[key]=null}
  }
  // Citizen ID is intentionally never returned in API responses.
  return out;
}
function saveScreening(p){
  const actor=authorized_();
  if(!p||typeof p!=='object')throw new Error('ข้อมูลไม่ถูกต้อง');
  const asString=(x,max)=>{const s=String(x??'').trim();if(s.length>max)throw new Error('ข้อมูลยาวเกินกำหนด');return s};
  const visitDate=asString(p.visitDate,10),hn=asString(p.hn,40),name=asString(p.fullName,160),citizenId=asString(p.citizenId,13),assessor=asString(p.assessor,120);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(visitDate)||!Number.isFinite(new Date(visitDate).getTime()))throw new Error('วันที่ไม่ถูกต้อง');
  if(!hn||!name)throw new Error('กรุณาระบุ HN และชื่อ');
  const age=Number(p.age);if(p.age===''||!Number.isInteger(age)||age<0||age>120)throw new Error('อายุไม่ถูกต้อง');
  if(citizenId&&!/^\d{13}$/.test(citizenId))throw new Error('เลขบัตรประชาชนต้องเป็นตัวเลข 13 หลัก');
  const q2=parseAnswers_(p.q2,2,'2Q',false),q9=parseAnswers_(p.q9,9,'9Q',true),q8=parseAnswers_(p.q8,9,'8Q',true);
  if(!allAnswered_(q2)||q2.some(v=>v!==0&&v!==1))throw new Error('กรอก 2Q ให้ครบ');
  if(q9 && (!allAnswered_(q9)||q9.some(v=>v<0||v>3)))throw new Error('กรอก 9Q ให้ครบ');
  if(q8){if(q8[2]!==1)q8[3]=null;if(!allAnswered_(q8,q8[2]===1?-1:3)||q8.some(v=>v!==null&&v!==0&&v!==1))throw new Error('กรอก 8Q ให้ครบ');}
  const s2=q2.reduce((a,b)=>a+b,0),s9=q9?q9.reduce((a,b)=>a+b,0):null;
  const s8=q8?q8.reduce((a,b,i)=>a+(b||0)*WEIGHTS_8Q[i],0):null;
  if(s2>0&&!q9)throw new Error('2Q เป็นบวก ต้องทำ 9Q');
  if(s9!==null&&s9>=7&&!q8)throw new Error('9Q ≥7 ต้องทำ 8Q');
  if(q9&&q9[8]>0&&!q8)throw new Error('ข้อ 9 ของ 9Q เป็นบวก ต้องประเมินต่อ');
  const lock=LockService.getScriptLock();lock.waitLock(15000);
  try{
    const row=[Utilities.getUuid(),new Date().toISOString(),visitDate,hn,name,age,citizenId,assessor,JSON.stringify(q2),q9?JSON.stringify(q9):'',q8?JSON.stringify(q8):'',s2,s9===null?'':s9,s8===null?'':s8,actor];
    // Write free-text as literal strings, never allow spreadsheet formulas from untrusted input.
    for(const i of [3,4,6,7])if(typeof row[i]==='string'&&/^[=+\-@]/.test(row[i]))row[i]="'"+row[i];
    sheet_().appendRow(row);
    return {ok:true,id:row[0]};
  }finally{lock.releaseLock()}
}

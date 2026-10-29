/* Backend DPD 2026. Pasang dalam projek Apps Script yang terikat pada fail
   Google Sheets maklum balas. Jangan padam helaian respons terdahulu.
   Jika projek berdiri sendiri, isikan SPREADSHEET_ID. */
const SPREADSHEET_ID = '1QXBoLKR_89TElfi8xCwrtDtLhL3ssQ42Sikz5Rhd4Qc';
const LEGACY_SHEET_NAME = 'RESPON_PRODUK_KMR_PUICE';
const SHEET_NAME = 'JejakKiblat_DPD2026';
const VERSION = 'DPD2026-v1';
const RATINGS = ['Sangat setuju','Setuju','Kurang setuju','Tidak setuju'];
const ROLES = ['MURID','GURU','IBU BAPA/PENJAGA','PENGUNJUNG'];
const QUESTIONS = [
 ['mudah','Mudah digunakan'],['arahan','Arahan jelas'],
 ['menarik','Reka bentuk menarik'],['nfc','Fungsi NFC'],
 ['bantu','Bantu faham'],['digital','Pembelajaran digital'],
 ['sebar','Potensi guna luas'],['puas','Puas hati']
];
const HEADERS = ['ID penghantaran','Timestamp','Tarikh','Masa','Versi','Peranan',...QUESTIONS.map(q=>q[1]),'Persetujuan','Cadangan'];

function doGet(e) {
  const p = e && e.parameter || {};
  // JSONP perlu nama fungsi yang sah. Tolak sebelum sebarang operasi tulis.
  if (p.callback && !/^[A-Za-z_$][\w$]{0,100}$/.test(p.callback)) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Panggilan tidak sah.'})).setMimeType(ContentService.MimeType.JSON);
  }
  let result;
  try {
    if(p.action === 'readProductFeedback') result = readFeedback_();
    else if(p.action === 'submitProductFeedback') result = submitFeedback_(p);
    else result = {status:'error',message:'Tindakan tidak dikenali.',apiVersion:2};
  } catch (error) {
    console.error(error);
    result = {status:'error',message:'Perkhidmatan tidak dapat menyelesaikan permintaan. Sila hubungi petugas.',apiVersion:2};
  }
  const json = JSON.stringify(result).replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
  return ContentService.createTextOutput(p.callback ? p.callback+'('+json+');' : json)
    .setMimeType(p.callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function workbook_(){
  const book = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if(!book)throw new Error('Tetapkan SPREADSHEET_ID untuk projek berdiri sendiri.');
  return book;
}

function doPost(e){
  try {return doGet({parameter:{...JSON.parse(e.postData.contents||'{}'),callback:''}});}
  catch(error){return ContentService.createTextOutput(JSON.stringify({status:'error',message:'Permintaan tidak sah.',apiVersion:2})).setMimeType(ContentService.MimeType.JSON);}
}

function readFeedback_(){
  const data=[];
  const book=workbook_();
  [LEGACY_SHEET_NAME,SHEET_NAME].map(name=>book.getSheetByName(name)).filter(Boolean).forEach(sheet=>{
    if(sheet.getLastRow()<2)return;
    const values=sheet.getDataRange().getDisplayValues();
    const headers=values[0].map(s=>s.trim());
    // Hanya helaian penilaian produk; jangan dedahkan helaian lain.
    if(!['Peranan','Mudah digunakan','Fungsi NFC','Puas hati'].every(h=>headers.includes(h)))return;
    const index=h=>headers.indexOf(h);
    values.slice(1).forEach(row=>{
      const item={Peranan:row[index('Peranan')]||'',Versi:index('Versi')>=0?row[index('Versi')]:'legacy'};
      QUESTIONS.forEach(q=>{
        const v=index(q[1])>=0?row[index(q[1])].trim():'';
        item[q[1]]=RATINGS.includes(v)?v:'';
      });
      if(QUESTIONS.some(q=>item[q[1]]))data.push(item);
    });
  });
  // Nama, kelas, cadangan, ID dan tarikh individu tidak keluar melalui API awam.
  return {status:'success',apiVersion:2,data};
}

function submitFeedback_(p){
  if(p.version!==VERSION||p.consent!=='yes'||!ROLES.includes(p.role))
    return {status:'error',message:'Versi borang, peranan atau persetujuan tidak sah.',apiVersion:2};
  if(!QUESTIONS.every(q=>RATINGS.includes(p[q[0]])))
    return {status:'error',message:'Sila jawab semua soalan.',apiVersion:2};
  if(!/^[A-Za-z0-9-]{10,100}$/.test(p.submissionId||''))
    return {status:'error',message:'ID penghantaran tidak sah.',apiVersion:2};
  const suggestion=String(p.suggestion||'').trim();
  if(suggestion.length>500)return {status:'error',message:'Cadangan melebihi 500 aksara.',apiVersion:2};
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const book=workbook_();let sheet=book.getSheetByName(SHEET_NAME);
    if(!sheet){sheet=book.insertSheet(SHEET_NAME);sheet.appendRow(HEADERS);sheet.setFrozenRows(1);}
    const existing=sheet.getRange(1,1,1,HEADERS.length).getDisplayValues()[0];
    if(!HEADERS.every((h,i)=>existing[i]===h))throw new Error('Lajur helaian DPD berbeza. Jangan tulis pada susunan yang salah.');
    const count=sheet.getLastRow();
    if(count>1&&sheet.getRange(2,1,count-1,1).createTextFinder(p.submissionId).matchEntireCell(true).findNext())
      return {status:'success',apiVersion:2,duplicate:true};
    const now=new Date();
    // Teks bebas tidak boleh menjadi formula Google Sheets.
    const safeSuggestion=/^[=+@-]/.test(suggestion)?"'"+suggestion:suggestion;
    sheet.appendRow([p.submissionId,now.toISOString(),Utilities.formatDate(now,'Asia/Kuala_Lumpur','yyyy-MM-dd'),Utilities.formatDate(now,'Asia/Kuala_Lumpur','HH:mm:ss'),VERSION,p.role,...QUESTIONS.map(q=>p[q[0]]),'Ya',safeSuggestion]);
    SpreadsheetApp.flush();return {status:'success',apiVersion:2};
  } finally {lock.releaseLock();}
}

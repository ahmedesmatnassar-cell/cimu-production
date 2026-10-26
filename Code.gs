// الصق الكود ده في: الشيت > إضافات (Extensions) > Apps Script
const TAB = "يومية PWA";
const HEAD = ["وقت الإرسال","التاريخ","القسم","المشرف","كود العامل","اسم العامل","المكان","المشروع","البند","وصف العمل","ساعات الوردية","ساعات السهرة","معرّف الإرسال"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const batch = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(TAB);
    if (!sh) { sh = ss.insertSheet(TAB); sh.appendRow(HEAD); sh.setFrozenRows(1); sh.setRightToLeft(true); }
    // منع التكرار لو الإرسال اتعاد
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, HEAD.length, last - 1, 1).getValues().flat();
      if (ids.indexOf(batch.id) !== -1) return ContentService.createTextOutput("dup");
    }
    const now = new Date();
    const out = batch.records.map(r => [now, r.date, r.dept, r.supervisor, r.code, r.worker, r.loc, r.project, r.item, r.desc, r.h1, r.h2, batch.id]);
    sh.getRange(sh.getLastRow() + 1, 1, out.length, HEAD.length).setValues(out);
    return ContentService.createTextOutput("ok");
  } finally { lock.releaseLock(); }
}

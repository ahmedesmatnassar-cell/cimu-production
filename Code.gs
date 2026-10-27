// النسخة 2 - استبدل بها الكود القديم كله في Apps Script ثم: نشر > إدارة عمليات النشر > تعديل > إصدار جديد
// ايميل التنبيهات: اكتب هنا أي ايميل (جوجل أو غيره). لأكتر من ايميل افصل بينهم بفاصلة. لو سبته فاضي هيروح لايميل حسابك.
const ALERT_EMAIL = "Ahmed.esmat@sectors-construction.com";
const TAB = "يومية PWA", CFG = "القوائم", REP = "تقارير", IDCOL = 13;
const HEAD = ["وقت الإرسال","التاريخ","القسم","المشرف","كود العامل","اسم العامل","المكان","المشروع","البند","وصف العمل","ساعات العمل","ساعات السهرة","معرّف الإرسال","المُرسِل","معرّف الجهاز","الجهاز"];
const PROJECTS = ["SPA","Mazzej-2","NXT","ICR-7","ICR-6","ICR-5","MCS","Factory"];
const SEED = {
 "النجارة":["عبد الظاهر محمد",[[754,"محمود علي السيد عبد القادر عرفة"],[816,"وجيه إسماعيل حفنى سلطان"],[8025,"محمد فايز صابر جمعة"]]],
 "الماكينات":["عبد الظاهر محمد",[[713,"عبد الظاهر محمد جمعة عبد الله (مشرف)"],[736,"بدير حسني أحمد بنداري"],[742,"أحمد متولي عبد القادر الطنطاوي"]]],
 "الدهانات":["عماد جمعة",[[763,"عماد جمعة حسين رمضان (مشرف)"],[722,"محمود محمد أحمد عبد المقصود"],[8012,"ناصر عويس شفيق محمد"],[8018,"مدبولي محمد مدبولي شافعي"]]],
 "التركيبات":["محمد بيومي",[[759,"محمد بيومي سعد بيومي (مشرف)"],[817,"طارق سيد محمود موسى"],[818,"سلامه سلامه محمد عبد الجليل"],[887,"محمد عبد الرحمن احمد محمد الأسود"],[8022,"اسلام احمد محمد احمد"],[8024,"عبد الله محمد راشد علي منصور"],[8027,"ايمن عابدين حسن عبد اللاه"]]]
};
const SUPCODES = [713,759,763];
const txt = s => ContentService.createTextOutput(s);
const json = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const ss = () => SpreadsheetApp.getActiveSpreadsheet();
const ds = d => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d).slice(0, 10).replace(/\//g, "-");

function tab_() {
  let sh = ss().getSheetByName(TAB);
  if (!sh) { sh = ss().insertSheet(TAB); sh.appendRow(HEAD); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  else if (sh.getRange(1, 14).getValue() === "") sh.getRange(1, 14, 1, 3).setValues([HEAD.slice(13)]);
  return sh;
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents), sh = tab_(), last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, IDCOL, last - 1, 1).getValues().flat().map(String) : [];
    if (b.action === "del") {
      for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === String(b.id)) sh.deleteRow(i + 2);
      return txt("ok");
    }
    if (ids.indexOf(String(b.id)) !== -1) return txt("dup");
    const now = new Date();
    const out = b.records.map(r => { const p = r.date.split("-"); return [now, new Date(+p[0], +p[1] - 1, +p[2]), r.dept, r.supervisor, r.code, r.worker, r.loc, r.project, r.item, r.desc, r.h1, r.h2, b.id, r.submitter || "", r.devId || "", r.device || ""]; });
    sh.getRange(last + 1, 1, out.length, HEAD.length).setValues(out);
    sh.getRange(last + 1, 2, out.length, 1).setNumberFormat("yyyy-mm-dd");
    return txt("ok");
  } finally { lock.releaseLock(); }
}
function doGet(e) {
  const a = (e.parameter || {}).action;
  if (a === "cfg") return json(cfg_());
  if (a === "hist") return json(hist_(e.parameter.dept));
  return txt("ok");
}
function cfgSheet_() {
  let sh = ss().getSheetByName(CFG);
  if (!sh) {
    sh = ss().insertSheet(CFG);
    const rows = [["القسم","الكود","الاسم","مشرف؟ (نعم)","اسم مشرف القسم","","المشاريع"]];
    Object.keys(SEED).forEach(d => SEED[d][1].forEach(w => rows.push([d, w[0], w[1], SUPCODES.indexOf(w[0]) > -1 ? "نعم" : "", SEED[d][0]])));
    sh.getRange(1, 1, rows.length, 5).setValues(rows.map(r => r.slice(0, 5)));
    sh.getRange(1, 7, PROJECTS.length + 1, 1).setValues([["المشاريع"]].concat(PROJECTS.map(p => [p])));
    sh.setFrozenRows(1); sh.setRightToLeft(true);
  }
  return sh;
}
function cfg_() {
  const v = cfgSheet_().getDataRange().getValues().slice(1), depts = {}, sups = [], proj = [];
  v.forEach(r => {
    if (r[0] && r[1] !== "" && r[2]) {
      (depts[r[0]] = depts[r[0]] || { sup: r[4] || "", workers: [] }).workers.push([Number(r[1]), String(r[2])]);
      if (String(r[3]).trim() === "نعم") sups.push(Number(r[1]));
    }
    if (r[6]) proj.push(String(r[6]));
  });
  return { depts: depts, sups: sups, projects: proj };
}
function hist_(dept) {
  const sh = tab_(), last = sh.getLastRow(); if (last < 2) return [];
  const tz = Session.getScriptTimeZone(), v = sh.getRange(2, 1, last - 1, 16).getValues(), lim = ds(new Date(Date.now() - 14 * 864e5)), g = {};
  v.forEach(r => {
    const d = ds(r[1]); if (r[2] !== dept || d < lim) return;
    const id = String(r[12]);
    const o = g[id] = g[id] || { id: id, date: d, t: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "HH:mm") : "", by: r[13], dv: r[15], n: 0, h1: 0, h2: 0, codes: [], rows: [] };
    const a = Number(r[10] || 0), b = Number(r[11] || 0);
    o.n++; o.h1 += a; o.h2 += b; if (o.codes.indexOf(r[4]) < 0) o.codes.push(r[4]);
    o.rows.push({ c: r[4], w: r[5], l: r[6], p: r[7], i: r[8], d: r[9], a: a, b: b });
  });
  return Object.keys(g).map(k => g[k]).sort((x, y) => x.date < y.date ? 1 : x.date > y.date ? -1 : (x.t < y.t ? 1 : -1));
}
// ===== قايمة CIMU في الشيت =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu("CIMU").addItem("إنشاء/تحديث التقارير", "buildReports").addItem("تشغيل تنبيه يومي 5 مساءً", "setupDailyReminder").addItem("فحص المتأخرين الآن", "checkMissing").addToUi();
}
function buildReports() {
  tab_(); let sh = ss().getSheetByName(REP) || ss().insertSheet(REP);
  sh.clear(); sh.setRightToLeft(true);
  sh.getRange("A1:B2").setValues([["من تاريخ", ""], ["إلى تاريخ", ""]]);
  sh.getRange("B1").setFormula("=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"); sh.getRange("B2").setFormula("=TODAY()");
  sh.getRange("B1:B2").setNumberFormat("yyyy-mm-dd");
  const R = "'" + TAB + "'!A:M", W = "\" where B >= date '\"&TEXT($B$1,\"yyyy-mm-dd\")&\"' and B <= date '\"&TEXT($B$2,\"yyyy-mm-dd\")&\"'";
  const q = (sel, extra, grp, lab) => '=QUERY(' + R + ',"select ' + sel + ' where ' + extra + ' and B >= date \'"&TEXT($B$1,"yyyy-mm-dd")&"\' and B <= date \'"&TEXT($B$2,"yyyy-mm-dd")&"\' group by ' + grp + ' label ' + lab + '",1)';
  const L = "sum(K) 'ساعات العمل', sum(L) 'ساعات السهرة'";
  sh.getRange("A4").setFormula(q("F, sum(K), sum(L)", "F is not null", "F", L));
  sh.getRange("E4").setFormula(q("H, sum(K), sum(L)", "H is not null", "H", L));
  sh.getRange("I4").setFormula(q("I, sum(K), sum(L)", "I is not null", "I", L));
  sh.getRange("M4").setFormula(q("F, count(F)", "G = 'أجازة'", "F", "count(F) 'أيام الأجازة'"));
  sh.getRange("A3").setValue("ساعات كل عامل"); sh.getRange("E3").setValue("ساعات كل مشروع"); sh.getRange("I3").setValue("ساعات كل بند"); sh.getRange("M3").setValue("الأجازات");
  sh.getRange("A3:M3").setFontWeight("bold"); sh.setActiveSelection("A1");
}
function checkMissing() {
  const sh = tab_(), last = sh.getLastRow(), today = ds(new Date()), seen = {};
  if (last > 1) sh.getRange(2, 2, last - 1, 4).getValues().forEach(r => { if (ds(r[0]) === today) seen[r[3]] = 1; });
  const miss = {}; let n = 0;
  const c = cfg_(); Object.keys(c.depts).forEach(d => c.depts[d].workers.forEach(w => { if (!seen[w[0]]) { (miss[d] = miss[d] || []).push(w[1]); n++; } }));
  if (!n) return;
  const body = Object.keys(miss).map(d => d + ": " + miss[d].join("، ")).join("\n");
  MailApp.sendEmail(ALERT_EMAIL.trim() || Session.getEffectiveUser().getEmail(), "يومية " + today + " - لم تُسجَّل (" + n + ")", body);
}
function setupDailyReminder() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "checkMissing").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("checkMissing").timeBased().everyDays(1).atHour(17).create();
  SpreadsheetApp.getUi().alert("تم. هيوصلك إيميل كل يوم الساعة 5 مساءً بالعمال اللي لم تُسجَّل ليهم يومية.");
}

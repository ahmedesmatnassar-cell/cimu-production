// النسخة 2 - استبدل بها الكود القديم كله في Apps Script ثم: نشر > إدارة عمليات النشر > تعديل > إصدار جديد
// ايميل التنبيهات: اكتب هنا أي ايميل (جوجل أو غيره). لأكتر من ايميل افصل بينهم بفاصلة. لو سبته فاضي هيروح لايميل حسابك.
// كلمة مرور الإعدادات في التطبيق: اكتبها هنا (لازم تتحدد وإلا الإعدادات متقفلة)
const ADMIN_PASS = "";
const ALERT_EMAIL = "Ahmed.esmat@sectors-construction.com";
const TAB = "يومية PWA", CFG = "القوائم", REP = "تقارير", IDCOL = 13;
const HEAD = ["وقت الإرسال","التاريخ","القسم","المشرف","كود العامل","اسم العامل","المكان","المشروع","البند","وصف العمل","ساعات العمل","ساعات السهرة","معرّف الإرسال","المُرسِل","معرّف الجهاز","الجهاز","آخر تعديل"];
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
  else if (sh.getRange(1, 14).getValue() === "" || sh.getRange(1, 17).getValue() === "") sh.getRange(1, 14, 1, 4).setValues([HEAD.slice(13)]);
  return sh;
}
function sums_(all, skip) { const m = {}; all.forEach((r, i) => { if (i === skip) return; const k = ds(r[1]) + "|" + r[4]; m[k] = (m[k] || 0) + Number(r[10] || 0); }); return m; }
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === "auth") return txt(!ADMIN_PASS ? "unset" : (b.pass === ADMIN_PASS ? "ok" : "no"));
    const admin = !!ADMIN_PASS && b.pass === ADMIN_PASS;
    if (b.action === "cfg_save") { if (!admin) return txt("no"); saveCfg_(b.cfg); return txt("ok"); }
    if (b.action === "dash") { if (!admin) return txt("no"); return json(dash_()); }
    const sh = tab_(), last = sh.getLastRow();
    const all = last > 1 ? sh.getRange(2, 1, last - 1, 13).getValues() : [];
    const ids = all.map(r => String(r[12]));
    if (b.action === "del") {
      for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === String(b.id)) sh.deleteRow(i + 2);
      return txt("ok");
    }
    if (b.action === "edit") {
      const pos = []; ids.forEach((x, i) => { if (x === String(b.id)) pos.push(i); });
      const i0 = pos[b.k]; if (i0 === undefined) return txt("nf");
      const r0 = all[i0], f = b.f, h1 = Number(f.a || 0), m = sums_(all, i0);
      if (h1 < 0 || (m[ds(r0[1]) + "|" + r0[4]] || 0) + h1 > 8) return txt("over:" + r0[5]);
      sh.getRange(i0 + 2, 7, 1, 6).setValues([[f.l, f.p, f.i, f.d, h1, Number(f.b || 0)]]);
      sh.getRange(i0 + 2, 17).setValue(new Date());
      return txt("ok");
    }
    if (ids.indexOf(String(b.id)) !== -1) return txt("dup");
    const today = ds(new Date()), lim = ds(new Date(Date.now() - 3 * 864e5));
    if (b.records.some(r => r.date > today || (r.date < lim && !admin))) return txt("date");
    const tot = sums_(all, -1);
    b.records.forEach(r => { const k = r.date + "|" + r.code; tot[k] = (tot[k] || 0) + Number(r.h1 || 0); });
    const over = []; b.records.forEach(r => { if (tot[r.date + "|" + r.code] > 8 && over.indexOf(r.worker) < 0) over.push(r.worker); });
    if (over.length) return txt("over:" + over.join("، "));
    const now = new Date();
    const out = b.records.map(r => { const p = r.date.split("-"); return [now, new Date(+p[0], +p[1] - 1, +p[2]), r.dept, r.supervisor, r.code, r.worker, r.loc, r.project, r.item, r.desc, r.h1, r.h2, b.id, r.submitter || "", r.devId || "", r.device || ""]; });
    sh.getRange(last + 1, 1, out.length, out[0].length).setValues(out);
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
function saveCfg_(c) {
  const sh = cfgSheet_(); sh.clearContents();
  const rows = [["القسم","الكود","الاسم","مشرف؟ (نعم)","اسم مشرف القسم"]];
  Object.keys(c.depts).forEach(d => c.depts[d].workers.forEach(w => rows.push([d, w[0], w[1], (c.sups || []).indexOf(w[0]) > -1 ? "نعم" : "", c.depts[d].sup || ""])));
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 7, c.projects.length + 1, 1).setValues([["المشاريع"]].concat(c.projects.map(p => [p])));
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
function dash_() {
  const sh = tab_(), last = sh.getLastRow(), c = cfg_(), today = ds(new Date()), mo = today.slice(0, 7);
  const v = last > 1 ? sh.getRange(2, 1, last - 1, 12).getValues() : [], seen = {}, P = {}, L = {};
  v.forEach(r => {
    const d = ds(r[1]), code = r[4];
    if (d === today) { const o = seen[code] = seen[code] || { h: 0, leave: false }; o.h += Number(r[10] || 0) + Number(r[11] || 0); if (r[6] === "أجازة") o.leave = true; }
    if (d.slice(0, 7) === mo) {
      if (r[6] === "أجازة") { (L[code] = L[code] || { w: r[5], days: {} }).days[d] = 1; }
      else if (r[7]) { const o = P[r[7]] = P[r[7]] || { p: r[7], a: 0, b: 0 }; o.a += Number(r[10] || 0); o.b += Number(r[11] || 0); }
    }
  });
  const t = []; Object.keys(c.depts).forEach(d => c.depts[d].workers.forEach(w => { const s = seen[w[0]]; t.push({ d: d, w: w[1], s: !s ? "none" : (s.leave ? "leave" : "ok"), h: s ? s.h : 0 }); }));
  return { date: today, today: t, projects: Object.keys(P).map(k => P[k]).sort((x, y) => y.a - x.a), leaves: Object.keys(L).map(k => ({ w: L[k].w, n: Object.keys(L[k].days).length })).sort((x, y) => y.n - x.n) };
}
function backupNow() {
  const it = DriveApp.getFoldersByName("CIMU-Backups"), folder = it.hasNext() ? it.next() : DriveApp.createFolder("CIMU-Backups");
  DriveApp.getFileById(ss().getId()).makeCopy("CIMU-Backup " + ds(new Date()), folder);
  const files = [], fi = folder.getFiles(); while (fi.hasNext()) files.push(fi.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()); files.slice(14).forEach(f => f.setTrashed(true));
}
function setupBackup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "backupNow").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("backupNow").timeBased().everyDays(1).atHour(2).create();
  SpreadsheetApp.getUi().alert("تم. نسخة احتياطية يومياً الساعة 2 صباحاً في مجلد CIMU-Backups (بيحتفظ بآخر 14 نسخة).");
}
function monthlyReport() {
  const n = new Date(), s = ds(new Date(n.getFullYear(), n.getMonth() - 1, 1)), e = ds(new Date(n.getFullYear(), n.getMonth(), 0));
  const sh = tab_(), last = sh.getLastRow(); if (last < 2) return;
  const W = {}; sh.getRange(2, 1, last - 1, 12).getValues().forEach(r => {
    const d = ds(r[1]); if (d < s || d > e) return;
    const o = W[r[4]] = W[r[4]] || { days: {}, leave: {}, a: 0, b: 0 };
    if (r[6] === "أجازة") o.leave[d] = 1; else { o.days[d] = 1; o.a += Number(r[10] || 0); o.b += Number(r[11] || 0); }
  });
  const c = cfg_(), rows = [["القسم", "الكود", "الاسم", "أيام العمل", "أيام الأجازة", "ساعات العمل", "ساعات السهرة"]];
  Object.keys(c.depts).forEach(d => c.depts[d].workers.forEach(w => { const o = W[w[0]] || { days: {}, leave: {}, a: 0, b: 0 }; rows.push([d, w[0], w[1], Object.keys(o.days).length, Object.keys(o.leave).length, o.a, o.b]); }));
  const csv = "\uFEFF" + rows.map(r => r.map(x => '"' + String(x).replace(/"/g, '""') + '"').join(",")).join("\n");
  const html = "<div dir='rtl'><h3>تقرير ساعات " + s.slice(0, 7) + "</h3><table border='1' cellpadding='4' style='border-collapse:collapse'>" + rows.map((r, i) => "<tr>" + r.map(x => (i ? "<td>" : "<th>") + x + (i ? "</td>" : "</th>")).join("") + "</tr>").join("") + "</table></div>";
  MailApp.sendEmail({ to: ALERT_EMAIL.trim() || Session.getEffectiveUser().getEmail(), subject: "تقرير ساعات " + s.slice(0, 7), htmlBody: html, attachments: [Utilities.newBlob(csv, "text/csv", "hours-" + s.slice(0, 7) + ".csv")] });
}
function setupMonthly() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "monthlyReport").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("monthlyReport").timeBased().onMonthDay(1).atHour(7).create();
  SpreadsheetApp.getUi().alert("تم. هيوصلك تقرير الشهر الماضي بالإيميل يوم 1 من كل شهر الساعة 7 صباحاً.");
}
function buildAttendance() {
  tab_(); cfgSheet_();
  const sh = ss().getSheetByName("الحضور") || ss().insertSheet("الحضور");
  sh.clear(); sh.setRightToLeft(true);
  const T = "'" + TAB + "'!", N = 60, R0 = 6, C = n => { let s = "", x = n; while (x > 0) { const m = (x - 1) % 26; s = String.fromCharCode(65 + m) + s; x = Math.floor((x - 1) / 26); } return s; };
  sh.getRange("A1:A3").setValues([["الشهر (اكتب أول يوم في الشهر)"], ["يوم الإجازة الأسبوعية (1=الأحد ... 6=الجمعة، 7=السبت)"], ["بداية العمل بالنظام (قبلها لا يُحسب غياب)"]]);
  sh.getRange("D1").setFormula("=DATE(YEAR(TODAY()),MONTH(TODAY()),1)"); sh.getRange("D2").setValue(6);
  sh.getRange("D3").setFormula("=IF(COUNT(" + T + "B:B)=0,TODAY(),MIN(" + T + "B:B))");
  sh.getRange("D1").setNumberFormat("yyyy-mm-dd"); sh.getRange("D3").setNumberFormat("yyyy-mm-dd");
  sh.getRange("A5:C5").setValues([["القسم", "الكود", "الاسم"]]);
  sh.getRange(5, 35, 1, 5).setValues([["أيام الحضور", "أيام الأجازة", "أيام الغياب", "ساعات العمل", "ساعات السهرة"]]);
  const f4 = [], f5 = [];
  for (let n = 1; n <= 31; n++) { const col = C(n + 3); f5.push('=IF(DAY(EOMONTH($D$1,0))>=' + n + ',$D$1+' + (n - 1) + ',"")'); f4.push('=IF(' + col + '5="","",TEXT(' + col + '5,"ddd"))'); }
  sh.getRange(4, 4, 1, 31).setFormulas([f4]); sh.getRange(5, 4, 1, 31).setFormulas([f5]).setNumberFormat("d");
  const F = [];
  for (let i = 0; i < N; i++) {
    const r = R0 + i, k = i + 2, row = ['=IF(\'' + CFG + '\'!C' + k + '="","",\'' + CFG + '\'!A' + k + ')', '=IF(\'' + CFG + '\'!C' + k + '="","",\'' + CFG + '\'!B' + k + ')', '=IF(\'' + CFG + '\'!C' + k + '="","",\'' + CFG + '\'!C' + k + ')'];
    for (let n = 1; n <= 31; n++) {
      const L = C(n + 3) + '$5';
      row.push('=IF(OR($B' + r + '="",' + L + '=""),"",IF(COUNTIFS(' + T + '$E:$E,$B' + r + ',' + T + '$B:$B,' + L + ',' + T + '$G:$G,"<>أجازة")>0,"✔",IF(COUNTIFS(' + T + '$E:$E,$B' + r + ',' + T + '$B:$B,' + L + ',' + T + '$G:$G,"أجازة")>0,"أ",IF(OR(' + L + '>TODAY(),' + L + '<$D$3),"",IF(WEEKDAY(' + L + ')=$D$2,"ج","غ")))))');
    }
    const rg = 'D' + r + ':AH' + r, sm = col => '=IF($B' + r + '="","",SUMIFS(' + T + '$' + col + ':$' + col + ',' + T + '$E:$E,$B' + r + ',' + T + '$B:$B,">="&$D$1,' + T + '$B:$B,"<="&EOMONTH($D$1,0)))';
    row.push('=IF($B' + r + '="","",COUNTIF(' + rg + ',"✔"))', '=IF($B' + r + '="","",COUNTIF(' + rg + ',"أ"))', '=IF($B' + r + '="","",COUNTIF(' + rg + ',"غ"))', sm("K"), sm("L"));
    F.push(row);
  }
  sh.getRange(R0, 1, N, 39).setFormulas(F);
  const rng = sh.getRange(R0, 4, N, 31), mk = (t, c) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(c).setRanges([rng]).build();
  sh.setConditionalFormatRules([mk("✔", "#c6efce"), mk("غ", "#ffc7ce"), mk("أ", "#ffeb9c"), mk("ج", "#d9d9d9")]);
  sh.setColumnWidths(4, 31, 30); sh.setColumnWidth(3, 200); sh.setFrozenRows(5); sh.setFrozenColumns(3);
  sh.getRange("A5:AM5").setFontWeight("bold"); sh.getRange("D4:AH4").setHorizontalAlignment("center"); sh.getRange(R0, 4, N, 31).setHorizontalAlignment("center");
  SpreadsheetApp.getUi().alert("تم إنشاء ورقة الحضور. ✔ حضور | أ أجازة | غ غياب | ج إجازة أسبوعية. غيّر الخلية D1 لأي شهر.");
}
// ===== قايمة CIMU في الشيت =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu("CIMU").addItem("إنشاء/تحديث التقارير", "buildReports").addItem("تشغيل تنبيه يومي 5 مساءً", "setupDailyReminder").addItem("فحص المتأخرين الآن", "checkMissing").addItem("إنشاء/تحديث ورقة الحضور والغياب", "buildAttendance").addItem("تشغيل نسخ احتياطي يومي", "setupBackup").addItem("نسخة احتياطية الآن", "backupNow").addItem("تشغيل التقرير الشهري بالإيميل", "setupMonthly").addItem("إرسال تقرير الشهر الماضي الآن", "monthlyReport").addToUi();
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

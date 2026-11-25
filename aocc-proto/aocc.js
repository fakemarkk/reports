/* aocc.js — รายงานประจำวัน AOCC SUP: โมเดล, พรีวิว A4 และไฟล์ Word */
(function (root) {
"use strict";

const M = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
const MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOC_CODE = "VTBS-RP-AOD-002";
const DOC_REV = "Rev00 Effective 01Oct26";
const POS_TH = "หน.พนง.งานกะ ศูนย์ปฏิบัติการ AIRSIDE OPERATION CONTROL CENTER (AOCC)";
const ATTACH = "(ตามเอกสารแนบ)";
const DELAY_SMS = " ให้ผู้บริหาร ทอท. ผ่านระบบโทรศัพท์มือถือข้อความสั้น (SMS) โดยผ่านศูนย์ SCC รับทราบเรียบร้อยแล้ว";
const TITLES = ["นาย","นางสาว","นาง"];
const DEF_PHONE = {brand:"Samsung", model:"A36 EE", storage:"128GB", imei:"3534 5615 809 8167"};
const ITEM_LABELS = [
  "การปฏิบัติงานของ งพบ.สบบ.ฝปข.",
  "การปฏิบัติงานของ งกบ.สบบ.ฝปข.",
  "การปฏิบัติงานของ งบจ.สบข.ฝปข."
];
const EVENT_TYPES = [
  {id:"divert", label:"Divert Flight", line:"เที่ยวบิน Divert Flight"},
  {id:"returnFlight", label:"Return Flight", line:"เที่ยวบิน Return Flight"},
  {id:"delay", label:"Delayed เกิน 3 ชั่วโมง", line:"เที่ยวบิน Delayed เกิน 3 ชั่วโมง"},
  {id:"returnTaxi", label:"Return Taxi", line:"เที่ยวบิน Return Taxi"},
  {id:"techLanding", label:"Technical Landing", line:"เที่ยวบิน Technical Landing"},
  {id:"other", label:"อื่นๆ", line:null}
];
const AGENCY = [
  {id:"fpk", label:"หน.พนง.เวร ฝปข.ทสภ. (06)"},
  {id:"fto", label:"หน.พนง.กะ ฝทอ.ทสภ. (07)"},
  {id:"scc", label:"พนง.ศูนย์สุวรรณภูมิ (SCC)"}
];

function s(v){ return String(v == null ? "" : v).trim(); }
function yy(d){ return String(d.getFullYear() + 543).slice(-2); }
function th(d){ return d.getDate() + " " + M[d.getMonth()] + " " + yy(d); }
function iso(d){ return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0"); }
function parseISO(v){
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s(v));
  if (!m) return null;
  return new Date(+m[1], +m[2]-1, +m[3]);
}
function addDays(d, n){ const x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); return x; }
function shiftInfo(d1, period){
  const night = !String(period || "17.00-08.00").startsWith("08");
  const a = parseISO(d1);
  const b = a ? (night ? addDays(a, 1) : a) : null;
  return {night:night, a:a, b:b, startT: night?"17.00":"08.00", endT: night?"08.00":"17.00", start:a?th(a):"", end:b?th(b):""};
}
function pronounOf(title){
  return title === "นาย" ? "กระผม" : "ดิฉัน";
}
function fullName(title, name){
  title = s(title); name = s(name);
  if (!title) return name;
  if (!name) return title;
  if (name.indexOf(title) === 0) return name;
  if (/^(นาย|นางสาว|นาง)$/.test(title)) return title + name;
  return title + " " + name;
}
function eventType(id){
  for (let i=0;i<EVENT_TYPES.length;i++) if (EVENT_TYPES[i].id===id) return EVENT_TYPES[i];
  return EVENT_TYPES[0];
}
function eventLine(ev, n){
  const typ = eventType(ev.type);
  const cnt = s(ev.count) || "0";
  let head;
  if (typ.id === "other") head = s(ev.other) || "อื่นๆ";
  else head = typ.line;
  let t = n + " " + head + " จำนวน " + cnt + " เที่ยวบิน";
  if (typ.id === "delay") t += DELAY_SMS;
  t += " " + ATTACH;
  return t;
}
function hasDelay(events){
  return (events||[]).some(function(e){ return e && e.type === "delay"; });
}
function normEvents(list){
  return (list||[]).map(function(e){
    return {type:e.type||"divert", count:s(e.count), other:s(e.other)};
  });
}

function defaultModel(){
  return {
    d1:"", period:"17.00-08.00",
    writerTitle:"นาย", writerName:"", pronoun:"กระผม", pronounAuto:true,
    fromTitle:"นาย", fromName:"",
    toTitle:"นาย", toName:"",
    ev1:"เรียบร้อย", note1:"-",
    ev2:"เรียบร้อย", note2:"-",
    ev3:"เรียบร้อย", note3:"-",
    no4:true, events:[],
    agFpkTitle:"นาย", agFpkName:"",
    agFtoTitle:"นาย", agFtoName:"",
    agSccTitle:"", agSccName:"",
    no5:true,
    ev5Title:"นาย", ev5Name:"",
    ev5FromDate:"", ev5FromTime:"", ev5ToDate:"", ev5ToTime:"",
    hiWhere:"ทางขับ สาย K", hiCount:"-",
    loWhere:"หลุมจอดฯ", loCount:"",
    ev5Other:"",
    phone: Object.assign({}, DEF_PHONE)
  };
}

function modelFromDraft(o){
  const m = defaultModel();
  if (!o) return m;
  Object.keys(m).forEach(function(k){
    if (k === "phone" || k === "events") return;
    if (o[k] != null) m[k] = o[k];
  });
  m.no4 = o.no4 !== false;
  m.no5 = o.no5 !== false;
  m.pronounAuto = o.pronounAuto !== false;
  if (m.pronounAuto) m.pronoun = pronounOf(m.writerTitle);
  m.events = m.no4 ? [] : normEvents(o.events);
  if (o.phone && typeof o.phone === "object") m.phone = Object.assign({}, DEF_PHONE, o.phone);
  return m;
}

function missingFields(m){
  const out = [];
  function need(ok, key, label){ if (!ok) out.push({k:key, label:label}); }
  need(!!parseISO(m.d1), "d1", "วันที่เริ่มผลัด");
  need(s(m.writerTitle) && s(m.writerName), "writerName", "ชื่อ AOCC SUP ผู้รายงาน");
  need(s(m.fromTitle) && s(m.fromName), "fromName", "ชื่อผู้มอบหน้าที่ (จาก)");
  need(s(m.toTitle) && s(m.toName), "toName", "ชื่อผู้รับมอบหน้าที่ (ให้กับ)");
  need(s(m.phone.brand), "phoneBrand", "ยี่ห้อโทรศัพท์เวร");
  need(s(m.phone.model), "phoneModel", "รุ่นโทรศัพท์เวร");
  need(s(m.phone.imei), "phoneImei", "หมายเลข IMEI");
  if (!m.no4) {
    const evs = (m.events||[]).filter(function(e){ return e && (e.type || e.count || e.other); });
    if (!evs.length) need(false, "events", "รายการข้อ 4");
    evs.forEach(function(e,i){
      if (!s(e.count)) need(false, "evCount"+i, "จำนวนเที่ยวบินข้อ 4."+(i+1));
      if (e.type === "other" && !s(e.other)) need(false, "evOther"+i, "รายละเอียดข้อ 4."+(i+1));
    });
    if (hasDelay(m.events)) {
      need(s(m.agFpkName), "agFpkName", "ชื่อ หน.พนง.เวร ฝปข.ทสภ. (06)");
      need(s(m.agFtoName), "agFtoName", "ชื่อ หน.พนง.กะ ฝทอ.ทสภ. (07)");
      need(s(m.agSccName), "agSccName", "ชื่อ พนง.ศูนย์สุวรรณภูมิ (SCC)");
    }
  }
  if (!m.no5) {
    need(s(m.ev5Name), "ev5Name", "ชื่อ หน.เวร ฝปข.ทสภ. (06) ข้อ 5");
    need(!!parseISO(m.ev5FromDate) && s(m.ev5FromTime), "ev5FromTime", "วันเวลาเริ่มต้นข้อ 5");
    need(!!parseISO(m.ev5ToDate) && s(m.ev5ToTime), "ev5ToTime", "วันเวลาสิ้นสุดข้อ 5");
  }
  return out;
}

function fileBase(m){
  const S = shiftInfo(m.d1, m.period);
  return "AOCC_SUP_" + (S.a ? iso(S.a).replace(/-/g,"") : "draft") + (S.night ? "_N" : "_D");
}

function phoneText(p){
  p = p || DEF_PHONE;
  const bits = ["ยี่ห้อ " + s(p.brand), "รุ่น " + s(p.model)];
  if (s(p.storage)) bits.push(s(p.storage));
  bits.push("หมายเลขเครื่อง IMEI : " + s(p.imei));
  return bits.join(" ");
}

function openingText(m){
  const S = shiftInfo(m.d1, m.period);
  const who = (m.pronoun || pronounOf(m.writerTitle)) + fullName(m.writerTitle, m.writerName);
  return "ด้วยเมื่อวันที่ " + (S.start || "........") + " เวลา " + S.startT + " น. " + who
    + " ได้รับมอบหน้าที่ " + POS_TH + " จาก " + fullName(m.fromTitle, m.fromName)
    + " ระหว่างปฏิบัติหน้าที่มีเหตุการณ์ควรเรียนให้ทราบดังนี้";
}
function closingText(m){
  const S = shiftInfo(m.d1, m.period);
  const pn = m.pronoun || pronounOf(m.writerTitle);
  return "ครั้นเมื่อวันที่ " + (S.end || "........") + " เวลา " + S.endT + " น. " + pn
    + " ได้มอบหน้าที่ " + POS_TH + " และเครื่องโทรศัพท์เคลื่อนที่ " + phoneText(m.phone)
    + " พร้อมอุปกรณ์ชาร์ตแบตฯ ให้กับ " + fullName(m.toTitle, m.toName)
    + " เพื่ออยู่ปฏิบัติหน้าที่ต่อไป";
}
function itemEvent(m, i){ return s(m["ev"+i]) || "เรียบร้อย"; }
function itemNote(m, i){ return s(m["note"+i]) || "-"; }
function ev4Lines(m){
  if (m.no4) return [];
  return (m.events||[]).map(function(e,i){ return eventLine(e, "4."+(i+1)); });
}
function ev5DatePhrase(m){
  const fd = parseISO(m.ev5FromDate), td = parseISO(m.ev5ToDate);
  return "วันที่ " + (fd?th(fd):"........") + " เวลา " + (s(m.ev5FromTime)||"......") + " น."
    + " ถึงวันที่ " + (td?th(td):"........") + " เวลา " + (s(m.ev5ToTime)||"......") + " น.";
}
function needPage2(m){ return !m.no4 || !m.no5; }

/* ========== HTML preview (A4 @96dpi, margins match the official memo) ========== */
/* 1440/15=96 top, 836/15≈55.73 right, 993/15=66.2 bottom, 1800/15=120 left */
function esc(t){ return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

function footerHTML(){
  return '<div class="a4-foot"><div class="code">'+esc(DOC_CODE)+'</div><div class="rev">'+esc(DOC_REV)+'</div></div>';
}
function tableHTML(m){
  function row(num, label, ev, note, multi){
    return "<tr"+(multi?' class="pair"':'')+"><td class='n'>"+esc(num)+"</td><td>"+label+"</td><td class='c'>"+esc(ev)+"</td><td class='c'>"+esc(note)+"</td></tr>";
  }
  let h = '<table class="rep"><thead><tr><th>ลำดับ</th><th>รายการ</th><th>เหตุการณ์</th><th>หมายเหตุ</th></tr></thead><tbody>';
  h += row("1.", ITEM_LABELS[0], itemEvent(m,1), itemNote(m,1));
  h += row("2.", ITEM_LABELS[1], itemEvent(m,2), itemNote(m,2));
  h += row("3.", ITEM_LABELS[2], itemEvent(m,3), itemNote(m,3));
  h += row("4.", "งานที่ผู้บังคับบัญชามอบหมายและสั่งการ", m.no4?"ไม่มี":"มี", m.no4?"-":"รายละเอียดด้านหลัง");
  h += row("5.", "รายการอนุมัติ/อื่นๆ", m.no5?"ไม่มี":"มี", m.no5?"-":"รายละเอียดด้านหลัง");
  h += "<tr class='spacer'><td></td><td></td><td></td><td></td></tr>";
  return h + "</tbody></table>";
}
function page2HTML(m){
  let h = "";
  if (!m.no4) {
    h += '<p class="h">ตามรายงานข้อ 4</p>';
    ev4Lines(m).forEach(function(line){ h += '<p class="ind">'+esc(line)+'</p>'; });
    if (hasDelay(m.events)) {
      h += '<p>ทั้งนี้ได้รายงานให้หน่วยงานที่เกี่ยวข้องให้ทราบดังนี้</p>';
      h += '<p class="ind">- '+esc(AGENCY[0].label)+"  "+esc(fullName(m.agFpkTitle, m.agFpkName))+"</p>";
      h += '<p class="ind">- '+esc(AGENCY[1].label)+"  "+esc(fullName(m.agFtoTitle, m.agFtoName))+"</p>";
      h += '<p class="ind">- '+esc(AGENCY[2].label)+"  "+esc(fullName(m.agSccTitle, m.agSccName))+"</p>";
    }
  }
  if (!m.no5) {
    h += '<p class="h">รายงานข้อ 5</p>';
    h += '<p class="ind">ได้รายงาน หน.เวร ฝปข.ทสภ. (06) '+esc(fullName(m.ev5Title, m.ev5Name))
      + " เกี่ยวกับข้อมูลอากาศยาน ขออนุญาตทำการทดสอบเครื่องยนต์ " + esc(ev5DatePhrase(m)) + " ดังนี้</p>";
    h += '<p class="ind">- อากาศยานขออนุญาตทำการทดสอบเครื่องยนต์กำลังสูง ณ '+esc(s(m.hiWhere)||"ทางขับ สาย K")
      + " จำนวน " + esc(s(m.hiCount)||"-") + " ลำ</p>";
    h += '<p class="ind">- อากาศยานขออนุญาตทำการทดสอบเครื่องยนต์รอบเดินเบา ณ '+esc(s(m.loWhere)||"หลุมจอดฯ")
      + " จำนวน " + esc(s(m.loCount)||"-") + " ลำ</p>";
    if (s(m.ev5Other)) h += '<p class="ind">'+esc(m.ev5Other)+"</p>";
    h += '<p class="ind att">'+esc(ATTACH)+"</p>";
  }
  return h;
}
function page1Body(m){
  return '<p class="org">ฝ่ายปฏิบัติการเขตการบิน</p>'
    + '<p class="subj"><b>เรื่อง</b>   รายงานประจำวัน</p>'
    + '<p class="subj"><b>เรียน</b>   ผอก.ฝปข.</p>'
    + '<p class="body indent">'+esc(openingText(m))+"</p>"
    + tableHTML(m)
    + '<p class="body indent">'+esc(closingText(m))+"</p>"
    + '<p class="body indent">จึงเรียนมาเพื่อทราบ</p>'
    + '<div class="sigs"><div class="sig">................................................ ผู้ส่ง</div>'
    + '<div class="sig">................................................ ผู้รับ</div></div>';
}
function pageHTML(m){
  const p1 = '<div class="a4 p1">'+page1Body(m)+footerHTML()+"</div>";
  if (!needPage2(m)) return p1;
  return p1 + '<div class="a4 p2">'+page2HTML(m)+footerHTML()+"</div>";
}

/* ========== Word / OOXML ========== */
const WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function xesc(t){ return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function rPr(bold){
  return "<w:rPr><w:rFonts w:ascii=\"TH Sarabun New\" w:hAnsi=\"TH Sarabun New\" w:eastAsia=\"TH Sarabun New\" w:cs=\"TH Sarabun New\"/>"
    + (bold?"<w:b/><w:bCs/>":"")
    + "<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>";
}
function wr(text, bold){
  return "<w:r>"+rPr(bold)+"<w:t xml:space=\"preserve\">"+xesc(text)+"</w:t></w:r>";
}
function pPr(opts){
  opts = opts || {};
  let h = "<w:pPr>";
  if (opts.ind) {
    const a = [];
    if (opts.ind.first) a.push("w:firstLine=\""+opts.ind.first+"\"");
    if (opts.ind.left) a.push("w:left=\""+opts.ind.left+"\"");
    if (opts.ind.right) a.push("w:right=\""+opts.ind.right+"\"");
    h += "<w:ind "+a.join(" ")+"/>";
  }
  if (opts.spc) {
    const a = [];
    if (opts.spc.before) a.push("w:before=\""+opts.spc.before+"\"");
    if (opts.spc.after != null) a.push("w:after=\""+opts.spc.after+"\"");
    if (opts.spc.line) a.push("w:line=\""+opts.spc.line+"\" w:lineRule=\"auto\"");
    h += "<w:spacing "+a.join(" ")+"/>";
  }
  if (opts.jc) h += "<w:jc w:val=\""+opts.jc+"\"/>";
  h += "</w:pPr>";
  return h;
}
function wp(runs, opts){ return "<w:p>"+pPr(opts)+runs+"</w:p>"; }
function wpText(text, opts, bold){ return wp(wr(text, bold), opts); }
function emptyP(opts){ return "<w:p>"+pPr(opts||{})+"</w:p>"; }

function tc(text, opts){
  opts = opts || {};
  const align = opts.jc || "center";
  const vAlign = opts.vAlign || "center";
  const w = opts.w;
  const paras = Array.isArray(text) ? text : [text];
  const body = paras.map(function(t){
    const bold = !!(opts.bold || (t && t.bold));
    const val = (t && typeof t === "object") ? t.t : t;
    return wp(wr(val == null ? "" : val, bold), {jc: (opts.pJc || align)});
  }).join("");
  const nowrap = opts.noWrap ? "<w:noWrap/>" : "";
  return "<w:tc><w:tcPr><w:tcW w:w=\""+w+"\" w:type=\"dxa\"/>"+nowrap+"<w:vAlign w:val=\""+vAlign+"\"/></w:tcPr>"+body+"</w:tc>";
}
function tr(cells, h){
  return "<w:tr><w:trPr><w:trHeight w:val=\""+h+"\"/></w:trPr>"+cells.join("")+"</w:tr>";
}

const COLS = [648, 3780, 1800, 2610];
function tableXml(m){
  const rows = [];
  rows.push(tr([
    tc("ลำดับ", {w:COLS[0], bold:true, noWrap:true}),
    tc("รายการ", {w:COLS[1], bold:true}),
    tc("เหตุการณ์", {w:COLS[2], bold:true}),
    tc("หมายเหตุ", {w:COLS[3], bold:true})
  ], 743));
  // items 1 and 2 share one physical row, like the official samples
  rows.push(tr([
    tc(["1.", "2."], {w:COLS[0], vAlign:"center", pJc:"left"}),
    tc([ITEM_LABELS[0], ITEM_LABELS[1]], {w:COLS[1]}),
    tc([itemEvent(m,1), itemEvent(m,2)], {w:COLS[2]}),
    tc([itemNote(m,1), itemNote(m,2)], {w:COLS[3]})
  ], 804));
  rows.push(tr([
    tc("3.", {w:COLS[0], pJc:"left"}),
    tc(ITEM_LABELS[2], {w:COLS[1]}),
    tc(itemEvent(m,3), {w:COLS[2]}),
    tc(itemNote(m,3), {w:COLS[3]})
  ], 381));
  rows.push(tr([
    tc("4.", {w:COLS[0], pJc:"left"}),
    tc("งานที่ผู้บังคับบัญชามอบหมายและสั่งการ", {w:COLS[1]}),
    tc(m.no4?"ไม่มี":"มี", {w:COLS[2]}),
    tc(m.no4?"-":"รายละเอียดด้านหลัง", {w:COLS[3]})
  ], 375));
  rows.push(tr([
    tc("5.", {w:COLS[0], pJc:"left"}),
    tc("รายการอนุมัติ/อื่นๆ", {w:COLS[1]}),
    tc(m.no5?"ไม่มี":"มี", {w:COLS[2]}),
    tc(m.no5?"-":"รายละเอียดด้านหลัง", {w:COLS[3]})
  ], 369));
  rows.push(tr([
    tc("", {w:COLS[0]}), tc("", {w:COLS[1]}), tc("", {w:COLS[2]}), tc("", {w:COLS[3]})
  ], 369));
  return "<w:tbl><w:tblPr><w:tblW w:w=\"8838\" w:type=\"dxa\"/><w:tblBorders>"
    + "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "</w:tblBorders><w:tblCellMar><w:top w:w=\"40\" w:type=\"dxa\"/><w:left w:w=\"40\" w:type=\"dxa\"/><w:bottom w:w=\"40\" w:type=\"dxa\"/><w:right w:w=\"40\" w:type=\"dxa\"/></w:tblCellMar><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"0\"/>"
    + "</w:tblPr><w:tblGrid>"
    + COLS.map(function(w){ return "<w:gridCol w:w=\""+w+"\"/>"; }).join("")
    + "</w:tblGrid>" + rows.join("") + "</w:tbl>";
}

function bodyXml(m){
  const indent = {ind:{first:"1134"}, jc:"thaiDistribute", spc:{after:"0", line:"240"}};
  const indentL = {ind:{first:"1134"}};
  let xml = "";
  xml += wpText("ฝ่ายปฏิบัติการเขตการบิน", {jc:"center", spc:{after:"0"}}, true);
  xml += wp(wr("เรื่อง", true)+wr("   รายงานประจำวัน"), {spc:{after:"0"}});
  xml += wp(wr("เรียน", true)+wr("   ผอก.ฝปข."), {spc:{after:"0"}});
  xml += emptyP();
  xml += wpText(openingText(m), indent);
  xml += emptyP();
  xml += tableXml(m);
  xml += emptyP();
  xml += wpText(closingText(m), indentL);
  xml += wpText("จึงเรียนมาเพื่อทราบ", indentL);
  xml += emptyP();
  xml += emptyP();
  xml += wpText("................................................ ผู้ส่ง", {jc:"right"});
  xml += emptyP();
  xml += wpText("................................................ ผู้รับ", {jc:"right"});
  if (needPage2(m)) {
    xml += "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>";
    if (!m.no4) {
      xml += wpText("ตามรายงานข้อ 4", {ind:{first:"720"}, spc:{before:"240"}}, true);
      ev4Lines(m).forEach(function(line){
        xml += wpText(line, {ind:{first:"720"}});
      });
      if (hasDelay(m.events)) {
        xml += wpText("ทั้งนี้ได้รายงานให้หน่วยงานที่เกี่ยวข้องให้ทราบดังนี้", {});
        xml += wpText("- "+AGENCY[0].label+"  "+fullName(m.agFpkTitle, m.agFpkName), {ind:{first:"720"}});
        xml += wpText("- "+AGENCY[1].label+"  "+fullName(m.agFtoTitle, m.agFtoName), {ind:{first:"720"}});
        xml += wpText("- "+AGENCY[2].label+"  "+fullName(m.agSccTitle, m.agSccName), {ind:{first:"720"}});
      }
    }
    if (!m.no5) {
      xml += wpText("รายงานข้อ 5", {spc:{before:"240"}}, true);
      xml += wpText("ได้รายงาน หน.เวร ฝปข.ทสภ. (06) "+fullName(m.ev5Title, m.ev5Name)
        + " เกี่ยวกับข้อมูลอากาศยาน ขออนุญาตทำการทดสอบเครื่องยนต์ " + ev5DatePhrase(m) + " ดังนี้",
        {ind:{first:"720"}});
      xml += wpText("- อากาศยานขออนุญาตทำการทดสอบเครื่องยนต์กำลังสูง ณ " + (s(m.hiWhere)||"ทางขับ สาย K")
        + " จำนวน " + (s(m.hiCount)||"-") + " ลำ", {});
      xml += wpText("- อากาศยานขออนุญาตทำการทดสอบเครื่องยนต์รอบเดินเบา ณ " + (s(m.loWhere)||"หลุมจอดฯ")
        + " จำนวน " + (s(m.loCount)||"-") + " ลำ", {});
      if (s(m.ev5Other)) xml += wpText(s(m.ev5Other), {ind:{first:"720"}});
      xml += wpText(ATTACH, {jc:"right"});
    }
  }
  xml += sectPr();
  return xml;
}

function sectPr(){
  // footerReference first (OOXML order), default footer on every page, no titlePg
  return "<w:sectPr>"
    + "<w:footerReference w:type=\"default\" r:id=\"rId5\"/>"
    + "<w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
    + "<w:pgMar w:top=\"1440\" w:right=\"836\" w:bottom=\"993\" w:left=\"1800\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
    + "<w:cols w:space=\"720\"/>"
    + "<w:docGrid w:linePitch=\"360\"/>"
    + "</w:sectPr>";
}

function documentXml(m){
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<w:document xmlns:w=\""+WNS+"\" xmlns:r=\""+RNS+"\">"
    + "<w:body>"+bodyXml(m)+"</w:body></w:document>";
}

function footerXml(){
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<w:ftr xmlns:w=\""+WNS+"\">"
    + wp(wr(DOC_CODE, true), {jc:"right", spc:{after:"0", line:"240"}})
    + wp(wr(DOC_REV, false), {jc:"right", spc:{after:"0", line:"240"}})
    + "</w:ftr>";
}

const CONTENT_TYPES = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
  + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
  + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
  + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
  + "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
  + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
  + "<Override PartName=\"/word/webSettings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml\"/>"
  + "<Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/>"
  + "<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>"
  + "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
  + "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
  + "</Types>";

const RELS_ROOT = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
  + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
  + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
  + "</Relationships>";

const RELS_DOC = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
  + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
  + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>"
  + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings\" Target=\"webSettings.xml\"/>"
  + "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/>"
  + "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>"
  + "</Relationships>";

const STYLES = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<w:styles xmlns:w=\""+WNS+"\">"
  + "<w:docDefaults><w:rPrDefault><w:rPr>"
  + "<w:rFonts w:ascii=\"TH Sarabun New\" w:hAnsi=\"TH Sarabun New\" w:eastAsia=\"TH Sarabun New\" w:cs=\"TH Sarabun New\"/>"
  + "<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>"
  + "</w:rPr></w:rPrDefault>"
  + "<w:pPrDefault><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>"
  + "</w:docDefaults>"
  + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">"
  + "<w:name w:val=\"Normal\"/><w:qFormat/>"
  + "<w:rPr><w:rFonts w:ascii=\"TH Sarabun New\" w:hAnsi=\"TH Sarabun New\" w:eastAsia=\"TH Sarabun New\" w:cs=\"TH Sarabun New\"/>"
  + "<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr></w:style>"
  + "</w:styles>";

const SETTINGS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<w:settings xmlns:w=\""+WNS+"\">"
  + "<w:zoom w:percent=\"100\"/>"
  + "<w:defaultTabStop w:val=\"720\"/>"
  + "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>"
  + "</w:settings>";

const WEB_SETTINGS = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<w:webSettings xmlns:w=\""+WNS+"\"><w:optimizeForBrowser/></w:webSettings>";

const FONT_TABLE = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<w:fonts xmlns:w=\""+WNS+"\">"
  + "<w:font w:name=\"TH Sarabun New\"><w:charset w:val=\"DE\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/></w:font>"
  + "</w:fonts>";

function coreXml(){
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
    + "<dc:title>รายงานประจำวัน AOCC SUP</dc:title>"
    + "<dc:creator></dc:creator>"
    + "<cp:lastModifiedBy></cp:lastModifiedBy>"
    + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">"+now+"</dcterms:created>"
    + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">"+now+"</dcterms:modified>"
    + "</cp:coreProperties>";
}
const APP_XML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
  + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
  + "<Application>Microsoft Office Word</Application>"
  + "<AppVersion>16.0000</AppVersion>"
  + "</Properties>";

async function buildDocx(JSZip, m, type){
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.folder("_rels").file(".rels", RELS_ROOT);
  zip.folder("docProps").file("core.xml", coreXml());
  zip.folder("docProps").file("app.xml", APP_XML);
  const word = zip.folder("word");
  word.folder("_rels").file("document.xml.rels", RELS_DOC);
  word.file("document.xml", documentXml(m));
  word.file("footer1.xml", footerXml());
  word.file("styles.xml", STYLES);
  word.file("settings.xml", SETTINGS);
  word.file("webSettings.xml", WEB_SETTINGS);
  word.file("fontTable.xml", FONT_TABLE);
  return zip.generateAsync({type: type || "blob", mimeType: MIME, compression: "DEFLATE"});
}

root.AoccLib = {
  M:M, MIME:MIME, DOC_CODE:DOC_CODE, DOC_REV:DOC_REV, POS_TH:POS_TH, ATTACH:ATTACH,
  TITLES:TITLES, DEF_PHONE:DEF_PHONE, ITEM_LABELS:ITEM_LABELS, EVENT_TYPES:EVENT_TYPES, AGENCY:AGENCY,
  s:s, th:th, iso:iso, parseISO:parseISO, addDays:addDays, shiftInfo:shiftInfo,
  pronounOf:pronounOf, fullName:fullName, eventType:eventType, eventLine:eventLine,
  hasDelay:hasDelay, defaultModel:defaultModel, modelFromDraft:modelFromDraft,
  missingFields:missingFields, fileBase:fileBase, needPage2:needPage2,
  openingText:openingText, closingText:closingText, ev4Lines:ev4Lines,
  pageHTML:pageHTML, documentXml:documentXml, footerXml:footerXml, buildDocx:buildDocx
};
})(typeof window !== "undefined" ? window : this);

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
  const t = s(title);
  if (t === "นางสาว" || t === "นาง") return "ดิฉัน";
  return "กระผม";
}
function isStdTitle(title){
  return /^(นาย|นางสาว|นาง)$/.test(s(title));
}
function fullName(title, name){
  title = s(title); name = s(name);
  if (!title) return name;
  if (!name) return title;
  if (name.indexOf(title) === 0) return name;
  return title + name;
}
function glueName(title, name){
  const fn = fullName(title, name);
  return fn ? fn.replace(/ /g, "\u00A0") : "";
}
/* Keep a Thai/Latin unit from wrapping mid-string in Word (NBSP for spaces, U+2060 between graphemes). Preview uses CSS nowrap instead. */
const TH_MARK = "[\u0300-\u036f\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]";
function unbreakWord(s){
  s = String(s == null ? "" : s).replace(/ /g, "\u00A0");
  if (!s) return s;
  const re = new RegExp("(?:\u00A0|[^\u00A0\u0300-\u036f\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]" + TH_MARK + "*)", "g");
  const parts = s.match(re);
  if (!parts) return s;
  let out = parts[0];
  for (let i = 1; i < parts.length; i++){
    if (parts[i] !== "\u00A0" && parts[i-1] !== "\u00A0") out += "\u2060";
    out += parts[i];
  }
  return out;
}
function glueNameXml(title, name){
  return unbreakWord(glueName(title, name));
}
function pronounName(pronoun, title, name){
  const fn = glueName(title, name);
  const p = s(pronoun);
  if (!p) return fn;
  // ต้นฉบับขึ้นต้นว่า "ดิฉัน …" — เว้นวรรคหลังสรรพนามเสมอ (กระผม นายทดสอบ / กระผม พ.อ.อ.สมชาย)
  // ชื่อ (คำนำหน้า+ชื่อ+นามสกุล) ติดกันด้วย NBSP ไม่ตัดกลางชื่อ
  return fn ? p + " " + fn : p;
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
    no5:true, ev5Text:"",
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
  need(s(m.writerTitle), "writerTitle", "คำนำหน้า / ยศ ผู้รายงาน");
  need(s(m.writerName), "writerName", "ชื่อ AOCC SUP ผู้รายงาน");
  need(s(m.fromTitle), "fromTitle", "คำนำหน้า / ยศ ผู้มอบหน้าที่");
  need(s(m.fromName), "fromName", "ชื่อผู้มอบหน้าที่ (จาก)");
  need(s(m.toTitle), "toTitle", "คำนำหน้า / ยศ ผู้รับมอบหน้าที่");
  need(s(m.toName), "toName", "ชื่อผู้รับมอบหน้าที่ (ให้กับ)");
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
    if (hasAgency(m.events)) {
      need(s(m.agFpkName), "agFpkName", "ชื่อ หน.พนง.เวร ฝปข.ทสภ. (06)");
      need(s(m.agFtoName), "agFtoName", "ชื่อ หน.พนง.กะ ฝทอ.ทสภ. (07)");
      need(s(m.agSccName), "agSccName", "ชื่อ พนง.ศูนย์สุวรรณภูมิ (SCC)");
    }
  }
  if (!m.no5) {
    need(s(m.ev5Text), "ev5Text", "รายการอนุมัติ/อื่นๆ ข้อ 5");
  }
  return out;
}

function hasAgency(events){
  return (events||[]).some(function(e){ return e && (e.type === "delay" || e.type === "techLanding"); });
}

function fileBase(m){
  const S = shiftInfo(m.d1, m.period);
  return "AOCC_SUP_" + (S.a ? iso(S.a).replace(/-/g,"") : "draft") + (S.night ? "_N" : "_D");
}

function phoneText(p){
  p = p || DEF_PHONE;
  let t = "ยี่ห้อ " + s(p.brand) + " รุ่น " + s(p.model);
  if (s(p.storage)) t += "  " + s(p.storage);
  t += " หมายเลขเครื่อง IMEI : " + s(p.imei);
  return t;
}

function nameForXml(title, name){
  return glueNameXml(title, name) || "........";
}
function openingLead(m, xml){
  const S = shiftInfo(m.d1, m.period);
  const who = xml
    ? (function(){
        const p = s(m.pronoun || pronounOf(m.writerTitle));
        const fn = glueNameXml(m.writerTitle, m.writerName);
        return fn ? (p ? p + " " + fn : fn) : (p || "........");
      })()
    : (pronounName(m.pronoun || pronounOf(m.writerTitle), m.writerTitle, m.writerName) || "........");
  return "ด้วยเมื่อวันที่ " + (S.start || "........") + " เวลา " + S.startT + " น. " + who + " ได้รับมอบหน้าที่";
}
function openingFrom(m, xml){
  const nm = xml ? nameForXml(m.fromTitle, m.fromName) : (glueName(m.fromTitle, m.fromName) || "........");
  return "จาก " + nm + " ระหว่างปฏิบัติหน้าที่มีเหตุการณ์ควรเรียนให้ทราบดังนี้";
}
function openingText(m){
  return openingLead(m, true).replace(/\s+$/,"") + " " + POS_TH + " " + openingFrom(m, true);
}
function closingLead(m){
  const S = shiftInfo(m.d1, m.period);
  const pn = m.pronoun || pronounOf(m.writerTitle);
  return "ครั้นเมื่อวันที่ " + (S.end || "........") + " เวลา " + S.endT + " น. " + pn + "ได้มอบหน้าที่ " + POS_TH + " ";
}
function closingMid(m){
  return "และเครื่องโทรศัพท์เคลื่อนที่ " + phoneText(m.phone)
    + " พร้อมอุปกรณ์ชาร์ตแบตฯ ให้กับ " + nameForXml(m.toTitle, m.toName);
}
function closingText(m){
  return closingLead(m) + closingMid(m) + " เพื่ออยู่ปฏิบัติหน้าที่ต่อไป";
}
function itemEvent(m, i){ return s(m["ev"+i]) || "เรียบร้อย"; }
function itemNote(m, i){ return s(m["note"+i]) || "-"; }
function ev4Lines(m){
  if (m.no4) return [];
  return (m.events||[]).map(function(e,i){ return eventLine(e, "4."+(i+1)); });
}
function ev5Lines(m){
  const raw = String(m && m.ev5Text != null ? m.ev5Text : "").replace(/\r\n/g,"\n").replace(/\r/g,"\n");
  if (!raw.trim()) return [];
  const lines = raw.split("\n");
  if (lines.length && lines[lines.length-1]==="") lines.pop();
  return lines;
}
function needPage2(m){ return !m.no4 || !m.no5; }

/* ========== HTML preview (A4 @96dpi, margins match the official memo) ========== */
/* 1440/15=96 top, 836/15≈55.73 right, 993/15=66.2 bottom, 1800/15=120 left */
function esc(t){ return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

function footerHTML(){
  return '<div class="a4-foot"><div class="code">'+esc(DOC_CODE)+'</div><div class="rev">'+esc(DOC_REV)+'</div></div>';
}
function tableHTML(m){
  function cell(cls, html){ return "<td class='"+cls+"'>"+html+"</td>"; }
  function bodyRow(n, label, ev, note){
    return "<tr>"+cell("n",esc(n))+cell("c",esc(label))+cell("c",esc(ev))+cell("c",esc(note))+"</tr>";
  }
  let h = '<table class="rep"><colgroup><col class="c1"><col class="c2"><col class="c3"><col class="c4"></colgroup>';
  h += "<thead><tr><th>ลำดับ</th><th>รายการ</th><th>เหตุการณ์</th><th>หมายเหตุ</th></tr></thead><tbody>";
  h += bodyRow("1.", ITEM_LABELS[0], itemEvent(m,1), itemNote(m,1));
  h += bodyRow("2.", ITEM_LABELS[1], itemEvent(m,2), itemNote(m,2));
  h += bodyRow("3.", ITEM_LABELS[2], itemEvent(m,3), itemNote(m,3));
  h += bodyRow("4.", "งานที่ผู้บังคับบัญชามอบหมายและสั่งการ", m.no4?"ไม่มี":"มี", m.no4?"-":"รายละเอียดด้านหลัง");
  h += bodyRow("5.", "รายการอนุมัติ/อื่นๆ", m.no5?"ไม่มี":"มี", m.no5?"-":"รายละเอียดด้านหลัง");
  return h + "</tbody></table>";
}
function page2HTML(m){
  let h = "";
  if (!m.no4) {
    h += '<p class="h">ตามรายงานข้อ 4</p>';
    ev4Lines(m).forEach(function(line){ h += '<p class="ind">'+esc(line)+'</p>'; });
    if (hasAgency(m.events)) {
      h += '<p>ทั้งนี้ได้รายงานให้หน่วยงานที่เกี่ยวข้องให้ทราบดังนี้</p>';
      h += '<p class="ind">- '+esc(AGENCY[0].label)+"  "+esc(fullName(m.agFpkTitle, m.agFpkName))+"</p>";
      h += '<p class="ind">- '+esc(AGENCY[1].label)+"  "+esc(fullName(m.agFtoTitle, m.agFtoName))+"</p>";
      h += '<p class="ind">- '+esc(AGENCY[2].label)+"  "+esc(fullName(m.agSccTitle, m.agSccName))+"</p>";
    }
  }
  if (!m.no5) {
    h += '<p class="h">5. รายการอนุมัติ/อื่นๆ</p>';
    ev5Lines(m).forEach(function(line){ h += '<p class="ind">'+(line?esc(line):"&nbsp;")+"</p>"; });
  }
  return h;
}
function nameSpan(title, name){
  const t = glueName(title, name) || "........";
  return '<span class="nm">'+esc(t)+"</span>";
}
function openingHTML(m){
  const S = shiftInfo(m.d1, m.period);
  const p = s(m.pronoun || pronounOf(m.writerTitle));
  return "ด้วยเมื่อวันที่ " + esc(S.start || "........") + " เวลา " + esc(S.startT) + " น. "
    + esc(p) + " " + nameSpan(m.writerTitle, m.writerName)
    + " ได้รับมอบหน้าที่ " + esc(POS_TH)
    + " จาก " + nameSpan(m.fromTitle, m.fromName)
    + " ระหว่างปฏิบัติหน้าที่มีเหตุการณ์ควรเรียนให้ทราบดังนี้";
}
function closingHTML(m){
  const S = shiftInfo(m.d1, m.period);
  const pn = s(m.pronoun || pronounOf(m.writerTitle));
  return "ครั้นเมื่อวันที่ " + esc(S.end || "........") + " เวลา " + esc(S.endT) + " น. "
    + esc(pn) + "ได้มอบหน้าที่ " + esc(POS_TH) + " และเครื่องโทรศัพท์เคลื่อนที่ "
    + esc(phoneText(m.phone)) + " พร้อมอุปกรณ์ชาร์ตแบตฯ ให้กับ "
    + nameSpan(m.toTitle, m.toName)
    + " เพื่ออยู่ปฏิบัติหน้าที่ต่อไป";
}
function page1Body(m){
  return '<p class="org">ฝ่ายปฏิบัติการเขตการบิน</p>'
    + '<p class="subj"><b>เรื่อง</b>   รายงานประจำวัน</p>'
    + '<p class="subj"><b>เรียน</b>   ผอก.ฝปข.</p>'
    + '<p class="spacer">&nbsp;</p>'
    + '<p class="body first">'+openingHTML(m)+"</p>"
    + '<p class="spacer">&nbsp;</p>'
    + tableHTML(m)
    + '<p class="spacer">&nbsp;</p>'
    + '<p class="body first">'+closingHTML(m)+"</p>"
    + '<p class="body first close-end">จึงเรียนมาเพื่อทราบ</p>'
    + '<div class="sigs"><div class="sig send">................................................................ผู้ส่ง</div>'
    + '<div class="sig recv">................................................................ผู้รับ</div></div>';
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
function rFonts(){
  return "<w:rFonts w:ascii=\"TH Sarabun New\" w:hAnsi=\"TH Sarabun New\" w:cs=\"TH Sarabun New\"/>";
}
function pageBreakP(){
  return "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr><w:r>"+rFonts()+"<w:br w:type=\"page\"/></w:r></w:p>";
}
function rPr(bold){
  return "<w:rPr>"+rFonts()+(bold?"<w:b/><w:bCs/>":"")+"<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>";
}
function wr(text, bold){
  return "<w:r>"+rPr(bold)+"<w:t xml:space=\"preserve\">"+xesc(text)+"</w:t></w:r>";
}
function wtab(bold){
  return "<w:r>"+rPr(!!bold)+"<w:tab/></w:r>";
}
function pPr(opts){
  opts = opts || {};
  let h = "<w:pPr>";
  if (opts.style) h += "<w:pStyle w:val=\""+opts.style+"\"/>";
  if (opts.tabs && opts.tabs.length){
    h += "<w:tabs>"+opts.tabs.map(function(t){ return "<w:tab w:val=\""+t.val+"\" w:pos=\""+t.pos+"\"/>"; }).join("")+"</w:tabs>";
  }
  if (opts.spc) {
    const a = [];
    if (opts.spc.before) a.push("w:before=\""+opts.spc.before+"\"");
    if (opts.spc.after != null) a.push("w:after=\""+opts.spc.after+"\"");
    if (opts.spc.line) a.push("w:line=\""+opts.spc.line+"\" w:lineRule=\""+(opts.spc.lineRule||"auto")+"\"");
    h += "<w:spacing "+a.join(" ")+"/>";
  }
  if (opts.ind) {
    const a = [];
    if (opts.ind.left != null) a.push("w:left=\""+opts.ind.left+"\"");
    if (opts.ind.right != null) a.push("w:right=\""+opts.ind.right+"\"");
    if (opts.ind.first != null) a.push("w:firstLine=\""+opts.ind.first+"\"");
    if (opts.ind.hanging != null) a.push("w:hanging=\""+opts.ind.hanging+"\"");
    h += "<w:ind "+a.join(" ")+"/>";
  }
  if (opts.jc) h += "<w:jc w:val=\""+opts.jc+"\"/>";
  const markBold = !!opts.pBold;
  h += "<w:rPr>"+rFonts()+(markBold?"<w:b/><w:bCs/>":"")+"<w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/></w:rPr>";
  h += "</w:pPr>";
  return h;
}
function wp(runs, opts){ return "<w:p>"+pPr(opts)+runs+"</w:p>"; }
function wpText(text, opts, bold){ return wp(wr(text, bold), opts); }
function emptyP(opts){ return "<w:p>"+pPr(opts||{})+"</w:p>"; }

function tcPr(opts){
  opts = opts || {};
  let h = "<w:tcPr><w:tcW w:w=\""+opts.w+"\" w:type=\"dxa\"/>";
  if (opts.header) h += "<w:tcBorders><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tcBorders>";
  if (opts.noWrap) h += "<w:noWrap/>";
  h += "<w:vAlign w:val=\"center\"/>";
  h += "</w:tcPr>";
  return h;
}
function tc(text, opts){
  opts = opts || {};
  const align = opts.jc || "center";
  const paras = Array.isArray(text) ? text : [text];
  const body = paras.map(function(t){
    const bold = !!(opts.bold || (t && t.bold));
    const val = (t && typeof t === "object") ? t.t : t;
    const pOpts = {jc: (opts.pJc || align), pBold: bold};
    if (opts.ind) pOpts.ind = opts.ind;
    return wp(wr(val == null ? "" : val, bold), pOpts);
  }).join("");
  return "<w:tc>"+tcPr(opts)+body+"</w:tc>";
}
function tr(cells, h){
  return "<w:tr><w:trPr><w:trHeight w:val=\""+h+"\"/></w:trPr>"+cells.join("")+"</w:tr>";
}

const COLS = [648, 3780, 1800, 2610];
const ROW_H = {head:743, body:925};
function tableXml(m){
  const rows = [];
  function dataRow(num, label, ev, note){
    return tr([
      tc(num, {w:COLS[0], jc:"center", ind:{left:"-108", right:"-108"}}),
      tc(label, {w:COLS[1], jc:"center"}),
      tc(ev, {w:COLS[2], jc:"center"}),
      tc(note, {w:COLS[3], jc:"center"})
    ], ROW_H.body);
  }
  rows.push(tr([
    tc("ลำดับ", {w:COLS[0], bold:true, header:true, jc:"center", noWrap:true, ind:{left:"-270", right:"-298"}}),
    tc("รายการ", {w:COLS[1], bold:true, header:true, jc:"center", ind:{left:"-108", right:"-108"}}),
    tc("เหตุการณ์", {w:COLS[2], bold:true, header:true, jc:"center", ind:{left:"-108", right:"-108"}}),
    tc("หมายเหตุ", {w:COLS[3], bold:true, header:true, jc:"center", ind:{left:"-108", right:"-108"}})
  ], ROW_H.head));
  rows.push(dataRow("1.", ITEM_LABELS[0], itemEvent(m,1), itemNote(m,1)));
  rows.push(dataRow("2.", ITEM_LABELS[1], itemEvent(m,2), itemNote(m,2)));
  rows.push(dataRow("3.", ITEM_LABELS[2], itemEvent(m,3), itemNote(m,3)));
  rows.push(dataRow("4.", "งานที่ผู้บังคับบัญชามอบหมายและสั่งการ", m.no4?"ไม่มี":"มี", m.no4?"-":"รายละเอียดด้านหลัง"));
  rows.push(dataRow("5.", "รายการอนุมัติ/อื่นๆ", m.no5?"ไม่มี":"มี", m.no5?"-":"รายละเอียดด้านหลัง"));
  return "<w:tbl><w:tblPr><w:tblW w:w=\"8838\" w:type=\"dxa\"/><w:tblBorders>"
    + "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:insideH w:val=\"nil\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>"
    + "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"
    + "</w:tblBorders>"
    + "<w:tblLook w:val=\"01E0\" w:firstRow=\"1\" w:lastRow=\"1\" w:firstColumn=\"1\" w:lastColumn=\"1\" w:noHBand=\"0\" w:noVBand=\"0\"/>"
    + "</w:tblPr><w:tblGrid>"
    + COLS.map(function(w){ return "<w:gridCol w:w=\""+w+"\"/>"; }).join("")
    + "</w:tblGrid>" + rows.join("") + "</w:tbl>";
}

function fillTpl(xml, map){
  Object.keys(map).forEach(function(k){
    xml = xml.split(k).join(xesc(String(map[k] == null ? "" : map[k])));
  });
  return xml;
}
function page1Map(m){
  return {
    "{{OPENING}}": openingText(m),
    "{{CLOSE}}": closingText(m),
    "{{EV1}}": itemEvent(m,1), "{{EV2}}": itemEvent(m,2), "{{EV3}}": itemEvent(m,3),
    "{{NOTE1}}": itemNote(m,1), "{{NOTE2}}": itemNote(m,2), "{{NOTE3}}": itemNote(m,3),
    "{{HAS4}}": m.no4 ? "ไม่มี" : "มี",
    "{{NOTE4}}": m.no4 ? "-" : "รายละเอียดด้านหลัง",
    "{{HAS5}}": m.no5 ? "ไม่มี" : "มี",
    "{{NOTE5}}": m.no5 ? "-" : "รายละเอียดด้านหลัง"
  };
}
function p2line(key, line, P){
  return fillTpl(P[key], {"{{LINE}}": line});
}
function page2Xml(m, P){
  let xml = "";
  if (!m.no4) {
    xml += P["p2_ev4_head.xml"];
    ev4Lines(m).forEach(function(line, i){
      xml += p2line(i === 0 ? "p2_ev4_first.xml" : "p2_ev4_next.xml", line, P);
    });
    if (hasAgency(m.events)) {
      xml += P["p2_agency_intro.xml"];
      xml += p2line("p2_agency_0.xml", "- "+AGENCY[0].label+"  "+fullName(m.agFpkTitle, m.agFpkName), P);
      xml += p2line("p2_agency_1.xml", "- "+AGENCY[1].label+"  "+fullName(m.agFtoTitle, m.agFtoName), P);
      xml += p2line("p2_agency_2.xml", "- "+AGENCY[2].label+"  "+fullName(m.agSccTitle, m.agSccName), P);
    }
    xml += P["p2_ev4_blank.xml"];
  }
  if (!m.no5) {
    xml += P["p2_ev5_head.xml"];
    ev5Lines(m).forEach(function(line){ xml += p2line("p2_ev5_body.xml", line, P); });
  }
  xml += P["p2_tail.xml"];
  return xml;
}
function bodyXml(m){
  const P = ooxmlParts();
  let xml = fillTpl(P["page1.xml"], page1Map(m));
  if (needPage2(m)) xml += pageBreakP() + page2Xml(m, P);
  xml += P["sect.xml"];
  return xml;
}
function sectPr(){
  const P = ooxmlParts();
  return P["sect.xml"];
}

function documentXml(m){
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<w:document xmlns:w=\""+WNS+"\" xmlns:r=\""+RNS+"\">"
    + "<w:body>"+bodyXml(m)+"</w:body></w:document>";
}

function footerXml(){
  const right = {jc:"right", ind:{right:"-334"}, spc:{after:"0", line:"240"}};
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<w:ftr xmlns:w=\""+WNS+"\">"
    + wp(wr(DOC_CODE, true), Object.assign({}, right, {pBold:true}))
    + wp(wr(DOC_REV, false), right)
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
  + "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>"
  + "<Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/>"
  + "<Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/>"
  + "<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>"
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
  + "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>"
  + "<Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>"
  + "<Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/>"
  + "</Relationships>";

function coreXml(){
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
    + "<dc:title>รายงานประจำวัน AOCC SUP</dc:title>"
    + "<dc:creator>ฝ่ายปฏิบัติการเขตการบิน</dc:creator>"
    + "<cp:lastModifiedBy>AOCC SUP</cp:lastModifiedBy>"
    + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">"+now+"</dcterms:created>"
    + "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">"+now+"</dcterms:modified>"
    + "</cp:coreProperties>";
}

function ooxmlParts(){
  const p = (typeof root !== "undefined" && root.AOCC_OOXML_PARTS) ? root.AOCC_OOXML_PARTS : (typeof globalThis !== "undefined" ? globalThis.AOCC_OOXML_PARTS : null);
  if (!p) throw new Error("AOCC_OOXML_PARTS is missing — load word/parts.js before aocc.js");
  return p;
}

async function buildDocx(JSZip, m, type){
  const P = ooxmlParts();
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.folder("_rels").file(".rels", RELS_ROOT);
  zip.folder("docProps").file("core.xml", coreXml());
  zip.folder("docProps").file("app.xml", P["app.xml"] || (
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
    + "<Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>"
  ));
  const word = zip.folder("word");
  word.folder("_rels").file("document.xml.rels", RELS_DOC);
  word.file("document.xml", documentXml(m));
  word.file("footer1.xml", footerXml());
  word.file("styles.xml", P["styles.xml"]);
  word.file("settings.xml", P["settings.xml"]);
  word.file("webSettings.xml", P["webSettings.xml"]);
  word.file("fontTable.xml", P["fontTable.xml"]);
  word.file("footnotes.xml", P["footnotes.xml"]);
  word.file("endnotes.xml", P["endnotes.xml"]);
  word.folder("theme").file("theme1.xml", P["theme1.xml"]);
  return zip.generateAsync({type: type || "blob", mimeType: MIME, compression: "DEFLATE"});
}

root.AoccLib = {
  M:M, MIME:MIME, DOC_CODE:DOC_CODE, DOC_REV:DOC_REV, POS_TH:POS_TH, ATTACH:ATTACH,
  TITLES:TITLES, DEF_PHONE:DEF_PHONE, ITEM_LABELS:ITEM_LABELS, EVENT_TYPES:EVENT_TYPES, AGENCY:AGENCY,
  s:s, th:th, iso:iso, parseISO:parseISO, addDays:addDays, shiftInfo:shiftInfo,
  pronounOf:pronounOf, isStdTitle:isStdTitle, fullName:fullName, glueName:glueName, pronounName:pronounName, eventType:eventType, eventLine:eventLine,
  hasDelay:hasDelay, hasAgency:hasAgency, defaultModel:defaultModel, modelFromDraft:modelFromDraft,
  missingFields:missingFields, fileBase:fileBase, needPage2:needPage2,
  openingText:openingText, closingText:closingText, openingLead:openingLead, openingFrom:openingFrom,
  ev4Lines:ev4Lines, ev5Lines:ev5Lines, pageHTML:pageHTML, documentXml:documentXml, footerXml:footerXml, buildDocx:buildDocx
};
})(typeof window !== "undefined" ? window : this);

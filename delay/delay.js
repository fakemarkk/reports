/* delay.js — ใช้ร่วมกันระหว่างหน้ารายงาน (delay/), หน้า Word (delay/word/) และชุดทดสอบ
   คำนวณเวลาล่าช้า, ข้อความไลน์, เอกสาร Word และหน้าจำลอง A4 */
(function(root){
"use strict";
const M=["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
const NB="\u00a0";
const BKK={name:"BANGKOK",code:"BKK"};
const CAUSES=["AIRCRAFT ROTATION","TECHNICAL PROBLEM","WEATHER","LATE ARRIVAL","OPERATIONAL REQUIREMENT"];
const DEF_POS="หน.พนง.งานกะ ศูนย์ควบคุมปฏิบัติการเขตการบิน AOCC";
// หน่วยงานของหัวหน้า AOCC (หมุนเวียน) ค่าเริ่มต้นตัวแรก
const UNITS=["งพบ.สบบ.ฝปข.ทสภ.", "งกบ.สบบ.ฝปข.ทสภ.", "งบจ.สบข.ฝปข.ทสภ."];
// เปลี่ยนชื่อหน่วยงานที่อยู่ในข้อความตำแหน่ง (ถ้ามี) เป็นหน่วยงานที่เลือก
function swapUnit(t,u){return String(t==null?"":t).split(/(งพบ\.สบบ\.ฝปข\.ทสภ\.|งกบ\.สบบ\.ฝปข\.ทสภ\.|งบจ\.สบข\.ฝปข\.ทสภ\.)/).map((x,i)=>i%2?u:x).join("")}
const s=x=>(x==null?"":String(x)).trim();
const U=x=>s(x).toUpperCase();
const yy=d=>String(d.getFullYear()+543).slice(-2);
// "24 ต.ค. 69"
function th(d){return d.getDate()+" "+M[d.getMonth()]+" "+yy(d)}
function iso(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0")}
function pd(x){return x?new Date(x+"T00:00:00"):null}
function addDays(d,n){const a=new Date(d);a.setDate(a.getDate()+n);return a}
function normTime(x){
  x=s(x).replace(/[:,]/g,".").replace(/[^0-9.]/g,""); if(!x) return "";
  let h,m;
  if(x.includes(".")){[h,m]=x.split(".");}
  else{if(x.length<=2){h=x;m="00"}else{h=x.slice(0,x.length-2);m=x.slice(-2)}}
  h=(h||"0").padStart(2,"0").slice(-2); m=(m||"0").padEnd(2,"0").slice(0,2);
  return h+"."+m;
}
function toMin(x){const t=normTime(x);if(!t)return null;const [h,m]=t.split(".").map(Number);if(h>23||m>59)return null;return h*60+m}
// ---------- ผลัด ----------
function shift(d){
  d=d||{};const night=!String(d.period||"17.00-08.00").startsWith("08");
  const a=pd(s(d.d1)), b=a?(night?addDays(a,1):a):null;
  const p=night?["17.00","08.00"]:["08.00","17.00"];
  let range="";
  if(a){range=!night?th(a):(a.getMonth()===b.getMonth()&&a.getFullYear()===b.getFullYear()?a.getDate()+"-"+th(b):th(a)+" - "+th(b))}
  return {night,a,b,p,range,start:a?th(a):"",end:b?th(b):""};
}
// ---------- เที่ยวบิน ----------
function newFlight(dir){dir=dir==="dep"?"dep":"arr";
  return {dir,flight:"",sched:"",est:"",off:"a",orgName:dir==="dep"?BKK.name:"",org:dir==="dep"?BKK.code:"",dstName:dir==="arr"?BKK.name:"",dst:dir==="arr"?BKK.code:"",cause:""}}
function labels(dir){return dir==="dep"?["SOBT","EOBT"]:["SIBT","EIBT"]}
const isBkk=(n,c)=>U(n)===BKK.name&&U(c)===BKK.code;
const isEmpty=(n,c)=>!s(n)&&!s(c);
// สลับขาเข้า/ขาออก: ย้ายค่าเริ่มต้น BANGKOK/BKK เฉพาะช่องที่ยังเป็นค่าเริ่มต้นหรือว่าง (ไม่ลบค่าที่ผู้ใช้พิมพ์เอง)
function setDir(f,dir){
  if(f.dir===dir) return f;
  const from=dir==="dep"?"dst":"org", to=dir==="dep"?"org":"dst";
  if(isBkk(f[from+"Name"],f[from])&&isEmpty(f[to+"Name"],f[to])){f[from+"Name"]="";f[from]=""}
  if(isEmpty(f[to+"Name"],f[to])&&!isBkk(f[from+"Name"],f[from])){f[to+"Name"]=BKK.name;f[to]=BKK.code}
  f.dir=dir;return f;
}
function hasData(f){const d=newFlight(f.dir);return ["flight","sched","est","cause","orgName","org","dstName","dst"].some(k=>s(f[k])!==s(d[k]))||(f.off&&f.off!=="a")}
function fmtDur(m){const h=Math.floor(m/60),mm=m%60;return h&&mm?h+" ชั่วโมง "+mm+" นาที":h?h+" ชั่วโมง":mm+" นาที"}
// off: "a" = อัตโนมัติ (ถ้าเวลาคาดการณ์น้อยกว่าเวลาตามตาราง บวก 1 วัน), "0" | "1" | "2" = กำหนดเอง
function calcDelay(f){
  const a=toMin(f&&f.sched), b=toMin(f&&f.est);
  if(a==null||b==null) return null;
  const off=f.off==null||f.off===""?"a":String(f.off);
  let min=b-a, days=0, auto=off==="a";
  if(auto){if(min<0){min+=1440;days=1}}
  else{days=+off||0;min+=1440*days}
  if(min<0) return {bad:true,min,days,auto,text:""};
  return {bad:false,min,days,auto,cross:days>0,under:min<180,text:fmtDur(min)};
}
function routeText(f){
  const names=[U(f.orgName),U(f.dstName)].filter(Boolean).join(" - ");
  const codes=U(f.org)&&U(f.dst)?" ("+U(f.org)+"-"+U(f.dst)+")":"";
  return names+codes;
}
// ---------- ข้อความไลน์ ----------
function lineText(d,name,shiftNo,unit){
  d=d||{};const S=shift(d),n=(d.flights&&d.flights.length)||0;
  const when=S.night?"วันที่ "+S.start+" ระหว่างเวลา 17.00 น. ถึงเวลา 08.00 น. ("+S.end+")":"วันที่ "+S.start+" ระหว่างเวลา 08.00 น. ถึงเวลา 17.00 น.";
  return [
"เรียน ผอก.ฝปข.ทสภ. และผู้บริหาร ฝปข.ทสภ. ",
when,
"",
"ขอรายงานเที่ยวบินล่าช้าเกิน 3 ชั่วโมง จำนวน "+n+" เที่ยวบิน พร้อมทั้งได้รายงานให้ผู้บริหาร ทอท.ผ่านศูนย์ SCC เรียบร้อยแล้ว",
"",
"จึงเรียนมาเพื่อทราบ",
s(name)||"XXXXXX",
"หน.พนง.งานกะ "+(UNITS.includes(unit)?unit:UNITS[0])+"(ผลัด"+s(shiftNo)+")"
  ].join("\n");
}
// ---------- แบบฟอร์ม Word ----------
// วันที่ส่ง SMS เริ่มต้น: วันนี้ ถ้ายังอยู่ระหว่างวันเริ่มผลัดถึงหนึ่งวันหลังจบผลัด ไม่งั้นใช้วันจบผลัด
function defaultSmsDate(d,now){
  const S=shift(d);now=now||new Date();if(!S.a) return iso(now);
  const t=pd(iso(now));return (t>=S.a&&t<=addDays(S.b,1))?iso(t):iso(S.b);
}
function nowTime(now){now=now||new Date();return String(now.getHours()).padStart(2,"0")+"."+String(now.getMinutes()).padStart(2,"0")}
// w = {smsDate, smsTime, name, pos, sigDate}
function model(d,w){
  d=d||{};w=w||{};const S=shift(d);
  const fl=(d.flights||[]).map((f,i)=>{const L=labels(f.dir),r=calcDelay(f);
    return {no:i+1,dir:f.dir==="dep"?"dep":"arr",L,flight:U(f.flight),sched:normTime(f.sched),est:normTime(f.est),delay:r&&!r.bad?r.text:"",res:r,route:routeText(f),cause:U(f.cause)}});
  const sd=pd(s(w.smsDate)),gd=pd(s(w.sigDate));
  return {range:S.range,smsTime:normTime(w.smsTime),subj:S.end,p:S.p,count:fl.length,flights:fl,
    name:s(w.name),pos:s(w.pos),sigDate:gd?th(gd):(sd?th(sd):""),smsDate:sd?th(sd):""};
}
// ช่องว่างไว้เขียนเพิ่มเมื่อไม่มีค่า (ไม่ใช้จุดไข่ปลา)
function blank(n){return NB+" ".repeat(Math.max(n-2,1))+NB}
const or=(v,n)=>v||blank(n);
// โครงบรรทัดของเอกสาร: ใช้ร่วมกันระหว่างไฟล์ Word และหน้าจำลอง A4
// k = ชนิดย่อหน้า, t = ข้อความ (สตริง หรือ [เลขข้อ, ข้อความ])
function lines(m){
  const L=[];
  for(let i=0;i<4;i++)L.push({k:"top",t:""});
  L.push({k:"ttl",t:"แบบฟอร์มการจัดส่งข้อมูลให้ผู้บริหาร ทอท."},{k:"ttl",t:"ผ่านระบบโทรศัพท์มือถือ (SMS)"});
  L.push({k:"to",t:"ถึง  ศูนย์รักษาความปลอดภัย ทสภ. (SCC)"});
  // บรรทัดบน = วันที่/เวลาที่ส่งรายงาน (วันเดียว) · บรรทัดเรื่อง = ช่วงวันที่ของผลัด
  L.push({k:"hd",t:"วันที่ "+or(m.smsDate,14)+"  เวลา "+or(m.smsTime,8)+" น."});
  L.push({k:"hd",t:"จาก ศูนย์ควบคุมปฏิบัติการเขตการบิน AOCC (โทร. 24103)"});
  L.push({k:"hd",t:"เรื่อง แจ้งข้อมูลเที่ยวบินล่าช้า  วันที่ "+or(m.range,12)+"  ช่วงเวลา "+m.p[0]+" - "+m.p[1]+" น."});
  L.push({k:"dash",t:"มีเที่ยวบินล่าช้ามากกว่า 3 ชั่วโมง จำนวน "+m.count+" เที่ยวบิน ดังนี้"});
  m.flights.forEach((f,i)=>{
    L.push({k:i?"fl":"fl1",t:[f.no+".","เที่ยวบิน "+or(f.flight,10)+"  "+f.L[0]+" "+or(f.sched,8)+" น. "+f.L[1]+" "+or(f.est,8)+" น."]});
    L.push({k:"sub",t:"รวมเวลาล่าช้า "+or(f.delay,16)});
    L.push({k:"sub",t:"เส้นทางบิน "+or(f.route,30)});
    L.push({k:"subl",t:"ล่าช้าเนื่องจาก "+or(f.cause,26),keep:i===m.flights.length-1}); // เที่ยวบินสุดท้ายไม่แยกหน้าจากช่องลงชื่อ
  });
  L.push({k:"sig",t:"(ลงชื่อ) "+or(m.name,30)},{k:"pos",t:or(m.pos,34)},{k:"date",t:or(m.sigDate,14)});
  return L;
}
// ความกว้างตัวอักษร TH SarabunPSK (ตัวปกติ) หน่วย 1/1000 em
const WA=[216, 147, 208, 403, 361, 585, 423, 120, 190, 190, 285, 411, 162, 216, 162, 270, 362, 362, 362, 362, 362, 362, 362, 362, 362, 362, 162, 162, 411, 411, 411, 283, 536, 400, 378, 406, 431, 351, 351, 425, 441, 147, 264, 376, 353, 548, 441, 486, 378, 487, 379, 352, 379, 466, 390, 588, 418, 366, 424, 196, 262, 196, 412, 352, 204, 344, 401, 331, 401, 374, 206, 311, 390, 143, 155, 316, 200, 601, 390, 398, 401, 401, 217, 282, 238, 390, 341, 507, 318, 337, 321, 208, 153, 208, 416], WT=[386, 378, 382, 393, 393, 408, 294, 367, 377, 380, 384, 519, 519, 425, 425, 343, 461, 532, 543, 391, 391, 378, 430, 335, 420, 428, 428, 381, 381, 447, 447, 425, 400, 375, 322, 378, 381, 425, 335, 393, 438, 381, 427, 454, 387, 372, 391, 357, 0, 316, 316, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 411, 203, 377, 237, 242, 244, 205, 399, 0, 0, 0, 0, 0, 0, 0, 0, 450, 449, 449, 449, 449, 449, 449, 449, 449, 449, 449, 522, 697];
function textW(t){let w=0;for(const ch of String(t)){const c=ch.codePointAt(0);
  w+=c>=0x20&&c<=0x7e?WA[c-0x20]:c===0xa0?216:c>=0xe01&&c<=0xe5b?WT[c-0xe01]:c===0x200b?0:c===0x2013?360:450}
  return w*16/1000*20} // twips ที่ 16pt
// บล็อกลงชื่อ (แบบ C): อยู่ฝั่งขวา จัดกึ่งกลางทุกบรรทัด ถ้าข้อความยาวบล็อกจะขยายไปทางซ้าย
const SIG_TW=9602, SIG_L=3810, SIG_PAD=360;
function sigLeft(m){const L=lines(m).filter(x=>x.k==="sig"||x.k==="pos"||x.k==="date");
  const need=Math.max(...L.map(x=>textW(x.t)))+SIG_PAD;
  return Math.max(0,Math.min(SIG_L,Math.floor(SIG_TW-need)))}
// ระยะย่อหน้า (twips): before/after, ind = เยื้องซ้าย, hang = เยื้องลอย
// gap ก่อน/หลังบรรทัด "- มีเที่ยวบิน..." = 1 บรรทัด 16pt (360tw), ระหว่างเที่ยวบิน = 1 บรรทัดว่าง (382tw)
function spec(k,sigL){return ({
  top:{after:160,bold:true}, ttl:{after:160,bold:true,jc:"center"}, to:{after:160}, hd:{after:20},
  dash:{before:360,after:20,ind:900},
  fl1:{before:360,after:20,ind:1650,hang:240,keep:true}, fl:{before:382,after:20,ind:1650,hang:240,keep:true},
  sub:{after:20,ind:1900,keep:true}, subl:{after:20,ind:1900},
  // บล็อกลงชื่อ: อยู่ในกรอบ (frame) เดียวกัน ยึดขวาล่างของขอบกระดาษ กว้าง = SIG_TW - sigL
  sig:{after:20,jc:"center",keep:true,frame:true}, pos:{after:160,jc:"center",keep:true,frame:true}, date:{after:0,jc:"center",keep:true,frame:true}
})[k]}
// กรอบลงชื่อ: ขวาล่างสุดของพื้นที่พิมพ์ (ชิดขอบล่าง) ข้อความไม่ไหลข้างกรอบ (notBeside) เว้นห่างเนื้อหาด้านบน 382tw
// ทุกย่อหน้าใช้ framePr เหมือนกันทุกตัว จึงรวมเป็นกรอบเดียวทั้งใน Word และ LibreOffice
function framePr(sigL){return '<w:framePr w:w="'+(SIG_TW-sigL)+'" w:hSpace="0" w:vSpace="382" w:wrap="notBeside" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="right" w:yAlign="bottom"/>'}
function xmlEsc(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,"")}
const FONT='<w:rFonts w:ascii="TH SarabunPSK" w:hAnsi="TH SarabunPSK" w:eastAsia="TH SarabunPSK" w:cs="TH SarabunPSK"/>';
function rPr(b){return '<w:rPr>'+FONT+(b?'<w:b/><w:bCs/>':'')+'<w:color w:val="000000"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'}
function run(t,b){return '<w:r>'+rPr(b)+'<w:t xml:space="preserve">'+xmlEsc(t)+'</w:t></w:r>'}
function para(x,sigL){
  const o=spec(x.k,sigL);
  let p='<w:pPr>'+(o.keep||x.keep?'<w:keepNext/>':'')+'<w:keepLines/>'+(o.frame?framePr(sigL):'')+'<w:spacing w:before="'+(o.before||0)+'" w:after="'+(o.after||0)+'" w:line="240" w:lineRule="auto"/>'
    +'<w:ind w:left="'+(o.ind||0)+'" w:right="0"'+(o.hang?' w:hanging="'+o.hang+'"':'')+'/><w:jc w:val="'+(o.jc||"left")+'"/>'+rPr(o.bold)+'</w:pPr>';
  if(Array.isArray(x.t)) p+=run(x.t[0],o.bold)+'<w:r>'+rPr(o.bold)+'<w:tab/></w:r>'+run(x.t[1],o.bold);
  else if(x.t) p+=run(x.t,o.bold);
  return '<w:p>'+p+'</w:p>';
}
// ย่อหน้าว่างสูง 1pt ปิดท้าย: เป็นจุดยึดของกรอบลงชื่อ (LibreOffice ยึดกรอบกับย่อหน้าถัดไป) ถ้าหน้าเต็ม
// เที่ยวบินสุดท้าย+กรอบลงชื่อจะขึ้นหน้าใหม่ด้วยกัน (keepNext) ไม่ทับรายการเที่ยวบิน
const ANCHOR_P='<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:ind w:left="0" w:right="0"/>'+rPr(false)+'</w:pPr></w:p>';
function bodyXml(m){const L=sigLeft(m);return lines(m).map(x=>para(x,L)).join("")+ANCHOR_P}
const MIME="application/vnd.openxmlformats-officedocument.wordprocessingml.document";
async function buildDocx(JSZip,b64,m,type){
  const zip=await JSZip.loadAsync(b64,{base64:true});
  const xml=await zip.file("word/document.xml").async("string");
  zip.file("word/document.xml",xml.replace("{{BODY}}",()=>bodyXml(m)));
  return zip.generateAsync({type:type||"blob",mimeType:MIME,compression:"DEFLATE"});
}
// ชื่อไฟล์ใช้วันที่ในบรรทัด "เรื่อง" (วันจบผลัด)
function fileBase(d){const S=shift(d);return "Delay_"+(S.b?iso(S.b).replace(/-/g,""):"")}
// ---------- หน้าจำลอง A4 (HTML) ----------
function h(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}
function pageHTML(m){
  const L=sigLeft(m)/15,isSig=x=>x.k==="sig"||x.k==="pos"||x.k==="date";
  const p=x=>{const t=Array.isArray(x.t)?'<span class="no">'+h(x.t[0])+'</span>'+h(x.t[1]):(x.t?h(x.t):NB);
    return '<p class="'+x.k+'">'+t+'</p>'};
  const all=lines(m);
  // บล็อกลงชื่อแยกเป็น .sigb ดันลงล่างสุดของหน้า (margin-top:auto) เยื้องซ้ายเท่า sigL ของไฟล์ Word
  return '<div class="a4-body">'+all.filter(x=>!isSig(x)).map(p).join("")
    +'<div class="sigb" style="margin-left:'+L.toFixed(1)+'px">'+all.filter(isSig).map(p).join("")+'</div></div>';
}
const api={M,BKK,CAUSES,DEF_POS,UNITS,swapUnit,th,iso,normTime,toMin,shift,newFlight,labels,setDir,hasData,fmtDur,calcDelay,routeText,lineText,defaultSmsDate,nowTime,model,lines,textW,sigLeft,bodyXml,buildDocx,fileBase,pageHTML,MIME};
if(typeof module!=="undefined"&&module.exports) module.exports=api; else root.DelayLib=api;
})(typeof window!=="undefined"?window:this);

/* rfword.js — ใช้ร่วมกันระหว่างหน้า word/ และชุดทดสอบ: แปลงข้อมูลรายงาน → ค่าในแบบฟอร์ม → เอกสาร Word / หน้าจำลอง A4 */
(function(root){
"use strict";
const M=["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
const NB="\u00a0";
const KEYS=["informer","infTime","airline","fdate","flight","sobt","atot0","rwy1","actype","reg","route","rwy2","reason","ref","crew","pax","bay","aldt","aibt","eobt","atot","other","recName","recPos"];
// ความยาวช่องว่างขีดเส้นใต้เมื่อไม่มีค่า (จำนวน non-breaking space)
const BLANK={informer:26,airline:22,fdate:14,flight:12,route:40,reason:30,ref:30,actype:16,reg:16,other:60,recName:30,recPos:34};
const s=x=>(x==null?"":String(x)).trim();
const U=x=>s(x).toUpperCase();
function th(d){return d.getDate()+" "+M[d.getMonth()]+" "+String(d.getFullYear()+543).slice(-2)}
function addDays(iso,n){const a=new Date(iso+"T00:00:00");a.setDate(a.getDate()+n);return a}
function normTime(x){
  x=s(x).replace(/[:,]/g,".").replace(/[^0-9.]/g,""); if(!x) return "";
  let h,m;
  if(x.includes(".")){[h,m]=x.split(".");}
  else{if(x.length<=2){h=x;m="00"}else{h=x.slice(0,x.length-2);m=x.slice(-2)}}
  h=(h||"0").padStart(2,"0").slice(-2); m=(m||"0").padEnd(2,"0").slice(0,2);
  return h+"."+m;
}
function isNight(d){return s(d&&d.period).startsWith("17")}
// รูปแบบเดียวกับบรรทัด "เส้นทางบิน" ในข้อความไลน์
function routeText(d){
  const names=[U(d.orgName),U(d.dstName)].filter(Boolean).join(" - ");
  return names+(U(d.dst)?" ("+[U(d.org),U(d.dst)].filter(Boolean).join(" - ")+")":"");
}
function daySig(d){return s(d.d1)+"|"+s(d.period)+"|"+s(d.atot0)}
// วันที่ทำการบิน = วันที่ของ ATOT เที่ยวแรก: 1 = วันแรกของผลัด, 2 = วันถัดไป (หลังเที่ยงคืน)
function autoDay(d){
  if(!isNight(d)) return 1;
  const h=parseInt(s(d.atot0).split(".")[0],10);
  return (!isNaN(h)&&h<12)?2:1;
}
function flightDay(d,w){
  if(!isNight(d)) return 1;
  if(w&&(w.fday==1||w.fday==2)&&w.fdaySig===daySig(d)) return +w.fday;
  return autoDay(d);
}
function flightDate(d,w){return s(d.d1)?addDays(s(d.d1),flightDay(d,w)-1):null}
function values(d,w,rec){
  d=d||{};w=w||{};rec=rec||{};
  const more=s(d.kind)==="more";
  const showPax=more||d.paxUnk===false;
  const fd=flightDate(d,w);
  return {
    informer:s(w.informer), infTime:normTime(w.infTime),
    airline:U(d.airline), fdate:fd?th(fd):"",
    flight:U(d.flight), sobt:normTime(w.sobt), atot0:normTime(d.atot0),
    rwy1:U(w.rwy1), actype:U(w.actype), reg:U(d.reg), route:routeText(d), rwy2:U(w.rwy2),
    reason:U(d.reason), ref:s(w.ref),
    crew:showPax?s(d.crew):"", pax:showPax?s(d.pax):"",
    bay:U(d.bay).replace(/^BAY\s*/,""), aldt:normTime(d.aldt), aibt:normTime(d.aibt),
    eobt:more?normTime(d.eobt):(d.unk===false?normTime(d.neobt):""),
    atot:more?normTime(d.atot):"",
    other:(w.other==null?"":String(w.other)).replace(/\r\n?/g,"\n").replace(/^\s+|\s+$/g,""),
    recName:s(rec.name), recPos:s(rec.pos)
  };
}
function padded(k,v){return v?NB+v+NB:NB.repeat(BLANK[k]||12)}
function xmlEsc(t){return t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,"")}
function fillXml(xml,vals){
  return xml.replace(/\{\{(\w+)\}\}/g,(m,k)=>xmlEsc(padded(k,s(vals[k])&&String(vals[k])))
    .split("\n").join(NB+'</w:t><w:br/><w:t xml:space="preserve">'+NB));
}
const MIME="application/vnd.openxmlformats-officedocument.wordprocessingml.document";
async function buildDocx(JSZip,b64,vals,type){
  const zip=await JSZip.loadAsync(b64,{base64:true});
  const xml=await zip.file("word/document.xml").async("string");
  zip.file("word/document.xml",fillXml(xml,vals));
  return zip.generateAsync({type:type||"blob",mimeType:MIME,compression:"DEFLATE"});
}
function fileBase(d,w){
  const fd=flightDate(d||{},w);
  const ds=fd?fd.getFullYear()+String(fd.getMonth()+1).padStart(2,"0")+String(fd.getDate()).padStart(2,"0"):"";
  const fl=U(d&&d.flight).replace(/[^A-Z0-9]/g,"")||"FLIGHT";
  return "ReturnFlight_"+fl+(ds?"_"+ds:"");
}
// ---------- หน้าจำลอง A4 (HTML) สำหรับพรีวิวและบันทึกเป็นรูป ----------
function h(t){return t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}
function pageHTML(vals){
  const u=k=>'<u>'+h(padded(k,s(vals[k])&&String(vals[k]))).split("\n").join(NB+"<br>"+NB)+'</u>';
  const n=(num,html,cls)=>'<p class="n'+(cls?" "+cls:"")+'"><span class="no">'+num+'</span>'+html+'</p>';
  const c=html=>'<p class="c gap">'+html+'</p>';
  return '<div class="a4-body">'
  +'<p class="ttl">แบบฟอร์มการบันทึกข้อมูล <u>Return Flight</u></p>'
  +n("1.",'ผู้แจ้ง '+u("informer")+' เมื่อเวลา '+u("infTime")+' น.',"j")
  +n("2.",'สายการบิน '+u("airline")+' วันที่ทำการบิน '+u("fdate"))
  +n("3.",'เที่ยวบินที่ '+u("flight")+' เวลาออกตามตารางการบิน (SOBT) '+u("sobt")+' น.',"gap")
  +c('เวลาทำการบินขึ้นจริง (ATOT) '+u("atot0")+' น.')
  +n("4.",'ทางวิ่งที่ใช้ในเที่ยวบินแรก '+u("rwy1"))
  +n("5.",'แบบอากาศยาน '+u("actype"))
  +n("6.",'เครื่องหมายสัญชาติและทะเบียนอากาศยาน '+u("reg"))
  +n("7.",'เส้นทางบิน '+u("route"))
  +n("8.",'ทางวิ่งที่ใช้ในเที่ยวขาเข้า '+u("rwy2"))
  +n("9.",'สาเหตุ '+u("reason"))
  +n("9.1.",'อ้างอิงจาก '+u("ref"),"l2")
  +n("10.",'จำนวนลูกเรือ '+u("crew")+' คน ผู้โดยสาร '+u("pax")+' คน')
  +n("11.",'หลุมจอดอากาศยาน '+u("bay")+' เวลาลง '+u("aldt")+' น.',"gap")
  +c('เวลาที่อากาศยานจอด ณ หลุมจอดอากาศยาน (AIBT) '+u("aibt")+' น.')
  +n("12.",'เวลาที่คาดว่าจะออกในเที่ยวที่สอง '+u("eobt")+' น.',"gap")
  +n("13.",'เวลาทำการบินขึ้นจริง (ATOT) '+u("atot")+' น.')
  +n("14.",'ข้อมูลอื่น ๆ เพิ่มเติม '+u("other"))
  +'<p class="e"></p>'
  +'<p class="sig">ผู้บันทึกข้อมูล '+u("recName")+'</p>'
  +'<p class="pos">ตำแหน่ง '+u("recPos")+'</p>'
  +'</div>'
  +'<div class="a4-foot"><p>ศูนย์ควบคุมปฏิบัติการเขตการบิน ฝ่ายปฏิบัติการเขตการบิน ท่าอากาศยานสุวรรณภูมิ  โทร.02-132-4103 แฟกซ์ 02-132-4119</p><p class="r">VTBS-FM-AOD-CA-001</p><p>Rev00 Effective 01Oct26</p></div>';
}
const api={KEYS,BLANK,th,normTime,routeText,daySig,autoDay,flightDay,flightDate,values,fillXml,buildDocx,fileBase,pageHTML,isNight,MIME};
if(typeof module!=="undefined"&&module.exports) module.exports=api; else root.RFWord=api;
})(typeof window!=="undefined"?window:this);

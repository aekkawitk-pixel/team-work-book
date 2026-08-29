(function(){
"use strict";

/* ==========================================================
   1. ค่าคงที่
   ========================================================== */
var MON=["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
var DAYF=["วันอาทิตย์","วันจันทร์","วันอังคาร","วันพุธ","วันพฤหัสบดี","วันศุกร์","วันเสาร์"];
var STATUS={received:"รับเรื่อง",in_progress:"กำลังทำ",blocked:"รอ/ติดปัญหา",
            completed:"เสร็จแล้ว",paused:"พักไว้"};
var STORDER=["received","in_progress","blocked","completed","paused"];
var STREP={received:"ยังไม่เริ่ม",in_progress:"กำลังดำเนินการ",blocked:"รอ/ติดปัญหา",
           completed:"เสร็จแล้ว",paused:"พักไว้"};
var STSYM={received:"○",in_progress:"◐",blocked:"!",completed:"✓",paused:"‖"};
var PRIO={high:"สูง",normal:"ปกติ",low:"ต่ำ"};
var CYCLE={"":"ไม่ต้องรายงาน",weekly:"ทุกสัปดาห์",biweekly:"ทุก 2 สัปดาห์",monthly:"ทุกเดือน"};
var CYSHORT={weekly:"รายงานทุกสัปดาห์",biweekly:"รายงานทุก 2 สัปดาห์",monthly:"รายงานทุกเดือน"};
var KEY="teamWorkBook", OLDKEY="workbook.state.v1", SELF="me", QUIET=4;
var SB=null,USER=null,UID="";
var TITLES={today:"วันนี้",tasks:"งานทั้งหมด",team:"ทีม",projects:"โครงการ",
            week:"สรุปสัปดาห์",print:"พิมพ์รายงาน",settings:"ตั้งค่า"};
var VIEWS=["today","tasks","team","projects","week","print","settings"];
var MOREV=["week","print","settings"];

/* ==========================================================
   2. ตัวช่วย
   ========================================================== */
function $(id){return document.getElementById(id);}
function d2(n){return n<10?"0"+n:""+n;}
function isoOf(d){return d.getFullYear()+"-"+d2(d.getMonth()+1)+"-"+d2(d.getDate());}
function parseD(s){var p=String(s).slice(0,10).split("-");return new Date(+p[0],+p[1]-1,+p[2]);}
function today(){return isoOf(new Date());}
function nowStamp(){var d=new Date();return isoOf(d)+"T"+d2(d.getHours())+":"+d2(d.getMinutes());}
function dOf(s){return (s||"").slice(0,10);}
function tOf(s){return (s||"").length>=16?s.slice(11,16):"";}
function fmtD(s){if(!s)return"—";var d=parseD(s);return d.getDate()+" "+MON[d.getMonth()];}
function fmtDY(s){if(!s)return"—";var d=parseD(s);return d.getDate()+" "+MON[d.getMonth()]+" "+(d.getFullYear()+543);}
function fmtFull(s){var d=parseD(s);return DAYF[d.getDay()]+"ที่ "+fmtDY(s);}
function fmtStamp(s){var t=tOf(s);return fmtDY(s)+(t?" • "+t:"");}
function dayDiff(a,b){return Math.round((parseD(b)-parseD(a))/86400000);}
function addD(s,n){var d=parseD(s);d.setDate(d.getDate()+n);return isoOf(d);}
function addM(s,n){var d=parseD(s),day=d.getDate();d.setDate(1);d.setMonth(d.getMonth()+n);
  var last=new Date(d.getFullYear(),d.getMonth()+1,0).getDate();
  d.setDate(Math.min(day,last));return isoOf(d);}
function weekStart(s){var d=parseD(s),w=(d.getDay()+6)%7;d.setDate(d.getDate()-w);return isoOf(d);}
function monthStart(s){return s.slice(0,8)+"01";}
function monthEnd(s){var d=parseD(s);return isoOf(new Date(d.getFullYear(),d.getMonth()+1,0));}
function uid(){
  if(window.crypto&&crypto.randomUUID)return crypto.randomUUID();
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,function(c){
    var r=Math.random()*16|0,v=c==="x"?r:(r&0x3|0x8);return v.toString(16);});
}
function esc(s){return String(s==null?"":s).replace(/[&<>"']/g,function(c){
  return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
function initials(n){n=(n||"").trim();return n?n.slice(0,2):"?";}
function byId(arr,id){for(var i=0;i<arr.length;i++)if(arr[i].id===id)return arr[i];return null;}
function uniq(a){var s={},o=[];a.forEach(function(x){if(x&&!s[x]){s[x]=1;o.push(x);}});return o;}
function normName(s){return (s||"").replace(/\s+/g," ").trim();}

/* ==========================================================
   2.5 นำเข้ารายชื่อผู้มอบหมายงานจากไฟล์ (.xlsx / .csv / .tsv)
   ========================================================== */
function parseRequesterFile(file){
  var isXlsx=/\.xlsx$/i.test(file.name)||file.type.indexOf("spreadsheet")>=0;
  if(isXlsx)return file.arrayBuffer().then(parseXlsxRequesters);
  return file.text().then(parseDelimitedRequesters);
}
function parseDelimitedRequesters(text){
  var lines=text.replace(/\r/g,"").split("\n").filter(function(l){return l.trim()!=="";});
  if(!lines.length)return [];
  var delim=lines[0].indexOf("\t")>=0?"\t":",";
  var rows=lines.map(function(l){return l.split(delim).map(function(c){
    return c.replace(/^"|"$/g,"").trim();});});
  return rowsToRequesters(rows);
}
function rowsToRequesters(rows){
  if(!rows.length)return [];
  var head=rows[0].map(function(h){return (h||"").trim();});
  var iName=-1,iDept=-1,iPos=-1;
  head.forEach(function(h,i){
    if(iName<0&&h.indexOf("ชื่อ")>=0)iName=i;
    if(iDept<0&&h.indexOf("แผนก")>=0)iDept=i;
    if(iPos<0&&h.indexOf("ตำแหน่ง")>=0)iPos=i;
  });
  var start=1;
  if(iName<0){iName=0;iDept=1;iPos=2;start=0;}
  var out=[];
  for(var i=start;i<rows.length;i++){
    var row=rows[i];
    var name=normName(row[iName]||"");
    if(!name)continue;
    out.push({name:name,department:iDept>=0?(row[iDept]||"").trim():"",
      position:iPos>=0?(row[iPos]||"").trim():""});
  }
  return out;
}
/* ---- xlsx = zip + xml — อ่านชีตแรกโดยไม่พึ่งไลบรารีภายนอก ---- */
function bytesToStr(bytes){return new TextDecoder("utf-8").decode(bytes);}
function xmlDecode(s){
  return (s||"").replace(/&#x([0-9a-fA-F]+);/g,function(_,h){return String.fromCodePoint(parseInt(h,16));})
    .replace(/&#(\d+);/g,function(_,d){return String.fromCodePoint(parseInt(d,10));})
    .replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"')
    .replace(/&apos;/g,"'").replace(/&amp;/g,"&");
}
function colToIndex(letters){
  var n=0;
  for(var i=0;i<letters.length;i++)n=n*26+(letters.charCodeAt(i)-64);
  return n-1;
}
function parseSharedStrings(xml){
  var out=[],siRe=/<si[^>]*>([\s\S]*?)<\/si>/g,m;
  while((m=siRe.exec(xml))){
    var t="",tRe=/<t[^>]*>([\s\S]*?)<\/t>/g,tm;
    while((tm=tRe.exec(m[1])))t+=tm[1];
    out.push(xmlDecode(t));
  }
  return out;
}
function parseSheetXml(xml,sst){
  var rows=[],rowRe=/<row\b[^>]*>([\s\S]*?)<\/row>/g,rm;
  var cellRe=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  while((rm=rowRe.exec(xml))){
    var cells=[],cm;
    cellRe.lastIndex=0;
    while((cm=cellRe.exec(rm[1]))){
      var attrs=cm[1],content=cm[2]||"";
      var rMatch=/r="([A-Z]+)\d+"/.exec(attrs);
      var col=rMatch?colToIndex(rMatch[1]):cells.length;
      var tMatch=/t="([^"]+)"/.exec(attrs),type=tMatch?tMatch[1]:"",val="";
      if(type==="inlineStr"){
        var isM=/<is>([\s\S]*?)<\/is>/.exec(content);
        if(isM){var tR=/<t[^>]*>([\s\S]*?)<\/t>/g,tm2,acc="";
          while((tm2=tR.exec(isM[1])))acc+=tm2[1];
          val=xmlDecode(acc);}
      }else{
        var vM=/<v>([\s\S]*?)<\/v>/.exec(content),raw=vM?vM[1]:"";
        val=type==="s"?(sst[parseInt(raw,10)]||""):xmlDecode(raw);
      }
      cells[col]=val;
    }
    var out=[];
    for(var i=0;i<cells.length;i++)out[i]=cells[i]||"";
    rows.push(out);
  }
  return rows;
}
function extractZipEntry(u8,entry){
  var dv=new DataView(u8.buffer,u8.byteOffset,u8.byteLength);
  var lp=entry.localOffset;
  var nameLen=dv.getUint16(lp+26,true),extraLen=dv.getUint16(lp+28,true);
  var dataStart=lp+30+nameLen+extraLen;
  var data=u8.subarray(dataStart,dataStart+entry.compSize);
  if(entry.method===0)return Promise.resolve(data.buffer.slice(data.byteOffset,data.byteOffset+data.byteLength));
  var stream=new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Response(stream).arrayBuffer();
}
function parseXlsxRequesters(buf){
  var u8=new Uint8Array(buf),dv=new DataView(buf);
  var eocd=-1,minI=Math.max(0,u8.length-22-65557);
  for(var i=u8.length-22;i>=minI;i--){
    if(dv.getUint32(i,true)===0x06054b50){eocd=i;break;}
  }
  if(eocd<0)throw new Error("ไม่ใช่ไฟล์ xlsx ที่ถูกต้อง");
  var cdOffset=dv.getUint32(eocd+16,true),cdCount=dv.getUint16(eocd+10,true);
  var entries={},p=cdOffset;
  for(var e=0;e<cdCount;e++){
    if(dv.getUint32(p,true)!==0x02014b50)break;
    var method=dv.getUint16(p+10,true),compSize=dv.getUint32(p+20,true);
    var nameLen=dv.getUint16(p+28,true),extraLen=dv.getUint16(p+30,true),commentLen=dv.getUint16(p+32,true);
    var lho=dv.getUint32(p+42,true);
    var name=bytesToStr(u8.subarray(p+46,p+46+nameLen));
    entries[name]={method:method,compSize:compSize,localOffset:lho};
    p+=46+nameLen+extraLen+commentLen;
  }
  var sheetName=Object.keys(entries).filter(function(n){
    return /^xl\/worksheets\/sheet\d+\.xml$/.test(n);}).sort()[0];
  if(!sheetName)throw new Error("ไม่พบชีตข้อมูลในไฟล์");
  return Promise.all([
    extractZipEntry(u8,entries[sheetName]),
    entries["xl/sharedStrings.xml"]?extractZipEntry(u8,entries["xl/sharedStrings.xml"]):Promise.resolve(null)
  ]).then(function(r){
    var sheetXml=bytesToStr(new Uint8Array(r[0]));
    var sst=r[1]?parseSharedStrings(bytesToStr(new Uint8Array(r[1]))):[];
    return rowsToRequesters(parseSheetXml(sheetXml,sst));
  });
}

/* ==========================================================
   3. ข้อมูล
   ========================================================== */
function blankDB(){
  return {version:2,updatedAt:null,
    settings:{owner:"",theme:"system",defaultDueDays:"",defaultReportCycle:"",
      recent:{member:[],project:[],requester:[]}},
    members:[],requesters:[],projects:[],tasks:[],notes:[],weeks:{}};
}
var DB=blankDB();

function normalize(db){
  if(!db||typeof db!=="object")return null;
  var b=blankDB();
  db.settings=db.settings||{};
  for(var k in b.settings)if(db.settings[k]===undefined)db.settings[k]=b.settings[k];
  var r=db.settings.recent||{};
  db.settings.recent={member:r.member||[],project:r.project||[],requester:r.requester||[]};
  ["members","requesters","projects","tasks","notes"].forEach(function(k){
    if(!Array.isArray(db[k]))db[k]=[];});
  db.weeks=db.weeks||{};
  db.members.forEach(function(m){m.nickname=m.nickname||"";m.active=m.active!==false;
    m.department=m.department||"";m.position=m.position||"";});
  db.requesters.forEach(function(x){x.active=x.active!==false;x.department=x.department||"";x.position=x.position||"";});
  db.projects.forEach(function(p){p.description=p.description||"";p.status=p.status||"active";
    p.startAt=p.startAt||"";p.endAt=p.endAt||"";p.archived=!!p.archived;});
  db.tasks.forEach(function(t){
    t.description=t.description||"";
    t.collaboratorIds=Array.isArray(t.collaboratorIds)?t.collaboratorIds:[];
    t.collaboratorIds=t.collaboratorIds.filter(function(x){return x!==t.ownerId;});
    t.status=STATUS[t.status]?t.status:"received";
    t.priority=PRIO[t.priority]?t.priority:"normal";
    t.receivedAt=t.receivedAt||today();
    t.dueAt=t.dueAt||"";
    t.reportCycle=CYCLE[t.reportCycle]?t.reportCycle:"";
    t.repeatNext=!!t.repeatNext;
    t.createdAt=t.createdAt||t.receivedAt;
    t.updatedAt=t.updatedAt||t.createdAt;
    t.completedAt=t.completedAt||"";
    t.updates=Array.isArray(t.updates)?t.updates:[];
    t.updates.forEach(function(u){u.createdAt=u.createdAt||today();u.taskId=t.id;});
  });
  return db;
}

/* แปลงข้อมูลรุ่นเดิม (v1) ให้เป็นโครงสร้างใหม่ */
function migrate(o){
  if(!o||typeof o!=="object")return null;
  if(o.version===2||o.settings)return normalize(o);
  var wasV1=true;
  if(!o.tasks&&!o.members&&!o.owner)return null;
  var db=blankDB();
  db.updatedAt=o.updatedAt||null;
  db.settings.owner=o.owner||"";
  db.members=(o.members||[]).map(function(m){
    return {id:m.id||uid(),name:m.name||"",nickname:"",active:m.active!==false};});
  var rmap={},pmap={};
  function req(name){
    if(!name)return "";
    if(!rmap[name]){var r={id:uid(),name:name,active:true};db.requesters.push(r);rmap[name]=r.id;}
    return rmap[name];
  }
  function prj(name){
    if(!name)return "";
    if(!pmap[name]){var p={id:uid(),name:name,description:"",status:"active",
      startAt:"",endAt:"",archived:false};db.projects.push(p);pmap[name]=p.id;}
    return pmap[name];
  }
  (o.sources||[]).forEach(req);
  (o.projects||[]).forEach(prj);
  var SM={"new":"received",doing:"in_progress",waiting:"blocked",done:"completed",hold:"paused"};
  var CM={w:"weekly","2w":"biweekly",m:"monthly"};
  (o.tasks||[]).forEach(function(t){
    var id=t.id||uid();
    db.tasks.push({
      id:id, title:t.title||"(ไม่มีชื่อ)", description:t.detail||"",
      requesterId:req(t.from), ownerId:t.owner||SELF,
      collaboratorIds:(t.helpers||[]).slice(), projectId:prj(t.project),
      status:SM[t.status]||"received", priority:PRIO[t.priority]?t.priority:"normal",
      receivedAt:(t.received||today())+(t.recvTime?"T"+t.recvTime:""),
      dueAt:t.due||"", reportCycle:CM[t.cycle]||"", repeatNext:!!t.repeat,
      createdAt:t.createdAt||t.received||today(),
      updatedAt:t.received||today(), completedAt:t.doneAt||"",
      updates:(t.updates||[]).map(function(u){
        return {id:u.id||uid(),taskId:id,message:u.text||"",
          status:SM[u.status]||"", createdAt:u.date||today()};})
    });
  });
  (o.notes||[]).forEach(function(n){
    db.notes.push({id:n.id||uid(),date:n.date||today(),text:n.text||"",projectId:prj(n.project)});});
  db.weeks=o.weeks||{};
  var out=normalize(db);
  if(out&&wasV1)out.__v1=true;
  return out;
}

/* ==========================================================
   4. ชั้นบันทึกข้อมูล — Supabase
   ========================================================== */
var syncTimer=null,syncing=false,dirty=false,snap=null,bootDone=false;
function setSave(kind,msg){
  ["dot-d","dot-m"].forEach(function(i){var e=$(i);if(e)e.className="dot"+(kind?" "+kind:"");});
  ["save-d","save-m"].forEach(function(i){var e=$(i);if(e)e.textContent=msg;});
}
function saveLocal(){try{localStorage.setItem(KEY+":"+UID,JSON.stringify(DB));}catch(e){}}
function touch(){
  DB.updatedAt=new Date().toISOString();
  saveLocal();
  if(!SB||!UID){setSave("wait","ยังไม่ได้เข้าสู่ระบบ");return;}
  dirty=true;setSave("wait","กำลังบันทึก…");
  clearTimeout(syncTimer);syncTimer=setTimeout(syncNow,700);
}
function clone(x){return JSON.parse(JSON.stringify(x));}
function localISO(s){
  if(!s)return null;
  var d=s.length>=16?new Date(s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10),+s.slice(11,13),+s.slice(14,16))
                    :new Date(+s.slice(0,4),+s.slice(5,7)-1,+s.slice(8,10),9,0);
  return d.toISOString();
}
function fromISO(iso){
  if(!iso)return today();
  var d=new Date(iso);
  return isoOf(d)+"T"+d2(d.getHours())+":"+d2(d.getMinutes());
}
/* ---------- รูปแบบแถวของแต่ละตาราง ---------- */
function rowMember(m,i){return {id:m.id,owner_id:UID,name:m.name,nickname:m.nickname||"",
  department:m.department||"",position:m.position||"",
  active:m.active!==false,is_self:m.id===SELF,sort_order:i};}
function rowRequester(r,i){return {id:r.id,owner_id:UID,name:r.name,department:r.department||"",
  position:r.position||"",active:r.active!==false,sort_order:i};}
function rowProject(p,i){return {id:p.id,owner_id:UID,name:p.name,description:p.description||"",
  status:p.status||"active",start_at:p.startAt||null,end_at:p.endAt||null,
  archived:!!p.archived,sort_order:i};}
function rowTask(t){return {id:t.id,owner_id:UID,title:t.title,description:t.description||"",
  requester_id:t.requesterId||null,
  assignee_id:(t.ownerId&&t.ownerId!==SELF)?t.ownerId:(SELF||null),
  project_id:t.projectId||null,status:t.status,priority:t.priority,
  received_at:localISO(t.receivedAt),due_at:t.dueAt||null,
  report_cycle:t.reportCycle||"",repeat_next:!!t.repeatNext,
  completed_at:t.completedAt||null};}
function rowNote(n){return {id:n.id,owner_id:UID,note_date:n.date,body:n.text||"",
  project_id:n.projectId||null};}
function rowUpdate(u){return {id:u.id,task_id:u.taskId,owner_id:UID,message:u.message||"",
  status:u.status||"",created_at:localISO(u.createdAt)};}
function flatUpdates(db){
  var o=[];db.tasks.forEach(function(t){(t.updates||[]).forEach(function(u){
    o.push({id:u.id,taskId:t.id,message:u.message,status:u.status,createdAt:u.createdAt});});});
  return o;
}
function flatCollabs(db){
  var o=[];db.tasks.forEach(function(t){collabs(t).forEach(function(mid){
    o.push({k:t.id+"|"+mid,task_id:t.id,member_id:mid,owner_id:UID});});});
  return o;
}
function diffRows(oldArr,newArr,keyFn,rowFn){
  var om={},ups=[],dels=[],seen={};
  (oldArr||[]).forEach(function(x,i){om[keyFn(x)]=JSON.stringify(rowFn(x,i));});
  newArr.forEach(function(x,i){
    var k=keyFn(x);seen[k]=1;
    var r=rowFn(x,i);
    if(om[k]!==JSON.stringify(r))ups.push(r);
  });
  Object.keys(om).forEach(function(k){if(!seen[k])dels.push(k);});
  return {ups:ups,dels:dels};
}
function syncNow(){
  if(!SB||!UID||syncing){if(dirty)clearTimeout(syncTimer),syncTimer=setTimeout(syncNow,900);return;}
  syncing=true;dirty=false;
  var cur=clone(DB),old=snap||{members:[],requesters:[],projects:[],tasks:[],notes:[],
    settings:{},weeks:{}};
  var byId=function(x){return x.id;};
  var dM=diffRows(old.members,cur.members,byId,rowMember);
  var dR=diffRows(old.requesters,cur.requesters,byId,rowRequester);
  var dP=diffRows(old.projects,cur.projects,byId,rowProject);
  var dT=diffRows(old.tasks,cur.tasks,byId,rowTask);
  var dN=diffRows(old.notes,cur.notes,byId,rowNote);
  var oldU=flatUpdates(old),newU=flatUpdates(cur);
  var dU=diffRows(oldU,newU,byId,rowUpdate);
  var oldC=flatCollabs(old),newC=flatCollabs(cur);
  var dC=diffRows(oldC,newC,function(x){return x.k;},function(x){
    return {task_id:x.task_id,member_id:x.member_id,owner_id:UID};});
  var setRow={owner_id:UID,display_name:cur.settings.owner||"",theme:cur.settings.theme||"system",
    default_due_days:String(cur.settings.defaultDueDays||""),
    default_report_cycle:cur.settings.defaultReportCycle||"",recent:cur.settings.recent};
  var setChanged=JSON.stringify(setRow)!==JSON.stringify(old.__setRow||{});
  var wkNew=Object.keys(cur.weeks).map(function(k){
    return {owner_id:UID,week_start:k,plan:cur.weeks[k].plan||"",risk:cur.weeks[k].risk||""};});
  var wkOld=Object.keys(old.weeks||{}).map(function(k){
    return {owner_id:UID,week_start:k,plan:old.weeks[k].plan||"",risk:old.weeks[k].risk||""};});
  var dW=diffRows(wkOld,wkNew,function(x){return x.week_start;},function(x){return x;});

  var steps=[];
  function up(tbl,rows){if(rows.length)steps.push(function(){return SB.from(tbl).upsert(rows);});}
  function del(tbl,ids){if(ids.length)steps.push(function(){return SB.from(tbl).delete().in("id",ids);});}
  /* ลบก่อน (ลูก -> แม่) */
  dC.dels.forEach(function(k){var p=k.split("|");
    steps.push(function(){return SB.from("task_collaborators").delete()
      .eq("task_id",p[0]).eq("member_id",p[1]);});});
  del("task_updates",dU.dels);
  del("tasks",dT.dels);
  del("notes",dN.dels);
  del("projects",dP.dels);
  del("requesters",dR.dels);
  del("members",dM.dels);
  dW.dels.forEach(function(k){
    steps.push(function(){return SB.from("week_notes").delete().eq("owner_id",UID).eq("week_start",k);});});
  /* แล้วค่อยเพิ่ม/แก้ (แม่ -> ลูก) */
  up("members",dM.ups); up("requesters",dR.ups); up("projects",dP.ups);
  up("tasks",dT.ups); up("task_collaborators",dC.ups); up("task_updates",dU.ups);
  up("notes",dN.ups);
  if(dW.ups.length)steps.push(function(){return SB.from("week_notes").upsert(dW.ups);});
  if(setChanged)steps.push(function(){return SB.from("settings").upsert(setRow);});

  if(!steps.length){syncing=false;setSave("ok","บันทึกแล้ว");return;}
  var i=0;
  function next(){
    if(i>=steps.length){
      snap=cur;snap.__setRow=setRow;syncing=false;
      setSave("ok","บันทึกแล้ว");
      if(dirty)syncNow();
      return;
    }
    steps[i++]().then(function(res){
      if(res&&res.error)throw res.error;
      next();
    }).catch(function(e){
      syncing=false;dirty=true;
      setSave("bad","บันทึกไม่สำเร็จ — จะลองใหม่");
      console.error("sync",e);
      clearTimeout(syncTimer);syncTimer=setTimeout(syncNow,5000);
    });
  }
  next();
}
function storeNote(){
  var n=$("store-note");if(!n)return;
  n.textContent=UID?("ข้อมูลเก็บบน Supabase บัญชี "+(USER&&USER.email||"")+" — เปิดจากเครื่องไหนก็เห็นข้อมูลชุดเดียวกัน"):
    "ยังไม่ได้เข้าสู่ระบบ";
}

/* ==========================================================
   5. ธีม
   ========================================================== */
var hostTheme=null;
function applyTheme(){
  var t=DB.settings.theme||"system";
  if(t==="system"){
    if(hostTheme)document.documentElement.setAttribute("data-theme",hostTheme);
    else document.documentElement.removeAttribute("data-theme");
  } else document.documentElement.setAttribute("data-theme",t);
  var seg=$("theme-seg");
  if(seg)Array.prototype.forEach.call(seg.children,function(b){
    b.setAttribute("aria-pressed",b.dataset.theme===t?"true":"false");});
}

/* ==========================================================
   6. ข้อมูลคำนวณ
   ========================================================== */
function memberLabel(id){
  if(!id||id===SELF)return DB.settings.owner?DB.settings.owner+" (ดูแลเอง)":"หัวหน้าดูแลเอง";
  var m=byId(DB.members,id);return m?m.name:"ไม่ระบุ";
}
function memberShort(id){
  if(!id||id===SELF)return DB.settings.owner||"หัวหน้า";
  var m=byId(DB.members,id);return m?(m.nickname||m.name):"ไม่ระบุ";
}
function projName(id){var p=byId(DB.projects,id);return p?p.name:"";}
function reqName(id){var r=byId(DB.requesters,id);return r?r.name:"";}
function activeMembers(){return DB.members.filter(function(m){return m.active!==false&&!m.isSelf;});}
function activeReq(){return DB.requesters.filter(function(r){return r.active!==false;});}
function liveProjects(){return DB.projects.filter(function(p){return !p.archived;});}
function isOpen(t){return t.status!=="completed";}
function isLate(t){return isOpen(t)&&t.dueAt&&t.dueAt<today();}
function lastUpdate(t){return t.updates.length?t.updates[t.updates.length-1]:null;}
function lastActivity(t){
  var d=dOf(t.receivedAt);
  t.updates.forEach(function(u){var x=dOf(u.createdAt);if(x>d)d=x;});
  return d;
}
function quietDays(t){return dayDiff(lastActivity(t),today());}
function cycleStep(c,d){return c==="monthly"?addM(d,1):addD(d,c==="biweekly"?14:7);}
function nextReport(t){
  if(!t.reportCycle||!isOpen(t))return "";
  return cycleStep(t.reportCycle,lastActivity(t));
}
function reportDue(t){var n=nextReport(t);return !!n&&n<=today();}
function collabs(t){return (t.collaboratorIds||[]).filter(function(x){return x!==t.ownerId;});}
function whoText(t){
  var c=collabs(t).map(memberShort).join(", ");
  return memberLabel(t.ownerId)+(c?" + "+c:"");
}
function involves(t,id){return t.ownerId===id||collabs(t).indexOf(id)>=0;}

function followUps(){
  var t0=today(),t1=addD(t0,1),t2=addD(t0,2),out=[];
  DB.tasks.forEach(function(t){
    if(t.status==="completed"||t.status==="paused")return;
    var r=null;
    if(t.dueAt&&t.dueAt<t0)r={k:0,cls:"r-over",txt:"เลยกำหนด "+dayDiff(t.dueAt,t0)+" วัน"};
    else if(t.status==="blocked")r={k:1,cls:"r-block",txt:"รอ / ติดปัญหา"};
    else if(reportDue(t))r={k:2,cls:"r-cycle",txt:"ถึงรอบรายงานแล้ว"};
    else if(t.dueAt===t0)r={k:3,cls:"r-due",txt:"ครบกำหนดวันนี้"};
    else if(t.dueAt===t1)r={k:4,cls:"r-due",txt:"พรุ่งนี้ครบกำหนด"};
    else if(t.dueAt&&t.dueAt<=t2)r={k:4,cls:"r-due",txt:"ครบกำหนดใน 2 วัน"};
    else if(quietDays(t)>QUIET)r={k:5,cls:"r-quiet",txt:"ไม่มีอัปเดต "+quietDays(t)+" วัน"};
    if(r){t.__r=r;out.push(t);}
  });
  out.sort(function(a,b){
    if(a.__r.k!==b.__r.k)return a.__r.k-b.__r.k;
    var da=a.dueAt||"9999",db=b.dueAt||"9999";
    if(da!==db)return da<db?-1:1;
    return 0;});
  return out;
}
function sortTasks(a,b){
  var w={high:0,normal:1,low:2};
  if(isLate(a)!==isLate(b))return isLate(a)?-1:1;
  if(w[a.priority]!==w[b.priority])return w[a.priority]-w[b.priority];
  var da=a.dueAt||"9999",db=b.dueAt||"9999";
  if(da!==db)return da<db?-1:1;
  return dOf(b.receivedAt)<dOf(a.receivedAt)?-1:1;
}
function remember(kind,id){
  if(!id)return;
  var l=DB.settings.recent[kind]||[];
  l=[id].concat(l.filter(function(x){return x!==id;})).slice(0,5);
  DB.settings.recent[kind]=l;
}

/* ==========================================================
   7. UI พื้นฐาน
   ========================================================== */
function toast(msg){
  var o=document.querySelector(".toast");if(o)o.remove();
  var d=document.createElement("div");d.className="toast";d.textContent=msg;
  document.body.appendChild(d);setTimeout(function(){d.remove();},2200);
}
var askCb=null;
function ask(title,body,yes,cb){
  $("ask-title").textContent=title;$("ask-body").textContent=body;
  $("ask-yes").textContent=yes||"ยืนยัน";askCb=cb;openOv("ov-ask");
  setTimeout(function(){$("ask-no").focus();},30);
}
var ovStack=[];
function openOv(id){
  var el=$(id);if(!el)return;
  el.hidden=false;
  if(ovStack.indexOf(id)<0)ovStack.push(id);
}
function closeOv(id){
  var el=$(id);if(el)el.hidden=true;
  ovStack=ovStack.filter(function(x){return x!==id;});
}
function closeTop(){
  if(!ovStack.length)return false;
  closeOv(ovStack[ovStack.length-1]);
  return true;
}
document.addEventListener("click",function(e){
  if(!e.target||!e.target.closest)return;
  var c=e.target.closest("[data-close]");
  if(c){var ov=c.closest(".ov");if(ov)closeOv(ov.id);return;}
  if(e.target.classList&&e.target.classList.contains("ov"))closeOv(e.target.id);
});
document.addEventListener("keydown",function(e){if(e.key==="Escape")closeTop();});
$("ask-no").addEventListener("click",function(){closeOv("ov-ask");askCb=null;});
$("ask-yes").addEventListener("click",function(){var c=askCb;closeOv("ov-ask");askCb=null;if(c)c();});

/* ==========================================================
   8. Picker (searchable select + ใช้ล่าสุด)
   ========================================================== */
var pick={kind:"",cb:null,cur:""};
function openPick(kind,cur,cb){
  pick={kind:kind,cb:cb,cur:cur||""};
  $("pick-title").textContent=
    kind==="member"?"เลือกผู้รับผิดชอบหลัก":kind==="project"?"เลือกโครงการ":"เลือกผู้มอบหมายงาน";
  $("pick-q").value="";
  $("pick-new").hidden=(kind==="member");
  $("pick-new").textContent=kind==="project"?"＋ เพิ่มโครงการใหม่":"＋ เพิ่มผู้มอบหมายงาน";
  renderPick();
  openOv("ov-pick");
  setTimeout(function(){$("pick-q").focus();},60);
}
function pickItems(kind){
  if(kind==="member")
    return [{id:SELF,name:DB.settings.owner?DB.settings.owner+" (ดูแลเอง)":"หัวหน้าดูแลเอง"}]
      .concat(activeMembers().map(function(m){
        return {id:m.id,name:m.name,side:m.nickname};}));
  if(kind==="project")
    return [{id:"",name:"ไม่ระบุโครงการ"}].concat(liveProjects().map(function(p){
      return {id:p.id,name:p.name};}));
  return [{id:"",name:"ไม่ระบุ"}].concat(activeReq().map(function(r){
    return {id:r.id,name:r.name};}));
}
function renderPick(){
  var q=($("pick-q").value||"").trim().toLowerCase();
  var items=pickItems(pick.kind);
  var rec=(DB.settings.recent[pick.kind]||[]).map(function(id){
    return items.filter(function(x){return x.id===id;})[0];}).filter(Boolean);
  function opt(x){
    return '<button class="opt" data-pickid="'+esc(x.id)+'" aria-selected="'+
      (x.id===pick.cur)+'"><span class="av">'+esc(initials(x.name))+"</span><span>"+
      esc(x.name)+"</span>"+(x.side?'<span class="side">'+esc(x.side)+"</span>":"")+"</button>";
  }
  var h="";
  if(!q&&rec.length)h+='<div class="optgroup">ใช้ล่าสุด</div>'+rec.map(opt).join("")+
    '<div class="optgroup">ทั้งหมด</div>';
  var list=items.filter(function(x){return !q||x.name.toLowerCase().indexOf(q)>=0;});
  h+=list.length?list.map(opt).join(""):'<div class="empty" style="margin:10px 0"><b>ไม่พบ</b><span>ลองพิมพ์คำอื่น หรือกดเพิ่มใหม่</span></div>';
  $("pick-list").innerHTML=h;
}
$("pick-q").addEventListener("input",renderPick);
$("pick-list").addEventListener("click",function(e){
  var b=e.target.closest("[data-pickid]");if(!b)return;
  var id=b.dataset.pickid;
  closeOv("ov-pick");
  if(pick.cb)pick.cb(id);
});
$("pick-new").addEventListener("click",function(){
  var name=($("pick-q").value||"").trim();
  var kind=pick.kind,cb=pick.cb;
  closeOv("ov-pick");
  openForm(kind==="project"?"เพิ่มโครงการ":"เพิ่มผู้มอบหมายงาน",
    [{key:"name",label:"ชื่อ",type:"text",value:name}],function(v){
      if(!v.name)return false;
      var id=uid();
      if(kind==="project")DB.projects.push({id:id,name:v.name,description:"",status:"active",
        startAt:"",endAt:"",archived:false});
      else DB.requesters.push({id:id,name:v.name,active:true});
      touch();renderAll();
      if(cb)cb(id);
      return true;
    });
});

/* ==========================================================
   9. Generic form overlay
   ========================================================== */
var formCb=null,formFields=[];
function openForm(title,fields,cb){
  formFields=fields;formCb=cb;
  $("form-title").textContent=title;
  $("form-body").innerHTML=fields.map(function(f){
    if(f.type==="toggle")
      return '<label class="chip" style="align-self:flex-start;cursor:pointer"><input type="checkbox" data-fk="'+
        f.key+'"'+(f.value?" checked":"")+' style="width:16px;height:16px;accent-color:var(--accent)">'+
        esc(f.label)+"</label>";
    if(f.type==="select")
      return '<div class="field"><label>'+esc(f.label)+'</label><select data-fk="'+f.key+'">'+
        f.options.map(function(o){return '<option value="'+esc(o[0])+'"'+
          (String(f.value)===String(o[0])?" selected":"")+">"+esc(o[1])+"</option>";}).join("")+
        "</select></div>";
    if(f.type==="textarea")
      return '<div class="field"><label>'+esc(f.label)+'</label><textarea data-fk="'+f.key+
        '">'+esc(f.value||"")+"</textarea></div>";
    return '<div class="field"><label>'+esc(f.label)+'</label><input type="'+
      (f.type||"text")+'" data-fk="'+f.key+'" value="'+esc(f.value||"")+'"></div>';
  }).join("");
  openOv("ov-form");
  setTimeout(function(){var i=$("form-body").querySelector("input,textarea,select");if(i)i.focus();},60);
}
$("form-save").addEventListener("click",function(){
  var v={};
  formFields.forEach(function(f){
    var el=$("form-body").querySelector('[data-fk="'+f.key+'"]');
    if(!el)return;
    v[f.key]=(el.type==="checkbox")?el.checked:el.value.trim();
  });
  var ok=formCb?formCb(v):true;
  if(ok!==false)closeOv("ov-form");
  else toast("กรอกข้อมูลให้ครบก่อน");
});

/* ==========================================================
   10. เพิ่ม / แก้ไขงาน
   ========================================================== */
var cd=null;
function newDraft(pre){
  var s=DB.settings;
  var due="";
  if(s.defaultDueDays!=="" && s.defaultDueDays!==null && s.defaultDueDays!==undefined)
    due=addD(today(),parseInt(s.defaultDueDays,10)||0);
  return {id:null,title:"",ownerId:(s.recent.member[0]||SELF),dueAt:due,
    collaboratorIds:[],requesterId:s.recent.requester[0]||"",
    projectId:(pre&&pre.projectId)||"",description:"",priority:"normal",
    reportCycle:s.defaultReportCycle||"",recvDate:today(),recvTime:"",repeatNext:false};
}
function openCreate(pre){
  cd=newDraft(pre);
  $("create-title").textContent="เพิ่มงานใหม่";
  $("c-save").textContent="บันทึกงาน";
  fillCreate();
  openOv("ov-create");
  setTimeout(function(){$("c-title").focus();},80);
}
function openEdit(t){
  cd={id:t.id,title:t.title,ownerId:t.ownerId,dueAt:t.dueAt,
    collaboratorIds:collabs(t).slice(),requesterId:t.requesterId||"",
    projectId:t.projectId||"",description:t.description||"",priority:t.priority,
    reportCycle:t.reportCycle||"",recvDate:dOf(t.receivedAt),recvTime:tOf(t.receivedAt),
    repeatNext:!!t.repeatNext};
  $("create-title").textContent="แก้ไขงาน";
  $("c-save").textContent="บันทึกการแก้ไข";
  fillCreate();
  openOv("ov-create");
}
function fillCreate(){
  $("c-title").value=cd.title;
  $("c-desc").value=cd.description;
  $("c-due").value=cd.dueAt;
  $("c-recv").value=cd.recvDate;
  $("c-time").value=cd.recvTime;
  $("c-repeat").checked=cd.repeatNext;
  $("c-pri").innerHTML=Object.keys(PRIO).map(function(k){
    return '<option value="'+k+'"'+(cd.priority===k?" selected":"")+">"+PRIO[k]+"</option>";}).join("");
  $("c-cycle").innerHTML=Object.keys(CYCLE).map(function(k){
    return '<option value="'+k+'"'+(cd.reportCycle===k?" selected":"")+">"+CYCLE[k]+"</option>";}).join("");
  $("c-more").hidden=true;
  $("c-more-btn").textContent="รายละเอียดเพิ่มเติม ▾";
  syncCreate();
}
function syncCreate(){
  $("c-owner-v").textContent=memberLabel(cd.ownerId);
  var rn=reqName(cd.requesterId),pn=projName(cd.projectId);
  $("c-req-v").textContent=rn||"ไม่ระบุ";
  $("c-req-v").className="v"+(rn?"":" ph-empty");
  $("c-proj-v").textContent=pn||"ไม่ระบุ";
  $("c-proj-v").className="v"+(pn?"":" ph-empty");
  var t0=today();
  var presets=[["วันนี้",t0],["พรุ่งนี้",addD(t0,1)],["+3 วัน",addD(t0,3)],["สัปดาห์หน้า",addD(t0,7)]];
  $("c-due-chips").innerHTML=presets.map(function(p){
    return '<button type="button" class="chip'+(cd.dueAt===p[1]?" on":"")+'" data-due="'+p[1]+'">'+
      p[0]+"</button>";}).join("")+
    (cd.dueAt?'<button type="button" class="chip" data-due="">ล้าง</button>':"");
  var list=[{id:SELF,name:"ตัวฉันเอง"}].concat(activeMembers())
    .filter(function(m){return m.id!==cd.ownerId;});
  $("c-helpers").innerHTML=list.length?list.map(function(m){
    return '<button type="button" class="chip" aria-pressed="'+
      (cd.collaboratorIds.indexOf(m.id)>=0)+'" data-help="'+m.id+'">'+
      esc(m.nickname||m.name)+"</button>";}).join("")
    :'<span class="dim" style="font-size:13px">ยังไม่มีรายชื่อทีม</span>';
}
$("c-more-btn").addEventListener("click",function(){
  var b=$("c-more");b.hidden=!b.hidden;
  this.textContent=b.hidden?"รายละเอียดเพิ่มเติม ▾":"ซ่อนรายละเอียด ▴";
});
$("c-due").addEventListener("change",function(){cd.dueAt=this.value;syncCreate();});
$("c-due-chips").addEventListener("click",function(e){
  var b=e.target.closest("[data-due]");if(!b)return;
  cd.dueAt=b.dataset.due;$("c-due").value=cd.dueAt;syncCreate();
});
$("c-helpers").addEventListener("click",function(e){
  var b=e.target.closest("[data-help]");if(!b)return;
  var id=b.dataset.help,i=cd.collaboratorIds.indexOf(id);
  if(i>=0)cd.collaboratorIds.splice(i,1);else cd.collaboratorIds.push(id);
  syncCreate();
});
document.addEventListener("click",function(e){
  if(!e.target||!e.target.closest)return;
  var p=e.target.closest("[data-pick]");if(!p)return;
  var kind=p.dataset.pick,tg=p.dataset.target;
  if(tg==="note"){
    openPick(kind,noteProj,function(id){noteProj=id;
      var v=$("note-proj").querySelector(".v");
      v.textContent=projName(id)||"ไม่ระบุ";
      v.className="v"+(id?"":" ph-empty");});
    return;
  }
  var key=tg==="c-owner"?"ownerId":tg==="c-req"?"requesterId":"projectId";
  openPick(kind,cd[key],function(id){
    cd[key]=id;
    if(key==="ownerId")cd.collaboratorIds=cd.collaboratorIds.filter(function(x){return x!==id;});
    syncCreate();});
});
$("c-save").addEventListener("click",function(){
  var title=$("c-title").value.trim();
  if(!title){toast("ใส่ชื่องานก่อน");$("c-title").focus();return;}
  cd.title=title;
  cd.description=$("c-desc").value.trim();
  cd.priority=$("c-pri").value;
  cd.reportCycle=$("c-cycle").value;
  cd.recvDate=$("c-recv").value||today();
  cd.recvTime=$("c-time").value||"";
  cd.repeatNext=cd.reportCycle?$("c-repeat").checked:false;
  var recv=cd.recvDate+(cd.recvTime?"T"+cd.recvTime:"");
  if(cd.id){
    var t=byId(DB.tasks,cd.id);
    if(t){
      t.title=cd.title;t.description=cd.description;t.ownerId=cd.ownerId;
      t.collaboratorIds=cd.collaboratorIds.slice();t.requesterId=cd.requesterId;
      t.projectId=cd.projectId;t.priority=cd.priority;t.reportCycle=cd.reportCycle;
      t.repeatNext=cd.repeatNext;t.receivedAt=recv;t.dueAt=cd.dueAt;t.updatedAt=nowStamp();
    }
    toast("บันทึกการแก้ไขแล้ว");
  } else {
    DB.tasks.push({id:uid(),title:cd.title,description:cd.description,
      requesterId:cd.requesterId,ownerId:cd.ownerId,
      collaboratorIds:cd.collaboratorIds.slice(),projectId:cd.projectId,
      status:"received",priority:cd.priority,receivedAt:recv,dueAt:cd.dueAt,
      reportCycle:cd.reportCycle,repeatNext:cd.repeatNext,
      createdAt:nowStamp(),updatedAt:nowStamp(),completedAt:"",updates:[]});
    toast("บันทึกงานแล้ว");
  }
  remember("member",cd.ownerId);
  remember("project",cd.projectId);
  remember("requester",cd.requesterId);
  closeOv("ov-create");
  touch();renderAll();
});

/* ==========================================================
   11. Quick update
   ========================================================== */
var uid_="",uStatus="";
function openUpdate(id){
  var t=byId(DB.tasks,id);if(!t)return;
  uid_=id;uStatus=t.status;
  $("u-task").textContent=t.title;
  $("u-msg").value="";
  renderUStatus();
  openOv("ov-update");
  setTimeout(function(){$("u-msg").focus();},80);
}
function renderUStatus(){
  $("u-status").innerHTML=STORDER.map(function(k){
    return '<button type="button" class="chip'+(uStatus===k?" on":"")+'" data-ust="'+k+'">'+
      STATUS[k]+"</button>";}).join("");
}
$("u-status").addEventListener("click",function(e){
  var b=e.target.closest("[data-ust]");if(!b)return;
  uStatus=b.dataset.ust;renderUStatus();
});
$("u-save").addEventListener("click",function(){
  var t=byId(DB.tasks,uid_);if(!t)return;
  var msg=$("u-msg").value.trim();
  if(!msg&&uStatus===t.status){toast("พิมพ์ความคืบหน้า หรือเปลี่ยนสถานะก่อน");return;}
  var closing=(uStatus==="completed"&&t.status!=="completed");
  t.updates.push({id:uid(),taskId:t.id,message:msg||("เปลี่ยนสถานะเป็น "+STATUS[uStatus]),
    status:uStatus,createdAt:nowStamp()});
  t.status=uStatus;t.updatedAt=nowStamp();
  if(uStatus==="completed"){if(!t.completedAt)t.completedAt=today();}
  else t.completedAt="";
  var made=closing?spawnNext(t):null;
  closeOv("ov-update");
  touch();renderAll();
  toast(made?"บันทึกอัปเดต · สร้างงานรอบถัดไปแล้ว":"บันทึกอัปเดตแล้ว");
});
function spawnNext(t){
  if(!t.repeatNext||!t.reportCycle)return null;
  var due="";
  if(t.dueAt){due=cycleStep(t.reportCycle,t.dueAt);
    for(var i=0;i<36&&due<today();i++)due=cycleStep(t.reportCycle,due);}
  var nt={id:uid(),title:t.title,description:t.description,requesterId:t.requesterId,
    ownerId:t.ownerId,collaboratorIds:collabs(t).slice(),projectId:t.projectId,
    status:"received",priority:t.priority,receivedAt:nowStamp(),dueAt:due,
    reportCycle:t.reportCycle,repeatNext:true,createdAt:nowStamp(),updatedAt:nowStamp(),
    completedAt:"",updates:[]};
  DB.tasks.push(nt);return nt;
}
function completeTask(id){
  var t=byId(DB.tasks,id);if(!t||t.status==="completed")return;
  t.status="completed";t.completedAt=today();t.updatedAt=nowStamp();
  t.updates.push({id:uid(),taskId:t.id,message:"ปิดงาน",status:"completed",createdAt:nowStamp()});
  var made=spawnNext(t);
  touch();renderAll();
  toast(made?"ปิดงานแล้ว · สร้างงานรอบถัดไปแล้ว":"ปิดงานแล้ว");
}

/* ==========================================================
   12. Task detail
   ========================================================== */
function openDetail(id){
  var t=byId(DB.tasks,id);if(!t)return;
  var kv=[
    ["สถานะ",'<span class="pill s-'+t.status+'">'+STATUS[t.status]+"</span>"],
    ["ความสำคัญ",esc(PRIO[t.priority])],
    ["ผู้รับผิดชอบ",esc(memberLabel(t.ownerId))],
    ["ผู้ร่วมงาน",collabs(t).length?esc(collabs(t).map(memberShort).join(", ")):"—"],
    ["โครงการ",esc(projName(t.projectId)||"—")],
    ["รับมาจาก",esc(reqName(t.requesterId)||"—")],
    ["วันที่รับ",esc(fmtStamp(t.receivedAt))],
    ["กำหนดส่ง",t.dueAt?'<span'+(isLate(t)?' style="color:var(--hot);font-weight:600"':"")+">"+
      esc(fmtDY(t.dueAt))+(isLate(t)?" (เลย "+dayDiff(t.dueAt,today())+" วัน)":"")+"</span>":"—"],
    ["รอบรายงาน",t.reportCycle?esc(CYCLE[t.reportCycle])+
      (isOpen(t)?' <span class="dim">· รอบถัดไป '+esc(fmtDY(nextReport(t)))+"</span>":""):"—"]
  ];
  if(t.completedAt)kv.push(["ปิดงานเมื่อ",esc(fmtDY(t.completedAt))]);
  var h='<div><h2 style="font-size:18px">'+esc(t.title)+"</h2>"+
    (t.description?'<p class="muted" style="margin-top:6px;white-space:pre-wrap">'+esc(t.description)+"</p>":"")+
    "</div>"+
    '<dl class="kv">'+kv.map(function(x){
      return "<dt>"+esc(x[0])+"</dt><dd>"+x[1]+"</dd>";}).join("")+"</dl>"+
    '<div><div class="eyebrow" style="margin-bottom:4px">ประวัติความคืบหน้า</div>'+
    (t.updates.length?'<ul class="tl">'+t.updates.slice().reverse().map(function(u){
      return "<li><div class=\"when\">"+esc(fmtStamp(u.createdAt))+"</div>"+
        '<div class="msg">'+esc(u.message)+"</div>"+
        (u.status?'<span class="pill s-'+u.status+'">'+STATUS[u.status]+"</span>":"")+"</li>";
      }).join("")+"</ul>":'<p class="dim" style="font-size:13.5px">ยังไม่มีการอัปเดต</p>')+"</div>";
  $("detail-body").innerHTML=h;
  $("detail-foot").innerHTML=
    '<button class="btn gh sm" data-act="del" data-id="'+t.id+'">ลบงาน</button>'+
    '<div style="display:flex;gap:8px">'+
      '<button class="btn gh" data-act="edit" data-id="'+t.id+'">แก้ไข</button>'+
      (t.status!=="completed"?'<button class="btn gh" data-act="done" data-id="'+t.id+'">เสร็จแล้ว</button>':"")+
      '<button class="btn" data-act="upd" data-id="'+t.id+'">อัปเดต</button>'+
    "</div>";
  openOv("ov-detail");
}

/* ==========================================================
   13. เรนเดอร์ row
   ========================================================== */
function rowHTML(t,showReason){
  var meta=[];
  var pn=projName(t.projectId);
  meta.push("<b>"+esc(memberShort(t.ownerId))+"</b>");
  if(pn)meta.push(esc(pn));
  if(t.dueAt)meta.push("ครบกำหนด "+fmtD(t.dueAt));
  var q=quietDays(t);
  meta.push(t.updates.length?("อัปเดตล่าสุด "+(q===0?"วันนี้":q+" วันก่อน")):"ยังไม่มีอัปเดต");
  var right='<span class="pill s-'+t.status+'">'+STATUS[t.status]+"</span>";
  if(t.priority==="high")right+='<span class="tag p-high">ด่วน</span>';
  if(t.reportCycle&&isOpen(t))right+='<span class="tag">'+CYSHORT[t.reportCycle]+"</span>";
  return '<div class="row" data-id="'+t.id+'">'+
    "<div>"+(showReason&&t.__r?'<span class="reason '+t.__r.cls+'"><i>●</i>'+esc(t.__r.txt)+"</span>":"")+
      '<div class="rt">'+esc(t.title)+"</div>"+
      '<div class="rm">'+meta.map(function(m){return "<span>"+m+"</span>";}).join("")+"</div></div>"+
    '<div class="rr">'+right+"</div>"+
    '<div class="acts">'+
      (t.status!=="completed"?'<button class="btn sm gh" data-act="upd" data-id="'+t.id+'">อัปเดต</button>':"")+
      '<button class="btn sm gh" data-act="open" data-id="'+t.id+'">ดูงาน</button>'+
      (t.status!=="completed"?'<button class="btn sm gh" data-act="done" data-id="'+t.id+'">เสร็จแล้ว</button>':"")+
    "</div></div>";
}
function listHTML(arr,showReason){
  return arr.length?'<div class="list card" style="padding:2px 15px">'+
    arr.map(function(t){return rowHTML(t,showReason);}).join("")+"</div>":"";
}
function emptyHTML(title,sub,btn){
  return '<div class="empty"><b>'+esc(title)+"</b><span>"+esc(sub)+"</span>"+
    (btn?'<div><button class="btn sm gh" data-act="'+btn[1]+'">'+esc(btn[0])+"</button></div>":"")+"</div>";
}

/* ==========================================================
   14. หน้าวันนี้
   ========================================================== */
function renderToday(){
  $("today-h1").textContent="วันนี้";
  $("today-sub").textContent=fmtFull(today())+
    (DB.settings.owner?" · สวัสดี "+DB.settings.owner:"");
  var f=followUps();
  var t0=today(),t2=addD(t0,2);
  var over=DB.tasks.filter(isLate).length;
  var near=DB.tasks.filter(function(t){
    return isOpen(t)&&t.status!=="paused"&&t.dueAt&&t.dueAt>=t0&&t.dueAt<=t2;}).length;
  var blocked=DB.tasks.filter(function(t){return t.status==="blocked";}).length;
  $("today-strip").innerHTML=
    st(f.length,"ต้องตาม",f.length?"acc":"")+
    st(over,"เลยกำหนด",over?"hot":"")+
    st(near,"ใกล้ครบกำหนด",near?"warn":"")+
    st(blocked,"รอ/ติดปัญหา",blocked?"warn":"");
  $("today-list").innerHTML=f.length?listHTML(f,true)
    :emptyHTML("วันนี้ไม่มีงานที่ต้องตาม 🎉","งานทุกอย่างอยู่ในแผน",["ดูงานทั้งหมด","goTasks"]);
  var notes=DB.notes.filter(function(n){return n.date===t0;});
  $("today-notes").innerHTML=notes.length?
    '<ul class="tl card pad">'+notes.map(function(n){
      return '<li style="display:flex;gap:10px;align-items:flex-start"><div style="flex:1">'+
        '<div class="when">'+fmtD(n.date)+(n.projectId?" · "+esc(projName(n.projectId)):"")+"</div>"+
        '<div class="msg">'+esc(n.text)+"</div></div>"+
        '<button class="btn sm gh" data-act="delnote" data-id="'+n.id+'">ลบ</button></li>';
      }).join("")+"</ul>":"";
}
function st(n,label,cls){
  return '<div class="st'+(cls?" "+cls:"")+'"><b>'+n+"</b><span>"+esc(label)+"</span></div>";
}

/* ==========================================================
   15. หน้างานทั้งหมด
   ========================================================== */
var flt={q:"",status:"open",ownerId:"",projectId:"",priority:"",
         recvFrom:"",recvTo:"",dueFrom:"",dueTo:""};
function filterCount(){
  var n=0;
  ["ownerId","projectId","priority","recvFrom","recvTo","dueFrom","dueTo"].forEach(function(k){
    if(flt[k])n++;});
  return n;
}
function filtered(){
  var q=flt.q.trim().toLowerCase();
  return DB.tasks.filter(function(t){
    if(flt.status==="open"){if(!isOpen(t))return false;}
    else if(flt.status!=="all"&&t.status!==flt.status)return false;
    if(flt.ownerId&&!involves(t,flt.ownerId))return false;
    if(flt.projectId)
      { if(flt.projectId==="__none"){if(t.projectId)return false;}
        else if(t.projectId!==flt.projectId)return false; }
    if(flt.priority&&t.priority!==flt.priority)return false;
    var rd=dOf(t.receivedAt);
    if(flt.recvFrom&&rd<flt.recvFrom)return false;
    if(flt.recvTo&&rd>flt.recvTo)return false;
    if(flt.dueFrom&&(!t.dueAt||t.dueAt<flt.dueFrom))return false;
    if(flt.dueTo&&(!t.dueAt||t.dueAt>flt.dueTo))return false;
    if(q){
      var hay=(t.title+" "+t.description+" "+projName(t.projectId)+" "+
        reqName(t.requesterId)+" "+memberShort(t.ownerId)).toLowerCase();
      if(hay.indexOf(q)<0)return false;
    }
    return true;
  }).sort(sortTasks);
}
function statusChips(){
  var opts=[["open","ยังไม่ปิด"],["all","ทั้งหมด"]].concat(STORDER.map(function(k){
    return [k,STATUS[k]];}));
  return opts.map(function(o){
    return '<button class="chip'+(flt.status===o[0]?" on":"")+'" data-st="'+o[0]+'">'+
      esc(o[1])+"</button>";}).join("");
}
function filterControls(){
  function sel(key,label,options){
    return '<div class="field"><label>'+esc(label)+'</label><select data-f="'+key+'">'+
      options.map(function(o){return '<option value="'+esc(o[0])+'"'+
        (flt[key]===o[0]?" selected":"")+">"+esc(o[1])+"</option>";}).join("")+"</select></div>";
  }
  function date(key,label){
    return '<div class="field"><label>'+esc(label)+'</label><input type="date" data-f="'+key+
      '" value="'+esc(flt[key])+'"></div>';
  }
  return sel("ownerId","ผู้เกี่ยวข้อง",[["","ทุกคน"],[SELF,memberLabel(SELF)]].concat(
      activeMembers().map(function(m){return [m.id,m.name];})))+
    sel("projectId","โครงการ",[["","ทุกโครงการ"]].concat(
      liveProjects().map(function(p){return [p.id,p.name];})).concat([["__none","ไม่ระบุโครงการ"]]))+
    sel("priority","ความสำคัญ",[["","ทุกระดับ"]].concat(
      Object.keys(PRIO).map(function(k){return [k,PRIO[k]];})))+
    date("recvFrom","รับงานตั้งแต่")+date("recvTo","รับงานถึง")+
    date("dueFrom","กำหนดส่งตั้งแต่")+date("dueTo","กำหนดส่งถึง");
}
function renderTasks(){
  $("status-chips").innerHTML=statusChips();
  var n=filterCount();
  $("filter-n").textContent=n?" ("+n+")":"";
  $("filter-bar").innerHTML=filterControls();
  $("filter-body").innerHTML=filterControls();
  var list=filtered();
  $("task-list").innerHTML=list.length?listHTML(list,false)
    :emptyHTML("ไม่พบงานตามเงื่อนไขนี้","ลองล้างตัวกรอง หรือเปลี่ยนคำค้นหา",["ล้างตัวกรอง","clearf"]);
}

/* ==========================================================
   16. หน้าทีม
   ========================================================== */
var sub=null;   /* {type:'member'|'project', id} */
function memberStats(id){
  var mine=DB.tasks.filter(function(t){return t.ownerId===id;});
  var ws=weekStart(today()),we=addD(ws,6);
  return {
    active:mine.filter(function(t){return t.status==="in_progress";}).length,
    open:mine.filter(isOpen).length,
    late:mine.filter(isLate).length,
    blocked:mine.filter(function(t){return t.status==="blocked";}).length,
    dueWeek:mine.filter(function(t){return isOpen(t)&&t.dueAt&&t.dueAt>=ws&&t.dueAt<=we;}).length,
    doneWeek:mine.filter(function(t){return t.status==="completed"&&t.completedAt>=ws&&t.completedAt<=we;}).length,
    helping:DB.tasks.filter(function(t){return isOpen(t)&&t.ownerId!==id&&collabs(t).indexOf(id)>=0;}).length
  };
}
function renderTeam(){
  if(sub&&sub.type==="member"){renderMemberDetail();return;}
  var list=[{id:SELF,name:(DB.settings.owner||"ตัวฉันเอง")+" — หัวหน้าแผนก"}].concat(activeMembers());
  $("team-grid").innerHTML=list.map(function(m){
    var s=memberStats(m.id);
    return '<button class="tile'+(s.late?" alert":"")+'" data-act="member" data-id="'+m.id+'">'+
      "<h3>"+esc(m.nickname?m.name+" ("+m.nickname+")":m.name)+"</h3>"+
      (s.helping?'<div class="sub">ช่วยงานของคนอื่นอีก '+s.helping+" ชิ้น</div>":"")+
      '<div class="cnt"><span><b>'+s.active+"</b>กำลังทำ</span>"+
        '<span class="h"><b>'+s.late+"</b>เลยกำหนด</span>"+
        '<span class="w"><b>'+s.blocked+"</b>ติดปัญหา</span>"+
        "<span><b>"+s.dueWeek+"</b>ครบสัปดาห์นี้</span></div></button>";
  }).join("");
}
function renderMemberDetail(){
  var id=sub.id,s=memberStats(id);
  var ws=weekStart(today());
  var mine=DB.tasks.filter(function(t){return t.ownerId===id;});
  var fu=followUps().filter(function(t){return involves(t,id);});
  var fuIds={};fu.forEach(function(t){fuIds[t.id]=1;});
  var doing=mine.filter(function(t){return t.status==="in_progress"&&!fuIds[t.id];}).sort(sortTasks);
  var blocked=mine.filter(function(t){return t.status==="blocked"&&!fuIds[t.id];}).sort(sortTasks);
  var done=mine.filter(function(t){return t.status==="completed"&&t.completedAt>=addD(ws,-7);})
    .sort(function(a,b){return a.completedAt<b.completedAt?1:-1;}).slice(0,8);
  var helping=DB.tasks.filter(function(t){return isOpen(t)&&t.ownerId!==id&&collabs(t).indexOf(id)>=0;});
  function block(title,arr){
    return arr.length?'<div class="sec"><div class="sec-h"><h2>'+esc(title)+" ("+arr.length+
      ")</h2></div>"+listHTML(arr,title==="ต้องตาม")+"</div>":"";
  }
  $("team-grid").innerHTML=
    '<div style="grid-column:1/-1">'+
      '<button class="btn gh sm" data-act="back" style="margin-bottom:14px">← กลับไปหน้าทีม</button>'+
      '<div class="sec"><h1>'+esc(memberShort(id))+"</h1></div>"+
      '<div class="sec strip">'+st(s.open,"งานที่ถืออยู่","")+st(s.late,"เลยกำหนด",s.late?"hot":"")+
        st(s.blocked,"ติดปัญหา",s.blocked?"warn":"")+st(s.doneWeek,"เสร็จสัปดาห์นี้","ok")+
        st(s.helping,"ไปช่วยคนอื่น","")+"</div>"+
      block("ต้องตาม",fu)+block("กำลังทำ",doing)+block("รอ / ติดปัญหา",blocked)+
      block("ไปช่วยงานคนอื่น",helping)+block("เสร็จล่าสุด",done)+
      (!fu.length&&!doing.length&&!blocked.length&&!done.length&&!helping.length?
        emptyHTML("ยังไม่มีงานของคนนี้","มอบงานให้เขาได้จากปุ่มเพิ่มงาน"):"")+
    "</div>";
}

/* ==========================================================
   17. หน้าโครงการ
   ========================================================== */
function projStats(id){
  var arr=DB.tasks.filter(function(t){return (t.projectId||"")===id;});
  var done=arr.filter(function(t){return t.status==="completed";}).length;
  return {all:arr.length,done:done,
    doing:arr.filter(function(t){return t.status==="in_progress";}).length,
    blocked:arr.filter(function(t){return t.status==="blocked";}).length,
    late:arr.filter(isLate).length,
    pct:arr.length?Math.round(done*100/arr.length):0,
    tasks:arr};
}
function renderProjects(){
  if(sub&&sub.type==="project"){renderProjectDetail();return;}
  var list=liveProjects().slice();
  var hasNone=DB.tasks.some(function(t){return !t.projectId;});
  var cards=list.map(function(p){return projCard(p.id,p.name);});
  if(hasNone)cards.push(projCard("","งานที่ไม่ระบุโครงการ"));
  $("proj-grid").innerHTML=cards.length?cards.join("")
    :'<div style="grid-column:1/-1">'+emptyHTML("ยังไม่มีโครงการ",
      "เพิ่มโครงการได้ที่หน้าตั้งค่า หรือกดเพิ่มตอนสร้างงานใหม่")+"</div>";
}
function projCard(id,name){
  var s=projStats(id);
  return '<button class="tile'+(s.late?" alert":"")+'" data-act="project" data-id="'+esc(id)+'">'+
    "<h3>"+esc(name)+"</h3>"+
    '<div class="bar"><i style="width:'+s.pct+'%"></i></div>'+
    '<div class="barline"><span>เสร็จ '+s.done+" / "+s.all+"</span><span>"+s.pct+"%</span></div>"+
    '<div class="cnt"><span><b>'+s.doing+"</b>กำลังทำ</span>"+
      '<span class="w"><b>'+s.blocked+"</b>ติดปัญหา</span>"+
      '<span class="h"><b>'+s.late+"</b>เลยกำหนด</span></div></button>";
}
function renderProjectDetail(){
  var id=sub.id,p=byId(DB.projects,id),s=projStats(id);
  var name=p?p.name:"งานที่ไม่ระบุโครงการ";
  var who=uniq(s.tasks.filter(isOpen).map(function(t){return memberShort(t.ownerId);}));
  var open=s.tasks.filter(isOpen).sort(sortTasks);
  var done=s.tasks.filter(function(t){return t.status==="completed";})
    .sort(function(a,b){return a.completedAt<b.completedAt?1:-1;}).slice(0,10);
  var kv="";
  if(p){
    var rows=[["สถานะ",p.status==="done"?"ปิดโครงการ":"กำลังดำเนินการ"],
      ["วันที่เริ่ม",p.startAt?fmtDY(p.startAt):"—"],
      ["กำหนดสิ้นสุด",p.endAt?fmtDY(p.endAt):"—"],
      ["ผู้เกี่ยวข้อง",who.length?who.join(", "):"—"]];
    kv='<dl class="kv card pad">'+rows.map(function(r){
      return "<dt>"+esc(r[0])+"</dt><dd>"+esc(r[1])+"</dd>";}).join("")+"</dl>";
  }
  $("proj-grid").innerHTML='<div style="grid-column:1/-1">'+
    '<button class="btn gh sm" data-act="back" style="margin-bottom:14px">← กลับไปหน้าโครงการ</button>'+
    '<div class="sec"><h1>'+esc(name)+"</h1>"+
      (p&&p.description?'<p class="muted" style="margin-top:5px">'+esc(p.description)+"</p>":"")+"</div>"+
    '<div class="sec strip">'+st(s.all,"งานทั้งหมด","")+st(s.done,"เสร็จแล้ว","ok")+
      st(s.doing,"กำลังทำ","")+st(s.blocked,"ติดปัญหา",s.blocked?"warn":"")+
      st(s.late,"เลยกำหนด",s.late?"hot":"")+"</div>"+
    '<div class="sec card pad"><div class="barline" style="margin-bottom:6px">'+
      "<span>ความคืบหน้า</span><span>"+s.pct+"%</span></div>"+
      '<div class="bar"><i style="width:'+s.pct+'%"></i></div></div>'+
    (kv?'<div class="sec">'+kv+"</div>":"")+
    '<div class="sec"><div class="sec-h"><h2>งานที่ยังไม่ปิด ('+open.length+")</h2>"+
      '<button class="btn sm" data-act="addinproj" data-id="'+esc(id)+'">＋ เพิ่มงานในโครงการนี้</button></div>'+
      (open.length?listHTML(open,false):emptyHTML("ไม่มีงานค้าง","งานในโครงการนี้ปิดครบแล้ว"))+"</div>"+
    (done.length?'<div class="sec"><div class="sec-h"><h2>เสร็จล่าสุด</h2></div>'+
      listHTML(done,false)+"</div>":"")+
    "</div>";
}

/* ==========================================================
   18. สรุปสัปดาห์
   ========================================================== */
var curWeek=weekStart(today());
function rangeData(from,to){
  var inR=function(d){return !!d&&d>=from&&d<=to;};
  return {from:from,to:to,
    done:DB.tasks.filter(function(t){return t.status==="completed"&&inR(t.completedAt);}),
    doing:DB.tasks.filter(function(t){return t.status==="received"||t.status==="in_progress";}),
    blocked:DB.tasks.filter(function(t){return t.status==="blocked";}),
    fresh:DB.tasks.filter(function(t){return inR(dOf(t.receivedAt));}),
    late:DB.tasks.filter(isLate),
    notes:DB.notes.filter(function(n){return inR(n.date);})};
}
function groupByProject(arr){
  var map={},order=[];
  arr.forEach(function(t){
    var k=t.projectId||"";
    if(!map[k]){map[k]=[];order.push(k);}
    map[k].push(t);});
  order.sort(function(a,b){
    if(!a)return 1; if(!b)return -1;
    var ia=DB.projects.findIndex(function(p){return p.id===a;});
    var ib=DB.projects.findIndex(function(p){return p.id===b;});
    return (ia<0?999:ia)-(ib<0?999:ib);});
  return order.map(function(k){return {id:k,name:projName(k)||"งานที่ไม่ระบุโครงการ",items:map[k]};});
}
function renderWeek(){
  var we=addD(curWeek,6),w=rangeData(curWeek,we);
  $("week-h1").textContent="สรุปสัปดาห์ "+fmtD(curWeek)+" – "+fmtDY(we);
  $("week-strip").innerHTML=st(w.done.length,"เสร็จแล้ว","ok")+
    st(w.doing.length,"กำลังดำเนินการ","")+st(w.blocked.length,"ติดปัญหา",w.blocked.length?"warn":"")+
    st(w.late.length,"เลยกำหนด",w.late.length?"hot":"")+st(w.fresh.length,"งานใหม่","");
  var wn=DB.weeks[curWeek]||{};
  $("wk-plan").value=wn.plan||"";$("wk-risk").value=wn.risk||"";

  var all=[].concat(w.done,w.doing,w.blocked);
  var seen={},uniqT=[];
  all.forEach(function(t){if(!seen[t.id]){seen[t.id]=1;uniqT.push(t);}});
  var groups=groupByProject(uniqT);
  var h='<div class="ov-sum">'+
    '<span class="pill s-completed">เสร็จแล้ว '+w.done.length+" งาน</span>"+
    '<span class="pill s-in_progress">กำลังดำเนินการ '+w.doing.length+" งาน</span>"+
    '<span class="pill s-blocked">ติดปัญหา '+w.blocked.length+" งาน</span>"+
    (w.late.length?'<span class="pill" style="background:var(--hot-soft);color:var(--hot);border-color:var(--hot)">เลยกำหนด '+w.late.length+" งาน</span>":"")+
    "</div>";
  groups.forEach(function(g){
    var d=g.items.filter(function(t){return t.status==="completed";});
    var p=g.items.filter(function(t){return t.status==="received"||t.status==="in_progress";}).sort(sortTasks);
    var b=g.items.filter(function(t){return t.status==="blocked";});
    if(!d.length&&!p.length&&!b.length)return;
    h+='<div class="prj"><h3>'+esc(g.name)+"</h3>";
    if(d.length)h+="<h4>เสร็จแล้ว</h4><ul>"+d.map(function(t){
      return "<li>"+esc(t.title)+" <small>— "+esc(whoText(t))+"</small></li>";}).join("")+"</ul>";
    if(p.length)h+="<h4>กำลังดำเนินการ</h4><ul>"+p.map(function(t){
      var u=lastUpdate(t);
      return "<li>"+esc(t.title)+" <small>— "+esc(memberShort(t.ownerId))+
        (t.dueAt?" · ครบกำหนด "+fmtD(t.dueAt):"")+(u?" · "+esc(u.message):"")+"</small></li>";}).join("")+"</ul>";
    if(b.length)h+="<h4>ติดปัญหา</h4><ul>"+b.map(function(t){
      var u=lastUpdate(t);
      return "<li>"+esc(t.title)+" <small>— "+esc(memberShort(t.ownerId))+
        (u?" · "+esc(u.message):"")+"</small></li>";}).join("")+"</ul>";
    h+="</div>";
  });
  if(w.fresh.length){
    var c={};w.fresh.forEach(function(t){c[t.status]=(c[t.status]||0)+1;});
    h+='<div class="prj"><h3>งานใหม่ที่รับเข้ามา ('+w.fresh.length+")</h3>"+
      '<div class="ov-sum">'+STORDER.filter(function(k){return c[k];}).map(function(k){
        return '<span class="pill s-'+k+'">'+STREP[k]+" "+c[k]+"</span>";}).join("")+"</div>"+
      "<ul>"+w.fresh.map(function(t){
        return "<li>"+esc(t.title)+' <span class="pill s-'+t.status+'">'+STREP[t.status]+"</span>"+
          " <small>— "+(t.requesterId?"จาก "+esc(reqName(t.requesterId))+" · ":"")+
          esc(memberShort(t.ownerId))+" · รับ "+fmtD(dOf(t.receivedAt))+"</small></li>";}).join("")+
      "</ul></div>";
  }
  if(w.notes.length)h+='<div class="prj"><h3>บันทึกของฉัน</h3><ul>'+w.notes.map(function(n){
    return "<li>"+esc(n.text)+" <small>— "+fmtD(n.date)+"</small></li>";}).join("")+"</ul></div>";
  if(wn.plan)h+='<div class="prj"><h3>แผนสัปดาห์หน้า</h3><ul>'+
    wn.plan.split("\n").filter(Boolean).map(function(x){return "<li>"+esc(x)+"</li>";}).join("")+"</ul></div>";
  if(wn.risk)h+='<div class="prj"><h3>ปัญหา / ต้องการการตัดสินใจ</h3><ul>'+
    wn.risk.split("\n").filter(Boolean).map(function(x){return "<li>"+esc(x)+"</li>";}).join("")+"</ul></div>";
  $("week-report").innerHTML=h;
  $("week-text").value=reportText(w,wn);
}
function reportText(w,wn){
  var L=[];
  L.push("สรุปงานประจำสัปดาห์");
  L.push(fmtD(w.from)+"–"+fmtDY(w.to));
  if(DB.settings.owner)L.push("โดย "+DB.settings.owner);
  L.push("");
  L.push("ภาพรวม");
  L.push("เสร็จแล้ว "+w.done.length+" งาน");
  L.push("กำลังดำเนินการ "+w.doing.length+" งาน");
  L.push("ติดปัญหา "+w.blocked.length+" งาน");
  L.push("เลยกำหนด "+w.late.length+" งาน");
  L.push("");
  var all=[].concat(w.done,w.doing,w.blocked),seen={},u=[];
  all.forEach(function(t){if(!seen[t.id]){seen[t.id]=1;u.push(t);}});
  groupByProject(u).forEach(function(g){
    var d=g.items.filter(function(t){return t.status==="completed";});
    var p=g.items.filter(function(t){return t.status==="received"||t.status==="in_progress";}).sort(sortTasks);
    var b=g.items.filter(function(t){return t.status==="blocked";});
    if(!d.length&&!p.length&&!b.length)return;
    L.push(g.id?("โครงการ "+g.name):g.name);
    if(d.length){L.push("เสร็จแล้ว");d.forEach(function(t){
      L.push("• "+t.title+" — "+memberShort(t.ownerId));});}
    if(p.length){L.push("กำลังดำเนินการ");p.forEach(function(t){
      var lu=lastUpdate(t);
      L.push("• "+t.title+" — "+memberShort(t.ownerId)+
        (t.dueAt?" ครบกำหนด "+fmtD(t.dueAt):"")+(lu?" ("+lu.message+")":""));});}
    if(b.length){L.push("ติดปัญหา");b.forEach(function(t){
      var lu=lastUpdate(t);
      L.push("• "+t.title+" — "+memberShort(t.ownerId)+(lu?" ("+lu.message+")":""));});}
    L.push("");
  });
  if(w.fresh.length){
    L.push("งานใหม่ที่รับเข้ามา ("+w.fresh.length+")");
    w.fresh.forEach(function(t){
      L.push("• "+t.title+(t.requesterId?" (จาก "+reqName(t.requesterId)+")":"")+
        " — "+memberShort(t.ownerId)+" → "+STREP[t.status]);});
    L.push("");
  }
  if(w.notes.length){L.push("บันทึกของฉัน");
    w.notes.forEach(function(n){L.push("• "+fmtD(n.date)+" "+n.text);});L.push("");}
  if(wn.plan){L.push("แผนสัปดาห์หน้า");
    wn.plan.split("\n").filter(Boolean).forEach(function(x){L.push("• "+x);});L.push("");}
  if(wn.risk){L.push("ต้องการการสนับสนุน / ตัดสินใจ");
    wn.risk.split("\n").filter(Boolean).forEach(function(x){L.push("• "+x);});L.push("");}
  return L.join("\n").trim();
}

/* ==========================================================
   19. หน้าพิมพ์
   ========================================================== */
function printScope(){
  var s=$("p-scope").value;
  if(s==="month")return {from:monthStart(today()),to:monthEnd(today()),
    label:"เดือน "+fmtDY(monthStart(today())).replace(/^\d+\s/,"")};
  if(s==="open")return {from:"0000-01-01",to:"9999-12-31",label:"ภาพรวมงานที่ยังไม่ปิด"};
  return {from:curWeek,to:addD(curWeek,6),label:"สัปดาห์ "+fmtD(curWeek)+" – "+fmtDY(addD(curWeek,6))};
}
function renderPrintControls(){
  $("p-proj").innerHTML='<option value="">ทุกโครงการ</option>'+
    liveProjects().map(function(p){return '<option value="'+p.id+'">'+esc(p.name)+"</option>";}).join("")+
    '<option value="__none">ไม่ระบุโครงการ</option>';
  $("p-member").innerHTML='<option value="">ทุกคน</option><option value="me">'+
    esc(memberLabel(SELF))+"</option>"+
    activeMembers().map(function(m){return '<option value="'+m.id+'">'+esc(m.name)+"</option>";}).join("");
  $("p-status").innerHTML='<option value="">ทุกสถานะ</option>'+
    STORDER.map(function(k){return '<option value="'+k+'">'+STATUS[k]+"</option>";}).join("");
}
function renderPrint(){
  var sc=printScope();
  $("p-wk-wrap").hidden=$("p-scope").value!=="week";
  var pf=$("p-proj").value,mf=$("p-member").value,sf=$("p-status").value;
  var gk=$("p-group").value;
  var d=rangeData(sc.from,sc.to);
  function keep(arr){
    return arr.filter(function(t){
      if(pf){if(pf==="__none"){if(t.projectId)return false;}else if(t.projectId!==pf)return false;}
      if(mf&&!involves(t,mf))return false;
      if(sf&&t.status!==sf)return false;
      return true;});
  }
  var done=keep(d.done),doing=keep(d.doing),blocked=keep(d.blocked),
      fresh=keep(d.fresh),late=keep(d.late);
  var notes=d.notes.filter(function(n){
    return (!pf||(pf==="__none"?!n.projectId:n.projectId===pf))&&!mf&&!sf;});
  var wn=DB.weeks[curWeek]||{};

  function grp(arr){
    if(gk==="owner"){
      var m={},o=[];
      arr.forEach(function(t){var k=memberShort(t.ownerId);
        if(!m[k]){m[k]=[];o.push(k);}m[k].push(t);});
      o.sort();
      return o.map(function(k){return {name:k,items:m[k]};});
    }
    return groupByProject(arr);
  }
  function stCell(t){
    return '<td><span class="st st-'+t.status+'"><i>'+STSYM[t.status]+"</i>"+
      esc(STREP[t.status])+"</span></td>";
  }
  function whoCell(t){
    var c=collabs(t).map(memberShort).join(", ");
    return "<td>"+esc(memberLabel(t.ownerId))+(c?' <span class="hp">+ '+esc(c)+"</span>":"")+"</td>";
  }
  function titleCell(t){
    return "<td>"+esc(t.title)+(t.reportCycle?' <span class="hp">('+esc(CYSHORT[t.reportCycle])+
      ")</span>":"")+"</td>";
  }
  function table(arr,head,row){
    if(!arr.length)return '<p class="none">ไม่มีรายการ</p>';
    return grp(arr).map(function(g){
      return "<h4>"+esc(g.name)+" ("+g.items.length+")</h4>"+
        '<div class="tw"><table><thead><tr>'+head.map(function(h){
          return "<th>"+esc(h)+"</th>";}).join("")+"</tr></thead><tbody>"+
        g.items.map(function(t){return "<tr>"+row(t).join("")+"</tr>";}).join("")+
        "</tbody></table></div>";
    }).join("");
  }
  function tally(arr){
    if(!arr.length)return "";
    var c={};arr.forEach(function(t){c[t.status]=(c[t.status]||0)+1;});
    return '<p class="tally"><b>'+arr.length+"</b> งาน"+
      ["completed","in_progress","received","blocked","paused"].filter(function(k){return c[k];})
      .map(function(k){return '<span class="st-'+k+'"><i>'+STSYM[k]+"</i>"+STREP[k]+" "+c[k]+"</span>";})
      .join("")+"</p>";
  }
  var h='<div class="sh-head"><div><h2>รายงานความคืบหน้างาน</h2><div class="sub">'+
    esc(sc.label)+(pf?" · "+esc(pf==="__none"?"ไม่ระบุโครงการ":projName(pf)):"")+
    (mf?" · เฉพาะ "+esc(memberShort(mf)):"")+(sf?" · เฉพาะ "+esc(STATUS[sf]):"")+
    '</div></div><div class="meta">'+(DB.settings.owner?"ผู้รายงาน: "+esc(DB.settings.owner)+"<br>":"")+
    "พิมพ์เมื่อ "+fmtDY(today())+"</div></div>";
  h+='<div class="kpis">'+
    '<div class="kpi"><b>'+done.length+"</b><span>เสร็จแล้ว</span></div>"+
    '<div class="kpi"><b>'+doing.length+"</b><span>กำลังดำเนินการ</span></div>"+
    '<div class="kpi"><b>'+blocked.length+"</b><span>ติดปัญหา</span></div>"+
    '<div class="kpi"><b>'+fresh.length+"</b><span>งานใหม่</span></div>"+
    '<div class="kpi"><b>'+late.length+"</b><span>เลยกำหนด</span></div></div>";
  h+="<section><h3>เสร็จแล้ว</h3>"+table(done,["งาน","ผู้รับผิดชอบ","รับเมื่อ","ปิดเมื่อ"],
    function(t){return [titleCell(t),whoCell(t),'<td class="n">'+fmtD(dOf(t.receivedAt))+"</td>",
      '<td class="n">'+fmtD(t.completedAt)+"</td>"];})+"</section>";
  h+="<section><h3>กำลังดำเนินการ</h3>"+table(doing.slice().sort(sortTasks),
    ["งาน","ผู้รับผิดชอบ","กำหนดส่ง","อัปเดตล่าสุด"],function(t){
      var u=lastUpdate(t);
      return [titleCell(t),whoCell(t),
        '<td class="n'+(isLate(t)?" late":"")+'">'+(t.dueAt?fmtD(t.dueAt):"—")+"</td>",
        "<td>"+(u?esc(u.message)+' <span class="n">('+fmtD(dOf(u.createdAt))+")</span>":
          '<span class="none">ยังไม่มีอัปเดต</span>')+"</td>"];})+"</section>";
  h+="<section><h3>ติดปัญหา / รออยู่</h3>"+table(blocked,["งาน","ผู้รับผิดชอบ","ติดตั้งแต่","ประเด็น"],
    function(t){var u=lastUpdate(t);
      return [titleCell(t),whoCell(t),'<td class="n">'+fmtD(lastActivity(t))+"</td>",
        "<td>"+(u?esc(u.message):"—")+"</td>"];})+"</section>";
  h+="<section><h3>งานใหม่ที่รับเข้ามา</h3>"+tally(fresh)+
    table(fresh,["งาน","รับมาจาก","ผู้รับผิดชอบ","วันเวลาที่รับ","สถานะตอนนี้"],function(t){
      return [titleCell(t),"<td>"+esc(reqName(t.requesterId)||"—")+"</td>",whoCell(t),
        '<td class="n">'+fmtD(dOf(t.receivedAt))+(tOf(t.receivedAt)?" "+tOf(t.receivedAt):"")+"</td>",
        stCell(t)];})+"</section>";
  if(late.length)h+="<section><h3>เลยกำหนด</h3>"+
    table(late,["งาน","ผู้รับผิดชอบ","กำหนดส่ง","เลยมาแล้ว"],function(t){
      return [titleCell(t),whoCell(t),'<td class="n">'+fmtD(t.dueAt)+"</td>",
        '<td class="n late">'+dayDiff(t.dueAt,today())+" วัน</td>"];})+"</section>";
  if(notes.length)h+='<section><h3>บันทึกของฉัน</h3><div class="tw"><table><thead><tr>'+
    "<th>วันที่</th><th>รายละเอียด</th><th>โครงการ</th></tr></thead><tbody>"+
    notes.map(function(n){return '<tr><td class="n">'+fmtD(n.date)+"</td><td>"+esc(n.text)+
      "</td><td>"+esc(projName(n.projectId)||"—")+"</td></tr>";}).join("")+
    "</tbody></table></div></section>";
  if($("p-scope").value==="week"){
    if(wn.plan)h+='<section><h3>แผนสัปดาห์หน้า</h3><p class="free">'+esc(wn.plan)+"</p></section>";
    if(wn.risk)h+='<section><h3>ปัญหา / เรื่องที่ต้องตัดสินใจ</h3><p class="free">'+esc(wn.risk)+"</p></section>";
  }
  h+='<div class="sh-foot"><span>สมุดคุมงานทีม</span><span>'+esc(sc.label)+"</span></div>";
  $("print-doc").innerHTML=h;
}

/* ==========================================================
   20. ตั้งค่า
   ========================================================== */
function renderSettings(){
  $("s-owner").value=DB.settings.owner||"";
  $("s-due").value=DB.settings.defaultDueDays===""?"":String(DB.settings.defaultDueDays);
  $("s-cycle").innerHTML=Object.keys(CYCLE).map(function(k){
    return '<option value="'+k+'"'+(DB.settings.defaultReportCycle===k?" selected":"")+">"+
      CYCLE[k]+"</option>";}).join("");
  var mlist=DB.members.filter(function(m){return !m.isSelf;});
  $("member-list").innerHTML=mlist.length?mlist.map(function(m,i){
    var n=DB.tasks.filter(function(t){return t.ownerId===m.id&&isOpen(t);}).length;
    var meta=[m.department,m.position].filter(Boolean).join(" · ");
    return '<div class="'+(m.active===false?"off":"")+'"><div class="nm"><b>'+esc(m.name)+"</b>"+
      (m.nickname?'<span>'+esc(m.nickname)+"</span>":"")+
      '<span>'+(meta?esc(meta)+" · ":"")+"งานค้าง "+n+(m.active===false?" · ไม่ใช้งาน":"")+"</span></div>"+
      '<button class="btn sm gh" data-act="edm" data-id="'+m.id+'">แก้ไข</button>'+
      '<button class="btn sm gh" data-act="rmm" data-id="'+m.id+'">ลบ</button></div>';
    }).join(""):'<div class="dim" style="font-size:13.5px">ยังไม่มีสมาชิก กด “＋ เพิ่มสมาชิก”</div>';
  $("requester-count").textContent=DB.requesters.length?"("+DB.requesters.length+" คน)":"";
  $("requester-list").innerHTML=DB.requesters.length?DB.requesters.map(function(r,i){
    var n=DB.tasks.filter(function(t){return t.requesterId===r.id;}).length;
    var meta=[r.department,r.position].filter(Boolean).join(" · ");
    return '<div class="'+(r.active===false?"off":"")+'"><div class="nm"><b>'+esc(r.name)+
      '</b><span>'+(meta?esc(meta)+" · ":"")+n+" งาน"+(r.active===false?" · ไม่ใช้งาน":"")+"</span></div>"+
      '<button class="btn sm gh" data-act="mvr" data-id="'+r.id+'" data-d="-1">↑</button>'+
      '<button class="btn sm gh" data-act="mvr" data-id="'+r.id+'" data-d="1">↓</button>'+
      '<button class="btn sm gh" data-act="edr" data-id="'+r.id+'">แก้ไข</button>'+
      '<button class="btn sm gh" data-act="rmr" data-id="'+r.id+'">ลบ</button></div>';
    }).join(""):'<div class="dim" style="font-size:13.5px">ยังไม่มีรายชื่อ</div>';
  $("project-list").innerHTML=DB.projects.length?DB.projects.map(function(p){
    var s=projStats(p.id);
    return '<div class="'+(p.archived?"off":"")+'"><div class="nm"><b>'+esc(p.name)+
      '</b><span>· '+s.done+"/"+s.all+" งาน"+(p.archived?" · เก็บเข้าคลัง":"")+"</span></div>"+
      '<button class="btn sm gh" data-act="edp" data-id="'+p.id+'">แก้ไข</button>'+
      '<button class="btn sm gh" data-act="arp" data-id="'+p.id+'">'+(p.archived?"กู้คืน":"เก็บ")+"</button>"+
      '<button class="btn sm gh" data-act="rmp" data-id="'+p.id+'">ลบ</button></div>';
    }).join(""):'<div class="dim" style="font-size:13.5px">ยังไม่มีโครงการ</div>';
  var cn=doneOlder(parseInt($("cl-age").value,10)).length;
  $("cl-n").textContent=cn;$("cl-done").disabled=!cn;
  applyTheme();storeNote();
}
function doneOlder(months){
  var cut=months?addM(today(),-months):"9999-12-31";
  return DB.tasks.filter(function(t){
    return t.status==="completed"&&(t.completedAt||dOf(t.receivedAt))<cut;});
}
function afterWipe(msg){
  sub=null;curWeek=weekStart(today());touch();renderAll();toast(msg);
}

/* ==========================================================
   21. เรนเดอร์รวม + นำทาง
   ========================================================== */
var view="today";
function renderAll(){
  var f=followUps().length,open=DB.tasks.filter(isOpen).length;
  $("bg-today").textContent=f;$("bg-today").className="bg"+(f?" hot":"");
  $("bg-tasks").textContent=open;
  $("bg-team").textContent=activeMembers().length;
  $("bg-proj").textContent=liveProjects().length;
  if(view==="today")renderToday();
  if(view==="tasks")renderTasks();
  if(view==="team")renderTeam();
  if(view==="projects")renderProjects();
  if(view==="week")renderWeek();
  if(view==="print"){renderPrintControls();renderPrint();}
  if(view==="settings")renderSettings();
  $("fab").hidden=(view==="print"||view==="settings"||view==="week");
}
function go(v,keepSub){
  if(!keepSub)sub=null;
  view=v;
  VIEWS.forEach(function(k){$("v-"+k).hidden=(k!==v);});
  Array.prototype.forEach.call($("nav").children,function(b){
    var on=b.dataset.go?b.dataset.go===v:(MOREV.indexOf(v)>=0);
    b.setAttribute("aria-current",on?"true":"false");});
  $("tb-title").textContent=TITLES[v]||"สมุดคุมงานทีม";
  closeOv("ov-more");
  try{sessionStorage.setItem("twb.view",v);}catch(e){}
  renderAll();window.scrollTo(0,0);
}
$("nav").addEventListener("click",function(e){
  var b=e.target.closest("button");if(!b)return;
  if(b.dataset.go)go(b.dataset.go);
  else if(b.dataset.more)openOv("ov-more");
});
$("ov-more").addEventListener("click",function(e){
  var b=e.target.closest("[data-go]");if(b)go(b.dataset.go);});
$("fab").addEventListener("click",function(){openCreate(null);});

/* ==========================================================
   22. Event delegation
   ========================================================== */
document.addEventListener("click",function(e){
  if(!e.target||!e.target.closest)return;
  var el=e.target.closest("[data-act]");if(!el)return;
  var a=el.dataset.act,id=el.dataset.id;
  if(a==="upd"){closeOv("ov-detail");openUpdate(id);return;}
  if(a==="open"){openDetail(id);return;}
  if(a==="edit"){closeOv("ov-detail");var t=byId(DB.tasks,id);if(t)openEdit(t);return;}
  if(a==="done"){closeOv("ov-detail");completeTask(id);return;}
  if(a==="del"){
    var dt=byId(DB.tasks,id);if(!dt)return;
    closeOv("ov-detail");
    ask("ลบงานนี้ถาวร?",dt.title+" — ประวัติการอัปเดตทั้งหมดจะหายไปด้วย กู้คืนไม่ได้","ลบงานนี้",
      function(){DB.tasks=DB.tasks.filter(function(x){return x.id!==id;});
        touch();renderAll();toast("ลบแล้ว");});
    return;
  }
  if(a==="member"){sub={type:"member",id:id};renderTeam();window.scrollTo(0,0);return;}
  if(a==="project"){sub={type:"project",id:id};renderProjects();window.scrollTo(0,0);return;}
  if(a==="back"){sub=null;renderAll();window.scrollTo(0,0);return;}
  if(a==="addinproj"){openCreate({projectId:id});return;}
  if(a==="goTasks"){go("tasks");return;}
  if(a==="clearf"){
    flt={q:"",status:"open",ownerId:"",projectId:"",priority:"",
      recvFrom:"",recvTo:"",dueFrom:"",dueTo:""};
    $("q").value="";renderTasks();return;
  }
  if(a==="delnote"){DB.notes=DB.notes.filter(function(n){return n.id!==id;});
    touch();renderToday();return;}
  /* --- settings --- */
  if(a==="edm"||a==="rmm"){
    var m=byId(DB.members,id);if(!m)return;
    if(a==="rmm"){
      var used=DB.tasks.filter(function(t){return t.ownerId===id;}).length;
      if(used)ask("เอา "+m.name+" ออกจากทีม?","มีงาน "+used+
        " ชิ้นที่เขาเป็นผู้รับผิดชอบ ระบบจะตั้งเป็น “ไม่ใช้งาน” แทนการลบ เพื่อรักษาประวัติงานเก่า",
        "เอาออกจากทีม",function(){m.active=false;touch();renderAll();toast("เอาออกแล้ว");});
      else ask("ลบ "+m.name+"?","ยังไม่มีงานผูกกับชื่อนี้ ลบได้เลย","ลบ",function(){
        DB.members=DB.members.filter(function(x){return x.id!==id;});
        touch();renderAll();toast("ลบแล้ว");});
      return;
    }
    openForm("แก้ไขสมาชิก",[
      {key:"name",label:"ชื่อ",type:"text",value:m.name},
      {key:"nickname",label:"ชื่อเล่น",type:"text",value:m.nickname},
      {key:"department",label:"แผนก",type:"text",value:m.department||""},
      {key:"position",label:"ตำแหน่ง",type:"text",value:m.position||""},
      {key:"active",label:"ใช้งานอยู่",type:"toggle",value:m.active!==false}],
      function(v){if(!v.name)return false;
        m.name=v.name;m.nickname=v.nickname;m.department=v.department;m.position=v.position;
        m.active=v.active;touch();renderAll();toast("บันทึกแล้ว");return true;});
    return;
  }
  if(a==="edr"||a==="rmr"||a==="mvr"){
    var r=byId(DB.requesters,id);if(!r)return;
    if(a==="mvr"){
      var i=DB.requesters.indexOf(r),j=i+parseInt(el.dataset.d,10);
      if(j<0||j>=DB.requesters.length)return;
      DB.requesters.splice(i,1);DB.requesters.splice(j,0,r);
      touch();renderSettings();return;
    }
    if(a==="rmr"){
      var u2=DB.tasks.filter(function(t){return t.requesterId===id;}).length;
      ask("ลบ "+r.name+"?",u2?("มีงาน "+u2+" ชิ้นที่รับมาจากคนนี้ งานเดิมจะแสดงเป็น “ไม่ระบุ”"):
        "ยังไม่มีงานผูกกับชื่อนี้","ลบ",function(){
        DB.requesters=DB.requesters.filter(function(x){return x.id!==id;});
        touch();renderAll();toast("ลบแล้ว");});
      return;
    }
    openForm("แก้ไขผู้มอบหมายงาน",[
      {key:"name",label:"ชื่อ",type:"text",value:r.name},
      {key:"department",label:"แผนก",type:"text",value:r.department||""},
      {key:"position",label:"ตำแหน่ง",type:"text",value:r.position||""},
      {key:"active",label:"ใช้งานอยู่",type:"toggle",value:r.active!==false}],
      function(v){if(!v.name)return false;r.name=v.name;r.department=v.department;r.position=v.position;
        r.active=v.active;touch();renderAll();toast("บันทึกแล้ว");return true;});
    return;
  }
  if(a==="edp"||a==="rmp"||a==="arp"){
    var p=byId(DB.projects,id);if(!p)return;
    if(a==="arp"){p.archived=!p.archived;touch();renderAll();
      toast(p.archived?"เก็บเข้าคลังแล้ว":"กู้คืนแล้ว");return;}
    if(a==="rmp"){
      var u3=DB.tasks.filter(function(t){return t.projectId===id;}).length;
      ask("ลบโครงการ "+p.name+"?",u3?("มีงาน "+u3+
        " ชิ้นในโครงการนี้ งานจะยังอยู่แต่กลายเป็น “ไม่ระบุโครงการ”"):"ยังไม่มีงานในโครงการนี้","ลบ",
        function(){DB.projects=DB.projects.filter(function(x){return x.id!==id;});
          DB.tasks.forEach(function(t){if(t.projectId===id)t.projectId="";});
          touch();renderAll();toast("ลบแล้ว");});
      return;
    }
    openForm("แก้ไขโครงการ",[
      {key:"name",label:"ชื่อโครงการ",type:"text",value:p.name},
      {key:"description",label:"รายละเอียด",type:"textarea",value:p.description},
      {key:"status",label:"สถานะ",type:"select",value:p.status,
        options:[["active","กำลังดำเนินการ"],["done","ปิดโครงการ"]]},
      {key:"startAt",label:"วันที่เริ่ม",type:"date",value:p.startAt},
      {key:"endAt",label:"กำหนดสิ้นสุด",type:"date",value:p.endAt}],
      function(v){if(!v.name)return false;
        p.name=v.name;p.description=v.description;p.status=v.status;
        p.startAt=v.startAt;p.endAt=v.endAt;
        touch();renderAll();toast("บันทึกแล้ว");return true;});
    return;
  }
});

/* ---- chips / filters ---- */
$("status-chips").addEventListener("click",function(e){
  var b=e.target.closest("[data-st]");if(!b)return;
  flt.status=b.dataset.st;renderTasks();});
$("q").addEventListener("input",function(){flt.q=this.value;renderTasks();});
$("filter-btn").addEventListener("click",function(){openOv("ov-filter");});
$("filter-clear").addEventListener("click",function(){
  ["ownerId","projectId","priority","recvFrom","recvTo","dueFrom","dueTo"].forEach(function(k){flt[k]="";});
  renderTasks();});
document.addEventListener("change",function(e){
  if(!e.target||!e.target.closest)return;
  var el=e.target.closest("[data-f]");if(!el)return;
  flt[el.dataset.f]=el.value;renderTasks();});

/* ---- notes ---- */
var noteProj="";
$("note-save").addEventListener("click",function(){
  var txt=$("note-text").value.trim();
  if(!txt){toast("พิมพ์สิ่งที่ทำก่อน");return;}
  DB.notes.push({id:uid(),date:$("note-date").value||today(),text:txt,projectId:noteProj});
  $("note-text").value="";
  touch();renderToday();toast("บันทึกแล้ว");
});

/* ---- week ---- */
$("wk-prev").addEventListener("click",function(){curWeek=addD(curWeek,-7);renderWeek();});
$("wk-next").addEventListener("click",function(){curWeek=addD(curWeek,7);renderWeek();});
$("wk-now").addEventListener("click",function(){curWeek=weekStart(today());renderWeek();});
["wk-plan","wk-risk"].forEach(function(id){
  $(id).addEventListener("change",function(){
    DB.weeks[curWeek]=DB.weeks[curWeek]||{};
    DB.weeks[curWeek][id==="wk-plan"?"plan":"risk"]=this.value;
    touch();renderWeek();});});
var showText=false;
$("wk-toggle").addEventListener("click",function(){
  showText=!showText;
  $("week-text").hidden=!showText;$("week-report").hidden=showText;
  this.textContent=showText?"ดูเป็นรายงาน":"ดูเป็นข้อความ";});
$("wk-copy").addEventListener("click",function(){
  var ta=$("week-text");
  var p=(navigator.clipboard&&navigator.clipboard.writeText)?
    navigator.clipboard.writeText(ta.value):Promise.reject();
  p.then(function(){toast("คัดลอกรายงานแล้ว");}).catch(function(){
    ta.hidden=false;$("week-report").hidden=true;showText=true;
    $("wk-toggle").textContent="ดูเป็นรายงาน";ta.select();toast("กด Ctrl+C เพื่อคัดลอก");});
});
if(navigator.share){
  $("wk-share").hidden=false;
  $("wk-share").addEventListener("click",function(){
    navigator.share({title:$("week-h1").textContent,text:$("week-text").value}).catch(function(){});});
}

/* ---- print ---- */
["p-scope","p-group","p-proj","p-member","p-status"].forEach(function(id){
  $(id).addEventListener("change",renderPrint);});
$("p-prev").addEventListener("click",function(){curWeek=addD(curWeek,-7);renderPrint();});
$("p-next").addEventListener("click",function(){curWeek=addD(curWeek,7);renderPrint();});
$("p-now").addEventListener("click",function(){curWeek=weekStart(today());renderPrint();});
$("p-print").addEventListener("click",function(){
  try{window.print();}catch(e){toast("ลองกด Ctrl+P");}});

/* ---- settings ---- */
$("s-owner").addEventListener("change",function(){
  DB.settings.owner=this.value.trim();
  var self=DB.members.filter(function(m){return m.id===SELF;})[0];
  if(self)self.name=DB.settings.owner||"หัวหน้าแผนก";
  touch();renderAll();});
$("s-due").addEventListener("change",function(){
  DB.settings.defaultDueDays=this.value;touch();});
$("s-cycle").addEventListener("change",function(){
  DB.settings.defaultReportCycle=this.value;touch();});
$("theme-seg").addEventListener("click",function(e){
  var b=e.target.closest("[data-theme]");if(!b)return;
  DB.settings.theme=b.dataset.theme;applyTheme();touch();});
$("add-member").addEventListener("click",function(){
  openForm("เพิ่มสมาชิก",[
    {key:"name",label:"ชื่อ",type:"text",value:""},
    {key:"nickname",label:"ชื่อเล่น (ไม่ใส่ก็ได้)",type:"text",value:""},
    {key:"department",label:"แผนก (ไม่ใส่ก็ได้)",type:"text",value:""},
    {key:"position",label:"ตำแหน่ง (ไม่ใส่ก็ได้)",type:"text",value:""}],
    function(v){if(!v.name)return false;
      DB.members.push({id:uid(),name:v.name,nickname:v.nickname,department:v.department,
        position:v.position,active:true});
      touch();renderAll();toast("เพิ่มสมาชิกแล้ว");return true;});});
$("add-requester").addEventListener("click",function(){
  openForm("เพิ่มผู้มอบหมายงาน",[
    {key:"name",label:"ชื่อ",type:"text",value:""},
    {key:"department",label:"แผนก (ไม่ใส่ก็ได้)",type:"text",value:""},
    {key:"position",label:"ตำแหน่ง (ไม่ใส่ก็ได้)",type:"text",value:""}],
    function(v){if(!v.name)return false;
      DB.requesters.push({id:uid(),name:v.name,department:v.department,position:v.position,active:true});
      touch();renderAll();toast("เพิ่มแล้ว");return true;});});
$("import-requesters").addEventListener("click",function(){$("requester-import-file").click();});
$("requester-import-file").addEventListener("change",function(){
  var f=this.files&&this.files[0];this.value="";if(!f)return;
  parseRequesterFile(f).then(function(list){
    if(!list.length){toast("ไม่พบรายชื่อในไฟล์");return;}
    var added=0,updated=0;
    list.forEach(function(item){
      var nm=normName(item.name);if(!nm)return;
      var ex=DB.requesters.filter(function(r){return normName(r.name)===nm;})[0];
      if(ex){
        if(item.department)ex.department=item.department;
        if(item.position)ex.position=item.position;
        updated++;
      }else{
        DB.requesters.push({id:uid(),name:nm,department:item.department||"",
          position:item.position||"",active:true});
        added++;
      }
    });
    touch();renderAll();
    toast("นำเข้าแล้ว — เพิ่มใหม่ "+added+" คน อัปเดต "+updated+" คน");
  }).catch(function(err){toast("อ่านไฟล์ไม่สำเร็จ: "+(err&&err.message||err));});
});
$("add-project").addEventListener("click",function(){
  openForm("เพิ่มโครงการ",[
    {key:"name",label:"ชื่อโครงการ",type:"text",value:""},
    {key:"description",label:"รายละเอียด",type:"textarea",value:""},
    {key:"startAt",label:"วันที่เริ่ม",type:"date",value:""},
    {key:"endAt",label:"กำหนดสิ้นสุด",type:"date",value:""}],
    function(v){if(!v.name)return false;
      DB.projects.push({id:uid(),name:v.name,description:v.description,status:"active",
        startAt:v.startAt,endAt:v.endAt,archived:false});
      touch();renderAll();toast("เพิ่มโครงการแล้ว");return true;});});
$("cl-age").addEventListener("change",renderSettings);
$("cl-done").addEventListener("click",function(){
  var hit=doneOlder(parseInt($("cl-age").value,10));
  if(!hit.length){toast("ไม่มีงานที่เข้าเงื่อนไข");return;}
  ask("ล้างงานที่ปิดแล้ว "+hit.length+" ชิ้น?",
    "งานที่ยังไม่ปิดไม่ถูกแตะต้อง แต่รายงานย้อนหลังจะไม่เห็นงานเหล่านี้อีก กู้คืนไม่ได้",
    "ล้าง "+hit.length+" งาน",function(){
      var ids={};hit.forEach(function(t){ids[t.id]=1;});
      DB.tasks=DB.tasks.filter(function(t){return !ids[t.id];});
      afterWipe("ล้างแล้ว "+hit.length+" งาน");});});
$("cl-tasks").addEventListener("click",function(){
  var n=DB.tasks.length,m=DB.notes.length;
  if(!n&&!m){toast("ยังไม่มีข้อมูลให้ล้าง");return;}
  ask("ล้างงานและบันทึกทั้งหมด?","จะลบงาน "+n+" ชิ้น บันทึก "+m+
    " รายการ และสรุปรายสัปดาห์ทั้งหมด — ชื่อคุณ ทีม ผู้มอบหมายงาน และโครงการยังอยู่ครบ กู้คืนไม่ได้",
    "ล้างงานทั้งหมด",function(){
      DB.tasks=[];DB.notes=[];DB.weeks={};afterWipe("ล้างงานทั้งหมดแล้ว");});});
$("cl-all").addEventListener("click",function(){
  ask("ล้างทุกอย่าง เริ่มต้นใหม่?",
    "กลับไปเป็นสมุดเปล่า ลบทั้งงาน บันทึก ทีม ผู้มอบหมายงาน และโครงการ กู้คืนไม่ได้ — ถ้ายังไม่สำรองไฟล์ กดยกเลิกก่อน",
    "ล้างทุกอย่าง",function(){
      var th=DB.settings.theme,ow=DB.settings.owner;
      var self=DB.members.filter(function(m){return m.id===SELF;})[0];
      DB=blankDB();DB.settings.theme=th;DB.settings.owner=ow;
      if(self)DB.members=[self];
      flt.q="";$("q").value="";afterWipe("เริ่มต้นใหม่แล้ว");});});
$("do-export").addEventListener("click",function(){
  var d=new Date();
  var name="team-task-book-backup-"+(d.getFullYear()+543)+"-"+d2(d.getMonth()+1)+"-"+d2(d.getDate())+".json";
  var blob=new Blob([JSON.stringify(DB,null,2)],{type:"application/json"});
  var a=document.createElement("a");
  a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
  setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},1000);
  toast("ดาวน์โหลดไฟล์สำรองแล้ว");
});
$("do-import").addEventListener("click",function(){$("import-file").click();});
$("import-file").addEventListener("change",function(){
  var f=this.files&&this.files[0];if(!f)return;
  var r=new FileReader();
  r.onload=function(){
    var parsed=null;
    try{parsed=migrate(JSON.parse(r.result));}catch(e){}
    if(!parsed||!parsed.settings){toast("ไฟล์ไม่ถูกต้อง");return;}
    ask("นำข้อมูลกลับเข้ามา?","ข้อมูลปัจจุบันจะถูกแทนที่ด้วยไฟล์นี้ ("+
      parsed.tasks.length+" งาน / "+parsed.members.length+" สมาชิก) กู้คืนไม่ได้","แทนที่ข้อมูล",
      function(){
        var self=DB.members.filter(function(m){return m.id===SELF;})[0];
        parsed=reidentify(parsed);
        parsed.members=parsed.members.filter(function(m){return m.id!==SELF;});
        if(self)parsed.members.unshift(self);
        parsed.settings.theme=DB.settings.theme;
        DB=parsed;applyTheme();afterWipe("นำข้อมูลกลับเข้ามาแล้ว");});
  };
  r.readAsText(f);this.value="";
});

/* ==========================================================
   23. เริ่มระบบ
   ========================================================== */
function reidentify(db){
  var map={};
  function nid(x){if(!x)return "";if(!map[x])map[x]=uid();return map[x];}
  db.members.forEach(function(m){m.id=nid(m.id);});
  db.requesters.forEach(function(r){r.id=nid(r.id);});
  db.projects.forEach(function(p){p.id=nid(p.id);});
  db.tasks.forEach(function(t){
    t.id=nid(t.id);
    t.ownerId=(!t.ownerId||t.ownerId==="me")?SELF:nid(t.ownerId);
    t.collaboratorIds=(t.collaboratorIds||[]).map(function(x){
      return (x==="me")?SELF:nid(x);});
    t.requesterId=t.requesterId?nid(t.requesterId):"";
    t.projectId=t.projectId?nid(t.projectId):"";
    (t.updates||[]).forEach(function(u){u.id=uid();u.taskId=t.id;});
  });
  db.notes.forEach(function(n){n.id=uid();n.projectId=n.projectId?nid(n.projectId):"";});
  var r=db.settings.recent||{};
  ["member","project","requester"].forEach(function(k){
    r[k]=(r[k]||[]).map(function(x){return x==="me"?SELF:(map[x]||"");}).filter(Boolean);});
  db.settings.recent=r;
  return db;
}
function newest(list){
  var best=null;
  list.forEach(function(x){
    var db=null;
    try{db=migrate(x);}catch(e){db=null;}
    if(!db)return;
    if(!best||(db.updatedAt||"")>(best.updatedAt||""))best=db;});
  return best;
}
/* ---------- Supabase / Auth ---------- */
var SBURL="__SB_URL__",SBKEY="__SB_KEY__";
function authMsg(t,bad){var e=$("auth-msg");e.textContent=t;e.style.color=bad?"var(--hot)":"var(--ink-2)";}
function showAuth(show){$("ov-auth").hidden=!show;$("app-root").style.visibility=show?"hidden":"visible";}

function rebuild(rows){
  var db=blankDB();
  var st=rows.settings||{};
  db.settings.owner=st.display_name||"";
  db.settings.theme=st.theme||"system";
  db.settings.defaultDueDays=st.default_due_days||"";
  db.settings.defaultReportCycle=st.default_report_cycle||"";
  db.settings.recent=st.recent||{member:[],project:[],requester:[]};
  var selfRow=null;
  (rows.members||[]).forEach(function(m){
    if(m.is_self){selfRow=m;SELF=m.id;}
    db.members.push({id:m.id,name:m.name,nickname:m.nickname||"",active:m.active!==false,
      department:m.department||"",position:m.position||"",
      isSelf:!!m.is_self});
  });
  db.requesters=(rows.requesters||[]).map(function(r){
    return {id:r.id,name:r.name,department:r.department||"",position:r.position||"",
      active:r.active!==false};});
  db.projects=(rows.projects||[]).map(function(p){
    return {id:p.id,name:p.name,description:p.description||"",status:p.status||"active",
      startAt:p.start_at||"",endAt:p.end_at||"",archived:!!p.archived};});
  var upMap={};
  (rows.updates||[]).forEach(function(u){
    (upMap[u.task_id]=upMap[u.task_id]||[]).push(
      {id:u.id,taskId:u.task_id,message:u.message||"",status:u.status||"",createdAt:fromISO(u.created_at)});});
  var coMap={};
  (rows.collabs||[]).forEach(function(c){
    (coMap[c.task_id]=coMap[c.task_id]||[]).push(c.member_id);});
  db.tasks=(rows.tasks||[]).map(function(t){
    return {id:t.id,title:t.title,description:t.description||"",
      requesterId:t.requester_id||"",ownerId:t.assignee_id||SELF,
      collaboratorIds:coMap[t.id]||[],projectId:t.project_id||"",
      status:t.status,priority:t.priority,receivedAt:fromISO(t.received_at),
      dueAt:t.due_at||"",reportCycle:t.report_cycle||"",repeatNext:!!t.repeat_next,
      createdAt:fromISO(t.created_at),updatedAt:fromISO(t.updated_at),
      completedAt:t.completed_at||"",updates:(upMap[t.id]||[])};});
  db.notes=(rows.notes||[]).map(function(n){
    return {id:n.id,date:n.note_date,text:n.body||"",projectId:n.project_id||""};});
  (rows.weeks||[]).forEach(function(w){
    db.weeks[w.week_start]={plan:w.plan||"",risk:w.risk||""};});
  if(!selfRow&&db.members.length===0){
    var sid=uid();SELF=sid;
    db.members.push({id:sid,name:db.settings.owner||"หัวหน้าแผนก",nickname:"",active:true,isSelf:true});
  }
  return normalize(db);
}
function loadAll(){
  setSave("wait","กำลังโหลดข้อมูล…");
  return Promise.all([
    SB.from("settings").select("*").maybeSingle(),
    SB.from("members").select("*").order("sort_order"),
    SB.from("requesters").select("*").order("sort_order"),
    SB.from("projects").select("*").order("sort_order"),
    SB.from("tasks").select("*").order("received_at"),
    SB.from("task_collaborators").select("*"),
    SB.from("task_updates").select("*").order("created_at"),
    SB.from("notes").select("*").order("note_date"),
    SB.from("week_notes").select("*")
  ]).then(function(r){
    r.forEach(function(x){if(x.error&&x.error.code!=="PGRST116")throw x.error;});
    DB=rebuild({settings:r[0].data,members:r[1].data,requesters:r[2].data,projects:r[3].data,
      tasks:r[4].data,collabs:r[5].data,updates:r[6].data,notes:r[7].data,weeks:r[8].data});
    snap=clone(DB);
    snap.__setRow={owner_id:UID,display_name:DB.settings.owner||"",theme:DB.settings.theme||"system",
      default_due_days:String(DB.settings.defaultDueDays||""),
      default_report_cycle:DB.settings.defaultReportCycle||"",recent:DB.settings.recent};
    saveLocal();applyTheme();
    $("note-date").value=today();
    bootDone=true;
    setSave("ok","พร้อมใช้งาน");
    renderAll();
  });
}
function onSignedIn(session){
  USER=session.user;UID=USER.id;
  $("who").textContent=USER.email||"";
  showAuth(false);
  storeNote();
  loadAll().catch(function(e){
    console.error(e);setSave("bad","โหลดข้อมูลไม่สำเร็จ");
    authMsg("โหลดข้อมูลไม่สำเร็จ: "+(e.message||e),true);
  });
}
function boot(){
  SB=window.supabase.createClient(SBURL,SBKEY,{auth:{persistSession:true,autoRefreshToken:true}});
  SB.auth.getSession().then(function(r){
    if(r.data&&r.data.session)onSignedIn(r.data.session);
    else {showAuth(true);setSave("","ยังไม่ได้เข้าสู่ระบบ");}
  });
  SB.auth.onAuthStateChange(function(ev,session){
    if(ev==="SIGNED_OUT"){USER=null;UID="";DB=blankDB();snap=null;showAuth(true);}
  });
}
$("auth-form").addEventListener("submit",function(e){
  e.preventDefault();
  var email=$("auth-email").value.trim(),pw=$("auth-pw").value;
  if(!email||!pw){authMsg("กรอกอีเมลและรหัสผ่านให้ครบ",true);return;}
  $("auth-go").disabled=true;
  var signup=$("auth-form").dataset.mode==="signup";
  authMsg(signup?"กำลังสมัคร…":"กำลังเข้าสู่ระบบ…");
  var p=signup?SB.auth.signUp({email:email,password:pw,
        options:{data:{display_name:$("auth-name").value.trim()}}})
    :SB.auth.signInWithPassword({email:email,password:pw});
  p.then(function(r){
    $("auth-go").disabled=false;
    if(r.error){authMsg(r.error.message,true);return;}
    if(r.data.session)onSignedIn(r.data.session);
    else authMsg("สมัครแล้ว — เปิดอีเมลเพื่อยืนยัน แล้วกลับมาเข้าสู่ระบบ");
  }).catch(function(err){$("auth-go").disabled=false;authMsg(String(err&&err.message||err),true);});
});
$("auth-toggle").addEventListener("click",function(){
  var f=$("auth-form"),su=f.dataset.mode==="signup";
  f.dataset.mode=su?"login":"signup";
  $("auth-title").textContent=su?"เข้าสู่ระบบ":"สมัครใช้งานครั้งแรก";
  $("auth-name-wrap").hidden=su;
  $("auth-go").textContent=su?"เข้าสู่ระบบ":"สมัครและเริ่มใช้งาน";
  this.textContent=su?"ยังไม่มีบัญชี? สมัครใช้งานครั้งแรก":"มีบัญชีแล้ว? เข้าสู่ระบบ";
  authMsg("");
});
$("more-logout").addEventListener("click",function(){closeOv("ov-more");$("logout").click();});
$("logout").addEventListener("click",function(){
  ask("ออกจากระบบ?","ข้อมูลที่บันทึกแล้วอยู่บนเซิร์ฟเวอร์ครบ เข้าใหม่เมื่อไหร่ก็เห็นเหมือนเดิม","ออกจากระบบ",
    function(){SB.auth.signOut();});
});

try{var v0=sessionStorage.getItem("twb.view");if(v0&&VIEWS.indexOf(v0)>=0)view=v0;}catch(e){}
$("note-date").value=today();
go(view);
boot();
})();

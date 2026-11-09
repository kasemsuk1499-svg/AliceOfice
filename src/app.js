(function(){
'use strict';
const C=()=>window.AliceCore;
const $=id=>document.getElementById(id);
const roles=[['planSTP','แผน STP 1 (CSV)'],['planPTR','แผน STP 2 / PTR (CSV)'],['formSTP','ใบแจ้งโหลด STP'],['formPTR','ใบแจ้งโหลด PTR'],['loadingSTP','Loading STP'],['loadingPTR','Loading PTR'],['exportSTP','Export STP'],['exportPTR','Export PTR'],['truck','ข้อมูลรถขนส่ง'],['ignore','ไม่ใช้งาน']];
let mode='prep',files=[],counter=0,prepared=null,outputs=[],busy=false;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=n=>Number(n||0).toLocaleString('th-TH');
const extension=n=>(n.match(/\.([^.]+)$/)||[])[1]?.toLowerCase()||'';
const mime=n=>extension(n)==='xlsx'?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':extension(n)==='csv'?'text/csv':'application/octet-stream';
function log(message){const el=$('logs');const now=new Date().toLocaleTimeString('th-TH',{hour12:false});el.textContent+='\n['+now+'] '+message;el.scrollTop=el.scrollHeight}
function status(text,kind='info'){$('status').textContent=text;$('status').className='status'+(kind==='error'?' error':kind==='good'?' good':'')}
function stat(n,bookings,units,warns){$('statFiles').textContent=n;$('statBookings').textContent=bookings??'—';$('statUnits').textContent=units??'—';$('statWarnings').textContent=warns??'—'}
function get(role){return files.filter(f=>f.role===role)}
function sole(role){const list=get(role);if(list.length>1)throw Error('มีไฟล์ประเภท '+roles.find(x=>x[0]===role)?.[1]+' มากกว่า 1 ไฟล์ กรุณาเลือกให้เหลือไฟล์ที่ถูกต้อง');return list[0]?.file||null}
function dt(iso){const [year,month,day]=iso.split('-');return day+'/'+month+'/'+year}
function stamp(iso){return dt(iso).replaceAll('/','-')}
function validDate(){const iso=$('workDate').value;if(!/^\d{4}-\d{2}-\d{2}$/.test(iso))throw Error('กรุณาเลือกวันที่ก่อน');return iso}
function renderFiles(){
 $('fileRows').innerHTML=files.length?files.map(e=>'<tr data-id="'+e.id+'"><td><strong>'+esc(e.file.name)+'</strong><div style="color:#91a2b5;font-size:10px">'+fmt(e.file.size/1024)+' KB</div></td><td><select class="role" aria-label="ประเภทไฟล์">'+roles.map(r=>'<option value="'+r[0]+'" '+(e.role===r[0]?'selected':'')+'>'+r[1]+'</option>').join('')+'</select></td><td><button class="remove" type="button" aria-label="ลบไฟล์">×</button></td></tr>').join(''):'<tr><td colspan="3" class="empty">ยังไม่ได้เพิ่มไฟล์</td></tr>';
 $('fileRows').querySelectorAll('tr[data-id]').forEach(tr=>{
  const ent=files.find(x=>String(x.id)===tr.dataset.id);
  tr.querySelector('select').addEventListener('change',e=>{ent.role=e.target.value;resetPreview();log('เปลี่ยนชนิดไฟล์ '+ent.file.name)});
  tr.querySelector('.remove').addEventListener('click',()=>{files=files.filter(x=>x!==ent);renderFiles();resetPreview()});
 });
 stat(files.length,prepared?.bookings,prepared?.units,prepared?.warnings?.length);
}
function resetPreview(){prepared=null;outputs=[];$('previewCard').classList.add('hidden');$('resultsCard').classList.add('hidden');$('statBookings').textContent='—';$('statUnits').textContent='—';$('statWarnings').textContent='—'}
async function acceptFiles(list){
 if(busy)return;
 for(const f of Array.from(list||[])){
  if(!f||!f.name)continue;
  if(f.name.startsWith('~$')||f.name.startsWith('.'))continue;
  if(extension(f.name)==='zip'){
   if(f.size>80*1024*1024){log('ข้าม ZIP ที่ใหญ่เกิน 80 MB: '+f.name);continue}
   if(!window.JSZip){status('JSZip ยังไม่โหลด กรุณาลองอีกครั้ง','error');return}
   try{
    const zip=await JSZip.loadAsync(await f.arrayBuffer());let added=0;
    for(const [path,z] of Object.entries(zip.files)){
     if(z.dir||/__MACOSX|\.DS_Store|\/\.\w/.test(path)||!/\.(csv|xlsx|xls)$/i.test(path))continue;
     const arr=await z.async('uint8array');if(arr.byteLength>32*1024*1024){log('ข้ามไฟล์ใหญ่เกิน 32MB: '+path);continue}
     const name=path.split('/').pop();
     files.push({id:++counter,file:new File([arr],name,{type:mime(name)}),role:C().classify({name})});added++;
    }
    log('นำเข้า ZIP '+f.name+' สำเร็จ '+added+' ไฟล์');
   }catch(e){log('ZIP ไม่ถูกต้อง: '+e.message);status('ไม่สามารถอ่าน ZIP '+f.name,'error')}
   continue;
  }
  if(!/\.(csv|xlsx|xls)$/i.test(f.name)){log('ข้ามไฟล์ '+f.name);continue}
  if(files.some(x=>x.file.name===f.name&&x.file.size===f.size)){log('ข้ามไฟล์ชื่อซ้ำ '+f.name);continue}
  files.push({id:++counter,file:f,role:C().classify(f)});
 }
 renderFiles();resetPreview();if(files.length)status('นำเข้า '+files.length+' ไฟล์แล้ว เลือกวันที่และกดตรวจสอบได้เลยค่ะ');
}
function showMode(next){
 mode=next;resetPreview();
 $('tabPrep').classList.toggle('active',mode==='prep');$('tabDaily').classList.toggle('active',mode==='daily');
 $('uploadTitle').textContent=mode==='prep'?'1. นำเข้าไฟล์สำหรับเตรียมโหลด':'1. นำเข้า Loading / Export และไฟล์ขนส่ง';
 $('uploadDesc').textContent=mode==='prep'?'ใช้ CSV แผน STP1/STP2 และแบบใบแจ้งโหลด STP/PTR (.xlsx)':'ใช้ Loading และ Export Excel ของวันที่ทำงาน พร้อมไฟล์รถขนส่งและ CSV แผน (ถ้ามี) เพื่อซิงก์ข้อมูลแต่ละคัน';
 $('rulesTitle').textContent=mode==='prep'?'กติกาเตรียมโหลด':'กติกา Daily Report';
 $('rulesPrep').classList.toggle('hidden',mode!=='prep');$('rulesDaily').classList.toggle('hidden',mode!=='daily');
 $('statUnitLabel').textContent=mode==='prep'?'จำนวนตู้':'จำนวนคัน';
 $('rulesNotice').innerHTML=mode==='prep'?'แม่แบบ STP/PTR ควรเป็น <b>ไฟล์ Excel เดิมที่มี Text Box</b> เพื่อรักษารูปแบบต้นฉบับ รวมทั้งหัวข้อ Gate No. ของ PTR':'ระบบจะซิงก์ข้อมูลจาก <b>Loading ไป Export</b> และเสริมเลขรถ/ตู้/Seal/Tare จากรายชื่อขนส่ง เมื่อจับคู่ Booking และลำดับรถได้แน่ชัด';
 status('เลือกไฟล์สำหรับ '+(mode==='prep'?'งานเตรียมโหลด':'Daily Report')+' แล้วกดตรวจสอบได้เลยค่ะ');
}
async function readPlans(iso){
 const result={STP:[],PTR:[],ignored:[],warnings:[]};
 for(const [role,kind] of [['planSTP','STP'],['planPTR','PTR']]){
  const file=sole(role);if(!file)continue;
  let text=await file.text();if(text.includes('\ufffd')&&window.TextDecoder){text=new TextDecoder('windows-874').decode(await file.arrayBuffer())}
  const records=C().csvParse(text);
  const r=C().planRecords(records,iso,kind);
  result[kind].push(...r.jobs);result.ignored.push(...r.ignored);
  log(kind+' แผนวันที่ '+dt(iso)+' '+r.jobs.length+' รายการ / ข้าม '+r.ignored.length+' รายการ');
 }
 return result;
}
function createPages(jobs){
 const pages=[];
 for(const job of jobs){
  for(let i=1;i<=job.fcl;i+=2){pages.push([{job,no:i},{job,no:i+1<=job.fcl?i+1:null}])}
 }
 return pages;
}
function sheetTitle(job,n){
 const productSuffix=job.booking==='035GX75579'?' '+job.product.replace(/[^A-Za-z0-9-]/g,'').slice(0,8):'';
 const base=(job.booking+productSuffix).slice(0,28);
 return (base+(n===1?'':' ('+n+')')).slice(0,31);
}
function parseRequiredCount(jobs){return jobs.reduce((sum,x)=>sum+x.fcl,0)}
function analyzePrep(plans){
 const jobs=[...plans.STP,...plans.PTR],warnings=[];
 if(!jobs.length)throw Error('ไม่พบแผนงานตรงวันที่และเงื่อนไข KNS — กรุณาตรวจวันที่หรือประเภท CSV');
 for(const kind of ['STP','PTR']){
  if(plans[kind].length&&!get(kind==='STP'?'formSTP':'formPTR').length)warnings.push('ยังไม่มีไฟล์แม่แบบใบแจ้งโหลด '+kind+' — จะยังสร้างไฟล์ฝั่งนี้ไม่ได้');
  for(const j of plans[kind]){
   if(!j.booking)warnings.push(kind+' มีแถวที่ Booking ว่าง');
   if(!j.packaging&&kind==='STP')warnings.push(j.booking+' ไม่มี Packaging and Batch no.');
   if(!j.bags)warnings.push(j.booking+' '+j.product+' ไม่ทราบจำนวนถุงต่อตู้ ต้องระบุก่อนสร้าง');
  }
 }
 const cards=jobs.flatMap(j=>Array.from({length:j.fcl},(_,i)=>({kind:j.kind,job:j,no:i+1,bags:j.bags||'',category:j.bags===20?'Pallet':'แบก'})));
 return {mode:'prep',jobs,plans,cards,warnings,bookings:new Set(jobs.map(j=>j.kind+':'+j.booking)).size,units:cards.length};
}
async function readManifest(file){
 const wb=await C().parseXlsx(file),found=[];
 const name=file.name;
 let booking=(name.match(/COAU\d{9,}|TH\d{8,}|\b\d{3}GX\d+\b|\b\d{10}\b/i)||[])[0]||'';
 let carrier=/MEECHOKE|มีโชค/i.test(name)?'MCT':/PANAWAN|พนาวรรณ/i.test(name)?'PA':/จิรภัทร/i.test(name)?'จิรภัทร':/KNS/i.test(name)?'KNS':'';
 for(const sheet of wb.SheetNames){
  const rows=C().sheetRows(wb,sheet);
  for(let index=0;index<rows.length;index++){
   const r=rows[index],no=Number(r[0]),cont=C().clean(r[4]);
   if(!Number.isInteger(no)||no<1||no>999||!cont||!/^[A-Z]{4}\d{7}$/i.test(cont))continue;
   const truck={booking,carNo:no,registration:[r[1],r[2]].map(C().clean).filter(Boolean).join('/'),driver:C().clean(r[3]),container:cont.toUpperCase(),seal:C().clean(r[5]),tare:Number(r[6])||null,carrier};
   found.push(truck);
  }
 }
 return found;
}
function chooseLoadingSheet(wb,iso){
 const [,mm,dd]=iso.split('-');
 const wanted=mm+'-'+dd;
 if(wb.SheetNames.includes(wanted))return wanted;
 const found=wb.SheetNames.filter(s=>{
  const rows=C().sheetRows(wb,s);
  return rows.some(row=>C().clean(row[7]).toLowerCase()==='no.'&&/container/i.test(C().clean(row[10])));
 });
 if(found.length===1)return found[0];
 throw Error('ไม่พบชีตวันที่ '+wanted+' ใน Loading หรือมีหลายชีตที่ระบุไม่ชัดเจน');
}
function extractLoading(wb,kind,iso,plan,manifests){
 const sheetName=chooseLoadingSheet(wb,iso),rows=C().sheetRows(wb,sheetName);
 const vehicles=[];
 let context={contract:'',booking:'',agent:'',etd:'',destination:'',returnPlace:''},current=null;
 const planMap=new Map(plan.map(j=>[j.booking,j]));
 function validBooking(book){return /^[A-Z0-9][A-Z0-9/-]{5,}$/i.test(book)&&/\d/.test(book)&&!/^(quantity|booking|date)/i.test(book)}
 function pushItem(row){
  const product=C().clean(row[12]||row[1]),batch=C().clean(row[14]),code=C().clean(row[15]);
  const contract=C().clean(row[0]);
  const bags=Number(row[18])||0,kg=Number(row[19])||0;
  if(product||batch||code||bags||kg)current.items.push({product,batch,code,bags,kg,contract:/^S?PTR?\d{4}|^TSP\d{4}/i.test(contract)?contract:''});
 }
 for(let k=0;k<rows.length;k++){
  const row=rows[k],no=Number(row[7]),book=C().clean(row[3]);
  const bookingPresent=validBooking(book);
  const header=C().clean(row[7]).toLowerCase()==='no.';
  if(bookingPresent){
   context={...context,booking:book,contract:C().clean(row[0])||context.contract,agent:C().clean(row[4])||context.agent,etd:C().clean(row[5])||context.etd,destination:'',returnPlace:''};
  }
  if(header){current=null;continue}
  const isCar=Number.isInteger(no)&&no>=1&&no<=200&&C().clean(row[7])!=='';
  if(isCar){
   if(!context.booking)continue;
   const planJob=planMap.get(context.booking);
   current={kind,sheetName,row:k+1,carNo:no,booking:context.booking,contract:context.contract||planJob?.contract||'',agent:context.agent||planJob?.agent||'',etd:planJob?.etd||context.etd,destination:planJob?.destination||context.destination,returnPlace:planJob?.place||context.returnPlace,carrier:C().clean(row[8]),registration:C().clean(row[9]),container:C().clean(row[10]),seal:C().clean(row[11]),weightCard:C().clean(row[20]),tare:Number(row[21])||null,gross:Number(row[22])||null,net:Number(row[23])||null,driver:'',items:[],sourceName:sheetName,hasManifest:false};
   vehicles.push(current);pushItem(row);
  }else if(current){
   if(bookingPresent&&book!==current.booking){current=null;continue}
   const f=C().clean(row[5]);
   if(/^[A-Z][A-Z -]{2,32}$/.test(f)&&!/ETD|FCL|CONTRACT/i.test(f)){
    context.destination=f;
    current.destination ||= f;
   }
   if(C().clean(row[12])||C().clean(row[14])||C().clean(row[15])||Number(row[18])||Number(row[19]))pushItem(row);
  }
 }
 for(const v of vehicles){
  const truck=manifests.find(m=>m.booking===v.booking&&m.carNo===v.carNo&&(m.container===v.container||!v.container))||
              manifests.find(m=>m.booking===v.booking&&m.carNo===v.carNo);
  if(truck){
   v.hasManifest=true;v.registration ||= truck.registration;v.container ||= truck.container;v.seal ||= truck.seal;v.tare ||= truck.tare;v.carrier ||= truck.carrier;v.driver ||= truck.driver;
  }
  v.contract=Array.from(new Set([v.contract,...v.items.map(i=>i.contract)].filter(Boolean))).join(' / ');
  v.product=Array.from(new Set(v.items.map(i=>i.product).filter(Boolean))).join(' / ');
  v.batch=Array.from(new Set(v.items.map(i=>i.batch).filter(Boolean))).join(' / ');
  v.code=v.items.filter(i=>i.code).map(i=>i.code+(i.bags?'('+i.bags+'Bag)':'')).join(' / ');
  v.bags=v.items.reduce((a,i)=>a+(i.bags||0),0);
  v.kg=v.items.reduce((a,i)=>a+(i.kg||0),0);
 }
 return {sheetName,vehicles};
}
async function analyzeDaily(iso,plans){
 const filesByKind={},all=[],warnings=[];
 const manifests=[];
 for(const en of get('truck'))try{const found=await readManifest(en.file);manifests.push(...found);log('ไฟล์ขนส่ง '+en.file.name+' พบ '+found.length+' คัน')}catch(e){warnings.push('ไฟล์ขนส่ง '+en.file.name+': '+e.message)}
 for(const kind of ['STP','PTR']){
  const load=sole('loading'+kind),exp=sole('export'+kind);
  if(!load&&!exp)continue;
  if(!load||!exp){warnings.push('มีไฟล์ '+kind+' ไม่ครบคู่ Loading / Export');continue}
  const wb=await C().parseXlsx(load);
  const parsed=extractLoading(wb,kind,iso,plans[kind],manifests);
  const allowed=plans[kind].length?new Set(plans[kind].map(j=>j.booking)):null;
  const vehicles=parsed.vehicles.filter(v=>!allowed||allowed.has(v.booking));
  all.push(...vehicles);filesByKind[kind]={load,exp,sheetName:parsed.sheetName,vehicles};
  log(kind+' พบ '+vehicles.length+' คันจาก Loading ('+parsed.sheetName+')');
 }
 if(!all.length)throw Error('ไม่พบรถใน Loading สำหรับจัดทำรายงาน — ตรวจชื่อชีตวันที่ และอัปโหลด Loading + Export ให้ครบคู่');
 const dup=new Set();
 for(const v of all){
  if(!v.booking)warnings.push(v.kind+' รถ '+v.carNo+' ไม่มี Booking');
  if(!v.container)warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังไม่มีเลข Container');
  if(!v.registration)warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังไม่มีทะเบียนรถ');
  if(!v.weightCard||!v.gross||!v.net)warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังมีข้อมูลชั่งไม่ครบ — จะเว้นว่าง');
  const key=v.kind+'|'+v.booking+'|'+v.carNo;
  if(dup.has(key))warnings.push('พบลำดับรถซ้ำ: '+key);
  dup.add(key);
 }
 return {mode:'daily',plans,filesByKind,cards:all,warnings,bookings:new Set(all.map(x=>x.kind+':'+x.booking)).size,units:all.length};
}
function renderPreview(result){
 prepared=result;outputs=[];$('resultsCard').classList.add('hidden');$('previewCard').classList.remove('hidden');
 $('previewHint').textContent=result.mode==='prep'?'ข้อมูลที่ผ่านเงื่อนไข Load = KNS และไม่ Cancel — แก้จำนวนถุงได้หากแผนมีการเปลี่ยนแปลง':'ข้อมูลจาก Loading รายคัน — สามารถแก้ช่องรถและบัตรชั่งก่อนสร้าง Export ได้';
 const head=result.mode==='prep'?['สาย','Booking','สินค้า','คันที่','ถุง / คัน','วิธีโหลด','ตรวจสอบ']:['สาย','Booking','คัน','ทะเบียน','เลขตู้','Seal','Tare','Gross','Net','Weight Card'];
 $('previewHead').innerHTML='<tr>'+head.map(x=>'<th>'+x+'</th>').join('')+'</tr>';
 if(result.mode==='prep'){
  $('previewRows').innerHTML=result.cards.map((c,i)=>'<tr><td><span class="pill">'+c.kind+'</span></td><td><b>'+esc(c.job.booking)+'</b></td><td>'+esc(c.job.product)+'</td><td>'+c.no+'/'+c.job.fcl+'</td><td><input type="number" min="1" step="1" class="bagInput" data-i="'+i+'" value="'+esc(c.bags)+'" placeholder="ระบุจำนวน"></td><td><span class="pill '+(!c.bags?'pending':'ok')+'" id="method'+i+'">'+esc(c.category)+'</span></td><td>'+(!c.bags?'<span class="pill pending">ตรวจจำนวน</span>':c.job.confidence==='review'?'<span class="pill pending">ทบทวน</span>':'<span class="pill ok">พร้อม</span>')+'</td></tr>').join('');
  $('previewRows').querySelectorAll('.bagInput').forEach(el=>el.addEventListener('change',()=>{const i=Number(el.dataset.i),n=Number(el.value);result.cards[i].bags=Number.isInteger(n)&&n>0?n:null;result.cards[i].category=n===20?'Pallet':'แบก';$('method'+i).textContent=result.cards[i].category}));
 }else{
  const editable=['registration','container','seal','tare','gross','net','weightCard'];
  $('previewRows').innerHTML=result.cards.map((v,i)=>'<tr><td><span class="pill">'+v.kind+'</span></td><td><b>'+esc(v.booking)+'</b></td><td>'+v.carNo+'</td>'+editable.map(key=>'<td><input data-i="'+i+'" data-key="'+key+'" value="'+esc(v[key]??'')+'" placeholder="'+(key==='weightCard'?'ยังไม่มี':'—')+'"></td>').join('')+'</tr>').join('');
  $('previewRows').querySelectorAll('input').forEach(el=>el.addEventListener('change',()=>{const v=result.cards[Number(el.dataset.i)],key=el.dataset.key;v[key]=['tare','gross','net'].includes(key)?(Number(el.value)||null):el.value.trim()}));
 }
 const warnings=[...result.warnings];
 $('warningsBox').classList.toggle('hidden',!warnings.length);
 $('warningsBox').innerHTML=warnings.length?'<b>รายการที่ต้องระวัง ('+warnings.length+')</b><div style="max-height:160px;overflow:auto">'+warnings.slice(0,24).map(w=>'<div>• '+esc(w)+'</div>').join('')+(warnings.length>24?'<div>… และอีก '+(warnings.length-24)+' รายการ</div>':'')+'</div>':'';
 $('generateHelp').textContent=result.mode==='prep'?'ไฟล์ผลลัพธ์จะเป็น .xlsx ใหม่สำหรับดาวน์โหลด โดยไม่เปลี่ยนไฟล์ต้นทาง':'ระบบจะสร้าง Excel จากไฟล์ที่อัปโหลด ไม่เขียนทับไฟล์ต้นฉบับในเครื่อง';
 stat(files.length,result.bookings,result.units,warnings.length);
 status('วิเคราะห์สำเร็จ '+result.units+' '+(result.mode==='prep'?'ตู้':'คัน')+' — ตรวจรายการก่อนกดสร้างไฟล์','good');
 $('previewCard').scrollIntoView({behavior:'smooth',block:'start'});
}
async function analyze(){
 if(busy)return;busy=true;$('analyzeBtn').disabled=true;resetPreview();
 try{
  const iso=validDate();
  const plan=await readPlans(iso);
  const result=mode==='prep'?analyzePrep(plan):await analyzeDaily(iso,plan);
  renderPreview(result);log('วิเคราะห์สำเร็จ: '+result.bookings+' Booking / '+result.units+' คัน / แจ้งเตือน '+result.warnings.length);
 }catch(e){status(e.message,'error');log('ERROR '+e.message);console.error(e)}
 finally{busy=false;$('analyzeBtn').disabled=false}
}
function getETD(job){
 const raw=job.etd||'';
 const month=(raw.match(/[A-Za-z]{3,}/)||[])[0];
 const day=(raw.match(/\d{1,2}/)||[])[0];
 return raw;
}
async function prepFile(kind,jobs,iso,template){
 const x=await C().ExcelPatch.load(template),pages=createPages(jobs),needed=pages.length;
 if(x.sheets.length>needed)throw Error('แม่แบบ '+kind+' มี '+x.sheets.length+' ชีต แต่ต้องใช้เพียง '+needed+' ชีต กรุณาใช้แม่แบบต้นฉบับที่มีชีตน้อยกว่าหรือเท่าจำนวนงาน เพื่อป้องกันหน้าเก่าค้าง');
 while(x.sheets.length<needed)await x.cloneSheet(x.sheets[x.sheets.length-1]);
 const names=new Map();
 for(let p=0;p<pages.length;p++){
  const sh=x.sheets[p],items=pages[p];
  for(let side=0;side<2;side++){
   const it=items[side],left=side===0;
   const seq=kind==='STP'?(left?8:9):(left?11:12);
   const date=kind==='STP'?(left?6:7):(left?8:9);
   const book=kind==='STP'?(left?10:11):(left?7:10);
   const contract=kind==='STP'?(left?12:13):(left?13:20);
   const etd=kind==='STP'?(left?14:15):(left?17:21);
   const close=kind==='STP'?(left?16:17):(left?19:22);
   const destination=kind==='STP'?(left?18:19):(left?18:23);
   const agent=kind==='STP'?(left?20:21):(left?15:24);
   const place=kind==='STP'?(left?22:23):(left?16:25);
   const detail=kind==='STP'?(left?28:29):(left?2:26);
   const quantity=kind==='STP'?(left?25:27):0;
   const product=kind==='PTR'?(left?14:27):0;
   const company=left?'G10':'S10',remark=left?'D22':'P22';
   const area=left?['A','E','H','J','K']:['M','Q','T','V','W'];
   for(let r=17;r<=21;r++)for(const col of area)await x.clearCell(sh,col+r);
   await x.clearCell(sh,company);
   await x.setCell(sh,remark,it?.job.product||'');
   if(!it){
    for(const [idx,fontsize] of [[seq,11],[date,10],[book,10],[contract,9],[etd,9],[close,9],[destination,9],[agent,9],[place,9],[detail,8],[quantity,9],[product,9]])if(idx)await x.setShape(sh,idx,'',fontsize);
    continue;
   }
   const job=it.job,n=it.no;
   const bags=job._cardBagValues?.[n-1]??job.bags;
   if(!bags)throw Error('ไม่พบจำนวนถุง '+job.booking+' คัน '+n);
   await x.setShape(sh,date,dt(iso),11);
   await x.setShape(sh,seq,n+'/'+job.fcl+'\n'+(bags===20?'Pallet':'แบก'),11);
   await x.setShape(sh,book,job.booking,12);
   await x.setShape(sh,contract,job.contract,job.contract.length>34?7.5:9.5);
   await x.setShape(sh,etd,getETD(job),10);
   await x.setShape(sh,close,job.closing+' ('+job.closingTime+')',9);
   await x.setShape(sh,destination,job.destination,9);
   await x.setShape(sh,agent,job.agent,9);
   await x.setShape(sh,place,job.place,9);
   if(kind==='STP'){
    await x.setShape(sh,detail,job.packaging,8.2);
    await x.setShape(sh,quantity,'Qty per Fcl : '+bags+' Bags',9);
   }else{
    const rect='Product Name:  \nBatch Number :  \nCode No :  \nMFG :  \nEXP:  \nQuantity :   '+bags+'  Bag';
    await x.setShape(sh,detail,rect,8.4);
    await x.setShape(sh,product,job.product,job.product.length>19?7:9);
   }
  }
  const job=items[0].job;
  const suffixKey=job.booking+'|'+job.product;
  names.set(suffixKey,(names.get(suffixKey)||0)+1);
  await x.rename(sh,sheetTitle(job,names.get(suffixKey)));
 }
 // Final check of all user-facing generated textbox details before saving.
 for(let p=0;p<pages.length;p++)for(let side=0;side<2;side++){
  const it=pages[p][side];if(!it)continue;
  const sh=x.sheets[p],idx=kind==='STP'?(side===0?28:29):(side===0?14:27);
  const text=await x.shapeText(sh,idx);
  const expected=kind==='STP'?it.job.packaging:it.job.product;
  if(text.replace(/\s+/g,'')!==expected.replace(/\s+/g,''))throw Error('ตรวจ Text Box ไม่ผ่าน: '+it.job.booking);
 }
 log('สร้าง '+kind+' '+pages.length+' ชีต / '+pages.flat().filter(v=>v?.no).length+' ตู้');
 return x.save('AliceOffice_ใบแจ้งโหลด_'+kind+'_'+stamp(iso)+'.xlsx');
}
async function createPrep(){
 const iso=validDate(),jobs=prepared.jobs;
 for(const j of jobs){j._cardBagValues=prepared.cards.filter(c=>c.job===j).map(c=>c.bags?Number(c.bags):null)}
 const out=[];
 for(const kind of ['STP','PTR']){
  const group=jobs.filter(j=>j.kind===kind);
  if(!group.length)continue;
  const template=sole('form'+kind);
  if(!template){log('ข้าม '+kind+' ไม่มีแม่แบบ');continue}
  if(group.some(j=>j._cardBagValues.some(v=>!v)))throw Error('กรุณาระบุจำนวนถุงให้ครบทุกคันก่อนสร้าง '+kind);
  if(kind==='STP'&&group.some(j=>!j.packaging))throw Error('รายการ STP ไม่มี Packaging and Batch no. กรุณาตรวจ CSV');
  out.push(await prepFile(kind,group,iso,template));
 }
 if(!out.length)throw Error('ยังไม่มีแม่แบบ Excel STP/PTR ที่ใช้สร้างไฟล์ได้');
 return out;
}
function findExportSheet(wb,iso){
 const [yyyy,mm]=iso.split('-');
 const mon=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][Number(mm)-1];
 const matched=wb.SheetNames.filter(s=>new RegExp('Exporting\\s+Report\\s+'+mon+'\\w*\\s+'+yyyy,'i').test(s));
 if(matched.length===1)return matched[0];
 if(matched.length>1)throw Error('พบชีต Export เดือน '+mon+' '+yyyy+' ซ้ำ กรุณาตรวจไฟล์');
 const fallback=wb.SheetNames.filter(s=>/exporting report/i.test(s));
 if(fallback.length===1)return fallback[0];
 throw Error('ไม่พบชีต Export เดือน '+mon+' '+yyyy+' (ไฟล์มี '+wb.SheetNames.length+' ชีต) กรุณาตรวจไฟล์ให้ตรงวัน');
}
function exportMap(kind){
 return kind==='STP'?{date:'C',contract:'E',booking:'F',agent:'G',etd:'H',receipt:'I',place:'J',product:'K',batch:'L',code:'M',bags:'N',kg:'O',tare:'P',gross:'Q',net:'R',carrier:'T',car:'W',registration:'X',driver:'Y',container:'Z',seal:'AA',weightCard:'AB'}:
 {date:'A',contract:'C',booking:'D',agent:'E',etd:'F',receipt:'G',place:'H',product:'I',batch:'J',code:'K',bags:'L',kg:'M',tare:'N',gross:'O',net:'P',carrier:'R',car:'U',registration:'V',driver:'W',container:'X',seal:'Y',weightCard:'Z'};
}
function colIndex(str){let n=0;for(const c of str)n=n*26+c.charCodeAt(0)-64;return n-1}
async function preserveNewRowStyle(patcher,sheet,source,target){
 const d=await patcher.worksheet(sheet),data=C().find(d.documentElement,'sheetData');
 const rows=C().l(data,'row');
 if(rows.some(r=>Number(r.getAttribute('r'))===target))return;
 const from=rows.find(r=>Number(r.getAttribute('r'))===source);
 if(!from)return;
 const clone=from.cloneNode(true);clone.setAttribute('r',String(target));
 for(const cell of C().l(clone,'c')){
  cell.setAttribute('r',cell.getAttribute('r').replace(/\d+$/,String(target)));
  C().l(cell,'f').concat(C().l(cell,'v'),C().l(cell,'is')).forEach(n=>n.remove());
  cell.removeAttribute('t');
 }
 const after=rows.find(r=>Number(r.getAttribute('r'))>target);data.insertBefore(clone,after||null);
}
async function createDaily(){
 const iso=validDate(),outputs=[];
 const required=['booking','carNo'];
 for(const kind of ['STP','PTR']){
  const source=prepared.filesByKind[kind];if(!source)continue;
  const list=source.vehicles;
  for(const v of list)if(!v.booking||!v.carNo)throw Error('Loading '+kind+' มี Booking หรือเลขคันไม่ครบ');
  const loadPatch=await C().ExcelPatch.load(source.load),expPatch=await C().ExcelPatch.load(source.exp);
  const loadSheet=loadPatch.sheetByName(source.sheetName);
  if(!loadSheet)throw Error('ไม่พบชีต Loading '+source.sheetName);
  const exportOriginal=await C().parseXlsx(source.exp),exportSheetName=findExportSheet(exportOriginal,iso);
  const expSheet=expPatch.sheetByName(exportSheetName);
  if(!expSheet)throw Error('ไม่พบชีต Export '+exportSheetName);
  const rows=C().sheetRows(exportOriginal,exportSheetName),map=exportMap(kind);
  const rowIndex=new Map(),regIndex=new Map(),usedTargets=new Set();
  let last=6,maxDayCar=0;
  const serialDate=C().excelDate(iso);
  const key=(book,code)=>C().clean(book)+'|'+C().clean(code).replace(/\s+/g,'').toUpperCase();
  for(let r=6;r<=rows.length;r++){
   const cells=rows[r-1]||[];
   const booking=C().clean(cells[colIndex(map.booking)]);
   const cont=C().clean(cells[colIndex(map.container)]);
   const reg=C().clean(cells[colIndex(map.registration)]);
   const date=Number(cells[colIndex(map.date)])||0;
   if(booking)last=Math.max(last,r);
   if(booking&&date===serialDate){
    if(cont)rowIndex.set(key(booking,cont),r);
    if(reg)regIndex.set(key(booking,reg),r);
    maxDayCar=Math.max(maxDayCar,Number(cells[colIndex(map.car)])||0);
   }
  }
  let next=last+1,matched=0,appended=0;
  for(const v of list){
   if(!v.container&&$('strictMode').value==='strict')throw Error('ยังไม่มีเลขตู้ '+v.booking+' คัน '+v.carNo+' กรุณากรอกก่อนสร้าง');
   const identity=key(v.booking,v.container);
   let target=v.container?rowIndex.get(identity):null;
   if(!target&&v.registration)target=regIndex.get(key(v.booking,v.registration));
   if(target){
    const oldCont=C().clean((rows[target-1]||[])[colIndex(map.container)]);
    if(oldCont&&v.container&&key(v.booking,oldCont)!==identity)throw Error('เลขตู้ขัดแย้งใน Export: '+v.booking+' คัน '+v.carNo+' ('+oldCont+' / '+v.container+')');
    if(usedTargets.has(target))throw Error('มีรถสองคันชี้ไปที่แถว Export เดียวกัน: '+v.booking+' '+v.container);
    matched++;
   }else{
    target=next++;appended++;await preserveNewRowStyle(expPatch,expSheet,Math.max(6,last),target);
   }
   usedTargets.add(target);
   const exportCarNumber=target<=rows.length?Number((rows[target-1]||[])[colIndex(map.car)])||++maxDayCar:++maxDayCar;
   const writableLoad={I:v.carrier,J:v.registration,K:v.container,L:v.seal,U:v.weightCard,V:v.tare,W:v.gross,X:v.net};
   for(const [col,val] of Object.entries(writableLoad))if(val!==null&&val!==undefined&&val!=='')await loadPatch.setCell(loadSheet,col+v.row,val);
   const obj={date:C().excelDate(iso),contract:v.contract,booking:v.booking,agent:v.agent,etd:(v.etd&&v.destination)?v.etd+'/'+v.destination:null,receipt:'',place:v.returnPlace,product:v.product,batch:v.batch,code:v.code,bags:v.bags||null,kg:v.kg||null,tare:v.tare,gross:v.gross,net:v.net,carrier:v.carrier,car:exportCarNumber,registration:v.registration,driver:v.driver,container:v.container,seal:v.seal,weightCard:v.weightCard};
   // Do not overwrite unknown fields with synthetic values.
   for(const [key,col] of Object.entries(map)){
    const value=obj[key];if(value!==null&&value!==undefined&&value!=='')await expPatch.setCell(expSheet,col+target,value);
   }
  }
  outputs.push(await loadPatch.save('AliceOffice_Loading_'+kind+'_'+stamp(iso)+'.xlsx'));
  outputs.push(await expPatch.save('AliceOffice_Export_'+kind+'_'+stamp(iso)+'.xlsx'));
  log('Daily '+kind+': '+list.length+' คัน, จับคู่แถวเดิม '+matched+', เพิ่มแถว '+appended);
 }
 if(!outputs.length)throw Error('ไม่พบคู่ไฟล์ Loading + Export ให้สร้างรายงาน');
 return outputs;
}
async function generate(){
 if(busy||!prepared)return;busy=true;$('generateBtn').disabled=true;status('กำลังสร้างไฟล์ Excel กรุณาอย่าปิดหน้านี้...');
 try{
  const results=prepared.mode==='prep'?await createPrep():await createDaily();
  outputs=results;
  $('downloads').innerHTML='';
  for(const out of results){
   const b=document.createElement('button');b.className='download';b.type='button';
   b.innerHTML='<span style="font-size:24px">📄</span><span><b>'+esc(out.filename)+'</b><small>'+fmt(Math.round(out.blob.size/1024))+' KB · Excel Workbook</small></span><span class="dl">↓</span>';
   b.addEventListener('click',()=>{C().saveBlob(out);log('ดาวน์โหลด '+out.filename)});
   $('downloads').appendChild(b);
  }
  $('resultsCard').classList.remove('hidden');$('resultsCard').scrollIntoView({behavior:'smooth',block:'start'});
  status('สร้างไฟล์สำเร็จ '+results.length+' ไฟล์ ตรวจสอบ Print Preview ก่อนนำไปใช้จริงนะคะ 💕','good');
  log('สร้างไฟล์เสร็จ '+results.length+' ไฟล์');
 }catch(e){status('ยังสร้างไฟล์ไม่ได้: '+e.message,'error');log('ERROR '+e.message);console.error(e)}
 finally{busy=false;$('generateBtn').disabled=false}
}
function init(){
 const today=new Date();const iso=[today.getFullYear(),String(today.getMonth()+1).padStart(2,'0'),String(today.getDate()).padStart(2,'0')].join('-');$('workDate').value=iso;
 $('pickFiles').onclick=()=>$('files').click();$('pickFolder').onclick=()=>$('directory').click();
 $('files').addEventListener('change',e=>{acceptFiles(e.target.files);e.target.value=''});
 $('directory').addEventListener('change',e=>{acceptFiles(e.target.files);e.target.value=''});
 $('dropZone').addEventListener('dragover',e=>{e.preventDefault();$('dropZone').classList.add('drag')});
 $('dropZone').addEventListener('dragleave',()=>$('dropZone').classList.remove('drag'));
 $('dropZone').addEventListener('drop',e=>{e.preventDefault();$('dropZone').classList.remove('drag');acceptFiles(e.dataTransfer.files)});
 $('tabPrep').onclick=()=>showMode('prep');$('tabDaily').onclick=()=>showMode('daily');
 $('analyzeBtn').onclick=analyze;$('generateBtn').onclick=generate;
 $('clearBtn').onclick=()=>{files=[];renderFiles();resetPreview();status('ล้างไฟล์แล้วค่ะ');log('ล้างรายการไฟล์')};
 $('workDate').addEventListener('change',resetPreview);
 renderFiles();log('ระบบพร้อม: เตรียมโหลด / Loading → Export');
 if(!window.JSZip||!window.XLSX){status('ไม่สามารถโหลดไลบรารี Excel ได้ กรุณาตรวจอินเทอร์เน็ตและรีเฟรชหน้า','error');log('ERROR: JSZip/XLSX CDN unavailable')}
}
document.addEventListener('DOMContentLoaded',init);
})();
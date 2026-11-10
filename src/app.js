(function(){
'use strict';
const C=()=>window.AliceCore;
const $=id=>document.getElementById(id);
const roles=[['planSTP','แผน STP 1 (CSV)'],['planPTR','แผน STP 2 / PTR (CSV)'],['formSTP','ใบแจ้งโหลด STP'],['formPTR','ใบแจ้งโหลด PTR'],['loadingSTP','Loading STP'],['loadingPTR','Loading PTR'],['loadingDomestic','Domestic / SSAMS Loading'],['planDomestic','แผนโหลด SSAMS (อ้างอิง)'],['planBatch','แผนแบช / ตำแหน่งจัดเก็บ (Rev)'],['exportSTP','Export STP'],['exportPTR','Export PTR'],['truck','ข้อมูลรถขนส่ง'],['ignore','ไม่ใช้งาน']];
let mode='prep',files=[],counter=0,prepared=null,outputs=[],busy=false;

const DAILY_EXTRA_GROUPS=[
 {title:'เอกสารและค่าขนส่ง',fields:[
  ['invoice','Invoice No.'],['contract','Contract No.'],['agent','Agent'],
  ['etd','ETD / Destination'],['receipt','Place of Receipt'],['returnPlace','Place of Return'],
  ['driver','Driver Name'],['carrier','Truck Company'],['contractorTeam','Contractor Team'],['price','Price (THB)'],
  ['overnight','Overnight Charge']
 ]},
 {title:'เวลาและติดตามงาน',fields:[
  ['arrive','Time Arrive','time'],['depart','Time Out','time'],
  ['delayCause','Root Cause Delay'],['loadStart','Start Loading','time'],
  ['loadEnd','Finish Loading','time'],['duration','Total Time','time','readonly']
 ]},
 {title:'ประเมินเอกสารและขนส่ง',fields:[
  ['documentDelivery','การจัดส่งเอกสาร','number'],['documentAccuracy','ความถูกต้องเอกสาร','number'],
  ['coordination','การประสานงาน','number'],['docTotal','Total Score เอกสาร','number','readonly'],
  ['docPct','Total % เอกสาร','number','readonly'],
  ['punctuality','ความตรงต่อเวลา','number'],['truckCondition','สภาพรถ','number'],
  ['compliance','ปฏิบัติตามกฎ','number'],['driverCourtesy','มารยาทคนขับ','number'],
  ['carrierTotal','Total Score ขนส่ง','number','readonly'],
  ['carrierPct','Total % ขนส่ง','number','readonly'],
  ['averageScore','Average Score','number','readonly'],
  ['assessmentTotal','Total Assessment','number','readonly']
 ]}
];
const DAILY_FIELDS=DAILY_EXTRA_GROUPS.flatMap(group=>group.fields);
const DAILY_DEFAULTS={
 receipt:'-',contractorTeam:'-',price:'-',overnight:'-',delayCause:'-',
 loadStart:'10:00',loadEnd:'10:50',
 documentDelivery:1,documentAccuracy:1,coordination:1,
 punctuality:1,truckCondition:1,compliance:1,driverCourtesy:1
};
const DAILY_RATING_KEYS=new Set(['documentDelivery','documentAccuracy','coordination','punctuality','truckCondition','compliance','driverCourtesy']);
function fixScore(value){return Math.round(Number(value||0)*10000)/10000}
function timeMinutes(value){
 const m=/^(\d{1,2}):(\d{2})$/.exec(String(value||''));if(!m)return null;
 const h=Number(m[1]),minute=Number(m[2]);
 return h<24&&minute<60?h*60+minute:null;
}
function asExcelTime(value){
 const minutes=timeMinutes(value);return minutes===null?null:minutes/1440;
}
function fromExcelTime(value){
 if(value===null||value===undefined||value==='')return '';
 if(typeof value==='number'&&Number.isFinite(value)){
  const minutes=Math.round((((value%1)+1)%1)*1440)%1440;
  return String(Math.floor(minutes/60)).padStart(2,'0')+':'+String(minutes%60).padStart(2,'0');
 }
 if(typeof value==='string'&&/^\d{1,2}:\d{2}/.test(value))return value.slice(0,5);
 return '';
}
function recalcDailyScores(extra){
 const n=key=>Number(extra[key])||0;
 extra.docTotal=fixScore(n('documentDelivery')+n('documentAccuracy')+n('coordination'));
 extra.docPct=fixScore(extra.docTotal/3);
 extra.carrierTotal=fixScore((n('punctuality')+n('truckCondition')+n('compliance')+n('driverCourtesy'))/4);
 extra.carrierPct=extra.carrierTotal;
 extra.averageScore=fixScore((extra.docPct+extra.carrierPct)/2);
 extra.assessmentTotal=extra.averageScore;
 const start=timeMinutes(extra.loadStart),end=timeMinutes(extra.loadEnd);
 extra.duration=start===null||end===null?'':fromExcelTime(((end-start+1440)%1440)/1440);
}
function dailySupplementColumns(kind){
 const cols=kind==='PTR'?
  ['AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS']:
  ['AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU'];
 const keys=['arrive','depart','delayCause','loadStart','loadEnd','duration',
  'documentDelivery','documentAccuracy','coordination','docTotal','docPct',
  'punctuality','truckCondition','compliance','driverCourtesy','carrierTotal','carrierPct',
  'averageScore','assessmentTotal'];
 return Object.fromEntries(keys.map((key,i)=>[key,cols[i]]));
}
function initializeDailyExtra(v){
 const extra={...DAILY_DEFAULTS,
  invoice:v.invoice||'-',contract:v.contract||'-',agent:v.agent||'-',
  etd:(v.etd&&v.destination)?v.etd+'/'+v.destination:(v.etd||'-'),
  returnPlace:v.returnPlace||'-',driver:v.driver||'-',carrier:v.carrier||'-'};
 recalcDailyScores(extra);return extra;
}
function matchExistingReportRow(rows,kind,v,dateSerial,claimed){
 const map=exportMap(kind),normal=x=>C().clean(x).toUpperCase().replace(/[^A-Z0-9]/g,'');
 const list=[];
 for(let r=7;r<=rows.length;r++){
  if(claimed.has(r))continue;
  const values=rows[r-1]||[];
  if(Number(values[colIndex(map.date)])!==dateSerial)continue;
  if(C().clean(values[colIndex(map.booking)])!==v.booking)continue;
  list.push({row:r,values});
 }
 const byUnique=fn=>{const m=list.filter(fn);return m.length===1?m[0]:null};
 let found=v.container&&normal(v.container)!=='-'?
  byUnique(x=>normal(x.values[colIndex(map.container)])===normal(v.container)):null;
 if(!found&&v.weightCard)found=byUnique(x=>normal(x.values[colIndex(map.weightCard)])===normal(v.weightCard));
 if(!found&&v.registration)found=byUnique(x=>normal(x.values[colIndex(map.registration)])===normal(v.registration)&&
  (v.kind!=='DOMESTIC'||normal(x.values[colIndex(map.place)])===normal(v.returnPlace)));
 if(!found&&v.kind==='DOMESTIC'&&!v.registration&&!v.weightCard)
  found=byUnique(x=>normal(x.values[colIndex(map.place)])===normal(v.returnPlace)&&
   normal(x.values[colIndex(map.product)])===normal(v.product)&&Number(x.values[colIndex(map.bags)])===v.bags);
 if(found)claimed.add(found.row);return found;
}
async function populateDailyExtras(iso,vehicles){
 const serial=C().excelDate(iso);
 for(const v of vehicles)v.extra=initializeDailyExtra(v);
 for(const kind of ['STP','PTR']){
  const subset=vehicles.filter(v=>v.exportKind===kind),file=sole('export'+kind);
  if(!subset.length||!file)continue;
  const wb=await C().parseXlsx(file),name=findExportSheet(wb,iso),rows=C().sheetRows(wb,name);
  const map=exportMap(kind),extraMap=dailySupplementColumns(kind),claimed=new Set();
  for(const v of subset){
   const found=matchExistingReportRow(rows,kind,v,serial,claimed);
   if(!found)continue;
   v.existingExportRow=found.row;
   for(const key of ['invoice','contract','agent','etd','receipt','driver','carrier']){
    const cell=found.values[colIndex(map[key])];
    if(cell!==null&&cell!==undefined&&String(cell).trim()!=='')v.extra[key]=String(cell);
   }
   const place=found.values[colIndex(map.place)];
   if(place!==null&&place!==undefined&&String(place).trim()!=='')v.extra.returnPlace=String(place);
   for(const key of ['contractorTeam','price','overnight']){
    const col=key==='contractorTeam'?(kind==='PTR'?'Q':'S'):
     key==='price'?(kind==='PTR'?'S':'U'):(kind==='PTR'?'T':'V');
    const value=found.values[colIndex(col)];
    if(value!==null&&value!==undefined&&String(value).trim()!=='')v.extra[key]=value;
   }
   for(const [key,col] of Object.entries(extraMap)){
    const value=found.values[colIndex(col)];
    if(value===null||value===undefined||String(value).trim()==='')continue;
    v.extra[key]=['arrive','depart','loadStart','loadEnd','duration'].includes(key)?fromExcelTime(value):value;
   }
   // Recompute dependent display values only after explicit changes; existing
   // report values are preserved until the user edits a score or time.
  }
 }
}

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
function resetPreview(){prepared=null;outputs=[];$('batchPlanCard')?.classList.add('hidden');$('previewCard').classList.add('hidden');$('resultsCard').classList.add('hidden');$('statBookings').textContent='—';$('statUnits').textContent='—';$('statWarnings').textContent='—'}
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
 $('uploadDesc').textContent=mode==='prep'?'ใช้ CSV STP1/STP2, แผนแบช Rev และแบบใบแจ้งโหลด STP/PTR (.xlsx)':'ใช้ Loading STP/PTR, Domestic/SSAMS, Export STP/PTR พร้อมไฟล์รถขนส่ง (ถ้ามี) — Loading ทุกประเภทที่ไม่ใช่ PTR ลง Export STP';
 $('rulesTitle').textContent=mode==='prep'?'กติกาเตรียมโหลด':'กติกา Daily Report';
 $('rulesPrep').classList.toggle('hidden',mode!=='prep');$('rulesDaily').classList.toggle('hidden',mode!=='daily');
 $('statUnitLabel').textContent=mode==='prep'?'จำนวนตู้':'จำนวนคัน';
 $('rulesNotice').innerHTML=mode==='prep'?'แผน Rev จะแสดงตำแหน่งคลัง Batch, Code และการจัดสรรแยกคัน โดยใช้สำหรับตรวจสอบ • แม่แบบ STP/PTR ควรเป็น <b>ไฟล์ Excel เดิมที่มี Text Box</b> เพื่อรักษารูปแบบต้นฉบับ รวมทั้งหัวข้อ Gate No. ของ PTR':'ระบบจะนำข้อมูล <b>Loading ทุกส่วน</b> ที่มีช่องตรงกันลง Export: PTR → Export PTR; ส่วน STP และ Domestic/SSAMS → Export STP โดยไม่คัดลอกข้อมูลจากวันก่อน';
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

function batchAllocations(raw){
 const out=[],regex=/(\d{1,3})\s*\(\s*([\d,]+(?:\.\d+)?)\s*\)/g;
 let m;while((m=regex.exec(String(raw||'')))){
  const carNo=Number(m[1]),qty=Number(m[2].replace(/,/g,''));
  if(carNo>=1&&carNo<=200&&qty>0)out.push({carNo,qty});
 }
 return out;
}
async function readBatchPlan(){
 const file=sole('planBatch');if(!file)return null;
 const wb=await C().parseXlsx(file),sections=[],warnings=[];
 let allocationRows=0,inventoryRows=0;
 for(const name of wb.SheetNames){
  const data=C().sheetRows(wb,name);
  const head=data.findIndex(r=>/^location$/i.test(C().clean(r[0]))&&/item\s*number/i.test(C().clean(r[1]))&&/qty\s*allocate/i.test(C().clean(r[3]))&&/^batch$/i.test(C().clean(r[6])));
  if(head===-1){warnings.push('อ่านหัวตารางแบชของ '+name+' ไม่ได้');continue}
  const isInventory=/^sheet1$/i.test(name),rows=[],notes=[],noted=new Set();
  for(let i=head+1;i<data.length;i++){
   const r=data[i],location=C().clean(r[0]),product=C().clean(r[1]),stock=Number(r[2]),
    allocationText=C().clean(r[3]),uom=C().clean(r[4]),status=C().clean(r[5]),lot=C().clean(r[6]);
   if(product&&lot&&Number.isFinite(stock)&&stock>=0){
    const sep=lot.indexOf('_'),batch=sep>=0?lot.slice(0,sep):lot,code=sep>=0?lot.slice(sep+1):'';
    const allocation=isInventory?[]:batchAllocations(allocationText);
    const allocated=allocation.reduce((sum,x)=>sum+x.qty,0);
    rows.push({sheet:name,row:i+1,location,product,stock,allocationText,uom,status,batch,code,allocation,allocated});
    if(isInventory)inventoryRows++;else if(allocated)allocationRows++;
    if(allocated>stock)warnings.push(name+' แถว '+(i+1)+' จัดสรร '+allocated+' เกินจำนวน '+stock+' '+uom);
    if(allocated&&status&&!/^available$/i.test(status))warnings.push(name+' แถว '+(i+1)+' จัดสรรจากสถานะ '+status);
   }else{
    for(const v of [r[0],r[1],r[2],r[3]]){
     const note=C().clean(v);if(note&&!noted.has(note)&&/โหลด|ตัดช็อต|รอโหลด|แจ้ง/i.test(note)){
      notes.push(note);noted.add(note);
     }
    }
   }
  }
  sections.push({name,isInventory,rows,notes,allocated:rows.reduce((sum,r)=>sum+r.allocated,0),
    cars:[...new Set(rows.flatMap(r=>r.allocation.map(a=>a.carNo)))].sort((a,b)=>a-b)});
 }
 if(!sections.length)throw Error('ไฟล์แผนแบชไม่มีหัวตาราง Location / Qty Allocate / Batch ที่ถูกต้อง');
 return {filename:file.name,sections,warnings,allocationRows,inventoryRows};
}
function compareBatchPlan(plan,vehicles){
 const notes=[],checks=[],key=x=>C().clean(x).replace(/\s/g,'').toUpperCase();
 const planned=plan.sections.filter(s=>!s.isInventory).flatMap(s=>s.rows);
 for(const v of vehicles){
  let used=0,totalPlan=0,totalLoaded=0;
  for(const item of v.items||[]){
   const batch=key(item.batch),code=key(item.code);
   const matches=planned.filter(r=>
    code&&key(r.code)===code&&(!batch||key(r.batch)===batch));
   if(!matches.length)continue;
   if(new Set(matches.map(x=>x.sheet)).size!==1)continue;
   const plannedQty=matches.reduce((sum,r)=>sum+r.allocation.filter(x=>x.carNo===v.carNo).reduce((n,x)=>n+x.qty,0),0);
   if(!plannedQty)continue;
   totalPlan+=plannedQty;totalLoaded+=Number(item.bags)||0;used++;
  }
  if(used){
   const ok=totalLoaded===totalPlan;
   checks.push({kind:v.kind,booking:v.booking,carNo:v.carNo,planned:totalPlan,loaded:totalLoaded,ok});
   if(!ok)notes.push(v.kind+' Booking '+v.booking+' คัน '+v.carNo+
    ' แบชตรงกันแต่จำนวนในแผน '+totalPlan+' / Loading '+totalLoaded+' — ตรวจสอบก่อนใช้');
  }
 }
 return {checks,notes};
}
function renderBatchPlan(plan,checks=[]){
 const box=$('batchPlanCard');if(!box)return;
 box.classList.toggle('hidden',!plan);if(!plan)return;
 const sections=plan.sections,sel=$('batchSheetFilter'),car=$('batchCarFilter'),search=$('batchSearch');
 $('batchSummary').textContent=plan.filename+' • '+plan.allocationRows+' แถวจัดสรร • '+
  plan.inventoryRows+' แถวสต็อกอ้างอิง • '+sections.filter(s=>!s.isInventory).length+' กลุ่มสินค้า';
 const selected=sel.value;
 sel.innerHTML=sections.map(x=>'<option value="'+esc(x.name)+'">'+esc(x.name)+
  (x.isInventory?' — สต็อกอ้างอิง':' — '+fmt(x.allocated)+' หน่วยจัดสรร')+'</option>').join('');
 sel.value=sections.some(x=>x.name===selected)?selected:(sections.find(x=>!x.isInventory)||sections[0]).name;
 const draw=(refreshCars=false)=>{
  const section=sections.find(x=>x.name===sel.value),carChoice=refreshCars?'':car.value;
  if(refreshCars||!car.options.length){
   car.innerHTML='<option value="">ทุกคัน/ตู้</option>'+section.cars.map(n=>'<option value="'+n+'">คัน/ตู้ '+n+'</option>').join('');
   car.value=section.cars.includes(Number(carChoice))?carChoice:'';
  }
  const n=Number(car.value),term=C().clean(search.value).toUpperCase();
  const rows=section.rows.filter(r=>(!n||r.allocation.some(x=>x.carNo===n))&&
   (!term||[r.product,r.batch,r.code,r.status,r.allocationText].join(' ').toUpperCase().includes(term)));
  $('batchRows').innerHTML=rows.slice(0,120).map(r=>
   '<tr><td>'+esc(r.product)+'</td><td>'+esc(r.batch)+'</td>'+
   '<td>'+esc(r.code)+'</td><td>'+fmt(r.stock)+'</td><td>'+esc(r.uom)+'</td>'+
   '<td>'+esc(r.status)+'</td><td><b>'+esc(n?r.allocation.filter(x=>x.carNo===n).map(x=>x.carNo+'('+x.qty+')').join(' '):r.allocationText)+'</b></td></tr>').join('');
  $('batchCount').textContent=fmt(rows.length)+' รายการ'+(rows.length>120?' (แสดง 120 รายการแรก — ค้นหาเพื่อดูรายการอื่น)':'');
  $('batchNotes').textContent=section.notes.length?'หมายเหตุ: '+section.notes.join(' • '):'';
 };
 sel.onchange=()=>draw(true);car.onchange=()=>draw();search.oninput=()=>draw();draw(true);
 const mismatches=checks.filter(x=>!x.ok);
 $('batchChecks').textContent=checks.length?
  'เทียบ Batch + Code + คันกับ Loading ได้ '+checks.length+' คัน • ตรง '+(checks.length-mismatches.length)+
  ' • ต่าง '+mismatches.length+' (ไม่แก้ข้อมูล Loading อัตโนมัติ)':
  'ยังจับคู่แผนแบชกับ Loading ไม่ได้ — ดูแผนเป็นข้อมูลอ้างอิง ไม่ถือว่าสินค้าโหลดแล้ว';
 const alerts=[...plan.warnings,...mismatches.map(x=>x.kind+' '+x.booking+' คัน '+x.carNo+
  ' แผน '+x.planned+' / Loading '+x.loaded)];
 $('batchAlerts').classList.toggle('hidden',!alerts.length);
 $('batchAlerts').textContent=alerts.slice(0,15).join(' • ')+(alerts.length>15?' และอีก '+(alerts.length-15)+' รายการ':'');
}

function missingPrepTemplates(plans){
 return ['STP','PTR'].filter(kind=>(plans[kind]||[]).length&&!get('form'+kind).length);
}
async function resolvePrepTemplates(plans){
 const needed=missingPrepTemplates(plans);
 if(!needed.length||!window.AliceCloud?.isAuthenticated?.())return;
 try{await window.AliceCloud.ensureTemplates(needed)}
 catch(e){log('CLOUD ไม่สามารถโหลดแม่แบบอัตโนมัติ: '+e.message)}
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
 if(!jobs.length&&!get('planBatch').length)throw Error('ไม่พบแผนงานตรงวันที่และเงื่อนไข KNS — กรุณาตรวจวันที่หรือประเภท CSV');
 for(const kind of ['STP','PTR']){
  if(plans[kind].length&&!get(kind==='STP'?'formSTP':'formPTR').length)warnings.push('ยังไม่มีไฟล์แม่แบบใบแจ้งโหลด '+kind+' — จะยังสร้างไฟล์ฝั่งนี้ไม่ได้');
  for(const j of plans[kind]){
   if(!j.booking)warnings.push(kind+' มีแถวที่ Booking ว่าง');
   if(!j.packaging&&kind==='STP')warnings.push(j.booking+' ไม่มี Packaging and Batch no.');
   if(!j.bags)warnings.push(j.booking+' '+j.product+' ไม่ทราบจำนวนถุงต่อตู้ ต้องระบุก่อนสร้าง');
  }
 }
 const cards=jobs.flatMap(j=>Array.from({length:j.fcl},(_,i)=>({kind:j.kind,job:j,no:i+1,bags:j.bags||'',category:j.bags===20?'Pallet':'แบก'})));
 // Keep warnings unrelated to bag counts; missing quantities are recalculated
 // as the user types, rather than remaining frozen at the initial CSV reading.
 const baseWarnings=warnings.filter(w=>!w.includes('ไม่ทราบจำนวนถุงต่อตู้'));
 return {mode:'prep',jobs,plans,cards,warnings,baseWarnings,planOnly:!cards.length,bookings:new Set(jobs.map(j=>j.kind+':'+j.booking)).size,units:cards.length};
}
function updatePrepInputs(result,{fromDOM=true}={}){
 if(!result||result.mode!=='prep')return [];
 const inputs=[...$('previewRows').querySelectorAll('.bagInput')];
 if(fromDOM)for(const el of inputs){
  const i=Number(el.dataset.i),card=result.cards[i];
  if(!card)continue;
  const raw=el.value.trim(),n=Number(raw);
  const valid=raw!==''&&Number.isSafeInteger(n)&&n>0;
  card.bags=valid?n:null;
  card.category=valid&&n===20?'Pallet':'แบก';
  el.setAttribute('aria-invalid',String(!valid));
 }
 const missing=[];
 const perJob=new Map();
 for(let i=0;i<result.cards.length;i++){
  const c=result.cards[i],valid=Number.isSafeInteger(c.bags)&&c.bags>0;
  const method=$('method'+i),row=inputs[i]?.closest('tr'),state=row?.querySelector('.prep-review-state');
  if(method){
   method.textContent=c.category;
   method.className='pill '+(!valid?'pending':c.bags===20?'ok':'');
  }
  if(state){
   state.textContent=valid?'พร้อม':'ตรวจจำนวน';
   state.className='pill '+(valid?'ok':'pending')+' prep-review-state';
  }
  if(!valid){
   missing.push(i);
   if(!perJob.has(c.job))perJob.set(c.job,[]);
   perJob.get(c.job).push(c.no);
  }
 }
 result.warnings=[...result.baseWarnings,...[...perJob].map(([job,carNos])=>
  job.booking+' '+job.product+' ต้องระบุจำนวนถุงคันที่ '+carNos.join(', '))];
 showPreviewWarnings(result);
 return missing;
}
function showPreviewWarnings(result){
 const warnings=[...result.warnings],box=$('warningsBox');
 box.classList.toggle('hidden',!warnings.length);
 box.innerHTML=warnings.length?'<b>รายการที่ต้องระวัง ('+warnings.length+')</b><div style="max-height:160px;overflow:auto">'+
  warnings.slice(0,24).map(w=>'<div>• '+esc(w)+'</div>').join('')+
  (warnings.length>24?'<div>… และอีก '+(warnings.length-24)+' รายการ</div>':'')+'</div>':'';
 stat(files.length,result.bookings,result.units,warnings.length);
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
function extractLoading(wb,kind,iso,plan,manifests,{domestic=false}={}){
 const sheetName=chooseLoadingSheet(wb,iso),rows=C().sheetRows(wb,sheetName);
 const vehicles=[],sourceRole=domestic?'loadingDomestic':'loading'+kind;
 let context={contract:'',booking:domestic?'-':'',agent:'',etd:'',destination:'',returnPlace:'',invoice:''},current=null,pageVehicles=[];
 const planMap=new Map(plan.map(j=>[j.booking,j]));
 function validBooking(book){return /^[A-Z0-9][A-Z0-9/-]{5,}$/i.test(book)&&/\d/.test(book)&&!/^(quantity|booking|date)/i.test(book)}
 function pushItem(row,carLine=false){
  const product=C().clean(row[12]||(carLine?row[1]:'')),batch=C().clean(row[14]),code=C().clean(row[15]);
  const packing=C().clean(row[13]),mfg=C().clean(row[16]),expiry=C().clean(row[17]);
  const contract=C().clean(row[0]);
  const bags=Number(row[18])||0,kg=Number(row[19])||0;
  if(product||batch||code||bags||kg)current.items.push({product,batch,code,packing,mfg,expiry,bags,kg,contract:/^S?PTR?\d{4}|^TSP\d{4}/i.test(contract)?contract:''});
 }
 for(let k=0;k<rows.length;k++){
  const row=rows[k],no=Number(row[7]),book=C().clean(row[3]);
  const bookingPresent=validBooking(book),header=C().clean(row[7]).toLowerCase()==='no.';
  if(header){
   current=null;pageVehicles=[];
   context={contract:'',booking:domestic?'-':'',agent:'',etd:'',destination:'',returnPlace:'',invoice:''};
   continue;
  }
  // Invoice / delivery labels can share the same row as a vehicle's cargo.
  // Apply these document values to every vehicle on the current printed page.
  const label=C().clean(row[0]).toLowerCase();
  if(label.includes('invoice no')){
   context.invoice=C().clean(row[1]);
   for(const vehicle of pageVehicles)vehicle.invoice=context.invoice;
  }
  if(label.includes('place of delivery')){
   context.returnPlace=C().clean(row[1]);
   for(const vehicle of pageVehicles)vehicle.returnPlace=context.returnPlace;
  }
  if(bookingPresent){
   context={...context,booking:book,contract:C().clean(row[0])||context.contract,
    agent:C().clean(row[4])||context.agent,etd:C().clean(row[5])||context.etd,destination:''};
  }
  const isCar=Number.isInteger(no)&&no>=1&&no<=200&&C().clean(row[7])!=='';
  if(isCar){
   if(!context.booking)continue;
   const registration=C().clean(row[9]),weightCard=C().clean(row[20]);
   // Some loading sheets repeat a car number on each product line rather than merging the cell.
   const sameCar=current&&current.carNo===no&&current.registration===registration&&current.weightCard===weightCard&&
    (registration||weightCard)&&current.sourceRole===sourceRole;
   if(sameCar){pushItem(row,true);continue}
   const planJob=planMap.get(context.booking);
   current={kind:domestic?'DOMESTIC':kind,exportKind:kind,sourceRole,sheetName,row:k+1,carNo:no,
    booking:context.booking,contract:context.contract||planJob?.contract||'',invoice:context.invoice,
    agent:context.agent||planJob?.agent||'',etd:planJob?.etd||context.etd,
    destination:planJob?.destination||context.destination,returnPlace:context.returnPlace||planJob?.place||'',
    receipt:'',carrier:C().clean(row[8]),registration,container:C().clean(row[10]),
    seal:C().clean(row[11]),weightCard,tare:Number(row[21])||null,gross:Number(row[22])||null,
    net:Number(row[23])||null,driver:'',items:[],sourceName:sheetName,hasManifest:false};
   vehicles.push(current);pageVehicles.push(current);pushItem(row,true);
  }else if(current){
   if(bookingPresent&&book!==current.booking){current=null;continue}
   const f=C().clean(row[5]);
   if(!domestic&&/^[A-Z][A-Z -]{2,32}$/.test(f)&&!/ETD|FCL|CONTRACT/i.test(f)){
    context.destination=f;current.destination ||= f;
   }
   // Totals at the bottom contain bag/kg numbers but no product/batch/code.
   // Skip those without losing cargo on rows also labeled invoice/delivery.
   if(C().clean(row[12])||C().clean(row[14])||C().clean(row[15])||C().clean(row[13]))pushItem(row);
  }
 }
 for(const v of vehicles){
  const truck=manifests.find(m=>m.booking===v.booking&&m.carNo===v.carNo&&(m.container===v.container||!v.container))||
              manifests.find(m=>m.booking===v.booking&&m.carNo===v.carNo);
  if(truck){
   v.hasManifest=true;v.registration ||= truck.registration;v.container ||= truck.container;
   v.seal ||= truck.seal;v.tare ||= truck.tare;v.carrier ||= truck.carrier;v.driver ||= truck.driver;
  }
  const distinct=values=>Array.from(new Set(values.map(C().clean).filter(Boolean)));
  v.contract=distinct([v.contract,...v.items.map(i=>i.contract)]).join(' / ');
  v.product=distinct(v.items.map(i=>i.product)).join(' / ');
  v.batch=distinct(v.items.map(i=>i.batch)).join(' / ');
  v.code=v.items.filter(i=>i.code).map(i=>i.code+(i.bags?'('+i.bags+'Bag)':'')).join(' / ');
  v.packing=distinct(v.items.map(i=>i.packing)).join(' / ');
  v.mfg=distinct(v.items.map(i=>i.mfg)).join(' / ');
  v.expiry=distinct(v.items.map(i=>i.expiry)).join(' / ');
  v.bags=v.items.reduce((a,i)=>a+(i.bags||0),0);
  v.kg=v.items.reduce((a,i)=>a+(i.kg||0),0);
 }
 return {sheetName,vehicles};
}

async function analyzeDaily(iso,plans){
 const filesByKind={},all=[],warnings=[],manifests=[];
 for(const en of get('truck'))try{
  const found=await readManifest(en.file);manifests.push(...found);
  log('ไฟล์ขนส่ง '+en.file.name+' พบ '+found.length+' คัน');
 }catch(e){warnings.push('ไฟล์ขนส่ง '+en.file.name+': '+e.message)}
 const sources=[
  {role:'loadingSTP',kind:'STP',exportRole:'exportSTP'},
  {role:'loadingDomestic',kind:'STP',exportRole:'exportSTP',domestic:true},
  {role:'loadingPTR',kind:'PTR',exportRole:'exportPTR'}
 ];
 for(const source of sources){
  const load=sole(source.role);if(!load)continue;
  const exp=sole(source.exportRole);
  if(!exp){warnings.push('มี '+source.role+' แต่ขาด '+source.exportRole);continue}
  const wb=await C().parseXlsx(load),group=source.domestic?'DOMESTIC':source.kind;
  const parsed=extractLoading(wb,source.kind,iso,source.domestic?[]:plans[source.kind],manifests,{domestic:!!source.domestic});
  // Loading is authoritative for DAILY. A CSV plan is only reference data;
  // it must never hide a genuinely loaded vehicle from the export.
  const vehicles=parsed.vehicles;
  all.push(...vehicles);
  filesByKind[group]={load,exp,sheetName:parsed.sheetName,vehicles};
  log(group+' พบ '+vehicles.length+' คันจาก Loading ('+parsed.sheetName+') → Export '+source.kind);
 }
 for(const kind of ['STP','PTR']){
  if(get('export'+kind).length&&!get('loading'+kind).length&&!(kind==='STP'&&get('loadingDomestic').length))
   warnings.push('มี Export '+kind+' แต่ยังไม่มี Loading สำหรับวันที่เลือก');
 }
 if(!all.length){
  const planned=[...plans.STP,...plans.PTR],plannedUnits=planned.reduce((n,j)=>n+(Number(j.fcl)||0),0);
  warnings.push('ยังไม่มีรายการรถเข้าจริงใน Loading วันที่เลือก — เก็บงานฉบับร่าง แล้วอัปเดตเมื่อมีรถ');
  if(plannedUnits)warnings.push('แผน STP/PTR มี '+plannedUnits+' ตู้ (ยังไม่ใช่ยอดรถเข้าจริง)');
  log('DAILY รอรถเข้าจริง: แผน '+plannedUnits+' ตู้ / Loading 0 คัน');
  return {mode:'daily',plans,filesByKind,cards:[],warnings,bookings:new Set(planned.map(j=>j.kind+':'+j.booking)).size,units:0,plannedUnits,draftOnly:true};
 }
 const dup=new Set();
 for(const v of all){
  if(!v.booking)warnings.push(v.kind+' รถ '+v.carNo+' ไม่มี Booking');
  if(!v.container&&v.kind!=='DOMESTIC')warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังไม่มีเลข Container');
  if(!v.registration)warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังไม่มีทะเบียนรถ');
  if(!v.weightCard||!v.gross||!v.net)warnings.push(v.kind+' '+v.booking+' คัน '+v.carNo+' ยังมีข้อมูลชั่งไม่ครบ — จะเว้นว่าง');
  const key=v.kind+'|'+v.booking+'|'+v.carNo;
  if(dup.has(key))warnings.push('พบลำดับรถซ้ำ: '+key);
  dup.add(key);
 }
 await populateDailyExtras(iso,all);
 return {mode:'daily',plans,filesByKind,cards:all,warnings,
  bookings:new Set(all.map(x=>x.kind+':'+x.booking)).size,units:all.length};
}


function dailyExtraPanel(v,i){
 const title='<summary>✏️ กรอกข้อมูล Export เพิ่มเติม · '+esc(v.kind)+' คัน '+esc(v.carNo)+' — คลิกเพื่อเปิด/ปิด</summary>';
 const groups=DAILY_EXTRA_GROUPS.map(group=>
  '<section class="daily-extra-group"><h4>'+esc(group.title)+'</h4><div class="daily-extra-grid">'+
  group.fields.map(([key,label,type,mode])=>{
   const val=v.extra?.[key]??'';
   const readonly=mode==='readonly';
   const inputType=type==='time'?'time':type==='number'?'number':'text';
   const attrs=type==='number'?' step="any"'+(DAILY_RATING_KEYS.has(key)?' min="0"':''):'';
   const autoNote=(Object.hasOwn(DAILY_DEFAULTS,key)?' · ค่าเริ่มต้น':'');
   return '<label class="daily-extra-field"><span>'+esc(label)+autoNote+'</span>'+
    '<input class="daily-extra-input" data-i="'+i+'" data-key="'+key+'" type="'+inputType+'"'+attrs+
    (readonly?' readonly tabindex="-1"':'')+' value="'+esc(val)+'" placeholder="'+(type==='time'?'--:--':'-')+'"></label>';
  }).join('')+'</div></section>'
 ).join('');
 return '<tr class="daily-details-row"><td colspan="10"><details class="daily-details">'+title+
  '<p class="daily-extra-help">ค่าเริ่มต้นอ้างอิงจากรายงานเก่าและแก้ไขได้ • เวลาเข้า/ออก และราคาให้กรอกตามจริง • ไม่ใช้ข้อมูลรถวันเก่ามาแทน</p>'+
  groups+'</details></td></tr>';
}
function renderPreview(result){
 prepared=result;outputs=[];$('resultsCard').classList.add('hidden');$('previewCard').classList.remove('hidden');
 $('previewHint').textContent=result.mode==='prep'?'ข้อมูลที่ผ่านเงื่อนไข Load = KNS และไม่ Cancel — แก้จำนวนถุงได้หากแผนมีการเปลี่ยนแปลง':
 'ข้อมูลจาก Loading พร้อมช่องแก้ไขรถ และรายละเอียด Export เพิ่มเติมรายคัน (คลิกเปิดใต้รถ)';
 const head=result.mode==='prep'?['สาย','Booking','สินค้า','คันที่','ถุง / คัน','วิธีโหลด','ตรวจสอบ']:
 ['สาย','Booking','คัน','ทะเบียน','เลขตู้','Seal','Tare','Gross','Net','Weight Card'];
 $('previewHead').innerHTML='<tr>'+head.map(x=>'<th>'+x+'</th>').join('')+'</tr>';
 if(result.mode==='prep'){
  $('previewRows').innerHTML=result.cards.map((c,i)=>'<tr><td><span class="pill">'+c.kind+'</span></td><td><b>'+esc(c.job.booking)+'</b></td><td>'+esc(c.job.product)+'</td><td>'+c.no+'/'+c.job.fcl+'</td><td><input type="number" min="1" step="1" inputmode="numeric" class="bagInput" data-i="'+i+'" value="'+esc(c.bags)+'" placeholder="ระบุจำนวน" aria-label="จำนวนถุง '+esc(c.job.booking)+' คัน '+c.no+'"></td><td><span class="pill '+(!c.bags?'pending':c.bags===20?'ok':'')+'" id="method'+i+'">'+esc(c.category)+'</span></td><td><span class="pill '+(!c.bags?'pending':'ok')+' prep-review-state">'+(!c.bags?'ตรวจจำนวน':'พร้อม')+'</span></td></tr>').join('');
  // "input" updates the review immediately, including pasted values.
  // "change" handles browser controls that commit only after losing focus.
  // Generate always reads the current DOM values again (autofill/scripts
  // may alter .value without firing either event).
  for(const el of $('previewRows').querySelectorAll('.bagInput')){
   el.addEventListener('input',()=>updatePrepInputs(result));
   el.addEventListener('change',()=>updatePrepInputs(result));
  }
 }else{
  const editable=['registration','container','seal','tare','gross','net','weightCard'];
  $('previewRows').innerHTML=result.draftOnly?'<tr><td colspan="10" class="empty">⏳ รอข้อมูลรถเข้าจริง — แผน '+fmt(result.plannedUnits||0)+' ตู้ / รถที่ยืนยันใน Loading 0 คัน บันทึกงานฉบับร่างไว้บน Alice Cloud ได้ แล้วเพิ่มไฟล์ Loading และทะเบียนรถเมื่อมีข้อมูล</td></tr>':
   result.cards.map((v,i)=>'<tr class="daily-vehicle-row"><td><span class="pill">'+esc(v.kind)+'</span></td><td><b>'+esc(v.booking)+'</b></td><td>'+v.carNo+'</td>'+
    editable.map(key=>'<td><input class="daily-main-input" data-i="'+i+'" data-key="'+key+'" value="'+esc(v[key]??'')+'" placeholder="'+(key==='weightCard'?'ยังไม่มี':'—')+'"></td>').join('')+'</tr>'+
    dailyExtraPanel(v,i)).join('');
  $('previewRows').querySelectorAll('.daily-main-input').forEach(el=>el.addEventListener('change',()=>{
   const v=result.cards[Number(el.dataset.i)],key=el.dataset.key;
   v[key]=['tare','gross','net'].includes(key)?(Number(el.value)||null):el.value.trim();
  }));
  $('previewRows').querySelectorAll('.daily-extra-input:not([readonly])').forEach(el=>el.addEventListener('change',()=>{
   const v=result.cards[Number(el.dataset.i)],key=el.dataset.key,value=el.value.trim();
   v.extra[key]=DAILY_RATING_KEYS.has(key)?(value===''?0:Number(value)):value;
   if(DAILY_RATING_KEYS.has(key)||key==='loadStart'||key==='loadEnd'){
    recalcDailyScores(v.extra);
    const panel=el.closest('.daily-details');
    for(const key2 of ['docTotal','docPct','carrierTotal','carrierPct','averageScore','assessmentTotal','duration']){
     const related=panel.querySelector('.daily-extra-input[data-key="'+key2+'"]');
     if(related)related.value=v.extra[key2]??'';
    }
   }
  }));
 }
 if(result.mode==='prep')updatePrepInputs(result);
 else showPreviewWarnings(result);
 $('generateBtn').disabled=!!result.draftOnly||!!result.planOnly;
 $('generateBtn').textContent=result.planOnly?'📋 แสดงรายละเอียดแบช (ยังไม่มีใบแจ้งโหลดให้สร้าง)':result.draftOnly?'⌛ รอรถเข้าจริงก่อนสร้าง Excel':'⚡ สร้างไฟล์ Excel';
 $('generateHelp').textContent=result.draftOnly?'ขณะนี้เป็นฉบับร่าง — กด “บันทึกงานวันนี้” ใน Alice Cloud เพื่อเก็บไฟล์และแผนไว้ก่อน ไม่สร้างรายงานรถที่ยังไม่เข้าจริง':
 result.mode==='prep'?'ไฟล์ผลลัพธ์จะเป็น .xlsx ใหม่สำหรับดาวน์โหลด โดยไม่เปลี่ยนไฟล์ต้นทาง':'ระบบจะสร้าง Excel จากไฟล์ที่อัปโหลด ไม่เขียนทับไฟล์ต้นฉบับในเครื่อง';
 status(result.planOnly?'อ่านแผนแบชสำเร็จ — ตรวจรายละเอียดตำแหน่งสินค้าและ Qty Allocate ในตารางด้านบน':result.draftOnly?'อ่านไฟล์สำเร็จ — รอข้อมูลรถเข้าจริง (แผน '+fmt(result.plannedUnits)+' ตู้ / Loading 0 คัน) สามารถบันทึกงานฉบับร่างใน Alice Cloud ได้':'วิเคราะห์สำเร็จ '+result.units+' '+(result.mode==='prep'?'ตู้':'คัน')+' — ตรวจรายการก่อนกดสร้างไฟล์',result.draftOnly?'info':'good');
 $('previewCard').scrollIntoView({behavior:'smooth',block:'start'});
}

async function analyze(){
 if(busy)return;busy=true;$('analyzeBtn').disabled=true;resetPreview();
 try{
  const iso=validDate();
  const plan=await readPlans(iso);
  if(mode==='prep')await resolvePrepTemplates(plan);
  const batchPlan=await readBatchPlan();
  const result=mode==='prep'?analyzePrep(plan):await analyzeDaily(iso,plan);
  result.batchPlan=batchPlan;
  if(batchPlan&&result.mode==='daily'&&result.cards.length){const comparison=compareBatchPlan(batchPlan,result.cards);result.batchChecks=comparison.checks;result.warnings.push(...comparison.notes)}
  renderPreview(result);renderBatchPlan(batchPlan,result.batchChecks||[]);
  if(result.planOnly){$('previewCard').classList.add('hidden');$('batchPlanCard').scrollIntoView({behavior:'smooth',block:'start'})}log('วิเคราะห์สำเร็จ: '+result.bookings+' Booking / '+result.units+' คัน / แจ้งเตือน '+result.warnings.length);
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
 // Read displayed input values immediately before generating, not just the
 // change-event state. This prevents falsely rejecting filled-in quantities.
 const missing=updatePrepInputs(prepared);
 if(missing.length){
  const idx=missing[0],card=prepared.cards[idx];
  const el=$('previewRows').querySelector('.bagInput[data-i="'+idx+'"]');
  el?.focus();
  el?.scrollIntoView({behavior:'smooth',block:'center'});
  throw Error('กรุณาระบุจำนวนถุงให้ครบ: '+card.job.booking+' '+card.job.product+
    ' คัน '+card.no+'/'+card.job.fcl+' (ยังขาด '+missing.length+' คัน)');
 }
 await resolvePrepTemplates(prepared.plans);
 const missing=missingPrepTemplates(prepared.plans);
 if(missing.length)throw Error('ยังไม่มีแม่แบบใบแจ้งโหลด '+missing.join(' และ ')+' (.xlsx) กรุณาอัปโหลด Excel ต้นฉบับและเลือกประเภทไฟล์เป็นใบแจ้งโหลด STP/PTR หรือเข้าสู่ระบบ Supabase เพื่อใช้แม่แบบที่เคยบันทึกไว้');
 for(const j of jobs){j._cardBagValues=prepared.cards.filter(c=>c.job===j).map(c=>c.bags?Number(c.bags):null)}
 const out=[];
 for(const kind of ['STP','PTR']){
  const group=jobs.filter(j=>j.kind===kind);
  if(!group.length)continue;
  const template=sole('form'+kind);
  if(!template)throw Error('ไม่พบแม่แบบ Excel '+kind+' กรุณาอัปโหลดแบบฟอร์มต้นฉบับ');
  if(extension(template.name)!=='xlsx')throw Error('แม่แบบ '+kind+' ต้องเป็นไฟล์ .xlsx เท่านั้น');
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
 return kind==='STP'?{date:'C',invoice:'D',contract:'E',booking:'F',agent:'G',etd:'H',receipt:'I',place:'J',product:'K',batch:'L',code:'M',bags:'N',kg:'O',tare:'P',gross:'Q',net:'R',carrier:'T',car:'W',registration:'X',driver:'Y',container:'Z',seal:'AA',weightCard:'AB'}:
 {date:'A',invoice:'B',contract:'C',booking:'D',agent:'E',etd:'F',receipt:'G',place:'H',product:'I',batch:'J',code:'K',bags:'L',kg:'M',tare:'N',gross:'O',net:'P',carrier:'R',car:'U',registration:'V',driver:'W',container:'X',seal:'Y',weightCard:'Z'};
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
 if(prepared?.draftOnly||!prepared?.cards?.length)throw Error('ยังไม่มีรถเข้าจริงใน Loading — บันทึกฉบับร่างแล้วอัปเดตเมื่อรถเข้า');
 const iso=validDate(),results=[],dateSerial=C().excelDate(iso);
 const normal=s=>String(s??'').toUpperCase().replace(/[^A-Z0-9]/g,'');
 const positive=v=>Number.isFinite(Number(v))&&Number(v)>0;
 for(const kind of ['STP','PTR']){
  // All non-PTR loading reports, including Domestic / SSAMS, share ONE STP Export workbook.
  const groups=(kind==='STP'?['STP','DOMESTIC']:['PTR'])
   .map(group=>({group,source:prepared.filesByKind[group]})).filter(x=>x.source);
  if(!groups.length)continue;
  const exp=sole('export'+kind);if(!exp)throw Error('ไม่มี Export '+kind+' รองรับ Loading');
  const loadPatches={};
  for(const {group,source} of groups)loadPatches[group]=await C().ExcelPatch.load(source.load);
  const expPatch=await C().ExcelPatch.load(exp);
  const original=await C().parseXlsx(exp),exportSheetName=findExportSheet(original,iso);
  const expSheet=expPatch.sheetByName(exportSheetName);
  if(!expSheet)throw Error('ไม่พบชีต Export '+exportSheetName);
  const rows=C().sheetRows(original,exportSheetName),map=exportMap(kind);
  const supplementMap=dailySupplementColumns(kind);
  const existing=[],claimed=new Set();
  let last=6,maxDailyNo=0;
  for(let r=7;r<=rows.length;r++){
   const values=rows[r-1]||[],booking=C().clean(values[colIndex(map.booking)]);
   const car=C().clean(values[colIndex(map.car)]),date=Number(values[colIndex(map.date)])||0;
   if(booking)last=Math.max(last,r);
   if(date!==dateSerial||!booking)continue;
   maxDailyNo=Math.max(maxDailyNo,Number(car)||0);
   existing.push({row:r,booking,container:normal(values[colIndex(map.container)]),
    registration:normal(values[colIndex(map.registration)]),
    weightCard:normal(values[colIndex(map.weightCard)]),
    place:normal(values[colIndex(map.place)]),
    product:normal(values[colIndex(map.product)]),
    code:normal(values[colIndex(map.code)]),
    bags:Number(values[colIndex(map.bags)])||0,car});
  }
  let next=last+1,matched=0,added=0;
  for(const {group,source} of groups){
   const loadPatch=loadPatches[group],loadSheet=loadPatch.sheetByName(source.sheetName);
   if(!loadSheet)throw Error('ไม่มีชีต Loading '+source.sheetName);
   const list=source.vehicles,seenInput=new Set();
   for(const v of list){
    const key=v.booking+'|'+v.carNo;
    if(seenInput.has(key))throw Error('Loading '+group+' มีลำดับรถซ้ำ '+key);
    seenInput.add(key);
    if(!v.container&&group!=='DOMESTIC'&&$('strictMode').value==='strict')
     throw Error('ยังไม่มีเลขตู้ '+v.booking+' คัน '+v.carNo);
    const candidates=existing.filter(r=>!claimed.has(r.row)&&r.booking===v.booking);
    // First match by container; domestic deliveries often have NO booking or container.
    let match=candidates.filter(r=>normal(v.container)&&r.container===normal(v.container));
    if(match.length>1)throw Error('เลขตู้ซ้ำ '+v.booking+' '+v.container);
    let target=match[0]||null;
    if(!target&&v.weightCard){
     match=candidates.filter(r=>r.weightCard&&r.weightCard===normal(v.weightCard));
     if(match.length>1)throw Error('Weight Card ซ้ำใน Export '+v.weightCard);
     target=match[0]||null;
    }
    if(!target&&v.registration){
     match=candidates.filter(r=>r.registration&&r.registration===normal(v.registration)&&
       (group!=='DOMESTIC'||r.place===normal(v.returnPlace)));
     if(match.length>1)throw Error('ทะเบียนซ้ำใน Export '+v.booking+' '+v.registration);
     target=match[0]||null;
    }
    if(!target&&group==='DOMESTIC'&&!v.weightCard&&!v.registration){
     match=candidates.filter(r=>r.place===normal(v.returnPlace)&&r.product===normal(v.product)&&
       r.code===normal(v.code)&&r.bags===v.bags);
     if(match.length===1)target=match[0];
    }
    // Never silently reuse a row when the identified container conflicts.
    if(target&&v.container&&normal(v.container)!==normal('-')&&
       target.container&&target.container!==normal(v.container))
     throw Error('เลขตู้ไม่ตรงกัน '+v.booking+' คัน '+v.carNo);
    let row,dayNo;
    if(target){row=target.row;dayNo=target.car;claimed.add(row);matched++}
    else{
     row=next++;dayNo=++maxDailyNo;added++;
     await preserveNewRowStyle(expPatch,expSheet,Math.max(7,last),row);
     claimed.add(row);
    }
    // Fill confirmed edits back into the corresponding Loading workbook.
    const updatedLoad={I:v.extra?.carrier??v.carrier,J:v.registration,K:v.container,L:v.seal,
     U:v.weightCard,V:v.tare,W:v.gross,X:v.net};
    for(const [col,val] of Object.entries(updatedLoad))
     if(val!==null&&val!==undefined&&val!=='')await loadPatch.setCell(loadSheet,col+v.row,val);
    const extra=v.extra||initializeDailyExtra(v);
    const desired={date:dateSerial,invoice:extra.invoice,contract:extra.contract,booking:v.booking,
     agent:extra.agent,etd:extra.etd,receipt:extra.receipt,place:extra.returnPlace,
     product:v.product,batch:v.batch,code:v.code,bags:v.bags||null,kg:v.kg||null,
     tare:v.tare,gross:v.gross,net:v.net,carrier:extra.carrier,car:dayNo,
     registration:v.registration,driver:extra.driver,container:v.container,seal:v.seal,
     weightCard:v.weightCard};
    const oldValues=rows[row-1]||[];
    const alwaysSync=new Set(['invoice','contract','booking','agent','etd','receipt','place',
     'tare','gross','net','weightCard','registration','container','seal',
     'product','batch','code','bags','kg','carrier','driver']);
    for(const [field,col] of Object.entries(map)){
     const val=desired[field];
     if(val===null||val===undefined||val==='')continue;
     if(['net','gross','tare'].includes(field)&&!positive(val))continue;
     const old=oldValues[colIndex(col)],hasOld=old!==null&&old!==undefined&&String(old).trim()!=='';
     if(!target||!hasOld||alwaysSync.has(field))await expPatch.setCell(expSheet,col+row,val);
    }
    // Additional fields available in the Export workbook, not in Loading.
    // Fixed historical values are exposed in the UI and may be overridden per vehicle.
    for(const [key,col] of Object.entries({
     contractorTeam:kind==='PTR'?'Q':'S',
     price:kind==='PTR'?'S':'U',
     overnight:kind==='PTR'?'T':'V',
     ...supplementMap
    })){
     let value=extra[key];
     if(value===null||value===undefined||String(value).trim()==='')continue;
     if(['arrive','depart','loadStart','loadEnd','duration'].includes(key)){
      value=asExcelTime(value);
      if(value===null)throw Error('รูปแบบเวลาไม่ถูกต้อง: '+key+' คัน '+v.carNo+' (ใช้ HH:MM)');
     }else if(['price','documentDelivery','documentAccuracy','coordination',
      'docTotal','docPct','punctuality','truckCondition','compliance','driverCourtesy',
      'carrierTotal','carrierPct','averageScore','assessmentTotal'].includes(key)&&value!=='-'){
      value=Number(value);
      if(!Number.isFinite(value))throw Error('ตัวเลขไม่ถูกต้อง: '+key+' คัน '+v.carNo);
     }
     await expPatch.setCell(expSheet,col+row,value);
    }
   }
   results.push(await loadPatch.save('AliceOffice_Loading_'+group+'_'+stamp(iso)+'.xlsx'));
   log('Daily '+group+': '+list.length+' คัน → Export '+kind);
  }
  results.push(await expPatch.save('AliceOffice_Export_'+kind+'_'+stamp(iso)+'.xlsx'));
  log('Export '+kind+': จับคู่แถวเดิม '+matched+', เพิ่มแถว '+added);
 }
 if(!results.length)throw Error('ไม่พบคู่ Loading และ Export ให้สร้างรายงาน');
 return results;
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
  if(window.AliceCloud?.isAuthenticated?.()){
   try{
    const cloudSaved=await window.AliceCloud.saveRun(cloudSnapshot());
    log('Supabase บันทึกงานและไฟล์ '+cloudSaved.saved+' ไฟล์');
    status('Excel สำเร็จ '+results.length+' ไฟล์ และซิงก์เข้าคลังส่วนตัวเรียบร้อยแล้ว ☁️','good');
   }catch(cloudError){
    log('CLOUD ERROR '+cloudError.message);
    status('สร้าง Excel สำเร็จแล้ว แต่ซิงก์ Supabase ไม่สำเร็จ: '+cloudError.message,'error');
   }
  }
 }catch(e){status('ยังสร้างไฟล์ไม่ได้: '+e.message,'error');log('ERROR '+e.message);console.error(e)}
 finally{busy=false;$('generateBtn').disabled=false}
}
function cloudSnapshot(){
 return {date:$('workDate').value,mode,files:files.map(e=>({file:e.file,role:e.role})),outputs:[...outputs],prepared};
}
async function addCloudFiles(entries,{replaceRole=false,preservePreview=false}={}){
 if(!Array.isArray(entries))throw Error('รายการไฟล์จากคลาวด์ไม่ถูกต้อง');
 for(const entry of entries){
  if(!entry.file||!entry.role)continue;
  if(replaceRole&&entry.role!=='truck')files=files.filter(x=>x.role!==entry.role);
  if(files.some(x=>x.role===entry.role&&x.file.name===entry.file.name&&x.file.size===entry.file.size))continue;
  files.push({id:++counter,file:entry.file,role:entry.role});
 }
 renderFiles();if(!preservePreview)resetPreview();
 status('โหลดไฟล์จากคลาวด์ '+entries.length+' รายการแล้ว ตรวจสอบประเภทไฟล์ได้เลยค่ะ','good');
}
async function restoreCloudRun({date,mode:nextMode,entries,cardData=[]}){
 files=[];outputs=[];showMode(nextMode);
 $('workDate').value=date;
 await addCloudFiles(entries,{replaceRole:false});
 if(nextMode==='daily'&&cardData.length){
  await analyze();
  if(prepared?.mode==='daily'&&!prepared.draftOnly){
   let restored=0;
   const identity=x=>String(x??'').replace(/[^a-z0-9ก-๙]/gi,'').toUpperCase();
   const used=new Set();
   for(const v of prepared.cards){
    const candidates=cardData.map((card,index)=>({card,index})).filter(({card,index})=>
     !used.has(index)&&card.flow===v.kind&&card.booking===v.booking&&Number(card.car_no)===Number(v.carNo));
    let match=candidates.find(({card})=>card.weightCard&&identity(card.weightCard)===identity(v.weightCard));
    if(!match)match=candidates.find(({card})=>card.container&&identity(card.container)===identity(v.container));
    if(!match&&candidates.length===1)match=candidates[0];
    if(!match)continue;
    used.add(match.index);
    const card=match.card;
    for(const key of ['registration','container','seal','tare','gross','net','weightCard']){
     if(card[key]!==undefined&&card[key]!==null)v[key]=card[key];
    }
    if(card.extra&&typeof card.extra==='object')v.extra={...v.extra,...card.extra};
    restored++;
   }
   renderPreview(prepared);
   status('เรียกคืนข้อมูลที่กรอกในเว็บ '+restored+' คัน จาก Alice Cloud แล้ว ตรวจสอบก่อนสร้างไฟล์','good');
  }
 }
}
window.AliceOfficeUI={snapshot:cloudSnapshot,addCloudFiles,restoreCloudRun,log,status};
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
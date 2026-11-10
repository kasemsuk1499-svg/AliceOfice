(function(){
'use strict';
const $=id=>document.getElementById(id);
const CONFIG=window.ALICE_OFFICE_CONFIG;
const BUCKET=CONFIG?.bucket||'alice-office-private';
const ROLES={formSTP:'แม่แบบใบแจ้ง STP',formPTR:'แม่แบบใบแจ้ง PTR',loadingSTP:'Loading STP ล่าสุด',loadingPTR:'Loading PTR ล่าสุด',exportSTP:'Export STP ล่าสุด',exportPTR:'Export PTR ล่าสุด',loadingDomestic:'Domestic / SSAMS ล่าสุด'};
const esc=s=>String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
const bytesToHex=array=>Array.from(new Uint8Array(array)).map(b=>b.toString(16).padStart(2,'0')).join('');
const ext=name=>String(name||'').match(/\.([a-z0-9]+)$/i)?.[1]?.toLowerCase()||'bin';
let db=null,user=null,channel=null,refreshTimer=null,connected=false;
function feedback(value,ok){const el=$('cloudFeedback');if(!el)return;el.textContent=value;el.className='cloud-feedback'+(ok===false?' error':ok===true?' good':'')}
function statusTag(value,online){const el=$('cloudIndicator');if(el){el.textContent=value;el.className='cloud-state'+(online?' online':'')}}
function note(message){try{window.AliceOfficeUI?.log?.('CLOUD '+message)}catch{}}
function signedIn(){return !!db&&!!user}
function account(){if(!signedIn())throw Error('กรุณาเข้าสู่ระบบ Supabase ก่อน');return user}
function getContext(){
 const snap=window.AliceOfficeUI?.snapshot?.();
 if(!snap)throw Error('หน้าเว็บยังไม่พร้อม');
 return snap;
}
async function digest(blob){
 if(!globalThis.crypto?.subtle)throw Error('ต้องเปิดเว็บด้วย HTTPS เพื่อใช้ระบบจัดเก็บไฟล์อย่างปลอดภัย');
 return bytesToHex(await crypto.subtle.digest('SHA-256',await blob.arrayBuffer()));
}
function mime(name){
 switch(ext(name)){
  case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  case 'xls':return 'application/vnd.ms-excel';
  case 'csv':return 'text/csv';
  default:return 'application/octet-stream';
 }
}
function validateRole(role){if(!/^[A-Za-z0-9_-]{1,40}$/.test(role))throw Error('ประเภทไฟล์ไม่ถูกต้อง '+role)}
async function storeFile(file,category,role,workDate=null,workflow=null){
 const u=account();validateRole(role);
 if(file.size>30*1024*1024)throw Error('ไฟล์ '+file.name+' เกิน 30MB');
 const hash=await digest(file),name=String(file.name||'untitled.xlsx');
 const part=category==='template'?'templates':('runs/'+workDate+'/'+workflow+'/'+category);
 const objectPath=[u.id,part,role,hash+'.'+ext(name)].join('/');
 const {error:uploadError}=await db.storage.from(BUCKET).upload(objectPath,file,{upsert:true,contentType:mime(name),cacheControl:'3600'});
 if(uploadError)throw new Error('อัปโหลด '+name+': '+uploadError.message);
 const record={owner_id:u.id,category,role,filename:name,object_path:objectPath,work_date:workDate,workflow,byte_size:file.size};
 const {error:dbError}=await db.from('alice_office_files').upsert(record,{onConflict:'owner_id,object_path'});
 if(dbError)throw new Error('บันทึกข้อมูลไฟล์ '+name+': '+dbError.message);
 return record;
}
async function saveTemplates(){
 const state=getContext(),selected=state.files.filter(f=>Object.prototype.hasOwnProperty.call(ROLES,f.role));
 if(!selected.length){feedback('ยังไม่มีแม่แบบ Excel ที่เลือกไว้ในหน้าเว็บ',false);return}
 $('cloudSaveTemplates').disabled=true;
 try{
  feedback('กำลังเก็บแม่แบบ '+selected.length+' ไฟล์...');
  for(const e of selected)await storeFile(e.file,'template',e.role);
  feedback('บันทึกแม่แบบสำเร็จ '+selected.length+' ไฟล์ และใช้ต่อในวันถัดไปได้',true);
  note('บันทึกแม่แบบ '+selected.map(x=>x.role).join(', '));
  await refresh();
 }catch(e){feedback(e.message,false);throw e}
 finally{$('cloudSaveTemplates').disabled=false}
}
function outputRole(filename){
 const s=String(filename);
 if(s.includes('Loading_DOMESTIC'))return 'loadingDomestic';
 for(const kind of ['STP','PTR']){
  if(s.includes('Export_'+kind))return 'export'+kind;
  if(s.includes('Loading_'+kind))return 'loading'+kind;
  if(s.includes('ใบแจ้งโหลด_'+kind))return 'generatedForm'+kind;
 }
 return 'result';
}
async function saveRun(state=getContext()){
 const u=account();
 if(!state.date||!['prep','daily'].includes(state.mode))throw Error('กรุณาเลือกวันที่และประเภทงานก่อนบันทึก');
 const inputs=state.files||[],outputs=state.outputs||[];
 if(!inputs.length&&!outputs.length)throw Error('ยังไม่มีไฟล์สำหรับบันทึก');
 feedback('กำลังบันทึกงานและไฟล์ Excel เข้าคลังส่วนตัว...');
 const cardData=(state.prepared?.cards||[]).map(card=>{
  if(state.mode==='prep')return {flow:card.kind,booking:card.job.booking,car_no:card.no,product:card.job.product,bags:card.bags,method:card.category};
  return {flow:card.kind,booking:card.booking,car_no:card.carNo,product:card.product,container:card.container,registration:card.registration,seal:card.seal,tare:card.tare,gross:card.gross,net:card.net,weightCard:card.weightCard,extra:{...(card.extra||{})}};
 });
 const snapshot={version:2,bookings:state.prepared?.bookings||0,units:state.prepared?.units||0,warnings:state.prepared?.warnings?.length||0,cards:cardData};
 const draft={owner_id:u.id,work_date:state.date,workflow:state.mode,payload:snapshot,updated_at:new Date().toISOString()};
 const {error:draftError}=await db.from('alice_office_drafts').upsert(draft,{onConflict:'owner_id,work_date,workflow'});
 if(draftError)throw new Error('บันทึกงาน: '+draftError.message);
 if(cardData.length){
  const vehicleRows=cardData.filter(c=>c.booking&&c.car_no).map(c=>({owner_id:u.id,work_date:state.date,flow:c.flow,booking:c.booking,car_no:c.car_no,data:c,updated_at:new Date().toISOString()}));
  const {error:rowsError}=await db.from('alice_office_vehicles').upsert(vehicleRows,{onConflict:'owner_id,work_date,flow,booking,car_no'});
  if(rowsError)throw new Error('บันทึกรถ: '+rowsError.message);
 }
 let saved=0;
 for(const e of inputs){
  if(e.role==='ignore')continue;
  await storeFile(e.file,'source',e.role,state.date,state.mode);saved++;
  // Always remember clean form templates; keep old versions for audit.
  if(e.role==='formSTP'||e.role==='formPTR')await storeFile(e.file,'template',e.role);
 }
 for(const output of outputs){
  const file=new File([output.blob],output.filename,{type:mime(output.filename)});
  const role=outputRole(output.filename);
  await storeFile(file,'output',role,state.date,state.mode);saved++;
  // An updated Export/Loading workbook is the latest continuity snapshot.
  if(role.startsWith('export')||role.startsWith('loading'))await storeFile(file,'template',role);
 }
 feedback('บันทึกงาน '+state.date+' สำเร็จ '+saved+' ไฟล์ — ซิงก์ข้ามอุปกรณ์แล้ว',true);
 note('บันทึกงาน '+state.date+'/'+state.mode+' '+saved+' ไฟล์');
 await refresh();
 return {saved};
}
async function loadObject(record){
 if(!record?.object_path?.startsWith(account().id+'/'))throw Error('เส้นทางไฟล์ไม่ใช่ของบัญชีนี้');
 const {data,error}=await db.storage.from(BUCKET).download(record.object_path);
 if(error)throw Error('เปิดไฟล์ '+record.filename+': '+error.message);
 return new File([data],record.filename,{type:mime(record.filename)});
}
async function restoreTemplate(record){
 try{
  feedback('กำลังโหลดแม่แบบ '+record.filename+'...');
  const file=await loadObject(record);
  await window.AliceOfficeUI.addCloudFiles([{file,role:record.role}],{replaceRole:true});
  feedback('โหลดแม่แบบ '+record.filename+' แล้ว พร้อมใช้สร้างงานวันใหม่',true);
  note('โหลดแม่แบบ '+record.role);
 }catch(e){feedback(e.message,false)}
}
async function ensureTemplates(kinds){
 if(!signedIn()||!Array.isArray(kinds)||!kinds.length)return [];
 const required=[...new Set(kinds.filter(k=>k==='STP'||k==='PTR').map(k=>'form'+k))];
 if(!required.length)return [];
 const {data,error}=await db.from('alice_office_files')
  .select('role,filename,object_path,created_at')
  .eq('owner_id',account().id).eq('category','template')
  .in('role',required).order('created_at',{ascending:false}).limit(100);
 if(error)throw Error('ค้นหาแม่แบบ Excel ในคลาวด์ไม่สำเร็จ: '+error.message);
 const entries=[],seen=new Set();
 for(const record of data||[]){
  if(seen.has(record.role))continue;
  seen.add(record.role);
  entries.push({role:record.role,file:await loadObject(record)});
 }
 if(entries.length){
  await window.AliceOfficeUI.addCloudFiles(entries,{replaceRole:false,preservePreview:true});
  feedback('ใช้แม่แบบ Excel จากคลาวด์อัตโนมัติ '+entries.map(x=>x.role.replace('form','')).join(', '),true);
  note('เรียกแม่แบบอัตโนมัติ '+entries.map(x=>x.role).join(', '));
 }
 return entries.map(x=>x.role.replace('form',''));
}
async function restoreRun(date,workflow){
 const u=account();
 if(!window.confirm('โหลดงาน '+date+' ('+workflow+') จากคลาวด์? ไฟล์ที่เลือกอยู่ในหน้านี้จะถูกแทนที่'))return;
 try{
  feedback('กำลังเรียกคืนงานวันที่ '+date+'...');
  const {data,error}=await db.from('alice_office_files').select('role,filename,object_path,created_at').eq('owner_id',u.id).eq('category','source').eq('work_date',date).eq('workflow',workflow).order('created_at',{ascending:false}).limit(150);
  if(error)throw Error(error.message);
  if(!data?.length)throw Error('ยังไม่มีไฟล์ต้นทางของงานนี้ในคลาวด์');
  // Repeated saves create new versions. Restore the NEWEST for each unique
  // Loading / Export / plan role, not every old version of the same role.
  const latest=[],seen=new Set();
  for(const row of data){
   if(row.role!=='truck'&&seen.has(row.role))continue;
   if(row.role!=='truck')seen.add(row.role);
   latest.push(row);
  }
  const entries=[];
  for(const row of latest)entries.push({role:row.role,file:await loadObject(row)});
  const {data:snapshots,error:snapshotError}=await db.from('alice_office_drafts')
   .select('payload').eq('owner_id',u.id).eq('work_date',date).eq('workflow',workflow).limit(1);
  if(snapshotError)throw Error('เรียกข้อมูลที่กรอกไว้ไม่สำเร็จ: '+snapshotError.message);
  await window.AliceOfficeUI.restoreCloudRun({date,mode:workflow,entries,cardData:snapshots?.[0]?.payload?.cards||[]});
  feedback('เรียกคืนงาน '+date+' สำเร็จ '+entries.length+' ไฟล์ รวมข้อมูลที่กรอกผ่านเว็บ',true);
  note('เรียกคืนงาน '+date+' '+entries.length+' ไฟล์');
 }catch(e){feedback(e.message,false)}
}
function renderFileList(records){
 const box=$('cloudTemplates');if(!box)return;
 const seen=new Set(),latest=[];
 for(const r of records||[]){if(seen.has(r.role))continue;seen.add(r.role);latest.push(r)}
 box.innerHTML=latest.length?latest.map((r,i)=>'<div class="cloud-entry"><span><strong>'+esc(ROLES[r.role]||r.role)+'</strong><small>'+esc(r.filename)+'</small></span><button data-index="'+i+'" type="button">ใช้ไฟล์</button></div>').join(''):'ยังไม่มีแม่แบบ กด “เก็บแม่แบบล่าสุด” หลังเพิ่มไฟล์';
 box.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>restoreTemplate(latest[Number(b.dataset.index)])));
}
function renderRuns(records){
 const box=$('cloudRuns');if(!box)return;
 box.innerHTML=records?.length?records.map((r,i)=>{
  const meta=r.payload||{};
  return '<div class="cloud-entry"><span><strong>'+esc(r.work_date)+' · '+(r.workflow==='prep'?'เตรียมโหลด':'Daily Report')+'</strong><small>'+Number(meta.units||0)+' คัน · อัปเดต '+esc(new Date(r.updated_at).toLocaleString('th-TH'))+'</small></span><button type="button" data-index="'+i+'">เปิด</button></div>';
 }).join(''):'ยังไม่มีงานที่บันทึกในคลาวด์';
 box.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>restoreRun(records[Number(b.dataset.index)].work_date,records[Number(b.dataset.index)].workflow)));
}
async function refresh(){
 if(!signedIn())return;
 const id=account().id;
 const [t,r]=await Promise.all([
  db.from('alice_office_files').select('role,filename,object_path,created_at').eq('owner_id',id).eq('category','template').order('created_at',{ascending:false}).limit(100),
  db.from('alice_office_drafts').select('work_date,workflow,payload,updated_at').eq('owner_id',id).order('updated_at',{ascending:false}).limit(20)
 ]);
 if(t.error)throw Error('ดึงแม่แบบไม่สำเร็จ: '+t.error.message);
 if(r.error)throw Error('ดึงประวัติงานไม่สำเร็จ: '+r.error.message);
 renderFileList(t.data);renderRuns(r.data);
}
function enqueueRefresh(){clearTimeout(refreshTimer);refreshTimer=setTimeout(()=>refresh().catch(e=>feedback(e.message,false)),500)}
function joinRealtime(){
 if(!signedIn())return;
 if(channel){db.removeChannel(channel);channel=null}
 channel=db.channel('alice-office-'+user.id)
  .on('postgres_changes',{event:'*',schema:'public',table:'alice_office_drafts',filter:'owner_id=eq.'+user.id},enqueueRefresh)
  .on('postgres_changes',{event:'*',schema:'public',table:'alice_office_files',filter:'owner_id=eq.'+user.id},enqueueRefresh)
  .on('postgres_changes',{event:'*',schema:'public',table:'alice_office_vehicles',filter:'owner_id=eq.'+user.id},()=>{
    feedback('มีข้อมูลรถอัปเดตจากอีกอุปกรณ์ หากกำลังแก้รายงานอยู่ ให้เปิดงานล่าสุดใหม่ก่อนสร้างไฟล์');
    enqueueRefresh();
  })
  .subscribe(state=>{if(state==='SUBSCRIBED')statusTag('เรียลไทม์พร้อม',true);else if(state==='CHANNEL_ERROR'||state==='TIMED_OUT')statusTag('เชื่อมเรียลไทม์ขัดข้อง',false)});
}
async function setUser(nextUser){
 user=nextUser||null;
 $('cloudGuest').classList.toggle('hidden',!!user);
 $('cloudUser').classList.toggle('hidden',!user);
 $('cloudAccount').textContent=user?.email||'';
 statusTag(user?'กำลังซิงก์...':'ยังไม่เข้าสู่ระบบ',false);
 if(channel){db.removeChannel(channel);channel=null}
 if(user){try{await refresh();joinRealtime();feedback('เข้าสู่ระบบแล้ว ไฟล์จะเก็บในพื้นที่ส่วนตัวของบัญชีนี้',true)}catch(e){feedback(e.message,false)}}
 else{renderFileList([]);renderRuns([]);feedback('ไฟล์ยังอยู่ในเครื่องนี้จนกว่าจะเข้าสู่ระบบ')}
}
async function login(){
 const email=$('cloudEmail').value.trim();
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){feedback('กรุณากรอกอีเมลให้ถูกต้อง',false);return}
 if(location.protocol!=='https:'){feedback('ระบบเข้าสู่ระบบคลาวด์ต้องเปิดผ่าน HTTPS (GitHub Pages) ก่อน',false);return}
 $('cloudSignIn').disabled=true;
 try{
  const redirectTo=location.origin+location.pathname;
  const {error}=await db.auth.signInWithOtp({email,options:{emailRedirectTo:redirectTo,shouldCreateUser:true}});
  if(error)throw error;
  feedback('ส่งลิงก์เข้าสู่ระบบแล้ว กรุณาตรวจอีเมลและกดลิงก์',true);
 }catch(e){feedback('ส่งอีเมลไม่สำเร็จ: '+e.message,false)}
 finally{$('cloudSignIn').disabled=false}
}
async function logout(){
 const {error}=await db.auth.signOut();
 if(error)feedback(error.message,false);else await setUser(null);
}
async function init(){
 if(!CONFIG?.publishableKey||!window.supabase?.createClient){statusTag('คลาวด์ไม่พร้อม',false);feedback('ไม่สามารถโหลดไลบรารี Supabase หรือคีย์การเชื่อมต่อ',false);return}
 db=window.supabase.createClient(CONFIG.url,CONFIG.publishableKey,{auth:{autoRefreshToken:true,persistSession:true,detectSessionInUrl:true}});
 $('cloudSignIn').addEventListener('click',login);
 $('cloudSignOut').addEventListener('click',logout);
 $('cloudSaveTemplates').addEventListener('click',()=>saveTemplates().catch(e=>note('ERROR '+e.message)));
 $('cloudSaveRun').addEventListener('click',()=>saveRun().catch(e=>{feedback(e.message,false);note('ERROR '+e.message)}));
 db.auth.onAuthStateChange((event,session)=>{setTimeout(()=>setUser(session?.user||null),0)});
 try{const {data,error}=await db.auth.getUser();if(error&&/invalid|refresh/i.test(error.message))note('Session expired');await setUser(data?.user||null)}
 catch(e){feedback('Supabase เชื่อมต่อขัดข้อง: '+e.message,false)}
}
window.AliceCloud={isAuthenticated:signedIn,saveRun,saveTemplates,ensureTemplates,refresh};
document.addEventListener('DOMContentLoaded',init);
})();
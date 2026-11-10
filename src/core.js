(function(){
'use strict';
const NS={r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',p:'http://schemas.openxmlformats.org/package/2006/relationships',s:'http://schemas.openxmlformats.org/spreadsheetml/2006/main',a:'http://schemas.openxmlformats.org/drawingml/2006/main',d:'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',ct:'http://schemas.openxmlformats.org/package/2006/content-types'};
const l=(node,name)=>Array.from(node?.childNodes||[]).filter(x=>x.nodeType===1&&x.localName===name);
const one=(node,name)=>l(node,name)[0]||null;
const find=(node,name)=>Array.from(node?.getElementsByTagName('*')||[]).find(x=>x.localName===name)||null;
const all=(node,name)=>Array.from(node?.getElementsByTagName('*')||[]).filter(x=>x.localName===name);
function xml(str){const doc=new DOMParser().parseFromString(str,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw Error('ไฟล์ XML ภายใน Excel ผิดรูปแบบ');return doc}
const serialize=doc=>new XMLSerializer().serializeToString(doc);
function join(base,path){if(path.startsWith('/'))return path.slice(1);const st=(base.slice(0,base.lastIndexOf('/')+1)+path).split('/');const out=[];st.forEach(p=>{if(p==='..')out.pop();else if(p&&p!=='.')out.push(p)});return out.join('/')}
const relPath=p=>p.slice(0,p.lastIndexOf('/')+1)+'_rels/'+p.slice(p.lastIndexOf('/')+1)+'.rels';
const attr=(n,name)=>n?.getAttribute(name)||'';
function colNumber(name){let v=0;for(const ch of name.toUpperCase())v=v*26+ch.charCodeAt(0)-64;return v}
function excelDate(ymd){const [y,m,d]=ymd.split('-').map(Number);return Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/86400000)}
const clean=v=>String(v??'').trim();
function csvParse(text){
 text=String(text).replace(/^\uFEFF/,'');
 const rows=[];let row=[],field='',quoted=false;
 for(let i=0;i<text.length;i++){
  const c=text[i];
  if(c==='"'){if(quoted&&text[i+1]==='"'){field+='"';i++}else quoted=!quoted}
  else if(c===','&&!quoted){row.push(field);field=''}
  else if((c==='\n'||c==='\r')&&!quoted){if(c==='\r'&&text[i+1]==='\n')i++;row.push(field);if(row.some(x=>x!==''))rows.push(row);row=[];field=''}
  else field+=c;
 }
 row.push(field);if(row.some(x=>x!==''))rows.push(row);
 const headers=rows.shift()||[];
 return rows.map(r=>Object.fromEntries(headers.map((h,i)=>[h.trim(),r[i]??''])));
}
function field(record,key){const k=Object.keys(record).find(x=>x.trim().toLowerCase()===key.trim().toLowerCase());return k?record[k]:''}
function cancelled(status){return /cancel|แคนเซิล|ยกเลิก/i.test(clean(status))}
function dateFcl(text,iso){
 const [y,m,d]=iso.split('-');const day=String(Number(d));const month=String(Number(m));const year=String(y);
 const matches=String(text||'').match(/\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{4}\s*[-–]\s*\d+\s*fcl/ig)||[];
 let total=0;
 for(const v of matches){const m=v.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\s*[-–]\s*(\d+)\s*fcl/i);if(m&&Number(m[1])===Number(day)&&Number(m[2])===Number(month)&&m[3]===year)total+=Number(m[4])}
 return total;
}
function parseProducts(raw){return clean(raw).split(/\r?\n+/).map(x=>x.trim()).filter(Boolean)}
function classify(file){
 const n=(file.name||'').toLowerCase();
 if(n.endsWith('.csv')){if(/stp\s*1|stp1/.test(n))return 'planSTP';if(/stp\s*2|stp2/.test(n))return 'planPTR';return 'planSTP'}
 if(/(?:แผนส่งออก|batch\s*allocat|stock\s*allocat|rev\s*\d+\s*แผน)/i.test(n)&&/\.xlsx?$/.test(n))return 'planBatch';
 if(/ssams/i.test(n)&&/\.xlsx?$/.test(n))return 'planDomestic';
 if(/domestic\s*report|domestic_report/i.test(n)&&/\.xlsx?$/.test(n))return 'loadingDomestic';
 if(n.includes('export')||n.includes('ส่งออก'))return n.includes('ptr')?'exportPTR':'exportSTP';
 if(n.includes('loading report'))return n.includes('ptr')?'loadingPTR':'loadingSTP';
 if(n.includes('ptr')&&(/ใบ|แจ้ง|weight|card/i.test(n)))return 'formPTR';
 if(n.includes('stp')&&(/ใบ|แจ้ง|weight|card/i.test(n)))return 'formSTP';
 if(/\.xlsx?$/.test(n))return 'truck';
 return 'ignore';
}
function quantityEstimate(r,totalFcl){
 const pack=clean(field(r,'Packing Type'));
 const qty=String(field(r,'Quantity (MT)')||'').split(/\r?\n/).map(Number).filter(n=>Number.isFinite(n)&&n>0).reduce((s,n)=>s+n,0);
 if(/850\s*kg|jumbo/i.test(pack))return {bags:20,weight:17000,confidence:'packing'};
 if(/25\s*kg|SPS-22|SPS-29|brown/i.test(pack)){
  if(qty>0&&totalFcl>0&&Math.abs(qty*1000/totalFcl-20000)<2)return {bags:800,weight:20000,confidence:'total'};
  if(qty>0&&totalFcl===1&&Math.round(qty*1000/25)===qty*1000/25)return {bags:Math.round(qty*1000/25),weight:qty*1000,confidence:'single'};
  return {bags:null,weight:null,confidence:'review'};
 }
 return {bags:null,weight:null,confidence:'review'};
}
function planRecords(rows,date,kind){
 const jobs=[],ignored=[];
 for(const r of rows){
  const n=dateFcl(field(r,'Date Load at factory (R1)'),date);
  if(!n)continue;
  const booking=clean(field(r,'Booking')),status=clean(field(r,'KRT Plan Status')),load=clean(field(r,'Load'));
  if(cancelled(status)||load.toUpperCase()!=='KNS'){ignored.push({booking,status,load,reason:cancelled(status)?'Cancel':'ไม่ใช่ KNS'});continue}
  const product=clean(field(r,'Product')).replace(/\r?\n/g,' / ');
  const etdraw=clean(field(r,'ETD'));
  const etd=(etdraw.match(/(?:ETD\s*:?\s*)?([A-Za-z]+\s*[.,]?\s*\d{1,2}\s*[.,]\s*\d{4})/i)||[])[1]||etdraw;
  const dest=etdraw.split('/').map(x=>x.trim()).filter(Boolean).slice(-1)[0]||'';
  const fclTxt=String(field(r,'Date Load at factory (R1)'));
  const allCounts=(fclTxt.match(/[-–]\s*(\d+)\s*fcl/ig)||[]).reduce((a,z)=>a+(Number((z.match(/\d+/)||[])[0])||0),0);
  const estimate=quantityEstimate(r,allCounts||n);
  jobs.push({kind,booking,product,contract:clean(field(r,'TPS Contract no.')).replace(/\r?\n/g,' / '),etd:etd.replace(/\s+/g,' '),destination:dest,agent:clean(field(r,'Agent')),place:clean(field(r,'Place of Delivery')),closing:clean(field(r,'Closing (Date)')),closingTime:clean(field(r,'Closing (Time)')),packaging:String(field(r,'Packaging and Batch no.')||''),packingType:clean(field(r,'Packing Type')),quantityMT:clean(field(r,'Quantity (MT)')),fcl:n,bags:estimate.bags,kg:estimate.weight,confidence:estimate.confidence,raw:r});
 }
 return {jobs,ignored};
}
function fileArrayBuffer(file){return typeof file.arrayBuffer==='function'?file.arrayBuffer():new Response(file).arrayBuffer()}
class ExcelPatch {
 constructor(zip){this.zip=zip;this.docs=new Map();this.sheets=[]}
 static async load(file){
  if(!window.JSZip)throw Error('โหลดไลบรารี JSZip ไม่สำเร็จ กรุณาตรวจอินเทอร์เน็ต');
  const zip=await JSZip.loadAsync(await fileArrayBuffer(file));
  const x=new ExcelPatch(zip);await x.init();return x;
 }
 async doc(path){if(this.docs.has(path))return this.docs.get(path);const f=this.zip.file(path);if(!f)throw Error('ไม่พบส่วนประกอบ '+path);const d=xml(await f.async('string'));this.docs.set(path,d);return d}
 async init(){
  this.book=await this.doc('xl/workbook.xml');
  this.bookRels=await this.doc('xl/_rels/workbook.xml.rels');
  this.contentTypes=await this.doc('[Content_Types].xml');
  const relationships=l(this.bookRels.documentElement,'Relationship');
  this.relMap=Object.fromEntries(relationships.map(v=>[attr(v,'Id'),attr(v,'Target')]));
  const list=one(this.book.documentElement,'sheets');
  this.sheets=l(list,'sheet').map(s=>({node:s,name:attr(s,'name'),id:Number(attr(s,'sheetId')),rid:s.getAttributeNS(NS.r,'id')||s.getAttribute('r:id')}));
  for(const sheet of this.sheets){sheet.path=join('xl/workbook.xml',this.relMap[sheet.rid]);sheet.row=null}
 }
 sheetByName(name){return this.sheets.find(s=>s.name===name)}
 async worksheet(sheet){return this.doc(sheet.path)}
 async cell(sheet,address,create=true){
  const d=await this.worksheet(sheet),sdata=find(d.documentElement,'sheetData');if(!sdata)throw Error('ไม่มี sheetData');
  const match=/^([A-Z]+)(\d+)$/.exec(address);if(!match)throw Error('เซลล์ไม่ถูกต้อง '+address);
  const ri=Number(match[2]),ci=colNumber(match[1]);let row=l(sdata,'row').find(r=>Number(attr(r,'r'))===ri);
  if(!row){if(!create)return null;row=d.createElementNS(NS.s,'row');row.setAttribute('r',ri);const before=l(sdata,'row').find(r=>Number(attr(r,'r'))>ri);sdata.insertBefore(row,before||null)}
  let cell=l(row,'c').find(c=>attr(c,'r')===address);
  if(!cell){if(!create)return null;cell=d.createElementNS(NS.s,'c');cell.setAttribute('r',address);const before=l(row,'c').find(c=>colNumber((attr(c,'r').match(/^[A-Z]+/)||[])[0])>ci);row.insertBefore(cell,before||null)}
  return cell;
 }
 async setCell(sheet,address,value){
  const c=await this.cell(sheet,address,true),d=c.ownerDocument;
  l(c,'f').concat(l(c,'v'),l(c,'is')).forEach(n=>n.remove());
  if(value===null||value===undefined||value===''){c.removeAttribute('t');return}
  if(typeof value==='number'&&Number.isFinite(value)){c.setAttribute('t','n');const n=d.createElementNS(NS.s,'v');n.textContent=String(value);c.appendChild(n)}
  else{c.setAttribute('t','inlineStr');const is=d.createElementNS(NS.s,'is'),t=d.createElementNS(NS.s,'t');t.setAttribute('xml:space','preserve');t.textContent=String(value);is.appendChild(t);c.appendChild(is)}
 }
 async clearCell(sheet,address){const c=await this.cell(sheet,address,false);if(c){l(c,'f').concat(l(c,'v'),l(c,'is')).forEach(n=>n.remove());c.removeAttribute('t')}}
 async shapeAnchors(sheet){
  if(sheet.anchors)return sheet.anchors;
  const sdoc=await this.worksheet(sheet);
  const drawing=one(sdoc.documentElement,'drawing');if(!drawing)throw Error('แม่แบบ '+sheet.name+' ไม่มี Drawing / Text Box');
  const rid=drawing.getAttributeNS(NS.r,'id')||drawing.getAttribute('r:id');
  const rp=relPath(sheet.path),rels=await this.doc(rp);
  const rel=l(rels.documentElement,'Relationship').find(r=>attr(r,'Id')===rid);
  if(!rel)throw Error('ไม่พบ drawing relationship ใน '+sheet.name);
  sheet.drawingPath=join(sheet.path,attr(rel,'Target'));
  sheet.drawingRelsPath=relPath(sheet.drawingPath);
  sheet.anchors=l((await this.doc(sheet.drawingPath)).documentElement,'oneCellAnchor').concat([]);
  // Shape order is the actual XML child order, matching Excel's Shapes collection.
  const doc=await this.doc(sheet.drawingPath);
  sheet.anchors=Array.from(doc.documentElement.childNodes).filter(n=>n.nodeType===1&&/CellAnchor$/.test(n.localName));
  return sheet.anchors;
 }
 async setShape(sheet,index,value,size=9){
  const anchors=await this.shapeAnchors(sheet),anchor=anchors[index-1];if(!anchor)throw Error('ไม่พบ Text Box ลำดับ '+index+' ใน '+sheet.name);
  const sp=one(anchor,'sp');if(!sp)throw Error('วัตถุลำดับ '+index+' ไม่ใช่กล่องข้อความ');
  const body=one(sp,'txBody');if(!body)throw Error('Text Box ไม่มี txBody');
  const doc=body.ownerDocument,contents=String(value??'');
  const firstRun=find(body,'rPr'),sample=firstRun?firstRun.cloneNode(true):null;
  // The supplied STP/PTR templates mix spAutoFit (grows a shape over its
  // neighbors) and noAutofit (clips text). Neither is safe for variable data.
  // Text-to-fit keeps the ORIGINAL textbox dimensions and the whole string.
  const bodyPr=one(body,'bodyPr');
  if(bodyPr){
   for(const option of ['spAutoFit','noAutofit','normAutofit'])
    l(bodyPr,option).forEach(n=>n.remove());
   bodyPr.appendChild(doc.createElementNS(NS.a,'a:normAutofit'));
   bodyPr.setAttribute('wrap',contents.includes('\n')?'square':'none');
   bodyPr.setAttribute('vertOverflow','clip');
   bodyPr.setAttribute('horzOverflow','clip');
   // Consistent insets for both halves of the form. The template has
   // asymmetric margins which previously consumed usable textbox space.
   bodyPr.setAttribute('lIns','12700');
   bodyPr.setAttribute('rIns','12700');
   bodyPr.setAttribute('tIns','0');
   bodyPr.setAttribute('bIns','0');
  }
  l(body,'p').forEach(p=>p.remove());
  const uniformSize=Math.round(Math.max(8,Math.min(11,Number(size)||9))*100);
  contents.split(/\r\n|\r|\n/).forEach(line=>{
   const p=doc.createElementNS(NS.a,'a:p'),r=doc.createElementNS(NS.a,'a:r');
   const rp=sample?doc.importNode(sample,true):doc.createElementNS(NS.a,'a:rPr');
   rp.setAttribute('lang','th-TH');rp.setAttribute('sz',String(uniformSize));
   // Force one Thai/Latin-capable face; cloned template runs can otherwise
   // mix theme fonts and different fallback faces in the same paragraph.
   for(const tag of ['latin','ea','cs']){
    let font=one(rp,tag);
    if(!font){font=doc.createElementNS(NS.a,'a:'+tag);rp.appendChild(font)}
    font.setAttribute('typeface','Tahoma');
   }
   r.appendChild(rp);
   const t=doc.createElementNS(NS.a,'a:t');t.textContent=line;t.setAttribute('xml:space','preserve');
   r.appendChild(t);p.appendChild(r);body.appendChild(p);
  });
  const cNvPr=find(sp,'cNvPr');if(!cNvPr)throw Error('ไม่พบข้อมูลรูปทรง');
 }
 async shapeText(sheet,index){const anchors=await this.shapeAnchors(sheet);const sp=one(anchors[index-1],'sp');return all(one(sp,'txBody'),'t').map(n=>n.textContent).join('\n')}
 async rename(sheet,name){
  if(name.length>31||/[\[\]:*?\/\\]/.test(name))throw Error('ชื่อชีตไม่ถูกต้อง '+name);
  if(this.sheets.some(s=>s!==sheet&&s.name.toLowerCase()===name.toLowerCase()))throw Error('ชื่อชีตซ้ำ: '+name);
  sheet.name=name;sheet.node.setAttribute('name',name);
 }
 async trimSheets(keep){
  if(!Number.isInteger(keep)||keep<1||keep>this.sheets.length)
   throw Error('จำนวนชีตที่ต้องเก็บไม่ถูกต้อง: '+keep);
  // Templates often have a fixed set of 10 blank form sheets. Keep only the
  // pages required by today's cars in the OUTPUT; never edit the source file.
  while(this.sheets.length>keep){
   const sheet=this.sheets.pop();
   sheet.node.remove();
   const relation=l(this.bookRels.documentElement,'Relationship').find(r=>attr(r,'Id')===sheet.rid);
   if(relation)relation.remove();
   this.zip.remove(sheet.path);
   this.docs.delete(sheet.path);
   const relationships=relPath(sheet.path);
   this.zip.remove(relationships);
   this.docs.delete(relationships);
   for(const override of l(this.contentTypes.documentElement,'Override')){
    if(attr(override,'PartName')==='/'+sheet.path)override.remove();
   }
  }
  // Defined names such as print areas use zero-based sheet indexes.
  const names=one(this.book.documentElement,'definedNames');
  if(names){
   for(const node of l(names,'definedName')){
    const idx=Number(attr(node,'localSheetId'));
    if(node.hasAttribute('localSheetId')&&Number.isInteger(idx)&&idx>=keep)node.remove();
   }
   if(!l(names,'definedName').length)names.remove();
  }
  const views=one(this.book.documentElement,'bookViews');
  if(views)for(const view of l(views,'workbookView')){
   const selected=Number(attr(view,'activeTab'));
   if(selected>=keep)view.setAttribute('activeTab','0');
   const first=Number(attr(view,'firstSheet'));
   if(first>=keep)view.setAttribute('firstSheet','0');
  }
 }
 async cloneSheet(source){
  const old=source.path;let nr=1;
  while(this.zip.file('xl/worksheets/sheet'+nr+'.xml'))nr++;
  const newPath='xl/worksheets/sheet'+nr+'.xml';
  const sourceContent=await this.zip.file(old).async('uint8array');
  this.zip.file(newPath,sourceContent);
  const sourceRels=relPath(old),destRels=relPath(newPath);
  if(this.zip.file(sourceRels)){
   const sourceRelDoc=await this.doc(sourceRels);
   const newRels=xml(serialize(sourceRelDoc));
   const drawingRel=l(newRels.documentElement,'Relationship').find(r=>/\/drawing$/.test(attr(r,'Type')));
   if(drawingRel){
    const oldDraw=join(old,attr(drawingRel,'Target'));
    let no=1;while(this.zip.file('xl/drawings/drawing'+no+'.xml'))no++;
    const newDraw='xl/drawings/drawing'+no+'.xml';
    this.zip.file(newDraw,await this.zip.file(oldDraw).async('uint8array'));
    const oldDrRel=relPath(oldDraw),newDrRel=relPath(newDraw);
    if(this.zip.file(oldDrRel))this.zip.file(newDrRel,await this.zip.file(oldDrRel).async('uint8array'));
    drawingRel.setAttribute('Target','../drawings/drawing'+no+'.xml');
    this.addOverride('/'+newDraw,'application/vnd.openxmlformats-officedocument.drawing+xml');
   }
   this.docs.set(destRels,newRels);
  }
  const maxSheet=Math.max(0,...this.sheets.map(s=>s.id))+1;
  const maxRid=Math.max(0,...l(this.bookRels.documentElement,'Relationship').map(e=>Number((attr(e,'Id').match(/\d+$/)||[])[0])||0))+1;
  const rid='rId'+maxRid;
  const newNode=this.book.createElementNS(NS.s,'sheet');
  newNode.setAttribute('name','Page '+maxSheet);
  newNode.setAttribute('sheetId',String(maxSheet));
  newNode.setAttributeNS(NS.r,'r:id',rid);
  one(this.book.documentElement,'sheets').appendChild(newNode);
  const rel=this.bookRels.createElementNS(NS.p,'Relationship');
  rel.setAttribute('Id',rid);
  rel.setAttribute('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
  rel.setAttribute('Target','worksheets/sheet'+nr+'.xml');
  this.bookRels.documentElement.appendChild(rel);
  const sheet={node:newNode,name:'Page '+maxSheet,id:maxSheet,rid,path:newPath};
  this.sheets.push(sheet);
  this.addOverride('/'+newPath,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  return sheet;
 }
 addOverride(name,type){
  const root=this.contentTypes.documentElement;
  if(l(root,'Override').some(e=>attr(e,'PartName')===name))return;
  const o=this.contentTypes.createElementNS(NS.ct,'Override');o.setAttribute('PartName',name);o.setAttribute('ContentType',type);root.appendChild(o);
 }
 async save(filename){
  // Excel's calcChain lists formula cells by address. If Loading/Export data are
  // synchronized, cached formulas can become stale. Remove the chain entirely,
  // including its relationship and content-type entry; Excel rebuilds it safely.
  if(this.zip.file('xl/calcChain.xml')){
   this.zip.remove('xl/calcChain.xml');
   for(const rel of l(this.bookRels.documentElement,'Relationship')){
    if(attr(rel,'Type').endsWith('/calcChain'))rel.remove();
   }
   for(const override of l(this.contentTypes.documentElement,'Override')){
    if(attr(override,'PartName')==='/xl/calcChain.xml')override.remove();
   }
   const cp=one(this.book.documentElement,'calcPr');
   if(cp){cp.setAttribute('fullCalcOnLoad','1');cp.setAttribute('forceFullCalc','1')}
  }
  for(const [path,d] of this.docs)this.zip.file(path,serialize(d));
  const blob=await this.zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
  return {filename,blob};
 }
}
function saveBlob(file){const url=URL.createObjectURL(file.blob);const a=document.createElement('a');a.href=url;a.download=file.filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000)}
function parseXlsx(file){if(!window.XLSX)throw Error('ไลบรารีอ่าน Excel ไม่พร้อม');return fileArrayBuffer(file).then(buf=>XLSX.read(buf,{type:'array',cellDates:false,raw:true}))}
function sheetRows(wb,name){const s=wb.Sheets[name];if(!s)throw Error('ไม่พบชีต '+name);const last=XLSX.utils.decode_range(s['!ref']||'A1').e;return XLSX.utils.sheet_to_json(s,{header:1,raw:true,defval:'',blankrows:true,range:{s:{r:0,c:0},e:last}})}
function textAt(r,col){return clean((r||[])[col])}
window.AliceCore={NS,l,one,find,all,xml,join,relPath,ExcelPatch,saveBlob,csvParse,field,cancelled,dateFcl,classify,quantityEstimate,parseProducts,planRecords,excelDate,parseXlsx,sheetRows,textAt,clean};
})();
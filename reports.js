/* Receipt reports: pure snapshot shared by PDF, workbook and UI. */
(function(root){
'use strict';
const compare=(a,b)=>String(a??'')<String(b??'')?-1:String(a??'')>String(b??'')?1:0;
function cents(value){
  if(value===null||value===undefined||! /^-?\d+(\.\d{1,2})?$/.test(String(value)))throw new Error('Invalid receipt amount; correct the record before exporting.');
  const n=Math.round(Number(value)*100);if(!Number.isSafeInteger(n))throw new Error('Receipt amount is too large.');return n;
}
function snapshot(input,filters={}){
 const f={month:'',company:'',status:'',search:'',includeExcluded:false,...filters};
 const matching=input.filter(r=>(!f.month||String(r.receipt_date||'').startsWith(f.month))&&(!f.company||r.company===f.company)&&(!f.status||r.status===f.status)&&(!f.search||`${r.vendor||''}\n${r.comment||''}`.toLowerCase().includes(f.search.trim().toLowerCase())));
 const rows=matching.filter(r=>f.includeExcluded||r.status!=='excluded').map(r=>({...r})).sort((a,b)=>compare(a.receipt_date,b.receipt_date)||compare(a.vendor,b.vendor)||compare(a.id,b.id));
 const sums=new Map();
 for(const r of rows){const currency=r.currency||'UNKNOWN';const n=cents(r.amount);if(!sums.has(currency))sums.set(currency,{currency,count:0,total:0,reimbursable:0,excludedCount:0,excluded:0});const t=sums.get(currency);if(r.status==='excluded'){t.excludedCount++;t.excluded+=n;}else{t.count++;t.total+=n;if(r.reimbursable)t.reimbursable+=n;}if(!Number.isSafeInteger(t.total)||!Number.isSafeInteger(t.excluded)||!Number.isSafeInteger(t.reimbursable))throw new Error('Report total is too large.');}
 const totals=[...sums.values()].sort((a,b)=>compare(a.currency,b.currency)).map(t=>({...t,total:t.total/100,reimbursable:t.reimbursable/100,excluded:t.excluded/100}));
 return {rows,totals,filters:f,omitted:matching.length-rows.length,generated:new Date().toISOString()};
}
function description(s){const f=s.filters;return [`Month: ${f.month||'All months'}`,`Company: ${f.company||'All companies'}`,`Status: ${f.status||'All statuses'}`,`Search: ${f.search||'(none)'}`,`Excluded: ${f.includeExcluded?'included as evidence only; separate excluded totals':'omitted'} (${s.omitted} omitted)`,`Records: ${s.rows.length} | Generated: ${s.generated}`];}
function nonempty(s){if(!s.rows.length)throw new Error('No receipts match this report. Change the filters or excluded option.');}
async function workbook(s){
 nonempty(s);if(!root.ExcelJS)throw new Error('Excel export library did not load. Reload and retry.');
 const wb=new root.ExcelJS.Workbook();wb.creator='360 Receipts';wb.created=new Date(s.generated);
 const summary=wb.addWorksheet('Summary');
 summary.addRow(['360 Receipts — verification report']);description(s).forEach(line=>summary.addRow([line]));
 summary.addRow(['Signed amounts preserved. Excluded evidence never contributes to report or reimbursable totals. No cross-currency grand total.']);
 summary.addRow([]);summary.addRow(['Currency','Included count','Report total','Reimbursable total','Excluded count','Excluded total (not counted)']);
 s.totals.forEach(t=>summary.addRow([t.currency,t.count,t.total,t.reimbursable,t.excludedCount,t.excluded]));
 summary.columns=[{width:24},{width:20},{width:20},{width:24},{width:20},{width:32}];
 for(let r=1;r<=8;r++){summary.mergeCells(r,1,r,6);summary.getRow(r).height=18*Math.max(1,Math.ceil(String(summary.getCell(r,1).value||'').length/115));}
 summary.getRow(10).font={bold:true,color:{argb:'FF1F5AB8'}};summary.getRow(10).height=32;
 const ws=wb.addWorksheet('Receipts',{views:[{state:'frozen',ySplit:1}]});
 const fields=[['Date','receipt_date',14],['Name','vendor',30],['Price','amount',18],['Currency','currency',12],['Category Name','company',28],['Category Code','category_code',18],['Comment','comment',50],['Reimbursable','reimbursable',16],['Receipt ID','id',38],['Ledger ID','ledger_receipt_id',24],['Status','status',18],['Counted in report','counted',20],['Source','source',14],['Original file','original_name',36],['Source path','storage_path',50],['SHA-256','sha256',68]];
 ws.columns=fields.map(([header,key,width])=>({header,key,width}));
 for(const r of s.rows){const record={};for(const [,key] of fields)record[key]=key==='receipt_date'?new Date(`${r[key]}T00:00:00Z`):key==='amount'?cents(r.amount)/100:key==='counted'?r.status!=='excluded':key==='reimbursable'?(r[key]?'Yes':'No'):String(r[key]??'');ws.addRow(record);}
 ws.getColumn('receipt_date').numFmt='m/d/yy';
 ws.autoFilter={from:'A1',to:{row:ws.rowCount,column:fields.length}};ws.getColumn('amount').numFmt='#,##0.00;[Red]-#,##0.00';
 for(const sh of [summary,ws]){sh.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};sh.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F5AB8'}};sh.eachRow(row=>{row.alignment={vertical:'top',wrapText:true};});}
 for(let r=11;r<=summary.rowCount;r++)for(const c of [3,4,6])summary.getCell(r,c).numFmt='#,##0.00;[Red]-#,##0.00';
 return wb.xlsx.writeBuffer();
}
// Smart Receipts precedent: A4 summary, six columns, four evidence panels.
// Browser fonts preserve Unicode; PDF evidence remains vector and all pages survive.
const W=595.28,H=841.89,M=32,BLUE='#0080ff';
const shortDate=v=>{const [y,m,d]=v.split('-');return `${Number(m)}/${Number(d)}/${y.slice(2)}`;};
const money=v=>Number(v).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
function surface(){const c=document.createElement('canvas');c.width=Math.round(W*2);c.height=Math.round(H*2);const x=c.getContext('2d');x.scale(2,2);x.fillStyle='white';x.fillRect(0,0,W,H);x.fillStyle=BLUE;x.fillRect(M,M,W-2*M,4);x.fillRect(M,H-32,W-2*M,1.5);x.fillStyle='black';x.font='italic 8px Arial';x.fillText('Report Generated using Receipt Tracker',M,H-19);return {c,x};}
function wrap(x,text,width){let lines=[''];for(const ch of String(text??'')){let i=lines.length-1;if(x.measureText(lines[i]+ch).width>width&&lines[i])lines.push(ch);else lines[i]+=ch;}return lines;}
async function addSurface(doc,c){const p=doc.addPage([W,H]);p.drawImage(await doc.embedPng(c.toDataURL()),{x:0,y:0,width:W,height:H});return p;}
async function summaryPages(doc,s){
 let c,x,y;const widths=[68,96,70,74,133,90];
 const headers=['Date','Name','Price','Currency','Category Name','Reimbursable'];
 const start=()=>{({c,x}=surface());x.fillStyle='black';x.font='bold 12px Arial';const dates=s.rows.map(r=>r.receipt_date);const months=[...new Set(dates.map(d=>d.slice(0,7)))];const month=s.filters.month||(months.length===1?months[0]:'');x.fillText(month?new Date(month+'-02T00:00:00Z').toLocaleString('en-US',{month:'long',timeZone:'UTC'}):'Receipts',M,64);x.font='10px Arial';const from=month?month+'-01':dates[0],to=month?new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).toISOString().slice(0,10):dates[dates.length-1];x.fillText(`From: ${shortDate(from)} To: ${shortDate(to)}`,M,79);y=93;x.font='bold 10px Arial';for(const t of s.totals){x.fillText(`${s.totals.length===1?'Grand Total':'Total'}: ${t.currency} ${money(t.total)}`,M,y);y+=13;}y=Math.max(y+28,133);};
 const row=async(values,fill,header=false)=>{x.font=`${header?'bold ':''}10px Arial`;const lines=values.map((v,i)=>wrap(x,v,widths[i]-8));const h=Math.max(22,...lines.map(a=>a.length*12+8));if(y+h>H-55){await addSurface(doc,c);start();await row(headers,'#cce3ff',true);}x.fillStyle=fill;x.fillRect(M,y,W-2*M,h);x.fillStyle=header?BLUE:'black';let left=M;lines.forEach((ls,i)=>{ls.forEach((t,j)=>x.fillText(t,left+(widths[i]-x.measureText(t).width)/2,y+14+j*12));left+=widths[i];});y+=h;};
 start();await row(headers,'#cce3ff',true);
 for(let i=0;i<s.rows.length;i++){const r=s.rows[i];await row([shortDate(r.receipt_date),r.vendor+(r.status==='excluded'?' [EXCLUDED]':''),money(r.amount),r.currency,r.company,r.reimbursable?'Yes':'No'],i%2?'white':'#eeeef3');}
 for(const t of s.totals)await row(['','',money(t.total),t.currency,'',''],'#cce3ff',true);
 // Filters and exclusions are explicit without replacing the historical table.
 x.font='8px Arial';for(const line of description(s).slice(0,5)){for(const part of wrap(x,line,W-2*M)){if(y+14>H-48){await addSurface(doc,c);start();}x.fillStyle='#444';x.fillText(part,M,y+16);y+=11;}}
 if(s.filters.includeExcluded)for(const t of s.totals){x.fillText(`Excluded (not counted): ${t.currency} ${money(t.excluded)}`,M,y+16);y+=12;}
 await addSurface(doc,c);
}
async function pdf(s,loadSource,progress=()=>{}){
 nonempty(s);if(!root.PDFLib)throw new Error('PDF export library did not load. Reload and retry.');
 const doc=await root.PDFLib.PDFDocument.create();doc.setTitle('Receipts report');doc.setProducer('Receipt Tracker 1.2');await summaryPages(doc,s);
 let sheet=null,slot=0,overlay=null;
 async function panel(asset,r,number,part,total,isPdf=false){
  if(slot%4===0){if(overlay){const image=await doc.embedPng(overlay.c.toDataURL());sheet.drawImage(image,{x:0,y:0,width:W,height:H});}sheet=await addSurface(doc,surface().c);overlay=surface();overlay.x.clearRect(0,0,W,H);}
  const idx=slot%4,col=idx%2,line=Math.floor(idx/2),left=M+col*270,top=51+line*377,bw=261,bh=344;
  const x=overlay.x;x.fillStyle='black';x.font='8px Arial';const label=`${number} • ${r.vendor} • ${shortDate(r.receipt_date)}${total>1?` • ${part}/${total}`:''}${r.status==='excluded'?' • EXCLUDED':''}`;const lines=wrap(x,label,bw);lines.forEach((t,i)=>x.fillText(t,left+(bw-x.measureText(t).width)/2,top+8+i*10));
  const h=bh-(lines.length-1)*10,fit=Math.min(bw/asset.width,h/asset.height);const opts={x:left+(bw-asset.width*fit)/2,y:H-top-20-(lines.length-1)*10-asset.height*fit,width:asset.width*fit,height:asset.height*fit};
  if(isPdf)sheet.drawPage(asset,opts);else sheet.drawImage(asset,opts);slot++;
 }
 for(let i=0;i<s.rows.length;i++){
  const r=s.rows[i];progress(i+1,s.rows.length);if(!r.storage_path)throw new Error(`Receipt ${r.id}: no source file. PDF not created.`);
  let blob;try{blob=await loadSource(r);}catch{throw new Error(`Receipt ${r.id}: source unavailable. No partial PDF was downloaded.`);}if(!blob?.size)throw new Error(`Receipt ${r.id}: empty source file. PDF not created.`);
  try{const bytes=await blob.arrayBuffer();
   if(r.mime_type==='application/pdf'||/\.pdf$/i.test(r.storage_path)){const src=await root.PDFLib.PDFDocument.load(bytes);if(!src.getPageCount())throw Error('Empty PDF');for(let j=0;j<src.getPageCount();j++){const [embedded]=await doc.embedPdf(src,[j]);await panel(embedded,r,i+1,j+1,src.getPageCount(),true);}}
   else{const bitmap=await createImageBitmap(blob);try{const c=document.createElement('canvas');c.width=bitmap.width;c.height=bitmap.height;c.getContext('2d').drawImage(bitmap,0,0);const asset=await doc.embedPng(c.toDataURL());await panel(asset,r,i+1,1,1);}finally{bitmap.close();}}
   // Full original bytes remain available even when the historical compact panel is small.
   await doc.attach(bytes,`${i+1}-${(r.original_name||r.storage_path.split('/').pop()).replace(/[\\/]/g,'_')}`,{mimeType:r.mime_type,description:`Original receipt ${i+1}`});
  }catch{throw new Error(`Receipt ${r.id}: unreadable, encrypted or unsupported source. Convert HEIC/HEIF to JPEG/PNG or unlock the PDF. No partial PDF was downloaded.`);}
 }
 if(overlay){const image=await doc.embedPng(overlay.c.toDataURL());sheet.drawImage(image,{x:0,y:0,width:W,height:H});}return doc.save();
}
const api={snapshot,description,workbook,pdf};root.ReceiptReports=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

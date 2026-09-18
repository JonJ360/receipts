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
 const fields=[['Receipt ID','id',38],['Ledger ID','ledger_receipt_id',24],['Date','receipt_date',14],['Vendor','vendor',30],['Amount','amount',18],['Currency','currency',12],['Company / reimbursement','company',28],['Category code','category_code',18],['Purpose / comment','comment',50],['Reimbursable','reimbursable',16],['Status','status',18],['Counted in report','counted',20],['Source','source',14],['Original file','original_name',36],['Source path','storage_path',50],['SHA-256','sha256',68]];
 ws.columns=fields.map(([header,key,width])=>({header,key,width}));
 for(const r of s.rows){const record={};for(const [,key] of fields)record[key]=key==='amount'?cents(r.amount)/100:key==='counted'?r.status!=='excluded':key==='reimbursable'?!!r[key]:String(r[key]??'');ws.addRow(record);}
 ws.autoFilter={from:'A1',to:{row:ws.rowCount,column:fields.length}};ws.getColumn('amount').numFmt='#,##0.00;[Red]-#,##0.00';
 for(const sh of [summary,ws]){sh.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};sh.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F5AB8'}};sh.eachRow(row=>{row.alignment={vertical:'top',wrapText:true};});}
 for(let r=11;r<=summary.rowCount;r++)for(const c of [3,4,6])summary.getCell(r,c).numFmt='#,##0.00;[Red]-#,##0.00';
 return wb.xlsx.writeBuffer();
}
// Render metadata with browser fonts: Unicode survives, with no HTML interpretation.
// Evidence PDFs are copied as original pages; images are embedded, never linked.
async function textPages(doc,lines){
 const width=1224,height=1584,pad=72,lineHeight=34;let canvas,ctx,y;
 function start(){canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,width,height);ctx.fillStyle='#222A33';ctx.font='24px Arial';y=pad;}
 async function flush(){const image=await doc.embedPng(canvas.toDataURL());const page=doc.addPage([612,792]);page.drawImage(image,{x:0,y:0,width:612,height:792});}
 start();for(const line of lines){let part='';for(const char of String(line)){if(ctx.measureText(part+char).width>width-2*pad){ctx.fillText(part,pad,y);y+=lineHeight;part='';if(y>height-pad){await flush();start();}}part+=char;}ctx.fillText(part,pad,y);y+=lineHeight;if(y>height-pad){await flush();start();}}
 if(y>pad)await flush();
}
async function imagePages(doc,blob){
 let bitmap;try{bitmap=await createImageBitmap(blob);}catch{throw new Error('Unsupported or damaged image. Convert HEIC/HEIF to JPEG or PNG before exporting.');}
 try{const scale=Math.min(1,1080/bitmap.width);const sliceHeight=Math.max(1,Math.floor(1440/scale));
 for(let top=0;top<bitmap.height;top+=sliceHeight){const h=Math.min(sliceHeight,bitmap.height-top);const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(bitmap.width*scale));canvas.height=Math.max(1,Math.round(h*scale));const ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(bitmap,0,top,bitmap.width,h,0,0,canvas.width,canvas.height);const img=await doc.embedPng(canvas.toDataURL());const page=doc.addPage([612,792]);const fit=Math.min(540/img.width,720/img.height);page.drawImage(img,{x:(612-img.width*fit)/2,y:756-img.height*fit,width:img.width*fit,height:img.height*fit});}
 }finally{bitmap.close();}
}
async function pdf(s,loadSource,progress=()=>{}){
 nonempty(s);if(!root.PDFLib)throw new Error('PDF export library did not load. Reload and retry.');
 const doc=await root.PDFLib.PDFDocument.create();doc.setTitle('360 Receipts report');doc.setProducer('360 Receipts 1.1');
 const lines=['360 RECEIPTS — REPORT',...description(s),'','Signed amounts retained. No cross-currency grand total.','Excluded amounts are evidence only, not counted in report totals.','',...s.totals.map(t=>`${t.currency}: ${t.count} counted | Total ${t.total.toFixed(2)} | Reimbursable ${t.reimbursable.toFixed(2)} | Excluded ${t.excludedCount}: ${t.excluded.toFixed(2)}`),'','RECEIPT INDEX (date / vendor / currency & amount / ID)',...s.rows.flatMap((r,i)=>[`${i+1}. ${r.receipt_date} | ${r.vendor} | ${r.currency} ${Number(r.amount).toFixed(2)}${r.status==='excluded'?' — EXCLUDED (not counted)':''}`,`    ${r.id}`])];
 await textPages(doc,lines);
 for(let i=0;i<s.rows.length;i++){
 const r=s.rows[i];progress(i+1,s.rows.length);
 if(!r.storage_path)throw new Error(`Receipt ${r.id}: no source file. PDF not created.`);
 let blob;try{blob=await loadSource(r);}catch{throw new Error(`Receipt ${r.id}: source unavailable or access expired. Sign in / reload and retry. No partial PDF was downloaded.`);}
 if(!blob?.size)throw new Error(`Receipt ${r.id}: empty source file. PDF not created.`);
 await textPages(doc,[`RECEIPT ${i+1} OF ${s.rows.length}`,`${r.receipt_date} | ${r.vendor}`,`${r.currency} ${Number(r.amount).toFixed(2)} | ${r.status}${r.status==='excluded'?' — NOT COUNTED':''}`,`Company / reimbursement: ${r.company||''}`,`Category code: ${r.category_code||''}`,`Reimbursable: ${r.reimbursable?'Yes':'No'}`,`Purpose: ${r.comment||''}`,`Receipt ID: ${r.id}`,`Ledger ID: ${r.ledger_receipt_id||''}`,`Source: ${r.source||''}`,`Original file: ${r.original_name||''}`,`Source path: ${r.storage_path}`,'','Original receipt evidence follows.']);
 try{if(r.mime_type==='application/pdf'||/\.pdf$/i.test(r.storage_path)){const src=await root.PDFLib.PDFDocument.load(await blob.arrayBuffer());if(!src.getPageCount())throw new Error('Empty PDF');const pages=await doc.copyPages(src,src.getPageIndices());pages.forEach(p=>doc.addPage(p));}else await imagePages(doc,blob);}catch{throw new Error(`Receipt ${r.id}: unreadable, encrypted or unsupported source. Convert HEIC/HEIF to JPEG/PNG or unlock the PDF. No partial PDF was downloaded.`);}
 }
 return doc.save();
}
const api={snapshot,description,workbook,pdf};root.ReceiptReports=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

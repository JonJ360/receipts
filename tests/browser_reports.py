"""Synthetic-only export tests. Writes artifacts outside the public repository."""
import json, pathlib, tempfile, base64
from playwright.sync_api import sync_playwright
import pymupdf as fitz, openpyxl
ROOT=pathlib.Path(__file__).resolve().parents[1]
OUT=pathlib.Path(tempfile.gettempdir())/'receipts-export-tests';OUT.mkdir(exist_ok=True)
with sync_playwright() as p:
 b=p.chromium.launch(channel='msedge',headless=True)
 page=b.new_page()
 page.goto('about:blank')
 for file in ['vendor/pdf-lib.min.js','vendor/exceljs.min.js','reports.js']:
  page.add_script_tag(path=str(ROOT/file))
 result=page.evaluate('''async()=>{
 const rows=[{id:'synthetic-image',receipt_date:'2026-09-01',vendor:'École <not HTML>',amount:'25.50',currency:'USD',company:'SMI',category_code:'IT',comment:'Synthetic verification only',status:'ready',storage_path:'image.png',mime_type:'image/png',reimbursable:true},
 {id:'synthetic-pdf',receipt_date:'2026-09-02',vendor:'PDF source',amount:'-5.50',currency:'USD',company:'Wash',status:'ready',storage_path:'two.pdf',mime_type:'application/pdf'},
 {id:'synthetic-excluded',receipt_date:'2026-09-03',vendor:'=1+1',amount:'999',currency:'USD',status:'excluded'}];
 const snap=ReceiptReports.snapshot(rows,{includeExcluded:true});
 const canvas=document.createElement('canvas');canvas.width=800;canvas.height=2400;const ctx=canvas.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,800,2400);ctx.fillStyle='blue';ctx.font='42px Arial';for(let y=60;y<2400;y+=120)ctx.fillText('SYNTHETIC RECEIPT '+y,25,y);
 const img=await (await fetch(canvas.toDataURL())).blob();
 const src=await PDFLib.PDFDocument.create();src.addPage().drawText('SYNTHETIC SOURCE PAGE ONE');src.addPage().drawText('SYNTHETIC SOURCE PAGE TWO');const source=new Blob([await src.save()],{type:'application/pdf'});
 const s=ReceiptReports.snapshot(rows.slice(0,2),{});
 const drawn=[];const originalFillText=CanvasRenderingContext2D.prototype.fillText;CanvasRenderingContext2D.prototype.fillText=function(text,...args){drawn.push(text);return originalFillText.call(this,text,...args)};
 const pdf=await ReceiptReports.pdf(s,async r=>r.mime_type==='image/png'?img:source);
 for(const label of ['Date','Name','Price','Currency','Category Name','Reimbursable','Report Generated using Receipt Tracker'])if(!drawn.includes(label))throw Error('Missing historical layout label: '+label);
 if(!drawn.includes('From: 9/1/26 To: 9/30/26'))throw Error('Monthly report must show complete calendar range');
 if(!drawn.some(t=>t.includes('1 • École <not HTML> • 9/1/26'))||!drawn.some(t=>t.includes('2/2')))throw Error('Missing chronological or multipage caption');
 if(drawn.includes('RECEIPT INDEX')||drawn.some(t=>t.includes('Original receipt evidence follows')))throw Error('Generic cover sheet returned');
 const xlsx=await ReceiptReports.workbook(snap);
 let missing=false;try{await ReceiptReports.pdf(s,async()=>{throw Error('missing')})}catch(e){missing=true}
 let empty=false;try{await ReceiptReports.pdf(ReceiptReports.snapshot([],{}),async()=>img)}catch(e){empty=true}
 let unsupported=false;try{await ReceiptReports.pdf(s,async()=>new Blob(['not an image']))}catch(e){unsupported=true}
 const encode=bytes=>{let s='';for(const x of new Uint8Array(bytes))s+=String.fromCharCode(x);return btoa(s)};
 return {pdf:encode(pdf),xlsx:encode(xlsx),missing,empty,unsupported};
 }''')
 for ext in ['pdf','xlsx']:(OUT/('synthetic-report.'+ext)).write_bytes(base64.b64decode(result.pop(ext)))
 assert all(result.values()),result
 doc=fitz.open(OUT/'synthetic-report.pdf');assert len(doc)==2,len(doc)
 assert abs(doc[0].rect.width-595.28)<0.1
 text=''.join(p.get_text() for p in doc);assert 'SYNTHETIC SOURCE PAGE ONE' in text and 'SYNTHETIC SOURCE PAGE TWO' in text
 assert sum(len(p.get_images()) for p in doc)>=2
 doc[0].get_pixmap().save(OUT/'summary.png');doc[1].get_pixmap().save(OUT/'evidence.png')
 wb=openpyxl.load_workbook(OUT/'synthetic-report.xlsx');assert wb.sheetnames==['Summary','Receipts']
 sheet=wb['Receipts'];assert sheet.max_row==4;assert not sheet._images
 assert [c.value for c in sheet[1]][:8]==['Date','Name','Price','Currency','Category Name','Category Code','Comment','Reimbursable']
 assert sheet['C2'].value==25.5 and sheet['C2'].data_type=='n'
 assert sheet['A2'].is_date
 assert any(c.value=='=1+1' and c.data_type=='s' for row in sheet for c in row)
 assert 'A1:F1' in [str(r) for r in wb['Summary'].merged_cells.ranges]
 summary=list(wb['Summary'].values);assert any(20 in row and 999 in row for row in summary),summary
 print(json.dumps({'checks':result,'pdf_pages':len(doc),'workbook_rows':sheet.max_row-1,'output':str(OUT)}))
 b.close()

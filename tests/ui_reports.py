"""End-to-end UI tests use a synthetic Supabase client; never write production."""
import pathlib,tempfile,threading,http.server,functools,json
from playwright.sync_api import sync_playwright
ROOT=pathlib.Path(__file__).resolve().parents[1];OUT=pathlib.Path(tempfile.gettempdir())/'receipts-export-tests';OUT.mkdir(exist_ok=True)
class Quiet(http.server.SimpleHTTPRequestHandler):
 def log_message(self,*args):pass
server=http.server.ThreadingHTTPServer(('127.0.0.1',0),functools.partial(Quiet,directory=str(ROOT)));threading.Thread(target=server.serve_forever,daemon=True).start()
FAKE="""window.failAuth=false;window.failRead=false;window.reads=0;window.supabase={createClient:()=>({auth:{getSession:async()=>({data:{session:{access_token:'synthetic',user:{email:'test@example.invalid'}}}}),refreshSession:async()=>({data:{session:{access_token:'synthetic',user:{email:'test@example.invalid'}}}}),getUser:async()=>({data:{user:window.failAuth?null:{id:'synthetic'}},error:window.failAuth?{message:'expired'}:null})},from:()=>({select(){return this},order(){return this},async range(a,b){window.reads++;return {error:window.failRead?{message:'synthetic read error'}:null,data:Array.from({length:1001},(_,i)=>({id:'synthetic-'+String(i).padStart(4,'0'),receipt_date:'2026-09-01',vendor:'Synthetic vendor',amount:'1.00',currency:'USD',company:'Test',status:i===1000?'excluded':'ready'})).slice(a,b+1)}}})})};"""
with sync_playwright() as p:
 b=p.chromium.launch(channel='msedge',headless=True);page=b.new_page(viewport={'width':1280,'height':900})
 page.route('https://cdn.jsdelivr.net/**',lambda r:r.fulfill(body=FAKE,content_type='text/javascript'))
 errors=[];page.on('pageerror',lambda e:errors.append(str(e)))
 page.goto(f'http://127.0.0.1:{server.server_port}/');page.locator('[data-view="list"]').click()
 page.locator('#kTotal').filter(has_text='USD 1000.00').wait_for()
 assert 'excluded' in page.locator('#listNote').inner_text()
 page.locator('#exportBtn').click();page.locator('#reportSummary').filter(has_text='1000').wait_for()
 assert page.locator('#exportPdf').is_enabled();assert 'USD 1000.00' in page.locator('#reportSummary').inner_text()
 page.screenshot(path=str(OUT/'desktop.png'))
 with page.expect_download() as download:page.locator('#exportExcel').click()
 download.value.save_as(OUT/'ui-report.xlsx')
 page.locator('#includeExcluded').check();assert 'Excluded' in page.locator('#reportSummary').inner_text()
 page.set_viewport_size({'width':390,'height':844});page.screenshot(path=str(OUT/'mobile.png'));assert page.evaluate('document.documentElement.scrollWidth<=innerWidth')
 page.locator('#closeExport').click();page.locator('#fltSearch').fill('absent');page.locator('#exportBtn').click();page.locator('#reportSummary').filter(has_text='0 receipts').wait_for();assert page.locator('#exportExcel').is_disabled()
 page.locator('#closeExport').click();page.evaluate('window.failAuth=true');page.locator('#exportBtn').click();page.locator('#exportMsg').filter(has_text='Sign in').wait_for();assert page.locator('#exportExcel').is_disabled()
 page.locator('#closeExport').click();page.evaluate('window.failAuth=false;window.failRead=true');page.locator('#exportBtn').click();page.locator('#exportMsg').filter(has_text='Unable').wait_for();assert page.locator('#exportExcel').is_disabled()
 assert not errors,errors
 print(json.dumps({'pagination':1001,'download':'ui-report.xlsx','desktop_mobile':True,'empty_auth_read_errors':True,'page_errors':errors}))
 b.close()
server.shutdown()

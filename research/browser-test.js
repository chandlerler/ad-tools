'use strict';
const {chromium}=require('playwright'),fs=require('fs'),path=require('path'),assert=require('assert'),http=require('http');
const XLSX=require('../vendor/xlsx.full.min.js');
function walk(p){return fs.readdirSync(p,{withFileTypes:true}).flatMap(e=>e.isDirectory()?walk(path.join(p,e.name)):[path.join(p,e.name)]);}
(async()=>{
 const paths=walk('/workspace/attachments'),zip=paths.find(p=>p.includes('67bad2de-')&&p.endsWith('.zip')),prev=paths.find(p=>p.endsWith('20261003_20261016_21_09_10_汇总.xlsx')),cur=paths.find(p=>p.endsWith('20261017_20261030_21_09_35_汇总.xlsx'));
 const server=http.createServer((req,res)=>{res.setHeader('Content-Type','text/html; charset=utf-8');res.end(fs.readFileSync('dist/广告数据工作台.html'));});await new Promise(r=>server.listen(0,'127.0.0.1',r));
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 try{
 const page=await browser.newPage({viewport:{width:1360,height:1050}});page.setDefaultTimeout(30000);const errors=[];page.on('pageerror',e=>errors.push(e.message));await page.goto('http://127.0.0.1:'+server.address().port);
 for(const [i,p] of [zip,prev,cur].entries())await page.locator('#file'+i).setInputFiles(p);
 await page.locator('#run').click();await page.locator('#ownersDialog[open]').waitFor();
 assert.equal(await page.locator('#ownerOptions input').count(),6);assert.equal(await page.locator('#ownerOptions input:checked').count(),1);assert.equal(await page.locator('#ownerOptions input:checked').inputValue(),'左春泽(zuochunze.1)');
 await page.screenshot({path:'research/owner-selection.png',fullPage:true});
 await page.locator('#ownerOptions input:checked').uncheck();await page.locator('#ownersConfirm').click();assert((await page.locator('#ownerError').innerText()).includes('至少'));
 await page.locator('#ownerOptions input[value="左春泽(zuochunze.1)"]').check();await page.locator('#ownersConfirm').click();await page.locator('#download').waitFor({state:'visible'});
 const downloaded=page.waitForEvent('download');await page.locator('#download').click();const d=await downloaded;await d.saveAs('/tmp/owner-browser-output.xlsx');
 const w=XLSX.read(fs.readFileSync('/tmp/owner-browser-output.xlsx'));assert.equal(w.SheetNames.length,3);
 const rows=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1,defval:null});assert.equal(rows.length,96);assert.equal(rows[0].length,27);assert.equal(rows.slice(1).reduce((n,r)=>n+r[3],0),345433);assert.equal(rows.filter(r=>String(r[0]).startsWith('未匹配SKU:')).length,13);
 for(const n of w.SheetNames.slice(1)){const r=XLSX.utils.sheet_to_json(w.Sheets[n],{header:1});assert.equal(r[0].filter(x=>x==='商品ID').length,2);assert(r.slice(1).every(x=>x[0]==='左春泽(zuochunze.1)'));}
 await page.screenshot({path:'research/screenshot.png',fullPage:true});
 // Multi-select is functional, and switching selection produces fresh data.
 await page.locator('#run').click();await page.locator('#ownersDialog[open]').waitFor();await page.locator('#ownerOptions input[value="叶云燕(yeyunyan.3)"]').check();await page.locator('#ownersConfirm').click();await page.locator('#download').waitFor({state:'visible'});assert((await page.locator('#warnings').innerText()).includes('叶云燕'));
 // Canceling selection must unblock the file controls without exporting stale results.
 await page.locator('#run').click();await page.locator('#ownersDialog[open]').waitFor();await page.locator('#ownersCancel').click();assert(await page.locator('#download').isHidden());assert(await page.locator('#run').isEnabled());
 // Original 13-day attachments must still be blocked, even with a valid new ZIP.
 await page.locator('#file1').setInputFiles(paths.find(p=>p.includes('20261001_20261013')&&p.endsWith('.xlsx')));await page.locator('#file2').setInputFiles(paths.find(p=>p.includes('20261014_20261026')&&p.endsWith('.xlsx')));await page.locator('#run').click();await page.locator('#alert[open]').waitFor();assert((await page.locator('#alertText').innerText()).includes('实际 13 天'));assert(await page.locator('#download').isHidden());await page.locator('#dismiss').click();
 // Drag/drop CSV still selects the intended slot.
 await page.locator('[data-slot="0"]').evaluate(el=>{const dt=new DataTransfer();dt.items.add(new File(['日期,SPU ID\n'],'drop.csv'));el.dispatchEvent(new DragEvent('drop',{dataTransfer:dt,bubbles:true,cancelable:true}));});assert((await page.locator('[data-slot="0"] .filename').innerText()).includes('drop.csv'));assert.deepEqual(errors,[]);
 console.log('PASS browser: six-owner checkbox dialog, recommended owner, empty-selection validation, multi-select, cancel, real ZIP mapping/export, duplicate headers, drag/drop and 14-day enforcement.');
 }finally{await browser.close();server.close();}
})().catch(e=>{console.error(e);process.exit(1)});

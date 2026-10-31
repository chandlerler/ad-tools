'use strict';
const fs=require('fs'),path=require('path'),assert=require('assert');
global.XLSX=require('../vendor/xlsx.full.min.js');global.fflate=require('../vendor/fflate.min.js');
const core=require('../core.js');
function walk(p){return fs.readdirSync(p,{withFileTypes:true}).flatMap(e=>e.isDirectory()?walk(path.join(p,e.name)):[path.join(p,e.name)]);}
const paths=walk('/workspace/attachments');
const inputs=[paths.find(p=>p.includes('67bad2de-')&&p.endsWith('.zip')),paths.find(p=>p.endsWith('20261003_20261016_21_09_10_汇总.xlsx')),paths.find(p=>p.endsWith('20261017_20261030_21_09_35_汇总.xlsx'))];
const asFile=p=>({name:path.basename(p),arrayBuffer:async()=>fs.readFileSync(p)});
(async()=>{
 const p=await core.prepareFiles(inputs.map(asFile),console.log);
 assert.equal(p.choices.length,6);assert.equal(p.choices[0].name,'左春泽(zuochunze.1)');assert.equal(p.choices[0].exposure,343065);
 assert.equal(p.prev.duplicates,1571);assert.equal(p.cur.duplicates,1389);
 assert.throws(()=>core.finishPrepared(p,[]),/至少选择/);
 const r=core.finishPrepared(p,['左春泽(zuochunze.1)']);
 assert.equal(r.summary.rows,51212);assert.equal(r.summary.unmatched,13);assert.equal(r.summary.total[0],345433);assert.equal(r.summary.total[1],24818);assert.equal(r.summary.total[3],1571);
 assert.deepEqual(r.wb.SheetNames,['广告搜索&推荐数据','广告成交数据（对比周期）','广告成交数据（现期）']);
 for(const name of r.wb.SheetNames.slice(1)){const rows=XLSX.utils.sheet_to_json(r.wb.Sheets[name],{header:1});assert.equal(rows[0].filter(h=>h==='商品ID').length,2);assert.equal(rows[0].filter(h=>h==='商品名称').length,2);assert(rows.slice(1).every(row=>row[0]==='左春泽(zuochunze.1)'));}
 const first=XLSX.utils.sheet_to_json(r.wb.Sheets[r.wb.SheetNames[0]],{header:1,defval:null});assert.equal(first[0].length,27);const missing=first.slice(1).filter(row=>String(row[0]).startsWith('未匹配SKU:'));assert.equal(missing.length,13);assert(missing.every(row=>row.slice(11).every(x=>x==null)));assert.equal(missing.reduce((a,r)=>a+r[3],0),2368);
 fs.writeFileSync('deliverables/广告分析_20261003-20261030_左春泽.xlsx',XLSX.write(r.wb,{type:'buffer',bookType:'xlsx',compression:true}));
 fs.writeFileSync('research/latest-validation.json',JSON.stringify({summary:r.summary,warnings:r.warnings,choices:p.choices},null,2));
 const all=core.finishPrepared(p,p.choices.map(c=>c.name));assert(all.summary.output>r.summary.output);assert.equal(all.summary.total[0],r.summary.total[0]);
 const other=core.finishPrepared(p,['叶云燕(yeyunyan.3)']);assert.equal(other.summary.total[0],2368);assert(other.warnings.some(s=>s.includes('已排除')));
 // Regression: old SPU exports still work, duplicate header aliases don't affect metric columns.
 const old=core.createAggregation();core.addCSV(old,new TextEncoder().encode('SPU_ID,花费,展现数,点击数,直接订单行,直接订单金额\n10001,10,100,20,3,30'),'old.csv');assert.deepEqual(old.byId.get('10001').v,[100,20,30,3,10]);
 // A SKU associated with two SPUs must remain unresolved, never guessed.
 const conflict=core.createAggregation();core.addCSV(conflict,new TextEncoder().encode('SKU ID,花费,展现数,点击数,直接订单行,直接订单金额\n99999,10,100,20,3,30'),'conflict.csv');
 const t=core.effectTable(fs.readFileSync(inputs[1]),'prev');const e=t.entries[0];t.entries=[{...e,id:'111',sku:'99999'},{...e,id:'222',sku:'99999'}];const result=core.finishPrepared({agg:conflict,prev:t,cur:{...t,entries:[]},choices:[{name:e.owner}]},[e.owner]);assert.equal(result.summary.unmatched,1);
 const fixture=XLSX.utils.book_new();XLSX.utils.book_append_sheet(fixture,XLSX.utils.aoa_to_sheet([['时间','2026-10-03 ~ 2026-10-16']]),'条件');
 const header=['运营人员','商品ID','商品名称','商品ID','商品名称','点击','消耗','成交父单量','成交金额','展现','三级品类'];const entry=['人员A','100','SPU名','200','SKU1',10,5,2,20,100,'类目'];
 XLSX.utils.book_append_sheet(fixture,XLSX.utils.aoa_to_sheet([header,entry,entry,['人员A','100','SPU名','201','SKU2',20,8,3,40,200,'类目']]),'数据');
 const parsed=core.effectTable(XLSX.write(fixture,{type:'buffer',bookType:'xlsx'}),'fixture');assert.equal(parsed.duplicates,1);assert.equal(parsed.entries.length,2);assert.equal(parsed.entries[0].sku,'200');assert.deepEqual(parsed.byId.get('100').v,[300,30,60,5,13]);
 console.log('PASS:',JSON.stringify(r.summary));
})().catch(e=>{console.error(e);process.exit(1)});

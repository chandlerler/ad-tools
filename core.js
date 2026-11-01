/* Local-only advertising report processor. Libraries supplied by the offline bundle. */
'use strict';
const REPORT_NAMES=['广告搜索&推荐数据','广告成交数据（对比周期）','广告成交数据（现期）'];
const METRICS=['曝光','点击','CTR','CVR','GMV','订单','CPC','ROI'];
const OUT_HEADERS=['SPU_ID','商品名称','三级类目',...['搜索','推荐','全量'].flatMap(p=>METRICS.map(m=>p+'_'+m))];
const canon=x=>String(x??'').replace(/^\uFEFF/,'').trim().replace(/[\s_]/g,'').toLowerCase();
function headerMap(row){const m=new Map();row.forEach((v,i)=>{const k=canon(v);if(k){if(m.has(k))throw Error('重复列名：'+v);m.set(k,i);}});return m;}
function column(m,names,required=true){for(const n of names)if(m.has(canon(n)))return m.get(canon(n));if(required)throw Error('缺少必需列：'+names.join(' / '));return -1;}
function idValue(v){const s=String(v??'').trim().replace(/^['\t]/,'').replace(/\.0$/,'');if(!/^\d+$/.test(s))throw Error('无效的商品/SPU ID：'+s);return s;}
function numberValue(v,label){if(typeof v==='number'&&Number.isFinite(v))return v;const s=String(v??'').trim().replace(/,/g,'');if(!s||s==='-')return 0;if(!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(s))throw Error(label+' 包含无效数值：'+s);return Number(s);}
function eachCSV(text,callback){let row=[],field='',quoted=false;for(let i=0;i<text.length;i++){const ch=text[i];if(ch==='"'){if(quoted&&text[i+1]==='"'){field+='"';i++;}else if(quoted)quoted=false;else if(!field)quoted=true;else field+='\"';}else if(!quoted&&(ch===','||ch==='\n'||ch==='\r')){row.push(field);field='';if(ch!==','){if(ch==='\r'&&text[i+1]==='\n')i++;if(row.some(x=>x.trim()))callback(row);row=[];}}else field+=ch;}if(quoted)throw Error('CSV 存在未闭合引号');if(field||row.length){row.push(field);if(row.some(x=>x.trim()))callback(row);}}
function decodeCSV(bytes){try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch(e){return new TextDecoder('gb18030',{fatal:true}).decode(bytes);}}
// Repeated 商品ID / 商品名称 headers are semantic: left is SPU, right is SKU.
function positions(headers,names){return headers.flatMap((v,i)=>names.some(n=>canon(n)===canon(v))?[i]:[]);}
function effectTable(buffer,label){
 const w=XLSX.read(buffer,{type:'array'});
 if(w.SheetNames.length<2)throw Error(label+'：至少需要两个 sheet，有效数据须在第二个 sheet');
 const rows=XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[1]],{header:1,defval:null,raw:true});
 const h=rows.findIndex(r=>positions(r,['商品ID','SPU ID']).length);
 if(h<0)throw Error(label+'：第二个 sheet 找不到商品ID列');
 const headers=rows[h],m=new Map();headers.forEach((v,i)=>{if(!m.has(canon(v)))m.set(canon(v),i);});
 const ids=positions(headers,['商品ID']),names=positions(headers,['商品名称']);
 const id=column(m,['SPU ID'],false)>=0?column(m,['SPU ID']):ids[0];
 const sku=column(m,['SKU ID'],false)>=0?column(m,['SKU ID']):(ids[1]??-1);
 const name=column(m,['SPU名称'],false)>=0?column(m,['SPU名称']):names[0];
 const cat=column(m,['三级品类','三级类目']);
 const owner=column(m,['运营人员','采销','采销名','采销姓名','采销人员'],false);
 if(id==null||name==null)throw Error(label+'：缺少SPU商品ID或商品名称');
 const measures=['展现','点击','成交金额','成交父单量','消耗'].map(n=>column(m,[n]));
 const entries=[],seen=new Set();let duplicates=0;
 for(let i=h+1;i<rows.length;i++){
  let row=headers.map((_,j)=>rows[i][j]??null);
  if(row.every(v=>v==null||v===''))continue;
  if(['合计','总计','汇总'].includes(String(row[id]).trim())||['合计','总计','汇总'].includes(String(row[cat]).trim()))continue;
  const spu=idValue(row[id]);row[id]=spu;
  let skuId=null;if(sku>=0&&row[sku]!=null&&String(row[sku]).trim()&&!['-','--'].includes(String(row[sku]).trim())){skuId=idValue(row[sku]);row[sku]=skuId;}
  row=row.map((v,j)=>{if(j!==id&&j!==sku&&typeof v==='string'&&/^[+-]?[\d,.]+%$/.test(v.trim()))return Number(v.replace(/[,%]/g,''))/100;return v;});
  const signature=JSON.stringify(row);if(seen.has(signature)){duplicates++;continue;}seen.add(signature);
  entries.push({id:spu,sku:skuId,owner:owner<0?'':String(row[owner]??'未标注采销').trim(),row,name:row[name],cat:row[cat],v:measures.map(j=>numberValue(row[j],headers[j]))});
 }
 if(!entries.length)throw Error(label+'：第二个 sheet 没有商品数据');
 let period='未提供';for(const r of XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{header:1}))if(r[0]==='时间')period=String(r[1]);
 return filterEffect({headers,m,entries,period,duplicates},null);
}
function filterEffect(table,selected){
 const entries=selected?table.entries.filter(e=>!e.owner||selected.has(e.owner)):table.entries;
 const byId=new Map();for(const e of entries){let g=byId.get(e.id);if(!g){g={name:e.name,cat:e.cat,v:[0,0,0,0,0],row:e.row};byId.set(e.id,g);}e.v.forEach((n,i)=>g.v[i]+=n);}
 return {...table,entries,byId,data:entries.map(e=>e.row)};
}
function createAggregation(){return {byId:new Map(),rows:0,files:0,periods:new Set(),total:[0,0,0,0,0]};}
function addCSV(agg,bytes,filename){
 let m,idx,line=0,width,kind;eachCSV(decodeCSV(bytes),row=>{
  line++;if(!m){m=headerMap(row);width=row.length;let id=column(m,['SPU ID','SPU_ID'],false);kind='SPU';
   if(id<0){id=column(m,['智能投放推广SKU ID','SKU ID'],false);kind='SKU';}
   if(id<0)throw Error(filename+'：搜索报表缺少 SPU ID 和 SKU ID，无法识别商品。');
   idx=[id,column(m,['展现数']),column(m,['点击数']),column(m,['直接订单金额']),column(m,['直接订单行']),column(m,['花费']),column(m,['推广计划'],false),column(m,['推广品类'],false),column(m,['日期'],false)];return;
  }
  if(row.length!==width)throw Error(filename+' 第'+line+'行列数与表头不一致');
  try{const sourceId=idValue(row[idx[0]]),key=kind==='SKU'?'SKU:'+sourceId:sourceId;let item=agg.byId.get(key);
   if(!item){item={kind,sourceId,v:[0,0,0,0,0],name:idx[6]>=0?row[idx[6]]:'',cat:idx[7]>=0?row[idx[7]].replace(/^[^-]*-/,''):''};agg.byId.set(key,item);}
   for(let k=0;k<5;k++){const n=numberValue(row[idx[k+1]],'指标');item.v[k]+=n;agg.total[k]+=n;}
   if(idx[8]>=0)agg.periods.add(row[idx[8]]);agg.rows++;
  }catch(e){throw Error(filename+' 第'+line+'行：'+e.message);}
 });if(!m)throw Error(filename+' 是空文件');agg.files++;
}
function mappingFor(tables){
 const sku=new Map(),spuOwners=new Map();for(const t of tables)for(const e of t.entries){
  if(!spuOwners.has(e.id))spuOwners.set(e.id,new Set());if(e.owner)spuOwners.get(e.id).add(e.owner);
  if(e.sku){if(!sku.has(e.sku))sku.set(e.sku,new Set());sku.get(e.sku).add(e.id);}
 }return {sku,spuOwners};
}
function ownerChoices(prev,cur,agg){
 const map=mappingFor([prev,cur]),choices=new Map();for(const t of [prev,cur])for(const e of t.entries)if(e.owner&&!choices.has(e.owner))choices.set(e.owner,{name:e.owner,exposure:0,matched:0});
 for(const [key,e] of agg.byId){const ids=e.kind==='SKU'?map.sku.get(e.sourceId):new Set([key]);const owners=new Set();for(const id of ids||[])for(const owner of map.spuOwners.get(id)||[])owners.add(owner);for(const owner of owners){const c=choices.get(owner);if(c){c.exposure+=e.v[0];c.matched++;}}}
 return [...choices.values()].sort((a,b)=>b.exposure-a.exposure||a.name.localeCompare(b.name));
}
function resolveSearch(agg,prev,cur,selected){
 const all=mappingFor([prev,cur]),filtered=mappingFor([filterEffect(prev,selected),filterEffect(cur,selected)]);
 const out={...agg,byId:new Map(),total:[0,0,0,0,0],unmatched:[],excluded:0,excludedExposure:0};
 for(const [raw,e] of agg.byId){let id=raw;
  if(e.kind==='SKU'){
   const candidates=filtered.sku.get(e.sourceId)||all.sku.get(e.sourceId);
   id=candidates?.size===1?[...candidates][0]:'未匹配SKU:'+e.sourceId;
   if(candidates?.size!==1)out.unmatched.push({sku:e.sourceId,reason:candidates?.size?'映射有冲突':'效果表无对应SKU'});
  }
  const owners=all.spuOwners.get(id);
  if(selected&&owners?.size&&![...owners].some(o=>selected.has(o))){out.excluded++;out.excludedExposure+=e.v[0];continue;}
  let g=out.byId.get(id);if(!g){g={...e,v:[0,0,0,0,0]};out.byId.set(id,g);}e.v.forEach((n,i)=>{g.v[i]+=n;out.total[i]+=n;});
 }
 return out;
}
const round4=n=>Math.round((n+Number.EPSILON)*10000)/10000;
function metrics(v){if(!v)return Array(8).fill(null);const [im,cl,gmv,orders,cost]=v;const ratio=(a,b)=>b>0?round4(a/b):null;return [im,cl,ratio(cl,im),ratio(orders,cl),Math.round(gmv*100)/100,orders,ratio(cost,cl),ratio(gmv,cost)];}
function assemble(agg,previous,current,mode){const warnings=[...(agg.notes||[])];const rows=[];const all=new Set([...agg.byId.keys(),...previous.byId.keys(),...current.byId.keys()]);const getTotals=(table,id)=>table.byId.get(id)?.v||[0,0,0,0,0];let missing=0,negative=0;
for(const id of all){const s=agg.byId.get(id),meta=current.byId.get(id)||previous.byId.get(id);const search=s?s.v:null;let total=null,recommend=null;if(mode==='estimate' && meta && s){const a=getTotals(previous,id),b=getTotals(current,id);total=a.map((v,i)=>v+b[i]);const diff=total.map((v,i)=>v-search[i]);if(diff.some(v=>v< -0.005))negative++;recommend=diff;}if(!meta)missing++;rows.push([id,meta?.name||s?.name||'',meta?.cat||s?.cat||'',...metrics(search),...metrics(recommend),...metrics(total)]);}
rows.sort((a,b)=>(b[3]??0)-(a[3]??0)||a[0].localeCompare(b[0]));if(mode!=='estimate')warnings.push('推荐和全量指标留空：三个输入不足以复现样例中的全量广告数据。');else warnings.push('全量来自两个周期效果表之和；推荐为全量减搜索。注意：效果表成交父单量与搜索直接订单行、成交金额与直接订单金额的口径可能不同。');if(negative)warnings.push(negative+' 个SPU存在负差额，已保留原值；相关分母非正时比率留空，请核对指标口径。');const dimensionMissing=[...all].filter(id=>!agg.byId.has(id)).length;if(dimensionMissing)warnings.push(dimensionMissing+' 个SPU没有搜索维度记录，首张表的搜索/推荐/全量指标留空；诊断GMV和毛利仍取两周期成交表。');if(missing)warnings.push(missing+' 个搜索商品未在效果表中匹配到，名称/类目回退至搜索表，全量和推荐留空。');warnings.push('搜索 CPC 使用搜索花费/搜索点击；ROI 使用直接订单金额/搜索花费。推荐 CPC 使用差额花费/差额点击；全量 CPC 使用全量消耗/全量点击。');
const wb=XLSX.utils.book_new();function append(headers,data,name){const ws=XLSX.utils.aoa_to_sheet([headers,...data]);ws['!autofilter']={ref:ws['!ref']};ws['!cols']=headers.map(h=>({wch:/商品名称/.test(h)?48:/ID/i.test(h)?19:Math.max(13,String(h).length*2+2)}));for(let r=1;r<=data.length;r++)for(let c=0;c<headers.length;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell&&typeof cell.v==='number')cell.z=/CTR|CVR|率|环比|同比/.test(headers[c])?'0.00%':Number.isInteger(cell.v)?'0':'0.0000';}XLSX.utils.book_append_sheet(wb,ws,name);return ws;}
const ws=append(OUT_HEADERS,rows,REPORT_NAMES[0]);ws.A1.c=[{a:'广告数据处理工具',t:['处理口径：按SPU ID聚合所有搜索CSV分片；订单与GMV使用直接订单行、直接订单金额。',...warnings,'搜索周期：'+[...agg.periods].join('、'),'对比周期：'+previous.period,'现期：'+current.period].join('\n')}];append(previous.headers,previous.data,REPORT_NAMES[1]);append(current.headers,current.data,REPORT_NAMES[2]);return {wb,warnings,summary:{files:agg.files,rows:agg.rows,spus:[...agg.byId.keys()].filter(k=>!k.startsWith('未匹配SKU:')).length,unmatched:agg.unmatched?.length||0,selectedOwners:agg.selectedOwners||[],output:rows.length,total:agg.total,previous:previous.period,current:current.period,search:[...agg.periods]}};}
async function prepareFiles(files,progress=()=>{},validate=true){progress('读取两个周期的效果表…');const prev=effectTable(await files[1].arrayBuffer(),'对比周期'),cur=effectTable(await files[2].arrayBuffer(),'现期');if(validate)validateEffectPeriods(prev,cur);const agg=createAggregation(),f=files[0];if(/\.csv$/i.test(f.name)){progress('汇总搜索 CSV…');addCSV(agg,new Uint8Array(await f.arrayBuffer()),f.name);}else if(/\.zip$/i.test(f.name)){const bytes=new Uint8Array(await f.arrayBuffer());let names=[];fflate.unzipSync(bytes,{filter:f=>{if(/\.csv$/i.test(f.name)&&!f.name.startsWith('__MACOSX/'))names.push(f.name);return false;}});if(!names.length)throw Error('ZIP 内没有 CSV 文件');for(let i=0;i<names.length;i++){const name=names[i];progress('汇总分片 '+(i+1)+' / '+names.length);const part=fflate.unzipSync(bytes,{filter:f=>f.name===name});addCSV(agg,part[name],name);}}else throw Error('搜索数据只支持 ZIP 或 CSV');if(!agg.rows)throw Error('搜索数据没有有效记录');if(validate)validateSearchPeriod(agg,prev,cur);return {agg,prev,cur,choices:ownerChoices(prev,cur,agg)};}
function finishPrepared(prepared,selectedOwners=null,mode='estimate'){
 const {agg,prev,cur,choices}=prepared;
 if(selectedOwners&&selectedOwners.some(n=>!choices.some(c=>c.name===n)))throw Error('采销选择已失效，请重新选择文件');
 if(choices.length&&(!selectedOwners||!selectedOwners.length))throw Error('请至少选择一位采销人员');
 const selected=selectedOwners?new Set(selectedOwners):null;
 const search=resolveSearch(agg,prev,cur,selected);search.selectedOwners=selectedOwners||[];
 search.notes=[];
 if(selectedOwners?.length)search.notes.push('本次采销范围：'+selectedOwners.join('、'));
 if(prev.duplicates||cur.duplicates)search.notes.push('效果表全组完全相同的重复行已去重：对比周期 '+prev.duplicates+' 行，现期 '+cur.duplicates+' 行；同一SPU的不同SKU正常累加。');
 if(search.unmatched.length)search.notes.push(search.unmatched.length+' 个SKU无法唯一映射，已保留为“未匹配SKU:编号”行，搜索指标保留，全量/推荐留空；这些SKU归属无法确认，未按采销排除。编号：'+search.unmatched.map(x=>x.sku).join('、'));
 if(search.excluded)search.notes.push('已排除属于未选采销的 '+search.excluded+' 个搜索商品（曝光 '+search.excludedExposure+'）。');
 return assemble(search,filterEffect(prev,selected),filterEffect(cur,selected),mode);
}
async function processFiles(files,mode='estimate',progress=()=>{},validate=true,selectedOwners=null){
 const p=await prepareFiles(files,progress,validate);
 const selected=selectedOwners||(p.choices.length?p.choices.filter(c=>c.matched>0||p.choices.length===1).map(c=>c.name):null);
 return finishPrepared(p,selected,mode);
}
function periodDates(text){const m=String(text).match(/(20\d{2})[-/]?(\d{2})[-/]?(\d{2})\s*[~～至—–]\s*(20\d{2})[-/]?(\d{2})[-/]?(\d{2})/);if(!m)throw Error('周期日期无法识别：'+text);const parts=m.slice(1).map(Number);const date=p=>{const t=Date.UTC(p[0],p[1]-1,p[2]),d=new Date(t);if(d.getUTCFullYear()!==p[0]||d.getUTCMonth()!==p[1]-1||d.getUTCDate()!==p[2])throw Error('周期日期无效：'+text);return t;};const start=date(parts.slice(0,3)),end=date(parts.slice(3));if(end<start)throw Error('周期结束日期早于开始日期：'+text);return {start,end,days:(end-start)/86400000+1};}
function validateEffectPeriods(prev,cur){const a=periodDates(prev.period),b=periodDates(cur.period);const issues=[];if(a.days!==14)issues.push('对比周期 '+prev.period+'，实际 '+a.days+' 天。');if(b.days!==14)issues.push('现期 '+cur.period+'，实际 '+b.days+' 天。');if(issues.length)throw Error('每个效果周期必须为14天（含起止日）。\n'+issues.join('\n')+'\n请重新导出两个14天周期的效果表后再处理。');if(a.end+86400000!==b.start)throw Error('两个效果周期必须按先后顺序连续且不重叠，请检查文件位置与日期。');}
function validateSearchPeriod(agg,prev,cur){const a=periodDates(prev.period),b=periodDates(cur.period);if(agg.periods.size!==1)throw Error('搜索周期必须唯一，且覆盖两个效果周期；当前检测到 '+agg.periods.size+' 个周期。');const s=periodDates([...agg.periods][0]);if(s.start!==a.start||s.end!==b.end)throw Error('搜索周期与两个效果周期的合并范围不一致，请重新导出覆盖28天的搜索数据。');}
if(typeof module!=='undefined')module.exports={prepareFiles,finishPrepared,filterEffect,mappingFor,eachCSV,addCSV,createAggregation,effectTable,assemble,processFiles,metrics,periodDates,validateEffectPeriods,validateSearchPeriod};

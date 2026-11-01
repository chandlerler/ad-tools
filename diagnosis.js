/* Auditable diagnosis rules. No network, no model calls, no inferred field aliases. */
'use strict';
const DIAG_GROUPS=['停投','先优化承接','ROI／预算调整','数据待补充'];
const diagCanon=v=>String(v??'').trim().replace(/[\s_]/g,'').toLowerCase();
const diagId=v=>{const s=String(v??'').trim().replace(/\.0$/,'');return /^\d+$/.test(s)&&/[1-9]/.test(s)?s:null;};
function diagNumber(v){if(v==null||String(v).trim()===''||['-','--','N/A','#N/A'].includes(String(v).trim()))return null;if(typeof v==='number')return Number.isFinite(v)?v:null;const s=String(v).trim().replace(/,/g,'');if(/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%?$/.test(s))return Number(s.replace('%',''))/(s.endsWith('%')?100:1);return null;}
function diagnosisSheets(wb){
 const schemas=wb.SheetNames.map(name=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null,raw:true});const h=rows.findIndex(r=>r.some(v=>['spuid','商品id'].includes(diagCanon(v))));if(h<0)return {name,headers:[],rows:[],fields:[],headerRow:null};const headers=rows[h],data=rows.slice(h+1).filter(r=>r.some(v=>v!=null&&v!==''));const fields=headers.map((n,i)=>({name:String(n??''),types:[...new Set(data.map(r=>r[i]==null?'空值':typeof r[i]==='number'?'数值':typeof r[i]==='string'&&/%$/.test(r[i].trim())?'百分比文本':typeof r[i]==='string'?'文本':typeof r[i]))]}));return {name,headers,rows:data,fields,headerRow:h+1};});
 const find=(test,label)=>{const hits=schemas.filter(test);if(hits.length!==1)throw Error('无法唯一识别'+label+'，请检查Sheet名称和字段。');return hits[0];};
 const search=find(s=>s.headers.some(h=>diagCanon(h)==='搜索ctr')&&s.headers.some(h=>diagCanon(h)==='推荐ctr'),'搜索/推荐Sheet');
 const previous=find(s=>/对比周期|前14天/.test(s.name)&&s.headers.some(h=>diagCanon(h)==='成交金额'),'对比周期Sheet');
 const current=find(s=>/现期|现周期|后14天/.test(s.name)&&s.headers.some(h=>diagCanon(h)==='成交金额'),'现周期Sheet');
 return {search,previous,current};
}
function diagnosisIndex(sheet,names,required=false){for(const name of names){const i=sheet.headers.findIndex(h=>diagCanon(h)===diagCanon(name));if(i>=0)return i;}if(required)throw Error(sheet.name+'缺少字段：'+names.join('/'));return -1;}
function analyzeDiagnosis(wb,meta={},options={}){
 const sheets=diagnosisSheets(wb),quality={excluded:[],duplicateSearch:[],invalidRates:[],conflicts:[],zeroClicks:[],negativeClicks:[],missingCurrent:[],missingPrevious:[],missingCurrentGMV:[],missingPreviousGMV:[],missingProfit:[],missingDimensions:[],partialCells:[],sumMismatches:[]};
 function effect(sheet){const id=diagnosisIndex(sheet,['SPU_ID','商品ID'],true),name=diagnosisIndex(sheet,['商品名称','SPU名称']),cat=diagnosisIndex(sheet,['三级品类','三级类目']),gmv=diagnosisIndex(sheet,['成交金额'],true);let profit=diagnosisIndex(sheet,['京喜预估投后履约毛利_折后']);let profitBasis='指定折后字段';if(profit<0&&options.confirmGenericProfit){profit=diagnosisIndex(sheet,['预估投后履约毛利']);profitBasis='用户确认：预估投后履约毛利与指定折后字段同口径';}if(profit<0)profitBasis='缺少指定毛利字段，未替代';sheet.profitBasis=profitBasis;sheet.profitField=profit<0?null:sheet.headers[profit];const map=new Map();
  for(const row of sheet.rows){const key=diagId(row[id]);if(!key){quality.excluded.push({sheet:sheet.name,id:String(row[id]??''),reason:'无法定位有效SPU'});continue;}let g=map.get(key);if(!g){g={id:key,name:row[name]||'',cat:row[cat]||'',gmv:0,profit:0,gmvCount:0,profitCount:0,count:0};map.set(key,g);}g.count++;if(!g.name)g.name=row[name]||'';if(!g.cat)g.cat=row[cat]||'';const a=diagNumber(row[gmv]),b=profit<0?null:diagNumber(row[profit]);if(a!==null){g.gmv+=a;g.gmvCount++;}if(b!==null){g.profit+=b;g.profitCount++;}}
  for(const g of map.values()){if(!g.gmvCount)g.gmv=null;else if(g.gmvCount<g.count)quality.partialCells.push({sheet:sheet.name,id:g.id,field:'成交金额',missing:g.count-g.gmvCount});if(!g.profitCount)g.profit=null;else if(g.profitCount<g.count)quality.partialCells.push({sheet:sheet.name,id:g.id,field:sheet.profitField,missing:g.count-g.profitCount});if(g.gmv!==null)g.gmv=Math.round(g.gmv*100)/100;if(g.profit!==null)g.profit=Math.round(g.profit*100)/100;}return map;
 }
 const prev=effect(sheets.previous),cur=effect(sheets.current),search=new Map();const idIndex=diagnosisIndex(sheets.search,['SPU_ID'],true),nameIndex=diagnosisIndex(sheets.search,['商品名称']),catIndex=diagnosisIndex(sheets.search,['三级类目','三级品类']);
 for(const row of sheets.search.rows){const id=diagId(row[idIndex]);if(!id){quality.excluded.push({sheet:sheets.search.name,id:String(row[idIndex]??''),reason:'无法定位有效SPU'});continue;}if(search.has(id)){quality.duplicateSearch.push(id);continue;}const value={id,name:row[nameIndex]||'',cat:row[catIndex]||'',raw:{}};for(const prefix of ['搜索','推荐','全量'])for(const m of ['曝光','点击','CTR','CVR','GMV','订单','CPC','ROI']){const k=prefix+'_'+m;value.raw[k]=diagNumber(row[diagnosisIndex(sheets.search,[k])]);}search.set(id,value);}
 // A duplicate SPU in a user-provided summary is not silently averaged: suppress its metrics.
 for(const id of new Set(quality.duplicateSearch)){const s=search.get(id);for(const k of Object.keys(s.raw))s.raw[k]=null;}
 const records=[];for(const id of new Set([...search.keys(),...prev.keys(),...cur.keys()])){
  const a=prev.get(id),b=cur.get(id),s=search.get(id);const r={id,name:s?.name||b?.name||a?.name||'',category:s?.cat||b?.cat||a?.cat||'未提供三级类目',raw:s?.raw||{},notes:[],previousGMV:a?a.gmv:0,currentGMV:b?b.gmv:0,previousProfit:a?.profit??null,currentProfit:b?.profit??null};
  if(!a){quality.missingPrevious.push(id);r.notes.push('对比周期成交表无记录，前14天GMV按0');}if(!b){quality.missingCurrent.push(id);r.notes.push('现周期成交表无记录，近14天GMV按0参与停投筛查；未推断下架或不可售');}
  if(a&&a.gmv===null){quality.missingPreviousGMV.push(id);r.notes.push('对比周期成交金额为空');}if(b&&b.gmv===null){quality.missingCurrentGMV.push(id);r.notes.push('现周期成交金额为空');}
  r.gmv28=r.previousGMV===null||r.currentGMV===null?null:Math.round((r.previousGMV+r.currentGMV)*100)/100;
  r.growth=r.previousGMV===null||r.currentGMV===null?null:r.currentGMV>r.previousGMV;
  r.growthLabel=r.growth===null?'待补充':r.growth?'增长':'不增长';r.remark=r.growth===null?'待补充':r.growth?'不需要调':'需调整';
  r.growthRate=r.previousGMV>0&&r.currentGMV!==null?(r.currentGMV-r.previousGMV)/r.previousGMV:null;
  if(r.previousGMV===0&&r.currentGMV>0)r.notes.push('前周期为0，现周期增长');
  r.loss=r.currentProfit===null?null:r.currentProfit<0;r.profitLabel=r.loss===null?'待补充':r.loss?'亏损':'不亏损';if(r.currentProfit===null)quality.missingProfit.push(id);
  function rate(k){const n=r.raw[k]??null;if(n!==null&&n<0){quality.invalidRates.push({id,field:k,value:n});return null;}return n;}
  r.searchCTR=rate('搜索_CTR');r.searchCVR=rate('搜索_CVR');r.recommendCTR=rate('推荐_CTR');r.recommendCVR=rate('推荐_CVR');r.totalCTR=rate('全量_CTR');r.totalCVR=rate('全量_CVR');
  const sc=r.raw['搜索_点击']??null,rc=r.raw['推荐_点击']??null,tc=r.raw['全量_点击']??null;
  r.totalClicks=sc!==null&&rc!==null?sc+rc:tc;r.recommendShare=null;
  if(sc!==null&&rc!==null&&tc!==null&&Math.abs(sc+rc-tc)>1e-8)quality.sumMismatches.push({id,field:'点击',supplied:tc,calculated:sc+rc});
  if(r.totalClicks===0){r.recommendShare=0;quality.zeroClicks.push(id);r.notes.push('全量点击为0，推荐点击占比按0降级');}
  else if(r.totalClicks>0&&rc!==null&&rc>=0&&(sc===null||sc>=0)&&rc<=r.totalClicks)r.recommendShare=rc/r.totalClicks;
  else if((sc!==null&&sc<0)||(rc!==null&&rc<0)||(r.totalClicks!==null&&r.totalClicks<0)){quality.negativeClicks.push(id);r.notes.push('点击差额为负，占比不作有效决策依据');}
  const complete=[r.searchCTR,r.searchCVR,r.recommendCTR,r.recommendCVR].every(x=>x!==null);
  if(!complete)quality.missingDimensions.push(id);
  r.searchDominant=complete?(r.searchCVR+1e-12>=r.recommendCVR*1.5&&r.searchCTR+1e-12>=.02&&r.searchCTR+1e-12>=r.recommendCTR*1.2):null;
  r.bigWord=options.bigWordCategories?.includes(r.category)||options.allCategoriesHaveBigWords===true;
  r.bigWordLabel=r.bigWord?'有大词（用户确认）':'大词前提待确认';r.placement=r.searchDominant===null?'待补充':r.searchDominant?(r.bigWord?'重点搜索主导':'搜索主导'):'非搜索主导';
  records.push(r);
 }
 const benchmarks=new Map();for(const r of records){if(r.category==='未提供三级类目')continue;if(!benchmarks.has(r.category))benchmarks.set(r.category,{category:r.category,ctr:[],cvr:[],products:0});const g=benchmarks.get(r.category);g.products++;if(r.totalCTR!==null)g.ctr.push(r.totalCTR);if(r.totalCVR!==null)g.cvr.push(r.totalCVR);}
 const mean=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;
 for(const g of benchmarks.values()){g.ctrMean=mean(g.ctr);g.cvrMean=mean(g.cvr);g.ctrCount=g.ctr.length;g.cvrCount=g.cvr.length;delete g.ctr;delete g.cvr;}
 for(const r of records){const b=benchmarks.get(r.category);r.ctrMean=b?.ctrMean??null;r.cvrMean=b?.cvrMean??null;r.lowCTR=r.totalCTR===null||r.ctrMean===null?null:r.totalCTR<r.ctrMean-1e-12;r.lowCVR=r.totalCVR===null||r.cvrMean===null?null:r.totalCVR<r.cvrMean-1e-12;r.stopReason='';r.stop=false;r.stopUnresolved=false;
  if(r.growth===true)r.stopStatus='不适用：GMV增长，仅提供建议';
  else if(r.growth===null){r.stopStatus='待补充：GMV数据不足';r.stopUnresolved=true;}
  else if(r.currentGMV===0){r.stop=true;r.stopReason='近14天广告GMV＝0';r.stopStatus='建议停投';}
  else if(r.gmv28!==null&&r.gmv28<200&&r.loss===true){r.stop=true;r.stopReason='近28天广告GMV＜200元，且现周期投后亏损';r.stopStatus='建议停投';}
  else if(r.gmv28!==null&&r.gmv28<200&&r.loss===null){r.stopStatus='待补充：低GMV商品缺少现周期毛利，暂无法完成停投筛查';r.stopUnresolved=true;}
  else r.stopStatus='未命中停投条件';
  r.optimizations=[];r.roiAction='';r.budgetAction='';
  if(r.stop){r.group='停投';r.action='建议停投：'+r.stopReason;}
  else if(r.stopUnresolved){r.group='数据待补充';r.action='补齐GMV／指定毛利后完成停投筛查';}
  else if(r.lowCTR===true||r.lowCVR===true){r.group='先优化承接';if(r.lowCTR)r.optimizations.push('优化主图、白底图、标题、短标、价格、活动');if(r.lowCVR)r.optimizations.push('优化商详、价格、好评');r.action=r.optimizations.join('；');}
  else if(r.lowCTR===null||r.lowCVR===null||r.recommendShare===null){r.group='数据待补充';r.action='补齐有效CTR／CVR、三级类目基准或推荐点击占比后再给出ROI／预算建议';}
  else if(r.recommendShare<.7){r.group='ROI／预算调整';r.roiAction='调低ROI';r.budgetAction='维稳预算';r.action=r.roiAction+'＋'+r.budgetAction;}
  else if(r.loss!==null){r.group='ROI／预算调整';r.roiAction='拉高ROI';r.budgetAction=r.loss?'降低预算':'拉高预算';r.action=r.roiAction+'＋'+r.budgetAction;}
  else{r.group='数据待补充';r.action='推荐点击占比≥70%，补齐现周期指定毛利后决定预算方向';}
 }
 records.sort((a,b)=>DIAG_GROUPS.indexOf(a.group)-DIAG_GROUPS.indexOf(b.group)||(b.currentGMV??0)-(a.currentGMV??0)||a.id.localeCompare(b.id));
 const counts=Object.fromEntries(DIAG_GROUPS.map(g=>[g,records.filter(r=>r.group===g).length]));
 return {version:'1.0',meta,options,records,benchmarks:[...benchmarks.values()],quality,counts,summary:{products:records.length,growing:records.filter(r=>r.growth).length,notGrowing:records.filter(r=>r.growth===false).length,excluded:quality.excluded.length,currentGMV:records.reduce((n,r)=>n+(r.currentGMV??0),0),previousGMV:records.reduce((n,r)=>n+(r.previousGMV??0),0)},schemas:Object.values(sheets).map(s=>({name:s.name,rows:s.rows.length,columns:s.headers.length,headerRow:s.headerRow,fields:s.fields,profitField:s.profitField??null,profitBasis:s.profitBasis??null}))};
}
if(typeof module!=='undefined')module.exports={analyzeDiagnosis,diagnosisSheets,diagNumber,DIAG_GROUPS};

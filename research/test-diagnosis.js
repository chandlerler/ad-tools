'use strict';
const fs=require('fs'),assert=require('assert');global.XLSX=require('../vendor/xlsx.full.min.js');
const {analyzeDiagnosis}=require('../diagnosis.js'),{renderDiagnosisHTML}=require('../report.js');
const h=['SPU_ID','商品名称','三级类目',...['搜索','推荐','全量'].flatMap(p=>['曝光','点击','CTR','CVR','GMV','订单','CPC','ROI'].map(k=>p+'_'+k))];
const columns=['商品名称','商品ID','三级品类','成交金额','京喜预估投后履约毛利_折后'];
function fixture(items,profitHeader=columns[4]){const w=XLSX.utils.book_new(),search=[h],prev=[[...columns.slice(0,4),profitHeader]],cur=[[...columns.slice(0,4),profitHeader]];for(const r of items){const o={SPU_ID:r.id,商品名称:r.name??'商品'+r.id,三级类目:r.cat??'独立类目'+r.id};for(const p of ['搜索','推荐','全量'])Object.assign(o,{[p+'_曝光']:1000,[p+'_点击']:p==='搜索'?20:p==='推荐'?80:100,[p+'_CTR']:.04,[p+'_CVR']:.03,[p+'_GMV']:999999});Object.assign(o,r.fields||{});if(!r.noSearch)search.push(h.map(k=>o[k]??null));if(!r.noPrev)prev.push(['成交名称'+r.id,r.id,o.三级类目,r.prev??200,r.prevProfit??10]);if(!r.noCur)cur.push(['成交名称'+r.id,r.id,o.三级类目,r.cur??300,r.profit===undefined?20:r.profit]);}for(const [name,data] of [['广告搜索&推荐数据',search],['广告成交数据（对比周期）',prev],['广告成交数据（现期）',cur]])XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet(data),name);return w;}
const items=[
 {id:'1',prev:0,cur:50,profit:-10},
 {id:'2',prev:100,cur:100,profit:-10,fields:{'搜索_点击':31,'推荐_点击':69}},
 {id:'3',prev:100,cur:99,profit:-10},
 {id:'4',prev:100,cur:0,profit:-10},
 {id:'5',prev:100,cur:99,profit:0,fields:{'搜索_点击':30,'推荐_点击':70}},
 {id:'6',prev:20,cur:30,profit:1,fields:{'搜索_CVR':.03,'推荐_CVR':.02,'搜索_CTR':.02,'推荐_CTR':.02/1.2}},
 {id:'7',fields:{'搜索_CVR':.03,'推荐_CVR':.02,'搜索_CTR':.0199,'推荐_CTR':.01}},
 {id:'8',fields:{'推荐_CVR':null}},
 {id:'9',noCur:true},
 {id:'10',noSearch:true},
 {id:'11',fields:{'搜索_点击':0,'推荐_点击':0,'全量_点击':0,'全量_CTR':null,'全量_CVR':null}},
 {id:'12',cat:'共基准',cur:0,fields:{'全量_CTR':.01,'全量_CVR':.01}},
 {id:'13',cat:'共基准',fields:{'全量_CTR':.02,'全量_CVR':.02}},
 {id:'14',cat:'共基准',fields:{'全量_CTR':.03,'全量_CVR':.03}},
 {id:'15',cat:'共基准',fields:{'全量_CTR':.01,'全量_CVR':.01}},
 {id:'16',cat:'共基准',fields:{'全量_CTR':.03,'全量_CVR':.03}},
 {id:'17',cat:'共基准',noSearch:true},
 {id:'18',name:'',noPrev:true},
 {id:'19',profit:null,prev:100,cur:99},
 {id:'未匹配SKU:123',noPrev:true,noCur:true}
];
const result=analyzeDiagnosis(fixture(items),{}, {allCategoriesHaveBigWords:true});const by=new Map(result.records.map(r=>[r.id,r]));
assert.equal(result.records.length,19);assert.equal(result.summary.excluded,1);assert.equal(new Set(result.records.map(r=>r.id)).size,19);
assert.equal(by.get('1').growth,true);assert.equal(by.get('1').stopStatus,'不适用：GMV增长，仅提供建议');assert.equal(by.get('1').stopReason,'');assert.equal(by.get('1').group,'ROI／预算调整');assert.equal(by.get('1').action,'拉高ROI＋降低预算');assert.equal(by.get('1').growthRate,null);assert.equal(by.get('1').remark,'不需要调');
assert.equal(by.get('2').gmv28,200);assert.equal(by.get('2').growth,false);assert.equal(by.get('2').stopStatus,'未命中停投条件');assert.equal(by.get('2').action,'调低ROI＋维稳预算');
assert.equal(by.get('3').group,'停投');assert(by.get('3').stopReason.includes('＜200'));assert.equal(by.get('4').stopReason,'近14天广告GMV＝0');assert.equal(by.get('5').action,'拉高ROI＋拉高预算');
assert.equal(by.get('6').placement,'重点搜索主导');assert.equal(by.get('7').placement,'非搜索主导');assert.equal(by.get('8').placement,'待补充');assert.equal(by.get('8').raw['推荐_CVR'],null);
assert.equal(by.get('9').currentGMV,0);assert.equal(by.get('9').group,'停投');assert(by.get('9').notes.some(x=>x.includes('现周期成交表无记录')));assert.equal(by.get('10').totalCTR,null);assert.equal(by.get('10').group,'数据待补充');
assert.equal(by.get('11').recommendShare,0);assert(result.quality.zeroClicks.includes('11'));
const baseline=result.benchmarks.find(b=>b.category==='共基准');assert(Math.abs(baseline.ctrMean-.02)<1e-12);assert.equal(baseline.ctrCount,5);assert.equal(baseline.products,6);assert.equal(by.get('13').lowCTR,false);assert.equal(by.get('15').group,'先优化承接');assert.equal(by.get('15').optimizations.length,2);assert.equal(by.get('15').roiAction,'');
assert.equal(by.get('18').name,'成交名称18');assert.equal(by.get('18').growth,true);assert.equal(by.get('19').group,'数据待补充');
for(const r of result.records){if(r.growth){assert.equal(r.stop,false);assert.equal(r.stopReason,'');}if(r.group!=='ROI／预算调整')assert.equal(r.roiAction,'');}
const generic=fixture([{id:'99',prev:100,cur:99,profit:-1}],'预估投后履约毛利');assert.equal(analyzeDiagnosis(generic).records[0].group,'数据待补充');assert.equal(analyzeDiagnosis(generic,{}, {confirmGenericProfit:true}).records[0].group,'停投');
const totals=fixture([{id:'90',cur:100,prev:100}]);const s=totals.Sheets['广告成交数据（现期）'];XLSX.utils.sheet_add_aoa(s,[['另SKU','90','独立类目90',50,-5]],{origin:-1});const sum=analyzeDiagnosis(totals).records[0];assert.equal(sum.currentGMV,150);assert.equal(sum.currentProfit,15);assert.equal(sum.gmv28,250);
const highCVR=analyzeDiagnosis(fixture([{id:'98',fields:{'全量_CVR':1.5}}]));assert.equal(highCVR.benchmarks[0].cvrMean,1.5);
const unsafe=analyzeDiagnosis(fixture([{id:'100',name:'</script><script>window.PWNED=1</script>'}]));const html=renderDiagnosisHTML(unsafe);assert(!html.includes('</script><script>window.PWNED'));fs.writeFileSync('/tmp/diagnosis-xss-test.html',html);
assert.equal(Object.values(result.counts).reduce((a,b)=>a+b),result.records.length);
console.log('PASS diagnosis boundaries: growth gate, zero/200/70% thresholds, AND equality, simple means including stopped products, null metrics, absent periods, priority, profit alias, SPU uniqueness and HTML escaping.');

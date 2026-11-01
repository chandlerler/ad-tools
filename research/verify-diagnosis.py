"""Independently recompute diagnosis from the actual intermediate XLSX."""
from pathlib import Path
import json, math, statistics, re, collections
import openpyxl
root=Path(__file__).resolve().parents[1]
w=openpyxl.load_workbook(root/'deliverables/广告分析_20261003-20261030_左春泽.xlsx',data_only=True)
expected=json.loads((root/'research/diagnosis-validation.json').read_text())
def num(x):
 if x is None or x=='':return None
 if isinstance(x,str):return float(x.rstrip('%'))/(100 if x.endswith('%') else 1)
 return float(x)
def read_period(sheet):
 rows=list(sheet.values);h=rows[0];out={}
 for r in rows[1:]:
  id=str(r[h.index('商品ID')]);g=out.setdefault(id,{'gmv':[],'profit':[]})
  for key,col in [('gmv','成交金额'),('profit','预估投后履约毛利')]:
   n=num(r[h.index(col)])
   if n is not None:g[key].append(n)
 return {k:{f:round(sum(v),2) if v else None for f,v in g.items()} for k,g in out.items()}
p,c=read_period(w.worksheets[1]),read_period(w.worksheets[2]);rows=list(w.worksheets[0].values);h=rows[0];raw={};excluded=0
for r in rows[1:]:
 id=str(r[h.index('SPU_ID')])
 if not re.fullmatch(r'\d+',id):excluded+=1;continue
 raw[id]=dict(zip(h,r))
assert excluded==expected['summary']['excluded']==13
actual={r['id']:r for r in expected['records']};assert set(actual)==set(raw)|set(p)|set(c);assert len(actual)==82
rates=collections.defaultdict(lambda:{'CTR':[],'CVR':[]})
for id,row in raw.items():
 for metric in ['CTR','CVR']:
  n=num(row['全量_'+metric])
  if n is not None and n>=0:rates[row['三级类目']][metric].append(n)
means={cat:{k:statistics.fmean(v) if v else None for k,v in values.items()} for cat,values in rates.items()}
def close(a,b):
 if a is None or b is None:assert a is b,(a,b)
 else:assert math.isclose(a,b,rel_tol=1e-9,abs_tol=1e-9),(a,b)
counts=collections.Counter()
for id,r in actual.items():
 prev=p[id]['gmv'] if id in p else 0;cur=c[id]['gmv'] if id in c else 0;profit=c[id]['profit'] if id in c else None;total=None if prev is None or cur is None else round(prev+cur,2);growth=None if prev is None or cur is None else cur>prev
 close(prev,r['previousGMV']);close(cur,r['currentGMV']);close(total,r['gmv28']);close(profit,r['currentProfit']);assert growth==r['growth']
 row=raw.get(id,{});cat=row.get('三级类目');m=means.get(cat,{'CTR':None,'CVR':None});low={}
 for metric in ['CTR','CVR']:
  n=num(row.get('全量_'+metric));valid=n is not None and n>=0
  low[metric]=None if not valid or m[metric] is None else n<m[metric]-1e-12
  close(m[metric],r['ctrMean' if metric=='CTR' else 'cvrMean']);assert low[metric]==r['low'+metric]
 sc,rc,tc=[num(row.get(k)) for k in ['搜索_点击','推荐_点击','全量_点击']];tc=sc+rc if sc is not None and rc is not None else tc
 share=0 if tc==0 else rc/tc if tc is not None and tc>0 and rc is not None and 0<=rc<=tc and (sc is None or sc>=0) else None
 close(share,r['recommendShare']);stop=False;unresolved=False
 if growth is None:unresolved=True
 elif not growth:
  if cur==0:stop=True;assert r['stopReason']=='近14天广告GMV＝0'
  elif total<200 and profit is not None and profit<0:stop=True
  elif total<200 and profit is None:unresolved=True
 else:assert r['stopStatus']=='不适用：GMV增长，仅提供建议' and r['stopReason']==''
 if stop:group='停投'
 elif unresolved:group='数据待补充'
 elif any(v is True for v in low.values()):group='先优化承接'
 elif any(v is None for v in low.values()) or share is None:group='数据待补充'
 elif share<.7:group='ROI／预算调整';assert r['action']=='调低ROI＋维稳预算'
 elif profit is not None:group='ROI／预算调整';assert r['action']==('拉高ROI＋降低预算' if profit<0 else '拉高ROI＋拉高预算')
 else:group='数据待补充'
 assert r['group']==group,(id,r['group'],group);counts[group]+=1
 vals=[num(row.get(k)) for k in ['搜索_CTR','搜索_CVR','推荐_CTR','推荐_CVR']]
 if any(x is None or x<0 for x in vals):assert r['placement']=='待补充'
 else:
  sctr,scvr,rctr,rcvr=vals;dominant=scvr+1e-12>=rcvr*1.5 and sctr+1e-12>=.02 and sctr+1e-12>=rctr*1.2
  assert r['placement']==('重点搜索主导' if dominant else '非搜索主导')
assert dict(counts)==expected['counts']
for b in expected['benchmarks']:
 for k in ['CTR','CVR']:assert b[k.lower()+'Count']==len(rates[b['category']][k])
print('PASS independent diagnosis verification:',dict(counts),'across',len(actual),'SPUs; all growth, stop, priority, placement, click share and category baseline decisions agree.')

from pathlib import Path
import csv,io,zipfile,collections,openpyxl,math
root=Path('/workspace/ad-tools');owner='左春泽(zuochunze.1)'
paths=list(Path('/workspace/attachments').rglob('*21_09*.xlsx'));paths.sort(key=lambda p:p.name)
mapping=collections.defaultdict(set);effect=collections.defaultdict(lambda:[0.0]*5);expected_rows=[]
for p in paths:
 w=openpyxl.load_workbook(p,read_only=True,data_only=True);s=w.worksheets[1];s.reset_dimensions();it=iter(s.values);headers=next(it);rows=list(dict.fromkeys(it));selected=[]
 for row in rows:
  mapping[str(row[11])].add(str(row[9]))
  if row[0]!=owner:continue
  selected.append(row)
  for i,c in enumerate(['展现','点击','成交金额','成交父单量','消耗']):effect[str(row[9])][i]+=row[headers.index(c)] or 0
 expected_rows.append((headers,selected))
search=collections.defaultdict(lambda:[0.0]*5)
z=zipfile.ZipFile(next(Path('/workspace/attachments/67bad2de-e6ce-47c3-9721-1d1e55ee8091').glob('*.zip')))
for name in z.namelist():
 for row in csv.DictReader(io.TextIOWrapper(z.open(name),encoding='utf-8-sig')):
  sku=row['智能投放推广SKU ID'];ids=mapping[sku];key=next(iter(ids)) if len(ids)==1 else '未匹配SKU:'+sku
  for i,c in enumerate(['展现数','点击数','直接订单金额','直接订单行','花费']):search[key][i]+=float(row[c])
w=openpyxl.load_workbook(root/'deliverables/广告分析_20261003-20261030_左春泽.xlsx',data_only=True)
rows=list(w.worksheets[0].values);assert len(rows[0])==27
assert {r[0] for r in rows[1:]}==set(search)|set(effect)
def close(a,b):assert math.isclose(a,b,rel_tol=1e-8,abs_tol=.005),(a,b)
for row in rows[1:]:
 key=row[0];s=search.get(key,[0]*5)
 for j,i in [(3,0),(4,1),(7,2),(8,3)]:close(row[j],s[i])
 if key in effect:
  total=effect[key]
  for j,i in [(19,0),(20,1),(23,2),(24,3)]:close(row[j],total[i]);close(row[j-8],total[i]-s[i])
 else:assert all(v is None for v in row[11:])
 for base,v in [(3,s)]+([(11,[a-b for a,b in zip(effect[key],s)]),(19,effect[key])] if key in effect else []):
  for offset,num,den in [(2,1,0),(3,3,1),(6,4,1),(7,2,4)]:
   actual=row[base+offset]
   if v[den]>0:assert abs(actual-v[num]/v[den])<=.000050001,(key,base,offset,actual,v)
   else:assert actual is None
for sheet,(headers,selected) in zip(w.worksheets[1:],expected_rows):
 data=list(sheet.values);assert tuple(data[0])==headers;assert len(data)-1==len(selected)
 for actual,source in zip(data[1:],selected):
  for a,b in zip(actual,source):
   if isinstance(b,str) and b.endswith('%'):close(a,float(b[:-1])/100)
   elif isinstance(b,(int,float)):close(a,b)
   else:assert a==b,(a,b)
print('PASS independent Python verification: every SPU/SKU output, additive metric, derived ratio, and both selected source sheets.')
print('Output rows',len(rows)-1,'matched search SPUs',len([k for k in search if not k.startswith('未匹配SKU:')]),'unmatched',len([k for k in search if k.startswith('未匹配SKU:')]),'period rows',[len(x[1]) for x in expected_rows])

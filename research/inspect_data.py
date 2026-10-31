import csv,io,zipfile,pathlib,collections,json
z=zipfile.ZipFile(next(pathlib.Path('/workspace/attachments').rglob('*.zip')))
a=collections.defaultdict(lambda:collections.defaultdict(float)); words=collections.Counter(); count=0
for name in z.namelist():
 if not name.endswith('.csv'):continue
 for r in csv.DictReader(io.TextIOWrapper(z.open(name),encoding='utf-8-sig')):
  count+=1; channel='推荐' if r['搜索词'] in ('','-') else '搜索';words[(r['搜索词'] if channel=='推荐' else '有搜索词',r['投放位置'])]+=1
  key=r['SPU ID']+'|'+channel
  for f in ['展现数','点击数','花费','直接订单行','直接订单金额','总订单行','总订单金额']:
   a[key][f]+=float(r[f])
  if count==1:print('HEADERS',list(r),flush=True)
 print(name.split('_')[-1],count,flush=True)
print('COUNTS',words,'GROUPS',len(a))
print('FIRST',[(k,v) for k,v in a.items() if k.startswith('10036052079763')])
pathlib.Path('/workspace/ad-tools/research/aggregate.json').write_text(json.dumps(a,ensure_ascii=False))

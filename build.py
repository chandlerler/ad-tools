from pathlib import Path
root=Path(__file__).parent
worker='\n'.join((root/p).read_text() for p in ['vendor/xlsx.full.min.js','vendor/fflate.min.js','core.js','diagnosis.js','report.js'])
worker+=r"""
let pending=null,reportOptions={},inputNames=[];
self.onmessage=async e=>{
 try{
  let result;
  if(e.data.type==='owners'){
   if(!pending)throw Error('文件状态已失效，请重新生成');
   postMessage({progress:'按所选采销汇总并生成Excel…'});
   result=finishPrepared(pending,e.data.owners);pending=null;
  }else{
   reportOptions=e.data.options||{};inputNames=e.data.files.map(f=>f.name);
   pending=await prepareFiles(e.data.files,progress=>postMessage({progress}));
   if(pending.choices.length>1){
    postMessage({chooseOwners:pending.choices});return;
   }
   result=finishPrepared(pending,pending.choices.length?pending.choices.map(c=>c.name):null);pending=null;
  }
  postMessage({progress:'按规则生成商品诊断报告…'});
  const diagnosis=analyzeDiagnosis(result.wb,{...result.summary,inputs:inputNames,processWarnings:result.warnings},reportOptions);
  const reportHtml=renderDiagnosisHTML(diagnosis);
  const buffer=XLSX.write(result.wb,{type:'array',bookType:'xlsx',compression:true});
  postMessage({buffer,reportHtml,diagnosisSummary:diagnosis.summary,warnings:result.warnings,summary:result.summary},[buffer]);
 }catch(error){postMessage({error:error.message});}
};
"""
# Avoid accidentally terminating the embedded script block.
worker=worker.replace('</script','<\\/script')
out=root/'dist';out.mkdir(exist_ok=True)
(out/'广告数据工作台.html').write_text((root/'index.template.html').read_text().replace('__WORKER_CODE__',worker))
print(out/'广告数据工作台.html')

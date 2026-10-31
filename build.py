from pathlib import Path
root=Path(__file__).parent
worker='\n'.join((root/p).read_text() for p in ['vendor/xlsx.full.min.js','vendor/fflate.min.js','core.js'])
worker+='''\nself.onmessage=async e=>{try{const result=await processFiles(e.data.files,'estimate',progress=>postMessage({progress}));const buffer=XLSX.write(result.wb,{type:'array',bookType:'xlsx',compression:true});postMessage({buffer,warnings:result.warnings,summary:result.summary},[buffer]);}catch(error){postMessage({error:error.message});}};'''
# Avoid accidentally terminating the embedded script block.
worker=worker.replace('</script','<\\/script')
out=root/'dist';out.mkdir(exist_ok=True)
(out/'广告数据工作台.html').write_text((root/'index.template.html').read_text().replace('__WORKER_CODE__',worker))
print(out/'广告数据工作台.html')

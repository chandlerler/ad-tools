"""Package the offline tool and source without bundling uploaded source reports."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
root = Path(__file__).parent
out = root / 'deliverables'
out.mkdir(exist_ok=True)
with ZipFile(out / '广告数据工作台_Windows离线版.zip', 'w', ZIP_DEFLATED) as archive:
    archive.write(root / 'dist/广告数据工作台.html', '广告数据工作台/广告数据工作台.html')
    archive.write(root / 'README.md', '广告数据工作台/使用说明.md')
    for item in sorted((root / 'vendor').glob('LICENSE*')):
        archive.write(item, '广告数据工作台/许可证/' + item.name)
with ZipFile(out / '广告数据工作台_源代码.zip', 'w', ZIP_DEFLATED) as archive:
    for name in ['core.js', 'index.template.html', 'build.py', 'package.py', 'README.md']:
        archive.write(root / name, name)
    for item in sorted((root / 'vendor').iterdir()):
        archive.write(item, 'vendor/' + item.name)
    for pattern in ['*.js', '*.py']:
        for item in sorted((root / 'research').glob(pattern)):
            archive.write(item, 'research/' + item.name)
print('Packaged offline tool and source ZIPs.')

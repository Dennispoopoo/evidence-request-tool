"""把 src/ 的程式與 node_modules 內的函式庫合併成單一離線 HTML 檔，輸出到 dist/。
使用方式：先執行 npm install，再執行 python build.py
"""
import pathlib

root = pathlib.Path(__file__).parent
html = (root / 'src/index.html').read_text(encoding='utf-8')
parts = {
    '/*@@CSS@@*/': root / 'src/app.css',
    '/*@@JSZIP@@*/': root / 'node_modules/jszip/dist/jszip.min.js',
    '/*@@PDFJS@@*/': root / 'node_modules/pdfjs-dist/legacy/build/pdf.min.js',
    '/*@@PDFWORKER@@*/': root / 'node_modules/pdfjs-dist/legacy/build/pdf.worker.min.js',
    '/*@@APP@@*/': root / 'src/app.js',
}
for marker, path in parts.items():
    content = path.read_text(encoding='utf-8')
    assert marker in html, f'找不到標記 {marker}'
    if marker != '/*@@CSS@@*/':
        assert '</script' not in content.lower(), f'{path} 含有 </script，無法內嵌'
    html = html.replace(marker, content)

out = root / 'dist/聲請調查證據表格產生器.html'
out.parent.mkdir(exist_ok=True)
out.write_text(html, encoding='utf-8')
print(f'已產生 {out}（{len(html.encode("utf-8")):,} bytes）')

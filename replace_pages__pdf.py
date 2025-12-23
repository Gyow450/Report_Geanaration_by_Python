"""批量替换pdf中的固定页"""
#!/usr/bin/env python3
from pathlib import Path
from pypdf import PdfReader, PdfWriter

root   = Path(__file__).parent
inp    = root / 'pdf_in'
repl   = PdfReader(root / 'replacement.pdf')

for idx, pdf in enumerate(sorted(inp.glob('*.pdf'))):
    old = PdfReader(pdf)
    w   = PdfWriter()
    w.add_page(repl.pages[idx])          # 新首页
    w.add_pages(old.pages[1:])           # 剩余旧页
    with pdf.open('wb') as f:
        w.write(f)
print('done')
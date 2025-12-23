import fitz, io
from pathlib import Path
from PIL import Image
from fastprogress import progress_bar

src = Path(r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\1400街道网格管线分布图')
pdfs:list[Path] = list(src.rglob('*.pdf'))
for pdf in progress_bar(pdfs):
    pix = fitz.open(pdf).load_page(0).get_pixmap(dpi=200, alpha=False, colorspace=fitz.csRGB)
    Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB').save(
        pdf.with_suffix('.jpg'), format='JPEG', quality=85, dpi=(200, 200)
    )
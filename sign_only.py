from mypackage import r_generator as rg
import win32com.client as win32
from pathlib import Path
from fastprogress import progress_bar

def sign_by_pic_name(doc,sign_path:Path|str):
    """单个docx文件里的签名图按照人名索引替换"""
    sign_path =Path(sign_path)
    name_list:list[str] = [n.stem for n in sign_path.glob('*') ]
    for shape in doc.InlineShapes:
        tag:str = shape.Title
        if tag in name_list:
            for ex_n in ('.jpg','.png','.jpeg'):
                p_name= (sign_path/tag).with_suffix(ex_n)
                if p_name.exists():
                    rg.replace_pictue(doc,str(p_name),shape)
                    break
        
if __name__ =='__main__':
    input_dir:Path=Path(r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\管网PE第一批')
    sign_path:Path=Path(r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\电子签名')
    docx_list:list[Path]=[p for p in input_dir.glob('*.docx') if (not p.name.startswith('~$') and not p.name.startswith('error'))]
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    # 全局关闭拼写/语法检查
    word.Options.CheckSpellingAsYouType = False   # 关闭实时拼写检查
    word.Options.CheckGrammarAsYouType = False    # 关闭实时语法检查
    word.Options.ContextualSpeller = False        # 关闭上下文拼写检查（Word 2010+）
    for docx_path in progress_bar(docx_list):
        doc=word.Documents.Open(str(docx_path))
        sign_by_pic_name(doc,sign_path)
        doc.Save()
        doc.Close(SaveChanges=False)
    word.Quit()
    # print('签名完成')
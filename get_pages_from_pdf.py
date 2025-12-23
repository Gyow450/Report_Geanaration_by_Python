#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
输入源文件所在文件夹、保存文件、页码，生成一个合并pdf文件
"""
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from mypackage import interraction_terminal
from fastprogress import progress_bar

def analyze_words(input_str:str)->list[int]:
    words =input_str.split(',')
    pages_list:list[int]=[]
    for word in words:
        if '~' in word:
            first_pages,last_pages = word.split('~')
            nums =[num for num  in range(int(first_pages),int(last_pages)+1)]
            pages_list += nums
        else:
            pages_list.append(int(word))
    lst = sorted(set(pages_list))
    a_list,b_list =[x for x in lst if x>0],[y for y in lst if y<0]
    return a_list+b_list

def get_pages():
    base_list=[
        (0,'数据源','',''),
        (1,'保存文件','pdf'),
        (4,'提取页面：可以输入±x或x~y，以“,”隔断',r'^(-?\d+(~(-?\d+))?(,|$))*$'),
        ]
    SETTING_DICT:dict[str,str] = interraction_terminal.set_argumments(base_list)
    INPUT_DIR = Path(SETTING_DICT['数据源'])
    OUTPUT_FILE = Path(SETTING_DICT['保存文件'])
    in_pages_str = SETTING_DICT['提取页面：可以输入±x或x~y，以“,”隔断']
    PAGES = analyze_words(in_pages_str)
    # OUTPUT_FILE= OUTPUT_DIR/'首页整合文件.pdf' 
    
    pdf_list = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdf_list:
        print("❌ 文件夹中没有 PDF 文件")
        return
    writer = PdfWriter()
    for pdf_path in progress_bar(pdf_list):
        try:
            reader = PdfReader(pdf_path)
            if reader.pages:                       # 非空
                for page in PAGES:
                    if page >= 0:
                        page +=-1
                    writer.add_page(reader.pages[page])
                    # print(f"已提取：{pdf_path.name}第{page}页")
        except Exception as e:
            print(f"跳过 {pdf_path.name} ：{e}")

    if not writer.pages:
        print("❌ 没有任何有效页面被添加")
        return

    with OUTPUT_FILE.open("wb") as f:
        writer.write(f)
    print(f"✅ 完成！→ {OUTPUT_FILE}  （共 {len(writer.pages)} 页）")
if __name__ == "__main__":
    get_pages()
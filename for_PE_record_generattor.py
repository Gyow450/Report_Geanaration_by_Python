"""
    用来生成PE管定检各种的原始记录
    默认的文件名为：“PE管开挖原始记录模版.docx”
    
"""
import openpyxl
#import datetime
#import re
import win32com.client as win32
#import math
import os
import r_generator as rg
from LOG_DATA import LOG_DICT


"""=========================编辑生成全部用于替换的列表索引文件replacements======================"""

def make_change_text_all(workbook ,row ,record_type):
    #   编辑抬头和结尾零星内容的替换文本

    sheet = workbook[record_type]
    log_dict = rg.get_col_in_sheet(sheet)  #获取表头索引
    replacements = list()
    
    replacements += rg.make_change_text_for_heading(sheet,row,record_type,log_dict)
    if record_type+'选项' in LOG_DICT:
        replacements += rg.make_change_text_for_option(sheet,row,record_type,log_dict)
    for ctrl_key in ['填表','填表1','填表2']:
        if record_type + ctrl_key in LOG_DICT:
            replacements += rg.make_change_text_for_table(sheet,row,record_type,log_dict,ctrl_key)

    return replacements

"""
========================执行替换=======================

"""
def do_replace(doc , replacements ):
    for target_text, replacement_text in replacements:
        rg.replace_text(doc, target_text, replacement_text )

def main():
    app_type = rg.check_office_installation()
    if app_type == None:
        print('未找到合适的应用以打开文档')

    path = os.getcwd()
    record_types = {'1':'宏观检查记录','2':'开挖检测记录','3':'穿、跨越检查记录'}
    record_key = input('请输入记录的种类：1-宏观检查记录、2-开挖检测记录、3-穿、跨越检查记录')
    record_type = record_types[record_key]
    doc_modle_path = path+'\\PE管'+record_type+'模板.docx'
    workbook = openpyxl.load_workbook("PE管定期检验数据汇总表.xlsx" )

    record_name = input('请输入需生成记录的编号：')

    sheet = workbook[record_type]
    rows = rg.get_rows_in_sheet(record_name , sheet)
    while len(rows) == 0:
        print('未查找到此记录编号')
        report_name = input('请确认记录编号：')
        rows = rg.get_rows_in_sheet(record_name , sheet )

    print('生成替换用文本')
    replacements = make_change_text_all(workbook ,rows[0] ,record_type)

    print('读取模板文件')
    word = win32.Dispatch("Word.Application")
    if app_type == "office":
        word = win32.Dispatch("Word.Application")
    elif app_type == "wps":
        word = win32.Dispatch("Kwps.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    doc = word.Documents.Open(doc_modle_path)
    print('替换内容')
    do_replace( doc , replacements )

##    print('替换图片')
##    image_path = path+'\\PE管定期检验数据汇总表-开挖检测记录-开挖检测记录_附件\\' + report_name +'.JPG'
##    # print(image_path)
##    insert_picture(doc ,  image_path , '+插入图片' )
    output_file = path + '\\' + record_name + '.docx'
    doc.SaveAs2(output_file)
    print(f"文档已保存为：{output_file}")

    doc.Close()
    word.Quit()

main()

from re import split
import re
import win32com.client as win32
from datetime import datetime
import pandas as pd
from src.mypackage.LOG_DATA_STEEL import LOG_DICT
from src.mypackage import interraction_terminal 
from src.mypackage import r_generator as rg

def replacements_for_all(data_all:dict)->dict[str,list[list[tuple]]]:
    f_dict={'开挖汇总':[],'开挖第一页':[],'开挖第二页':[]}
    for key,data_dict in data_all.items():
        #   填开挖汇总
        f_dict['开挖汇总'].append([
            ('+管道名称',data_dict['管道名称']),
            ('+管道规格',data_dict['管道规格']), 
            ('+检测日期',data_dict['检测日期'].strftime('%Y年%m月%d日')),
            ('+地表状况', data_dict['地表状况']), 
            ('+环境条件', data_dict['环境条件']), 
            ('+检验情况：', 
            f"""\
检验情况：
    1.开挖点坐标（{data_dict['探坑坐标 X']}，{data_dict['探坑坐标 Y']}）
    2.防腐层{data_dict['防腐层破损情况描述'].split('（')[0]}
    3.管道{data_dict['本体腐蚀情况情况描述']}""" 
            ),
            ('+检验结论：',f"防腐层外观评定为{data_dict['防腐层破损情况描述'].split('（')[-1].replace('）','')}",) 
        ])

        #   填开挖第一页
        
        temp_list=[]
        for key in ['管道名称','管道规格','检测管段','探坑编号','管道埋深','天气条件','近参比点位','防腐层类型','地形、地貌、地物描述','土壤颜色','结构']:
            temp_list.append((f"+{key}",data_dict[key]))
        for key,options in LOG_DICT['开挖勾选'].items():
            t_texts=[f"{key}（√）" if option in data_dict[key].split(',') else f"{key}（ ）" for option in options ]
            temp_list.append((f"+{key}",'、'.joint(t_texts)))
        f_dict['开挖第一页'].append(temp_list)

        #   填开挖第二页
        f_dict['开挖第二页'].append([
            ('+管道名称',data_dict['管道名称']),
            ('+管道规格',data_dict['管道规格']), 
            ('+检测日期',data_dict['检测日期'].strftime('%Y年%m月%d日')),
            ('+探坑位置', data_dict['探坑位置']), 
            ('+探坑编号', data_dict['探坑编号']), 
            ('+环境条件', data_dict['环境条件']), 
            ('+防腐层破损情况描述', data_dict['防腐层破损情况描述']), 
            ('+防腐层测厚设备名称及编号', data_dict['防腐层测厚设备名称及编号']), 
            ('+FC1L0', data_dict['FC1L0']), 
            ('+FC1L3', data_dict['FC1L3']), 
            ('+FC1L6', data_dict['FC1L6']), 
            ('+FC1L9', data_dict['FC1L9']), 
            ('+管道本体腐蚀情况描述', data_dict['本体腐蚀情况情况描述']), 
            ('+管子测厚用设备及编号', data_dict['管子测厚用设备及编号']), 
            ('+C1L0', data_dict['C1L0']), 
            ('+C1L3', data_dict['C1L3']), 
            ('+C1L6', data_dict['C1L6']), 
            ('+C1L9', data_dict['C1L9']), 
        ])
    return f_dict

def replace_text_in_table(doc,table,any_list:list[tuple],type:str)->None:
    """在确定的表格内部作替换"""
    for log_tuple in any_list:
        a =log_tuple[0]
        row,col = LOG_DICT[type][log_tuple[0]]
        cell = table.Cell(row,col)
        if len(log_tuple[1:]) ==1:      # 普通一对一替换
            if log_tuple[1] is None:
                cell.Range.text = '/'
            elif isinstance(log_tuple[1],datetime.datetime):
                cell.Range.text = log_tuple[1].strftime("%Y年%m月%d日")
            else:
                cell.Range.text = log_tuple[1]
        else:                           # 追加下划线的替换
            cell.Range.text = log_tuple[1]
            after_text = log_tuple[2]
            cell.Range.InsertAfter(after_text)
            start = cell.Range.End - len(after_text)-1
            end   = cell.Range.End
            doc.Range(start, end).Font.Underline = 1 # 1 = wdUnderlineSingle

def do_replace_in_son_report(doc,any_dict:dict[str,list[list[tuple]]]):
    """执行分项报告表格写入"""
    i:int = 0
    j:int = 0
    k:int = 0
    for table in doc.Tables:
        title_name:str = table.Title
        if title_name == '开挖汇总':
            replace_text_in_table(doc,table,any_dict['开挖汇总'][i],'开挖汇总')  
            n+=1
        elif title_name == '开挖第一页':
            replace_text_in_table(doc,table,any_dict['开挖第一页'][j],'开挖第一页')  
            i+=1
        elif title_name == '开挖第二页':
            replace_text_in_table(doc,table,any_dict['开挖第二页'][k],'开挖第二页')  
            j+=1
        else:
            pass

if __name__ == '__main__':
    # 1. 获取用户输入
    set_list:list[tuple[int,str,str|bool,str|bool]]=[
        (2,'模板文件','docx',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\钢管\管网\开挖模板.docx'),
        (0,'数据源所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\钢管\管网'),
        # (0,'签名图片所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\电子签名'),
        (0,'输出文件所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\钢管\管网\输出'),
        # (3,'是否生成概述段落',False,True),
        # (3,'是否写入管道清单',False,True),
        # (3,'是否写入管道路由图',False,True),
        # (3,'是否生成签字',False,False),    
        # (3,'是否转pdf',False,False),    
    ]
    CONFIG = interraction_terminal.set_argumments(set_list)
    
    # 2. 读取数据生成DF
    df=pd.read_excel(f"{CONFIG['数据源所在']}\原始数据.xlsx",sheet_name='开挖记录',index_col=0)
    dict_all=df.to_dict('index')
    # 3. 整理数据，生成替换索引
    
    replacements_dict:dict = {}
    replacements_dict|=replacements_for_all(dict_all)

    # 4. 将报告内容写入Word文档，并保存文档
    app_type = rg.check_office_installation()
    if app_type == None:
        print('未找到合适的应用以打开文档')
    print('读取模板文件')
    
    if app_type == "office":
        word = win32.Dispatch("Word.Application")
    elif app_type == "wps":
        word = win32.Dispatch("Kwps.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    word.Options.CheckSpellingAsYouType = False   # 关闭实时拼写检查
    word.Options.CheckGrammarAsYouType = False    # 关闭实时语法检查
    word.Options.ContextualSpeller = False        # 关闭上下文拼写检查（Word 2010+）
    doc_modle_path = f"{CONFIG['模板文件']}"
    doc = word.Documents.Open(doc_modle_path)
    #   扩张表格
    rg.copy_and_insert_report_bookmark(doc , '开挖直接检验报告', len(dict_all))
    #   填写表格内容
    do_replace_in_son_report(doc,replacements_dict)
    #   替换图片
    ...
#   改写为win32库
import openpyxl
import datetime
import re
import win32com.client as win32
import math
import os
# import random
from src.mypackage import r_generator as rg

#   输入：待查询字符，工作表，返回：所在行号字符串列表
def get_rows_in_sheet(report_name , sheet , col_num ='A'):
    """返回确定列中，符合行号的列表"""                                                        
    rows_in_sheet = []           #开始检查该工作表中此报告的重数                            
    for row in sheet[col_num]:                                                                  
          if row.value == report_name:
              rows_in_sheet.append(str(row.row))
    return rows_in_sheet

def get_col_in_sheet( sheet, row:str = '1'):
    """默认返回工作表第一行的 名称：列字母 字典"""
    log_dict:dict = {}
    for cell in sheet[1]:
        log_dict[cell.value] = cell.column_letter
    return log_dict

# 返回字典——查询的键：行号的列表
def get_dict_from_col(report_name , sheet ,col_num_list =['B']):
    """多重查询，返回的是字典{多行关键字的元组：行号}"""
    name_row_list = []
    name_set = set()
    name_rows_dict = {}     #存放最终结果
    rows = get_rows_in_sheet(report_name , sheet )

    for row in rows:

        mid_list = []
        for col in col_num_list:
            v = sheet[col+row].value    #第一层（目标列）
            mid_list.append( v )
        name_set.add( tuple(mid_list) )
        name_row_list.append(( tuple(mid_list) , row ))

    for name in name_set:
        name_rows_dict[name] = []
   

    for name,row in name_row_list:
        name_rows_dict[name].append( row )       


    sorted_dict = dict(sorted( name_rows_dict.items() ))  

    return sorted_dict
    
#   用于生成整段文本的函数（用于项目概况）
def get_target_string(worksheet , row_num , key_cols ,ctrl_num = 0 , begin_with = "" , shut_with = "，"  , end_with = "。"):
    sheet = worksheet
    rn = row_num                                 
    ts = key_cols
    bw = begin_with
    sw = shut_with 
    ew = end_with
    k=''
    f_str = ""
    for c in ts:
        v = sheet[c+rn].value
        if ctrl_num ==1:
            k = sheet[c+"1"].value
        else:
            k = ""
                
        if v == None or v == "/":
            f_str += (sw+k+bw+"不明")                              #输出"不明"               
        elif isinstance( v , int):
            if k == "长度":
                f_str += (sw+k+bw+str(v)+"m")                      #输出XXm
            elif k == "实际使用年限":
                f_str += (sw+k+bw+str(v)+"年")                     #输出XXXX年
        elif  isinstance( v ,datetime.datetime):
            f_str += (sw+k+bw+v.strftime("%Y年%m月%d日"))           #输出日期
        else:
            f_str += (sw+k+bw+ v)

    if re.match(r"[\u4e00-\u9fff0-9]" , f_str[::-1]) == None:
        f_str = f_str[:-1]
    f_str += ew                                           #修正首尾字符
    
    return f_str                                          #输出字符串

#   检查文本末尾，删去所有标点

def check_text(in_text):
    text = in_text
    while text[-1] in ['；', '，' ,'、','：','。'] :
        text = text[:-1]
    
    return text


#   这个函数输入：工作表，单个行号，文本索引来获取：一段文本字符串

def get_text_by_log(worksheet, row_num ,key_cols):
    sheet = worksheet
    rn =row_num
    result = ""
    for word in key_cols:
        if word[0].isupper():
            v = sheet[word+rn].value
            if v == None or v == "/":
                result += '不明'
            elif isinstance( v , int):
                if sheet[word+'1'].value == "长度":
                    result += (str(v)+"m")                      #输出XXm
                elif sheet[word+'1'].value == "实际使用年限":
                    result += (str(v)+"年")                     #输出XXXX年
            elif  isinstance( v ,datetime.datetime):
                result += v.strftime("%Y年%m月%d日" )            #日期格式
            else:
                result += v
        else:
            result += word

    return result

#   考虑多重时，获取替换的文本
def get_replace_text(report_name, sheet, log_sheet ,col_num =[ 'B' ], row = ''):

    new_text = ''
    if row == '':
        name_rows_dict = get_dict_from_col(report_name , sheet ,col_num)
      
        
        for name in name_rows_dict:
            new_text += get_text_by_log(sheet , name_rows_dict[name][0], log_sheet)
            new_text = check_text(new_text)
            new_text += '。'
    else:
       new_text += get_text_by_log(sheet , row , log_sheet)
       new_text = check_text(new_text)
       new_text += '。'

    return new_text

"""

==================================以上处理仅Excel中的值=====================================


"""
#   获取匹配文本所在页码
def get_text_page(doc,target_text):

    selection = doc.Application.Selection

    selection.Find.ClearFormatting()
    selection.Find.Replacement.ClearFormatting()
    selection.Find.MatchCase = False  # 忽略大小写
    selection.Find.MatchWholeWord = True  # 匹配整个单词
    selection.Find.MatchWildcards = False  # 不使用通配符
    selection.Find.Wrap = 1

    selection.Find.Execute(target_text)
    page_number = selection.Information(3)

    return page_number

    
#   替换里的文本
def replace_text(doc, target_text, replacement_text ,r = 1):
    # 清除之前的查找格式
    doc.Content.Find.ClearFormatting()
    doc.Content.Find.Replacement.ClearFormatting()

    # 执行查找和替换操作

    
    # 长文本处理
    max_length = 250
    old_text_len = len(target_text)
    new_text_len = len(replacement_text)

    if new_text_len < max_length:
        doc.Content.Find.Execute(
            FindText=target_text,
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=1,
            Format=False,
            ReplaceWith=replacement_text,
            Replace=r  # 替换所有匹配项或第一项
        )
    else:
        # 计算每次替换的片段长度
        segment_length = max_length - old_text_len
        segment_count = math.ceil(new_text_len / segment_length)

        for i in range(segment_count):
            if i < segment_count - 1:
                # 非最后一段，加上旧文本以便继续查找
                segment = replacement_text[i * segment_length:(i + 1) * segment_length] + target_text
            else:
                # 最后一段
                segment = replacement_text[i * segment_length:]

            doc.Content.Find.Execute(
                FindText=target_text,
                MatchCase=False,
                MatchWholeWord=False,
                MatchWildcards=False,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Forward=True,
                Wrap=1,
                Format=False,
                ReplaceWith=segment,
                Replace=r  # 替换所有匹配项或第一项
            )
 


 
#   扩张段落用函数

def copy_paragraph(doc, target_text ,times):
     """扩张段落"""
     insert_text ='\n' + target_text
   

     # 遍历文档中的所有段落
     for para in doc.Paragraphs:
        if target_text in para.Range.Text:
            # 找到包含目标文本的第一个段落

            # 获取目标段落的文本内容
            para_text = para.Range.Text

            # 找到目标文本在段落中的位置
            start_index = para_text.find(target_text)
            if start_index != -1:
                # 计算目标文本的结束位置
                end_index = start_index + len(target_text)

                # 创建一个 Range 对象，表示目标文本的位置
                target_range = para.Range
                target_range.Start = target_range.Start + start_index
                target_range.End = target_range.Start + len(target_text)

                # 在目标文本后插入新文本
                target_range.Text = target_text + insert_text * ( times - 1 )

                # 找到第一个符合条件的段落后直接退出循环
                break

#   扩张单个表格
                
def copy_and_insert_report(doc , target_text,times ,pages = 0):
    """扩张单（多）页表格"""
    selection = doc.Application.Selection
    selection.Find.Execute(target_text)
    print(target_text+'\t数量：'+str(times-1))
    # 获取目标段落所在的页码，使用整数值 3 表示 wdActiveEndPageNumber
    page_number = selection.Information(3)

    for _ in range(0,times-1):
        
        # 使用 GoTo 方法定位到目标页的起始位置
        target_page_start = doc.GoTo(1, 1, page_number).Start  # 1 表示 wdGoToPage，1 表示 wdGoToAbsolute

        # 使用 GoTo 方法定位到目标页的结束位置（即下一页的起始位置）
        target_page_end = doc.GoTo(1, 1, page_number + 1 + pages).Start

        # 获取目标页的内容范围
        target_range = doc.Range(target_page_start, target_page_end)

        # 复制目标页的内容
        target_range.Copy()
        new_page_range = doc.Range(target_page_end, target_page_end)
        new_page_range.Paste()
      
    # 删除控制用关键字
    replace_text(doc, target_text, '' , 2)
            

#   删除指定页

def delete_page_by_text(doc, target_text):
    """删除关键字所在的整个页"""

    selection = doc.Application.Selection
    selection.Find.Execute(target_text)
    # 获取目标段落所在的页码，使用整数值 3 表示 wdActiveEndPageNumber
    page_number = selection.Information(3)            
        
    # 使用 GoTo 方法定位到目标页的起始位置
    target_page_start = doc.GoTo(1, 1, page_number).Start  # 1 表示 wdGoToPage，1 表示 wdGoToAbsolute

    # 使用 GoTo 方法定位到目标页的结束位置（即下一页的起始位置）
    target_page_end = doc.GoTo(1, 1, page_number + 1).Start

    # 获取目标页的内容范围
    if target_text == '删除渗透报告':
        target_range = doc.Range(target_page_start, target_page_end-1)
    elif target_text == '删除磁粉报告':
        target_range = doc.Range(target_page_start, target_page_end)
    else:
        target_page_end = doc.Content.End
        target_range = doc.Range(target_page_end-6, target_page_end)
    # 复制目标页的内容
    target_range.Delete()
    

#   对所有内容实施多重替换

def multi_replace(workbook, doc ,report_name):
    sheet = workbook['资料审查']
    do_multi_replace(doc, sheet, report_name, '多重资料审查+')
    
    sheet = workbook['庭院钢管宏观检查']
    do_multi_replace(doc, sheet, report_name, '多重庭院检查+')
                
    sheet = workbook['立管宏观检查']
    do_multi_replace(doc, sheet, report_name, '多重立管检查+')

    sheet = workbook['泄漏检测']
    do_multi_replace_plus(doc, sheet, report_name, '多重泄漏评估+' ,['B','C'], '++' , 6)

    sheet = workbook['破损点检测']
    do_multi_replace_plus(doc, sheet, report_name, '多重破损+' ,['B','D'], '++' , 8)   

    sheet = workbook['阴保测试']
    do_multi_replace(doc, sheet, report_name, '多重阴保报告+')

    sheet = workbook['壁厚测定']
    do_multi_replace(doc, sheet, report_name, '多重壁厚测定+')

    sheet = workbook['开挖检测']
    do_multi_replace(doc, sheet, report_name, '&')

#   扩张所有表格

def expand_all_tables(workbook, doc, report_name):

    #   资料审查报告
    times = len(get_rows_in_sheet(report_name , workbook['资料审查'] ))
    if times>1:
        copy_and_insert_report(doc , '复制资料审查报告', times)
    else:
        replace_text(doc, '复制资料审查报告','')    


    #   庭院钢管宏观检查
    times = len(get_rows_in_sheet(report_name , workbook['庭院钢管宏观检查'] )) 
    if times>1:
        copy_and_insert_report(doc , '复制庭院钢管宏观检查报告', times)
    elif times==1:
        replace_text(doc, '复制庭院钢管宏观检查报告','')
    else:
        delete_page_by_text(doc,'复制庭院钢管宏观检查报告')


    #   立管宏观检查
    times =len(get_rows_in_sheet(report_name , workbook['立管宏观检查'] ))
    if times>1:
        copy_and_insert_report(doc , '复制立管宏观检查报告', times)
    elif times==1:
        replace_text(doc, '复制立管宏观检查报告','')
    else:
        delete_page_by_text(doc,'复制立管宏观检查报告')

    
    #   泄漏评估：需按照管道名称、管道位置扩张 
    times = len(get_dict_from_col(report_name , workbook['泄漏检测'], ['B','C']))
    if times>1:
        copy_and_insert_report(doc , '复制泄漏评估报告', times)
    else:
        replace_text(doc, '复制泄漏评估报告','')


    #   防腐层：需按照管道名称、检测管段扩张 
    times = len(get_dict_from_col(report_name , workbook['破损点检测'], ['B','D']))
    if times>1:
        copy_and_insert_report(doc , '复制破损评估报告', times)
    elif times==1:
        replace_text(doc, '复制破损评估报告','')  
    else:
        delete_page_by_text(doc,'复制破损评估报告')

    #   阴保有效性
    times = len(get_rows_in_sheet(report_name , workbook['阴保测试'] ))
    if times>1:
        copy_and_insert_report(doc , '复制阴保评估报告', times)
    elif times==1:
        replace_text(doc, '复制阴保评估报告','')  
    else:
        delete_page_by_text(doc,'复制阴保评估报告')

    #   壁厚测定
    times = len(get_rows_in_sheet(report_name , workbook['壁厚测定'] ))
    if times>1:
        copy_and_insert_report(doc , '复制壁厚测定报告', times)
    else:
        replace_text(doc, '复制壁厚测定报告','')  
       
    #   磁粉、渗透不复制，按需删除
    # if len(get_rows_in_sheet(report_name , workbook['磁粉检测'] )) == 0:
    #     delete_page_by_text(doc, '删除磁粉报告')
    # else:
    #     replace_text(doc, '删除磁粉报告','')  
       
    # if len(get_rows_in_sheet(report_name , workbook['渗透检测'] )) == 0:
    #     delete_page_by_text(doc, '删除渗透报告')
    # else:
    #     replace_text(doc, '删除渗透报告','')  
       
    #   开挖评估，需扩张两页
    times = len(get_rows_in_sheet(report_name , workbook['开挖检测'] ))
    if times>1:
        copy_and_insert_report(doc , '复制开挖报告', times ,1)
    else:
        replace_text(doc, '复制开挖报告','')  
       
    #   整理删除页面
    delete_page_by_text(doc, '待删除')

   
#   编辑产生固定替换的文本

def make_change_text_const(workbook ,report_name):
    #   编辑抬头和结尾零星内容的替换文本

    sheet = workbook['封面']
    rows = get_rows_in_sheet(report_name ,sheet)
    
    replacements = [
        ("封面+管道类型", sheet['C'+rows[0]].value),
        ("封面+管道名称", sheet['D'+rows[0]].value),
        ("封面+管道名称", sheet['D'+rows[0]].value),
        ("封面+管道位置", sheet['E'+rows[0]].value),
        ("封面+评估日期", sheet['F'+rows[0]].value),
        ("封面+使用单位", sheet['B'+rows[0]].value),
        ("封面+使用单位", sheet['B'+rows[0]].value),
        ("审核人名",   sheet['H'+rows[0]].value),
    ]
    
    sheet = workbook['资料审查']

    rows = get_rows_in_sheet(report_name ,sheet)
    
    replacements.append(("资料审查+记数", str(len(rows))))

    #   编辑资料审查内容
    for row in rows:
        new_p = get_replace_text(report_name , sheet , 
            [
            'C' , '，管道类别：' , 'B' , '，管道材质：' , 'T' , '，管道规格：' , 'P' , '，设计压力：' , 'L' ,
            '，防腐层材料' , 'U' ,'，长度：' , 'E' , '，设计单位：' , 'F' , '，安装单位：' , 'H'  , '，竣工验收日期：',
           'K' , '，投用日期：' , 'M' , '，实际使用年限：' ,  'O'
             ],
            ['B'] , row)
        replacements.append(("复制写入审查概况", new_p))
    
    new_p = get_replace_text(report_name, sheet, ['C','Y'] , ['C'])
    
    replacements.append(("写入审查问题",'通过资料审查发现：'+ new_p))
    
    #print('编辑宏观检查')
    sheet = workbook["庭院钢管宏观检查"] 
    if len(get_rows_in_sheet(report_name ,sheet)) > 0:
        new_p = get_replace_text(report_name, sheet, ['B','N'])
    else:
        new_p = ''
    sheet = workbook["立管宏观检查"]   
    new_p += get_replace_text(report_name, sheet, ['B','N'])
    
    replacements.append(("写入宏观检查",'通过宏观检查发现：'+ new_p))

    #print('编辑泄漏检测')
    sheet = workbook['泄漏检测']
    new_p = get_replace_text(report_name, sheet, ['B','H'] , [ 'B' ,  'C' ])
    replacements.append(('写入泄漏评估','通过泄漏评估发现：'+ new_p))

    #print('编辑破损检测')
    sheet= workbook['破损点检测']
    if len(get_rows_in_sheet(report_name ,sheet)) > 0:
        new_p = get_replace_text(report_name, sheet, ['M'] , [ 'B' , 'D' ])
        replacements.append(('写入防腐层破损评估', '通过破损点抽查发现：'+ new_p))
    else:
        replacements+=[('写入防腐层破损评估', ''),('外防腐层破损评估','')]

    #print('编辑阴保评估')
    sheet = workbook['阴保测试']
    if len(get_rows_in_sheet(report_name ,sheet)) > 0:
        new_p = get_replace_text(report_name, sheet, ['B','G'])
        replacements.append(('写入阴保评估','通过阴极保护评估发现：'+ new_p))
    else:
        replacements+=[('写入阴保评估','通过阴极保护评估发现：'+ new_p),('阴极保护评估','')]


    #print('编辑开挖检测')
    sheet = workbook['开挖检测']

    new_p = get_replace_text(report_name, sheet, ['B', 'G' , '开挖坑内管道防腐层' , 'W' , '。管道本体' ,'AL'])
    replacements.append(('写入开挖检测','通过开挖直接评估发现：'+ new_p))

    #print('编辑壁厚测定')
    sheet = workbook['壁厚测定']

    new_p = get_replace_text(report_name, sheet, ['G','I'])
    replacements.append(('写入壁厚测定','通过壁厚测定发现：'+ new_p))

    #print('编辑磁粉检测')
    # sheet = workbook['磁粉检测']
    # rows = get_rows_in_sheet(report_name ,sheet)
    # if len(get_rows_in_sheet(report_name ,sheet)) > 0:
    #     ##   文本
    #     new_p = get_replace_text(report_name, sheet, ['对','H','进行磁粉检测，结果为：','G'])
    #     replacements.append(('写入磁粉检测', new_p))

    #     ##   表格
    #     for row in get_rows_in_sheet(report_name ,sheet):
    #         for cell in sheet[1]:
    #             v = sheet[cell.column_letter+row].value
    #             if cell.value == None:
    #                 break
    #             if isinstance( v ,datetime.datetime):
    #                 replacements.append(('磁粉检测+'+cell.value, v.strftime("%Y年%m月%d日") ))
    #             else:
    #                 replacements.append(('磁粉检测+'+cell.value, v ))

    # else:
    #     replacements.append(('写入磁粉检测', ''))

    # #print('编辑渗透检测')
    # sheet = workbook['渗透检测']
    # if len(get_rows_in_sheet(report_name ,sheet)) > 0:
    #     ##   文本
    #     new_p = get_replace_text(report_name, sheet, ['对','H','进行渗透检测，结果为：','G'])
    #     replacements.append(('写入渗透检测', new_p))

    #     ##   表格
    #     for row in get_rows_in_sheet(report_name ,sheet):
    #         for cell in sheet[1]:
    #             v = sheet[cell.column_letter+row].value
    #             if cell.value == None:
    #                 break
    #             if isinstance( v ,datetime.datetime):
    #                 replacements.append(('渗透检测+'+cell.value, v.strftime("%Y年%m月%d日") ))
    #             else:
    #                 replacements.append(('渗透检测+'+cell.value, v ))
    # else:
    #     replacements.append(('写入渗透检测', ''))

    #   编辑结论内容

    ##   编辑评估主要问题的替换内容
    sheet = workbook['评估情况表']

    new_p = get_replace_text(report_name, sheet, ['通过对以上单项检测评估结果进行综合评定，该条管道主要存在的问题为：','H'])
    replacements.append(('写入评估主要问题' ,new_p))

    ##   安全管控措施

    new_p = get_replace_text(report_name, sheet, ['I'])
    replacements.append(('写入管控措施', new_p))

    ##  评估结果

    new_p = get_replace_text(report_name, sheet, ['G'] )
    replacements.append(('评估情况表+评估结果', new_p))

    return replacements

#   替换页面往后的文本

def replace_text_in_page(doc, page_number, find_text, replace_text):
    
    # 获取当前页面的开始位置
    target_page_start = doc.GoTo(1, 1, page_number).Start  # wdActiveEndPageNumber = 3

    # 移动到页面末尾
    target_page_end = doc.GoTo(1, 1, page_number ).End


    # 创建页面范围
    rng_page = doc.Range(target_page_start, target_page_end)

    # 设置查找选项
    rng_page.Find.ClearFormatting()
    rng_page.Find.Replacement.ClearFormatting()


    # 在当前页面范围内查找并替换文本
    rng_page.Find.Execute(
                    FindText=find_text,
                    MatchCase=False,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Forward=True,
                    Wrap=1,
                    Format=False,
                    ReplaceWith=replace_text,
                    Replace=1  # 替换所有匹配项或第一项
                )
      
    
    #print(f"在第 {page_number} 页中成功替换文本。")



#   多重替换用函数普通版
def do_multi_replace(doc, sheet, report_name , key_word):

    #   遍历数据所在行
    for row in get_rows_in_sheet(report_name ,sheet):

        #   由关键词头定位页数
        page_number = get_text_page(doc , key_word)
        #print('替换页数'+str(page_number))

        #   遍历表头
        for cell in sheet[1]:

            #   处理值类型的合法性，并按顺序匹配关键字写入
            if cell.value ==None:
                break
            v = sheet[cell.column_letter + row].value
            if v == None:
                replace_text(doc, key_word+cell.value, '/' )
            elif isinstance( v ,datetime.datetime):
                replace_text(doc, key_word+cell.value, v.strftime("%Y年%m月%d日") )
            elif isinstance( v ,int) or isinstance( v ,float):
                replace_text(doc, key_word+cell.value, str(v) )
            else:
                replace_text(doc, key_word+cell.value, v )

            #   需写入固定两次的
            if key_word+cell.value in (
                '多重庭院检查+管道位置',
                '&管道名称',
                '&管道规格',
                '&探坑编号',
                '&组别',
                '&评估日期',
                '&审核日期'
                ):
                if isinstance( v ,datetime.datetime):
                    replace_text(doc, key_word+cell.value, v.strftime("%Y年%m月%d日") )
                elif isinstance( v ,int) or isinstance( v ,float):
                    replace_text(doc, key_word+cell.value, str(v) )
                else:
                    replace_text(doc, key_word+cell.value, v )



#   多重替换用函数加强版，用于泄漏评估、破损评估
def do_multi_replace_plus(doc, sheet, report_name, key_word ,key_list , key_symbol , space_number  ):


    name_rows_dict = get_dict_from_col(report_name, sheet ,key_list)

    #   遍历所有（管道）名称
    for name in name_rows_dict:
               
        #   由关键词头定位页数
        page_number = get_text_page(doc , key_word)
        #print('替换页数'+str(page_number))

        #   只执行一遍
        row0 = name_rows_dict[name][0]
           
           
        for cell in sheet[1]:
            
            #   处理值类型的合法性，并按顺序匹配关键字写入
            if cell.value ==None:
                break
            v = sheet[cell.column_letter + row0].value
            if v == None:
                pass
            elif isinstance( v ,datetime.datetime):
                replace_text(doc, key_word+cell.value, v.strftime("%Y年%m月%d日") )
            elif isinstance( v ,int) or isinstance( v ,float):
                replace_text(doc, key_word+cell.value, str(v) )
            else:
                replace_text(doc, key_word+cell.value, v )

                

        count = 0
        for row in name_rows_dict[name]:
            count += 1
            for cell in sheet[1]:
                if cell.value ==None:
                    break
                v = sheet[cell.column_letter + row].value
                if v == None:
                    pass
                elif isinstance( v ,datetime.datetime):
                    replace_text(doc,  key_symbol+cell.value, v.strftime("%Y年%m月%d日") )
                elif isinstance( v ,int) or isinstance( v ,float):
                    replace_text(doc,  key_symbol+cell.value, str(v) )
                else:
                    replace_text(doc,  key_symbol+cell.value, v )
            

        #   遍历完每个name，再填入剩余空位space_number-count
        for _ in range(space_number - count):
            for cell in sheet[1]:
                if cell.value ==None:
                    break

                replace_text(doc, key_symbol+cell.value, '/' )

def do_replace_dig_pic(doc,report_name:str,path:str):
    image_name:str = f"{path}\\开挖记录\\{report_name}"
    ex_names ={'.jpg','.png'}
    for ex_name in ex_names:
        image_path = f"{image_name}{ex_name}"
        if os.path.exists(image_path):
            rg.insert_picture(doc,image_path,'开挖检测',120)
            break

def do_sign(doc,workbook,report_name,path)->None:
    """实施编制签名索引并替换图片"""
    sign_list:list[str] = []
    for book_name in ['评估情况表','资料审查','立管宏观检查','泄漏检测','破损点检测','阴保测试','壁厚测定','开挖检测']:
        sheet =workbook[book_name]
        log_dict = rg.get_col_in_sheet(sheet)
        rows = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
        sign_list.append(sheet[log_dict['组别']+rows[0]].value)
        sign_list.append('刘畅')
    i:int = 0
    for shape in doc.InlineShapes:
        if shape.Title == '签字':
            rg.replace_pictue(doc,f"{path}\\电子签名\\{sign_list[i]}.png",shape,40)
            i+=1

"""
========================处理文本框、勾选项、页眉等=================================                

"""


def other_things(workbook , doc , report_name):

    # 勾选框等等

    sheet = workbook['资料审查']
    log_dict = get_col_in_sheet(sheet)
    rows = get_rows_in_sheet(report_name , sheet )
    sum_lenth = 0 
    for row in rows:
        sum_lenth += sheet[log_dict['长度']+row].value
    



    sheet = workbook['评估情况表']
    rows = get_rows_in_sheet(report_name , sheet )
    log_dict = get_col_in_sheet(sheet)
    row = rows[0]
    replacements = [
        ('评估情况表+对象简述' ,  sheet[log_dict['对象简述'] + row].value ),
        ('评估情况表+介质类型' ,  sheet[log_dict['介质类型'] + row].value ),
        ('评估情况表+长度' ,  str( sum_lenth ) ),
        ('评估情况表+管材类别' ,  sheet[log_dict['管材类别']+ row].value ),

        ]

    v =  sheet[log_dict['评估结果'] + row].value
    if v == '符合安全运行要求':
        replacements += ( ('$$' ,'☑') , ('$$' ,'□'),('$$' ,'□') ,('$$' ,'□') )
    elif v == '落实安全管控措施，可继续运行':
        replacements += ( ('$$' ,'□') , ('$$' ,'☑'),('$$' ,'□') ,('$$' ,'□') )
    elif v == '限期改造':
        replacements += ( ('$$' ,'□') , ('$$' ,'□'),('$$' ,'☑') ,('$$' ,'□') )
    else:
        replacements += ( ('$$' ,'□') , ('$$' ,'□'),('$$' ,'□') ,('$$' ,'☑') )


    v = sheet[log_dict['主要问题'] + row].value
    if '材质落后' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
        
    if '使用年限较长' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
        
    if '腐蚀泄漏严重' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
    if '防腐状况较差' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
    if '建构筑物占压' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
    if '处于/临近地质灾害易发区域' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)
    if '处于/临近人员密集区' in v:
        replacements += (( '$$' ,'☑'),)
    else:
        replacements += (( '$$' ,'□'),)

    parts = v.split('：')  
    replacements += (( '$$' ,parts[1]),)

    for target_text, replacement_text in replacements:
        replace_text(doc, target_text, replacement_text)

    sheet = workbook['封面']
    rows = get_rows_in_sheet(report_name , sheet)
    row =rows[0]
    log_dict = get_col_in_sheet(sheet)
    #   文本框
    for shape in doc.Shapes:

        if shape.TextFrame.HasText:
            if '编制人名' in shape.TextFrame.TextRange.Text:
                # 替换文本框中的内容
                shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace('编制人名', sheet[log_dict['评估']+ row].value)
            if '审核人名' in shape.TextFrame.TextRange.Text:
                # 替换文本框中的内容
                shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace('审核人名', sheet[log_dict['审核']+ row].value)


    #   替换页眉等其他内容
    replace_text(doc, "审核人名", workbook['封面'][log_dict['审核']+rows[0]].value , 2)

    for section in doc.Sections:
        # 遍历节中的所有页眉
        for header in section.Headers:
            # 替换页眉中的文本
            if "报告编码" in header.Range.Text:
                header.Range.Find.ClearFormatting()
                header.Range.Find.Replacement.ClearFormatting()
                header.Range.Find.Execute(
                    FindText="报告编码",
                    MatchCase=False,
                    MatchWholeWord=True,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Forward=True,
                    Wrap=1,  # wdFindStop
                    Format=False,
                    ReplaceWith=report_name,
                    Replace=2  # wdReplaceAll
                )
                
"""

========================单体表格主函数=====================================


"""

def solo_main(word,doc, output_file:str ,workbook ,report_name:str,path:str):
    

    # 扩张段落
    print('扩张段落')
    times = len ( get_rows_in_sheet( report_name , workbook['资料审查']) )
    copy_paragraph(doc , '复制写入审查概况' ,times)
    
    # 扩张表格
    print('复制表格')
    expand_all_tables(workbook, doc, report_name)

    # 多重替换操作
    print('填写表格内容')
    multi_replace(workbook, doc ,report_name)
    
    #   获取固定替换的文本
    replacements = make_change_text_const(workbook ,report_name)
    print('替换文本')

    # 执行固定替换操作
    # rows = get_rows_in_sheet(report_name , workbook['封面'])


    for target_text, replacement_text in replacements:
        replace_text(doc, target_text, replacement_text)

    other_things(workbook , doc , report_name)

    # 替换开挖图片
    print('替换图片及签名')
    do_replace_dig_pic(doc,report_name,path)
    do_sign(doc,workbook,report_name,path.replace('\\成华区评估',''))
    

     # 移动到文档的末端
    selection = word.Selection
    selection.EndKey(6)  # 6 表示 wdStory，即整个文档
    # 更新所有域（页码）
    doc.Fields.Update()

    # 保存为新文件
    doc.SaveAs2(f"{output_file}.docx", FileFormat=16)  # 16 表示docx 17 表示 PDF
    # doc.SaveAs2(f"{output_file}.pdf", FileFormat=17) 
    
    print(f"文档已保存为：{output_file}")

    doc.Close(SaveChanges=False)



    

# 读取表格
# workbook = openpyxl.load_workbook("2025成华区评估数据.xls" )
# sheet = workbook['封面']
# report_names = []
# for row in sheet["A"]:
#     cell_value = row.value
#     if cell_value != "报告编号":
#         report_names.append(cell_value)

# report_name =report_names[7]    #   测试用



#solo_main('E:\\成渝特检\\模板.docx', 'E:\\成渝特检\\moedl.docx' ,workbook, report_name)


def main():
    path = 'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管'

    doc_modle_path = f"{path}\\成华区评估\\模板-成华25-pdf.docx"
    workbook = workbook = openpyxl.load_workbook(f"{path}\\成华区评估\\2025年成华区评估数据0730-910-修正01.xlsx" )

    # 获取所有文件名称列表
    all_reports = []
    for cell in workbook['封面']['A']:
        all_reports.append(cell.value)

    all_reports = all_reports[1:]
    
    # 启动 Word 应用程序
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息


    # for report_name in [all_reports[20],all_reports[41],all_reports[85]]:
    for report_name in all_reports:
    
        
        # 加载模板文档
        doc = word.Documents.Open(doc_modle_path)
        output_file = f"{path}\\成华区评估\\输出结果\\{report_name}"
        solo_main(word,doc, output_file ,workbook ,report_name,f"{path}\\成华区评估")

    

    # 关闭文档并退出wod

    word.Quit()

main()

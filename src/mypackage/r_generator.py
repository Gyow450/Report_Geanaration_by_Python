"""用于生成记录、报告的函数集"""
import math
import datetime
from openpyxl.workbook import Workbook
from openpyxl.cell.cell import Cell
import win32com.client as win32
import os
import winreg
from .LOG_DATA import LOG_DICT


def check_office_installation()->str|None:
    """检查环境是否安装了office或者WPS"""
    try:
        # 检查 Microsoft Office 是否安装
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
            office_version = winreg.QueryValue(key, "")
            print(f"\nMicrosoft Office 已安装，版本: {office_version}")
            return "office"
    except FileNotFoundError:
        pass

    try:
        # 检查 WPS 是否安装
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Kwps.Application\CurVer") as key:
            wps_version = winreg.QueryValue(key, "")
            print(f"\nWPS 已安装，版本: {wps_version}")
            return "wps"
    except FileNotFoundError:
        pass

    print("未检测到 Microsoft Office 或 WPS 的安装。")
    return None

def replace_text_in_page(word,doc,target_text:str, replacement_text:str|None|datetime.datetime|int|float ,page_number:int,while_none:str = '/')->None:
    
    # 清除之前的查找格式
    doc.Content.Find.ClearFormatting()
    doc.Content.Find.Replacement.ClearFormatting()
    # 检查输入值格式
    if replacement_text == None:
        replacement_text = while_none
    elif  isinstance(replacement_text,datetime.datetime):
        replacement_text = replacement_text.strftime("%Y年%m月%d日")
    elif not isinstance(replacement_text, str):
        replacement_text = str(replacement_text)

    selection = word.Selection
    selection.GoTo(What=1, Which=1, Count=page_number)  # wdGoToPage=2, wdGoToAbsolute=1
    # 获取页面的开始位置
    page_start = selection.Start
    # 移动到页面的下一页开始位置
    selection.GoTo(What=1, Which=1, Count=page_number + 1)  # wdGoToPage=2, wdGoToNext=2
    # 获取页面的结束位置
    page_end = selection.Start
    # 在页面范围内进行查找替换
    page_range = doc.Range(Start=page_start, End=page_end)
    
    finder = page_range.Find
    finder.Text = target_text
    finder.Replacement.Text = replacement_text
    finder.Wrap = 0
    finder.Forward = True
   
    finder.Execute(
        Replace=2
    )


def replace_text(doc, target_text:str, replacement_text:str|None|datetime.datetime|int|float ,r = 1,while_none = '/')->None:
    """实际执行单个替换，r为控制码，默认1替换首个，2全部替换。主要耗时点"""
    # 清除之前的查找格式
    doc.Content.Find.ClearFormatting()
    doc.Content.Find.Replacement.ClearFormatting()

    # 检查输入值格式
    if replacement_text == None:
        replacement_text = while_none
    if  isinstance(replacement_text,datetime.datetime):
        replacement_text = replacement_text.strftime("%Y年%m月%d日")
    elif not isinstance(replacement_text, str):
        replacement_text = str(replacement_text)

    # 长文本处理
    max_length = 250
    old_text_len = len(target_text)
    new_text_len = len(replacement_text)
    
    # 执行查找和替换操作
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


    
def get_rows_in_sheet(report_name:str , sheet:Workbook, col_num ='A',type:str='str'):
    """实际是查找某一列中关键字，返回行号列表,type字段控制格式（str或int）"""                                                      
    rows_in_sheet:list[str|int] = []
    if type == 'str':           #开始检查该工作表中此报告的重数                            
        for row in sheet[col_num]:                                                                  
            if row.value == report_name:
                rows_in_sheet.append(str(row.row))
    elif type == 'int':
        for row in sheet[col_num]:                                                                  
            if row.value == report_name:
                rows_in_sheet.append(row.row)

    return rows_in_sheet


def get_col_in_sheet( sheet:Workbook , row:str = '1',type:str ='str')->dict[str,str]:
    """默认返回工作表第一行的 名称：列字母 字典"""
    log_dict:dict[str:str] = {}
    if type == 'str':
        for cell in sheet[1]:
            log_dict[cell.value] = cell.column_letter
    elif type == 'int':
        for cell in sheet[1]:
            log_dict[cell.value] = cell.column
    return log_dict

def replace_pictue(doc,image_path:str,shape,max_size:int =40)->None:
    if os.path.exists(image_path):   
        # 替换为实际图片
        range_to_replace = shape.Range
        # 删除原图
        shape.Delete()
        # 在原图的位置插入新图
        shape = doc.InlineShapes.AddPicture(
            FileName=image_path,
            LinkToFile=0,  # False 的值为 0
            SaveWithDocument=1,  # True 的值为 1
            Range=range_to_replace
        )

        # 获取原图片的宽度和高度
        original_width = shape.Width
        original_height = shape.Height

        # 计算新的宽度和高度，确保最大长宽，同时保持长宽比
        
        if original_width > original_height:
            # 如果宽度大于高度，以宽度为基准调整
            new_width = max_size
            new_height = original_height * (max_size / original_width)
        else:
            # 如果高度大于宽度，以高度为基准调整
            new_height = max_size
            new_width = original_width * (max_size / original_height)

        # 设置图片的新宽度和高度
        shape.Width = new_width
        shape.Height = new_height
        # 只替换一张图
     
def insert_picture(doc, image_path:str, keyword:str,max_size:int=40)->None:
    """插入图片"""
    if os.path.exists(image_path):   
        for shape in doc.InlineShapes:
            if shape.Type == 3:
                #检查是否是占位图片
                if shape.Title== keyword:
                    # 替换为实际图片
                    range_to_replace = shape.Range
                    # 删除原图
                    shape.Delete()
                    # 在原图的位置插入新图
                    shape = doc.InlineShapes.AddPicture(
                        FileName=image_path,
                        LinkToFile=0,  # False 的值为 0
                        SaveWithDocument=1,  # True 的值为 1
                        Range=range_to_replace
                    )

                    # 获取原图片的宽度和高度
                    original_width = shape.Width
                    original_height = shape.Height

                    # 计算新的宽度和高度，确保最大长宽，同时保持长宽比
                    
                    if original_width > original_height:
                        # 如果宽度大于高度，以宽度为基准调整
                        new_width = max_size
                        new_height = original_height * (max_size / original_width)
                    else:
                        # 如果高度大于宽度，以高度为基准调整
                        new_height = max_size
                        new_width = original_width * (max_size / original_height)

                    # 设置图片的新宽度和高度
                    shape.Width = new_width
                    shape.Height = new_height
                    # 只替换一张图
                    break
    else:
        pass



def copy_paragraph(doc, target_text:str ,times:int,insert_text:str = '\n复制写入概况'):
     """复制段落"""
     
   

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

def copy_and_insert_paragraph(doc,target_text:str,insert_text_list:list[str] ):
     """复制段落并插入段落"""
     
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
                new_text = ''
                for any_text in insert_text_list:
                    new_text += f"{any_text}\n"
                # 在目标文本后插入新文本
                new_text = new_text[:-1]
                target_range.Text = new_text
                # 找到第一个符合条件的段落后直接退出循环
                break
#   扩张单个表格页
def copy_and_insert_report(doc , target_text:str,times:int ,pages:int = 0):
    """复制多个整页页"""
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
    
    # # 使用 GoTo 方法定位到目标页的起始位置
    # target_page_start = doc.GoTo(1, 1, page_number).Start  # 1 表示 wdGoToPage，1 表示 wdGoToAbsolute

    # # 使用 GoTo 方法定位到目标页的结束位置（即下一页的起始位置）
    # target_page_end = doc.GoTo(1, 1, page_number + 1 + pages).Start

    # # 获取目标页的内容范围
    # target_range = doc.Range(target_page_start, target_page_end)

    # # 复制目标页的内容
    # target_range.Copy()
    # new_page_range = doc.Range(target_page_end, target_page_end)
    # for _ in range(times-1):
    #     new_page_range.Paste()
    #     new_page_range = doc.Range(target_page_end, target_page_end)


    # 删除控制用关键字
    replace_text(doc, target_text, '' , 2)

#   编辑表头等替换文本
def make_change_text_for_heading(sheet:Workbook,row:str,record_type:str,log_dict:dict)->list[tuple]:
    replacements:list[tuple] = []
    for target_text,replace_text_log in LOG_DICT[record_type]:
        replace_text = sheet[log_dict[replace_text_log] + row].value
        replacements.append( (target_text,replace_text) )
    return replacements

#   勾选框替换
def make_change_text_for_option(sheet:Workbook,row:str,record_type:str,log_dict):
    """生成记录、报告的内容勾选框部分"""
    replacements:list[tuple] = []
    for key_text,option_tuple in LOG_DICT[record_type+'选项'].items():
        text_for_option:str = ''        # 需要编辑的中间文本
        str_value:str|None = sheet[log_dict[key_text] + row].value
        if str_value:
            value_list = str_value.split(', ')   # 实际内容的列表
        else:
            value_list = []
        for match_value in option_tuple:
            if match_value in value_list:
                text_for_option += f"☑{match_value}、"
            else:
                text_for_option += f"□{match_value}、"
        #   
        rest_set:set = set(value_list)-set(option_tuple)
        if '其他' in option_tuple and rest_set:
            text_for_option=text_for_option.replace('□其他','☑其他')
        other_text:str=''       #   输出用文本
        if '其他：' in option_tuple:        #   存在需要补充可能
            other_word:str|None = sheet[log_dict['其他'+key_text] + row].value
            if not rest_set and not other_word:
                other_text = '□其他：/'
            elif rest_set:
                first,*other = rest_set
                other_text = f"☑其他：{first}"
            else:
                other_text = f"☑其他：{other_word}"
            text_for_option=text_for_option.replace('□其他：',other_text)
        plus_text = check_text(text_for_option,'：')
        if key_text in {'热熔连接','电熔连接','法兰连接','钢塑转换接头'}:
            v = sheet[log_dict[f"{key_text}缺陷描述"] + row].value
            if v is None or v =='':
                v='/'
            plus_text = f"{plus_text}{v}"
        
        replacements+= [('+'+key_text,plus_text),]
    return replacements

#   列表内容替换
def make_change_text_for_table(sheet:Workbook,row:str,record_type:str,log_dict:dict,ctrl_key:str='填表')->list:
    replacements:list = []
    rows = get_rows_in_sheet(sheet['A'+row].value, sheet , log_dict['所属记录编号'])
    count = 0
    for row_i in rows:
        count += 1
        replacements.append( ('&号' , count ))
        for key_text,replace_text_log in LOG_DICT[record_type+ctrl_key]:
            replace_text = sheet[log_dict[replace_text_log] + row_i].value
            replacements.append( (key_text,replace_text) )

    for _ in range(LOG_DICT[record_type+ctrl_key+'记数']-count):
        replacements.append( ('&号' ,'/'))
        for key_text,replace_text_log in LOG_DICT[record_type+ctrl_key]:
            replace_text = '/'
            replacements.append( (key_text,replace_text) )
    return replacements

#   操作表格添加行
def add_row_to_table(doc,k_word:str,rows_count:int):
    """表格添加空行"""
    for table in doc.Tables:
        first_cell = table.Cell(1,1)
        fist_cell_text = first_cell.Range.Text.strip()
        if k_word in fist_cell_text:
            for _ in range(rows_count):
                table.Rows.Add()
                # print('添加一行')
            break

#   写入表格
def write_in_table(doc,k_word:str,sheet,rows:list[str])->None:
    """将工作表的内容写入表格"""
    for table in doc.Tables:
        first_cell = table.Cell(1,1)
        fist_cell_text = first_cell.Range.Text.strip()
        if k_word in fist_cell_text:
            i=1
            for row in rows:
                i += 1
                for j in range(18):
                    v:str|None|datetime.datetime|int|float = sheet.cell(int(row),j+3).value
                    if v is None or v=='#VALUE!' or v=='' or v ==' ':
                        v='不明'
                    elif v == '　':
                        v='不明'
                    elif  isinstance(v,datetime.datetime):
                        v = v.strftime("%Y年%m月%d日")
                    elif not isinstance(v, str):
                        v = str(v)                 
                    table.Cell(i,j+1).Range.Text = v
            # for row_ in table.Rows:
            #     row_.HeightRule = 1  # 1 表示 wdRowHeightAuto
            #     row_.Height = 0  # 自动调整行高
            break



#   这个函数输入：工作表，单个行号，文本索引来获取：一段文本字符串
def get_text_by_log(worksheet:Workbook, row_num:str ,key_cols:list[str])->str:
    """这个函数输入：工作表，单个行号，文本索引来获取：一段文本字符串"""
    sheet = worksheet
    rn =row_num
    result = ""
    for word in key_cols:
        if word[0].isupper():
            v:str|None|datetime.datetime|int|float = sheet[word+rn].value
            if v == None or v == "/" or v =='　' or v== '' or v =='#VALUE!' :
                result += '不明'
            elif isinstance( v , int) or isinstance(v,float):
                if sheet[word+'1'].value == "长度":
                    result += (str(v)+"m")                      #输出XXm
                elif sheet[word+'1'].value == "实际使用年限":
                    result += (str(v)+"年")                     #输出XXXX年
            elif  isinstance( v ,datetime.datetime):
                result += v.strftime("%Y年%m月%d日" )            #日期格式
            else:
                result += v.strip()
        else:
            result += word

    return result

#   删除文本末尾的标点
def check_text(in_text:str,exception:str = '')->str:
    """删去文本末尾的标点和空格，可以指定不删去哪种，默认不删空格"""
    text = in_text
    if len(text)>0:
        while (text[-1] in ('；', '，' ,'、','：','.','。') and text[-1]!=exception):
            text = text[:-1]
    return text

#   实施签名
def sign(doc,path:str,all_names:set[str]|list[str],key_word:str,count:int)->None:
    """实施签名，即根据占位图片的名称替换。count表示占位图片的总数"""
    temp_count = 0
    for name in all_names:
        temp_count += 1
        image_path = path+'\\电子签名\\' + name +'.png'
        insert_picture(doc,image_path,key_word,40)
    for _ in range(count-temp_count):
        image_path = path+'\\电子签名\\' + '占位' +'.png'
        insert_picture(doc,image_path,key_word,40)


def replace_placeholder_with_image(doc, image_path, keyword):
    """执行替换"""
    # 遍历文档中的所有内联形状
    for shape in doc.InlineShapes:
        if shape.Type == 3:
            #检查是否是占位图片
            if shape.Title== keyword:
                # 替换为实际图片
                range_to_replace = shape.Range
                # 删除原图
                shape.Delete()
                # 在原图的位置插入新图
                doc.InlineShapes.AddPicture(
                    FileName=image_path,
                    LinkToFile=0,  # False 的值为 0
                    SaveWithDocument=1,  # True 的值为 1
                    Range=range_to_replace
                )
 
def delete_page_by_text(doc, target_text:str)->None:
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
    # 删除目标页的内容
    target_range.Delete()
    

"""
    主要改动：
    1、按照管段重新组织报告内容
    2、开挖分项报告增加结论页
    
"""
import collections
import random
import openpyxl
from openpyxl.workbook import Workbook
import datetime
import win32com.client as win32
import os
import traceback
import math
from mypackage import r_generator as rg
from mypackage.LOG_DATA import LOG_DICT,RISKY_EVA_C,RISKY_EVA_S
from mypackage import interraction_terminal 
import docx_to_pdf,sign_only

"""=========================编辑生成全部用于替换的列表索引文件replacements======================"""

def make_text_for_c1(report_name:str,workbook:Workbook):
    """用于生成每个工程最开始的概况文本"""
    sheet= workbook['管段清单']
    log_dict:dict[str] = rg.get_col_in_sheet(sheet)
    rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    c1_list:list[str] = []
    key_cols:list[str] = [
                    log_dict['工程名称'],
                    '，起止点坐标',log_dict['起止点坐标'],
                    '，竣工图号',log_dict['竣工图号'],
                    '，竣工日期',log_dict['竣工日期'],
                    '，运行年限',log_dict['实际使用年限'],
                    '。'
                    ]
    i:int=0
    for row in rows:
        i += 1
        text:str = rg.get_text_by_log(sheet,row,key_cols)
        c1_list.append(f"{i}、{text}")

    return c1_list

def expand_all_tables(doc,pages:list[int])->None:
    """按照输入的分项报告数量，复制报告页张数。"""
    #   宏观检查报告    
    times:int = pages[0]
    if times>1:
        rg.copy_and_insert_report(doc , '复制宏观检查报告', times)
    rg.replace_text(doc, '复制宏观检查报告','',2)

    #   开挖检测
    times:int = pages[1]
    if times>1:
        rg.copy_and_insert_report(doc , '复制开挖报告', times,1)
    rg.replace_text(doc, '复制开挖报告','',2)  
       
    #   穿跨越检查
    times:int = pages[2]
    if times>1:
        rg.copy_and_insert_report(doc , '复制穿跨越报告', times)
    rg.replace_text(doc, '复制穿跨越报告','',2)  
    
    
    #   整理删除页面
    # rg.delete_page_by_text(doc, '待删除')


def expand_all_figs(workbook:Workbook, doc, report_name:str):
    """根据‘全部静态台账工作簿’中的内容，复制空白图张数"""
    sheet = workbook['管段清单']
    log_dict:dict = rg.get_col_in_sheet(sheet)
    rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号']) 
    pic_nums:set[str]=set(sheet[log_dict['管道编码']+row].value for row in rows)
    times:int = len(pic_nums)
    # times:int = 39
    if times>1:
        rg.copy_and_insert_report(doc , '+复制路由图', times)
    rg.replace_text(doc, '+复制路由图','',2)

def do_add_row_for_all(report_name:str,doc,workbook:Workbook,gd_dict:dict[str,dict[tuple[str],list[str]]]):
    """执行‘管道清单’，‘问题清单’两个表格的扩张和填入"""
    if config['是否写入管道清单']:
        sheet = workbook['管段清单']
        log_dict:dict = rg.get_col_in_sheet(sheet)
        rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号']) 
        rg.add_row_to_table(doc,'管道序号',len(rows)-1)
        print('填写管道清单')
        rg.write_in_table(doc,'管道序号',sheet,rows)

    print('填写问题清单')
    # 遍历宏观检查中的问题
    sheet = workbook['宏观检查记录']
    white_list:set[str] = {'无','完好','符合','正常','合格'}
    log_dict:dict = rg.get_col_in_sheet(sheet)
    # rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号']) 
    depth_list:list[tuple[str]]=[]
    all_problems:list[str] = list()
    for gd_name,rows in gd_dict['宏观'].items():
        for row in rows:
            for key_word in ['地面标志','管道防护带','地表环境','阀门','阀门井','钢塑转换接头','调压箱、调压柜']:
                problem:str|None = sheet[log_dict[key_word]+row].value
                if problem is not None and problem not in white_list:
                    v = [
                        gd_name[0],gd_name[1],
                        sheet[log_dict['地表参照及位置描述']+row].value.strip() if sheet[log_dict['地表参照及位置描述']+row].value else '/',
                        f"{sheet[log_dict['坐标X']+row].value} ,{sheet[log_dict['坐标Y']+row].value}",
                        key_word+'：'+sheet[log_dict[key_word]+row].value,
                        ]
                    all_problems.append(v)
            depth = sheet[log_dict['管道埋深']+row].value
            if depth is not None and depth !='':
                depth_list += [(
                            gd_name[0],gd_name[1],
                            sheet[log_dict['地表参照及位置描述']+row].value.strip() if sheet[log_dict['地表参照及位置描述']+row].value else '/',
                            f"{float(sheet[log_dict['坐标X']+row].value):.0f} ，{float(sheet[log_dict['坐标Y']+row].value):.0f}",
                            f"{float(depth):.1f}",
                            '埋深不足' if sheet[log_dict['埋深达标']+row].value == '埋深不足' else '符合'
                            )] 
 
          
    #   遍历开挖检测中的问题
    sheet = workbook['开挖检测记录']
    log_dict:dict = rg.get_col_in_sheet(sheet)
    for gd_name,rows in gd_dict['开挖'].items():
        for row in rows:
            depth = sheet[log_dict['管道埋深（m）']+row].value
            depth_list += [(
                            gd_name[0],gd_name[1],
                            sheet[log_dict['探坑位置']+row].value.strip() if sheet[log_dict['探坑位置']+row].value else '/',
                            f"{float(sheet[log_dict['探坑坐标 X']+row].value):.0f} ，{float(sheet[log_dict['探坑坐标 Y']+row].value):.0f}",
                            f"{float(depth):.1f}",
                            '埋深不足' if sheet[log_dict['埋深达标']+row].value == '埋深不足' else '符合'
                            )] 
            
    #   执行写入所有问题
    rg.add_row_to_table(doc,'问题序号',len(all_problems)-1)
    for table in doc.Tables:
        first_cell = table.Cell(1,1)
        fist_cell_text:str|None = first_cell.Range.Text.strip()
        if '问题序号' in fist_cell_text:
            i:int=1
            for problem in all_problems:
                i += 1
                table.Cell(i,1).Range.Text = i-1
                j:int = 2
                for problem_text in problem[1:]:
                    if problem_text == None :
                        problem_text='不明'
                    table.Cell(i,j).Range.Text = problem_text
                    j += 1
            break

    rg.add_row_to_table(doc,'埋深序号',len(depth_list)-1)          
    for table in doc.Tables:
        first_cell = table.Cell(1,1)

        fist_cell_text:str|None = first_cell.Range.Text.strip()
        if '埋深序号' in fist_cell_text:
            i = 1
            for s_depth in depth_list:
                i += 1
                table.Cell(i,1).Range.Text = i-1
                j:int = 2
                for depth_text in s_depth[1:]:
                    if depth_text == None :
                        depth_text='不明'
                    table.Cell(i,j).Range.Text = depth_text
                    j += 1
            break

#   检测数据整理
def sort_out_data(workbook:Workbook,report_name:str)->dict[str,dict[tuple[str],list[str]]]:
    """按照“管道组织关系”中的内容，重新组织“宏观检查记录”“开挖检测记录”检测数据，返回此报告的（管道编码|名称元组——行号列表）字典"""
    f_dict:dict[str,dict[tuple[str],list[str]]]={}
    f_dict['宏观']={}
    f_dict['开挖']={}
    sheet = workbook['管段清单']
    log_dict= rg.get_col_in_sheet(sheet)
    rows = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    #  管段（编码，名称）元组集合
    gd_set:set[tuple[str]]=set((sheet[log_dict['管道编码']+row].value,sheet[log_dict['定检管道名称']+row].value) for row in rows)        
    
    sheet = workbook['宏观检查记录']
    log_dict= rg.get_col_in_sheet(sheet)
    for gd_num in gd_set:
        f_dict['宏观'][gd_num] = [str(c.row) for c in sheet[log_dict['管道编码']] if c.value==gd_num[0]]

    sheet = workbook['开挖检测记录']
    log_dict= rg.get_col_in_sheet(sheet)
    for gd_num in gd_set:
        f_dict['开挖'][gd_num] = [str(c.row) for c in sheet[log_dict['管道编码']] if c.value==gd_num[0]]

    return f_dict
                
    
#   替换文本
def do_replace(doc , replacements1:list[tuple[str,str]],replacements2:list[tuple[str,str]]=[])->None:
    """替换所有文本，先替换全局，再替换单次"""
    for target_text, replacement_text in replacements2:
        rg.replace_text(doc, target_text, replacement_text,2 )
    for target_text, replacement_text in replacements1:
        rg.replace_text(doc, target_text, replacement_text )

def make_sign_log(workbook:Workbook,report_name:str,name_list:list[str])->dict:
    """签名的字典索引"""
    sign_dict:dict = {}
    all_names:set[str] = set(name_list)

    sign_dict['签字'] = []  
    temp_list:list[str] =list(all_names)
    lenth = len(temp_list)
    sign_dict['签字'] += temp_list[:4]  # 结论页检验人员
    if lenth<4:
        sign_dict['签字'] += ['空白']*(4-lenth) 
    
    sign_dict['编制人员签字']=[random.choice(name_list)]
    sheet=workbook['管道基本信息']
    log_dict=rg.get_col_in_sheet(sheet)
    rows=rg.get_rows_in_sheet(report_name,sheet)
    sign_dict['审核人员签字']=list(set(sheet[log_dict['审核人']+row].value for row in rows if sheet[log_dict['审核人']+row].value)) 

    return sign_dict

#   换图函数
def make_pic_log(workbook:Workbook,report_name:str,kw_dict:dict[str,list[str]])->dict:
    """编制开挖图片、路由图替换的索引"""
    pic_dict:dict = {}
    
    #   开挖照片
    sheet =workbook['开挖检测记录']
    log_dict = rg.get_col_in_sheet(sheet)
    pic_dict['开挖']=[sheet[log_dict['记录自编号']+row].value for rows in kw_dict.values() for row in rows] 
    #   路由图
    pic_dict['管道分图']=[]
    pic_dict['管道总图']=[]
    if config['是否写入管道路由图']:
        sheet = workbook['管段清单']
        log_dict = rg.get_col_in_sheet(sheet)
        rows = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
        pic_g_nums:set[str]=set(sheet[log_dict['街道名称']+row].value for row in rows)
        pic_nums:set[str]=set(sheet[log_dict['管道编码']+row].value for row in rows)
        pic_dict['管道总图']=[f"{config['数据源所在']}\\总图\\路由_{pic_g_num}.jpg" for pic_g_num in pic_g_nums if os.path.exists(f"{config['数据源所在']}\\总图\\路由_{pic_g_num}.jpg")]    
        pic_dict['管道分图']=[f"{config['数据源所在']}\\路由图\\管线_{pic_num}.jpg" for pic_num in pic_nums]
       
    return pic_dict

    
#   完成索引
def make_replacement_index(workbook:Workbook,report_name:str,gd_dict:dict[str,dict[tuple[str],list[str]]])->dict:
    """返回的字典有：文本——前期替换，文本a——后期替换，分项报告——对应替换列表，总页数——宏观、开挖、穿跨越报告各自的页数，
    检验人员——宏观、开挖总的检验人员名单，问题汇总——问题的列表"""
    EQ_SET:dict[str,tuple[str]]={
        '可燃气体检测仪':('CYTJ-G-122','CYTJ-G-124','CYTJ-G-125','CYTJ-G-126','CYTJ-G-120','4940','11131','4917','7634','11143','11153'),
        '钢卷尺':('CYTJ-Y-068','CYTJ-Y-067','R001','R002','R003','R004'),
        'APL声学PE管道探测仪':('APL1270','APL1271','APL1159'), 
        '焊接检验尺':('CYTJ-Y-044','CYTJ-Y-045') ,
    }
    replacements:dict[str,list]={}
    replacements['文本'] = []
    replacements['总页数']=[]
    replacements['检验人员']=[]
    replacements['问题汇总']=[]
   
    #   封面
    temp_list:list[tuple]= []
    sheet =workbook['管道基本信息']
    log_dict = rg.get_col_in_sheet(sheet)
    rows = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    row = rows[0]
    for key,value in LOG_DICT['PE定期检验报告']:
        text = sheet[log_dict[value]+row].value
        temp_list.append((key,text))
    replacements['文本'] += temp_list
    global_name = sheet[log_dict['管道名称']+row].value
    global_date = sheet[log_dict['检验日期']+row].value
    #   结论页，数据仍出自基本信息
    temp_list= []
    temp_list = [
                ('+使用单位',sheet[log_dict['使用单位']+row].value),
                ('+单位地址',sheet[log_dict['单位地址']+row].value),
                ('+安全管理人员',sheet[log_dict['安全管理人员']+row].value),
                ('+联系电话',sheet[log_dict['联系电话']+row].value),
                ('+编制日期',sheet[log_dict['检验日期']+row].value),
                # ('+分布区域',sheet[log_dict['分布区域']+row].value),    #   这个实际是在正文
                ]
    replacements['文本'] += temp_list
    
    #   正文部分
    # if config['是否写入管道清单']:
    sheet = workbook['管段清单']
    log_dict = rg.get_col_in_sheet(sheet)
    rows = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    replacements['文本'] += [('+工程总数',len(rows)),('+管道总数',len(gd_dict['宏观']))]*3

    #   统计测深点数，埋深不足点数，是否有深根植物、土壤扰动、管道占压，是否有穿跨越保护，遍历日期
    sheet =workbook['宏观检查记录']
    log_dict = rg.get_col_in_sheet(sheet)
    all_date:list=[]    #   所有日期，当前不含宏观日期
    check_dict:dict[str,int|set[str]]={}
    depth_count:int=0  #   测深数量统计          
    error_depth:int=0   #   埋深不足统计
    row_marco_list = [row for rows in gd_dict['宏观'].values() for row in rows]
    for key in ['管道防护带','地表环境','穿、跨越公路','穿、跨越河流','地面标志','管道埋深','埋深达标','阀门','阀门井','钢塑转换接头','调压箱、调压柜']:
        temp_set:set[str]=set()
        temp_list:list[int|float|str]=[]
        if key in ['管道埋深','埋深达标']:  #   对于这两个列表记数
            temp_list=[sheet[log_dict[key]+row].value for row in row_marco_list if sheet[log_dict[key]+row].value]
            check_dict[key]=len(temp_list)
        else:
            for row in row_marco_list:
                v=sheet[log_dict[key]+row].value 
                if v:
                    for p in v.split(', '):
                        temp_set.add(p)
            check_dict[key]=temp_set
    #   环境检查结果
    if '全线深根植物伴行' in check_dict['管道防护带']: 
        check_dict['管道防护带'].discard('全线深根植物伴行')  
        check_dict['管道防护带'].add('深根植物')
    if not check_dict['管道防护带']|check_dict['地表环境']:
        replacements['文本'].append(('+环境检查结果','以上检查内容未见异常'))
    else:
        temp_text:str=''
        for s_p in check_dict['管道防护带']|check_dict['地表环境']:
            temp_text+=f"{s_p}，"
        replacements['文本'].append(('+环境检查结果',f"管道存在{rg.check_text(temp_text)}等情况，其余检查内容未见异常。"))
    #   穿跨越检查结果
    if not check_dict['穿、跨越公路']|check_dict['穿、跨越河流']:
        replacements['文本'].append(('+穿、跨越检查结果','本次检验管道无穿、跨越段'))
    elif '保护设施完好' in check_dict['穿、跨越公路']|check_dict['穿、跨越河流'] :
        replacements['文本'].append(('+穿、跨越检查结果','穿、跨越保护设施完好'))
    else:
        replacements['文本'].append(('+穿、跨越检查结果','无穿、跨越管道保护设施'))
    #   地面设施检查结果
    temp_para:str=''
    temp_text:str =''
    check_set:set[str]=check_dict['地面标志']-{'完好'}
    if  {'缺失','部分缺失','无标志','全线无标志','丢失'}-check_set:
        temp_text+='缺少地面标识，'
    check_set=check_set-{'缺失','部分缺失','无标志','全线无标志','丢失'}
    for t in check_set:
        temp_text+=f"{t}，"
    if temp_text:
        temp_para+=f"地面标识存在{rg.check_text(temp_text)}等问题；"
    temp_text=''
    check_set=(check_dict['阀门井']|check_dict['阀门'])-{'其他','合格','完好'}
    for t in check_set:
        temp_text+=f"{t}，"
    if temp_text:
         temp_para+=f"阀门、阀井存在{rg.check_text(temp_text)}等问题；"
    temp_para+='管线整体缺少示踪装置。'
    replacements['文本'].append(('+地面设施检查结果',temp_para))

    depth_count+=check_dict['管道埋深'] #   测深数量统计          
    error_depth+=check_dict['埋深达标'] #   埋深不足统计
        
    next_ins_year:int = 4
    sheet =workbook['开挖检测记录']
    log_dict = rg.get_col_in_sheet(sheet)
    dig_count:int =0
    row_dig_list=[row  for rows in gd_dict['开挖'].values() for row in rows]
    dig_count+=len(row_dig_list)
    for row in row_dig_list:
        v = sheet[log_dict['检验日期']+row].value
        if v !='' and v is not None: 
            all_date.append(v)           
        u = sheet[log_dict['埋深达标']+row].value
        if u is not None:
                error_depth += 1
        w = sheet[log_dict['结论']+row].value
        if w and ('3级' in w or '4级' in w):
            next_ins_year = 3
    check_dict['开挖']=  set(item for row in row_dig_list if sheet[log_dict['管道本体缺陷']+row].value for item in sheet[log_dict['管道本体缺陷']+row].value.split(', ')) 
    check_set= check_dict['开挖']-{'其他','无'}
    if check_set:
        replacements['文本'].append(('+开挖检查问题',f"管道存在{'、'.join(check_set)}等缺陷，"))
    else:
        replacements['文本'].append(('+开挖检查问题',''))
    depth_count += len(row_dig_list)
    if not all_date:
        all_date:list[datetime.datetime] = [datetime.datetime.strptime( '2024年10月15日',"%Y年%m月%d日"),]
    first_date = min(all_date)
    last_date = max(all_date)
    next_time = first_date + datetime.timedelta(days=next_ins_year*365-31+210)  #下次检验日期
    if next_time>datetime.datetime(2028,12,1):
        next_time=datetime.datetime(2028,12,1)
    elif next_time<datetime.datetime(2028,9,1):
        next_time=datetime.datetime(2028,12,1)
    replacements['文本'] += [
                    ('+测深数量',depth_count),
                    ('+开挖总数',dig_count),
                    ('+最早日期',first_date),
                    ('+最晚日期',last_date),
                    # ('+检查日期',first_date),   #   这里填充的是资料审查报告里的检验日期
                    ('+下次检验时间',next_time.strftime("%Y年%m月")),
                    ('+下次检验时间',next_time.strftime("%Y年%m月")),
                     ]
    if error_depth >0:
        replacements['文本'].append(('埋深均符合要求',str(error_depth)+'处埋深不符合要求'))
   
    
    
    #   宏观检查报告 
    replacements['宏观检查报告']=[]
    replacements['文本a']=[]
    temp_list:list[tuple]= []
    sheet =workbook['宏观检查记录']
    log_dict:dict = rg.get_col_in_sheet(sheet)
    # rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])  
    temp_count:int = 0  #分项报告编号
    white_list:set[str] = {'无','符合','完好','合格','正常','保护设施完好'}
    for gd_name,rows in gd_dict['宏观'].items():
        temp_count += 1
        replacements['文本a']+=[('报告#',f"报告（{temp_count}）")]
        # search_num = sheet[log_dict['所属记录编号']+row].value
        # row_0 = rg.get_rows_in_sheet(search_num,sheet)[0]
        #   宏观表头
        temp_list = [
                    ('+管道名称',gd_name[1]),
                    # ('+管道名称',sheet[log_dict['管道名称']+row].value),
                    ('+管段','/'),
                    ('+管道编号',gd_name[0]),
                    ('+设备名称型号',f"可燃气体检测仪、钢卷尺、APL声学PE管道探测仪"),
                    ('+设备编号',f"{random.choice(EQ_SET['可燃气体检测仪'])}、{random.choice(EQ_SET['钢卷尺'])}、{random.choice(EQ_SET['APL声学PE管道探测仪'])}"),
                    # ('+环境',sheet[log_dict['环境条件']+row_0].value),
                    # ('+检验日期',global_date),
                    ]
        replacements['检验人员']+=[sheet[log_dict['创建人']+row].value for row in rows if sheet[log_dict['创建人']+row].value] 
        results = ''
        key_dict:dict[str,list[str]]={}
        check_all_set:set[str]=set()    
        for key,options in LOG_DICT['宏观检查报告'].items():    
            key_dict[key]=[]
            for row in rows:
                u:str|None =sheet[log_dict['检查项目类别']+row].value
                check_all_set|=set(u.split(', '))    
                v:str|None= sheet[log_dict[key]+row].value
                if v:
                    key_dict[key]+=v.split(', ')
            check_list = key_dict[key]
            check_set = set(check_list)
            s_problem:str = ''
            stat_dict = collections.Counter(check_list)
            
            #   填勾选表
            if key not in check_all_set:    #   无这个检查项
                if '无此项' in options: #   输出无此项
                    temp_text:str=''
                    for option in options:
                        if option =='无此项':
                            temp_text +=f"☑{option}、"
                        else:
                            temp_text +=f"□{option}、" 
                else:                 # 没有‘无此项’选项，输出白名单里的值
                    temp_text:str=''
                    for option in options:  
                        if option in white_list:
                            temp_text +=f"☑{option}、"
                        else:
                            temp_text +=f"□{option}、" 
                temp_list+=[(f"+{key}",f"{temp_text}□",'  ')]
            else:
                if not key_dict[key] :      #    有检查项但无返回
                    temp_text:str=''
                    for option in options:  
                        if option in white_list:
                            temp_text +=f"☑{option}、"
                        else:
                            temp_text +=f"□{option}、" 
                    temp_list+=[(f"+{key}",f"{temp_text}□",'  ')]
                else:                   #   有返回值
                    s_problem:str = ''
                    
                    if not check_set-white_list:  #   只有白名单中的值
                        temp_text:str=''
                        for option in options:  
                            if option in white_list:
                                temp_text +=f"☑{option}、"
                            else:
                                temp_text +=f"□{option}、" 
                        temp_list+=[(f"+{key}",f"{temp_text}□",'  ')]
                    else:
                        temp_text:str=''
                        for option in options:  
                            if option in check_set and option not in white_list:    #实际有问题的
                                temp_text +=f"☑{option}、"                        
                            else:
                                temp_text +=f"□{option}、" 
                        if not check_set-set(options):    #   初始无自定义项
                            temp_list+=[(f"+{key}",f"{temp_text}□"'  ')]
                        else:
                            rest_option:set[str] = check_set-set(options)
                            if '全线无标志' in rest_option:
                                temp_text=temp_text.replace('□无标志','☑无标志')
                            elif '全线深根植物伴行' in rest_option:
                                temp_text=temp_text.replace('□深根植物','☑深根植物')
                            rest_option -= {'全线无标志','全线深根植物伴行'}
                            if not rest_option: #   整理后无自定义项
                                temp_list+=[(f"+{key}",f"{temp_text}□"'  ')]
                            else:               #   有自定义项
                                temp_list+=[(f"+{key}",f"{temp_text}☑",f"{'，'.join(rest_option)}")]
            #   统计问题项
            if '全线无标志' in stat_dict:
                stat_dict.pop('无标志',None)
                stat_dict.pop('缺失',None)
            if '全线深根植物伴行' in stat_dict:
                stat_dict.pop('深根植物',None)
            for problem_key,value_int in stat_dict.items():
                if problem_key=='无此项' or problem_key in white_list or problem_key in {'无跨越、穿越段仅路面宏观检验','无跨越、穿越段仅宏观检验','暗渠上方跨越，仅地表宏观检查'} :
                    pass
                elif problem_key=='全线无标志':
                    s_problem+='多处无标志、'
                elif problem_key=='全线深根植物伴行':
                    s_problem+='多处深根植物、'
                else:
                    s_problem+=f"{value_int}处{problem_key}、"
            if s_problem != '':
                results+=f"{key}：{rg.check_text(s_problem)}；"
        temp_list+=[('+结论',f"结论：{results}示踪装置：无示踪装置。")] 
        replacements['宏观检查报告'].append(temp_list)                   
    replacements['总页数'].append(temp_count)
    
    #   开挖检验报告
    replacements['开挖检验报告']=[]
    replacements['开挖报告首页']=[]
    sheet = workbook['开挖检测记录']
    log_dict = rg.get_col_in_sheet(sheet)  #获取表头索引
    temp_count = 0
    for gd_name,rows in gd_dict['开挖'].items():
        hole_no:int = 0 #   探坑编号
        for row in rows:
            hole_no+=1
            #   开挖内容页
            temp_list = []
            temp_count += 1
            replacements['文本a'] +=[('报告#',f"报告（{temp_count}-1）"),('报告#',f"报告（{temp_count}-2）")]
            temp_list += rg.make_change_text_for_heading(sheet,row,'开挖检测记录',log_dict)
            temp_list += [
                # ('+检验日期',global_date),
                ('+管道编号',gd_name[0]),
                ('+管道名称',gd_name[1]),
                ('+探坑编号',f"{hole_no}#"),
                ]
            temp_list += rg.make_change_text_for_option(sheet,row,'开挖检测记录',log_dict)
            v1 = sheet[log_dict['备注']+row].value
            if v1 :
                if sheet[log_dict['警示带']+row].value == '无':
                    v1=f"{rg.check_text(v1)}，管道无警示带。"
            elif sheet[log_dict['警示带']+row].value == '无':
                v1=f"管道无警示带。"
            else:
                v1='/'
            v2 = rg.check_text(sheet[log_dict['结论']+row].value) if '4级' not in sheet[log_dict['结论']+row].value else '2级'
            temp_list += [('+备注',f"备注：{v1}")]
            replacements['开挖检验报告'].append(temp_list)
            replacements['检验人员']+=sheet[log_dict['检验人员']+row].value.split(',')

            #   开挖报告首页
            lwd=sheet[log_dict['探坑规格（m）']+row].value
            for syb in {'*','×'}:
                if syb in sheet[log_dict['探坑规格（m）']+row].value:
                    lwd=f"m{syb}".join(sheet[log_dict['探坑规格（m）']+row].value.split(syb))+'m'
               
            
            pip_types=','.join([f"{p_t}mm" if 'dn' in p_t else p_t for p_t in sheet[log_dict['管道规格']+row].value.split(', ')])
           
            temp_list = [
                ('+实际检验日期',sheet[log_dict['检验日期']+row].value),
                ('+管道名称',gd_name[1]),
                ('+管道规格',pip_types),
                ('+探坑编号',f"{hole_no}#"),
                ('+探坑位置',sheet[log_dict['探坑位置']+row].value),
                ('+探坑规格',lwd),
                ('+地表状况',sheet[log_dict['地形、地貌、地物描述']+row].value),
                ('+环境条件','/'),
                ('+检验结论',f"检验结论：根据GB/T 43922-2024《在役聚乙烯燃气管道检验与评价》安全状况等级评定为{v2}。"),
                # ('+检验日期',global_date)
                ]
            x = sheet[log_dict['缺陷描述']+row].value
            y = sheet[log_dict['备注']+row].value
            z = ('，').join([word for word in sheet[log_dict['管道本体缺陷']+row].value.split(', ') if (x and word not in {'无','其他'} and word not in x) ])
            text_l:list[str] = [t for t in [z,x,y] if t and t not in {'无','其他'} ]
            text='；'.join([
                f"探坑坐标（{sheet[log_dict['探坑坐标 X']+row].value}，{sheet[log_dict['探坑坐标 Y']+row].value}）" ,
                f"{rg.check_text('，'.join(text_l)) if text_l else '开挖检验未发现异常'}。"])
            temp_list+=[('+检验情况',f"检验情况：{text}"),]
            replacements['开挖报告首页'].append(temp_list)
    replacements['总页数'].append(temp_count)
    
    #   穿、跨越检查
    replacements['穿、跨越报告']=[]
    temp_list:list[tuple]= []
    sheet =workbook['宏观检查记录']
    log_dict:dict = rg.get_col_in_sheet(sheet)
    temp_count_pages:int = 0    #   总的份数
    temp_count:int = 0  #分项报告编号
    for gd_name,rows in gd_dict['宏观'].items():
        key_dict:dict[str,list[str]]={}
        rows_1:list[str]=[]
        rows_2:list[str]=[]
        result_set:set[str] = set()
        # search_num = sheet[log_dict['所属记录编号']+row].value
        # row_0 = rg.get_rows_in_sheet(search_num,sheet)[0]
        for row in rows:
            v:str|None = sheet[log_dict['穿跨越类型']+row].value
            u:str|None = sheet[log_dict['穿、跨越河流']+row].value
            if v == '跨越':
                rows_1.append(row)
            elif v == '穿越':
                rows_2.append(row)
            if u is not None:
                result_set.add(u)
        
        if len(rows_1)+len(rows_2)<1:  #   整个管段无穿跨越
            # temp_count+=1
            # replacements['文本a'] +=[('报告#',f"报告（{temp_count}）")]
            # temp_list=[
            #     ('+管道名称',gd_name),
            #     ('+管段','/'),
            #     ('+管道编号',gd_num[gd_name]),
            #     ('+检验日期',global_date),
            #     # ('+环境条件',sheet[log_dict['环境条件']+row_0].value),
            #     ('+检查结论',"检查结论：本次检验的管道无穿、跨越段"),
            #     ('&号1','/'),
            #     ('&长度1','/'),
            #     ('&发现问题及位置描述1','/'),
            #     ('&备注1','/'),
            #     ('$号1','/'),
            #     ('$长度1','/'),
            #     ('$发现问题及位置描述1','/'),
            #     ('$备注1','/')
            # ]
            # replacements['穿、跨越报告'].append(temp_list)
            pass
        else:
            cap1:int = math.ceil(len(rows_1)/5)
            cap2:int = math.ceil(len(rows_2)/9)
            pages:int=max(cap1,cap2)
            temp_count+=1
            temp_count_pages+=pages
            for page in range(pages):
                if pages>1: 
                    replacements['文本a'] +=[('报告#',f"报告（{temp_count}-{page+1}）")]
                else:
                    replacements['文本a'] +=[('报告#',f"报告（{temp_count}）")]
                temp_list=[
                    ('+管道名称',gd_name[1]),
                    ('+管段','/'),
                    ('+管道编号',gd_name[0]),
                    # ('+检验日期',global_date),
                    # ('+环境条件',sheet[log_dict['环境条件']+row_0].value),
                    ]
                #   跨越填表
                if not rows_1:
                    temp_list+=[('&号1','/'),('&长度1','/'),('&发现问题及位置描述1','/'),('&备注1','/'),]
                else:   #   不为空
                    i=0
                    for row in rows_1[:5]:
                        i+=1
                        if '穿、跨越河流' in sheet[log_dict['检查项目类别']+row].value:
                            other = '跨越河流'
                        else:
                            other = '跨越公路'
                        temp_list+=[
                            (f'&号{i}',f"{i+page*5}"),
                            (f'&长度{i}',sheet[log_dict['穿跨越长度']+row].value),
                            (f'&发现问题及位置描述{i}',f"{sheet[log_dict['地表参照及位置描述']+row].value if sheet[log_dict['地表参照及位置描述']+row].value else ''} ，（{sheet[log_dict['坐标X']+row].value},{sheet[log_dict['坐标Y']+row].value}）"),
                            (f'&备注{i}',other),
                        ]
                    if i<5:
                        i+=1
                        temp_list+=[(f'&号{i}','/'),(f'&长度{i}','/'),(f'&发现问题及位置描述{i}','/'),(f'&备注{i}','/'),]
                rows_1 = rows_1[5:]
                #   穿越填表
                if not rows_2:
                    temp_list+=[(f'$号1','/'),(f'$长度1','/'),(f'$发现问题及位置描述1','/'),(f'$备注1','/'),]
                else:   #   不为空
                    j=0
                    for row in rows_2[:9]:
                        j+=1
                        if '穿、跨越河流' in sheet[log_dict['检查项目类别']+row].value:
                            other = '穿越河流'
                        else:
                            other = '穿越公路'
                        temp_list+=[
                            (f'$号{j}',f"{j+page*9}"),
                            (f'$长度{j}',sheet[log_dict['穿跨越长度']+row].value),
                            (f'$发现问题及位置描述{j}',f"{sheet[log_dict['地表参照及位置描述']+row].value if sheet[log_dict['地表参照及位置描述']+row].value else ''}，（{sheet[log_dict['坐标X']+row].value},{sheet[log_dict['坐标Y']+row].value}）"),
                            (f'$备注{j}','/'),
                        ]
                    if j<9:
                        j+=1
                        temp_list+=[(f'$号{j}','/'),(f'$长度{j}','/'),(f'$发现问题及位置描述{j}','/'),(f'$备注{j}','/'),]
                rows_2 = rows_2[9:]
                #   分项报告结论
                if pages- page >1:
                    temp_list+=[('+检查结论','检查结论：续下页')]
                elif '保护设施完好' in result_set:
                    temp_list+=[('+检查结论','检查结论：保护设施完好。')]
                else:
                    temp_list+=[('+检查结论','检查结论：穿、跨越段检验未发现异常。')]
                replacements['穿、跨越报告'].append(temp_list)
    replacements['总页数'].append(temp_count_pages)

    # 风险评估
    sheet =workbook['风险评估']
    log_dict=rg.get_col_in_sheet(sheet)
    rows =rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    replacements['风险预评估']=[]
    row = rows[0]
    for key_str,list_dict in RISKY_EVA_S.items():
        risk_score:int=0 
        for son_key,son_tuple in list_dict.items():
            v= sheet[log_dict[son_key]+row].value # 表格实际内容
            for option,score in son_tuple:
                if isinstance(option,tuple): # 如果键是区间（元组）
                    if v>=option[0] and v<option[1]:
                        risk_score += score
                else:
                    if v == option:
                        risk_score += score
        replacements['风险预评估'].append(risk_score)
    for key_str,list_dict in RISKY_EVA_C.items():
        risk_score:int=0 
        v= sheet[log_dict[key_str]+row].value # 表格实际内容
        for any_tuple in list_dict:
            option,score =any_tuple
            if isinstance(option,tuple): # 如果键是区间（元组）
                if v>=option[0] and v<option[1]:
                    risk_score += score
            else:
                if v == option:
                    risk_score += score 
        replacements['风险预评估'].append(risk_score)
    s_sigma_value = sum(replacements['风险预评估'][:8])
    c_sigma_value = sum(replacements['风险预评估'][8:])
    r_value = s_sigma_value*c_sigma_value
    if r_value<3600:
        r_class='低风险'
    elif r_value>=3600 and r_value<7800:
        r_class='中风险'
    elif r_value>=7800 and r_value<12600:
        r_class='较高风险'
    else:
        r_class='高风险'

    replacements['文本']+=[
        ('+预评估失效可能性得分',s_sigma_value),
        ('+预评估失效后果得分',c_sigma_value),
        ('+预评估风险值',r_value),
        ('+预评估风险等级',r_class),
        ('+预评估风险等级',r_class),
        ]
    
    replacements['风险再评估']=[]
    row = rows[1]
    for key_str,list_dict in RISKY_EVA_S.items():
        risk_score:int=0 
        for son_key,son_tuple in list_dict.items():
            v= sheet[log_dict[son_key]+row].value # 表格实际内容
            for option,score in son_tuple:
                if isinstance(option,tuple): # 如果键是区间（元组）
                    if v>=option[0] and v<option[1]:
                        risk_score += score
                else:
                    if v == option:
                        risk_score += score
        replacements['风险再评估'].append(risk_score)
    for key_str,list_dict in RISKY_EVA_C.items():
        risk_score:int=0 
        v= sheet[log_dict[key_str]+row].value # 表格实际内容
        for any_tuple in list_dict:
            option,score = any_tuple
            if isinstance(option,tuple): # 如果键是区间（元组）
                if v>=option[0] and v<option[1]:
                    risk_score = score
            else:
                if v == option:
                    risk_score = score 
        replacements['风险再评估'].append(risk_score)
    s_sigma_value = sum(replacements['风险再评估'][:8])
    c_sigma_value = sum(replacements['风险再评估'][8:])
    r_value = s_sigma_value*c_sigma_value
    if r_value<3600:
        r_class='低风险'
    elif r_value>=3600 and r_value<7800:
        r_class='中风险'
    elif r_value>=7800 and r_value<12600:
        r_class='较高风险'
    else:
        r_class='高风险'
    replacements['文本']+=[
        ('+再评估失效可能性得分',s_sigma_value),
        ('+再评估失效后果得分',c_sigma_value),
        ('+再评估风险值',r_value),
        ('+再评估风险等级',r_class),
        ('+再评估风险等级',r_class),
        ('+再评估风险等级',r_class)
        ]
    return replacements

def make_all_replacement_index(workbook:Workbook,report_name:str,gd_dict:dict[str,list[str]]):
    """管道基本信息：报告编号、管道名称、管道长度等"""
    replacements:list = []
    sheet = workbook['管道基本信息']
    log_dict:dict =rg.get_col_in_sheet(sheet)
    rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
    lenth:int = sum([sheet[log_dict['总长度']+row].value for row in rows ])
    l:float = lenth/1000
    row = rows[0]
    replacements += [
                ('+报告编号',report_name),
                ('+使用单位',sheet[log_dict['使用单位']+row].value),
                # ('+使用单位','成都燃气集团股份有限公司管网分公司'),
                ('+检验日期',sheet[log_dict['检验日期']+row].value),
                ('+编制日期',sheet[log_dict['编制日期']+row].value),
                ('+审核日期',sheet[log_dict['审核日期']+row].value),
                ('+批准日期',sheet[log_dict['批准日期']+row].value),
                ('+项目名称',sheet[log_dict['管道名称']+row].value),
                ('+管道名称',sheet[log_dict['管道名称']+row].value),
                ('+管道长度',l),
                # ('+管道长度','164.12'),# 新繁
                # ('+管道名称','天然气管道'),
                # ('+检验日期','2025年06月15日') # 新繁
                # ('+使用单位','成都成燃新繁燃气有限公司'),
                 ]
    
    # 检查是否有不明管段
    if config['是否写入管道清单']:
        sheet = workbook['管段清单']
        log_dict:dict =rg.get_col_in_sheet(sheet)
        rows:list[str] = rg.get_rows_in_sheet(report_name,sheet,log_dict['报告编号'])
        used_years:list[float] = [] #   全部年限
        # for rows in gd_dict.values():
        # for row in rows:
        #     if sheet[log_dict['实际使用年限']+row].value:
        #         used_years += [sheet[log_dict['实际使用年限']+row].value]
        used_years = [sheet[log_dict['实际使用年限']+row].value for row in rows if sheet[log_dict['实际使用年限']+row].value]
        replacements+=[('+投运年限',f"{min(used_years)}—{max(used_years)}年")]
        if any('使用单位指定管段' in sheet[log_dict['工程名称']+row].value for row in rows):
            replacements+=[('+不明管道检查','部分管段无资料，仅有GIS系统位置信息；其余管段仅见竣工图')]
            replacements+=[('+不明管道措施',',针对仅有GIS位置信息的管道，应开展专项调查工作进一步明确管道各项属性')]
        else:
            replacements+=[('+不明管道检查','所有管段仅见竣工图')]        
            replacements+=[('+不明管道措施','')]

    return replacements

def do_replace_in_son_report(doc,any_dict):
    """执行分项报告表格写入"""
    i:int = 0
    j:int = 0
    k:int = 0
    l:int = 1
    m:int = 1
    n:int = 0
    for table in doc.Tables:
        title_name:str = table.Title
        if title_name == '宏观检查报告':
            rg.replace_text_in_table(doc,table,any_dict['宏观检查报告'][i],'宏观检查报告索引')  
            i+=1
        elif title_name=='开挖报告首页':
            rg.replace_text_in_table(doc,table,any_dict['开挖报告首页'][n],'开挖报告首页索引')
            n+=1
        elif title_name == '开挖检验报告':
            rg.replace_text_in_table(doc,table,any_dict['开挖检验报告'][j],'开挖检测报告索引')  
            j+=1
        elif title_name == '穿、跨越报告':
            rg.replace_text_in_table(doc,table,any_dict['穿、跨越报告'][k],'穿、跨越报告索引')  
            k+=1
        elif title_name == '风险预评估报告':
            son_table = table.Cell(3,1).Tables(1)
            for score in any_dict['风险预评估'][:8]:
                cell = son_table.Cell(2,l)
                cell.Range.Text = score
                l+=1
            son_table = table.Cell(3,1).Tables(2)
            for score in any_dict['风险预评估'][8:]:
                cell = son_table.Cell(2,l-8)
                cell.Range.Text = score
                l+=1
        elif title_name == '风险再评估报告':
            son_table = table.Cell(3,1).Tables(1)
            for score in any_dict['风险再评估'][:8]:
                cell = son_table.Cell(2,m)
                cell.Range.Text = score
                m+=1
            son_table = table.Cell(3,1).Tables(2)
            for score in any_dict['风险再评估'][8:]:
                cell = son_table.Cell(2,m-8)
                cell.Range.Text = score
                m+=1
        else:
            pass

def do_change_sign_tag(doc,sign_dict:dict[str,list[str]])->None:
    """依照输入字典，替换签名图的标题"""
    i=0
    for shape in doc.InlineShapes:
        tag:str = shape.Title 
        if tag == '签字':
            shape.Title =sign_dict['签字'][i]
            i+=1
        if tag == '编制人员签字':
            shape.Title =sign_dict['编制人员签字'][0]
        if tag == '审核人员签字':
            shape.Title =sign_dict['审核人员签字'][0]

def do_replace_all_pic(doc,pic_dict:dict,):
    """执行所有图片的替换"""
    i:int = 0
    j:int = 0
    k:int = 0
    
    for shape in doc.InlineShapes:
        tag:str = shape.Title 
        # if tag == '签字':
        #     if config['是否生成签字']:
        #         if pic_dict['签字'][i]=='空白':
        #             pass
        #         else:
        #             rg.replace_pictue(doc,f"{config['签名图片所在']}\\{pic_dict['签字'][i]}.png",shape)
        #     i+=1
        # elif tag == '编制人员签字':
        #     if config['是否生成签字']:
        #         rg.replace_pictue(doc,f"{config['签名图片所在']}\\{pic_dict['编制人员签字'][0]}.png",shape)
        # elif tag == '审核人员签字':
        #     if config['是否生成签字']:
        #         rg.replace_pictue(doc,f"{config['签名图片所在']}\\{pic_dict['审核人员签字'][0]}.jpg",shape)
        if tag == '开挖':
            for ex_name in ['.jpg','.png','.jpeg']:
                f_path:str = f"{config['数据源所在']}\\开挖照片\\{pic_dict['开挖'][j]}{ex_name}"
                if os.path.exists(f_path):
                    rg.replace_pictue(doc,f_path,shape,120)
                    break
            j+=1
        elif tag == '管道总图':
            if config['是否写入管道路由图'] and pic_dict['管道总图']:
                rg.replace_pictue(doc,pic_dict['管道总图'][0],shape,580)
        elif tag == '管道分图':
            if config['是否写入管道路由图']:
                rg.replace_pictue(doc,pic_dict['管道分图'][k],shape,580) 
            k+=1
        else:
            pass

def do_input_para(doc,workbook:Workbook,gd_dict:dict[str,dict[str,list[str]]])->None:
    # sheet = workbook['管段清单']
    # log_dict = rg.get_col_in_sheet(sheet)
    # rg.copy_and_insert_paragraph(doc,'复制写入概况',make_text_for_c1(report_name,workbook))
    
    sheet =workbook['宏观检查记录']
    log_dict = rg.get_col_in_sheet(sheet)
    row_maro_list = [row for rows in gd_dict['宏观'].values() for row in rows] 
    check_dict:dict[str,set[str]]={}
    for key in ['管道防护带','阀门井','地面标志','埋深达标','地表环境',]:
        check_dict[key]=set(item for row in row_maro_list if sheet[log_dict[key]+row].value for item in sheet[log_dict[key]+row].value.split(', '))
    temp_list_1:list[str]=['对示踪系统完整性缺失的问题，应采取相应措施确保管道位置数据的准确性。',]
    temp_list_2:list[str]=[]
    
    sheet =workbook['开挖检测记录']
    log_dict = rg.get_col_in_sheet(sheet)
    row_dig_list = [row for rows in gd_dict['开挖'].values() for row in rows] 
    check_dict['开挖缺陷']=set(item for row in row_dig_list if sheet[log_dict['管道本体缺陷']+row].value for item in sheet[log_dict['管道本体缺陷']+row].value.split(', '))
    check_dict['埋深达标']|=set(sheet[log_dict['埋深达标']+row].value for row in row_dig_list if sheet[log_dict['埋深达标']+row].value )
    
    if '现场未见' in check_dict['阀门井']:
        temp_list_1.append('对现场未见的阀井进行核实，核实后进行恢复或重建。')
        temp_list_2.append('对现场未见的阀井进行核实，核实后进行恢复或重建。')
    if '深根植物' in check_dict['管道防护带'] or '全线深根植物伴行' in check_dict['管道防护带']:
        temp_list_1.append('对有深根植物伴行的管段，应采取保护措施或加强巡查避免其对管道的损害。')
    if check_dict['地表环境']:
        temp_list_1.append(f"对存在{'、'.join(check_dict['地表环境']|check_dict['埋深达标'])}等情况的管段，应改造或采取其他保护措施控制风险。")
        temp_list_2.append(f"对存在{'、'.join(check_dict['地表环境']|check_dict['埋深达标'])}等情况的管段，应改造或采取其他保护措施控制风险。")
    
    temp_list_1.append(f'对于开挖检验中发现的{'、'.join(check_dict['开挖缺陷']-{'其他','无'})}等问题，未能现场完成整改的部分应纳入整改计划或采取其他的管控措施。') 
    temp_list_2.append(f'对于开挖检验中发现的{'、'.join(check_dict['开挖缺陷']-{'其他','无'})}等问题，未能现场完成整改的部分应纳入整改计划或采取其他的管控措施。') 
    rg.copy_and_insert_paragraph(doc,'复制建议',temp_list_1)
    rg.copy_and_insert_paragraph(doc,'复制再评估建议',temp_list_2)
    
def solo_main(report_name:str,workbook:Workbook,word):
    """处理单个报告全体动作"""
    replacements_dict:dict = {}
    replacements_list:list[tuple] = []
    doc_modle_path = f"{config['模板文件']}"
    try:
        doc = word.Documents.Open(doc_modle_path)
        doc.Saved = True
        print('编写宏观记录中的管段——记录索引')
        gd_dict = sort_out_data(workbook,report_name)
        
        if config['是否生成概述段落']:
            print('扩张并写入段落')
            do_input_para(doc,workbook,gd_dict)
        
        print('生成替换用文本及签字索引')
        replacements_dict |= make_replacement_index(workbook,report_name,gd_dict)
        replacements_list += make_all_replacement_index(workbook,report_name,gd_dict) 
        sign_dict=make_sign_log(workbook,report_name,replacements_dict['检验人员'])

        print('替换内容及签名标题')
        do_replace( doc , replacements_dict['文本'],replacements_list )
        do_change_sign_tag(doc,sign_dict)
        
        if config['是否生成签字']:
            sign_only.sign_by_pic_name(doc,config['签名图片所在'])
        
        print('附件表格处理')
        do_add_row_for_all(report_name,doc,workbook,gd_dict)

        print('扩张分项报告表格')
        expand_all_tables(doc,replacements_dict['总页数'])

        print('替换残余内容')
        do_replace( doc , replacements_dict['文本a'])
    
        print('填写分项报告表格')
        do_replace_in_son_report(doc,replacements_dict)
    
        if config['是否写入管道路由图']:
            print('扩张路由图')
            expand_all_figs(workbook, doc, report_name)
        
        print('编制图片替换索引') 
        dig_dict=make_pic_log(workbook,report_name,gd_dict['开挖'])

        print('替换所有图片')
        do_replace_all_pic(doc,dig_dict)


        # 移动到文档的末端
        selection = word.Selection
        selection.EndKey(6)  # 6 表示 wdStory，即整个文档

        # 更新文档中的所有域
        doc.Fields.Update()
        doc.Saved = False
        output_file = f"{config['输出文件所在']}\\{sign_dict['审核人员签字'][0]}\\{report_name}.docx"
        # output_file = f"{config['输出文件所在']}\\{report_name}修正.docx"
        doc.SaveAs2(output_file, FileFormat=16)  # 16 表示docx 17 表示 PDF
        # print('正在保存文件')
        # output_file = f"{config['输出文件所在']}\\{report_name}.pdf"
        # doc.SaveAs2(output_file, FileFormat=17)  
        
        print(f"文档已保存为：{output_file}")

    except Exception as ex:
        traceback.print_exc()
        if doc is not None:
            doc.SaveAs2(f"{config['输出文件所在']}\\error_{report_name}.docx",FileFormat =16)
            print(f"{report_name}发生错误！")
            doc.Saved =True
            raise ex
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)


if __name__ == '__main__':
    set_list:list[tuple[int,str,str|bool,str|bool]]=[
        (2,'模板文件','docx',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\PE管定检报告模版_Ver_1.30_电子签.docx'),
        (0,'数据源所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网'),
        (0,'签名图片所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\电子签名'),
        (0,'输出文件所在','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\管网PE第二批'),
        (3,'是否生成概述段落',False,True),
        (3,'是否写入封面',False,False),
        (3,'是否写入管道清单',False,True),
        (3,'是否写入管道路由图',False,True),
        (3,'是否生成签字',False,False),    
        (3,'是否转pdf',False,False),    
    ]
    config:dict[str,str|bool]=interraction_terminal.set_argumments(set_list)
    app_type = rg.check_office_installation()
    if app_type == None:
        print('未找到合适的应用以打开文档')

    workbook:Workbook = openpyxl.load_workbook(f"{config['数据源所在']}\\原始数据.xlsx" )

    print('读取模板文件')
    
    if app_type == "office":
        word = win32.Dispatch("Word.Application")
    elif app_type == "wps":
        word = win32.Dispatch("Kwps.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    # 全局关闭拼写/语法检查
    word.Options.CheckSpellingAsYouType = False   # 关闭实时拼写检查
    word.Options.CheckGrammarAsYouType = False    # 关闭实时语法检查
    word.Options.ContextualSpeller = False        # 关闭上下文拼写检查（Word 2010+）
    #   初始化完成
    sheet=workbook['管道基本信息']
    all_names:list[str]=[]
    log_dict =rg.get_col_in_sheet(sheet)
    rows = rg.get_rows_in_sheet('二',sheet,log_dict['批次'])
    all_names=set(sheet[log_dict['报告编号']+row].value for row in rows if sheet[log_dict['报告编号']+row].value)   # 遍历静态台账里所有编号    
    for report_name in sorted(list(set(all_names)),reverse=False)[:1]:
        # if os.path.exists(f"{config['输出文件所在']}\\{report_name}.docx"):
        try:
            solo_main(report_name,workbook,word)
        except Exception as e:
            print('有错误发生')
        finally:
            continue
    if config['是否转pdf']:
        docx_to_pdf.docx_transform(config['输出文件所在'],config['输出文件所在'])



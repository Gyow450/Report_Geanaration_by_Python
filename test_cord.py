import pandas as pd
from pandas import DataFrame,Series
from datetime import datetime
# from numpy import nan
dfs:dict[str,DataFrame]=pd.read_excel(r"E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\管网840\原始数据.xlsx",sheet_name=None)
# print(dfs['管道基本信息'])
df:DataFrame = dfs['管道基本信息']
all_names = df.loc[:,'检验日期']
for name in all_names:
    output_name:str=''
    if pd.isna(name):
         output_name = '/'
    elif isinstance(name,(str,int,float)):
        output_name = name
    elif isinstance(name,datetime):
        output_name = name.strftime(r'%Y年%m月%d日')
    print(output_name,type(name))
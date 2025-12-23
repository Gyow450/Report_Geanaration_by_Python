from datetime import datetime
import pandas as pd
from pandas import DataFrame,Series
from numpy import nan
from LOG_DATA_FOR_CLASS import LOG_DICT

class Record():
    """抽象类，type为Any，只有检验人员和检验日期属性"""
    instances=[]
    p_statement:dict[str,list[str]]=dict()
    
    def __init__(self,data:DataFrame,record_type:str = '未定义')->None:
        """按照数据源，记录类型初始化"""
        self.name:str = data.loc[:,'记录自编号'].dropna().values[0]
        self.data:DataFrame = data
        self.inspection_date:datetime = data.loc[:,'检验日期'].dropna().values[0]
        self.inspecter:str = data.loc[:,'检验人员'].dropna().values[0]
        self.type:str = record_type
        self.heading:dict[str,str|datetime|None] = {}
        self.replacement:list[tuple[str,str|datetime|int|float]] = []
        self.problem:dict[str,str] = {}
        self.pages:int = 1
        # 表头的赋值，替换索引
        for key,log in LOG_DICT[self.type]:
            any_series:Series = data.loc[:,log].dropna()
            if  any_series.empty:
                self.heading[log] = '/'   
            else:
                self.heading[log] = any_series.values[0]
            self.replacement += [(key,self.heading[log])]
       
        self.replacement += [('+检验日期',self.inspection_date)]
        Record.instances.append(self)

    def __str__(self):
        return self.name

    
    def get_type(record):
        """返回记录类型的str"""
        return(record.type)
    
    def get_date(record):
        """返回检验日期的datetime"""
        return(record.inspection_date)
    
    def get_inspecter_names(record):
        """返回检验人员的list[str]"""
        inspecters:list[str] = record.inspecter.split(',')
        return inspecters

    @classmethod
    def clear_instances(cls):
        cls.instances.clear()

class MacroI_record(Record):
    """宏观检查记录，实例化时将DateFrame数据写入data属性,然后读取表头信息，写入替换索引"""
    
    def __init__(self, data:DataFrame):
        super().__init__(data,'宏观检查记录')
        self.name:str = data.loc[:,'所属记录编号'].dropna().values[0]
        
   

    def analyze(self):
        """
        应该统计各种问题，完成自身包括结论的替换索引，并且留出对外部的接口，统计单个问题的数量、完成的测深点数等
        """
        df = self.data.dropna(subset='检查项目类别')
        check_list:list[tuple[str,str]] = LOG_DICT['宏观检查记录填表']
        for i in range(len(df)):
            for key,log in check_list:
                if log == '检查项目类别':
                    ...
        

class EI_record(Record):
    """开挖检验记录，实例化时即将DateFrame数据写入data属性"""
    
    def __init__(self, data:DataFrame):
        super().__init__(data,'开挖检测记录')
         
        
    
    def analyze(self):
        df = self.data.dropna(subset='检查项目类别')
        self.problem:dict[str,str] = dict() 
    
class Report():
    def __init__(self,name:str,m_records:list[MacroI_record],e_records:list[EI_record]):
        self.name=name
        

if __name__ == '__main__':
 
    df:DataFrame = pd.read_excel('PE管定期检验数据汇总表.xlsx','宏观检查记录')
    cord = 'PE宏观检查记录202505287223'
    tf:bool = (df['记录自编号'] == cord) | (df['所属记录编号'] == cord)
    df1 = df.loc[tf,:]
    record1:MacroI_record = MacroI_record(df1)
    cord = 'PE宏观检查记录202505287214'
    tf:bool = (df['记录自编号'] == cord) | (df['所属记录编号'] == cord)
    df2 = df.loc[tf,:]
    record1:MacroI_record = MacroI_record(df2)
    # MacroI_record.clear_instances()
    print(f"\n当前宏观检查记录数量为：{len(MacroI_record.instances)}")
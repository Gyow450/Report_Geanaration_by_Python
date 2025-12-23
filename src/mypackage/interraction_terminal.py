"""互动窗口，得到参数设置，包括文件夹，文件名，其他控制参数等
类型码：0-文件夹，1-保存文件，2-打开文件，3-布尔，4-字符串"""

import sys
import re
import tkinter as tk
from tkinter import filedialog

def set_argumments(anything:list[tuple[int,str,str|bool,str|bool]])->dict[str,str|bool]:
    """按照特征码，关键字，扩展名|正则表达式构成元组"""
    def set_dir_value(key_word):
        final_dict[key_word]=filedialog.askdirectory(title=f'选择{key_word}文件夹')
        any_var[key_word].set(final_dict[key_word])
    
    def set_savefile_value(key_word,ex_info:str):
        ex_names=ex_info.split(',')
        final_dict[key_word]=filedialog.asksaveasfilename(title=f'选择{key_word}文件',
                                                          defaultextension=f'.{ex_names[0]}',
                                                          filetypes=[(f'{ex_name}文件',f'.{ex_name}') for ex_name in ex_names]
                                                          )
        any_var[key_word].set(final_dict[key_word])

    def set_openfile_value(key_word,ex_info:str):
        ex_names=ex_info.split(',')
        final_dict[key_word]=filedialog.askopenfilename(title=f'选择{key_word}文件',
                                                          defaultextension=f'.{ex_names[0]}',
                                                          filetypes=[(f'{ex_name}文件',f'.{ex_name}') for ex_name in ex_names]
                                                          )
        any_var[key_word].set(final_dict[key_word])
    
    def validate_input(entry,pattern):
        # 使用传入的正则表达式验证输入内容
        new_value = entry.get()
        if  bool(re.fullmatch(pattern, new_value)):
            entry.config(bg='white')       # 白背景
        else:
            entry.config(bg='#ffdddd')       # 淡红背景
            entry.focus_set()                # 焦点回去
    
    def on_ok():
        for name in final_dict.keys():
            final_dict[name]=any_var[name].get()
        root.destroy()

    def on_cancel():
        root.quit()
        sys.exit(0)
    
    final_dict:dict[str,str|bool]={}
    any_var:dict[str,tk.StringVar|tk.BooleanVar]={}
    root=tk.Tk()
    root.title('选择运行参数')
    
    i=-1 #row
    j=0 #column
    for temp_tuple in anything:
        i+=1
        type_num =temp_tuple[0]
        key_word =temp_tuple[1]
        ex_info = temp_tuple[2]
        if len(temp_tuple)>3:
            variable_value = temp_tuple[3]
        elif isinstance(ex_info,bool):
            variable_value = False
        else:
            variable_value = ''
        if isinstance(ex_info,bool):
            final_dict[key_word]=variable_value
            any_var[key_word]=tk.BooleanVar(value=final_dict[key_word])
        else:
            final_dict[key_word]=variable_value
            any_var[key_word]=tk.StringVar(value=final_dict[key_word])
        if type_num==0:
            tk.Label(root,text=f'选择{key_word}所在文件夹').grid(row=i,column=j)
            tk.Entry(root,textvariable=any_var[key_word],width=80,state='readonly').grid(row=i,column=j+1)
            tk.Button(root,text='选择文件夹',command=lambda key_word=key_word:set_dir_value(key_word)).grid(row=i,column=j+2)      
        elif type_num==1:
            tk.Label(root,text=f'选择{key_word}{ex_info}文件').grid(row=i,column=j)
            tk.Entry(root,textvariable=any_var[key_word],width=80,state='readonly').grid(row=i,column=j+1)
            tk.Button(root,text='选择文件',command=lambda key_word=key_word,ex_info=ex_info:set_savefile_value(key_word,ex_info)).grid(row=i,column=j+2)
        elif type_num==2:
            tk.Label(root,text=f'选择{key_word}{ex_info}文件').grid(row=i,column=j)
            tk.Entry(root,textvariable=any_var[key_word],width=80,state='readonly',).grid(row=i,column=j+1)
            tk.Button(root,text='选择文件',command=lambda key_word=key_word,ex_info=ex_info:set_openfile_value(key_word,ex_info)).grid(row=i,column=j+2)  
        elif type_num==3:
            tk.Checkbutton(root, text=key_word, variable=any_var[key_word]).grid(row=i, column=j)
        else:
            tk.Label(root,text=f'输入{key_word}').grid(row=i,column=j)
            entry=tk.Entry(root,textvariable=any_var[key_word],width=90,state='normal')
            entry.grid(row=i,column=j+1)
            var = any_var[key_word]
            var.trace_add('write',lambda *_,ent=entry,p=ex_info:validate_input(ent,p))
            # entry.bind('<Return>',lambda e,ent=entry,p=ex_info:validate_input(ent,p))
    
    i+=1
    tk.Button(root,text='确定',command=on_ok).grid(row=i,column=j)
    tk.Button(root,text='取消',command=on_cancel).grid(row=i,column=j+1)
    root.mainloop()
    return final_dict

if __name__=="__main__":
    a_list=[
        (0,'数据源',''),
        (1,'输出','pdf,docx'),
        (3,'是否写入概述',False),
        (4,'数字参数',r'^(-?\d+(~(-?\d+))?(,|$))*$'),
    ]
    print(set_argumments(a_list))
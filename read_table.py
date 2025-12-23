import win32com.client as win32
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

def read_table()->None:
    def pick_file():
        any_dict['file_path'] = filedialog.askopenfilename(title='选择模板文件')
        file_path_var.set(any_dict['file_path'])
    
    def on_ok():
        any_dict['file_path'] =file_path_var.get()
        any_dict['t_title'] =t_title_var.get()
        root.destroy()
    
    root = tk.Tk()
    root.title("读取的表格模板")
    any_dict:dict[str:str]={}
    any_dict['file_path'] = os.getcwd()
    file_path_var = tk.StringVar(value=any_dict['file_path'])
    
    tk.Label(root,text='模板文件路径').pack()
    tk.Entry(root,textvariable=file_path_var,width=80, state="readonly").pack()
    tk.Button(root, text="选择模板文件", command=pick_file).pack()

    any_dict['t_title']='任意表格标题'
    t_title_var = tk.StringVar(value=any_dict['t_title'])
    tk.Label(root,text='表格标题').pack()
    tk.Entry(root,textvariable=t_title_var,width=80, state="normal").pack()
    
    tk.Button(root, text="确定", width=10, bg="green", fg="white", command=on_ok).pack()

    root.mainloop()
    
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    
    doc = word.Documents.Open(any_dict['file_path'])
    cell_dict:dict[str,tuple[int,int]] ={}
    for table in doc.Tables:
        name = table.Title
        if name == any_dict['t_title']:
            for cell in table.Range.Cells:
               
                # 去掉 Word 单元格自带的 2 个不可见字符：\r、\x07
                text = cell.Range.Text.rstrip('\r\x07')
                if '+' in text or '&' in text or '$' in text:
                    cell_dict[text]=(cell.RowIndex,cell.ColumnIndex)
                    # if text == '+检验日期':
                    #     cell.Range.Text ='替换掉'
                    #     cell.Range.InsertAfter("我是谁")
                    #     start = cell.Range.End - len("我是谁")-1
                    #     end   = cell.Range.End
                    #     doc.Range(start, end).Font.Underline = 1 # 1 = wdUnderlineSingle

    # output_file = f"{path}\\99999999999.docx"
    # doc.SaveAs2(output_file, FileFormat=16)  # 16 表示docx 17 表示 PDF
    # print(f"文档已保存为：{output_file}")
    print(cell_dict)
    doc.Close(SaveChanges=False)
if __name__ == "__main__":
    read_table()
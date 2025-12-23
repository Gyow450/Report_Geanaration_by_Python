import json
import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

# 配置文件路径
CFG_FILE = "config.json"

# 默认配置
DEFAULT = {
    "file": "",
    "dir": "",
    "sign_path":"",
    "out_path":"",
    "para": True,
    "cover": True,
    "list": False,
    "pip_fig": True,
    "sign":True
}

def load_cfg():
    """加载配置"""
    if os.path.exists(CFG_FILE):
        with open(CFG_FILE, "r") as f:
            return json.load(f)
    return DEFAULT

def save_cfg(cfg):
    """保存配置"""
    with open(CFG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

def show_config_dialog():
    """显示配置对话框，返回用户选择"""
    def pick_file():
        cfg["file"] = filedialog.askopenfilename(title="选择模板文件")
        file_var.set(cfg["file"] or "（未选择）")

    def pick_dir():
        cfg["dir"] = filedialog.askdirectory(title="选择数据源文件夹")
        dir_var.set(cfg["dir"] or "（未选择）")
    
    def pick_sign_path():
        cfg["sign_path"] = filedialog.askdirectory(title="选择签名图片文件夹")
        sign_path_var.set(cfg["sign_path"] or "（未选择）")

    def pick_out_path():
        cfg["out_path"] = filedialog.askdirectory(title="选择输出文件位置")
        out_path_var.set(cfg["out_path"] or "（未选择）")
    
    def on_ok():
        cfg["para"] = para_var.get()
        cfg["cover"] = cover_var.get()
        cfg["list"] = list_var.get()
        cfg["pip_fig"] = pip_fig_var.get()
        cfg["sign"] = sign_var.get()
        save_cfg(cfg)
        root.destroy()

    def on_cancel():
        root.quit()
        sys.exit(0)

    # 加载配置
    cfg = load_cfg()

    # 主窗口
    root = tk.Tk()
    root.title("生成报告参数设置")

    # 设置变量
    file_var = tk.StringVar(value=cfg["file"] or "（未选择）")
    dir_var = tk.StringVar(value=cfg["dir"] or "（未选择）")
    sign_path_var = tk.StringVar(value=cfg["sign_path"] or "（未选择）")
    out_path_var = tk.StringVar(value=cfg["out_path"] or "（未选择）")
    para_var = tk.BooleanVar(value=cfg["para"])
    cover_var = tk.BooleanVar(value=cfg["cover"])
    list_var = tk.BooleanVar(value=cfg["list"])
    pip_fig_var = tk.BooleanVar(value=cfg["pip_fig"])
    sign_var = tk.BooleanVar(value=cfg["sign"])
    

    # 文件选择
    tk.Label(root, text="模板文件：").grid(row=0, column=0, sticky="e", padx=5, pady=6)
    tk.Entry(root, textvariable=file_var, width=80, state="readonly").grid(row=0, column=1, padx=5)
    tk.Button(root, text="选择模板文件", command=pick_file).grid(row=0, column=2, padx=5)

    # 目录选择
    tk.Label(root, text="数据源目录：").grid(row=1, column=0, sticky="e", padx=5, pady=6)
    tk.Entry(root, textvariable=dir_var, width=80, state="readonly").grid(row=1, column=1, padx=5)
    tk.Button(root, text="选择数据源目录", command=pick_dir).grid(row=1, column=2, padx=5)
    
    # 签名目录选择
    tk.Label(root, text="签名所在目录：").grid(row=2, column=0, sticky="e", padx=5, pady=6)
    tk.Entry(root, textvariable=sign_path_var, width=80, state="readonly").grid(row=2, column=1, padx=5)
    tk.Button(root, text="选择签名所在目录", command=pick_sign_path).grid(row=2, column=2, padx=5)
    
    # 输出位置
    tk.Label(root, text="输出文件目录：").grid(row=3, column=0, sticky="e", padx=5, pady=6)
    tk.Entry(root, textvariable=out_path_var, width=80, state="readonly").grid(row=3, column=1, padx=5)
    tk.Button(root, text="选择输出文件目录", command=pick_out_path).grid(row=3, column=2, padx=5)
    
    # 分隔线
    ttk.Separator(root, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", pady=8)
    
    # 复选框
    tk.Checkbutton(root, text="写入概述", variable=para_var).grid(row=5, column=1, sticky="w", padx=10)
    tk.Checkbutton(root, text="填写封面", variable=cover_var).grid(row=6, column=1, sticky="w", padx=10)
    tk.Checkbutton(root, text="生成管道清单", variable=list_var).grid(row=7, column=1, sticky="w", padx=10)
    tk.Checkbutton(root, text="替换管道路由图", variable=pip_fig_var).grid(row=8, column=1, sticky="w", padx=10)
    tk.Checkbutton(root, text="写入签名", variable=sign_var).grid(row=9, column=1, sticky="w", padx=10)

    # 按钮区
    tk.Button(root, text="确定", width=10, bg="green", fg="white", command=on_ok).grid(row=10, column=0, padx=10, pady=10)
    tk.Button(root, text="取消", width=10, command=on_cancel).grid(row=10, column=2, padx=10, pady=10)

    root.mainloop()

    return cfg

# 主程序
if __name__ == "__main__":
    config = show_config_dialog()
    print("最终配置：", config)
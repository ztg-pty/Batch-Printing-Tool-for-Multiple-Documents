import os
import re
import win32print
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import json
from pathlib import Path
import ctypes
import sys
import time

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_sorted_files(directory, file_types):
    """获取目录中按特定规则排序的文件列表"""
    files = []
    pattern = re.compile(r'(注：)?(\d+)-(\d+)(.*?)\.(doc|docx|pdf|txt|xls|xlsx|ppt|pptx|log|png|jpg|jpeg|ico|img)', re.IGNORECASE)
    
    for filename in os.listdir(directory):
        full_path = os.path.join(directory, filename)
        
        if os.path.isfile(full_path):
            match = pattern.search(filename)
            ext = os.path.splitext(filename)[1].lower()
            
            if ext in file_types:
                if match:
                    try:
                        num1 = int(match.group(2))
                        num2 = int(match.group(3))
                        files.append((full_path, num1, num2, filename))
                    except ValueError:
                        files.append((full_path, 0, 0, filename))
                else:
                    files.append((full_path, float('inf'), float('inf'), filename))
    
    files.sort(key=lambda x: (x[1], x[2], x[3]))
    return [{"path": f[0], "name": f[3]} for f in files]

def print_file_with_default_app(file_path, printer_name):
    """使用系统默认关联软件打印（核心逻辑）"""
    try:
        # 1. 前置检查：文件是否存在
        if not os.path.exists(file_path):
            return False, f"文件不存在: {os.path.basename(file_path)}"
        
        # 2. 管理员权限下设置默认打印机
        win32print.SetDefaultPrinter(printer_name)
        
        # 3. 调用系统默认软件执行打印（核心）
        os.startfile(file_path, "print")
        
        # 4. 短暂等待，确保打印指令发送成功
        time.sleep(0.5)
        return True, f"已通过默认软件发送打印任务: {os.path.basename(file_path)}"
    
    except OSError as e:
        if e.winerror == 1155:
            ext = os.path.splitext(file_path)[1].lower()
            return False, f"打印失败 {os.path.basename(file_path)}: 系统未找到{ext}格式的默认打开程序，请先设置默认软件"
        else:
            return False, f"打印失败 {os.path.basename(file_path)}: {str(e)}"
    except Exception as e:
        return False, f"打印失败 {os.path.basename(file_path)}: {str(e)}"

def get_printers():
    """获取系统中可用的打印机列表"""
    printers = []
    try:
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            printers.append(printer[2])
    except Exception as e:
        print(f"获取打印机失败: {e}")
    return printers

def save_settings(printer_name):
    """保存用户设置"""
    settings = {"last_printer": printer_name}
    try:
        config_dir = Path.home() / ".document_printer"
        config_dir.mkdir(exist_ok=True)
        config_file = config_dir / "settings.json"
        
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False)
    except Exception as e:
        print(f"保存设置失败: {e}")

def load_settings():
    """加载用户设置"""
    try:
        config_dir = Path.home() / ".document_printer"
        config_file = config_dir / "settings.json"
        
        if config_file.exists():
            with open(config_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                return settings.get("last_printer", "")
    except Exception as e:
        print(f"加载设置失败: {e}")
    return ""

def print_files_gui():
    """图形界面主函数"""
    global log_text
    
    root = tk.Tk()
    root.title("多格式文档多选打印工具（默认软件版）")
    root.geometry("800x700")
    root.resizable(True, True)
    
    default_font = ('Microsoft YaHei UI', 10)
    root.option_add("*Font", default_font)
    
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 选择目录
    ttk.Label(main_frame, text="选择文档目录:").grid(row=0, column=0, sticky=tk.W, pady=5)
    directory_var = tk.StringVar()
    directory_entry = ttk.Entry(main_frame, textvariable=directory_var, width=50)
    directory_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
    
    def browse_directory():
        directory = filedialog.askdirectory()
        if directory:
            directory_var.set(directory)
            clear_log()
            update_log(f"已选择目录: {directory}")
            update_file_list()
    
    ttk.Button(main_frame, text="浏览...", command=browse_directory).grid(row=0, column=2, padx=5, pady=5)
    
    # 文件类型过滤
    file_types = {
        "Word文档 (.doc, .docx)": ['.doc', '.docx'],
        "Excel文档 (.xls, .xlsx)": ['.xls', '.xlsx'],
        "PowerPoint演示文稿 (.ppt, .pptx)": ['.ppt', '.pptx'],
        "PDF文档 (.pdf)": ['.pdf'],
        "文本/日志文件 (.txt, .log)": ['.txt', '.log'],
        "图像文件 (.png, .jpg, .ico, .img)": ['.png', '.jpg', '.jpeg', '.ico', '.img'],
        "所有支持的格式": ['.doc', '.docx', '.pdf', '.txt', '.xls', '.xlsx', '.ppt', '.pptx', '.log', '.png', '.jpg', '.jpeg', '.ico', '.img']
    }
    file_type_var = tk.StringVar(value="所有支持的格式")
    file_type_combo = ttk.Combobox(main_frame, textvariable=file_type_var, values=list(file_types.keys()), width=47)
    file_type_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
    file_type_combo.bind("<<ComboboxSelected>>", lambda event: update_file_list())
    
    # 选择打印机
    ttk.Label(main_frame, text="选择打印机:").grid(row=2, column=0, sticky=tk.W, pady=5)
    printer_var = tk.StringVar()
    printer_combobox = ttk.Combobox(main_frame, textvariable=printer_var, width=50)
    printer_combobox.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
    
    def load_printers():
        printers = get_printers()
        printer_combobox['values'] = printers
        
        last_printer = load_settings()
        if last_printer and last_printer in printers:
            printer_var.set(last_printer)
            update_log(f"已加载上次使用的打印机: {last_printer}")
        elif printers:
            printer_var.set(printers[0])
        
        update_log(f"已加载 {len(printers)} 台打印机")
    
    def on_printer_selected(event=None):
        save_settings(printer_var.get())
        update_log(f"已选择打印机: {printer_var.get()}")
    
    ttk.Button(main_frame, text="加载打印机", command=load_printers).grid(row=2, column=2, padx=5, pady=5)
    printer_combobox.bind("<<ComboboxSelected>>", on_printer_selected)
    
    # 多选文件列表
    ttk.Label(main_frame, text="文件列表（按住Ctrl/Shift可多选）:").grid(row=3, column=0, sticky=tk.W, pady=5)
    file_listbox = tk.Listbox(main_frame, width=80, height=15, selectmode=tk.EXTENDED)
    file_listbox.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
    scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=file_listbox.yview)
    scrollbar.grid(row=4, column=3, sticky=(tk.N, tk.S))
    file_listbox.config(yscrollcommand=scrollbar.set)
    
    file_paths = {}
    
    def update_file_list():
        file_listbox.delete(0, tk.END)
        file_paths.clear()
        directory = directory_var.get()
        
        if not directory or not os.path.isdir(directory):
            update_log("错误: 无效的目录路径")
            return
        
        selected_types = file_types[file_type_var.get()]
        update_log(f"开始扫描目录: {directory} (类型: {', '.join(selected_types)})")
        files = get_sorted_files(directory, selected_types)
        
        if not files:
            update_log("警告: 在指定目录中未找到符合条件的文档文件")
            messagebox.showinfo("提示", f"在指定目录中未找到符合条件的{'/'.join(selected_types)}文件")
            return
        
        update_log(f"找到 {len(files)} 个匹配的文档文件")
        
        for i, file in enumerate(files):
            file_listbox.insert(tk.END, file["name"])
            file_paths[i] = file["path"]
    
    # 全选/取消全选
    def select_all():
        file_listbox.select_set(0, tk.END)
    
    def deselect_all():
        file_listbox.selection_clear(0, tk.END)
    
    ttk.Button(main_frame, text="全选", command=select_all).grid(row=5, column=0, padx=5, pady=5)
    ttk.Button(main_frame, text="取消全选", command=deselect_all).grid(row=5, column=1, padx=5, pady=5)
    
    # 打印进度
    ttk.Label(main_frame, text="打印进度:").grid(row=6, column=0, sticky=tk.W, pady=5)
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(main_frame, variable=progress_var, length=500)
    progress_bar.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=5)
    
    # 状态信息
    status_var = tk.StringVar()
    status_var.set("就绪")
    ttk.Label(main_frame, textvariable=status_var).grid(row=7, column=0, columnspan=3, sticky=tk.W, pady=5)
    
    # 打印选中文件（核心：使用默认软件打印）
    def start_printing():
        directory = directory_var.get()
        printer = printer_var.get()
        
        if not directory or not os.path.isdir(directory):
            messagebox.showerror("错误", "请选择有效的文档目录")
            return
        
        if not printer:
            messagebox.showerror("错误", "请选择打印机")
            return
        
        selected_indices = file_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请至少选择一个文件")
            return
        
        selected_files = [file_paths[i] for i in selected_indices]
        update_log(f"准备打印 {len(selected_files)} 个文件（使用系统默认软件）")
        
        threading.Thread(target=print_files_thread, args=(selected_files, printer), daemon=True).start()
    
    ttk.Button(main_frame, text="打印选中文件", command=start_printing, style='Accent.TButton').grid(row=8, column=0, columnspan=3, pady=10)
    
    def print_files_thread(files, printer):
        total_files = len(files)
        success_count = 0
        for i, file in enumerate(files):
            # 更新进度和状态
            root.after(0, lambda i=i: progress_var.set((i / total_files) * 100))
            root.after(0, lambda f=file: status_var.set(f"正在打印: {os.path.basename(f)}"))
            
            # 核心：调用默认软件打印函数
            success, message = print_file_with_default_app(file, printer)
            root.after(0, lambda m=message: update_log(m))
            
            if success:
                success_count += 1
            time.sleep(1)  # 给系统默认软件缓冲时间
        
        # 打印完成后更新状态
        root.after(0, lambda: progress_var.set(100))
        root.after(0, lambda: status_var.set(f"打印完成！成功: {success_count}/{total_files}"))
        root.after(0, lambda: messagebox.showinfo("完成", f"已完成 {total_files} 个文件的打印（成功 {success_count} 个）"))
    
    # 日志功能
    ttk.Label(main_frame, text="操作日志:").grid(row=9, column=0, sticky=tk.W, pady=5)
    log_text = tk.Text(main_frame, width=90, height=12, wrap=tk.WORD)
    log_text.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
    log_scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=log_text.yview)
    log_scrollbar.grid(row=10, column=3, sticky=(tk.N, tk.S))
    log_text.config(yscrollcommand=log_scrollbar.set)
    
    def update_log(message):
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)
    
    def clear_log():
        log_text.delete(1.0, tk.END)
    
    ttk.Button(main_frame, text="清空日志", command=clear_log).grid(row=11, column=0, pady=5)
    
    # 版权信息
    copyright_frame = ttk.Frame(root)
    copyright_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
    ttk.Label(copyright_frame, text="© 2026 文档打印工具（默认软件版）").pack(side=tk.LEFT)
    
    # 设置列和行的权重
    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(4, weight=1)
    main_frame.rowconfigure(10, weight=1)
    
    # 添加样式
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Microsoft YaHei UI', 10, 'bold'))
    
    # 初始化
    load_printers()
    root.mainloop()

if __name__ == "__main__":
    # 检查管理员权限（确保能设置默认打印机）
    if not is_admin():
        # 重新以管理员身份运行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        print_files_gui()
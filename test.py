import os
import tkinter as tk
from tkinter import filedialog, messagebox
from zipfile import ZipFile
import py7zr
import gzip
import shutil
from pyunpack import Archive
import patoolib

# 解压ZIP文件
def extract_zip(file_path, dest_folder):
    try:
        with ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(dest_folder)
        messagebox.showinfo("成功", f"ZIP文件解压到 {dest_folder} 成功")
    except Exception as e:
        messagebox.showerror("错误", f"解压失败: {e}")

# 解压RAR文件
def extract_rar(file_path, dest_folder):
    try:
        patoolib.extract_archive(file_path, outdir=dest_folder)
        messagebox.showinfo("成功", f"RAR文件解压到 {dest_folder} 成功")
    except Exception as e:
        messagebox.showerror("错误", f"解压失败: {e}")

# 解压7Z文件
def extract_7z(file_path, dest_folder):
    try:
        with py7zr.SevenZipFile(file_path, mode='r') as z:
            z.extractall(path=dest_folder)
        messagebox.showinfo("成功", f"7Z文件解压到 {dest_folder} 成功")
    except Exception as e:
        messagebox.showerror("错误", f"解压失败: {e}")

# 解压GZ文件
def extract_gz(file_path, dest_folder):
    try:
        with gzip.open(file_path, 'rb') as f_in:
            with open(os.path.join(dest_folder, os.path.basename(file_path).replace('.gz', '')), 'wb') as f_out:
                shutil.copyfileobj(f_in, f_out)
        messagebox.showinfo("成功", f"GZ文件解压到 {dest_folder} 成功")
    except Exception as e:
        messagebox.showerror("错误", f"解压失败: {e}")

# 获取当前目录下的压缩包
def scan_compressed_files(directory):
    compressed_files = []
    for file in os.listdir(directory):
        if file.endswith(('.zip', '.rar', '.7z', '.gz')):
            compressed_files.append(file)
    return compressed_files

# 解压操作
def extract_file(file_path, directory):
    # 创建解压后的文件夹
    dest_folder = os.path.join(directory, os.path.splitext(os.path.basename(file_path))[0])
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.zip':
        extract_zip(file_path, dest_folder)
    elif file_extension == '.rar':
        extract_rar(file_path, dest_folder)
    elif file_extension == '.7z':
        extract_7z(file_path, dest_folder)
    elif file_extension == '.gz':
        extract_gz(file_path, dest_folder)
    else:
        messagebox.showerror("错误", "不支持的压缩格式")

# 创建图形界面
def create_gui():
    # 设置DPI感知
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    # 全局目录变量
    current_directory = os.getcwd()

    def update_file_list():
        # 刷新文件列表
        compressed_files = scan_compressed_files(current_directory)
        listbox.delete(0, tk.END)
        for file in compressed_files:
            listbox.insert(tk.END, file)
        # 更新当前目录标签
        current_dir_label.config(text=f"当前目录: {current_directory}")

    def on_select_extract():
        selected_file = listbox.get(tk.ACTIVE)
        if selected_file:
            file_path = os.path.join(current_directory, selected_file)
            extract_file(file_path, current_directory)

    def on_folder_select():
        nonlocal current_directory
        folder = filedialog.askdirectory()
        if folder:
            current_directory = folder
            update_file_list()

    def on_scan():
        update_file_list()

    # 创建窗口
    root = tk.Tk()
    root.title("解压软件")
    root.geometry("1000x1100")

    # 创建一个标签
    label = tk.Label(root, text="快速解压工具", font=("Helvetica", 16))
    label.pack(pady=10)

    # 创建一个框架来放置按钮
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # 刷新按钮
    refresh_button = tk.Button(button_frame, text="刷新", command=on_scan)
    refresh_button.pack(side=tk.LEFT, padx=5)

    # 选择文件夹按钮
    folder_button = tk.Button(button_frame, text="选择文件夹", command=on_folder_select)
    folder_button.pack(side=tk.LEFT, padx=5)

    # 当前目录标签
    current_dir_label = tk.Label(root, text=f"当前目录: {current_directory}")
    current_dir_label.pack(pady=5)

    # 获取当前目录下的所有压缩包
    listbox = tk.Listbox(root, width=50, height=15)
    listbox.pack(pady=20)

    # 点击解压按钮时触发
    extract_button = tk.Button(root, text="解压选中的文件", command=on_select_extract)
    extract_button.pack(pady=10)

    # 软件信息
    info_label = tk.Label(root, text="开发者: 黄伟鑫\n版本: v1.0")
    info_label.pack(pady=10)

    # 初始扫描
    update_file_list()

    root.mainloop()

# 启动GUI
if __name__ == "__main__":
    create_gui()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 最近访问文件夹查看器
功能：
- 读取Windows最近访问的文件夹
- 按访问时间排序显示
- 支持搜索过滤
- 单击复制路径，双击打开文件夹
"""

import tkinter as tk
from tkinter import ttk, messagebox
import winreg
import os
import subprocess
import pyperclip
from datetime import datetime
import threading
import re
import win32com.client
import glob
import pystray
from PIL import Image, ImageDraw
import keyboard
import json
import time


class RecentFoldersViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows 最近访问文件夹查看器")
        self.root.geometry("1000x600")
        self.root.minsize(600, 400)
        
        # 存储文件夹数据
        self.folders_data = []
        self.filtered_data = []
        # 记录已打开的文件夹和打开次数
        self.opened_folders = set()
        self.open_history = {}  # {path: {'count': 打开次数, 'last_opened': 最后打开时间}}
        # 文件夹注释
        self.folder_comments = {}  # {path: comment}
        
        # 配置文件路径
        self.config_dir = os.path.join(os.path.expanduser("~"), ".recent_folders_viewer")
        self.config_file = os.path.join(self.config_dir, "config.json")
        
        # 系统托盘相关
        self.tray_icon = None
        self.is_hidden = False
        
        # 创建配置目录
        self.create_config_dir()
        # 加载配置
        self.load_config()
        
        self.setup_ui()
        self.setup_window_icon()
        self.setup_tray()
        self.setup_global_hotkey()
        self.load_recent_folders()
        
        # 让搜索框获得默认焦点
        self.root.after(100, lambda: self.search_entry.focus_set())
        
        # 绑定程序关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 搜索框架
        search_frame = ttk.Frame(main_frame)
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        # 搜索标签和输入框
        ttk.Label(search_frame, text="搜索过滤:").grid(row=0, column=0, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 刷新按钮
        refresh_btn = ttk.Button(search_frame, text="刷新", command=self.refresh_folders)
        refresh_btn.grid(row=0, column=2, padx=(5, 0))
        
        # 创建水平分割面板
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 左侧文件夹列表框架
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=2)  # 左侧占2/3
        
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # 创建文件夹列表Treeview
        columns = ('path', 'comment')
        self.tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=15)
        
        # 定义列标题和宽度
        self.tree.heading('path', text='文件夹路径')
        self.tree.heading('comment', text='注释')
        
        self.tree.column('path', width=400, anchor='w')
        self.tree.column('comment', width=200, anchor='w')
        
        # 左侧滚动条
        left_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=left_scrollbar.set)
        
        # 左侧布局
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        left_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 右侧文件预览框架
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)  # 右侧占1/3
        
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        # 右侧标题
        self.preview_title = ttk.Label(right_frame, text="", font=('', 10, 'bold'))
        self.preview_title.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 右侧文件列表
        file_columns = ('name', 'type', 'size')
        self.file_tree = ttk.Treeview(right_frame, columns=file_columns, show='headings', height=15)
        
        # 定义文件列表列标题和宽度
        self.file_tree.heading('name', text='文件名')
        self.file_tree.heading('type', text='类型')
        self.file_tree.heading('size', text='大小')
        
        self.file_tree.column('name', width=200, anchor='w')
        self.file_tree.column('type', width=80, anchor='center')
        self.file_tree.column('size', width=80, anchor='e')
        
        # 右侧滚动条
        right_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=right_scrollbar.set)
        
        # 右侧布局
        self.file_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        right_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # 绑定事件
        self.tree.bind('<Button-1>', self.on_single_click)
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Return>', self.on_enter_key)  # 绑定回车键
        self.tree.bind('<KeyPress>', self.on_tree_key_press)  # 绑定其他按键
        self.tree.bind('<<TreeviewSelect>>', self.on_folder_select)  # 绑定选择事件
        self.tree.bind('<Button-3>', self.show_context_menu)  # 绑定右键菜单
        
        # 绑定文件列表双击事件
        self.file_tree.bind('<Double-1>', self.on_file_double_click)
        
        # 为搜索框绑定键盘导航
        self.search_entry.bind('<Down>', self.focus_to_tree)
        self.search_entry.bind('<Return>', self.focus_to_tree)
        
        # 绑定全局快捷键
        self.root.bind('<Control-f>', self.focus_to_search)
        self.root.bind('<Control-F>', self.focus_to_search)  # 大小写都支持
        self.root.bind('<Escape>', self.hide_to_tray)  # ESC键隐藏到托盘
        
        # 绑定左右方向键切换两栏焦点（绑定到具体控件而不是全局）
        self.tree.bind('<Right>', self.on_tree_right_key)
        self.file_tree.bind('<Left>', self.on_file_tree_left_key)
        
        # 存储当前焦点状态
        self.current_panel = 'left'  # 'left' 或 'right'
        
        # 绑定窗口事件
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)  # 关闭按钮隐藏到托盘
        self.root.bind('<Unmap>', self.on_window_minimize)  # 最小化事件
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="编辑注释", command=self.edit_comment)
        self.context_menu.add_command(label="删除注释", command=self.delete_comment)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="打开文件夹", command=self.open_selected_folder)
        self.context_menu.add_command(label="复制路径", command=self.copy_selected_path)
    
    def get_recent_folders_from_lnk_files(self):
        """从Windows Recent文件夹的.lnk文件读取最近访问的文件夹"""
        folders = []
        
        try:
            # 获取Recent文件夹路径
            appdata = os.environ.get('APPDATA')
            if not appdata:
                return folders
                
            recent_path = os.path.join(appdata, 'Microsoft', 'Windows', 'Recent')
            if not os.path.exists(recent_path):
                return folders
            
            # 创建Shell对象来解析快捷方式
            shell = win32com.client.Dispatch("WScript.Shell")
            
            # 获取所有.lnk文件
            lnk_files = glob.glob(os.path.join(recent_path, '*.lnk'))
            
            for lnk_file in lnk_files:
                try:
                    # 解析快捷方式
                    shortcut = shell.CreateShortCut(lnk_file)
                    target_path = shortcut.Targetpath
                    
                    # 检查目标是否是文件夹
                    if target_path and os.path.exists(target_path) and os.path.isdir(target_path):
                        # 获取文件的修改时间作为访问时间
                        file_stat = os.stat(lnk_file)
                        access_time = datetime.fromtimestamp(file_stat.st_mtime)
                        
                        folders.append({
                            'path': target_path,
                            'access_time': access_time,
                            'exists': True
                        })
                    elif target_path:
                        # 如果目标是文件，获取其父目录
                        parent_dir = os.path.dirname(target_path)
                        if parent_dir and os.path.exists(parent_dir) and os.path.isdir(parent_dir):
                            file_stat = os.stat(lnk_file)
                            access_time = datetime.fromtimestamp(file_stat.st_mtime)
                            
                            folders.append({
                                'path': parent_dir,
                                'access_time': access_time,
                                'exists': True
                            })
                            
                except Exception as e:
                    # 跳过无法解析的快捷方式
                    continue
            
        except Exception as e:
            print(f"读取Recent文件夹时出错: {e}")
        
        return folders
    
    def load_recent_folders(self):
        """加载最近访问的文件夹"""
        def load_in_thread():
            # 显示加载提示
            self.root.after(0, self.show_folders_loading)
            
            # 使用字典来存储文件夹信息，以路径为键进行去重
            folder_dict = {}
            
            # 只从Recent文件夹的.lnk文件获取（这是真正的最近文件夹）
            try:
                recent_folders = self.get_recent_folders_from_lnk_files()
                total_found = len(recent_folders)
                
                # 分批处理文件夹数据
                for i, folder in enumerate(recent_folders):
                    # 标准化路径（解决大小写和路径分隔符问题）
                    normalized_path = os.path.normpath(folder['path']).lower()
                    if normalized_path not in folder_dict:
                        folder_dict[normalized_path] = folder
                    else:
                        # 如果路径已存在，保留访问时间更新的那个
                        if folder['access_time'] > folder_dict[normalized_path]['access_time']:
                            folder_dict[normalized_path] = folder
                    
                    # 每处理50个文件夹就更新一次进度
                    if (i + 1) % 50 == 0 or i == total_found - 1:
                        progress = min(100, int((i + 1) / total_found * 100))
                        self.root.after(0, self.update_folders_loading_progress, progress, len(folder_dict))
                        # 给UI一点时间响应
                        import time
                        time.sleep(0.01)
                        
            except Exception as e:
                print(f"从Recent文件夹读取失败: {e}")
                self.root.after(0, lambda: self.show_folders_loading_error(f"读取失败: {str(e)}"))
                return
            
            # 转换为列表并按优先级排序（打开次数+访问时间）
            folder_info = list(folder_dict.values())
            folder_info = self.sort_folders_by_priority(folder_info)
            
            # 分批更新UI
            self.root.after(0, self.update_folder_list_batched, folder_info)
        
        # 在后台线程中加载
        threading.Thread(target=load_in_thread, daemon=True).start()
    
    def show_folders_loading(self):
        """显示文件夹列表加载中的提示"""
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 显示加载提示
        self.tree.insert('', 'end', values=("正在扫描最近访问的文件夹...",), tags=("loading",))
        
        # 配置加载样式
        self.tree.tag_configure("loading", foreground="#4A90E2", font=('', 9, 'italic'))
    
    def update_folders_loading_progress(self, progress, found_count):
        """更新文件夹加载进度"""
        # 更新第一个项目的文本显示进度
        children = self.tree.get_children()
        if children:
            first_item = children[0]
            self.tree.item(first_item, values=(f"正在扫描... {progress}% (已找到 {found_count} 个文件夹)",))
    
    def show_folders_loading_error(self, error_msg):
        """显示文件夹加载错误"""
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 显示错误信息
        self.tree.insert('', 'end', values=(f"加载失败: {error_msg}",), tags=("error",))
        self.tree.tag_configure("error", foreground="red")
        
        # 已移除状态栏相关功能
    
    def update_folder_list_batched(self, folders_data):
        """分批更新文件夹列表，避免UI卡顿"""
        self.folders_data = folders_data
        
        # 清空现有列表
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 如果没有数据，显示提示
        if not folders_data:
            self.tree.insert('', 'end', values=("未找到最近访问的文件夹",), tags=("empty",))
            self.tree.tag_configure("empty", foreground="#888888", font=('', 10, 'italic'))
            
            # 已移除状态栏相关功能
            return
        
        # 分批添加文件夹到列表
        batch_size = 20  # 每批20个文件夹
        self.add_folders_batch(folders_data, 0, batch_size)
    
    def add_folders_batch(self, folders_data, start_idx, batch_size):
        """分批添加文件夹到列表"""
        end_idx = min(start_idx + batch_size, len(folders_data))
        
        # 添加当前批次的文件夹
        for i in range(start_idx, end_idx):
            folder = folders_data[i]
            # 根据状态和是否已打开设置不同的标签
            if folder['path'] in self.opened_folders:
                tags = ("opened_exists",) if folder['exists'] else ("opened_not_exists",)
            else:
                tags = ("exists",) if folder['exists'] else ("not_exists",)
            
            self.tree.insert('', 'end', values=(folder['path'],), tags=tags)
        
        # 配置标签样式
        self.tree.tag_configure("exists", foreground="black")
        self.tree.tag_configure("not_exists", foreground="gray")
        self.tree.tag_configure("opened_exists", foreground="#4A90E2")  # 淡蓝色
        self.tree.tag_configure("opened_not_exists", foreground="#6BA3F0")  # 稍亮的淡蓝色
        
        # 更新进度
        progress = min(100, int(end_idx / len(folders_data) * 100))
        loaded_count = end_idx
        
        # 如果还有更多数据，继续处理下一批
        if end_idx < len(folders_data):
            # 调度下一批（给UI一些时间响应）
            self.root.after(20, lambda: self.add_folders_batch(folders_data, end_idx, batch_size))
        else:
            # 所有批次完成，应用过滤器并恢复状态
            self.filtered_data = folders_data.copy()
            
            # 已移除状态栏相关功能
    
    
    def apply_filter(self):
        """应用搜索过滤"""
        # 清空现有项目
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 获取搜索文本
        search_text = self.search_var.get().lower()
        
        # 过滤数据
        if search_text:
            self.filtered_data = [
                folder for folder in self.folders_data
                if search_text in folder['path'].lower()
            ]
        else:
            self.filtered_data = self.folders_data.copy()
        
        # 添加过滤后的项目
        for folder in self.filtered_data:
            # 根据状态和是否已打开设置不同的标签
            if folder['path'] in self.opened_folders:
                tags = ("opened_exists",) if folder['exists'] else ("opened_not_exists",)
            else:
                tags = ("exists",) if folder['exists'] else ("not_exists",)
            
            # 获取该文件夹的注释
            comment = self.folder_comments.get(folder['path'], "")
            
            self.tree.insert('', 'end', values=(
                folder['path'],
                comment
            ), tags=tags)
        
        # 配置标签样式
        self.tree.tag_configure("exists", foreground="black")
        self.tree.tag_configure("not_exists", foreground="gray")
        self.tree.tag_configure("opened_exists", foreground="#4A90E2")  # 淡蓝色
        self.tree.tag_configure("opened_not_exists", foreground="#6BA3F0")  # 稍亮的淡蓝色
        
        # 更新状态已移除
    
    def on_search_change(self, *args):
        """搜索文本变化时的回调"""
        self.apply_filter()
    
    def on_single_click(self, event):
        """单击事件：选中项目（不再复制路径）"""
        # 保留单击选中功能，但移除自动复制路径的行为
        # 现在用户需要通过右键菜单来复制路径
        pass
    
    def on_double_click(self, event):
        """双击事件：在文件管理器中打开文件夹"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if item:
            path = self.tree.item(item, 'values')[0]
            try:
                if os.path.exists(path):
                    # 在文件管理器中打开（移除check=True避免误报错误）
                    subprocess.run(['explorer', path])
                    
                    # 记录文件夹打开历史
                    self.record_folder_open(path)
                    
                    # 将该文件夹移到最前面并更新访问时间
                    self.move_folder_to_top(path)
                else:
                    messagebox.showwarning("警告", f"文件夹不存在: {path}")
            except Exception as e:
                messagebox.showerror("错误", f"打开文件夹失败: {str(e)}")
    
    def move_folder_to_top(self, path):
        """将指定文件夹移到列表最前面"""
        # 找到目标文件夹并更新其访问时间
        for folder in self.folders_data:
            if folder['path'] == path:
                folder['access_time'] = datetime.now()
                break
        
        # 重新排序：已打开的文件夹优先，然后按访问时间排序
        self.folders_data.sort(key=lambda x: (
            x['path'] not in self.opened_folders,  # 已打开的文件夹在前（False < True）
            -x['access_time'].timestamp()  # 时间倒序
        ))
        
        # 刷新显示
        self.apply_filter()
    
    def focus_to_tree(self, event):
        """从搜索框焦点转到列表"""
        if self.tree.get_children():
            # 如果列表有项目，选中第一个并获得焦点
            first_item = self.tree.get_children()[0]
            self.tree.selection_set(first_item)
            self.tree.focus_set()
            self.tree.focus(first_item)
            return 'break'  # 阻止默认行为
    
    def on_enter_key(self, event):
        """回车键事件：打开选中的文件夹"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if item:
            # 复用双击事件的逻辑
            self.on_double_click(event)
            return 'break'
    
    def on_tree_key_press(self, event):
        """处理列表中的按键事件"""
        # 如果是字母数字键，将焦点转回搜索框并插入字符
        if event.char and event.char.isprintable() and not event.state & 0x4:  # 不是Ctrl组合键
            self.search_entry.focus_set()
            # 将当前字符添加到搜索框
            current_text = self.search_var.get()
            self.search_var.set(current_text + event.char)
            # 将光标移到末尾
            self.search_entry.icursor(tk.END)
            return 'break'
        elif event.keysym == 'BackSpace':
            # 退格键：回到搜索框并删除最后一个字符
            self.search_entry.focus_set()
            current_text = self.search_var.get()
            if current_text:
                self.search_var.set(current_text[:-1])
            self.search_entry.icursor(tk.END)
            return 'break'
    
    def focus_to_search(self, event):
        """Ctrl+F快捷键：聚焦到搜索框并全选文本"""
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)  # 全选搜索框中的文本
        return 'break'  # 阻止默认行为
    
    def get_icon_path(self, filename):
        """获取图标文件路径"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, filename)
    
    def load_icon_image(self, size=64):
        """加载图标图像"""
        try:
            if size <= 16:
                icon_path = self.get_icon_path('app_icon_16.png')
            elif size <= 32:
                icon_path = self.get_icon_path('app_icon_32.png')
            else:
                icon_path = self.get_icon_path('app_icon_64.png')
            
            if os.path.exists(icon_path):
                return Image.open(icon_path)
            else:
                # 如果文件不存在，创建备用图标
                return self.create_fallback_icon(size)
        except Exception as e:
            print(f"加载图标失败: {e}")
            return self.create_fallback_icon(size)
    
    def create_fallback_icon(self, size=64):
        """创建备用图标（当图标文件不存在时）"""
        image = Image.new('RGB', (size, size), color='white')
        draw = ImageDraw.Draw(image)
        
        # 按比例缩放文件夹形状
        scale = size / 64
        draw.rectangle([int(10*scale), int(20*scale), int(54*scale), int(50*scale)], 
                      fill='#FFD700', outline='#B8860B', width=max(1, int(2*scale)))
        draw.rectangle([int(10*scale), int(15*scale), int(25*scale), int(25*scale)], 
                      fill='#FFD700', outline='#B8860B', width=max(1, int(2*scale)))
        
        return image
    
    def setup_window_icon(self):
        """设置窗口图标"""
        try:
            # 使用PNG文件并同时设置iconbitmap和iconphoto
            png_path_32 = self.get_icon_path('app_icon_32.png')
            ico_path = self.get_icon_path('app_icon.ico')
            
            # 设置窗口图标（标题栏显示）
            if os.path.exists(png_path_32):
                photo = tk.PhotoImage(file=png_path_32)
                self.root.iconphoto(True, photo)
                # 保存引用以防止被垃圾回收
                self.window_icon = photo
            
            # 设置任务栏图标（使用iconbitmap）
            if os.path.exists(ico_path):
                try:
                    self.root.iconbitmap(ico_path)
                except Exception as e:
                    print(f"设置ICO图标失败: {e}")
                    # 如果ICO失败，尝试重新创建更好的ICO文件
                    self.create_better_ico()
            else:
                # 如果ICO文件不存在，创建一个
                self.create_better_ico()
            
            # 如果都不存在，创建备用图标
            if not os.path.exists(png_path_32) and not os.path.exists(ico_path):
                print("图标文件不存在，使用备用图标")
                fallback_icon = self.create_fallback_icon(32)
                
                # 保存备用图标为临时文件
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    fallback_icon.save(tmp.name, 'PNG')
                    photo = tk.PhotoImage(file=tmp.name)
                    self.root.iconphoto(True, photo)
                    self.window_icon = photo
                    
                    # 清理临时文件
                    import atexit
                    atexit.register(lambda: os.unlink(tmp.name) if os.path.exists(tmp.name) else None)
                
        except Exception as e:
            print(f"设置窗口图标失败: {e}")
    
    def create_better_ico(self):
        """创建更好的ICO文件来解决任务栏图标问题"""
        try:
            # 加载原始图标
            icon_64 = self.load_icon_image(64)
            
            # 创建多个尺寸的图标
            sizes = [16, 24, 32, 48, 64]
            images = []
            
            for size in sizes:
                resized = icon_64.resize((size, size), Image.Resampling.LANCZOS)
                images.append(resized)
            
            # 保存为ICO文件
            ico_path = self.get_icon_path('app_icon.ico')
            icon_64.save(ico_path, format='ICO', sizes=[(s, s) for s in sizes])
            
            # 立即尝试使用新创建的ICO文件
            self.root.iconbitmap(ico_path)
            print("重新创建ICO文件并设置成功")
            
        except Exception as e:
            print(f"创建更好的ICO文件失败: {e}")
    
    def setup_tray(self):
        """设置系统托盘"""
        try:
            # 创建托盘菜单
            menu = pystray.Menu(
                pystray.MenuItem("显示窗口", self.show_window, default=True),
                pystray.MenuItem("刷新列表", self.refresh_folders),
                pystray.MenuItem("退出", self.quit_app)
            )
            
            # 加载托盘图标（从文件加载）
            icon_image = self.load_icon_image(64)
            self.tray_icon = pystray.Icon(
                "recent_folders", 
                icon_image, 
                "最近文件夹查看器", 
                menu
            )
            
        except Exception as e:
            print(f"设置系统托盘失败: {e}")
    
    def hide_to_tray(self, event=None):
        """隐藏到系统托盘"""
        if not self.is_hidden:
            self.root.withdraw()  # 隐藏窗口
            self.is_hidden = True
            
            # 启动托盘图标（在后台线程中）
            if self.tray_icon and not self.tray_icon.visible:
                threading.Thread(target=self.tray_icon.run, daemon=True).start()
        
        return 'break'  # 阻止默认行为
    
    def on_window_minimize(self, event):
        """窗口最小化事件"""
        # 检查是否是真正的最小化（而不是其他unmap事件）
        if self.root.state() == 'iconic':
            self.hide_to_tray()
    
    def show_window(self, icon=None, item=None):
        """从托盘显示窗口或将已显示的窗口置顶"""
        if self.is_hidden:
            # 如果窗口被隐藏，显示它
            self.root.deiconify()  # 显示窗口
            self.is_hidden = False
        
        # 无论窗口是否已显示，都将其置顶并获得焦点
        self.root.lift()  # 置顶
        self.root.focus_force()  # 强制获得焦点
        self.root.attributes('-topmost', True)  # 临时置为最顶层
        self.root.after(100, lambda: self.root.attributes('-topmost', False))  # 100ms后取消最顶层
        
        # 让搜索框获得焦点
        self.search_entry.focus_set()
    
    def setup_global_hotkey(self):
        """设置全局快捷键"""
        try:
            # 注册全局快捷键 Ctrl+9
            keyboard.add_hotkey('ctrl+9', self.on_global_hotkey)
        except Exception as e:
            print(f"设置全局快捷键失败: {e}")
    
    def on_global_hotkey(self):
        """全局快捷键回调：显示窗口"""
        try:
            # 使用after方法确保在主线程中执行UI操作
            self.root.after(0, self.show_window)
        except Exception as e:
            print(f"全局快捷键处理失败: {e}")
    
    def on_tray_double_click(self, icon=None, item=None):
        """托盘图标双击事件：显示窗口"""
        try:
            # 使用after方法确保在主线程中执行UI操作
            self.root.after(0, self.show_window)
        except Exception as e:
            print(f"托盘双击处理失败: {e}")
    
    def quit_app(self, icon=None, item=None):
        """退出应用程序"""
        try:
            # 清理全局快捷键
            keyboard.unhook_all_hotkeys()
        except:
            pass
        
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.quit()
        self.root.destroy()
    
    def on_folder_select(self, event):
        """文件夹选择事件：加载文件夹内容到右侧预览"""
        selected_items = self.tree.selection()
        if not selected_items:
            # 没有选中项，清空文件预览
            self.clear_file_preview()
            return
        
        # 获取选中的文件夹路径
        item = selected_items[0]
        folder_path = self.tree.item(item, 'values')[0]
        
        # 更新预览标题
        folder_name = os.path.basename(folder_path) or folder_path
        self.preview_title.config(text=f"{folder_name}")
        
        # 在后台线程中加载文件列表
        threading.Thread(target=self.load_folder_contents, args=(folder_path,), daemon=True).start()
    
    def clear_file_preview(self):
        """清空文件预览"""
        # 清空文件列表
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 重置标题
        self.preview_title.config(text="")
    
    def load_folder_contents(self, folder_path):
        """在后台线程中加载文件夹内容"""
        try:
            if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
                self.root.after(0, self.clear_file_preview)
                return
            
            # 立即显示加载提示
            self.root.after(0, self.show_loading_preview, folder_path)
            
            max_items = 300  # 减少到300个以提升性能
            batch_size = 50   # 分批处理，每批50个
            
            # 获取文件夹中的项目
            try:
                # 使用scandir代替listdir，性能更好
                with os.scandir(folder_path) as entries:
                    folders = []
                    files = []
                    total_count = 0
                    
                    # 快速分类并统计总数
                    for entry in entries:
                        total_count += 1
                        try:
                            if entry.is_dir(follow_symlinks=False):
                                if len(folders) < max_items:
                                    folders.append(entry.name)
                            else:
                                if len(files) < max_items:
                                    files.append(entry.name)
                            
                            # 如果已经收集够了，就不继续遍历了
                            if len(folders) + len(files) >= max_items and total_count > max_items:
                                # 快速计算剩余数量
                                remaining_entries = list(entries)
                                total_count += len(remaining_entries)
                                break
                                
                        except (OSError, PermissionError):
                            continue
                
                # 排序（只排序需要显示的部分）
                folders.sort(key=str.lower)
                files.sort(key=str.lower)
                
                # 合并并限制数量
                selected_items = folders[:max_items]
                remaining_slots = max_items - len(selected_items)
                if remaining_slots > 0:
                    selected_items.extend(files[:remaining_slots])
                
                is_truncated = total_count > len(selected_items)
                
                # 分批处理文件信息获取
                files_data = []
                self.load_files_in_batches(folder_path, selected_items, batch_size, total_count, is_truncated)
                
            except PermissionError:
                self.root.after(0, lambda: self.show_preview_error("权限不足，无法访问此文件夹"))
            except Exception as e:
                self.root.after(0, lambda: self.show_preview_error(f"加载失败: {str(e)}"))
                
        except Exception as e:
            self.root.after(0, lambda: self.show_preview_error(f"发生错误: {str(e)}"))
    
    def load_files_in_batches(self, folder_path, items, batch_size, total_count, is_truncated):
        """分批加载文件信息，避免UI卡顿"""
        files_data = []
        
        def process_batch(start_idx):
            batch_data = []
            end_idx = min(start_idx + batch_size, len(items))
            
            for i in range(start_idx, end_idx):
                item_name = items[i]
                item_path = os.path.join(folder_path, item_name)
                
                try:
                    # 使用lstat避免跟随符号链接，性能更好
                    stat_info = os.lstat(item_path)
                    
                    if os.path.isdir(item_path):
                        # 文件夹
                        item_type = "文件夹"
                        size_str = "-"
                    else:
                        # 文件
                        _, ext = os.path.splitext(item_name)
                        item_type = ext.upper()[1:] if ext else "文件"
                        
                        # 快速格式化文件大小
                        size = stat_info.st_size
                        if size < 1024:
                            size_str = f"{size} B"
                        elif size < 1048576:  # 1024 * 1024
                            size_str = f"{size >> 10:.0f} KB"  # 使用位运算
                        elif size < 1073741824:  # 1024 * 1024 * 1024
                            size_str = f"{size >> 20:.1f} MB"
                        else:
                            size_str = f"{size >> 30:.1f} GB"
                    
                    batch_data.append({
                        'name': item_name,
                        'type': item_type,
                        'size': size_str,
                        'is_dir': os.path.isdir(item_path),
                        'path': item_path
                    })
                    
                except (OSError, PermissionError):
                    # 跳过无法访问的文件
                    continue
            
            return batch_data
        
        def process_next_batch(start_idx=0):
            if start_idx >= len(items):
                # 所有批次处理完成，排序并更新UI
                files_data.sort(key=lambda x: (not x['is_dir'], x['name'].lower()))
                self.root.after(0, self.update_file_preview, files_data, total_count, is_truncated)
                return
            
            # 处理当前批次
            batch_data = process_batch(start_idx)
            files_data.extend(batch_data)
            
            # 更新进度
            progress = min(100, int((start_idx + batch_size) / len(items) * 100))
            self.root.after(0, self.update_loading_progress, progress)
            
            # 调度下一批次（给UI一些时间响应）
            self.root.after(10, lambda: process_next_batch(start_idx + batch_size))
        
        # 开始处理第一批
        process_next_batch()
    
    def show_loading_preview(self, folder_path):
        """显示加载中的提示"""
        # 清空现有项目
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 显示加载提示
        folder_name = os.path.basename(folder_path) or folder_path
        self.preview_title.config(text=f"{folder_name}")
        
        self.file_tree.insert('', 'end', values=(
            "正在加载...",
            "",
            ""
        ), tags=("loading",))
        
        # 配置加载样式
        self.file_tree.tag_configure("loading", foreground="#4A90E2", font=('', 9, 'italic'))
        
        # 已删除状态栏相关功能
    
    def update_loading_progress(self, progress):
        """更新加载进度"""
        # 更新第一个项目的文本显示进度
        children = self.file_tree.get_children()
        if children:
            first_item = children[0]
            self.file_tree.item(first_item, values=(f"正在加载... {progress}%", "", ""))
    
    def update_file_preview(self, files_data, total_items=None, is_truncated=False):
        """在主线程中更新文件预览"""
        # 清空现有项目
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 如果文件夹为空，显示提示信息
        if not files_data:
            self.file_tree.insert('', 'end', values=(
                "文件夹为空",
                "",
                ""
            ), tags=("empty",))
            
            # 配置空文件夹样式
            self.file_tree.tag_configure("empty", foreground="#888888", font=('', 10, 'italic'))
            
            # 已移除状态栏相关功能
            return
        
        # 添加文件项目
        for file_info in files_data:
            # 根据文件类型设置不同的标签
            if file_info['is_dir']:
                tags = ("folder",)
                # 文件夹前面添加emoji
                display_name = f"📁 {file_info['name']}"
            else:
                tags = ("file",)
                display_name = file_info['name']
            
            self.file_tree.insert('', 'end', values=(
                display_name,
                file_info['type'],
                file_info['size']
            ), tags=tags)
        
        # 如果有截断，添加提示信息
        if is_truncated and total_items:
            remaining = total_items - len(files_data)
            self.file_tree.insert('', 'end', values=(
                f"... 还有 {remaining} 个项目未显示",
                "提示",
                ""
            ), tags=("info",))
        
        # 配置标签样式
        self.file_tree.tag_configure("folder", foreground="black")    # 文件夹用黑色
        self.file_tree.tag_configure("file", foreground="black")      # 文件用黑色
        self.file_tree.tag_configure("info", foreground="#888888", font=('', 9, 'italic'))  # 提示信息用灰色斜体
        
        # 已移除状态栏相关功能
    
    def show_preview_error(self, error_msg):
        """显示预览错误信息"""
        # 清空文件列表
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 添加错误信息项
        self.file_tree.insert('', 'end', values=(error_msg, "", ""), tags=("error",))
        self.file_tree.tag_configure("error", foreground="red")
    
    def on_file_double_click(self, event):
        """文件列表双击事件：打开文件或文件夹"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        values = self.file_tree.item(item, 'values')
        
        if len(values) < 3:
            return  # 错误信息项，不处理
        
        displayed_name = values[0]
        file_type = values[1]
        
        # 如果是文件夹（带emoji），需要去掉emoji前缀
        if displayed_name.startswith("📁 "):
            actual_name = displayed_name[2:]  # 去掉 "📁 " 前缀
        else:
            actual_name = displayed_name
        
        # 获取当前选中的文件夹路径
        selected_folder_items = self.tree.selection()
        if not selected_folder_items:
            return
        
        folder_path = self.tree.item(selected_folder_items[0], 'values')[0]
        file_path = os.path.join(folder_path, actual_name)
        
        try:
            if os.path.exists(file_path):
                # 使用系统默认程序打开文件/文件夹
                os.startfile(file_path)
                
                # 记录文件夹打开历史（因为打开了文件夹中的文件）
                self.record_folder_open(folder_path)
                
                # 将该文件夹移到最前面
                self.move_folder_to_top(folder_path)
                
                # 已移除状态栏显示功能
            else:
                messagebox.showwarning("警告", f"文件不存在: {actual_name}")
        except Exception as e:
            messagebox.showerror("错误", f"打开文件失败: {str(e)}")
    
    def create_config_dir(self):
        """创建配置目录"""
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
        except Exception as e:
            print(f"创建配置目录失败: {e}")
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                # 加载打开历史
                self.open_history = config.get('open_history', {})
                
                # 加载文件夹注释
                self.folder_comments = config.get('folder_comments', {})
                
                # 重建 opened_folders 集合
                self.opened_folders = set(self.open_history.keys())
                
                print(f"配置加载成功，包含 {len(self.open_history)} 条历史记录和 {len(self.folder_comments)} 条注释")
            else:
                print("配置文件不存在，使用默认设置")
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            self.open_history = {}
            self.opened_folders = set()
            self.folder_comments = {}
    
    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'open_history': self.open_history,
                'folder_comments': self.folder_comments,
                'last_saved': time.time()
            }
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            print(f"配置保存成功，包含 {len(self.open_history)} 条历史记录和 {len(self.folder_comments)} 条注释")
        except Exception as e:
            print(f"保存配置文件失败: {e}")
    
    def record_folder_open(self, folder_path):
        """记录文件夹打开历史"""
        current_time = time.time()
        
        if folder_path in self.open_history:
            # 增加打开次数
            self.open_history[folder_path]['count'] += 1
            self.open_history[folder_path]['last_opened'] = current_time
        else:
            # 首次打开
            self.open_history[folder_path] = {
                'count': 1,
                'first_opened': current_time,
                'last_opened': current_time
            }
        
        # 添加到已打开集合
        self.opened_folders.add(folder_path)
        
        # 保存配置
        self.save_config()
    
    def get_folder_priority_score(self, folder_data):
        """计算文件夹优先级分数，用于排序"""
        folder_path = folder_data['path']
        
        # 基础分数：最近访问时间（转换为分数，越近分数越高）
        base_score = folder_data['access_time'].timestamp()
        
        # 如果在打开历史中，根据打开次数和最后打开时间计算加分
        if folder_path in self.open_history:
            history = self.open_history[folder_path]
            
            # 打开次数加分（每次+1000分）
            count_bonus = history['count'] * 1000
            
            # 最后打开时间加分（如果最后打开时间比系统记录的访问时间更新，使用最后打开时间）
            last_opened_score = history['last_opened']
            if last_opened_score > base_score:
                base_score = last_opened_score
            
            # 频率加分：最近经常使用的文件夹额外加分
            days_since_first = (time.time() - history['first_opened']) / 86400  # 转换为天数
            if days_since_first > 0:
                frequency_bonus = (history['count'] / max(days_since_first, 1)) * 500  # 平均每天打开次数 * 500
            else:
                frequency_bonus = history['count'] * 500
            
            return base_score + count_bonus + frequency_bonus
        
        return base_score
    
    def sort_folders_by_priority(self, folders_data):
        """根据优先级排序文件夹列表"""
        # 为每个文件夹计算优先级分数并排序
        folders_data.sort(key=self.get_folder_priority_score, reverse=True)
        return folders_data
    
    def on_closing(self):
        """程序关闭时的处理"""
        # 保存配置
        self.save_config()
        
        # 清理全局快捷键
        try:
            keyboard.unhook_all_hotkeys()
        except:
            pass
        
        # 停止托盘图标
        if self.tray_icon:
            self.tray_icon.stop()
        
        # 关闭窗口
        self.root.destroy()
    
    def refresh_folders(self):
        """刷新文件夹列表"""
        self.load_recent_folders()
    
    def on_tree_right_key(self, event):
        """在左侧列表中按下右方向键时切换到右侧面板"""
        # 阻止默认的右方向键行为（可能会触发其他事件）
        
        # 设置当前面板为右侧
        self.current_panel = 'right'
        
        # 将焦点设置到右侧的文件列表
        self.file_tree.focus_set()
        
        # 如果右侧列表有项目但没有选中项，选中第一个
        if self.file_tree.get_children() and not self.file_tree.selection():
            first_item = self.file_tree.get_children()[0]
            self.file_tree.selection_set(first_item)
            self.file_tree.focus(first_item)
        
        return 'break'  # 阻止默认行为和事件传播
    
    def on_file_tree_left_key(self, event):
        """在右侧列表中按下左方向键时切换到左侧面板"""
        # 设置当前面板为左侧
        self.current_panel = 'left'
        
        # 清除右侧列表的选中状态
        self.file_tree.selection_remove(self.file_tree.selection())
        
        # 将焦点设置到左侧的文件夹列表
        self.tree.focus_set()
        
        # 如果左侧列表有项目但没有选中项，选中第一个
        if self.tree.get_children() and not self.tree.selection():
            first_item = self.tree.get_children()[0]
            self.tree.selection_set(first_item)
            self.tree.focus(first_item)
        
        return 'break'  # 阻止默认行为和事件传播
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        # 获取点击的项目
        item = self.tree.identify_row(event.y)
        if item:
            # 选中该项目
            self.tree.selection_set(item)
            self.tree.focus(item)
            
            # 获取路径
            path = self.tree.item(item, 'values')[0]
            
            # 根据是否有注释更新菜单项状态
            if path in self.folder_comments and self.folder_comments[path]:
                self.context_menu.entryconfig(1, state="normal")  # 删除注释菜单项
            else:
                self.context_menu.entryconfig(1, state="disabled")  # 禁用删除注释菜单项
            
            # 显示菜单
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
    
    def edit_comment(self):
        """编辑选中文件夹的注释"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        path = self.tree.item(item, 'values')[0]
        
        # 获取当前注释
        current_comment = self.folder_comments.get(path, "")
        
        # 创建编辑对话框
        self.show_comment_dialog(path, current_comment)
    
    def show_comment_dialog(self, path, current_comment):
        """显示注释编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("编辑文件夹注释")
        dialog.geometry("500x300")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 居中显示
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + 50,
            self.root.winfo_rooty() + 50
        ))
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 文件夹路径标签
        folder_name = os.path.basename(path) or path
        ttk.Label(main_frame, text=f"文件夹: {folder_name}", font=('', 10, 'bold')).pack(anchor="w")
        ttk.Label(main_frame, text=path, font=('', 8), foreground="gray").pack(anchor="w", pady=(0, 10))
        
        # 注释输入框
        ttk.Label(main_frame, text="注释:").pack(anchor="w")
        
        # 创建文本框和滚动条
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True, pady=(5, 10))
        
        comment_text = tk.Text(text_frame, wrap="word", height=8)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=comment_text.yview)
        comment_text.configure(yscrollcommand=scrollbar.set)
        
        comment_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 设置当前注释内容
        comment_text.insert("1.0", current_comment)
        comment_text.focus_set()
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        def save_comment():
            """保存注释"""
            new_comment = comment_text.get("1.0", "end-1c").strip()
            if new_comment:
                self.folder_comments[path] = new_comment
            else:
                # 如果注释为空，删除该注释
                if path in self.folder_comments:
                    del self.folder_comments[path]
            
            # 保存配置
            self.save_config()
            
            # 刷新显示
            self.update_folder_display()
            
            dialog.destroy()
        
        def cancel_edit():
            """取消编辑"""
            dialog.destroy()
        
        # 保存和取消按钮
        ttk.Button(button_frame, text="保存", command=save_comment).pack(side="right", padx=(5, 0))
        ttk.Button(button_frame, text="取消", command=cancel_edit).pack(side="right")
        
        # 绑定快捷键
        dialog.bind('<Control-Return>', lambda e: save_comment())
        dialog.bind('<Escape>', lambda e: cancel_edit())
        
        # 选中所有文本便于编辑
        comment_text.select_range("1.0", "end")
    
    def delete_comment(self):
        """删除选中文件夹的注释"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        path = self.tree.item(item, 'values')[0]
        
        if path in self.folder_comments:
            if messagebox.askyesno("确认删除", "确定要删除这个文件夹的注释吗？"):
                del self.folder_comments[path]
                self.save_config()
                self.update_folder_display()
    
    def open_selected_folder(self):
        """打开选中的文件夹"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # 复用双击事件的逻辑
        fake_event = type('Event', (), {})()
        self.on_double_click(fake_event)
    
    def copy_selected_path(self):
        """复制选中文件夹的路径"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        path = self.tree.item(item, 'values')[0]
        try:
            pyperclip.copy(path)
            # 可以添加一个简短的提示
            self.root.title("Windows 最近访问文件夹查看器 - 路径已复制")
            self.root.after(2000, lambda: self.root.title("Windows 最近访问文件夹查看器"))
        except Exception as e:
            messagebox.showerror("错误", f"复制到剪贴板失败: {str(e)}")
    
    def update_folder_display(self):
        """更新文件夹显示（用于在编辑注释后刷新显示）"""
        # 重新应用过滤器以更新显示
        self.apply_filter()


def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = RecentFoldersViewer(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("启动错误", f"程序启动失败: {str(e)}")


if __name__ == "__main__":
    main()

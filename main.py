import win32gui
import win32con
import win32com.client
import os
import time
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
from PIL import Image, ImageGrab
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WindowMonitor:
    def __init__(self, save_dir="pic", check_interval=2, log_callback=None):
        # 确保保存目录存在
        self.save_dir = save_dir
        if not os.path.exists(self.save_dir):
            os.makedirs(self.save_dir)
            logger.info(f"创建截图保存目录: {self.save_dir}")
            if log_callback:
                log_callback(f"创建截图保存目录: {self.save_dir}")
        
        self.check_interval = check_interval  # 检查间隔（秒）
        self.known_windows = set()  # 已记录的窗口句柄集合
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.running = False
        self.log_callback = log_callback
        
    def log(self, message):
        """记录日志并可选地调用回调函数"""
        logger.info(message)
        if self.log_callback:
            self.log_callback(message)
            
    def get_all_window_handles(self):
        """获取当前所有窗口的句柄"""
        window_handles = []
        
        def enum_windows_callback(hwnd, ctx):
            # 只处理可见窗口
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                window_handles.append(hwnd)
        
        win32gui.EnumWindows(enum_windows_callback, None)
        return window_handles
    
    def get_window_info(self, hwnd):
        """获取窗口信息"""
        title = win32gui.GetWindowText(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        return {
            'hwnd': hwnd,
            'title': title,
            'rect': rect
        }
    
    def capture_window(self, hwnd, title):
        """捕获窗口截图并保存"""
        try:
            # 确保窗口在前台
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)  # 等待窗口切换完成
            
            # 获取窗口位置和大小
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect
            width = right - left
            height = bottom - top
            
            # 捕获窗口截图
            img = ImageGrab.grab(bbox=(left, top, right, bottom))
            
            # 生成保存文件名（使用时间戳和窗口标题）
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            # 清理文件名中的非法字符
            safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '_', '-', '.'))[:50]
            filename = f"{timestamp}_{safe_title}.png"
            filepath = os.path.join(self.save_dir, filename)
            
            # 保存截图
            img.save(filepath)
            self.log(f"成功截取窗口 '{title}' 的截图并保存为: {filepath}")
            
        except Exception as e:
            error_msg = f"截取窗口 '{title}' 截图时出错: {str(e)}"
            logger.error(error_msg)
            if self.log_callback:
                self.log_callback(error_msg)
    
    def monitor_windows(self):
        """监控新窗口的主循环"""
        self.running = True
        self.log("开始监控新窗口...")
        self.log(f"检查间隔: {self.check_interval}秒")
        self.log(f"截图保存目录: {os.path.abspath(self.save_dir)}")
        
        # 初始化已知窗口列表
        self.known_windows = set(self.get_all_window_handles())
        
        try:
            while self.running:
                # 获取当前所有窗口
                current_windows = set(self.get_all_window_handles())
                
                # 找出新出现的窗口
                new_windows = current_windows - self.known_windows
                
                # 处理新窗口
                for hwnd in new_windows:
                    window_info = self.get_window_info(hwnd)
                    title = window_info['title']
                    self.log(f"发现新窗口: '{title}' (句柄: {hwnd})")
                    self.capture_window(hwnd, title)
                
                # 更新已知窗口列表
                self.known_windows = current_windows
                
                # 等待一段时间后再次检查，同时允许提前终止
                wait_time = 0
                while wait_time < self.check_interval and self.running:
                    time.sleep(0.1)
                    wait_time += 0.1
                
        except Exception as e:
            error_msg = f"监控过程中出错: {str(e)}"
            logger.error(error_msg)
            if self.log_callback:
                self.log_callback(error_msg)
        
        if self.running:
            self.running = False
            
    def stop_monitoring(self):
        """停止监控"""
        self.running = False
        self.log("停止监控新窗口")

class WindowMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("窗口监控工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # 设置中文字体支持
        self.setup_fonts()
        
        # 创建监控器实例（尚未启动）
        self.monitor = None
        self.monitor_thread = None
        
        # 创建界面元素
        self.create_widgets()
        
        # 窗口关闭时的处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_fonts(self):
        """设置中文字体支持"""
        try:
            # 尝试设置中文字体
            self.default_font = ('Microsoft YaHei UI', 10)
            self.button_font = ('Microsoft YaHei UI', 10, 'bold')
            self.title_font = ('Microsoft YaHei UI', 12, 'bold')
        except:
            # 如果出错，使用默认字体
            self.default_font = ('Arial', 10)
            self.button_font = ('Arial', 10, 'bold')
            self.title_font = ('Arial', 12, 'bold')
    
    def create_widgets(self):
        """创建界面元素"""
        # 创建标题标签
        title_label = tk.Label(self.root, text="窗口监控工具", font=self.title_font, pady=10)
        title_label.pack(fill=tk.X)
        
        # 创建操作按钮框架
        button_frame = tk.Frame(self.root, pady=10)
        button_frame.pack(fill=tk.X, padx=20)
        
        # 创建开始按钮
        self.start_button = tk.Button(
            button_frame, 
            text="开始监控", 
            font=self.button_font, 
            command=self.start_monitoring,
            width=15, 
            bg="#4CAF50", 
            fg="white",
            relief=tk.RAISED
        )
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        # 创建停止按钮
        self.stop_button = tk.Button(
            button_frame, 
            text="停止监控", 
            font=self.button_font, 
            command=self.stop_monitoring,
            width=15, 
            bg="#f44336", 
            fg="white",
            state=tk.DISABLED,
            relief=tk.RAISED
        )
        self.stop_button.pack(side=tk.LEFT, padx=10)
        
        # 创建退出按钮
        self.exit_button = tk.Button(
            button_frame, 
            text="退出", 
            font=self.button_font, 
            command=self.on_closing,
            width=15, 
            bg="#2196F3", 
            fg="white",
            relief=tk.RAISED
        )
        self.exit_button.pack(side=tk.RIGHT, padx=10)
        
        # 创建日志显示区域
        log_label = tk.Label(self.root, text="运行日志:", font=self.default_font, pady=5)
        log_label.pack(fill=tk.X, padx=20)
        
        self.log_text = scrolledtext.ScrolledText(
            self.root, 
            font=self.default_font, 
            wrap=tk.WORD, 
            bg="#f0f0f0",
            bd=1,
            relief=tk.SUNKEN
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        self.log_text.config(state=tk.DISABLED)
        
        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = tk.Label(
            self.root, 
            textvariable=self.status_var, 
            font=self.default_font, 
            bd=1, 
            relief=tk.SUNKEN, 
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def append_log(self, message):
        """向日志区域添加消息"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # 自动滚动到最后一行
        self.log_text.config(state=tk.DISABLED)
    
    def start_monitoring(self):
        """开始监控新窗口"""
        if self.monitor_thread and self.monitor_thread.is_alive():
            messagebox.showinfo("提示", "监控已经在运行中")
            return
        
        # 创建监控器实例
        self.monitor = WindowMonitor(log_callback=self.append_log)
        
        # 在新线程中运行监控器
        self.monitor_thread = threading.Thread(target=self.monitor.monitor_windows)
        self.monitor_thread.daemon = True  # 设置为守护线程，主线程结束时自动终止
        self.monitor_thread.start()
        
        # 更新按钮状态和状态栏
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.status_var.set("监控中...")
        
    def stop_monitoring(self):
        """停止监控新窗口"""
        if self.monitor and self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor.stop_monitoring()
            self.monitor_thread.join(2.0)  # 等待线程结束，最多等待2秒
        
        # 更新按钮状态和状态栏
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_var.set("已停止监控")
    
    def on_closing(self):
        """窗口关闭时的处理"""
        # 停止监控
        if self.monitor and hasattr(self.monitor, 'running') and self.monitor.running:
            self.stop_monitoring()
        
        # 确认退出
        if messagebox.askyesno("确认退出", "确定要退出程序吗？"):
            self.root.destroy()

if __name__ == "__main__":
    # 创建主窗口并启动应用
    root = tk.Tk()
    app = WindowMonitorApp(root)
    root.mainloop()
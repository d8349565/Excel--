#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具 GUI 控制面板
使用wxPython创建的图形界面控制应用
"""

import wx
import wx.lib.newevent
import subprocess
import threading
import time
import json
import os
import sys
import webbrowser
from datetime import datetime
import configparser
import requests

# 自定义事件
UpdateStatusEvent, EVT_UPDATE_STATUS = wx.lib.newevent.NewEvent()

class ConfigManager:
    """配置管理器"""
    
    def __init__(self):
        self.config_file = os.path.join(os.path.dirname(__file__), 'gui_config.ini')
        self.config = configparser.ConfigParser()
        self.load_config()
    
    def load_config(self):
        """加载配置"""
        self.config.read(self.config_file, encoding='utf-8')
        
        # 默认配置
        if not self.config.has_section('server'):
            self.config.add_section('server')
            self.config.set('server', 'host', '0.0.0.0')  # 改为0.0.0.0
            self.config.set('server', 'port', '5000')
            self.config.set('server', 'debug', 'False')
            self.config.set('server', 'auto_start', 'False')
            
        if not self.config.has_section('auth'):
            self.config.add_section('auth')
            self.config.set('auth', 'access_password', '123456')
            self.config.set('auth', 'admin_password', 'admin2025')
            
        if not self.config.has_section('limits'):
            self.config.add_section('limits')
            self.config.set('limits', 'max_file_size', '100')
            self.config.set('limits', 'max_concurrent_tasks', '5')
            self.config.set('limits', 'file_retention_days', '1')
            self.config.set('limits', 'preview_rows', '50')
            
        self.save_config()
    
    def save_config(self):
        """保存配置"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
    
    def get(self, section, key, fallback=None):
        """获取配置值"""
        return self.config.get(section, key, fallback=fallback)
    
    def set(self, section, key, value):
        """设置配置值"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
        self.save_config()

class ServerManager:
    """服务器管理器"""
    
    def __init__(self, parent):
        self.parent = parent
        self.process = None
        self.running = False
        self.monitor_thread = None
        self.stop_monitoring = threading.Event()
        
    def start_server(self, config_manager):
        """启动服务器"""
        if self.running:
            return False, "服务器已在运行"
            
        try:
            # 设置环境变量
            env = os.environ.copy()
            env['ACCESS_PASSWORD'] = config_manager.get('auth', 'access_password')
            env['ADMIN_PASSWORD'] = config_manager.get('auth', 'admin_password')
            env['SERVER_HOST'] = config_manager.get('server', 'host')
            env['SERVER_PORT'] = config_manager.get('server', 'port')
            env['SERVER_DEBUG'] = config_manager.get('server', 'debug')
            
            # 传递系统限制参数
            env['MAX_FILE_SIZE_MB'] = config_manager.get('limits', 'max_file_size')
            env['MAX_CONCURRENT_TASKS'] = config_manager.get('limits', 'max_concurrent_tasks')
            
            # 构建启动命令
            host = config_manager.get('server', 'host')
            port = config_manager.get('server', 'port')
            debug = config_manager.get('server', 'debug') == 'True'
            
            script_dir = os.path.dirname(__file__)
            app_path = os.path.join(script_dir, 'app.py')
            
            cmd = [sys.executable, app_path]
            
            # 启动进程
            self.process = subprocess.Popen(
                cmd,
                env=env,
                cwd=script_dir,
                stdout=subprocess.DEVNULL,  # 不捕获输出，避免缓冲区阻塞
                stderr=subprocess.DEVNULL,  # 不捕获错误输出
                stdin=subprocess.DEVNULL,   # 不提供输入
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            self.running = True
            
            # 启动监控线程
            self.start_monitoring()
            
            # 显示访问地址
            access_host = host
            if host == '0.0.0.0':
                import socket
                try:
                    hostname = socket.gethostname()
                    access_host = socket.gethostbyname(hostname)
                    if access_host == '127.0.0.1':
                        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                        try:
                            s.connect(('8.8.8.8', 80))
                            access_host = s.getsockname()[0]
                        except:
                            access_host = '127.0.0.1'
                        finally:
                            s.close()
                except:
                    access_host = '127.0.0.1'
            
            return True, f"服务器启动成功，访问地址: http://{access_host}:{port}"
            
        except Exception as e:
            return False, f"启动失败: {str(e)}"
    
    def stop_server(self):
        """停止服务器"""
        if not self.running:
            return False, "服务器未运行"
            
        try:
            self.stop_monitoring_thread()
            
            if self.process:
                self.process.terminate()
                
                # 等待进程结束
                try:
                    self.process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    self.process.kill()
                    
            self.process = None
            self.running = False
            
            return True, "服务器已停止"
            
        except Exception as e:
            return False, f"停止失败: {str(e)}"
    
    def start_monitoring(self):
        """启动监控线程"""
        self.stop_monitoring.clear()
        self.monitor_thread = threading.Thread(target=self._monitor_server)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
    
    def stop_monitoring_thread(self):
        """停止监控线程"""
        self.stop_monitoring.set()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
    
    def _monitor_server(self):
        """监控服务器状态"""
        startup_time = time.time()
        health_check_interval = 10  # 每10秒检查一次健康状态
        last_health_check = 0
        
        while not self.stop_monitoring.is_set():
            current_time = time.time()
            
            if self.process:
                poll = self.process.poll()
                if poll is not None:
                    # 进程已结束
                    self.running = False
                    self.process = None
                    
                    # 发送状态更新事件
                    if current_time - startup_time < 5:
                        # 如果启动后很快就结束，可能是启动失败
                        evt = UpdateStatusEvent(status="stopped", message=f"服务器启动失败，退出代码: {poll}")
                    else:
                        evt = UpdateStatusEvent(status="stopped", message="服务器进程意外停止")
                    wx.PostEvent(self.parent, evt)
                    break
                
                # 定期进行健康检查（启动5秒后开始）
                if (current_time - startup_time > 5 and 
                    current_time - last_health_check > health_check_interval):
                    
                    last_health_check = current_time
                    
                    # 获取配置的端口
                    try:
                        port = self.parent.config_manager.get('server', 'port', '5000')
                        if not self.is_server_responding('127.0.0.1', port):
                            # 服务器无响应，尝试重启
                            wx.PostEvent(self.parent, UpdateStatusEvent(
                                status="warning", 
                                message="检测到服务器无响应，正在尝试重启..."
                            ))
                            
                            # 重启服务器
                            wx.CallAfter(self.parent.restart_server_if_needed)
                            break
                    except Exception as e:
                        wx.PostEvent(self.parent, UpdateStatusEvent(
                            status="warning", 
                            message=f"健康检查失败: {str(e)}"
                        ))
            
            time.sleep(2)
    
    def is_server_responding(self, host='127.0.0.1', port='5000'):
        """检查服务器是否响应"""
        try:
            # 设置较短的超时时间
            response = requests.get(f'http://{host}:{port}/', timeout=3)
            return response.status_code == 200
        except requests.exceptions.RequestException:
            # 网络请求异常
            return False
        except Exception:
            # 其他异常
            return False

class MainFrame(wx.Frame):
    """主窗口"""
    
    def __init__(self):
        super().__init__(None, title="Excel工具控制面板", size=(800, 600))
        
        # 初始化管理器
        self.config_manager = ConfigManager()
        self.server_manager = ServerManager(self)
        
        # 绑定自定义事件
        self.Bind(EVT_UPDATE_STATUS, self.on_status_update)
        
        # 创建界面
        self.create_ui()
        
        # 设置图标（如果存在）
        self.setup_icon()
        
        # 强制刷新任务栏图标
        self.refresh_taskbar_icon()
        
        # 居中显示
        self.Center()
        
        # 检查自动启动
        if self.config_manager.get('server', 'auto_start') == 'True':
            wx.CallAfter(self.start_server)
    
    def setup_icon(self):
        """设置窗口和任务栏图标"""
        try:
            # 尝试多种图标格式和位置
            script_dir = os.path.dirname(__file__)
            icon_files = [
                os.path.join(script_dir, 'static', 'app_icon.ico'),
                os.path.join(script_dir, 'static', 'icon.ico'),
                os.path.join(script_dir, 'static', 'icon.png'),
                os.path.join(script_dir, 'app_icon.ico'),
                os.path.join(script_dir, 'icon.ico'),
                os.path.join(script_dir, 'icon.png')
            ]
            
            icon_set = False
            for icon_path in icon_files:
                if os.path.exists(icon_path):
                    try:
                        if icon_path.endswith('.ico'):
                            # ICO格式 - 最佳支持Windows任务栏
                            icon = wx.Icon(icon_path, wx.BITMAP_TYPE_ICO)
                        elif icon_path.endswith('.png'):
                            # PNG格式
                            icon = wx.Icon()
                            icon.LoadFile(icon_path, wx.BITMAP_TYPE_PNG)
                        
                        if icon.IsOk():
                            self.SetIcon(icon)
                            # 同时设置任务栏图标
                            if hasattr(self, 'taskbar_icon'):
                                self.taskbar_icon.SetIcon(icon, "Excel工具控制面板")
                            
                            print(f"成功设置图标: {icon_path}")
                            icon_set = True
                            break
                    except Exception as e:
                        print(f"加载图标失败 {icon_path}: {e}")
                        continue
            
            if not icon_set:
                print("未找到可用的图标文件")
                # 创建默认图标
                self.create_default_icon()
            
        except Exception as e:
            print(f"设置图标时出错: {e}")
            
    def create_default_icon(self):
        """创建默认图标"""
        try:
            # 创建一个简单的默认图标
            import wx.lib.embeddedimage
            
            # 创建32x32的位图
            bmp = wx.Bitmap(32, 32)
            dc = wx.MemoryDC(bmp)
            
            # 绘制简单图标
            dc.SetBackground(wx.Brush(wx.Colour(76, 175, 80)))  # 绿色背景
            dc.Clear()
            
            # 绘制边框
            dc.SetPen(wx.Pen(wx.Colour(255, 255, 255), 2))
            dc.DrawRectangle(2, 2, 28, 28)
            
            # 绘制网格
            dc.SetPen(wx.Pen(wx.Colour(255, 255, 255), 1))
            for i in range(1, 4):
                x = 8 + i * 6
                y = 8 + i * 6
                dc.DrawLine(8, y, 24, y)
                dc.DrawLine(x, 8, x, 24)
            
            dc.SelectObject(wx.NullBitmap)
            
            # 转换为图标
            icon = wx.Icon()
            icon.CopyFromBitmap(bmp)
            
            if icon.IsOk():
                self.SetIcon(icon)
                print("使用默认生成的图标")
                
        except Exception as e:
            print(f"创建默认图标失败: {e}")
    
    def refresh_taskbar_icon(self):
        """强制刷新任务栏图标"""
        try:
            if os.name == 'nt':  # Windows系统
                import ctypes
                from ctypes import wintypes
                
                # 获取窗口句柄
                hwnd = self.GetHandle()
                
                # 强制刷新窗口图标
                user32 = ctypes.windll.user32
                
                # 获取当前图标
                icon_handle = user32.SendMessageW(hwnd, 0x7F, 1, 0)  # WM_GETICON, ICON_BIG
                if icon_handle:
                    # 重新设置图标
                    user32.SendMessageW(hwnd, 0x80, 1, icon_handle)  # WM_SETICON, ICON_BIG
                    user32.SendMessageW(hwnd, 0x80, 0, icon_handle)  # WM_SETICON, ICON_SMALL
                
                print("已刷新任务栏图标")
                
        except Exception as e:
            print(f"刷新任务栏图标失败: {e}")
    
    def create_ui(self):
        """创建用户界面"""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 创建各个面板
        self.create_header_panel(panel, main_sizer)
        self.create_server_panel(panel, main_sizer)
        self.create_config_panel(panel, main_sizer)
        self.create_status_panel(panel, main_sizer)
        self.create_button_panel(panel, main_sizer)
        
        panel.SetSizer(main_sizer)
    
    def create_header_panel(self, parent, sizer):
        """创建头部面板"""
        header_panel = wx.Panel(parent)
        header_panel.SetBackgroundColour(wx.Colour(240, 248, 255))
        
        header_sizer = wx.BoxSizer(wx.VERTICAL)
        
        title = wx.StaticText(header_panel, label="Excel/CSV 汇总工具控制面板")
        title_font = title.GetFont()
        title_font.PointSize += 4
        title_font = title_font.Bold()
        title.SetFont(title_font)
        
        subtitle = wx.StaticText(header_panel, label="管理服务器运行状态和配置参数")
        subtitle.SetForegroundColour(wx.Colour(100, 100, 100))
        
        header_sizer.Add(title, 0, wx.ALL | wx.CENTER, 10)
        header_sizer.Add(subtitle, 0, wx.ALL | wx.CENTER, 5)
        
        header_panel.SetSizer(header_sizer)
        sizer.Add(header_panel, 0, wx.EXPAND | wx.ALL, 5)
    
    def create_server_panel(self, parent, sizer):
        """创建服务器控制面板"""
        box = wx.StaticBox(parent, label="服务器控制")
        box_sizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        
        # 端口设置
        addr_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        addr_sizer.Add(wx.StaticText(parent, label="端口:"), 0, wx.ALL | wx.CENTER, 5)
        
        self.port_ctrl = wx.SpinCtrl(parent, size=(80, -1), 
                                   min=1000, max=65535)
        self.port_ctrl.SetValue(int(self.config_manager.get('server', 'port')))
        addr_sizer.Add(self.port_ctrl, 0, wx.ALL, 5)
        
        self.debug_cb = wx.CheckBox(parent, label="调试模式")
        self.debug_cb.SetValue(self.config_manager.get('server', 'debug') == 'True')
        addr_sizer.Add(self.debug_cb, 0, wx.ALL | wx.CENTER, 5)
        
        self.auto_start_cb = wx.CheckBox(parent, label="启动时自动开启服务")
        self.auto_start_cb.SetValue(self.config_manager.get('server', 'auto_start') == 'True')
        addr_sizer.Add(self.auto_start_cb, 0, wx.ALL | wx.CENTER, 5)
        
        box_sizer.Add(addr_sizer, 0, wx.ALL, 5)
        
        # 控制按钮
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.start_btn = wx.Button(parent, label="启动服务器")
        self.start_btn.Bind(wx.EVT_BUTTON, self.on_start_server)
        btn_sizer.Add(self.start_btn, 0, wx.ALL, 5)
        
        self.stop_btn = wx.Button(parent, label="停止服务器")
        self.stop_btn.Bind(wx.EVT_BUTTON, self.on_stop_server)
        self.stop_btn.Enable(False)
        btn_sizer.Add(self.stop_btn, 0, wx.ALL, 5)
        
        self.restart_btn = wx.Button(parent, label="重启服务器")
        self.restart_btn.Bind(wx.EVT_BUTTON, self.on_restart_server)
        self.restart_btn.Enable(False)
        btn_sizer.Add(self.restart_btn, 0, wx.ALL, 5)
        
        self.open_btn = wx.Button(parent, label="打开网页")
        self.open_btn.Bind(wx.EVT_BUTTON, self.on_open_web)
        self.open_btn.Enable(False)
        btn_sizer.Add(self.open_btn, 0, wx.ALL, 5)
        
        box_sizer.Add(btn_sizer, 0, wx.ALL, 5)
        
        sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)
    
    def create_config_panel(self, parent, sizer):
        """创建配置面板"""
        notebook = wx.Notebook(parent)
        
        # 认证配置页
        auth_panel = wx.Panel(notebook)
        self.create_auth_config(auth_panel)
        notebook.AddPage(auth_panel, "用户认证")
        
        # 限制配置页
        limits_panel = wx.Panel(notebook)
        self.create_limits_config(limits_panel)
        notebook.AddPage(limits_panel, "系统限制")
        
        sizer.Add(notebook, 1, wx.EXPAND | wx.ALL, 5)
    
    def create_auth_config(self, parent):
        """创建认证配置"""
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 访问密码
        access_box = wx.StaticBox(parent, label="用户访问密码")
        access_sizer = wx.StaticBoxSizer(access_box, wx.VERTICAL)
        
        access_grid = wx.FlexGridSizer(2, 2, 10, 10)
        access_grid.Add(wx.StaticText(parent, label="访问密码:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.access_pwd_ctrl = wx.TextCtrl(parent, value=self.config_manager.get('auth', 'access_password'),
                                         style=wx.TE_PASSWORD, size=(200, -1))
        access_grid.Add(self.access_pwd_ctrl, 1, wx.EXPAND)
        
        access_sizer.Add(access_grid, 0, wx.ALL | wx.EXPAND, 10)
        sizer.Add(access_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        # 管理员密码
        admin_box = wx.StaticBox(parent, label="管理员密码")
        admin_sizer = wx.StaticBoxSizer(admin_box, wx.VERTICAL)
        
        admin_grid = wx.FlexGridSizer(2, 2, 10, 10)
        admin_grid.Add(wx.StaticText(parent, label="管理员密码:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.admin_pwd_ctrl = wx.TextCtrl(parent, value=self.config_manager.get('auth', 'admin_password'),
                                        style=wx.TE_PASSWORD, size=(200, -1))
        admin_grid.Add(self.admin_pwd_ctrl, 1, wx.EXPAND)
        
        admin_sizer.Add(admin_grid, 0, wx.ALL | wx.EXPAND, 10)
        sizer.Add(admin_sizer, 0, wx.EXPAND | wx.ALL, 5)
        
        parent.SetSizer(sizer)
    
    def create_limits_config(self, parent):
        """创建限制配置"""
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        grid = wx.FlexGridSizer(2, 2, 15, 10)  # 改为2行
        grid.AddGrowableCol(1)
        
        # 最大文件大小
        grid.Add(wx.StaticText(parent, label="单文件大小限制 (MB):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.file_size_ctrl = wx.SpinCtrl(parent, size=(100, -1), min=1, max=1000)
        self.file_size_ctrl.SetValue(int(self.config_manager.get('limits', 'max_file_size')))
        grid.Add(self.file_size_ctrl, 0, wx.EXPAND)
        
        # 最大并发任务数
        grid.Add(wx.StaticText(parent, label="最大并发任务数:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.concurrent_tasks_ctrl = wx.SpinCtrl(parent, size=(100, -1), min=1, max=20)
        self.concurrent_tasks_ctrl.SetValue(int(self.config_manager.get('limits', 'max_concurrent_tasks')))
        grid.Add(self.concurrent_tasks_ctrl, 0, wx.EXPAND)
        
        sizer.Add(grid, 0, wx.ALL | wx.EXPAND, 20)
        parent.SetSizer(sizer)
    
    def create_status_panel(self, parent, sizer):
        """创建状态面板"""
        box = wx.StaticBox(parent, label="服务器状态")
        box_sizer = wx.StaticBoxSizer(box, wx.VERTICAL)
        
        self.status_text = wx.TextCtrl(parent, style=wx.TE_MULTILINE | wx.TE_READONLY, size=(-1, 120))
        self.status_text.SetFont(wx.Font(9, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        
        self.log_message("Excel工具控制面板已启动")
        
        box_sizer.Add(self.status_text, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)
    
    def create_button_panel(self, parent, sizer):
        """创建底部按钮面板"""
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        save_btn = wx.Button(parent, label="保存配置")
        save_btn.Bind(wx.EVT_BUTTON, self.on_save_config)
        btn_sizer.Add(save_btn, 0, wx.ALL, 5)
        
        reset_btn = wx.Button(parent, label="重置配置")
        reset_btn.Bind(wx.EVT_BUTTON, self.on_reset_config)
        btn_sizer.Add(reset_btn, 0, wx.ALL, 5)
        
        btn_sizer.AddStretchSpacer()
        
        about_btn = wx.Button(parent, label="关于")
        about_btn.Bind(wx.EVT_BUTTON, self.on_about)
        btn_sizer.Add(about_btn, 0, wx.ALL, 5)
        
        exit_btn = wx.Button(parent, label="退出")
        exit_btn.Bind(wx.EVT_BUTTON, self.on_exit)
        btn_sizer.Add(exit_btn, 0, wx.ALL, 5)
        
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 5)
    
    def log_message(self, message):
        """记录日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_text = f"[{timestamp}] {message}\n"
        wx.CallAfter(self.status_text.AppendText, log_text)
    
    def on_start_server(self, event):
        """启动服务器"""
        self.save_server_config()
        success, message = self.server_manager.start_server(self.config_manager)
        
        if success:
            self.start_btn.Enable(False)
            self.stop_btn.Enable(True)
            self.restart_btn.Enable(True)
            self.open_btn.Enable(True)
        
        self.log_message(message)
    
    def on_stop_server(self, event):
        """停止服务器"""
        success, message = self.server_manager.stop_server()
        
        if success:
            self.start_btn.Enable(True)
            self.stop_btn.Enable(False)
            self.restart_btn.Enable(False)
            self.open_btn.Enable(False)
        
        self.log_message(message)
    
    def on_restart_server(self, event):
        """重启服务器"""
        self.log_message("正在重启服务器...")
        
        # 先停止服务器
        success, message = self.server_manager.stop_server()
        if success:
            # 更新UI状态
            self.start_btn.Enable(False)
            self.stop_btn.Enable(False)
            self.restart_btn.Enable(False)
            self.open_btn.Enable(False)
            
            # 延迟1秒后重新启动
            wx.CallLater(1000, self.restart_server_complete)
        else:
            self.log_message(f"重启失败: {message}")
    
    def restart_server_complete(self):
        """完成服务器重启"""
        success, message = self.server_manager.start_server(self.config_manager)
        
        if success:
            self.start_btn.Enable(False)
            self.stop_btn.Enable(True)
            self.restart_btn.Enable(True)
            self.open_btn.Enable(True)
            self.log_message("服务器重启完成")
        else:
            self.start_btn.Enable(True)
            self.stop_btn.Enable(False)
            self.restart_btn.Enable(False)
            self.open_btn.Enable(False)
            self.log_message(f"重启失败: {message}")
    
    def on_open_web(self, event):
        """打开网页"""
        port = self.port_ctrl.GetValue()
        
        # 服务器固定监听0.0.0.0，使用本机IP访问
        import socket
        try:
            # 获取本机IP地址
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            # 如果获取失败，使用127.0.0.1
            if local_ip == '127.0.0.1':
                # 尝试另一种方法获取真实IP
                s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                try:
                    s.connect(('8.8.8.8', 80))
                    local_ip = s.getsockname()[0]
                except:
                    local_ip = '127.0.0.1'
                finally:
                    s.close()
            url = f"http://{local_ip}:{port}"
        except:
            url = f"http://127.0.0.1:{port}"
        
        try:
            webbrowser.open(url)
            self.log_message(f"已打开网页: {url}")
        except Exception as e:
            self.log_message(f"打开网页失败: {str(e)}")
    
    def on_save_config(self, event):
        """保存配置"""
        self.save_server_config()
        self.save_auth_config()
        self.save_limits_config()
        self.log_message("配置已保存")
        wx.MessageBox("配置保存成功！", "提示", wx.OK | wx.ICON_INFORMATION)
    
    def save_server_config(self):
        """保存服务器配置"""
        # 服务器地址固定使用0.0.0.0，不允许修改
        self.config_manager.set('server', 'host', '0.0.0.0')
        self.config_manager.set('server', 'port', str(self.port_ctrl.GetValue()))
        self.config_manager.set('server', 'debug', str(self.debug_cb.GetValue()))
        self.config_manager.set('server', 'auto_start', str(self.auto_start_cb.GetValue()))
    
    def save_auth_config(self):
        """保存认证配置"""
        self.config_manager.set('auth', 'access_password', self.access_pwd_ctrl.GetValue())
        self.config_manager.set('auth', 'admin_password', self.admin_pwd_ctrl.GetValue())
    
    def save_limits_config(self):
        """保存限制配置"""
        self.config_manager.set('limits', 'max_file_size', str(self.file_size_ctrl.GetValue()))
        self.config_manager.set('limits', 'max_concurrent_tasks', str(self.concurrent_tasks_ctrl.GetValue()))
    
    def on_reset_config(self, event):
        """重置配置"""
        if wx.MessageBox("确定要重置所有配置吗？", "确认", wx.YES_NO | wx.ICON_QUESTION) == wx.YES:
            # 删除配置文件
            if os.path.exists(self.config_manager.config_file):
                os.remove(self.config_manager.config_file)
            
            # 重新加载默认配置
            self.config_manager.load_config()
            
            # 更新界面
            self.load_config_to_ui()
            
            self.log_message("配置已重置为默认值")
            wx.MessageBox("配置重置成功！", "提示", wx.OK | wx.ICON_INFORMATION)
    
    def load_config_to_ui(self):
        """将配置加载到界面"""
        # 服务器配置
        # 服务器地址固定为0.0.0.0，不在界面显示
        self.port_ctrl.SetValue(int(self.config_manager.get('server', 'port')))
        self.debug_cb.SetValue(self.config_manager.get('server', 'debug') == 'True')
        self.auto_start_cb.SetValue(self.config_manager.get('server', 'auto_start') == 'True')
        
        # 认证配置
        self.access_pwd_ctrl.SetValue(self.config_manager.get('auth', 'access_password'))
        self.admin_pwd_ctrl.SetValue(self.config_manager.get('auth', 'admin_password'))
        
        # 限制配置
        self.file_size_ctrl.SetValue(int(self.config_manager.get('limits', 'max_file_size')))
        self.concurrent_tasks_ctrl.SetValue(int(self.config_manager.get('limits', 'max_concurrent_tasks')))
    
    def on_about(self, event):
        """显示关于对话框"""
        info = wx.adv.AboutDialogInfo()
        info.SetName("Excel工具控制面板")
        info.SetVersion("1.0.0")
        info.SetDescription("Excel/CSV文件汇总工具的图形控制面板\n\n"
                           "功能特性:\n"
                           "• 服务器启动/停止控制\n"
                           "• 实时状态监控\n"
                           "• 参数配置管理\n"
                           "• 用户认证设置")
        info.SetCopyright("© 2025 Excel工具项目")
        info.SetWebSite("https://github.com/d8349565/Excel--")
        
        wx.adv.AboutBox(info)
    
    def on_exit(self, event):
        """退出应用"""
        if self.server_manager.running:
            if wx.MessageBox("服务器正在运行，确定要退出吗？", "确认退出", 
                           wx.YES_NO | wx.ICON_QUESTION) == wx.NO:
                return
            
            self.server_manager.stop_server()
        
        self.Close()
    
    def on_status_update(self, event):
        """处理状态更新事件"""
        if event.status == "stopped":
            self.start_btn.Enable(True)
            self.stop_btn.Enable(False)
            self.restart_btn.Enable(False)
            self.open_btn.Enable(False)
        
        self.log_message(event.message)
    
    def restart_server_if_needed(self):
        """在检测到问题时重启服务器"""
        if self.server_manager.running:
            self.log_message("正在重启服务器...")
            
            # 停止当前服务器
            success, message = self.server_manager.stop_server()
            if success:
                # 等待一秒后重新启动
                wx.CallLater(1000, self.restart_server_delayed)
            else:
                self.log_message(f"重启失败: {message}")
    
    def restart_server_delayed(self):
        """延迟重启服务器"""
        success, message = self.server_manager.start_server(self.config_manager)
        if success:
            self.start_btn.Enable(False)
            self.stop_btn.Enable(True)
            self.restart_btn.Enable(True)
            self.open_btn.Enable(True)
            self.log_message("服务器重启成功")
        else:
            self.log_message(f"服务器重启失败: {message}")
            # 更新按钮状态
            self.start_btn.Enable(True)
            self.stop_btn.Enable(False)
            self.restart_btn.Enable(False)
            self.open_btn.Enable(False)
    
    def start_server(self):
        """程序启动时自动启动服务器"""
        if not self.server_manager.running:
            self.on_start_server(None)

class ExcelToolControllerApp(wx.App):
    """应用程序类"""
    
    def OnInit(self):
        # 设置应用程序ID（Windows任务栏分组）
        try:
            # Windows系统设置应用程序ID
            if os.name == 'nt':
                import ctypes
                # 设置应用程序用户模型ID
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ExcelTool.GUI.Controller.1.0")
        except Exception as e:
            print(f"设置应用程序ID失败: {e}")
        
        # 设置应用程序名称
        self.SetAppName("Excel工具控制面板")
        self.SetAppDisplayName("Excel工具控制面板")
        
        frame = MainFrame()
        frame.Show(True)
        
        # 设置主框架
        self.SetTopWindow(frame)
        
        return True

if __name__ == '__main__':
    app = ExcelToolControllerApp()
    app.MainLoop()

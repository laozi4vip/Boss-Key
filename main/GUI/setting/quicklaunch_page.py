"""
快捷启动设置页面
"""
import wx
import wx.lib.scrolledpanel as scrolled
from core.config import Config
import core.tools as tools
import os
import json
import psutil
import ctypes

# Windows API
user32 = ctypes.windll.user32
SW_MINIMIZE = 6
SW_RESTORE = 9
SW_TOPMOST = 0x0008

class QuickLaunchPage(scrolled.ScrolledPanel):
    def __init__(self, parent):
        super().__init__(parent)
        self.init_ui()
        self.setup_scrolling()
        
    def init_ui(self):
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 标题
        title = wx.StaticText(self, label="快捷启动设置", style=wx.ALIGN_CENTER)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL|wx.ALIGN_CENTER, 10)
        
        # 说明
        desc = wx.StaticText(self, label="设置快捷键来快速启动或切换程序窗口")
        sizer.Add(desc, 0, wx.ALL, 5)
        
        # 程序列表
        list_box = wx.StaticBox(self, label="已配置的程序")
        list_sizer = wx.StaticBoxSizer(list_box, wx.VERTICAL)
        
        # 列表控件
        self.program_list = wx.ListCtrl(self, style=wx.LC_REPORT|wx.LC_SINGLE_SEL)
        self.program_list.InsertColumn(0, "程序名称", width=150)
        self.program_list.InsertColumn(1, "程序路径", width=250)
        self.program_list.InsertColumn(2, "热键", width=100)
        
        list_sizer.Add(self.program_list, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(list_sizer, 1, wx.EXPAND|wx.ALL, 10)
        
        # 按钮
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.add_btn = wx.Button(self, label="添加程序")
        self.add_btn.Bind(wx.EVT_BUTTON, self.on_add)
        
        self.from_running_btn = wx.Button(self, label="从运行程序绑定")
        self.from_running_btn.Bind(wx.EVT_BUTTON, self.on_from_running)
        
        self.del_btn = wx.Button(self, label="删除")
        self.del_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        
        self.hotkey_btn = wx.Button(self, label="设置热键")
        self.hotkey_btn.Bind(wx.EVT_BUTTON, self.on_set_hotkey)
        
        btn_sizer.Add(self.add_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.from_running_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.del_btn, 0, wx.ALL, 5)
        btn_sizer.Add(self.hotkey_btn, 0, wx.ALL, 5)
        
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)
        
        self.SetSizer(sizer)
        
        # 加载配置
        self.load_config()
    
    def setup_scrolling(self):
        self.SetupScrolling()
    
    def load_config(self):
        """加载配置"""
        config_path = os.path.join(Config.root_path, "quicklaunch_config.json")
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except:
                self.config = {"programs": []}
        else:
            self.config = {"programs": []}
        
        self.refresh_list()
    
    def save_config(self):
        """保存配置"""
        config_path = os.path.join(Config.root_path, "quicklaunch_config.json")
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            # 保存后重新绑定热键
            if hasattr(Config, 'HotkeyListener') and Config.HotkeyListener:
                Config.HotkeyListener.reBind()
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def refresh_list(self):
        """刷新列表"""
        self.program_list.DeleteAllItems()
        for p in self.config.get('programs', []):
            self.program_list.Append([p.get('name', ''), p.get('path', ''), p.get('hotkey', '')])
    
    def on_add(self, event):
        """添加程序"""
        dialog = wx.Dialog(self, title="添加程序", size=(500, 200))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 程序名称
        name_sizer = wx.BoxSizer(wx.HORIZONTAL)
        name_sizer.Add(wx.StaticText(panel, label="程序名称:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        name_ctrl = wx.TextCtrl(panel)
        name_sizer.Add(name_ctrl, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(name_sizer, 0, wx.EXPAND|wx.ALL, 5)
        
        # 程序路径
        path_sizer = wx.BoxSizer(wx.HORIZONTAL)
        path_sizer.Add(wx.StaticText(panel, label="程序路径:"), 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5)
        path_ctrl = wx.TextCtrl(panel)
        path_sizer.Add(path_ctrl, 1, wx.EXPAND|wx.ALL, 5)
        
        def on_browse(event):
            dlg = wx.FileDialog(panel, "选择程序", wildcard="*.exe", style=wx.FD_OPEN)
            if dlg.ShowModal() == wx.ID_OK:
                path_ctrl.SetValue(dlg.GetPath())
                if not name_ctrl.GetValue():
                    name_ctrl.SetValue(os.path.basename(dlg.GetPath()).replace('.exe', ''))
            dlg.Destroy()
        
        browse_btn = wx.Button(panel, label="浏览")
        browse_btn.Bind(wx.EVT_BUTTON, on_browse)
        path_sizer.Add(browse_btn, 0, wx.ALL, 5)
        sizer.Add(path_sizer, 0, wx.EXPAND|wx.ALL, 5)
        
        # 按钮
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        ok_btn = wx.Button(panel, wx.ID_OK)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL)
        btn_sizer.Add(ok_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        
        if dialog.ShowModal() == wx.ID_OK:
            name = name_ctrl.GetValue().strip()
            path = path_ctrl.GetValue().strip()
            if name and path:
                self.config.setdefault('programs', []).append({
                    'name': name,
                    'path': path,
                    'hotkey': ''
                })
                self.save_config()
                self.refresh_list()
        
        dialog.Destroy()
    
    def on_from_running(self, event):
        """从运行程序绑定 - 使用进程枚举"""
        running = []
        
        # 直接遍历所有进程获取可执行程序路径
        for proc in psutil.process_iter(['exe', 'name']):
            try:
                exe = proc.info.get('exe')
                if exe and exe.endswith('.exe'):
                    exe_name = os.path.basename(exe)
                    # 避免重复
                    if exe_name not in [p['exe'] for p in running]:
                        # 获取窗口标题
                        title = self.get_process_window_title(proc)
                        if title:
                            running.append({
                                'title': title,
                                'exe': exe_name,
                                'path': exe
                            })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        if not running:
            wx.MessageBox("没有找到运行的程序", "提示")
            return
        
        # 按名称排序
        running.sort(key=lambda x: x['title'] or x['exe'])
        
        # 显示选择对话框
        dialog = wx.Dialog(self, title="选择程序", size=(500, 400))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        sizer.Add(wx.StaticText(panel, label="双击选择程序:"), 0, wx.ALL, 5)
        
        listbox = wx.ListBox(panel, size=(-1, 300))
        for p in running:
            display = f"{p['title']} - {p['exe']}" if p['title'] else p['exe']
            listbox.Append(display)
        sizer.Add(listbox, 1, wx.EXPAND|wx.ALL, 5)
        
        selected = [None]
        
        def on_select(event):
            selection = listbox.GetSelection()
            if selection != wx.NOT_FOUND:
                selected[0] = running[selection]
        
        listbox.Bind(wx.EVT_LISTBOX_DCLICK, on_select)
        
        def ok():
            if selected[0]:
                self.config.setdefault('programs', []).append({
                    'name': selected[0]['exe'].replace('.exe', ''),
                    'path': selected[0]['path'],
                    'hotkey': ''
                })
                self.save_config()
                self.refresh_list()
                dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()
    
    def get_process_window_title(self, proc):
        """获取进程的主窗口标题"""
        try:
            pid = proc.pid
            windows = []
            
            def callback(hwnd, wins):
                try:
                    if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                        # 检查是否属于该进程
                        window_pid = ctypes.c_ulong()
                        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                        if window_pid.value == pid:
                            length = user32.GetWindowTextLengthW(hwnd)
                            if length > 0:
                                title = ctypes.create_unicode_buffer(length + 1)
                                user32.GetWindowTextW(hwnd, title, length + 1)
                                wins.append(title.value)
                except:
                    pass
                return True
            
            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
            user32.EnumWindows(WNDENUMPROC(callback), windows)
            
            # 返回第一个非空标题
            for title in windows:
                if title and len(title) > 0:
                    return title
        except:
            pass
        return ""
                self.save_config()
                self.refresh_list()
                dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()
    
    def on_delete(self, event):
        """删除程序"""
        selection = self.program_list.GetFirstSelected()
        if selection >= 0:
            self.config['programs'].pop(selection)
            self.save_config()
            self.refresh_list()
    
    def on_set_hotkey(self, event):
        """设置热键"""
        selection = self.program_list.GetFirstSelected()
        if selection < 0:
            wx.MessageBox("请先选择程序", "提示")
            return
        
        dialog = wx.Dialog(self, title="设置热键", size=(300, 120))
        panel = wx.Panel(dialog)
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        sizer.Add(wx.StaticText(panel, label="按下热键组合 (如 Alt+1):"), 0, wx.ALL|wx.ALIGN_CENTER, 10)
        
        hotkey_ctrl = wx.TextCtrl(panel, size=(200, 30))
        sizer.Add(hotkey_ctrl, 0, wx.ALIGN_CENTER|wx.ALL, 10)
        
        def on_key(event):
            keys = []
            key = event.keysym.lower()
            
            if key in ('alt_l', 'alt_r', 'shift_l', 'shift_r', 'ctrl_l', 'ctrl_r', 'ctrl', 'shift', 'alt', 'mode'):
                return
            
            if event.state & 0x1: keys.append('alt')
            if event.state & 0x4: keys.append('shift')
            if event.state & 0x8: keys.append('ctrl')
            
            key = key.upper()
            if key:
                keys.append(key)
            
            if keys:
                hotkey_ctrl.SetValue('+'.join(keys))
        
        hotkey_ctrl.Bind(wx.EVT_KEY_UP, on_key)
        
        def ok():
            hotkey = hotkey_ctrl.GetValue().strip()
            if hotkey and selection < len(self.config.get('programs', [])):
                self.config['programs'][selection]['hotkey'] = hotkey
                self.save_config()
                self.refresh_list()
                dialog.Destroy()
        
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(wx.Button(panel, wx.ID_OK, "确定"), 0, wx.ALL, 5)
        btn_sizer.Add(wx.Button(panel, wx.ID_CANCEL, "取消"), 0, wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER|wx.ALL, 5)
        
        panel.SetSizer(sizer)
        dialog.ShowModal()
        dialog.Destroy()

"""
QuickLaunch - 快捷启动器
"""
import sys
import os
import json
import subprocess
import ctypes
import time
import psutil

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "quicklaunch_config.json")
LOG_FILE = os.path.join(BASE_DIR, "quicklaunch.log")

def log(msg):
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n")
    except:
        pass

log("启动")

user32 = ctypes.windll.user32

SW_MINIMIZE = 6
SW_RESTORE = 9

DEFAULT_CONFIG = {"programs": []}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return DEFAULT_CONFIG.copy()

def save_config_to_file(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def find_window(exe_name):
    exe_name = exe_name.lower()
    windows = []
    
    def callback(hwnd, wins):
        try:
            if hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    wins.append(hwnd)
        except:
            pass
        return True
    
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p)
    try:
        user32.EnumWindows(WNDENUMPROC(callback), windows)
    except Exception as e:
        log(f"EnumWindows error: {e}")
    
    for hwnd in windows:
        pid = ctypes.c_ulong()
        try:
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            proc = psutil.Process(pid.value)
            if exe_name in proc.name().lower():
                return hwnd
        except:
            pass
    return None

def is_minimized(hwnd):
    try:
        return user32.IsIconic(hwnd)
    except:
        return False

def restore(hwnd):
    try:
        if is_minimized(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
    except:
        pass

def minimize(hwnd):
    try:
        user32.ShowWindow(hwnd, SW_MINIMIZE)
    except:
        pass

def launch(path):
    if os.path.exists(path):
        subprocess.Popen(path)
        return True
    return False

def toggle(program):
    path = program.get('path', '')
    if not path:
        return False
    exe = os.path.basename(path)
    hwnd = find_window(exe)
    log(f"切换 {exe}, hwnd={hwnd}")
    if hwnd:
        if is_minimized(hwnd):
            restore(hwnd)
        else:
            minimize(hwnd)
        return True
    else:
        return launch(path)

# 热键防抖
last_triggered = {}
COOLDOWN = 500  # 毫秒

def check_hotkeys():
    """使用定时器检测热键，带防抖"""
    config = load_config()
    now = time.time() * 1000
    
    # 检测修饰键
    alt = user32.GetAsyncKeyState(0x12) & 0x8000
    ctrl = user32.GetAsyncKeyState(0x11) & 0x8000
    shift = user32.GetAsyncKeyState(0x10) & 0x8000
    
    # 检测数字键
    for i, vk in enumerate([0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x30]):
        if user32.GetAsyncKeyState(vk) & 0x8000:
            key = str(i % 10)
            
            # 构建期望的热键
            expected = None
            if alt: expected = f"alt+{key}"
            elif ctrl: expected = f"ctrl+{key}"
            elif shift: expected = f"shift+{key}"
            
            if expected:
                # 检查冷却时间
                last_time = last_triggered.get(expected, 0)
                if now - last_time > COOLDOWN:
                    # 匹配并触发
                    for program in config.get('programs', []):
                        hotkey = program.get('hotkey', '').lower().replace(' ', '')
                        if hotkey == expected:
                            log(f"触发热键: {hotkey}")
                            toggle(program)
                            last_triggered[expected] = now
                            break
            break

def main():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    
    log("初始化GUI")
    
    root = tk.Tk()
    root.title("QuickLaunch - 快捷启动器")
    root.geometry("600x400")
    
    config = load_config()
    
    # 程序列表
    tk.Label(root, text="程序列表", font=("Arial", 14, "bold")).pack(pady=10)
    
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    tree = ttk.Treeview(frame, columns=("name", "path", "hotkey"), show='headings')
    tree.heading("name", text="程序名称")
    tree.heading("path", text="程序路径")
    tree.heading("hotkey", text="热键")
    tree.column("name", width=120)
    tree.column("path", width=300)
    tree.column("hotkey", width=80)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scrollbar.set)
    
    def refresh():
        tree.delete(*tree.get_children())
        for p in config.get('programs', []):
            tree.insert('', tk.END, values=(p.get('name',''), p.get('path',''), p.get('hotkey','')))
    
    def add_program():
        dialog = tk.Toplevel(root)
        dialog.title("添加程序")
        dialog.geometry("500x150")
        
        tk.Label(dialog, text="名称:").grid(row=0, column=0)
        name_entry = tk.Entry(dialog, width=40)
        name_entry.grid(row=0, column=1, padx=5)
        
        tk.Label(dialog, text="路径:").grid(row=1, column=0)
        path_entry = tk.Entry(dialog, width=40)
        path_entry.grid(row=1, column=1, padx=5)
        
        def browse():
            f = filedialog.askopenfilename(filetypes=[("exe", "*.exe")])
            if f:
                path_entry.delete(0, tk.END)
                path_entry.insert(0, f)
                if not name_entry.get():
                    name_entry.insert(0, os.path.basename(f).replace('.exe',''))
        
        tk.Button(dialog, text="浏览", command=browse).grid(row=1, column=2, padx=5)
        
        def ok():
            if name_entry.get().strip() and path_entry.get().strip():
                config.setdefault('programs', []).append({
                    'name': name_entry.get().strip(),
                    'path': path_entry.get().strip(),
                    'hotkey': ''
                })
                refresh()
                dialog.destroy()
        
        tk.Button(dialog, text="确定", command=ok).grid(row=2, column=1, pady=5)
    
    def delete_program():
        sel = tree.selection()
        if sel:
            idx = tree.index(sel[0])
            config['programs'].pop(idx)
            refresh()
    
    def set_hotkey():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择程序")
            return
        
        idx = tree.index(sel[0])
        
        dialog = tk.Toplevel(root)
        dialog.title("设置热键")
        dialog.geometry("300x100")
        
        tk.Label(dialog, text="热键 (如 alt+1):").pack(pady=10)
        hotkey_var = tk.StringVar()
        tk.Entry(dialog, textvariable=hotkey_var, width=20).pack()
        
        def ok():
            if idx < len(config['programs']):
                config['programs'][idx]['hotkey'] = hotkey_var.get().strip()
                refresh()
            dialog.destroy()
        
        tk.Button(dialog, text="确定", command=ok).pack(pady=5)
    
    def on_save():
        save_config_to_file(config)
        messagebox.showinfo("提示", "已保存!")
    
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)
    
    tk.Button(btn_frame, text="添加", command=add_program, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="删除", command=delete_program, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="设置热键", command=set_hotkey, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="保存", command=on_save, width=10).pack(side=tk.LEFT, padx=5)
    
    status_label = tk.Label(root, text="运行中... 按热键切换程序", fg="gray")
    status_label.pack(pady=5)
    
    refresh()
    
    # 定时检测热键
    def check():
        check_hotkeys()
        root.after(200, check)
    
    check()
    
    log("完成")
    root.mainloop()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        log(f"错误: {e}")
        import traceback
        traceback.print_exc()
        input("按回车...")

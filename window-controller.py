import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import win32process
import psutil
import json
import os
import ctypes
from tkinter import simpledialog
import sys

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # 如果不是管理員權限，則重新以管理員權限執行
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

class WindowController:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("視窗控制程式")
        
        # 設定檔路徑
        self.config_file = "window_settings.txt"
        
        # 主視窗大小和位置
        self.window_items = []
        self.load_settings()
        
        # 建立主要框架
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 新增按鈕
        ttk.Button(self.main_frame, text="新增視窗", command=self.add_window).grid(row=0, column=0, pady=5)
        
        # 建立列表框架
        self.list_frame = ttk.Frame(self.main_frame)
        self.list_frame.grid(row=1, column=0, pady=5)
        
        # 控制按鈕框架
        control_frame = ttk.Frame(self.main_frame)
        control_frame.grid(row=2, column=0, pady=5)
        
        # 控制按鈕
        ttk.Button(control_frame, text="定位", command=lambda: self.apply_settings("position")).grid(row=0, column=0, padx=5)
        ttk.Button(control_frame, text="調整大小", command=lambda: self.apply_settings("size")).grid(row=0, column=1, padx=5)
        ttk.Button(control_frame, text="縮小", command=lambda: self.apply_settings("minimize")).grid(row=0, column=2, padx=5)
        ttk.Button(control_frame, text="顯示", command=lambda: self.apply_settings("restore")).grid(row=0, column=3, padx=5)
        ttk.Button(control_frame, text="刪除", command=self.delete_selected).grid(row=0, column=4, padx=5)
        
        # 視窗關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 更新列表顯示
        self.update_list()

    def get_window_list(self):
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(pid)
                        windows.append((title, hwnd))
                    except:
                        pass
            return True
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    def add_window(self):
        # 創建新視窗
        dialog = tk.Toplevel(self.root)
        dialog.title("選擇視窗")
        dialog.geometry("300x100")
        
        # 獲取當前視窗列表
        windows = self.get_window_list()
        window_titles = [w[0] for w in windows]
        
        # 下拉選單
        combo = ttk.Combobox(dialog, values=window_titles, state="readonly")
        combo.grid(row=0, column=0, padx=10, pady=10)
        
        def on_confirm():
            selected_index = combo.current()
            if selected_index >= 0:
                hwnd = windows[selected_index][1]
                rect = win32gui.GetWindowRect(hwnd)
                
                # 添加到列表
                self.window_items.append({
                    'name': windows[selected_index][0],
                    'hwnd': hwnd,
                    'x': rect[0],
                    'y': rect[1],
                    'width': rect[2] - rect[0],
                    'height': rect[3] - rect[1],
                    'checked': False
                })
                self.update_list()
                dialog.destroy()
        
        ttk.Button(dialog, text="確認", command=on_confirm).grid(row=1, column=0, pady=10)

    def update_list(self):
        # 清除現有列表
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        
        # 重新建立列表
        for i, item in enumerate(self.window_items):
            frame = ttk.Frame(self.list_frame)
            frame.grid(row=i, column=0, pady=2)
            
            # Checkbox
            var = tk.BooleanVar(value=item['checked'])
            cb = ttk.Checkbutton(frame, variable=var)
            cb.grid(row=0, column=0, padx=2)
            
            # 綁定checkbox狀態變更事件
            def on_checkbox_change(index=i, variable=var):
                self.window_items[index]['checked'] = variable.get()
            
            var.trace_add('write', lambda *args, index=i, variable=var: on_checkbox_change(index, variable))
            
            # 視窗名稱
            ttk.Label(frame, text=item['name']).grid(row=0, column=1, padx=2)
            
            # X座標
            ttk.Label(frame, text="X:").grid(row=0, column=2, padx=2)
            x_entry = ttk.Entry(frame, width=5)
            x_entry.insert(0, str(item['x']))
            x_entry.grid(row=0, column=3, padx=2)
            
            # Y座標
            ttk.Label(frame, text="Y:").grid(row=0, column=4, padx=2)
            y_entry = ttk.Entry(frame, width=5)
            y_entry.insert(0, str(item['y']))
            y_entry.grid(row=0, column=5, padx=2)
            
            # 寬度
            ttk.Label(frame, text="寬:").grid(row=0, column=6, padx=2)
            width_entry = ttk.Entry(frame, width=5)
            width_entry.insert(0, str(item['width']))
            width_entry.grid(row=0, column=7, padx=2)
            
            # 高度
            ttk.Label(frame, text="高:").grid(row=0, column=8, padx=2)
            height_entry = ttk.Entry(frame, width=5)
            height_entry.insert(0, str(item['height']))
            height_entry.grid(row=0, column=9, padx=2)
            
            # 綁定輸入框變更事件
            def update_value(index, field, value):
                try:
                    self.window_items[index][field] = int(value)
                except ValueError:
                    pass
            
            x_entry.bind('<FocusOut>', lambda e, i=i: update_value(i, 'x', x_entry.get()))
            y_entry.bind('<FocusOut>', lambda e, i=i: update_value(i, 'y', y_entry.get()))
            width_entry.bind('<FocusOut>', lambda e, i=i: update_value(i, 'width', width_entry.get()))
            height_entry.bind('<FocusOut>', lambda e, i=i: update_value(i, 'height', height_entry.get()))

    def apply_settings(self, action):
        for item in self.window_items:
            if item['checked']:
                try:
                    hwnd = win32gui.FindWindow(None, item['name'])
                    if hwnd:
                        if action == "position":
                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, item['x'], item['y'], 
                                                item['width'], item['height'], 
                                                win32con.SWP_NOSIZE)
                        elif action == "size":
                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, item['x'], item['y'], 
                                                item['width'], item['height'], 
                                                win32con.SWP_NOMOVE)
                        elif action == "minimize":
                            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                        elif action == "restore":
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                except:
                    messagebox.showerror("錯誤", f"無法控制視窗: {item['name']}")

    def delete_selected(self):
        self.window_items = [item for item in self.window_items if not item['checked']]
        self.update_list()

    def save_settings(self):
        settings = {
            'window_items': [{k: v for k, v in item.items() if k != 'hwnd'} 
                           for item in self.window_items],
            'window_geometry': {
                'x': self.root.winfo_x(),
                'y': self.root.winfo_y(),
                'width': self.root.winfo_width(),
                'height': self.root.winfo_height()
            }
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)

    def load_settings(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    
                self.window_items = settings.get('window_items', [])
                
                # 設置主視窗位置和大小
                geometry = settings.get('window_geometry', {})
                if geometry:
                    self.root.geometry(f"{geometry['width']}x{geometry['height']}+"
                                     f"{geometry['x']}+{geometry['y']}")
            except:
                pass

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = WindowController()
    app.run()

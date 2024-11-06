import tkinter as tk
from tkinter import ttk
import json
import win32gui
import win32con
import os
from tkinter import messagebox

class WindowManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("視窗定位管理器")
        self.root.geometry("700x400")
        
        self.window_list = []
        self.config_file = "window_config.txt"
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.create_ui()
        self.load_config()
    
    def on_closing(self):
        # 關閉前儲存設定
        self.save_config()
        self.root.destroy()
        
    def get_window_position(self, window_name):
        """獲取指定視窗的當前位置"""
        position = [0, 0]  # 預設值
        
        def callback(hwnd, context):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title == window_name:
                    rect = win32gui.GetWindowRect(hwnd)
                    position[0], position[1] = rect[0], rect[1]
                    return False  # 找到視窗後停止搜尋
            return True

        try:
            win32gui.EnumWindows(callback, None)
        except Exception:
            pass
            
        return position[0], position[1]
        
    def create_ui(self):
        # 新增按鈕
        add_button = ttk.Button(self.root, text="新增視窗", command=self.show_add_window)
        add_button.pack(pady=10)
        
        # 建立列表框架
        self.list_frame = ttk.Frame(self.root)
        self.list_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        # 設定列寬
        for i in range(5):
            self.list_frame.grid_columnconfigure(i, weight=1)
        
        # 列表標題
        headers = ["選擇", "應用程式名稱", "X軸", "Y軸", "操作"]
        for i, header in enumerate(headers):
            label = ttk.Label(self.list_frame, text=header)
            label.grid(row=0, column=i, padx=5, sticky="w" if i == 1 else "")
            
        # 執行按鈕
        execute_button = ttk.Button(self.root, text="執行", command=self.execute_positioning)
        execute_button.pack(pady=10)

    def delete_window(self, index):
        # 刪除UI元素
        for widget in self.list_frame.grid_slaves():
            if int(widget.grid_info()["row"]) == index + 1:
                widget.destroy()
        
        # 從列表中移除
        del self.window_list[index]
        
        # 重新整理UI列表
        for i, window_info in enumerate(self.window_list, start=1):
            for widget in self.list_frame.grid_slaves():
                if int(widget.grid_info()["row"]) == i + 1:
                    widget.grid(row=i + 1)

        # 儲存配置
        self.save_config()

    def add_window_to_list(self, window_name, x=0, y=0, checked=False):
        # 檢查是否已存在相同名稱的視窗
        for window_info in self.window_list:
            if window_info["name"] == window_name:
                messagebox.showwarning("警告", "此視窗已在列表中")
                return False
                
        row = len(self.window_list) + 1
        
        # 核取方塊
        var = tk.BooleanVar(value=checked)  # 使用傳入的checked狀態
        check = ttk.Checkbutton(self.list_frame, variable=var)
        check.grid(row=row, column=0)
        
        # 應用程式名稱（靠左對齊）
        name_label = ttk.Label(self.list_frame, text=window_name)
        name_label.grid(row=row, column=1, sticky="w", padx=5)
        
        # X軸輸入
        x_entry = ttk.Entry(self.list_frame, width=10)
        x_entry.grid(row=row, column=2)
        x_entry.insert(0, str(x))
        
        # Y軸輸入
        y_entry = ttk.Entry(self.list_frame, width=10)
        y_entry.grid(row=row, column=3)
        y_entry.insert(0, str(y))
        
        # 刪除按鈕
        index = len(self.window_list)
        delete_button = ttk.Button(self.list_frame, text="刪除",
                                command=lambda idx=index: self.delete_window(idx))
        delete_button.grid(row=row, column=4)
        
        # 儲存視窗資訊
        window_info = {
            "name": window_name,
            "checkbox": var,
            "x_entry": x_entry,
            "y_entry": y_entry
        }
        self.window_list.append(window_info)
        
        # 儲存配置
        self.save_config()
        return True
        
    def execute_positioning(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        for window_info in self.window_list:
            if window_info["checkbox"].get():  # 檢查是否勾選
                try:
                    x = int(window_info["x_entry"].get())
                    y = int(window_info["y_entry"].get())
                    
                    # 檢查是否超出螢幕範圍
                    if x >= screen_width or y >= screen_height:
                        messagebox.showwarning("警告", 
                            f"視窗 '{window_info['name']}' 的位置超出螢幕範圍")
                        continue
                    
                    # 找到視窗並移動位置
                    def callback(hwnd, context):
                        if win32gui.GetWindowText(hwnd) == window_info["name"]:
                            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, win32con.SWP_NOSIZE)
                    
                    # 使用EnumWindows遍歷所有視窗以尋找符合的應用程式名稱
                    win32gui.EnumWindows(callback, None)
                    
                except ValueError:
                    messagebox.showerror("錯誤", "座標必須是數字")
        
        # 儲存配置
        self.save_config()
        
    def show_add_window(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("新增視窗")
        add_window.geometry("300x150")
        add_window.transient(self.root)  # 設置為主視窗的子視窗
        add_window.grab_set()  # 模態視窗
        
        # 將視窗置中
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 計算置中的位置
        center_x = root_x + (root_width // 2) - (300 // 2)
        center_y = root_y + (root_height // 2) - (150 // 2)
        add_window.geometry(f"+{center_x}+{center_y}")
        
        # 獲取當前開啟的視窗列表
        windows = []
        def enum_windows_callback(hwnd, windows_list):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and title not in windows_list:
                    windows_list.append(title)
        
        win32gui.EnumWindows(enum_windows_callback, windows)
        
        # 移除空標題的視窗
        windows = [w for w in windows if w.strip()]
        
        # 下拉選單
        ttk.Label(add_window, text="選擇應用程式:").pack(pady=10)
        window_combobox = ttk.Combobox(add_window, values=sorted(windows), width=0)
        window_combobox.pack(pady=5, padx=10, fill=tk.X)
        
        def confirm():
            selected = window_combobox.get()
            if selected:
                # 獲取當前視窗位置
                x, y = self.get_window_position(selected)
                if self.add_window_to_list(selected, x, y):
                    add_window.destroy()
            else:
                messagebox.showwarning("警告", "請選擇一個視窗")
        
        ttk.Button(add_window, text="確認", command=confirm).pack(pady=10)
        
    def save_config(self):
        config = []
        for window_info in self.window_list:
            config.append({
                "name": window_info["name"],
                "x": window_info["x_entry"].get(),
                "y": window_info["y_entry"].get(),
                "checked": window_info["checkbox"].get()  # 儲存checkbox狀態
            })
            
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
            
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    
                for window_config in config:
                    checkbox_state = window_config.get("checked", False)  # 讀取checkbox狀態
                    self.add_window_to_list(
                        window_config["name"],
                        int(window_config["x"]), 
                        int(window_config["y"]),
                        checkbox_state  # 傳遞狀態給add_window_to_list
                    )
                    
            except json.JSONDecodeError:
                messagebox.showerror("錯誤", "配置文件格式錯誤")

                
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = WindowManager()
    app.run()
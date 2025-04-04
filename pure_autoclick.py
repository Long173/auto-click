import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import json
import os
import datetime

# Import các module tùy chọn
try:
    import win32gui
    import win32api
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    messagebox.showwarning("Missing Dependency", 
                          "win32gui module not found. Please install pywin32: pip install pywin32")

try:
    import pyautogui
    HAS_PYAUTOGUI = True
except ImportError:
    HAS_PYAUTOGUI = False
    messagebox.showwarning("Missing Dependency", 
                          "pyautogui module not found. Please install: pip install pyautogui")

class AutoClickApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Free Auto Clicker")
        self.root.geometry("600x470")  # Tăng chiều rộng từ 580 lên 600, tăng chiều cao từ 450 lên 470
        self.root.resizable(False, False)
        
        # Biến lưu trữ chung
        self.click_tabs = []  # Danh sách các tab click
        self.current_tab = None  # Tab hiện tại đang được chọn
        
        # Tạo giao diện
        self.create_ui()
        
        # Thiết lập file cấu hình tự động
        self.auto_config_file = os.path.join(os.getcwd(), "autoclick_auto_save.json")
        
        # Kiểm tra xem có file cấu hình tự động không
        if os.path.exists(self.auto_config_file):
            # Tự động tải cấu hình khi mở phần mềm
            self.load_auto_config()
        else:
            # Tạo tab mặc định đầu tiên
            self.add_click_tab("Tab 1")
        
        # Bắt sự kiện đóng cửa sổ để lưu cấu hình
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Khôi phục vị trí cửa sổ nếu có
        self.restore_window_position()
        
    def create_ui(self):
        # Tab Control chính
        self.main_tab_control = ttk.Notebook(self.root)
        
        # Tab Auto Click
        self.click_container = ttk.Frame(self.main_tab_control)
        self.main_tab_control.add(self.click_container, text="Auto Click")
        
        # Tab Cài đặt 
        self.settings_tab = ttk.Frame(self.main_tab_control)
        self.main_tab_control.add(self.settings_tab, text="Cài đặt")
        
        self.main_tab_control.pack(expand=1, fill="both", padx=5, pady=5)
        
        # Hiển thị các nút quản lý tab ngay dưới tab control chính
        tab_manager_frame = ttk.LabelFrame(self.click_container, text="Quản lý tab")
        tab_manager_frame.pack(fill="x", padx=5, pady=5)
        
        self.tab_control_frame = ttk.Frame(tab_manager_frame)
        self.tab_control_frame.pack(fill="x", padx=5, pady=5)
        
        # Sắp xếp lại các nút có kích thước vừa phải hơn
        add_tab_btn = ttk.Button(self.tab_control_frame, text="Thêm tab", command=self.add_new_tab_dialog, width=10)
        add_tab_btn.pack(side=tk.LEFT, padx=2, pady=5)
        
        rename_tab_btn = ttk.Button(self.tab_control_frame, text="Đổi tên", command=self.rename_current_tab, width=10)
        rename_tab_btn.pack(side=tk.LEFT, padx=2, pady=5)
        
        del_tab_btn = ttk.Button(self.tab_control_frame, text="Xóa tab", command=self.delete_current_tab, width=10)
        del_tab_btn.pack(side=tk.LEFT, padx=2, pady=5)
        
        # Tạo Tab Control cho các tab bên trong tab Auto Click
        tab_games_frame = ttk.LabelFrame(self.click_container, text="Danh sách tab")
        tab_games_frame.pack(expand=1, fill="both", padx=5, pady=5)
        
        self.click_tab_control = ttk.Notebook(tab_games_frame)
        self.click_tab_control.pack(expand=1, fill="both", padx=5, pady=5)
        
        # Bind khi chuyển tab
        self.click_tab_control.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Thêm chức năng kéo thả tab
        self.enable_tab_reordering(self.click_tab_control)
        
        # ===== SETTINGS TAB =====
        # Phím tắt
        hotkey_frame = ttk.LabelFrame(self.settings_tab, text="Phím tắt")
        hotkey_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(hotkey_frame, text="Home:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(hotkey_frame, text="Bắt cửa sổ hiện tại").grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(hotkey_frame, text="End:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(hotkey_frame, text="Bắt đầu/Dừng auto click").grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(hotkey_frame, text="PgUp:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(hotkey_frame, text="Thêm vị trí chuột hiện tại").grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Thông tin phần mềm
        info_frame = ttk.LabelFrame(self.settings_tab, text="Thông tin")
        info_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(info_frame, text="Free Auto Clicker - Phiên bản 1.0").pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text="Phần mềm tự động click không chiếm chuột").pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text="Phím tắt Home và Page Up hoạt động cả khi cửa sổ phần mềm không được chọn").pack(anchor=tk.W, padx=5, pady=2)
        ttk.Label(info_frame, text="Kéo và thả để thay đổi vị trí các tab").pack(anchor=tk.W, padx=5, pady=2)
        
        # Đăng ký phím tắt toàn cục
        self.register_global_hotkeys()
        
        # Gắn phím tắt trong ứng dụng
        self.root.bind("<End>", lambda e: self.toggle_clicking())
        self.root.bind("<Prior>", lambda e: self.add_current_position())  # Page Up
        
    def register_global_hotkeys(self):
        """Đăng ký phím tắt toàn cục"""
        if HAS_WIN32:
            # Tạo thread riêng để lắng nghe phím tắt toàn cục
            hotkey_thread = threading.Thread(target=self.global_hotkey_listener, daemon=True)
            hotkey_thread.start()
        
    def global_hotkey_listener(self):
        """Lắng nghe phím tắt toàn cục"""
        import keyboard
        keyboard.add_hotkey('home', self.handle_home_hotkey)
        keyboard.add_hotkey('page up', self.handle_pageup_hotkey)
        keyboard.wait()  # Chạy vô hạn
        
    def handle_home_hotkey(self):
        """Xử lý khi nhấn phím Home từ bất kỳ đâu"""
        self.root.after(0, lambda: self.select_current_window())
        
    def handle_pageup_hotkey(self):
        """Xử lý khi nhấn phím Page Up từ bất kỳ đâu"""
        self.root.after(0, lambda: self.add_current_position())
    
    def add_new_tab_dialog(self):
        """Hiển thị hộp thoại tạo tab mới"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Tạo tab mới")
        dialog.geometry("300x120")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Đặt vị trí hộp thoại ở giữa cửa sổ chính
        self.center_dialog(dialog)
        
        # Icon cho hộp thoại
        try:
            dialog.iconbitmap(self.root.iconbitmap())
        except:
            pass
        
        ttk.Label(dialog, text="Tên tab:").pack(padx=10, pady=5)
        
        name_var = tk.StringVar(value=f"Tab {len(self.click_tabs) + 1}")
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(padx=10, pady=5)
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        # Tạo frame cho nút để canh giữa
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def add_and_close():
            name = name_var.get().strip()
            if name:
                self.add_click_tab(name)
                dialog.destroy()
        
        # Tạo nút với kích thước lớn hơn và rõ ràng hơn
        create_button = tk.Button(button_frame, text="Tạo tab", command=add_and_close, 
                                width=15, bg="#4CAF50", fg="white", 
                                relief=tk.RAISED, font=("Arial", 10, "bold"))
        create_button.pack(pady=5)
        
        # Enter để xác nhận
        dialog.bind("<Return>", lambda e: add_and_close())
        
    def rename_current_tab(self):
        """Đổi tên tab hiện tại"""
        if not self.current_tab:
            return
            
        current_idx = self.click_tab_control.index(self.click_tab_control.select())
        if current_idx < 0 or current_idx >= len(self.click_tabs):
            return
            
        current_name = self.click_tab_control.tab(current_idx, "text")
        # Loại bỏ kí hiệu nếu đang chạy
        if current_name.startswith("▶ "):
            current_name = current_name[2:]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Đổi tên tab")
        dialog.geometry("300x100")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Đặt vị trí hộp thoại ở giữa cửa sổ chính
        self.center_dialog(dialog)
        
        ttk.Label(dialog, text="Tên mới:").pack(padx=10, pady=5)
        
        name_var = tk.StringVar(value=current_name)
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        name_entry.pack(padx=10, pady=5)
        name_entry.select_range(0, tk.END)
        name_entry.focus()
        
        def rename_and_close():
            new_name = name_var.get().strip()
            if new_name:
                # Giữ lại kí hiệu nếu tab đang chạy
                tab_data = self.click_tabs[current_idx]
                if tab_data["is_clicking"]:
                    new_name = f"▶ {new_name}"
                self.click_tab_control.tab(current_idx, text=new_name)
                dialog.destroy()
        
        ttk.Button(dialog, text="Đổi tên", command=rename_and_close).pack(pady=10)
        
        # Enter để xác nhận
        dialog.bind("<Return>", lambda e: rename_and_close())
    
    def center_dialog(self, dialog):
        """Đặt vị trí của hộp thoại ở trung tâm cửa sổ chính"""
        # Đợi cho cửa sổ được tạo
        dialog.update_idletasks()
        
        # Tính toán vị trí
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        main_x = self.root.winfo_rootx()
        main_y = self.root.winfo_rooty()
        
        dialog_width = dialog.winfo_width()
        dialog_height = dialog.winfo_height()
        
        # Vị trí chính giữa cửa sổ chính
        x = main_x + (main_width - dialog_width) // 2
        y = main_y + (main_height - dialog_height) // 2
        
        # Đặt vị trí của cửa sổ dialog
        dialog.geometry(f"+{x}+{y}")
        
    def delete_current_tab(self):
        """Xóa tab hiện tại"""
        if len(self.click_tabs) <= 1:
            messagebox.showwarning("Cảnh báo", "Không thể xóa tab duy nhất")
            return
            
        current_idx = self.click_tab_control.index(self.click_tab_control.select())
        if current_idx < 0 or current_idx >= len(self.click_tabs):
            return
            
        # Dừng click nếu đang chạy
        tab_data = self.click_tabs[current_idx]
        if tab_data.get("is_clicking", False):
            tab_data["is_clicking"] = False
            
        # Xóa tab
        self.click_tab_control.forget(current_idx)
        self.click_tabs.pop(current_idx)
        
        # Chọn tab khác nếu có
        if self.click_tabs:
            new_idx = min(current_idx, len(self.click_tabs) - 1)
            self.click_tab_control.select(new_idx)
            self.current_tab = self.click_tabs[new_idx]
            
    def on_tab_changed(self, event):
        """Xử lý khi chuyển tab"""
        current_idx = self.click_tab_control.index(self.click_tab_control.select())
        if 0 <= current_idx < len(self.click_tabs):
            self.current_tab = self.click_tabs[current_idx]

    def add_click_tab(self, name):
        """Thêm tab click mới"""
        # Tạo frame cho tab mới
        tab_frame = ttk.Frame(self.click_tab_control)
        self.click_tab_control.add(tab_frame, text=name)
        
        # Tạo dữ liệu cho tab
        tab_data = {
            "name": name,
            "frame": tab_frame,
            "target_windows": [],
            "selected_window": None,
            "click_positions": [],
            "is_clicking": False,
            "click_thread": None,
            "delay": 1.0,
            "repeat_count": 0,
            "max_repeats": 0,
            "require_active": False,  # Giữ lại trong dữ liệu với giá trị mặc định là False
            "widgets": {}  # Lưu các widget của tab
        }
        
        # Thêm vào danh sách tab
        self.click_tabs.append(tab_data)
        
        # Tạo giao diện cho tab
        self.create_tab_ui(tab_data)
        
        # Chọn tab mới tạo
        self.click_tab_control.select(len(self.click_tabs) - 1)
        self.current_tab = tab_data
        
        # Load cửa sổ hiện tại
        self.refresh_windows(tab_data)
        
        return tab_data

    def create_tab_ui(self, tab_data):
        """Tạo giao diện cho tab"""
        frame = tab_data["frame"]
        widgets = tab_data["widgets"]
        
        # Chia giao diện thành hai phần
        left_frame = ttk.Frame(frame)
        left_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=2, pady=2)
        
        right_frame = ttk.Frame(frame)
        right_frame.pack(side=tk.RIGHT, fill="y", padx=2, pady=2)
        
        # Frame chọn cửa sổ - bên trái
        window_frame = ttk.LabelFrame(left_frame, text="Chọn cửa sổ mục tiêu")
        window_frame.pack(fill="x", padx=3, pady=3)
        
        # Combobox chọn cửa sổ
        window_var = tk.StringVar()
        window_combo = ttk.Combobox(window_frame, textvariable=window_var, state="readonly", width=30)
        window_combo.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Nút làm mới
        refresh_btn = ttk.Button(window_frame, text="Làm mới", 
                                command=lambda: self.refresh_windows(tab_data))
        refresh_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Hiển thị cửa sổ đã chọn
        selected_window_label = ttk.Label(left_frame, text="Chưa chọn cửa sổ (Nhấn Home để bắt cửa sổ)", wraplength=300)
        selected_window_label.pack(anchor=tk.W, padx=3, pady=1)
        
        # Frame vị trí click - bên trái
        click_frame = ttk.LabelFrame(left_frame, text="Danh sách vị trí click")
        click_frame.pack(fill="both", expand=True, padx=3, pady=3)
        
        # Tạo container mới với cuộn
        listbox_container = ttk.Frame(click_frame)
        listbox_container.pack(fill="both", expand=True, padx=2, pady=2)
        
        # Listbox cho vị trí click
        click_listbox = tk.Listbox(listbox_container, width=40, height=6)  # Giảm chiều cao từ 8 xuống 6
        click_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(listbox_container, orient="vertical", command=click_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        click_listbox.config(yscrollcommand=scrollbar.set)
        
        # Bind sự kiện cuộn chuột cho Windows
        click_listbox.bind('<MouseWheel>', lambda e: click_listbox.yview_scroll(int(-1*(e.delta/120)), "units"))
        # Bind sự kiện cuộn chuột cho Linux
        click_listbox.bind('<Button-4>', lambda e: click_listbox.yview_scroll(-1, "units"))
        click_listbox.bind('<Button-5>', lambda e: click_listbox.yview_scroll(1, "units"))
        # Binding double-click để đổi tên vị trí
        click_listbox.bind('<Double-1>', lambda e: self.rename_position(tab_data))
        
        # Frame nút bấm cho vị trí click - bên trái
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill="x", padx=3, pady=2)  # Giảm padding
        
        # Bỏ nút thêm vị trí hiện tại, chỉ giữ lại thông tin phím tắt
        pos_label = ttk.Label(btn_frame, text="PgUp: Thêm vị trí")
        pos_label.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Nút đổi tên vị trí đã chọn
        rename_pos_btn = ttk.Button(btn_frame, text="Đổi tên", 
                                  command=lambda: self.rename_position(tab_data), width=8)
        rename_pos_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Nút xóa vị trí đã chọn
        del_pos_btn = ttk.Button(btn_frame, text="Xóa vị trí", 
                                command=lambda: self.delete_position(tab_data), width=8)
        del_pos_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Nút xóa tất cả
        clear_pos_btn = ttk.Button(btn_frame, text="Xóa tất cả", 
                                  command=lambda: self.clear_positions(tab_data), width=10)
        clear_pos_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Frame cấu hình - bên trái
        config_frame = ttk.Frame(left_frame)
        config_frame.pack(fill="x", padx=3, pady=2)
        
        # Nút lưu cấu hình tab hiện tại
        save_tab_btn = ttk.Button(config_frame, text="Lưu cấu hình", 
                                 command=lambda: self.save_tab_config(tab_data), width=15)
        save_tab_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Nút tải cấu hình cho tab
        load_tab_btn = ttk.Button(config_frame, text="Tải cấu hình", 
                                 command=lambda: self.load_tab_config(tab_data), width=15)
        load_tab_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # Frame điều khiển - bên phải
        control_frame = ttk.LabelFrame(right_frame, text="Điều khiển")
        control_frame.pack(fill="both", expand=True, padx=3, pady=3)
        
        # Nút bắt đầu/dừng
        start_stop_btn = ttk.Button(control_frame, text="Bắt đầu", 
                                   command=lambda: self.toggle_clicking(tab_data), width=12)
        start_stop_btn.pack(padx=5, pady=5)
        
        # Thời gian chờ
        delay_frame = ttk.Frame(control_frame)
        delay_frame.pack(fill="x", padx=5, pady=5)  # Tăng padding dọc
        ttk.Label(delay_frame, text="Thời gian chờ:").pack(side=tk.LEFT)
        delay_var = tk.StringVar(value="1.0")
        
        # Tăng kích thước của trường nhập thời gian chờ
        delay_entry = ttk.Entry(delay_frame, textvariable=delay_var, width=8)  # Tăng từ 5 lên 8
        delay_entry.pack(side=tk.RIGHT)
        
        # Số lần lặp lại
        repeat_frame = ttk.Frame(control_frame)
        repeat_frame.pack(fill="x", padx=5, pady=5)  # Tăng padding dọc
        ttk.Label(repeat_frame, text="Lặp:").pack(side=tk.LEFT)
        repeat_var = tk.StringVar(value="0")
        
        # Tăng kích thước của trường nhập số lần lặp
        repeat_entry = ttk.Entry(repeat_frame, textvariable=repeat_var, width=8)  # Tăng từ 5 lên 8
        repeat_entry.pack(side=tk.RIGHT)
        
        # Hiển thị số lần đã lặp
        count_frame = ttk.Frame(control_frame)
        count_frame.pack(fill="x", padx=5, pady=5)  # Tăng padding dọc
        ttk.Label(count_frame, text="Đã lặp:").pack(side=tk.LEFT)
        count_label = ttk.Label(count_frame, text="0")
        count_label.pack(side=tk.RIGHT)
        
        # Lưu các widget vào tab_data
        widgets.update({
            "window_var": window_var,
            "window_combo": window_combo,
            "selected_window_label": selected_window_label,
            "click_listbox": click_listbox,
            "delay_var": delay_var,
            "repeat_var": repeat_var,
            "count_label": count_label,
            "start_stop_btn": start_stop_btn
        })
        
    def refresh_windows(self, tab_data=None):
        """Làm mới danh sách cửa sổ"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        tab_data["target_windows"] = []
        window_titles = []
        
        if HAS_WIN32:
            def enum_windows_callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if title:
                        # Thêm HWND vào tiêu đề để phân biệt các cửa sổ cùng tên
                        display_title = f"{title} [HWND: {hwnd}]"
                        results.append((hwnd, display_title))
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            for hwnd, display_title in windows:
                tab_data["target_windows"].append(hwnd)
                window_titles.append(display_title)
        
        window_combo = tab_data["widgets"]["window_combo"]
        window_combo['values'] = window_titles
        if window_titles:
            window_combo.current(0)
    
    def select_current_window(self, tab_data=None):
        """Chọn cửa sổ hiện tại đang active"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        if not HAS_WIN32:
            messagebox.showwarning("Missing Dependency", "Tính năng này yêu cầu module win32gui")
            return
        
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            
            if title:
                # Thêm HWND vào tiêu đề để phân biệt các cửa sổ cùng tên
                current_time = datetime.datetime.now().strftime("%H:%M:%S")
                display_title = f"{title} [HWND: {hwnd}] [{current_time}]"
                
                # Kiểm tra xem cửa sổ đã có trong danh sách chưa
                if hwnd not in tab_data["target_windows"]:
                    tab_data["target_windows"].append(hwnd)
                    
                    # Cập nhật combobox
                    window_combo = tab_data["widgets"]["window_combo"]
                    values = list(window_combo['values'])
                    values.append(display_title)
                    window_combo['values'] = values
                    window_combo.current(len(values)-1)
                else:
                    # Chọn cửa sổ trong danh sách
                    idx = tab_data["target_windows"].index(hwnd)
                    tab_data["widgets"]["window_combo"].current(idx)
                
                # Lưu cửa sổ đã chọn
                tab_data["selected_window"] = hwnd
                tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {display_title}")
                
                # Cập nhật tiêu đề ứng dụng tạm thời thay vì hiện popup
                self.root.title(f"Free Auto Clicker - Đã bắt cửa sổ: {title}")
                self.root.after(1500, lambda: self.root.title("Free Auto Clicker"))
                
                # Không hiển thị thông báo nữa
                # messagebox.showinfo("Thành công", f"Đã chọn cửa sổ: {display_title}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể chọn cửa sổ hiện tại: {e}")
    
    def add_current_position(self, tab_data=None):
        """Thêm vị trí chuột hiện tại vào danh sách"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        if not HAS_PYAUTOGUI:
            messagebox.showwarning("Missing Dependency", "Tính năng này yêu cầu module pyautogui")
            return
            
        try:
            pos = pyautogui.position()
            
            # Hiển thị hộp thoại để người dùng nhập tên cho vị trí click
            dialog = tk.Toplevel(self.root)
            dialog.title("Đặt tên vị trí")
            dialog.geometry("300x120")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Canh giữa hộp thoại với cửa sổ chính
            self.center_dialog(dialog)
            
            # Frame chứa các thành phần
            frame = ttk.Frame(dialog, padding=10)
            frame.pack(fill="both", expand=True)
            
            # Label và Entry để nhập tên
            ttk.Label(frame, text="Nhập tên cho vị trí (x={}, y={}):".format(pos[0], pos[1])).pack(pady=(0, 5))
            
            name_var = tk.StringVar(value=f"Vị trí {len(tab_data['click_positions']) + 1}")
            name_entry = ttk.Entry(frame, textvariable=name_var, width=30)
            name_entry.pack(pady=5)
            name_entry.select_range(0, tk.END)
            name_entry.focus_set()
            
            # Frame chứa các nút
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(pady=5)
            
            def add_and_close():
                position_name = name_var.get().strip()
                if not position_name:
                    position_name = f"Vị trí {len(tab_data['click_positions']) + 1}"
                
                # Lưu vị trí với tên đã đặt
                click_pos = {
                    "name": position_name,
                    "x": pos[0],
                    "y": pos[1]
                }
                tab_data["click_positions"].append(click_pos)
                tab_data["widgets"]["click_listbox"].insert(tk.END, f"{position_name}: x={pos[0]}, y={pos[1]}")
                
                # Cuộn xuống mục mới thêm
                tab_data["widgets"]["click_listbox"].see(tk.END)
                
                # Thông báo nhỏ
                self.root.title(f"Free Auto Clicker - Đã thêm vị trí '{position_name}' ({pos[0]}, {pos[1]})")
                self.root.after(1500, lambda: self.root.title("Free Auto Clicker"))
                
                # Tự động lưu cấu hình
                self.save_auto_config()
                
                dialog.destroy()
            
            # Thêm nút OK và Cancel
            ttk.Button(btn_frame, text="OK", command=add_and_close, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Hủy", command=dialog.destroy, width=10).pack(side=tk.LEFT)
            
            # Binding phím Enter để xác nhận
            dialog.bind("<Return>", lambda e: add_and_close())
            dialog.bind("<Escape>", lambda e: dialog.destroy())
            
            # Đợi đến khi hộp thoại đóng
            dialog.wait_window()
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể thêm vị trí chuột: {e}")
    
    def delete_position(self, tab_data=None):
        """Xóa vị trí được chọn trong listbox"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        try:
            click_listbox = tab_data["widgets"]["click_listbox"]
            selected_idx = click_listbox.curselection()[0]
            position = tab_data["click_positions"][selected_idx]
            position_name = position.get("name", f"Vị trí {selected_idx+1}")
            
            # Xóa vị trí khỏi listbox và danh sách
            click_listbox.delete(selected_idx)
            tab_data["click_positions"].pop(selected_idx)
            
            # Thông báo nhỏ
            self.root.title(f"Free Auto Clicker - Đã xóa vị trí '{position_name}'")
            self.root.after(1500, lambda: self.root.title("Free Auto Clicker"))
            
            # Tự động lưu cấu hình
            self.save_auto_config()
        except (IndexError, TypeError):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một vị trí để xóa")
    
    def clear_positions(self, tab_data=None):
        """Xóa tất cả vị trí click"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        if not tab_data["click_positions"]:
            messagebox.showinfo("Thông báo", "Không có vị trí nào để xóa")
            return
            
        count = len(tab_data["click_positions"])
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa tất cả {count} vị trí?"):
            tab_data["click_positions"] = []
            tab_data["widgets"]["click_listbox"].delete(0, tk.END)
            
            # Thông báo nhỏ
            self.root.title(f"Free Auto Clicker - Đã xóa tất cả {count} vị trí")
            self.root.after(1500, lambda: self.root.title("Free Auto Clicker"))
            
            # Tự động lưu cấu hình
            self.save_auto_config()
    
    def toggle_clicking(self, tab_data=None):
        """Bắt đầu hoặc dừng quá trình tự động click"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        if tab_data["is_clicking"]:
            self.stop_clicking(tab_data)
        else:
            self.start_clicking(tab_data)
    
    def start_clicking(self, tab_data=None):
        """Bắt đầu quá trình tự động click"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        if not tab_data["click_positions"]:
            messagebox.showwarning("Cảnh báo", "Vui lòng thêm ít nhất một vị trí click")
            return
        
        if not tab_data["selected_window"] and HAS_WIN32:
            try:
                window_combo = tab_data["widgets"]["window_combo"]
                selected_idx = window_combo.current()
                if selected_idx >= 0:
                    tab_data["selected_window"] = tab_data["target_windows"][selected_idx]
                    title = window_combo.get()
                    tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {title}")
            except Exception:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một cửa sổ mục tiêu")
                return
        
        try:
            # Lấy các thông số và đảm bảo chúng được chuyển đổi đúng định dạng
            try:
                tab_data["delay"] = float(tab_data["widgets"]["delay_var"].get())
                if tab_data["delay"] <= 0:
                    tab_data["delay"] = 0.1  # Đặt giá trị mặc định nếu không hợp lệ
            except ValueError:
                tab_data["delay"] = 1.0  # Đặt giá trị mặc định nếu không thể chuyển đổi
                tab_data["widgets"]["delay_var"].set("1.0")
                
            try:
                tab_data["max_repeats"] = int(tab_data["widgets"]["repeat_var"].get())
                if tab_data["max_repeats"] < 0:
                    tab_data["max_repeats"] = 0  # Đặt giá trị mặc định nếu không hợp lệ
            except ValueError:
                tab_data["max_repeats"] = 0  # Đặt giá trị mặc định nếu không thể chuyển đổi
                tab_data["widgets"]["repeat_var"].set("0")
                
            tab_data["repeat_count"] = 0
            tab_data["widgets"]["count_label"].config(text="0")
            
            # Đổi trạng thái
            tab_data["is_clicking"] = True
            tab_data["widgets"]["start_stop_btn"].config(text="Dừng")
            
            # Thêm kí hiệu vào tiêu đề tab
            tab_idx = self.click_tabs.index(tab_data)
            current_tab_name = self.click_tab_control.tab(tab_idx, "text")
            if not current_tab_name.startswith("▶ "):
                self.click_tab_control.tab(tab_idx, text=f"▶ {current_tab_name}")
            
            # Bắt đầu thread click
            tab_data["click_thread"] = threading.Thread(
                target=lambda: self.click_loop(tab_data), 
                daemon=True
            )
            tab_data["click_thread"].start()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể bắt đầu auto click: {e}")
            tab_data["is_clicking"] = False
            tab_data["widgets"]["start_stop_btn"].config(text="Bắt đầu")
    
    def stop_clicking(self, tab_data=None):
        """Dừng quá trình tự động click"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        tab_data["is_clicking"] = False
        tab_data["widgets"]["start_stop_btn"].config(text="Bắt đầu")
        
        # Xóa kí hiệu ở tiêu đề tab
        tab_idx = self.click_tabs.index(tab_data)
        current_tab_name = self.click_tab_control.tab(tab_idx, "text")
        if current_tab_name.startswith("▶ "):
            self.click_tab_control.tab(tab_idx, text=current_tab_name[2:])
    
    def click_loop(self, tab_data):
        """Vòng lặp thực hiện click"""
        while tab_data["is_clicking"]:
            try:
                if HAS_WIN32 and tab_data["selected_window"]:
                    # Kiểm tra xem cửa sổ còn tồn tại không
                    if not win32gui.IsWindow(tab_data["selected_window"]):
                        self.root.after(0, lambda: messagebox.showwarning("Cảnh báo", "Cửa sổ đã chọn không còn tồn tại"))
                        self.root.after(0, lambda: self.stop_clicking(tab_data))
                        break
                
                # Thực hiện click tại từng vị trí
                for pos in tab_data["click_positions"]:
                    if not tab_data["is_clicking"]:
                        break
                    
                    # Click ảo (không di chuyển chuột thật)
                    if HAS_WIN32 and tab_data["selected_window"]:
                        self.virtual_click(tab_data, pos["x"], pos["y"])
                    # Nếu không có win32gui, dùng pyautogui (sẽ di chuyển chuột)
                    elif HAS_PYAUTOGUI:
                        current_pos = pyautogui.position()  # Lưu vị trí chuột hiện tại
                        pyautogui.click(pos["x"], pos["y"])
                        pyautogui.moveTo(current_pos[0], current_pos[1])  # Khôi phục vị trí chuột
                    
                    # Đợi giữa các click
                    time.sleep(tab_data["delay"])
                
                # Tăng số lượt lặp
                tab_data["repeat_count"] += 1
                count_label = tab_data["widgets"]["count_label"]
                self.root.after(0, lambda: count_label.config(text=str(tab_data["repeat_count"])))
                
                # Kiểm tra điều kiện dừng
                if tab_data["max_repeats"] > 0 and tab_data["repeat_count"] >= tab_data["max_repeats"]:
                    self.root.after(0, lambda tab=tab_data: self.stop_clicking(tab))
                    break
            except Exception as e:
                print(f"Error in click loop: {e}")
                self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi trong quá trình auto click: {e}"))
                self.root.after(0, lambda tab=tab_data: self.stop_clicking(tab))
                break
    
    def virtual_click(self, tab_data, x, y):
        """Thực hiện click ảo tại tọa độ x, y mà không di chuyển chuột thật"""
        if not HAS_WIN32 or not tab_data["selected_window"]:
            return False
        
        try:
            # Kiểm tra xem cửa sổ còn tồn tại không
            if not win32gui.IsWindow(tab_data["selected_window"]):
                self.root.after(0, lambda: messagebox.showwarning("Cảnh báo", "Cửa sổ đã chọn không còn tồn tại"))
                return False
                
            # Chuyển tọa độ màn hình sang tọa độ của cửa sổ mục tiêu
            client_pos = win32gui.ScreenToClient(tab_data["selected_window"], (x, y))
            
            # Gửi thông điệp click chuột đến cửa sổ bằng PostMessage thay vì SendMessage
            # PostMessage không yêu cầu cửa sổ phải active và không block thread
            lParam = win32api.MAKELONG(client_pos[0], client_pos[1])
            
            # Gửi tin nhắn chuột xuống và lên (click)
            win32api.PostMessage(tab_data["selected_window"], win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
            time.sleep(0.02)  # Đợi một chút
            win32api.PostMessage(tab_data["selected_window"], win32con.WM_LBUTTONUP, 0, lParam)
            
            return True
        except Exception as e:
            error_msg = f"Lỗi virtual click: {e}"
            print(error_msg)
            # Hiển thị thông báo lỗi khi không thể click
            if "SetForegroundWindow" in str(e):
                self.root.after(0, lambda: messagebox.showwarning("Lưu ý", 
                    "Không thể đưa cửa sổ lên phía trước, nhưng vẫn tiếp tục thực hiện click."))
            return False
    
    def save_config(self):
        """Lưu cấu hình hiện tại"""
        # Kiểm tra tất cả các tab xem có vị trí click nào không
        has_clicks = False
        for tab in self.click_tabs:
            if tab["click_positions"]:
                has_clicks = True
                break
                
        if not has_clicks:
            messagebox.showwarning("Cảnh báo", "Không có vị trí click nào để lưu")
            return
        
        try:
            # Lưu cấu hình cho tất cả các tab
            all_tabs_config = []
            
            for tab in self.click_tabs:
                # Lấy tiêu đề từ widget notebook
                tab_idx = self.click_tabs.index(tab)
                tab_name = self.click_tab_control.tab(tab_idx, "text")
                
                # Tạo cấu hình cho tab này
                tab_config = {
                    "name": tab_name,
                    "window_title": tab["widgets"]["window_combo"].get() if tab["selected_window"] else "",
                    "click_positions": [(pos["x"], pos["y"]) for pos in tab["click_positions"]],
                    "delay": float(tab["widgets"]["delay_var"].get()),
                    "repeats": int(tab["widgets"]["repeat_var"].get()),
                    "require_active": tab["require_active"]
                }
                
                all_tabs_config.append(tab_config)
            
            config = {
                "version": "1.0",
                "tabs": all_tabs_config
            }
            
            # Cho phép người dùng chọn vị trí và đặt tên file
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")],
                initialdir=os.getcwd(),
                initialfile="autoclick_config.json",
                title="Lưu cấu hình"
            )
            
            if not file_path:  # Người dùng đã hủy
                return
                
            with open(file_path, "w") as f:
                json.dump(config, f)
            
            messagebox.showinfo("Thành công", f"Đã lưu cấu hình thành công vào file:\n{os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu cấu hình: {e}")
    
    def load_config(self):
        """Tải cấu hình đã lưu"""
        try:
            # Cho phép người dùng chọn file cấu hình
            file_path = filedialog.askopenfilename(
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")],
                initialdir=os.getcwd(),
                title="Chọn file cấu hình"
            )
            
            if not file_path:  # Người dùng đã hủy
                return
                
            if not os.path.exists(file_path):
                messagebox.showinfo("Thông báo", "Không tìm thấy file cấu hình")
                return
            
            with open(file_path, "r") as f:
                config = json.load(f)
            
            # Kiểm tra phiên bản (để tương thích ngược với định dạng cũ)
            if "version" in config and "tabs" in config:
                # Định dạng mới với nhiều tab
                tabs_config = config["tabs"]
                
                # Xóa tất cả các tab hiện tại
                for i in range(len(self.click_tabs)):
                    # Dừng click nếu đang chạy
                    if self.click_tabs[0]["is_clicking"]:
                        self.click_tabs[0]["is_clicking"] = False
                    self.click_tab_control.forget(0)
                
                self.click_tabs = []
                
                # Tạo tab mới từ cấu hình
                for tab_config in tabs_config:
                    tab_name = tab_config.get("name", f"Tab {len(self.click_tabs) + 1}")
                    tab_data = self.add_click_tab(tab_name)
                    
                    # Tải vị trí click
                    tab_data["click_positions"] = []
                    click_listbox = tab_data["widgets"]["click_listbox"]
                    click_listbox.delete(0, tk.END)
                    
                    for pos in tab_config.get("click_positions", []):
                        # Đảm bảo pos là tuple và có hai giá trị
                        if isinstance(pos, (list, tuple)) and len(pos) >= 2:
                            x, y = int(pos[0]), int(pos[1])
                            tab_data["click_positions"].append({"x": x, "y": y})
                            click_listbox.insert(tk.END, f"x={x}, y={y}")
                    
                    # Tải các cài đặt khác
                    tab_data["widgets"]["delay_var"].set(str(tab_config.get("delay", 1.0)))
                    tab_data["widgets"]["repeat_var"].set(str(tab_config.get("repeats", 0)))
                    tab_data["require_active"] = tab_config.get("require_active", False)
                    
                    # Tìm cửa sổ theo tiêu đề nếu có
                    window_title = tab_config.get("window_title", "")
                    if window_title and HAS_WIN32:
                        self.refresh_windows(tab_data)
                        window_combo = tab_data["widgets"]["window_combo"]
                        for i, title in enumerate(window_combo['values']):
                            if title == window_title:
                                window_combo.current(i)
                                tab_data["selected_window"] = tab_data["target_windows"][i]
                                tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {title}")
                                break
            else:
                # Định dạng cũ chỉ có một tab
                # Xóa tất cả các tab hiện tại
                for i in range(len(self.click_tabs)):
                    # Dừng click nếu đang chạy
                    if self.click_tabs[0]["is_clicking"]:
                        self.click_tabs[0]["is_clicking"] = False
                    self.click_tab_control.forget(0)
                
                self.click_tabs = []
                
                # Tạo tab mới
                tab_data = self.add_click_tab("Tab 1")
                
                # Tải vị trí click
                tab_data["click_positions"] = []
                click_listbox = tab_data["widgets"]["click_listbox"]
                click_listbox.delete(0, tk.END)
                
                for pos in config.get("click_positions", []):
                    # Đảm bảo pos là tuple và có hai giá trị
                    if isinstance(pos, (list, tuple)) and len(pos) >= 2:
                        x, y = int(pos[0]), int(pos[1])
                        tab_data["click_positions"].append({"x": x, "y": y})
                        click_listbox.insert(tk.END, f"x={x}, y={y}")
                
                # Tải các cài đặt khác
                tab_data["widgets"]["delay_var"].set(str(config.get("delay", 1.0)))
                tab_data["widgets"]["repeat_var"].set(str(config.get("repeats", 0)))
                tab_data["require_active"] = config.get("require_active", False)
                
                # Tìm cửa sổ theo tiêu đề nếu có
                window_title = config.get("window_title", "")
                if window_title and HAS_WIN32:
                    self.refresh_windows(tab_data)
                    window_combo = tab_data["widgets"]["window_combo"]
                    for i, title in enumerate(window_combo['values']):
                        if title == window_title:
                            window_combo.current(i)
                            tab_data["selected_window"] = tab_data["target_windows"][i]
                            tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {title}")
                            break
            
            # Chọn tab đầu tiên
            if self.click_tabs:
                self.click_tab_control.select(0)
                self.current_tab = self.click_tabs[0]
            
            messagebox.showinfo("Thành công", f"Đã tải cấu hình thành công từ file:\n{os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải cấu hình: {e}")

    def load_auto_config(self):
        """Tự động tải cấu hình từ file lưu tự động"""
        try:
            if os.path.exists(self.auto_config_file):
                with open(self.auto_config_file, "r") as f:
                    config = json.load(f)
                
                # Cấu hình mới với tabs và phiên bản
                if "version" in config and "tabs" in config:
                    tabs_config = config["tabs"]
                    
                    # Xóa tất cả các tab hiện tại
                    for i in range(len(self.click_tabs)):
                        # Dừng click nếu đang chạy
                        if self.click_tabs[0]["is_clicking"]:
                            self.click_tabs[0]["is_clicking"] = False
                        self.click_tab_control.forget(0)
                    
                    self.click_tabs = []
                    
                    # Tạo tab mới từ cấu hình
                    for tab_config in tabs_config:
                        tab_name = tab_config.get("name", f"Tab {len(self.click_tabs) + 1}")
                        # Xóa kí hiệu "▶ " nếu có trong tên tab (vì tab đang không chạy khi khởi động)
                        if tab_name.startswith("▶ "):
                            tab_name = tab_name[2:]
                            
                        tab_data = self.add_click_tab(tab_name)
                        
                        # Tải vị trí click
                        tab_data["click_positions"] = []
                        click_listbox = tab_data["widgets"]["click_listbox"]
                        click_listbox.delete(0, tk.END)
                        
                        for pos in tab_config.get("click_positions", []):
                            if isinstance(pos, dict) and "x" in pos and "y" in pos:
                                # Định dạng mới với tên vị trí
                                click_pos = {
                                    "name": pos.get("name", f"Vị trí {len(tab_data['click_positions']) + 1}"),
                                    "x": int(pos["x"]),
                                    "y": int(pos["y"])
                                }
                                tab_data["click_positions"].append(click_pos)
                                click_listbox.insert(tk.END, f"{click_pos['name']}: x={click_pos['x']}, y={click_pos['y']}")
                            elif isinstance(pos, (list, tuple)) and len(pos) >= 2:
                                # Định dạng cũ chỉ có tọa độ x, y
                                x, y = int(pos[0]), int(pos[1])
                                click_pos = {
                                    "name": f"Vị trí {len(tab_data['click_positions']) + 1}",
                                    "x": x,
                                    "y": y
                                }
                                tab_data["click_positions"].append(click_pos)
                                click_listbox.insert(tk.END, f"{click_pos['name']}: x={x}, y={y}")
                        
                        # Tải các cài đặt khác
                        tab_data["widgets"]["delay_var"].set(str(tab_config.get("delay", 1.0)))
                        tab_data["widgets"]["repeat_var"].set(str(tab_config.get("repeats", 0)))
                        tab_data["require_active"] = tab_config.get("require_active", False)
                        
                        # Tìm cửa sổ theo tiêu đề nếu có
                        window_title = tab_config.get("window_title", "")
                        if window_title and HAS_WIN32:
                            self.refresh_windows(tab_data)
                            window_combo = tab_data["widgets"]["window_combo"]
                            for i, title in enumerate(window_combo['values']):
                                if title == window_title:
                                    window_combo.current(i)
                                    tab_data["selected_window"] = tab_data["target_windows"][i]
                                    tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {title}")
                                    break
                                    
                # Khôi phục vị trí cửa sổ
                self.restore_window_position()
                
        except Exception as e:
            print(f"Không thể tải cấu hình tự động: {e}")
    
    def save_auto_config(self):
        """Tự động lưu cấu hình khi đóng phần mềm"""
        try:
            # Lưu cấu hình cho tất cả các tab
            all_tabs_config = []
            
            for tab in self.click_tabs:
                # Lấy tiêu đề từ widget notebook
                tab_idx = self.click_tabs.index(tab)
                tab_name = self.click_tab_control.tab(tab_idx, "text")
                
                # Tạo cấu hình cho tab này
                tab_config = {
                    "name": tab_name,
                    "window_title": tab["widgets"]["window_combo"].get() if tab["selected_window"] else "",
                    "click_positions": [{"name": pos.get("name", f"Vị trí {i+1}"), "x": pos["x"], "y": pos["y"]} 
                                       for i, pos in enumerate(tab["click_positions"])],
                    "delay": float(tab["widgets"]["delay_var"].get()),
                    "repeats": int(tab["widgets"]["repeat_var"].get()),
                    "require_active": tab["require_active"]
                }
                
                all_tabs_config.append(tab_config)
            
            # Lưu vị trí cửa sổ
            window_x = self.root.winfo_x()
            window_y = self.root.winfo_y()
            window_width = self.root.winfo_width()
            window_height = self.root.winfo_height()
            
            config = {
                "version": "1.0",
                "tabs": all_tabs_config,
                "window_position": {
                    "x": window_x,
                    "y": window_y,
                    "width": window_width,
                    "height": window_height
                }
            }
            
            with open(self.auto_config_file, "w") as f:
                json.dump(config, f)
                
        except Exception as e:
            print(f"Không thể lưu cấu hình tự động: {e}")
    
    def restore_window_position(self):
        """Khôi phục vị trí cửa sổ từ cấu hình đã lưu"""
        try:
            if os.path.exists(self.auto_config_file):
                with open(self.auto_config_file, "r") as f:
                    config = json.load(f)
                
                if "window_position" in config:
                    pos = config["window_position"]
                    # Chờ cửa sổ được hiển thị hoàn chỉnh
                    self.root.update_idletasks()
                    
                    # Đặt kích thước (nếu cần)
                    if "width" in pos and "height" in pos:
                        self.root.geometry(f"{pos['width']}x{pos['height']}")
                    
                    # Đặt vị trí cửa sổ
                    if "x" in pos and "y" in pos:
                        # Kiểm tra vị trí có nằm trong màn hình không
                        screen_width = self.root.winfo_screenwidth()
                        screen_height = self.root.winfo_screenheight()
                        
                        x = pos["x"]
                        y = pos["y"]
                        
                        # Đảm bảo cửa sổ nằm trong màn hình
                        if x < 0:
                            x = 0
                        if y < 0:
                            y = 0
                        if x > screen_width - 100:
                            x = screen_width - 600  # Thay đổi từ 580 thành 600
                        if y > screen_height - 100:
                            y = screen_height - 470  # Thay đổi từ 450 thành 470
                        
                        self.root.geometry(f"+{x}+{y}")
        except Exception as e:
            print(f"Không thể khôi phục vị trí cửa sổ: {e}")
    
    def on_closing(self):
        """Xử lý khi đóng phần mềm"""
        # Dừng tất cả các tab đang chạy
        for tab in self.click_tabs:
            if tab.get("is_clicking", False):
                tab["is_clicking"] = False
        
        # Lưu cấu hình tự động
        self.save_auto_config()
        
        # Đóng phần mềm
        self.root.destroy()

    def save_tab_config(self, tab_data):
        """Lưu cấu hình cho tab cụ thể"""
        if not tab_data["click_positions"]:
            messagebox.showwarning("Cảnh báo", "Không có vị trí click nào để lưu")
            return
            
        try:
            # Lấy tiêu đề tab
            tab_idx = self.click_tabs.index(tab_data)
            tab_name = self.click_tab_control.tab(tab_idx, "text")
            
            # Tạo cấu hình cho tab này
            tab_config = {
                "name": tab_name,
                "window_title": tab_data["widgets"]["window_combo"].get() if tab_data["selected_window"] else "",
                "click_positions": [(pos["x"], pos["y"]) for pos in tab_data["click_positions"]],
                "delay": float(tab_data["widgets"]["delay_var"].get()),
                "repeats": int(tab_data["widgets"]["repeat_var"].get()),
                "require_active": tab_data["require_active"]
            }
            
            # Đóng gói trong định dạng tương thích
            config = {
                "version": "1.0",
                "tabs": [tab_config]
            }
            
            # Cho phép người dùng chọn vị trí và đặt tên file
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")],
                initialdir=os.getcwd(),
                initialfile=f"{tab_name}_config.json",
                title=f"Lưu cấu hình cho '{tab_name}'"
            )
            
            if not file_path:  # Người dùng đã hủy
                return
                
            with open(file_path, "w") as f:
                json.dump(config, f)
            
            messagebox.showinfo("Thành công", f"Đã lưu cấu hình cho '{tab_name}' thành công")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu cấu hình tab: {e}")
    
    def load_tab_config(self, tab_data):
        """Tải cấu hình cho tab cụ thể"""
        try:
            # Lấy tiêu đề tab
            tab_idx = self.click_tabs.index(tab_data)
            tab_name = self.click_tab_control.tab(tab_idx, "text")
            
            # Cho phép người dùng chọn file cấu hình
            file_path = filedialog.askopenfilename(
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")],
                initialdir=os.getcwd(),
                title=f"Tải cấu hình cho '{tab_name}'"
            )
            
            if not file_path:  # Người dùng đã hủy
                return
                
            if not os.path.exists(file_path):
                messagebox.showinfo("Thông báo", "Không tìm thấy file cấu hình")
                return
            
            with open(file_path, "r") as f:
                config = json.load(f)
            
            # Lấy cấu hình tab
            tab_config = None
            if "version" in config and "tabs" in config:
                # Lấy tab đầu tiên trong file cấu hình, nếu có nhiều
                if config["tabs"]:
                    tab_config = config["tabs"][0]
            else:
                # Định dạng cũ, lấy trực tiếp
                tab_config = config
            
            if not tab_config:
                messagebox.showwarning("Cảnh báo", "File cấu hình không chứa dữ liệu hợp lệ")
                return
                
            # Tải vị trí click
            tab_data["click_positions"] = []
            click_listbox = tab_data["widgets"]["click_listbox"]
            click_listbox.delete(0, tk.END)
            
            for pos in tab_config.get("click_positions", []):
                tab_data["click_positions"].append({"x": pos["x"], "y": pos["y"]})
                click_listbox.insert(tk.END, f"x={pos['x']}, y={pos['y']}")
            
            # Tải các cài đặt khác
            tab_data["widgets"]["delay_var"].set(str(tab_config.get("delay", 1.0)))
            tab_data["widgets"]["repeat_var"].set(str(tab_config.get("repeats", 0)))
            tab_data["require_active"] = tab_config.get("require_active", False)
            
            # Tìm cửa sổ theo tiêu đề nếu có
            window_title = tab_config.get("window_title", "")
            if window_title and HAS_WIN32:
                self.refresh_windows(tab_data)
                window_combo = tab_data["widgets"]["window_combo"]
                for i, title in enumerate(window_combo['values']):
                    if title == window_title:
                        window_combo.current(i)
                        tab_data["selected_window"] = tab_data["target_windows"][i]
                        tab_data["widgets"]["selected_window_label"].config(text=f"Đã chọn: {title}")
                        break
            
            # Đổi tên tab nếu có tên trong file cấu hình và người dùng muốn
            if "name" in tab_config and tab_config["name"] != tab_name:
                if messagebox.askyesno("Đổi tên tab", f"Đổi tên tab từ '{tab_name}' thành '{tab_config['name']}'?"):
                    self.click_tab_control.tab(tab_idx, text=tab_config["name"])
            
            messagebox.showinfo("Thành công", f"Đã tải cấu hình cho tab thành công")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải cấu hình tab: {e}")

    def enable_tab_reordering(self, notebook):
        """Thêm chức năng kéo thả để sắp xếp lại các tab"""
        
        def reorder_tab(event):
            try:
                # Kiểm tra nếu đang click vào khu vực không phải tab
                clicked_tab = notebook.tk.call(notebook._w, "identify", "tab", event.x, event.y)
                
                # Nếu không có tab nào được nhấn, thoát
                if clicked_tab == "":
                    return
                
                # Chuyển đổi từ chuỗi thành số
                clicked_tab = int(clicked_tab)
                
                # Kiểm tra xem index có hợp lệ không
                if clicked_tab < 0 or clicked_tab >= len(self.click_tabs):
                    return
                
                # Lưu tên hiện tại của tab đang được kéo
                tab_name = notebook.tab(clicked_tab, "text")
                
                # Lưu lại dữ liệu tab
                source_tab_data = self.click_tabs[clicked_tab]
                
                # Lưu trạng thái đánh dấu đang chạy
                is_running = tab_name.startswith("▶ ")
                
                # Tạo biến lưu trữ vị trí cuối cùng của chuột khi di chuyển
                self._drag_tab_index = clicked_tab
                self._last_mouse_x = event.x
                self._last_mouse_y = event.y
                
                # Ghi nhớ tab đang được kéo để quản lý tốt hơn
                self._dragging = True
            except Exception as e:
                print(f"Error starting tab drag: {e}")
                # Đảm bảo reset trạng thái nếu có lỗi
                if hasattr(self, '_drag_tab_index'):
                    del self._drag_tab_index
                if hasattr(self, '_dragging'):
                    self._dragging = False
        
        def drag_tab(event):
            try:
                # Kiểm tra xem có tab đang được kéo không
                if not hasattr(self, '_drag_tab_index') or not self._dragging:
                    return
                
                # Xác định tab đích
                target_tab = notebook.tk.call(notebook._w, "identify", "tab", event.x, event.y)
                
                # Nếu không có tab đích, thoát
                if target_tab == "":
                    return
                
                # Chuyển đổi từ chuỗi thành số
                target_tab = int(target_tab)
                
                # Kiểm tra xem index có hợp lệ không
                if target_tab < 0 or target_tab >= len(self.click_tabs):
                    return
                
                # Nếu tab đích giống tab nguồn, thoát
                if target_tab == self._drag_tab_index:
                    return
                
                # Chuyển tab từ vị trí cũ sang vị trí mới
                self.move_tab(self._drag_tab_index, target_tab)
                
                # Cập nhật vị trí
                self._drag_tab_index = target_tab
                
                # Lưu lại vị trí chuột
                self._last_mouse_x = event.x
                self._last_mouse_y = event.y
            except Exception as e:
                print(f"Error dragging tab: {e}")
        
        def release_tab(event):
            try:
                # Đảm bảo reset trạng thái kéo
                self._dragging = False
                
                # Xóa biến lưu trữ
                if hasattr(self, '_drag_tab_index'):
                    del self._drag_tab_index
                if hasattr(self, '_last_mouse_x'):
                    del self._last_mouse_x
                if hasattr(self, '_last_mouse_y'):
                    del self._last_mouse_y
                    
                # Tự động lưu cấu hình sau khi thay đổi vị trí tab
                self.save_auto_config()
            except Exception as e:
                print(f"Error releasing tab: {e}")
        
        # Gắn các sự kiện kéo thả cho notebook
        notebook.bind("<ButtonPress-1>", reorder_tab)
        notebook.bind("<B1-Motion>", drag_tab)
        notebook.bind("<ButtonRelease-1>", release_tab)
    
    def move_tab(self, from_index, to_index):
        """Di chuyển tab từ vị trí này sang vị trí khác"""
        try:
            # Kiểm tra tính hợp lệ của vị trí
            if (from_index == to_index or 
                from_index < 0 or to_index < 0 or 
                from_index >= len(self.click_tabs) or 
                to_index >= len(self.click_tabs)):
                return
            
            # Lấy dữ liệu của tab nguồn một cách an toàn
            source_tab_data = self.click_tabs[from_index]
            source_frame = source_tab_data["frame"]
            source_tab_text = self.click_tab_control.tab(from_index, "text")
            
            # Bảo vệ để tránh xóa tab cuối cùng
            if len(self.click_tabs) <= 1:
                return
                
            # Sử dụng try-except để xử lý lỗi khi xóa hoặc chèn tab
            try:
                # Xóa tab nguồn
                self.click_tab_control.forget(from_index)
                self.click_tabs.pop(from_index)
                
                # Điều chỉnh chỉ số đích nếu cần
                if from_index < to_index:
                    to_index -= 1
                
                # Chèn tab vào vị trí mới
                self.click_tab_control.insert(to_index, source_frame, text=source_tab_text)
                self.click_tabs.insert(to_index, source_tab_data)
                
                # Chọn tab vừa di chuyển
                self.click_tab_control.select(to_index)
                self.current_tab = source_tab_data
            except Exception as e:
                # Nếu có lỗi khi di chuyển, phục hồi tab đã xóa
                if from_index not in range(len(self.click_tabs) + 1):
                    # Nếu không thể khôi phục vị trí cũ, thêm vào cuối
                    self.click_tab_control.add(source_frame, text=source_tab_text)
                    self.click_tabs.append(source_tab_data)
                else:
                    # Khôi phục lại vị trí cũ
                    self.click_tab_control.insert(from_index, source_frame, text=source_tab_text)
                    self.click_tabs.insert(from_index, source_tab_data)
                
                # Log lỗi
                print(f"Error moving tab, restored at position: {e}")
                
                # Đảm bảo chọn một tab hợp lệ
                if self.click_tabs:
                    self.click_tab_control.select(0)
                    self.current_tab = self.click_tabs[0]
        except Exception as e:
            print(f"Critical error moving tab: {e}")
            # Làm mới toàn bộ giao diện để đảm bảo không mất dữ liệu
            if len(self.click_tabs) == 0:
                # Nếu đã mất tất cả tab, tạo tab mặc định
                self.add_click_tab("Tab 1")

    def rename_position(self, tab_data=None):
        """Đổi tên vị trí click đã chọn"""
        if tab_data is None:
            tab_data = self.current_tab
            
        if tab_data is None:
            return
            
        try:
            click_listbox = tab_data["widgets"]["click_listbox"]
            selected_idx = click_listbox.curselection()[0]
            position = tab_data["click_positions"][selected_idx]
            old_name = position.get("name", f"Vị trí {selected_idx+1}")
            
            # Hiển thị hộp thoại để người dùng nhập tên mới
            dialog = tk.Toplevel(self.root)
            dialog.title("Đổi tên vị trí")
            dialog.geometry("300x120")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Canh giữa hộp thoại với cửa sổ chính
            self.center_dialog(dialog)
            
            # Frame chứa các thành phần
            frame = ttk.Frame(dialog, padding=10)
            frame.pack(fill="both", expand=True)
            
            # Label và Entry để nhập tên
            ttk.Label(frame, text=f"Đổi tên cho vị trí (x={position['x']}, y={position['y']}):").pack(pady=(0, 5))
            
            name_var = tk.StringVar(value=old_name)
            name_entry = ttk.Entry(frame, textvariable=name_var, width=30)
            name_entry.pack(pady=5)
            name_entry.select_range(0, tk.END)
            name_entry.focus_set()
            
            # Frame chứa các nút
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(pady=5)
            
            def rename_and_close():
                new_name = name_var.get().strip()
                if not new_name:
                    new_name = old_name
                
                # Cập nhật tên mới cho vị trí
                position["name"] = new_name
                
                # Cập nhật hiển thị trong listbox
                click_listbox.delete(selected_idx)
                click_listbox.insert(selected_idx, f"{new_name}: x={position['x']}, y={position['y']}")
                click_listbox.selection_set(selected_idx)
                
                # Thông báo nhỏ
                self.root.title(f"Free Auto Clicker - Đã đổi tên vị trí thành '{new_name}'")
                self.root.after(1500, lambda: self.root.title("Free Auto Clicker"))
                
                # Tự động lưu cấu hình
                self.save_auto_config()
                
                dialog.destroy()
            
            # Thêm nút OK và Cancel
            ttk.Button(btn_frame, text="OK", command=rename_and_close, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="Hủy", command=dialog.destroy, width=10).pack(side=tk.LEFT)
            
            # Binding phím Enter để xác nhận
            dialog.bind("<Return>", lambda e: rename_and_close())
            dialog.bind("<Escape>", lambda e: dialog.destroy())
            
            # Đợi đến khi hộp thoại đóng
            dialog.wait_window()
            
        except (IndexError, TypeError):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một vị trí để đổi tên")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoClickApp(root)
    root.mainloop() 
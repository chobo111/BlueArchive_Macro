import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import keyboard
import threading
import ctypes
from ctypes import wintypes
import json
import os
import time
import webbrowser
import atexit
import sys
import os

def resource_path(relative_path):
    try:
        # PyInstaller에 의해 임시폴더에 설치된 경로
        base_path = sys._MEIPASS
    except Exception:
        # 일반 파이썬으로 실행 시 경로
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False

def create_tray_icon():
    if os.path.exists("icon.ico"):
        return Image.open("icon.ico")
    else:
        # 파일이 없을 때를 대비한 기본 사각형 유지
        image = Image.new('RGB', (64, 64), color=(43, 43, 43))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill="#28a745")
        return image

COLOR_SUCCESS = ("#28a745", "#2E8B57")  
COLOR_DANGER = ("#dc3545", "#FF6B6B")   
COLOR_DARK = ("#e2e2e2", "#2B2B2B")     
COLOR_TEXT = ("black", "white")         
COLOR_SETTING = ("#d1d1d1", "#555555")  

PUL = ctypes.POINTER(ctypes.c_ulong)
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort), ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]
class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong), ("wParamL", ctypes.c_short), ("wParamH", ctypes.c_ushort)]
class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long), ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]
class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput), ("mi", MouseInput), ("hi", HardwareInput)]
class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", Input_I)]

def click_mouse_sendinput(duration_ms):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.mi = MouseInput(0, 0, 0, 0x0002, 0, ctypes.pointer(extra)) 
    x = Input(ctypes.c_ulong(0), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
    time.sleep(duration_ms / 1000.0)
    ii_.mi = MouseInput(0, 0, 0, 0x0004, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(0), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

# 1. 매크로 컴포넌트 클래스
class MacroRow:
    def __init__(self, master_frame, index, data, app):
        self.master = master_frame
        self.app = app
        self.index = index
        self.widgets = [] 
        self.drag_win = None
        
        self.x_var = ctk.StringVar()
        self.y_var = ctk.StringVar()
        self.memo_var = ctk.StringVar()
        self.mode_var = ctk.StringVar()
        
        self.action_mode = data.get("action_mode", "move_only")

        self._build_widgets()
        self.update_data(data)

    def _build_widgets(self):
        row_pos = self.index + 2

        self.edit_frame = ctk.CTkFrame(self.master, fg_color="transparent")
        
        self.lbl_drag = ctk.CTkLabel(self.edit_frame, text="☰", width=20, cursor="hand2", text_color="gray", font=("Arial", 16))
        self.lbl_drag.pack(side="left", padx=(0, 5))
        self.lbl_drag.bind("<ButtonPress-1>", self.on_drag_start)
        self.lbl_drag.bind("<B1-Motion>", self.on_drag_motion)
        self.lbl_drag.bind("<ButtonRelease-1>", self.on_drag_release)

        self.check_var = tk.BooleanVar(value=False)
        self.chk_delete = ctk.CTkCheckBox(self.edit_frame, text="", variable=self.check_var, width=20, checkbox_width=20, checkbox_height=20)
        self.chk_delete.pack(side="left")

        self.btn_move = ctk.CTkButton(self.master, text="", width=80, fg_color=COLOR_DARK, text_color=COLOR_TEXT, command=lambda: self.app.start_hotkey_listen(self, "move"))
        self.btn_move.grid(row=row_pos, column=1, padx=5, pady=2)

        self.ent_x = ctk.CTkEntry(self.master, textvariable=self.x_var, width=60, justify="center", validate="key", validatecommand=self.app.vcmd)
        self.ent_x.grid(row=row_pos, column=2, padx=2)

        self.ent_y = ctk.CTkEntry(self.master, textvariable=self.y_var, width=60, justify="center", validate="key", validatecommand=self.app.vcmd)
        self.ent_y.grid(row=row_pos, column=3, padx=2)

        modes = ["좌표로 커서만 이동", "좌표로 스킬 사용", "커서 위치로 스킬 사용"]
        self.combo_mode = ctk.CTkOptionMenu(
            self.master, 
            values=modes, 
            variable=self.mode_var, 
            width=170, 
            command=self.on_mode_change,
            fg_color=COLOR_DARK, 
            button_color=COLOR_DARK, 
            button_hover_color=COLOR_SETTING,
            text_color=COLOR_TEXT
        )
        self.combo_mode.grid(row=row_pos, column=4, padx=5)

        self.btn_save = ctk.CTkButton(self.master, text="", width=80, fg_color=COLOR_DARK, text_color=COLOR_TEXT, command=lambda: self.app.start_hotkey_listen(self, "save"))
        self.btn_save.grid(row=row_pos, column=5, padx=5)

        self.ent_memo = ctk.CTkEntry(self.master, textvariable=self.memo_var, width=100)
        self.ent_memo.grid(row=row_pos, column=6, padx=(5, 10), sticky="w") 
        
        self.widgets.extend([self.edit_frame, self.btn_move, self.ent_x, self.ent_y, self.combo_mode, self.btn_save, self.ent_memo])
        self.memo_var.trace_add("write", self._auto_resize_memo)

        # [수정] 오직 X, Y 열에 포커스가 있을 때만 CTRL+S 동작 허용 로직
        def on_xy_focus_in(e):
            self.app.focused_row = self
            self.app.log(f"[{self.index + 1}행 선택됨] {self.app.capture_hk.upper()}로 좌표 지정 가능")

        def on_memo_focus_in(e):
            self.app.log(f"[{self.index + 1}행 메모 입력 중] 메모를 입력하세요.")

        def on_focus_out(e):
            def _check():
                if not self.app.winfo_exists(): return
                focus = self.app.focus_get()
                
                # CustomTkinter 위젯 특성상 내부 _entry도 확인
                is_in_xy = (focus in [self.ent_x, self.ent_y] or 
                           (hasattr(self.ent_x, '_entry') and focus in [self.ent_x._entry, self.ent_y._entry]))
                
                # XY 칸을 벗어났다면 focused_row 해제 및 로그 초기화
                if not is_in_xy:
                    self.app.focused_row = None
                    if self.app.target_combo.get() == "전체 화면":
                        self.app.log("⚠️ 목표 프로그램을 지정해 주세요.")
                    else:
                        self.app.log(" 준비 완료")
            self.app.after(100, _check)

        # X, Y 칸에만 캡처 활성화 바인딩
        for w in [self.ent_x, self.ent_y]:
            w.bind("<FocusIn>", on_xy_focus_in)
            w.bind("<FocusOut>", on_focus_out)
        
        # 메모 칸은 단순 안내만 표시
        self.ent_memo.bind("<FocusIn>", on_memo_focus_in)

    def on_mode_change(self, choice):
        mode_map_rev = {
            "좌표로 커서만 이동": "move_only",
            "좌표로 스킬 사용": "move_and_click",
            "커서 위치로 스킬 사용": "target_and_return_click"
        }
        self.action_mode = mode_map_rev.get(choice, "move_only")
        self.app.save_settings()

    def on_drag_start(self, event):
        self.lbl_drag.configure(text_color=COLOR_DANGER)
        
        self.drag_win = tk.Toplevel(self.app)
        self.drag_win.wm_overrideredirect(True)
        self.drag_win.attributes('-alpha', 0.8)
        self.drag_win.attributes('-topmost', True)
        
        hk = self.btn_move.cget("text")
        memo = self.memo_var.get()
        display_text = f"이동 중: {hk}" + (f" ({memo})" if memo else "")
        
        lbl = ctk.CTkLabel(self.drag_win, text=display_text, fg_color=COLOR_DARK, text_color=COLOR_TEXT, corner_radius=5, padx=10, pady=5, font=("Malgun Gothic", 12, "bold"))
        lbl.pack()
        self.drag_win.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")

        self.app.drop_indicator_win = tk.Toplevel(self.app)
        self.app.drop_indicator_win.wm_overrideredirect(True)
        self.app.drop_indicator_win.attributes('-topmost', True)
        self.app.drop_indicator_win.configure(bg="red")
        self.app.drop_indicator_win.geometry("0x0+-100+-100")
        
    def on_drag_motion(self, event):
        if self.drag_win and self.drag_win.winfo_exists():
            self.drag_win.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")
            
        drop_y = event.y_root
        target_idx = len(self.app.macro_rows)
        
        for i, r in enumerate(self.app.macro_rows):
            try:
                mid_y = r.btn_move.winfo_rooty() + (r.btn_move.winfo_height() / 2)
                if drop_y < mid_y:
                    target_idx = i
                    break
            except Exception: pass
            
        try:
            x = self.app.grid_frame.winfo_rootx() + 20
            width = self.app.grid_frame.winfo_width() - 40 
            
            if target_idx < len(self.app.macro_rows):
                y = self.app.macro_rows[target_idx].btn_move.winfo_rooty() - 3
            else:
                y = self.app.macro_rows[-1].btn_move.winfo_rooty() + self.app.macro_rows[-1].btn_move.winfo_height() + 2
                
            if hasattr(self.app, 'drop_indicator_win') and self.app.drop_indicator_win.winfo_exists():
                self.app.drop_indicator_win.geometry(f"{width}x3+{x}+{y}")
        except Exception: pass
        
    def on_drag_release(self, event):
        self.lbl_drag.configure(text_color="gray")
        
        if self.drag_win and self.drag_win.winfo_exists():
            self.drag_win.destroy()
            self.drag_win = None
            
        if hasattr(self.app, 'drop_indicator_win') and self.app.drop_indicator_win.winfo_exists():
            self.app.drop_indicator_win.destroy()
            self.app.drop_indicator_win = None
            
        drop_y = event.y_root
        target_idx = len(self.app.macro_rows)
        
        for i, r in enumerate(self.app.macro_rows):
            try:
                mid_y = r.btn_move.winfo_rooty() + (r.btn_move.winfo_height() / 2)
                if drop_y < mid_y:
                    target_idx = i
                    break
            except Exception: pass
            
        self.app.reorder_rows(self.index, target_idx)

    def set_edit_mode(self, is_edit):
        if is_edit:
            self.edit_frame.grid(row=self.index+2, column=0, padx=(0, 2), sticky="e")
        else:
            self.edit_frame.grid_remove()
            self.check_var.set(False)

    def update_data(self, data):
        self.btn_move.configure(text=data.get("move_hk", f"F{self.index+1}").upper())
        self.btn_save.configure(text=data.get("save_hk", "").upper())
        
        coords = data.get("coords", "0, 0").split(",")
        xv, yv = (coords[0].strip(), coords[1].strip()) if len(coords) == 2 else ("0", "0")
        self.x_var.set(xv)
        self.y_var.set(yv)
        self.memo_var.set(data.get("memo", ""))
        
        am = data.get("action_mode", "move_only")
        if am in ["default", "target_only"]: am = "move_only"
        self.action_mode = am 
        
        mode_map = {
            "move_only": "좌표로 커서만 이동",
            "move_and_click": "좌표로 스킬 사용",
            "target_and_return_click": "커서 위치로 스킬 사용"
        }
        self.mode_var.set(mode_map.get(self.action_mode, "좌표로 커서만 이동"))
        
        if hasattr(self, 'check_var'):
            self.check_var.set(False)

    def _auto_resize_memo(self, *args):
        text = self.memo_var.get()
        if not text:
            self.ent_memo.configure(width=100)
            return
            
        w = sum(13 if ord(c) > 127 else 8 for c in text) + 20
        self.ent_memo.configure(width=min(max(100, w), 400))

    def get_data(self):
        return {
            "move_hk": self.btn_move.cget("text"),
            "save_hk": self.btn_save.cget("text"),
            "coords": f"{self.x_var.get()}, {self.y_var.get()}",
            "memo": self.memo_var.get(),
            "action_mode": getattr(self, "action_mode", "move_only") 
        }

    def destroy(self):
        for w in self.widgets: w.destroy()


# --- [ 2. Main App ] ---
class MouseMacroApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Blue Archive Macro")
        self.iconbitmap(resource_path("icon.ico"))
        self.geometry("720x500") 
        self.minsize(100, 100) 
        
        self.resizable(True, True) 

        self.global_config_file = "app_config.json"
        self.current_file = "macro.json" 
        
        self.global_pause = False     
        self.is_settings_open = False 
        self.is_edit_mode = False     
        self.is_topmost = False
        self.is_listening = False 
        self.cancel_listen_func = None 
        self.listen_ignore_click = False 
        
        self.hook_id = None
        self.registered_hotkeys = []
        self._apply_hotkeys_job = None  
        
        self.mouse_lock = threading.Lock()
        self.active_hotkeys = set()
        
        self.focused_row = None
        self.capture_hotkey_hook = None
        self.tray_icon = None
        
        self.app_theme = "System"
        self.capture_hk = "CTRL+S" 
        self.pause_hk = "CTRL+TAB"
        self.exclusive_hook = True      
        self.show_realtime = True    
        self.click_delay = 1
        self.click_duration = 50
        self.return_delay = 10 
        self.after_action_delay = 10 
        self.auto_fire = True
        
        self.last_target_window = "전체 화면"
        self.last_resize_ratio = "16:9"
        self.last_resize_res = ""
        self.target_resolution = "알 수 없음"

        self.macro_rows = []
        self.vcmd = (self.register(self.validate_numeric_input), '%P')

        self.load_global_config()
        ctk.set_appearance_mode(self.app_theme)
        ctk.set_default_color_theme("blue")
        
        self.setup_menu() 
        self.setup_ui()
        
        self.target_combo.set(self.last_target_window)
        self.refresh_windows() 
        self.load_macro_profile(self.current_file)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind_all("<Button-1>", self.on_root_click)
        self.track_mouse()
        
        self.app_has_focus = True
        self.bind_all("<FocusIn>", self._on_app_focus_in)
        self.bind_all("<FocusOut>", self._on_app_focus_out)
        
        atexit.register(self._cleanup_hooks)

        if self.target_combo.get() == "전체 화면":
            self.log("⚠️ 목표 프로그램을 지정해 주세요.")
        else:
            self.log(" 준비 완료")

    def _on_app_focus_in(self, event):
        if not getattr(self, 'app_has_focus', False):
            self.app_has_focus = True
            if self.global_pause:
                self.request_apply_hotkeys()

    def _on_app_focus_out(self, event):
        def _check():
            if not self.focus_displayof():
                if getattr(self, 'app_has_focus', True):
                    self.app_has_focus = False
                    if self.global_pause:
                        self.request_apply_hotkeys()
        self.after(100, _check)

    def _cleanup_hooks(self):
        try: keyboard.unhook_all()
        except Exception: pass

    def minimize_to_tray(self):
        if not HAS_TRAY:
            messagebox.showerror("라이브러리 필요", "트레이 최소화 기능을 사용하려면 추가 라이브러리가 필요합니다.\n명령 프롬프트(CMD)에서 아래 명령어를 실행해주세요:\n\npip install pystray Pillow")
            return
            
        self.withdraw() 
        
        image = create_tray_icon()
        menu = pystray.Menu(
            pystray.MenuItem('창 열기', self.restore_from_tray),
            pystray.MenuItem('프로그램 종료', self.quit_from_tray)
        )
        self.tray_icon = pystray.Icon("MouseMacro", image, "몰루좌표 (실행 중)", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_from_tray(self, icon=None, item=None):
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self.deiconify)
        
    def quit_from_tray(self, icon=None, item=None):
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self.on_closing)

    def enable_capture_hotkey(self):
        if not self.capture_hk or self.capture_hk.lower() == "없음": return
        if self.capture_hotkey_hook: return

        def _do_capture():
            if getattr(self, 'focused_row', None):
                self.execute_capture(self.focused_row)

        try:
            self.capture_hotkey_hook = keyboard.add_hotkey(
                self.capture_hk.lower(), 
                _do_capture, 
                suppress=self.exclusive_hook
            )
        except Exception as e:
            print(f"캡처 핫키 등록 실패: {e}")

    def disable_capture_hotkey(self):
        if self.capture_hotkey_hook:
            try:
                keyboard.remove_hotkey(self.capture_hotkey_hook)
            except Exception: pass
            self.capture_hotkey_hook = None

    def load_global_config(self):
        if os.path.exists(self.global_config_file):
            try:
                with open(self.global_config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    last_file = config.get("last_json_path", "macro.json")
                    if os.path.exists(last_file):
                        self.current_file = last_file
                    
                    app_set = config.get("app_settings", {})
                    self.app_theme = app_set.get("app_theme", "System")
                    self.click_delay = app_set.get("click_delay", 1)
                    self.click_duration = app_set.get("click_duration", 100)
                    self.return_delay = app_set.get("return_delay", 10)
                    
                    self.after_action_delay = app_set.get("after_action_delay", 10)
                    self.auto_fire = app_set.get("auto_fire", True)
                    
                    self.capture_hk = app_set.get("capture_hk", "CTRL+S").upper()
                    self.pause_hk = app_set.get("pause_hk", "CTRL+TAB").upper()
                    self.exclusive_hook = app_set.get("exclusive_hook", True) 
                    self.show_realtime = app_set.get("show_realtime", True)
                    self.last_target_window = app_set.get("last_target_window", "전체 화면")
                    self.last_resize_ratio = app_set.get("last_resize_ratio", "16:9")
                    self.last_resize_res = app_set.get("last_resize_res", "")
            except Exception as e:
                self.log(f"⚠️ 전역 설정 로드 오류 (기본값 사용)")

    def save_global_config(self, event=None):
        config = {
            "last_json_path": self.current_file,
            "app_settings": {
                "app_theme": self.app_theme,
                "click_delay": self.click_delay,
                "click_duration": self.click_duration,
                "return_delay": self.return_delay,
                "after_action_delay": self.after_action_delay,
                "auto_fire": self.auto_fire,
                "capture_hk": self.capture_hk.upper(),
                "pause_hk": getattr(self, 'pause_hk', 'CTRL+TAB').upper(),
                "exclusive_hook": self.exclusive_hook, 
                "show_realtime": self.show_realtime,
                "last_target_window": self.target_combo.get() if hasattr(self, 'target_combo') else "전체 화면",
                "last_resize_ratio": self.last_resize_ratio,
                "last_resize_res": self.last_resize_res
            }
        }
        try:
            with open(self.global_config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception:
             self.log(f"⚠️ 전역 설정 저장 오류")

    def save_settings(self):
        data_list = [row.get_data() for row in self.macro_rows]
        export = {"target_resolution": self.target_resolution, "macros": data_list}
        try:
            with open(self.current_file, "w", encoding="utf-8") as f: 
                json.dump(export, f, ensure_ascii=False, indent=4)
            self.save_global_config()
            self.log("💾 매크로 데이터 저장 완료")
        except Exception as e:
             messagebox.showerror("저장 오류", f"파일을 저장하는 중 오류가 발생했습니다.\n{e}")
             self.log("⚠️ 저장 실패")

    def get_initial_presets(self):
        return [
            {"move_hk": "Q", "coords": "0, 0", "save_hk": "SHIFT+Q", "memo": "", "action_mode": "target_and_return_click"},
            {"move_hk": "W", "coords": "0, 0", "save_hk": "SHIFT+W", "memo": "", "action_mode": "target_and_return_click"},
            {"move_hk": "E", "coords": "0, 0", "save_hk": "SHIFT+E", "memo": "", "action_mode": "target_and_return_click"},
            {"move_hk": "F1", "coords": "0, 0", "save_hk": "SHIFT+F1", "memo": "", "action_mode": "move_only"},
            {"move_hk": "F2", "coords": "0, 0", "save_hk": "SHIFT+F2", "memo": "", "action_mode": "move_only"},
            {"move_hk": "F3", "coords": "0, 0", "save_hk": "SHIFT+F3", "memo": "", "action_mode": "move_only"}
        ]

    def load_macro_profile(self, filepath):
        if not os.path.exists(filepath):
            self.target_resolution = "알 수 없음"
            self.refresh_ui(self.get_initial_presets())
            return
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.target_resolution = data.get("target_resolution", "알 수 없음")
                macros = data.get("macros", self.get_initial_presets())
                self.refresh_ui(macros)
        except Exception as e:
             messagebox.showerror("로드 오류", f"파일을 불러오는 중 오류가 발생했습니다.\n{e}")
             self.refresh_ui(self.get_initial_presets())

    def new_file(self):
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f:
            self.current_file = f
            self.file_name_var.set(os.path.splitext(os.path.basename(f))[0])
            self.target_resolution = "알 수 없음"
            self.refresh_ui(self.get_initial_presets())
            self.save_settings()

    def save_as_settings(self):
        f = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if f: 
            self.current_file = f
            self.file_name_var.set(os.path.splitext(os.path.basename(f))[0])
            self.save_settings()

    def load_settings(self):
        f = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if f:
            self.current_file = f
            self.file_name_var.set(os.path.splitext(os.path.basename(f))[0])
            self.load_macro_profile(self.current_file)

    def execute_capture(self, row):
        if getattr(self, 'is_settings_open', False): 
            return
            
        pt = wintypes.POINT(); ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        hwnd = self.get_target_hwnd()
        if hwnd: ctypes.windll.user32.ScreenToClient(hwnd, ctypes.byref(pt))
        w, h = self.get_current_target_size(); self.target_resolution = f"{w}x{h}"
        
        def _update_ui():
            row.x_var.set(str(pt.x))
            row.y_var.set(str(pt.y))
            self.log(f"📸 {row.index+1}행 좌표 저장 완료")
            
        self.after(0, _update_ui)

    def execute_move(self, row):
        if getattr(self, 'global_pause', False) or getattr(self, 'is_settings_open', False) or getattr(self, 'is_listening', False): 
            return
            
        hk_str = row.btn_move.cget("text").lower()
        
        if hk_str in self.active_hotkeys:
            return
        self.active_hotkeys.add(hk_str)
            
        try:
            x_str, y_str = row.x_var.get(), row.y_var.get()
            x = int(x_str) if x_str not in ('', '-') else 0
            y = int(y_str) if y_str not in ('', '-') else 0
            
            hwnd = self.get_target_hwnd()
            
            def _worker():
                try:
                    while True:
                        if getattr(self, 'global_pause', False) or getattr(self, 'is_listening', False):
                            break
                            
                        with self.mouse_lock:
                            orig_pt = wintypes.POINT()
                            ctypes.windll.user32.GetCursorPos(ctypes.byref(orig_pt))
                            
                            if hwnd:
                                pt = wintypes.POINT(x, y)
                                ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
                                ctypes.windll.user32.SetCursorPos(pt.x, pt.y)
                            else: 
                                ctypes.windll.user32.SetCursorPos(x, y)
                            
                            mode = getattr(row, 'action_mode', 'move_only')
                            
                            if mode == "target_and_return_click":
                                time.sleep(self.click_delay / 1000.0)
                                click_mouse_sendinput(self.click_duration)
                                
                                time.sleep(self.return_delay / 1000.0)
                                ctypes.windll.user32.SetCursorPos(orig_pt.x, orig_pt.y)
                                
                                time.sleep(self.click_delay / 1000.0)
                                click_mouse_sendinput(self.click_duration)

                            elif mode == "move_and_click":
                                time.sleep(self.click_delay / 1000.0)
                                click_mouse_sendinput(self.click_duration)
                                
                                time.sleep(self.return_delay / 1000.0)
                                ctypes.windll.user32.SetCursorPos(orig_pt.x, orig_pt.y)
                                
                            elif mode == "move_only":
                                pass 
                            
                            if self.after_action_delay > 0:
                                time.sleep(self.after_action_delay / 1000.0)
                                
                        if getattr(self, 'global_pause', False) or getattr(self, 'is_listening', False):
                            break
                            
                        if not self.auto_fire or mode != "target_and_return_click":
                            while hk_str and keyboard.is_pressed(hk_str):
                                if getattr(self, 'global_pause', False) or getattr(self, 'is_listening', False):
                                    break
                                time.sleep(0.01)
                            break
                            
                        if not hk_str or not keyboard.is_pressed(hk_str):
                            break
                finally:
                    if hk_str in self.active_hotkeys:
                        self.active_hotkeys.remove(hk_str)

            threading.Thread(target=_worker, daemon=True).start()
            
        except ValueError:
            self.log("⚠️ X, Y 좌표값을 확인해주세요.")
            if hk_str in self.active_hotkeys:
                self.active_hotkeys.remove(hk_str)
        except Exception as e:
            self.log(f"⚠️ 이동 실행 오류: {e}")
            if hk_str in self.active_hotkeys:
                self.active_hotkeys.remove(hk_str)

    def check_duplicate_hk(self, new_hk, exclude_row=None, exclude_mode=None, is_global=False):
        if not new_hk: return False
        nhk = new_hk.lower()
        if not is_global:
            if getattr(self, 'capture_hk', '').lower() == nhk: return True
            if getattr(self, 'pause_hk', '').lower() == nhk: return True
            
        for r in self.macro_rows:
            if not (r == exclude_row and exclude_mode == "move"):
                if r.btn_move.cget("text").lower() == nhk: return True
            if not (r == exclude_row and exclude_mode == "save"):
                if r.btn_save.cget("text").lower() == nhk: return True
        return False

    def register_single_hotkey(self, hk_str, callback):
        if not hk_str or "대기" in hk_str: return
        hk_lower = hk_str.lower()
        try:
            keyboard.add_hotkey(hk_lower, callback, suppress=self.exclusive_hook)
            self.registered_hotkeys.append(hk_lower)
        except Exception as e:
            print(f"핫키 등록 실패 ({hk_str}): {e}")

    def request_apply_hotkeys(self):
        if self._apply_hotkeys_job is not None:
            self.after_cancel(self._apply_hotkeys_job)
        self._apply_hotkeys_job = self.after(150, self._do_apply_all_hotkeys)

    def _do_apply_all_hotkeys(self):
        self._apply_hotkeys_job = None
        self.apply_all_hotkeys()

    def apply_all_hotkeys(self):
        if getattr(self, 'is_listening', False):
            return
            
        try: keyboard.unhook_all()
        except Exception: pass
        self.registered_hotkeys = []
        self.capture_hotkey_hook = None
        
        if getattr(self, 'focused_row', None):
            self.enable_capture_hotkey()

        if getattr(self, 'is_settings_open', False):
            return

        if getattr(self, 'capture_hk', "") and self.capture_hk.lower() != "없음":
            self.register_single_hotkey(self.capture_hk, self._global_capture_action)

        if getattr(self, 'global_pause', False) and not getattr(self, 'app_has_focus', False):
            return

        if getattr(self, 'pause_hk', "") and self.pause_hk.lower() != "없음":
            self.register_single_hotkey(self.pause_hk, self.toggle_pause)

        if getattr(self, 'global_pause', False):
            return

        for row in self.macro_rows:
            s_hk = row.btn_save.cget("text")
            self.register_single_hotkey(s_hk, lambda r=row: self.execute_capture(r))

        for row in self.macro_rows:
            m_hk = row.btn_move.cget("text")
            self.register_single_hotkey(m_hk, lambda r=row: self.execute_move(r))

    def _global_capture_action(self):
        if getattr(self, 'focused_row', None):
            self.execute_capture(self.focused_row)
        else:
            self.log("⚠️ 좌표를 지정할 곳을 먼저 선택하세요.")

    def capture_hotkey(self, prompt_msg, callback_on_success, current_hk=""):
        if getattr(self, 'is_listening', False): return
        
        try: keyboard.unhook_all()
        except Exception: pass
        self.registered_hotkeys = []
        self.capture_hotkey_hook = None
        
        self.is_listening = True
        self.listen_ignore_click = True
        self.after(200, lambda: setattr(self, 'listen_ignore_click', False))
        
        self.log(prompt_msg)
        
        def finish_hk(new_hk):
            self.is_listening = False
            if self.hook_id:
                try: keyboard.unhook(self.hook_id)
                except Exception: pass
                self.hook_id = None
            callback_on_success(new_hk)
            
        pressed = set()
        def _listen(ev):
            try:
                if getattr(self, "cancel_listen", False): 
                    self.cancel_listen = False
                    finish_hk(current_hk)
                    return
                if ev.event_type == keyboard.KEY_DOWN:
                    pressed.add(ev.name.lower())
                    if ev.name.lower() == 'esc': finish_hk(current_hk)
                elif ev.event_type == keyboard.KEY_UP:
                    n = ev.name.lower()
                    if n not in {'shift', 'ctrl', 'alt', 'esc', 'right shift', 'right ctrl', 'right alt'}:
                        comb = []
                        if any(k in pressed for k in ['ctrl', 'right ctrl', 'left ctrl']): comb.append('Ctrl')
                        if any(k in pressed for k in ['shift', 'right shift', 'left shift']): comb.append('Shift')
                        if any(k in pressed for k in ['alt', 'right alt', 'left alt']): comb.append('Alt')
                        comb.append(n.upper())
                        finish_hk('+'.join(comb).upper())
                    if n in pressed: pressed.remove(n)
            except Exception as e:
                print(f"Hook Error: {e}")
                finish_hk(current_hk)

        try:
            self.hook_id = keyboard.hook(_listen, suppress=self.exclusive_hook)
            self.cancel_listen_func = lambda: finish_hk(current_hk)
        except Exception as e:
            messagebox.showerror("에러", f"단축키 입력을 대기할 수 없습니다: {e}")
            finish_hk(current_hk)

    def start_hotkey_listen(self, row, mode):
        if getattr(self, 'is_listening', False):
            if callable(getattr(self, 'cancel_listen_func', None)):
                self.cancel_listen_func()
                
        btn = row.btn_move if mode == "move" else row.btn_save
        old_hk = btn.cget("text")
        msg = "커서를 이동시킬 단축키를 입력하세요." if mode == "move" else "좌표를 지정할 단축키를 입력하세요."
        btn.configure(text="입력 대기...", fg_color=("#B8860B", "#D4A017"), text_color="white")
        
        def on_finish(new_hk):
            if new_hk and new_hk.lower() != old_hk.lower():
                if self.check_duplicate_hk(new_hk, exclude_row=row, exclude_mode=mode):
                    self.log(f"⚠️'{new_hk.upper()}'는 이미 사용 중인 단축키입니다.")
                    new_hk = old_hk
                else: 
                    if self.target_combo.get() == "전체 화면":
                        self.log("⚠️ 목표 프로그램을 지정해 주세요.")
                    else:
                        self.log(" 준비 완료") 
            else:
                if self.target_combo.get() == "전체 화면":
                    self.log("⚠️ 목표 프로그램을 지정해 주세요.")
                else:
                    self.log(" 준비 완료") 
                
            btn.configure(text=new_hk.upper() if new_hk else old_hk.upper(), fg_color=COLOR_DARK, text_color=COLOR_TEXT)
            self.request_apply_hotkeys() 
            
        self.capture_hotkey(msg, on_finish, old_hk)

    def on_target_change(self, value):
        self.save_global_config()
        if value == "전체 화면":
            self.log("⚠️ 목표 프로그램을 지정해 주세요.")
        else:
            self.log(" 준비 완료")

    def setup_menu(self):
        self.menubar = tk.Menu(self)
        
        fm = tk.Menu(self.menubar, tearoff=0)
        fm.add_command(label="📄 새 파일", command=self.new_file)
        fm.add_command(label="💾 저장", command=self.save_settings)
        fm.add_command(label="📝 다른 이름으로 저장", command=self.save_as_settings)
        fm.add_command(label="📂 불러오기", command=self.load_settings)
        fm.add_separator()
        fm.add_command(label="⬇️ 트레이로 최소화", command=self.minimize_to_tray)
        self.menubar.add_cascade(label="파일", menu=fm)
        
        sm = tk.Menu(self.menubar, tearoff=0)
        sm.add_command(label="환경 설정", command=self.open_global_settings_window)
        sm.add_command(label="매크로 설정", command=self.open_macro_settings_window)
        self.menubar.add_cascade(label="설정", menu=sm)
        
        tm = tk.Menu(self.menubar, tearoff=0)
        tm.add_command(label="🪟 창 크기 조절", command=self.open_resize_window)
        tm.add_command(label="🔄 좌표 파일 변환", command=self.open_convert_window)
        self.menubar.add_cascade(label="도구", menu=tm)
        
        im = tk.Menu(self.menubar, tearoff=0)
        im.add_command(label="ℹ️ 프로그램 정보", command=self.open_info_window)
        self.menubar.add_cascade(label="정보", menu=im)
        
        self.config(menu=self.menubar)

    def open_info_window(self):
        win = ctk.CTkToplevel(self)
        win.title("프로그램 정보")
        win.geometry("300x180")
        win.attributes("-topmost", True)  
        win.grab_set()
        
        ctk.CTkLabel(win, text="Blue Archive Macro", font=("Malgun Gothic", 14, "bold")).pack(pady=(20, 0))
        ctk.CTkLabel(win, text="V.1", font=("Malgun Gothic", 14, "bold")).pack(pady=(0, 5))
        
        link = ctk.CTkLabel(win, text="🔗 GitHub", text_color=("#1f538d", "#569cd6"), cursor="hand2")
        link.pack(pady=5)
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/chobo111/BlueArchive_Macro"))
        
        ctk.CTkButton(win, text="닫기", width=80, command=win.destroy).pack(pady=10)

    def toggle_pause(self):
        self.global_pause = not self.global_pause
        if self.global_pause:
            self.btn_pause.configure(text=f"⏸ 정지됨 ({self.pause_hk.upper()})", fg_color=COLOR_DANGER)
            self.log("⏸ 단축키 정지됨")
        else:
            self.btn_pause.configure(text=f"▶ 작동 중 ({self.pause_hk.upper()})", fg_color=COLOR_SUCCESS)
            if self.target_combo.get() == "전체 화면":
                self.log("⚠️ 목표 프로그램을 지정해 주세요.")
            else:
                self.log("▶ 단축키 활성화.")
            
        self.request_apply_hotkeys()

    def toggle_edit_mode(self):
        self.is_edit_mode = not self.is_edit_mode
        if self.is_edit_mode:
            self.btn_edit.configure(text="✅ 완료", text_color=COLOR_SUCCESS)
            self.btn_delete_selected.pack(side="left", padx=(0, 5)) 
            self.log("삭제할 행을 선택하거나 드래그해 순서를 변경하세요")
        else:
            self.btn_edit.configure(text="✏️ 편집", text_color=COLOR_TEXT)
            self.btn_delete_selected.pack_forget() 
            if self.target_combo.get() == "전체 화면":
                self.log("⚠️ 목표 프로그램을 지정해 주세요.")
            else:
                self.log("편집 완료")
        
        for row in self.macro_rows:
            row.set_edit_mode(self.is_edit_mode)

    def reorder_rows(self, from_idx, to_idx):
        if from_idx == to_idx: return
        data_list = [row.get_data() for row in self.macro_rows]
        item = data_list.pop(from_idx)
        
        if from_idx < to_idx:
            to_idx -= 1
            
        data_list.insert(to_idx, item)
        self.refresh_ui(data_list)
        self.save_settings()

    def delete_selected_rows(self):
        to_delete = [i for i, row in enumerate(self.macro_rows) if row.check_var.get()]
        
        if not to_delete:
            self.log("⚠️ 삭제할 행을 선택해주세요.")
            return
            
        if len(self.macro_rows) - len(to_delete) < 1:
            self.log("⚠️ 최소 1개의 행은 남겨야 합니다.")
            return
            
        data_list = [row.get_data() for i, row in enumerate(self.macro_rows) if i not in to_delete]
        self.refresh_ui(data_list)
        self.save_settings()
        self.log(f"🗑️ {len(to_delete)}개의 행이 삭제되었습니다.")

    def add_row(self):
        new_data = {
            "move_hk": "", 
            "coords": "0, 0", 
            "save_hk": "", 
            "memo": "", 
            "action_mode": "move_only"
        }
        self.macro_rows.append(MacroRow(self.grid_frame, len(self.macro_rows), new_data, self))
        if self.is_edit_mode:
            self.macro_rows[-1].set_edit_mode(True)
        self.request_apply_hotkeys()

    def setup_ui(self):
        self.ribbon = ctk.CTkFrame(self, height=40, corner_radius=0, fg_color="transparent")
        self.ribbon.pack(side="top", fill="x", padx=10, pady=(5,0))
        
        self.file_name_var = ctk.StringVar(value=os.path.splitext(os.path.basename(self.current_file))[0])
        ctk.CTkLabel(self.ribbon, text="📄 파일:").pack(side="left", padx=5)
        self.f_entry = ctk.CTkEntry(self.ribbon, textvariable=self.file_name_var, width=130)
        self.f_entry.pack(side="left")
        self.f_entry.bind("<Return>", lambda e: self.rename_file())
        
        self.btn_pause = ctk.CTkButton(self.ribbon, text=f"▶ 작동 중 ({self.pause_hk.upper()})", width=120, fg_color=COLOR_SUCCESS, text_color=("white", "white"), font=("Malgun Gothic", 12, "bold"), command=self.toggle_pause)
        self.btn_pause.pack(side="left", padx=10)
        
        self.btn_pin = ctk.CTkButton(self.ribbon, text="📌", width=30, command=self.toggle_topmost, fg_color="transparent", text_color=COLOR_TEXT, border_width=1)
        self.btn_pin.pack(side="left", padx=5)

        self.target_frame = ctk.CTkFrame(self, height=40, fg_color="transparent")
        self.target_frame.pack(side="top", fill="x", padx=10)
        
        self.target_combo = ctk.CTkComboBox(self.target_frame, width=200, values=["전체 화면"], command=self.on_target_change)
        self.target_combo.pack(side="left", padx=5)
        
        self.target_combo.bind("<Enter>", lambda e: self.refresh_windows())
        self.target_combo.bind("<Button-1>", lambda e: self.refresh_windows())
        
        self.target_size_var = ctk.StringVar(value="크기: 모니터 전체")
        ctk.CTkLabel(self.target_frame, textvariable=self.target_size_var, text_color=("gray50", "gray")).pack(side="left", padx=10)
        
        self.realtime_pos_var = ctk.StringVar(value="0, 0")
        self.realtime_pos_label = ctk.CTkLabel(self.target_frame, textvariable=self.realtime_pos_var, width=80, corner_radius=5, fg_color=("gray80", "#444444"), text_color=COLOR_TEXT, font=("Arial", 11, "bold"))
        if self.show_realtime: self.realtime_pos_label.pack(side="left", padx=20)

        self.bottom_container = ctk.CTkFrame(self, height=25, fg_color=("gray85", "#2D2D30"))
        self.bottom_container.pack(side="bottom", fill="x")
        self.status_var = ctk.StringVar(value=" 준비 완료")
        ctk.CTkLabel(self.bottom_container, textvariable=self.status_var, text_color=COLOR_SUCCESS).pack(side="left", padx=5)

        self.grid_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.grid_frame.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        
        btn_frame = ctk.CTkFrame(self.grid_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=0, columnspan=7, padx=5, pady=(0, 5), sticky="w")
        
        ctk.CTkButton(btn_frame, text="➕", width=30, height=24, command=self.add_row).pack(side="left", padx=(0, 2))
        self.btn_edit = ctk.CTkButton(btn_frame, text="✏️ 편집", width=55, height=24, fg_color="transparent", text_color=COLOR_TEXT, border_width=1, command=self.toggle_edit_mode)
        self.btn_edit.pack(side="left", padx=(0, 5))
        
        self.btn_delete_selected = ctk.CTkButton(btn_frame, text="🗑️ 삭제", width=55, height=24, fg_color="transparent", hover_color=COLOR_DANGER, text_color=COLOR_DANGER, border_width=1, command=self.delete_selected_rows)
        self.btn_delete_selected.pack_forget()
        
        ctk.CTkLabel(self.grid_frame, text="작동 단축키", font=("Malgun Gothic", 12, "bold")).grid(row=1, column=1, padx=5)

        headers = [("X", 2), ("Y", 3), ("작동 방식", 4), ("좌표 지정 단축키", 5), ("메모", 6)]
        for text, col in headers: 
            ctk.CTkLabel(self.grid_frame, text=text, font=("Malgun Gothic", 12, "bold")).grid(row=1, column=col, padx=5, sticky="w" if col==6 else "")

    def refresh_ui(self, data_list):
        while len(self.macro_rows) < len(data_list):
            new_idx = len(self.macro_rows)
            dummy_data = {"move_hk": f"F{new_idx+1}", "coords": "0, 0", "save_hk": "", "memo": "", "action_mode": "move_only"}
            self.macro_rows.append(MacroRow(self.grid_frame, new_idx, dummy_data, self))
            
        while len(self.macro_rows) > len(data_list):
            row = self.macro_rows.pop()
            row.destroy()
            
        for i, d in enumerate(data_list):
            self.macro_rows[i].update_data(d)
            if self.is_edit_mode:
                self.macro_rows[i].set_edit_mode(True)
            
        self.request_apply_hotkeys()

    def open_resize_window(self):
        hwnd = self.get_target_hwnd()
        if not hwnd: messagebox.showwarning("타겟 없음", "크기를 조절할 창을 먼저 선택하세요."); return
        win = ctk.CTkToplevel(self); win.title("창 크기 조절"); win.geometry("450x480")
        win.attributes("-topmost", True) 
        win.grab_set()
        
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        p169 = ["640x360", "854x480", "960x540", "1024x576", "1152x648", "1280x720", "1366x768", "1600x900", "1920x1080", "2560x1440", "3840x2160"]
        p43 = ["640x480", "800x600", "960x720", "1024x768", "1152x864", "1280x960", "1400x1050", "1600x1200"]
        f169 = [r for r in p169 if int(r.split('x')[0]) <= sw and int(r.split('x')[1]) <= sh]
        f43 = [r for r in p43 if int(r.split('x')[0]) <= sw and int(r.split('x')[1]) <= sh]
        r_var = ctk.StringVar(value=self.last_resize_ratio); p_combo = ctk.CTkComboBox(win, width=250, state="readonly")
        cf = ctk.CTkFrame(win, fg_color="transparent")
        ctk.CTkLabel(cf, text="W:").pack(side="left"); cw = ctk.CTkEntry(cf, width=80, validate="key", validatecommand=self.vcmd); cw.pack(side="left", padx=5)
        ctk.CTkLabel(cf, text="H:").pack(side="left"); ch = ctk.CTkEntry(cf, width=80, validate="key", validatecommand=self.vcmd); ch.pack(side="left", padx=5)
        def up_ui():
            self.last_resize_ratio = r_var.get()
            if self.last_resize_ratio == "16:9": 
                p_combo.configure(values=f169); p_combo.set(self.last_resize_res if self.last_resize_res in f169 else (f169[-1] if f169 else "")); p_combo.pack(pady=15); cf.pack_forget()
            elif self.last_resize_ratio == "4:3": 
                p_combo.configure(values=f43); p_combo.set(self.last_resize_res if self.last_resize_res in f43 else (f43[-1] if f43 else "")); p_combo.pack(pady=15); cf.pack_forget()
            else: p_combo.pack_forget(); cf.pack(pady=15)
        f = ctk.CTkFrame(win, fg_color="transparent"); f.pack(pady=10)
        for r in ["16:9", "4:3", "직접 입력"]: ctk.CTkRadioButton(f, text=r, variable=r_var, value=r, command=up_ui).pack(side="left", padx=10)
        up_ui()
        def do():
            try:
                if r_var.get() == "직접 입력": dw, dh = int(cw.get()), int(ch.get())
                else: self.last_resize_res = p_combo.get(); dw, dh = map(int, self.last_resize_res.split('x'))
                self.save_global_config(); rect = wintypes.RECT(0, 0, dw, dh)
                ctypes.windll.user32.AdjustWindowRectEx(ctypes.byref(rect), ctypes.windll.user32.GetWindowLongW(hwnd, -16), 0, ctypes.windll.user32.GetWindowLongW(hwnd, -20))
                ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, rect.right-rect.left, rect.bottom-rect.top, 0x0002); win.destroy()
            except ValueError: messagebox.showerror("오류", "해상도 숫자를 확인해주세요.")
            except Exception as e: messagebox.showerror("오류", f"창 크기 조절 실패: {e}")
        ctk.CTkButton(win, text="적용", fg_color=COLOR_SUCCESS, command=do).pack(side="bottom", pady=20)

    def open_convert_window(self):
        win = ctk.CTkToplevel(self); win.title("좌표 변환"); win.geometry("380x420")
        win.attributes("-topmost", True) 
        win.grab_set()
        
        ctk.CTkLabel(win, text="[ 다른 해상도에서 지정된 좌표 파일을 변환. ]", font=("Malgun Gothic", 12, "bold")).pack(pady=15)
        swv, shv = ctk.StringVar(), ctk.StringVar()
        if "x" in getattr(self, "target_resolution", ""):
            p = self.target_resolution.split("x"); swv.set(p[0].strip()); shv.set(p[1].strip())
        f1, f2 = ctk.CTkFrame(win, fg_color="transparent"), ctk.CTkFrame(win, fg_color="transparent")
        f1.pack(pady=5); f2.pack(pady=10)
        ctk.CTkLabel(f1, text="원본:").pack(side="left"); ctk.CTkEntry(f1, textvariable=swv, width=60, validate="key", validatecommand=self.vcmd).pack(side="left")
        ctk.CTkLabel(f1, text="x").pack(side="left"); ctk.CTkEntry(f1, textvariable=shv, width=60, validate="key", validatecommand=self.vcmd).pack(side="left")
        cw, ch = self.get_current_target_size(); dwv, dhv = ctk.StringVar(value=str(cw)), ctk.StringVar(value=str(ch))
        ctk.CTkLabel(f2, text="변경:").pack(side="left"); ctk.CTkEntry(f2, textvariable=dwv, width=60, validate="key", validatecommand=self.vcmd).pack(side="left")
        ctk.CTkLabel(f2, text="x").pack(side="left"); ctk.CTkEntry(f2, textvariable=dhv, width=60, validate="key", validatecommand=self.vcmd).pack(side="left")
        def do():
            try:
                sw, sh, dw, dh = int(swv.get()), int(shv.get()), int(dwv.get()), int(dhv.get())
                if sw == 0 or sh == 0: return
                rx, ry = dw/sw, dh/sh
                for r in self.macro_rows:
                    nx, ny = int(round(int(r.x_var.get())*rx)), int(round(int(r.y_var.get())*ry))
                    r.x_var.set(str(nx)); r.y_var.set(str(ny))
                self.target_resolution = f"{dw}x{dh}"; self.save_settings(); win.destroy()
            except ValueError: messagebox.showerror("오류", "해상도 숫자를 입력해주세요.")
            except Exception as e: messagebox.showerror("오류", f"변환 실패: {e}")
        ctk.CTkButton(win, text="변환 저장", fg_color=COLOR_SUCCESS, command=do).pack(side="bottom", pady=20)

    def open_global_settings_window(self):
        self.is_settings_open = True
        self.request_apply_hotkeys()
        
        win = ctk.CTkToplevel(self); win.title("환경 설정"); win.geometry("400x380")
        win.attributes("-topmost", True) 
        win.grab_set()
        
        def on_close():
            self.is_settings_open = False
            self.request_apply_hotkeys() 
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

        ctk.CTkLabel(win, text="[ 단축키 설정 ]", font=("Malgun Gothic", 12, "bold")).pack(pady=(15, 5))
        
        self.temp_capture_hk = self.capture_hk
        self.temp_pause_hk = self.pause_hk
        
        hk_frame = ctk.CTkFrame(win, fg_color="transparent")
        hk_frame.pack(fill="x", padx=20, pady=5)
        
        f_cap = ctk.CTkFrame(hk_frame, fg_color="transparent")
        f_cap.pack(side="left", expand=True)
        ctk.CTkLabel(f_cap, text="좌표 캡처").pack()
        btn_cap = ctk.CTkButton(f_cap, text=self.temp_capture_hk.replace("Control-", "Ctrl+").upper() if self.temp_capture_hk else "없음", width=120, fg_color=COLOR_DARK, text_color=COLOR_TEXT)
        btn_cap.pack(pady=5)
        
        f_pause = ctk.CTkFrame(hk_frame, fg_color="transparent")
        f_pause.pack(side="right", expand=True)
        ctk.CTkLabel(f_pause, text="작동/정지 (전체)").pack()
        btn_pause_hk = ctk.CTkButton(f_pause, text=self.temp_pause_hk.replace("Control-", "Ctrl+").upper() if self.temp_pause_hk else "없음", width=120, fg_color=COLOR_DARK, text_color=COLOR_TEXT)
        btn_pause_hk.pack(pady=5)
        
        def create_listener(hk_type, btn_widget):
            def start_listen():
                if getattr(self, 'is_listening', False):
                    if callable(getattr(self, 'cancel_listen_func', None)):
                        self.cancel_listen_func()
                        
                msg = f"{'좌표 캡처' if hk_type == 'cap' else '작동/정지'} 단축키를 입력하세요."
                current_temp = self.temp_capture_hk if hk_type == 'cap' else self.temp_pause_hk
                
                def on_finish(new_hk):
                    if new_hk and new_hk.lower() != current_temp.lower():
                        if self.check_duplicate_hk(new_hk, is_global=True):
                            self.log(f"⚠️'{new_hk.upper()}'는 이미 사용 중인 단축키입니다.")
                            new_hk = current_temp
                    
                    if hk_type == 'cap': self.temp_capture_hk = new_hk
                    else: self.temp_pause_hk = new_hk
                        
                    btn_widget.configure(text=new_hk.upper() if new_hk else "없음", fg_color=COLOR_DARK, text_color=COLOR_TEXT)
                    
                    if self.target_combo.get() == "전체 화면":
                        self.log("⚠️ 목표 프로그램을 지정해 주세요.")
                    else:
                        self.log(" 준비 완료")
                    
                btn_widget.configure(text="입력 대기...", fg_color=("#B8860B", "#D4A017"), text_color="white")
                self.capture_hotkey(msg, on_finish, current_temp)
            return start_listen

        btn_cap.configure(command=create_listener('cap', btn_cap))
        btn_pause_hk.configure(command=create_listener('pause', btn_pause_hk))
        
        exclusive_var = tk.BooleanVar(value=self.exclusive_hook)
        ctk.CTkCheckBox(win, text="독점 핫키 사용(타 프로그램의 단축키를 무시합니다.)", variable=exclusive_var).pack(pady=10)
        
        win.bind("<Button-1>", lambda e: self.on_root_click(e))
        
        ctk.CTkLabel(win, text="[ 프로그램 테마 ]", font=("Malgun Gothic", 12, "bold")).pack(pady=(15, 5))
        
        theme_combo = ctk.CTkOptionMenu(
            win, 
            values=["System", "Light", "Dark"], 
            command=lambda v: ctk.set_appearance_mode(v),
            fg_color=COLOR_DARK, 
            button_color=COLOR_DARK, 
            button_hover_color=COLOR_SETTING, 
            text_color=COLOR_TEXT
        )
        theme_combo.set(self.app_theme)
        theme_combo.pack(pady=5)
        
        chk_v = tk.BooleanVar(value=self.show_realtime); ctk.CTkCheckBox(win, text="실시간 좌표 표시 활성화", variable=chk_v).pack(pady=10)
        
        def sc():
            self.capture_hk = self.temp_capture_hk.upper()
            self.pause_hk = self.temp_pause_hk.upper()
            self.exclusive_hook = exclusive_var.get()
            self.show_realtime = chk_v.get()
            self.app_theme = theme_combo.get()
            
            if self.show_realtime: self.realtime_pos_label.pack(side="left", padx=20)
            else: self.realtime_pos_label.pack_forget()
            
            if self.global_pause:
                self.btn_pause.configure(text=f"⏸ 정지됨 ({self.pause_hk.upper()})")
            else:
                self.btn_pause.configure(text=f"▶ 작동 중 ({self.pause_hk.upper()})")
                
            ctk.set_appearance_mode(self.app_theme) 
            self.save_global_config() 
            self.is_settings_open = False
            self.request_apply_hotkeys() 
            win.destroy()
            
        ctk.CTkButton(win, text="저장", command=sc).pack(side="bottom", pady=20)

    def open_macro_settings_window(self):
        self.is_settings_open = True
        self.request_apply_hotkeys()
        
        win = ctk.CTkToplevel(self)
        win.title("매크로 설정")
        win.geometry("350x320")
        win.attributes("-topmost", True) 
        win.grab_set()
        
        def on_close():
            self.is_settings_open = False
            self.request_apply_hotkeys() 
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

        ctk.CTkLabel(win, text="[ 동작 딜레이 설정 ]", font=("Malgun Gothic", 12, "bold")).pack(pady=(15, 10))

        def add_e(t, v):
            f = ctk.CTkFrame(win, fg_color="transparent"); f.pack(pady=5)
            ctk.CTkLabel(f, text=t).pack(side="left", padx=5)
            sv = ctk.StringVar(value=str(v))
            ctk.CTkEntry(f, textvariable=sv, width=60, justify="center", validate="key", validatecommand=self.vcmd).pack(side="left")
            ctk.CTkLabel(f, text="ms", text_color="gray").pack(side="left", padx=3)
            return sv

        d1 = add_e("이동 후 클릭 대기:", self.click_delay)
        d2 = add_e("클릭 유지 간격:", self.click_duration)
        d3 = add_e("복귀 전 대기 시간:", self.return_delay)
        
        d4 = add_e("동작 완료 후 대기:", self.after_action_delay)
        
        auto_fire_var = tk.BooleanVar(value=self.auto_fire)
        ctk.CTkCheckBox(win, text="광클 활성화 (커서 위치로 스킬 사용 전용)", variable=auto_fire_var).pack(pady=10)
        
        def sc():
            try: 
                val1, val2, val3, val4 = d1.get(), d2.get(), d3.get(), d4.get()
                self.click_delay = int(val1) if val1 not in ('', '-') else 1
                self.click_duration = int(val2) if val2 not in ('', '-') else 100
                self.return_delay = int(val3) if val3 not in ('', '-') else 1
                
                self.after_action_delay = int(val4) if val4 not in ('', '-') else 0
            except ValueError: pass
            
            self.auto_fire = auto_fire_var.get()
            
            self.save_global_config() 
            self.is_settings_open = False
            self.request_apply_hotkeys() 
            win.destroy()
            
        ctk.CTkButton(win, text="저장 및 적용", command=sc).pack(side="bottom", pady=20)

    def get_target_hwnd(self):
        t = self.target_combo.get()
        return None if t == "전체 화면" else ctypes.windll.user32.FindWindowW(None, t)

    def get_current_target_size(self):
        h = self.get_target_hwnd()
        if h: rect = wintypes.RECT(); ctypes.windll.user32.GetClientRect(h, ctypes.byref(rect)); return rect.right, rect.bottom
        return self.winfo_screenwidth(), self.winfo_screenheight()

    def refresh_windows(self):
        titles = ["전체 화면"]
        def cb(hwnd, lp):
            if ctypes.windll.user32.IsWindowVisible(hwnd):
                length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
                    if buff.value and buff.value not in titles: titles.append(buff.value)
            return True
        
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_void_p)
        cb_func = EnumWindowsProc(cb)
        ctypes.windll.user32.EnumWindows(cb_func, 0)
        
        curr = self.target_combo.get(); self.target_combo.configure(values=titles); self.target_combo.set(curr if curr in titles else "전체 화면")

    def track_mouse(self):
        try:
            h = self.get_target_hwnd()
            if h: 
                r = wintypes.RECT()
                if ctypes.windll.user32.GetClientRect(h, ctypes.byref(r)):
                    self.target_size_var.set(f"크기: {r.right} x {r.bottom}")
                else:
                    self.target_size_var.set("크기: 확인 불가")
            else: 
                self.target_size_var.set("크기: 전체 모니터")
                
            if self.show_realtime:
                pt = wintypes.POINT()
                if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
                    if h: ctypes.windll.user32.ScreenToClient(h, ctypes.byref(pt))
                    self.realtime_pos_var.set(f"{pt.x}, {pt.y}")
                    
            self.after(50, self.track_mouse)
            
        except Exception as e:
            self.log(f"⚠️ 좌표 추적 에러: {e}")
            self.after(50, self.track_mouse) 

    def toggle_topmost(self):
        self.is_topmost = not self.is_topmost; self.attributes("-topmost", self.is_topmost); self.btn_pin.configure(fg_color=COLOR_SUCCESS if self.is_topmost else "transparent")

    def log(self, m): self.status_var.set(f" {m}")
    
    def validate_numeric_input(self, P): return (P == "" or P == "-" or P.lstrip("-").isdigit())
    
    def on_closing(self):
        try: self.save_settings()
        except Exception: pass
        try: keyboard.unhook_all()
        except Exception: pass
        if hasattr(self, 'tray_icon') and self.tray_icon:
            try: self.tray_icon.stop()
            except Exception: pass
        self.quit(); self.destroy(); os._exit(0)

    def on_root_click(self, event): 
        if getattr(self, 'listen_ignore_click', False): return
        if getattr(self, 'is_listening', False):
            if callable(getattr(self, 'cancel_listen_func', None)):
                self.cancel_listen_func() 
        if not isinstance(event.widget, (ctk.CTkEntry, tk.Entry, ctk.CTkComboBox, ctk.CTkOptionMenu, ctk.CTkCheckBox)):
            self.focus_set()

    def rename_file(self):
        nb = self.file_name_var.get().strip()
        op = self.current_file
        ob = os.path.splitext(os.path.basename(op))[0]
        
        if not nb or nb == ob: 
            self.file_name_var.set(ob)
            return
            
        np = os.path.join(os.path.dirname(op), nb + ".json")
        if os.path.exists(np): 
            messagebox.showwarning("경고", "이미 존재하는 파일 이름입니다.")
            self.file_name_var.set(ob)
            return
            
        try:
            if os.path.exists(op): 
                os.rename(op, np)
            self.current_file = np
            self.save_settings()
            self.log("📄 파일 이름 변경 완료")
        except OSError as e: 
            messagebox.showerror("오류", f"파일 이름 변경 중 오류 발생:\n{e}")
            self.file_name_var.set(ob)

if __name__ == "__main__":
    try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception: ctypes.windll.user32.SetProcessDPIAware()
    app = MouseMacroApp(); app.mainloop()
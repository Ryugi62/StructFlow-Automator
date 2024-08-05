import os
import subprocess
import time
import psutil
import win32gui
import win32con
import win32process
import customtkinter as ctk
import pyperclip
from tkinter import filedialog, Listbox, TclError
from dotenv import load_dotenv


class MidasWindowManager:
    def __init__(self):
        self.original_position = None
        self.original_size = None
        self.midas_hwnd = None

    def is_midas_gen_open(self, file_path):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    process = psutil.Process(pid)
                    try:
                        if any(
                            file_path.lower() in cmd.lower()
                            for cmd in process.cmdline()
                        ):
                            hwnds.append(hwnd)
                    except psutil.AccessDenied:
                        # Skip this process if we don't have permission to access its information
                        pass
                except psutil.NoSuchProcess:
                    pass
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        if hwnds:
            self.midas_hwnd = hwnds[0]
        return bool(hwnds)

    def open_midas_gen_file(self, file_path, x, y, width, height):
        midas_gen_executable = "C:\\Program Files\\MIDAS\\MODS\\Midas Gen\\MidasGen.exe"

        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = win32con.SW_HIDE
        proc = subprocess.Popen(
            [midas_gen_executable, file_path], startupinfo=startupinfo
        )

        while not self.is_midas_gen_open(file_path):
            print("Waiting for Midas Gen to open...")
            time.sleep(5)

        time.sleep(15)
            
        while not self.is_midas_gen_open(file_path):
            print("Waiting for Midas Gen to open...")
            time.sleep(5)

    def save_original_position_and_size(self, hwnd):
        rect = win32gui.GetWindowRect(hwnd)
        self.original_position = (rect[0], rect[1])
        self.original_size = (rect[2] - rect[0], rect[3] - rect[1])

    def set_window_position_and_size(self, hwnd, x, y, width, height):
        # 최소화 되어 있을 경우 복원
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,
            x,
            y,
            width,
            height,
            win32con.SWP_NOACTIVATE,
        )

        # z-order를 최하위로 설정
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_BOTTOM,
            0,
            0,
            0,
            0,
            win32con.SWP_NOSIZE | win32con.SWP_NOMOVE,
        )

    def restore_original_position_and_size(self):
        if self.midas_hwnd and self.original_position and self.original_size:
            try:
                top_hwnd = self.get_top_level_parent(self.midas_hwnd)
                self.set_window_position_and_size(
                    top_hwnd,
                    self.original_position[0],
                    self.original_position[1],
                    self.original_size[0],
                    self.original_size[1],
                )
            except win32gui.error as e:
                print(f"Failed to restore original position and size: {e}")

    def close_midas_gen(self):
        if self.midas_hwnd:
            win32gui.PostMessage(self.midas_hwnd, win32con.WM_CLOSE, 0, 0)

    def get_top_level_parent(self, hwnd):
        parent = hwnd
        while True:
            new_parent = win32gui.GetParent(parent)
            if new_parent == 0:
                return parent
            parent = new_parent

    def get_hwnds_for_pid(self, pid):
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds


class SingletonApp:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls, *args, **kwargs)
        return cls._instance


class OrderSelectionWidget(ctk.CTkToplevel):
    def __init__(self, parent, checked_items, file_entries):
        super().__init__(parent)
        self.parent = parent
        self.title("순서 선택")
        self.geometry("500x600")
        self.configure(fg_color="#2b2b2b")

        self.checked_items = checked_items
        self.file_entries = file_entries
        self.create_widgets()

        self.lift()
        self.focus_force()
        self.grab_set()
        self.attributes("-topmost", True)

    def create_widgets(self):
        title_label = ctk.CTkLabel(
            self, text="선택된 항목 순서 조정", font=("Roboto Medium", 20)
        )
        title_label.pack(pady=(20, 10))

        list_frame = ctk.CTkFrame(self)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.listbox = Listbox(
            list_frame,
            width=50,
            height=15,
            bg="#3a3a3a",
            fg="white",
            selectbackground="#1f6aa5",
        )
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")

        self.listbox.configure(yscrollcommand=scrollbar.set)

        for item in self.checked_items:
            self.listbox.insert("end", item)

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=20, fill="x")

        self.up_button = ctk.CTkButton(
            button_frame, text="위로", command=self.move_up, width=100
        )
        self.up_button.pack(side="left", padx=10)

        self.down_button = ctk.CTkButton(
            button_frame, text="아래로", command=self.move_down, width=100
        )
        self.down_button.pack(side="left", padx=10)

        self.start_button = ctk.CTkButton(
            button_frame, text="시작", command=self.start_program, width=100
        )
        self.start_button.pack(side="right", padx=10)

        files_frame = ctk.CTkFrame(self, fg_color="transparent")
        files_frame.pack(pady=10, fill="x")

        files_label = ctk.CTkLabel(
            files_frame, text="파일 위치", font=("Roboto Medium", 16)
        )
        files_label.pack(pady=(10, 5))

        for label, path in self.file_entries.items():
            file_label = ctk.CTkLabel(
                files_frame, text=f"{label}: {path}", font=("Roboto", 12)
            )
            file_label.pack(anchor="w", padx=20, pady=2)

    def move_up(self):
        try:
            selected = self.listbox.curselection()[0]
            if selected > 0:
                text = self.listbox.get(selected)
                self.listbox.delete(selected)
                self.listbox.insert(selected - 1, text)
                self.listbox.select_set(selected - 1)
        except (IndexError, TclError):
            pass

    def move_down(self):
        try:
            selected = self.listbox.curselection()[0]
            if selected < self.listbox.size() - 1:
                text = self.listbox.get(selected)
                self.listbox.delete(selected)
                self.listbox.insert(selected + 1, text)
                self.listbox.select_set(selected + 1)
        except (IndexError, TclError):
            pass

    def start_program(self):
        ordered_items = self.listbox.get(0, "end")
        print("순서:", ordered_items)

        # SimpleMouseTracker 실행
        self.parent.run_simple_mouse_tracker(ordered_items)

        self.grab_release()
        self.destroy()


class App(SingletonApp, ctk.CTk):
    WINDOW_GEOMETRY = "1280x768"
    WINDOW_TITLE = "Midas Linker"
    ICON_PATH = "./icons/StructFlow-Automator-Icon.ico"
    TABS_DATA = [("신규", 1)]
    BASIC_INFO_LABELS = ["태양광 명칭", "풍속 (m/s)", "설하중 (kN/m²)", "노풍도"]
    FILE_LOCATIONS = ["태양광", "건물", "디자인"]

    def __init__(self):
        super().__init__()
        load_dotenv()
        ctk.set_appearance_mode("dark")
        self.file_entries = {}
        self.tab_buttons = {}
        self.tabs_content = {}
        self.checkboxes = {}
        self.checkbox_functions = {
            "타입분할(태양광)": self.run_type_division_solar,
            "타입분할(건물)": self.run_type_division_building,
            "건물 / 태양광 통합": self.run_building_solar_integration,
            "크레인": self.run_crane,
            "지진": self.run_earthquake,
            "바닥 활하중": self.run_floor_load,
            "기타 고정하중": self.run_dead_load,
            "펄린": self.run_purlin,
            "기본형": self.run_basic_solar,
            "부착형": self.run_attached_solar,
            "알류미늄": self.run_aluminum_solar,
            "토지위(푸팅)": self.run_footing,
            "토지위(파일)": self.run_pile,
            "슬라브위": self.run_slab,
            "건물위": self.run_on_building,
            "접합부": self.run_joint,
            "안전로프": self.run_safety_rope,
        }
        self.configure_gui()
        self.create_layout()
        self.show_tab(1)
        self.json_directory = os.path.join(os.path.dirname(__name__), "json_scripts")
        self.ensure_json_directory()
        self.window_manager = MidasWindowManager()

    def configure_gui(self):
        self.geometry(self.WINDOW_GEOMETRY)
        self.title(self.WINDOW_TITLE)
        self.iconbitmap(self.ICON_PATH)
        self.resizable(False, False)

    def create_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.create_side_frame()
        self.create_main_frame()

    def create_side_frame(self):
        side_frame = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color="#333333")
        side_frame.grid(row=0, column=0, sticky="nswe")
        side_frame.grid_propagate(False)
        for name, index in self.TABS_DATA:
            button = ctk.CTkButton(
                side_frame,
                text=name,
                height=50,
                command=lambda idx=index: self.show_tab(idx),
                fg_color="#444444",
            )
            button.pack(pady=(20, 0) if index == 1 else 0, fill="x")
            self.tab_buttons[index] = button
        self.update_button_styles(1)

    def create_main_frame(self):
        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(0, weight=1)
        self.tabs_content = {1: self.create_tab_content_new(main_frame)}

    def update_button_styles(self, active_index):
        for index, button in self.tab_buttons.items():
            button.configure(corner_radius=0)
            button.configure(fg_color="#1f6aa5" if index == active_index else "#2b2b2b")

    def show_tab(self, tab_number):
        self.update_button_styles(tab_number)
        for frame in self.tabs_content.values():
            frame.grid_remove()
        self.tabs_content[tab_number].grid()

    def create_tab_content_new(self, parent):
        frame = ctk.CTkFrame(parent, width=1100, height=728, fg_color="#2b2b2b")
        frame.grid(sticky="nsew")
        frame.grid_propagate(False)

        sections = [
            ("파일 위치", 0.05, self.add_file_location_entries, 0.05),
            ("기본 정보", 0.30, self.add_basic_info_labels_and_entries, 0.05),
            ("모델링 형태", 0.55, self.add_modeling_type_checkboxes, 0.05),
            ("건물", 0.30, self.add_division_settings, 0.4),
            ("태양광 형태", 0.55, self.add_solar_type_checkboxes, 0.4),
            ("건물 정보", 0.30, self.add_advanced_building_info, 0.7),
            ("태양광 기타해석", 0.55, self.add_analysis_options, 0.7),
        ]

        for section in sections:
            text, rel_y, func, rel_x = section
            self.add_section_label(frame, text, rel_y, rel_x)
            func(frame, rel_y, rel_x)

        self.add_log_box(frame)
        self.add_create_button(frame)

        return frame

    def add_section_label(self, parent, text, y, x=0.05):
        ctk.CTkLabel(parent, text=text, font=("Roboto Medium", 20)).place(
            relx=x, rely=y, anchor=ctk.W
        )

    def open_file_dialog(self, entry_widget):
        file_path = filedialog.askopenfilename()
        if file_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, file_path)

    def add_file_location_entries(self, parent, start_y, start_x):
        y_offset = start_y + 0.05
        for label in self.FILE_LOCATIONS:
            entry = ctk.CTkEntry(parent, width=300)
            self.file_entries[label] = entry
            entry.place(relx=start_x + 0.1, rely=y_offset, anchor=ctk.W)
            ctk.CTkLabel(parent, text=f"{label} 파일").place(
                relx=start_x, rely=y_offset, anchor=ctk.W
            )
            button = ctk.CTkButton(
                parent,
                text="파일 선택",
                command=lambda e=entry: self.open_file_dialog(e),
                width=100,
                fg_color="#555555",
            )
            button.place(relx=start_x + 0.4, rely=y_offset, anchor=ctk.W)
            y_offset += 0.07

    def add_basic_info_labels_and_entries(self, parent, start_y, start_x):
        y_offset = start_y + 0.05
        for label in self.BASIC_INFO_LABELS:
            entry = ctk.CTkEntry(parent, width=180)
            entry.place(relx=start_x + 0.1, rely=y_offset, anchor=ctk.W)
            ctk.CTkLabel(parent, text=label).place(
                relx=start_x, rely=y_offset, anchor=ctk.W
            )
            y_offset += 0.05

    def add_modeling_type_checkboxes(self, parent, start_y, start_x):
        y_offset = start_y + 0.05
        options = ["타입분할(태양광)", "타입분할(건물)", "건물 / 태양광 통합"]
        for option in options:
            checkbox = ctk.CTkCheckBox(parent, text=option)
            checkbox.place(relx=start_x, rely=y_offset, anchor=ctk.W)
            self.checkboxes[option] = checkbox
            y_offset += 0.05

    def add_division_settings(self, parent, start_y, start_x):
        labels = ["크레인", "지진", "바닥 활하중", "기타 고정하중", "펄린"]
        self.add_checkboxes(parent, labels, start_y, start_x)

    def add_solar_type_checkboxes(self, parent, start_y, start_x):
        labels = ["기본형", "부착형", "알류미늄"]
        self.add_checkboxes(parent, labels, start_y, start_x)

    def add_advanced_building_info(self, parent, start_y, start_x):
        labels = ["토지위(푸팅)", "토지위(파일)", "슬라브위", "건물위"]
        self.add_checkboxes(parent, labels, start_y, start_x)

    def add_analysis_options(self, parent, start_y, start_x):
        labels = ["접합부", "안전로프"]
        self.add_checkboxes(parent, labels, start_y, start_x)

    def add_checkboxes(self, parent, labels, start_y, start_x):
        y_offset = start_y + 0.05
        for label in labels:
            checkbox = ctk.CTkCheckBox(parent, text=label)
            checkbox.place(relx=start_x, anchor=ctk.W, rely=y_offset)
            self.checkboxes[label] = checkbox
            y_offset += 0.05

    def add_log_box(self, parent):
        log_frame = ctk.CTkFrame(parent, width=850, height=20, fg_color="#2b2b2b")
        log_frame.place(relx=0.05, rely=0.91, anchor=ctk.W)
        self.log_box = ctk.CTkTextbox(log_frame, width=850, height=90)
        self.log_box.pack()

    def add_create_button(self, parent):
        create_button = ctk.CTkButton(
            parent,
            text="생성",
            height=90,
            command=self.open_order_selection,
            fg_color="#1f6aa5",
        )
        create_button.place(relx=0.9, rely=0.91, anchor=ctk.CENTER)

    def open_order_selection(self):
        checked_items = [
            label for label, checkbox in self.checkboxes.items() if checkbox.get()
        ]
        file_entries = {
            label: entry.get()
            for label, entry in self.file_entries.items()
            if entry.get()
        }
        if checked_items or file_entries:
            OrderSelectionWidget(self, checked_items, file_entries)
        else:
            print("체크된 항목이 없습니다.")

    def ensure_json_directory(self):
        if not os.path.exists(self.json_directory):
            os.makedirs(self.json_directory)

    def minimize_all_windows(self):
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == os.getpid():  # 현재 프로세스의 창만 최소화
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return True

        win32gui.EnumWindows(callback, None)

    def run_simple_mouse_tracker(self, ordered_items):
        # 모든 gui 최소화
        self.minimize_all_windows()

        for item in ordered_items:
            if item in self.checkbox_functions:
                self.checkbox_functions[item]()

        # 모든 gui 복원
        self.deiconify()

    def run_json_file(self, json_file):
        json_path = os.path.join(self.json_directory, json_file)
        if os.path.exists(json_path):
            subprocess.run(["python", "SimpleMouseTracker.py", json_path])
        else:
            print(f"Warning: JSON file {json_file} not found.")

    def save_clipboard_to_file(self, file_path):
        try:
            content = pyperclip.paste()
        except Exception as e:
            print(f"클립보드 접근 중 오류 발생: {e}")
            return False

        if not content.strip():
            print(
                f"경고: 클립보드가 비어 있습니다. 파일 {file_path}이(가) 저장되지 않았습니다."
            )
            return False

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
        except IOError as e:
            print(f"파일 저장 중 오류 발생: {e}")
            return False

        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            print(f"파일 {file_path}이(가) 성공적으로 저장되었습니다.")
            return True
        else:
            print(f"경고: 파일 {file_path}이(가) 비어 있거나 생성되지 않았습니다.")
            return False

    def clipboard_clear(self):
        pyperclip.copy("")

    def run_type_division_solar(self):
        while True:
            print("타입분할(태양광) 작업 시작")
            solar_file = self.file_entries["태양광"].get()
            if not solar_file:
                print("태양광 파일이 없습니다.")
                return

            if not self.window_manager.is_midas_gen_open(solar_file):
                self.window_manager.open_midas_gen_file(solar_file, 17, 14, 1906, 1028)
                self.window_manager.save_original_position_and_size(
                    self.window_manager.midas_hwnd
                )
                self.window_manager.set_window_position_and_size(
                    self.window_manager.midas_hwnd, 17, 14, 1906, 1028
                )
            else:
                print("Midas Gen이 이미 열려 있습니다.")
                self.window_manager.save_original_position_and_size(
                    self.window_manager.midas_hwnd
                )
                self.window_manager.set_window_position_and_size(
                    self.window_manager.midas_hwnd, 17, 14, 1906, 1028
                )

            self.run_json_file("display.json")
            self.run_json_file("calculate.json")

            # Steel code check
            self.clipboard_clear()
            self.run_json_file("steel_code_check.json")
            if not self.save_clipboard_to_file("solar_steel_code_check.txt"):
                print("Steel code check failed. Restarting from the beginning.")
                continue

            # Cold formed steel code check
            self.clipboard_clear()
            self.run_json_file("cold_formed_steel_code_check.json")
            if not self.save_clipboard_to_file(
                "solar_cold_formed_steel_code_check.txt"
            ):
                print(
                    "Cold formed steel code check failed. Restarting from the beginning."
                )
                continue

            # Table
            self.clipboard_clear()
            self.run_json_file("table.json")
            if not self.save_clipboard_to_file("solar_table.txt"):
                print("Table generation failed. Restarting from the beginning.")
                continue

            self.window_manager.restore_original_position_and_size()
            self.window_manager.close_midas_gen()
            print("타입분할(태양광) 작업 완료")
            break

    def run_type_division_building(self):
        while True:
            print("타입분할(건물) 작업 시작")
            building_file = self.file_entries["건물"].get()
            if not building_file:
                print("건물 파일이 없습니다.")
                return

            if not self.window_manager.is_midas_gen_open(building_file):
                self.window_manager.open_midas_gen_file(
                    building_file, 17, 14, 1906, 1028
                )
                self.window_manager.save_original_position_and_size(
                    self.window_manager.midas_hwnd
                )
                self.window_manager.set_window_position_and_size(
                    self.window_manager.midas_hwnd, 17, 14, 1906, 1028
                )
            else:
                print("Midas Gen이 이미 열려 있습니다.")
                self.window_manager.save_original_position_and_size(
                    self.window_manager.midas_hwnd
                )
                self.window_manager.set_window_position_and_size(
                    self.window_manager.midas_hwnd, 17, 14, 1906, 1028
                )

            self.run_json_file("display.json")
            self.run_json_file("calculate.json")

            # Steel code check
            self.clipboard_clear()
            self.run_json_file("steel_code_check.json")
            if not self.save_clipboard_to_file("building_steel_code_check.txt"):
                print("Steel code check failed. Restarting from the beginning.")
                continue

            # Cold formed steel code check
            self.clipboard_clear()
            self.run_json_file("cold_formed_steel_code_check.json")
            if not self.save_clipboard_to_file("building_cold_formed_steel_code_check.txt"):
                print("Cold formed steel code check failed. Restarting from the beginning.")
                continue

            # Table
            self.clipboard_clear()
            self.run_json_file("table.json")
            if not self.save_clipboard_to_file("building_table.txt"):
                print("Table generation failed. Restarting from the beginning.")
                continue

            self.window_manager.restore_original_position_and_size()
            self.window_manager.close_midas_gen()
            print("타입분할(건물) 작업 완료")
            break

    def run_building_solar_integration(self):
        print("건물 / 태양광 통합 작업 시작")
        self.run_json_file("calculate.json")
        # 여기에 중간 작업 추가 가능
        self.run_json_file("test.json")
        print("건물 / 태양광 통합 작업 완료")

    def run_crane(self):
        print("크레인 작업 시작")
        self.run_json_file("crane.json")
        print("크레인 작업 완료")

    def run_earthquake(self):
        print("지진 작업 시작")
        self.run_json_file("earthquake.json")
        print("지진 작업 완료")

    def run_floor_load(self):
        print("바닥 활하중 작업 시작")
        self.run_json_file("floor_load.json")
        print("바닥 활하중 작업 완료")

    def run_dead_load(self):
        print("기타 고정하중 작업 시작")
        self.run_json_file("dead_load.json")
        print("기타 고정하중 작업 완료")

    def run_purlin(self):
        print("펄린 작업 시작")
        self.run_json_file("purlin.json")
        print("펄린 작업 완료")

    def run_basic_solar(self):
        print("기본형 태양광 작업 시작")
        self.run_json_file("basic_solar.json")
        print("기본형 태양광 작업 완료")

    def run_attached_solar(self):
        print("부착형 태양광 작업 시작")
        self.run_json_file("attached_solar.json")
        print("부착형 태양광 작업 완료")

    def run_aluminum_solar(self):
        print("알류미늄 태양광 작업 시작")
        self.run_json_file("aluminum_solar.json")
        print("알류미늄 태양광 작업 완료")

    def run_footing(self):
        print("토지위(푸팅) 작업 시작")
        self.run_json_file("footing.json")
        print("토지위(푸팅) 작업 완료")

    def run_pile(self):
        print("토지위(파일) 작업 시작")
        self.run_json_file("pile.json")
        print("토지위(파일) 작업 완료")

    def run_slab(self):
        print("슬라브위 작업 시작")
        self.run_json_file("slab.json")
        print("슬라브위 작업 완료")

    def run_on_building(self):
        print("건물위 작업 시작")
        self.run_json_file("on_building.json")
        print("건물위 작업 완료")

    def run_joint(self):
        print("접합부 작업 시작")
        self.run_json_file("joint.json")
        print("접합부 작업 완료")

    def run_safety_rope(self):
        print("안전로프 작업 시작")
        self.run_json_file("safety_rope.json")
        print("안전로프 작업 완료")


if __name__ == "__main__":
    app = App()
    app.mainloop()

# app.py

import os
import subprocess
import time
import psutil
import win32gui
import win32con
import win32process
from tkinter import filedialog, Listbox, TclError
import customtkinter as ctk
from dotenv import load_dotenv


class SingletonApp:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls, *args, **kwargs)
        return cls._instance


class OrderSelectionWidget(ctk.CTkToplevel):
    def __init__(self, parent, checked_items, file_entries):
        super().__init__(parent)
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

        # Midas Gen 파일을 실행
        if "태양광" in self.file_entries:
            file_path = self.file_entries["태양광"]
            self.open_midas_gen_file(file_path, 17, 14, 1906, 1028)
            self.get_sunlight_data()
        elif "건물" in self.file_entries:
            file_path = self.file_entries["건물"]
            self.open_midas_gen_file(file_path, 17, 14, 1906, 1028)
            self.get_building_data()
        elif "디자인" in self.file_entries:
            file_path = self.file_entries["디자인"]
            self.open_midas_design_file(file_path, 17, 14, 1906, 1028)

        self.grab_release()
        self.destroy()

    def open_midas_gen_file(self, file_path, x, y, width, height):
        midas_gen_executable = "C:\\Program Files\\MIDAS\\MODS\\Midas Gen\\MidasGen.exe"  # 마이다스 Gen 실행 파일의 경로로 변경 필요
        proc = subprocess.Popen([midas_gen_executable, file_path])

        # 윈도우가 열릴 때까지 대기
        time.sleep(30)

        # proc으로부터 pid 가져오기
        pid = proc.pid

        # psutil을 사용하여 프로세스 객체 가져오기
        p = psutil.Process(pid)

        # 프로세스의 모든 창 핸들 가져오기
        def get_hwnds_for_pid(pid):
            def callback(hwnd, hwnds):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if found_pid == pid:
                        hwnds.append(hwnd)
                return True

            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            return hwnds

        hwnds = get_hwnds_for_pid(pid)

        # 첫 번째 핸들에 대해 창 크기 및 위치 설정
        if hwnds:
            win32gui.MoveWindow(hwnds[0], x, y, width, height, True)
        else:
            print("윈도우 핸들을 찾을 수 없습니다.")


    def open_midas_design_file(self, file_path, x, y, width, height):
        midas_design_executable = "C:\\Program Files\\MIDAS\\MODS\\Midas Design+\\Design+.exe"  # 마이다스 Design+ 실행 파일의 경로로 변경 필요
        proc = subprocess.Popen([midas_design_executable, file_path])

        # 윈도우가 열릴 때까지 대기
        time.sleep(30)

        # proc으로부터 pid 가져오기
        pid = proc.pid

        # psutil을 사용하여 프로세스 객체 가져오기
        p = psutil.Process(pid)

        # 프로세스의 모든 창 핸들 가져오기
        def get_hwnds_for_pid(pid):
            def callback(hwnd, hwnds):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                    if found_pid == pid:
                        hwnds.append(hwnd)
                return True

            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            return hwnds
        
        hwnds = get_hwnds_for_pid(pid)

        # 첫 번째 핸들에 대해 창 크기 및 위치 설정
        if hwnds:
            win32gui.MoveWindow(hwnds[0], x, y, width, height, True)
        else:
            print("윈도우 핸들을 찾을 수 없습니다.")

    def get_sunlight_data(self):
        # 태양광 데이터 갖고오기
        pass

    def get_building_data(self):
        # 건물 데이터 갖고오기
        pass


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
        self.configure_gui()
        self.create_layout()
        self.show_tab(1)

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
        options = ["타입분할", "건물 / 태양광 통합"]
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
            checkbox.place(relx=start_x, rely=y_offset, anchor=ctk.W)
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


if __name__ == "__main__":
    app = App()
    app.mainloop()

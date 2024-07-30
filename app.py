import os
from tkinter import filedialog, Listbox, Scrollbar
import customtkinter as ctk
from dotenv import load_dotenv
import subprocess

class SingletonApp:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls, *args, **kwargs)
        return cls._instance

class OrderSelectionWidget(ctk.CTkToplevel):
    def __init__(self, parent, checked_items):
        super().__init__(parent)
        self.title("순서 선택")
        self.geometry("400x500")

        self.checked_items = checked_items
        
        # Frame for listbox and scrollbar
        list_frame = ctk.CTkFrame(self)
        list_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Listbox with scrollbar
        self.listbox = Listbox(list_frame, width=50, height=20)
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")

        self.listbox.config(yscrollcommand=scrollbar.set)

        for item in checked_items:
            self.listbox.insert(ctk.END, item)

        self.up_button = ctk.CTkButton(self, text="위로", command=self.move_up)
        self.up_button.pack(side=ctk.LEFT, padx=10, pady=10)

        self.down_button = ctk.CTkButton(self, text="아래로", command=self.move_down)
        self.down_button.pack(side=ctk.LEFT, padx=10, pady=10)

        self.start_button = ctk.CTkButton(self, text="시작", command=self.start_program)
        self.start_button.pack(side=ctk.RIGHT, padx=10, pady=10)

    def move_up(self):
        selected = self.listbox.curselection()
        if selected and selected[0] > 0:
            text = self.listbox.get(selected[0])
            self.listbox.delete(selected[0])
            self.listbox.insert(selected[0] - 1, text)
            self.listbox.select_set(selected[0] - 1)

    def move_down(self):
        selected = self.listbox.curselection()
        if selected and selected[0] < self.listbox.size() - 1:
            text = self.listbox.get(selected[0])
            self.listbox.delete(selected[0])
            self.listbox.insert(selected[0] + 1, text)
            self.listbox.select_set(selected[0] + 1)

    def start_program(self):
        ordered_items = self.listbox.get(0, ctk.END)
        print("순서:", ordered_items)
        # 여기에 특정 프로그램을 실행하는 코드를 추가하세요
        # 예: subprocess.run(["프로그램_경로", "인자1", "인자2"])
        self.destroy()

class App(SingletonApp, ctk.CTk):
    WINDOW_GEOMETRY = "1280x768"
    WINDOW_TITLE = "Midas Linker"
    ICON_PATH = "./icons/StructFlow-Automator-Icon.ico"
    TABS_DATA = [("신규", 1)]
    BASIC_INFO_LABELS = [
        "폴더 위치",
        "태양광 명칭",
        "풍속 (m/s)",
        "설하중 (kN/m²)",
        "노풍도",
    ]

    def __init__(self):
        super().__init__()
        load_dotenv()
        ctk.set_appearance_mode("dark")
        self.address_entry = None
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
        side_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        side_frame.grid(row=0, column=0, sticky="nswe")
        side_frame.grid_propagate(False)
        for name, index in self.TABS_DATA:
            button = ctk.CTkButton(
                side_frame,
                text=name,
                height=50,
                command=lambda idx=index: self.show_tab(idx),
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
            ("기본 정보", 0.05, self.add_basic_info_labels_and_entries, 0.05),
            ("모델링 형태", 0.55, self.add_modeling_type_checkboxes, 0.05),
            ("건물", 0.05, self.add_division_settings, 0.4),
            ("태양광 형태", 0.55, self.add_solar_type_checkboxes, 0.4),
            ("건물 정보", 0.05, self.add_advanced_building_info, 0.7),
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

    def open_folder_dialog(self, entry_widget):
        folder_path = filedialog.askdirectory()
        if folder_path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, folder_path)

    def add_basic_info_labels_and_entries(self, parent, start_y, start_x):
        y_offset = start_y + 0.1
        for label in self.BASIC_INFO_LABELS:
            entry = ctk.CTkEntry(parent, width=180)
            if label == "폴더 위치":
                entry = ctk.CTkEntry(parent, width=110)
                self.address_entry = entry
                button = ctk.CTkButton(
                    parent,
                    text="폴더 선택",
                    command=lambda e=entry: self.open_folder_dialog(e),
                    width=10,
                )
                button.place(relx=start_x + 0.2, rely=y_offset, anchor=ctk.W)
            entry.place(relx=start_x + 0.1, rely=y_offset, anchor=ctk.W)
            ctk.CTkLabel(parent, text=label).place(
                relx=start_x, rely=y_offset, anchor=ctk.W
            )
            y_offset += 0.075

    def add_modeling_type_checkboxes(self, parent, start_y, start_x):
        y_offset = start_y + 0.1
        options = ["타입분할", "건물 / 태양광 통합"]
        for option in options:
            checkbox = ctk.CTkCheckBox(parent, text=option)
            checkbox.place(relx=start_x, rely=y_offset, anchor=ctk.W)
            self.checkboxes[option] = checkbox
            y_offset += 0.075

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
        y_offset = start_y + 0.1
        for label in labels:
            checkbox = ctk.CTkCheckBox(parent, text=label)
            checkbox.place(relx=start_x, rely=y_offset, anchor=ctk.W)
            self.checkboxes[label] = checkbox
            y_offset += 0.075

    def add_log_box(self, parent):
        log_frame = ctk.CTkFrame(parent, width=850, height=20, fg_color="#2b2b2b")
        log_frame.place(relx=0.05, rely=0.91, anchor=ctk.W)
        self.log_box = ctk.CTkTextbox(log_frame, width=850, height=90)
        self.log_box.pack()

    def add_create_button(self, parent):
        create_button = ctk.CTkButton(parent, text="생성", height=90, command=self.open_order_selection)
        create_button.place(relx=0.9, rely=0.91, anchor=ctk.CENTER)

    def open_order_selection(self):
        checked_items = [label for label, checkbox in self.checkboxes.items() if checkbox.get()]
        if checked_items:
            OrderSelectionWidget(self, checked_items)
        else:
            print("체크된 항목이 없습니다.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
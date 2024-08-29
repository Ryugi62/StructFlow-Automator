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
import pandas as pd
import sys
import requests
import re


class RedirectText:
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert("end", string)
        self.output.see("end")

    def flush(self):
        pass


class FileHandler:
    @staticmethod
    def extract_purlin_girth_data(file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Purlin _ Girth")
            df = df.iloc[2:].reset_index(drop=True)
            df.columns = [
                "CHK",
                "CHK.1",
                "Apply Member To",
                "Material",
                "Member Type",
                "Section",
                "Unnamed: 6",
                "Unnamed: 7",
                "Unnamed: 8",
                "Unnamed: 9",
                "Unnamed: 10",
                "Unnamed: 11",
                "Unnamed: 12",
                "Unnamed: 13",
                "Span",
                "Unnamed: 15",
                "Unnamed: 16",
                "Unbraced Length",
                "Unnamed: 18",
                "Factor",
                "Unnamed: 20",
                "Design Load",
                "Unnamed: 22",
                "Unnamed: 23",
                "Unnamed: 24",
                "Unnamed: 25",
                "Unnamed: 26",
                "Unnamed: 27",
                "Unnamed: 28",
                "Unnamed: 29",
                "Unnamed: 30",
                "Unnamed: 31",
                "Unnamed: 32",
                "Unnamed: 33",
                "Unnamed: 34",
                "Unnamed: 35",
                "Defl. Criteria",
                "Width-Thick Ratio",
                "Unnamed: 38",
                "Unnamed: 39",
                "Moment Strength",
                "Unnamed: 41",
                "Unnamed: 42",
                "Unnamed: 43",
                "Unnamed: 44",
                "Unnamed: 45",
                "Unnamed: 46",
                "Shear Strength",
                "Unnamed: 48",
                "Unnamed: 49",
                "Unnamed: 50",
                "Unnamed: 51",
                "Unnamed: 52",
                "Deflection",
                "Unnamed: 54",
            ]

            df["Deflection"] = (
                df["Deflection"].astype(str).str.extract(r"(\d+\.\d+)").astype(float)
            )
            df["Ratio"] = (
                df["Unnamed: 54"].astype(str).str.extract(r"(\d+\.\d+)").astype(float)
            )
            df["Calculated Span/300"] = (df["Span"] * 1000 / 300).round(2)
            output_df = df[["Deflection", "Calculated Span/300", "Ratio"]].dropna()
            result = output_df.to_string(index=False, header=False)
            pyperclip.copy(result)
            print("Data successfully extracted and copied to clipboard.")
            print("Extracted data:")
            print(result)
        except Exception as e:
            print(f"An error occurred while extracting data: {str(e)}")


class MidasWindowManager:
    def __init__(self):
        self.original_position = None
        self.original_size = None
        self.midas_hwnd = None

    def is_midas_gen_open(self, file_path):
        hwnds = self._get_hwnds_by_filepath(file_path)
        if hwnds:
            self.midas_hwnd = hwnds[0]
        return bool(hwnds)

    def _get_hwnds_by_filepath(self, file_path):
        hwnds = []

        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    process = psutil.Process(pid)
                    if any(
                        file_path.lower() in cmd.lower() for cmd in process.cmdline()
                    ):
                        hwnds.append(hwnd)
                except (psutil.AccessDenied, psutil.NoSuchProcess):
                    pass
            return True

        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    def open_midas_gen_file(self, file_path):
        midas_gen_executable = "C:\\Program Files\\MIDAS\\MODS\\Midas Gen\\MidasGen.exe"
        if not os.path.exists(midas_gen_executable):
            print("Midas Gen executable not found.")
            return False

        if self.is_midas_gen_open(file_path):
            print("Midas Gen is already open.")
            return True

        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = win32con.SW_HIDE
        subprocess.Popen([midas_gen_executable, file_path], startupinfo=startupinfo)

        while not self.is_midas_gen_open(file_path):
            print("Waiting for Midas Gen to open...")
            time.sleep(5)

        print("Midas Gen opened successfully.")
        time.sleep(15)
        return True

    def save_original_position_and_size(self):
        if self.midas_hwnd:
            rect = win32gui.GetWindowRect(self.midas_hwnd)
            self.original_position = (rect[0], rect[1])
            self.original_size = (rect[2] - rect[0], rect[3] - rect[1])

    def set_ui_position_and_size(self, hwnd, ini_file):
        try:
            window_title = win32gui.GetWindowText(hwnd)
            exe_path = os.path.join(
                os.path.dirname(__file__), "WindowLayoutManager.exe"
            )
            ini_file = os.path.join(os.path.dirname(__file__), "midas_gen.ini")

            success, _ = self.run_window_layout_manager(
                exe_path, window_title, ini_file
            )
            if success:
                print("Window layout restoration process completed successfully.")
            else:
                print("Window layout restoration failed or timed out.")
        except Exception as e:
            print(f"Failed to set UI position and size: {e}")

    def run_window_layout_manager(self, exe_path, window_title, ini_file, timeout=300):
        command = [exe_path, window_title, ini_file]
        try:
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            start_time = time.time()
            output_lines = []
            while True:
                output = process.stdout.readline()
                if output == "" and process.poll() is not None:
                    break
                if output:
                    output_lines.append(output.strip())
                    print(output_lines[-1])

                if time.time() - start_time > timeout:
                    print(f"Timeout after {timeout} seconds. Terminating process.")
                    process.terminate()
                    return False, output_lines

            if process.poll() == 0:
                print("Process completed successfully.")
                return True, output_lines
            else:
                print(f"Error occurred: {process.stderr.read()}")
                return False, output_lines
        except Exception as e:
            print(f"Subprocess failed: {e}")
            return False, []

    def restore_original_position_and_size(self):
        if self.midas_hwnd and self.original_position and self.original_size:
            try:
                top_hwnd = self.get_top_level_parent(self.midas_hwnd)
                if self.original_position == "maximized":
                    win32gui.ShowWindow(top_hwnd, win32con.SW_MAXIMIZE)
                else:
                    self.set_window_position_and_size(
                        top_hwnd,
                        self.original_position[0],
                        self.original_position[1],
                        self.original_size[0],
                        self.original_size[1],
                    )
            except win32gui.error as e:
                pass

    def minimize_window(self):
        if self.midas_hwnd:
            win32gui.ShowWindow(self.midas_hwnd, win32con.SW_MINIMIZE)

    def restore_window(self):
        if self.midas_hwnd:
            win32gui.ShowWindow(self.midas_hwnd, win32con.SW_RESTORE)

    def set_window_position_and_size(self, hwnd, x, y, width, height):
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,
            x,
            y,
            width,
            height,
            win32con.SWP_SHOWWINDOW,
        )

    def close_midas_gen(self):
        if self.midas_hwnd:
            win32gui.PostMessage(self.midas_hwnd, win32con.WM_CLOSE, 0, 0)
            self.midas_hwnd = None

    def get_top_level_parent(self, hwnd):
        parent = hwnd
        while True:
            new_parent = win32gui.GetParent(parent)
            if new_parent == 0:
                return parent
            parent = new_parent


class OrderSelectionWidget(ctk.CTkToplevel):
    def __init__(self, parent, checked_items, file_entries):
        super().__init__(parent)
        self.parent = parent
        self.checked_items = checked_items
        self.file_entries = file_entries
        self.configure_ui()

    def configure_ui(self):
        self.title("순서 선택")
        self.geometry("500x600")
        self.configure(fg_color="#2b2b2b")
        self.create_widgets()
        self.grab_set()
        self.focus_set()

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
        self.parent.window_manager.minimize_window()

        ordered_items = self.listbox.get(0, "end")
        for item in ordered_items:
            if item in self.parent.checkbox_functions:
                try:
                    self.parent.checkbox_functions[item]()
                except Exception as e:
                    print(f"Error in {item}: {str(e)}")

        self.parent.window_manager.restore_window()

        self.grab_release()
        self.destroy()


class App(ctk.CTk):
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
        self.checkbox_functions = self.get_checkbox_functions()
        self.configure_gui()
        self.create_layout()
        self.show_tab(1)
        self.json_directory = os.path.join(os.path.dirname(__name__), "json_scripts")
        self.ensure_json_directory()
        self.window_manager = MidasWindowManager()
        self.log_redirector = RedirectText(self.log_box)
        sys.stdout = self.log_redirector
        sys.stderr = self.log_redirector

    def get_checkbox_functions(self):
        return {
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
        file_path = filedialog.askopenfilename(
            filetypes=[("MGB Files", "*.mgb"), ("MDPB Files", "*.mdpb")]
        )
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

    def run_simple_mouse_tracker(self, ordered_items):
        for item in ordered_items:
            if item in self.checkbox_functions:
                try:
                    self.checkbox_functions[item]()
                except Exception as e:
                    print(f"Error in {item}: {str(e)}")

    def open_solar_file(self):
        solar_file = self.file_entries["태양광"].get()
        if not solar_file:
            print("태양광 파일이 없습니다.")
            return

        if not self.window_manager.open_midas_gen_file(solar_file):
            print("Midas Gen 파일이 열리지 않았습니다.")
            return

        self.window_manager.save_original_position_and_size()

        self.window_manager.set_ui_position_and_size(
            self.window_manager.midas_hwnd, "midas_gen.ini"
        )

        self.window_manager.set_window_position_and_size(
            self.window_manager.midas_hwnd, 50, 50, 1800, 930
        )

    def get_satellite_image_info(self):
        dotenv_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "StructFlow-Automator-Private",
            ".env",
        )
        load_dotenv(dotenv_path)

        api_key = os.getenv("GOOGLE_MAPS_API_KEY")
        if not api_key:
            raise ValueError("GOOGLE_MAPS_API_KEY not found in environment variables")

        return api_key

    def get_coordinates(self, address, api_key):
        base_url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {"address": address, "key": api_key}
        response = requests.get(base_url, params=params)
        data = response.json()

        if data["status"] == "OK":
            location = data["results"][0]["geometry"]["location"]
            return location["lat"], location["lng"]
        else:
            raise ValueError(
                f"Could not find coordinates for the given address. Status: {data['status']}"
            )

    def download_satellite_image(self, address, api_key, zoom="18", size="608x325"):
        lat, lng = self.get_coordinates(address, api_key)
        offset = 0.001 * (21 - int(zoom))
        box = f"{lat-offset},{lng-offset}|{lat-offset},{lng+offset}|{lat+offset},{lng+offset}|{lat+offset},{lng-offset}|{lat-offset},{lng-offset}"

        base_url = "https://maps.googleapis.com/maps/api/staticmap?"
        params = {
            "center": f"{lat},{lng}",
            "zoom": zoom,
            "size": size,
            "maptype": "satellite",
            "key": api_key,
            "markers": f"color:red|label:A|{lat},{lng}",
            "path": f"color:0xFFFF00FF|weight:5|{box}",
        }

        url = base_url + "&".join(f"{k}={v}" for k, v in params.items())
        response = requests.get(url)
        if response.status_code != 200:
            raise Exception(
                f"Failed to download image. Status code: {response.status_code}"
            )

        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        filename = os.path.join(temp_dir, "satellite_image.jpg")
        with open(filename, "wb") as f:
            f.write(response.content)

        print(f"Satellite image saved to: {filename}")

    def get_satellite_image(self, address=None):
        if not address:
            return
        while True:
            api_key = self.get_satellite_image_info()
            self.download_satellite_image(address, api_key)

            if not os.path.exists(self.get_temp_file_path("satellite_image.jpg")):
                continue
            break

    def get_address(self, key):
        if key == "sollar":
            file_path = self.file_entries["태양광"].get()
        elif key == "building":
            file_path = self.file_entries["건물"].get()
        else:
            return None

        match = re.search(
            r"([가-힣]+\s[가-힣]+\s[가-힣]+\s[가-힣]+\s[0-9-]+)", file_path
        )
        if match:
            address = match.group(1)
            address = re.sub(r"\s*[0-9-]+호.*$", "", address)
            return address
        else:
            print(f"주소를 찾을 수 없습니다: {file_path}")
            return None

    def clear_temp_dir(self):
        # clear temp dir
        temp_dir = os.path.join(os.path.dirname(__file__), "temp")
        for file in os.listdir(temp_dir):
            file_path = os.path.join(temp_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)

    def run_type_division_solar(self):
        try:
            self.clear_temp_dir()

            # 마이다스 프로그램 실행
            self.open_solar_file()

            # 위치도 이미지 생성
            self.get_satellite_image(address=self.get_address("sollar"))

            # 마이다스 내부 자료 생성 자동화 시작
            self.run_set_display()
            self.run_steel_code_check()
            self.run_cold_formed_steel_check()
            self.generate_table()
            self.generate_dummy_image()
            self.generate_boundaries_type()

            return
            self.set_reaction_force_moments()
        finally:
            return
            self.window_manager.restore_original_position_and_size()
            self.window_manager.close_midas_gen()
            print("타입분할(태양광) 작업 완료")

    def run_type_division_building(self):
        print("타입분할(건물) 작업 시작")
        try:
            pass
        finally:
            print("타입분할(건물) 작업 완료")

    def run_building_solar_integration(self):
        print("건물 / 태양광 통합 작업 시작")
        try:
            pass
        finally:
            print("건물 / 태양광 통합 작업 완료")

    def run_crane(self):
        print("크레인 작업 시작")
        try:
            pass
        finally:
            print("크레인 작업 완료")

    def run_earthquake(self):
        print("지진 작업 시작")
        try:
            pass
        finally:
            print("지진 작업 완료")

    def run_floor_load(self):
        print("바닥 활하중 작업 시작")
        try:
            pass
        finally:
            print("바닥 활하중 작업 완료")

    def run_dead_load(self):
        print("기타 고정하중 작업 시작")
        try:
            pass
        finally:
            print("기타 고정하중 작업 완료")

    def run_purlin(self):
        print("펄린 작업 시작")
        try:
            pass
        finally:
            print("펄린 작업 완료")

    def run_basic_solar(self):
        print("기본형 태양광 작업 시작")
        try:
            pass
        finally:
            print("기본형 태양광 작업 완료")

    def run_attached_solar(self):
        print("부착형 태양광 작업 시작")
        try:
            pass
        finally:
            print("부착형 태양광 작업 완료")

    def run_aluminum_solar(self):
        print("알류미늄 태양광 작업 시작")
        try:
            pass
        finally:
            print("알류미늄 태양광 작업 완료")

    def run_footing(self):
        print("토지위(푸팅) 작업 시작")
        try:
            pass
        finally:
            print("토지위(푸팅) 작업 완료")

    def run_pile(self):
        print("토지위(파일) 작업 시작")
        try:
            pass
        finally:
            print("토지위(파일) 작업 완료")

    def run_slab(self):
        print("슬라브위 작업 시작")
        try:
            pass
        finally:
            print("슬라브위 작업 완료")

    def run_on_building(self):
        print("건물위 작업 시작")
        try:
            pass
        finally:
            print("건물위 작업 완료")

    def run_joint(self):
        print("접합부 작업 시작")
        try:
            pass
        finally:
            print("접합부 작업 완료")

    def run_safety_rope(self):
        print("안전로프 작업 시작")
        try:
            pass
        finally:
            print("안전로프 작업 완료")

    def run_set_display(self):
        self.run_json_file("display.json")

    def run_steel_code_check(self):
        while True:
            try:
                # Open and copy steel code check
                self.run_json_file("open_widget_steel_code_check.json")
                self.clear_clipboard()
                self.run_json_file("copy_txt_steel_code_check.json")

                if not self.save_clipboard_to_file("solar_steel_code_check.txt"):
                    continue

                self.run_json_file("create_img_steel_code_check.json")
                if not os.path.exists(self.get_temp_file_path("100.emf")):
                    continue
            finally:
                self.run_json_file("close_steel_code_check.json")
            break

    def run_cold_formed_steel_check(self):
        while True:
            try:
                self.run_json_file("open_widget_cold_formed_steel_code_check.json")
                self.clear_clipboard()
                self.run_json_file("copy_txt_cold_formed_steel_code_check.json")

                if not self.save_clipboard_to_file(
                    "solar_cold_formed_steel_code_check.txt"
                ):
                    continue

                self.run_json_file("create_img_cold_formed_steel_code_check.json")
                if not os.path.exists(self.get_temp_file_path("201.emf")):
                    continue
            finally:
                self.run_json_file("close_cold_formed_steel_code_check.json")
            break

    def generate_table(self):
        while True:
            try:
                self.clear_clipboard()
                self.run_json_file("create_table.json")

                if not self.save_clipboard_to_file("solar_table.txt"):
                    continue
            finally:
                self.run_json_file("close_table.json")
            break

    def generate_dummy_image(self):
        while True:
            self.clear_clipboard()
            self.run_json_file("unactive_dummy.json")
            if not os.path.exists(self.get_temp_file_path("unactive_dummy.jpg")):
                continue
            break

    def generate_boundaries_type(self):
        while True:
            self.run_json_file("boundaries_type.json")
            if not os.path.exists(self.get_temp_file_path("boundaries_type.jpg")):
                continue
            break

    def set_reaction_force_moments(self):
        while True:
            self.run_json_file("set_reaction_force_moments.json")
            self.run_json_file("create_img_reaction_force_moments.json")
            if not os.path.exists(
                self.get_temp_file_path("reaction_force_moments.jpg")
            ):
                continue
            break

    def run_json_file(self, json_file):
        json_path = os.path.join(self.json_directory, json_file)
        if os.path.exists(json_path):
            subprocess.run(["python", "SimpleMouseTracker.py", json_path])
        else:
            print(f"Warning: JSON file {json_file} not found.")

    def clear_clipboard(self):
        pyperclip.copy("")

    def save_clipboard_to_file(self, file_name):
        try:
            content = pyperclip.paste()
            if not content.strip():
                return False
            file_path = self.get_temp_file_path(file_name)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)
            return True
        except Exception as e:
            print(f"Error saving clipboard content: {e}")
            return False

    def get_temp_file_path(self, file_name):
        temp_directory = os.path.join(os.path.dirname(__file__), "temp")
        if not os.path.exists(temp_directory):
            os.makedirs(temp_directory)
        return os.path.join(temp_directory, file_name)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    app = App()
    app.mainloop()

import win32gui
import win32process
import win32api
import win32con
import json
import os
import psutil
import ctypes
from ctypes import wintypes
import time

class WindowManager:
    def __init__(self):
        self.windows_info = {}
        self.log_file = "window_manager.log"

    def log(self, message):
        with open(self.log_file, "a") as f:
            f.write(f"{message}\n")
        print(message)

    def enum_windows_callback(self, hwnd, program_name):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            exe_name = self.get_process_name(pid)
            if exe_name and program_name.lower() in exe_name.lower():
                self.collect_window_info(hwnd)

    def collect_window_info(self, hwnd):
        parent_hwnd = win32gui.GetParent(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        
        left, top, right, bottom = rect
        width = right - left
        height = bottom - top

        # Get DPI for the window
        dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
        dpi_scale = dpi / 96.0

        if parent_hwnd:
            parent_rect = win32gui.GetWindowRect(parent_hwnd)
            relative_left = (left - parent_rect[0]) / (parent_rect[2] - parent_rect[0])
            relative_top = (top - parent_rect[1]) / (parent_rect[3] - parent_rect[1])
        else:
            relative_left = left
            relative_top = top

        # Get window style and extended style
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

        self.windows_info[hwnd] = {
            "parent": parent_hwnd,
            "relative_left": relative_left,
            "relative_top": relative_top,
            "width": width,
            "height": height,
            "dpi_scale": dpi_scale,
            "style": style,
            "ex_style": ex_style,
            "class_name": win32gui.GetClassName(hwnd),
            "title": win32gui.GetWindowText(hwnd)
        }

        self.log(f"Collected info for window {hwnd}: {self.windows_info[hwnd]}")

    def get_process_name(self, pid):
        try:
            process = psutil.Process(pid)
            return process.name()
        except Exception as e:
            self.log(f"[ERROR] Failed to get process name for pid {pid}: {e}")
            return ""

    def list_running_programs(self):
        process_list = {}
        win32gui.EnumWindows(self.enum_windows_callback_list, process_list)
        return process_list

    def enum_windows_callback_list(self, hwnd, process_list):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process_name = self.get_process_name(pid)
            if process_name and process_name not in process_list:
                process_list[process_name] = pid

    def save_window_positions(self, program_name):
        self.windows_info = {}
        win32gui.EnumWindows(self.enum_windows_callback, program_name)
        if not self.windows_info:
            self.log(f"[ERROR] No windows found for program: {program_name}")
            return

        with open(f"{program_name}_window_positions.json", "w") as f:
            json.dump(self.windows_info, f)
        self.log(f"[INFO] Window positions saved for program: {program_name}")

    def load_window_positions(self, program_name):
        if os.path.exists(f"{program_name}_window_positions.json"):
            with open(f"{program_name}_window_positions.json", "r") as f:
                self.windows_info = json.load(f)

            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

            # 메인 창 먼저 복원
            main_windows = [hwnd for hwnd, info in self.windows_info.items() if not info["parent"]]
            for hwnd in main_windows:
                self.restore_window(int(hwnd), self.windows_info[hwnd], screen_width, screen_height)
                time.sleep(0.1)  # 메인 창 복원 후 잠시 대기

            # 자식 창들 복원
            child_windows = [hwnd for hwnd, info in self.windows_info.items() if info["parent"]]
            for hwnd in child_windows:
                self.restore_window(int(hwnd), self.windows_info[hwnd], screen_width, screen_height)
                time.sleep(0.05)  # 각 자식 창 복원 후 잠시 대기

        else:
            self.log(f"[ERROR] No saved positions found for program: {program_name}")

    def restore_window(self, hwnd, pos, screen_width, screen_height):
        try:
            if not win32gui.IsWindow(hwnd):
                self.log(f"[WARNING] Window {hwnd} no longer exists. Skipping...")
                return

            if pos["parent"]:
                parent_rect = win32gui.GetWindowRect(int(pos["parent"]))
                left = int(parent_rect[0] + pos["relative_left"] * (parent_rect[2] - parent_rect[0]))
                top = int(parent_rect[1] + pos["relative_top"] * (parent_rect[3] - parent_rect[1]))
            else:
                left = int(pos["relative_left"])
                top = int(pos["relative_top"])

            # Apply DPI scaling
            dpi_scale = pos.get("dpi_scale", 1.0)
            left = int(left / dpi_scale)
            top = int(top / dpi_scale)
            width = int(pos["width"] / dpi_scale)
            height = int(pos["height"] / dpi_scale)

            # 화면 경계 확인
            left = max(0, min(left, screen_width - width))
            top = max(0, min(top, screen_height - height))

            self.log(f"[DEBUG] Restoring window: hwnd={hwnd}, left={left}, top={top}, width={width}, height={height}")
            
            # Set window style and extended style
            win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, pos["style"])
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, pos["ex_style"])
            
            # Move and resize the window
            flags = win32con.SWP_NOZORDER | win32con.SWP_FRAMECHANGED
            if "tool" in pos["class_name"].lower() or "palette" in pos["class_name"].lower():
                flags |= win32con.SWP_NOACTIVATE  # 도구 창은 활성화하지 않음

            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, left, top, width, height, flags)
            
            # Force redraw
            win32gui.RedrawWindow(hwnd, None, None, 
                                  win32con.RDW_INVALIDATE | win32con.RDW_ERASE | win32con.RDW_FRAME | win32con.RDW_ALLCHILDREN)

        except Exception as e:
            self.log(f"[ERROR] Failed to move window {hwnd}: {e}")

    def print_window_hierarchy(self, program_name):
        self.windows_info = {}
        win32gui.EnumWindows(self.enum_windows_callback, program_name)
        
        def print_hierarchy(hwnd, level=0):
            info = self.windows_info.get(str(hwnd), {})
            indent = "  " * level
            self.log(f"{indent}Window {hwnd}:")
            self.log(f"{indent}  Title: {info.get('title', 'N/A')}")
            self.log(f"{indent}  Class: {info.get('class_name', 'N/A')}")
            self.log(f"{indent}  Position: ({info.get('relative_left', 'N/A')}, {info.get('relative_top', 'N/A')})")
            self.log(f"{indent}  Size: {info.get('width', 'N/A')}x{info.get('height', 'N/A')}")
            self.log(f"{indent}  Style: {info.get('style', 'N/A')}")
            self.log(f"{indent}  Extended Style: {info.get('ex_style', 'N/A')}")
            
            for child_hwnd, child_info in self.windows_info.items():
                if child_info.get('parent') == hwnd:
                    print_hierarchy(int(child_hwnd), level + 1)

        root_windows = [int(hwnd) for hwnd, info in self.windows_info.items() if not info.get('parent')]
        for root_hwnd in root_windows:
            print_hierarchy(root_hwnd)

if __name__ == "__main__":
    manager = WindowManager()

    while True:
        print("\n현재 실행 중인 프로그램 목록:")
        process_list = manager.list_running_programs()

        if not process_list:
            manager.log("[ERROR] 실행 중인 프로그램을 찾을 수 없습니다.")
            break

        for i, process_name in enumerate(process_list.keys(), 1):
            print(f"{i}. {process_name}")

        choice = input("\n프로그램을 선택하세요 (번호 입력, q 입력 시 종료): ")

        if choice.lower() == "q":
            break

        try:
            choice = int(choice)
            if choice < 1 or choice > len(process_list):
                manager.log("[ERROR] 잘못된 선택입니다.")
                continue

            selected_program = list(process_list.keys())[choice - 1]
            print(f"\n선택한 프로그램: {selected_program}")

            action = input("1: 위치 저장, 2: 위치 복구, 3: 창 계층 구조 출력\n선택하세요: ")

            if action == "1":
                manager.save_window_positions(selected_program)
            elif action == "2":
                manager.load_window_positions(selected_program)
            elif action == "3":
                manager.print_window_hierarchy(selected_program)
            else:
                manager.log("[ERROR] 잘못된 선택입니다.")
        except ValueError:
            manager.log("[ERROR] 유효한 번호를 입력하세요.")
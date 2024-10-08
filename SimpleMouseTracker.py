import sys
import os
import time
import json
import logging
import cv2
import numpy as np
import psutil
import win32gui
import win32process
import win32api
import win32con
import win32ui
import threading
from logging.handlers import RotatingFileHandler
from ctypes import windll
import subprocess
import pywintypes
import keyboard


# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
LOG_FILE = "mouse_tracker.log"
SAMPLE_TARGETS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__name__)),
    "StructFlow-Automator-Private",
    "sample_targets",
)

# Maximum retry count for click events
MAX_RETRY_COUNT = 5
MINIMUM_CLICK_DELAY = 2  # 최소 대기시간 2초


def configure_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[RotatingFileHandler(LOG_FILE, maxBytes=10**6, backupCount=3)],
    )


configure_logging()


class InputBlocker:
    def __init__(self, hwnd):
        self.hwnd = hwnd
        self.overlay_hwnd = None
        self.stop_event = threading.Event()
        self.class_name = "OverlayWindowClass"
        self.update_thread = None

    def create_overlay(self):
        def overlay_window_proc(hwnd, msg, wparam, lparam):
            if msg == win32con.WM_CLOSE:
                win32gui.DestroyWindow(hwnd)
            elif msg in [
                win32con.WM_LBUTTONDOWN,
                win32con.WM_RBUTTONDOWN,
                win32con.WM_MBUTTONDOWN,
                win32con.WM_LBUTTONUP,
                win32con.WM_RBUTTONUP,
                win32con.WM_MBUTTONUP,
                win32con.WM_MOUSEMOVE,
                win32con.WM_MOUSEWHEEL,
                win32con.WM_KEYDOWN,
                win32con.WM_KEYUP,
                win32con.WM_CHAR,
                win32con.WM_SYSKEYDOWN,
                win32con.WM_SYSKEYUP,
            ]:
                return 0  # 모든 입력 이벤트를 차단
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = overlay_window_proc
        wc.lpszClassName = self.class_name
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)

        try:
            win32gui.RegisterClass(wc)
        except win32gui.error as e:
            if e.winerror != 1410:  # "Class already exists" error
                raise

        rect = win32gui.GetWindowRect(self.hwnd)
        try:
            # 오버레이 창을 생성할 때 WS_EX_LAYERED 속성을 추가해야 합니다.
            self.overlay_hwnd = win32gui.CreateWindowEx(
                win32con.WS_EX_LAYERED
                | win32con.WS_EX_TOPMOST
                | win32con.WS_EX_NOACTIVATE,
                self.class_name,
                "Overlay",
                win32con.WS_POPUP,
                rect[0] - 5,
                rect[1] - 5,
                rect[2] - rect[0] + 10,
                rect[3] - rect[1] + 10,
                0,
                0,
                0,
                None,
            )
        except win32gui.error as e:
            logging.error(f"Failed to create overlay window: {e}")
            return

        # 노란색 설정 (RGB: 255, 255, 0)
        yellow_color = win32api.RGB(255, 255, 0)

        # SetLayeredWindowAttributes 호출 시 오버레이 창이 WS_EX_LAYERED 속성을 가지고 있어야 합니다.
        win32gui.SetLayeredWindowAttributes(
            self.overlay_hwnd,
            yellow_color,
            128,  # 알파값 (0-255, 여기서는 50% 투명도)
            win32con.LWA_COLORKEY | win32con.LWA_ALPHA,
        )

        win32gui.ShowWindow(self.overlay_hwnd, win32con.SW_SHOW)

        # "blocked" 텍스트 그리기
        hdc = win32gui.GetDC(self.overlay_hwnd)
        font = win32ui.CreateFont(
            {"name": "Arial", "height": 40, "weight": win32con.FW_BOLD}
        )
        win32gui.SelectObject(hdc, font.GetSafeHandle())
        win32gui.SetTextColor(hdc, win32api.RGB(0, 0, 0))  # 검은색 텍스트
        text_rect = (0, 0, rect[2] - rect[0], rect[3] - rect[1])
        win32gui.DrawText(
            hdc,
            "blocked",
            -1,
            text_rect,
            win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE,
        )
        win32gui.ReleaseDC(self.overlay_hwnd, hdc)

        # x2, y2, width2, height2 = Gen 2024 - [C:\\Users\\xorjf\\OneDrive\\바탕 화면\\3764-(에스비일렉트릭)경북 예천군 용궁면 덕계리 380-1, 381 도성기1~5호태양광발전소(축사위)-完 (1)\\02-(마이다스)구조계산\\태양광\\[태양광]도성기1~5] - [MIDAS/Gen]"의
        # 사이즈를 가져와서 그 사이즈로 오버레이 창을 생성합니다.
        self.window_hwnd = sys.argv[2]
        x2, y2, width2, height2 = win32gui.GetWindowRect(self.window_hwnd)
        self.overlay2_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST | win32con.WS_EX_NOACTIVATE,
            self.class_name,
            "Overlay",
            win32con.WS_POPUP,
            x2 - 5,
            y2 - 5,
            width2 - x2 + 10,
            height2 - y2 + 10,
            0,
            0,
            0,
            None,
        )

        # 노란색 설정 (RGB: 255, 255, 0)
        yellow_color = win32api.RGB(255, 255, 0)

        # SetLayeredWindowAttributes 호출 시 오버레이 창이 WS_EX_LAYERED 속성을 가지고 있어야 합니다.
        win32gui.SetLayeredWindowAttributes(
            self.overlay2_hwnd,
            yellow_color,
            128,  # 알파값 (0-255, 여기서는 50% 투명도)
            win32con.LWA_COLORKEY | win32con.LWA_ALPHA,
        )

        win32gui.ShowWindow(self.overlay2_hwnd, win32con.SW_SHOW)

        # "blocked" 텍스트 그리기
        hdc = win32gui.GetDC(self.overlay2_hwnd)
        font = win32ui.CreateFont(
            {"name": "Arial", "height": 40, "weight": win32con.FW_BOLD}
        )
        win32gui.SelectObject(hdc, font.GetSafeHandle())
        win32gui.SetTextColor(hdc, win32api.RGB(0, 0, 0))
        text_rect = (0, 0, 1920, 1080)
        win32gui.DrawText(
            hdc,
            "blocked",
            -1,
            text_rect,
            win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE,
        )
        win32gui.ReleaseDC(self.overlay2_hwnd, hdc)

        self.update_thread = threading.Thread(target=self.update_z_order_loop)
        self.update_thread.start()

        # 해당 input blocker의 z-order를 hwnd_bottom으로 이동시킵니다.
        win32gui.SetWindowPos(
            self.overlay_hwnd,
            win32con.HWND_BOTTOM,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
        )
        win32gui.SetWindowPos(
            self.overlay2_hwnd,
            win32con.HWND_BOTTOM,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
        )

        while not self.stop_event.is_set():
            win32gui.PumpWaitingMessages()

    def update_z_order_loop(self):
        while not self.stop_event.is_set():
            self.update_z_order()
            time.sleep(0.1)  # 업데이트 주기를 0.1초로 늘림

    def is_hwnd_bottom(self, hwnd):
        """
        제공된 hwnd가 Z-order의 HWND_BOTTOM인지 확인합니다.
        """
        # 가장 아래에 있는 창을 가져옵니다.
        bottom_hwnd = win32gui.GetWindow(win32gui.GetDesktopWindow(), win32con.GW_CHILD)

        # bottom_hwnd가 None이면 창이 없는 경우이므로 종료
        if bottom_hwnd == 0:
            print("There are no windows.")
            return False

        # Z-order에서 가장 아래에 있는 창을 탐색
        while bottom_hwnd:
            next_hwnd = win32gui.GetWindow(bottom_hwnd, win32con.GW_HWNDNEXT)
            if next_hwnd == 0:  # 더 이상 아래로 갈 창이 없으면 가장 아래 창임
                break
            bottom_hwnd = next_hwnd

        # 가장 아래에 있는 hwnd와 제공된 hwnd를 비교
        if hwnd == bottom_hwnd:
            print(f"hwnd {hwnd} is at HWND_BOTTOM.")
            return True
        else:
            print(f"hwnd {hwnd} is NOT at HWND_BOTTOM.")
            return False

    def update_z_order(self):
        if (
            self.overlay_hwnd
            and win32gui.IsWindow(self.overlay_hwnd)
            and win32gui.IsWindow(self.hwnd)
        ):
            try:
                # 만약 self.hwnd의 z-order가 HWND_BOTTOM이 아니라면, HWND_BOTTOM으로 이동시킵니다.
                if self.is_hwnd_bottom(self.hwnd):
                    print("Overlay window z-order is already at the bottom.")
                    return

                # self.hwnd도 win32con.HWND_BOTTOM으로 이동시킵니다.
                win32gui.SetWindowPos(
                    sys.argv[2],
                    win32con.HWND_BOTTOM,
                    0,
                    0,
                    0,
                    0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
                )
                win32gui.SetWindowPos(
                    self.hwnd,
                    win32con.HWND_BOTTOM,
                    0,
                    0,
                    0,
                    0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
                )
            except win32gui.error as e:
                logging.error(f"Failed to update z-order: {e}")

    def stop(self):
        self.stop_event.set()
        if self.update_thread:
            self.update_thread.join()
        if self.overlay_hwnd and win32gui.IsWindow(self.overlay_hwnd):
            try:
                win32gui.PostMessage(self.overlay_hwnd, win32con.WM_CLOSE, 0, 0)
            except pywintypes.error as e:
                logging.error(f"Failed to close overlay window: {e}")
        self.overlay_hwnd = None


class AutoMouseTracker:
    def __init__(self, script_path):
        self.script_path = script_path
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.capture_thread = None
        self.dark_mode = False
        self.custom_image_path = None
        self.running_script = False
        self.load_script()
        self.lock = threading.Lock()
        self.script_completed = threading.Event()
        self.active_blockers = []
        self.before_hwnd = None
        self.before_window = None

    def load_script(self):
        try:
            with open(self.script_path, "r", encoding="utf-8") as f:
                self.script = json.load(f)
            logging.info(f"Script loaded from {self.script_path}")
        except Exception as e:
            logging.error(f"Failed to load script: {e}")
            raise

    def play_script(self):
        if not self.script:
            return

        self.running_script = True
        self.progress_bar_value = 0

        def process_event(index):
            if not self.running_script:
                logging.info("Script playback stopped.")
                self.script_completed.set()  # Ensure the event is set here
                return

            if index >= len(self.script):
                self.running_script = False
                self.progress_bar_value = len(self.script)
                logging.info("Script playback completed.")
                self.script_completed.set()  # Ensure the event is set here
                return

            event = self.script[index]

            # 제공된 sys.argv[2]에 가림막을 생성합니다.
            # 만약 제곧된 self.before_hwnd가 더이상 존재하지 않는다면, sys.argv[2]를 사용합니다.

            if self.before_hwnd is not None and win32gui.IsWindow(self.before_hwnd):
                blocker = InputBlocker(self.before_hwnd)
                self.active_blockers.append(blocker)
                blocker_thread = threading.Thread(target=blocker.create_overlay)
                blocker_thread.start()
            else:
                blocker = InputBlocker(sys.argv[2])
                self.active_blockers.append(blocker)
                blocker_thread = threading.Thread(target=blocker.create_overlay)
                blocker_thread.start()

            if event.get("condition") == "이미지 찾을때까지 계속 기다리기":
                while True:
                    try:
                        hwnd = self.find_target_hwnd(event)
                        if hwnd and self.check_image_presence(event, hwnd):
                            logging.info("Image found.")
                            break
                    except RuntimeError:
                        logging.warning("Target hwnd not found, retrying...")
                    time.sleep(1)

            if event.get("condition") in [
                "이미지가 있으면 스킵",
                "이미지가 없으면 스킵",
            ]:
                hwnd = self.find_target_hwnd(event)
                image_present = self.check_image_presence(event, hwnd)

                if (event["condition"] == "이미지가 있으면 스킵" and image_present) or (
                    event["condition"] == "이미지가 없으면 스킵" and not image_present
                ):
                    logging.info("Skipping click based on image presence condition.")
                    time.sleep(0.5)
                    process_event(index + 1)
                    return

            # 대기하는동한 생성한 가림막 제거
            blocker.stop()

            repeat_count = event.get("repeat_count", 1)
            for _ in range(repeat_count):
                if not self.process_single_event(event):
                    break

            self.progress_bar_value = index + 1

            logging.info(f"Processed event {index + 1}/{len(self.script)}")
            process_event(index + 1)

        threading.Thread(target=process_event, args=(0,)).start()

    def process_single_event(self, event):
        hwnd = self.find_target_hwnd(event)
        if hwnd is None or hwnd == 0:
            logging.warning(f"Failed to find hwnd for event: {event}")
            return False

        top_parent = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)

        self.before_hwnd = top_parent

        if win32gui.GetWindowText(top_parent) == "다른 이름으로 저장":
            print("다른 이름으로 저장 창이 열려있습니다.")
            self.block_all_keyboard_input()
        else:
            self.unblock_all_keyboard_input()

        blocker = None

        blocker = InputBlocker(top_parent)
        self.active_blockers.append(blocker)
        blocker_thread = threading.Thread(target=blocker.create_overlay)
        blocker_thread.start()

        self.set_window_to_bottom(top_parent)
        time.sleep(0.5)

        try:
            if not self.check_image_presence(event, hwnd):
                logging.info("No image match found.")
                return False

            click_delay = event.get("click_delay", 0)
            if click_delay > 0:
                time.sleep(click_delay / 1000)

            if event.get("auto_update_target", False):
                self.update_target_position(event, hwnd)

            click_x = event["relative_x"] + event.get("click_offset_x", 0)
            click_y = event["relative_y"] + event.get("click_offset_y", 0)

            self.send_click_event(
                click_x,
                click_y,
                hwnd,
                event["move_cursor"],
                event.get("double_click", False),
                event["button"],
            )

            keyboard_input = event.get("keyboard_input", "")
            if keyboard_input:
                self.send_keyboard_input(keyboard_input, hwnd)

            # 이벤트 처리 후 잠시 대기
            time.sleep(0.5)

            return True
        finally:
            if blocker:
                blocker.stop()
                self.active_blockers.remove(blocker)

    def verify_click(self, event, hwnd):
        time.sleep(1)
        if "verify_image" in event:
            return self.check_image_presence({"image": event["verify_image"]}, hwnd)
        return True

    def is_mouse_button_pressed(self):
        return win32api.GetAsyncKeyState(win32con.VK_LBUTTON) < 0

    def is_mouse_moving(self):
        current_pos = win32api.GetCursorPos()
        if current_pos != self.current:
            self.current = current_pos
            return True
        return False

    def is_keyboard_event_active(self):
        for i in range(8, 256):
            if win32api.GetAsyncKeyState(i):
                return True
        return False

    def find_target_hwnd(self, event):
        try:
            if event.get("window_class") == "ComboBox":
                self.set_ui_position_and_size(
                    event["program_name"], "set_save_window.ini"
                )

            hwnds = self.find_hwnd(
                event["program_name"],
                event["window_class"],
                event["window_name"],
                event["depth"],
                event["program_path"],
                event["window_title"],
                event["window_rect"],
                event.get("ignore_pos_size", False),
            )
            return hwnds[0] if hwnds else None
        except RuntimeError as e:
            logging.warning(f"No valid hwnd found for event: {event}. Exception: {e}")
            return None

    def find_hwnd(
        self,
        program_name,
        window_class=None,
        window_name=None,
        depth=0,
        program_path=None,
        window_title=None,
        window_rect=None,
        ignore_pos_size=False,
    ):
        hwnds = []

        def callback(hwnd, _):
            if not self.is_valid_window(hwnd):
                return True
            current_depth = self.get_window_depth(hwnd)
            if current_depth > depth:
                return True
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if self.is_valid_process(found_pid, program_name, program_path):
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
                if self.matches_window_criteria(
                    class_name,
                    window_text,
                    window_class,
                    window_name,
                    window_title,
                    window_rect,
                    hwnd,
                    ignore_pos_size,
                ):
                    hwnds.append(hwnd)
                hwnds.extend(
                    self.find_all_child_windows(
                        hwnd,
                        window_class,
                        window_name,
                        window_text,
                        current_depth + 1,
                        window_title,
                        window_rect,
                        ignore_pos_size,
                    )
                )
            return True

        win32gui.EnumWindows(callback, None)

        if not hwnds:
            hwnds = self.find_hwnd_top_level(
                program_name, window_class, window_name, program_path, window_title
            )

        if not hwnds:
            raise RuntimeError("No valid hwnd found at the top level or in children")

        return hwnds

    def find_hwnd_top_level(
        self,
        program_name,
        window_class=None,
        window_name=None,
        program_path=None,
        window_title=None,
    ):
        hwnds = []

        def callback(hwnd, _):
            if self.is_valid_window(hwnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if self.is_valid_process(found_pid, program_name, program_path):
                    class_name = win32gui.GetClassName(hwnd)
                    window_text = win32gui.GetWindowText(hwnd)
                    if self.matches_window_criteria(
                        class_name, window_text, window_class, window_name, window_title
                    ):
                        hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(callback, None)
        return hwnds

    def find_all_child_windows(
        self,
        parent_hwnd,
        window_class=None,
        window_name=None,
        window_text=None,
        current_depth=0,
        window_title=None,
        window_rect=None,
        ignore_pos_size=False,
    ):
        child_windows = []

        def enum_child_proc(hwnd, _):
            if self.is_valid_window(hwnd):
                if self.matches_window_criteria(
                    win32gui.GetClassName(hwnd),
                    win32gui.GetWindowText(hwnd),
                    window_class,
                    window_name,
                    window_title,
                    window_rect,
                    hwnd,
                    ignore_pos_size,
                ):
                    child_windows.append(hwnd)
                    child_windows.extend(
                        self.find_all_child_windows(
                            hwnd,
                            window_class,
                            window_name,
                            window_text,
                            current_depth + 1,
                            window_title,
                            window_rect,
                            ignore_pos_size,
                        )
                    )
            return True

        win32gui.EnumChildWindows(parent_hwnd, enum_child_proc, None)
        return child_windows

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

    def set_ui_position_and_size(self, hwnd, ini_file):
        try:
            window_title = "다른 이름으로 저장"
            window_hwnd = win32gui.FindWindow(None, window_title)

            if window_hwnd == 0 or not win32gui.IsWindow(window_hwnd):
                logging.warning(
                    "Window not found or invalid, skipping UI position and size setting."
                )
                return

            # win32con.HWND_BOTTOM으로 설정하여 창을 최하위로 이동시킵니다.
            win32gui.SetWindowPos(
                window_hwnd,
                win32con.HWND_BOTTOM,
                0,
                0,
                0,
                0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
            )

            exe_path = os.path.join(
                os.path.dirname(__file__), "WindowLayoutManager.exe"
            )
            ini_file = os.path.join(os.path.dirname(__file__), ini_file)

            success, _ = self.run_window_layout_manager(
                exe_path, window_title, ini_file
            )
            if success:
                print("Window layout restoration process completed successfully.")
            else:
                print("Window layout restoration failed or timed out.")

            win32gui.SetWindowPos(
                window_hwnd,
                win32con.HWND_TOP,
                58,
                63,
                1018,
                625,
                win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE,
            )
        except Exception as e:
            logging.error(f"Failed to set UI position and size: {e}")

    def send_click_event(
        self, relative_x, relative_y, hwnd, move_cursor, double_click, button
    ):
        if not self.is_valid_window(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return False

        lParam = self.get_lparam(relative_x, relative_y, hwnd)

        try:
            while (
                self.is_mouse_button_pressed()
                or self.is_mouse_moving()
                or self.is_keyboard_event_active()
            ):
                time.sleep(0.1)

            if move_cursor:
                self.move_cursor(relative_y)

            self.simulate_click(button, hwnd, lParam, double_click)
            logging.debug("Click event sent successfully.")
            return True
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")
            return False

    def block_all_keyboard_input(self):
        try:
            for i in range(256):
                keyboard.block_key(i)
            logging.info("All keyboard inputs blocked successfully")
        except Exception as e:
            logging.error(f"Failed to block keyboard inputs: {e}")

    def unblock_all_keyboard_input(self):
        try:
            if keyboard.is_blocked(0):
                keyboard.unblock_all()
                logging.info("All keyboard inputs unblocked successfully")
        except Exception as e:
            logging.error(f"Failed to unblock keyboard inputs: {e}")

    def send_keyboard_input(self, text, hwnd):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        current_dir = os.path.join(current_dir, "temp")
        if not os.path.exists(current_dir):
            os.makedirs(current_dir)

        text = os.path.join(current_dir, text)
        logging.info(f"Keyboard input: {text}")

        try:
            # ctrl + a를 눌러 모든 텍스트를 선택합니다.
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
            # 지우기 키를 눌러 선택된 텍스트를 삭제합니다.
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_BACK, 0)

            for char in text:
                # char를 입력할 때 hwnd를 활성화 하지 않고 입력합니다.
                win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                time.sleep(0.05)
        except Exception as e:
            logging.error(f"Failed to send keyboard input: {e}")

    def check_image_presence(self, event, hwnd):
        target_image_paths = event["image"].get("target_paths", [])
        if not target_image_paths:
            return False

        current_image = self.capture_window_image(hwnd)
        if current_image is None:
            return False

        current_image_gray = cv2.cvtColor(current_image, cv2.COLOR_BGR2GRAY)

        similarity_threshold = event.get("similarity_threshold", 0.6)
        for target_image_info in target_image_paths:
            target_image = cv2.imread(target_image_info["path"], cv2.IMREAD_GRAYSCALE)
            if target_image is None:
                continue

            result = cv2.matchTemplate(
                current_image_gray, target_image, cv2.TM_CCOEFF_NORMED
            )
            _, max_val, _, _ = cv2.minMaxLoc(result)

            if max_val >= similarity_threshold:
                return True

        return False

    def capture_window_image(self, hwnd):
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width, height = right - left, bottom - top

            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

            if result == 0:
                windll.gdi32.BitBlt(
                    saveDC.GetSafeHdc(),
                    0,
                    0,
                    width,
                    height,
                    mfcDC.GetSafeHdc(),
                    0,
                    0,
                    win32con.SRCCOPY,
                )

            signedIntsArray = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(signedIntsArray, dtype="uint8")
            img.shape = (height, width, 4)

            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)

            return img
        except Exception as e:
            logging.error(f"Error capturing window image: {e}")
            return None

    @staticmethod
    def simulate_click(button, hwnd, lParam, double_click):
        if button == "left":
            AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
            time.sleep(0.1)
            AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)
            if double_click:
                time.sleep(0.1)
                AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
                time.sleep(0.1)
                AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)
        elif button == "right":
            AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_RBUTTONDOWN)
            time.sleep(0.1)
            AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_RBUTTONUP)
            if double_click:
                time.sleep(0.1)
                AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_RBUTTONDOWN)
                time.sleep(0.1)
                AutoMouseTracker.simulate_mouse_event(hwnd, lParam, WM_RBUTTONUP)

    @staticmethod
    def simulate_mouse_event(hwnd, lParam, event_type):
        if event_type in [WM_LBUTTONDOWN, WM_RBUTTONDOWN]:
            win32gui.PostMessage(hwnd, event_type, win32con.MK_LBUTTON, lParam)
        else:
            win32gui.PostMessage(hwnd, event_type, 0, lParam)

        win32gui.PumpWaitingMessages()

    def get_window_depth(self, hwnd):
        depth = 0
        while hwnd != 0:
            hwnd = win32gui.GetParent(hwnd)
            depth += 1
        return depth

    def is_valid_process(self, pid, program_name, program_path=None):
        try:
            process = psutil.Process(pid)
            if program_path:
                return (
                    process.exe().lower().split("/")[-1]
                    == program_path.lower().split("/")[-1]
                )

            return process.name().lower() == program_name.lower()
        except psutil.NoSuchProcess:
            return False

    def is_valid_window(self, hwnd):
        return (
            win32gui.IsWindow(hwnd)
            and win32gui.IsWindowEnabled(hwnd)
            and win32gui.IsWindowVisible(hwnd)
        )

    def matches_window_criteria(
        self,
        class_name,
        window_text,
        window_class,
        window_name,
        window_title,
        window_rect=None,
        hwnd=None,
        ignore_pos_size=False,
    ):
        class_match = window_class in class_name if window_class else True
        name_match = window_name == window_text if window_name else True
        title_match = window_title == window_text if window_title else True
        rect_match = True

        if window_rect and not ignore_pos_size and hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            tolerance = 10  # 오차 범위 ±10 픽셀

            rect_match = (
                abs(window_rect[0] - rect[0]) <= tolerance
                and abs(window_rect[1] - rect[1]) <= tolerance
                and abs(window_rect[2] - rect[2]) <= tolerance
                and abs(window_rect[3] - rect[3]) <= tolerance
            )

        return class_match and name_match and title_match and rect_match

    def check_conditions(self, event, hwnd):
        if event.get("condition") == "Image Present":
            if not self.check_image_presence(event, hwnd):
                logging.info("Condition 'Image Present' not met, skipping event.")
                return False
        elif event.get("condition") == "Image Not Present":
            if self.check_image_presence(event, hwnd):
                logging.info("Condition 'Image Not Present' not met, skipping event.")
                return False
        return True

    def update_target_position(self, event, hwnd):
        img = self.capture_window_image(hwnd)
        if img is not None:
            for image_path in event.get("image_paths", []):
                target_image = cv2.imread(image_path, cv2.IMREAD_COLOR)
                if target_image is not None:
                    result = cv2.matchTemplate(img, target_image, cv2.TM_CCOEFF_NORMED)
                    _, max_val, _, max_loc = cv2.minMaxLoc(result)
                    if max_val >= event.get("similarity_threshold", 0.6):
                        x, y, width, height = (
                            max_loc[0],
                            max_loc[1],
                            target_image.shape[1],
                            target_image.shape[0],
                        )
                        event["relative_x"] = x + width // 2
                        event["relative_y"] = y + height // 2
                        break

    def move_cursor(self, current_y):
        current_x, current_y = win32api.GetCursorPos()
        screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        new_y = current_y + 20 if current_y + 25 < screen_height else current_y - 20
        win32api.SetCursorPos((current_x, new_y))
        time.sleep(0.1)

    def get_lparam(self, relative_x, relative_y, hwnd):
        lParam = win32api.MAKELONG(relative_x, relative_y)
        if win32gui.GetClassName(hwnd) == "#32768":
            left, top, _, _ = win32gui.GetWindowRect(hwnd)
            screen_x, screen_y = left + relative_x, top + relative_y
            time.sleep(0.5)
            lParam = win32api.MAKELONG(screen_x, screen_y)
        return lParam

    def set_window_to_bottom(self, top_parent):
        try:
            # Move all related windows to the bottom
            self.set_window_and_children_to_bottom(top_parent)

            logging.info(f"Window {top_parent} and all related windows moved to bottom")
            return True
        except Exception as e:
            logging.error(f"Failed to set windows to bottom: {e}")
            return False

    def set_window_and_children_to_bottom(self, hwnd):
        # Move the current window to the bottom
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_BOTTOM,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
        )

        # Recursively set all child windows to bottom
        def enum_child_windows(child_hwnd, _):
            self.set_window_and_children_to_bottom(child_hwnd)
            return True

        win32gui.EnumChildWindows(hwnd, enum_child_windows, None)

    def wait_for_completion(self):
        self.script_completed.wait()

    # Cleanup function
    def cleanup(self):
        for blocker in self.active_blockers:
            blocker.stop()
        self.active_blockers.clear()
        if self.capture_thread:
            self.capture_thread.join()
        logging.shutdown()
        self.script_completed.set()  # Ensure the event is set to allow exit


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python AutoMouseTracker.py <path_to_script.json>")
        sys.exit(1)

    script_path = sys.argv[1]
    tracker = AutoMouseTracker(script_path)
    tracker.play_script()

    # Main program execution at the end
    tracker.wait_for_completion()
    tracker.cleanup()
    print("Script execution completed. Exiting program.")
    time.sleep(1)
    os._exit(0)  # Force exit if sys.exit does not work

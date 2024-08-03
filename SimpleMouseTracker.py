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

# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
LOG_FILE = "mouse_tracker.log"
SAMPLE_TARGETS_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "StructFlow-Automator-Private",
    "sample_targets",
)

# Maximum retry count for click events
MAX_RETRY_COUNT = 5


def configure_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[RotatingFileHandler(LOG_FILE, maxBytes=10**6, backupCount=3)],
    )


configure_logging()


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
                return

            if index >= len(self.script):
                self.running_script = False
                self.progress_bar_value = len(self.script)
                logging.info("Script playback completed.")
                return

            event = self.script[index]

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

            repeat_count = event.get("repeat_count", 1)
            for _ in range(repeat_count):
                if not self.process_single_event(event):
                    break

            self.progress_bar_value = index + 1

            logging.info(f"Processed event {index + 1}/{len(self.script)}")
            time.sleep(0.5)
            process_event(index + 1)

        threading.Thread(target=process_event, args=(0,)).start()

    def process_single_event(self, event):
        hwnd = self.find_target_hwnd(event)
        if hwnd is None:
            logging.warning(f"Failed to find hwnd for event: {event}")
            return False

        if not self.check_conditions(event, hwnd):
            return False

        click_delay = event.get("click_delay", 0)
        if click_delay > 0:
            time.sleep(click_delay / 1000)

        if event.get("auto_update_target", False):
            self.update_target_position(event, hwnd)

        success = False
        attempt = 0
        while not success and attempt < MAX_RETRY_COUNT:
            success = self.send_click_event(
                event["relative_x"],
                event["relative_y"],
                hwnd,
                event["move_cursor"],
                event.get("double_click", False),
                event["button"],
            )
            if not success:
                logging.warning(f"Click event failed, retrying... (Attempt {attempt + 1})")
                attempt += 1
                time.sleep(1)  # Wait before retrying

        if not success:
            logging.error("Click event failed after maximum retries.")
            return False

        keyboard_input = event.get("keyboard_input", "")
        if keyboard_input:
            self.send_keyboard_input(keyboard_input, hwnd)

        return True

    def find_target_hwnd(self, event):
        try:
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

    def send_click_event(
        self, relative_x, relative_y, hwnd, move_cursor, double_click, button
    ):
        if not self.is_valid_window(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return False

        if move_cursor:
            self.move_cursor(relative_y)

        lParam = self.get_lparam(relative_x, relative_y, hwnd)

        try:
            self.simulate_click(button, hwnd, lParam, double_click)
            logging.debug("Click event sent successfully.")
            return True
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")
            return False

    def send_keyboard_input(self, text, hwnd):
        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.05)

    def check_image_presence(self, event, hwnd):
        target_image_paths = event["image"].get("target_paths", [])
        if not target_image_paths:
            return False

        current_image = self.capture_window_image(hwnd)
        if current_image is None:
            return False

        similarity_threshold = event.get("similarity_threshold", 0.6)
        for target_image_info in target_image_paths:
            target_image = cv2.imread(target_image_info["path"], cv2.IMREAD_COLOR)
            if target_image is None:
                continue

            result = cv2.matchTemplate(
                current_image, target_image, cv2.TM_CCOEFF_NORMED
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
                return process.exe().lower() == program_path.lower()
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
        if window_rect and not ignore_pos_size:
            rect = win32gui.GetWindowRect(hwnd)
            rect_match = (
                window_rect[0] == rect[0]
                and window_rect[1] == rect[1]
                and window_rect[2] == rect[2]
                and window_rect[3] == rect[3]
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
                        event["relative_x"], event["relative_y"] = max_loc
                        break

    def move_cursor(self, current_y):
        current_x, current_y = win32api.GetCursorPos()
        screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        move_direction = 10 if current_y == 0 else -10
        new_y = max(0, min(screen_height - 1, current_y + move_direction))
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


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python AutoMouseTracker.py <path_to_script.json>")
        sys.exit(1)

    script_path = sys.argv[1]
    tracker = AutoMouseTracker(script_path)
    tracker.play_script()

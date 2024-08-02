# SimpleMouseTracker.py

import sys
import os
import time
import json
import logging
import numpy as np
import cv2
from logging.handlers import RotatingFileHandler
from ctypes import windll
import psutil
import win32gui
import win32process
import win32api
import win32con
import win32ui

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

def configure_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[RotatingFileHandler(LOG_FILE, maxBytes=10**6, backupCount=3)],
    )

configure_logging()

class ImageCapture:
    def __init__(self, hwnd):
        self.hwnd = hwnd

    def capture_window_image(self):
        try:
            left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
            width = right - left
            height = bottom - top

            hwndDC = win32gui.GetWindowDC(self.hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            result = windll.user32.PrintWindow(self.hwnd, saveDC.GetSafeHdc(), 0)

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
            win32gui.ReleaseDC(self.hwnd, hwndDC)

            return img
        except Exception as e:
            logging.error(f"Error capturing window image: {e}")
            return None

class MouseTracker:
    def __init__(self, script_path):
        self.script_path = script_path
        self.click_events = self.load_script()
        self.current_program_hwnd = win32gui.GetForegroundWindow()

    def load_script(self):
        with open(self.script_path, "r") as file:
            return json.load(file)

    def play_script(self):
        if not self.click_events:
            return

        for index, event in enumerate(self.click_events):
            self.process_single_event(event)
            time.sleep(1)  # Adding delay between events

    def process_single_event(self, event):
        condition = event.get("condition")
        if condition == "이미지 찾을때까지 계속 기다리기":
            hwnd = None
            timeout = 5 * 60
            start_time = time.time()
            while time.time() - start_time < timeout:
                hwnd = self.find_target_hwnd(event)
                if hwnd and self.check_image_presence(event, hwnd):
                    logging.info("Image found, proceeding with click.")
                    break
                logging.info("Waiting for image to appear...")
                time.sleep(1)
        else:
            hwnd = self.find_target_hwnd(event)

        if not hwnd or hwnd == 0:
            logging.warning(f"Failed to find hwnd for event: {event}")
            return

        if condition == "이미지가 있으면 스킵":
            if self.check_image_presence(event, hwnd):
                logging.info("Image found, skipping click.")
                return

        elif condition == "이미지가 없으면 스킵":
            if not self.check_image_presence(event, hwnd):
                logging.info("Image not found, skipping click.")
                return

        click_delay = event.get("click_delay", 0)
        if click_delay > 0:
            time.sleep(click_delay / 1000)

        if event.get("auto_update_target", False):
            capture = ImageCapture(hwnd)
            img = capture.capture_window_image()
            if img is not None:
                for image_path in event.get("image_paths", []):
                    target_image = cv2.imread(image_path, cv2.IMREAD_COLOR)
                    if target_image is not None:
                        result = cv2.matchTemplate(
                            img, target_image, cv2.TM_CCOEFF_NORMED
                        )
                        _, max_val, _, max_loc = cv2.minMaxLoc(result)
                        if max_val >= event.get("similarity_threshold", 0.6):
                            event["relative_x"], event["relative_y"] = max_loc
                            break

        self.send_click_event(
            event["relative_x"],
            event["relative_y"],
            hwnd,
            event["move_cursor"],
            event.get("double_click", False),
            event["button"],
        )

        keyboard_input = event.get("keyboard_input", "")
        if keyboard_input:
            self.send_keyboard_input(keyboard_input, hwnd)

    def send_click_event(
        self, relative_x, relative_y, hwnd, move_cursor, double_click, button
    ):
        if not hwnd or not win32gui.IsWindow(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return

        if move_cursor:
            current_x, current_y = win32api.GetCursorPos()
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

            move_direction = 10 if current_y == 0 else -10
            new_y = max(0, min(screen_height - 1, current_y + move_direction))
            win32api.SetCursorPos((current_x, new_y))
            time.sleep(0.1)

        lParam = win32api.MAKELONG(relative_x, relative_y)

        if win32gui.GetClassName(hwnd) == "#32768":
            left, top, _, _ = win32gui.GetWindowRect(hwnd)
            screen_x, screen_y = left + relative_x, top + relative_y

            time.sleep(0.5)

            lParam = win32api.MAKELONG(screen_x, screen_y)

        try:
            print(hwnd, relative_x, relative_y, button)
            if button == "left":
                self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
                time.sleep(0.1)
                self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)

                if double_click:
                    time.sleep(0.1)
                    self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
                    time.sleep(0.1)
                    self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)
            elif button == "right":
                self.simulate_mouse_event(hwnd, lParam, WM_RBUTTONDOWN)
                time.sleep(0.1)
                self.simulate_mouse_event(hwnd, lParam, WM_RBUTTONUP)

                if double_click:
                    time.sleep(0.1)
                    self.simulate_mouse_event(hwnd, lParam, WM_RBUTTONDOWN)
                    time.sleep(0.1)
                    self.simulate_mouse_event(hwnd, lParam, WM_RBUTTONUP)

            logging.debug("Click event sent successfully.")
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")

    def send_keyboard_input(self, text, hwnd):
        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.05)

    def check_image_presence(self, event, hwnd):
        target_image_paths = event.get("image_paths", [])
        if not target_image_paths:
            return False

        current_image = ImageCapture(hwnd).capture_window_image()
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
            if hwnds:
                return hwnds[0]
        except RuntimeError as e:
            logging.warning(f"No valid hwnd found for event: {event}. Exception: {e}")
            return None

        logging.warning(f"No valid hwnd found for event: {event}")
        return None

    def is_valid_process(self, pid, program_name, program_path=None):
        try:
            process = psutil.Process(pid)
            if program_path:
                return process.exe().lower() == program_path.lower()
            return process.name().lower() == program_name.lower()
        except psutil.NoSuchProcess:
            return False

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
            if (
                not win32gui.IsWindow(hwnd)
                or not win32gui.IsWindowEnabled(hwnd)
                or not win32gui.IsWindowVisible(hwnd)
            ):
                return True
            current_depth = self.get_window_depth(hwnd)
            if current_depth > depth:
                return True
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if self.is_valid_process(found_pid, program_name, program_path):
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
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
                if (
                    class_match
                    and name_match
                    and title_match
                    and current_depth == depth
                    and rect_match
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
            if not win32gui.IsWindowEnabled(hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if self.is_valid_process(found_pid, program_name, program_path):
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
                class_match = window_class in class_name if window_class else True
                name_match = window_name == window_text if window_name else True
                title_match = window_title == window_text if window_title else True
                if class_match and name_match and title_match:
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
            if (
                not win32gui.IsWindow(hwnd)
                or not win32gui.IsWindowEnabled(hwnd)
                or not win32gui.IsWindowVisible(hwnd)
            ):
                return True
            match_class = (
                window_class in win32gui.GetClassName(hwnd) if window_class else True
            )
            match_text = (
                window_text == win32gui.GetWindowText(hwnd) if window_text else True
            )
            match_name = (
                window_name == win32gui.GetWindowText(hwnd) if window_name else True
            )
            match_title = (
                window_title == win32gui.GetWindowText(hwnd) if window_title else True
            )
            rect_match = True
            if window_rect and not ignore_pos_size:
                rect = win32gui.GetWindowRect(hwnd)
                rect_match = (
                    window_rect[0] == rect[0]
                    and window_rect[1] == rect[1]
                    and window_rect[2] == rect[2]
                    and window_rect[3] == rect[3]
                )
            if match_class and (match_text or match_name or match_title) and rect_match:
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

    @staticmethod
    def simulate_mouse_event(hwnd, lParam, event_type):
        if event_type == win32con.WM_LBUTTONDOWN:
            win32gui.PostMessage(hwnd, event_type, win32con.MK_LBUTTON, lParam)
        elif event_type == win32con.WM_RBUTTONDOWN:
            win32gui.PostMessage(hwnd, event_type, win32con.MK_RBUTTON, lParam)
        else:
            win32gui.PostMessage(hwnd, event_type, 0, lParam)

    def get_window_depth(self, hwnd):
        depth = 0
        while hwnd != 0:
            hwnd = win32gui.GetParent(hwnd)
            depth += 1
        return depth

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python.exe .\\SimpleMouseTracker.py .\\calcurate.json")
        sys.exit(1)

    script_path = sys.argv[1]
    tracker = MouseTracker(script_path)
    tracker.play_script()

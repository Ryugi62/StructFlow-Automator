import sys
import json
import time
import logging
import ctypes
from ctypes import wintypes
import numpy as np
import cv2
import psutil
import win32gui
import win32process
import win32api
import win32con
import win32ui
from PyQt5.QtWidgets import (
    QApplication,
    QLabel,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QFileDialog,
    QSlider,
    QMessageBox,
    QListWidget,
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPixmap, QImage
from pynput import keyboard, mouse
import logging.handlers
import os

# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
LOG_FILE = "mouse_tracker.log"

# Logging configuration 수정
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=10**6, backupCount=3)
    ],
)

user32 = ctypes.windll.user32


class ImageCaptureThread(QThread):
    image_captured = pyqtSignal(np.ndarray)
    condition_met = pyqtSignal()

    def __init__(self, hwnd, exclude_hwnd):
        super().__init__()
        self.hwnd = hwnd
        self.exclude_hwnd = exclude_hwnd
        self.running = True
        self.condition = None
        self.start_time = time.time()

    def run(self):
        while self.running:
            if time.time() - self.start_time > 300:
                logging.error("Timeout error: Search exceeded 5 minutes.")
                self.running = False
                return
            if self.hwnd == self.exclude_hwnd:
                time.sleep(0.1)
                continue
            try:
                img = self.capture_window_image(self.hwnd)
                if img is not None:
                    self.image_captured.emit(img)
                    if self.condition and self.condition(img):
                        self.condition_met.emit()
                        self.running = False
                        return
                time.sleep(1)
            except Exception as e:
                logging.error(f"Error in ImageCaptureThread: {e}")

    def stop(self):
        self.running = False
        self.wait()

    def capture_window_image(self, hwnd):
        try:
            window_rect = win32gui.GetWindowRect(hwnd)
            width, height = (
                window_rect[2] - window_rect[0],
                window_rect[3] - window_rect[1],
            )
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype="uint8").reshape((height, width, 4))
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)

            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            return img
        except Exception as e:
            logging.error(f"Error capturing window image: {e}")
            return None

    def set_condition(self, condition_func):
        self.condition = condition_func


class MouseTracker(QWidget):

    def __init__(self):
        super().__init__()
        self.init_variables()
        self.init_ui()
        self.init_listeners()
        self.init_capture_thread()

    def init_ui(self):
        self.layout = QVBoxLayout()
        self.program_label = QLabel("Current Program: None")
        self.program_label.setWordWrap(True)
        self.target_label = QLabel("Target Window: None")
        self.target_label.setWordWrap(True)
        self.current_target_label = QLabel("Current Target Relative Position: (0, 0)")
        self.image_label = QLabel()
        self.image_label.setMaximumSize(640, 480)
        self.image_label.setStyleSheet("background-color: white;")

        self.event_list = QListWidget()

        self.layout.addWidget(self.program_label)
        self.layout.addWidget(self.target_label)
        self.layout.addWidget(self.current_target_label)
        self.layout.addWidget(self.image_label)
        self.layout.addWidget(self.event_list)

        self.record_button = QPushButton("Start Recording")
        self.record_button.clicked.connect(self.toggle_recording)
        self.layout.addWidget(self.record_button)

        self.save_button = QPushButton("Save Script")
        self.save_button.clicked.connect(self.save_script)
        self.layout.addWidget(self.save_button)

        self.load_button = QPushButton("Load Script")
        self.load_button.clicked.connect(self.load_script)
        self.layout.addWidget(self.load_button)

        self.play_button = QPushButton("Play Script")
        self.play_button.clicked.connect(self.play_script)
        self.layout.addWidget(self.play_button)

        self.speed_slider = QSlider(Qt.Horizontal)
        self.speed_slider.setRange(1, 100)
        self.speed_slider.setValue(50)
        self.speed_slider.valueChanged.connect(self.update_speed_factor)
        self.layout.addWidget(self.speed_slider)

        self.setLayout(self.layout)
        self.setWindowTitle("Mouse Tracker")
        self.setGeometry(100, 100, 800, 600)
        self.show()

    def init_variables(self):
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.capture_thread = None

    def init_listeners(self):
        self.mouse_listener = mouse.Listener(
            on_move=self.on_move, on_click=self.on_click
        )
        self.mouse_listener.start()
        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.keyboard_listener.start()
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_program_label)
        self.timer.start(1000)

    def init_capture_thread(self):
        self.capture_thread = ImageCaptureThread(
            self.current_program_hwnd, self.winId()
        )
        self.capture_thread.image_captured.connect(self.update_image_label)
        self.capture_thread.condition_met.connect(self.on_condition_met)
        self.capture_thread.start()

    def on_move(self, x, y):
        self.current = (x, y)
        self.update_current_target_label()

    def on_click(self, x, y, button, pressed):
        if pressed:
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd:
                self.print_window_hierarchy(hwnd)
            if pressed and self.recording and not self.is_own_window(hwnd):
                self.record_click_event(x, y, hwnd)

    def is_own_window(self, hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        return window_title == "Mouse Tracker"

    def on_press(self, key):
        if key == keyboard.Key.f9:
            self.toggle_recording()
        elif key == keyboard.Key.f10:
            self.play_script()

    def record_click_event(self, x, y, hwnd):
        if hwnd == self.current_program_hwnd or self.is_own_window(hwnd):
            return
        window_rect = win32gui.GetWindowRect(hwnd)
        relative_x, relative_y = x - window_rect[0], y - window_rect[1]
        self.update_program_label(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        current_program, program_path = self.get_program_info(pid)
        window_name, window_class = win32gui.GetWindowText(hwnd), win32gui.GetClassName(
            hwnd
        )
        depth = self.get_window_depth(hwnd)

        img = self.capture_thread.capture_window_image(hwnd)
        timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
        image_filename = f"{current_program}_{timestamp}_{relative_x}_{relative_y}.png"
        full_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "StructFlow-Automator-Private",
            "sample_targets",
            image_filename,
        )
        cv2.imwrite(full_path, cv2.cvtColor(img, cv2.COLOR_RGB2BGR))

        event_info = {
            "relative_x": relative_x,
            "relative_y": relative_y,
            "program_name": current_program,
            "program_path": program_path,
            "window_name": window_name,
            "window_class": window_class,
            "depth": depth,
            "time": time.time() - self.start_time,
            "image": {
                "path": full_path,
                "image_program_name": current_program,
                "image_program_path": program_path,
                "image_window_class": window_class,
                "image_window_name": window_name,
                "image_depth": 0,
                "wait_for_image": False,
                "wait_method": "time",  # 추가: time 또는 image
            },
        }

        self.update_image_label(img)
        self.capture_thread.hwnd = hwnd
        logging.info(f"Recording click event: {event_info}")
        self.click_events.append(event_info)
        self.event_list.addItem(
            f"Click at ({relative_x}, {relative_y}) in {current_program}"
        )

    def save_image(self, img, event_info):
        timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
        filename = f"{event_info['program_name']}_{timestamp}_{event_info['relative_x']}_{event_info['relative_y']}.png"
        full_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "StructFlow-Automator-Private",
            "sample_targets",
            filename,
        )
        cv2.imwrite(full_path, cv2.cvtColor(img, cv2.COLOR_RGB2BGR))
        event_info["image"]["path"] = full_path

    def get_window_depth(self, hwnd):
        depth = 0
        while hwnd != 0:
            hwnd = win32gui.GetParent(hwnd)
            depth += 1
        return depth

    def get_program_info(self, pid):
        for proc in psutil.process_iter(["pid", "name", "exe"]):
            if proc.info["pid"] == pid:
                return proc.info["name"], proc.info["exe"]
        return "Unknown", "Unknown"

    @pyqtSlot(np.ndarray)
    def update_image_label(self, img):
        height, width = img.shape[:2]
        max_width, max_height = 640, 480
        if width > max_width or height > max_height:
            aspect_ratio = width / height
            if width > height:
                new_width, new_height = max_width, int(max_width / aspect_ratio)
            else:
                new_height, new_width = max_height, int(max_height * aspect_ratio)
            img = cv2.resize(img, (new_width, new_height))
        qimg = QImage(
            img.data, img.shape[1], img.shape[0], img.strides[0], QImage.Format_RGB888
        )
        self.image_label.setPixmap(QPixmap.fromImage(qimg))

    def update_program_label(self, hwnd=None):
        if hwnd is None:
            hwnd = win32gui.WindowFromPoint(self.current)
        if hwnd == 0 or not win32gui.IsWindow(hwnd):
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        current_program, program_path = self.get_program_info(pid)
        window_name, window_class = win32gui.GetWindowText(hwnd), win32gui.GetClassName(
            hwnd
        )

        self.program_label.setText(
            f"Current Program: {current_program} ({program_path})"
        )
        self.target_label.setText(
            f"Target Window: {window_name} (hwnd: {hwnd}, class: {window_class})"
        )

    def update_current_target_label(self):
        hwnd = win32gui.WindowFromPoint(self.current)
        if hwnd == self.current_program_hwnd or hwnd == 0:
            return
        window_rect = win32gui.GetWindowRect(hwnd)
        relative_x, relative_y = (
            self.current[0] - window_rect[0],
            self.current[1] - window_rect[1],
        )
        self.current_target_label.setText(
            f"Current Target Relative Position: ({relative_x}, {relative_y})"
        )

    def toggle_recording(self):
        self.recording = not self.recording
        self.record_button.setText(
            "Stop Recording" if self.recording else "Start Recording"
        )
        if self.recording:
            self.start_time = time.time()
            self.click_events = []
            self.event_list.clear()

    def save_script(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Script", "", "JSON Files (*.json);;All Files (*)"
        )
        if filename:
            with open(filename, "w") as file:
                json.dump(self.click_events, file)

    def load_script(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Load Script", "", "JSON Files (*.json);;All Files (*)"
        )
        if filename:
            with open(filename, "r") as file:
                self.click_events = json.load(file)
            self.event_list.clear()
            for event in self.click_events:
                self.event_list.addItem(
                    f"Click at ({event['relative_x']}, {event['relative_y']}) in {event['program_name']}"
                )

    def play_script(self):
        if not self.click_events:
            return

        self.set_buttons_enabled(False)

        try:
            start_time = self.click_events[0]["time"]
            last_event_time = start_time

            for i, event in enumerate(self.click_events):
                wait_method = event["image"].get("wait_method", "time")
                full_image_path = event["image"]["path"]
                if wait_method == "image":
                    self.wait_for_image(full_image_path)
                else:
                    elapsed_time = (event["time"] - last_event_time) * self.speed_factor
                    if elapsed_time > 0:
                        time.sleep(elapsed_time)
                last_event_time = event["time"]

                hwnd = self.find_target_hwnd(event)
                if hwnd is None or hwnd == 0:
                    logging.warning(f"Invalid hwnd for event: {event}")
                    continue

                self.update_image_label(self.capture_thread.capture_window_image(hwnd))
                self.capture_thread.hwnd = hwnd
                self.event_list.setCurrentRow(i)
                self.send_click_event(event["relative_x"], event["relative_y"], hwnd)
                QApplication.processEvents()

        except RuntimeError as e:
            logging.error(f"Error playing script: {e}")
            self.show_error_message(str(e))
        finally:
            self.set_buttons_enabled(True)

    def wait_for_image(self, image_path):
        if not os.path.exists(image_path):
            logging.error(f"Image file not found: {image_path}")
            return

        target_image = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if target_image is None:
            logging.error(f"Failed to load image: {image_path}")
            return

        while True:
            img = self.capture_thread.capture_window_image(self.current_program_hwnd)
            if img is None:
                time.sleep(1)
                continue

            # 이미지를 같은 크기로 변환하여 비교
            if img.shape[:2] != target_image.shape[:2]:
                target_image_resized = cv2.resize(target_image, (img.shape[1], img.shape[0]))
            else:
                target_image_resized = target_image

            result = cv2.matchTemplate(img, target_image_resized, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            if max_val > 0.9:  # 유사도가 90% 이상이면 조건 만족으로 판단
                break
            time.sleep(1)

    def wait_for_image(self, image_path):
        if not os.path.exists(image_path):
            logging.error(f"Image file not found: {image_path}")
            return

        target_image = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if target_image is None:
            logging.error(f"Failed to load image: {image_path}")
            return

        while True:
            img = self.capture_thread.capture_window_image(self.current_program_hwnd)
            if img is None:
                time.sleep(1)
                continue

            if img.shape[:2] != target_image.shape[:2]:
                target_image_resized = cv2.resize(
                    target_image, (img.shape[1], img.shape[0])
                )
            else:
                target_image_resized = target_image

            result = cv2.matchTemplate(img, target_image_resized, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            if max_val > 0.8:
                break
            time.sleep(1)

    def set_buttons_enabled(self, enabled):
        self.record_button.setEnabled(enabled)
        self.save_button.setEnabled(enabled)
        self.load_button.setEnabled(enabled)
        self.play_button.setEnabled(enabled)

    def find_target_hwnd(self, event):
        hwnds = self.find_hwnd(
            event["program_name"],
            event["window_class"],
            event["window_name"],
            event["depth"],
            event["program_path"],
        )
        if hwnds:
            return hwnds[0]

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
    ):
        hwnds = []

        def callback(hwnd, _):
            if not win32gui.IsWindow(hwnd):
                return True
            current_depth = self.get_window_depth(hwnd)
            if current_depth > depth:
                return True
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if self.is_valid_process(found_pid, program_name, program_path):
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
                class_match = window_class in class_name if window_class else True
                name_match = window_name in window_text if window_name else True
                if class_match and name_match and current_depth == depth:
                    hwnds.append(hwnd)
                hwnds.extend(
                    self.find_all_child_windows(
                        hwnd, window_class, window_name, current_depth + 1
                    )
                )
            return True

        win32gui.EnumWindows(callback, None)

        if not hwnds:
            hwnds = self.find_hwnd_top_level(
                program_name, window_class, window_name, program_path
            )

        if not hwnds:
            raise RuntimeError("No valid hwnd found at the top level or in children")

        return hwnds

    def find_hwnd_top_level(
        self, program_name, window_class=None, window_name=None, program_path=None
    ):
        hwnds = []

        def callback(hwnd, _):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if self.is_valid_process(found_pid, program_name, program_path):
                class_name = win32gui.GetClassName(hwnd)
                window_text = win32gui.GetWindowText(hwnd)
                class_match = window_class in class_name if window_class else True
                name_match = window_name in window_text if window_name else True
                if class_match and name_match:
                    hwnds.append(hwnd)
            return True

        win32gui.EnumWindows(callback, None)
        return hwnds

    def find_all_child_windows(
        self, parent_hwnd, window_class=None, window_text=None, current_depth=0
    ):
        child_windows = []

        def enum_child_proc(hwnd, _):
            if not win32gui.IsWindow(hwnd):
                return True
            match_class = (
                window_class in win32gui.GetClassName(hwnd) if window_class else True
            )
            match_text = (
                window_text in win32gui.GetWindowText(hwnd) if window_text else True
            )
            if match_class and match_text:
                child_windows.append(hwnd)
                child_windows.extend(
                    self.find_all_child_windows(
                        hwnd, window_class, window_text, current_depth + 1
                    )
                )
            return True

        win32gui.EnumChildWindows(parent_hwnd, enum_child_proc, None)
        return child_windows

    def send_click_event(self, relative_x, relative_y, hwnd):
        if not hwnd or not win32gui.IsWindow(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return

        lParam = win32api.MAKELONG(relative_x, relative_y)

        try:
            self.simulate_mouse_event(hwnd, lParam, WM_MOUSEMOVE)
            self.activate_window(hwnd)
            time.sleep(0.1)

            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
            time.sleep(0.05)

            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)
            time.sleep(0.05)

            logging.debug("Click event sent successfully.")
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")

    def activate_window(self, hwnd):
        try:
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_CLICKACTIVE, 0)
            time.sleep(0.05)
        except Exception as e:
            logging.error(f"Failed to activate window: {e}")

    @staticmethod
    def simulate_mouse_event(hwnd, lParam, event_type):
        win32api.PostMessage(hwnd, event_type, 0, lParam)

    def update_speed_factor(self, value):
        self.speed_factor = max(0.1, value / 50.0)

    def closeEvent(self, event):
        if self.capture_thread is not None:
            self.capture_thread.stop()
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
        super().closeEvent(event)

    @pyqtSlot(str)
    def show_error_message(self, message):
        QMessageBox.critical(self, "Error", message)

    def print_window_hierarchy(self, hwnd):
        hierarchy = self.find_window_hierarchy(hwnd)
        self.display_window_hierarchy(hierarchy)

    def find_window_hierarchy(self, hwnd):
        hierarchy = []
        current_hwnd = hwnd

        while current_hwnd:
            window_title = win32gui.GetWindowText(current_hwnd)
            window_class = win32gui.GetClassName(current_hwnd)
            hierarchy.append((window_title, window_class, current_hwnd))
            current_hwnd = win32gui.GetParent(current_hwnd)

        return hierarchy[::-1]

    def display_window_hierarchy(self, hierarchy):
        indent = "   "
        for i, (title, class_name, hwnd) in enumerate(hierarchy):
            print(
                f"{indent * i}Window Title: {title}, Window Class: {class_name}, HWND: {hwnd}"
            )

    def start_search_for_condition(self):
        def condition(img):
            target_image = cv2.imread("path_to_target_image.png")  # 목표 이미지 경로
            result = cv2.matchTemplate(img, target_image, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            return max_val > 0.8  # 유사도가 0.8 이상이면 조건 만족으로 판단

        self.capture_thread.set_condition(condition)
        self.capture_thread.start()

    def on_condition_met(self):
        logging.info("Condition met, stopping search.")
        self.handle_timeout_error()

    def handle_timeout_error(self):
        logging.error("Timeout occurred during image search.")
        QMessageBox.critical(self, "Error", "Timeout occurred during image search.")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        ex = MouseTracker()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Error in main: {e}")

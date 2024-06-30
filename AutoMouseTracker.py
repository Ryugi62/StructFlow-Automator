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
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPixmap, QImage
from pynput import keyboard, mouse

# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202

# Logging configuration
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
LOG_FILE = "mouse_tracker.log"

logging.basicConfig(
    level=LOG_LEVEL,
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler(LOG_FILE),
    ],
)

user32 = ctypes.windll.user32


class ImageCaptureThread(QThread):
    image_captured = pyqtSignal(np.ndarray)

    def __init__(self, hwnd, exclude_hwnd):
        super().__init__()
        self.hwnd = hwnd
        self.exclude_hwnd = exclude_hwnd
        self.running = True

    def run(self):
        while self.running:
            if self.hwnd == self.exclude_hwnd:
                time.sleep(0.1)
                continue
            try:
                img = self.capture_window_image(self.hwnd)
                if img is not None:
                    self.image_captured.emit(img)
                time.sleep(0.5)
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


class MouseTracker(QWidget):

    def __init__(self):
        super().__init__()
        self.initVariables()
        self.initUI()
        self.initListeners()

    def initUI(self):
        self.layout = QVBoxLayout()
        self.program_label = QLabel("Current Program: None")
        self.program_label.setWordWrap(True)
        self.target_label = QLabel("Target Window: None")
        self.target_label.setWordWrap(True)
        self.current_target_label = QLabel("Current Target Relative Position: (0, 0)")
        self.image_label = QLabel()
        self.image_label.setMaximumSize(640, 480)
        self.image_label.setStyleSheet("background-color: white;")

        self.layout.addWidget(self.program_label)
        self.layout.addWidget(self.target_label)
        self.layout.addWidget(self.current_target_label)
        self.layout.addWidget(self.image_label)

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

        self.capture_thread = ImageCaptureThread(
            self.current_program_hwnd, self.winId()
        )
        self.capture_thread.image_captured.connect(self.update_image_label)
        self.capture_thread.start()

    def initVariables(self):
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.capture_thread = None

    def initListeners(self):
        self.mouse_listener = mouse.Listener(
            on_move=self.on_move, on_click=self.on_click
        )
        self.mouse_listener.start()
        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.keyboard_listener.start()
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_program_label)
        self.timer.start(1000)

    def on_move(self, x, y):
        self.current = (x, y)
        self.update_current_target_label()

    def on_click(self, x, y, button, pressed):
        if pressed:
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd:
                self.print_window_hierarchy(hwnd)
            if pressed and self.recording:
                self.record_click_event(x, y)

    def on_press(self, key):
        if key == keyboard.Key.f9:
            self.toggle_recording()
        elif key == keyboard.Key.f10:
            self.play_script()

    def record_click_event(self, x, y):
        hwnd = win32gui.WindowFromPoint((x, y))
        if hwnd == self.current_program_hwnd or hwnd == 0:
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

        logging.info(
            f"Recording click event: Program: {current_program}, Window Name: {window_name}, Window Class: {window_class}, Depth: {depth}"
        )

        self.click_events.append(
            {
                "relative_x": relative_x,
                "relative_y": relative_y,
                "program_name": current_program,
                "program_path": program_path,
                "window_name": window_name,
                "window_class": window_class,
                "depth": depth,
                "time": time.time() - self.start_time,
            }
        )

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
        if hwnd == 0:
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

    def play_script(self):
        if not self.click_events:
            return

        self.record_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.load_button.setEnabled(False)
        self.play_button.setEnabled(False)

        try:
            start_time = self.click_events[0]["time"]
            last_event_time = start_time

            for event in self.click_events:
                elapsed_time = (event["time"] - last_event_time) * self.speed_factor
                if elapsed_time > 0:
                    time.sleep(elapsed_time)
                last_event_time = event["time"]

                hwnd = self.find_target_hwnd(event)
                if hwnd is None or hwnd == 0:
                    logging.warning(f"Invalid hwnd for event: {event}")
                    continue

                self.send_click_event(event["relative_x"], event["relative_y"], hwnd)
                QApplication.processEvents()  # Update the UI

        except RuntimeError as e:
            logging.error(f"Error playing script: {e}")
            self.show_error_message(str(e))
        finally:
            self.record_button.setEnabled(True)
            self.save_button.setEnabled(True)
            self.load_button.setEnabled(True)
            self.play_button.setEnabled(True)

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

            self.activate_window(hwnd)
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


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        ex = MouseTracker()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Error in main: {e}")

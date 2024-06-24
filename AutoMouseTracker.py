import sys
import pyperclip
import win32gui
import win32process
import win32api
import win32con
import win32ui
import json
import time
import psutil
import cv2
import numpy as np
from PyQt5.QtWidgets import (
    QApplication,
    QLabel,
    QVBoxLayout,
    QWidget,
    QPushButton,
    QFileDialog,
    QSlider,
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPixmap, QImage
from pynput import keyboard, mouse
import ctypes
import logging

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("mouse_tracker.log"),
        logging.StreamHandler(sys.stdout),
    ],
)


class ImageCaptureThread(QThread):
    image_captured = pyqtSignal(np.ndarray)

    def __init__(self, hwnd, exclude_hwnd):
        super().__init__()
        self.hwnd = hwnd
        self.exclude_hwnd = exclude_hwnd
        self.running = True

    def run(self):
        try:
            while self.running:
                if self.hwnd == self.exclude_hwnd:
                    time.sleep(0.1)
                    continue

                try:
                    window_rect = win32gui.GetWindowRect(self.hwnd)
                except Exception as e:
                    logging.error(f"Error in GetWindowRect: {e}")
                    self.stop()
                    continue

                width = window_rect[2] - window_rect[0]
                height = window_rect[3] - window_rect[1]
                hwndDC = win32gui.GetWindowDC(self.hwnd)
                mfcDC = win32ui.CreateDCFromHandle(hwndDC)
                saveDC = mfcDC.CreateCompatibleDC()
                saveBitMap = win32ui.CreateBitmap()
                saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
                saveDC.SelectObject(saveBitMap)
                user32.PrintWindow(self.hwnd, saveDC.GetSafeHdc(), 1)
                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                img = np.frombuffer(bmpstr, dtype="uint8")
                img.shape = (height, width, 4)
                img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
                self.image_captured.emit(img)
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(self.hwnd, hwndDC)
                time.sleep(0.5)  # 시간 간격을 늘려 CPU 사용량 줄이기
        except Exception as e:
            logging.error(f"Error in ImageCaptureThread: {e}")

    def stop(self):
        self.running = False
        self.wait()


class MouseTracker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.capture_thread = None
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.hwnd_cache = {}
        self.hwnds = []

        self.mouse_listener = mouse.Listener(
            on_move=self.on_move, on_click=self.on_click
        )
        self.mouse_listener.start()

        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.keyboard_listener.start()

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_program_label)
        self.timer.start(1000)

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

    def on_move(self, x, y):
        try:
            self.current = (x, y)
            self.update_current_target_label()
            hwnd = win32gui.WindowFromPoint((x, y))
        except Exception as e:
            logging.error(f"Error in on_move: {e}")

    def on_click(self, x, y, button, pressed):
        try:
            if pressed and self.recording:
                hwnd = win32gui.WindowFromPoint((x, y))
                if hwnd == self.current_program_hwnd or hwnd == 0:
                    return
                window_rect = win32gui.GetWindowRect(hwnd)
                relative_x = x - window_rect[0]
                relative_y = y - window_rect[1]
                self.update_program_label(hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                current_program = "Unknown"
                program_path = "Unknown"
                for proc in psutil.process_iter(["pid", "name", "exe"]):
                    if proc.info["pid"] == pid:
                        current_program = proc.info["name"]
                        program_path = proc.info["exe"]
                        break

                window_name = win32gui.GetWindowText(hwnd) or "No Name"
                window_class = win32gui.GetClassName(hwnd)

                logging.info(
                    f"Recording click event: Program: {current_program}, Window Name: {window_name}, Window Class: {window_class}"
                )

                self.click_events.append(
                    {
                        "relative_x": relative_x,
                        "relative_y": relative_y,
                        "program_name": current_program,
                        "window_name": window_name,
                        "window_class": window_class,
                        "hwnd": hwnd,
                        "time": time.time() - self.start_time,
                    }
                )
        except Exception as e:
            logging.error(f"Error in on_click: {e}")

    def on_press(self, key):
        try:
            if key == keyboard.Key.f9:
                self.toggle_recording()
            elif key == keyboard.Key.f10:
                self.play_script()
        except Exception as e:
            logging.error(f"Error in on_press: {e}")

    @pyqtSlot(np.ndarray)
    def update_image_label(self, img):
        try:
            max_width, max_height = 640, 480
            height, width, channel = img.shape
            if width > max_width or height > max_height:
                aspect_ratio = width / height
                if width > height:
                    new_width = max_width
                    new_height = int(max_width / aspect_ratio)
                else:
                    new_height = max_height
                    new_width = int(max_height * aspect_ratio)
                img = cv2.resize(img, (new_width, new_height))
            height, width, channel = img.shape
            bytes_per_line = 3 * width
            qimg = QImage(img.data, width, height, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.image_label.setPixmap(pixmap)
        except Exception as e:
            logging.error(f"Error in update_image_label: {e}")

    def update_program_label(self, hwnd=None):
        try:
            if hwnd is None:
                x, y = self.current
                hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd == 0:
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            current_program = "Unknown"
            program_path = "Unknown"
            for proc in psutil.process_iter(["pid", "name", "exe"]):
                if proc.info["pid"] == pid:
                    current_program = proc.info["name"]
                    program_path = proc.info["exe"]
                    break

            window_name = win32gui.GetWindowText(hwnd) or "No Name"
            window_class = win32gui.GetClassName(hwnd)

            self.program_label.setText(
                f"Current Program: {current_program} ({program_path})"
            )
            self.target_label.setText(
                f"Target Window: {window_name} (hwnd: {hwnd}, class: {window_class})"
            )
        except Exception as e:
            logging.error(f"Error in update_program_label: {e}")

    def update_current_target_label(self):
        try:
            x, y = self.current
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd == self.current_program_hwnd or hwnd == 0:
                return
            window_rect = win32gui.GetWindowRect(hwnd)
            relative_x = x - window_rect[0]
            relative_y = y - window_rect[1]
            self.current_target_label.setText(
                f"Current Target Relative Position: ({relative_x}, {relative_y})"
            )
        except Exception as e:
            logging.error(f"Error in update_current_target_label: {e}")

    def toggle_recording(self):
        try:
            self.recording = not self.recording
            self.record_button.setText(
                "Stop Recording" if self.recording else "Start Recording"
            )
            if self.recording:
                self.start_time = time.time()
                self.click_events = []
        except Exception as e:
            logging.error(f"Error in toggle_recording: {e}")

    def save_script(self):
        try:
            options = QFileDialog.Options()
            filename, _ = QFileDialog.getSaveFileName(
                self,
                "Save Script",
                "",
                "JSON Files (*.json);;All Files (*)",
                options=options,
            )
            if filename:
                with open(filename, "w") as file:
                    json.dump(self.click_events, file)
        except Exception as e:
            logging.error(f"Error in save_script: {e}")

    def load_script(self):
        try:
            options = QFileDialog.Options()
            filename, _ = QFileDialog.getOpenFileName(
                self,
                "Load Script",
                "",
                "JSON Files (*.json);;All Files (*)",
                options=options,
            )
            if filename:
                with open(filename, "r") as file:
                    self.click_events = json.load(file)
        except Exception as e:
            logging.error(f"Error in load_script: {e}")

    def play_script(self, *args):
        if not self.click_events:
            return
        start_time = self.click_events[0]["time"]
        for event in self.click_events:
            relative_x = event["relative_x"]
            relative_y = event["relative_y"]
            hwnd = self.get_valid_hwnd(event)
            if hwnd is None or hwnd == 0:
                logging.warning(f"Invalid hwnd for event: {event}")
                continue
            t = event["time"]
            logging.info(
                f"Clicking at ({relative_x}, {relative_y}) in {t - start_time:.2f}s"
            )
            time.sleep((t - start_time) * self.speed_factor)
            self.send_click_event(relative_x, relative_y, hwnd)
            start_time = t

    def click_on_hwnd(hwnd, x, y):
        try:
            win32gui.PostMessage(
                hwnd,
                win32con.WM_LBUTTONDOWN,
                win32con.MK_LBUTTON,
                win32api.MAKELONG(x, y),
            )
            time.sleep(0.1)  # Small delay to mimic actual click
            win32gui.PostMessage(
                hwnd, win32con.WM_LBUTTONUP, None, win32api.MAKELONG(x, y)
            )
            logging.info(f"Clicked at ({x}, {y}) on hwnd: {hwnd}")
        except Exception as e:
            logging.error(f"Failed to click on hwnd: {hwnd}, error: {e}")

    def find_hwnd(self, program_name, window_class=None, window_name=None):
        hwnds = []

        def callback(hwnd, extra):
            if not win32gui.IsWindow(hwnd):
                return True

            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(found_pid)
                if process.name().lower() == program_name.lower():
                    class_name = win32gui.GetClassName(hwnd)
                    window_text = win32gui.GetWindowText(hwnd) or "No Name"

                    class_match = window_class in class_name if window_class else True
                    name_match = window_name in window_text if window_name else True

                    if class_match and name_match:
                        hwnds.append(hwnd)
                    # 자식 윈도우 검색
                    child_windows = self.find_child_windows(
                        hwnd, window_class, window_name
                    )
                    hwnds.extend(child_windows)
            except psutil.NoSuchProcess:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        return hwnds

    @staticmethod
    def find_child_windows(parent_hwnd, class_name=None, window_text=None):
        child_windows = []

        def enum_child_proc(hwnd, lParam):
            if not win32gui.IsWindow(hwnd):
                return True

            match_class = True
            match_text = True

            if class_name:
                match_class = win32gui.GetClassName(hwnd) == class_name
            if window_text:
                match_text = win32gui.GetWindowText(hwnd) == window_text

            if match_class and match_text:
                child_windows.append(hwnd)

            return True

        win32gui.EnumChildWindows(parent_hwnd, enum_child_proc, None)
        return child_windows

    def get_valid_hwnd(self, event):
        cache_key = (event["program_name"], event["window_class"], event["window_name"])

        hwnd = event.get("hwnd")
        if hwnd and win32gui.IsWindow(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid)
                if process.name().lower() == event["program_name"].lower():
                    logging.debug(f"Using cached hwnd: {hwnd}")
                    return hwnd
            except psutil.NoSuchProcess:
                pass

        if cache_key in self.hwnd_cache:
            cached_hwnd = self.hwnd_cache[cache_key]
            if win32gui.IsWindow(cached_hwnd):
                _, pid = win32process.GetWindowThreadProcessId(cached_hwnd)
                try:
                    process = psutil.Process(pid)
                    if process.name().lower() == event["program_name"].lower():
                        logging.debug(f"Using cached hwnd from cache: {cached_hwnd}")
                        return cached_hwnd
                except psutil.NoSuchProcess:
                    pass

        hwnds = self.find_hwnd(
            event["program_name"], event["window_class"], event["window_name"]
        )
        if hwnds:
            hwnd = hwnds[0]
            self.hwnd_cache[cache_key] = hwnd
            logging.debug(f"Using found hwnd: {hwnd}")
            return hwnd

        # 윈도우 이름이 없거나 찾지 못한 경우
        hwnds = self.find_hwnd(event["program_name"], event["window_class"])
        if hwnds:
            hwnd = hwnds[0]
            self.hwnd_cache[cache_key] = hwnd
            logging.debug(f"Using found hwnd with class only: {hwnd}")
            return hwnd

        logging.warning(f"No valid hwnd found for event: {event}")
        return None

    def send_click_event(self, relative_x, relative_y, hwnd):
        if not hwnd or not win32gui.IsWindow(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return

        lParam = win32api.MAKELONG(relative_x, relative_y)
        logging.info(f"Clicking at ({relative_x}, {relative_y})")

        win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        time.sleep(0.125)

        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
        time.sleep(0.125)

        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        time.sleep(0.075)

        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)
        time.sleep(0.075)

        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
        time.sleep(0.125)

    def update_speed_factor(self, value):
        try:
            self.speed_factor = value / 50.0
        except Exception as e:
            logging.error(f"Error in update_speed_factor: {e}")

    def closeEvent(self, event):
        if self.capture_thread is not None:
            self.capture_thread.stop()
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
        super().closeEvent(event)


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        ex = MouseTracker()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Error in main: {e}")

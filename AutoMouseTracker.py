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
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QImage
from pynput import keyboard, mouse
import ctypes

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32


class ImageCaptureThread(QThread):
    image_captured = pyqtSignal(np.ndarray)

    def __init__(self, hwnd):
        super().__init__()
        self.hwnd = hwnd
        self.running = True

    def run(self):
        try:
            while self.running:
                window_rect = win32gui.GetWindowRect(self.hwnd)
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
                time.sleep(0.1)
        except Exception as e:
            print(f"Error in ImageCaptureThread: {e}")

    def stop(self):
        self.running = False


class MouseTracker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.capture_thread = None

        self.mouse_listener = mouse.Listener(
            on_move=self.on_move, on_click=self.on_click
        )
        self.mouse_listener.start()

        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.keyboard_listener.start()

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_program_label)
        self.timer.start(500)  # 500ms마다 업데이트

    def initUI(self):
        self.layout = QVBoxLayout()

        self.program_label = QLabel("Current Program: None")
        self.program_label.setWordWrap(True)
        self.target_label = QLabel("Target Window: None")
        self.target_label.setWordWrap(True)
        self.current_target_label = QLabel("Current Target Relative Position: (0, 0)")
        self.image_label = QLabel()
        self.image_label.setMaximumSize(1000, 800)
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
        except Exception as e:
            print(f"Error in on_move: {e}")

    def on_click(self, x, y, button, pressed):
        try:
            if pressed and self.recording:
                hwnd = win32gui.WindowFromPoint((x, y))
                window_rect = win32gui.GetWindowRect(hwnd)
                relative_x = x - window_rect[0]
                relative_y = y - window_rect[1]
                self.update_program_label(hwnd)
                self.click_events.append(
                    (relative_x, relative_y, hwnd, time.time() - self.start_time)
                )
        except Exception as e:
            print(f"Error in on_click: {e}")

    def on_press(self, key):
        try:
            if key == keyboard.Key.f10:  # 녹화 시작/정지 단축키
                self.toggle_recording()
            elif key == keyboard.Key.f11:  # 스크립트 재생 단축키
                self.play_script()
        except Exception as e:
            print(f"Error in on_press: {e}")

    def update_program_label(self, hwnd=None):
        try:
            if hwnd is None:
                x, y = self.current
                hwnd = win32gui.WindowFromPoint((x, y))
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

            target_info = {
                "program": current_program,
                "hwnd": hwnd,
                "class": window_class,
                "path": program_path,
                "window_name": window_name,
            }
            pyperclip.copy(json.dumps(target_info))

            # Start capturing the target window's image
            if self.capture_thread is not None:
                self.capture_thread.stop()
            self.capture_thread = ImageCaptureThread(hwnd)
            self.capture_thread.image_captured.connect(self.update_image_label)
            self.capture_thread.start()
        except Exception as e:
            print(f"Error in update_program_label: {e}")

    def update_image_label(self, img):
        try:
            max_width, max_height = 1000, 800
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
            print(f"Error in update_image_label: {e}")

    def update_current_target_label(self):
        try:
            x, y = self.current
            hwnd = win32gui.WindowFromPoint((x, y))
            window_rect = win32gui.GetWindowRect(hwnd)
            relative_x = x - window_rect[0]
            relative_y = y - window_rect[1]
            self.current_target_label.setText(
                f"Current Target Relative Position: ({relative_x}, {relative_y})"
            )
        except Exception as e:
            print(f"Error in update_current_target_label: {e}")

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
            print(f"Error in toggle_recording: {e}")

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
            print(f"Error in save_script: {e}")

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
            print(f"Error in load_script: {e}")

    def play_script(self):
        try:
            if not self.click_events:
                return
            start_time = self.click_events[0][3]
            for relative_x, relative_y, hwnd, t in self.click_events:
                print(
                    f"Clicking at ({relative_x}, {relative_y}) in {t - start_time:.2f}s"
                )
                time.sleep((t - start_time) * self.speed_factor)
                self.send_click_event(relative_x, relative_y, hwnd)
                start_time = t
        except Exception as e:
            print(f"Error in play_script: {e}")

    def send_click_event(self, relative_x, relative_y, hwnd):
        try:
            # Redraw the window image with a red dot at the click position
            self.update_image_with_click(hwnd, relative_x, relative_y)

            lParam = win32api.MAKELONG(relative_x, relative_y)

            print(f"Clicking at ({relative_x}, {relative_y})")

            # 백그라운드에서 활성화
            win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
            win32api.Sleep(125)

            # 백그라운드 마우스 이동 이벤트 전송
            win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
            win32api.Sleep(125)

            # 마우스 클릭 다운 이벤트 전송
            win32gui.PostMessage(
                hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam
            )
            win32api.Sleep(75)  # 클릭 유지 시간

            # 마우스 클릭 업 이벤트 전송
            win32gui.PostMessage(
                hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam
            )
            win32api.Sleep(75)  # 클릭 유지 시간

            # 백그라운드 마우스 이동 이벤트 전송
            win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
            win32api.Sleep(125)
        except Exception as e:
            print(f"Error in send_click_event: {e}")

    def update_image_with_click(self, hwnd, relative_x, relative_y):
        try:
            window_rect = win32gui.GetWindowRect(hwnd)
            width = window_rect[2] - window_rect[0]
            height = window_rect[3] - window_rect[1]
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            result = user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype="uint8")
            img.shape = (height, width, 4)
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
            cv2.circle(img, (relative_x, relative_y), 5, (255, 0, 0), -1)
            self.update_image_label(img)
        except Exception as e:
            print(f"Error in update_image_with_click: {e}")

    def update_speed_factor(self, value):
        try:
            self.speed_factor = value / 50.0  # 기본 속도는 1.0, 0.02 ~ 2.0 배속 조절
        except Exception as e:
            print(f"Error in update_speed_factor: {e}")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        ex = MouseTracker()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error in main: {e}")

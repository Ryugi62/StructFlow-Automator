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
                if hwnd == self.current_program_hwnd or hwnd == 0:
                    print("No hwnd found")
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

                print(
                    f"Recording click event: Program: {current_program}, Window Name: {window_name}, Window Class: {window_class}"
                )

                self.click_events.append(
                    {
                        "relative_x": relative_x,
                        "relative_y": relative_y,
                        "program_name": current_program,
                        "window_name": window_name,
                        "window_class": window_class,
                        "hwnd": hwnd,  # Store the hwnd
                        "time": time.time() - self.start_time,
                    }
                )
        except Exception as e:
            print(f"Error in on_click: {e}")

    def on_press(self, key):
        try:
            if key == keyboard.Key.f9:  # 녹화 시작/정지 단축키
                self.toggle_recording()
            elif key == keyboard.Key.f10:  # 스크립트 재생 단축키
                self.play_script()
        except Exception as e:
            print(f"Error in on_press: {e}")

    @pyqtSlot(np.ndarray)
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
                    new_width = int(max.height * aspect_ratio)
                img = cv2.resize(img, (new_width, new_height))
            height, width, channel = img.shape
            bytes_per_line = 3 * width
            qimg = QImage(img.data, width, height, bytes_per_line, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            self.image_label.setPixmap(pixmap)
        except Exception as e:
            print(f"Error in update_image_label: {e}")

    def update_program_label(self, hwnd=None):
        try:
            if hwnd is None:
                x, y = self.current
                hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd == 0:
                print("No hwnd found")
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

            # Start capturing the target window's image
            if self.capture_thread is not None:
                self.capture_thread.stop()
            self.capture_thread = ImageCaptureThread(hwnd)
            self.capture_thread.image_captured.connect(self.update_image_label)
            self.capture_thread.start()
        except Exception as e:
            print(f"Error in update_program_label: {e}")

    def update_current_target_label(self):
        try:
            x, y = self.current
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd == self.current_program_hwnd or hwnd == 0:
                print("No hwnd found")
                return
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
            start_time = self.click_events[0]["time"]
            for event in self.click_events:
                relative_x = event["relative_x"]
                relative_y = event["relative_y"]
                hwnd = self.get_valid_hwnd(event)
                if hwnd is None or hwnd == 0:
                    print(f"Invalid hwnd for event: {event}")
                    continue
                t = event["time"]
                print(
                    f"Clicking at ({relative_x}, {relative_y}) in {t - start_time:.2f}s"
                )
                time.sleep((t - start_time) * self.speed_factor)
                self.send_click_event(relative_x, relative_y, hwnd)
                start_time = t
        except Exception as e:
            print(f"Error in play_script: {e}")

    def get_valid_hwnd(self, event):
        hwnd = event.get("hwnd")
        if hwnd and win32gui.IsWindow(hwnd):
            return hwnd

        # Fallback to finding hwnd by program name and window class
        hwnd = self.find_hwnd(event["program_name"], event["window_class"])
        return hwnd

    def find_hwnd(self, program_name, window_class):
        try:
            hwnds = []

            def callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process_name = None
                    for proc in psutil.process_iter(["pid", "name"]):
                        if proc.info["pid"] == pid:
                            process_name = proc.info["name"]
                            break
                    hwnd_window_class = win32gui.GetClassName(hwnd)
                    if (
                        process_name == program_name
                        and window_class.split(":")[0] in hwnd_window_class
                    ):
                        hwnds.append(hwnd)

            win32gui.EnumWindows(callback, None)
            if hwnds:
                for hwnd in hwnds:
                    child_hwnd = self.find_child_hwnd_recursive(hwnd, window_class)
                    if child_hwnd:
                        return child_hwnd
                return hwnds[0]
            else:
                print(
                    f"No hwnd found for program: {program_name}, window class: {window_class}"
                )
                return None
        except Exception as e:
            print(f"Error in find_hwnd: {e}")
            return None

    def find_child_hwnd_recursive(self, parent_hwnd, window_class):
        try:
            hwnds = []

            def callback(hwnd, hwnds):
                hwnd_window_class = win32gui.GetClassName(hwnd)
                if window_class.split(":")[0] in hwnd_window_class:
                    hwnds.append(hwnd)
                else:
                    child_hwnd = self.find_child_hwnd_recursive(hwnd, window_class)
                    if child_hwnd:
                        hwnds.append(child_hwnd)

            win32gui.EnumChildWindows(parent_hwnd, callback, hwnds)
            return hwnds[0] if hwnds else None
        except Exception as e:
            print(f"Error in find_child_hwnd_recursive: {e}")
            return None

    def find_child_hwnd(self, parent_hwnd, window_class):
        try:
            hwnds = []

            def callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    hwnd_window_class = win32gui.GetClassName(hwnd)
                    if window_class.split(":")[0] in hwnd_window_class:
                        hwnds.append(hwnd)

            win32gui.EnumChildWindows(parent_hwnd, callback, None)
            if hwnds:
                return hwnds[0]
            else:
                return None
        except Exception as e:
            print(f"Error in find_child_hwnd: {e}")
            return None

    def find_parent_hwnd(child_hwnd):
        parent_hwnd = win32gui.GetParent(child_hwnd)
        if parent_hwnd == 0:
            parent_hwnd = win32gui.GetAncestor(child_hwnd, win32con.GA_ROOT)
        return parent_hwnd

    def send_click_event(self, relative_x, relative_y, hwnd):
        try:
            if not hwnd or not win32gui.IsWindow(hwnd):
                print(f"Invalid hwnd: {hwnd}")
                return

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
        print(f"Error in main: {e}")

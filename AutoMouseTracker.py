import json
import time
import logging
import ctypes
import os
import uuid
from ctypes import wintypes
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
    QProgressBar,
    QStatusBar,
    QMenuBar,
    QAction,
    QGridLayout,
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPixmap, QImage, QIcon
import numpy as np
import cv2
import psutil
import win32gui
import win32process
import win32api
import win32con
import win32ui
from pynput import keyboard, mouse
from logging.handlers import RotatingFileHandler
import pywintypes
import sys
from ctypes import windll


# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
LOG_FILE = "mouse_tracker.log"

# Logging configuration
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[RotatingFileHandler(LOG_FILE, maxBytes=10**6, backupCount=3)],
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
            # 윈도우 크기 가져오기
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            # 윈도우 DC와 메모리 DC 생성
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            # 비트맵 객체 생성
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

            # 비트맵에 윈도우 내용 복사 (PrintWindow 사용)
            result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

            if result == 0:
                print(f"PrintWindow failed, trying BitBlt")
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

            # 비트맵을 numpy 배열로 변환
            signedIntsArray = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(signedIntsArray, dtype="uint8")
            img.shape = (height, width, 4)

            # 색상 공간 변환 (BGRA to BGR)
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

            # 리소스 해제
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

    def init_variables(self):
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.capture_thread = None
        self.dark_mode = False

    def init_ui(self):
        self.layout = QGridLayout()
        self.create_labels()
        self.create_buttons()
        self.create_sliders()
        self.create_progress_bar()
        self.create_status_bar()
        self.create_menu_bar()
        self.setLayout(self.layout)
        self.setWindowTitle("Mouse Tracker")
        self.setGeometry(100, 100, 800, 600)
        self.show()

    def create_labels(self):
        self.program_label = QLabel("Current Program: None")
        self.program_label.setWordWrap(True)
        self.target_label = QLabel("Target Window: None")
        self.target_label.setWordWrap(True)
        self.current_target_label = QLabel("Current Target Relative Position: (0, 0)")
        self.image_label = QLabel()
        self.image_label.setMaximumSize(640, 480)
        self.image_label.setStyleSheet(
            "background-color: white; border: 1px solid #ddd;"
        )
        self.event_list = QListWidget()

        self.layout.addWidget(self.program_label, 0, 0, 1, 2)
        self.layout.addWidget(self.target_label, 1, 0, 1, 2)
        self.layout.addWidget(self.current_target_label, 2, 0, 1, 2)
        self.layout.addWidget(self.image_label, 3, 0, 1, 2)
        self.layout.addWidget(self.event_list, 4, 0, 1, 2)

    def create_buttons(self):
        self.record_button = self.create_button(
            "Start Recording", "icons/record.png", self.toggle_recording
        )
        self.save_button = self.create_button(
            "Save Script", "icons/save.png", self.save_script
        )
        self.load_button = self.create_button(
            "Load Script", "icons/load.png", self.load_script
        )
        self.play_button = self.create_button(
            "Play Script", "icons/play.png", self.play_script
        )

        self.layout.addWidget(self.record_button, 5, 0)
        self.layout.addWidget(self.save_button, 5, 1)
        self.layout.addWidget(self.load_button, 6, 0)
        self.layout.addWidget(self.play_button, 6, 1)

    def create_button(self, text, icon_path, callback):
        button = QPushButton(text)
        button.setIcon(QIcon(icon_path))
        button.clicked.connect(callback)
        return button

    def create_sliders(self):
        self.speed_slider = QSlider(Qt.Horizontal)
        self.speed_slider.setRange(1, 100)
        self.speed_slider.setValue(50)
        self.speed_slider.valueChanged.connect(self.update_speed_factor)
        self.layout.addWidget(self.speed_slider, 7, 0, 1, 2)

    def create_progress_bar(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar, 8, 0, 1, 2)

    def create_status_bar(self):
        self.status_bar = QStatusBar()
        self.layout.addWidget(self.status_bar, 9, 0, 1, 2)

    def create_menu_bar(self):
        self.menu_bar = QMenuBar()
        self.view_menu = self.menu_bar.addMenu("View")
        self.dark_mode_action = QAction("Toggle Dark Mode", self)
        self.dark_mode_action.triggered.connect(self.toggle_dark_mode)
        self.view_menu.addAction(self.dark_mode_action)
        self.layout.setMenuBar(self.menu_bar)

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
            try:
                hwnd = win32gui.WindowFromPoint((x, y))
                if hwnd:
                    self.print_window_hierarchy(hwnd)
                if pressed and self.recording and not self.is_own_window(hwnd):
                    self.record_click_event(x, y, hwnd, move_cursor=False)
            except Exception as e:
                logging.error(f"Error in on_click: {e}")

    def is_own_window(self, hwnd):
        window_title = win32gui.GetWindowText(hwnd)
        return window_title == "Mouse Tracker"

    def on_press(self, key):
        if key == keyboard.Key.f9:
            self.toggle_recording()
        elif key == keyboard.Key.f10:
            self.play_script()

    def record_click_event(self, x, y, hwnd, move_cursor):
        if hwnd == self.current_program_hwnd or self.is_own_window(hwnd):
            return
        window_rect = win32gui.GetWindowRect(hwnd)
        relative_x, relative_y = x - window_rect[0], y - window_rect[1]
        width, height = window_rect[2] - window_rect[0], window_rect[3] - window_rect[1]
        self.update_program_label(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        current_program, program_path = self.get_program_info(pid)
        window_name, window_class = win32gui.GetWindowText(hwnd), win32gui.GetClassName(
            hwnd
        )
        depth = self.get_window_depth(hwnd)

        img = self.capture_thread.capture_window_image(hwnd)
        if img is not None:
            # 전체 이미지 저장
            unique_id = uuid.uuid4().hex
            image_filename = (
                f"{current_program}_{unique_id}_{relative_x}_{relative_y}.png"
            )
            save_dir = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "StructFlow-Automator-Private",
                "sample_targets",
            )
            os.makedirs(save_dir, exist_ok=True)
            full_path = os.path.join(save_dir, image_filename)
            cv2.imwrite(full_path, cv2.cvtColor(img, cv2.COLOR_RGB2BGR))

            # 클릭한 위치 주변의 작은 영역을 추출
            region_size = 50  # 추출할 영역의 크기 (50x50 픽셀)
            x1 = max(0, relative_x - region_size // 2)
            y1 = max(0, relative_y - region_size // 2)
            x2 = min(img.shape[1], x1 + region_size)
            y2 = min(img.shape[0], y1 + region_size)
            target_region = img[y1:y2, x1:x2]

            target_image_filename = f"{current_program}_{unique_id}_target.png"
            target_full_path = os.path.join(save_dir, target_image_filename)
            cv2.imwrite(
                target_full_path, cv2.cvtColor(target_region, cv2.COLOR_RGB2BGR)
            )

        window_title = win32gui.GetWindowText(hwnd)

        event_info = {
            "relative_x": relative_x,
            "relative_y": relative_y,
            "program_name": current_program,
            "program_path": program_path,
            "window_name": window_name,
            "window_class": window_class,
            "window_title": window_title,
            "depth": depth,
            "time": time.time() - self.start_time,
            "window_rect": window_rect,
            "move_cursor": move_cursor,
            "image": {
                "path": full_path,
                "target_path": target_full_path,
                "image_program_name": current_program,
                "image_program_path": program_path,
                "image_window_class": window_class,
                "image_window_name": window_name,
                "image_depth": depth,
                "wait_for_image": True,
                "wait_method": "image",
                "window_title": window_title,
                "comparison_threshold": 0.74,  # 기본 임계값 추가
            },
        }

        self.update_image_label(img)
        self.capture_thread.hwnd = hwnd
        logging.info(f"Recording click event: {event_info}")
        self.click_events.append(event_info)
        self.event_list.addItem(
            f"Click at ({relative_x}, {relative_y}) in {current_program} ({window_title})"
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
        if img is None:
            logging.error("Failed to capture image; img is None")
            return

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
        if hwnd == self.current_program_hwnd:
            return
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        current_program, program_path = self.get_program_info(pid)

        window_name = win32gui.GetWindowText(hwnd)
        if window_name == "":
            window_name = "None"
        window_class = win32gui.GetClassName(hwnd)
        if window_class == "":
            window_class = "None"

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
        self.status_bar.showMessage(
            "Recording started" if self.recording else "Recording stopped"
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
            self.status_bar.showMessage(f"Script saved to {filename}")

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
            self.status_bar.showMessage(f"Script loaded from {filename}")

    def wait_for_target_image(
        self, target_image_path, hwnd, timeout=300, threshold=0.74
    ):
        start_time = time.time()
        target_image = cv2.imread(target_image_path, cv2.IMREAD_COLOR)

        if target_image is None:
            logging.error(f"Failed to load target image: {target_image_path}")
            return False

        while time.time() - start_time < timeout:
            try:
                current_image = self.capture_thread.capture_window_image(hwnd)
                if current_image is None:
                    logging.warning("Failed to capture current window image")
                    time.sleep(0.5)
                    continue

                # 이미지 크기 조정
                if current_image.shape[:2] != target_image.shape[:2]:
                    current_image = cv2.resize(
                        current_image, (target_image.shape[1], target_image.shape[0])
                    )

                # 여러 매칭 방법 시도
                methods = [cv2.TM_CCOEFF_NORMED, cv2.TM_SQDIFF_NORMED]
                for method in methods:
                    result = cv2.matchTemplate(current_image, target_image, method)
                    if method == cv2.TM_SQDIFF_NORMED:
                        min_val, _, min_loc, _ = cv2.minMaxLoc(result)
                        match_val = (
                            1 - min_val
                        )  # TM_SQDIFF_NORMED에서는 값이 작을수록 일치
                    else:
                        _, max_val, _, max_loc = cv2.minMaxLoc(result)
                        match_val = max_val

                    print(
                        f"Image matching method: {method}, similarity: {match_val:.4f}"
                    )

                    if match_val >= threshold:
                        print(
                            f"Target image found. Method: {method}, Confidence: {match_val:.4f}"
                        )
                        return True

            except Exception as e:
                logging.error(f"Error while searching for target image: {e}")

            time.sleep(0.8)

        print(f"Target image not found within {timeout} seconds")
        return False

    def play_script(self):
        if not self.click_events:
            return

        self.set_buttons_enabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(self.click_events))

        def process_event(index):
            if index >= len(self.click_events):
                self.set_buttons_enabled(True)
                self.progress_bar.setValue(len(self.click_events))
                logging.info("Script playback completed.")
                return

            event = self.click_events[index]
            start_time = time.time()
            max_wait_time = 300  # 5 minutes in seconds

            def try_process_event():
                nonlocal start_time
                try:
                    wait_for_image = event["image"].get("wait_for_image", False)
                    wait_method = event["image"].get("wait_method", "")
                    target_image_path = event["image"].get("target_path")
                    threshold = event["image"].get(
                        "comparison_threshold", 0.74
                    )  # 임계값 가져오기

                    hwnd = None
                    if wait_for_image and wait_method == "image":
                        while time.time() - start_time < max_wait_time:
                            hwnd = self.find_target_hwnd(event)
                            if hwnd is not None and hwnd != 0:
                                break
                            time.sleep(2)  # Wait 2 seconds before trying again

                        if hwnd is None or hwnd == 0:
                            logging.warning(f"Failed to find hwnd for event: {event}")
                            QTimer.singleShot(
                                500, try_process_event
                            )  # Retry after 0.5 seconds
                            return

                        while time.time() - start_time < max_wait_time:
                            if target_image_path and os.path.exists(target_image_path):
                                remaining_time = max_wait_time - (
                                    time.time() - start_time
                                )
                                if self.wait_for_target_image(
                                    target_image_path,
                                    hwnd,
                                    timeout=remaining_time,
                                    threshold=threshold,  # 임계값 사용
                                ):
                                    logging.info(
                                        f"Target image found: {target_image_path}"
                                    )
                                    break
                            time.sleep(2)  # Wait 2 seconds before trying again

                        if time.time() - start_time >= max_wait_time:
                            logging.warning(
                                f"Failed to find target image within 5 minutes: {event}"
                            )
                            QTimer.singleShot(
                                500, try_process_event
                            )  # Retry after 0.5 seconds
                            return
                    else:
                        hwnd = self.find_target_hwnd(event)
                        if hwnd is None or hwnd == 0:
                            logging.warning(f"Failed to find hwnd for event: {event}")
                            QTimer.singleShot(
                                500, try_process_event
                            )  # Retry after 0.5 seconds
                            return

                    self.send_click_event(
                        event["relative_x"],
                        event["relative_y"],
                        hwnd,
                        event["move_cursor"],
                    )
                    logging.info(
                        f"Click event sent: x={event['relative_x']}, y={event['relative_y']}"
                    )

                    img = self.capture_thread.capture_window_image(hwnd)
                    if img is not None:
                        self.update_image_label(img)
                    self.capture_thread.hwnd = hwnd
                    self.event_list.setCurrentRow(index)
                    self.progress_bar.setValue(index + 1)

                    logging.info(
                        f"Processed event {index + 1}/{len(self.click_events)}"
                    )
                    QTimer.singleShot(
                        500, lambda: process_event(index + 1)
                    )  # Proceed to next event after 0.5 seconds

                except Exception as e:
                    logging.error(f"Error processing event {index}: {e}")
                    QTimer.singleShot(500, try_process_event)  # Retry after 0.5 seconds

            QTimer.singleShot(100, try_process_event)

        process_event(0)

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
            event["window_title"],
            event["window_rect"],
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
        window_title=None,
        window_rect=None,
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
                name_match = window_name == window_text if window_name else True
                title_match = window_title == window_text if window_title else True
                rect_match = True
                if window_rect:
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
    ):
        child_windows = []

        def enum_child_proc(hwnd, _):
            if not win32gui.IsWindow(hwnd):
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
            if window_rect:
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
                    )
                )
            return True

        win32gui.EnumChildWindows(parent_hwnd, enum_child_proc, None)
        return child_windows

    def send_click_event(self, relative_x, relative_y, hwnd, move_cursor):
        if not hwnd or not win32gui.IsWindow(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return

        lParam = win32api.MAKELONG(relative_x, relative_y)
        if move_cursor:
            now_x, now_y = win32api.GetCursorPos()
            now_x += 10
            now_y += 10
            max_w, max_h = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
            if now_x > max_w:
                now_x -= 20
            if now_y > max_h:
                now_y -= 20

            try:
                win32api.SetCursorPos((now_x, now_y))
                time.sleep(0.1)
            except Exception as e:
                logging.error(f"Failed to move cursor: {e}")

        try:
            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
            time.sleep(0.1)

            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)
            time.sleep(0.1)

            logging.debug("Click event sent successfully.")
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")

    def activate_window(self, hwnd):
        try:
            win32gui.SendMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_CLICKACTIVE, 0)
            time.sleep(0.1)
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
        try:
            hierarchy = self.find_window_hierarchy(hwnd)
            self.display_window_hierarchy(hierarchy)
        except Exception as e:
            logging.error(f"Error in print_window_hierarchy: {e}")

    def find_window_hierarchy(self, hwnd):
        hierarchy = []
        current_hwnd = hwnd

        while current_hwnd:
            try:
                window_title = win32gui.GetWindowText(current_hwnd)
                window_class = win32gui.GetClassName(current_hwnd)
                hierarchy.append((window_title, window_class, current_hwnd))
                current_hwnd = win32gui.GetParent(current_hwnd)
            except pywintypes.error as e:
                logging.warning(f"Error getting window info: {e}")
                break
            except Exception as e:
                logging.warning(f"Unexpected error in find_window_hierarchy: {e}")
                break

        return hierarchy[::-1]

    def display_window_hierarchy(self, hierarchy):
        indent = "   "
        for i, (title, class_name, hwnd) in enumerate(hierarchy):
            logging.info(
                f"{indent * i}Window Title: {title}, Window Class: {class_name}, HWND: {hwnd}"
            )

    def start_search_for_condition(self):
        def condition(img):
            target_image = cv2.imread("path_to_target_image.png")
            result = cv2.matchTemplate(img, target_image, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            return max_val > 0.8

        self.capture_thread.set_condition(condition)
        self.capture_thread.start()

    def on_condition_met(self):
        logging.info("Condition met, stopping search.")
        self.handle_timeout_error()

    def handle_timeout_error(self):
        logging.error("Timeout occurred during image search.")
        QMessageBox.critical(self, "Error", "Timeout occurred during image search.")

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        if self.dark_mode:
            self.setStyleSheet("background-color: #2E2E2E; color: #FFFFFF;")
        else:
            self.setStyleSheet("background-color: #FFFFFF; color: #000000;")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        ex = MouseTracker()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Error in main: {e}")

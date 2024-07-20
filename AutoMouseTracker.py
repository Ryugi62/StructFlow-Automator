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
    QListWidgetItem,
    QProgressBar,
    QStatusBar,
    QMenuBar,
    QAction,
    QGridLayout,
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLineEdit,
    QSpinBox,
    QComboBox,
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPixmap, QImage, QIcon
import numpy as np
import cv2
import logging
import json
import time
import os
import uuid
import sys
from ctypes import windll
import psutil
import win32gui
import win32process
import win32api
import win32con
import win32ui
from pynput import mouse, keyboard
from logging.handlers import RotatingFileHandler

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
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = right - left
            height = bottom - top

            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()

            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)

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

    def set_condition(self, condition_func):
        self.condition = condition_func


class EventSettingsDialog(QDialog):
    def __init__(self, event, parent=None):
        super(EventSettingsDialog, self).__init__(parent)
        self.event = event
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.setWindowTitle("Event Settings")

        self.move_cursor_checkbox = QCheckBox("Move Cursor")
        self.move_cursor_checkbox.setChecked(self.event.get("move_cursor", False))
        layout.addWidget(self.move_cursor_checkbox)

        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Click Delay (ms):"))
        self.delay_spinbox = QSpinBox()
        self.delay_spinbox.setRange(0, 10000)
        self.delay_spinbox.setValue(self.event.get("click_delay", 0))
        delay_layout.addWidget(self.delay_spinbox)
        layout.addLayout(delay_layout)

        self.double_click_checkbox = QCheckBox("Double Click")
        self.double_click_checkbox.setChecked(self.event.get("double_click", False))
        layout.addWidget(self.double_click_checkbox)

        keyboard_layout = QHBoxLayout()
        keyboard_layout.addWidget(QLabel("Keyboard Input:"))
        self.keyboard_input = QLineEdit(self.event.get("keyboard_input", ""))
        keyboard_layout.addWidget(self.keyboard_input)
        layout.addLayout(keyboard_layout)

        condition_layout = QHBoxLayout()
        condition_layout.addWidget(QLabel("Conditional Execution:"))
        self.condition_combo = QComboBox()
        self.condition_combo.addItems(
            [
                "None",
                "Image Present",
                "Image Not Present",
                "Wait up to 5 minutes for Image",
            ]
        )
        self.condition_combo.setCurrentText(self.event.get("condition", "None"))
        condition_layout.addWidget(self.condition_combo)
        layout.addLayout(condition_layout)

        repeat_layout = QHBoxLayout()
        repeat_layout.addWidget(QLabel("Repeat Count:"))
        self.repeat_spinbox = QSpinBox()
        self.repeat_spinbox.setRange(1, 1000)
        self.repeat_spinbox.setValue(self.event.get("repeat_count", 1))
        repeat_layout.addWidget(self.repeat_spinbox)
        layout.addLayout(repeat_layout)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_settings)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.close)
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def save_settings(self):
        self.event["move_cursor"] = self.move_cursor_checkbox.isChecked()
        self.event["click_delay"] = self.delay_spinbox.value()
        self.event["double_click"] = self.double_click_checkbox.isChecked()
        self.event["keyboard_input"] = self.keyboard_input.text()
        self.event["condition"] = self.condition_combo.currentText()
        self.event["repeat_count"] = self.repeat_spinbox.value()
        self.accept()


class MouseTracker(QWidget):
    def __init__(self):
        super().__init__()
        self.init_variables()
        self.init_ui()
        self.init_listeners()
        self.init_capture_thread()
        self.init_emergency_stop()

    def init_variables(self):
        self.current = (0, 0)
        self.click_events = []
        self.recording = False
        self.speed_factor = 1.0
        self.current_program_hwnd = win32gui.GetForegroundWindow()
        self.capture_thread = None
        self.dark_mode = False
        self.custom_image_path = None
        self.running_script = False

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
        self.event_list.itemClicked.connect(self.update_image_on_selection)

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
        self.add_custom_image_button = self.create_button(
            "Add Custom Image", "icons/custom_image.png", self.add_custom_image
        )
        self.settings_button = self.create_button(
            "Event Settings", "icons/settings.png", self.show_event_settings
        )
        self.stop_button = self.create_button(
            "Emergency Stop", "icons/stop.png", self.emergency_stop
        )

        self.layout.addWidget(self.record_button, 5, 0)
        self.layout.addWidget(self.save_button, 5, 1)
        self.layout.addWidget(self.load_button, 6, 0)
        self.layout.addWidget(self.play_button, 6, 1)
        self.layout.addWidget(self.add_custom_image_button, 7, 0, 1, 2)
        self.layout.addWidget(self.settings_button, 8, 0, 1, 2)
        self.layout.addWidget(self.stop_button, 9, 0, 1, 2)

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
        self.layout.addWidget(self.speed_slider, 10, 0, 1, 2)

    def create_progress_bar(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.layout.addWidget(self.progress_bar, 11, 0, 1, 2)

    def create_status_bar(self):
        self.status_bar = QStatusBar()
        self.layout.addWidget(self.status_bar, 12, 0, 1, 2)

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

            sizes = [(30, 30), (50, 50), (70, 70)]
            target_image_paths = []

            for size in sizes:
                region_size = size[0]
                x1 = max(0, relative_x - region_size // 2)
                y1 = max(0, relative_y - region_size // 2)
                x2 = min(img.shape[1], x1 + region_size)
                y2 = min(img.shape[0], y1 + region_size)
                target_region = img[y1:y2, x1:x2]

                target_image_filename = (
                    f"{current_program}_{unique_id}_target_{size[0]}x{size[1]}.png"
                )
                target_full_path = os.path.join(save_dir, target_image_filename)
                cv2.imwrite(
                    target_full_path, cv2.cvtColor(target_region, cv2.COLOR_RGB2BGR)
                )
                target_image_paths.append({"path": target_full_path, "size": size})

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
                "target_paths": target_image_paths,
                "image_program_name": current_program,
                "image_program_path": program_path,
                "image_window_class": window_class,
                "image_window_name": window_name,
                "image_depth": depth,
                "wait_for_image": True,
                "wait_method": "image",
                "window_title": window_title,
                "comparison_threshold": 0.6,
            },
        }

        self.update_image_label(img)
        self.capture_thread.hwnd = hwnd
        logging.info(f"Recording click event: {event_info}")
        self.click_events.append(event_info)
        list_item = QListWidgetItem(
            f"Click at ({relative_x}, {relative_y}) in {current_program} ({window_title})"
        )
        list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
        list_item.setCheckState(Qt.Unchecked)
        self.event_list.addItem(list_item)

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
                list_item = QListWidgetItem(
                    f"Click at ({event['relative_x']}, {event['relative_y']}) in {event['program_name']}"
                )
                list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
                list_item.setCheckState(Qt.Unchecked)
                self.event_list.addItem(list_item)
            self.status_bar.showMessage(f"Script loaded from {filename}")

    def play_script(self):
        if not self.click_events:
            return

        self.running_script = True
        self.set_buttons_enabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(self.click_events))

        def process_event(index):
            if not self.running_script:
                self.set_buttons_enabled(True)
                logging.info("Script playback stopped.")
                return

            if index >= len(self.click_events):
                self.set_buttons_enabled(True)
                self.running_script = False
                self.progress_bar.setValue(len(self.click_events))
                logging.info("Script playback completed.")
                return

            event = self.click_events[index]

            if event.get("condition") == "Wait up to 5 minutes for Image":
                start_time = time.time()
                while time.time() - start_time <= 300:
                    try:
                        hwnd = self.find_target_hwnd(event)
                        if hwnd and self.check_image_presence(event, hwnd):
                            logging.info("Image found within 5 minutes.")
                            break
                    except RuntimeError:
                        logging.warning("Target hwnd not found, retrying...")
                    time.sleep(1)
                else:
                    logging.warning(
                        "Image not found within 5 minutes, stopping script."
                    )
                    self.stop_script_execution()
                    return

            repeat_count = event.get("repeat_count", 1)
            for _ in range(repeat_count):
                if not self.process_single_event(event):
                    break

            self.event_list.setCurrentRow(index)
            self.progress_bar.setValue(index + 1)

            logging.info(f"Processed event {index + 1}/{len(self.click_events)}")
            QTimer.singleShot(500, lambda: process_event(index + 1))

        process_event(0)

    def process_single_event(self, event):
        hwnd = self.find_target_hwnd(event)
        if hwnd is None or hwnd == 0:
            logging.warning(f"Failed to find hwnd for event: {event}")
            return False

        if event.get("condition") == "Image Present":
            if not self.check_image_presence(event, hwnd):
                logging.info("Condition 'Image Present' not met, skipping event.")
                return False
        elif event.get("condition") == "Image Not Present":
            if self.check_image_presence(event, hwnd):
                logging.info("Condition 'Image Not Present' not met, skipping event.")
                return False

        click_delay = event.get("click_delay", 0)
        if click_delay > 0:
            time.sleep(click_delay / 1000)

        self.send_click_event(
            event["relative_x"],
            event["relative_y"],
            hwnd,
            event["move_cursor"],
            event.get("double_click", False),
        )

        keyboard_input = event.get("keyboard_input", "")
        if keyboard_input:
            self.send_keyboard_input(keyboard_input, hwnd)

        img = self.capture_thread.capture_window_image(hwnd)
        if img is not None:
            self.update_image_label(img)
        self.capture_thread.hwnd = hwnd

        return True

    def send_click_event(self, relative_x, relative_y, hwnd, move_cursor, double_click):
        if not hwnd or not win32gui.IsWindow(hwnd):
            logging.warning(f"Invalid hwnd: {hwnd}")
            return

        if move_cursor:

            current_x, current_y = win32api.GetCursorPos()
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

            # Determine direction to move (up or down)
            move_direction = 10 if current_y == 0 else -10

            # Calculate new y position
            new_y = max(0, min(screen_height - 1, current_y + move_direction))

            # Move cursor
            win32api.SetCursorPos((current_x, new_y))
            time.sleep(0.1)

        lParam = win32api.MAKELONG(relative_x, relative_y)

        try:
            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
            time.sleep(0.1)
            self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)

            if double_click:
                time.sleep(0.1)
                self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONDOWN)
                time.sleep(0.1)
                self.simulate_mouse_event(hwnd, lParam, WM_LBUTTONUP)

            logging.debug("Click event sent successfully.")
        except Exception as e:
            logging.error(f"Failed to send click event: {e}")

    def send_keyboard_input(self, text, hwnd):
        for char in text:
            win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
            time.sleep(0.05)

    def check_image_presence(self, event, hwnd):
        target_image_paths = event["image"].get("target_paths", [])
        if not target_image_paths:
            return False

        current_image = self.capture_thread.capture_window_image(hwnd)
        if current_image is None:
            return False

        for target_image_info in target_image_paths:
            target_image = cv2.imread(target_image_info["path"], cv2.IMREAD_COLOR)
            if target_image is None:
                continue

            result = cv2.matchTemplate(
                current_image, target_image, cv2.TM_CCOEFF_NORMED
            )
            _, max_val, _, _ = cv2.minMaxLoc(result)

            if max_val >= event["image"].get("comparison_threshold", 0.6):
                return True

        return False

    def move_cursor(self, x, y, hwnd):
        screen_x, screen_y = win32gui.ClientToScreen(hwnd, (x, y))
        win32api.SetCursorPos((screen_x, screen_y))

    def init_emergency_stop(self):
        def on_move(x, y):
            if x == 0 and y == 0:  # Top-left corner
                self.emergency_stop()

        self.mouse_listener = mouse.Listener(on_move=on_move)
        self.mouse_listener.start()

    def emergency_stop(self):
        logging.info("Emergency stop triggered!")
        self.stop_script_execution()
        QMessageBox.warning(
            self, "Emergency Stop", "Script execution has been stopped."
        )

    def stop_script_execution(self):
        self.running_script = False
        self.set_buttons_enabled(True)
        self.progress_bar.setValue(0)
        logging.info("Script execution stopped.")

    def set_buttons_enabled(self, enabled):
        self.record_button.setEnabled(enabled)
        self.save_button.setEnabled(enabled)
        self.load_button.setEnabled(enabled)
        self.play_button.setEnabled(enabled)
        self.stop_button.setEnabled(not enabled)

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

    def update_image_on_selection(self, item):
        index = self.event_list.row(item)
        event = self.click_events[index]
        image_path = event.get("image", {}).get("path", None)
        if image_path and os.path.exists(image_path):
            img = cv2.imread(image_path)
            self.update_image_label(img)

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

    def init_custom_image_ui(self):
        self.add_custom_image_button = self.create_button(
            "Add Custom Image", "icons/custom_image.png", self.add_custom_image
        )
        self.layout.addWidget(self.add_custom_image_button, 7, 0, 1, 2)

    def add_custom_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Image File", "", "Image Files (*.png *.jpg *.bmp)"
        )
        if file_path:
            self.custom_image_path = file_path
            self.status_bar.showMessage(f"Custom image set: {file_path}")
            logging.info(f"Custom image set: {file_path}")

            for index in range(self.event_list.count()):
                item = self.event_list.item(index)
                if item.checkState() == Qt.Checked:
                    event = self.click_events[index]
                    if "custom_images" not in event:
                        event["custom_images"] = []
                    event["custom_images"].append(file_path)

                    self.event_list.item(index).setText(
                        f"{item.text()} [Custom Image Added]"
                    )

    def show_event_settings(self):
        selected_items = self.event_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select an event to edit.")
            return

        selected_item = selected_items[0]
        index = self.event_list.row(selected_item)
        event = self.click_events[index]

        dialog = EventSettingsDialog(event, self)
        if dialog.exec_() == QDialog.Accepted:
            self.status_bar.showMessage(f"Event settings updated for index {index}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    tracker = MouseTracker()
    tracker.show()
    sys.exit(app.exec_())

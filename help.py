import sys
import logging
import win32gui
import win32process
import psutil
import win32api
import win32con
from pynput import mouse
from PyQt5.QtWidgets import (
    QApplication,
    QLabel,
    QVBoxLayout,
    QWidget,
    QTextEdit,
    QMainWindow,
    QPushButton,
)
from PyQt5.QtCore import Qt

logging.basicConfig(
    level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s"
)


class WindowInfoExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()

    def init_ui(self):
        self.setWindowTitle("Window Info Extractor")
        self.setGeometry(100, 100, 800, 600)

        self.info_display = QTextEdit(self)
        self.info_display.setReadOnly(True)
        self.info_display.setGeometry(10, 10, 780, 540)

        self.clear_button = QPushButton("Clear", self)
        self.clear_button.setGeometry(350, 560, 100, 30)
        self.clear_button.clicked.connect(self.clear_display)

    def clear_display(self):
        self.info_display.clear()

    def on_click(self, x, y, button, pressed):
        if pressed:
            hwnd = win32gui.WindowFromPoint((x, y))
            if hwnd:
                info = self.get_window_info(hwnd)
                self.display_info(info)

    def get_window_info(self, hwnd):
        try:
            window_text = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            window_rect = win32gui.GetWindowRect(hwnd)
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            process_name = process.name()
            process_exe = process.exe()
            parent_hwnd = win32gui.GetParent(hwnd)
            menu = win32gui.GetMenu(hwnd)
            children = []
            win32gui.EnumChildWindows(
                hwnd, lambda child_hwnd, param: param.append(child_hwnd), children
            )

            def bool_to_str(value):
                return "Yes" if value else "No"

            def is_maximized(hwnd):
                placement = win32gui.GetWindowPlacement(hwnd)
                return placement[1] == win32con.SW_MAXIMIZE

            info = (
                f"Window Handle (HWND): {hwnd}\n"
                f"Window Title: {window_text}\n"
                f"Class Name: {class_name}\n"
                f"Window Rect: {window_rect}\n"
                f"Style: {style:#010x}\n"
                f"Extended Style: {ex_style:#010x}\n"
                f"Thread ID: {thread_id}\n"
                f"Process ID (PID): {pid}\n"
                f"Process Name: {process_name}\n"
                f"Process Executable: {process_exe}\n"
                f"Parent Window Handle: {parent_hwnd}\n"
                f"Parent Window Title: {win32gui.GetWindowText(parent_hwnd) if parent_hwnd else 'N/A'}\n"
                f"Parent Class Name: {win32gui.GetClassName(parent_hwnd) if parent_hwnd else 'N/A'}\n"
                f"Menu: {menu}\n"
                f"Number of Child Windows: {len(children)}\n"
                f"Is Visible: {bool_to_str(win32gui.IsWindowVisible(hwnd))}\n"
                f"Is Enabled: {bool_to_str(win32gui.IsWindowEnabled(hwnd))}\n"
                f"Is Minimized: {bool_to_str(win32gui.IsIconic(hwnd))}\n"
                f"Is Maximized: {bool_to_str(is_maximized(hwnd))}\n"  # Use GetWindowPlacement to check if maximized
                f"Is Window: {bool_to_str(win32gui.IsWindow(hwnd))}\n"
                f"Is Child: {bool_to_str(win32gui.IsChild(parent_hwnd, hwnd))}\n"
            )

            for i, child in enumerate(children):
                child_text = win32gui.GetWindowText(child)
                child_class = win32gui.GetClassName(child)
                child_rect = win32gui.GetWindowRect(child)
                info += (
                    f"  Child {i + 1} Handle: {child}\n"
                    f"  Child {i + 1} Title: {child_text}\n"
                    f"  Child {i + 1} Class: {child_class}\n"
                    f"  Child {i + 1} Rect: {child_rect}\n"
                )

            return info
        except Exception as e:
            logging.error(f"Error retrieving window info: {e}")
            return f"Error retrieving window info: {e}"

    def display_info(self, info):
        self.info_display.append(info)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WindowInfoExtractor()
    window.show()
    sys.exit(app.exec_())

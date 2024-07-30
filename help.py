import win32gui
import ctypes
from pynput import mouse
from pywinauto import Desktop, Application
from tkinter import Tk, messagebox
import json


def get_window_info(hwnd):
    """주어진 HWND의 윈도우 정보를 가져오는 함수"""
    if hwnd:
        window_text = win32gui.GetWindowText(hwnd)
        window_class = win32gui.GetClassName(hwnd)
        rect = win32gui.GetWindowRect(hwnd)

        # DWM API를 사용하여 확장된 프레임 경계를 가져옴
        rect_extended = ctypes.wintypes.RECT()
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        dwmapi = ctypes.WinDLL("dwmapi")
        dwmapi.DwmGetWindowAttribute(
            hwnd,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect_extended),
            ctypes.sizeof(rect_extended),
        )

        return {
            "HWND": hwnd,
            "Title": window_text,
            "Class": window_class,
            "Position": rect,
            "Extended Frame Bounds": (
                rect_extended.left,
                rect_extended.top,
                rect_extended.right,
                rect_extended.bottom,
            ),
        }
    else:
        return None


def get_control_info(control):
    """주어진 컨트롤의 상세 정보를 가져오는 함수"""
    try:
        rectangle = control.rectangle()
        mid_point = rectangle.mid_point()
        rect_dict = {
            "left": rectangle.left,
            "top": rectangle.top,
            "right": rectangle.right,
            "bottom": rectangle.bottom,
            "mid_point": {"x": mid_point.x, "y": mid_point.y},
        }
    except:
        rect_dict = "N/A"

    info = {
        "Control Type": control.friendly_class_name(),
        "Control Text": control.window_text(),
        "Is Visible": safe_get_attribute(control, "is_visible"),
        "Is Enabled": safe_get_attribute(control, "is_enabled"),
        "Rectangle": rect_dict,
        "Automation Id": safe_get_attribute(control, "automation_id"),
        "Control Id": safe_get_attribute(control, "control_id"),
        "Is Keyboard Focusable": safe_get_attribute(control, "is_keyboard_focusable"),
        "Is Offscreen": safe_get_attribute(control, "is_offscreen"),
        "Control Type (legacy)": control.legacy_properties().get("ControlType", "N/A"),
    }
    return info


def safe_get_attribute(control, attribute):
    """컨트롤 속성을 안전하게 가져오는 함수"""
    try:
        return getattr(control, attribute)()
    except Exception:
        return "N/A"


def dump_control_info(control, depth=0):
    """주어진 컨트롤과 모든 자식 컨트롤의 정보를 재귀적으로 가져오는 함수"""
    details = []
    control_info = get_control_info(control)
    details.append(
        "  " * depth + json.dumps(control_info, indent=4, ensure_ascii=False)
    )
    for child in control.children():
        details.extend(dump_control_info(child, depth + 1))
    return details


def show_target_info(hwnd):
    """타겟 정보를 표시하는 함수"""
    info = get_window_info(hwnd)
    if info:
        app = Desktop(backend="uia").window(handle=hwnd)
        details = dump_control_info(app)

        info_message = (
            f"HWND: {info['HWND']}\n"
            f"Title: {info['Title']}\n"
            f"Class: {info['Class']}\n"
            f"Position: {info['Position']}\n"
            f"Extended Frame Bounds: {info['Extended Frame Bounds']}\n"
            f"Details:\n" + "\n".join(details)
        )

        print(info_message)  # 콘솔에 출력

        root = Tk()
        root.withdraw()  # Tkinter 창 숨기기
        messagebox.showinfo("Target Information", info_message)
        root.destroy()
    else:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Error", "Target not found!")
        root.destroy()


def on_click(x, y, button, pressed):
    """마우스 클릭 이벤트 핸들러"""
    if pressed:
        hwnd = win32gui.WindowFromPoint((x, y))
        show_target_info(hwnd)


def start_mouse_listener():
    """마우스 리스너 시작 함수"""
    listener = mouse.Listener(on_click=on_click)
    listener.start()
    listener.join()


if __name__ == "__main__":
    start_mouse_listener()

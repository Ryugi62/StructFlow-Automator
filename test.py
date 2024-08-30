import win32gui, win32process, win32con, win32api
import time
import threading
from ctypes import windll, c_bool, c_int, WINFUNCTYPE, POINTER, Structure, CFUNCTYPE, c_void_p
import tkinter as tk
from tkinter import ttk
from collections import defaultdict

# Windows API 함수 및 상수 정의
user32 = windll.user32
kernel32 = windll.kernel32
SetWindowsHookEx, UnhookWindowsHookEx, CallNextHookEx = user32.SetWindowsHookExA, user32.UnhookWindowsHookEx, user32.CallNextHookEx
GetModuleHandle, SetWindowPos = kernel32.GetModuleHandleW, user32.SetWindowPos

WH_KEYBOARD_LL, WH_MOUSE_LL = 13, 14
WM_KEYDOWN, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN = 0x0100, 0x0200, 0x0201, 0x0204

# Windows 구조체 정의
class POINT(Structure):
    _fields_ = [("x", c_int), ("y", c_int)]

class MSLLHOOKSTRUCT(Structure):
    _fields_ = [("pt", POINT), ("mouseData", c_int), ("flags", c_int), ("time", c_int), ("dwExtraInfo", POINTER(c_int))]

# 전역 변수
window_list = []
blocked_windows = set()
overlays = {}
keyboard_hook = mouse_hook = None

class OverlayWindow:
    def __init__(self, target_hwnd):
        self.target_hwnd = target_hwnd
        self.overlay = tk.Toplevel()
        self.overlay.overrideredirect(True)
        self.overlay.attributes("-topmost", True, "-alpha", 0.5)
        
        tk.Label(self.overlay, text="BLOCKED", fg="red", bg="yellow", font=("Arial", 24)).pack(expand=True, fill="both")
        
        self.overlay.update()
        self.hwnd = windll.user32.GetParent(self.overlay.winfo_id())
        self.update_position()
        
    def update_position(self):
        try:
            x, y, right, bottom = win32gui.GetWindowRect(self.target_hwnd)
            width, height = right - x, bottom - y
            self.overlay.geometry(f"{width}x{height}+{x}+{y}")
            
            SetWindowPos(self.hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                         win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
            SetWindowPos(self.target_hwnd, self.hwnd, 0, 0, 0, 0,
                         win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
        except Exception as e:
            print(f"Error updating overlay position: {e}")
        
    def show(self):
        self.update_position()
        self.overlay.deiconify()
        
    def hide(self):
        self.overlay.withdraw()

def update_window_list():
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            windows.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True

    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows

def periodic_update():
    global window_list
    while True:
        window_list = update_window_list()
        time.sleep(0.5)

def event_handler(nCode, wParam, lParam, event_type):
    if nCode >= 0:
        if event_type == 'keyboard' and wParam == WM_KEYDOWN:
            hwnd = user32.GetForegroundWindow()
        elif event_type == 'mouse' and wParam in [WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_RBUTTONDOWN]:
            mouse_struct = ctypes.cast(lParam, POINTER(MSLLHOOKSTRUCT)).contents
            hwnd = user32.WindowFromPoint(mouse_struct.pt)
        else:
            return CallNextHookEx(None, nCode, wParam, lParam)
        
        if hwnd in blocked_windows:
            return 1
    return CallNextHookEx(None, nCode, wParam, lParam)

def keyboard_handler(nCode, wParam, lParam):
    return event_handler(nCode, wParam, lParam, 'keyboard')

def mouse_handler(nCode, wParam, lParam):
    return event_handler(nCode, wParam, lParam, 'mouse')

def start_hooks():
    global keyboard_hook, mouse_hook
    keyboard_handler_func = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))(keyboard_handler)
    mouse_handler_func = CFUNCTYPE(c_int, c_int, c_int, POINTER(c_void_p))(mouse_handler)

    keyboard_hook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboard_handler_func, GetModuleHandle(None), 0)
    mouse_hook = SetWindowsHookEx(WH_MOUSE_LL, mouse_handler_func, GetModuleHandle(None), 0)

def stop_hooks():
    global keyboard_hook, mouse_hook
    if keyboard_hook:
        UnhookWindowsHookEx(keyboard_hook)
    if mouse_hook:
        UnhookWindowsHookEx(mouse_hook)

def block_window(hwnd):
    if hwnd not in blocked_windows:
        blocked_windows.add(hwnd)
        overlays[hwnd] = OverlayWindow(hwnd)
        overlays[hwnd].show()
        
        def keep_overlay_updated():
            while hwnd in blocked_windows:
                try:
                    overlays[hwnd].update_position()
                except Exception as e:
                    print(f"Error updating overlay: {e}")
                time.sleep(0.5)

        threading.Thread(target=keep_overlay_updated, daemon=True).start()

def unblock_window(hwnd):
    if hwnd in blocked_windows:
        blocked_windows.remove(hwnd)
        if hwnd in overlays:
            overlays[hwnd].hide()

class WindowBlockerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Window Event Blocker")

        self.tree = ttk.Treeview(root, columns=("Window", "Status", "HWND"), show="headings")
        self.tree.heading("Window", text="Window")
        self.tree.heading("Status", text="Status")
        self.tree.heading("HWND", text="HWND")
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure("Blocked", foreground="red")
        self.tree.tag_configure("Normal", foreground="black")

        tk.Button(root, text="Toggle Block", command=self.toggle_block).pack()

        self.update_tree()

    def update_tree(self):
        selected_hwnds = [self.tree.item(item)["values"][2] for item in self.tree.selection()]

        self.tree.delete(*self.tree.get_children())

        for hwnd, title in window_list:
            status = "Blocked" if hwnd in blocked_windows else "Normal"
            item = self.tree.insert("", "end", values=(title, status, hwnd), tags=(status,))
            if hwnd in selected_hwnds:
                self.tree.selection_add(item)

        self.root.after(500, self.update_tree)

    def toggle_block(self):
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            hwnd = item['values'][2]
            if hwnd in blocked_windows:
                unblock_window(hwnd)
            else:
                block_window(hwnd)

def main():
    threading.Thread(target=periodic_update, daemon=True).start()
    start_hooks()

    root = tk.Tk()
    WindowBlockerGUI(root)
    root.mainloop()

    stop_hooks()

if __name__ == "__main__":
    main()
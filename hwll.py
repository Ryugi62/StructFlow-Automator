import sys
import os
import time
import json
import logging
import cv2
import numpy as np
import win32gui
import win32process
import win32api
import win32con
import psutil
import ctypes
from ctypes import windll, wintypes

# Constants
WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
LOG_FILE = "script_executor_debug.log"

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', filename=LOG_FILE, filemode='w')
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger('').addHandler(console)

def load_script(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            return json.load(file)
    except Exception as e:
        logging.error(f"Error loading script: {e}")
        return None

def is_valid_process(pid, program_name, program_path=None):
    try:
        process = psutil.Process(pid)
        if program_path:
            return process.exe().lower() == program_path.lower()
        return process.name().lower() == program_name.lower()
    except psutil.NoSuchProcess:
        return False

def get_window_depth(hwnd):
    depth = 0
    while hwnd != 0:
        hwnd = win32gui.GetParent(hwnd)
        depth += 1
    return depth

def find_hwnd(program_name, window_class=None, window_name=None, depth=0, program_path=None, window_title=None, window_rect=None, ignore_pos_size=False):
    hwnds = []

    def callback(hwnd, _):
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowEnabled(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return True
        current_depth = get_window_depth(hwnd)
        if depth != 0 and current_depth != depth:
            return True
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
        if is_valid_process(found_pid, program_name, program_path):
            class_name = win32gui.GetClassName(hwnd)
            window_text = win32gui.GetWindowText(hwnd)
            class_match = window_class in class_name if window_class else True
            name_match = window_name in window_text if window_name else True
            title_match = window_title in window_text if window_title else True
            rect_match = True
            if window_rect and not ignore_pos_size:
                rect = win32gui.GetWindowRect(hwnd)
                rect_match = (window_rect[0] == rect[0] and window_rect[1] == rect[1] and
                              window_rect[2] == rect[2] and window_rect[3] == rect[3])
            if class_match and name_match and title_match and rect_match:
                hwnds.append(hwnd)
                logging.debug(f"Found matching window - HWND: {hwnd}, Class: {class_name}, Text: {window_text}")
        return True

    win32gui.EnumWindows(callback, None)
    return hwnds

def send_input_mouse_event(x, y, button, double_click):
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.1)

    input_events = []

    if button == "left":
        down_flag = 0x0002  # MOUSEEVENTF_LEFTDOWN
        up_flag = 0x0004    # MOUSEEVENTF_LEFTUP
    elif button == "right":
        down_flag = 0x0008  # MOUSEEVENTF_RIGHTDOWN
        up_flag = 0x0010    # MOUSEEVENTF_RIGHTUP
    else:
        logging.error(f"Unsupported button: {button}")
        return

    input_events.append(wintypes.INPUT(type=1, ki=wintypes.MOUSEINPUT(dwFlags=down_flag)))
    input_events.append(wintypes.INPUT(type=1, ki=wintypes.MOUSEINPUT(dwFlags=up_flag)))

    if double_click:
        input_events.extend([
            wintypes.INPUT(type=1, ki=wintypes.MOUSEINPUT(dwFlags=down_flag)),
            wintypes.INPUT(type=1, ki=wintypes.MOUSEINPUT(dwFlags=up_flag))
        ])

    nInputs = len(input_events)
    LPINPUT = wintypes.INPUT * nInputs
    pInputs = LPINPUT(*input_events)
    cbSize = ctypes.c_int(ctypes.sizeof(wintypes.INPUT))
    result = ctypes.windll.user32.SendInput(nInputs, pInputs, cbSize)

    if result != nInputs:
        logging.error(f"SendInput failed. Result: {result}, Last error: {ctypes.get_last_error()}")

def send_click_event(relative_x, relative_y, hwnd, move_cursor, double_click, button):
    if not hwnd or not win32gui.IsWindow(hwnd):
        logging.warning(f"Invalid hwnd: {hwnd}")
        return

    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        x = left + relative_x
        y = top + relative_y
        
        logging.debug(f"Sending click event - X: {x}, Y: {y}, Button: {button}, Double Click: {double_click}")
        send_input_mouse_event(x, y, button, double_click)
        logging.debug("Click event sent successfully.")
    except Exception as e:
        logging.error(f"Failed to send click event: {e}")

def send_keyboard_input(text, hwnd):
    for char in text:
        win32api.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
        time.sleep(0.05)

def capture_window_image(hwnd):
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32gui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32gui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        img = np.frombuffer(bmpstr, dtype='uint8')
        img.shape = (height, width, 4)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    except Exception as e:
        logging.error(f"Error capturing window image: {e}")
        return None

def check_image_presence(event, hwnd):
    target_image_paths = event["image"].get("target_paths", [])
    if not target_image_paths:
        return False

    current_image = capture_window_image(hwnd)
    if current_image is None:
        return False

    similarity_threshold = event.get("similarity_threshold", 0.6)
    for target_image_info in target_image_paths:
        target_image = cv2.imread(target_image_info["path"], cv2.IMREAD_COLOR)
        if target_image is None:
            continue

        result = cv2.matchTemplate(current_image, target_image, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(result)

        if max_val >= similarity_threshold:
            logging.debug(f"Image found with similarity: {max_val}")
            return True

    logging.debug("Image not found")
    return False

def process_event(event):
    logging.info(f"Processing event: {event}")
    hwnds = find_hwnd(
        event["program_name"],
        event["window_class"],
        event["window_name"],
        event["depth"],
        event["program_path"],
        event["window_title"],
        event["window_rect"],
        event.get("ignore_pos_size", False)
    )

    if not hwnds:
        logging.warning(f"Failed to find hwnd for event: {event}")
        return False

    hwnd = hwnds[0]  # Take the first hwnd if multiple are found
    logging.info(f"Selected HWND: {hwnd}")

    if event.get("condition") == "이미지 찾을때까지 계속 기다리기":
        logging.info("Waiting for image...")
        start_time = time.time()
        while True:
            if check_image_presence(event, hwnd):
                logging.info("Image found.")
                break
            if time.time() - start_time > 60:  # 1-minute timeout
                logging.warning("Timeout waiting for image.")
                return False
            time.sleep(1)

    if event.get("condition") == "이미지가 있으면 스킵":
        if check_image_presence(event, hwnd):
            logging.info("Image found, skipping click.")
            return True

    if event.get("condition") == "이미지가 없으면 스킵":
        if not check_image_presence(event, hwnd):
            logging.info("Image not found, skipping click.")
            return True

    click_delay = event.get("click_delay", 0)
    if click_delay > 0:
        logging.debug(f"Waiting for {click_delay}ms before click")
        time.sleep(click_delay / 1000)

    send_click_event(
        event["relative_x"],
        event["relative_y"],
        hwnd,
        event["move_cursor"],
        event.get("double_click", False),
        event["button"]
    )

    keyboard_input = event.get("keyboard_input", "")
    if keyboard_input:
        logging.debug(f"Sending keyboard input: {keyboard_input}")
        send_keyboard_input(keyboard_input, hwnd)

    return True

def execute_script(script):
    for i, event in enumerate(script):
        logging.info(f"Executing event {i+1}/{len(script)}")
        if not process_event(event):
            logging.warning(f"Failed to process event: {event}")
        time.sleep(0.5)  # Small delay between events

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script_executor.py <script_file.json>")
        sys.exit(1)

    script_file = sys.argv[1]
    script = load_script(script_file)

    if script:
        print(f"Executing script from {script_file}")
        execute_script(script)
        print("Script execution completed")
    else:
        print("Failed to load script")
import argparse
import win32gui
import win32con
import json
import os

def get_window_info(hwnd):
    """
    주어진 핸들의 창 정보를 가져옵니다.
    """
    if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowEnabled(hwnd) or not win32gui.IsWindowVisible(hwnd):
        return None
    
    rect = win32gui.GetWindowRect(hwnd)
    window_info = {
        "hwnd": hwnd,
        "left": rect[0],
        "top": rect[1],
        "right": rect[2],
        "bottom": rect[3],
        "is_visible": win32gui.IsWindowVisible(hwnd),
        "children": []
    }

    # 자식 창 탐색 및 정보 추가
    def enum_child_proc(child_hwnd, child_list):
        child_info = get_window_info(child_hwnd)
        if child_info:
            child_list.append(child_info)
        return True

    win32gui.EnumChildWindows(hwnd, enum_child_proc, window_info["children"])
    return window_info

def save_window_info(window_info, file_path):
    """
    창의 정보를 JSON 파일로 저장합니다.
    """
    if window_info is not None:
        with open(file_path, "w") as f:
            json.dump(window_info, f, indent=4)
        print(f"Window info saved to {file_path}")

def load_window_info(file_path):
    """
    JSON 파일에서 창 정보를 불러옵니다.
    """
    if os.path.exists(file_path):
        with open(file_path, "r") as f:
            return json.load(f)
    return None

def restore_window_position(hwnd, window_info):
    """
    저장된 정보를 기반으로 창의 위치와 크기를 복원합니다.
    """
    if window_info is None:
        print("No window information to restore.")
        return

    if win32gui.IsWindow(hwnd):
        win32gui.SetWindowPos(
            hwnd,
            0,  # Z-order
            window_info["left"],
            window_info["top"],
            window_info["right"] - window_info["left"],
            window_info["bottom"] - window_info["top"],
            win32con.SWP_NOZORDER  # Z-order를 변경하지 않음
        )
        print(f"Window position restored for HWND {hwnd}")

        # 자식 창 위치 복원
        for child_info in window_info["children"]:
            child_hwnd = win32gui.FindWindowEx(hwnd, None, None, None)
            restore_window_position(child_hwnd, child_info)
    else:
        print(f"Invalid hwnd: {hwnd}")

def find_window_by_title(title):
    return win32gui.FindWindow(None, title)

def main():
    parser = argparse.ArgumentParser(description="Save and restore window positions.")
    parser.add_argument("command", choices=["save", "restore"], help="Command to execute: save or restore")
    parser.add_argument("window_title", help="The title of the window to target")
    parser.add_argument("--file", default="window_info.json", help="The file to save/load window info")

    args = parser.parse_args()

    hwnd = find_window_by_title(args.window_title)
    if hwnd == 0:
        print(f"Window with title '{args.window_title}' not found.")
        return

    if args.command == "save":
        window_info = get_window_info(hwnd)
        if window_info:
            save_window_info(window_info, args.file)
    elif args.command == "restore":
        window_info = load_window_info(args.file)
        restore_window_position(hwnd, window_info)

if __name__ == "__main__":
    main()

import json
import win32gui
import win32process
import psutil
import os


def get_hwnds_for_pid(pid):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
        return True

    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds


def find_child_windows(hwnd):
    def callback(hwnd, hwnds):
        hwnds.append(hwnd)
        return True

    hwnds = []
    win32gui.EnumChildWindows(hwnd, callback, hwnds)
    return hwnds


def find_window_hwnd(program_path, window_class, depth, window_rect):
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            if proc.info["exe"] and os.path.normpath(
                proc.info["exe"].lower()
            ) == os.path.normpath(program_path.lower()):
                print(f"Found process: {proc.info['name']} (PID: {proc.info['pid']})")
                hwnds = get_hwnds_for_pid(proc.info["pid"])
                print(f"Found {len(hwnds)} top-level windows for this process")

                for hwnd in hwnds:
                    print(f"Checking top-level window: {hwnd}")
                    print(f"Window class: {win32gui.GetClassName(hwnd)}")
                    print(f"Window text: {win32gui.GetWindowText(hwnd)}")
                    print(f"Window rect: {win32gui.GetWindowRect(hwnd)}")

                    # Check all child windows
                    child_windows = find_child_windows(hwnd)
                    print(f"Found {len(child_windows)} child windows")

                    for child_hwnd in child_windows:
                        child_class = win32gui.GetClassName(child_hwnd)
                        child_rect = win32gui.GetWindowRect(child_hwnd)

                        # if (
                        #     not win32gui.IsWindow(hwnd)
                        #     or not win32gui.IsWindowEnabled(hwnd)
                        #     or not win32gui.IsWindowVisible(hwnd)
                        # ):

                        if (
                            win32gui.IsWindow(child_hwnd)
                            and win32gui.IsWindowEnabled(child_hwnd)
                            and win32gui.IsWindowVisible(child_hwnd)
                        ):
                            print(f"Child window is enabled and visible")
                        else:
                            print(f"Child window is not enabled and visible")
                            continue

                        print(f"  Child window: {child_hwnd}")
                        print(f"  Child class: {child_class}")
                        print(f"  Child rect: {child_rect}")

                        if child_class == window_class and child_rect == tuple(
                            window_rect
                        ):
                            print(f"Found matching window: {child_hwnd}")
                            return child_hwnd
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return None


def set_window_size(hwnd, width, height):
    # 모니터에서 현재 hwnd의 절대적인 x, y 좌표를 구함
    
    win32gui.MoveWindow(hwnd, 0, 0, width, height, True)



def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_file = os.path.join(script_dir, "json_scripts", "111.json")

    with open(json_file, "r") as f:
        data = json.load(f)

    for item in data:
        program_path = item["program_path"]
        window_class = item["window_class"]
        depth = item["depth"]
        window_rect = item["window_rect"]

        print(f"Searching for window:")
        print(f"  Program: {program_path}")
        print(f"  Window Class: {window_class}")
        print(f"  Depth: {depth}")
        print(f"  Expected Window Rect: {window_rect}")

        hwnd = find_window_hwnd(program_path, window_class, depth, window_rect)

        if hwnd:
            print(f"Found window handle: {hwnd}")
            width, height = (
                window_rect[2] - window_rect[0],
                window_rect[3] - window_rect[1],
            )
            set_window_size(hwnd, 269, 989)
            print(f"Window size set to {width}x{height}")
        else:
            print("Window not found")


if __name__ == "__main__":
    main()

import win32gui, win32api, win32con


def find_child_windows(parent_handle):
    """주어진 부모 창에 대한 모든 자식 창의 핸들을 찾아 반환합니다."""
    child_handles = []

    def enum_child_proc(hwnd, _):
        child_handles.append(hwnd)
        return True

    win32gui.EnumChildWindows(parent_handle, enum_child_proc, None)
    return child_handles


# MIDAS Gen 프로그램의 메인 윈도우 타이틀 설정
main_window_title = "Gen 2024"
main_window_handle = win32gui.FindWindow(None, main_window_title)

if main_window_handle:
    # "MIDAS/Gen" 타이틀을 가진 자식 창만 필터링
    midas_gen_handles = [
        hwnd
        for hwnd in find_child_windows(main_window_handle)
        if win32gui.GetWindowText(hwnd) == "MIDAS/Gen"
    ]

    for hwnd in midas_gen_handles:
        # 클릭할 위치 지정
        lParam = win32api.MAKELONG(842, 78)

        # 백그라운드에서 활성화
        win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
        win32api.Sleep(125)

        # 백그라운드 마우스 이동 이벤트 전송
        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
        win32api.Sleep(125)

        # 마우스 클릭 다운 이벤트 전송
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam)
        win32api.Sleep(75)  # 클릭 유지 시간

        # 마우스 클릭 업 이벤트 전송
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, lParam)
        win32api.Sleep(75)  # 클릭 유지 시간

        # 백그라운드 마우스 이동 이벤트 전송
        win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam)
        win32api.Sleep(125)

else:
    print("지정된 이름의 창을 찾을 수 없습니다.")

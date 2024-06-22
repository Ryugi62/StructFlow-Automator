import win32gui
import win32api
import win32con


def find_child_windows(parent_handle):
    """주어진 부모 창에 대한 모든 자식 창의 핸들을 찾아 반환합니다."""
    child_handles = []

    def enum_child_proc(hwnd, _):
        child_handles.append(hwnd)
        return True

    win32gui.EnumChildWindows(parent_handle, enum_child_proc, None)
    return child_handles


def click_on_child_window(window_title, child_window_text, x, y):
    """주어진 타이틀의 창에서 특정 자식 창을 찾아 주어진 좌표를 클릭합니다."""
    main_window_handle = win32gui.FindWindow(None, window_title)

    if main_window_handle:
        # 지정된 텍스트를 가진 자식 창만 필터링
        target_handles = [
            hwnd
            for hwnd in find_child_windows(main_window_handle)
            if win32gui.GetWindowText(hwnd) == child_window_text
        ]

        for hwnd in target_handles:
            # 클릭할 위치 지정
            lParam = win32api.MAKELONG(x, y)

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
    else:
        print("지정된 이름의 창을 찾을 수 없습니다.")


# 사용 예제
window_title = r"Gen 2024 - [C:\Users\xorjf\Desktop\3764-(에스비일렉트릭)경북 예천군 용궁면 덕계리 380-1, 381 도성기1~5호태양광발전소(축사위)-完\02-(마이다스)구조계산\태양광\[태양광]도성기1~5] - [MIDAS/Gen]"
child_window_text = "MIDAS/Gen"
x = 842
y = 78

click_on_child_window(window_title, child_window_text, x, y)

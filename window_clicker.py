import win32gui
import win32api
import win32con
import traceback


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
    try:
        main_window_handle = win32gui.FindWindow(None, window_title)

        if main_window_handle == 0:
            print("지정된 이름의 창을 찾을 수 없습니다.")
            return

        try:
            target_handles = [
                hwnd
                for hwnd in find_child_windows(main_window_handle)
                if win32gui.GetWindowText(hwnd) == child_window_text
            ]

            if not target_handles:
                print(
                    f"지정된 텍스트 '{child_window_text}'를 가진 자식 창을 찾을 수 없습니다."
                )
                return

            for hwnd in target_handles:
                try:
                    # 클릭할 위치 지정
                    lParam = win32api.MAKELONG(x, y)

                    # 백그라운드에서 활성화
                    win32gui.PostMessage(
                        hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0
                    )
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

                except Exception as e:
                    print(f"자식 창 '{child_window_text}' 클릭 시 오류 발생: {e}")
                    traceback.print_exc()

        except Exception as e:
            print(f"자식 창을 찾는 중 오류 발생: {e}")
            traceback.print_exc()

    except Exception as e:
        print(f"메인 창 '{window_title}'를 찾는 중 오류 발생: {e}")
        traceback.print_exc()


# 테스트 호출
click_on_child_window("메인 창 제목", "자식 창 텍스트", 10, 10)

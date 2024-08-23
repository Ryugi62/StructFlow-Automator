#NoEnv
#SingleInstance, Force
SetWorkingDir %A_ScriptDir%
SetBatchLines, -1
CoordMode, Mouse, Screen

; 전역 변수
global clickPositions := []
global isRecording := false
global guiHwnd := 0

; GUI 생성
Gui, Font, s9, Malgun Gothic
Gui, Add, Button, x10 y10 w100 h30 gStartStopRecording vRecordBtn, 기록 시작
Gui, Add, Button, x120 y10 w100 h30 gPlayback, 재생
Gui, Add, ListView, x10 y50 w570 h200 vClickList, #|X|Y|창 제목|컨트롤|클래스
Gui, Add, StatusBar,, 준비
Gui, Show, w590 h300, 백그라운드 클릭 기록기

; GUI 윈도우 핸들 저장
WinGet, guiHwnd, ID, A

return

GuiClose:
ExitApp

StartStopRecording:
    if (isRecording := !isRecording) {
        clickPositions := []
        GuiControl,, RecordBtn, 기록 중지
        SetTimer, WatchMouse, 10
        SB_SetText("기록 중...")
        LV_Delete()
    } else {
        SetTimer, WatchMouse, Off
        GuiControl,, RecordBtn, 기록 시작
        SB_SetText("기록 완료. " . clickPositions.Length() . "개의 클릭 저장됨.")
    }
return

WatchMouse:
    if (GetKeyState("LButton", "P")) {
        MouseGetPos, x, y, windowHwnd, control, 2
        ; GUI 클릭 무시
        if (windowHwnd != guiHwnd and windowHwnd) {
            WinGetTitle, windowTitle, ahk_id %windowHwnd%
            WinGetClass, windowClass, ahk_id %windowHwnd%
            clickPositions.Push({x: x, y: y, hwnd: windowHwnd, title: windowTitle, control: control, class: windowClass})

            LV_Add("", clickPositions.Length(), x, y, windowTitle, control, windowClass)
        }

        KeyWait, LButton
        Sleep, 100
    }
return

Playback:
    if (clickPositions.Length() = 0) {
        MsgBox, 저장된 클릭이 없습니다.
        return
    }

    SB_SetText("재생 중...")

    for index, pos in clickPositions {
        ; 클라이언트 영역 좌표로 변환
        VarSetCapacity(point, 8, 0)
        NumPut(pos.x, point, 0, "Int")
        NumPut(pos.y, point, 4, "Int")
        DllCall("ScreenToClient", "Ptr", pos.hwnd, "Ptr", &point)
        clientX := NumGet(point, 0, "Int")
        clientY := NumGet(point, 4, "Int")

        ; 윈도우 메시지 전송 (마우스 이동 없이 클릭 이벤트만 전송)
        DllCall("PostMessage", "Ptr", pos.hwnd, "UInt", 0x201, "Ptr", 0, "Ptr", (clientY << 16) | (clientX & 0xFFFF)) ; WM_LBUTTONDOWN
        Sleep, 10
        DllCall("PostMessage", "Ptr", pos.hwnd, "UInt", 0x202, "Ptr", 0, "Ptr", (clientY << 16) | (clientX & 0xFFFF)) ; WM_LBUTTONUP

        LV_Modify(index, "Select")
        Sleep, 100 ; 클릭 간 약간의 지연
    }

    SB_SetText("재생 완료.")
return

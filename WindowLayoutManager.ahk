#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

global targetProgram := ""
global targetWindow := ""
global restoreFile := ""
global isCliMode := false

; 명령줄 인수 처리
if (A_Args.Length() = 3) {
    targetProgram := A_Args[1]
    targetWindow := A_Args[2]
    restoreFile := A_Args[3]
    isCliMode := true
    GoSub, AutoRestore
} else if (A_Args.Length() >= 2) {
    targetWindow := A_Args[1]
    restoreFile := A_Args[2]
    isCliMode := true
    GoSub, AutoRestore
} else {
    GoSub, ShowGUI
}

return

AutoRestore:
    if (FindAndSelectWindow(targetWindow)) {
        RestoreWindowInfo(restoreFile)
    } else {
        MsgBox, 지정된 창을 찾을 수 없습니다: %targetWindow%
    }
ExitApp

ShowGUI:
    Gui, Add, Text,, 대상 프로그램:
    Gui, Add, DropDownList, vSelectedProgram gUpdateSelected
    Gui, Add, Button, gSaveWindowInfo, 위치/크기 저장
    Gui, Add, Button, gRestoreWindowInfo, 위치/크기 복원
    Gui, Add, Button, gRefreshList, 프로그램 목록 새로고침
    Gui, Show,, 창 위치/크기 관리자

    GoSub, RefreshList
return

UpdateSelected:
    Gui, Submit, NoHide
    targetProgram := SelectedProgram
return

SaveWindowInfo:
    if (targetProgram = "") {
        MsgBox, 프로그램을 선택하세요.
        return
    }
    WinGetPos, x, y, w, h, %targetProgram%
    FileSelectFile, file, S16,, 창 정보 저장, INI (*.ini)
    if (file = "")
        return
    IniWrite, %x%, %file%, MainWindow, X
    IniWrite, %y%, %file%, MainWindow, Y
    IniWrite, %w%, %file%, MainWindow, Width
    IniWrite, %h%, %file%, MainWindow, Height

    SaveUIElements(targetProgram, file)

    MsgBox, 창 정보와 내부 UI 요소 정보가 저장되었습니다.
return

RestoreWindowInfo:
    if (!isCliMode) {
        if (targetProgram = "") {
            MsgBox, 프로그램을 선택하세요.
            return
        }
        FileSelectFile, file, 3,, 창 정보 불러오기, INI (*.ini)
        if (file = "")
            return
    } else {
        file := restoreFile
    }
    RestoreWindowInfo(file)
return

RestoreWindowInfo(file) {
    if (!WinExist(targetProgram)) {
        MsgBox, 지정된 창을 찾을 수 없습니다: %targetProgram%
        return
    }

    IniRead, x, %file%, MainWindow, X, ERROR
    IniRead, y, %file%, MainWindow, Y, ERROR
    IniRead, w, %file%, MainWindow, Width, ERROR
    IniRead, h, %file%, MainWindow, Height, ERROR

    if ((x = "ERROR" or y = "ERROR" or w = "ERROR" or h = "ERROR") and !isCliMode) {
        MsgBox, 파일에서 창 정보를 읽을 수 없습니다.
        return
    }

    ; 현재 창 크기 확인
    WinGetPos, currentX, currentY, currentW, currentH, %targetProgram%

    ; 최소 크기 설정 (예: 100x100)
    w := (w < 100) ? 100 : w
    h := (h < 100) ? 100 : h

    WinMove, %targetProgram%,, %x%, %y%, %w%, %h%

    RestoreUIElements(targetProgram, file)

    if (isCliMode) {
        FileAppend, 창 위치/크기 복원 완료 - X: %x%, Y: %y%, W: %w%, H: %h%`n, *
    } else {
        MsgBox, 창 위치와 내부 UI 요소 위치가 복원되었습니다.
    }
}

RefreshList:
    WinGet, id, List,,, Program Manager
    GuiControl,, SelectedProgram, |
    Loop, %id%
    {
        this_id := id%A_Index%
        WinGetTitle, title, ahk_id %this_id%
        if (title != "")
            GuiControl,, SelectedProgram, %title%
    }

    ; 명령줄 인수로 받은 창 자동 선택
    if (targetWindow != "") {
        if (FindAndSelectWindow(targetWindow)) {
            GuiControl, Choose, SelectedProgram, %targetProgram%
        }
    }
return

FindAndSelectWindow(windowName) {
    WinGet, id, List,,, Program Manager
    bestMatch := ""
    bestMatchScore := 0
    Loop, %id%
    {
        this_id := id%A_Index%
        WinGetTitle, title, ahk_id %this_id%
        if (title != "") {
            matchScore := CalculateMatchScore(title, windowName)
            if (matchScore > bestMatchScore) {
                bestMatch := title
                bestMatchScore := matchScore
            }
        }
    }
    if (bestMatch != "") {
        targetProgram := bestMatch
        if (isCliMode) {
            FileAppend, 선택된 창: %targetProgram%`n, *
        }
        return true
    }
return false
}

CalculateMatchScore(str1, str2) {
    score := 0
    Loop, Parse, str2, %A_Space%
    {
        if (InStr(str1, A_LoopField))
            score += StrLen(A_LoopField)
    }
return score
}

SaveUIElements(winTitle, file) {
    WinGet, controlList, ControlList, %winTitle%
    Loop, Parse, controlList, `n
    {
        ControlGetPos, cx, cy, cw, ch, %A_LoopField%, %winTitle%
        IniWrite, %cx%, %file%, UIElements, %A_LoopField%_X
        IniWrite, %cy%, %file%, UIElements, %A_LoopField%_Y
        IniWrite, %cw%, %file%, UIElements, %A_LoopField%_Width
        IniWrite, %ch%, %file%, UIElements, %A_LoopField%_Height
    }
}

RestoreUIElements(winTitle, file) {
    WinGet, controlList, ControlList, %winTitle%
    Loop, Parse, controlList, `n
    {
        IniRead, cx, %file%, UIElements, %A_LoopField%_X, ERROR
        IniRead, cy, %file%, UIElements, %A_LoopField%_Y, ERROR
        IniRead, cw, %file%, UIElements, %A_LoopField%_Width, ERROR
        IniRead, ch, %file%, UIElements, %A_LoopField%_Height, ERROR
        if (cx != "ERROR" and cy != "ERROR" and cw != "ERROR" and ch != "ERROR") {
            ControlMove, %A_LoopField%, %cx%, %cy%, %cw%, %ch%, %winTitle%
        }
    }
}

GuiClose:
ExitApp
#NoEnv
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%

global targetProgram := ""

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

; 내부 UI 요소 정보 저장
SaveUIElements(targetProgram, file)

MsgBox, 창 정보와 내부 UI 요소 정보가 저장되었습니다.
return

RestoreWindowInfo:
if (targetProgram = "") {
    MsgBox, 프로그램을 선택하세요.
    return
}
FileSelectFile, file, 3,, 창 정보 불러오기, INI (*.ini)
if (file = "")
    return
IniRead, x, %file%, MainWindow, X
IniRead, y, %file%, MainWindow, Y
IniRead, w, %file%, MainWindow, Width
IniRead, h, %file%, MainWindow, Height
WinMove, %targetProgram%,, %x%, %y%, %w%, %h%

; 내부 UI 요소 정보 복원
RestoreUIElements(targetProgram, file)

MsgBox, 창 위치와 내부 UI 요소 위치가 복원되었습니다.
return

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
return

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
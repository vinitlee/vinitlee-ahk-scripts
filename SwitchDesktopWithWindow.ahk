#Requires AutoHotkey v2.0

#SingleInstance Force

dllPath := A_ScriptDir "\VirtualDesktopAccessor.dll"
if !FileExist(dllPath) {
    MsgBox "VirtualDesktopAccessor.dll not found in script directory."
    ExitApp
}

; Load the DLL
dll := DllCall("LoadLibrary", "Str", dllPath, "Ptr")

; Hotkeys
#!^Right:: MoveWithWindow(1)
#!^Left:: MoveWithWindow(-1)

#+^Right:: MoveAndRefocus(1)
#+^Left:: MoveAndRefocus(-1)

PosMod(a, b) {
    return Mod(a + b, b)
}

MoveWindowToDesktop(direction) {
    try hwnd := WinGetID("A")
    catch {
        newFocus := RefocusTop()
        return -1
    }
    if !hwnd
        return

    ; Get number of desktops
    count := DllCall(dllPath "\GetDesktopCount", "Int")

    ; Find current desktop index
    currIndex := DllCall(dllPath "\GetWindowDesktopNumber", "Ptr", hwnd)

    ; Calculate new index
    newIndex := PosMod(currIndex + direction, count)

    ; Move window to new desktop
    DllCall(dllPath "\MoveWindowToDesktopNumber", "Ptr", hwnd, "Int", newIndex)

    return newIndex
}

MoveAndRefocus(direction) {
    MoveWindowToDesktop(direction)
    RefocusTop()
}

RefocusTop() {
    thisHwnd := A_ScriptHwnd

    hwnd := DllCall("GetTopWindow", "Ptr", 0, "UPtr")

    while hwnd {
        isVisible := DllCall("IsWindowVisible", "Ptr", hwnd)
        exStyle := DllCall("GetWindowLongPtr", "Ptr", hwnd, "Int", -20, "UPtr")  ; GWL_EXSTYLE = -20
        style := DllCall("GetWindowLongPtr", "Ptr", hwnd, "Int", -16, "UPtr")   ; GWL_STYLE = -16
        hasOwner := DllCall("GetWindow", "Ptr", hwnd, "UInt", 4, "UPtr") != 0   ; GW_OWNER = 4
        class := DllCall("GetWindowProcessHandle", "Ptr", hwnd)

        currentId := DllCall(dllPath "\GetCurrentDesktopNumber")
        windowId := DllCall(dllPath "\GetWindowDesktopNumber", "Ptr", hwnd)

        if currentId = windowId {
            ; Window is on current virtual desktop
            ; OutputDebug(WinGetTitle(hwnd))

            ; 0x10000000 = WS_VISIBLE (already checked by IsWindowVisible)
            ; 0x80 = WS_EX_TOOLWINDOW

            if hwnd != thisHwnd && isVisible && (style & 0x10000000) && !(exStyle & 0x80) && !hasOwner {
                if WinExist(hwnd) {
                    DllCall("SetForegroundWindow", "Ptr", hwnd)

                    return hwnd
                }
            }
        }

        hwnd := DllCall("GetWindow", "Ptr", hwnd, "UInt", 2, "UPtr")  ; GW_HWNDNEXT = 2
    }
}

GoToDesktop(n) {
    ; Switch to new desktop
    DllCall(dllPath "\GoToDesktopNumber", "Int", n)
}

MoveWithWindow(direction) {
    newIndex := MoveWindowToDesktop(direction)
    GoToDesktop(newIndex)
}

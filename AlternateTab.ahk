#Requires AutoHotkey v2.0

; Win+Tab = Focus next window by Z-order (push current to bottom)
#Tab:: CycleForward()

; Win+Shift+Tab = Bring backmost window to front
#+Tab:: CycleBackward()

SetWindowBottom(hwnd) {
    ; HWND_BOTTOM = 1
    ; SWP_NOSIZE | SWP_NOMOVE = 0x0001 | 0x0002 = 0x0003
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 1, "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0003)
}

CycleForward() {
    hwnds := WinGetList()
    if hwnds.Length < 2
        return

    active := WinExist("A")

    for i, hwnd in hwnds {
        if hwnd == active {
            SetWindowBottom(hwnd)
            Sleep(100)

            ; Try each remaining hwnd in list
            loop hwnds.Length - i {
                offset := i + A_Index
                if offset > hwnds.Length
                    break
                nextHwnd := hwnds[offset]
                if !IsRealWindow(nextHwnd)
                    continue
                ShowWindowInfo(nextHwnd)
                WinActivate("ahk_id " nextHwnd)
                return
            }

            MsgBox("No valid window found above current.")
            return
        }
    }
}

CycleBackward() {
    hwnds := WinGetList()
    if hwnds.Length < 2
        return

    ; Iterate from back to front
    loop hwnds.Length {
        i := hwnds.Length - A_Index + 1
        hwnd := hwnds[i]
        if !IsRealWindow(hwnd)
            continue
        if WinExist("ahk_id " hwnd) && !WinActive("ahk_id " hwnd) {
            WinActivate("ahk_id " hwnd)
            return
        }
    }
}

IsRealWindow(hwnd) {
    try {
        class := WinGetClass("ahk_id " hwnd)
        proc := ProcessGetName(WinGetPID("ahk_id " hwnd))
        style := WinGetStyle("ahk_id " hwnd)
        title := WinGetTitle("ahk_id " hwnd)

        ; Skip known phantom classes
        if class = "InternetExplorer_hidden"
            return false

        ; Skip known phantom processes
        if proc = "Adobe Desktop Service.exe"
            return false

        ; Skip invisible windows
        if !(style & 0x10000000)  ; WS_VISIBLE
            return false

        ; Optional: Skip untitled windows
        ; if !title
        ;     return false

        return true
    } catch {
        return false
    }
}

ShowWindowInfo(hwnd) {
    try {
        title := WinGetTitle("ahk_id " hwnd)
        class := WinGetClass("ahk_id " hwnd)
        proc := ProcessGetName(WinGetPID("ahk_id " hwnd))
        min := WinGetMinMax("ahk_id " hwnd)
        visible := WinGetStyle("ahk_id " hwnd) & 0x10000000 ? "Yes" : "No"

        ; Get window rect (position + size)
        WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
        dimensions := w " x " h

        msg := "Switching to:`n"
        msg .= "Title: " title "`n"
        msg .= "Class: " class "`n"
        msg .= "HWND: " hwnd "`n"
        msg .= "Process: " proc "`n"
        msg .= "Minimized: " (min = 1 ? "Yes" : "No") "`n"
        msg .= "Visible: " visible "`n"
        msg .= "Size: " dimensions

        ToolTip(msg)
        SetTimer(() => ToolTip(), -2000)
    } catch as e {
        ToolTip("Error reading window info:`n" e.Message)
    }
}

#Requires AutoHotkey v2.0

global session := Map(
    "active", false,
    "index", 0,
    "hwnds", [],
    "winDown", false
)

; === Main Triggers ===
#Tab:: StartCycle(1)
#+Tab:: StartCycle(-1)

; Detect Win key release to end session
~*LWin up:: EndCycle()
~*RWin up:: EndCycle()

; === Core Functions ===

StartCycle(direction := 1) {
    global session
    while !session["active"] {
        tempList := GetRealWindows()
        FilterWindowList(&tempList)
        if tempList.Length = 0
            return

        session["hwnds"] := tempList
        ; LogWindowList(session["hwnds"], "Filtered Window Stack")
        session["index"] := 0
        session["active"] := true
        session["winDown"] := true
    }

    session["index"] += direction

    listLen := session["hwnds"].Length
    if listLen = 0
        return

    session["index"] := Mod(session["index"], listLen)
    if session["index"] < 0
        session["index"] += listLen

    ActivateFromSession()

}

EndCycle() {
    global session
    if session["active"] {
        session["active"] := false
        session["index"] := 0
        session["hwnds"] := []
        session["winDown"] := false
    }
}

ActivateFromSession() {
    global session

    targetIndex := session["index"]
    hwnds := session["hwnds"]
    target := hwnds[targetIndex + 1]  ; AHK arrays are 1-based

    if !WinExist("ahk_id " target)
        return

    ; Bring target to top, then restore others in original order
    ; Put target last so it ends on top
    otherHwnds := hwnds.Clone()
    otherHwnds.RemoveAt(targetIndex + 1)

    ; Restore original Z-order for all others
    for hwnd in otherHwnds {
        SetWindowBehind(hwnd)
    }

    ; Bring selected window to front
    WinActivate("ahk_id " target)
}

; === Helpers ===

LogWindowList(hwnds, label := "Window List") {
    output := label "`n----------------------`n"
    for i, hwnd in hwnds {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            class := WinGetClass("ahk_id " hwnd)
            proc := ProcessGetName(WinGetPID("ahk_id " hwnd))
            output .= i ": " title "`n"
            output .= "    HWND: " hwnd "`n"
            output .= "    Class: " class "`n"
            output .= "    Process: " proc "`n`n"
        } catch {
            output .= i ": [error reading hwnd " hwnd "]`n`n"
        }
    }
    ; Print to console or MsgBox
    OutputDebug(output)  ; ⬅️ For DebugView or IDE
    MsgBox(output)       ; ⬅️ For visual popup
}

FilterWindowList(&hwnds) {
    newList := []
    for hwnd in hwnds {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            class := WinGetClass("ahk_id " hwnd)
            proc := ProcessGetName(WinGetPID("ahk_id " hwnd))
            style := WinGetStyle("ahk_id " hwnd)
            visible := (style & 0x10000000)

            ; Filter logic — customize here
            if !visible
                continue
            if class = "InternetExplorer_hidden"
                continue
            if class = "Shell_TrayWnd"
                continue
            if class = "Progman"
                continue
            if proc = "Adobe Desktop Service.exe"
                continue
            ; Optional: skip untitled or 0x0 windows
            ; if !title
            ;     continue
            ; WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
            ; if w = 0 and h = 0
            ;     continue

            newList.Push(hwnd)
        } catch {
            continue
        }
    }
    hwnds := newList
}

GetRealWindows() {
    hwnds := WinGetList()
    real := []
    for hwnd in hwnds {
        try {
            title := WinGetTitle("ahk_id " hwnd)
            class := WinGetClass("ahk_id " hwnd)
            proc := ProcessGetName(WinGetPID("ahk_id " hwnd))
            style := WinGetStyle("ahk_id " hwnd)
            visible := (style & 0x10000000)

            if !visible
                continue
            if class = "InternetExplorer_hidden"
                continue
            if proc = "Adobe Desktop Service.exe"
                continue
            ; Optional: ignore untitled windows
            ; if !title
            ;     continue

            real.Push(hwnd)
        } catch {
            continue
        }
    }
    return real
}

SetWindowBehind(hwnd) {
    ; HWND_BOTTOM = 1, SWP_NOSIZE|SWP_NOMOVE|SWP_NOACTIVATE = 0x0013
    DllCall("SetWindowPos", "ptr", hwnd, "ptr", 1, "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x0013)
}

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

    hwnds := session["hwnds"]
    targetIndex := session["index"]
    target := hwnds[targetIndex + 1]

    if !WinExist("ahk_id " target)
        return

    ; Find current topmost window (before we activate new one)
    currentTop := WinGetList()[1]
    if currentTop == target {
        WinActivate("ahk_id " target)  ; still bring to front to ensure focus
        return
    }

    ; Move the window that was frontmost before back to its old Z-order position
    originalPos := 0
    for i, hwnd in hwnds {
        if hwnd == currentTop {
            originalPos := i
            break
        }
    }
    if originalPos {
        ; Insert old front just before the target's position in the frozen list
        ; (This approximates "putting it back" without reordering the full list)
        refHwnd := hwnds[originalPos + 1]  ; hwnd to insert *behind*
        if refHwnd && refHwnd != target
            SetWindowBehindRelative(currentTop, refHwnd)
        else
            SetWindowBehind(currentTop)
    }

    ; Bring the new window to front
    WinActivate("ahk_id " target)
}

; === Helpers ===

SetWindowBehindRelative(hwndToMove, insertBehindHwnd) {
    ; SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE = 0x13
    DllCall("SetWindowPos", "ptr", hwndToMove, "ptr", insertBehindHwnd, "int", 0, "int", 0, "int", 0, "int", 0, "uint",
        0x13)
}

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
            exstyle := DllCall("GetWindowLongPtr", "ptr", hwnd, "int", -20, "ptr")  ; GWL_EXSTYLE = -20
            hasOwner := DllCall("GetWindow", "ptr", hwnd, "uint", 4, "ptr") != 0  ; GW_OWNER = 4

            ; 1. Must be visible
            if !(style & 0x10000000)  ; WS_VISIBLE
                continue

            ; 2. Must not be tool window
            if (exstyle & 0x80)  ; WS_EX_TOOLWINDOW
                continue

            ; 3. Must be top-level (not owned popup)
            if hasOwner
                continue

            ; 4. Optional: must have a title
            ; if !title or Trim(title) = ""
            ;     continue

            ; 5. Optional: must be app window
            ; if !(exstyle & 0x40000)  ; WS_EX_APPWINDOW
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

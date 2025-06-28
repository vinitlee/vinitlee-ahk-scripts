#Requires AutoHotkey v2.0

global session := Map(
    "active", false,
    "index", 0,
    "hwnds", [],
    "winDown", false,
    "lastIndex", 0
)

; === Main Triggers ===
#`:: StartCycle(1)
#+`:: StartCycle(-1)

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

    session["lastIndex"] := session["index"]
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
        session["lastIndex"] := 0
    }
    HideHighlight()
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
    if WinExist("ahk_id " currentTop) {
        if currentTop == target {
            WinActivate("ahk_id " target)  ; still bring to front to ensure focus
            return
        }
    }

    originalPos := session["lastIndex"]
    if originalPos {
        ; refPos := Mod(originalPos - 1 + session["hwnds"].Length, session["hwnds"].Length)
        if session["index"] > session["lastIndex"]  ; forward
            refPos := Mod(originalPos - 1 + hwnds.Length, hwnds.Length)
        else  ; backward
            refPos := Mod(originalPos + 1, hwnds.Length)
        refHwnd := hwnds[refPos + 1]  ; hwnd to insert *behind*
        if refHwnd && refHwnd != target
            SetWindowBehindRelative(currentTop, refHwnd)
        else
            SetWindowBehind(currentTop)
    }

    ; Bring the new window to front
    WinActivate("ahk_id " target)
    ; ShowHighlight(target)
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
    ; OutputDebug(output)  ; ⬅️ For DebugView or IDE
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

; global highlightGui := CreateHighlightGui()

; CreateHighlightGui() {
;     guiRect := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 +E0x80020 +LastFound", "Highlight")
;     guiRect.BackColor := "Red"
;     WinSetTransparent(150, guiRect.Hwnd)
;     return guiRect
; }

; ShowHighlight(hwnd) {
;     global highlightGui

;     if !WinExist("ahk_id " hwnd)
;         return

;     WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
;     border := 10

;     ; Resize and move GUI
;     highlightGui.Move(x - border, y - border, w + 2 * border, h + 2 * border)
;     highlightGui.Show("NA")
; }

; HideHighlight() {
;     global highlightGui
;     try highlightGui.Hide()
; }

global highlightGui := CreateBorderGui()

CreateBorderGui() {
    newGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 +E0x80020 +E0x80000 +LastFound", "Highlight")
    newGui.BackColor := "Fuchsia"  ; placeholder background
    newGui.Opt("+AlwaysOnTop")  ; ensure topmost
    ; WinSetTransparent(50, newGui.Hwnd)  ; nearly invisible base
    hwnd := newGui.Hwnd
    DllCall("SetLayeredWindowAttributes", "ptr", hwnd, "uint", 0xFF00FF, "uint", 0, "uint", 0x1)  ; 0xFF00FF = Fuchsia

    ; Create red border with 4 sides
    newGui.MarginX := 0
    newGui.MarginY := 0
    newGui.SetFont("W0")  ; remove any default padding

    newGui.Add("Text", "x0 y0 w0 h0 BackgroundRed vTopBorder")
    newGui.Add("Text", "x0 y0 w0 h0 BackgroundRed vLeftBorder")
    newGui.Add("Text", "x0 y0 w0 h0 BackgroundRed vRightBorder")
    newGui.Add("Text", "x0 y0 w0 h0 BackgroundRed vBottomBorder")

    return newGui
}

ShowHighlight(hwnd) {
    global highlightGui

    if !WinExist("ahk_id " hwnd)
        return

    WinGetPos(&x, &y, &w, &h, "ahk_id " hwnd)
    border := 3  ; thickness of the border lines

    highlightGui.Move(x - border, y - border, w + border * 2, h + border * 2)

    ; Resize each border
    highlightGui["TopBorder"].Move(0, 0, w + border * 2, border)
    highlightGui["LeftBorder"].Move(0, 0, border, h + border * 2)
    highlightGui["RightBorder"].Move(w + border, 0, border, h + border * 2)
    highlightGui["BottomBorder"].Move(0, h + border, w + border * 2, border)

    ; Reassert AlwaysOnTop in case another window fought us
    WinSetAlwaysOnTop(true, highlightGui.Hwnd)

    highlightGui.Show("NA")
}

HideHighlight() {
    global highlightGui
    try highlightGui.Hide()
}

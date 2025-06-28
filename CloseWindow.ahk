#Requires AutoHotkey v2.0

; Win+Q closes current window
#q:: WinClose("A")

; Optional: Win+W closes current *tab*, useful if apps support Ctrl+W natively
#w:: Send("^w")
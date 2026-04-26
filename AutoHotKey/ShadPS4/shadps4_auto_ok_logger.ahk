#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode(2)

;=========================
; CONFIG
;=========================
logFile     := A_ScriptDir "\shadps4_patch_log.txt"
shadTitle   := "PKG Extraction"
processName := "shadps4.exe"
enabled     := true

;=========================
; STARTUP
;=========================
TrayTip("Script started. Ctrl+Alt+S = toggle, Ctrl+Alt+Q = quit.")
SetTimer(WatchShadDialogs, 200)

;=========================
; HOTKEYS
;=========================
^!s:: {
    global enabled
    enabled := !enabled
    state := enabled ? "ENABLED" : "DISABLED"
    TrayTip("Auto-click is now " state)
}

^!q:: {
    TrayTip("Exiting script...")
    Sleep(500)
    ExitApp()
}

;=========================
; MAIN FUNCTION
;=========================
WatchShadDialogs() {
    global enabled, shadTitle, processName, logFile
    
    if !enabled
        return
    
    try {
        winList := WinGetList(shadTitle)
    } catch {
        return
    }
    
    for hwnd in winList {
        try {
            pName := WinGetProcessName(hwnd)
        } catch {
            continue
        }
        
        if (pName != processName)
            continue
        
        try {
            dlgTitle := WinGetTitle(hwnd)
            dlgText  := ControlGetText("Static2", hwnd)  ; Changed to Static2 for message text
        } catch {
            dlgTitle := "N/A"
            dlgText  := "N/A"
        }
        
        ; Verify this is a PKG Extraction dialog (could be patch, DLC, or game install)
        if (!InStr(dlgText, "PKG"))
            continue
        
        now := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        logEntry := "[" now "] Detected Patch Dialog`r`n"
            . "Window ID: " hwnd "`r`n"
            . "Title    : " dlgTitle "`r`n"
            . "Text     : " dlgText "`r`n"
            . "Action   : Clicked button`r`n"
            . "----------------------------------------`r`n"
        
        try {
            FileAppend(logEntry, logFile, "UTF-8")
        }
        
        ; Try to click either OK or YES button
        buttonClicked := false
        try {
            ; First try Button1 (usually OK or YES)
            if ControlGetText("Button1", hwnd) {
                ControlClick("Button1", hwnd)
                buttonClicked := true
            }
        }
        
        ; If Button1 didn't work, try Button2 (sometimes YES is the second button)
        if (!buttonClicked) {
            try {
                if ControlGetText("Button2", hwnd) {
                    ControlClick("Button2", hwnd)
                    buttonClicked := true
                }
            }
        }
        
        if (buttonClicked)
            Sleep(300)
    }
}
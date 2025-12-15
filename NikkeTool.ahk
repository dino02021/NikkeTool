
#Requires AutoHotkey v2.0
#UseHook  ; 強制鍵鼠熱鍵使用 hook，避免滑鼠鍵在送出點擊時讀不到實體狀態
; ============================================================
; 以系統管理員身分重新啟動
; ============================================================
if !A_IsAdmin {
    try Run('*RunAs "' A_ScriptFullPath '"')
    ExitApp()
}

; ===================================================
; 全域設定與預設值
; ============================================================
SettingsDir      := A_MyDocuments "\NikkeToolSettings"
global SettingsFile := SettingsDir "\NikkeToolSettings.ini"
global AutoStartLink := A_Startup "\NikkeToolStarter.lnk"

; 延遲預設
global EscDelayMs  := 220
global Click1_HoldMs := 225
global Click1_GapMs := 25
global Click2_HoldMs := 240
global Click2_GapMs := 40
global Click3_HoldMs := 240
global Click3_GapMs := 40
global KeySpamDelayMs  := 17

; 預設綁定鍵
global KeySpamD      := "F13"
global KeySpamS      := "F14"
global KeySpamA      := "F15"
global KeyEscDouble  := "F16"
global KeyClick1 := "F17"
global KeyClick2 := "F18"
global KeyClick3 := "F19"

; 功能啟用
global IsSpamDEnabled      := false
global IsSpamSEnabled      := false
global IsSpamAEnabled      := false
global IsEscDoubleEnabled  := false
global IsClick1Enabled := true
global IsClick2Enabled := true
global IsClick3Enabled := false
global ClickBtn1 := "LButton"
global ClickBtn2 := "LButton"
global ClickBtn3 := "LButton"

; 自動啟動
global AutoStartEnabled := false

; QPC
global QPCFreq := 0

; 滑鼠鎖定
global IsCursorLocked := false
global EnableCursorLock := false

; 熱鍵資料
global HotkeyCurrentMap := Map()
global HotkeyHandlerMap := Map()
global HotkeyBaseKeyMap := Map()

; GUI 控制項
global MainGui
global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyClick1, EditKeyClick2, EditKeyClick3
global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkClick1, ChkClick2, ChkClick3
global EditEscDelay, EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
global LblEscDelay, LblClick1Hold, LblClick1Gap, LblClick2Hold, LblClick2Gap, LblClick3Hold, LblClick3Gap
global TxtStatus, ChkAutoStart, TxtNikkeStatus, ChkCursorLock
global DelayCtrlMap := Map()
global FeatureCtrlMap := Map()
global TxtClick1Info, TxtClick2Info, TxtClick3Info, TxtClick1Warn, TxtClick2Warn, TxtClick3Warn
global Click1WarnPosX, Click1WarnPosY, Click2WarnPosX, Click2WarnPosY, Click3WarnPosX, Click3WarnPosY, WarnOffsetXPx
global BtnClick1Side, BtnClick2Side, BtnClick3Side

; 綁定狀態
global IsBinding := false
global BindingActionId := ""
global BindingDisplayCtrl := ""
global BindingInputHook

global AppVersion := "v1.04"

; ============================================================
; 初始化
; ============================================================
Init() {
    global SettingsDir
    if !DirExist(SettingsDir) {
        DirCreate(SettingsDir)
    }

    LoadSettings()
    ApplyAutoStart()
    UpdateAllHotkeys()
    UpdateContextHotkeys()
    BuildGui()
    SetTimer(CursorLockTick, 200)
    OnExit(UnlockCursor)
    A_IconTip := "Nikke小工具 " AppVersion " - Yabi"
}

; ============================================================
; 遊戲前景判斷
; ============================================================
IsNikkeForeground() {
    return WinActive("ahk_exe nikke.exe") ? true : false
}

IsScriptEnabledForContext() {
    return IsNikkeForeground()
}

CursorLockTick(*) {
    global EnableCursorLock
    if EnableCursorLock && IsNikkeForeground() {
        LockCursorToNikke()
    } else {
        UnlockCursor()
    }
}

LockCursorToNikke() {
    global IsCursorLocked
    try {
        hwnd := WinExist("ahk_exe nikke.exe")
        if !hwnd {
            UnlockCursor()
            return
        }
        ; 取客戶區座標，避免碰到邊框 resize 熱區
        WinGetClientPos(&x, &y, &w, &h, "ahk_id " hwnd)
        rect := Buffer(16, 0)
        NumPut("Int", x, rect, 0)
        NumPut("Int", y, rect, 4)
        NumPut("Int", x + w, rect, 8)
        NumPut("Int", y + h, rect, 12)
        if DllCall("ClipCursor", "Ptr", rect.Ptr, "Int") {
            IsCursorLocked := true
        }
    } catch {
        ; 若失敗則解鎖
        UnlockCursor()
    }
}

UnlockCursor(*) {
    global IsCursorLocked
    if IsCursorLocked {
        DllCall("ClipCursor", "Ptr", 0)
        IsCursorLocked := false
    }
}

ToggleCursorLock(state) {
    global EnableCursorLock
    EnableCursorLock := (state != 0)
    SaveKeySettings()
    if EnableCursorLock && IsNikkeForeground() {
        LockCursorToNikke()
    } else {
        UnlockCursor()
    }
}

UpdateContextHotkeys() {
    global HotkeyBaseKeyMap
    wantForeground := IsNikkeForeground()
    passThroughNeeded := !wantForeground
    for id, _ in HotkeyBaseKeyMap {
        ApplyHotkeyState(id, passThroughNeeded)
    }
}

UpdateNikkeStatus(*) {
    global TxtNikkeStatus
    if !IsSet(TxtNikkeStatus)
        return

    if IsNikkeForeground() {
        TxtNikkeStatus.Value := "已偵測到遊戲前景執行中 (nikke.exe)"
        TxtNikkeStatus.Opt("cGreen")
    } else {
        TxtNikkeStatus.Value := "尚未偵測到遊戲或是背景執行中 (nikke.exe)"
        TxtNikkeStatus.Opt("cRed")
    }
    UpdateContextHotkeys()
}

; ============================================================
; 高精度計時
; ============================================================
InitQPC() {
    global QPCFreq
    if (QPCFreq = 0) {
        ok := DllCall("QueryPerformanceFrequency", "Int64*", &QPCFreq)
        if (!ok || QPCFreq = 0) {
            MsgBox("QueryPerformanceFrequency 失敗，QPCFreq = " QPCFreq)
        }
    }
}

BusyWaitMs(ms) {
    global QPCFreq
    InitQPC()
    if (QPCFreq = 0) {
        Sleep(ms)
        return
    }
    start := 0, now := 0
    DllCall("QueryPerformanceCounter", "Int64*", &start)
    target := start + (QPCFreq * ms // 1000)
    while true {
        DllCall("QueryPerformanceCounter", "Int64*", &now)
        if (now >= target)
            break
    }
}

BusyWaitMsCancel(ms, cancelKey) {
    global QPCFreq
    InitQPC()
    if (QPCFreq = 0) {
        Sleep(ms)
        return false
    }
    start := 0, now := 0
    DllCall("QueryPerformanceCounter", "Int64*", &start)
    target := start + (QPCFreq * ms // 1000)
    while true {
        if (cancelKey != "" && !GetKeyState(cancelKey, "P"))
            return false
        DllCall("QueryPerformanceCounter", "Int64*", &now)
        if (now >= target)
            break
    }
    return true
}

WaitMs(ms) => BusyWaitMs(ms)
WaitMsCancel(ms, cancelKey) => BusyWaitMsCancel(ms, cancelKey)

GetClickButtonLabel(btn) {
    return (btn = "RButton") ? "✓ 右鍵" : "✓ 左鍵"
}

SetClickButtonText(seq, btn) {
    global BtnClick1Side, BtnClick2Side, BtnClick3Side
    label := GetClickButtonLabel(btn)
    switch seq {
        case 1:
            if IsSet(BtnClick1Side)
                BtnClick1Side.Text := label
        case 2:
            if IsSet(BtnClick2Side)
                BtnClick2Side.Text := label
        case 3:
            if IsSet(BtnClick3Side)
                BtnClick3Side.Text := label
    }
}

ToggleClickSide(seq) {
    global ClickBtn1, ClickBtn2, ClickBtn3
    btn := "LButton"
    switch seq {
        case 1:
            ClickBtn1 := (ClickBtn1 = "LButton") ? "RButton" : "LButton"
            btn := ClickBtn1
        case 2:
            ClickBtn2 := (ClickBtn2 = "LButton") ? "RButton" : "LButton"
            btn := ClickBtn2
        case 3:
            ClickBtn3 := (ClickBtn3 = "LButton") ? "RButton" : "LButton"
            btn := ClickBtn3
    }
    SetClickButtonText(seq, btn)
    SaveKeySettings()
}

; ============================================================
; 設定載入 / 儲存
; ============================================================
LoadSettings() {
    global SettingsFile, AutoStartLink
    global EscDelayMs, Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyClick1, KeyClick2, KeyClick3
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global ClickBtn1, ClickBtn2, ClickBtn3
    global AutoStartEnabled, EnableCursorLock

    if !FileExist(SettingsFile)
        return

    try EscDelayMs  := Integer(IniRead(SettingsFile, "Delays", "EscDelayMs",  EscDelayMs))
    try Click1_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click1_HoldMs", Click1_HoldMs))
    try Click1_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click1_GapMs", Click1_GapMs))
    try Click2_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click2_HoldMs", Click2_HoldMs))
    try Click2_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click2_GapMs", Click2_GapMs))
    try Click3_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click3_HoldMs", Click3_HoldMs))
    try Click3_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click3_GapMs", Click3_GapMs))

    try KeySpamD      := IniRead(SettingsFile, "Keys", "DSpam",      KeySpamD)
    try KeySpamS      := IniRead(SettingsFile, "Keys", "SSpam",      KeySpamS)
    try KeySpamA      := IniRead(SettingsFile, "Keys", "ASpam",      KeySpamA)
    try KeyEscDouble  := IniRead(SettingsFile, "Keys", "EscDouble",  KeyEscDouble)
    try KeyClick1 := IniRead(SettingsFile, "Keys", "ClickSeq1", KeyClick1)
    try KeyClick2 := IniRead(SettingsFile, "Keys", "ClickSeq2", KeyClick2)
    try KeyClick3 := IniRead(SettingsFile, "Keys", "ClickSeq3", KeyClick3)

    try ClickBtn1 := IniRead(SettingsFile, "Buttons", "ClickSeq1_Button", ClickBtn1)
    try ClickBtn2 := IniRead(SettingsFile, "Buttons", "ClickSeq2_Button", ClickBtn2)
    try ClickBtn3 := IniRead(SettingsFile, "Buttons", "ClickSeq3_Button", ClickBtn3)

    try IsSpamDEnabled      := (Integer(IniRead(SettingsFile, "Enable", "DSpam",      IsSpamDEnabled      ? 1 : 0)) != 0)
    try IsSpamSEnabled      := (Integer(IniRead(SettingsFile, "Enable", "SSpam",      IsSpamSEnabled      ? 1 : 0)) != 0)
    try IsSpamAEnabled      := (Integer(IniRead(SettingsFile, "Enable", "ASpam",      IsSpamAEnabled      ? 1 : 0)) != 0)
    try IsEscDoubleEnabled  := (Integer(IniRead(SettingsFile, "Enable", "EscDouble",  IsEscDoubleEnabled  ? 1 : 0)) != 0)
    try IsClick1Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq1", IsClick1Enabled ? 1 : 0)) != 0)
    try IsClick2Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq2", IsClick2Enabled ? 1 : 0)) != 0)
    try IsClick3Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq3", IsClick3Enabled ? 1 : 0)) != 0)

    try AutoStartEnabled := (Integer(IniRead(SettingsFile, "General", "AutoStart", FileExist(AutoStartLink) ? 1 : 0)) != 0)
    try EnableCursorLock := (Integer(IniRead(SettingsFile, "General", "CursorLock", EnableCursorLock ? 1 : 0)) != 0)
}

SaveKeySettings() {
    global SettingsFile
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyClick1, KeyClick2, KeyClick3
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global AutoStartEnabled, EnableCursorLock
    global EscDelayMs, Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global ClickBtn1, ClickBtn2, ClickBtn3

    IniWrite(EscDelayMs,  SettingsFile, "Delays", "EscDelayMs")
    IniWrite(Click1_HoldMs, SettingsFile, "Delays", "Click1_HoldMs")
    IniWrite(Click1_GapMs, SettingsFile, "Delays", "Click1_GapMs")
    IniWrite(Click2_HoldMs, SettingsFile, "Delays", "Click2_HoldMs")
    IniWrite(Click2_GapMs, SettingsFile, "Delays", "Click2_GapMs")
    IniWrite(Click3_HoldMs, SettingsFile, "Delays", "Click3_HoldMs")
    IniWrite(Click3_GapMs, SettingsFile, "Delays", "Click3_GapMs")

    IniWrite(KeySpamD,      SettingsFile, "Keys", "DSpam")
    IniWrite(KeySpamS,      SettingsFile, "Keys", "SSpam")
    IniWrite(KeySpamA,      SettingsFile, "Keys", "ASpam")
    IniWrite(KeyEscDouble,  SettingsFile, "Keys", "EscDouble")
    IniWrite(KeyClick1, SettingsFile, "Keys", "ClickSeq1")
    IniWrite(KeyClick2, SettingsFile, "Keys", "ClickSeq2")
    IniWrite(KeyClick3, SettingsFile, "Keys", "ClickSeq3")

    IniWrite(ClickBtn1, SettingsFile, "Buttons", "ClickSeq1_Button")
    IniWrite(ClickBtn2, SettingsFile, "Buttons", "ClickSeq2_Button")
    IniWrite(ClickBtn3, SettingsFile, "Buttons", "ClickSeq3_Button")

    IniWrite(IsSpamDEnabled      ? 1 : 0, SettingsFile, "Enable", "DSpam")
    IniWrite(IsSpamSEnabled      ? 1 : 0, SettingsFile, "Enable", "SSpam")
    IniWrite(IsSpamAEnabled      ? 1 : 0, SettingsFile, "Enable", "ASpam")
    IniWrite(IsEscDoubleEnabled  ? 1 : 0, SettingsFile, "Enable", "EscDouble")
    IniWrite(IsClick1Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq1")
    IniWrite(IsClick2Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq2")
    IniWrite(IsClick3Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq3")

    IniWrite(AutoStartEnabled ? 1 : 0, SettingsFile, "General", "AutoStart")
    IniWrite(EnableCursorLock ? 1 : 0, SettingsFile, "General", "CursorLock")
}

ApplyAutoStart() {
    global AutoStartEnabled, AutoStartLink
    if AutoStartEnabled {
        if !FileExist(AutoStartLink) {
            FileCreateShortcut(A_ScriptFullPath, AutoStartLink, A_ScriptDir)
        }
    } else if FileExist(AutoStartLink) {
        FileDelete(AutoStartLink)
    }
}

ToggleAutoStart(state) {
    global AutoStartEnabled
    AutoStartEnabled := (state != 0)
    SaveKeySettings()
    ApplyAutoStart()
}

; ============================================================
; CPS / 顯示與驗證
; ============================================================
UpdateCpsInfo() {
    global Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global TxtClick1Info, TxtClick2Info, TxtClick3Info, TxtClick1Warn, TxtClick2Warn, TxtClick3Warn
    global Click1WarnPosX, Click1WarnPosY, Click2WarnPosX, Click2WarnPosY, Click3WarnPosX, Click3WarnPosY, WarnOffsetXPx

    cycle1 := Click1_HoldMs + Click1_GapMs
    if (cycle1 > 0) {
        cps1 := 1000.0 / cycle1
        TxtClick1Info.Value := Format("連點1：{1} + {2} = {3} ms (約 {4:.2f} 次/秒)", Click1_HoldMs, Click1_GapMs, cycle1, cps1)
        TxtClick1Info.Opt("cBlack")
        if (cps1 > 4.1) {
            TxtClick1Warn.Value := "#超速警告"
            TxtClick1Warn.Visible := true
            TxtClick1Warn.Move(Click1WarnPosX + WarnOffsetXPx, Click1WarnPosY)
        } else {
            TxtClick1Warn.Value := ""
            TxtClick1Warn.Visible := false
        }
    } else {
        TxtClick1Info.Value := "連點1：設定錯誤 (總時間為 0)"
        TxtClick1Info.Opt("cBlack")
        TxtClick1Warn.Value := ""
        TxtClick1Warn.Visible := false
    }

    cycle2 := Click2_HoldMs + Click2_GapMs
    if (cycle2 > 0) {
        cps2 := 1000.0 / cycle2
        TxtClick2Info.Value := Format("連點2：{1} + {2} = {3} ms (約 {4:.2f} 次/秒)", Click2_HoldMs, Click2_GapMs, cycle2, cps2)
        TxtClick2Info.Opt("cBlack")
        if (cps2 > 4.1) {
            TxtClick2Warn.Value := "#超速警告"
            TxtClick2Warn.Visible := true
            TxtClick2Warn.Move(Click2WarnPosX + WarnOffsetXPx, Click2WarnPosY)
        } else {
            TxtClick2Warn.Value := ""
            TxtClick2Warn.Visible := false
        }
    } else {
        TxtClick2Info.Value := "連點2：設定錯誤 (總時間為 0)"
        TxtClick2Info.Opt("cBlack")
        TxtClick2Warn.Value := ""
        TxtClick2Warn.Visible := false
    }

    cycle3 := Click3_HoldMs + Click3_GapMs
    if (cycle3 > 0) {
        cps3 := 1000.0 / cycle3
        TxtClick3Info.Value := Format("連點3：{1} + {2} = {3} ms (約 {4:.2f} 次/秒)", Click3_HoldMs, Click3_GapMs, cycle3, cps3)
        TxtClick3Info.Opt("cBlack")
        if (cps3 > 4.1) {
            TxtClick3Warn.Value := "#超速警告"
            TxtClick3Warn.Visible := true
            TxtClick3Warn.Move(Click3WarnPosX + WarnOffsetXPx, Click3WarnPosY)
        } else {
            TxtClick3Warn.Value := ""
            TxtClick3Warn.Visible := false
        }
    } else {
        TxtClick3Info.Value := "連點3：設定錯誤 (總時間為 0)"
        TxtClick3Info.Opt("cBlack")
        TxtClick3Warn.Value := ""
        TxtClick3Warn.Visible := false
    }
}

UpdateCpsVisibility() {
    global IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global TxtClick1Info, TxtClick2Info, TxtClick3Info, TxtClick1Warn, TxtClick2Warn, TxtClick3Warn

    if IsSet(TxtClick1Info) {
        if IsClick1Enabled {
            TxtClick1Info.Visible := true
        } else {
            TxtClick1Info.Visible := false
            TxtClick1Warn.Visible := false
        }
    }

    if IsSet(TxtClick2Info) {
        if IsClick2Enabled {
            TxtClick2Info.Visible := true
        } else {
            TxtClick2Info.Visible := false
            TxtClick2Warn.Visible := false
        }
    }

    if IsSet(TxtClick3Info) {
        if IsClick3Enabled {
            TxtClick3Info.Visible := true
        } else {
            TxtClick3Info.Visible := false
            TxtClick3Warn.Visible := false
        }
    }
}

UpdateFeatureRowState(id) {
    global FeatureCtrlMap
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled

    if !FeatureCtrlMap.Has(id)
        return

    enabled := true
    switch id {
        case "DSpam":      enabled := IsSpamDEnabled
        case "SSpam":      enabled := IsSpamSEnabled
        case "ASpam":      enabled := IsSpamAEnabled
        case "EscDouble":  enabled := IsEscDoubleEnabled
        case "ClickSeq1":  enabled := IsClick1Enabled
        case "ClickSeq2":  enabled := IsClick2Enabled
        case "ClickSeq3":  enabled := IsClick3Enabled
    }

    for _, ctrl in FeatureCtrlMap[id] {
        try {
            if (ctrl.Type = "Text") {
                ctrl.SetFont(enabled ? "cBlack" : "cGray")
            } else {
                ctrl.Enabled := enabled
            }
        }
    }
}

ApplyFeatureRowStates() {
    global FeatureCtrlMap
    for id, _ in FeatureCtrlMap {
        UpdateFeatureRowState(id)
    }
}

; ============================================================
; 熱鍵綁定與狀態
; ============================================================
UpdateAllHotkeys() {
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyClick1, KeyClick2, KeyClick3
    BindHotkey("DSpam",      KeySpamD,      HandleSpamD)
    BindHotkey("SSpam",      KeySpamS,      HandleSpamS)
    BindHotkey("ASpam",      KeySpamA,      HandleSpamA)
    BindHotkey("EscDouble",  KeyEscDouble,  HandleEscDouble)
    BindHotkey("ClickSeq1", KeyClick1, HandleClick1)
    BindHotkey("ClickSeq2", KeyClick2, HandleClick2)
    BindHotkey("ClickSeq3", KeyClick3, HandleClick3)
}

BindHotkey(id, keyName, func) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global HotkeyHandlerMap, HotkeyBaseKeyMap

    enabled := true
    switch id {
        case "DSpam":       enabled := IsSpamDEnabled
        case "SSpam":       enabled := IsSpamSEnabled
        case "ASpam":       enabled := IsSpamAEnabled
        case "EscDouble":   enabled := IsEscDoubleEnabled
        case "ClickSeq1":  enabled := IsClick1Enabled
        case "ClickSeq2":  enabled := IsClick2Enabled
        case "ClickSeq3":  enabled := IsClick3Enabled
    }

    HotkeyHandlerMap[id] := func
    HotkeyBaseKeyMap[id] := (enabled && keyName != "") ? keyName : ""
    ApplyHotkeyState(id, !IsScriptEnabledForContext())
}

NormalizeHotkeyName(name) {
    while (name != "" && SubStr(name, 1, 1) = "~")
        name := SubStr(name, 2)
    return name
}

ApplyHotkeyState(id, passThrough) {
    global HotkeyBaseKeyMap, HotkeyHandlerMap, HotkeyCurrentMap

    base := HotkeyBaseKeyMap.Has(id) ? HotkeyBaseKeyMap[id] : ""
    handler := HotkeyHandlerMap.Has(id) ? HotkeyHandlerMap[id] : ""
    current := HotkeyCurrentMap.Has(id) ? HotkeyCurrentMap[id] : ""

    if (base = "" || handler = "") {
        if (current != "") {
            try Hotkey(current, "Off")
            HotkeyCurrentMap[id] := ""
        }
        return
    }

    normalized := NormalizeHotkeyName(base)
    isMouseKey := InStr(normalized, "Button")
    ; 統一使用 $ 強制 hook，滑鼠再加 * 避免修飾鍵影響；背景放行時再加 ~
    if isMouseKey {
        newHotkey := passThrough ? "~*$" normalized : "*$" normalized
    } else {
        newHotkey := passThrough ? "~$" normalized : "$" normalized
    }

    if (current = newHotkey)
        return  ; 不重綁相同熱鍵，避免迴圈中被關掉

    if (current != "")
        try Hotkey(current, "Off")

    Hotkey(newHotkey, handler, "On")
    HotkeyCurrentMap[id] := newHotkey
}

; ============================================================
; 熱鍵行為
; ============================================================
HandleSpamD(*) {
    global KeySpamD, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    while GetKeyState(KeySpamD, "P") {
        if !IsScriptEnabledForContext()
            break
        Send "d"
        WaitMs(KeySpamDelayMs)
    }
}

HandleSpamS(*) {
    global KeySpamS, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    while GetKeyState(KeySpamS, "P") {
        if !IsScriptEnabledForContext()
            break
        Send "s"
        WaitMs(KeySpamDelayMs)
    }
}

HandleSpamA(*) {
    global KeySpamA, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    while GetKeyState(KeySpamA, "P") {
        if !IsScriptEnabledForContext()
            break
        Send "a"
        WaitMs(KeySpamDelayMs)
    }
}

HandleEscDouble(*) {
    global EscDelayMs, KeyEscDouble
    if !IsScriptEnabledForContext()
        return
    Send "{Esc}"
    WaitMs(EscDelayMs)
    Send "{Esc}"
}

HandleClick1(*) {
    global Click1_HoldMs, Click1_GapMs, KeyClick1, ClickBtn1
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn1
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        if GetKeyState(eachButton, "P") {
            Send "{" eachButton " up}"
            released := true
        }
    }
    if released
        WaitMs(Click1_GapMs)
    while GetKeyState(KeyClick1, "P") {
        Send btnDown
        if !WaitMsCancel(Click1_HoldMs, KeyClick1) {
            Send btnUp
            break
        }
        Send btnUp
        if !WaitMsCancel(Click1_GapMs, KeyClick1)
            break
    }
}

HandleClick2(*) {
    global Click2_HoldMs, Click2_GapMs, KeyClick2, ClickBtn2
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn2
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        if GetKeyState(eachButton, "P") {
            Send "{" eachButton " up}"
            released := true
        }
    }
    if released
        WaitMs(Click2_GapMs)
    while GetKeyState(KeyClick2, "P") {
        Send btnDown
        if !WaitMsCancel(Click2_HoldMs, KeyClick2) {
            Send btnUp
            break
        }
        Send btnUp
        if !WaitMsCancel(Click2_GapMs, KeyClick2)
            break
    }
}

HandleClick3(*) {
    global Click3_HoldMs, Click3_GapMs, KeyClick3, ClickBtn3
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn3
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        if GetKeyState(eachButton, "P") {
            Send "{" eachButton " up}"
            released := true
        }
    }
    if released
        WaitMs(Click3_GapMs)
    while GetKeyState(KeyClick3, "P") {
        Send btnDown
        if !WaitMsCancel(Click3_HoldMs, KeyClick3) {
            Send btnUp
            break
        }
        Send btnUp
        if !WaitMsCancel(Click3_GapMs, KeyClick3)
            break
    }
}

; ============================================================
; 勾選事件與延遲顯示
; ============================================================
SetFeatureEnabled(id, state) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global TxtStatus

    enabled := (state != 0)
    switch id {
        case "DSpam":      IsSpamDEnabled      := enabled
        case "SSpam":      IsSpamSEnabled      := enabled
        case "ASpam":      IsSpamAEnabled      := enabled
        case "EscDouble":  IsEscDoubleEnabled  := enabled
        case "ClickSeq1": IsClick1Enabled := enabled
        case "ClickSeq2": IsClick2Enabled := enabled
        case "ClickSeq3": IsClick3Enabled := enabled
    }

    SaveKeySettings()
    UpdateAllHotkeys()
    SetDelayControlsEnabled()
    UpdateFeatureRowState(id)
    UpdateCpsInfo()
    UpdateCpsVisibility()
    TxtStatus.Value := id " 已" (enabled ? "啟用" : "停用")
}

SetDelayControlsEnabled() {
    global DelayCtrlMap
    global IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled

    for id, ctrls in DelayCtrlMap {
        visible := true
        switch id {
            case "EscDouble":  visible := IsEscDoubleEnabled
            case "ClickSeq1": visible := IsClick1Enabled
            case "ClickSeq2": visible := IsClick2Enabled
            case "ClickSeq3": visible := IsClick3Enabled
        }
        for _, ctrl in ctrls {
            try {
                if (ctrl.Type = "Edit") {
                    ctrl.Enabled := visible
                } else if (ctrl.Type = "Text") {
                    ctrl.SetFont(visible ? "cBlack" : "cGray")
                }
            }
        }
    }
}

; ============================================================
; 延遲套用與驗證
; ============================================================
ApplyDelayConfig(*) {
    global EscDelayMs, Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global EditEscDelay, EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
    global TxtStatus

    mins := [200, 200, 17, 200, 17, 200, 17]
    labels := ["ESC 延遲", "連點1 按壓時間", "連點1 休息間隔", "連點2 按壓時間", "連點2 休息間隔", "連點3 按壓時間", "連點3 休息間隔"]
    inputs := [EditEscDelay.Value, EditClick1Hold.Value, EditClick1Gap.Value, EditClick2Hold.Value, EditClick2Gap.Value, EditClick3Hold.Value, EditClick3Gap.Value]
    parsed := []

    for idx, val in inputs {
        try v := Integer(val)
        catch {
            TxtStatus.Value := "延遲設定錯誤：" labels[idx] " 需要數字"
            TxtStatus.Opt("cRed")
            return
        }
        if (v < mins[idx]) {
            TxtStatus.Value := "延遲設定錯誤：" labels[idx] " 最低 " mins[idx] " ms"
            TxtStatus.Opt("cRed")
            return
        }
        parsed.Push(v)
    }

    EscDelayMs  := parsed[1]
    Click1_HoldMs := parsed[2]
    Click1_GapMs := parsed[3]
    Click2_HoldMs := parsed[4]
    Click2_GapMs := parsed[5]
    Click3_HoldMs := parsed[6]
    Click3_GapMs := parsed[7]

    SaveKeySettings()
    UpdateCpsInfo()
    UpdateCpsVisibility()
    TxtStatus.Value := "延遲已套用並儲存，下方 CPS 已更新。"
    TxtStatus.Opt("cGreen")
}

; ============================================================
; 綁定鍵捕捉
; ============================================================
StartCaptureBinding(id, ctrl) {
    global IsBinding, BindingActionId, BindingDisplayCtrl, TxtStatus, BindingInputHook
    if IsBinding
        return
    IsBinding      := true
    BindingActionId      := id
    BindingDisplayCtrl := ctrl
    TxtStatus.Value   := "請按要綁定的鍵或滑鼠按鍵..."

    BindingInputHook := InputHook("L1", "")
    BindingInputHook.KeyOpt("{All}", "E")
    BindingInputHook.OnEnd := BindingKeyboardEnd
    BindingInputHook.Start()

    OnMessage(0x0201, WM_LButtonDown)
    OnMessage(0x0204, WM_RBUTTONDOWN)
    OnMessage(0x0207, WM_MBUTTONDOWN)
    OnMessage(0x020B, WM_XBUTTONDOWN)
}

BindingKeyboardEnd(ih, *) {
    global IsBinding
    if !IsBinding
        return
    key := ih.EndKey
    if (key != "")
        FinishBinding(key)
}

WM_LButtonDown(*)    => FinishBinding("LButton")
WM_RBUTTONDOWN(*)    => FinishBinding("RButton")
WM_MBUTTONDOWN(*)    => FinishBinding("MButton")
WM_XBUTTONDOWN(wParam, *) {
    ; WM_XBUTTONDOWN 的高位元組標示按下的是哪顆側鍵：1= XButton1, 2= XButton2
    btn := (((wParam >> 16) & 0xFFFF) == 1) ? "XButton1" : "XButton2"
    FinishBinding(btn)
}

FinishBinding(newKey) {
    global IsBinding, BindingActionId, BindingDisplayCtrl, BindingInputHook
    global TxtStatus
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyClick1, KeyClick2, KeyClick3

    if !IsBinding
        return

    IsBinding := false
    try BindingInputHook.Stop()

    OnMessage(0x0201, WM_LButtonDown, 0)
    OnMessage(0x0204, WM_RBUTTONDOWN, 0)
    OnMessage(0x0207, WM_MBUTTONDOWN, 0)
    OnMessage(0x020B, WM_XBUTTONDOWN, 0)

    switch BindingActionId {
        case "DSpam":      KeySpamD      := newKey
        case "SSpam":      KeySpamS      := newKey
        case "ASpam":      KeySpamA      := newKey
        case "EscDouble":  KeyEscDouble  := newKey
        case "ClickSeq1": KeyClick1 := newKey
        case "ClickSeq2": KeyClick2 := newKey
        case "ClickSeq3": KeyClick3 := newKey
    }

    BindingDisplayCtrl.Value := newKey
    SaveKeySettings()
    UpdateAllHotkeys()
    TxtStatus.Value := "「" BindingActionId "」已綁定為：" newKey
}

; ============================================================
; 資料夾 / 匯出 / 匯入
; ============================================================
OpenSettingsFolder(*) {
    global SettingsDir, TxtStatus
    Run(SettingsDir)
    TxtStatus.Value := "已開啟設定資料夾。"
}

ExportSettings(*) {
    global SettingsFile, TxtStatus
    if !FileExist(SettingsFile)
        SaveKeySettings()
    path := FileSelect("S16", "", "選擇匯出位置", "INI 檔案 (*.ini)")
    if (path = "")
        return
    FileCopy(SettingsFile, path, true)
    TxtStatus.Value := "已匯出設定到：" path
}

ImportSettings(*) {
    global SettingsFile, TxtStatus
    src := FileSelect("16", "", "選擇要匯入的設定檔", "INI 檔案 (*.ini)")
    if (src = "")
        return
    FileCopy(src, SettingsFile, true)
    LoadSettings()
    SaveKeySettings()
    ApplyAutoStart()
    RefreshUI()
    UpdateAllHotkeys()
    TxtStatus.Value := "已匯入設定並套用。"
}

; ============================================================
; GUI
; ============================================================
BuildGui() {
    global MainGui, TxtStatus, ChkAutoStart
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyClick1, EditKeyClick2, EditKeyClick3
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkClick1, ChkClick2, ChkClick3
    global EditEscDelay, EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
    global LblEscDelay, LblClick1Hold, LblClick1Gap, LblClick2Hold, LblClick2Gap, LblClick3Hold, LblClick3Gap
    global DelayCtrlMap, TxtClick1Info, TxtClick2Info, TxtClick3Info, TxtClick1Warn, TxtClick2Warn, TxtClick3Warn
    global Click1WarnPosX, Click1WarnPosY, Click2WarnPosX, Click2WarnPosY, Click3WarnPosX, Click3WarnPosY, WarnOffsetXPx
    global TxtNikkeStatus
    global BtnClick1Side, BtnClick2Side, BtnClick3Side

    MainGui := Gui("+AlwaysOnTop")
    MainGui.Title := "Nikke小工具 " AppVersion " - Yabi"

    MainGui.Add("Text", "xs yp+15", "綁定按鍵 (觸發鍵)：")

    lblD := MainGui.Add("Text", "xs yp+30", "D 連點：")
    EditKeySpamD := MainGui.Add("Edit", "x+24 w90 ReadOnly yp-5", KeySpamD)
    btnBindD  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindD.OnEvent("Click", (*) => StartCaptureBinding("DSpam", EditKeySpamD))
    ChkSpamD := MainGui.Add("CheckBox", (IsSpamDEnabled ? "Checked " : "") "x+5 yp+6", "啟用")
    ChkSpamD.OnEvent("Click", (*) => SetFeatureEnabled("DSpam", ChkSpamD.Value))
    FeatureCtrlMap["DSpam"] := [lblD, EditKeySpamD, btnBindD]

    lblS := MainGui.Add("Text", "xs yp+30", "S 連點：")
    EditKeySpamS := MainGui.Add("Edit", "x+26 w90 ReadOnly yp-5", KeySpamS)
    btnBindS  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindS.OnEvent("Click", (*) => StartCaptureBinding("SSpam", EditKeySpamS))
    ChkSpamS := MainGui.Add("CheckBox", (IsSpamSEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkSpamS.OnEvent("Click", (*) => SetFeatureEnabled("SSpam", ChkSpamS.Value))
    FeatureCtrlMap["SSpam"] := [lblS, EditKeySpamS, btnBindS]

    lblA := MainGui.Add("Text", "xs yp+30", "A 連點：")
    EditKeySpamA := MainGui.Add("Edit", "x+24 w90 ReadOnly yp-5", KeySpamA)
    btnBindA  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindA.OnEvent("Click", (*) => StartCaptureBinding("ASpam", EditKeySpamA))
    ChkSpamA := MainGui.Add("CheckBox", (IsSpamAEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkSpamA.OnEvent("Click", (*) => SetFeatureEnabled("ASpam", ChkSpamA.Value))
    FeatureCtrlMap["ASpam"] := [lblA, EditKeySpamA, btnBindA]

    lblEsc := MainGui.Add("Text", "xs yp+30", "ESC x2：")
    EditKeyEscDouble := MainGui.Add("Edit", "x+23 w90 ReadOnly yp-5", KeyEscDouble)
    btnBindEsc  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindEsc.OnEvent("Click", (*) => StartCaptureBinding("EscDouble", EditKeyEscDouble))
    ChkEscDouble := MainGui.Add("CheckBox", (IsEscDoubleEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkEscDouble.OnEvent("Click", (*) => SetFeatureEnabled("EscDouble", ChkEscDouble.Value))
    FeatureCtrlMap["EscDouble"] := [lblEsc, EditKeyEscDouble, btnBindEsc]

    MainGui.Add("Text", "xs yp+30 w380 h2 0x10", "")

    lblClick1 := MainGui.Add("Text", "xs yp+20", "連點1：")
    EditKeyClick1 := MainGui.Add("Edit", "x+29 w90 ReadOnly yp-5", KeyClick1)
    btnBindL1  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindL1.OnEvent("Click", (*) => StartCaptureBinding("ClickSeq1", EditKeyClick1))
    ChkClick1 := MainGui.Add("CheckBox", (IsClick1Enabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkClick1.OnEvent("Click", (*) => SetFeatureEnabled("ClickSeq1", ChkClick1.Value))
    BtnClick1Side := MainGui.Add("Button", "x+5 w50 yp-5", GetClickButtonLabel(ClickBtn1))
    BtnClick1Side.OnEvent("Click", (*) => ToggleClickSide(1))
    FeatureCtrlMap["ClickSeq1"] := [lblClick1, EditKeyClick1, btnBindL1, BtnClick1Side]

    lblClick2 := MainGui.Add("Text", "xs yp+35", "連點2：")
    EditKeyClick2 := MainGui.Add("Edit", "x+29 w90 ReadOnly yp-5", KeyClick2)
    btnBindL2  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindL2.OnEvent("Click", (*) => StartCaptureBinding("ClickSeq2", EditKeyClick2))
    ChkClick2 := MainGui.Add("CheckBox", (IsClick2Enabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkClick2.OnEvent("Click", (*) => SetFeatureEnabled("ClickSeq2", ChkClick2.Value))
    BtnClick2Side := MainGui.Add("Button", "x+5 w50 yp-5", GetClickButtonLabel(ClickBtn2))
    BtnClick2Side.OnEvent("Click", (*) => ToggleClickSide(2))
    FeatureCtrlMap["ClickSeq2"] := [lblClick2, EditKeyClick2, btnBindL2, BtnClick2Side]

    lblClick3 := MainGui.Add("Text", "xs yp+35", "連點3：")
    EditKeyClick3 := MainGui.Add("Edit", "x+29 w90 ReadOnly yp-5", KeyClick3)
    btnBindL3  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindL3.OnEvent("Click", (*) => StartCaptureBinding("ClickSeq3", EditKeyClick3))
    ChkClick3 := MainGui.Add("CheckBox", (IsClick3Enabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkClick3.OnEvent("Click", (*) => SetFeatureEnabled("ClickSeq3", ChkClick3.Value))
    BtnClick3Side := MainGui.Add("Button", "x+5 w50 yp-5", GetClickButtonLabel(ClickBtn3))
    BtnClick3Side.OnEvent("Click", (*) => ToggleClickSide(3))
    FeatureCtrlMap["ClickSeq3"] := [lblClick3, EditKeyClick3, btnBindL3, BtnClick3Side]

    MainGui.Add("Text", "xs yp+40 w380 h2 0x10", "")
    MainGui.Add("Text", "xs yp+20", "延遲設定：")

    LblEscDelay := MainGui.Add("Text", "xs yp+30", "ESC：兩次 ESC 中間延遲 (ms)")
    EditEscDelay  := MainGui.Add("Edit", "w120", EscDelayMs)

    LblClick1Hold := MainGui.Add("Text", , "連點1：按壓時間 (ms)")
    EditClick1Hold  := MainGui.Add("Edit", "w120", Click1_HoldMs)
    LblClick1Gap := MainGui.Add("Text", , "連點1：休息間隔 (ms)")
    EditClick1Gap  := MainGui.Add("Edit", "w120", Click1_GapMs)

    LblClick2Hold := MainGui.Add("Text", , "連點2：按壓時間 (ms)")
    EditClick2Hold  := MainGui.Add("Edit", "w120", Click2_HoldMs)
    LblClick2Gap := MainGui.Add("Text", , "連點2：休息間隔 (ms)")
    EditClick2Gap  := MainGui.Add("Edit", "w120", Click2_GapMs)

    LblClick3Hold := MainGui.Add("Text", , "連點3：按壓時間 (ms)")
    EditClick3Hold  := MainGui.Add("Edit", "w120", Click3_HoldMs)
    LblClick3Gap := MainGui.Add("Text", , "連點3：休息間隔 (ms)")
    EditClick3Gap  := MainGui.Add("Edit", "w120", Click3_GapMs)

    DelayCtrlMap["EscDouble"]  := [LblEscDelay, EditEscDelay]
    DelayCtrlMap["ClickSeq1"] := [LblClick1Hold, EditClick1Hold, LblClick1Gap, EditClick1Gap]
    DelayCtrlMap["ClickSeq2"] := [LblClick2Hold, EditClick2Hold, LblClick2Gap, EditClick2Gap]
    DelayCtrlMap["ClickSeq3"] := [LblClick3Hold, EditClick3Hold, LblClick3Gap, EditClick3Gap]
    SetDelayControlsEnabled()

    btnApply := MainGui.Add("Button", "w120", "套用延遲")
    btnApply.OnEvent("Click", ApplyDelayConfig)

    TxtStatus := MainGui.Add("Text", "w380 cGreen", "")

    TxtClick1Info := MainGui.Add("Text", "xs yp+20 w235", "")
    TxtClick1Warn := MainGui.Add("Text", "xp+237 yp w180 cRed", "")
    TxtClick2Info := MainGui.Add("Text", "xs yp+20 w235", "")
    TxtClick2Warn := MainGui.Add("Text", "xp+237 yp w180 cRed", "")
    TxtClick3Info := MainGui.Add("Text", "xs yp+20 w235", "")
    TxtClick3Warn := MainGui.Add("Text", "xp+237 yp w180 cRed", "")
    WarnOffsetXPx := 0
    TxtClick1Warn.GetPos(&Click1WarnPosX, &Click1WarnPosY)
    TxtClick2Warn.GetPos(&Click2WarnPosX, &Click2WarnPosY)
    TxtClick3Warn.GetPos(&Click3WarnPosX, &Click3WarnPosY)
    UpdateCpsInfo()
    UpdateCpsVisibility()
    ApplyFeatureRowStates()

    MainGui.Add("Text", "xs yp+25 w380 h2 0x10", "")

    ChkAutoStart := MainGui.Add("CheckBox", "xs yp+20", "開機時自動啟動")
    ChkAutoStart.Value := AutoStartEnabled ? 1 : 0
    ChkAutoStart.OnEvent("Click", (*) => ToggleAutoStart(ChkAutoStart.Value))

    ChkCursorLock := MainGui.Add("CheckBox", "xs yp+20", "遊戲中鎖定滑鼠鼠標")
    ChkCursorLock.Value := EnableCursorLock ? 1 : 0
    ChkCursorLock.OnEvent("Click", (*) => ToggleCursorLock(ChkCursorLock.Value))

    btnOpenFolder := MainGui.Add("Button", "xs yp+25 w140", "開啟設定資料夾")
    btnOpenFolder.OnEvent("Click", OpenSettingsFolder)
    btnExport := MainGui.Add("Button", "x+10 w90", "匯出設定")
    btnExport.OnEvent("Click", ExportSettings)
    btnImport := MainGui.Add("Button", "x+10 w90", "匯入設定")
    btnImport.OnEvent("Click", ImportSettings)

    MainGui.Add("Text", "xs yp+40 w380 h2 0x10", "")
    TxtNikkeStatus := MainGui.Add("Text", "xs yp+10 w380", "")
    UpdateNikkeStatus()
    SetTimer(UpdateNikkeStatus, 500)

    MainGui.Show()
}

RefreshUI() {
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyClick1, KeyClick2, KeyClick3
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global EscDelayMs, Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global AutoStartEnabled
    global BtnClick1Side, BtnClick2Side, BtnClick3Side
    global ClickBtn1, ClickBtn2, ClickBtn3
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyClick1, EditKeyClick2, EditKeyClick3
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkClick1, ChkClick2, ChkClick3
    global EditEscDelay, EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
    global ChkAutoStart, TxtStatus, ChkCursorLock, EnableCursorLock

    EditKeySpamD.Value   := KeySpamD
    EditKeySpamS.Value   := KeySpamS
    EditKeySpamA.Value   := KeySpamA
    EditKeyEscDouble.Value := KeyEscDouble
    EditKeyClick1.Value  := KeyClick1
    EditKeyClick2.Value  := KeyClick2
    EditKeyClick3.Value  := KeyClick3

    ChkSpamD.Value   := IsSpamDEnabled      ? 1 : 0
    ChkSpamS.Value   := IsSpamSEnabled      ? 1 : 0
    ChkSpamA.Value   := IsSpamAEnabled      ? 1 : 0
    ChkEscDouble.Value := IsEscDoubleEnabled  ? 1 : 0
    ChkClick1.Value  := IsClick1Enabled ? 1 : 0
    ChkClick2.Value  := IsClick2Enabled ? 1 : 0
    ChkClick3.Value  := IsClick3Enabled ? 1 : 0
    SetClickButtonText(1, ClickBtn1)
    SetClickButtonText(2, ClickBtn2)
    SetClickButtonText(3, ClickBtn3)

    EditEscDelay.Value  := EscDelayMs
    EditClick1Hold.Value := Click1_HoldMs
    EditClick1Gap.Value := Click1_GapMs
    EditClick2Hold.Value := Click2_HoldMs
    EditClick2Gap.Value := Click2_GapMs
    EditClick3Hold.Value := Click3_HoldMs
    EditClick3Gap.Value := Click3_GapMs

    ChkAutoStart.Value := AutoStartEnabled ? 1 : 0
    UpdateCpsInfo()
    UpdateCpsVisibility()
    UpdateNikkeStatus()
    ApplyFeatureRowStates()
    TxtStatus.Value := "已重新載入設定。"
}

; ============================================================
; GUI 快捷鍵與托盤
; ============================================================
^!g::ShowGui()
ShowGui(*) {
    global MainGui
    MainGui.Show()
}

A_TrayMenu.Delete()
A_TrayMenu.Add("開啟控制面板 (&O)", ShowGui)
A_TrayMenu.Add()
A_TrayMenu.Add("退出 (&X)", (*) => ExitApp())
A_TrayMenu.Default := "開啟控制面板 (&O)"  ; 雙擊托盤圖示直接開面板

; ============================================================
; 執行
; ============================================================
Init()

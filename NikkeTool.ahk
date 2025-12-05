#Requires AutoHotkey v2.0

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
global LClick1_HoldMs := 225
global LClick1_GapMs := 25
global LClick2_HoldMs := 240
global LClick2_GapMs := 40
global KeySpamDelayMs  := 17

; 預設綁定鍵
global KeySpamD      := "F13"
global KeySpamS      := "F14"
global KeySpamA      := "F15"
global KeyEscDouble  := "F16"
global KeyLClick1 := "F17"
global KeyLClick2 := "F18"

; 功能啟用
global IsSpamDEnabled      := false
global IsSpamSEnabled      := false
global IsSpamAEnabled      := false
global IsEscDoubleEnabled  := false
global IsLClick1Enabled := true
global IsLClick2Enabled := true

; 自動啟動
global AutoStartEnabled := false

; QPC
global QPCFreq := 0

; 熱鍵資料
global HotkeyCurrentMap := Map()
global HotkeyHandlerMap := Map()
global HotkeyBaseKeyMap := Map()

; GUI 控制項
global MainGui
global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyLClick1, EditKeyLClick2
global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkLClick1, ChkLClick2
global EditEscDelay, EditLClick1Hold, EditLClick1Gap, EditLClick2Hold, EditLClick2Gap
global LblEscDelay, LblLClick1Hold, LblLClick1Gap, LblLClick2Hold, LblLClick2Gap
global TxtStatus, ChkAutoStart, TxtNikkeStatus
global DelayCtrlMap := Map()
global TxtLClick1Info, TxtLClick2Info, TxtLClick1Warn, TxtLClick2Warn
global LClick1WarnPosX, LClick1WarnPosY, LClick2WarnPosX, LClick2WarnPosY, WarnOffsetXPx

; 綁定狀態
global IsBinding := false
global BindingActionId := ""
global BindingDisplayCtrl := ""
global BindingInputHook

global AppVersion := "v1.0"

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
    A_IconTip := "Nikke小工具 " AppVersion " - Yabi"
}

; ============================================================
; 遊戲前景判斷
; ============================================================
IsNikkeForeground() {
    try {
        return WinGetProcessName("A") = "nikke.exe"
    } catch {
        return false
    }
}

IsScriptEnabledForContext() {
    return IsNikkeForeground()
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

; ============================================================
; 設定載入 / 儲存
; ============================================================
LoadSettings() {
    global SettingsFile, AutoStartLink
    global EscDelayMs, LClick1_HoldMs, LClick1_GapMs, LClick2_HoldMs, LClick2_GapMs
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyLClick1, KeyLClick2
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled
    global AutoStartEnabled

    if !FileExist(SettingsFile)
        return

    try EscDelayMs  := Integer(IniRead(SettingsFile, "Delays", "EscDelayMs",  EscDelayMs))
    try LClick1_HoldMs := Integer(IniRead(SettingsFile, "Delays", "LClick1_HoldMs", LClick1_HoldMs))
    try LClick1_GapMs := Integer(IniRead(SettingsFile, "Delays", "LClick1_GapMs", LClick1_GapMs))
    try LClick2_HoldMs := Integer(IniRead(SettingsFile, "Delays", "LClick2_HoldMs", LClick2_HoldMs))
    try LClick2_GapMs := Integer(IniRead(SettingsFile, "Delays", "LClick2_GapMs", LClick2_GapMs))

    try KeySpamD      := IniRead(SettingsFile, "Keys", "DSpam",      KeySpamD)
    try KeySpamS      := IniRead(SettingsFile, "Keys", "SSpam",      KeySpamS)
    try KeySpamA      := IniRead(SettingsFile, "Keys", "ASpam",      KeySpamA)
    try KeyEscDouble  := IniRead(SettingsFile, "Keys", "EscDouble",  KeyEscDouble)
    try KeyLClick1 := IniRead(SettingsFile, "Keys", "LClickSeq1", KeyLClick1)
    try KeyLClick2 := IniRead(SettingsFile, "Keys", "LClickSeq2", KeyLClick2)

    try IsSpamDEnabled      := (Integer(IniRead(SettingsFile, "Enable", "DSpam",      IsSpamDEnabled      ? 1 : 0)) != 0)
    try IsSpamSEnabled      := (Integer(IniRead(SettingsFile, "Enable", "SSpam",      IsSpamSEnabled      ? 1 : 0)) != 0)
    try IsSpamAEnabled      := (Integer(IniRead(SettingsFile, "Enable", "ASpam",      IsSpamAEnabled      ? 1 : 0)) != 0)
    try IsEscDoubleEnabled  := (Integer(IniRead(SettingsFile, "Enable", "EscDouble",  IsEscDoubleEnabled  ? 1 : 0)) != 0)
    try IsLClick1Enabled := (Integer(IniRead(SettingsFile, "Enable", "LClickSeq1", IsLClick1Enabled ? 1 : 0)) != 0)
    try IsLClick2Enabled := (Integer(IniRead(SettingsFile, "Enable", "LClickSeq2", IsLClick2Enabled ? 1 : 0)) != 0)

    try AutoStartEnabled := (Integer(IniRead(SettingsFile, "General", "AutoStart", FileExist(AutoStartLink) ? 1 : 0)) != 0)
}

SaveKeySettings() {
    global SettingsFile
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyLClick1, KeyLClick2
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled
    global AutoStartEnabled
    global EscDelayMs, LClick1_HoldMs, LClick1_GapMs, LClick2_HoldMs, LClick2_GapMs

    IniWrite(EscDelayMs,  SettingsFile, "Delays", "EscDelayMs")
    IniWrite(LClick1_HoldMs, SettingsFile, "Delays", "LClick1_HoldMs")
    IniWrite(LClick1_GapMs, SettingsFile, "Delays", "LClick1_GapMs")
    IniWrite(LClick2_HoldMs, SettingsFile, "Delays", "LClick2_HoldMs")
    IniWrite(LClick2_GapMs, SettingsFile, "Delays", "LClick2_GapMs")

    IniWrite(KeySpamD,      SettingsFile, "Keys", "DSpam")
    IniWrite(KeySpamS,      SettingsFile, "Keys", "SSpam")
    IniWrite(KeySpamA,      SettingsFile, "Keys", "ASpam")
    IniWrite(KeyEscDouble,  SettingsFile, "Keys", "EscDouble")
    IniWrite(KeyLClick1, SettingsFile, "Keys", "LClickSeq1")
    IniWrite(KeyLClick2, SettingsFile, "Keys", "LClickSeq2")

    IniWrite(IsSpamDEnabled      ? 1 : 0, SettingsFile, "Enable", "DSpam")
    IniWrite(IsSpamSEnabled      ? 1 : 0, SettingsFile, "Enable", "SSpam")
    IniWrite(IsSpamAEnabled      ? 1 : 0, SettingsFile, "Enable", "ASpam")
    IniWrite(IsEscDoubleEnabled  ? 1 : 0, SettingsFile, "Enable", "EscDouble")
    IniWrite(IsLClick1Enabled ? 1 : 0, SettingsFile, "Enable", "LClickSeq1")
    IniWrite(IsLClick2Enabled ? 1 : 0, SettingsFile, "Enable", "LClickSeq2")

    IniWrite(AutoStartEnabled ? 1 : 0, SettingsFile, "General", "AutoStart")
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
    global LClick1_HoldMs, LClick1_GapMs, LClick2_HoldMs, LClick2_GapMs
    global TxtLClick1Info, TxtLClick2Info, TxtLClick1Warn, TxtLClick2Warn
    global LClick1WarnPosX, LClick1WarnPosY, LClick2WarnPosX, LClick2WarnPosY, WarnOffsetXPx

    cycle1 := LClick1_HoldMs + LClick1_GapMs
    if (cycle1 > 0) {
        cps1 := 1000.0 / cycle1
        TxtLClick1Info.Value := Format("左鍵連點1：{1} + {2} = {3} ms (約 {4:.2f} 次/秒)", LClick1_HoldMs, LClick1_GapMs, cycle1, cps1)
        TxtLClick1Info.Opt("cBlack")
        if (cps1 > 4.1) {
            TxtLClick1Warn.Value := "#超速警告"
            TxtLClick1Warn.Visible := true
            TxtLClick1Warn.Move(LClick1WarnPosX + WarnOffsetXPx, LClick1WarnPosY)
        } else {
            TxtLClick1Warn.Value := ""
            TxtLClick1Warn.Visible := false
        }
    } else {
        TxtLClick1Info.Value := "左鍵連點1：設定錯誤 (總時間為 0)"
        TxtLClick1Info.Opt("cBlack")
        TxtLClick1Warn.Value := ""
        TxtLClick1Warn.Visible := false
    }

    cycle2 := LClick2_HoldMs + LClick2_GapMs
    if (cycle2 > 0) {
        cps2 := 1000.0 / cycle2
        TxtLClick2Info.Value := Format("左鍵連點2：{1} + {2} = {3} ms (約 {4:.2f} 次/秒)", LClick2_HoldMs, LClick2_GapMs, cycle2, cps2)
        TxtLClick2Info.Opt("cBlack")
        if (cps2 > 4.1) {
            TxtLClick2Warn.Value := "#超速警告"
            TxtLClick2Warn.Visible := true
            TxtLClick2Warn.Move(LClick2WarnPosX + WarnOffsetXPx, LClick2WarnPosY)
        } else {
            TxtLClick2Warn.Value := ""
            TxtLClick2Warn.Visible := false
        }
    } else {
        TxtLClick2Info.Value := "左鍵連點2：設定錯誤 (總時間為 0)"
        TxtLClick2Info.Opt("cBlack")
        TxtLClick2Warn.Value := ""
        TxtLClick2Warn.Visible := false
    }
}

UpdateCpsVisibility() {
    global IsLClick1Enabled, IsLClick2Enabled
    global TxtLClick1Info, TxtLClick2Info, TxtLClick1Warn, TxtLClick2Warn

    if IsSet(TxtLClick1Info) {
        if IsLClick1Enabled {
            TxtLClick1Info.Visible := true
        } else {
            TxtLClick1Info.Visible := false
            TxtLClick1Warn.Visible := false
        }
    }

    if IsSet(TxtLClick2Info) {
        if IsLClick2Enabled {
            TxtLClick2Info.Visible := true
        } else {
            TxtLClick2Info.Visible := false
            TxtLClick2Warn.Visible := false
        }
    }
}

; ============================================================
; 熱鍵綁定與狀態
; ============================================================
UpdateAllHotkeys() {
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyLClick1, KeyLClick2
    BindHotkey("DSpam",      KeySpamD,      HandleSpamD)
    BindHotkey("SSpam",      KeySpamS,      HandleSpamS)
    BindHotkey("ASpam",      KeySpamA,      HandleSpamA)
    BindHotkey("EscDouble",  KeyEscDouble,  HandleEscDouble)
    BindHotkey("LClickSeq1", KeyLClick1, HandleLClick1)
    BindHotkey("LClickSeq2", KeyLClick2, HandleLClick2)
}

BindHotkey(id, keyName, func) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled
    global HotkeyHandlerMap, HotkeyBaseKeyMap

    enabled := true
    switch id {
        case "DSpam":       enabled := IsSpamDEnabled
        case "SSpam":       enabled := IsSpamSEnabled
        case "ASpam":       enabled := IsSpamAEnabled
        case "EscDouble":   enabled := IsEscDoubleEnabled
        case "LClickSeq1":  enabled := IsLClick1Enabled
        case "LClickSeq2":  enabled := IsLClick2Enabled
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
    newHotkey := passThrough ? "~" normalized : normalized

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

HandleLClick1(*) {
    global LClick1_HoldMs, LClick1_GapMs, KeyLClick1
    if !IsScriptEnabledForContext()
        return
    ; 若左鍵已被按住，先補一次放開，避免卡住
    if GetKeyState("LButton", "P") {
        Send "{LButton up}"
        WaitMs(LClick1_GapMs)  ; 放掉後補休息間隔
    }
    if GetKeyState(KeyLClick1, "P") {
        while GetKeyState(KeyLClick1, "P") {
            if !IsScriptEnabledForContext()
                break
            Send "{LButton down}"
            WaitMs(LClick1_HoldMs)
            Send "{LButton up}"
            WaitMs(LClick1_GapMs)
        }
    }
}

HandleLClick2(*) {
    global LClick2_HoldMs, LClick2_GapMs, KeyLClick2
    if !IsScriptEnabledForContext()
        return
    ; 若左鍵已被按住，先補一次放開，避免卡住
    if GetKeyState("LButton", "P") {
        Send "{LButton up}"
        WaitMs(LClick2_GapMs)  ; 放掉後補休息間隔
    }
    if GetKeyState(KeyLClick2, "P") {
        while GetKeyState(KeyLClick2, "P") {
            if !IsScriptEnabledForContext()
                break
            Send "{LButton down}"
            WaitMs(LClick2_HoldMs)
            Send "{LButton up}"
            WaitMs(LClick2_GapMs)
        }
    }
}

; ============================================================
; 勾選事件與延遲顯示
; ============================================================
SetFeatureEnabled(id, state) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled
    global TxtStatus

    enabled := (state != 0)
    switch id {
        case "DSpam":      IsSpamDEnabled      := enabled
        case "SSpam":      IsSpamSEnabled      := enabled
        case "ASpam":      IsSpamAEnabled      := enabled
        case "EscDouble":  IsEscDoubleEnabled  := enabled
        case "LClickSeq1": IsLClick1Enabled := enabled
        case "LClickSeq2": IsLClick2Enabled := enabled
    }

    SaveKeySettings()
    UpdateAllHotkeys()
    SetDelayControlsEnabled()
    UpdateCpsInfo()
    UpdateCpsVisibility()
    TxtStatus.Value := id " 已" (enabled ? "啟用" : "停用")
}

SetDelayControlsEnabled() {
    global DelayCtrlMap
    global IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled

    for id, ctrls in DelayCtrlMap {
        visible := true
        switch id {
            case "EscDouble":  visible := IsEscDoubleEnabled
            case "LClickSeq1": visible := IsLClick1Enabled
            case "LClickSeq2": visible := IsLClick2Enabled
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
    global EscDelayMs, LClick1_HoldMs, LClick1_GapMs, LClick2_HoldMs, LClick2_GapMs
    global EditEscDelay, EditLClick1Hold, EditLClick1Gap, EditLClick2Hold, EditLClick2Gap
    global TxtStatus

    mins := [200, 200, 17, 200, 17]
    labels := ["ESC 延遲", "左鍵連點1 按壓時間", "左鍵連點1 休息間隔", "左鍵連點2 按壓時間", "左鍵連點2 休息間隔"]
    inputs := [EditEscDelay.Value, EditLClick1Hold.Value, EditLClick1Gap.Value, EditLClick2Hold.Value, EditLClick2Gap.Value]
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
    LClick1_HoldMs := parsed[2]
    LClick1_GapMs := parsed[3]
    LClick2_HoldMs := parsed[4]
    LClick2_GapMs := parsed[5]

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
    btn := (wParam & 0x0001) ? "XButton1" : "XButton2"
    FinishBinding(btn)
}

FinishBinding(newKey) {
    global IsBinding, BindingActionId, BindingDisplayCtrl, BindingInputHook
    global TxtStatus
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyLClick1, KeyLClick2

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
        case "LClickSeq1": KeyLClick1 := newKey
        case "LClickSeq2": KeyLClick2 := newKey
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
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyLClick1, EditKeyLClick2
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkLClick1, ChkLClick2
    global EditEscDelay, EditLClick1Hold, EditLClick1Gap, EditLClick2Hold, EditLClick2Gap
    global LblEscDelay, LblLClick1Hold, LblLClick1Gap, LblLClick2Hold, LblLClick2Gap
    global DelayCtrlMap, TxtLClick1Info, TxtLClick2Info, TxtLClick1Warn, TxtLClick2Warn
    global LClick1WarnPosX, LClick1WarnPosY, LClick2WarnPosX, LClick2WarnPosY, WarnOffsetXPx
    global TxtNikkeStatus

    MainGui := Gui("+AlwaysOnTop")
    MainGui.Title := "Nikke小工具 " AppVersion " - Yabi"

    MainGui.Add("Text", "Section", "綁定按鍵 (觸發鍵)：")

    MainGui.Add("Text", "xs yp+30", "D 連點：")
    EditKeySpamD := MainGui.Add("Edit", "x+24 w90 ReadOnly yp-5", KeySpamD)
    btnBindD  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindD.OnEvent("Click", (*) => StartCaptureBinding("DSpam", EditKeySpamD))
    ChkSpamD := MainGui.Add("CheckBox", (IsSpamDEnabled ? "Checked " : "") "x+5 yp+6", "啟用")
    ChkSpamD.OnEvent("Click", (*) => SetFeatureEnabled("DSpam", ChkSpamD.Value))

    MainGui.Add("Text", "xs yp+30", "S 連點：")
    EditKeySpamS := MainGui.Add("Edit", "x+26 w90 ReadOnly yp-5", KeySpamS)
    btnBindS  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindS.OnEvent("Click", (*) => StartCaptureBinding("SSpam", EditKeySpamS))
    ChkSpamS := MainGui.Add("CheckBox", (IsSpamSEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkSpamS.OnEvent("Click", (*) => SetFeatureEnabled("SSpam", ChkSpamS.Value))

    MainGui.Add("Text", "xs yp+30", "A 連點：")
    EditKeySpamA := MainGui.Add("Edit", "x+24 w90 ReadOnly yp-5", KeySpamA)
    btnBindA  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindA.OnEvent("Click", (*) => StartCaptureBinding("ASpam", EditKeySpamA))
    ChkSpamA := MainGui.Add("CheckBox", (IsSpamAEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkSpamA.OnEvent("Click", (*) => SetFeatureEnabled("ASpam", ChkSpamA.Value))

    MainGui.Add("Text", "xs yp+30", "ESC x2：")
    EditKeyEscDouble := MainGui.Add("Edit", "x+23 w90 ReadOnly yp-5", KeyEscDouble)
    btnBindEsc  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindEsc.OnEvent("Click", (*) => StartCaptureBinding("EscDouble", EditKeyEscDouble))
    ChkEscDouble := MainGui.Add("CheckBox", (IsEscDoubleEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkEscDouble.OnEvent("Click", (*) => SetFeatureEnabled("EscDouble", ChkEscDouble.Value))

    MainGui.Add("Text", "xs yp+30", "左鍵連點1：")
    EditKeyLClick1 := MainGui.Add("Edit", "x+5 w90 ReadOnly yp-5", KeyLClick1)
    btnBindL1  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindL1.OnEvent("Click", (*) => StartCaptureBinding("LClickSeq1", EditKeyLClick1))
    ChkLClick1 := MainGui.Add("CheckBox", (IsLClick1Enabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkLClick1.OnEvent("Click", (*) => SetFeatureEnabled("LClickSeq1", ChkLClick1.Value))

    MainGui.Add("Text", "xs yp+30", "左鍵連點2：")
    EditKeyLClick2 := MainGui.Add("Edit", "x+5 w90 ReadOnly yp-5", KeyLClick2)
    btnBindL2  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindL2.OnEvent("Click", (*) => StartCaptureBinding("LClickSeq2", EditKeyLClick2))
    ChkLClick2 := MainGui.Add("CheckBox", (IsLClick2Enabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkLClick2.OnEvent("Click", (*) => SetFeatureEnabled("LClickSeq2", ChkLClick2.Value))

    MainGui.Add("Text", "xs yp+40 w380 h2 0x10", "")
    MainGui.Add("Text", "xs yp+10", "延遲設定：")

    LblEscDelay := MainGui.Add("Text", "xs yp+30", "ESC：兩次 ESC 中間延遲 (ms)")
    EditEscDelay  := MainGui.Add("Edit", "w120", EscDelayMs)

    LblLClick1Hold := MainGui.Add("Text", , "左鍵連點1：左鍵按壓時間 (ms)")
    EditLClick1Hold  := MainGui.Add("Edit", "w120", LClick1_HoldMs)
    LblLClick1Gap := MainGui.Add("Text", , "左鍵連點1：休息間隔 (ms)")
    EditLClick1Gap  := MainGui.Add("Edit", "w120", LClick1_GapMs)

    LblLClick2Hold := MainGui.Add("Text", , "左鍵連點2：左鍵按壓時間 (ms)")
    EditLClick2Hold  := MainGui.Add("Edit", "w120", LClick2_HoldMs)
    LblLClick2Gap := MainGui.Add("Text", , "左鍵連點2：休息間隔 (ms)")
    EditLClick2Gap  := MainGui.Add("Edit", "w120", LClick2_GapMs)

    DelayCtrlMap["EscDouble"]  := [LblEscDelay, EditEscDelay]
    DelayCtrlMap["LClickSeq1"] := [LblLClick1Hold, EditLClick1Hold, LblLClick1Gap, EditLClick1Gap]
    DelayCtrlMap["LClickSeq2"] := [LblLClick2Hold, EditLClick2Hold, LblLClick2Gap, EditLClick2Gap]
    SetDelayControlsEnabled()

    btnApply := MainGui.Add("Button", "w120", "套用延遲")
    btnApply.OnEvent("Click", ApplyDelayConfig)

    TxtStatus := MainGui.Add("Text", "w380 cGreen", "")

    TxtLClick1Info := MainGui.Add("Text", "xs yp+20 w235", "")
    TxtLClick1Warn := MainGui.Add("Text", "xp+237 yp w180 cRed", "")
    TxtLClick2Info := MainGui.Add("Text", "xs yp+20 w235", "")
    TxtLClick2Warn := MainGui.Add("Text", "xp+237 yp w180 cRed", "")
    WarnOffsetXPx := 0
    TxtLClick1Warn.GetPos(&LClick1WarnPosX, &LClick1WarnPosY)
    TxtLClick2Warn.GetPos(&LClick2WarnPosX, &LClick2WarnPosY)
    UpdateCpsInfo()
    UpdateCpsVisibility()

    MainGui.Add("Text", "xs yp+25 w380 h2 0x10", "")

    ChkAutoStart := MainGui.Add("CheckBox", "xs yp+25", "開機時自動啟動")
    ChkAutoStart.Value := AutoStartEnabled ? 1 : 0
    ChkAutoStart.OnEvent("Click", (*) => ToggleAutoStart(ChkAutoStart.Value))

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
    global KeySpamD, KeySpamS, KeySpamA, KeyEscDouble, KeyLClick1, KeyLClick2
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsEscDoubleEnabled, IsLClick1Enabled, IsLClick2Enabled
    global EscDelayMs, LClick1_HoldMs, LClick1_GapMs, LClick2_HoldMs, LClick2_GapMs
    global AutoStartEnabled
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyEscDouble, EditKeyLClick1, EditKeyLClick2
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkEscDouble, ChkLClick1, ChkLClick2
    global EditEscDelay, EditLClick1Hold, EditLClick1Gap, EditLClick2Hold, EditLClick2Gap
    global ChkAutoStart, TxtStatus

    EditKeySpamD.Value   := KeySpamD
    EditKeySpamS.Value   := KeySpamS
    EditKeySpamA.Value   := KeySpamA
    EditKeyEscDouble.Value := KeyEscDouble
    EditKeyLClick1.Value  := KeyLClick1
    EditKeyLClick2.Value  := KeyLClick2

    ChkSpamD.Value   := IsSpamDEnabled      ? 1 : 0
    ChkSpamS.Value   := IsSpamSEnabled      ? 1 : 0
    ChkSpamA.Value   := IsSpamAEnabled      ? 1 : 0
    ChkEscDouble.Value := IsEscDoubleEnabled  ? 1 : 0
    ChkLClick1.Value  := IsLClick1Enabled ? 1 : 0
    ChkLClick2.Value  := IsLClick2Enabled ? 1 : 0

    EditEscDelay.Value  := EscDelayMs
    EditLClick1Hold.Value := LClick1_HoldMs
    EditLClick1Gap.Value := LClick1_GapMs
    EditLClick2Hold.Value := LClick2_HoldMs
    EditLClick2Gap.Value := LClick2_GapMs

    ChkAutoStart.Value := AutoStartEnabled ? 1 : 0

    UpdateCpsInfo()
    UpdateCpsVisibility()
    UpdateNikkeStatus()
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

; ============================================================
; 執行
; ============================================================
Init()


#Requires AutoHotkey v2.0
#SingleInstance Off
#UseHook  ; 強制鍵鼠熱鍵使用 hook，避免滑鼠鍵在送出點擊時讀不到實體狀態
; ============================================================
; 以系統管理員身分重新啟動
; ============================================================
if !A_IsAdmin {
    try Run('*RunAs "' A_ScriptFullPath '"')
    ExitApp()
}

ClosePreviousInstance()

; ===================================================
; 全域設定與預設值
; ============================================================
SettingsDir      := A_MyDocuments "\NikkeToolSettings"
global SettingsFile := SettingsDir "\NikkeToolSettings.ini"
global AutoStartLink := A_Startup "\NikkeToolStarter.lnk"

; 延遲預設
global Click1_HoldMs := 225
global Click1_GapMs := 25
global Click2_HoldMs := 240
global Click2_GapMs := 40
global Click3_HoldMs := 240
global Click3_GapMs := 40
global KeySpamDelayMs  := 34

; 預設綁定鍵
global KeySpamD      := "F13"
global KeySpamS      := "F14"
global KeySpamA      := "F15"
global KeyClick1 := "F17"
global KeyClick2 := "F18"
global KeyClick3 := "F19"
global KeyPanic := "F20"

; 功能啟用
global IsSpamDEnabled      := false
global IsSpamSEnabled      := false
global IsSpamAEnabled      := false
global IsClick1Enabled := true
global IsClick2Enabled := true
global IsClick3Enabled := false
global ClickBtn1 := "LButton"
global ClickBtn2 := "LButton"
global ClickBtn3 := "LButton"

; 自動啟動
global AutoStartEnabled := false
global IsPanicEnabled := true

; QPC (僅用於量測)
global QPCFreq := 0


; 滑鼠鎖定
global IsCursorLocked := false
global EnableCursorLock := false
global EnableGlobalHotkeys := false

; 熱鍵資料
global HotkeyCurrentMap := Map()
global HotkeyHandlerMap := Map()
global HotkeyBaseKeyMap := Map()
global LastForegroundState := false

; GUI 控制項
global MainGui
global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyClick1, EditKeyClick2, EditKeyClick3, EditKeyPanic
global ChkSpamD, ChkSpamS, ChkSpamA, ChkClick1, ChkClick2, ChkClick3, ChkPanic
global EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
global TxtStatus, ChkAutoStart, TxtNikkeStatus
global ChkGlobalHotkeys
global DelayCtrlMap := Map()
global FeatureCtrlMap := Map()
global TxtClick1Info, TxtClick2Info, TxtClick3Info, TxtClick1Warn, TxtClick2Warn, TxtClick3Warn
global Click1WarnPosX, Click1WarnPosY, Click2WarnPosX, Click2WarnPosY, Click3WarnPosX, Click3WarnPosY, WarnOffsetXPx
global BtnClick1Side, BtnClick2Side, BtnClick3Side
global KeyStateMap := Map()
global HotkeyReleaseMap := Map()
global HotkeyStateMap := Map()
global ActiveHotkeyOwner := ""
global PendingHotkeyId := ""
global PendingHotkeyKey := ""

; 綁定狀態
global IsBinding := false
global BindingActionId := ""
global BindingDisplayCtrl := ""
global BindingInputHook

global AppVersion := "v3.01"

; ============================================================
; 初始化
; ============================================================
Init() {
    global SettingsDir
    if !DirExist(SettingsDir) {
        DirCreate(SettingsDir)
    }
    ok := DllCall("winmm\timeBeginPeriod", "UInt", 1)
    LogEvent("SYS", "timeBeginPeriod", "init", "ok=" (ok = 0) " err=" A_LastError)
    ProcessSetPriority("AboveNormal")

    LoadSettings()
    ApplyAutoStart()
    SetupKeyStateHook()
    UpdateAllHotkeys()
    LastForegroundState := (EnableGlobalHotkeys || IsNikkeForeground())
    ApplyContextState(LastForegroundState)
    BuildGui()
    SetTimer(CursorLockTick, 200)
    OnExit(UnlockCursor)
    A_IconTip := "Nikke小工具 " AppVersion " - Yabi"
}

ClosePreviousInstance() {
    DetectHiddenWindows true
    for hwnd in WinGetList("ahk_class AutoHotkey") {
        if (hwnd = A_ScriptHwnd)
            continue
        title := WinGetTitle("ahk_id " hwnd)
        if (title = "")
            continue
        if InStr(title, A_ScriptFullPath) {
            pid := WinGetPID("ahk_id " hwnd)
            try ProcessClose(pid)
        }
    }
}

; ============================================================
; 遊戲前景判斷
; ============================================================
IsNikkeForeground() {
    return WinActive("ahk_exe nikke.exe") ? true : false
}

IsScriptEnabledForContext() {
    return EnableGlobalHotkeys || IsNikkeForeground()
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
        hwnd := WinActive("ahk_exe nikke.exe")
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
    LogEvent("SYS", "CursorLock", "toggle", "state=" (EnableCursorLock ? "ON" : "OFF"))
    if EnableCursorLock && IsNikkeForeground() {
        LockCursorToNikke()
    } else {
        UnlockCursor()
    }
}

ToggleGlobalHotkeys(state) {
    global EnableGlobalHotkeys
    EnableGlobalHotkeys := (state != 0)
    SaveKeySettings()
    LogEvent("SYS", "GlobalHotkeys", "toggle", "state=" (EnableGlobalHotkeys ? "ON" : "OFF"))
    ApplyContextState(EnableGlobalHotkeys || IsNikkeForeground())
}

ApplyContextState(wantForeground) {
    global HotkeyBaseKeyMap, LastForegroundState
    passThroughNeeded := !(EnableGlobalHotkeys || IsNikkeForeground())
    for id, _ in HotkeyBaseKeyMap {
        if !passThroughNeeded
            BindReleaseHotkey(id)
        ApplyHotkeyState(id, passThroughNeeded)
    }
    if (wantForeground != LastForegroundState) {
        LogEvent("CTX", "-", "hotkeys", (wantForeground ? "ENABLED (fg/global)" : "DISABLED (bg)"))
    }
    LastForegroundState := wantForeground
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
    ApplyContextState(EnableGlobalHotkeys || IsNikkeForeground())
}

; ============================================================
; 計時等待
; ============================================================
WaitMsg(ms) {
    ; MsgWaitForMultipleObjectsEx 只等待 timeout，不等待任何 handle
    DllCall("MsgWaitForMultipleObjectsEx", "UInt", 0, "Ptr", 0, "UInt", ms, "UInt", 0, "UInt", 0x00000004)
}

WaitMsCancel(ms, cancelKey) {
    global QPCFreq
    InitQPC()
    if (QPCFreq = 0) {
        Sleep(ms)
        return true
    }
    start := QpcNow()
    target := start + (QPCFreq * ms // 1000)
    while true {
        if (cancelKey != "" && !IsHookDown(cancelKey)) {
            return false
        }
        now := QpcNow()
        remainingMs := (target - now) * 1000 / QPCFreq
        if (remainingMs <= 0)
            break
        if (remainingMs >= 16) {
            WaitMsg(14)
        } else if (remainingMs >= 2) {
            WaitMsg(1)
        } else
            WaitMsg(0)
    }
    return true
}

InitQPC() {
    global QPCFreq
    if (QPCFreq = 0) {
        ok := DllCall("QueryPerformanceFrequency", "Int64*", &QPCFreq)
        if (!ok || QPCFreq = 0) {
            LogEvent("SYS", "QPC", "initFail", "freq=" QPCFreq)
        }
    }
}

QpcNow() {
    global QPCFreq
    InitQPC()
    now := 0
    DllCall("QueryPerformanceCounter", "Int64*", &now)
    return now
}

QpcDiffMs(start, now) {
    global QPCFreq
    if (QPCFreq = 0)
        return 0
    return (now - start) * 1000 / QPCFreq
}

SetHookStateFlag(key, isDown) {
    global KeyStateMap
    norm := NormalizeHotkeyName(key)
    KeyStateMap[norm] := isDown ? true : false
}

IsHookDown(key) {
    global KeyStateMap
    norm := NormalizeHotkeyName(key)
    flagHook := KeyStateMap.Has(norm) ? KeyStateMap[norm] : false
    return flagHook
}

ShouldKeepRunning(id, key) {
    ; isKeyDown 為真且 token 未被更新時才繼續，任何 up 或新啟動都會跳出
    reasons := ""
    if (HotkeyStateMap.Has(id) && !HotkeyStateMap[id])
        reasons .= "hotkeyUp,"
    if !GetKeyState(key, "P")
        reasons .= "physicalUp,"
    if !IsHookDown(key)
        reasons .= "keyUp,"
    if (reasons != "") {
        LogEvent("RUN", id, "stop", "reason=" RTrim(reasons, ",") " key=" key)
        return false
    }
    return true
}


AcquireHotkeyOwner(id, key) {
    global ActiveHotkeyOwner, PendingHotkeyId, PendingHotkeyKey
    if (ActiveHotkeyOwner = "" || ActiveHotkeyOwner = id) {
        ActiveHotkeyOwner := id
        LogEvent("OWNER", id, "acquire")
        return true
    }
    static lastOwnerQueueLog := 0
    PendingHotkeyId := id
    PendingHotkeyKey := key
    if (A_TickCount - lastOwnerQueueLog >= 1000) {
        LogEvent("OWNER", id, "queue", "owner=" ActiveHotkeyOwner)
        lastOwnerQueueLog := A_TickCount
    }
    return false
}

ReleaseHotkeyOwner(id) {
    global ActiveHotkeyOwner
    if (ActiveHotkeyOwner = id)
        ActiveHotkeyOwner := ""
    LogEvent("OWNER", id, "release")
    TryStartPendingHotkey()
}

TryStartPendingHotkey() {
    global PendingHotkeyId, PendingHotkeyKey, HotkeyHandlerMap
    if (PendingHotkeyId = "")
        return
    id := PendingHotkeyId
    key := PendingHotkeyKey
    PendingHotkeyId := ""
    PendingHotkeyKey := ""
    if (key != "" && IsHookDown(key)) {
        LogEvent("OWNER", id, "startPending")
        if (HotkeyHandlerMap.Has(id)) {
            func := HotkeyHandlerMap[id]
            SetTimer((*) => func.Call(), -1)
        }
    } else {
        LogEvent("OWNER", id, "dropPending")
    }
}

EnsureOwnerAlive() {
    global ActiveHotkeyOwner
    if (ActiveHotkeyOwner = "")
        return
    ownerKey := GetTriggerKey(ActiveHotkeyOwner)
    if (ownerKey = "" || !IsHookDown(ownerKey)) {
        LogEvent("OWNER", ActiveHotkeyOwner, "cleared")
        ActiveHotkeyOwner := ""
    }
}

SetupKeyStateHook() {
    static KeyStateHook
    if IsSet(KeyStateHook)
        return
    ; 使用 InputHook 追蹤鍵盤按鍵的 down/up
    KeyStateHook := InputHook("V")
    KeyStateHook.OnKeyDown := (ih, vk, sc) => OnKeyStateChange(vk, sc, true)
    KeyStateHook.OnKeyUp   := (ih, vk, sc) => OnKeyStateChange(vk, sc, false)
    KeyStateHook.Start()
}


OnKeyStateChange(vk, sc, isDown) {
    global ActiveHotkeyOwner, HotkeyStateMap
    name := GetKeyName(Format("vk{:02X}sc{:03X}", vk, sc))
    if (name = "")
        name := GetKeyName(Format("sc{:03X}", sc))
    if (name = "")
        return
    LogEvent("HOOK", name, (isDown ? "down" : "up"), "vk=" vk " sc=" sc)
    if (!isDown) {
        owner := ActiveHotkeyOwner
        if (owner != "" && HotkeyStateMap.Has(owner) && HotkeyStateMap[owner]) {
            ownerKey := GetTriggerKey(owner)
            if (ownerKey != "" && NormalizeHotkeyName(ownerKey) = NormalizeHotkeyName(name)) {
                HotkeyStateMap[owner] := false
                LogEvent("HOOK", owner, "fallbackUp", "key=" name)
            }
        }
    }
    SetHookStateFlag(name, isDown)
}

GetTriggerKey(id) {
    global HotkeyBaseKeyMap, HotkeyCurrentMap
    if HotkeyBaseKeyMap.Has(id) && (HotkeyBaseKeyMap[id] != "")
        return HotkeyBaseKeyMap[id]
    if HotkeyCurrentMap.Has(id)
        return HotkeyCurrentMap[id]
    return ""
}

ForcePreemptOthers(currentId) {
    for _, id in ["DSpam", "SSpam", "ASpam", "ClickSeq1", "ClickSeq2", "ClickSeq3"] {
        if (id = currentId)
            continue
        key := GetTriggerKey(id)
        if (key != "") {
            Send "{" key " up}"
            SetHookStateFlag(key, false)
        }
    }
}

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

LogEvent(category, id, action, detail := "") {
    msg := "LOG | " category " | " id " | " action
    if (detail != "")
        msg .= " | " detail
    LogDebug(msg)
}

LogDebug(msg) {
    static LogFile := SettingsDir "\NikkeToolDebug.log"
    ts := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    try FileAppend(ts " - " msg "`n", LogFile, "UTF-8")
}

; ============================================================
; 設定載入 / 儲存
; ============================================================
LoadSettings() {
    global SettingsFile, AutoStartLink
    global Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global KeySpamD, KeySpamS, KeySpamA, KeyClick1, KeyClick2, KeyClick3, KeyPanic
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled, IsPanicEnabled
    global ClickBtn1, ClickBtn2, ClickBtn3
    global AutoStartEnabled, EnableCursorLock, EnableGlobalHotkeys

    if !FileExist(SettingsFile)
        return

    try Click1_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click1_HoldMs", Click1_HoldMs))
    try Click1_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click1_GapMs", Click1_GapMs))
    try Click2_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click2_HoldMs", Click2_HoldMs))
    try Click2_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click2_GapMs", Click2_GapMs))
    try Click3_HoldMs := Integer(IniRead(SettingsFile, "Delays", "Click3_HoldMs", Click3_HoldMs))
    try Click3_GapMs := Integer(IniRead(SettingsFile, "Delays", "Click3_GapMs", Click3_GapMs))

    try KeySpamD      := IniRead(SettingsFile, "Keys", "DSpam",      KeySpamD)
    try KeySpamS      := IniRead(SettingsFile, "Keys", "SSpam",      KeySpamS)
    try KeySpamA      := IniRead(SettingsFile, "Keys", "ASpam",      KeySpamA)
    try KeyClick1 := IniRead(SettingsFile, "Keys", "ClickSeq1", KeyClick1)
    try KeyClick2 := IniRead(SettingsFile, "Keys", "ClickSeq2", KeyClick2)
    try KeyClick3 := IniRead(SettingsFile, "Keys", "ClickSeq3", KeyClick3)
    try KeyPanic := IniRead(SettingsFile, "Keys", "Panic", KeyPanic)

    try ClickBtn1 := IniRead(SettingsFile, "Buttons", "ClickSeq1_Button", ClickBtn1)
    try ClickBtn2 := IniRead(SettingsFile, "Buttons", "ClickSeq2_Button", ClickBtn2)
    try ClickBtn3 := IniRead(SettingsFile, "Buttons", "ClickSeq3_Button", ClickBtn3)

    try IsSpamDEnabled      := (Integer(IniRead(SettingsFile, "Enable", "DSpam",      IsSpamDEnabled      ? 1 : 0)) != 0)
    try IsSpamSEnabled      := (Integer(IniRead(SettingsFile, "Enable", "SSpam",      IsSpamSEnabled      ? 1 : 0)) != 0)
    try IsSpamAEnabled      := (Integer(IniRead(SettingsFile, "Enable", "ASpam",      IsSpamAEnabled      ? 1 : 0)) != 0)
    try IsClick1Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq1", IsClick1Enabled ? 1 : 0)) != 0)
    try IsClick2Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq2", IsClick2Enabled ? 1 : 0)) != 0)
    try IsClick3Enabled := (Integer(IniRead(SettingsFile, "Enable", "ClickSeq3", IsClick3Enabled ? 1 : 0)) != 0)
    try IsPanicEnabled := (Integer(IniRead(SettingsFile, "Enable", "Panic", IsPanicEnabled ? 1 : 0)) != 0)

    try AutoStartEnabled := (Integer(IniRead(SettingsFile, "General", "AutoStart", FileExist(AutoStartLink) ? 1 : 0)) != 0)
    try EnableCursorLock := (Integer(IniRead(SettingsFile, "General", "CursorLock", EnableCursorLock ? 1 : 0)) != 0)
    try EnableGlobalHotkeys := (Integer(IniRead(SettingsFile, "General", "GlobalHotkeys", EnableGlobalHotkeys ? 1 : 0)) != 0)
}

SaveKeySettings() {
    global SettingsFile
    global KeySpamD, KeySpamS, KeySpamA, KeyClick1, KeyClick2, KeyClick3, KeyPanic
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled, IsPanicEnabled
    global AutoStartEnabled, EnableCursorLock, EnableGlobalHotkeys
    global Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global ClickBtn1, ClickBtn2, ClickBtn3

    IniWrite(Click1_HoldMs, SettingsFile, "Delays", "Click1_HoldMs")
    IniWrite(Click1_GapMs, SettingsFile, "Delays", "Click1_GapMs")
    IniWrite(Click2_HoldMs, SettingsFile, "Delays", "Click2_HoldMs")
    IniWrite(Click2_GapMs, SettingsFile, "Delays", "Click2_GapMs")
    IniWrite(Click3_HoldMs, SettingsFile, "Delays", "Click3_HoldMs")
    IniWrite(Click3_GapMs, SettingsFile, "Delays", "Click3_GapMs")

    IniWrite(KeySpamD,      SettingsFile, "Keys", "DSpam")
    IniWrite(KeySpamS,      SettingsFile, "Keys", "SSpam")
    IniWrite(KeySpamA,      SettingsFile, "Keys", "ASpam")
    IniWrite(KeyClick1, SettingsFile, "Keys", "ClickSeq1")
    IniWrite(KeyClick2, SettingsFile, "Keys", "ClickSeq2")
    IniWrite(KeyClick3, SettingsFile, "Keys", "ClickSeq3")
    IniWrite(KeyPanic, SettingsFile, "Keys", "Panic")

    IniWrite(ClickBtn1, SettingsFile, "Buttons", "ClickSeq1_Button")
    IniWrite(ClickBtn2, SettingsFile, "Buttons", "ClickSeq2_Button")
    IniWrite(ClickBtn3, SettingsFile, "Buttons", "ClickSeq3_Button")

    IniWrite(IsSpamDEnabled      ? 1 : 0, SettingsFile, "Enable", "DSpam")
    IniWrite(IsSpamSEnabled      ? 1 : 0, SettingsFile, "Enable", "SSpam")
    IniWrite(IsSpamAEnabled      ? 1 : 0, SettingsFile, "Enable", "ASpam")
    IniWrite(IsClick1Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq1")
    IniWrite(IsClick2Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq2")
    IniWrite(IsClick3Enabled ? 1 : 0, SettingsFile, "Enable", "ClickSeq3")
    IniWrite(IsPanicEnabled ? 1 : 0, SettingsFile, "Enable", "Panic")

    IniWrite(AutoStartEnabled ? 1 : 0, SettingsFile, "General", "AutoStart")
    IniWrite(EnableCursorLock ? 1 : 0, SettingsFile, "General", "CursorLock")
    IniWrite(EnableGlobalHotkeys ? 1 : 0, SettingsFile, "General", "GlobalHotkeys")
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
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled, IsPanicEnabled

    if !FeatureCtrlMap.Has(id)
        return

    enabled := true
    switch id {
        case "DSpam":      enabled := IsSpamDEnabled
        case "SSpam":      enabled := IsSpamSEnabled
        case "ASpam":      enabled := IsSpamAEnabled
        case "ClickSeq1":  enabled := IsClick1Enabled
        case "ClickSeq2":  enabled := IsClick2Enabled
        case "ClickSeq3":  enabled := IsClick3Enabled
        case "Panic":      enabled := IsPanicEnabled
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
    global KeySpamD, KeySpamS, KeySpamA, KeyClick1, KeyClick2, KeyClick3, KeyPanic
    BindHotkey("DSpam",      KeySpamD,      HandleSpamD)
    BindHotkey("SSpam",      KeySpamS,      HandleSpamS)
    BindHotkey("ASpam",      KeySpamA,      HandleSpamA)
    BindHotkey("ClickSeq1", KeyClick1, HandleClick1)
    BindHotkey("ClickSeq2", KeyClick2, HandleClick2)
    BindHotkey("ClickSeq3", KeyClick3, HandleClick3)
    BindHotkey("Panic", KeyPanic, HandlePanic)
}

BindHotkey(id, keyName, func) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsPanicEnabled
    global HotkeyHandlerMap, HotkeyBaseKeyMap, HotkeyReleaseMap

    enabled := true
    switch id {
        case "DSpam":       enabled := IsSpamDEnabled
        case "SSpam":       enabled := IsSpamSEnabled
        case "ASpam":       enabled := IsSpamAEnabled
        case "ClickSeq1":  enabled := IsClick1Enabled
        case "ClickSeq2":  enabled := IsClick2Enabled
        case "ClickSeq3":  enabled := IsClick3Enabled
        case "Panic":      enabled := IsPanicEnabled
    }

    HotkeyHandlerMap[id] := func
    HotkeyBaseKeyMap[id] := (enabled && keyName != "") ? keyName : ""
    LogEvent("HOTKEY", id, "bind", "key=" (HotkeyBaseKeyMap[id] = "" ? "EMPTY" : HotkeyBaseKeyMap[id]) " enabled=" enabled)
    BindReleaseHotkey(id)
    ApplyHotkeyState(id, !IsScriptEnabledForContext())
}

NormalizeHotkeyName(name) {
    while (name != "" && SubStr(name, 1, 1) = "~")
        name := SubStr(name, 2)
    return name
}

ApplyHotkeyState(id, passThrough) {
    global HotkeyBaseKeyMap, HotkeyHandlerMap, HotkeyCurrentMap, HotkeyReleaseMap

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

    if passThrough {
        if (current != "") {
            try Hotkey(current, "Off")
            HotkeyCurrentMap[id] := ""
        }
        release := HotkeyReleaseMap.Has(id) ? HotkeyReleaseMap[id] : ""
        if (release != "") {
            try Hotkey(release, "Off")
            HotkeyReleaseMap[id] := ""
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

    if (current != "") {
        try Hotkey(current, "Off")
        catch as e {
            LogEvent("HOTKEY", id, "offFail", "hotkey=" current " err=" e.Message)
        }
    }

    try {
        Hotkey(newHotkey, handler, "On")
        HotkeyCurrentMap[id] := newHotkey
    } catch as e {
        LogEvent("HOTKEY", id, "onFail", "hotkey=" newHotkey " err=" e.Message)
        HotkeyCurrentMap[id] := ""
    }
}

BindReleaseHotkey(id) {
    global HotkeyBaseKeyMap, HotkeyReleaseMap
    base := HotkeyBaseKeyMap.Has(id) ? HotkeyBaseKeyMap[id] : ""
    current := HotkeyReleaseMap.Has(id) ? HotkeyReleaseMap[id] : ""

    if (base = "") {
        if (current != "") {
            try Hotkey(current, "Off")
            HotkeyReleaseMap[id] := ""
        }
        return
    }

    normalized := NormalizeHotkeyName(base)
    releaseHotkey := "*$" normalized " up"

    if (current = releaseHotkey)
        return

    if (current != "") {
        try Hotkey(current, "Off")
        HotkeyReleaseMap[id] := ""
    }

    try {
        Hotkey(releaseHotkey, (*) => OnHotkeyReleased(id, normalized), "On")
        HotkeyReleaseMap[id] := releaseHotkey
    } catch as e {
        LogEvent("HOTKEY", id, "releaseBindFail", "hotkey=" releaseHotkey " err=" e.Message)
        HotkeyReleaseMap[id] := ""
    }
}

OnHotkeyReleased(id, baseKey) {
    LogEvent("HOTKEY", id, "up", "key=" baseKey)
    if (id = "DSpam" || id = "SSpam" || id = "ASpam") {
        HotkeyStateMap[id] := false
        SetHookStateFlag(baseKey, false)
        return
    }
    if (id = "ClickSeq1" || id = "ClickSeq2" || id = "ClickSeq3") {
        HotkeyStateMap[id] := false
    }
    SetHookStateFlag(baseKey, false)
}

HandlePanic(*) {
    global ActiveHotkeyOwner, PendingHotkeyId, PendingHotkeyKey, HotkeyStateMap
    LogEvent("PANIC", "-", "trigger")
    ActiveHotkeyOwner := ""
    PendingHotkeyId := ""
    PendingHotkeyKey := ""

    for id in ["DSpam", "SSpam", "ASpam", "ClickSeq1", "ClickSeq2", "ClickSeq3"] {
        HotkeyStateMap[id] := false
        key := GetTriggerKey(id)
        if (key != "") {
            Send "{" key " up}"
            SetHookStateFlag(key, false)
        }
    }
    for btn in ["LButton", "RButton", "MButton", "XButton1", "XButton2"] {
        Send "{" btn " up}"
    }
}

; ============================================================
; 熱鍵行為
; ============================================================
HandleSpamD(*) {
    global KeySpamD, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    HotkeyStateMap["DSpam"] := true
    SetHookStateFlag(KeySpamD, true)
    if !AcquireHotkeyOwner("DSpam", KeySpamD)
        return
    if !IsHookDown(KeySpamD) {
        ReleaseHotkeyOwner("DSpam")
        return
    }
    try {
        while ShouldKeepRunning("DSpam", KeySpamD) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "DSpam", "stop", "reason=context")
                break
            }
            Send "d"
            if !WaitMsCancel(KeySpamDelayMs, KeySpamD) {
                break
            }
        }
    } finally {
        ; 觸發鍵不再補 up，改由 Hook 狀態判斷
        SetHookStateFlag(KeySpamD, false)
        ReleaseHotkeyOwner("DSpam")
    }
}

HandleSpamS(*) {
    global KeySpamS, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    HotkeyStateMap["SSpam"] := true
    SetHookStateFlag(KeySpamS, true)
    if !AcquireHotkeyOwner("SSpam", KeySpamS)
        return
    if !IsHookDown(KeySpamS) {
        ReleaseHotkeyOwner("SSpam")
        return
    }
    try {
        while ShouldKeepRunning("SSpam", KeySpamS) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "SSpam", "stop", "reason=context")
                break
            }
            Send "s"
            if !WaitMsCancel(KeySpamDelayMs, KeySpamS) {
                break
            }
        }
    } finally {
        SetHookStateFlag(KeySpamS, false)
        ReleaseHotkeyOwner("SSpam")
    }
}

HandleSpamA(*) {
    global KeySpamA, KeySpamDelayMs
    if !IsScriptEnabledForContext()
        return
    HotkeyStateMap["ASpam"] := true
    SetHookStateFlag(KeySpamA, true)
    if !AcquireHotkeyOwner("ASpam", KeySpamA)
        return
    if !IsHookDown(KeySpamA) {
        ReleaseHotkeyOwner("ASpam")
        return
    }
    try {
        while ShouldKeepRunning("ASpam", KeySpamA) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "ASpam", "stop", "reason=context")
                break
            }
            Send "a"
            if !WaitMsCancel(KeySpamDelayMs, KeySpamA) {
                break
            }
        }
    } finally {
        SetHookStateFlag(KeySpamA, false)
        ReleaseHotkeyOwner("ASpam")
    }
}

HandleClick1(*) {
    global Click1_HoldMs, Click1_GapMs, KeyClick1, ClickBtn1
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    HotkeyStateMap["ClickSeq1"] := true
    SetHookStateFlag(KeyClick1, true)
    if !AcquireHotkeyOwner("ClickSeq1", KeyClick1)
        return
    if !IsHookDown(KeyClick1) {
        ReleaseHotkeyOwner("ClickSeq1")
        return
    }
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn1
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        Send "{" eachButton " up}"
        released := true
    }
    if released {
        WaitMsCancel(Click1_GapMs, KeyClick1)
    }
    try {
        while ShouldKeepRunning("ClickSeq1", KeyClick1) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "ClickSeq1", "stop", "reason=context")
                break
            }
            Send btnDown
            holdStart := QpcNow()
            holdOk := WaitMsCancel(Click1_HoldMs, KeyClick1)
            if !holdOk {
                break
            }
            Send btnUp
            gapStart := QpcNow()
            gapOk := WaitMsCancel(Click1_GapMs, KeyClick1)
            if !gapOk
                break
        }
    } finally {
        Send btnUp  ; 確保滑鼠按鍵鬆開
        HotkeyStateMap["ClickSeq1"] := false
        SetHookStateFlag(KeyClick1, false)
        ReleaseHotkeyOwner("ClickSeq1")
    }
}

HandleClick2(*) {
    global Click2_HoldMs, Click2_GapMs, KeyClick2, ClickBtn2
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    HotkeyStateMap["ClickSeq2"] := true
    SetHookStateFlag(KeyClick2, true)
    if !AcquireHotkeyOwner("ClickSeq2", KeyClick2)
        return
    if !IsHookDown(KeyClick2) {
        ReleaseHotkeyOwner("ClickSeq2")
        return
    }
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn2
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        Send "{" eachButton " up}"
        released := true
    }
    if released {
        WaitMsCancel(Click2_GapMs, KeyClick2)
    }
    try {
        while ShouldKeepRunning("ClickSeq2", KeyClick2) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "ClickSeq2", "stop", "reason=context")
                break
            }
            Send btnDown
            holdStart := QpcNow()
            holdOk := WaitMsCancel(Click2_HoldMs, KeyClick2)
            if !holdOk {
                break
            }
            Send btnUp
            gapStart := QpcNow()
            gapOk := WaitMsCancel(Click2_GapMs, KeyClick2)
            if !gapOk
                break
        }
    } finally {
        Send btnUp  ; 確保滑鼠按鍵鬆開
        HotkeyStateMap["ClickSeq2"] := false
        SetHookStateFlag(KeyClick2, false)
        ReleaseHotkeyOwner("ClickSeq2")
    }
}

HandleClick3(*) {
    global Click3_HoldMs, Click3_GapMs, KeyClick3, ClickBtn3
    allowed := IsScriptEnabledForContext()
    if !allowed
        return
    HotkeyStateMap["ClickSeq3"] := true
    SetHookStateFlag(KeyClick3, true)
    if !AcquireHotkeyOwner("ClickSeq3", KeyClick3)
        return
    if !IsHookDown(KeyClick3) {
        ReleaseHotkeyOwner("ClickSeq3")
        return
    }
    ; 若左右鍵任一已按住，先放開並等待一次休息間隔，避免卡住
    button := ClickBtn3
    btnDown := "{" button " down}"
    btnUp   := "{" button " up}"
    released := false
    for eachButton in ["LButton", "RButton"] {
        Send "{" eachButton " up}"
        released := true
    }
    if released {
        WaitMsCancel(Click3_GapMs, KeyClick3)
    }
    try {
        while ShouldKeepRunning("ClickSeq3", KeyClick3) {
            if !IsScriptEnabledForContext() {
                LogEvent("RUN", "ClickSeq3", "stop", "reason=context")
                break
            }
            Send btnDown
            holdStart := QpcNow()
            holdOk := WaitMsCancel(Click3_HoldMs, KeyClick3)
            if !holdOk {
                break
            }
            Send btnUp
            gapStart := QpcNow()
            gapOk := WaitMsCancel(Click3_GapMs, KeyClick3)
            if !gapOk
                break
        }
    } finally {
        Send btnUp  ; 確保滑鼠按鍵鬆開
        HotkeyStateMap["ClickSeq3"] := false
        SetHookStateFlag(KeyClick3, false)
        ReleaseHotkeyOwner("ClickSeq3")
    }
}

; ============================================================
; 勾選事件與延遲顯示
; ============================================================
SetFeatureEnabled(id, state) {
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled, IsPanicEnabled
    global TxtStatus

    enabled := (state != 0)
    switch id {
        case "DSpam":      IsSpamDEnabled      := enabled
        case "SSpam":      IsSpamSEnabled      := enabled
        case "ASpam":      IsSpamAEnabled      := enabled
        case "ClickSeq1": IsClick1Enabled := enabled
        case "ClickSeq2": IsClick2Enabled := enabled
        case "ClickSeq3": IsClick3Enabled := enabled
        case "Panic":     IsPanicEnabled := enabled
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
    global IsClick1Enabled, IsClick2Enabled, IsClick3Enabled

    for id, ctrls in DelayCtrlMap {
        visible := true
        switch id {
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
    global Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
    global TxtStatus

    mins := [200, 17, 200, 17, 200, 17]
    labels := ["連點1 按壓時間", "連點1 休息間隔", "連點2 按壓時間", "連點2 休息間隔", "連點3 按壓時間", "連點3 休息間隔"]
    inputs := [EditClick1Hold.Value, EditClick1Gap.Value, EditClick2Hold.Value, EditClick2Gap.Value, EditClick3Hold.Value, EditClick3Gap.Value]
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

    Click1_HoldMs := parsed[1]
    Click1_GapMs := parsed[2]
    Click2_HoldMs := parsed[3]
    Click2_GapMs := parsed[4]
    Click3_HoldMs := parsed[5]
    Click3_GapMs := parsed[6]

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
    global KeySpamD, KeySpamS, KeySpamA, KeyClick1, KeyClick2, KeyClick3, KeyPanic

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
        case "ClickSeq1": KeyClick1 := newKey
        case "ClickSeq2": KeyClick2 := newKey
        case "ClickSeq3": KeyClick3 := newKey
        case "Panic":     KeyPanic := newKey
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
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyClick1, EditKeyClick2, EditKeyClick3, EditKeyPanic
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkClick1, ChkClick2, ChkClick3, ChkPanic
    global EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
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

    lblPanic := MainGui.Add("Text", "xs yp+20", "逃脫鍵：")
    EditKeyPanic := MainGui.Add("Edit", "x+23 w90 ReadOnly yp-5", KeyPanic)
    btnBindPanic  := MainGui.Add("Button", "x+5 w80 yp-1", "變更")
    btnBindPanic.OnEvent("Click", (*) => StartCaptureBinding("Panic", EditKeyPanic))
    ChkPanic := MainGui.Add("CheckBox", (IsPanicEnabled ? "Checked " : "") "x+5 yp+5", "啟用")
    ChkPanic.OnEvent("Click", (*) => SetFeatureEnabled("Panic", ChkPanic.Value))
    FeatureCtrlMap["Panic"] := [lblPanic, EditKeyPanic, btnBindPanic]

    MainGui.Add("Text", "xs yp+30 w380 h2 0x10", "")
    MainGui.Add("Text", "xs yp+20", "延遲設定：")

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

    ChkCursorLock := MainGui.Add("CheckBox", "xs yp+20", "鎖定滑鼠於遊戲視窗內")
    ChkCursorLock.Value := EnableCursorLock ? 1 : 0
    ChkCursorLock.OnEvent("Click", (*) => ToggleCursorLock(ChkCursorLock.Value))

    ChkGlobalHotkeys := MainGui.Add("CheckBox", "xs yp+20", "全域啟用熱鍵")
    ChkGlobalHotkeys.Value := EnableGlobalHotkeys ? 1 : 0
    ChkGlobalHotkeys.OnEvent("Click", (*) => ToggleGlobalHotkeys(ChkGlobalHotkeys.Value))

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
    global KeySpamD, KeySpamS, KeySpamA, KeyClick1, KeyClick2, KeyClick3, KeyPanic
    global IsSpamDEnabled, IsSpamSEnabled, IsSpamAEnabled, IsClick1Enabled, IsClick2Enabled, IsClick3Enabled
    global Click1_HoldMs, Click1_GapMs, Click2_HoldMs, Click2_GapMs, Click3_HoldMs, Click3_GapMs
    global AutoStartEnabled
    global BtnClick1Side, BtnClick2Side, BtnClick3Side
    global ClickBtn1, ClickBtn2, ClickBtn3
    global EditKeySpamD, EditKeySpamS, EditKeySpamA, EditKeyClick1, EditKeyClick2, EditKeyClick3, EditKeyPanic
    global ChkSpamD, ChkSpamS, ChkSpamA, ChkClick1, ChkClick2, ChkClick3
    global EditClick1Hold, EditClick1Gap, EditClick2Hold, EditClick2Gap, EditClick3Hold, EditClick3Gap
    global ChkAutoStart, TxtStatus, EnableCursorLock, ChkGlobalHotkeys, EnableGlobalHotkeys

    EditKeySpamD.Value   := KeySpamD
    EditKeySpamS.Value   := KeySpamS
    EditKeySpamA.Value   := KeySpamA
    EditKeyClick1.Value  := KeyClick1
    EditKeyClick2.Value  := KeyClick2
    EditKeyClick3.Value  := KeyClick3
    if IsSet(EditKeyPanic)
        EditKeyPanic.Value := KeyPanic

    ChkSpamD.Value   := IsSpamDEnabled      ? 1 : 0
    ChkSpamS.Value   := IsSpamSEnabled      ? 1 : 0
    ChkSpamA.Value   := IsSpamAEnabled      ? 1 : 0
    ChkClick1.Value  := IsClick1Enabled ? 1 : 0
    ChkClick2.Value  := IsClick2Enabled ? 1 : 0
    ChkClick3.Value  := IsClick3Enabled ? 1 : 0
    if IsSet(ChkPanic)
        ChkPanic.Value := IsPanicEnabled ? 1 : 0
    SetClickButtonText(1, ClickBtn1)
    SetClickButtonText(2, ClickBtn2)
    SetClickButtonText(3, ClickBtn3)

    EditClick1Hold.Value := Click1_HoldMs
    EditClick1Gap.Value := Click1_GapMs
    EditClick2Hold.Value := Click2_HoldMs
    EditClick2Gap.Value := Click2_GapMs
    EditClick3Hold.Value := Click3_HoldMs
    EditClick3Gap.Value := Click3_GapMs

    ChkAutoStart.Value := AutoStartEnabled ? 1 : 0
    if IsSet(ChkGlobalHotkeys)
        ChkGlobalHotkeys.Value := EnableGlobalHotkeys ? 1 : 0
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

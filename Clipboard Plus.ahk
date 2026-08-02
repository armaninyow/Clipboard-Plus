;@Ahk2Exe-SetMainIcon icon.ico
#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
;  CLIPBOARD PLUS v1.1  |  AHK v2  |  Hotkey: Win+V
;  Custom drawn items inside a child-Gui scroll viewport
;  Child Gui clips its own children — menu bar never overlapped
;  Settings + history persisted to ClipboardManager.ini
; ============================================================

; ---------- Config ----------
global INI_FILE      := A_ScriptDir "\ClipboardManager.ini"
global CLIP_DIR      := A_ScriptDir "\ClipboardData"
global THUMB_DIR     := A_ScriptDir "\ClipboardData\Thumbs"
global MaxItems      := 25
global MAX_CLIP_BYTES := 8 * 1024 * 1024   ; 8MB cap on ClipboardAll() capture; beyond this, store plain text only
global PlainTextMode := false
global KeepOpen      := false
global ShowPinned    := false

; ---------- Runtime ----------
global ClipHistory   := []
global NextID        := 1
global LastClipText  := ""
global IgnoreNextClip := 0
global ManagerGui    := ""
global ScrollGui     := ""   ; child Gui — the scroll viewport
global ItemControls  := []
global ScrollOffset  := 0
global TooltipTimer  := ""
global ExpandedIds   := []
global GDIP_TOKEN    := 0
global LVDisplay     := []

; Menu bar buttons
global TooltipLastHwnd := 0
global BtnPin      := ""
global BtnClear    := ""
global BtnSettings := ""
global BtnWinClip  := ""
global MouseTracked := false

; Layout
global MENUBAR_H := 26
global SEP_H     := 1
global LIST_TOP  := 0    ; set after chrome built = MENUBAR_H + SEP_H
global ITEM_PAD  := 10
global MIN_ITEM_H := 36
global MAX_ITEM_H := 85
global MAX_EXPANDED_CHARS := 700   ; hard backstop for expanded view, independent of line-count estimate
global HoveredItemId := 0
global HoveredExpandBtnId := 0
global HoveredHeaderBtn := ""

; ============================================================
;  INI  — LOAD / SAVE
; ============================================================
LoadConfig() {
    global INI_FILE, CLIP_DIR, MaxItems, PlainTextMode, KeepOpen, ShowPinned
    global ClipHistory, NextID
    try MaxItems      := Integer(IniRead(INI_FILE, "Settings", "MaxItems",      25))
    try PlainTextMode := Integer(IniRead(INI_FILE, "Settings", "PlainTextMode", 0)) = 1
    try KeepOpen      := Integer(IniRead(INI_FILE, "Settings", "KeepOpen",      0)) = 1
    try ShowPinned    := Integer(IniRead(INI_FILE, "Settings", "ShowPinned",    0)) = 1
    if !DirExist(CLIP_DIR)
        DirCreate(CLIP_DIR)
    if !DirExist(THUMB_DIR)
        DirCreate(THUMB_DIR)
    ClipHistory := []
    count := 0
    try count := Integer(IniRead(INI_FILE, "History", "Count", 0))
    Loop count {
        try {
            pinned := Integer(IniRead(INI_FILE, "History", "Pinned" . A_Index, 0)) = 1
            ; Full text is stored on disk (avoids the ~32KB per-value INI limit
            ; that used to silently corrupt the History section on large copies).
            ; Fall back to the INI value (older files / short items).
            textFile := CLIP_DIR "\" A_Index ".txt"
            if FileExist(textFile) {
                text := FileRead(textFile, "UTF-8")
            } else {
                text := IniRead(INI_FILE, "History", "Text" . A_Index, "")
                text := StrReplace(text, "\n", "`n")
                text := StrReplace(text, "\r", "`r")
            }
            ; Load binary clipboard data if saved
            clipFile := CLIP_DIR "\" A_Index ".clip"
            clip := ""
            if FileExist(clipFile) {
                buf := Buffer(FileGetSize(clipFile))
                f   := FileOpen(clipFile, "r")
                f.RawRead(buf)
                f.Close()
                clip := ClipboardAll(buf, buf.Size)
            }
            isImage  := false
            thumb    := ""
            try isImage := Integer(IniRead(INI_FILE, "History", "IsImage" . A_Index, 0)) = 1
            if isImage {
                thumbFile := THUMB_DIR "\" A_Index ".png"
                thumb := FileExist(thumbFile) ? thumbFile : ""
            }
            ClipHistory.Push({text: text, clip: clip, thumb: thumb, isImage: isImage, pinned: pinned, id: NextID++})
        }
    }
    ; Remove any PNG files in Thumbs not referenced by loaded history
    PurgeOrphanedThumbs()
}

SaveConfig() {
    global INI_FILE, CLIP_DIR, MaxItems, PlainTextMode, KeepOpen, ShowPinned, ClipHistory
    IniWrite(MaxItems,            INI_FILE, "Settings", "MaxItems")
    IniWrite(PlainTextMode?1:0,   INI_FILE, "Settings", "PlainTextMode")
    IniWrite(KeepOpen?1:0,        INI_FILE, "Settings", "KeepOpen")
    IniWrite(ShowPinned?1:0,      INI_FILE, "Settings", "ShowPinned")
    try IniDelete(INI_FILE, "History")
    IniWrite(ClipHistory.Length,  INI_FILE, "History", "Count")
    if !DirExist(CLIP_DIR)
        DirCreate(CLIP_DIR)
    ; Delete old clip files first
    Loop Files, CLIP_DIR "\*.clip"
        FileDelete(A_LoopFileFullPath)
    Loop ClipHistory.Length {
        try {
            item := ClipHistory[A_Index]

            ; Full text always goes to its own file — INI values have a hard
            ; ~32KB limit and silently corrupt the whole [History] section
            ; (which used to wipe pins/other items) when a huge copy overflows it.
            try {
                f := FileOpen(CLIP_DIR "\" A_Index ".txt", "w", "UTF-8")
                f.Write(item.text)
                f.Close()
            }
            ; Keep a small bounded preview in the INI too, just for safety/debugging.
            preview := SubStr(item.text, 1, 500)
            safe := StrReplace(preview, "`n", "\n")
            safe := StrReplace(safe,    "`r", "\r")
            IniWrite(safe,            INI_FILE, "History", "Text"   . A_Index)
            IniWrite(item.pinned?1:0, INI_FILE, "History", "Pinned" . A_Index)
            IniWrite((item.HasProp("isImage") && item.isImage) ? 1 : 0, INI_FILE, "History", "IsImage" . A_Index)
            ; Save binary clipboard data
            if (item.HasProp("clip") && item.clip != "" && !(item.clip is String)) {
                f := FileOpen(CLIP_DIR "\" A_Index ".clip", "w")
                f.RawWrite(item.clip)
                f.Close()
            }
            ; Copy thumb to indexed filename so load can find it
            if (item.HasProp("isImage") && item.isImage && item.HasProp("thumb") && item.thumb != "" && FileExist(item.thumb)) {
                destThumb := THUMB_DIR "\" A_Index ".png"
                if (item.thumb != destThumb)
                    FileCopy(item.thumb, destThumb, 1)
            }
        } catch {
            ; Don't let one bad item abort the save of everything else.
            continue
        }
    }
}

; Last line of defense: log unexpected errors instead of letting the whole
; script die (which was also what left the INI in a half-written state).
OnError(GlobalErrorHandler)
GlobalErrorHandler(e, mode) {
    try {
        FileAppend(FormatTime(, "yyyy-MM-dd HH:mm:ss") " | " e.Message " | " e.What " | line " e.Line "`n",
            A_ScriptDir "\ClipboardPlus_error.log", "UTF-8")
    }
    return true   ; true = suppress the error and keep running
}

LoadConfig()
GdipStart()

; ============================================================
;  TRAY
; ============================================================
A_TrayMenu.Delete()
A_TrayMenu.Add("Open Clipboard Plus", (*) => ShowOrRefresh())
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Open Clipboard Plus"
A_IconTip := "Clipboard Plus v1.1"
if A_IsCompiled
    TraySetIcon(A_ScriptFullPath, 1)
else if FileExist(A_ScriptDir "\icon.ico")
    TraySetIcon(A_ScriptDir "\icon.ico")

; ============================================================
;  WIN+V
; ============================================================
try RegWrite(0, "REG_DWORD", "HKCU\Software\Microsoft\Clipboard", "EnableClipboardHistory")
OnExit(OnAppExit)
OnAppExit(*) {
    SaveConfig()
    GdipShutdown()
    try RegWrite(1, "REG_DWORD", "HKCU\Software\Microsoft\Clipboard", "EnableClipboardHistory")
}

#v:: {
    CoordMode("Mouse", "Screen")
    MouseGetPos(&_mx, &_my)
    ShowOrRefresh(_mx, _my)
}

ShowOrRefresh(mx := -1, my := -1) {
    global ManagerGui, ScrollOffset
    if (mx = -1)
        MouseGetPos(&mx, &my)
    if (ManagerGui = "") {
        BuildGui()
        PlaceAndShow(mx, my)
        SyncPinButton()
        return
    }
    if (WinExist("ahk_id " ManagerGui.Hwnd) && WinActive("ahk_id " ManagerGui.Hwnd)) {
        ManagerGui.Hide()
        return
    }
    ScrollOffset := 0
    RebuildItems()
    PlaceAndShow(mx, my)
    SyncPinButton()
}

SyncPinButton() {
    global BtnPin, ShowPinned, HoveredHeaderBtn
    if (BtnPin = "")
        return
    BtnPin.Opt("Background" HeaderBtnColor("pin", HoveredHeaderBtn = "pin"))
    DllCall("InvalidateRect", "Ptr", BtnPin.Hwnd, "Ptr", 0, "Int", 1)
    DllCall("UpdateWindow",   "Ptr", BtnPin.Hwnd)
}


; ============================================================
;  GDI+ HELPERS
; ============================================================
global GDIP_TOKEN := 0

GdipStart() {
    global GDIP_TOKEN
    if (GDIP_TOKEN != 0)
        return
    DllCall("LoadLibrary", "Str", "gdiplus")
    si := Buffer(24, 0)
    NumPut("UInt", 1, si, 0)
    DllCall("gdiplus\GdiplusStartup", "Ptr*", &tok := 0, "Ptr", si, "Ptr", 0)
    GDIP_TOKEN := tok
}

GdipShutdown() {
    global GDIP_TOKEN
    if (GDIP_TOKEN = 0)
        return
    DllCall("gdiplus\GdiplusShutdown", "Ptr", GDIP_TOKEN)
    GDIP_TOKEN := 0
}

; Save clipboard image as PNG thumbnail with transparency composited over bg color
; bgColor is 0xAARRGGBB — use 0xFF282828 for dark item background
SaveClipImageThumb(filePath, thumbW := 200, thumbH := 120, bgColor := 0xFF282828) {
    GdipStart()
    if !DllCall("OpenClipboard", "Ptr", 0)
        return false

    ; Try CF_DIBV5 (17) first — preserves alpha channel
    ; Fall back to CF_DIB (8) if not available
    bmp := 0
    hData := DllCall("GetClipboardData", "UInt", 17, "Ptr")  ; CF_DIBV5
    if hData {
        pData := DllCall("GlobalLock", "Ptr", hData, "Ptr")
        if pData {
            biSize    := NumGet(pData + 0,  "UInt")
            biWidth   := NumGet(pData + 4,  "Int")
            biHeight  := Abs(NumGet(pData + 8, "Int"))
            biBitCount := NumGet(pData + 14, "UShort")
            biCompression := NumGet(pData + 16, "UInt")
            biClrUsed := NumGet(pData + 32, "UInt")
            palColors := (biBitCount <= 8) ? (biClrUsed ? biClrUsed : (1 << biBitCount)) : 0
            ; For BI_BITFIELDS (3), palette area holds 3 DWORD masks
            if (biCompression = 3)
                palColors := Max(palColors, 3)
            palSize   := palColors * 4
            pixOffset := biSize + palSize
            stride    := ((biWidth * biBitCount + 31) // 32) * 4
            ; PixelFormat32bppARGB = 0x26200A, 32bppRGB = 0x22009
            fmt := (biBitCount = 32) ? 0x26200A : 0x22009
            DllCall("gdiplus\GdipCreateBitmapFromScan0",
                "Int", biWidth, "Int", biHeight, "Int", stride,
                "Int", fmt, "Ptr", pData + pixOffset, "Ptr*", &bmp)
            DllCall("GlobalUnlock", "Ptr", hData)
        }
    }
    ; Fallback: CF_DIB
    if (!bmp) {
        hData := DllCall("GetClipboardData", "UInt", 8, "Ptr")
        if hData {
            pData := DllCall("GlobalLock", "Ptr", hData, "Ptr")
            if pData {
                DllCall("gdiplus\GdipCreateBitmapFromGdiDib", "Ptr", pData, "Ptr", pData + NumGet(pData+0,"UInt") + (NumGet(pData+32,"UInt") ? NumGet(pData+32,"UInt") : 0)*4, "Ptr*", &bmp)
                DllCall("GlobalUnlock", "Ptr", hData)
            }
        }
    }
    DllCall("CloseClipboard")
    if !bmp
        return false

    ; Get dimensions
    DllCall("gdiplus\GdipGetImageWidth",  "Ptr", bmp, "UInt*", &origW := 0)
    DllCall("gdiplus\GdipGetImageHeight", "Ptr", bmp, "UInt*", &origH := 0)
    if (origW = 0 || origH = 0) {
        DllCall("gdiplus\GdipDisposeImage", "Ptr", bmp)
        return false
    }

    ; Compute scaled size preserving aspect ratio
    scale := Min(thumbW / origW, thumbH / origH)
    newW  := Max(1, Round(origW * scale))
    newH  := Max(1, Round(origH * scale))

    ; Create destination bitmap (32bpp ARGB)
    DllCall("gdiplus\GdipCreateBitmapFromScan0",
        "Int", newW, "Int", newH, "Int", 0, "Int", 0x26200A, "Ptr", 0, "Ptr*", &thumb := 0)
    DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", thumb, "Ptr*", &g := 0)
    DllCall("gdiplus\GdipSetInterpolationMode", "Ptr", g, "Int", 7)

    ; Fill background with item bg color so transparent areas look correct
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", bgColor, "Ptr*", &brush := 0)
    DllCall("gdiplus\GdipFillRectangleI",  "Ptr", g, "Ptr", brush, "Int", 0, "Int", 0, "Int", newW, "Int", newH)
    DllCall("gdiplus\GdipDeleteBrush",     "Ptr", brush)

    ; Draw image over background — flip vertically to correct bottom-up DIB storage
    DllCall("gdiplus\GdipDrawImageRectI", "Ptr", g, "Ptr", bmp,
        "Int", 0, "Int", newH, "Int", newW, "Int", -newH)
    DllCall("gdiplus\GdipDeleteGraphics", "Ptr", g)
    DllCall("gdiplus\GdipDisposeImage",   "Ptr", bmp)

    ; Save as PNG
    pngClsid := Buffer(16)
    CLSIDFromString("{557CF406-1A04-11D3-9A73-0000F81EF32E}", pngClsid)
    DllCall("gdiplus\GdipSaveImageToFile", "Ptr", thumb, "WStr", filePath, "Ptr", pngClsid, "Ptr", 0)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", thumb)
    return true
}

CLSIDFromString(str, buf) {
    DllCall("ole32\CLSIDFromString", "WStr", str, "Ptr", buf)
}

; ============================================================
;  CLIPBOARD HOOK
; ============================================================
OnClipboardChange(ClipChanged)

ClipChanged(DataType) {
    global ClipHistory, LastClipText, PlainTextMode, MaxItems, ManagerGui, NextID
    global CLIP_DIR, THUMB_DIR, IgnoreNextClip

    if (IgnoreNextClip > 0) {
        IgnoreNextClip -= 1
        return
    }

    ; ── Image capture (DataType=2, CF_DIB available) ──────────────
    if (DataType = 2 && DllCall("IsClipboardFormatAvailable", "UInt", 8)) {
        clip := ClipboardAll()
        if (clip.Size = 0)
            return
        ; Ensure dirs exist
        if !DirExist(CLIP_DIR)
            DirCreate(CLIP_DIR)
        if !DirExist(THUMB_DIR)
            DirCreate(THUMB_DIR)
        id   := NextID++
        ; Generate thumbnail — clipboard still open from ClipboardAll() above
        thumb := ""
        thumbFile := THUMB_DIR "\" id ".png"
        if SaveClipImageThumb(thumbFile)
            thumb := thumbFile
        ; Trim oldest unpinned if at limit
        while (ClipHistory.Length >= MaxItems) {
            removed := false
            Loop ClipHistory.Length {
                ri := ClipHistory.Length - A_Index + 1
                if (!ClipHistory[ri].pinned) {
                    ClipHistory.RemoveAt(ri)
                    removed := true
                    break
                }
            }
            if (!removed)
                break
        }
        ClipHistory.InsertAt(1, {text: "[Image]", clip: clip, thumb: thumb,
            isImage: true, pinned: false, id: id})
        if (ManagerGui != "" && WinExist("ahk_id " ManagerGui.Hwnd)) {
            RebuildItems()
            DllCall("InvalidateRect", "Ptr", ScrollGui.Hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow",   "Ptr", ScrollGui.Hwnd)
        }
        ShowCopyTooltip()
        return
    }

    ; ── Text capture ───────────────────────────────────────────────
    if (DataType != 1)
        return
    if (DllCall("IsClipboardFormatAvailable", "UInt", 15))
        return
    try {
        text := A_Clipboard
    } catch {
        return
    }
    if (text = "" || text = LastClipText)
        return
    ; Windows 11's double-click-anywhere text selection can copy a button's
    ; own caption (e.g. this app's header emoji buttons) straight to the
    ; clipboard, bypassing our code entirely. Recognize and ignore that.
    static OwnUiCaptions := ["📌", "❌", "⚙️", "🪟"]
    for caption in OwnUiCaptions {
        if (text = caption)
            return
    }
    LastClipText := text
    if (PlainTextMode) {
        A_Clipboard := text
        text := A_Clipboard
    }
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].text = text) {
            moved := ClipHistory.RemoveAt(A_Index)
            moved.clip := PlainTextMode ? text : SafeClipboardAll()
            ClipHistory.InsertAt(1, moved)
            if (ManagerGui != "" && WinExist("ahk_id " ManagerGui.Hwnd))
                RebuildItems()
            ShowCopyTooltip()
            return
        }
    }
    while (ClipHistory.Length >= MaxItems) {
        removed := false
        Loop ClipHistory.Length {
            ri := ClipHistory.Length - A_Index + 1
            if (!ClipHistory[ri].pinned) {
                ClipHistory.RemoveAt(ri)
                removed := true
                break
            }
        }
        if (!removed)
            break
    }
    clip := PlainTextMode ? text : SafeClipboardAll()
    ClipHistory.InsertAt(1, {text: text, clip: clip, pinned: false, id: NextID++})
    if (ManagerGui != "" && WinExist("ahk_id " ManagerGui.Hwnd))
        RebuildItems()
    ShowCopyTooltip()
}

; Capture full clipboard formatting, but bail out to plain-text-only storage
; if it's unreasonably large or errors — this is what used to crash the script
; on very large copies.
SafeClipboardAll() {
    global MAX_CLIP_BYTES
    try {
        clip := ClipboardAll()
        if (clip.Size > MAX_CLIP_BYTES)
            return ""   ; too big — caller keeps plain text only, no formatted copy
        return clip
    } catch {
        return ""
    }
}

ShowCopyTooltip() {
    ShowTooltip("Copied!", 1500)
}

ShowTooltip(text, ms := 1000) {
    global TooltipTimer
    ToolTip(text)
    if (TooltipTimer != "")
        SetTimer(TooltipTimer, 0)
    TooltipTimer := () => ToolTip()
    SetTimer(TooltipTimer, -ms)
}

; ============================================================
;  BUILD GUI
; ============================================================
BuildGui() {
    global ManagerGui, ScrollGui, MENUBAR_H, LIST_TOP
    global BtnPin, BtnClear, BtnSettings, BtnWinClip

    ManagerGui := Gui("+Resize +MinSize240x200 -MaximizeBox -MinimizeBox -DPIScale", "Clipboard Plus v1.1")
    ManagerGui.BackColor := "1C1C1C"
    ManagerGui.MarginX   := 0
    ManagerGui.MarginY   := 0

    ; ── Menu bar ──────────────────────────────────────────────────────
    ; 4 emoji buttons right-aligned, no separator
    btnW := 32
    ManagerGui.SetFont("s12", "Segoe UI Emoji")

    BtnPin := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "📌")
    BtnPin.OnEvent("Click",       (*) => OnBtnPin())

    BtnClear := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "❌")
    BtnClear.OnEvent("Click",       (*) => OnBtnClear())

    BtnSettings := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "⚙️")
    BtnSettings.OnEvent("Click",       (*) => OnBtnSettings())

    BtnWinClip := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "🪟")
    BtnWinClip.OnEvent("Click",       (*) => OnBtnWinClip())

    ; No separator — menu bar blends into background
    LIST_TOP := MENUBAR_H

    ; ── Scroll viewport (child Gui) ───────────────────────────
    ; The child Gui is a real HWND child — Windows clips its children
    ; to its own bounds, so items drawn inside can never bleed into the menu bar.
    ScrollGui := Gui("+Parent" ManagerGui.Hwnd " -Caption -Border -DPIScale")
    ScrollGui.BackColor := "1C1C1C"
    ScrollGui.MarginX   := 0
    ScrollGui.MarginY   := 0
    ScrollGui.SetFont("s9 cE0E0E0 w400", "Segoe UI")

    ; Position and size the child Gui within parent
    DllCall("SetWindowPos", "Ptr", ScrollGui.Hwnd,
        "Ptr", 0,
        "Int", 0, "Int", LIST_TOP,
        "Int", 360, "Int", 400 - LIST_TOP,
        "UInt", 0x0040)   ; SWP_SHOWWINDOW

    ManagerGui.OnEvent("Size",   OnGuiSize)
    ManagerGui.OnEvent("Close",  (*) => ManagerGui.Hide())
    ManagerGui.OnEvent("Escape", (*) => ManagerGui.Hide())

    ; Tooltip on hover via WM_MOUSEMOVE
    OnMessage(0x0200, OnMouseMove)
    ; Mousewheel on child or parent
    OnMessage(0x020A, OnMouseWheel)

    RebuildItems()
}

; ============================================================
;  PLACE AND SHOW
; ============================================================
PlaceAndShow(mx := -1, my := -1) {
    global ManagerGui
    if (mx = -1) {
        CoordMode("Mouse", "Screen")
        MouseGetPos(&mx, &my)
    }
    w := 360, h := 400
    ; Flip horizontally: open to the right of cursor unless it would clip
    x := (mx + w > A_ScreenWidth)  ? mx - w : mx
    ; Flip vertically: open below cursor unless it would clip
    y := (my + h > A_ScreenHeight) ? my - h : my
    ManagerGui.Show("x" x " y" y " w" w " h" h)
    MoveMenuButtons(w)
}

; ============================================================
;  RESIZE
; ============================================================
OnGuiSize(guiObj, minMax, w, h) {
    global ScrollGui, LIST_TOP
    if (minMax = -1)
        return
    MoveMenuButtons(w)
    DllCall("SetWindowPos", "Ptr", ScrollGui.Hwnd,
        "Ptr", 0,
        "Int", 0, "Int", LIST_TOP,
        "Int", w, "Int", h - LIST_TOP,
        "UInt", 0x0004)   ; SWP_NOZORDER
    RedrawMenuButtons()
    RebuildItems()
}

MoveMenuButtons(w) {
    global BtnPin, BtnClear, BtnSettings, BtnWinClip, MENUBAR_H
    if (BtnPin = "")
        return
    btnW  := 32
    flags := 0x0004   ; SWP_NOZORDER
    ; Right-to-left: 🪟 ⚙ ❌ 📌
    DllCall("SetWindowPos","Ptr",BtnWinClip.Hwnd, "Ptr",0, "Int",w - btnW*1, "Int",0, "Int",btnW,"Int",MENUBAR_H,"UInt",flags)
    DllCall("SetWindowPos","Ptr",BtnSettings.Hwnd,"Ptr",0, "Int",w - btnW*2, "Int",0, "Int",btnW,"Int",MENUBAR_H,"UInt",flags)
    DllCall("SetWindowPos","Ptr",BtnClear.Hwnd,   "Ptr",0, "Int",w - btnW*3, "Int",0, "Int",btnW,"Int",MENUBAR_H,"UInt",flags)
    DllCall("SetWindowPos","Ptr",BtnPin.Hwnd,     "Ptr",0, "Int",w - btnW*4, "Int",0, "Int",btnW,"Int",MENUBAR_H,"UInt",flags)
}

RedrawMenuButtons() {
    global BtnPin, BtnClear, BtnSettings, BtnWinClip
    if (BtnPin = "")
        return
    for btn in [BtnPin, BtnClear, BtnSettings, BtnWinClip]
        DllCall("InvalidateRect", "Ptr", btn.Hwnd, "Ptr", 0, "Int", 1)
    DllCall("UpdateWindow", "Ptr", BtnPin.Hwnd)
}

; ============================================================
;  ITEM DRAWING inside ScrollGui
; ============================================================
GetScrollW() {
    rc := Buffer(16, 0)
    DllCall("GetClientRect", "Ptr", ScrollGui.Hwnd, "Ptr", rc)
    return NumGet(rc, 8, "Int")
}

GetScrollH() {
    rc := Buffer(16, 0)
    DllCall("GetClientRect", "Ptr", ScrollGui.Hwnd, "Ptr", rc)
    return NumGet(rc, 12, "Int")
}

; Windows Text controls only wrap at whitespace. A single long token (URL,
; long word, no-space blob) will just overflow the box instead of wrapping.
; A bare line-feed isn't reliably treated as a hard break by the control's
; own auto-wrap — it can get reflowed/merged with neighboring lines. Real
; space characters DO work reliably (that's how normal multi-word text
; already wraps), so insert one every N characters within long tokens.
WrapLongTokens(text, pxWidth, fontSize) {
    breakEvery := 50   ; target line length for unbroken words/URLs
    out := ""
    Loop Parse, text, "`n" {
        lineText := A_LoopField
        outLine := ""
        Loop Parse, lineText, " ", " " {
            token := A_LoopField
            if (token = "") {
                outLine .= " "
                continue
            }
            if (StrLen(token) > breakEvery) {
                broken := ""
                pos := 1
                len := StrLen(token)
                while (pos <= len) {
                    broken .= (broken = "" ? "" : " ") SubStr(token, pos, breakEvery)
                    pos += breakEvery
                }
                outLine .= (outLine = "" ? "" : " ") broken
            } else {
                outLine .= (outLine = "" ? "" : " ") token
            }
        }
        out .= (out = "" ? "" : "`n") outLine
    }
    return out
}

CalcTextLines(text, pxWidth, fontSize) {
    charsPerLine := Max(1, Floor(pxWidth / (fontSize * 0.62)))
    totalLines   := 0
    Loop Parse, text, "`n" {
        segment := StrLen(A_LoopField)
        totalLines += Max(1, Ceil(segment / charsPerLine))
    }
    return Max(1, totalLines)
}

CalcItemHeight(text, pxWidth, fontSize) {
    global MIN_ITEM_H
    lines := Min(3, CalcTextLines(text, pxWidth, fontSize))
    return Max(MIN_ITEM_H, lines * (fontSize + 6) + 16)
}

CalcItemHeightFull(text, pxWidth, fontSize) {
    global MIN_ITEM_H
    lines := Min(10, CalcTextLines(text, pxWidth, fontSize))
    return Max(MIN_ITEM_H, lines * (fontSize + 6) + 24)
}

; Returns text truncated to fit within maxLines, appending "..." if truncated
TruncateToLines(text, pxWidth, fontSize, maxLines := 10) {
    charsPerLine := Max(1, Floor(pxWidth / (fontSize * 0.62)))
    outputLines  := []
    totalLines   := 0
    Loop Parse, text, "`n" {
        segment   := A_LoopField
        segChars  := StrLen(segment)
        segLines  := Max(1, Ceil(segChars / charsPerLine))
        if (totalLines + segLines >= maxLines) {
            remaining := maxLines - totalLines
            if (remaining <= 0)
                break
            ; Fit only what remaining lines allow
            allowed := remaining * charsPerLine
            if (StrLen(segment) > allowed)
                segment := SubStr(segment, 1, Max(1, allowed - 3)) "..."
            outputLines.Push(segment)
            totalLines += remaining
            break
        }
        outputLines.Push(segment)
        totalLines += segLines
    }
    result := ""
    for i, ln in outputLines
        result .= (i > 1 ? "`n" : "") ln
    ; If original had more content than what we kept, ensure "..." suffix
    orig := StrReplace(text, "`r", "")
    if (result != orig && !RegExMatch(result, "\.\.\.$"))
        result := RTrim(result) "..."
    return result
}

TextExceedsBox(text, pxWidth, fontSize) {
    return (CalcTextLines(text, pxWidth, fontSize) > 3)
}

RebuildItems() {
    global ScrollGui, ItemControls, ScrollOffset, ClipHistory, ShowPinned, ITEM_PAD, ExpandedIds, HoveredItemId, HoveredExpandBtnId

    ; Destroy old controls (not just hide) — hiding accumulates them toward the 8192 control limit
    for ic in ItemControls {
        for key in ["sep", "lbl", "expandBtn", "pinIcon", "bg"] {
            if (ic.HasProp(key) && ic.%key% != "") {
                try ic.%key%.Visible := false
                try {
                    hwnd := ic.%key%.Hwnd
                    DllCall("DestroyWindow", "Ptr", hwnd)
                }
            }
        }
    }
    ItemControls := []
    HoveredItemId := 0
    HoveredExpandBtnId := 0

    w := GetScrollW()
    h := GetScrollH()

    display := []
    for item in ClipHistory {
        if (ShowPinned && !item.pinned)
            continue
        display.Push(item)
    }

    global LVDisplay := display

    yPos := ITEM_PAD - ScrollOffset

    for item in display {
        yPos += DrawItem(item, yPos, w)
        yPos += ITEM_PAD
    }

    if (display.Length = 0) {
        lbl := ScrollGui.Add("Text",
            "x0 y20 w" w " h30 Center cAAAAAA", "No clipboard items yet.")
        ItemControls.Push({bg:"", lbl:lbl, sep:"", expandBtn:"", pinIcon:"", yTop:-9999, itemH:30})
    }
}

DrawItem(item, yTop, w) {
    global ScrollGui, ItemControls, ITEM_PAD, ExpandedIds, MIN_ITEM_H, MAX_EXPANDED_CHARS

    expandBtnW := 22
    expandBtnH := 16
    innerL     := 10
    innerR     := ITEM_PAD
    fontSize   := 9

    bgX := ITEM_PAD
    bgW := Max(1, w - ITEM_PAD * 2)

    lblW := Max(1, bgW - innerL - innerR - expandBtnW - 2)

    preview  := StrReplace(item.text, "`r", "")
    preview  := StrReplace(preview,   "`r`n", "`n")
    fullText := WrapLongTokens(preview, lblW, fontSize)

    needsExpand := (!item.HasProp("isImage") || !item.isImage) && TextExceedsBox(fullText, lblW, fontSize)

    isExpanded := false
    for eid in ExpandedIds {
        if (eid = item.id) {
            isExpanded := true
            break
        }
    }

    ; Always truncate to what's actually shown — passing the full (possibly huge)
    ; text as a control caption makes Windows reject control creation outright.
    displayText := fullText
    if (isExpanded && needsExpand) {
        displayText := TruncateToLines(fullText, lblW, fontSize, 10)
        if (StrLen(displayText) > MAX_EXPANDED_CHARS) {
            displayText := SubStr(displayText, 1, MAX_EXPANDED_CHARS - 3) "..."
        }
    }
    else if (!isExpanded)
        displayText := TruncateToLines(fullText, lblW, fontSize, 3)

    if (item.HasProp("isImage") && item.isImage)
        itemH := 110   ; fixed thumbnail height
    else
        itemH := isExpanded
            ? CalcItemHeightFull(displayText, lblW, fontSize)
            : CalcItemHeight(fullText, lblW, fontSize)

    ; Safety clamp: Windows rejects controls taller than ~32700px or shorter than 1px
    itemH := Max(MIN_ITEM_H, Min(itemH, 32700))

    h := GetScrollH()
    if (yTop + itemH < 0 || yTop > h)
        return itemH

    clrStr := item.pinned ? "2D3B55" : "282828"
    ebX    := bgX + bgW - innerR - expandBtnW + 2

    bg := ScrollGui.Add("Text",
        "x" bgX " y" yTop " w" bgW " h" itemH " +0x200 Background" clrStr, "")

    ; Expand button (top-right)
    expandBtn := ""
    if (needsExpand) {
        expandBtn := ScrollGui.Add("Text",
            "x" ebX " y" yTop+4 " w" expandBtnW " h" expandBtnH
            " Center cFFFFFF Background333333 +0x80",
            isExpanded ? "🔺" : "🔻")
        expandBtn.SetFont("s8", "Segoe UI Emoji")
    }

    ; Pin icon below expand button (or at top if no expand)
    pinIcon := ""
    if (item.pinned) {
        piY := needsExpand ? yTop + 4 + expandBtnH + 2 : yTop + 4
        pinIcon := ScrollGui.Add("Text",
            "x" ebX " y" piY " w" expandBtnW " h16 Center cFFFFFF Background" clrStr, "📌")
        pinIcon.SetFont("s6", "Segoe UI Emoji")
    }

    ; Text label or image thumbnail
    if (item.HasProp("isImage") && item.isImage && item.HasProp("thumb") && item.thumb != "" && FileExist(item.thumb)) {
        ; Show thumbnail — fix height, auto width to preserve aspect ratio
        thumbH_px := Max(1, itemH - 16)
        lbl := ScrollGui.Add("Pic",
            "x" bgX+innerL " y" yTop+8 " w-1 h" thumbH_px
            " Background" clrStr, item.thumb)
        DllCall("InvalidateRect", "Ptr", lbl.Hwnd, "Ptr", 0, "Int", 1)
    } else {
        lblH := Max(1, itemH - 16)
        try {
            lbl := ScrollGui.Add("Text",
                "x" bgX+innerL " y" yTop+8 " w" lblW " h" lblH
                " cE0E0E0 Background" clrStr, displayText)
        } catch {
            lbl := ScrollGui.Add("Text",
                "x" bgX+innerL " y" yTop+8 " w" lblW " h" lblH
                " cE0E0E0 Background" clrStr, "[Large item — content too big to preview]")
        }
        lbl.SetFont("s" fontSize, "Segoe UI")
    }

    sep := ScrollGui.Add("Text",
        "x" bgX " y" yTop+itemH " w" bgW " h1 Background2A2A2A")

    ic := {bg:bg, lbl:lbl, sep:sep, expandBtn:expandBtn, pinIcon:pinIcon,
           item:item, yTop:yTop, itemH:itemH}
    ItemControls.Push(ic)

    capturedItem := item
    lbl.OnEvent("Click",      (*) => OnItemLeftClick(capturedItem))
    bg.OnEvent("ContextMenu", (*) => OnItemRightClick(capturedItem))
    lbl.OnEvent("ContextMenu",(*) => OnItemRightClick(capturedItem))
    if (needsExpand && expandBtn != "")
        expandBtn.OnEvent("Click", (*) => ToggleExpand(capturedItem.id))

    return itemH
}

ToggleExpand(id) {
    global ExpandedIds
    for i, eid in ExpandedIds {
        if (eid = id) {
            ExpandedIds.RemoveAt(i)
            RebuildItems()
            return
        }
    }
    ExpandedIds.Push(id)
    RebuildItems()
}

; Lighter tint of the normal card color, used to signal "this is hoverable/clickable"
ItemCardColor(item, hovered) {
    if (item.pinned)
        return hovered ? "3D4E6E" : "2D3B55"
    else
        return hovered ? "383838" : "282828"
}

SetItemHoverState(ic, hovered) {
    clr := ItemCardColor(ic.item, hovered)
    ; Paint the background and label first, and flush that paint synchronously
    ; (UpdateWindow) before touching anything else — otherwise their repaint
    ; can still be pending when we redraw the icons, and land on top of them
    ; afterward.
    for ctrl in [ic.bg, ic.lbl] {
        if (ctrl = "")
            continue
        try {
            ctrl.Opt("Background" clr)
            DllCall("InvalidateRect", "Ptr", ctrl.Hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow",   "Ptr", ctrl.Hwnd)
        }
    }
    ; Now update the pin icon's color and redraw both overlapping foreground
    ; controls last, guaranteed after the background/label paint is done.
    if (ic.pinIcon != "") {
        try ic.pinIcon.Opt("Background" clr)
    }
    for ctrl in [ic.pinIcon, ic.expandBtn] {
        if (ctrl = "")
            continue
        try {
            DllCall("InvalidateRect", "Ptr", ctrl.Hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow",   "Ptr", ctrl.Hwnd)
        }
    }
}

SetExpandBtnHoverState(ic, hovered) {
    if (ic.expandBtn = "")
        return
    clr := ItemCardColor(ic.item, hovered)
    try {
        ic.expandBtn.Opt("Background" clr)
        DllCall("InvalidateRect", "Ptr", ic.expandBtn.Hwnd, "Ptr", 0, "Int", 1)
    }
}

; Header buttons are normally dark gray, except Pin which turns blue while
; "Show Only Pinned" is active — hover should lighten whichever is current.
HeaderBtnColor(name, hovered) {
    global ShowPinned
    isActive := (name = "pin" && ShowPinned)
    if (isActive)
        return hovered ? "3D4E6E" : "2D3B55"
    else
        return hovered ? "2A2A2A" : "1C1C1C"
}

SetHeaderBtnHoverState(btn, name, hovered) {
    if (btn = "")
        return
    try {
        btn.Opt("Background" HeaderBtnColor(name, hovered))
        DllCall("InvalidateRect", "Ptr", btn.Hwnd, "Ptr", 0, "Int", 1)
    }
}

; ============================================================
;  MENU BAR — TOOLTIP HOVER
; ============================================================
OnMouseMove(wParam, lParam, msg, hwnd) {
    global ManagerGui, BtnPin, BtnClear, BtnSettings, BtnWinClip, ShowPinned
    global TooltipLastHwnd, ItemControls, HoveredItemId, HoveredExpandBtnId, HoveredHeaderBtn

    ; Header button hover highlight (📌 ❌ ⚙️ 🪟)
    headerBtns := Map("pin", BtnPin, "clear", BtnClear, "settings", BtnSettings, "winclip", BtnWinClip)
    newHeaderHover := ""
    for name, btn in headerBtns {
        if (btn != "" && hwnd = btn.Hwnd) {
            newHeaderHover := name
            break
        }
    }
    if (newHeaderHover != HoveredHeaderBtn) {
        if (HoveredHeaderBtn != "" && headerBtns.Has(HoveredHeaderBtn))
            SetHeaderBtnHoverState(headerBtns[HoveredHeaderBtn], HoveredHeaderBtn, false)
        if (newHeaderHover != "")
            SetHeaderBtnHoverState(headerBtns[newHeaderHover], newHeaderHover, true)
        HoveredHeaderBtn := newHeaderHover
    }

    ; Expand/collapse button — its own independent hover highlight, not the whole card.
    hitExpandIc := ""
    for ic in ItemControls {
        if (ic.HasProp("item") && ic.expandBtn != "" && hwnd = ic.expandBtn.Hwnd) {
            hitExpandIc := ic
            break
        }
    }
    newExpandHoverId := (hitExpandIc != "") ? hitExpandIc.item.id : 0
    if (newExpandHoverId != HoveredExpandBtnId) {
        if (HoveredExpandBtnId != 0) {
            for ic in ItemControls {
                if (ic.HasProp("item") && ic.item.id = HoveredExpandBtnId) {
                    SetExpandBtnHoverState(ic, false)
                    break
                }
            }
        }
        if (hitExpandIc != "")
            SetExpandBtnHoverState(hitExpandIc, true)
        HoveredExpandBtnId := newExpandHoverId
    }

    ; Card hover highlight — deliberately excludes the expand button (handled above)
    ; so hovering it doesn't also light up the whole card.
    hitIc := ""
    for ic in ItemControls {
        if (!ic.HasProp("item"))
            continue
        if ((ic.bg != "" && hwnd = ic.bg.Hwnd)
            || (ic.lbl != "" && hwnd = ic.lbl.Hwnd)
            || (ic.pinIcon != "" && hwnd = ic.pinIcon.Hwnd)) {
            hitIc := ic
            break
        }
    }
    newHoverId := (hitIc != "") ? hitIc.item.id : 0
    if (newHoverId != HoveredItemId) {
        if (HoveredItemId != 0) {
            for ic in ItemControls {
                if (ic.HasProp("item") && ic.item.id = HoveredItemId) {
                    SetItemHoverState(ic, false)
                    break
                }
            }
        }
        if (hitIc != "")
            SetItemHoverState(hitIc, true)
        HoveredItemId := newHoverId
    }

    if (ManagerGui = "")
        return
    if (hwnd = TooltipLastHwnd)
        return
    TooltipLastHwnd := hwnd
    ToolTip("")
    if (BtnPin != "" && hwnd = BtnPin.Hwnd)
        ShowTooltip(ShowPinned ? "Show All" : "Show Only Pinned")
    else if (BtnClear != "" && hwnd = BtnClear.Hwnd)
        ShowTooltip("Clear All Unpinned")
    else if (BtnSettings != "" && hwnd = BtnSettings.Hwnd)
        ShowTooltip("Settings")
    else if (BtnWinClip != "" && hwnd = BtnWinClip.Hwnd)
        ShowTooltip("Windows 11 Clipboard")
}

ResetTooltipCache() {
    global TooltipLastHwnd
    TooltipLastHwnd := 0
}

; ============================================================
;  SCROLL
; ============================================================
OnMouseWheel(wParam, lParam, msg, hwnd) {
    global ManagerGui, ScrollGui, ScrollOffset, ClipHistory, ShowPinned
    global ITEM_PAD, ExpandedIds

    if (ManagerGui = "")
        return

    ; Accept wheel on parent or child
    guiHwnd    := ManagerGui.Hwnd
    scrollHwnd := ScrollGui.Hwnd
    chk := hwnd
    Loop 8 {
        if (chk = guiHwnd || chk = scrollHwnd)
            break
        chk := DllCall("GetParent", "Ptr", chk, "Ptr")
        if (chk = 0)
            return
    }
    if (chk != guiHwnd && chk != scrollHwnd)
        return

    delta := (wParam >> 16) & 0xFFFF
    if (delta > 32767)
        delta -= 65536

    if (delta > 0) {
        ScrollOffset := Max(0, ScrollOffset - 60)
    } else {
        w          := GetScrollW()
        h          := GetScrollH()
        pinColW    := 22
        expandBtnW := 22
        innerL     := 10
        innerR     := ITEM_PAD
        bgW        := w - ITEM_PAD * 2
        textW      := bgW - innerL - innerR - expandBtnW - 2
        totalH     := ITEM_PAD
        for item in ClipHistory {
            if (ShowPinned && !item.pinned)
                continue
            preview := StrReplace(item.text, "`r", "")
            preview := StrReplace(preview,   "`r`n", "`n")
            preview := WrapLongTokens(preview, textW, 9)
            isExpanded := false
            for eid in ExpandedIds {
                if (eid = item.id) {
                    isExpanded := true
                    break
                }
            }
            ih := isExpanded
                ? CalcItemHeightFull(TruncateToLines(preview, textW, 9, 10), textW, 9)
                : CalcItemHeight(preview, textW, 9)
            ih := Min(ih, 32700)
            totalH += ih + ITEM_PAD
        }
        maxScroll    := Max(0, totalH - h + ITEM_PAD)
        ScrollOffset := Min(maxScroll, ScrollOffset + 60)
    }

    RebuildItems()
    return 0
}


; ============================================================
;  HELPER — center a dialog over the main window
; ============================================================
CenterOnGui(parentGui, dlgW, dlgH) {
    parentGui.GetPos(&px, &py, &pw, &ph)
    x := px + (pw - dlgW) // 2
    y := py + (ph - dlgH) // 2
    return "x" x " y" y
}

; ============================================================
;  MENU BAR BUTTON HANDLERS
; ============================================================

; 📌  Show Only Pinned — highlighted when active
OnBtnPin() {
    global ShowPinned, ScrollOffset, BtnPin, HoveredHeaderBtn
    ShowPinned   := !ShowPinned
    ScrollOffset := 0
    BtnPin.Opt("Background" HeaderBtnColor("pin", HoveredHeaderBtn = "pin"))
    DllCall("InvalidateRect", "Ptr", BtnPin.Hwnd, "Ptr", 0, "Int", 1)
    DllCall("UpdateWindow",   "Ptr", BtnPin.Hwnd)
    ; Update tooltip immediately while still hovering
    ShowTooltip(ShowPinned ? "Show All" : "Show Only Pinned")
    ResetTooltipCache()
    RebuildItems()
}

; ❌  Clear All Unpinned — confirmation window
OnBtnClear() {
    global ManagerGui
    ToolTip("")
    dlgW := 300, dlgH := 110
    dlg := Gui("+Owner" ManagerGui.Hwnd " -MaximizeBox -MinimizeBox", "Clear Clipboard")
    dlg.BackColor := "1C1C1C"
    dlg.SetFont("s9 cE0E0E0", "Segoe UI")
    dlg.MarginX := 0
    dlg.MarginY := 0
    dlg.Add("Text", "x0 y20 w" dlgW " Center", "Clear all unpinned items from clipboard history?")
    btnYes := dlg.Add("Button", "x70 y60 w70 h28", "Yes")
    btnNo  := dlg.Add("Button", "x160 y60 w70 h28", "No")
    closeDlg := () => (ManagerGui.Opt("-Disabled"), dlg.Destroy())
    btnYes.OnEvent("Click", (*) => (closeDlg(), DoClearAll()))
    btnNo.OnEvent("Click",  (*) => closeDlg())
    dlg.OnEvent("Close",    (*) => closeDlg())
    dlg.OnEvent("Escape",   (*) => closeDlg())
    ManagerGui.Opt("+Disabled")
    dlg.Show("w" dlgW " h" dlgH " " CenterOnGui(ManagerGui, dlgW, dlgH))
}

DoClearAll() {
    global ClipHistory, ScrollOffset, THUMB_DIR
    newHistory := []
    for item in ClipHistory {
        if (item.pinned)
            newHistory.Push(item)
        else
            DeleteItemThumb(item)
    }
    ClipHistory  := newHistory
    ScrollOffset := 0
    ; Purge any leftover PNG files in Thumbs that no longer belong to any item
    PurgeOrphanedThumbs()
    RebuildItems()
}

; ⚙️  Settings window
OnBtnSettings() {
    global ManagerGui, PlainTextMode, KeepOpen, MaxItems
    ToolTip("")
    dlgW := 280, dlgH := 190
    pad  := 24    ; left/right padding
    ctlW := dlgW - pad * 2   ; 232px — all controls same width & x
    dlg := Gui("+Owner" ManagerGui.Hwnd " -MaximizeBox -MinimizeBox", "Settings")
    dlg.BackColor := "1C1C1C"
    dlg.SetFont("s9 cE0E0E0", "Segoe UI")
    cbPlain := dlg.Add("Checkbox", "x" pad " y20 w" ctlW " Background1C1C1C cE0E0E0",
        "Plain Text Mode")
    cbPlain.Value := PlainTextMode ? 1 : 0
    cbKeep := dlg.Add("Checkbox", "x" pad " y46 w" ctlW " Background1C1C1C cE0E0E0",
        "Keep Open After Paste")
    cbKeep.Value := KeepOpen ? 1 : 0
    ; "Set Maximum Items" row: label + edit on same line
    lblMax := dlg.Add("Text",  "x" pad " y76 w160 h22 +0x200", "Set Maximum Items")
    txtMax := dlg.Add("Edit",  "x" pad+160 " y74 w" ctlW-160 " c000000", MaxItems)
    btnExport := dlg.Add("Button", "x" pad " y106 w" (ctlW//2 - 4) " h28", "Export Backup")
    btnImport := dlg.Add("Button", "x" pad+(ctlW//2+4) " y106 w" (ctlW//2 - 4) " h28", "Import Backup")
    btnSave    := dlg.Add("Button", "x" pad " y142 w" (ctlW//2 - 4) " h28", "Save")
    btnDiscard := dlg.Add("Button", "x" pad+(ctlW//2+4) " y142 w" (ctlW//2 - 4) " h28", "Discard")
    closeDlg := () => (ManagerGui.Opt("-Disabled"), dlg.Destroy())
    btnSave.OnEvent("Click", (*) => SaveSettings(dlg, cbPlain, cbKeep, txtMax, closeDlg))
    btnDiscard.OnEvent("Click", (*) => closeDlg())
    btnExport.OnEvent("Click", (*) => ExportBackup(dlg))
    btnImport.OnEvent("Click", (*) => ImportBackup(dlg, closeDlg))
    dlg.OnEvent("Close",  (*) => closeDlg())
    dlg.OnEvent("Escape", (*) => closeDlg())
    ManagerGui.Opt("+Disabled")
    dlg.Show("w" dlgW " h" dlgH " " CenterOnGui(ManagerGui, dlgW, dlgH))
}

SaveSettings(dlg, cbPlain, cbKeep, txtMax, closeDlg) {
    global PlainTextMode, KeepOpen, MaxItems
    PlainTextMode := cbPlain.Value = 1
    KeepOpen      := cbKeep.Value  = 1
    n := Integer(txtMax.Value)
    if (n >= 1 && n <= 200)
        MaxItems := n
    closeDlg()
}

; ============================================================
;  BACKUP — EXPORT / IMPORT (pinned + unpinned, full text, images)
; ============================================================
; A plain .ini can't safely hold this data on its own — full clip text and
; binary clipboard/image data are stored as separate files precisely because
; INI values are capped at ~32KB (that's what used to corrupt the history and
; wipe pins on large copies). So a backup bundles the INI *and* its data
; folder together into a single .zip, which preserves everything losslessly.
ExportBackup(dlg) {
    global INI_FILE, CLIP_DIR, ManagerGui
    SaveConfig()   ; flush current in-memory state (including pins) to disk first
    defaultName := "ClipboardPlus_Backup_" FormatTime(, "yyyyMMdd_HHmmss") ".zip"
    try savePath := FileSelect("S16", defaultName, "Export Clipboard Plus Backup", "Zip Files (*.zip)")
    if (!IsSet(savePath) || savePath = "")
        return
    if !InStr(savePath, ".zip")
        savePath .= ".zip"
    try FileDelete(savePath)
    psCmd := "Compress-Archive -Path '" INI_FILE "','" CLIP_DIR "' -DestinationPath '" savePath "' -Force"
    RunWait('powershell.exe -NoProfile -WindowStyle Hidden -Command "' psCmd '"', , "Hide")
    if FileExist(savePath)
        MsgBox("Backup exported to:`n" savePath, "Export Complete", "Iconi")
    else
        MsgBox("Export failed. Please try again.", "Export Failed", "IconX")
}

ImportBackup(dlg, closeDlg) {
    global INI_FILE, CLIP_DIR, ManagerGui
    try zipPath := FileSelect(1, , "Import Clipboard Plus Backup", "Zip Files (*.zip)")
    if (!IsSet(zipPath) || zipPath = "")
        return
    result := MsgBox("Importing will replace your current clipboard history, including pinned items. Continue?",
        "Import Backup", "YesNo Icon!")
    if (result != "Yes")
        return
    tempDir := A_Temp "\ClipboardPlusImport_" A_TickCount
    try DirCreate(tempDir)
    psCmd := "Expand-Archive -Path '" zipPath "' -DestinationPath '" tempDir "' -Force"
    RunWait('powershell.exe -NoProfile -WindowStyle Hidden -Command "' psCmd '"', , "Hide")
    importedIni  := tempDir "\ClipboardManager.ini"
    importedData := tempDir "\ClipboardData"
    if !FileExist(importedIni) {
        MsgBox("That file doesn't look like a valid Clipboard Plus backup.", "Import Failed", "IconX")
        try DirDelete(tempDir, true)
        return
    }
    try FileDelete(INI_FILE)
    try FileCopy(importedIni, INI_FILE, true)
    try DirDelete(CLIP_DIR, true)
    if DirExist(importedData)
        try DirCopy(importedData, CLIP_DIR, true)
    try DirDelete(tempDir, true)
    LoadConfig()
    RebuildItems()
    MsgBox("Backup imported successfully.", "Import Complete", "Iconi")
    closeDlg()
}

; 🪟  Open Windows 11 clipboard
OnBtnWinClip() {
    ToolTip("")
    Send("#v")
}

ClearAllItems(*) {
    global ClipHistory, ScrollOffset
    newHistory := []
    for item in ClipHistory {
        if (item.pinned)
            newHistory.Push(item)
    }
    ClipHistory  := newHistory
    ScrollOffset := 0
    RebuildItems()
}

; ============================================================
;  ITEM INTERACTIONS
; ============================================================
OnItemLeftClick(item) {
    global LastClipText, ManagerGui, KeepOpen, ClipHistory, PlainTextMode, IgnoreNextClip
    LastClipText   := item.text
    IgnoreNextClip := 2
    if (!PlainTextMode && item.HasProp("clip") && item.clip != "" && !(item.clip is String))
        A_Clipboard := item.clip   ; restore full formatting
    else
        A_Clipboard := item.text   ; plain text only
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].id = item.id) {
            itm := ClipHistory.RemoveAt(A_Index)
            ClipHistory.InsertAt(1, itm)
            break
        }
    }
    if (!KeepOpen)
        ManagerGui.Hide()
    Sleep(80)
    Send("^v")
}

OnItemRightClick(item) {
    global ExpandedIds
    isExpanded := false
    for eid in ExpandedIds {
        if (eid = item.id) {
            isExpanded := true
            break
        }
    }
    m := Menu()
    m.Add(item.pinned ? "Unpin" : "Pin",      (*) => TogglePin(item))
    m.Add("Delete",                            (*) => DeleteItem(item))
    m.Add()
    m.Add(isExpanded ? "Collapse" : "Expand", (*) => ToggleExpandItem(item))
    m.Add()
    m.Add("Move Up",   (*) => MoveItem(item, -1))
    m.Add("Move Down", (*) => MoveItem(item, +1))
    m.Show()
}

ToggleExpandItem(item) {
    global ExpandedIds
    for i, eid in ExpandedIds {
        if (eid = item.id) {
            ExpandedIds.RemoveAt(i)
            RebuildItems()
            return
        }
    }
    ExpandedIds.Push(item.id)
    RebuildItems()
}

TogglePin(item) {
    global ClipHistory
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].id = item.id) {
            ClipHistory[A_Index].pinned := !ClipHistory[A_Index].pinned
            break
        }
    }
    RebuildItems()
}

DeleteItem(item) {
    global ClipHistory
    DeleteItemThumb(item)
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].id = item.id) {
            ClipHistory.RemoveAt(A_Index)
            break
        }
    }
    RebuildItems()
}

DeleteItemThumb(item) {
    if (item.HasProp("isImage") && item.isImage && item.HasProp("thumb") && item.thumb != "")
        try FileDelete(item.thumb)
}

; Delete any PNG files in THUMB_DIR not referenced by any current ClipHistory item
PurgeOrphanedThumbs() {
    global THUMB_DIR, ClipHistory
    if !DirExist(THUMB_DIR)
        return
    ; Collect all thumb paths still in use
    usedThumbs := Map()
    for item in ClipHistory {
        if (item.HasProp("thumb") && item.thumb != "")
            usedThumbs[item.thumb] := true
    }
    Loop Files, THUMB_DIR "\*.png" {
        if !usedThumbs.Has(A_LoopFileFullPath)
            try FileDelete(A_LoopFileFullPath)
    }
}

MoveItem(item, delta) {
    global ClipHistory
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].id = item.id) {
            cur  := A_Index
            dest := cur + delta
            if (dest < 1 || dest > ClipHistory.Length)
                return
            tmp               := ClipHistory[cur]
            ClipHistory[cur]  := ClipHistory[dest]
            ClipHistory[dest] := tmp
            RebuildItems()
            return
        }
    }
}

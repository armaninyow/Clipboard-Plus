;@Ahk2Exe-SetMainIcon icon.ico
#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
;  CLIPBOARD PLUS  |  AHK v2  |  Hotkey: Win+V
;  Custom drawn items inside a child-Gui scroll viewport
;  Child Gui clips its own children — menu bar never overlapped
;  Settings + history persisted to ClipboardManager.ini
; ============================================================

; ---------- Config ----------
global INI_FILE      := A_ScriptDir "\ClipboardManager.ini"
global CLIP_DIR      := A_ScriptDir "\ClipboardData"
global THUMB_DIR     := A_ScriptDir "\ClipboardData\Thumbs"
global MaxItems      := 25
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
    count := 0
    try count := Integer(IniRead(INI_FILE, "History", "Count", 0))
    Loop count {
        try {
            text   := IniRead(INI_FILE, "History", "Text"   . A_Index)
            pinned := Integer(IniRead(INI_FILE, "History", "Pinned" . A_Index, 0)) = 1
            text   := StrReplace(text, "\n", "`n")
            text   := StrReplace(text, "\r", "`r")
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
        item := ClipHistory[A_Index]
        safe := StrReplace(item.text, "`n", "\n")
        safe := StrReplace(safe,      "`r", "\r")
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
    }
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
A_IconTip := "Clipboard Plus"
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
    global BtnPin, ShowPinned
    if (BtnPin = "")
        return
    BtnPin.Opt("Background" (ShowPinned ? "2D3B55" : "1C1C1C"))
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
    text := A_Clipboard
    if (text = "" || text = LastClipText)
        return
    LastClipText := text
    if (PlainTextMode) {
        A_Clipboard := text
        text := A_Clipboard
    }
    Loop ClipHistory.Length {
        if (ClipHistory[A_Index].text = text) {
            moved := ClipHistory.RemoveAt(A_Index)
            moved.clip := PlainTextMode ? text : ClipboardAll()
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
    clip := PlainTextMode ? text : ClipboardAll()
    ClipHistory.InsertAt(1, {text: text, clip: clip, pinned: false, id: NextID++})
    if (ManagerGui != "" && WinExist("ahk_id " ManagerGui.Hwnd))
        RebuildItems()
    ShowCopyTooltip()
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

    ManagerGui := Gui("+Resize +MinSize240x200 -MaximizeBox -MinimizeBox -DPIScale", "Clipboard Plus")
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
    BtnPin.OnEvent("DoubleClick",  (*) => OnBtnPin())

    BtnClear := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "❌")
    BtnClear.OnEvent("Click",       (*) => OnBtnClear())
    BtnClear.OnEvent("DoubleClick", (*) => OnBtnClear())

    BtnSettings := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "⚙️")
    BtnSettings.OnEvent("Click",       (*) => OnBtnSettings())
    BtnSettings.OnEvent("DoubleClick", (*) => OnBtnSettings())

    BtnWinClip := ManagerGui.Add("Text",
        "x0 y0 w" btnW " h" MENUBAR_H " +0x200 +0x80 Center Background1C1C1C cFFFFFF", "🪟")
    BtnWinClip.OnEvent("Click",       (*) => OnBtnWinClip())
    BtnWinClip.OnEvent("DoubleClick", (*) => OnBtnWinClip())

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
    lines := CalcTextLines(text, pxWidth, fontSize)
    return Max(MIN_ITEM_H, lines * (fontSize + 6) + 24)
}

TextExceedsBox(text, pxWidth, fontSize) {
    return (CalcTextLines(text, pxWidth, fontSize) > 3)
}

RebuildItems() {
    global ScrollGui, ItemControls, ScrollOffset, ClipHistory, ShowPinned, ITEM_PAD, ExpandedIds

    ; Hide old controls
    for ic in ItemControls {
        for key in ["bg", "lbl", "sep", "expandBtn", "pinIcon"] {
            if (ic.HasProp(key) && ic.%key% != "")
                try ic.%key%.Visible := false
        }
    }
    ItemControls := []

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
    global ScrollGui, ItemControls, ITEM_PAD, ExpandedIds, MIN_ITEM_H

    expandBtnW := 22
    expandBtnH := 16
    innerL     := 10
    innerR     := ITEM_PAD
    fontSize   := 9

    bgX := ITEM_PAD
    bgW := w - ITEM_PAD * 2

    lblW := bgW - innerL - innerR - expandBtnW - 2

    preview  := StrReplace(item.text, "`r", "")
    preview  := StrReplace(preview,   "`r`n", "`n")
    fullText := preview

    needsExpand := (!item.HasProp("isImage") || !item.isImage) && TextExceedsBox(fullText, lblW, fontSize)

    isExpanded := false
    for eid in ExpandedIds {
        if (eid = item.id) {
            isExpanded := true
            break
        }
    }

    if (item.HasProp("isImage") && item.isImage)
        itemH := 110   ; fixed thumbnail height
    else
        itemH := isExpanded
            ? CalcItemHeightFull(fullText, lblW, fontSize)
            : CalcItemHeight(fullText, lblW, fontSize)

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
        thumbH_px := itemH - 16
        lbl := ScrollGui.Add("Pic",
            "x" bgX+innerL " y" yTop+8 " w-1 h" thumbH_px
            " Background" clrStr, item.thumb)
        DllCall("InvalidateRect", "Ptr", lbl.Hwnd, "Ptr", 0, "Int", 1)
    } else {
        lbl := ScrollGui.Add("Text",
            "x" bgX+innerL " y" yTop+8 " w" lblW " h" itemH-16
            " cE0E0E0 Background" clrStr, fullText)
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

; ============================================================
;  MENU BAR — TOOLTIP HOVER
; ============================================================
OnMouseMove(wParam, lParam, msg, hwnd) {
    global ManagerGui, BtnPin, BtnClear, BtnSettings, BtnWinClip, ShowPinned
    global TooltipLastHwnd
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
            isExpanded := false
            for eid in ExpandedIds {
                if (eid = item.id) {
                    isExpanded := true
                    break
                }
            }
            ih := isExpanded
                ? CalcItemHeightFull(preview, textW, 9)
                : CalcItemHeight(preview, textW, 9)
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
    global ShowPinned, ScrollOffset, BtnPin
    ShowPinned   := !ShowPinned
    ScrollOffset := 0
    BtnPin.Opt("Background" (ShowPinned ? "2D3B55" : "1C1C1C"))
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
    global ClipHistory, ScrollOffset
    newHistory := []
    for item in ClipHistory {
        if (item.pinned)
            newHistory.Push(item)
        else
            DeleteItemThumb(item)
    }
    ClipHistory  := newHistory
    ScrollOffset := 0
    RebuildItems()
}

; ⚙️  Settings window
OnBtnSettings() {
    global ManagerGui, PlainTextMode, KeepOpen, MaxItems
    ToolTip("")
    dlgW := 280, dlgH := 160
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
    btnSave    := dlg.Add("Button", "x" pad " y112 w" (ctlW//2 - 4) " h28", "Save")
    btnDiscard := dlg.Add("Button", "x" pad+(ctlW//2+4) " y112 w" (ctlW//2 - 4) " h28", "Discard")
    closeDlg := () => (ManagerGui.Opt("-Disabled"), dlg.Destroy())
    btnSave.OnEvent("Click", (*) => SaveSettings(dlg, cbPlain, cbKeep, txtMax, closeDlg))
    btnDiscard.OnEvent("Click", (*) => closeDlg())
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

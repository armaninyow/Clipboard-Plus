# 📋 Clipboard Plus



Clipboard Plus is a modern clipboard history manager for Windows. Replaces the default Win+V clipboard with a dark-themed, scrollable history panel that supports rich text, images, pinning, and persistent storage across restarts.

---

## Features

- **Win+V hotkey**: opens the clipboard history panel near your mouse cursor; press again to close it
- **Text & image support**: captures both text and image clipboard entries
- **Image thumbnails**: displays thumbnails for image entries with correct alpha/transparency handling
- **Rich text paste**: restores original formatting (RTF, HTML, etc.) when pasting, not just plain text
- **Plain Text Mode**: optionally strip formatting and paste as plain text only
- **Pin items**: pin important entries so they survive "Clear All" and stay at the top when filtering
- **Show Only Pinned**: filter the list to pinned items only
- **Expand / Collapse**: expand long text entries to see the full content
- **Persistent history**: clipboard history, images, and settings are saved to disk and restored on next launch
- **Multi-monitor aware**: window opens at the cursor and flips to stay on screen near edges
- **Resizable window**: drag to resize; items reflow automatically
- **Context menu**: right-click any item to Pin/Unpin, Delete, Expand/Collapse, Move Up/Down
- **Settings**: Plain Text Mode, Keep Open After Paste, Max Items (1–200)
- **Dark theme**: clean dark UI throughout

---

## Usage

| Action | Result |
|---|---|
| `Win+V` | Open / close the clipboard panel |
| Click an item | Paste it into the active window |
| Right-click an item | Pin, Delete, Expand/Collapse, Move Up/Down |
| `📌` button | Toggle Show Only Pinned |
| `❌` button | Clear all unpinned items (with confirmation) |
| `⚙️` button | Open Settings |
| `🪟` button | Open Windows 11 native clipboard |
| Scroll wheel | Scroll through history |
| Resize window | Items reflow to new width |

---

## Settings

Open the ⚙️ settings panel from the toolbar:

| Setting | Description |
|---|---|
| **Plain Text Mode** | Strip all formatting; always paste as plain text |
| **Keep Open After Paste** | Don't hide the panel after clicking an item |
| **Max Items** | Maximum number of clipboard entries to keep (1–200, default 25) |

---

## File Structure

The following files and folders are created next to the `.exe` on first run:

```
ClipboardPlus.exe
ClipboardManager.ini        ← settings and text history
ClipboardData\
├── 1.clip, 2.clip, ...     ← full clipboard data per item
└── Thumbs\
    └── 1.png, 2.png, ...   ← image thumbnails
```

---

## Installation

1. Download the latest `.exe` from [Releases](https://github.com/armaninyow/Clipboard-Plus/releases)
2. Place the `.exe` anywhere you like and run it

No installer needed. No external files required.

---

## Requirements

* **Operating System**: Windows 10 or Windows 11.
* **Standalone**: No installation or external assets required.

---

## License

### CC0 1.0 Universal
This project is licensed under the **Creative Commons Legal Code CC0 1.0 Universal**. 

To the extent possible under law, the author(s) have dedicated all copyright and related and neighboring rights to this software to the public domain worldwide. This software is distributed without any warranty.

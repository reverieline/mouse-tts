# Mouse TTS

A lightweight Windows utility that copies and reads selected text aloud when you press a configured mouse button. 

![Platform](https://img.shields.io/badge/platform-Windows-blue) ![Python](https://img.shields.io/badge/python-3.8%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green)

## Features

- **Global mouse hook** — works in any application
- **Configurable trigger** — any mouse button (left, right, middle, X1, X2) with optional Ctrl / Shift / Alt modifiers
- **Event suppression** — optionally block the trigger button from reaching other apps (usefull for X1/X2 side buttons)
- **Voice settings** — choose from any installed Windows SAPI voice; adjust rate and pitch
- **Text filters** — regex-based exclusion patterns to skip brackets, URLs, or any other unwanted text
- **Clipboard safe** — saves and restores your clipboard content around each trigger
- **System tray** — minimize to tray and keep listening in the background
- **Autostart** — optional Windows Registry entry to launch on login

## Requirements

- Windows 10 / 11
- Python 3.8+ (for running from source)
- At least one Windows SAPI voice installed (built-in voices work fine)

## Installation

### Option A — Standalone executable (no Python needed)

1. Download `mouse_tts.exe` from Release.
2. Run it — no installation required.

### Option B — Run from source

```bash
git clone https://github.com/reverieline/mouse-tts.git
cd mouse-tts
pip install -r requirements.txt
python mouse_tts.py
```

### Option C — Build the executable yourself

```bash
git clone https://github.com/reverieline/mouse-tts.git
cd mouse-tts
build.bat
```

The compiled binary will be at `dist\mouse_tts.exe`.

## Usage

1. Launch `mouse_tts.exe` (or `python mouse_tts.py`).
2. In the **Button** section, click **Detect…** and press the mouse button (+ any modifier keys) you want to use as your trigger.
3. Choose a **Voice** and adjust **Rate** / **Pitch** to taste.
4. Optionally add **Exclude patterns** (one regex per line) for text you want silently skipped.
5. Click **Save & Apply** — the app starts listening immediately.
6. Select text anywhere and press your trigger button to hear it read aloud.
7. Minimize the window to minimize to the system tray. Right-click the tray icon to open settings or quit.

## Safety notes

- **Left / right button suppression without a modifier is blocked** — suppressing those without a modifier would make your primary buttons unusable.
- The hook runs with minimal overhead in a dedicated thread and does not poll continuously.

## Dependencies

| Package       | Purpose                       |
| ------------- | ----------------------------- |
| `pywin32`     | Windows SAPI COM voice access |
| `pynput`      | Mouse input capture           |
| `pyperclip`   | Clipboard read/write          |
| `pystray`     | System tray icon              |
| `Pillow`      | Icon image handling           |
| `pyinstaller` | Build standalone `.exe`       |

## License

MIT © [Alex Gavr](https://github.com/reverieline) — see [LICENSE](LICENSE) for details.

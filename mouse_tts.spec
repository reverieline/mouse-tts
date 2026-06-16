# PyInstaller spec for mouse_tts — single windowless .exe with GUI tray
# Run from the windows\ directory: pyinstaller mouse_tts.spec

from PyInstaller.utils.hooks import collect_all

pynput_datas, pynput_bins, pynput_hidden = collect_all("pynput")
pystray_datas, pystray_bins, pystray_hidden = collect_all("pystray")

block_cipher = None

a = Analysis(
    ["mouse_tts.py"],
    pathex=[],
    binaries=pynput_bins + pystray_bins,
    datas=pynput_datas + pystray_datas + [("icon.ico", ".")],
    hiddenimports=pynput_hidden + pystray_hidden + [
        "win32com",
        "win32com.client",
        "win32com.client.gencache",
        "win32com.server",
        "win32com.server.util",
        "win32api",
        "win32con",
        "pywintypes",
        "win32clipboard",
        "PIL._tkinter_finder",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="mouse_tts",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

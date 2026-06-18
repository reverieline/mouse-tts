"""Global aux-button text-to-speech helper for Windows — GUI edition.

Opens a settings window on launch. Closing minimizes to the system tray.
The mouse hook runs silently in the background.
"""

from __future__ import annotations

import configparser
import ctypes
import html
import logging
import re
import sys
import threading
import time
import winreg
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from ctypes import wintypes
from tkinter import ttk, messagebox, scrolledtext
import tkinter as tk

import pyperclip
import pystray
from PIL import Image, ImageDraw
from pynput import keyboard, mouse

# ── constants ────────────────────────────────────────────────────────────────

CONFIG_FILENAME = "config.ini"
LOG_FILENAME = "mouse_tts.log"
AUTOSTART_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_APP_NAME = "MouseTTS"
XMLNS_SPEECH = "http://www.w3.org/2001/10/synthesis"
SVSF_ASYNC = 1
SVSF_PURGE_BEFORE_SPEAK = 2
SVSF_IS_XML = 8
DEFAULT_RATE = 0
DEFAULT_PITCH = 0
DEFAULT_BUTTON = "x2"
DEFAULT_MODIFIER = "none"
DEFAULT_SUPPRESS_TRIGGER = True
DEFAULT_EXCLUDE_PATTERN_STRINGS: tuple[str, ...] = (
    r"\[.*\]",
    r"(?:https?://|www\.)\S+",
)

LOGGER = logging.getLogger("mouse_tts")
LOGGING_CONFIGURED = False

ERROR_ALREADY_EXISTS = 183
SINGLE_INSTANCE_MUTEX = "MouseTTS_SingleInstance"

WH_MOUSE_LL = 14
HC_ACTION = 0
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_MBUTTONDOWN = 0x0207
WM_MBUTTONUP = 0x0208
WM_XBUTTONDOWN = 0x020B
WM_XBUTTONUP = 0x020C
WM_QUIT = 0x0012
XBUTTON1 = 0x0001
XBUTTON2 = 0x0002

BUTTON_ALIASES = {
    "left": mouse.Button.left,
    "right": mouse.Button.right,
    "middle": mouse.Button.middle,
    "x1": mouse.Button.x1,
    "x2": mouse.Button.x2,
    "xbutton1": mouse.Button.x1,
    "xbutton2": mouse.Button.x2,
    "thumb1": mouse.Button.x1,
    "thumb2": mouse.Button.x2,
    "side1": mouse.Button.x1,
    "side2": mouse.Button.x2,
}

MODIFIER_VK_CODES = {
    "ctrl": (0x11, 0xA2, 0xA3),
    "shift": (0x10, 0xA0, 0xA1),
    "alt": (0x12, 0xA4, 0xA5),
}

# ── Win32 setup ───────────────────────────────────────────────────────────────

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

kernel32.GetCurrentThreadId.restype = wintypes.DWORD
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = ctypes.c_void_p
kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
kernel32.CreateMutexW.restype = ctypes.c_void_p

user32.SetWindowsHookExW.argtypes = [ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD]
user32.SetWindowsHookExW.restype = ctypes.c_void_p
user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
user32.CallNextHookEx.restype = ctypes.c_ssize_t
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), ctypes.c_void_p, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.TranslateMessage.restype = wintypes.BOOL
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.restype = ctypes.c_ssize_t
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostThreadMessageW.restype = wintypes.BOOL


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


LowLevelMouseProc = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

# ── data class ────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Settings:
    button_name: str
    modifiers: tuple[str, ...]
    suppress: bool
    voice_id: str
    voice_name: str
    rate: int
    pitch: int
    launch_minimized: bool
    exclude_pattern_strings: tuple[str, ...]
    exclude_patterns: tuple[re.Pattern, ...]

    @property
    def button(self) -> mouse.Button:
        return BUTTON_ALIASES[self.button_name]

# ── core: speech ──────────────────────────────────────────────────────────────


class SpeechController:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._generation = 0
        self._active_cancel: threading.Event | None = None

    def begin_trigger(self) -> int:
        with self._lock:
            self._generation += 1
            generation = self._generation
            active_cancel = self._active_cancel
        if active_cancel is not None:
            active_cancel.set()
        return generation

    def is_current(self, generation: int) -> bool:
        with self._lock:
            return generation == self._generation

    def speak(self, settings: Settings, text: str, generation: int) -> bool:
        import win32com.client

        sanitized = normalize_for_speech(text)
        sanitized = filter_for_speech(sanitized, settings.exclude_patterns)
        if not sanitized or not self.is_current(generation):
            return False

        stop_event = threading.Event()
        with self._lock:
            if generation != self._generation:
                return False
            self._active_cancel = stop_event

        try:
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Rate = settings.rate
            token = find_voice_token(speaker, settings)
            if token is not None:
                speaker.Voice = token

            xml = (
                f'<speak version="1.0" xmlns="{XMLNS_SPEECH}">'
                f'<prosody pitch="{settings.pitch:+d}">'
                f"{html.escape(sanitized)}</prosody></speak>"
            )
            try:
                speaker.Speak(xml, SVSF_ASYNC | SVSF_IS_XML)
            except Exception:
                LOGGER.exception("SAPI XML speak failed; retrying without XML")
                speaker.Speak(sanitized, SVSF_ASYNC)

            while True:
                if stop_event.is_set() or not self.is_current(generation):
                    try:
                        speaker.Speak("", SVSF_ASYNC | SVSF_PURGE_BEFORE_SPEAK)
                    except Exception:
                        LOGGER.exception("Failed to purge SAPI speech")
                    return False
                try:
                    if speaker.WaitUntilDone(10):
                        return True
                except Exception:
                    LOGGER.exception("SAPI WaitUntilDone failed")
                    return False
        finally:
            with self._lock:
                if self._active_cancel is stop_event:
                    self._active_cancel = None

# ── core: clipboard ───────────────────────────────────────────────────────────


class ClipboardSession:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._generation = 0
        self._saved: str | None = None

    def begin_trigger(self, generation: int) -> None:
        with self._lock:
            if self._generation == 0:
                self._saved = safe_paste()
            self._generation = generation

    def finish_trigger(self, generation: int) -> None:
        with self._lock:
            if generation != self._generation:
                return
            saved = self._saved
            self._generation = 0
            self._saved = None
        if saved is not None:
            safe_copy(saved)

# ── core: mouse hook ──────────────────────────────────────────────────────────


class MouseTriggerHook:
    def __init__(
        self,
        settings: Settings,
        speech: SpeechController,
        clipboard: ClipboardSession,
        suppress: bool,
    ) -> None:
        self.settings = settings
        self.speech = speech
        self.clipboard = clipboard
        self.suppress = suppress
        self._hook_handle = None
        self._thread_id = 0
        self._stop_event = threading.Event()
        self._proc = LowLevelMouseProc(self._hook_proc)

    def run(self) -> None:
        self._thread_id = kernel32.GetCurrentThreadId()
        self._stop_event.clear()
        hinstance = kernel32.GetModuleHandleW(None)
        hook = user32.SetWindowsHookExW(WH_MOUSE_LL, self._proc, hinstance, 0)
        if not hook:
            raise OSError(ctypes.get_last_error(), "Failed to install mouse hook")
        self._hook_handle = hook
        msg = wintypes.MSG()
        try:
            while not self._stop_event.is_set():
                r = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if r == 0:
                    break
                if r == -1:
                    raise OSError(ctypes.get_last_error(), "Hook message loop failed")
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            if self._hook_handle:
                user32.UnhookWindowsHookEx(self._hook_handle)
                self._hook_handle = None

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread_id:
            user32.PostThreadMessageW(self._thread_id, WM_QUIT, 0, 0)

    def _hook_proc(self, n_code: int, w_param: int, l_param: int) -> int:
        if n_code == HC_ACTION:
            md = ctypes.cast(l_param, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
            btn = _btn_for_msg(int(w_param), md.mouseData)
            if btn is not None and btn == self.settings.button:
                modifier_ok = _modifier_active(self.settings.modifiers)
                if _is_down(int(w_param)) and modifier_ok:
                    gen = self.speech.begin_trigger()
                    self.clipboard.begin_trigger(gen)
                    threading.Thread(
                        target=_trigger_worker_entry,
                        args=(self.settings, self.speech, self.clipboard, gen),
                        daemon=True,
                    ).start()
                if self.suppress and modifier_ok:
                    return 1
        return user32.CallNextHookEx(self._hook_handle, n_code, w_param, l_param)

# ── GUI application ───────────────────────────────────────────────────────────


class App:
    def __init__(self) -> None:
        self.config_path = _config_path()
        self.speech = SpeechController()
        self.clipboard = ClipboardSession()
        self._hook: MouseTriggerHook | None = None
        self._hook_thread: threading.Thread | None = None
        self._capture_before: tuple[str, tuple[str, ...]] = (DEFAULT_BUTTON, ())
        self._capturing = False

        try:
            self._voices: list[dict] = _list_voices()
        except Exception:
            LOGGER.exception("Failed to list SAPI voices")
            self._voices = []

        self.root = tk.Tk()
        self.root.title("Mouse TTS")
        self.root.iconbitmap(str(_resource_path("icon.ico")))
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._quit)
        self._build_ui()

        initial: Settings | None = None
        try:
            if self.config_path.exists():
                initial = read_settings(self.config_path)
        except Exception:
            LOGGER.exception("Failed to read settings from %s", self.config_path)
        self._populate(initial)

        if initial is not None:
            self._start_hook(initial)
            self._set_status("Listening")
            if initial.launch_minimized:
                self.root.withdraw()
        else:
            self._set_status("Not configured — save settings to start")

        self._tray = pystray.Icon(
            "mouse_tts",
            _tray_icon(),
            "Mouse TTS",
            menu=pystray.Menu(
                pystray.MenuItem("Show settings", self._tray_show, default=True),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Exit", self._tray_quit),
            ),
        )
        self._tray.run_detached()

        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def run(self) -> None:
        self.root.mainloop()

    # ── layout ────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        f = ttk.Frame(self.root, padding=16)
        f.pack(fill="both", expand=True)
        r = 0

        # Trigger
        ttk.Label(f, text="Trigger", font=("Segoe UI", 10, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(0, 2)); r += 1

        ttk.Label(f, text="Button:").grid(row=r, column=0, sticky="w", padx=(16, 8), pady=2)
        self._btn_name: str = DEFAULT_BUTTON
        self._mod_names: tuple[str, ...] = ()
        self._combo_var = tk.StringVar()
        self._capture_listener: mouse.Listener | None = None
        cap_f = ttk.Frame(f)
        cap_f.grid(row=r, column=1, sticky="w", pady=2); r += 1
        ttk.Label(cap_f, textvariable=self._combo_var,
                  width=28, relief="groove", anchor="center", padding=(4, 2)).pack(side="left")
        self._btn_detect_btn = ttk.Button(cap_f, text="Detect…", width=9,
                                           command=self._start_capture)
        self._btn_detect_btn.pack(side="left", padx=(6, 0))

        self._suppress_var = tk.BooleanVar(value=DEFAULT_SUPPRESS_TRIGGER)
        ttk.Checkbutton(f, text="Block button event (prevent click from reaching other apps)",
                        variable=self._suppress_var).grid(
            row=r, column=0, columnspan=2, sticky="w", padx=(16, 0), pady=2); r += 1

        ttk.Separator(f, orient="horizontal").grid(
            row=r, column=0, columnspan=2, sticky="ew", pady=8); r += 1

        # Voice
        ttk.Label(f, text="Voice", font=("Segoe UI", 10, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(0, 2)); r += 1

        voice_names = [v["name"] for v in self._voices] if self._voices else ["(no voices found)"]
        ttk.Label(f, text="Voice:").grid(row=r, column=0, sticky="w", padx=(16, 8), pady=2)
        self._voice_var = tk.StringVar()
        self._voice_cb = ttk.Combobox(f, textvariable=self._voice_var,
                                      values=voice_names, state="readonly", width=36)
        self._voice_cb.grid(row=r, column=1, sticky="w", pady=2); r += 1
        if not self._voices:
            self._voice_cb.configure(state="disabled")

        ttk.Label(f, text="Rate:").grid(row=r, column=0, sticky="w", padx=(16, 8), pady=2)
        rf = ttk.Frame(f)
        rf.grid(row=r, column=1, sticky="w", pady=2); r += 1
        self._rate_var = tk.IntVar(value=DEFAULT_RATE)
        self._rate_lbl = ttk.Label(rf, text=f"{DEFAULT_RATE:+d}", width=4)
        ttk.Scale(rf, from_=-10, to=10, orient="horizontal", length=180, variable=self._rate_var,
                  command=lambda v: self._rate_lbl.config(text=f"{round(float(v)):+d}")).pack(side="left")
        self._rate_lbl.pack(side="left", padx=(4, 0))

        ttk.Label(f, text="Pitch:").grid(row=r, column=0, sticky="w", padx=(16, 8), pady=2)
        pf = ttk.Frame(f)
        pf.grid(row=r, column=1, sticky="w", pady=2); r += 1
        self._pitch_var = tk.IntVar(value=DEFAULT_PITCH)
        self._pitch_lbl = ttk.Label(pf, text=f"{DEFAULT_PITCH:+d}", width=4)
        ttk.Scale(pf, from_=-20, to=20, orient="horizontal", length=180, variable=self._pitch_var,
                  command=lambda v: self._pitch_lbl.config(text=f"{round(float(v)):+d}")).pack(side="left")
        self._pitch_lbl.pack(side="left", padx=(4, 0))

        ttk.Separator(f, orient="horizontal").grid(
            row=r, column=0, columnspan=2, sticky="ew", pady=8); r += 1

        # Exclude patterns
        ttk.Label(f, text="Exclude patterns", font=("Segoe UI", 10, "bold")).grid(
            row=r, column=0, columnspan=2, sticky="w", pady=(0, 2)); r += 1
        ttk.Label(f, text="One regex per line:", foreground="gray").grid(
            row=r, column=0, columnspan=2, sticky="w", padx=(16, 0)); r += 1
        self._patterns = scrolledtext.ScrolledText(f, width=46, height=4, font=("Consolas", 9))
        self._patterns.grid(row=r, column=0, columnspan=2, sticky="ew", padx=(16, 0), pady=(2, 0)); r += 1

        ttk.Separator(f, orient="horizontal").grid(
            row=r, column=0, columnspan=2, sticky="ew", pady=8); r += 1

        # Launch options
        self._launch_min_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Launch minimized to tray",
                        variable=self._launch_min_var).grid(
            row=r, column=0, columnspan=2, sticky="w", padx=(16, 0)); r += 1

        self._autostart_var = tk.BooleanVar(value=False)
        self._autostart_cb = ttk.Checkbutton(f, text="Start with Windows",
                                              variable=self._autostart_var,
                                              command=self._on_autostart_toggle)
        if not getattr(sys, "frozen", False):
            self._autostart_cb.config(state="disabled")
        self._autostart_cb.grid(row=r, column=0, columnspan=2, sticky="w", padx=(16, 0)); r += 1

        ttk.Separator(f, orient="horizontal").grid(
            row=r, column=0, columnspan=2, sticky="ew", pady=8); r += 1

        # Status + buttons
        self._status_var = tk.StringVar(value="—")
        ttk.Label(f, textvariable=self._status_var, foreground="gray").grid(
            row=r, column=0, columnspan=2, sticky="w"); r += 1

        bf = ttk.Frame(f)
        bf.grid(row=r, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(bf, text="Minimize to tray", command=self._hide).pack(side="right", padx=(8, 0))
        ttk.Button(bf, text="Save & Apply", command=self._save).pack(side="right")

    def _populate(self, s: Settings | None) -> None:
        self._btn_name = s.button_name if s else DEFAULT_BUTTON
        self._mod_names = s.modifiers if s else ()
        self._combo_var.set(self._combo_label())
        self._suppress_var.set(s.suppress if s else DEFAULT_SUPPRESS_TRIGGER)
        rate = s.rate if s else DEFAULT_RATE
        pitch = s.pitch if s else DEFAULT_PITCH
        self._rate_var.set(rate)
        self._pitch_var.set(pitch)
        self._rate_lbl.config(text=f"{rate:+d}")
        self._pitch_lbl.config(text=f"{pitch:+d}")
        self._patterns.delete("1.0", "end")
        patterns = s.exclude_pattern_strings if s else DEFAULT_EXCLUDE_PATTERN_STRINGS
        self._patterns.insert("1.0", "\n".join(patterns))
        self._launch_min_var.set(s.launch_minimized if s else False)
        self._autostart_var.set(_get_autostart())
        if s and s.voice_name:
            self._voice_var.set(s.voice_name)
        elif self._voices:
            self._voice_var.set(self._voices[0]["name"])

    # ── actions ───────────────────────────────────────────────────────────────

    def _save(self) -> None:
        if not self._voices:
            messagebox.showerror("No voices", "No SAPI voices found.", parent=self.root)
            return
        try:
            settings = self._read_form()
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc), parent=self.root)
            return
        try:
            write_settings(self.config_path, settings)
        except OSError as exc:
            messagebox.showerror("Save failed", str(exc), parent=self.root)
            return
        self._start_hook(settings)
        self._set_status("Listening")

    def _read_form(self) -> Settings:
        btn = self._btn_name
        mods = self._mod_names
        voice_name = self._voice_var.get()
        rate = int(self._rate_var.get())
        pitch = int(self._pitch_var.get())
        raw = self._patterns.get("1.0", "end").strip()

        _validate_button(btn)
        _validate_modifiers(mods)
        _validate_rate(rate)
        _validate_pitch(pitch)

        suppress = self._suppress_var.get()
        if suppress and not mods and btn in ("left", "right"):
            raise ValueError(
                f"Blocking the {btn} mouse button without a modifier key will make it "
                f"unusable system-wide.\n\n"
                f"Either add a modifier key (e.g. Ctrl + {btn}) or uncheck "
                f"'Block button event'."
            )

        voice_id = next((v["id"] for v in self._voices if v["name"] == voice_name), "")
        if not voice_id and not voice_name:
            raise ValueError("Please select a voice.")

        pat_strings = _parse_pattern_strings(raw)
        pat_compiled = _compile_patterns(pat_strings)

        return Settings(
            button_name=btn, modifiers=mods,
            suppress=suppress,
            voice_id=voice_id, voice_name=voice_name,
            rate=rate, pitch=pitch,
            launch_minimized=self._launch_min_var.get(),
            exclude_pattern_strings=pat_strings,
            exclude_patterns=pat_compiled,
        )

    def _on_autostart_toggle(self) -> None:
        try:
            _set_autostart(self._autostart_var.get())
        except OSError as exc:
            messagebox.showwarning("Autostart", f"Could not update autostart: {exc}", parent=self.root)
            self._autostart_var.set(_get_autostart())

    def _combo_label(self) -> str:
        parts = [m.capitalize() for m in self._mod_names]
        parts.append(self._btn_name)
        return " + ".join(parts)

    def _start_capture(self) -> None:
        self._capture_before = (self._btn_name, self._mod_names)
        self._capturing = True
        self._combo_var.set("hold keys + click…")
        self._btn_detect_btn.config(text="Cancel", command=self._cancel_capture)
        # Delay so the click that triggered this button isn't immediately captured
        self.root.after(250, self._begin_capture_listen)

    def _poll_capture_keys(self) -> None:
        if not self._capturing:
            return
        mods = _detect_modifiers()
        if mods:
            self._combo_var.set(" + ".join(m.capitalize() for m in mods) + " + …")
        else:
            self._combo_var.set("hold keys + click…")
        self.root.after(50, self._poll_capture_keys)

    def _begin_capture_listen(self) -> None:
        if not self._capturing:
            return  # was cancelled during the delay

        self._poll_capture_keys()

        def on_click(x, y, button, pressed):
            if not pressed:
                return
            name = {
                mouse.Button.left: "left",
                mouse.Button.right: "right",
                mouse.Button.middle: "middle",
                mouse.Button.x1: "x1",
                mouse.Button.x2: "x2",
            }.get(button)
            if name:
                mods = _detect_modifiers()
                self.root.after(0, lambda: self._on_combo_captured(name, mods))
                return False  # stop listener

        self._capture_listener = mouse.Listener(on_click=on_click)
        self._capture_listener.start()

    def _on_combo_captured(self, btn_name: str, mod_names: tuple[str, ...]) -> None:
        self._capturing = False
        self._btn_name = btn_name
        self._mod_names = mod_names
        self._combo_var.set(self._combo_label())
        self._btn_detect_btn.config(text="Detect…", command=self._start_capture)
        self._capture_listener = None

    def _cancel_capture(self) -> None:
        self._capturing = False
        if self._capture_listener is not None:
            self._capture_listener.stop()
            self._capture_listener = None
        self._btn_name, self._mod_names = self._capture_before
        self._combo_var.set(self._combo_label())
        self._btn_detect_btn.config(text="Detect…", command=self._start_capture)

    def _hide(self) -> None:
        self.root.withdraw()

    def _show(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_show(self, icon=None, item=None) -> None:
        self.root.after(0, self._show)

    def _tray_quit(self, icon=None, item=None) -> None:
        self.root.after(0, self._quit)

    def _quit(self) -> None:
        if self._capture_listener is not None:
            self._capture_listener.stop()
            self._capture_listener = None
        self._stop_hook()
        try:
            self._tray.stop()
        except Exception:
            LOGGER.exception("Failed to stop tray icon")
        self.root.destroy()

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self._status_var.set(text))

    # ── hook lifecycle ────────────────────────────────────────────────────────

    def _start_hook(self, settings: Settings) -> None:
        self._stop_hook()
        self._hook = MouseTriggerHook(settings, self.speech, self.clipboard, settings.suppress)
        self._hook_thread = threading.Thread(target=self._hook.run, daemon=True)
        self._hook_thread.start()

    def _stop_hook(self) -> None:
        if self._hook is not None:
            self._hook.stop()
            if self._hook_thread is not None:
                self._hook_thread.join(timeout=2.0)
        self._hook = None
        self._hook_thread = None

# ── helper functions ──────────────────────────────────────────────────────────


def _resource_path(name: str) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / name  # type: ignore[attr-defined]
    return Path(__file__).parent / name


def _work_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path.cwd()


def _log_path() -> Path:
    return _work_dir() / LOG_FILENAME


def _configure_logging() -> None:
    global LOGGING_CONFIGURED
    if LOGGING_CONFIGURED:
        return
    log_path = _log_path()
    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s")
    )
    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(logging.INFO)
    root.addHandler(handler)
    root.addHandler(logging.StreamHandler(sys.stderr))
    LOGGING_CONFIGURED = True
    _install_exception_hooks()
    LOGGER.info("Logging to %s", log_path)


def _install_exception_hooks() -> None:
    def _sys_excepthook(exc_type, exc_value, exc_traceback) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        LOGGER.critical(
            "Unhandled exception",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    def _thread_excepthook(args) -> None:
        if issubclass(args.exc_type, KeyboardInterrupt):
            return
        LOGGER.critical(
            "Unhandled exception in thread %s",
            args.thread.name if args.thread is not None else "<unknown>",
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )

    sys.excepthook = _sys_excepthook
    if hasattr(threading, "excepthook"):
        threading.excepthook = _thread_excepthook  # type: ignore[assignment]


def _config_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / CONFIG_FILENAME
    return Path(CONFIG_FILENAME)


def _tray_icon() -> Image.Image:
    return Image.open(_resource_path("icon.ico"))


@contextmanager
def _com_initialized():
    import pythoncom

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def _list_voices() -> list[dict]:
    with _com_initialized():
        import win32com.client

        spv = win32com.client.Dispatch("SAPI.SpVoice")
        voices = spv.GetVoices()
        return [
            {"id": str(voices.Item(i).Id), "name": str(voices.Item(i).GetDescription())}
            for i in range(voices.Count)
        ]


def find_voice_token(speaker, settings: Settings):
    voices = speaker.GetVoices()
    for i in range(voices.Count):
        t = voices.Item(i)
        if settings.voice_id and str(t.Id) == settings.voice_id:
            return t
    for i in range(voices.Count):
        t = voices.Item(i)
        if settings.voice_name and str(t.GetDescription()) == settings.voice_name:
            return t
    return None


def _trigger_worker(
    settings: Settings,
    speech: SpeechController,
    clipboard: ClipboardSession,
    generation: int,
) -> None:
    try:
        text = _copy_selected(timeout=1.5)
        if not speech.is_current(generation) or not text or not text.strip():
            return
        text = normalize_for_speech(text)
        text = filter_for_speech(text, settings.exclude_patterns)
        if text:
            speech.speak(settings, text, generation)
    except Exception:
        LOGGER.exception("Trigger worker failed")
        raise
    finally:
        clipboard.finish_trigger(generation)


def _trigger_worker_entry(
    settings: Settings,
    speech: SpeechController,
    clipboard: ClipboardSession,
    generation: int,
) -> None:
    with _com_initialized():
        _trigger_worker(settings, speech, clipboard, generation)


def _copy_selected(timeout: float = 1.5) -> str:
    before = int(user32.GetClipboardSequenceNumber())
    _ctrl_c()
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        if int(user32.GetClipboardSequenceNumber()) != before:
            time.sleep(0.05)
            return safe_paste()
        time.sleep(0.05)
    return ""


def _ctrl_c() -> None:
    ctrl = keyboard.Controller()
    c = keyboard.KeyCode.from_vk(0x43)
    # Release Shift/Alt first — if they're physically held as trigger modifiers,
    # the target app would otherwise see Ctrl+Shift+C or Ctrl+Alt+C instead of Ctrl+C.
    for key in (keyboard.Key.shift, keyboard.Key.alt, keyboard.Key.alt_gr):
        ctrl.release(key)
    with ctrl.pressed(keyboard.Key.ctrl):
        ctrl.press(c)
        ctrl.release(c)
    time.sleep(0.1)


def safe_paste() -> str:
    try:
        return pyperclip.paste()
    except pyperclip.PyperclipException:
        LOGGER.exception("Clipboard paste failed")
        return ""


def safe_copy(text: str) -> None:
    try:
        pyperclip.copy(text)
    except pyperclip.PyperclipException:
        LOGGER.exception("Clipboard copy failed")


def _btn_for_msg(message: int, mouse_data: int) -> mouse.Button | None:
    if message in (WM_LBUTTONDOWN, WM_LBUTTONUP):
        return mouse.Button.left
    if message in (WM_RBUTTONDOWN, WM_RBUTTONUP):
        return mouse.Button.right
    if message in (WM_MBUTTONDOWN, WM_MBUTTONUP):
        return mouse.Button.middle
    if message in (WM_XBUTTONDOWN, WM_XBUTTONUP):
        hiword = (mouse_data >> 16) & 0xFFFF
        if hiword == XBUTTON1:
            return mouse.Button.x1
        if hiword == XBUTTON2:
            return mouse.Button.x2
    return None


def _is_down(message: int) -> bool:
    return message in (WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN)


def _modifier_active(modifiers: tuple[str, ...]) -> bool:
    return all(
        any(bool(user32.GetAsyncKeyState(vk) & 0x8000) for vk in MODIFIER_VK_CODES[m])
        for m in modifiers
    )


def _detect_modifiers() -> tuple[str, ...]:
    return tuple(
        m for m, vk_codes in MODIFIER_VK_CODES.items()
        if any(bool(user32.GetAsyncKeyState(vk) & 0x8000) for vk in vk_codes)
    )


def _get_autostart() -> bool:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, AUTOSTART_APP_NAME)
        winreg.CloseKey(key)
        return True
    except OSError:
        LOGGER.exception("Failed to read autostart registry state")
        return False


def _fix_autostart_path() -> None:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0,
            winreg.KEY_QUERY_VALUE | winreg.KEY_SET_VALUE,
        )
        try:
            value, _ = winreg.QueryValueEx(key, AUTOSTART_APP_NAME)
            stored_path = value.strip().strip('"')
            if stored_path != sys.executable:
                winreg.SetValueEx(key, AUTOSTART_APP_NAME, 0, winreg.REG_SZ, f'"{sys.executable}"')
        except OSError:
            LOGGER.exception("Failed to repair autostart registry value")
        finally:
            winreg.CloseKey(key)
    except OSError:
        LOGGER.exception("Failed to open autostart registry key")


def _set_autostart(enable: bool) -> None:
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_KEY, 0, winreg.KEY_SET_VALUE)
    try:
        if enable:
            winreg.SetValueEx(key, AUTOSTART_APP_NAME, 0, winreg.REG_SZ, f'"{sys.executable}"')
        else:
            try:
                winreg.DeleteValue(key, AUTOSTART_APP_NAME)
            except OSError:
                LOGGER.exception("Failed to remove autostart registry value")
    finally:
        winreg.CloseKey(key)


def read_settings(config_path: Path) -> Settings:
    p = configparser.ConfigParser()
    p.read(config_path, encoding="utf-8")
    if not p.has_section("mouse") or not p.has_section("tts"):
        raise ValueError("Config missing required sections.")
    btn = p.get("mouse", "button", fallback=DEFAULT_BUTTON).strip().lower()
    mod_str = p.get("mouse", "modifier", fallback="").strip().lower()
    modifiers = _parse_modifiers(mod_str)
    suppress = p.getboolean("mouse", "suppress", fallback=DEFAULT_SUPPRESS_TRIGGER)
    voice_id = p.get("tts", "voice_id", fallback="").strip()
    voice_name = p.get("tts", "voice_name", fallback="").strip()
    rate = p.getint("tts", "rate", fallback=DEFAULT_RATE)
    pitch = p.getint("tts", "pitch", fallback=DEFAULT_PITCH)
    launch_minimized = p.getboolean("app", "launch_minimized", fallback=False)
    raw_pats = p.get("tts", "exclude_patterns", fallback="")
    _validate_button(btn)
    _validate_modifiers(modifiers)
    _validate_rate(rate)
    _validate_pitch(pitch)
    if not voice_id and not voice_name:
        raise ValueError("No TTS voice configured.")
    pat_strings = _parse_pattern_strings(raw_pats)
    return Settings(
        button_name=btn, modifiers=modifiers,
        suppress=suppress,
        voice_id=voice_id, voice_name=voice_name,
        rate=rate, pitch=pitch,
        launch_minimized=launch_minimized,
        exclude_pattern_strings=pat_strings,
        exclude_patterns=_compile_patterns(pat_strings),
    )


def write_settings(config_path: Path, settings: Settings) -> None:
    p = configparser.ConfigParser()
    p["mouse"] = {"button": settings.button_name, "modifier": "+".join(settings.modifiers), "suppress": str(settings.suppress).lower()}
    p["tts"] = {
        "voice_id": settings.voice_id,
        "voice_name": settings.voice_name,
        "rate": str(settings.rate),
        "pitch": str(settings.pitch),
        "exclude_patterns": "\n".join(settings.exclude_pattern_strings),
    }
    p["app"] = {
        "launch_minimized": str(settings.launch_minimized).lower(),
    }
    with config_path.open("w", encoding="utf-8") as fh:
        p.write(fh)


def _parse_pattern_strings(raw: str) -> tuple[str, ...]:
    return tuple(
        line.strip() for line in raw.splitlines()
        if line.strip() and not line.strip().startswith(("#", ";"))
    )


def _compile_patterns(patterns: tuple[str, ...]) -> tuple[re.Pattern, ...]:
    result = []
    for pat in patterns:
        try:
            result.append(re.compile(pat))
        except re.error as exc:
            raise ValueError(f"Invalid regex '{pat}': {exc}") from exc
    return tuple(result)


def filter_for_speech(text: str, patterns: tuple[re.Pattern, ...]) -> str:
    for p in patterns:
        text = p.sub("", text)
    return normalize_for_speech(text)


def normalize_for_speech(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\t\f\v]+", " ", text)
    lines = [re.sub(r" {2,}", " ", line).strip() for line in text.split("\n")]
    return "\n".join(l for l in lines if l).strip()


def _validate_button(name: str) -> None:
    if name not in BUTTON_ALIASES:
        raise ValueError(f"Unsupported mouse button '{name}'.")


def _validate_modifiers(modifiers: tuple[str, ...]) -> None:
    for name in modifiers:
        if name not in MODIFIER_VK_CODES:
            raise ValueError(f"Unsupported modifier '{name}'.")


def _parse_modifiers(s: str) -> tuple[str, ...]:
    if not s or s == "none":
        return ()
    return tuple(m for m in (part.strip() for part in s.split("+")) if m)


def _validate_rate(rate: int) -> None:
    if not -10 <= rate <= 10:
        raise ValueError("Rate must be between -10 and 10.")


def _validate_pitch(pitch: int) -> None:
    if not -20 <= pitch <= 20:
        raise ValueError("Pitch must be between -20 and 20.")


def main() -> int:
    _mutex = kernel32.CreateMutexW(None, True, SINGLE_INSTANCE_MUTEX)
    if ctypes.get_last_error() == ERROR_ALREADY_EXISTS:
        return 0
    _configure_logging()
    if getattr(sys, "frozen", False):
        _fix_autostart_path()
    App().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

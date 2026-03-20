import threading
import ctypes

import pyautogui

TRANSPARENT_KEY = "#ff00ff"
MOVE_DOWN_PIXELS = 65
ICON_GAP_PIXELS = 10

try:
    import tkinter as tk
except Exception:
    tk = None


class VoiceStateIndicator:
    def __init__(self):
        self._enabled = tk is not None
        self._root = None
        self._thread = None
        self._ready = threading.Event()
        self._canvas = None
        self._emoji_item = None
        self._hide_after_id = None

    def _get_work_area(self):
        if tk is None:
            return 0, 0, 0, 0

        try:
            if hasattr(ctypes, "windll") and hasattr(ctypes.windll, "user32"):
                class RECT(ctypes.Structure):
                    _fields_ = [
                        ("left", ctypes.c_long),
                        ("top", ctypes.c_long),
                        ("right", ctypes.c_long),
                        ("bottom", ctypes.c_long),
                    ]

                rect = RECT()
                SPI_GETWORKAREA = 0x0030
                ok = ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
                if ok:
                    return rect.left, rect.top, rect.right, rect.bottom
        except Exception:
            pass

        screen_w, screen_h = pyautogui.size()
        return 0, 0, screen_w, screen_h

    def _compute_geometry(self):
        left, top, right, bottom = self._get_work_area()
        work_w = max(1, right - left)
        work_h = max(1, bottom - top)

        icon = int(max(56, min(84, round(min(work_w, work_h) * 0.035))))
        margin = max(8, int(round(icon * 0.18)))

        x = left + margin + icon + ICON_GAP_PIXELS
        y = bottom - icon - margin + MOVE_DOWN_PIXELS
        return icon, x, y

    def _cancel_auto_hide(self):
        if self._root is None:
            return
        if self._hide_after_id is not None:
            try:
                self._root.after_cancel(self._hide_after_id)
            except Exception:
                pass
            self._hide_after_id = None

    def _set_symbol(self, symbol, auto_hide_ms=None):
        if not self._enabled or self._root is None or self._canvas is None or self._emoji_item is None:
            return

        def _apply():
            self._cancel_auto_hide()
            self._canvas.itemconfigure(self._emoji_item, text=symbol)
            self._root.deiconify()
            if auto_hide_ms is not None and auto_hide_ms > 0:
                self._hide_after_id = self._root.after(auto_hide_ms, self._root.withdraw)

        try:
            self._root.after(0, _apply)
        except Exception:
            pass

    def _run_ui(self):
        try:
            self._root = tk.Tk()
            self._root.overrideredirect(True)
            self._root.attributes("-topmost", True)
            self._root.configure(bg=TRANSPARENT_KEY)
            try:
                self._root.wm_attributes("-transparentcolor", TRANSPARENT_KEY)
            except Exception:
                pass

            icon_size, x, y = self._compute_geometry()
            card = tk.Frame(
                self._root,
                bg=TRANSPARENT_KEY,
                bd=0,
                relief="flat",
                highlightthickness=0,
                highlightbackground=TRANSPARENT_KEY,
                width=icon_size,
                height=icon_size,
            )
            card.pack(fill="both", expand=True)
            card.pack_propagate(False)

            self._canvas = tk.Canvas(card, width=icon_size, height=icon_size, bg=TRANSPARENT_KEY, highlightthickness=0)
            self._canvas.place(x=0, y=0)

            label_font = max(14, int(round(icon_size * 0.34)))
            self._emoji_item = self._canvas.create_text(
                int(icon_size * 0.50),
                int(icon_size * 0.50),
                text="🫧",
                fill="#ffffff",
                font=("Segoe UI Emoji", label_font, "bold"),
            )

            self._root.geometry(f"{icon_size}x{icon_size}+{x}+{y}")
            self._root.withdraw()
            self._ready.set()
            self._root.mainloop()
        except Exception as e:
            print(f"Voice indicator error: {e}")
            self._enabled = False
            self._ready.set()

    def _ensure_started(self):
        if not self._enabled:
            return False
        if self._thread is None:
            self._thread = threading.Thread(target=self._run_ui, daemon=True)
            self._thread.start()
        self._ready.wait(timeout=1.0)
        return self._root is not None

    def show_listening(self):
        if not self._ensure_started():
            return
        self._set_symbol("🫧")

    def show_recognized(self, auto_hide_ms=900):
        if not self._ensure_started():
            return
        self._set_symbol("⚙️", auto_hide_ms=auto_hide_ms)

    def hide(self):
        if not self._enabled or self._root is None:
            return

        def _apply():
            self._cancel_auto_hide()
            self._root.withdraw()

        try:
            self._root.after(0, _apply)
        except Exception:
            pass

    def close(self):
        if not self._enabled or self._root is None:
            return

        def _apply():
            self._cancel_auto_hide()
            self._root.destroy()

        try:
            self._root.after(0, _apply)
        except Exception:
            pass

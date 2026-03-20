import threading
import ctypes

import pyautogui

TRANSPARENT_KEY = "#ff00ff"
MOVE_DOWN_PIXELS = 65

try:
    import tkinter as tk
except Exception:
    tk = None


class DesktopAuthIndicator:
    def __init__(self):
        self._enabled = tk is not None
        self._root = None
        self._thread = None
        self._ready = threading.Event()

    def minimize_all_windows(self):
        # Windows-only: trigger Show Desktop (Win + D) once auth succeeds.
        try:
            if not hasattr(ctypes, "windll") or not hasattr(ctypes.windll, "user32"):
                return
            user32 = ctypes.windll.user32
            VK_LWIN = 0x5B
            VK_D = 0x44
            KEYEVENTF_KEYUP = 0x0002
            user32.keybd_event(VK_LWIN, 0, 0, 0)
            user32.keybd_event(VK_D, 0, 0, 0)
            user32.keybd_event(VK_D, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass

    def _get_work_area(self):
        # Return desktop work area (excluding taskbar) on Windows.
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

        # Fallback to full screen if work area query is unavailable.
        screen_w, screen_h = pyautogui.size()
        return 0, 0, screen_w, screen_h

    def _compute_geometry(self):
        left, top, right, bottom = self._get_work_area()
        work_w = max(1, right - left)
        work_h = max(1, bottom - top)

        # Scale slightly with monitor size while keeping a compact icon look.
        icon = int(max(56, min(84, round(min(work_w, work_h) * 0.035))))
        margin = max(8, int(round(icon * 0.18)))

        # Place near bottom-left of usable area, similar to taskbar app icon region.
        x = left + margin
        y = bottom - icon - margin + MOVE_DOWN_PIXELS
        return icon, x, y

    def _run_ui(self):
        try:
            self._root = tk.Tk()
            self._root.overrideredirect(True)
            self._root.attributes("-topmost", True)
            self._root.configure(bg=TRANSPARENT_KEY)
            try:
                # Windows: make key color fully transparent to remove black background card.
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

            canvas = tk.Canvas(card, width=icon_size, height=icon_size, bg=TRANSPARENT_KEY, highlightthickness=0)
            canvas.place(x=0, y=0)

            label_font = max(14, int(round(icon_size * 0.34)))
            canvas.create_text(
                int(icon_size * 0.50),
                int(icon_size * 0.50),
                text="👌",
                fill="#ffffff",
                font=("Segoe UI Emoji", label_font, "bold"),
            )

            self._root.geometry(f"{icon_size}x{icon_size}+{x}+{y}")
            self._root.withdraw()
            self._ready.set()
            self._root.mainloop()
        except Exception as e:
            print(f"Desktop indicator error: {e}")
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

    def show(self):
        if not self._ensure_started():
            return
        try:
            self._root.after(0, self._root.deiconify)
        except Exception:
            pass

    def hide(self):
        if not self._enabled or self._root is None:
            return
        try:
            self._root.after(0, self._root.withdraw)
        except Exception:
            pass

    def close(self):
        if not self._enabled or self._root is None:
            return
        try:
            self._root.after(0, self._root.destroy)
        except Exception:
            pass

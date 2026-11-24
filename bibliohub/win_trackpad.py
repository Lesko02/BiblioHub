"""Trackpad scrolling on Windows.

Tk turns a mouse wheel into <MouseWheel>. A precision trackpad does not:
two-finger drag is WM_POINTERWHEEL, which Tk never delivers. A thread hook
sees that message (and the ordinary wheel message) for every window of this
process and reports the notch delta.
"""

import ctypes
import sys

if sys.platform != "win32":
    def install(callback):
        return False

    def uninstall():
        return None
else:
    WH_GETMESSAGE = 3
    WH_CALLWNDPROC = 4
    PM_REMOVE = 0x0001
    WM_MOUSEWHEEL = 0x020A
    WM_POINTERWHEEL = 0x024E
    WM_GESTURE = 0x0119
    GID_PAN = 4
    GF_BEGIN = 0x00000001

    class _POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    class _MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd", ctypes.c_void_p),
            ("message", ctypes.c_uint),
            ("wParam", ctypes.c_size_t),
            ("lParam", ctypes.c_size_t),
            ("time", ctypes.c_ulong),
            ("pt", _POINT),
        ]

    class _CWPSTRUCT(ctypes.Structure):
        _fields_ = [
            ("lParam", ctypes.c_size_t),
            ("wParam", ctypes.c_size_t),
            ("message", ctypes.c_uint),
            ("hwnd", ctypes.c_void_p),
        ]

    class _GESTUREINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_uint),
            ("dwFlags", ctypes.c_uint),
            ("dwID", ctypes.c_uint),
            ("hwndTarget", ctypes.c_void_p),
            ("ptsLocation", ctypes.c_long),
            ("dwInstanceID", ctypes.c_uint),
            ("dwSequenceID", ctypes.c_uint),
            ("ullArguments", ctypes.c_ulonglong),
            ("cbExtraArgs", ctypes.c_uint),
        ]

    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    _HOOKPROC = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t, ctypes.c_int, ctypes.c_size_t, ctypes.c_void_p,
    )
    _user32.SetWindowsHookExW.argtypes = [
        ctypes.c_int, _HOOKPROC, ctypes.c_void_p, ctypes.c_ulong,
    ]
    _user32.SetWindowsHookExW.restype = ctypes.c_void_p
    _user32.CallNextHookEx.argtypes = [
        ctypes.c_void_p, ctypes.c_int, ctypes.c_size_t, ctypes.c_void_p,
    ]
    _user32.CallNextHookEx.restype = ctypes.c_ssize_t
    _user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
    _user32.UnhookWindowsHookEx.restype = ctypes.c_int
    _user32.GetGestureInfo.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    _user32.GetGestureInfo.restype = ctypes.c_int
    _user32.CloseGestureInfoHandle.argtypes = [ctypes.c_void_p]
    _user32.CloseGestureInfoHandle.restype = ctypes.c_int
    _kernel32.GetCurrentThreadId.restype = ctypes.c_ulong

    _state = {"hooks": [], "procs": [], "callback": None, "pan_y": None}

    def _wheel_delta(wparam: int) -> int:
        return ctypes.c_short((int(wparam) >> 16) & 0xFFFF).value

    def _note(message, wparam):
        log = _state.get("log")
        if log is not None:
            log.append(message)
        if message in (WM_MOUSEWHEEL, WM_POINTERWHEEL):
            callback = _state["callback"]
            if callback is not None:
                callback(_wheel_delta(wparam))

    def _get_proc(nCode, wParam, lParam):
        # Posted wheel messages, including a normal mouse wheel.
        if nCode >= 0 and int(wParam) == PM_REMOVE and lParam:
            msg = ctypes.cast(lParam, ctypes.POINTER(_MSG)).contents
            _note(msg.message, msg.wParam)
        return _user32.CallNextHookEx(None, nCode, wParam, lParam)

    def _pan(hgesture):
        info = _GESTUREINFO()
        info.cbSize = ctypes.sizeof(_GESTUREINFO)
        if not _user32.GetGestureInfo(hgesture, ctypes.byref(info)):
            return
        try:
            if info.dwID != GID_PAN:
                return
            y = info.ptsLocation & 0xFFFF
            if y >= 0x8000:
                y -= 0x10000
            if info.dwFlags & GF_BEGIN or _state["pan_y"] is None:
                _state["pan_y"] = y
                return
            dy = y - _state["pan_y"]
            _state["pan_y"] = y
            callback = _state["callback"]
            if dy and callback is not None:
                # Finger pixels, scaled so a short swipe moves about as much as a wheel notch.
                callback(int(dy * 2.5))
        finally:
            _user32.CloseGestureInfoHandle(hgesture)

    def _call_proc(nCode, wParam, lParam):
        # Sent messages. A precision trackpad scroll is WM_POINTERWHEEL or a pan
        # gesture, and those are sent, so the get-message hook never sees them.
        if nCode >= 0 and lParam:
            info = ctypes.cast(lParam, ctypes.POINTER(_CWPSTRUCT)).contents
            log = _state.get("log")
            if log is not None:
                log.append(info.message)
            if info.message == WM_POINTERWHEEL:
                _note(info.message, info.wParam)
            elif info.message == WM_GESTURE and info.lParam:
                _pan(info.lParam)
        return _user32.CallNextHookEx(None, nCode, wParam, lParam)

    def install(callback) -> bool:
        """Report each wheel/trackpad notch to callback(delta). delta is a Windows wheel step (120 per notch)."""
        _state["callback"] = callback
        if _state["hooks"]:
            return True
        thread = _kernel32.GetCurrentThreadId()
        ok = False
        for kind, function in ((WH_GETMESSAGE, _get_proc), (WH_CALLWNDPROC, _call_proc)):
            proc = _HOOKPROC(function)
            hook = _user32.SetWindowsHookExW(kind, proc, None, thread)
            if hook:
                _state["hooks"].append(hook)
                _state["procs"].append(proc)
                ok = True
        return ok

    def uninstall() -> None:
        hooks = list(_state["hooks"])
        _state["hooks"].clear()
        _state["procs"].clear()
        _state["callback"] = None
        for hook in hooks:
            _user32.UnhookWindowsHookEx(hook)

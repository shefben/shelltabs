#!/usr/bin/env python3
# explorer_gui_dump_crosswin_compat.py
# Cross-Windows Explorer GUI dumper with Python 3.8 compatibility for Vista/7,
# and full features on 8/8.1/10/11 with Python 3.9+.
#
# Usage:
#   python explorer_gui_dump_crosswin_compat.py [--uia] [--resolve] [--outdir DIR]
#
# Notes:
#   - Run elevated for best results.
#   - All files are written with latin-1 encoding.
#   - Timestamp format everywhere: "%m/%d/%Y %H:%M:%S"

import os
import sys
import argparse
import ctypes
import ctypes.wintypes as wintypes
import datetime
import json
import traceback

# Optional packages (graceful if missing)
try:
    import psutil
except Exception:
    psutil = None

try:
    import pefile
except Exception:
    pefile = None

# Optional for UI Automation
try:
    import comtypes
    import comtypes.client as cc
    from comtypes import GUID
except Exception:
    comtypes = None

# ---------- Compatibility gates ----------
PY38 = (sys.version_info.major, sys.version_info.minor) <= (3, 8)
winver = sys.getwindowsversion()  # (major, minor, build, platform, service_pack)
WIN_MAJOR, WIN_MINOR = winver.major, winver.minor
# Vista = 6.0, Win7 = 6.1, Win8 = 6.2, 8.1 = 6.3, Win10/11 = 10.0+
IS_VISTA_OR_7 = (WIN_MAJOR == 6 and WIN_MINOR in (0, 1))
LEGACY_MODE = PY38 or IS_VISTA_OR_7  # keep later-version code path untouched; just degrade politely when needed

TS_FMT = "%m/%d/%Y %H:%M:%S"
ENCODING = "latin-1"  # per your constraint

def now_ts():
    return datetime.datetime.now().strftime(TS_FMT)

# ---------- WinAPI ----------
user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
psapi = ctypes.WinDLL("psapi", use_last_error=True)
advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
try:
    oleacc = ctypes.WinDLL("oleacc")
except Exception:
    oleacc = None

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowThreadProcessId.restype = wintypes.DWORD
GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))

EnumWindows = user32.EnumWindows
EnumWindows.argtypes = (WNDENUMPROC, wintypes.LPARAM)
EnumWindows.restype = wintypes.BOOL

EnumChildWindows = user32.EnumChildWindows
EnumChildWindows.argtypes = (wintypes.HWND, WNDENUMPROC, wintypes.LPARAM)
EnumChildWindows.restype = wintypes.BOOL

GetClassNameW = user32.GetClassNameW
GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
GetClassNameW.restype = ctypes.c_int

GetWindowTextLengthW = user32.GetWindowTextLengthW
GetWindowTextLengthW.argtypes = (wintypes.HWND,)
GetWindowTextLengthW.restype = ctypes.c_int

GetWindowTextW = user32.GetWindowTextW
GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
GetWindowTextW.restype = ctypes.c_int

GetWindowLongPtrW = user32.GetWindowLongPtrW
GetWindowLongPtrW.argtypes = (wintypes.HWND, ctypes.c_int)
GetWindowLongPtrW.restype = ctypes.c_void_p

GetClassLongPtrW = user32.GetClassLongPtrW
GetClassLongPtrW.argtypes = (wintypes.HWND, ctypes.c_int)
GetClassLongPtrW.restype = ctypes.c_void_p

GetMenu = user32.GetMenu
GetMenu.argtypes = (wintypes.HWND,)
GetMenu.restype = wintypes.HMENU

GetMenuItemCount = user32.GetMenuItemCount
GetMenuItemCount.argtypes = (wintypes.HMENU,)
GetMenuItemCount.restype = ctypes.c_int

GetMenuStringW = user32.GetMenuStringW
GetMenuStringW.argtypes = (wintypes.HMENU, ctypes.c_uint, wintypes.LPWSTR, ctypes.c_int, ctypes.c_uint)
GetMenuStringW.restype = ctypes.c_int

GetWindowRect = user32.GetWindowRect
GetWindowRect.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.RECT))
GetWindowRect.restype = wintypes.BOOL

IsWindow = user32.IsWindow
IsWindow.argtypes = (wintypes.HWND,)
IsWindow.restype = wintypes.BOOL

GWL_WNDPROC = -4
GWL_HINSTANCE = -6
GCL_WNDPROC = -24
GCL_HMODULE = -16

TH32CS_SNAPMODULE = 0x00000008
TH32CS_SNAPMODULE32 = 0x00000010

class MODULEENTRY32(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("th32ModuleID", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("GlblcntUsage", wintypes.DWORD),
        ("ProccntUsage", wintypes.DWORD),
        ("modBaseAddr", ctypes.POINTER(ctypes.c_byte)),
        ("modBaseSize", wintypes.DWORD),
        ("hModule", wintypes.HMODULE),
        ("szModule", ctypes.c_char * 256),
        ("szExePath", ctypes.c_char * 260),
    ]

CreateToolhelp32Snapshot = kernel32.CreateToolhelp32Snapshot
CreateToolhelp32Snapshot.restype = wintypes.HANDLE
CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)

Module32First = kernel32.Module32First
Module32First.argtypes = (wintypes.HANDLE, ctypes.POINTER(MODULEENTRY32))
Module32First.restype = wintypes.BOOL

Module32Next = kernel32.Module32Next
Module32Next.argtypes = (wintypes.HANDLE, ctypes.POINTER(MODULEENTRY32))
Module32Next.restype = wintypes.BOOL

CloseHandle = kernel32.CloseHandle
CloseHandle.argtypes = (wintypes.HANDLE,)
CloseHandle.restype = wintypes.BOOL

IsWow64Process = kernel32.IsWow64Process
IsWow64Process.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.BOOL))
IsWow64Process.restype = wintypes.BOOL

OpenProcess = kernel32.OpenProcess
OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
OpenProcess.restype = wintypes.HANDLE

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010

# Registry
RegOpenKeyEx = advapi32.RegOpenKeyExW
RegOpenKeyEx.argtypes = (wintypes.HKEY, wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(wintypes.HKEY))
RegOpenKeyEx.restype = wintypes.LONG

RegEnumKeyEx = advapi32.RegEnumKeyExW
RegEnumKeyEx.argtypes = (wintypes.HKEY, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD),
                         ctypes.POINTER(wintypes.DWORD), wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD),
                         ctypes.c_void_p)
RegEnumKeyEx.restype = wintypes.LONG

RegCloseKey = advapi32.RegCloseKey
RegCloseKey.argtypes = (wintypes.HKEY,)
RegCloseKey.restype = wintypes.LONG

HKEY_LOCAL_MACHINE = wintypes.HKEY(0x80000002)
KEY_READ = 0x20019

# MSAA
if oleacc is not None:
    AccessibleObjectFromWindow = oleacc.AccessibleObjectFromWindow
    # IA GUID for IAccessible
    # Note: comtypes may be unavailable on Python 3.8 minimal boxes; we only need GUID struct if comtypes exists.
    if comtypes:
        IID_IAccessible = GUID("{618736E0-3C3D-11CF-810C-00AA00389B71}")
    else:
        # define a tiny GUID surrogate for the call
        class _GUID(ctypes.Structure):
            _fields_ = [("Data1", ctypes.c_ulong),
                        ("Data2", ctypes.c_ushort),
                        ("Data3", ctypes.c_ushort),
                        ("Data4", ctypes.c_ubyte * 8)]
        IID_IAccessible = _GUID(0x618736E0, 0x3C3D, 0x11CF,
                                (ctypes.c_ubyte * 8)(0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71))
    AccessibleObjectFromWindow.argtypes = (wintypes.HWND, wintypes.DWORD, ctypes.POINTER(type(IID_IAccessible)), ctypes.POINTER(ctypes.c_void_p))
    AccessibleObjectFromWindow.restype = ctypes.HRESULT

# ---------- helpers ----------
def safe_str(s):
    # Decode using latin-1 per constraint; never fancy.
    if s is None:
        return ""
    if isinstance(s, bytes):
        try:
            return s.decode(ENCODING, "replace")
        except Exception:
            return s.decode("ascii", "replace")
    return str(s)

def mkdir_p(path):
    os.makedirs(path, exist_ok=True)

def dump_text(path, text):
    with open(path, "w", encoding=ENCODING, errors="replace") as f:
        f.write(text)

def dump_json(path, data):
    # latin-1 with ensure_ascii=False still yields 8-bit text; good enough for tooling that reads bytes.
    with open(path, "w", encoding=ENCODING, errors="replace") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)

# ---------- processes & windows ----------
def find_explorer_pids():
    procs = []
    if psutil:
        for p in psutil.process_iter(["pid", "name", "exe"]):
            try:
                name = (p.info.get("name") or "").lower()
                if name == "explorer.exe":
                    procs.append((p.info["pid"], p.info.get("exe") or ""))
            except Exception:
                pass
    else:
        # Simple fallback via tasklist
        try:
            import subprocess
            out = subprocess.check_output(["tasklist", "/fi", "imagename eq explorer.exe", "/fo", "csv"], universal_newlines=True)
            for line in out.splitlines():
                if '"' in line:
                    parts = [p.strip('"') for p in line.split('","') if p.strip()]
                    if parts and parts[0].lower() == "explorer.exe":
                        pid = int(parts[1])
                        procs.append((pid, ""))
        except Exception:
            pass
    return procs

def enum_top_windows_for_pid(target_pid):
    hwnds = []
    @WNDENUMPROC
    def _cb(hwnd, lparam):
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        try:
            if pid.value == target_pid and user32.IsWindowVisible(hwnd):
                hwnds.append(hwnd)
        except Exception:
            pass
        return True
    EnumWindows(_cb, 0)
    return hwnds

def enum_children(hwnd):
    children = []
    @WNDENUMPROC
    def _cb(child, lparam):
        children.append(child)
        return True
    try:
        EnumChildWindows(hwnd, _cb, 0)
    except Exception:
        pass
    return children

def get_window_basic_info(hwnd):
    try:
        txt_len = GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(max(1, txt_len + 1))
        GetWindowTextW(hwnd, buf, txt_len + 1)
        text = buf.value
    except Exception:
        text = ""
    try:
        clsbuf = ctypes.create_unicode_buffer(256)
        GetClassNameW(hwnd, clsbuf, 256)
        clsname = clsbuf.value
    except Exception:
        clsname = ""
    pid = wintypes.DWORD()
    tid = GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    procid = pid.value
    rect = wintypes.RECT()
    ok = GetWindowRect(hwnd, ctypes.byref(rect))
    rect_tup = (rect.left, rect.top, rect.right, rect.bottom) if ok else None
    try:
        wndproc = GetWindowLongPtrW(hwnd, GWL_WNDPROC)
    except Exception:
        wndproc = None
    try:
        class_wndproc = GetClassLongPtrW(hwnd, GCL_WNDPROC)
    except Exception:
        class_wndproc = None
    try:
        hinst = GetWindowLongPtrW(hwnd, GWL_HINSTANCE)
    except Exception:
        hinst = None
    # menu
    hmenu = GetMenu(hwnd)
    menu_items = []
    if hmenu:
        cnt = GetMenuItemCount(hmenu)
        for i in range(max(0, cnt)):
            buf = ctypes.create_unicode_buffer(512)
            ret = GetMenuStringW(hmenu, i, buf, 512, 0x400)
            if ret > 0:
                menu_items.append(buf.value)
    return {
        "hwnd": int(hwnd),
        "hwnd_hex": hex(hwnd),
        "class": clsname,
        "text": text,
        "process_id": procid,
        "thread_id": tid,
        "rect": rect_tup,
        "wndproc": (hex(wndproc) if wndproc else None),
        "class_wndproc": (hex(class_wndproc) if class_wndproc else None),
        "hinstance": (hex(hinst) if hinst else None),
        "menu_items": menu_items,
    }

def recurse_window_tree(hwnd, depth=0, max_depth=50, seen=None):
    if seen is None:
        seen = set()
    if depth > max_depth:
        return []
    if hwnd in seen:
        return []
    seen.add(hwnd)
    info = get_window_basic_info(hwnd)
    children = enum_children(hwnd)
    info["children"] = [recurse_window_tree(c, depth+1, max_depth, seen) for c in children]
    return info

# ---------- modules & address resolution ----------
TH32_SNAP = TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32

def list_modules_for_pid(pid):
    mods = []
    hSnap = CreateToolhelp32Snapshot(TH32_SNAP, pid)
    if not hSnap or hSnap == wintypes.HANDLE(-1).value:
        return mods
    me32 = MODULEENTRY32()
    me32.dwSize = ctypes.sizeof(MODULEENTRY32)
    success = Module32First(hSnap, ctypes.byref(me32))
    while success:
        try:
            baseptr = ctypes.addressof(me32.modBaseAddr.contents) if me32.modBaseAddr else None
            mods.append({
                "szModule": safe_str(me32.szModule),
                "szExePath": safe_str(me32.szExePath),
                "baseaddr": baseptr,
                "baseaddr_hex": (hex(baseptr) if baseptr else None),
                "size": int(me32.modBaseSize),
                "hModule": int(me32.hModule),
            })
        except Exception:
            pass
        success = Module32Next(hSnap, ctypes.byref(me32))
    CloseHandle(hSnap)
    return mods

def is_process_wow64(pid):
    try:
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
        if not hProc:
            return False
        wow = wintypes.BOOL()
        ok = IsWow64Process(hProc, ctypes.byref(wow))
        CloseHandle(hProc)
        if not ok:
            return False
        return bool(wow.value)
    except Exception:
        return False

def resolve_addr_to_module(addr_hex, modules):
    try:
        addr = int(addr_hex, 16)
    except Exception:
        return None
    for m in modules:
        base = m.get("baseaddr")
        size = m.get("size") or 0
        if base and (base <= addr < base + size):
            off = addr - base
            return {"module": m.get("szModule"), "path": m.get("szExePath"),
                    "baseaddr_hex": m.get("baseaddr_hex"), "offset": hex(off)}
    # last-resort hex compare
    for m in modules:
        try:
            base_hex = m.get("baseaddr_hex")
            if base_hex and addr >= int(base_hex, 16):
                off = addr - int(base_hex, 16)
                return {"module": m.get("szModule"), "path": m.get("szExePath"),
                        "baseaddr_hex": base_hex, "offset": hex(off)}
        except Exception:
            pass
    return None

def try_resolve_export_symbol(module_path, offset_hex):
    if not pefile:
        return None
    try:
        pe = pefile.PE(module_path, fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_EXPORT']])
        exports = getattr(pe, "DIRECTORY_ENTRY_EXPORT", None)
        if not exports:
            return None
        off = int(offset_hex, 16)
        for exp in exports.symbols:
            if exp.address is None:
                continue
            if exp.address == off:
                return {"export_name": safe_str(exp.name) if exp.name else None, "ordinal": exp.ordinal}
        # closest lower RVA
        closest = None
        for exp in exports.symbols:
            if exp.address is None:
                continue
            if exp.address <= off:
                if closest is None or exp.address > closest.address:
                    closest = exp
        if closest:
            return {"export_name": safe_str(closest.name) if closest.name else None,
                    "ordinal": closest.ordinal, "closest_rva": hex(closest.address)}
    except Exception:
        return None
    return None

# ---------- registry hints ----------
def dump_known_shell_ext_keys():
    keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions",
        r"SOFTWARE\Classes\*\shellex",
        r"SOFTWARE\Classes\Folder\shellex",
        r"SOFTWARE\Classes\Directory\shellex",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers",
    ]
    collected = {}
    for k in keys:
        sk = wintypes.HKEY()
        rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, k, 0, KEY_READ, ctypes.byref(sk))
        if rc != 0:
            collected[k] = {"error": "RegOpenKeyEx rc=" + str(rc)}
            continue
        items = []
        i = 0
        while True:
            name_buf = ctypes.create_unicode_buffer(512)
            name_len = wintypes.DWORD(512)
            rc = RegEnumKeyEx(sk, i, name_buf, ctypes.byref(name_len), None, None, None, None)
            if rc != 0:
                break
            items.append(name_buf.value)
            i += 1
        RegCloseKey(sk)
        collected[k] = items
    return collected

# ---------- UIA/MSAA ----------
def dump_uia_tree_for_hwnd(hwnd, outpath):
    if LEGACY_MODE:
        dump_text(outpath, "UIA skipped in LEGACY_MODE unless explicitly enabled and available.\n")
        return
    if not comtypes:
        dump_text(outpath, "UIA unavailable: comtypes not installed.\n")
        return
    try:
        uia = cc.CreateObject("UIAutomationClient.CUIAutomation", dynamic=True)
        root = uia.ElementFromHandle(int(hwnd))
        def rec(e, depth):
            out = []
            name = safe_str(getattr(e, "CurrentName", "")) if hasattr(e, "CurrentName") else ""
            cls = safe_str(getattr(e, "CurrentClassName", "")) if hasattr(e, "CurrentClassName") else ""
            control = getattr(e, "CurrentControlType", None)
            control_t = int(control) if control is not None else None
            out.append({"depth": depth, "name": name, "class": cls, "control_type": control_t})
            walker = uia.ControlViewWalker
            child = walker.GetFirstChildElement(e)
            while child:
                out.extend(rec(child, depth + 1))
                child = walker.GetNextSiblingElement(child)
            return out
        tree = rec(root, 0)
        dump_json(outpath, tree)
    except Exception as ex:
        dump_text(outpath, "UIA walk failed: " + str(ex) + "\n" + traceback.format_exc())

def msaa_dump_for_hwnd(hwnd, outpath):
    if oleacc is None:
        dump_text(outpath, "MSAA oleacc not available.\n")
        return
    try:
        pacc = ctypes.c_void_p()
        OBJID_CLIENT = (-4) & 0xffffffff
        hr = AccessibleObjectFromWindow(wintypes.HWND(hwnd), OBJID_CLIENT, ctypes.byref(IID_IAccessible), ctypes.byref(pacc))
        if hr != 0:
            dump_text(outpath, "AccessibleObjectFromWindow failed hr=" + str(hr) + "\n")
            return
        dump_text(outpath, "MSAA IAccessible pointer: " + (hex(pacc.value) if pacc.value else "0x0") + "\n")
    except Exception as ex:
        dump_text(outpath, "MSAA failed: " + str(ex) + "\n" + traceback.format_exc())

# ---------- formatting ----------
def prettify_window_tree(info, indent=0):
    s = []
    pad = "  " * indent
    s.append(pad + "HWND " + info.get('hwnd_hex', '0x0') +
             " Class=" + safe_str(info.get('class')) +
             " Text=" + repr(safe_str(info.get('text'))) +
             " PID=" + str(info.get('process_id')) +
             " WNDPROC=" + safe_str(info.get('wndproc')) +
             " CLASS_WNDPROC=" + safe_str(info.get('class_wndproc')))
    if info.get("rect"):
        s.append(pad + "  RECT=" + repr(info["rect"]))
    if info.get("menu_items"):
        for m in info["menu_items"]:
            s.append(pad + "    Menu: " + safe_str(m))
    for c in info.get("children", []):
        if isinstance(c, dict):
            s.append(prettify_window_tree(c, indent + 1))
        else:
            s.append(pad + "  <bad child entry>")
    return "\n".join(s)

# ---------- main ----------
def run_dump(args):
    # Outdir
    ts_short = datetime.datetime.now().strftime("%m-%d-%Y_%H%M%S")
    outdir_base = args.outdir or ("explorer_gui_dump_" + ts_short)
    mkdir_p(outdir_base)

    meta = {
        "generated_at": now_ts(),
        "python": sys.version,
        "platform": sys.platform,
        "legacy_mode": bool(LEGACY_MODE),
        "os_version": {"major": WIN_MAJOR, "minor": WIN_MINOR, "build": winver.build},
        "options": {"uia": bool(args.uia), "resolve": bool(args.resolve)}
    }
    dump_json(os.path.join(outdir_base, "meta.json"), meta)

    explorer_procs = find_explorer_pids()
    if not explorer_procs:
        # fallback: dump all visible top-level hwnds
        all_hwnds = []
        @WNDENUMPROC
        def _cb(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                all_hwnds.append(hwnd)
            return True
        EnumWindows(_cb, 0)
        dump_text(os.path.join(outdir_base, "all_top_windows.txt"), "\n".join([hex(h) for h in all_hwnds]))
        return

    # class names of interest across versions
    known_explorer_classes = [
        "CabinetWClass", "ExploreWClass", "WorkerW", "Progman",
        "Shell_TrayWnd", "DirectUIHWND", "DUIViewWndClassName",
        "ApplicationFrameWindow", "ATL:Window:1"
    ]

    all_summary = {}

    for pid, exe in explorer_procs:
        summary = {"pid": pid, "exe": exe, "collected_at": now_ts()}
        try:
            summary["is_wow64"] = is_process_wow64(pid)

            hwnds = enum_top_windows_for_pid(pid)
            summary["top_windows_count"] = len(hwnds)

            trees = []
            pretty_parts = []
            for h in hwnds:
                tree = recurse_window_tree(h)
                trees.append(tree)
                pretty_parts.append(prettify_window_tree(tree))
                pretty_parts.append("-" * 80)
            dump_text(os.path.join(outdir_base, "explorer_%d_windows.txt" % pid), "\n".join(pretty_parts))
            dump_json(os.path.join(outdir_base, "explorer_%d_windows.json" % pid), trees)

            # modules
            mods = list_modules_for_pid(pid)
            summary["modules_count"] = len(mods)
            dump_json(os.path.join(outdir_base, "explorer_%d_modules.json" % pid), mods)
            mod_lines = []
            for m in mods:
                mod_lines.append("%s %s base=%s size=%s" % (m.get("szModule"), m.get("szExePath"),
                                                            m.get("baseaddr_hex"), str(m.get("size"))))
            dump_text(os.path.join(outdir_base, "explorer_%d_modules.txt" % pid), "\n".join(mod_lines))

            # proc map
            proc_map = []
            def collect_procs(node):
                proc_map.append({
                    "hwnd": node.get("hwnd"),
                    "hwnd_hex": node.get("hwnd_hex"),
                    "class": node.get("class"),
                    "wndproc": node.get("wndproc"),
                    "class_wndproc": node.get("class_wndproc"),
                    "text": node.get("text"),
                })
                for c in node.get("children", []):
                    if isinstance(c, dict):
                        collect_procs(c)
            for t in trees:
                if isinstance(t, dict):
                    collect_procs(t)
            dump_json(os.path.join(outdir_base, "explorer_%d_proc_map.json" % pid), proc_map)

            # registry hints
            reg = dump_known_shell_ext_keys()
            dump_json(os.path.join(outdir_base, "explorer_%d_registry_shellext_hints.json" % pid), reg)

            # resolve addresses to modules/exports only when requested
            if args.resolve:
                resolved = []
                for entry in proc_map:
                    for field in ("wndproc", "class_wndproc"):
                        val = entry.get(field)
                        if val:
                            r = resolve_addr_to_module(val, mods)
                            if r:
                                info = {"hwnd": entry.get("hwnd"), "field": field, "addr": val, "module": r}
                                if pefile and r.get("path"):
                                    sym = try_resolve_export_symbol(r.get("path"), r.get("offset"))
                                    if sym:
                                        info["export"] = sym
                                resolved.append(info)
                dump_json(os.path.join(outdir_base, "explorer_%d_resolved_wndprocs.json" % pid), resolved)
                summary["resolved_wndprocs_count"] = len(resolved)

            # UIA: by default skipped in LEGACY_MODE unless user explicitly passed --uia
            if args.uia:
                for h in hwnds:
                    dump_uia_tree_for_hwnd(h, os.path.join(outdir_base, "explorer_%d_uia_hwnd_%d.json" % (pid, int(h))))

            # MSAA is useful on Vista/7 too
            for h in hwnds:
                msaa_dump_for_hwnd(h, os.path.join(outdir_base, "explorer_%d_msaa_hwnd_%d.txt" % (pid, int(h))))

            if psutil:
                try:
                    p = psutil.Process(pid)
                    more = {
                        "cmdline": p.cmdline(),
                        "exe": p.exe(),
                        "username": p.username(),
                        "create_time": datetime.datetime.fromtimestamp(p.create_time()).strftime(TS_FMT),
                    }
                    dump_json(os.path.join(outdir_base, "explorer_%d_process_info.json" % pid), more)
                except Exception:
                    pass

        except Exception as e:
            dump_text(os.path.join(outdir_base, "explorer_%d_error.txt" % pid),
                      "Exception while collecting: " + str(e) + "\n\n" + traceback.format_exc())
        all_summary[pid] = summary

    dump_json(os.path.join(outdir_base, "summary.json"), all_summary)
    print("Dump completed. Output in folder: " + outdir_base)

# ---------- CLI ----------
def parse_args():
    p = argparse.ArgumentParser(description="Explorer GUI dumper (Vista/7/8/8.1/10/11) with Python 3.8 compat.")
    p.add_argument("--uia", action="store_true", help="Attempt UI Automation (requires comtypes).")
    p.add_argument("--resolve", action="store_true", help="Resolve WNDPROC/class WNDPROC to module+export (pefile optional).")
    p.add_argument("--outdir", type=str, help="Output directory base.")
    return p.parse_args()

if __name__ == "__main__":
    try:
        args = parse_args()
        run_dump(args)
    except Exception as e:
        print("Fatal error:", e)
        print(traceback.format_exc())
        sys.exit(1)

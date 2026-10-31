# -*- coding: utf-8 -*-
"""
Monitor detection diagnostics.
Run on the target host:  python diag_monitor.py
Writes full results to diag_output.txt next to this script (UTF-8).
Console output stays English/ASCII to avoid GBK mojibake.
"""
import os
import sys
import ctypes
import datetime
from ctypes import wintypes

out = []


def log(s=""):
    out.append(str(s))


# --- [1] Registry: HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY ----------------
log("=" * 64)
log("[1] Registry Enum\\DISPLAY tree (EDID storage)")
try:
    import winreg
    base = r"SYSTEM\CurrentControlSet\Enum\DISPLAY"
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as root:
        i = 0
        while True:
            try:
                pnp = winreg.EnumKey(root, i)
                i += 1
            except OSError:
                break
            log("  PNP code: %s" % pnp)
            try:
                with winreg.OpenKey(root, pnp) as dc:
                    j = 0
                    while True:
                        try:
                            inst = winreg.EnumKey(dc, j)
                            j += 1
                        except OSError:
                            break
                        log("    instance: %s" % inst)
                        try:
                            with winreg.OpenKey(dc, inst + "\\Device Parameters") as dp:
                                try:
                                    edid, _ = winreg.QueryValueEx(dp, "EDID")
                                    log("      EDID: %d bytes, header=%s"
                                        % (len(edid), bytes(edid[:8]).hex()))
                                except OSError as e:
                                    log("      EDID: <none> (%s)" % e)
                        except OSError as e:
                            log("      Device Parameters: <open failed> %s" % e)
            except OSError as e:
                log("    <open failed> %s" % e)
except Exception as e:
    log("  registry error: %r" % (e,))

# --- [2] WMI: root\wmi WmiMonitorID -------------------------------------------
log("")
log("=" * 64)
log("[2] WMI root\\wmi -> WmiMonitorID")
try:
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    loc = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    svc = loc.ConnectServer(".", "root\\wmi")

    def chars(v):
        try:
            return "".join(chr(c) for c in (v or []) if c != 0)
        except Exception:
            return ""

    n = 0
    for item in svc.ExecQuery("SELECT * FROM WmiMonitorID"):
        props = {}
        for p in item.Properties_:
            try:
                props[p.Name] = p.Value
            except Exception:
                props[p.Name] = "<err>"
        n += 1
        log("  entry %d:" % n)
        log("    InstanceName : %s" % props.get("InstanceName"))
        log("    Active       : %s" % props.get("Active"))
        log("    Name         : %r" % chars(props.get("UserFriendlyName")))
        log("    MfgCode      : %r" % chars(props.get("ManufacturerName")))
        log("    Serial       : %r" % chars(props.get("SerialNumberID")))
        log("    Year/Week    : %s / %s" % (props.get("YearOfManufacture"),
                                           props.get("WeekOfManufacture")))
    if n == 0:
        log("  <no entries>")
except Exception as e:
    log("  wmi error: %r" % (e,))

# --- [3] Win32: EnumDisplayDevices --------------------------------------------
log("")
log("=" * 64)
log("[3] Win32 EnumDisplayDevices (adapters and their monitors)")
try:
    class DISPLAY_DEVICEW(ctypes.Structure):
        _fields_ = [("cb", wintypes.DWORD),
                    ("DeviceName", wintypes.WCHAR * 32),
                    ("DeviceString", wintypes.WCHAR * 128),
                    ("StateFlags", wintypes.DWORD),
                    ("DeviceID", wintypes.WCHAR * 128),
                    ("DeviceKey", wintypes.WCHAR * 128)]

    u32 = ctypes.windll.user32
    i = 0
    while True:
        dd = DISPLAY_DEVICEW()
        dd.cb = ctypes.sizeof(DISPLAY_DEVICEW)
        if not u32.EnumDisplayDevicesW(None, i, ctypes.byref(dd), 0):
            break
        i += 1
        log("  adapter: %s | %s" % (dd.DeviceName, dd.DeviceString))
        log("    DeviceID  : %s" % dd.DeviceID)
        log("    StateFlags: 0x%08X (active=%s)"
            % (dd.StateFlags, bool(dd.StateFlags & 0x1)))
        j = 0
        while True:
            md = DISPLAY_DEVICEW()
            md.cb = ctypes.sizeof(DISPLAY_DEVICEW)
            if not u32.EnumDisplayDevicesW(dd.DeviceName, j, ctypes.byref(md), 0):
                break
            j += 1
            log("    monitor: %s | %s" % (md.DeviceName, md.DeviceString))
            log("      DeviceID  : %s" % md.DeviceID)
            log("      StateFlags: 0x%08X" % md.StateFlags)
except Exception as e:
    log("  enum error: %r" % (e,))

log("")
log("generated: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])),
                    "diag_output.txt")
with open(path, "w", encoding="utf-8-sig") as f:
    f.write("\n".join(out))
print("Diagnostics written to:")
print(path)

# -*- coding: utf-8 -*-
"""
硬件检测工具 (HWInfo Tool)
==========================
Windows 7 ~ Windows 11 (32/64 位) 硬件信息检测小工具。
- 依赖: 仅 pywin32 + Python 标准库 + Tkinter
- WMI 访问: 直接使用 SWbemLocator COM 接口 (不经过 moniker 解析,
  避免 wmi 第三方库在 PyInstaller 冻结环境下报 0x800401E4 的问题)
- 打包: PyInstaller 单文件 EXE, 绿色免安装
- 容错: 任何一项检测失败均显示 "未知" / "未检测到", 不崩溃
"""

import os
import sys
import socket
import ctypes
import datetime
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "硬件检测工具"
APP_VERSION = "1.5.0"
UNKNOWN = "未知"
NOT_FOUND = "未检测到"

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def safe(fn, default=UNKNOWN):
    """执行 fn(), 任何异常或空值都返回 default, 保证不崩溃。"""
    try:
        v = fn()
        if v is None:
            return default
        if isinstance(v, str) and not v.strip():
            return default
        return v
    except Exception:
        return default


def fmt_bytes(n):
    """字节数 -> 人类可读容量。"""
    try:
        n = float(n)
    except Exception:
        return UNKNOWN
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024.0 or unit == "TB":
            return "%.2f %s" % (n, unit)
        n /= 1024.0
    return UNKNOWN


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def cjk_width(s):
    """计算字符串显示宽度 (中文字符计 2)。"""
    return sum(2 if ord(c) > 0x2E7F else 1 for c in str(s))


# ---------------------------------------------------------------------------
# WMI 底层访问 (SWbemLocator, PyInstaller 冻结环境可靠)
# ---------------------------------------------------------------------------

class WmiClient(object):
    """基于 win32com 的 WMI 查询封装。必须在调用线程内创建 (内部完成
    CoInitialize)。任何失败只记录 error, 不抛异常。"""

    def __init__(self):
        self.available = False
        self.error = None
        self._locator = None
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            self._locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            self.available = True
        except Exception as e:
            self.error = str(e)

    def query(self, wql, namespace="root\\cimv2"):
        """执行 WQL 查询, 返回 [ {属性名: 值}, ... ]。失败返回 []。"""
        if not self.available:
            return []
        try:
            svc = self._locator.ConnectServer(".", namespace)
            rows = []
            for item in svc.ExecQuery(wql):
                row = {}
                for prop in item.Properties_:
                    try:
                        v = prop.Value
                    except Exception:
                        v = None
                    if isinstance(v, tuple):
                        v = list(v)
                    row[prop.Name] = v
                rows.append(row)
            return rows
        except Exception:
            return []


# ---------------------------------------------------------------------------
# 硬件信息采集
# ---------------------------------------------------------------------------

class HardwareCollector(object):
    """逐项采集硬件信息。所有方法均内部容错。"""

    DDR_TYPES = {
        20: "DDR", 21: "DDR2", 22: "DDR2 FB-DIMM", 24: "DDR3",
        25: "DDR3", 26: "DDR4", 27: "LPDDR4", 30: "LPDDR5", 34: "DDR5",
    }

    MONITOR_VENDORS = {
        "DEL": "戴尔 Dell", "LEN": "联想 Lenovo", "HWP": "惠普 HP",
        "HPN": "惠普 HP", "AOC": "冠捷 AOC", "SAM": "三星 Samsung",
        "LGD": "LG", "AUO": "友达 AUO", "BOE": "京东方 BOE",
        "CMN": "群创 Innolux", "ACR": "宏碁 Acer", "BNQ": "明基 BenQ",
        "PHL": "飞利浦 Philips", "HKC": "惠科 HKC", "MSI": "微星 MSI",
        "AUS": "华硕 ASUS", "XMI": "小米 Xiaomi", "HWV": "华为 Huawei",
    }

    def __init__(self):
        self.wmi = WmiClient()
        self.wmi_error = None if self.wmi.available else self.wmi.error

    # -- 通用 -----------------------------------------------------------------
    @staticmethod
    def _fmt_wmi_time(s):
        # WMI 时间格式: 20241001120000.000000+480
        try:
            return "%s-%s-%s %s:%s:%s" % (
                s[0:4], s[4:6], s[6:8], s[8:10], s[10:12], s[12:14])
        except Exception:
            return UNKNOWN

    @staticmethod
    def _edid_str(char_list):
        """EDID 字符数组 (uint16 列表) -> 字符串, 截断于 0。"""
        try:
            chars = []
            for c in char_list or []:
                if c == 0:
                    break
                chars.append(chr(c))
            return "".join(chars).strip()
        except Exception:
            return ""

    # -- 计算机 / 操作系统 ------------------------------------------------------
    def collect_system(self):
        items = []
        items.append(("计算机名称", safe(lambda: socket.gethostname())))
        items.append(("当前用户", safe(lambda: os.environ.get("USERNAME", ""))))
        items.append(("管理员权限", "是" if is_admin() else "否（普通权限，部分信息可能不完整）"))

        rows = self.wmi.query("SELECT Caption, BuildNumber, OSArchitecture, Version, InstallDate FROM Win32_OperatingSystem")
        if rows:
            o = rows[0]
            items.append(("操作系统", safe(lambda: "%s（内部版本 %s，%s）" % (
                o.get("Caption"), o.get("BuildNumber"), o.get("OSArchitecture")))))
            items.append(("系统版本号", safe(lambda: o.get("Version"))))
            items.append(("安装日期", safe(lambda: self._fmt_wmi_time(o.get("InstallDate")))))
        else:
            import platform
            items.append(("操作系统", safe(lambda: "%s %s (%s)" % (
                platform.system(), platform.version(), platform.machine()))))
        return items

    # -- CPU -------------------------------------------------------------------
    def collect_cpu(self):
        items = []
        rows = self.wmi.query(
            "SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, "
            "L2CacheSize, L3CacheSize, Manufacturer, SocketDesignation FROM Win32_Processor")
        if not rows:
            import platform
            items.append(("型号", safe(lambda: platform.processor(), NOT_FOUND)))
            return items
        c = rows[0]
        items.append(("型号", safe(lambda: str(c.get("Name")).strip())))
        items.append(("核心 / 线程", "%s / %s" % (
            safe(lambda: c.get("NumberOfCores")), safe(lambda: c.get("NumberOfLogicalProcessors")))))
        items.append(("最大主频", safe(lambda: "%s MHz" % c.get("MaxClockSpeed"))))
        items.append(("二级缓存", safe(lambda: "%s KB" % c.get("L2CacheSize"))))
        items.append(("三级缓存", safe(lambda: "%s KB" % c.get("L3CacheSize"))))
        items.append(("制造商", safe(lambda: c.get("Manufacturer"))))
        items.append(("插槽", safe(lambda: c.get("SocketDesignation"))))
        if len(rows) > 1:
            items.append(("处理器数量", "%d 路" % len(rows)))
        return items

    # -- 主板 / BIOS -------------------------------------------------------------
    def collect_mainboard(self):
        items = []
        bb = self.wmi.query("SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard")
        if bb:
            b = bb[0]
            items.append(("主板制造商", safe(lambda: str(b.get("Manufacturer")).strip())))
            items.append(("主板型号", safe(lambda: str(b.get("Product")).strip())))
            items.append(("主板序列号", safe(lambda: str(b.get("SerialNumber")).strip())))
        else:
            items.append(("主板", NOT_FOUND))
        bios = self.wmi.query("SELECT SMBIOSBIOSVersion, Manufacturer, ReleaseDate, SerialNumber FROM Win32_BIOS")
        if bios:
            b = bios[0]
            items.append(("BIOS 版本", safe(lambda: b.get("SMBIOSBIOSVersion"))))
            items.append(("BIOS 厂商", safe(lambda: b.get("Manufacturer"))))
            items.append(("BIOS 日期", safe(lambda: self._fmt_wmi_time(b.get("ReleaseDate"))[:10])))
        else:
            items.append(("BIOS", NOT_FOUND))
        return items

    # -- 内存 --------------------------------------------------------------------
    def collect_memory(self):
        items = []
        cs = self.wmi.query("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
        if cs:
            items.append(("总容量", safe(lambda: fmt_bytes(int(cs[0].get("TotalPhysicalMemory") or 0)))))
        rows = self.wmi.query(
            "SELECT Capacity, Speed, ConfiguredClockSpeed, DeviceLocator, Manufacturer, "
            "PartNumber, SMBIOSMemoryType, SerialNumber FROM Win32_PhysicalMemory")
        if not rows:
            items.append(("内存条", NOT_FOUND))
            return items
        items.append(("已用插槽", "%d 条" % len(rows)))
        for i, m in enumerate(rows, 1):
            memtype = self.DDR_TYPES.get(safe(lambda: int(m.get("SMBIOSMemoryType") or 0), 0), UNKNOWN)
            items.append(("模块 %d（%s）" % (i, safe(lambda: m.get("DeviceLocator"))), ""))
            items.append(("  容量", safe(lambda: fmt_bytes(int(m.get("Capacity") or 0)))))
            items.append(("  规格", memtype))
            items.append(("  标称 / 实际频率", "%s / %s MHz" % (
                safe(lambda: m.get("Speed")), safe(lambda: m.get("ConfiguredClockSpeed")))))
            items.append(("  制造商", safe(lambda: str(m.get("Manufacturer")).strip())))
            items.append(("  型号", safe(lambda: str(m.get("PartNumber")).strip(), "—")))
        return items

    # -- 显卡 --------------------------------------------------------------------
    @staticmethod
    def _vram_from_registry(gpu_name):
        """从注册表读取显卡真实显存 (64 位 qwMemorySize)。
        WMI 的 AdapterRAM 为 32 位, >4GB 会溢出/变负, 不可靠。只读不写。"""
        import winreg
        base = r"SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as root:
                i = 0
                while True:
                    try:
                        sub = winreg.EnumKey(root, i)
                        i += 1
                    except OSError:
                        break
                    try:
                        with winreg.OpenKey(root, sub) as k:
                            desc, _ = winreg.QueryValueEx(k, "DriverDesc")
                            if desc and str(desc).strip().lower() == str(gpu_name).strip().lower():
                                try:
                                    v, _ = winreg.QueryValueEx(k, "HardwareInformation.qwMemorySize")
                                    if int(v) > 0:
                                        return int(v)
                                except OSError:
                                    return None
                    except OSError:
                        continue
        except OSError:
            pass
        return None

    def collect_gpu(self):
        items = []
        rows = self.wmi.query(
            "SELECT Name, AdapterRAM, DriverVersion, VideoProcessor, PNPDeviceID, "
            "CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate "
            "FROM Win32_VideoController")
        # 只保留 PCI 总线上的物理显卡, 过滤远程桌面/串流软件的虚拟显示适配器
        virt = [g for g in rows if not str(g.get("PNPDeviceID") or "").upper().startswith("PCI")]
        phys = [g for g in rows if str(g.get("PNPDeviceID") or "").upper().startswith("PCI")]
        rows = phys or rows
        if not rows:
            items.append(("显卡", NOT_FOUND))
            return items
        for i, g in enumerate(rows, 1):
            name = safe(lambda: str(g.get("Name")).strip())
            items.append(("显卡 %d" % i, name))
            vram = safe(lambda: self._vram_from_registry(name), None)
            if vram:
                items.append(("  显存", fmt_bytes(vram)))
            else:
                raw = safe(lambda: int(g.get("AdapterRAM") or 0), 0)
                if raw < 0:
                    raw += 1 << 32  # 有符号溢出修正
                if raw:
                    items.append(("  显存", fmt_bytes(raw) + "（WMI 读数，超 4GB 可能不准）"))
                else:
                    items.append(("  显存", UNKNOWN))
            items.append(("  驱动版本", safe(lambda: g.get("DriverVersion"))))
            items.append(("  当前分辨率", safe(lambda: "%s x %s @ %s Hz" % (
                g.get("CurrentHorizontalResolution"), g.get("CurrentVerticalResolution"),
                g.get("CurrentRefreshRate")))))
        if virt:
            names = "、".join(self._dedupe([safe(lambda: str(v.get("Name")).strip()) for v in virt]))
            items.append(("说明", "检测到虚拟显示适配器（%s），当前可能处于远程会话，"
                                 "以上物理显卡分辨率可能不代表本机实际显示输出" % names))
        return items

    # -- 硬盘 --------------------------------------------------------------------
    # MSFT_PhysicalDisk.MediaType: 3=HDD, 4=SSD, 5=SCM (root\microsoft\windows\storage, Win8+)
    DISK_MEDIA_TYPES = {3: "机械硬盘 (HDD)", 4: "固态硬盘 (SSD)", 5: "SCM"}

    def _physical_disk_media(self):
        """{硬盘型号: 介质描述}, Win7 或无权限时返回空 dict。"""
        result = {}
        rows = self.wmi.query("SELECT FriendlyName, MediaType FROM MSFT_PhysicalDisk",
                              namespace="root\\microsoft\\windows\\storage")
        for r in rows:
            name = safe(lambda: str(r.get("FriendlyName")).strip(), "")
            mt = safe(lambda: int(r.get("MediaType") or 0), 0)
            if name and mt in self.DISK_MEDIA_TYPES:
                result[name] = self.DISK_MEDIA_TYPES[mt]
        return result

    def collect_disk(self):
        items = []
        media_map = self._physical_disk_media()
        rows = self.wmi.query("SELECT Model, Size, InterfaceType, MediaType, SerialNumber, PNPDeviceID FROM Win32_DiskDrive")
        if not rows:
            items.append(("硬盘", NOT_FOUND))
        else:
            for i, d in enumerate(rows, 1):
                model = safe(lambda: str(d.get("Model")).strip())
                size = safe(lambda: int(d.get("Size") or 0), 0)
                # WMI 常把 NVMe 误报为 SCSI, 依据 PNP 设备 ID 修正
                iface = safe(lambda: d.get("InterfaceType"))
                if "NVME" in str(d.get("PNPDeviceID") or "").upper():
                    iface = "NVMe (PCIe)"
                items.append(("硬盘 %d" % i, model))
                items.append(("  容量", fmt_bytes(size) if size else UNKNOWN))
                items.append(("  接口类型", iface))
                # 优先 SSD/HDD 精确介质 (Win8+ 存储命名空间), 回退 WMI 笼统描述
                media = media_map.get(model) or safe(lambda: str(d.get("MediaType")).strip())
                items.append(("  介质", media))
        parts = self.wmi.query(
            "SELECT DeviceID, FileSystem, Size, FreeSpace, VolumeName FROM Win32_LogicalDisk WHERE DriveType = 3")
        for p in parts:
            size = safe(lambda: int(p.get("Size") or 0), 0)
            free = safe(lambda: int(p.get("FreeSpace") or 0), 0)
            if size:
                used_pct = (size - free) * 100.0 / size
                vol = safe(lambda: str(p.get("VolumeName")).strip(), "")
                label = "分区 %s" % p.get("DeviceID") + ("（%s）" % vol if vol else "")
                items.append((label, "%s · %s · 已用 %.1f%%" % (
                    safe(lambda: p.get("FileSystem")), fmt_bytes(size), used_pct)))
        return items

    # -- 显示器 ------------------------------------------------------------------
    @staticmethod
    def _parse_edid(data):
        """解析 128 字节 EDID: 厂商代码 / 型号名 / 序列号 / 出厂年周。"""
        try:
            data = bytes(data)
            if len(data) < 128 or data[0:8] != b"\x00\xff\xff\xff\xff\xff\xff\x00":
                return None
            mid = (data[8] << 8) | data[9]
            code = "".join(chr(((mid >> s) & 31) + 64) for s in (10, 5, 0))
            name, serial_str = "", ""
            for off in (54, 72, 90, 108):  # 4 个描述符块
                blk = data[off:off + 18]
                if blk[0] == 0 and blk[1] == 0 and blk[2] == 0:
                    if blk[3] == 0xFC and not name:  # 显示器名称
                        name = blk[5:17].split(b"\x0a")[0].decode("ascii", "ignore").strip()
                    elif blk[3] == 0xFF and not serial_str:  # 序列号字符串
                        serial_str = blk[5:17].split(b"\x0a")[0].decode("ascii", "ignore").strip()
            serial_num = int.from_bytes(data[12:16], "little")
            # 屏幕物理尺寸 (cm) -> 对角线英寸
            wcm, hcm = data[21], data[22]
            size = int(round(((wcm ** 2 + hcm ** 2) ** 0.5) / 2.54)) if (wcm and hcm) else 0
            return {
                "vendor_code": code,
                "name": name,
                "serial_str": serial_str,
                "serial_num": serial_num,
                "week": data[16],
                "year": data[17] + 1990,
                "size_inch": size,
            }
        except Exception:
            return None

    @classmethod
    def _edid_from_registry(cls):
        """读取注册表 Enum\\DISPLAY 下所有显示器 EDID (Windows 见过的全部显示器,
        不受远程会话影响)。只读操作。"""
        import winreg
        results = []
        base = r"SYSTEM\CurrentControlSet\Enum\DISPLAY"
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base) as root:
                i = 0
                while True:
                    try:
                        pnp = winreg.EnumKey(root, i)
                        i += 1
                    except OSError:
                        break
                    try:
                        with winreg.OpenKey(root, pnp) as dc:
                            j = 0
                            while True:
                                try:
                                    inst = winreg.EnumKey(dc, j)
                                    j += 1
                                except OSError:
                                    break
                                try:
                                    with winreg.OpenKey(dc, inst + "\\Device Parameters") as dp:
                                        edid, _ = winreg.QueryValueEx(dp, "EDID")
                                        info = cls._parse_edid(edid)
                                        if info:
                                            info["pnp"] = pnp
                                            results.append(info)
                                except OSError:
                                    continue
                    except OSError:
                        continue
        except OSError:
            pass
        return results

    @staticmethod
    def _active_monitor_map():
        """EnumDisplayDevices: {显示器PNP代码: (适配器名称, 是否PCI物理卡)}。
        用于区分 当前物理连接 / 远程会话虚拟显示器 / 历史记录。"""
        import ctypes
        from ctypes import wintypes
        result = {}
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
                if not (dd.StateFlags & 0x1):  # DISPLAY_DEVICE_ACTIVE
                    continue
                adapter_phys = str(dd.DeviceID).upper().startswith("PCI")
                j = 0
                while True:
                    md = DISPLAY_DEVICEW()
                    md.cb = ctypes.sizeof(DISPLAY_DEVICEW)
                    if not u32.EnumDisplayDevicesW(dd.DeviceName, j, ctypes.byref(md), 0):
                        break
                    j += 1
                    parts = str(md.DeviceID).split("\\")
                    if len(parts) >= 2 and parts[0].upper() == "MONITOR":
                        result[parts[1].upper()] = (str(dd.DeviceString), adapter_phys)
        except Exception:
            pass
        return result

    def collect_monitor(self):
        items = []
        reg_mons = self._edid_from_registry()
        active = self._active_monitor_map()
        if reg_mons:
            # 排序: 物理显示器优先 (无论当前是否激活), 远程/虚拟显示器排最后
            def sort_key(m):
                a = active.get(str(m.get("pnp", "")).upper())
                if a and not a[1]:
                    return 2  # 挂在虚拟适配器上的远程镜像显示器
                return 0 if a else 1
            reg_mons.sort(key=sort_key)
            for i, m in enumerate(reg_mons, 1):
                code = m.get("vendor_code", "")
                vendor = self.MONITOR_VENDORS.get(code.upper(), code or UNKNOWN)
                name = m.get("name") or "通用即插即用监视器"
                a = active.get(str(m.get("pnp", "")).upper())
                if a and a[1]:
                    status = "当前连接（物理显示器）"
                elif a:
                    status = "当前连接（%s · 远程/虚拟显示适配器）" % a[0]
                else:
                    status = "未激活（历史记录，或远程会话下物理显示器休眠）"
                items.append(("显示器 %d" % i, name))
                items.append(("  品牌", vendor))
                items.append(("  状态", status))
                sn = m.get("serial_str") or ("%08X" % m["serial_num"] if m.get("serial_num") else "")
                if sn:
                    items.append(("  序列号", sn))
                y, w = m.get("year"), m.get("week")
                if y and 1990 <= y <= 2100:
                    items.append(("  出厂日期", "%d 年第 %d 周" % (y, w) if w else "%d 年" % y))
            return items
        # 注册表读取失败时回退 WMI
        found = False
        rows = self.wmi.query(
            "SELECT UserFriendlyName, ManufacturerName, ProductCodeID, SerialNumberID, "
            "YearOfManufacture, WeekOfManufacture FROM WmiMonitorID", namespace="root\\wmi")
        for i, m in enumerate(rows, 1):
            name = self._edid_str(m.get("UserFriendlyName"))
            if not name:
                continue
            found = True
            mfg_code = self._edid_str(m.get("ManufacturerName"))
            vendor = self.MONITOR_VENDORS.get(mfg_code.upper(), mfg_code or UNKNOWN)
            items.append(("显示器 %d" % i, name))
            items.append(("  品牌", vendor))
            sn = self._edid_str(m.get("SerialNumberID"))
            if sn:
                items.append(("  序列号", sn))
            year = safe(lambda: int(m.get("YearOfManufacture") or 0), 0)
            week = safe(lambda: int(m.get("WeekOfManufacture") or 0), 0)
            if year:
                items.append(("  出厂日期", "%d 年第 %d 周" % (year, week) if week else "%d 年" % year))
        if not found:
            items.append(("显示器", NOT_FOUND))
        return items

    # -- 网卡 --------------------------------------------------------------------
    def collect_network(self):
        items = []
        rows = self.wmi.query(
            "SELECT Description, MACAddress, IPAddress, DefaultIPGateway, DHCPEnabled "
            "FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
        if not rows:
            items.append(("网卡（已启用）", NOT_FOUND))
            return items
        for i, n in enumerate(rows, 1):
            ips = n.get("IPAddress") or []
            v4 = [ip for ip in ips if ":" not in ip]
            v6 = [ip for ip in ips if ":" in ip]
            items.append(("网卡 %d" % i, safe(lambda: str(n.get("Description")).strip())))
            items.append(("  MAC 地址", safe(lambda: n.get("MACAddress"))))
            items.append(("  IPv4 地址", "、".join(v4) if v4 else NOT_FOUND))
            if v6:
                items.append(("  IPv6 地址", "、".join(v6)))
            gw = n.get("DefaultIPGateway") or []
            if gw:
                items.append(("  默认网关", "、".join(gw)))
            items.append(("  DHCP", safe(lambda: "启用" if n.get("DHCPEnabled") else "静态")))
        return items

    # -- 外设 --------------------------------------------------------------------
    @staticmethod
    def _dedupe(seq):
        """去重且保持顺序。"""
        seen, out = set(), []
        for x in seq:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out

    def collect_peripherals(self):
        items = []
        mice = self._dedupe([str(m.get("Name")) for m in self.wmi.query("SELECT Name FROM Win32_PointingDevice") if m.get("Name")])
        items.append(("鼠标", "、".join(mice) if mice else NOT_FOUND))
        kbs = self._dedupe([str(k.get("Name")) for k in self.wmi.query("SELECT Name FROM Win32_Keyboard") if k.get("Name")])
        items.append(("键盘", "、".join(kbs) if kbs else NOT_FOUND))
        snds = self._dedupe([str(s.get("Name")) for s in self.wmi.query("SELECT Name FROM Win32_SoundDevice") if s.get("Name")])
        items.append(("音频设备", "、".join(snds) if snds else NOT_FOUND))
        return items

    # -- 简约概览 ------------------------------------------------------------------
    CN_CORES = {1: "单核", 2: "双核", 3: "三核", 4: "四核", 5: "五核", 6: "六核",
                8: "八核", 10: "十核", 12: "十二核", 14: "十四核", 16: "十六核",
                20: "二十核", 24: "二十四核", 32: "三十二核"}

    VIRTUAL_NIC_KEYS = ("tailscale", "tap-", "wintun", "vmware", "virtualbox",
                        "hyper-v", "zerotier", "virtual", "bluetooth", "wan miniport")

    @staticmethod
    def _clean_vendor(name):
        """厂商全名 -> 简短形式。"""
        n = str(name or "").strip()
        for suffix in (" Technology Co., Ltd.", " Co., Ltd.", " Technologies",
                       " Technology", " Corporation", " Corp.", " Inc.",
                       " Ltd.", " Company"):
            n = n.replace(suffix, "")
        return n.strip() or UNKNOWN

    def collect_summary(self):
        """简约概览数据: {"title", "subtitle", "rows": [(项目, 值), ...]}"""
        result = {"title": safe(lambda: socket.gethostname()) + " 台式电脑",
                  "subtitle": "", "rows": []}
        rows = result["rows"]

        o = self.wmi.query("SELECT Caption, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
        if o:
            result["subtitle"] = safe(lambda: "%s（内部版本 %s，%s）" % (
                o[0].get("Caption"), o[0].get("BuildNumber"), o[0].get("OSArchitecture")))

        # 处理器
        c = self.wmi.query("SELECT Name, NumberOfCores, NumberOfLogicalProcessors FROM Win32_Processor")
        if c:
            cores = safe(lambda: int(c[0].get("NumberOfCores") or 0), 0)
            threads = safe(lambda: int(c[0].get("NumberOfLogicalProcessors") or 0), 0)
            extra = ""
            if cores and threads:
                extra = "（%s %d 线程）" % (self.CN_CORES.get(cores, "%d 核" % cores), threads)
            rows.append(("处理器", safe(lambda: str(c[0].get("Name")).strip()) + extra))
        else:
            rows.append(("处理器", NOT_FOUND))

        # 内存
        mems = self.wmi.query("SELECT Capacity, Speed, Manufacturer, SMBIOSMemoryType FROM Win32_PhysicalMemory")
        if mems:
            total = sum(safe(lambda m=m: int(m.get("Capacity") or 0), 0) for m in mems)
            total_gb = int(round(total / 1024.0 ** 3))
            m0 = mems[0]
            per_gb = safe(lambda: int(int(m0.get("Capacity") or 0) / 1024 ** 3), 0)
            ddr = self.DDR_TYPES.get(safe(lambda: int(m0.get("SMBIOSMemoryType") or 0), 0), "")
            detail = "%s %s %sMHz %dGB × %d" % (
                self._clean_vendor(m0.get("Manufacturer")), ddr,
                safe(lambda: m0.get("Speed"), "?"), per_gb, len(mems))
            rows.append(("内存", "%d GB（%s）" % (total_gb, detail)))
        else:
            rows.append(("内存", NOT_FOUND))

        # 显卡 (物理卡)
        gpus = self.wmi.query("SELECT Name, AdapterRAM, PNPDeviceID FROM Win32_VideoController")
        g_phys = [g for g in gpus if str(g.get("PNPDeviceID") or "").upper().startswith("PCI")] or gpus
        if g_phys:
            gname = safe(lambda: str(g_phys[0].get("Name")).strip())
            vram = safe(lambda: self._vram_from_registry(gname), None)
            if not vram:
                raw = safe(lambda: int(g_phys[0].get("AdapterRAM") or 0), 0)
                if raw < 0:
                    raw += 1 << 32
                vram = raw or None
            vram_gb = int(round(vram / 1024.0 ** 3)) if vram else 0
            rows.append(("显卡", gname + ("（%d GB）" % vram_gb if vram_gb else "")))
        else:
            rows.append(("显卡", NOT_FOUND))

        # 主板
        bb = self.wmi.query("SELECT Manufacturer, Product FROM Win32_BaseBoard")
        if bb:
            rows.append(("主板", "%s %s" % (
                self._clean_vendor(bb[0].get("Manufacturer")),
                safe(lambda: str(bb[0].get("Product")).strip()))))
        else:
            rows.append(("主板", NOT_FOUND))

        # 显示器 (物理优先, 排除远程虚拟镜像)
        reg_mons = self._edid_from_registry()
        active = self._active_monitor_map()
        phys_mons = []
        for m in reg_mons:
            a = active.get(str(m.get("pnp", "")).upper())
            if a and not a[1]:
                continue  # 挂在虚拟适配器上的远程镜像
            phys_mons.append(m)
        phys_mons.sort(key=lambda m: 0 if active.get(str(m.get("pnp", "")).upper()) else 1)
        if phys_mons:
            m = phys_mons[0]
            vendor = self.MONITOR_VENDORS.get(str(m.get("vendor_code", "")).upper(),
                                              m.get("vendor_code") or "")
            name = m.get("name") or "显示器"
            label = " ".join(x for x in (vendor, name) if x)
            size = m.get("size_inch") or 0
            rows.append(("显示器", label + ("（%d 英寸）" % size if size else "")))
        else:
            rows.append(("显示器", NOT_FOUND))

        # 主硬盘
        disks = self.wmi.query("SELECT Model, Size FROM Win32_DiskDrive")
        if disks:
            model = safe(lambda: str(disks[0].get("Model")).strip())
            cap_gb = safe(lambda: int(round(int(disks[0].get("Size") or 0) / 1024.0 ** 3)), 0)
            media = self._physical_disk_media().get(model, "")
            media_short = media.split(" ")[0] if media else ""
            tail = "%d GB" % cap_gb + (" / " + media_short if media_short else "")
            rows.append(("主硬盘", "%s（%s）" % (model, tail)))
        else:
            rows.append(("主硬盘", NOT_FOUND))

        # 网卡 (排除虚拟适配器)
        nics = self.wmi.query("SELECT Description FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
        phys_nics = [n for n in nics
                     if not any(k in str(n.get("Description", "")).lower()
                                for k in self.VIRTUAL_NIC_KEYS)]
        pick = phys_nics or nics
        rows.append(("网卡", safe(lambda: str(pick[0].get("Description")).strip()) if pick else NOT_FOUND))
        return result

    # -- 汇总 --------------------------------------------------------------------
    def collect_all(self):
        """返回 [(分区名, [(项目, 值), ...]), ...] 有序结构。"""
        sections = []
        jobs = [
            ("计算机 / 操作系统", self.collect_system),
            ("处理器 (CPU)", self.collect_cpu),
            ("主板 / BIOS", self.collect_mainboard),
            ("内存", self.collect_memory),
            ("显卡 (GPU)", self.collect_gpu),
            ("硬盘", self.collect_disk),
            ("显示器", self.collect_monitor),
            ("网卡", self.collect_network),
            ("外设", self.collect_peripherals),
        ]
        for title, fn in jobs:
            try:
                sections.append((title, fn()))
            except Exception:
                sections.append((title, [("检测结果", UNKNOWN)]))
        return sections


# ---------------------------------------------------------------------------
# 报告文本生成
# ---------------------------------------------------------------------------

def build_report_text(sections):
    lines = []
    now = datetime.datetime.now()
    lines.append("=" * 56)
    lines.append("硬件检测报告")
    lines.append("生成时间：%s" % now.strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("检测工具：%s v%s" % (APP_NAME, APP_VERSION))
    lines.append("运行权限：%s" % ("管理员" if is_admin() else "普通用户"))
    lines.append("=" * 56)
    for title, items in sections:
        lines.append("")
        lines.append("【%s】" % title)
        lines.append("-" * 56)
        width = max([cjk_width(k) for k, _ in items] + [4])
        for k, v in items:
            pad = " " * (width - cjk_width(k))
            # 值为空的分组标题行去掉行尾冒号
            lines.append(("%s%s：%s" % (k, pad, v)).rstrip("： "))
    lines.append("")
    lines.append("=" * 56)
    lines.append("报告结束")
    return "\n".join(lines)


def build_summary_text(summary):
    """简约概览 -> 可复制的纯文本。"""
    lines = [summary.get("title", "")]
    if summary.get("subtitle"):
        lines.append(summary["subtitle"])
    lines.append("")
    width = max([cjk_width(k) for k, _ in summary["rows"]] + [4])
    for k, v in summary["rows"]:
        pad = " " * (width - cjk_width(k))
        lines.append("%s%s  %s" % (k, pad, v))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        self.title("%s v%s" % (APP_NAME, APP_VERSION))
        self.geometry("760x560")
        self.minsize(640, 480)
        self.sections = []
        self.report_text = ""
        self.summary = None
        self.summary_text = ""
        self._build_ui()
        self.after(100, self.refresh)

    def _build_ui(self):
        top = ttk.Frame(self, padding=(8, 8, 8, 4))
        top.pack(side=tk.TOP, fill=tk.X)

        self.btn_refresh = ttk.Button(top, text="重新检测", command=self.refresh)
        self.btn_refresh.pack(side=tk.LEFT)
        self.btn_copy = ttk.Button(top, text="复制当前页", command=self.copy_all)
        self.btn_copy.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_export = ttk.Button(top, text="导出报告", command=self.export_report)
        self.btn_export.pack(side=tk.LEFT, padx=(8, 0))

        perm = "管理员权限" if is_admin() else "普通权限（如需完整信息可右键以管理员身份运行）"
        ttk.Label(top, text=perm, foreground="#666666").pack(side=tk.RIGHT)

        mid = ttk.Frame(self, padding=(8, 0, 8, 4))
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(mid)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 页签 1: 简约概览
        tab_s = ttk.Frame(self.notebook)
        self.notebook.add(tab_s, text=" 简约概览 ")
        self.txt_summary = tk.Text(tab_s, wrap=tk.NONE, bd=0, padx=20, pady=16,
                                   font=("Microsoft YaHei UI", 11),
                                   state=tk.DISABLED, cursor="arrow")
        ssb = ttk.Scrollbar(tab_s, orient="vertical", command=self.txt_summary.yview)
        self.txt_summary.configure(yscrollcommand=ssb.set)
        self.txt_summary.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ssb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_summary.tag_configure("title", font=("Microsoft YaHei UI", 16, "bold"))
        self.txt_summary.tag_configure("sub", font=("Microsoft YaHei UI", 10),
                                       foreground="#666666")
        self.txt_summary.tag_configure("label", font=("Microsoft YaHei UI", 11, "bold"))

        # 页签 2: 详细参数
        tab_d = ttk.Frame(self.notebook)
        self.notebook.add(tab_d, text=" 详细参数 ")
        self.tree = ttk.Treeview(tab_d, columns=("value",), show="tree headings",
                                 selectmode="browse")
        self.tree.heading("#0", text="项目")
        self.tree.heading("value", text="值")
        self.tree.column("#0", width=260, stretch=False)
        self.tree.column("value", width=460, stretch=True)
        vsb = ttk.Scrollbar(tab_d, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("section", font=("Microsoft YaHei UI", 10, "bold"))

        self.status = tk.StringVar(value="正在检测硬件信息，请稍候…")
        bar = ttk.Label(self, textvariable=self.status, anchor=tk.W,
                        padding=(8, 4), foreground="#444444")
        bar.pack(side=tk.BOTTOM, fill=tk.X)

    def set_busy(self, busy):
        state = tk.DISABLED if busy else tk.NORMAL
        for b in (self.btn_refresh, self.btn_copy, self.btn_export):
            b.configure(state=state)

    def refresh(self):
        self.set_busy(True)
        self.status.set("正在检测硬件信息，请稍候…")
        self.tree.delete(*self.tree.get_children())

        def work():
            collector = HardwareCollector()
            sections = collector.collect_all()
            report = build_report_text(sections)
            summary = collector.collect_summary()
            summary_text = build_summary_text(summary)
            wmi_err = collector.wmi_error
            self.after(0, lambda: self._fill(sections, report, summary, summary_text, wmi_err))

        threading.Thread(target=work, daemon=True).start()

    def _fill(self, sections, report, summary, summary_text, wmi_err):
        self.sections = sections
        self.report_text = report
        self.summary = summary
        self.summary_text = summary_text
        # 简约概览页
        self.txt_summary.configure(state=tk.NORMAL)
        self.txt_summary.delete("1.0", tk.END)
        self.txt_summary.insert(tk.END, summary.get("title", "") + "\n", "title")
        self.txt_summary.insert(tk.END, summary.get("subtitle", "") + "\n\n", "sub")
        width = max([cjk_width(k) for k, _ in summary["rows"]] + [4])
        for k, v in summary["rows"]:
            pad = " " * (width - cjk_width(k))
            self.txt_summary.insert(tk.END, k + pad + "  ", "label")
            self.txt_summary.insert(tk.END, v + "\n\n")
        self.txt_summary.configure(state=tk.DISABLED)
        # 详细参数页
        for title, items in sections:
            node = self.tree.insert("", tk.END, text=title, values=("",),
                                    open=True, tags=("section",))
            for k, v in items:
                self.tree.insert(node, tk.END, text=k, values=(v,))
        if wmi_err:
            self.status.set("检测完成（WMI 组件不可用：%s，结果可能不完整）" % wmi_err)
        else:
            self.status.set("检测完成：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.set_busy(False)

    def copy_all(self):
        """复制当前页签内容: 简约概览 或 详细报告。"""
        try:
            on_summary = self.notebook.index(self.notebook.select()) == 0
        except Exception:
            on_summary = False
        text = self.summary_text if on_summary else self.report_text
        what = "简约概览" if on_summary else "详细报告"
        if not text:
            messagebox.showinfo(APP_NAME, "暂无检测结果，请先完成检测。")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.set("%s内容已复制到剪贴板。" % what)

    def export_report(self):
        if not self.report_text:
            messagebox.showinfo(APP_NAME, "暂无检测结果，请先完成检测。")
            return
        hostname = safe(lambda: socket.gethostname(), "PC")
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = "硬件检测报告_%s_%s.txt" % (hostname, stamp)
        path = filedialog.asksaveasfilename(
            title="导出检测报告",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("文本文件", "*.txt")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write(self.report_text)
            self.status.set("报告已导出：%s" % path)
            messagebox.showinfo(APP_NAME, "报告已导出：\n%s" % path)
        except Exception as e:
            messagebox.showerror(APP_NAME, "导出失败：%s" % e)


if __name__ == "__main__":
    # Windows 高 DPI 适配 (失败不影响运行)
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    try:
        App().mainloop()
    except Exception:
        # 最后兜底: 任何未预期错误写入 exe 同目录日志, 不弹崩溃窗
        try:
            log = os.path.join(os.path.dirname(os.path.abspath(
                getattr(sys, "_MEIPASS", __file__))), "hwinfo_error.log")
            with open(log, "a", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass

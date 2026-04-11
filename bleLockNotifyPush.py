# coding=utf-8

import asyncio, ctypes, gc, httpx, os, subprocess, sys, threading, time
from ctypes import wintypes
from bleak import BleakScanner

import winrt.windows.ui.notifications.management as mgmt
import winrt.windows.ui.notifications as notifications


# Windows 蓝牙自动锁屏（感应钥匙） 通知转发
# 装依赖 pip install httpx bleak winrt-Windows.UI.Notifications winrt-Windows.UI.Notifications.Management winrt-Windows.Data.Xml.Dom
# zyyme 20260327


# 手环或手机的蓝牙mac地址 为空时会扫描并输出所有设备
TARGET_MAC = "AA:BB:CC:DD:EE:FF"
# 低于这个信号强度会自动锁屏
RSSI_THRESHOLD = -85
# 检查间隔
CHECK_INTERVAL = 60
# 通知转发的webhook ios可以用bark 两个{}是标题和内容的占位符 留空则不推送
WEBHOOK_URL = "https://api.day.app//{}/{}"
# 筛选应用名称 为空则全推送
FILTER_APP_NAMES = ["微信", "企业微信"]
# 检查闪烁提醒的进程名 为空则不监听
FLASH_WATCH_PROCESS_NAMES = ["Weixin.exe", "WXWork.exe"]

# 闪烁提醒通知冷却时间 秒
FLASH_NOTIFY_COOLDOWN_SECONDS = 10
# 本地 Windows Toast 通知使用的 AppID
TOAST_APP_ID = "bleLockNotifyPush"
# 锁屏时如果检测到媒体正在播放，则先发送一次播放/暂停键
PAUSE_MEDIA_ON_LOCK = True

# 打包用 参数传入
# TARGET_MAC = sys.argv[1] if len(sys.argv) > 1 else ""
# RSSI_THRESHOLD = int(sys.argv[2]) if len(sys.argv) > 2 else -80
# CHECK_INTERVAL = int(sys.argv[3]) if len(sys.argv) > 3 else 60
# WEBHOOK_URL = sys.argv[4] if len(sys.argv) > 4 else ""
# FILTER_APP_NAMES = sys.argv[5].split(',') if len(sys.argv) > 5 else []
# FLASH_WATCH_PROCESS_NAMES = sys.argv[6].split(',') if len(sys.argv) > 6 else []


# 实时 RSSI
current_device_rssi = None
last_seen_time = 0
flash_last_notify_time = {}
POWERSHELL_EXE = os.path.join(
    os.environ.get("SystemRoot", r"C:\Windows"),
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "powershell.exe",
)

MEDIA_STATUS_CORE_AUDIO_POWERSHELL_SCRIPT = r"""
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace AudioSessionProbe {
    public enum EDataFlow { eRender = 0, eCapture = 1, eAll = 2 }
    public enum ERole { eConsole = 0, eMultimedia = 1, eCommunications = 2 }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumeratorComObject {}

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator {
        int EnumAudioEndpoints(EDataFlow dataFlow, uint dwStateMask, out object ppDevices);
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice {
        int Activate(ref Guid iid, uint dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    }

    [ComImport, Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioMeterInformation {
        int GetPeakValue(out float pfPeak);
        int GetMeteringChannelCount(out int pnChannelCount);
        int GetChannelsPeakValues(int u32ChannelCount, IntPtr afPeakValues);
        int QueryHardwareSupport(out int pdwHardwareSupportMask);
    }

    public static class Detector {
        private const uint CLSCTX_ALL = 23;

        public static float GetRenderPeakMax(int sampleCount, int sampleDelayMs) {
            var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
            IMMDevice device;
            if (enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device) != 0 || device == null) {
                return 0f;
            }

            Guid iid = typeof(IAudioMeterInformation).GUID;
            object meterObj;
            if (device.Activate(ref iid, CLSCTX_ALL, IntPtr.Zero, out meterObj) != 0 || meterObj == null) {
                return 0f;
            }

            var meter = (IAudioMeterInformation)meterObj;
            float peakMax = 0f;
            for (int i = 0; i < sampleCount; i++) {
                float peak;
                if (meter.GetPeakValue(out peak) == 0 && peak > peakMax) {
                    peakMax = peak;
                }
                if (i < sampleCount - 1 && sampleDelayMs > 0) {
                    Thread.Sleep(sampleDelayMs);
                }
            }
            return peakMax;
        }
    }
}
"@ -ErrorAction Stop

    $peakThreshold = 0.0015
    $peakMax = [AudioSessionProbe.Detector]::GetRenderPeakMax(6, 70)

    Write-Output ([string]::Format("{0:F6}", $peakMax))

    if ($peakMax -ge $peakThreshold) {
        Write-Output 'PLAYING'
    }
    else {
        Write-Output 'NOT_PLAYING'
    }
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
"""


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_EXTENDEDKEY = 0x0001
VK_F14 = 0x7D
VK_MEDIA_PLAY_PAUSE = 0xB3
HSHELL_FLASH = 0x8006
GA_ROOTOWNER = 3
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WM_CLOSE = 0x0010
WM_DESTROY = 0x0002
LRESULT = ctypes.c_ssize_t
PUL = ctypes.POINTER(ctypes.c_ulong)
WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
)


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", PUL),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", PUL),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.c_ulong),
        ("wParamL", ctypes.c_short),
        ("wParamH", ctypes.c_ushort),
    ]


class INPUTUNION(ctypes.Union):
    _fields_ = [
        ("ki", KEYBDINPUT),
        ("mi", MOUSEINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("ii", INPUTUNION),
    ]


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HANDLE),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HANDLE),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
user32.SendInput.restype = wintypes.UINT
user32.LockWorkStation.argtypes = []
user32.LockWorkStation.restype = wintypes.BOOL
user32.DefWindowProcW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.DefWindowProcW.restype = LRESULT
user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = wintypes.ATOM
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD,
    wintypes.LPCWSTR,
    wintypes.LPCWSTR,
    wintypes.DWORD,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HWND,
    wintypes.HMENU,
    wintypes.HINSTANCE,
    wintypes.LPVOID,
]
user32.CreateWindowExW.restype = wintypes.HWND
user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL
user32.UnregisterClassW.argtypes = [wintypes.LPCWSTR, wintypes.HINSTANCE]
user32.UnregisterClassW.restype = wintypes.BOOL
user32.RegisterWindowMessageW.argtypes = [wintypes.LPCWSTR]
user32.RegisterWindowMessageW.restype = wintypes.UINT
user32.RegisterShellHookWindow.argtypes = [wintypes.HWND]
user32.RegisterShellHookWindow.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [
    ctypes.POINTER(wintypes.MSG),
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT,
]
user32.GetMessageW.restype = ctypes.c_int
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.TranslateMessage.restype = wintypes.BOOL
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.restype = LRESULT
user32.PostMessageW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.PostMessageW.restype = wintypes.BOOL
user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.PostQuitMessage.restype = None
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND
user32.GetWindowThreadProcessId.argtypes = [
    wintypes.HWND,
    ctypes.POINTER(wintypes.DWORD),
]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int

kernel32.OpenProcess.argtypes = [
    wintypes.DWORD,
    wintypes.BOOL,
    wintypes.DWORD,
]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE


def send_virtual_key(vk_code, is_extended=False):
    """发送一次指定的虚拟按键（按下 + 抬起）。"""
    key_down_flags = KEYEVENTF_EXTENDEDKEY if is_extended else 0
    key_up_flags = KEYEVENTF_KEYUP | key_down_flags
    extra = ctypes.c_ulong(0)
    key_down = INPUT(
        type=INPUT_KEYBOARD,
        ii=INPUTUNION(
            ki=KEYBDINPUT(
                wVk=vk_code,
                dwFlags=key_down_flags,
                dwExtraInfo=ctypes.pointer(extra),
            )
        ),
    )
    key_up = INPUT(
        type=INPUT_KEYBOARD,
        ii=INPUTUNION(
            ki=KEYBDINPUT(
                wVk=vk_code,
                dwFlags=key_up_flags,
                dwExtraInfo=ctypes.pointer(extra),
            )
        ),
    )
    sent_down = user32.SendInput(1, ctypes.byref(key_down), ctypes.sizeof(INPUT))
    sent_up = user32.SendInput(1, ctypes.byref(key_up), ctypes.sizeof(INPUT))
    if sent_down != 1 or sent_up != 1:
        raise ctypes.WinError(ctypes.get_last_error())


def query_media_playback_status():
    """使用 CoreAudio 检查系统媒体状态，返回是否播放bool。"""
    try:
        powershell_exe = POWERSHELL_EXE if os.path.exists(POWERSHELL_EXE) else "powershell"
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            [
                powershell_exe,
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                MEDIA_STATUS_CORE_AUDIO_POWERSHELL_SCRIPT,
            ],
            capture_output=True,
            text=True,
            timeout=5,
            creationflags=creationflags,
        )
        if result.returncode != 0:
            stderr = result.stderr.strip() or result.stdout or "未知错误"
            print(f"CoreAudio 媒体状态检查失败: {stderr}")
            return False

        print(f"CoreAudio: {result.stdout.strip()}")

        return "NOT_PLAYING" not in result.stdout
    except Exception as e:
        print(f"CoreAudio 媒体状态检查异常: {e}")
        return False


async def pause_media_if_playing():
    """锁屏前如检测到媒体正在播放，则发送一次播放/暂停键。"""
    if not PAUSE_MEDIA_ON_LOCK:
        return

    media_status = await asyncio.to_thread(query_media_playback_status)
    if media_status:
        try:
            # 媒体按键属于扩展键，需带 KEYEVENTF_EXTENDEDKEY 才能在部分设备上生效。
            send_virtual_key(VK_MEDIA_PLAY_PAUSE, is_extended=True)
            print("检测到媒体正在播放，已发送播放/暂停按键")
            await asyncio.sleep(1)
            media_status = await asyncio.to_thread(query_media_playback_status)
            if media_status:
                print("还有声音，再按一次")
                send_virtual_key(VK_MEDIA_PLAY_PAUSE, is_extended=True)
        except Exception as e:
            print(f"暂停媒体失败: {e}")




async def lock_workstation_with_media_pause():
    """锁屏前先尝试暂停正在播放的媒体。"""
    await pause_media_if_playing()
    if not user32.LockWorkStation():
        print(f"锁屏失败: {ctypes.WinError(ctypes.get_last_error())}")


def wake_screen():
    """模拟一次 F14 按键，尝试点亮已熄灭的显示器。"""
    try:
        send_virtual_key(VK_F14)
    except Exception as e:
        print(f"亮屏失败: {e}")


def format_process_name(process_name):
    return os.path.splitext(os.path.basename(process_name))[0] or process_name


def get_watched_process_name(process_name):
    if not process_name:
        return None
    for target in FLASH_WATCH_PROCESS_NAMES:
        if process_name.lower() == target.lower():
            return target
    return None


def get_window_title(hwnd):
    if not hwnd:
        return ""
    length = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(max(length + 1, 256))
    user32.GetWindowTextW(hwnd, buf, len(buf))
    return buf.value.strip()


def get_process_name_for_hwnd(hwnd):
    if not hwnd:
        return None, None

    root_hwnd = user32.GetAncestor(hwnd, GA_ROOTOWNER) or hwnd
    process_id = wintypes.DWORD()
    if not user32.GetWindowThreadProcessId(root_hwnd, ctypes.byref(process_id)):
        return None, root_hwnd
    if not process_id.value:
        return None, root_hwnd

    process_handle = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION,
        False,
        process_id.value,
    )
    if not process_handle:
        return None, root_hwnd

    try:
        size = wintypes.DWORD(1024)
        path_buf = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(
            process_handle,
            0,
            path_buf,
            ctypes.byref(size),
        ):
            return None, root_hwnd
        return os.path.basename(path_buf.value), root_hwnd
    finally:
        kernel32.CloseHandle(process_handle)


def resolve_flash_event(hwnd):
    process_name, root_hwnd = get_process_name_for_hwnd(hwnd)
    watched_process_name = get_watched_process_name(process_name)
    if not watched_process_name:
        return None

    body = get_window_title(root_hwnd) or "检测到闪烁提醒"
    return watched_process_name, body


def send_local_toast(title, body):
    """发送本地 Windows Toast 通知。"""
    try:
        toast_xml = notifications.ToastNotificationManager.get_template_content(
            notifications.ToastTemplateType.TOAST_TEXT02
        )
        text_nodes = toast_xml.get_elements_by_tag_name("text")
        if text_nodes.length < 2:
            raise RuntimeError("Toast 模板缺少文本节点")

        text_nodes.item(0).append_child(toast_xml.create_text_node(title))
        text_nodes.item(1).append_child(toast_xml.create_text_node(body))

        create_with_id = getattr(
            notifications.ToastNotificationManager,
            "create_toast_notifier_with_id",
            None,
        )
        if create_with_id:
            notifier = create_with_id(TOAST_APP_ID)
        else:
            notifier = notifications.ToastNotificationManager.create_toast_notifier(
                TOAST_APP_ID
            )
        notifier.show(notifications.ToastNotification(toast_xml))
    except Exception as e:
        print(f"本地通知失败: {e}")


def get_notification_texts(notification):
    """提取通知里的文本内容，取不到时返回空列表。"""
    visual = None
    bindings = None
    binding = None
    text_elements = None
    try:
        visual = notification.visual
        bindings = visual.bindings
        if not bindings or len(bindings) == 0:
            return []
        binding = bindings[0]
        text_elements = binding.get_text_elements()
        return [item.text for item in text_elements]
    except Exception:
        return []
    finally:
        text_elements = None
        binding = None
        bindings = None
        visual = None


async def get_next_notification_text_snapshot(listener, processed_ids):
    """每次只提取一条通知的正文文本，避免同批 WinRT 文本对象一起释放时崩溃。"""
    notifs = await listener.get_notifications_async(notifications.NotificationKinds.TOAST)
    try:
        for user_notification in notifs:
            notification_id = user_notification.id
            if notification_id in processed_ids:
                user_notification = None
                continue
            return {
                "id": notification_id,
                "texts": get_notification_texts(user_notification.notification),
            }
        return None
    finally:
        notifs = None
        gc.collect()


async def get_notification_app_name(listener, notification_id):
    """在新的 WinRT 查询里读取 app_info，避免和正文文本对象混用。"""
    notifs = await listener.get_notifications_async(notifications.NotificationKinds.TOAST)
    try:
        for user_notification in notifs:
            if user_notification.id != notification_id:
                user_notification = None
                continue
            try:
                return user_notification.app_info.display_info.display_name
            except Exception:
                return None
        return None
    finally:
        notifs = None
        gc.collect()


async def send_webhook(app, title, content):
    print("转发通知:", title, content)
    # payload = {"app": app, "title": title, "content": content}
    try:
        async with httpx.AsyncClient() as client:
            await client.get(WEBHOOK_URL.format(app + ":" + title + time.strftime("(%I:%M)", time.localtime()), content), timeout=20)
    except Exception as e:
        print(f"Webhook 失败: {e}")


async def handle_flash_event(process_name, body):
    try:
        now = time.monotonic()
        last_notify_time = flash_last_notify_time.get(process_name, 0)
        if now - last_notify_time < FLASH_NOTIFY_COOLDOWN_SECONDS:
            return
        flash_last_notify_time[process_name] = now

        display_name = format_process_name(process_name)
        title = f"{display_name} 有新提醒"
        body = body or "检测到闪烁提醒"
        print("检测到闪烁提醒:", process_name, body)
        send_local_toast(title, body)
        if WEBHOOK_URL:
            await send_webhook(display_name, "新提醒", body)
    except Exception as e:
        print(f"闪烁提醒处理异常: {e}")


async def monitor_notifications():
    """监听并推送 Windows 通知"""
    try:
        listener = mgmt.UserNotificationListener.current
        access = await listener.request_access_async()
        if access != mgmt.UserNotificationListenerAccessStatus.ALLOWED:
            print("错误：未获得通知访问权限。")
            return

        print("正在监听通知:", ", ".join(FILTER_APP_NAMES) if FILTER_APP_NAMES else "全部应用")
        print("推送到:", WEBHOOK_URL)
        processed_ids = set()
        startup_notifs = await listener.get_notifications_async(notifications.NotificationKinds.TOAST)
        try:
            for user_notification in startup_notifs:
                processed_ids.add(user_notification.id)
            print(f"启动时已忽略现有通知: {len(processed_ids)} 条")
        finally:
            startup_notifs = None
            gc.collect()

        while True:
            try:
                while True:
                    item = await get_next_notification_text_snapshot(listener, processed_ids)
                    if not item:
                        break

                    notification_id = item["id"]
                    app_name = await get_notification_app_name(listener, notification_id)
                    if not app_name:
                        # 程序发送的通知没有appinfo
                        processed_ids.add(notification_id)
                        continue

                    print("通知应用:", app_name)
                    if not FILTER_APP_NAMES or any(target in app_name for target in FILTER_APP_NAMES):
                        texts = item["texts"]
                        t = texts[0] if len(texts) > 0 else ""
                        c = texts[1] if len(texts) > 1 else ""
                        await send_webhook(app_name, t, c)

                    processed_ids.add(notification_id)

                    # 文本和 app_info 分两次查询，每条之间稍作让步。
                    await asyncio.sleep(0.2)

                if len(processed_ids) > 300:
                    current_notifs = await listener.get_notifications_async(
                        notifications.NotificationKinds.TOAST
                    )
                    try:
                        processed_ids = {user_notification.id for user_notification in current_notifs}
                    finally:
                        current_notifs = None
                        gc.collect()
            except Exception as e:
                print(f"获取通知异常: {e}")
            await asyncio.sleep(5)
    except Exception as e:
        print(f"通知监控异常: {e}")


async def monitor_shell_flash(loop):
    """监听指定程序的窗口闪烁提醒。"""
    if not FLASH_WATCH_PROCESS_NAMES:
        return

    print("正在监听闪烁提醒:", ", ".join(FLASH_WATCH_PROCESS_NAMES))
    ready_event = threading.Event()
    state = {"error": None, "hwnd": None, "thread": None, "wndproc": None}

    def process_flash_message(hwnd):
        payload = resolve_flash_event(hwnd)
        if not payload:
            return
        process_name, body = payload
        asyncio.create_task(handle_flash_event(process_name, body))

    def flash_thread_target():
        shellhook_message_id = 0
        class_name = f"BleLockNotifyPushShellHook_{os.getpid()}"
        hinstance = kernel32.GetModuleHandleW(None)
        if not hinstance:
            state["error"] = f"GetModuleHandleW 失败: {ctypes.WinError(ctypes.get_last_error())}"
            ready_event.set()
            return

        @WNDPROC
        def window_proc(hwnd, msg, wparam, lparam):
            if msg == shellhook_message_id and int(wparam) == HSHELL_FLASH:
                hwnd_value = int(lparam)
                if not loop.is_closed():
                    try:
                        loop.call_soon_threadsafe(process_flash_message, hwnd_value)
                    except RuntimeError:
                        pass
                return 0
            if msg == WM_CLOSE:
                user32.DestroyWindow(hwnd)
                return 0
            if msg == WM_DESTROY:
                user32.PostQuitMessage(0)
                return 0
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        state["wndproc"] = window_proc
        wnd_class = WNDCLASSW()
        wnd_class.lpfnWndProc = window_proc
        wnd_class.hInstance = hinstance
        wnd_class.lpszClassName = class_name

        atom = user32.RegisterClassW(ctypes.byref(wnd_class))
        if not atom:
            state["error"] = f"RegisterClassW 失败: {ctypes.WinError(ctypes.get_last_error())}"
            ready_event.set()
            return

        hwnd = user32.CreateWindowExW(
            0,
            class_name,
            class_name,
            0,
            0,
            0,
            0,
            0,
            None,
            None,
            hinstance,
            None,
        )
        if not hwnd:
            state["error"] = f"CreateWindowExW 失败: {ctypes.WinError(ctypes.get_last_error())}"
            user32.UnregisterClassW(class_name, hinstance)
            ready_event.set()
            return

        shellhook_message_id = user32.RegisterWindowMessageW("SHELLHOOK")
        if not shellhook_message_id:
            state["error"] = f"RegisterWindowMessageW 失败: {ctypes.WinError(ctypes.get_last_error())}"
            user32.DestroyWindow(hwnd)
            user32.UnregisterClassW(class_name, hinstance)
            ready_event.set()
            return

        if not user32.RegisterShellHookWindow(hwnd):
            state["error"] = f"RegisterShellHookWindow 失败: {ctypes.WinError(ctypes.get_last_error())}"
            user32.DestroyWindow(hwnd)
            user32.UnregisterClassW(class_name, hinstance)
            ready_event.set()
            return

        state["hwnd"] = hwnd
        state["class_name"] = class_name
        state["hinstance"] = hinstance
        ready_event.set()

        msg = wintypes.MSG()
        result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        while result > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
            result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)

        if result == -1:
            print("闪烁提醒消息循环异常")

        state["hwnd"] = None
        user32.UnregisterClassW(class_name, hinstance)

    flash_thread = threading.Thread(target=flash_thread_target, daemon=True)
    state["thread"] = flash_thread
    flash_thread.start()
    await asyncio.to_thread(ready_event.wait)

    if state["error"]:
        print(f"闪烁提醒监控启动失败: {state['error']}")
        return

    try:
        await asyncio.Future()
    except asyncio.CancelledError:
        if state["hwnd"]:
            user32.PostMessageW(state["hwnd"], WM_CLOSE, 0, 0)
        await asyncio.to_thread(flash_thread.join, 2)
        raise


def detection_callback(device, advertisement_data):
    """蓝牙回调：实时更新目标设备的 RSSI"""
    global current_device_rssi, last_seen_time
    if device.address.upper() == TARGET_MAC.upper():
        current_device_rssi = advertisement_data.rssi
        last_seen_time = time.monotonic()


async def monitor_ble():
    """逻辑判定：根据回调数据决定是否锁屏"""
    global current_device_rssi
    print(f"正在监听设备: {TARGET_MAC}")

    # 启动持续扫描
    scanner = None
    try:
        scanner = BleakScanner(detection_callback=detection_callback)
        await scanner.start()

        while True:
            await asyncio.sleep(CHECK_INTERVAL)

            now = time.monotonic()
            # 判定 1: 如果超时没收到广播包，视为离开
            if now - last_seen_time > CHECK_INTERVAL:
                if current_device_rssi is None:
                    print("找不到设备")
                else:
                    print("找不到设备 -> 锁屏")
                    current_device_rssi = None
                    await lock_workstation_with_media_pause()
            # 判定 2: 收到信号但太弱
            elif current_device_rssi is not None and current_device_rssi < RSSI_THRESHOLD:
                print(f"信号太弱 ({current_device_rssi} dBm) -> 锁屏")
                current_device_rssi = None
                await lock_workstation_with_media_pause()
            else:
                print(f"当前信号强度:{current_device_rssi} dBm")
                wake_screen()
    finally:
        if scanner is not None:
            print("扫描停止")
            try:
                await scanner.stop()
            except Exception as e:
                print(f"停止扫描失败: {e}")


async def scan_and_list_devices():
    """扫描模式：列出所有设备及其 RSSI"""
    try:
        print("正在搜寻周围的 BLE 设备 (10s)...")
        # return_adv=True 确保拿到 AdvertisementData 对象
        devices_dict = await BleakScanner.discover(timeout=10.0, return_adv=True)

        print("\n" + "=" * 70)
        print(f"{'名称':<30} | {'MAC 地址':<20} | RSSI")
        print("-" * 70)
        for addr, (device, adv) in devices_dict.items():
            name = device.name if device.name else "Unknown"
            print(f"{name:<30} | {addr:<20} | {adv.rssi} dBm")
        print("=" * 70 + "\n")
    except Exception as e:
        print(f"扫描异常: {e}")


async def main():
    tasks = []
    if not TARGET_MAC:
        await scan_and_list_devices()
    else:
        tasks.append(monitor_ble())
    if FLASH_WATCH_PROCESS_NAMES:
        loop = asyncio.get_running_loop()
        tasks.append(monitor_shell_flash(loop))
    if WEBHOOK_URL:
        tasks.append(monitor_notifications())
    if tasks:
        await asyncio.gather(*tasks)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n退出")

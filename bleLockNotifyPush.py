# coding=utf-8

import asyncio, ctypes, httpx, sys
from bleak import BleakScanner

import winrt.windows.ui.notifications.management as mgmt
import winrt.windows.ui.notifications as notifications


# Windows 蓝牙自动锁屏（感应钥匙） 通知转发
# 装依赖 pip install httpx bleak winrt-Windows.UI.Notifications winrt-Windows.UI.Notifications.Management
# zyyme 20260327


# 手环或手机的蓝牙mac地址 为空时会扫描并输出所有设备
TARGET_MAC = ""
# 低于这个信号强度会自动锁屏
RSSI_THRESHOLD = -75
# 检查间隔
CHECK_INTERVAL = 60
# 通知转发的webhook ios可以用bark 两个{}是标题和内容的占位符 留空则不推送
WEBHOOK_URL = "https://api.day.app/abcd/{}/{}"
# 筛选应用名称 为空则全推送
FILTER_APP_NAMES = ["微信", "企业微信"]


# 打包用 参数传入
# TARGET_MAC = sys.argv[1] if len(sys.argv) > 1 else ""
# RSSI_THRESHOLD = int(sys.argv[2]) if len(sys.argv) > 2 else -80
# CHECK_INTERVAL = int(sys.argv[3]) if len(sys.argv) > 3 else 60
# WEBHOOK_URL = sys.argv[4] if len(sys.argv) > 4 else ""
# FILTER_APP_NAMES = sys.argv[5].split(',') if len(sys.argv) > 5 else []

# 实时 RSSI
current_device_rssi = None
last_seen_time = 0

async def send_webhook(app, title, content):
    print("转发通知:", title, content)
    # payload = {"app": app, "title": title, "content": content}
    try:
        async with httpx.AsyncClient() as client:
            await client.get(WEBHOOK_URL.format(app  + ":" + title, content), timeout=20)
    except Exception as e:
        print(f"Webhook 失败: {e}")

async def monitor_notifications():
    """监听并推送 Windows 通知"""
    try:
        listener = mgmt.UserNotificationListener.current
        access = await listener.request_access_async()
        if access != mgmt.UserNotificationListenerAccessStatus.ALLOWED:
            print("错误：未获得通知访问权限。")
            return

        processed_ids = set()
        while True:
            notifs = await listener.get_notifications_async(notifications.NotificationKinds.TOAST)
            for n in notifs:
                if n.id not in processed_ids:
                    app_name = n.app_info.display_info.display_name
                    print("通知应用:",app_name)
                    if not FILTER_APP_NAMES or any(target in app_name for target in FILTER_APP_NAMES):
                        visual = n.notification.visual
                        bindings = visual.bindings
                        if bindings and len(bindings) > 0:
                            texts = bindings[0].get_text_elements()
                            t = texts[0].text if len(texts) > 0 else ""
                            c = texts[1].text if len(texts) > 1 else ""
                            await send_webhook(app_name, t, c)
                    processed_ids.add(n.id)
            if len(processed_ids) > 100: processed_ids.clear()
            await asyncio.sleep(2)
    except Exception as e:
        print(f"通知监控异常: {e}")

def detection_callback(device, advertisement_data):
    """蓝牙回调：实时更新目标设备的 RSSI"""
    global current_device_rssi, last_seen_time
    if device.address.upper() == TARGET_MAC.upper():
        current_device_rssi = advertisement_data.rssi
        last_seen_time = asyncio.get_event_loop().time()

async def monitor_ble():
    """逻辑判定：根据回调数据决定是否锁屏"""
    global current_device_rssi
    print(f"正在监听设备: {TARGET_MAC}")
    
    # 启动持续扫描
    scanner = BleakScanner(detection_callback=detection_callback)
    await scanner.start()
    
    try:
        while True:
            await asyncio.sleep(CHECK_INTERVAL)

            now = asyncio.get_event_loop().time()
            # 判定 1: 如果超时没收到广播包，视为离开
            if now - last_seen_time > CHECK_INTERVAL:
                print("找不到设备 -> 锁屏")
                ctypes.windll.user32.LockWorkStation()
                current_device_rssi = None # 重置
                # 离开时翻倍延迟
                await asyncio.sleep(CHECK_INTERVAL)
            # 判定 2: 收到信号但太弱
            elif current_device_rssi is not None and current_device_rssi < RSSI_THRESHOLD:
                print(f"信号太弱 ({current_device_rssi} dBm) -> 锁屏")
                ctypes.windll.user32.LockWorkStation()
            else:
                print(f"当前信号强度:{current_device_rssi} dBm")
    finally:
        print("扫描停止")
        await scanner.stop()

async def scan_and_list_devices():
    """扫描模式：列出所有设备及其 RSSI"""
    print("正在搜寻周围的 BLE 设备 (10s)...")
    # return_adv=True 确保拿到 AdvertisementData 对象
    devices_dict = await BleakScanner.discover(timeout=10.0, return_adv=True)
    
    print("\n" + "="*70)
    print(f"{'名称':<30} | {'MAC 地址':<20} | RSSI")
    print("-" * 70)
    for addr, (device, adv) in devices_dict.items():
        name = device.name if device.name else "Unknown"
        print(f"{name:<30} | {addr:<20} | {adv.rssi} dBm")
    print("="*70 + "\n")
    sys.exit(0)

async def main():
    if not TARGET_MAC:
        await scan_and_list_devices()
    else:
        # await send_webhook("应用名称", "标题", "内容")
        if WEBHOOK_URL:
            await asyncio.gather(
                monitor_ble(),
                monitor_notifications()
            )
        else:
            await monitor_ble()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n退出")

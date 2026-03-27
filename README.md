# Windows 蓝牙自动锁屏（感应钥匙） 通知转发

需要Python3  
安装依赖
```
pip install httpx bleak winrt-Windows.UI.Notifications winrt-Windows.UI.Notifications.Management
```

使用文本编辑器打开文件，顶部的注释会教你怎么配置

测试可用后，可以将后缀改成pyw（后台运行，不会显示任何界面），放到启动文件夹或是配置到“任务计划程序”中作为开机启动

注意信号强度是一个负数，越接近0信号越强，如果频繁提示“找不到设备”，可以加大CHECK_INTERVAL，因为设备的广播间隔比较大

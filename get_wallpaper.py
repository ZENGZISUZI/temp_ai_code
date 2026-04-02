import win32api
import win32con
import win32gui

# 从注册表获取壁纸路径
import ctypes

SPI_GETDESKWALLPAPER = 0x0073
buffer = ctypes.create_unicode_buffer(260)
ctypes.windll.user32.SystemParametersInfoW(SPI_GETDESKWALLPAPER, 260, buffer, 0)
print(buffer.value)

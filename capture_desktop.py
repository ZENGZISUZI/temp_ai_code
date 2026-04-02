import pyautogui
from datetime import datetime
import os

# 截取整个屏幕
screenshot = pyautogui.screenshot()

# 保存
save_path = r"D:\AI\current_wallpaper.png"
screenshot.save(save_path)
print(f"截图已保存: {save_path}")
print(f"尺寸: {screenshot.size[0]}x{screenshot.size[1]}")

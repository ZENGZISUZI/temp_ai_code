import pyautogui
import os
from datetime import datetime

save_dir = r'D:\AI\screenshots'
os.makedirs(save_dir, exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
path = os.path.join(save_dir, f'screenshot_{timestamp}.png')

pyautogui.screenshot().save(path)
print(f'截图完成: {path}')

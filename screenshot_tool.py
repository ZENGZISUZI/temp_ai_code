"""
电脑截屏工具 v2.0
作者：狗腿子 🐕
功能：全屏截图、区域截图、选择窗口截图、保存到文件
"""

import pyautogui
import pygetwindow as gw
from PIL import Image
import os
from datetime import datetime


def take_screenshot(save_path=None, region=None):
    """
    截取屏幕
    
    Args:
        save_path: 保存路径
        region: 截图区域 (left, top, width, height)，None为全屏
    
    Returns:
        PIL.Image 对象
    """
    if region:
        screenshot = pyautogui.screenshot(region=region)
    else:
        screenshot = pyautogui.screenshot()
    
    if save_path:
        screenshot.save(save_path)
        print(f"截图已保存: {save_path}")
    
    return screenshot


def get_screen_size():
    """获取屏幕尺寸"""
    return pyautogui.size()


def list_windows():
    """列出所有可见窗口"""
    windows = gw.getAllWindows()
    visible_windows = []
    for i, win in enumerate(windows):
        if win.visible and win.title:
            visible_windows.append({
                'index': i,
                'title': win.title,
                'window': win
            })
    return visible_windows


def capture_window(window, save_path=None):
    """
    截取指定窗口
    
    Args:
        window: pygetwindow 窗口对象
        save_path: 保存路径
    
    Returns:
        PIL.Image 对象
    """
    # 获取窗口位置和大小
    left = window.left
    top = window.top
    width = window.width
    height = window.height
    
    # 截取窗口区域
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    
    if save_path:
        screenshot.save(save_path)
        print(f"截图已保存: {save_path}")
    
    return screenshot


def main():
    # 配置区域
    save_dir = r"D:\AI\screenshots"  # 截图保存目录
    filename_format = "screenshot_{time}.png"  # 文件名格式
    
    # 创建保存目录
    os.makedirs(save_dir, exist_ok=True)
    
    # 生成文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = filename_format.format(time=timestamp)
    save_path = os.path.join(save_dir, filename)
    
    print("=" * 50)
    print("截屏工具 v2.0 🐕")
    print("=" * 50)
    print("请选择截图模式：")
    print("1. 全屏截图")
    print("2. 选择窗口截图")
    print("0. 退出")
    print("-" * 50)
    
    choice = input("请输入选项 (0/1/2): ").strip()
    
    if choice == "0":
        print("已退出")
        return
    
    elif choice == "1":
        # 全屏截图
        screen_width, screen_height = get_screen_size()
        print(f"\n屏幕尺寸: {screen_width} x {screen_height}")
        print("正在截屏...")
        screenshot = take_screenshot(save_path)
        print(f"截图完成！尺寸: {screenshot.size[0]} x {screenshot.size[1]}")
        print(f"保存位置: {save_path}")
    
    elif choice == "2":
        # 选择窗口截图
        print("\n正在获取窗口列表...")
        windows = list_windows()
        
        if not windows:
            print("未找到可见窗口")
            return
        
        print("\n可用窗口列表：")
        print("-" * 50)
        for win in windows:
            print(f"[{win['index']}] {win['title']}")
        print("-" * 50)
        
        try:
            window_index = int(input("请输入窗口编号: ").strip())
            
            # 查找选中的窗口
            selected_window = None
            for win in windows:
                if win['index'] == window_index:
                    selected_window = win['window']
                    break
            
            if selected_window:
                # 激活窗口并截图
                selected_window.activate()
                print(f"\n正在截取窗口: {selected_window.title}")
                screenshot = capture_window(selected_window, save_path)
                print(f"截图完成！尺寸: {screenshot.size[0]} x {screenshot.size[1]}")
                print(f"保存位置: {save_path}")
            else:
                print("无效的窗口编号")
                
        except ValueError:
            print("请输入有效的数字")
        except Exception as e:
            print(f"截图失败: {e}")
    
    else:
        print("无效选项")


if __name__ == '__main__':
    main()

"""
电脑截屏工具 v2.5
作者：狗腿子 🐕
功能：全屏截图、选择窗口截图、保存到文件
修复：绕过 SetForegroundWindow 权限限制
"""

import win32gui
import win32ui
import win32con
import win32process
import ctypes
from PIL import Image
import os
import time
from datetime import datetime


def take_full_screenshot(save_path=None):
    """
    全屏截图
    
    Args:
        save_path: 保存路径
    
    Returns:
        PIL.Image 对象
    """
    # 获取屏幕尺寸
    hwnd = win32gui.GetDesktopWindow()
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    
    # 创建设备上下文
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    
    # 创建位图
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)
    
    # 截图
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
    
    # 转换为 PIL Image
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    screenshot = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )
    
    # 清理资源
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    
    if save_path:
        screenshot.save(save_path)
        print(f"截图已保存: {save_path}")
    
    return screenshot


def capture_window(hwnd, save_path=None):
    """
    截取指定窗口
    
    Args:
        hwnd: 窗口句柄
        save_path: 保存路径
    
    Returns:
        PIL.Image 对象
    """
    # 获取窗口位置和大小
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    
    print(f"窗口位置: ({left}, {top}), 大小: {width}x{height}")
    
    if width <= 0 or height <= 0:
        raise ValueError("窗口大小无效")
    
    # 创建设备上下文
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    
    # 创建位图
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)
    
    # 截图
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
    
    # 转换为 PIL Image
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    screenshot = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )
    
    # 清理资源
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    
    if save_path:
        screenshot.save(save_path)
        print(f"截图已保存: {save_path}")
    
    return screenshot


def force_set_foreground_window(hwnd):
    """
    强制将窗口置于前台（绕过 Windows 权限限制）
    
    Args:
        hwnd: 窗口句柄
    """
    # 方法1: 使用 AttachThreadInput 绕过限制
    foreground_thread = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())[0]
    current_thread = win32api.GetCurrentThreadId()
    target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]
    
    # 附加线程输入
    if current_thread != target_thread:
        ctypes.windll.user32.AttachThreadInput(current_thread, target_thread, True)
        if foreground_thread != target_thread:
            ctypes.windll.user32.AttachThreadInput(foreground_thread, target_thread, True)
    
    # 显示窗口
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.BringWindowToTop(hwnd)
    win32gui.SetForegroundWindow(hwnd)
    win32gui.SetFocus(hwnd)
    
    # 分离线程输入
    if current_thread != target_thread:
        ctypes.windll.user32.AttachThreadInput(current_thread, target_thread, False)
        if foreground_thread != target_thread:
            ctypes.windll.user32.AttachThreadInput(foreground_thread, target_thread, False)


def restore_window(hwnd):
    """
    恢复最小化的窗口
    
    Args:
        hwnd: 窗口句柄
    """
    # 检查窗口是否最小化
    if win32gui.IsIconic(hwnd):
        print("窗口已最小化，正在恢复...")
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)
    
    # 尝试将窗口置于前台
    try:
        force_set_foreground_window(hwnd)
        time.sleep(0.3)
        print("窗口已激活")
    except Exception as e:
        print(f"注意: 无法将窗口置于前台 ({e})，尝试直接截图...")


def list_windows():
    """列出所有可见窗口"""
    windows = []
    
    def enum_window_proc(hwnd, results):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                results.append({
                    'hwnd': hwnd,
                    'title': title
                })
    
    win32gui.EnumWindows(enum_window_proc, windows)
    return windows


# 需要导入 win32api
import win32api


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
    print("截屏工具 v2.5 🐕")
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
        print("\n正在截屏...")
        screenshot = take_full_screenshot(save_path)
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
        for i, win in enumerate(windows):
            print(f"[{i}] {win['title']}")
        print("-" * 50)
        
        try:
            window_index = int(input("请输入窗口编号: ").strip())
            
            if 0 <= window_index < len(windows):
                selected = windows[window_index]
                hwnd = selected['hwnd']
                
                print(f"\n正在截取窗口: {selected['title']}")
                
                # 恢复窗口（如果最小化）
                restore_window(hwnd)
                
                screenshot = capture_window(hwnd, save_path)
                print(f"截图完成！尺寸: {screenshot.size[0]} x {screenshot.size[1]}")
                print(f"保存位置: {save_path}")
            else:
                print("无效的窗口编号")
                
        except ValueError:
            print("请输入有效的数字")
        except Exception as e:
            print(f"截图失败: {e}")
            import traceback
            traceback.print_exc()
    
    else:
        print("无效选项")


if __name__ == '__main__':
    main()

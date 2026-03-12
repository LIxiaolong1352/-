import pyautogui
import time
import threading
import keyboard

class AutoClicker:
    def __init__(self):
        self.clicking = False
        self.interval = 0.1  # 点击间隔（秒）
        
    def click_loop(self):
        while self.clicking:
            pyautogui.click()
            time.sleep(self.interval)
    
    def start(self):
        if not self.clicking:
            self.clicking = True
            thread = threading.Thread(target=self.click_loop)
            thread.daemon = True
            thread.start()
            print("连点器已启动！按 F8 停止")
    
    def stop(self):
        self.clicking = False
        print("连点器已停止！按 F7 启动")

def main():
    clicker = AutoClicker()
    
    print("=" * 40)
    print("简单连点器")
    print("=" * 40)
    print("F7 - 启动连点")
    print("F8 - 停止连点")
    print("F9 - 退出程序")
    print("=" * 40)
    
    # 注册热键
    keyboard.add_hotkey('f7', clicker.start)
    keyboard.add_hotkey('f8', clicker.stop)
    keyboard.add_hotkey('f9', lambda: exit())
    
    print("\n等待热键...")
    keyboard.wait('f9')

if __name__ == "__main__":
    main()

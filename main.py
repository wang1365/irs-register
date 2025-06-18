# This is a sample Python script.
import threading
import time

import psutil
import configparser
from pywinauto import Application
import win32com.client


def get_pid_by_name(process_name):
    for proc in psutil.process_iter(['name', 'pid']):
        if process_name.lower() in proc.info['name'].lower():
            return proc.info['pid']
    return None


class AutoRegister:
    def __init__(self):
        config = configparser.ConfigParser()
        config.read('config.ini')

        self.found = False
        self.esc = False
        self.shell = win32com.client.Dispatch("WScript.Shell")

        # 创建结果文件, 文件命名为 result_{时间戳，格式为年-月-日 时分秒}.txt
        self.result_file = open(f'result_{datetime.now().strftime("%Y-%m-%d_%H_%M_%S")}.txt', 'a')

        self.init_product_key = config['DEFAULT']['ProductKey']
        self.reg_key_suffix = config['DEFAULT']['RegistrationKeySuffix']
        self.reg_key_start = int(config['DEFAULT']['RegistrationKeyStart'])
        self.reg_key_len = len(config['DEFAULT']['RegistrationKeyStart'])
        self.reg_key_count = int(config['DEFAULT']['RegistrationKeyCount'])

        # 获取记事本进程ID
        pid = get_pid_by_name('irsLINK_Server')

        # 连接到已运行的记事本应用
        self.app = Application(backend="win32").connect(process=pid)

        # 现在可以操作该应用了
        self.window = self.app.window(title_re='IRS Multi-Store Registration.*')
        self.window.set_focus()

        self.product_key = self.window.child_window(control_id=6)
        self.reg_key = self.window.child_window(control_id=7)
        self.save_btn = self.window.child_window(control_id=8)

    def log(self, msg):
        print(msg)
        self.result_file.write(msg + '\n')
        self.result_file.flush()

    @staticmethod
    def click(ctl):
        retry = 3
        while retry > 0:
            try:
                ctl.click()
                break
            except RuntimeError as e:
                retry -= 1
                print('====> retry for', e)

    def click_save_with_kb(self):
        self.shell.SendKeys('%s')

    def click_enter_with_kb(self):
        self.shell.SendKeys('{ENTER}')

    def find_result_dlg(self):
        retry = 5
        while retry > 0:
            try:
                result = self.app.window(title='IRS Registration')
                result.wait('exists', timeout=1, retry_interval=1)
                return result
            except RuntimeError as e:
                retry -= 1
                print('====> not find result dlg, retry for', e)
                time.sleep(1)

        print('====> retry failed, re-click SAVE button')
        self.click(self.save_btn)
        time.sleep(1)
        result = self.app.window(title='IRS Registration')
        result.wait('exists', timeout=60, retry_interval=1)

        return result

    def try_key(self, key):
        self.window.wait('exists', timeout=30, retry_interval=1)
        self.reg_key.set_text(key)
        # self.click(self.save_btn)
        self.click_save_with_kb()
        self.click_enter_with_kb()

        # result = self.find_result_dlg()
        # try:
        #     success = not result.child_window(title='Invalid Registration Key !').exists()
        #     if success:
        #         with open('result.txt', 'a') as f:
        #             f.write(f'{key}: {success}\n')
        #
        #     result.close()
        # except Exception as e:
        #     # 打印异常堆栈
        #     import traceback
        #     print('====>', e)
        #     traceback.print_exc()

    def run(self):
        from datetime import datetime
        start = datetime.now()
        self.product_key.set_text(self.init_product_key)

        # 创建一个线程来检测self.window是否存在,如果不存在则说明注册成功，key_found设置为True
        def check_window_exists():
            while self.found:
                if not self.window.exists(timeout=5, retry_interval=1):
                    time.sleep(1)
                    if not self.window.exists(timeout=3, retry_interval=1):
                        self.found = True
                        self.log('====> key found, exiting...')
                        break
                time.sleep(1)

        import threading
        threading.Thread(target=check_window_exists).start()

        self.listen_esc()

        for i in range(self.reg_key_count):
            key_prefix = str(self.reg_key_start + i).zfill(self.reg_key_len)
            key = key_prefix + self.reg_key_suffix
            self.log(f'{i+1}/{self.reg_key_count} - {key} - {datetime.now() - start}')
            self.try_key(key)
            if self.found or self.esc:
                break

        end = datetime.now()
        self.log(f'total time:{end - start}')

    def listen_esc(self):
        from pynput import keyboard
        def on_press(key):
            try:
                # 检查按下的键是否为Esc键
                if key == keyboard.Key.esc:
                    self.log('Esc 键被按下了！')
                    self.esc = True
                    # 如果需要在Esc键被按下时停止监听，可以返回 `False`
                    return False  # 返回 False 可以停止监听
            except AttributeError:
                # 非特殊按键，无需处理
                pass

        def run_listener():
            # 创建一个监听器来监听键盘事件
            with keyboard.Listener(on_press=on_press) as listener:
                listener.join()

        threading.Thread(target=run_listener).start()
        # 创建一个监听器来监听键盘事件


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    from datetime import datetime

    start = datetime.now()

    # 如果当前日期晚于2025-06-20，直接退出
    if datetime.now() > datetime(2026, 5, 20):
        print('>>>')
    else:
        try:
            AutoRegister().run()
        except Exception as e:
            # 打印异常堆栈
            import traceback

            traceback.print_exc()

    time.sleep(1)
    # 等待用户输入
    input("Press Enter to exit...")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

# This is a sample Python script.
import time

import psutil
import configparser
from pywinauto import Application


def get_pid_by_name(process_name):
    for proc in psutil.process_iter(['name', 'pid']):
        if process_name.lower() in proc.info['name'].lower():
            return proc.info['pid']
    return None

class AutoRegister:
    def __init__(self):
        config = configparser.ConfigParser()
        config.read('config.ini')

        self.init_product_key = config['DEFAULT']['ProductKey']
        self.reg_key_suffix = config['DEFAULT']['RegistrationKeySuffix']
        self.reg_key_start = int(config['DEFAULT']['RegistrationKeyStart'])
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
        self.click(self.save_btn)

        result = self.find_result_dlg()
        try:
            success = not result.child_window(title='Invalid Registration Key !').exists()
            if success:
                with open('result.txt', 'a') as f:
                    f.write(f'{key}: {success}\n')

            result.close()
        except Exception as e:
            # 打印异常堆栈
            import traceback
            print('====>', e)
            traceback.print_exc()

    def run(self):
        from datetime import datetime
        start = datetime.now()
        self.product_key.set_text(self.init_product_key)

        random_keys = [f'{i}{self.reg_key_suffix}' for i in range(self.reg_key_start, self.reg_key_start + self.reg_key_count)]
        for i, key in enumerate(random_keys):
            print(f'{i}/{self.reg_key_count} - {key} - {datetime.now() - start}')
            self.try_key(key)

        end = datetime.now()
        print('total time:', end - start)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    from datetime import datetime

    start = datetime.now()

    # 如果当前日期晚于2025-06-20，直接退出
    if datetime.now() > datetime(2025, 6, 20):
        print('license expired, exiting...')
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

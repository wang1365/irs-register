# -*- coding: utf-8 -*-
import configparser
import ctypes
import getpass
import os
import platform
import queue
import socket
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, messagebox
from tkinter import ttk
import tkinter as tk

import psutil


AUTHOR_NAME = "Wang Xiaochuan"
AUTHOR_EMAIL = "wangxiaochuan01@163.com"
CONFIG_PATH = Path("config.ini")
WINDOW_SIZE = "1180x760"
WINDOW_MIN_WIDTH = 1080
WINDOW_MIN_HEIGHT = 700
DEFAULT_FONT_SIZE = 12
TEXT_FONT_SIZE = 11
CONFIG_KEYS = (
    "ProductKey",
    "RegistrationKeySuffix",
    "RegistrationKeyStart",
    "RegistrationKeyCount",
)


def enable_dpi_awareness():
    if platform.system() != "Windows":
        return

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def get_pid_by_name(process_name):
    for proc in psutil.process_iter(["name", "pid"]):
        name = proc.info.get("name") or ""
        if process_name.lower() in name.lower():
            return proc.info["pid"]
    return None


def format_bytes(size):
    try:
        size = float(size)
    except (TypeError, ValueError):
        return "N/A"

    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.2f} {unit}"
        size /= 1024

    return "N/A"


def get_network_summary():
    addresses = []
    try:
        for interface, addr_list in psutil.net_if_addrs().items():
            for addr in addr_list:
                if addr.family == socket.AF_INET and not addr.address.startswith("127."):
                    addresses.append(f"{interface}: {addr.address}")
    except Exception:
        return "N/A"

    return ", ".join(addresses) if addresses else "N/A"


def get_system_info_lines():
    uname = platform.uname()
    memory = psutil.virtual_memory()
    disk = psutil.disk_usage(os.path.abspath(os.sep))

    return [
        ("Author", AUTHOR_NAME),
        ("Email", AUTHOR_EMAIL),
        ("Date Time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("User", getpass.getuser()),
        ("Host Name", socket.gethostname()),
        ("Operating System", f"{uname.system} {uname.release} ({uname.version})"),
        ("Machine", uname.machine),
        ("Processor", uname.processor or platform.processor() or "N/A"),
        ("CPU Cores", f"physical={psutil.cpu_count(logical=False)}, logical={psutil.cpu_count(logical=True)}"),
        ("Memory", f"total={format_bytes(memory.total)}, available={format_bytes(memory.available)}"),
        ("Disk", f"total={format_bytes(disk.total)}, free={format_bytes(disk.free)}"),
        ("Network", get_network_summary()),
    ]


def print_startup_info():
    print("=" * 60)
    print("Auto Registry Startup Information")
    for key, value in get_system_info_lines():
        print(f"{key}: {value}")
    print("=" * 60)


def load_config_values(path=CONFIG_PATH):
    parser = configparser.ConfigParser()
    parser.optionxform = str
    parser.read(path, encoding="utf-8")
    defaults = parser["DEFAULT"]

    return {
        "ProductKey": defaults.get("ProductKey", ""),
        "RegistrationKeySuffix": defaults.get("RegistrationKeySuffix", ""),
        "RegistrationKeyStart": defaults.get("RegistrationKeyStart", "00001"),
        "RegistrationKeyCount": defaults.get("RegistrationKeyCount", "10000"),
    }


def save_config_values(values, path=CONFIG_PATH):
    parser = configparser.ConfigParser()
    parser.optionxform = str
    parser["DEFAULT"] = {key: str(values.get(key, "")) for key in CONFIG_KEYS}

    with open(path, "w", encoding="utf-8") as config_file:
        parser.write(config_file)


def validate_config_values(values):
    errors = []
    if not values["ProductKey"].strip():
        errors.append("ProductKey is required.")
    if not values["RegistrationKeySuffix"].strip():
        errors.append("RegistrationKeySuffix is required.")
    if not values["RegistrationKeyStart"].strip().isdigit():
        errors.append("RegistrationKeyStart must be a number.")
    if not values["RegistrationKeyCount"].strip().isdigit():
        errors.append("RegistrationKeyCount must be a number.")
    elif int(values["RegistrationKeyCount"]) <= 0:
        errors.append("RegistrationKeyCount must be greater than 0.")

    return errors


class AutoRegister:
    def __init__(self, config_values=None, stop_event=None, log_callback=None):
        from pywinauto import Application
        import win32com.client

        values = config_values or load_config_values()

        self.found = False
        self.esc = False
        self.stop_event = stop_event or threading.Event()
        self.log_callback = log_callback
        self.shell = win32com.client.Dispatch("WScript.Shell")

        self.result_file = open(
            f'result_{datetime.now().strftime("%Y-%m-%d_%H_%M_%S")}.txt',
            "a",
            encoding="utf-8",
        )

        self.init_product_key = values["ProductKey"]
        self.reg_key_suffix = values["RegistrationKeySuffix"]
        self.reg_key_start = int(values["RegistrationKeyStart"])
        self.reg_key_len = len(values["RegistrationKeyStart"])
        self.reg_key_count = int(values["RegistrationKeyCount"])

        pid = get_pid_by_name("irsLINK_Server")
        if pid is None:
            raise RuntimeError("Process irsLINK_Server was not found. Start the target application first.")

        self.app = Application(backend="win32").connect(process=pid)
        self.window = self.app.window(title_re="IRS Multi-Store Registration.*")
        self.window.set_focus()

        self.product_key = self.window.child_window(control_id=6)
        self.reg_key = self.window.child_window(control_id=7)
        self.save_btn = self.window.child_window(control_id=8)

    def should_stop(self):
        return self.esc or self.stop_event.is_set()

    def stop(self):
        self.esc = True
        self.stop_event.set()

    def log(self, msg):
        print(msg)
        self.result_file.write(msg + "\n")
        self.result_file.flush()
        if self.log_callback:
            self.log_callback(msg)

    @staticmethod
    def click(ctl):
        retry = 3
        while retry > 0:
            try:
                ctl.click()
                break
            except RuntimeError as e:
                retry -= 1
                print("====> retry for", e)

    def click_save_with_kb(self):
        self.shell.SendKeys("%s")

    def click_enter_with_kb(self):
        self.shell.SendKeys("{ENTER}")

    def find_result_dlg(self):
        retry = 5
        while retry > 0 and not self.should_stop():
            try:
                result = self.app.window(title="IRS Registration")
                result.wait("exists", timeout=1, retry_interval=1)
                return result
            except RuntimeError as e:
                retry -= 1
                print("====> not find result dlg, retry for", e)
                time.sleep(1)

        print("====> retry failed, re-click SAVE button")
        self.click(self.save_btn)
        time.sleep(0.5)
        result = self.app.window(title="IRS Registration")
        result.wait("exists", timeout=60, retry_interval=1)

        return result

    def try_key(self, key):
        self.window.wait("exists", timeout=30, retry_interval=1)
        self.reg_key.set_text(key)
        self.click_save_with_kb()
        self.click_enter_with_kb()

    def run(self):
        start = datetime.now()
        self.product_key.set_text(self.init_product_key)

        def check_window_exists():
            while not self.should_stop() and not self.found:
                if not self.window.exists(timeout=5, retry_interval=1):
                    time.sleep(1)
                    if not self.window.exists(timeout=3, retry_interval=1):
                        self.found = True
                        self.log("====> key found, exiting...")
                        break
                time.sleep(1)

        threading.Thread(target=check_window_exists, daemon=True).start()
        self.listen_esc()

        for i in range(self.reg_key_count):
            if self.should_stop() or self.found:
                break

            key_prefix = str(self.reg_key_start + i).zfill(self.reg_key_len)
            key = key_prefix + self.reg_key_suffix
            self.log(f"{i + 1}/{self.reg_key_count} - {key} - {datetime.now() - start}")
            self.try_key(key)

        end = datetime.now()
        if self.should_stop():
            self.log("====> stopped by user")
        self.log(f"total time:{end - start}")

    def listen_esc(self):
        from pynput import keyboard

        def on_press(key):
            try:
                if key == keyboard.Key.esc:
                    self.log("Esc pressed, stopping...")
                    self.stop()
                    return False
            except AttributeError:
                pass

        def run_listener():
            with keyboard.Listener(on_press=on_press) as listener:
                listener.join()

        threading.Thread(target=run_listener, daemon=True).start()


class AutoRegisterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Registry")
        self.configure_fonts()
        self.root.geometry(self.get_initial_geometry())
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)

        self.log_queue = queue.Queue()
        self.worker = None
        self.stop_event = threading.Event()
        self.register = None
        self.config_vars = {key: tk.StringVar() for key in CONFIG_KEYS}

        self.build_ui()
        self.load_config_into_form()
        self.refresh_system_info()
        self.set_running(False)
        self.root.update_idletasks()
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)
        self.root.after(200, self.drain_log_queue)

    def configure_fonts(self):
        default_font = ("Segoe UI", DEFAULT_FONT_SIZE)
        text_font = ("Consolas", TEXT_FONT_SIZE)
        self.root.option_add("*Font", default_font)
        self.root.option_add("*Text.Font", text_font)

    def get_initial_geometry(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = max(WINDOW_MIN_WIDTH, int(screen_width * 0.58))
        height = max(WINDOW_MIN_HEIGHT, int(screen_height * 0.68))
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        return f"{width}x{height}+{x}+{y}"

    def build_ui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=BOTH, expand=True, padx=12, pady=12)

        config_frame = ttk.LabelFrame(main_pane, text="Configuration")
        info_frame = ttk.LabelFrame(main_pane, text="System Information")
        main_pane.add(config_frame, weight=1)
        main_pane.add(info_frame, weight=2)

        form = ttk.Frame(config_frame)
        form.pack(fill=X, padx=12, pady=12)

        labels = {
            "ProductKey": "Product Key",
            "RegistrationKeySuffix": "Registration Key Suffix",
            "RegistrationKeyStart": "Registration Key Start",
            "RegistrationKeyCount": "Registration Key Count",
        }

        for row, key in enumerate(CONFIG_KEYS):
            ttk.Label(form, text=labels[key]).grid(row=row, column=0, sticky=tk.W, pady=6)
            entry = ttk.Entry(form, textvariable=self.config_vars[key])
            entry.grid(row=row, column=1, sticky=tk.EW, padx=(10, 0), pady=6)
        form.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(config_frame)
        button_frame.pack(fill=X, padx=12, pady=(0, 12))

        self.start_button = ttk.Button(button_frame, text="Start", command=self.start_register)
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_register)
        self.close_button = ttk.Button(button_frame, text="Close", command=self.close_app)
        self.start_button.pack(side=LEFT, padx=(0, 8))
        self.stop_button.pack(side=LEFT, padx=(0, 8))
        self.close_button.pack(side=RIGHT)

        log_frame = ttk.LabelFrame(config_frame, text="Run Log")
        log_frame.pack(fill=BOTH, expand=True, padx=12, pady=(0, 12))
        self.log_text = tk.Text(log_frame, height=12, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=BOTH, expand=True, padx=8, pady=8)

        top_info = ttk.Frame(info_frame)
        top_info.pack(fill=X, padx=12, pady=12)
        ttk.Button(top_info, text="Refresh", command=self.refresh_system_info).pack(side=RIGHT)

        self.info_text = tk.Text(info_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.info_text.pack(fill=BOTH, expand=True, padx=12, pady=(0, 12))

    def load_config_into_form(self):
        try:
            values = load_config_values()
        except Exception as e:
            messagebox.showerror("Config Load Failed", str(e))
            values = load_config_values(Path("__missing_config__.ini"))

        for key, value in values.items():
            self.config_vars[key].set(value)

    def get_form_values(self):
        return {key: var.get().strip() for key, var in self.config_vars.items()}

    def refresh_system_info(self):
        lines = [f"{key}: {value}" for key, value in get_system_info_lines()]
        self.info_text.configure(state=tk.NORMAL)
        self.info_text.delete("1.0", END)
        self.info_text.insert(END, "\n".join(lines))
        self.info_text.configure(state=tk.DISABLED)

    def append_log(self, msg):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(END, f"{datetime.now().strftime('%H:%M:%S')} {msg}\n")
        self.log_text.see(END)
        self.log_text.configure(state=tk.DISABLED)

    def queue_log(self, msg):
        self.log_queue.put(msg)

    def drain_log_queue(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.append_log(msg)

        if self.worker and not self.worker.is_alive():
            self.worker = None
            self.register = None
            self.set_running(False)

        self.root.after(200, self.drain_log_queue)

    def set_running(self, running):
        self.start_button.configure(state=tk.DISABLED if running else tk.NORMAL)
        self.stop_button.configure(state=tk.NORMAL if running else tk.DISABLED)
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)

    def restore_window_geometry(self, geometry):
        if self.root.winfo_exists():
            self.root.geometry(geometry)

    def start_register(self):
        current_geometry = self.root.geometry()
        values = self.get_form_values()
        errors = validate_config_values(values)
        if errors:
            messagebox.showerror("Invalid Configuration", "\n".join(errors))
            return

        try:
            save_config_values(values)
        except Exception as e:
            messagebox.showerror("Config Save Failed", str(e))
            return

        self.stop_event = threading.Event()
        self.set_running(True)
        self.root.geometry(current_geometry)
        self.append_log("Configuration saved. Starting...")

        def run_worker():
            try:
                self.register = AutoRegister(
                    config_values=values,
                    stop_event=self.stop_event,
                    log_callback=self.queue_log,
                )
                self.register.run()
            except Exception:
                self.queue_log(traceback.format_exc())

        self.worker = threading.Thread(target=run_worker, daemon=True)
        self.worker.start()
        for delay in (100, 500, 1000):
            self.root.after(delay, lambda geometry=current_geometry: self.restore_window_geometry(geometry))

    def stop_register(self):
        self.stop_event.set()
        if self.register:
            self.register.stop()
        self.append_log("Stopping...")

    def close_app(self):
        if self.worker and self.worker.is_alive():
            self.stop_register()
            if not messagebox.askokcancel("Close Application", "A task is still running. Close anyway?"):
                return
        self.root.destroy()


def main():
    enable_dpi_awareness()
    print_startup_info()
    root = tk.Tk()
    AutoRegisterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

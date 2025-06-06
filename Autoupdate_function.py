import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import shutil
import requests
from packaging import version  # pip install packaging

CURRENT_VERSION = "v0422"
VERSION_FILE = r"\\wectinfo02\pp00\yplu\version.txt"

def get_update_url(latest_version):
    return f"file://wectinfo02/pp00/yplu/Booking_{latest_version}.7z"

class UpdateApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("版本更新工具")
        self.geometry("400x200")
        self.resizable(False, False)
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.create_widgets()

    def create_widgets(self):
        self.status_label = ttk.Label(self, text="等待檢查更新...", font=("Arial", 12))
        self.status_label.pack(pady=20)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.check_button = ttk.Button(self, text="檢查更新", command=self.start_update_check)
        self.check_button.pack(pady=10)

    def update_status(self, message):
        self.status_label.config(text=message)

    def start_update_check(self):
        self.check_button.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self.check_for_update, daemon=True).start()

    def check_for_update(self):
        try:
            with open(VERSION_FILE, "r") as file:
                

                #################自行修改#################
                
                #讀取第一行數字是否MATCH
                latest_version = file.read().strip()

                # # 強制取第二行作為最新版本
                # lines = file.readlines()
                # latest_version = lines[1].strip()

                #################自行修改#################

            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                self.safe_update_status(f"發現新版本：{latest_version}")
                self.download_update(latest_version)
            else:
                self.safe_update_status("目前已是最新版本")
                self.safe_enable_button()
        except Exception as e:
            self.safe_update_status(f"檢查更新失敗：{e}")
            self.safe_enable_button()

    def download_update(self, latest_version):
        self.safe_update_status("開始下載更新...")
        UPDATE_URL = get_update_url(latest_version)
        update_path = f"Booking_{latest_version}.7z"
        if UPDATE_URL.startswith("file://"):
            # file:// URL, convert to UNC path
            local_path = UPDATE_URL.replace("file://", r"\\")
            try:
                # Get file size to update progress bar
                total_size = os.path.getsize(local_path)
                self.safe_set_progress_max(total_size)
                with open(local_path, "rb") as src, open(update_path, "wb") as dst:
                    chunk_size = 1024 * 10
                    bytes_copied = 0
                    while True:
                        chunk = src.read(chunk_size)
                        if not chunk:
                            break
                        dst.write(chunk)
                        bytes_copied += len(chunk)
                        self.safe_update_progress(bytes_copied)
                self.safe_update_status("下載完成，請進行安裝")
            except Exception as e:
                self.safe_update_status(f"本機複製更新失敗：{e}")
        else:
            try:
                response = requests.get(UPDATE_URL, stream=True)
                total_size = int(response.headers.get("Content-Length", 0))
                if total_size:
                    self.safe_set_progress_max(total_size)
                else:
                    self.progress.config(mode="indeterminate")
                    self.safe_update_status("下載進度未知，開始下載...")
                    self.progress.start(10)
                with open(update_path, "wb") as f:
                    bytes_downloaded = 0
                    chunk_size = 1024 * 10
                    for chunk in response.iter_content(chunk_size=chunk_size):
                        if chunk:
                            f.write(chunk)
                            bytes_downloaded += len(chunk)
                            if total_size:
                                self.safe_update_progress(bytes_downloaded)
                if not total_size:
                    self.progress.stop()
                self.safe_update_status("下載完成，請進行安裝")
            except Exception as e:
                self.safe_update_status(f"網路下載更新失敗：{e}")
        self.safe_enable_button()

    # Methods to safely update the UI from other threads.
    def safe_update_status(self, message):
        self.after(0, lambda: self.update_status(message))

    def safe_update_progress(self, value):
        self.after(0, lambda: self.progress.config(value=value))

    def safe_set_progress_max(self, max_value):
        self.after(0, lambda: self.progress.config(maximum=max_value))

    def safe_enable_button(self):
        self.after(0, lambda: self.check_button.config(state="normal"))

if __name__ == "__main__":
    app = UpdateApp()
    app.mainloop()

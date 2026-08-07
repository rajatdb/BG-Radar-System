import os
import sys
import json
import urllib.request
import threading
import subprocess
import time
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
import ctypes

class InAppSeamlessUpdater:
    def __init__(self, current_version="v3.4", product_code="NWR_BG_RADAR_PRO", parent_app=None):
        self.current_version = current_version
        self.product_code = product_code
        self.parent_app = parent_app
        
        # GitHub Manifest URL
        self.manifest_url = "https://raw.githubusercontent.com/rajatdb/BG-Radar-System/main/version_manifest.json"
        
        if getattr(sys, 'frozen', False):
            self.current_exe = sys.executable
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.current_exe = os.path.abspath(__file__)
            self.base_dir = os.path.dirname(os.path.abspath(__file__))

    def fetch_manifest(self):
        """Fetches remote release manifest directly from GitHub"""
        try:
            req = urllib.request.Request(
                self.manifest_url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            )
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = resp.read().decode('utf-8')
                return json.loads(data)
        except Exception:
            # Fallback to local manifest if offline / local testing
            manifest_path = os.path.join(self.base_dir, "version_manifest.json")
            if os.path.exists(manifest_path):
                try:
                    with open(manifest_path, "r", encoding="utf-8") as f:
                        return json.load(f)
                except Exception:
                    pass
            return None

    def check_and_update(self):
        """Checks version and initiates live update if available"""
        manifest = self.fetch_manifest()
        if not manifest or manifest.get("product_code") != self.product_code:
            return

        latest_ver = manifest.get("latest_version", self.current_version)
        download_url = manifest.get("download_url", "")

        if latest_ver != self.current_version and download_url:
            notes = "\n• " + "\n• ".join(manifest.get("update_notes", ["Performance improvements"]))
            msg = (
                f"🚀 OFFICIAL SOFTWARE UPDATE AVAILABLE!\n\n"
                f"📌 Installed Version: {self.current_version}\n"
                f"✨ New Version: {latest_ver} ({manifest.get('release_date', 'IST')})\n\n"
                f"What's New:{notes}\n\n"
                f"Would you like to update the running software automatically now?"
            )
            
            choice = messagebox.askyesno("NWR Portal Auto-Updater", msg, parent=self.parent_app)
            if choice:
                self.show_download_progress_dialog(download_url, latest_ver)

    def show_download_progress_dialog(self, download_url, new_version):
        """Creates progress window during update"""
        dialog = tk.Toplevel(self.parent_app)
        dialog.title("Installing Updates...")
        dialog.geometry("420x180")
        dialog.resizable(False, False)
        dialog.configure(bg="#061d33")
        dialog.transient(self.parent_app)
        dialog.grab_set()

        try:
            dialog.geometry("+%d+%d" % (
                self.parent_app.winfo_rootx() + self.parent_app.winfo_width()//2 - 210,
                self.parent_app.winfo_rooty() + self.parent_app.winfo_height()//2 - 90
            ))
        except Exception:
            pass

        lbl_title = tk.Label(
            dialog, text=f"Downloading Official Update {new_version}",
            font=("Segoe UI", 11, "bold"), fg="#d97706", bg="#061d33"
        )
        lbl_title.pack(pady=(20, 10))

        progress = ttk.Progressbar(dialog, mode='determinate', length=340)
        progress.pack(pady=10)

        lbl_status = tk.Label(
            dialog, text="Connecting to Railway Update Server...",
            font=("Segoe UI", 9, "italic"), fg="#cbd5e1", bg="#061d33"
        )
        lbl_status.pack(pady=(5, 10))

        threading.Thread(
            target=self.download_and_apply_update,
            args=(download_url, progress, lbl_status, dialog),
            daemon=True
        ).start()

    def download_and_apply_update(self, download_url, progress, status_lbl, dialog):
        """Downloads updated EXE with full Redirect Handling & Integrity Check"""
        try:
            temp_dir = tempfile.gettempdir()
            temp_new_exe = os.path.join(temp_dir, "BG_Radar_New.exe")

            # Remove old leftover file if present
            if os.path.exists(temp_new_exe):
                try:
                    os.remove(temp_new_exe)
                except Exception:
                    pass

            # Setup urllib request with proper headers to follow GitHub AWS S3 Redirects
            opener = urllib.request.build_opener(urllib.request.HTTPRedirectHandler)
            req = urllib.request.Request(
                download_url,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                    'Accept': 'application/octet-stream'
                }
            )

            with opener.open(req, timeout=60) as response:
                # Get final content length after redirect
                total_size = response.getheader('Content-Length')
                total_size = int(total_size) if total_size else 0

                downloaded = 0
                block_size = 16384  # 16KB buffer for faster download

                with open(temp_new_exe, 'wb') as out_file:
                    while True:
                        buffer = response.read(block_size)
                        if not buffer:
                            break
                        downloaded += len(buffer)
                        out_file.write(buffer)

                        if total_size > 0:
                            percent = int((downloaded / total_size) * 100)
                            dialog.after(0, lambda p=percent: progress.config(value=p))
                            dialog.after(0, lambda d=downloaded, t=total_size: status_lbl.config(
                                text=f"Downloaded {d // 1024} KB / {t // 1024} KB ({percent}%)"
                            ))
                        else:
                            dialog.after(0, lambda d=downloaded: status_lbl.config(
                                text=f"Downloaded {d // 1024} KB..."
                            ))

            # --- VALIDATION: Ensure file is a genuine executable (> 2 MB) ---
            file_size = os.path.getsize(temp_new_exe)
            if file_size < 2000000:
                raise ValueError(f"Downloaded file size ({file_size} bytes) is too small. Expected full executable.")

            # Validate Windows PE EXE Header (First 2 bytes must be 'MZ')
            with open(temp_new_exe, 'rb') as f:
                header = f.read(2)
                if header != b'MZ':
                    raise ValueError("Downloaded file is not a valid Windows Binary Executable (Corrupted/HTML page received).")

            dialog.after(0, lambda: status_lbl.config(text="Applying Update & Restarting Portal...", fg="#4ade80"))
            time.sleep(1)

            self.execute_batch_restart(temp_new_exe)

        except Exception as e:
            dialog.after(0, lambda: messagebox.showerror("Update Error", f"Failed to apply auto-update:\n{e}"))
            dialog.after(0, dialog.destroy)

    def execute_batch_restart(self, temp_new_exe):
        """Generates elevated script to replace Program Files executable and restart"""
        temp_dir = tempfile.gettempdir()
        bat_file = os.path.join(temp_dir, "apply_update.bat")
        exe_name = os.path.basename(self.current_exe)
        
        # Ensured application process termination before moving binary
        bat_content = f"""@echo off
timeout /t 2 /nobreak > nul
taskkill /f /im "{exe_name}" >nul 2>&1
move /y "{temp_new_exe}" "{self.current_exe}"
start "" "{self.current_exe}"
del "%~f0"
"""
        with open(bat_file, "w") as f:
            f.write(bat_content)

        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", bat_file, None, None, 0)
        except Exception:
            subprocess.Popen([bat_file], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        if self.parent_app:
            self.parent_app.after(100, self.parent_app.destroy)
        else:
            sys.exit(0)
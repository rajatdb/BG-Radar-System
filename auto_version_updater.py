import os
import sys
import json
import urllib.request
import threading
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
        
        # Manifest URL on GitHub
        self.manifest_url = "https://raw.githubusercontent.com/rajatdb/BG-Radar-System/main/version_manifest.json"
        
        if getattr(sys, 'frozen', False):
            self.current_exe = sys.executable
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.current_exe = os.path.abspath(__file__)
            self.base_dir = os.path.dirname(os.path.abspath(__file__))

    def fetch_manifest(self):
        """Fetches live remote release manifest directly from GitHub with cache buster"""
        try:
            # 🚀 CACHE BUSTER: Prevents GitHub CDN 5-minute caching
            cache_buster_url = f"{self.manifest_url}?t={int(time.time())}"
            req = urllib.request.Request(
                cache_buster_url,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            )
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = resp.read().decode('utf-8')
                return json.loads(data)
        except Exception:
            return None

    def check_and_update(self):
        """Checks current vs remote version and prompts user if update is available"""
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
                f"Would you like to auto-update the system now?"
            )
            
            choice = messagebox.askyesno("NWR Portal Auto-Updater", msg, parent=self.parent_app)
            if choice:
                self.show_download_progress_dialog(download_url, latest_ver)

    def show_download_progress_dialog(self, download_url, new_version):
        """Displays in-app downloading progress window"""
        dialog = tk.Toplevel(self.parent_app)
        dialog.title("Updating Application...")
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
            dialog, text=f"Updating to Version {new_version}",
            font=("Segoe UI", 11, "bold"), fg="#d97706", bg="#061d33"
        )
        lbl_title.pack(pady=(20, 10))

        progress = ttk.Progressbar(dialog, mode='determinate', length=340)
        progress.pack(pady=10)

        lbl_status = tk.Label(
            dialog, text="Downloading official setup package...",
            font=("Segoe UI", 9, "italic"), fg="#cbd5e1", bg="#061d33"
        )
        lbl_status.pack(pady=(5, 10))

        threading.Thread(
            target=self.download_and_run_inno_setup,
            args=(download_url, progress, lbl_status, dialog),
            daemon=True
        ).start()

    def download_and_run_inno_setup(self, download_url, progress, status_lbl, dialog):
        """Downloads Inno Setup executable into Windows Temp and triggers silent update"""
        try:
            # 📁 DOWNLOAD TO WINDOWS TEMP DIRECTORY
            temp_dir = tempfile.gettempdir()
            installer_path = os.path.join(temp_dir, "BG_Radar_System_Update_Setup.exe")

            # Clean leftover installer from previous attempts
            if os.path.exists(installer_path):
                try:
                    os.remove(installer_path)
                except Exception:
                    pass

            opener = urllib.request.build_opener(urllib.request.HTTPRedirectHandler)
            req = urllib.request.Request(
                download_url,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                    'Accept': 'application/octet-stream'
                }
            )

            with opener.open(req, timeout=120) as response:
                total_size = response.getheader('Content-Length')
                total_size = int(total_size) if total_size else 0

                downloaded = 0
                block_size = 32768

                with open(installer_path, 'wb') as out_file:
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
                                text=f"Downloading Update: {d // (1024*1024)} MB / {t // (1024*1024)} MB ({percent}%)"
                            ))

            # Integrity Check: Ensure binary is valid
            file_size = os.path.getsize(installer_path)
            if file_size < 1000000:
                raise ValueError("Downloaded update installer file is corrupted or incomplete.")

            dialog.after(0, lambda: status_lbl.config(text="Launching Installer from Temp...", fg="#4ade80"))
            time.sleep(0.5)
            dialog.after(0, dialog.destroy)

            # 🚀 EXECUTE INNO SETUP SILENTLY IN ADMIN MODE
            # /SILENT = Minimalist updating UI without manual Next/License clicks
            inno_flags = '/SILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
            ctypes.windll.shell32.ShellExecuteW(None, "runas", installer_path, inno_flags, None, 1)

            # 🚨 HARD EXIT: Instantly release file lock on active executables
            if self.parent_app:
                self.parent_app.after(100, self.parent_app.destroy)
            
            time.sleep(0.2)
            os._exit(0)

        except Exception as e:
            dialog.after(0, lambda: messagebox.showerror("Update Error", f"Failed to apply auto-update:\n{e}"))
            dialog.after(0, dialog.destroy)
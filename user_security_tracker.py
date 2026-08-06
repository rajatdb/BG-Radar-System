# user_security_tracker.py
import getpass
import os
import platform
import socket
import urllib.request
import json
from datetime import datetime

class DeviceUserSecurityTracker:
    def __init__(self):
        self.username = self.get_system_username()
        self.computer_name = socket.gethostname()
        self.os_info = f"{platform.system()} {platform.release()}"
        self.ip_address = self.get_local_ip()

    @staticmethod
    def get_system_username():
        try:
            user = getpass.getuser()
            return user.upper() if user else "OFFICIAL_STAFF"
        except Exception:
            return os.environ.get("USERNAME", "OFFICIAL_STAFF").upper()

    @staticmethod
    def get_local_ip():
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    def get_device_info(self):
        """Returns structured metadata of the client system"""
        return {
            "computer_name": self.computer_name,
            "username": self.username,
            "ip_address": self.ip_address,
            "os_info": self.os_info,
            "last_seen": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
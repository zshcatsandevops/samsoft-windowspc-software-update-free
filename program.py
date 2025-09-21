#!/usr/bin/env python3
"""
Samsoft Update Manager (DebloatMy11 Companion)
- Windows Update–like GUI
- Auto-elevates to run as Administrator
- Installs PSWindowsUpdate automatically if missing
- Checks, downloads, installs updates
"""

import sys, os, ctypes, subprocess, threading
import tkinter as tk
from tkinter import ttk, messagebox

# ---------- Auto-elevation ----------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit()

# ---------- GUI Application ----------
class UpdateManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Samsoft Update Manager")
        self.root.geometry("700x450")
        self.root.configure(bg="#1e1e1e")

        self.status_var = tk.StringVar(value="Idle.")
        self.create_ui()

    def create_ui(self):
        frame = tk.Frame(self.root, bg="#1e1e1e")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(frame, text="Samsoft Update Manager", font=("Consolas", 16, "bold")).pack(pady=10)

        self.check_btn = ttk.Button(frame, text="Check for Updates", command=self.check_updates)
        self.check_btn.pack(pady=5)

        self.install_btn = ttk.Button(frame, text="Install Updates", command=self.install_updates, state="disabled")
        self.install_btn.pack(pady=5)

        log_frame = tk.LabelFrame(frame, text="Update Log", bg="#1e1e1e", fg="white")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.log_text = tk.Text(log_frame, wrap="word", bg="#0d0d0d", fg="#00ff00",
                                insertbackground="white", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.status_var.set(msg)

    def run_powershell(self, command):
        completed = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                                   capture_output=True, text=True)
        return completed.stdout.strip(), completed.stderr.strip()

    def ensure_module(self):
        """Check if PSWindowsUpdate is installed, install if missing."""
        check_cmd = "Get-Module -ListAvailable -Name PSWindowsUpdate"
        out, _ = self.run_powershell(check_cmd)
        if not out.strip():
            self.log("[INFO] PSWindowsUpdate not found. Installing...")
            install_cmd = "Install-PackageProvider -Name NuGet -Force; Install-Module PSWindowsUpdate -Force"
            out, err = self.run_powershell(install_cmd)
            if err:
                self.log(f"[ERROR] Failed to install PSWindowsUpdate: {err}")
                return False
            self.log("[OK] PSWindowsUpdate installed successfully.")
        return True

    def check_updates(self):
        threading.Thread(target=self._check_updates_thread, daemon=True).start()

    def _check_updates_thread(self):
        self.log("[CHECK] Looking for updates...")

        if not self.ensure_module():
            return

        cmd = "Import-Module PSWindowsUpdate; Get-WindowsUpdate -MicrosoftUpdate"
        out, err = self.run_powershell(cmd)

        if err:
            self.log(f"[ERROR] {err}")
            return

        if not out.strip():
            self.log("[ERROR] ALL GOOD [Y] PROCESS COMPLETED")  # Custom message when no updates
        else:
            self.log("[FOUND] Updates available:\n" + out)
            self.install_btn.config(state="normal")

    def install_updates(self):
        threading.Thread(target=self._install_updates_thread, daemon=True).start()

    def _install_updates_thread(self):
        self.log("[INSTALL] Installing updates...")

        if not self.ensure_module():
            return

        cmd = "Import-Module PSWindowsUpdate; Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot"
        out, err = self.run_powershell(cmd)
        if err:
            self.log(f"[ERROR] {err}")
            messagebox.showerror("Install Error", err)
            return
        self.log("[DONE] Updates installed. A restart may be required.")
        messagebox.showinfo("Updates", "Updates installed. Restart your PC if required.")
        self.install_btn.config(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateManagerApp(root)
    root.mainloop()

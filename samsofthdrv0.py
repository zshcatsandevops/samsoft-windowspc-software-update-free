#!/usr/bin/env python3
"""
Samsoft Update Manager CE (Community Edition)
- Inspired by WSUS Offline Update
- Online + Offline update modes
- Local repository for Windows/Office/.NET updates
- Tkinter GUI frontend
"""

import sys
import os
import ctypes
import subprocess
import threading
import time
import json
import textwrap  # Added for dedent to fix multi-line PS command parsing
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

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

# ---------- Configuration ----------
REPO_DIR = os.path.join(os.getcwd(), "SamsoftRepo")
CONFIG_FILE = os.path.join(REPO_DIR, "config.json")
os.makedirs(REPO_DIR, exist_ok=True)

# Default configuration
DEFAULT_CONFIG = {
    "repo_path": REPO_DIR,
    "update_categories": {
        "windows": True,
        "office": True,
        "dotnet": True,
        "vcredist": False
    },
    "auto_reboot": False
}

def load_config():
    """Load configuration from file or create default"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config):
    """Save configuration to file"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

# ---------- GUI Application ----------
class UpdateManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Samsoft Update Manager CE")
        self.root.geometry("900x600")
        self.root.configure(bg="#1e1e1e")
        
        self.config = load_config()
        self.repo_path = self.config.get("repo_path", REPO_DIR)
        self.status_var = tk.StringVar(value="Idle.")
        self.progress_var = tk.IntVar(value=0)
        self.pswindowsupdate_available = False  # Initialize here
        
        self.create_ui()
        self.check_pswindowsupdate()

    def create_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left panel for controls
        left_frame = ttk.LabelFrame(main_frame, text="Controls")
        left_frame.pack(side="left", fill="y", padx=(0, 10))
        
        # Right panel for log
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        
        # Control buttons
        ttk.Button(left_frame, text="Check Online", command=self.check_updates).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Download to Repo", command=self.download_updates).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Install Online", command=self.install_updates).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Install Offline (Repo)", command=self.install_offline).pack(pady=5, fill="x")
        
        ttk.Separator(left_frame, orient="horizontal").pack(fill="x", pady=10)
        
        ttk.Button(left_frame, text="Update Office (C2R)", command=self.update_office).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Update .NET", command=self.update_dotnet).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Update VC++ Redists", command=self.update_vcredist).pack(pady=5, fill="x")
        
        ttk.Separator(left_frame, orient="horizontal").pack(fill="x", pady=10)
        
        # Settings
        settings_frame = ttk.LabelFrame(left_frame, text="Settings")
        settings_frame.pack(fill="x", pady=5)
        
        self.auto_reboot_var = tk.BooleanVar(value=self.config.get("auto_reboot", False))
        ttk.Checkbutton(settings_frame, text="Auto Reboot", 
                       variable=self.auto_reboot_var,
                       command=self.toggle_auto_reboot).pack(anchor="w", pady=2)
        
        ttk.Button(settings_frame, text="Change Repo Path", 
                  command=self.change_repo_path).pack(fill="x", pady=5)
        
        # Progress bar
        progress_frame = ttk.Frame(left_frame)
        progress_frame.pack(fill="x", pady=10)
        
        ttk.Label(progress_frame, text="Progress:").pack(anchor="w")
        ttk.Progressbar(progress_frame, variable=self.progress_var, 
                       maximum=100).pack(fill="x", pady=5)
        
        # Log frame
        log_frame = ttk.LabelFrame(right_frame, text="Update Log")
        log_frame.pack(fill="both", expand=True)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(text_frame, wrap="word", bg="#0d0d0d", fg="#00ff00",
                               insertbackground="white", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Status bar
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom", padx=5, pady=2)

    # ---------- Helpers ----------
    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.status_var.set(msg)
        self.root.update_idletasks()

    def run_powershell(self, command, capture_output=True):
        try:
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                capture_output=capture_output, text=True, timeout=3600
            )
            stdout = completed.stdout if completed.stdout else ""
            stderr = completed.stderr if completed.stderr else ""
            return stdout.strip(), stderr.strip(), completed.returncode
        except subprocess.TimeoutExpired:
            return "", "Command timed out after 1 hour", 1
        except Exception as e:
            return "", f"Error executing PowerShell: {str(e)}", 1

    def check_pswindowsupdate(self):
        """Check if PSWindowsUpdate module is available"""
        self.log("[INFO] Checking for PSWindowsUpdate module...")
        check_cmd = "Get-Module -ListAvailable -Name PSWindowsUpdate"
        out, err, code = self.run_powershell(check_cmd)
        if not out.strip() and not err:
            self.log("[INFO] PSWindowsUpdate module not found. It will be installed when needed.")
            self.pswindowsupdate_available = False
        else:
            self.pswindowsupdate_available = True
            self.log("[OK] PSWindowsUpdate module is available.")

    def ensure_module(self):
        if self.pswindowsupdate_available:
            return True
            
        self.log("[INFO] Installing PSWindowsUpdate...")
        install_cmd = textwrap.dedent("""
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            Install-PackageProvider -Name NuGet -Force
            Install-Module PSWindowsUpdate -Force -AcceptLicense -Scope AllUsers
        """)
        _, err, code = self.run_powershell(install_cmd)
        if err and "already exists" not in err and "NoMatch" not in err:
            self.log(f"[ERROR] Failed to install PSWindowsUpdate: {err}")
            return False
        
        self.pswindowsupdate_available = True
        self.log("[OK] PSWindowsUpdate installed successfully.")
        return True

    def update_progress(self, value):
        self.progress_var.set(value)
        self.root.update_idletasks()

    def toggle_auto_reboot(self):
        self.config["auto_reboot"] = self.auto_reboot_var.get()
        save_config(self.config)

    def change_repo_path(self):
        new_path = filedialog.askdirectory(initialdir=self.repo_path, title="Select Repository Directory")
        if new_path:
            self.repo_path = new_path
            self.config["repo_path"] = new_path
            save_config(self.config)
            self.log(f"[INFO] Repository path changed to: {new_path}")

    # ---------- Windows ----------
    def check_updates(self):
        threading.Thread(target=self._check_updates_thread, daemon=True).start()

    def _check_updates_thread(self):
        self.log("[CHECK] Searching online for updates...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        self.update_progress(30)
        cmd = "Import-Module PSWindowsUpdate; Get-WUList -MicrosoftUpdate -Verbose"
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(70)
        if err and "0x80240024" not in err:  # Ignore "no updates available" error
            self.log(f"[ERROR] {err}")
        elif not out.strip() or "0x80240024" in err:
            self.log("[OK] System is up to date.")
        else:
            self.log("[FOUND] Available updates:")
            self.log(out)
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def download_updates(self):
        threading.Thread(target=self._download_thread, daemon=True).start()

    def _download_thread(self):
        self.log(f"[DOWNLOAD] Saving updates to {self.repo_path}...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        self.update_progress(20)
        
        # Create download directory if it doesn't exist
        download_dir = os.path.join(self.repo_path, "Downloads")
        os.makedirs(download_dir, exist_ok=True)
        
        self.update_progress(30)
        cmd = textwrap.dedent(f"""
            Import-Module PSWindowsUpdate
            $updates = Get-WUList -MicrosoftUpdate
            if ($updates) {{
                Save-WUUpdates -Updates $updates -DirectoryPath "{download_dir}" -Verbose
            }}
        """)
        
        self.update_progress(50)
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(80)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] Updates downloaded to repo.")
            # Create a manifest of downloaded updates
            self._create_update_manifest(download_dir)
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def _create_update_manifest(self, download_dir):
        """Create a JSON manifest of downloaded updates"""
        manifest_path = os.path.join(self.repo_path, "updates_manifest.json")
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            Get-WUList -MicrosoftUpdate | Select-Object Title, KB, Size, Status | ConvertTo-Json
        """)
        
        out, err, code = self.run_powershell(cmd)
        if out and not err:
            try:
                updates = json.loads(out)
                with open(manifest_path, 'w') as f:
                    json.dump(updates, f, indent=2)
                self.log("[INFO] Created update manifest.")
            except:
                self.log("[WARNING] Could not create update manifest.")

    def install_updates(self):
        threading.Thread(target=self._install_online_thread, daemon=True).start()

    def _install_online_thread(self):
        self.log("[INSTALL] Installing updates online...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        # FIXED: Pre-check for available updates to avoid exit code 1 on empty list
        check_cmd = "Import-Module PSWindowsUpdate; $updates = Get-WUList -MicrosoftUpdate; if ($updates) { $updates | ConvertTo-Json } else { '[]' }"
        out, err, code = self.run_powershell(check_cmd)
        
        try:
            updates_list = json.loads(out)
            if not updates_list or len(updates_list) == 0:
                self.log("[OK] No updates available to install. System is up to date.")
                self.update_progress(100)
                time.sleep(1)
                self.update_progress(0)
                return
        except json.JSONDecodeError:
            self.log(f"[ERROR] Failed to parse update list: {err}")
            self.update_progress(0)
            return
        
        self.log(f"[INFO] Found {len(updates_list)} updates to install. Proceeding...")
        self.update_progress(50)
        
        reboot_param = "-AutoReboot" if self.config.get("auto_reboot", False) else ""
        cmd = f"Import-Module PSWindowsUpdate; Install-WUUpdates -MicrosoftUpdate -AcceptAll {reboot_param} -Verbose"
        
        # Don't capture output for installation to see real-time progress
        _, err, code = self.run_powershell(cmd, capture_output=False)
        
        self.update_progress(80)
        if code != 0:
            self.log(f"[ERROR] Installation failed with exit code {code}. Check log for details.")
            if err:
                self.log(f"[DETAIL] {err}")
        else:
            self.log("[DONE] Online updates installed successfully.")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def install_offline(self):
        threading.Thread(target=self._install_offline_thread, daemon=True).start()

    def _install_offline_thread(self):
        self.log(f"[OFFLINE] Applying updates from {self.repo_path}...")
        self.update_progress(10)
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        if not os.path.exists(download_dir) or not os.listdir(download_dir):
            self.log("[ERROR] No updates found in repository. Please download updates first.")
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        # Install all .msu files in the download directory
        msu_files = [f for f in os.listdir(download_dir) if f.endswith('.msu')]
        
        if not msu_files:
            self.log("[ERROR] No .msu files found in repository. Please download Windows updates.")
            self.update_progress(0)
            return
            
        self.log(f"[INFO] Found {len(msu_files)} update files to install.")
        
        self.update_progress(40)
        success_count = 0
        
        for i, msu_file in enumerate(msu_files):
            msu_path = os.path.join(download_dir, msu_file)
            self.log(f"[INSTALL] Installing {msu_file}...")
            
            try:
                # Use DISM to install the update
                result = subprocess.run(
                    ["dism", "/online", "/add-package", f"/packagepath:{msu_path}", "/quiet", "/norestart"],
                    capture_output=True, text=True, timeout=600
                )
                
                if result.returncode == 0:
                    self.log(f"[OK] Successfully installed {msu_file}")
                    success_count += 1
                else:
                    self.log(f"[ERROR] Failed to install {msu_file}: {result.stderr}")
                    
            except subprocess.TimeoutExpired:
                self.log(f"[ERROR] Timeout installing {msu_file}")
            except Exception as e:
                self.log(f"[ERROR] Error installing {msu_file}: {str(e)}")
                
            # Update progress
            progress = 40 + (i * 50 / len(msu_files))
            self.update_progress(int(progress))
        
        self.update_progress(90)
        self.log(f"[SUMMARY] Successfully installed {success_count} of {len(msu_files)} updates.")
        
        if success_count > 0 and self.config.get("auto_reboot", False):
            self.log("[INFO] System will reboot in 30 seconds to complete installation.")
            subprocess.run(["shutdown", "/r", "/t", "30"])
        
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    # ---------- Office ----------
    def update_office(self):
        threading.Thread(target=self._update_office_thread, daemon=True).start()

    def _update_office_thread(self):
        self.log("[OFFICE] Running Office Click-to-Run updater...")
        self.update_progress(30)
        
        # Check multiple possible paths for OfficeC2RClient
        possible_paths = [
            r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
            r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
        ]
        
        office_path = None
        for path in possible_paths:
            if os.path.exists(path):
                office_path = path
                break
        
        if not office_path:
            self.log("[ERROR] Office Click-to-Run client not found. Is Office installed?")
            self.update_progress(0)
            return
            
        self.update_progress(60)
        try:
            result = subprocess.run([office_path, "/update", "user"], 
                                  capture_output=True, text=True, timeout=1200)
            
            self.update_progress(90)
            if result.returncode == 0:
                self.log("[DONE] Office updated successfully.")
            else:
                self.log(f"[ERROR] Office update failed: {result.stderr}")
                
        except subprocess.TimeoutExpired:
            self.log("[ERROR] Office update timed out.")
        except Exception as e:
            self.log(f"[ERROR] Office update error: {str(e)}")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    # ---------- Extras ----------
    def update_dotnet(self):
        if not self.ensure_module():
            self.log("[ERROR] Cannot update .NET without PSWindowsUpdate module.")
            return
        self.log("[.NET] Triggering .NET update (using Windows Update).")
        # FIXED: Use dedent for clean parsing; filter by title for .NET-specific updates; conditional install
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            $updates = Get-WUList -MicrosoftUpdate | Where-Object { $_.Title -like '* .NET*' }
            if ($updates) {
                Install-WUUpdates -Updates $updates -AcceptAll -Verbose
            } else {
                Write-Output 'No .NET updates available.'
            }
        """)
        threading.Thread(target=self._run_powershell_thread, args=("Updating .NET", cmd), daemon=True).start()

    def update_vcredist(self):
        self.log("[VC++] Updating Visual C++ Redistributables...")
        # Check if winget is available
        try:
            subprocess.run(["where", "winget"], check=True, capture_output=True)
            # Use winget to update VC++ redistributables
            cmd = "winget upgrade --id Microsoft.VCRedist.* --silent --accept-package-agreements --accept-source-agreements"
            threading.Thread(target=self._run_powershell_thread, args=("Updating VC++ Redistributables", cmd), daemon=True).start()
        except (subprocess.CalledProcessError, FileNotFoundError):
            self.log("[INFO] Winget not found. Using alternative method for VC++ Redistributables.")
            # Alternative method using PowerShell
            cmd = textwrap.dedent("""
                $ProgressPreference = 'SilentlyContinue'
                Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vc_redist.x64.exe" -OutFile "$env:TEMP\\vc_redist.x64.exe"
                Start-Process -Wait -FilePath "$env:TEMP\\vc_redist.x64.exe" -ArgumentList "/install", "/quiet", "/norestart"
                Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vc_redist.x86.exe" -OutFile "$env:TEMP\\vc_redist.x86.exe"
                Start-Process -Wait -FilePath "$env:TEMP\\vc_redist.x86.exe" -ArgumentList "/install", "/quiet", "/norestart"
            """)
            threading.Thread(target=self._run_powershell_thread, args=("Updating VC++ Redistributables", cmd), daemon=True).start()

    def _run_powershell_thread(self, description, command):
        self.log(f"[INFO] {description}...")
        self.update_progress(30)
        
        out, err, code = self.run_powershell(command)
        
        self.update_progress(80)
        if code != 0:
            self.log(f"[ERROR] {err if err else 'Command failed'}")
        else:
            self.log(f"[DONE] {description} completed.")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateManagerApp(root)
    root.mainloop()

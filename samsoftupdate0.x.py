#!/usr/bin/env python3
"""
Samsoft Update Manager CE (Community Edition) - Fixed Version
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
import textwrap
import queue
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
        except Exception as e:
            print(f"Error loading config: {e}")
            return DEFAULT_CONFIG
    return DEFAULT_CONFIG

def save_config(config):
    """Save configuration to file"""
    try:
        os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

# ---------- GUI Application ----------
class UpdateManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Samsoft Update Manager CE")
        self.root.geometry("900x600")
        self.root.configure(bg="#1e1e1e")
        
        # Thread-safe message queue for GUI updates
        self.message_queue = queue.Queue()
        
        self.config = load_config()
        self.repo_path = self.config.get("repo_path", REPO_DIR)
        self.status_var = tk.StringVar(value="Idle.")
        self.progress_var = tk.IntVar(value=0)
        self.pswindowsupdate_available = False
        self.running_operation = False  # Prevent multiple simultaneous operations
        
        self.create_ui()
        self.check_pswindowsupdate()
        self.process_message_queue()

    def create_ui(self):
        # Apply dark theme style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#1e1e1e")
        style.configure("TLabelFrame", background="#1e1e1e", foreground="#ffffff")
        style.configure("TLabelFrame.Label", background="#1e1e1e", foreground="#ffffff")
        style.configure("TButton", background="#3c3c3c", foreground="#ffffff")
        style.configure("TCheckbutton", background="#1e1e1e", foreground="#ffffff")
        style.configure("TLabel", background="#1e1e1e", foreground="#ffffff")
        
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Left panel for controls
        left_frame = ttk.LabelFrame(main_frame, text="Controls")
        left_frame.pack(side="left", fill="y", padx=(0, 10))
        
        # Right panel for log
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)
        
        # Control buttons - store references for enabling/disabling
        self.buttons = []
        
        btn_check = ttk.Button(left_frame, text="Check Online", command=self.check_updates)
        btn_check.pack(pady=5, fill="x")
        self.buttons.append(btn_check)
        
        btn_download = ttk.Button(left_frame, text="Download to Repo", command=self.download_updates)
        btn_download.pack(pady=5, fill="x")
        self.buttons.append(btn_download)
        
        btn_install = ttk.Button(left_frame, text="Install Online", command=self.install_updates)
        btn_install.pack(pady=5, fill="x")
        self.buttons.append(btn_install)
        
        btn_offline = ttk.Button(left_frame, text="Install Offline (Repo)", command=self.install_offline)
        btn_offline.pack(pady=5, fill="x")
        self.buttons.append(btn_offline)
        
        ttk.Separator(left_frame, orient="horizontal").pack(fill="x", pady=10)
        
        btn_office = ttk.Button(left_frame, text="Update Office (C2R)", command=self.update_office)
        btn_office.pack(pady=5, fill="x")
        self.buttons.append(btn_office)
        
        btn_dotnet = ttk.Button(left_frame, text="Update .NET", command=self.update_dotnet)
        btn_dotnet.pack(pady=5, fill="x")
        self.buttons.append(btn_dotnet)
        
        btn_vcredist = ttk.Button(left_frame, text="Update VC++ Redists", command=self.update_vcredist)
        btn_vcredist.pack(pady=5, fill="x")
        self.buttons.append(btn_vcredist)
        
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
        
        ttk.Button(settings_frame, text="Clear Log", 
                  command=self.clear_log).pack(fill="x", pady=5)
        
        # Progress bar
        progress_frame = ttk.Frame(left_frame)
        progress_frame.pack(fill="x", pady=10)
        
        ttk.Label(progress_frame, text="Progress:").pack(anchor="w")
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate')
        self.progress_bar.pack(fill="x", pady=5)
        
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
        
        # Initial log message
        self.log("[INFO] Samsoft Update Manager CE initialized.")
        self.log(f"[INFO] Repository path: {self.repo_path}")

    # ---------- Thread-safe GUI updates ----------
    def process_message_queue(self):
        """Process messages from threads to update GUI safely"""
        try:
            while True:
                msg_type, msg_data = self.message_queue.get_nowait()
                
                if msg_type == "log":
                    self._log_safe(msg_data)
                elif msg_type == "progress":
                    self._update_progress_safe(msg_data)
                elif msg_type == "status":
                    self.status_var.set(msg_data)
                elif msg_type == "enable_buttons":
                    self._set_buttons_state(True)
                elif msg_type == "disable_buttons":
                    self._set_buttons_state(False)
                    
        except queue.Empty:
            pass
        finally:
            # Schedule the next check
            self.root.after(100, self.process_message_queue)

    def _log_safe(self, msg):
        """Thread-safe log update"""
        timestamp = time.strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {msg}"
        self.log_text.insert("end", formatted_msg + "\n")
        self.log_text.see("end")
        self.status_var.set(msg)

    def _update_progress_safe(self, value):
        """Thread-safe progress update"""
        self.progress_var.set(value)

    def _set_buttons_state(self, enabled):
        """Enable or disable all buttons"""
        state = "normal" if enabled else "disabled"
        for button in self.buttons:
            button.configure(state=state)
        self.running_operation = not enabled

    # ---------- Public methods for thread use ----------
    def log(self, msg):
        """Queue a log message (thread-safe)"""
        self.message_queue.put(("log", msg))

    def update_progress(self, value):
        """Queue a progress update (thread-safe)"""
        self.message_queue.put(("progress", value))

    def set_status(self, msg):
        """Queue a status update (thread-safe)"""
        self.message_queue.put(("status", msg))

    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.delete(1.0, "end")
        self.log("[INFO] Log cleared.")

    # ---------- Helpers ----------
    def run_powershell(self, command, capture_output=True, timeout=3600):
        """Execute PowerShell command with improved error handling"""
        try:
            # Clean up the command
            clean_command = textwrap.dedent(command).strip()
            
            # Build PowerShell arguments
            ps_args = [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy", "Bypass",
                "-Command", clean_command
            ]
            
            if capture_output:
                completed = subprocess.run(
                    ps_args,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                    shell=False
                )
                stdout = completed.stdout if completed.stdout else ""
                stderr = completed.stderr if completed.stderr else ""
                return stdout.strip(), stderr.strip(), completed.returncode
            else:
                # For long-running commands, don't capture output
                completed = subprocess.run(
                    ps_args,
                    timeout=timeout,
                    shell=False
                )
                return "", "", completed.returncode
                
        except subprocess.TimeoutExpired:
            return "", f"Command timed out after {timeout} seconds", 1
        except Exception as e:
            return "", f"Error executing PowerShell: {str(e)}", 1

    def check_pswindowsupdate(self):
        """Check if PSWindowsUpdate module is available"""
        self.log("[INFO] Checking for PSWindowsUpdate module...")
        check_cmd = "Get-Module -ListAvailable -Name PSWindowsUpdate | Select-Object -First 1 | Format-List"
        out, err, code = self.run_powershell(check_cmd, timeout=30)
        
        if out and "Name" in out and "PSWindowsUpdate" in out:
            self.pswindowsupdate_available = True
            self.log("[OK] PSWindowsUpdate module is available.")
        else:
            self.log("[INFO] PSWindowsUpdate module not found. It will be installed when needed.")
            self.pswindowsupdate_available = False

    def ensure_module(self):
        """Ensure PSWindowsUpdate module is installed"""
        if self.pswindowsupdate_available:
            return True
            
        self.log("[INFO] Installing PSWindowsUpdate module...")
        
        # Try to install the module with improved error handling
        install_cmd = """
            $ErrorActionPreference = 'Stop'
            try {
                # Set repository as trusted
                if (!(Get-PSRepository -Name PSGallery).Trusted) {
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                }
                
                # Install NuGet provider if needed
                $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
                if (-not $nuget) {
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
                }
                
                # Install module
                if (!(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
                    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers
                }
                
                # Import module to verify
                Import-Module PSWindowsUpdate -ErrorAction Stop
                Write-Output "SUCCESS: Module installed and imported"
            }
            catch {
                Write-Error $_.Exception.Message
            }
        """
        
        out, err, code = self.run_powershell(install_cmd, timeout=120)
        
        if "SUCCESS" in out:
            self.pswindowsupdate_available = True
            self.log("[OK] PSWindowsUpdate installed successfully.")
            return True
        else:
            self.log(f"[ERROR] Failed to install PSWindowsUpdate: {err}")
            return False

    def toggle_auto_reboot(self):
        """Toggle auto-reboot setting"""
        self.config["auto_reboot"] = self.auto_reboot_var.get()
        save_config(self.config)
        self.log(f"[INFO] Auto-reboot {'enabled' if self.auto_reboot_var.get() else 'disabled'}")

    def change_repo_path(self):
        """Change repository path"""
        new_path = filedialog.askdirectory(initialdir=self.repo_path, title="Select Repository Directory")
        if new_path:
            self.repo_path = new_path
            self.config["repo_path"] = new_path
            save_config(self.config)
            self.log(f"[INFO] Repository path changed to: {new_path}")

    def _operation_wrapper(self, operation_func):
        """Wrapper to handle button state and exceptions"""
        if self.running_operation:
            self.log("[WARNING] Another operation is already running. Please wait.")
            return
            
        try:
            self.message_queue.put(("disable_buttons", None))
            operation_func()
        except Exception as e:
            self.log(f"[ERROR] Operation failed: {str(e)}")
        finally:
            self.message_queue.put(("enable_buttons", None))

    # ---------- Windows Update Operations ----------
    def check_updates(self):
        """Check for available updates"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._check_updates_thread), daemon=True).start()

    def _check_updates_thread(self):
        """Thread function for checking updates"""
        self.log("[CHECK] Searching online for updates...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.log("[ERROR] Cannot proceed without PSWindowsUpdate module.")
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        # Get update list with better error handling
        cmd = """
            Import-Module PSWindowsUpdate -ErrorAction Stop
            $updates = Get-WUList -MicrosoftUpdate -Verbose
            if ($updates) {
                $updates | Format-Table -Property Title, KB, Size, Status -AutoSize | Out-String
            } else {
                Write-Output "No updates available"
            }
        """
        
        out, err, code = self.run_powershell(cmd, timeout=180)
        
        self.update_progress(70)
        
        if code == 0:
            if "No updates available" in out or not out.strip():
                self.log("[OK] System is up to date.")
            else:
                self.log("[FOUND] Available updates:")
                # Split output into lines for better formatting
                for line in out.split('\n'):
                    if line.strip():
                        self.log(line.strip())
        else:
            if "0x80240024" in err:  # No updates available error code
                self.log("[OK] System is up to date.")
            else:
                self.log(f"[ERROR] Failed to check updates: {err}")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def download_updates(self):
        """Download updates to repository"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._download_thread), daemon=True).start()

    def _download_thread(self):
        """Thread function for downloading updates"""
        self.log(f"[DOWNLOAD] Saving updates to {self.repo_path}...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.log("[ERROR] Cannot proceed without PSWindowsUpdate module.")
            self.update_progress(0)
            return
            
        self.update_progress(20)
        
        # Create download directory
        download_dir = os.path.join(self.repo_path, "Downloads")
        os.makedirs(download_dir, exist_ok=True)
        
        self.update_progress(30)
        
        # Download updates with better error handling
        cmd = f"""
            Import-Module PSWindowsUpdate -ErrorAction Stop
            $downloadPath = '{download_dir}'
            
            # Get list of available updates
            $updates = Get-WUList -MicrosoftUpdate
            
            if ($updates) {{
                Write-Output "Found $($updates.Count) updates to download"
                
                # Download each update
                foreach ($update in $updates) {{
                    try {{
                        Write-Output "Downloading: $($update.Title)"
                        $update | Save-WUUpdates -DirectoryPath $downloadPath -Verbose
                    }}
                    catch {{
                        Write-Warning "Failed to download: $($update.Title)"
                    }}
                }}
                Write-Output "Download complete"
            }} else {{
                Write-Output "No updates to download"
            }}
        """
        
        self.update_progress(50)
        out, err, code = self.run_powershell(cmd, timeout=1800)  # 30 minutes timeout
        
        self.update_progress(80)
        
        if out:
            for line in out.split('\n'):
                if line.strip():
                    self.log(line.strip())
        
        if code == 0:
            self.log("[DONE] Updates downloaded to repository.")
            self._create_update_manifest(download_dir)
        else:
            self.log(f"[ERROR] Download failed: {err}")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def _create_update_manifest(self, download_dir):
        """Create a JSON manifest of downloaded updates"""
        try:
            manifest_path = os.path.join(self.repo_path, "updates_manifest.json")
            
            # Get list of downloaded files
            files = []
            if os.path.exists(download_dir):
                for file in os.listdir(download_dir):
                    if file.endswith(('.msu', '.cab', '.exe')):
                        file_path = os.path.join(download_dir, file)
                        files.append({
                            "filename": file,
                            "size": os.path.getsize(file_path),
                            "modified": os.path.getmtime(file_path)
                        })
            
            # Save manifest
            manifest = {
                "download_date": time.strftime("%Y-%m-%d %H:%M:%S"),
                "files": files,
                "count": len(files)
            }
            
            with open(manifest_path, 'w') as f:
                json.dump(manifest, f, indent=2)
                
            self.log(f"[INFO] Created update manifest with {len(files)} files.")
            
        except Exception as e:
            self.log(f"[WARNING] Could not create update manifest: {str(e)}")

    def install_updates(self):
        """Install updates online"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._install_online_thread), daemon=True).start()

    def _install_online_thread(self):
        """Thread function for installing updates online"""
        self.log("[INSTALL] Installing updates online...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.log("[ERROR] Cannot proceed without PSWindowsUpdate module.")
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        # Check for available updates first
        check_cmd = """
            Import-Module PSWindowsUpdate -ErrorAction Stop
            $updates = Get-WUList -MicrosoftUpdate
            if ($updates) {
                Write-Output "$($updates.Count)"
            } else {
                Write-Output "0"
            }
        """
        
        out, err, code = self.run_powershell(check_cmd, timeout=120)
        
        try:
            update_count = int(out.strip()) if out.strip().isdigit() else 0
        except:
            update_count = 0
        
        if update_count == 0:
            self.log("[OK] No updates available to install. System is up to date.")
            self.update_progress(100)
            time.sleep(1)
            self.update_progress(0)
            return
        
        self.log(f"[INFO] Found {update_count} updates to install. This may take some time...")
        self.update_progress(50)
        
        # Install updates
        reboot_param = "-AutoReboot" if self.config.get("auto_reboot", False) else ""
        install_cmd = f"""
            Import-Module PSWindowsUpdate -ErrorAction Stop
            Install-WindowsUpdate -MicrosoftUpdate -AcceptAll {reboot_param} -Verbose -Confirm:$false
        """
        
        # Run installation without capturing output for real-time progress
        self.log("[INFO] Installing updates... (This may take a while)")
        out, err, code = self.run_powershell(install_cmd, capture_output=False, timeout=7200)  # 2 hours timeout
        
        self.update_progress(90)
        
        if code == 0:
            self.log("[DONE] Online updates installed successfully.")
            if self.config.get("auto_reboot", False):
                self.log("[INFO] System will reboot automatically if required.")
        else:
            self.log(f"[ERROR] Installation completed with issues. Some updates may have failed.")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def install_offline(self):
        """Install updates from repository"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._install_offline_thread), daemon=True).start()

    def _install_offline_thread(self):
        """Thread function for installing updates offline"""
        self.log(f"[OFFLINE] Installing updates from {self.repo_path}...")
        self.update_progress(10)
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        
        if not os.path.exists(download_dir):
            self.log("[ERROR] Download directory does not exist. Please download updates first.")
            self.update_progress(0)
            return
        
        # Find all update files
        update_files = []
        for ext in ['.msu', '.cab']:
            update_files.extend([f for f in os.listdir(download_dir) if f.endswith(ext)])
        
        if not update_files:
            self.log("[ERROR] No update files found in repository. Please download updates first.")
            self.update_progress(0)
            return
            
        self.log(f"[INFO] Found {len(update_files)} update files to install.")
        self.update_progress(30)
        
        success_count = 0
        skip_count = 0
        fail_count = 0
        
        for i, update_file in enumerate(update_files):
            update_path = os.path.join(download_dir, update_file)
            self.log(f"[{i+1}/{len(update_files)}] Installing {update_file}...")
            
            try:
                # Use DISM for .cab files, wusa for .msu files
                if update_file.endswith('.cab'):
                    result = subprocess.run(
                        ["dism.exe", "/online", "/add-package", f"/packagepath:{update_path}", "/quiet", "/norestart"],
                        capture_output=True, text=True, timeout=600
                    )
                else:  # .msu file
                    result = subprocess.run(
                        ["wusa.exe", update_path, "/quiet", "/norestart"],
                        capture_output=True, text=True, timeout=600
                    )
                
                if result.returncode == 0:
                    self.log(f"  [OK] Successfully installed")
                    success_count += 1
                elif result.returncode == 2359302:  # Already installed
                    self.log(f"  [SKIP] Already installed")
                    skip_count += 1
                else:
                    self.log(f"  [ERROR] Failed with code {result.returncode}")
                    fail_count += 1
                    
            except subprocess.TimeoutExpired:
                self.log(f"  [ERROR] Installation timeout")
                fail_count += 1
            except Exception as e:
                self.log(f"  [ERROR] {str(e)}")
                fail_count += 1
                
            # Update progress
            progress = 30 + (i * 60 / len(update_files))
            self.update_progress(int(progress))
        
        self.update_progress(90)
        
        # Summary
        self.log("=" * 50)
        self.log(f"[SUMMARY] Installation complete:")
        self.log(f"  - Successful: {success_count}")
        self.log(f"  - Skipped (already installed): {skip_count}")
        self.log(f"  - Failed: {fail_count}")
        
        if success_count > 0:
            if self.config.get("auto_reboot", False):
                self.log("[INFO] System will reboot in 30 seconds to complete installation.")
                subprocess.run(["shutdown", "/r", "/t", "30", "/c", "Rebooting to complete update installation"])
            else:
                self.log("[INFO] Please reboot to complete installation.")
        
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    # ---------- Additional Update Operations ----------
    def update_office(self):
        """Update Microsoft Office"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._update_office_thread), daemon=True).start()

    def _update_office_thread(self):
        """Thread function for updating Office"""
        self.log("[OFFICE] Checking for Office Click-to-Run...")
        self.update_progress(30)
        
        # Check multiple possible paths for OfficeC2RClient
        possible_paths = [
            r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
            r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
            r"C:\Program Files\Microsoft Office\root\Office16\OfficeC2RClient.exe"
        ]
        
        office_path = None
        for path in possible_paths:
            if os.path.exists(path):
                office_path = path
                self.log(f"[INFO] Found Office at: {path}")
                break
        
        if not office_path:
            self.log("[ERROR] Office Click-to-Run client not found. Is Office installed?")
            self.update_progress(0)
            return
            
        self.update_progress(60)
        
        try:
            self.log("[INFO] Starting Office update process...")
            
            # First, try to update silently
            result = subprocess.run(
                [office_path, "/update", "user", "/displaylevel", "false"],
                capture_output=True, text=True, timeout=1200  # 20 minutes
            )
            
            self.update_progress(90)
            
            if result.returncode == 0:
                self.log("[DONE] Office updated successfully.")
            else:
                # Try alternative update method
                self.log("[INFO] Trying alternative update method...")
                result = subprocess.run(
                    [office_path, "/update", "user"],
                    capture_output=True, text=True, timeout=1200
                )
                
                if result.returncode == 0:
                    self.log("[DONE] Office updated successfully.")
                else:
                    self.log(f"[WARNING] Office update completed with code: {result.returncode}")
                    
        except subprocess.TimeoutExpired:
            self.log("[ERROR] Office update timed out after 20 minutes.")
        except Exception as e:
            self.log(f"[ERROR] Office update error: {str(e)}")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def update_dotnet(self):
        """Update .NET Framework"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._update_dotnet_thread), daemon=True).start()

    def _update_dotnet_thread(self):
        """Thread function for updating .NET"""
        self.log("[.NET] Checking for .NET Framework updates...")
        self.update_progress(20)
        
        if not self.ensure_module():
            self.log("[ERROR] Cannot update .NET without PSWindowsUpdate module.")
            self.update_progress(0)
            return
        
        self.update_progress(40)
        
        # Search for .NET updates
        cmd = """
            Import-Module PSWindowsUpdate -ErrorAction Stop
            $updates = Get-WUList -MicrosoftUpdate | Where-Object { 
                $_.Title -match '\\.NET' -or 
                $_.Title -match 'Framework' 
            }
            
            if ($updates) {
                Write-Output "Found $($updates.Count) .NET updates:"
                $updates | ForEach-Object { Write-Output "  - $($_.Title)" }
                
                Write-Output ""
                Write-Output "Installing .NET updates..."
                Install-WindowsUpdate -Updates $updates -AcceptAll -Verbose -Confirm:$false
                Write-Output "Installation complete"
            } else {
                Write-Output "No .NET updates available"
            }
        """
        
        self.update_progress(60)
        out, err, code = self.run_powershell(cmd, timeout=1800)  # 30 minutes
        
        self.update_progress(90)
        
        if out:
            for line in out.split('\n'):
                if line.strip():
                    self.log(line.strip())
        
        if code == 0:
            if "No .NET updates available" in out:
                self.log("[OK] .NET Framework is up to date.")
            else:
                self.log("[DONE] .NET updates installed successfully.")
        else:
            self.log(f"[ERROR] .NET update failed: {err}")
            
        self.update_progress(100)
        time.sleep(1)
        self.update_progress(0)

    def update_vcredist(self):
        """Update Visual C++ Redistributables"""
        if not self.running_operation:
            threading.Thread(target=lambda: self._operation_wrapper(self._update_vcredist_thread), daemon=True).start()

    def _update_vcredist_thread(self):
        """Thread function for updating VC++ Redistributables"""
        self.log("[VC++] Updating Visual C++ Redistributables...")
        self.update_progress(20)
        
        # Check if winget is available
        winget_available = False
        try:
            result = subprocess.run(["where", "winget"], capture_output=True, text=True)
            winget_available = result.returncode == 0
        except:
            winget_available = False
        
        self.update_progress(40)
        
        if winget_available:
            self.log("[INFO] Using winget to update VC++ Redistributables...")
            
            # Get list of installed VC++ redistributables
            list_cmd = "winget list --id Microsoft.VCRedist"
            result = subprocess.run(list_cmd, shell=True, capture_output=True, text=True, timeout=60)
            
            if result.stdout:
                self.log("[INFO] Installed VC++ Redistributables:")
                for line in result.stdout.split('\n'):
                    if "Microsoft.VCRedist" in line:
                        self.log(f"  - {line.strip()}")
            
            self.update_progress(60)
            
            # Update all VC++ redistributables
            update_cmd = "winget upgrade --id Microsoft.VCRedist --all --silent --accept-package-agreements --accept-source-agreements"
            result = subprocess.run(update_cmd, shell=True, capture_output=True, text=True, timeout=600)
            
            if result.returncode == 0:
                self.log("[DONE] VC++ Redistributables updated via winget.")
            else:
                self.log("[WARNING] Some VC++ updates may have failed.")
                
        else:
            self.log("[INFO] Winget not found. Downloading latest VC++ Redistributables...")
            
            # Download and install latest VC++ redistributables
            vcredist_urls = {
                "x64": "https://aka.ms/vs/17/release/vc_redist.x64.exe",
                "x86": "https://aka.ms/vs/17/release/vc_redist.x86.exe",
                "ARM64": "https://aka.ms/vs/17/release/vc_redist.arm64.exe"
            }
            
            self.update_progress(60)
            
            for arch, url in vcredist_urls.items():
                if arch == "ARM64" and not self._is_arm64():
                    continue
                    
                self.log(f"[INFO] Downloading VC++ {arch}...")
                
                download_cmd = f"""
                    $ProgressPreference = 'SilentlyContinue'
                    $url = '{url}'
                    $output = Join-Path $env:TEMP 'vc_redist.{arch}.exe'
                    
                    try {{
                        Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing
                        Write-Output "Downloaded to: $output"
                        
                        # Install silently
                        Start-Process -FilePath $output -ArgumentList '/install', '/quiet', '/norestart' -Wait
                        Write-Output "Installed {arch} redistributable"
                        
                        # Clean up
                        Remove-Item $output -Force -ErrorAction SilentlyContinue
                    }}
                    catch {{
                        Write-Error "Failed to download/install {arch}: $_"
                    }}
                """
                
                out, err, code = self.run_powershell(download_cmd, timeout=300)
                
                if code == 0:
                    self.log(f"  [OK] {arch} redistributable updated")
                else:
                    self.log(f"  [ERROR] Failed to update {arch}: {err}")
            
        self.update_progress(100)
        self.log("[DONE] VC++ Redistributables update complete.")
        time.sleep(1)
        self.update_progress(0)

    def _is_arm64(self):
        """Check if system is ARM64"""
        try:
            import platform
            return platform.machine().endswith('ARM64')
        except:
            return False


# ---------- Main Entry Point ----------
def main():
    """Main application entry point"""
    try:
        root = tk.Tk()
        app = UpdateManagerApp(root)
        
        # Handle window close event
        def on_closing():
            if app.running_operation:
                if messagebox.askokcancel("Quit", "An operation is running. Do you want to quit anyway?"):
                    root.destroy()
            else:
                root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Start the main loop
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Error", f"Application error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

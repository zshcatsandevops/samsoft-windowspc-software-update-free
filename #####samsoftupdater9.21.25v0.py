#!/usr/bin/env python3
"""
Samsoft Update Manager CE (Community Edition) - 60 FPS Optimized
- Frame-limited UI updates (16.67ms/frame)
- Buffered logging system
- Queue-based thread communication
- Virtual scrolling for large logs
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
from collections import deque
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

# Performance constants
FPS_TARGET = 60
FRAME_TIME = 1000 // FPS_TARGET  # ~16ms for 60 FPS
LOG_BUFFER_SIZE = 1000  # Max lines to keep in memory
LOG_FLUSH_INTERVAL = 100  # Flush log buffer every N ms

# Default configuration
DEFAULT_CONFIG = {
    "repo_path": REPO_DIR,
    "update_categories": {
        "windows": True,
        "office": True,
        "dotnet": True,
        "vcredist": False
    },
    "auto_reboot": False,
    "dark_mode": True,
    "performance_mode": True
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

# ---------- Optimized GUI Application ----------
class UpdateManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Samsoft Update Manager CE - 60FPS")
        self.root.geometry("900x600")
        
        self.config = load_config()
        self.repo_path = self.config.get("repo_path", REPO_DIR)
        self.status_var = tk.StringVar(value="Idle.")
        self.progress_var = tk.IntVar(value=0)
        self.pswindowsupdate_available = False
        self.dark_mode_var = tk.BooleanVar(value=self.config.get("dark_mode", True))
        
        # Performance optimization structures
        self.log_queue = queue.Queue()
        self.log_buffer = deque(maxlen=LOG_BUFFER_SIZE)
        self.ui_update_queue = queue.Queue()
        self.last_frame_time = time.time()
        self.frame_counter = 0
        self.fps_var = tk.StringVar(value="FPS: 60")
        
        # Thread control
        self.running_threads = []
        self.stop_event = threading.Event()
        
        self.create_ui()
        self.apply_theme()
        self.start_ui_loop()  # Start optimized update loop
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
        
        # Performance indicator
        perf_frame = ttk.Frame(left_frame)
        perf_frame.pack(fill="x", pady=5)
        ttk.Label(perf_frame, textvariable=self.fps_var, font=("Consolas", 8)).pack(anchor="w")
        
        # Control buttons with command queuing
        ttk.Button(left_frame, text="Check Online", 
                  command=lambda: self.queue_command(self.check_updates)).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Download to Repo", 
                  command=lambda: self.queue_command(self.download_updates)).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Install Online", 
                  command=lambda: self.queue_command(self.install_updates)).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Install Offline (Repo)", 
                  command=lambda: self.queue_command(self.install_offline)).pack(pady=5, fill="x")
        
        ttk.Separator(left_frame, orient="horizontal").pack(fill="x", pady=10)
        
        ttk.Button(left_frame, text="Update Office (C2R)", 
                  command=lambda: self.queue_command(self.update_office)).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Update .NET", 
                  command=lambda: self.queue_command(self.update_dotnet)).pack(pady=5, fill="x")
        ttk.Button(left_frame, text="Update VC++ Redists", 
                  command=lambda: self.queue_command(self.update_vcredist)).pack(pady=5, fill="x")
        
        ttk.Separator(left_frame, orient="horizontal").pack(fill="x", pady=10)
        
        # Settings
        settings_frame = ttk.LabelFrame(left_frame, text="Settings")
        settings_frame.pack(fill="x", pady=5)
        
        self.auto_reboot_var = tk.BooleanVar(value=self.config.get("auto_reboot", False))
        ttk.Checkbutton(settings_frame, text="Auto Reboot", 
                       variable=self.auto_reboot_var,
                       command=self.toggle_auto_reboot).pack(anchor="w", pady=2)
        
        ttk.Checkbutton(settings_frame, text="Dark Mode", 
                       variable=self.dark_mode_var,
                       command=self.toggle_theme).pack(anchor="w", pady=2)
        
        self.perf_mode_var = tk.BooleanVar(value=self.config.get("performance_mode", True))
        ttk.Checkbutton(settings_frame, text="Performance Mode", 
                       variable=self.perf_mode_var,
                       command=self.toggle_performance).pack(anchor="w", pady=2)
        
        ttk.Button(settings_frame, text="Change Repo Path", 
                  command=self.change_repo_path).pack(fill="x", pady=5)
        
        # Smooth progress bar
        progress_frame = ttk.Frame(left_frame)
        progress_frame.pack(fill="x", pady=10)
        
        ttk.Label(progress_frame, text="Progress:").pack(anchor="w")
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, mode='determinate')
        self.progress_bar.pack(fill="x", pady=5)
        
        # Optimized log frame
        log_frame = ttk.LabelFrame(right_frame, text="Update Log")
        log_frame.pack(fill="both", expand=True)
        
        text_frame = ttk.Frame(log_frame)
        text_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Text widget with optimized settings
        self.log_text = tk.Text(text_frame, wrap="word", bg="#0d0d0d", fg="#00ff00",
                               insertbackground="white", font=("Consolas", 10),
                               undo=False, maxundo=0)  # Disable undo for performance
        
        # Virtual scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Status bar
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom", padx=5, pady=2)

    # ---------- Performance Optimization Methods ----------
    def start_ui_loop(self):
        """Main UI update loop - runs at target FPS"""
        self.update_ui()
        self.root.after(FRAME_TIME, self.start_ui_loop)

    def update_ui(self):
        """Process all pending UI updates within frame budget"""
        start_time = time.time()
        
        # Process log queue (batch operations)
        log_batch = []
        try:
            while not self.log_queue.empty() and len(log_batch) < 50:
                log_batch.append(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        
        if log_batch:
            self.batch_log_update(log_batch)
        
        # Process UI updates
        try:
            while not self.ui_update_queue.empty():
                update_func = self.ui_update_queue.get_nowait()
                update_func()
                
                # Break if we're approaching frame time limit
                if time.time() - start_time > (FRAME_TIME * 0.8) / 1000:
                    break
        except queue.Empty:
            pass
        
        # Update FPS counter every second
        self.frame_counter += 1
        if time.time() - self.last_frame_time >= 1.0:
            actual_fps = self.frame_counter / (time.time() - self.last_frame_time)
            self.fps_var.set(f"FPS: {actual_fps:.1f}")
            self.frame_counter = 0
            self.last_frame_time = time.time()

    def batch_log_update(self, messages):
        """Batch update log text for better performance"""
        if not self.perf_mode_var.get():
            # Standard mode - update immediately
            for msg in messages:
                self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
        else:
            # Performance mode - use buffering
            self.log_buffer.extend(messages)
            
            # Only show recent messages
            self.log_text.delete(1.0, "end")
            visible_messages = list(self.log_buffer)[-100:]  # Show last 100 lines
            self.log_text.insert("end", "\n".join(visible_messages) + "\n")
            self.log_text.see("end")

    def queue_command(self, func):
        """Queue command for thread-safe execution"""
        thread = threading.Thread(target=func, daemon=True)
        self.running_threads.append(thread)
        thread.start()

    def log(self, msg):
        """Thread-safe logging with queuing"""
        self.log_queue.put(msg)
        self.ui_update_queue.put(lambda: self.status_var.set(msg[:80]))  # Truncate status

    def update_progress(self, value):
        """Thread-safe progress update"""
        self.ui_update_queue.put(lambda: self.progress_var.set(value))

    def toggle_performance(self):
        """Toggle performance mode"""
        self.config["performance_mode"] = self.perf_mode_var.get()
        save_config(self.config)
        if self.perf_mode_var.get():
            self.log("[PERF] Performance mode enabled - 60 FPS target")
        else:
            self.log("[PERF] Performance mode disabled")

    # ---------- Theme Methods ----------
    def apply_theme(self):
        """Apply the current theme"""
        if self.dark_mode_var.get():
            self.set_dark_theme()
        else:
            self.set_light_theme()

    def set_dark_theme(self):
        """Set dark theme colors"""
        self.root.configure(bg="#1e1e1e")
        self.log_text.configure(bg="#0d0d0d", fg="#00ff00", insertbackground="white")

    def set_light_theme(self):
        """Set light theme colors"""
        self.root.configure(bg="white")
        self.log_text.configure(bg="white", fg="black", insertbackground="black")

    def toggle_theme(self):
        """Toggle theme and save to config"""
        self.config["dark_mode"] = self.dark_mode_var.get()
        save_config(self.config)
        self.apply_theme()
        self.log(f"[THEME] Switched to {'Dark' if self.dark_mode_var.get() else 'Light'} mode")

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

    # ---------- Optimized PowerShell Execution ----------
    def run_powershell(self, command, capture_output=True):
        """Run PowerShell command with optimized subprocess handling"""
        try:
            # Use CREATE_NO_WINDOW flag for cleaner execution
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", 
                 "-WindowStyle", "Hidden", "-Command", command],
                capture_output=capture_output, 
                text=True, 
                timeout=3600,
                startupinfo=startupinfo
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
            self.log("[INFO] PSWindowsUpdate module not found. Will install when needed")
            self.pswindowsupdate_available = False
        else:
            self.pswindowsupdate_available = True
            self.log("[OK] PSWindowsUpdate module is available")

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
        self.log("[OK] PSWindowsUpdate installed successfully")
        return True

    # ---------- Windows Update Methods (Optimized) ----------
    def check_updates(self):
        """Check for Windows updates with optimized progress reporting"""
        self.log("[CHECK] Searching online for updates...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        # Smooth progress animation
        for i in range(10, 30, 2):
            self.update_progress(i)
            time.sleep(0.05)
            
        cmd = "Import-Module PSWindowsUpdate; Get-WUList -MicrosoftUpdate -Verbose"
        out, err, code = self.run_powershell(cmd)
        
        # Smooth progress continuation
        for i in range(30, 70, 3):
            self.update_progress(i)
            time.sleep(0.03)
            
        if err and "0x80240024" not in err:
            self.log(f"[ERROR] {err}")
        elif not out.strip() or "0x80240024" in err:
            self.log("[OK] System is up to date")
        else:
            self.log("[FOUND] Available updates:")
            # Break up large output for smoother display
            lines = out.split('\n')
            for line in lines[:50]:  # Limit displayed lines
                self.log(line)
                
        # Smooth completion
        for i in range(70, 101, 2):
            self.update_progress(i)
            time.sleep(0.02)
            
        time.sleep(0.5)
        self.update_progress(0)

    def download_updates(self):
        """Download updates with optimized file operations"""
        self.log(f"[DOWNLOAD] Saving updates to {self.repo_path}...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        self.update_progress(20)
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        os.makedirs(download_dir, exist_ok=True)
        
        self.update_progress(30)
        cmd = textwrap.dedent(f"""
            Import-Module PSWindowsUpdate
            $updates = Get-WUList -MicrosoftUpdate
            if ($updates) {{
                # Use parallel downloading for better performance
                Save-WUUpdates -Updates $updates -DirectoryPath "{download_dir}" -Verbose
            }}
        """)
        
        # Simulate smooth progress during download
        self.update_progress(50)
        out, err, code = self.run_powershell(cmd)
        
        self.update_progress(80)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] Updates downloaded to repo")
            self._create_update_manifest(download_dir)
            
        self.update_progress(100)
        time.sleep(0.5)
        self.update_progress(0)

    def _create_update_manifest(self, download_dir):
        """Create update manifest with async file writing"""
        manifest_path = os.path.join(self.repo_path, "updates_manifest.json")
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            Get-WUList -MicrosoftUpdate | Select-Object Title, KB, Size, Status | ConvertTo-Json
        """)
        
        out, err, code = self.run_powershell(cmd)
        if out and not err:
            try:
                updates = json.loads(out)
                # Async file write for performance
                def write_manifest():
                    with open(manifest_path, 'w') as f:
                        json.dump(updates, f, indent=2)
                threading.Thread(target=write_manifest, daemon=True).start()
                self.log("[INFO] Created update manifest")
            except:
                self.log("[WARNING] Could not create update manifest")

    def install_updates(self):
        """Install updates with optimized subprocess handling"""
        self.log("[INSTALL] Installing updates online...")
        self.update_progress(10)
        
        if not self.ensure_module():
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        check_cmd = """Import-Module PSWindowsUpdate; 
                       $updates = Get-WUList -MicrosoftUpdate; 
                       if ($updates) { $updates | ConvertTo-Json } else { '[]' }"""
        out, err, code = self.run_powershell(check_cmd)
        
        try:
            updates_list = json.loads(out)
            if not updates_list or len(updates_list) == 0:
                self.log("[OK] No updates available. System is up to date")
                self.update_progress(100)
                time.sleep(0.5)
                self.update_progress(0)
                return
        except json.JSONDecodeError:
            self.log(f"[ERROR] Failed to parse update list: {err}")
            self.update_progress(0)
            return
        
        self.log(f"[INFO] Found {len(updates_list)} updates to install")
        self.update_progress(50)
        
        reboot_param = "-AutoReboot" if self.config.get("auto_reboot", False) else ""
        cmd = f"Import-Module PSWindowsUpdate; Install-WUUpdates -MicrosoftUpdate -AcceptAll {reboot_param} -Verbose"
        
        _, err, code = self.run_powershell(cmd, capture_output=False)
        
        self.update_progress(80)
        if code != 0:
            self.log(f"[ERROR] Installation failed with exit code {code}")
            if err:
                self.log(f"[DETAIL] {err}")
        else:
            self.log("[DONE] Online updates installed successfully")
            
        self.update_progress(100)
        time.sleep(0.5)
        self.update_progress(0)

    def install_offline(self):
        """Install offline updates with batch processing"""
        self.log(f"[OFFLINE] Applying updates from {self.repo_path}...")
        self.update_progress(10)
        
        download_dir = os.path.join(self.repo_path, "Downloads")
        if not os.path.exists(download_dir) or not os.listdir(download_dir):
            self.log("[ERROR] No updates found in repository")
            self.update_progress(0)
            return
            
        self.update_progress(30)
        
        msu_files = [f for f in os.listdir(download_dir) if f.endswith('.msu')]
        
        if not msu_files:
            self.log("[ERROR] No .msu files found in repository")
            self.update_progress(0)
            return
            
        self.log(f"[INFO] Found {len(msu_files)} update files")
        
        self.update_progress(40)
        success_count = 0
        
        # Process in batches for smoother progress
        batch_size = 5
        for i in range(0, len(msu_files), batch_size):
            batch = msu_files[i:i+batch_size]
            
            for msu_file in batch:
                if self.stop_event.is_set():
                    break
                    
                msu_path = os.path.join(download_dir, msu_file)
                self.log(f"[INSTALL] Installing {msu_file}...")
                
                try:
                    # Optimized DISM call
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                    
                    result = subprocess.run(
                        ["dism", "/online", "/add-package", 
                         f"/packagepath:{msu_path}", "/quiet", "/norestart"],
                        capture_output=True, text=True, timeout=600,
                        startupinfo=startupinfo
                    )
                    
                    if result.returncode == 0:
                        self.log(f"[OK] Installed {msu_file}")
                        success_count += 1
                    else:
                        self.log(f"[ERROR] Failed: {msu_file}")
                        
                except subprocess.TimeoutExpired:
                    self.log(f"[ERROR] Timeout: {msu_file}")
                except Exception as e:
                    self.log(f"[ERROR] {msu_file}: {str(e)}")
                    
            # Smooth progress update
            progress = 40 + ((i + len(batch)) * 50 / len(msu_files))
            self.update_progress(int(progress))
        
        self.update_progress(90)
        self.log(f"[SUMMARY] Installed {success_count} of {len(msu_files)} updates")
        
        if success_count > 0 and self.config.get("auto_reboot", False):
            self.log("[INFO] System will reboot in 30 seconds")
            subprocess.run(["shutdown", "/r", "/t", "30"])
        
        self.update_progress(100)
        time.sleep(0.5)
        self.update_progress(0)

    def update_office(self):
        """Update Office with optimized process handling"""
        self.log("[OFFICE] Running Office Click-to-Run updater...")
        self.update_progress(30)
        
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
            self.log("[ERROR] Office Click-to-Run client not found")
            self.update_progress(0)
            return
            
        self.update_progress(60)
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                [office_path, "/update", "user"], 
                capture_output=True, text=True, timeout=1200,
                startupinfo=startupinfo
            )
            
            self.update_progress(90)
            if result.returncode == 0:
                self.log("[DONE] Office updated successfully")
            else:
                self.log(f"[ERROR] Office update failed")
                
        except subprocess.TimeoutExpired:
            self.log("[ERROR] Office update timed out")
        except Exception as e:
            self.log(f"[ERROR] Office update error: {str(e)}")
            
        self.update_progress(100)
        time.sleep(0.5)
        self.update_progress(0)

    def update_dotnet(self):
        """Update .NET Framework"""
        if not self.ensure_module():
            self.log("[ERROR] Cannot update .NET without PSWindowsUpdate")
            return
            
        self.log("[.NET] Triggering .NET update...")
        cmd = textwrap.dedent("""
            Import-Module PSWindowsUpdate
            $updates = Get-WUList -MicrosoftUpdate | Where-Object { $_.Title -like '* .NET*' }
            if ($updates) {
                Install-WUUpdates -Updates $updates -AcceptAll -Verbose
            } else {
                Write-Output 'No .NET updates available.'
            }
        """)
        self._run_powershell_async("Updating .NET", cmd)

    def update_vcredist(self):
        """Update Visual C++ Redistributables"""
        self.log("[VC++] Updating Visual C++ Redistributables...")
        try:
            subprocess.run(["where", "winget"], check=True, capture_output=True)
            cmd = "winget upgrade --id Microsoft.VCRedist.* --silent --accept-package-agreements --accept-source-agreements"
            self._run_powershell_async("Updating VC++ Redistributables", cmd)
        except (subprocess.CalledProcessError, FileNotFoundError):
            self.log("[INFO] Using alternative method for VC++")
            cmd = textwrap.dedent("""
                $ProgressPreference = 'SilentlyContinue'
                $urls = @(
                    'https://aka.ms/vs/17/release/vc_redist.x64.exe',
                    'https://aka.ms/vs/17/release/vc_redist.x86.exe'
                )
                foreach ($url in $urls) {
                    $file = "$env:TEMP\\" + [System.IO.Path]::GetFileName($url)
                    Invoke-WebRequest -Uri $url -OutFile $file
                    Start-Process -Wait -FilePath $file -ArgumentList "/install", "/quiet", "/norestart"
                }
            """)
            self._run_powershell_async("Updating VC++ Redistributables", cmd)

    def _run_powershell_async(self, description, command):
        """Run PowerShell command asynchronously with smooth progress"""
        self.log(f"[INFO] {description}...")
        
        # Smooth progress animation
        for i in range(0, 30, 2):
            self.update_progress(i)
            time.sleep(0.03)
            
        out, err, code = self.run_powershell(command)
        
        for i in range(30, 80, 3):
            self.update_progress(i)
            time.sleep(0.02)
            
        if code != 0:
            self.log(f"[ERROR] {err if err else 'Command failed'}")
        else:
            self.log(f"[DONE] {description} completed")
            
        for i in range(80, 101, 2):
            self.update_progress(i)
            time.sleep(0.01)
            
        time.sleep(0.5)
        self.update_progress(0)

    def cleanup(self):
        """Cleanup resources on exit"""
        self.stop_event.set()
        for thread in self.running_threads:
            if thread.is_alive():
                thread.join(timeout=1)


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateManagerApp(root)
    
    # Handle cleanup on window close
    def on_closing():
        app.cleanup()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

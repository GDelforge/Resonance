import os
import sys
import subprocess
import time
import winreg
import shutil
import ctypes
from pathlib import Path

# We can safely import these because Summon.ps1 ensured they exist
from rich.console import Console
from rich.prompt import Confirm
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.panel import Panel
from rich.style import Style
from rich.theme import Theme
from rich.traceback import install
from rich.table import Table
from rich.align import Align

# --- THEME CONFIGURATION ---
custom_theme = Theme({
    "info": "dim white",
    "warning": "yellow",
    "error": "bold red",
    "success": "bold green",
    "arcane": "magenta",
    "spell": "italic violet",
})
console = Console(theme=custom_theme)

# Install rich traceback handler for prettier error debugging
install(show_locals=True)

class Incantator:
    def __init__(self):
        self.user = os.environ.get('USERNAME')
        self.script_root = Path(__file__).parent
        self.data_path = Path(r"C:\Data")
        self.share_name = "Data$"
        self.drive_letter = "R:"
        self.drive_label = "Codex"

    def banner(self):
        console.clear()
        title = Panel(f"[arcane]~~~ THE GRAND CONJURATION (PYTHON EDITION) ~~~[/arcane]\n[dim]Apprentice: {self.user}[/dim]", border_style="magenta", padding=(1, 2))
        console.print(Align.center(title))
        time.sleep(1.5)

    def run_ps(self, cmd, description=None):
        """Executes a raw PowerShell command."""
        if description:
            console.print(f"[info]  > {description}...[/info]")
        
        full_cmd = ["powershell", "-NoProfile", "-Command", cmd]
        result = subprocess.run(full_cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            console.print(f"[error]  ! Spell failed: {result.stderr.strip()}[/error]")
            return False
        return True

    def set_reg_key(self, path, name, value, reg_type=winreg.REG_DWORD):
        """Sets a registry key value safely."""
        try:
            # Handle shorthand
            root_map = {"HKCU": winreg.HKEY_CURRENT_USER, "HKLM": winreg.HKEY_LOCAL_MACHINE}
            root_str, subkey = path.split(":\\", 1)
            root = root_map.get(root_str, winreg.HKEY_CURRENT_USER)

            # Attempt to open/create with specific write permission (KEY_SET_VALUE)
            # This avoids Access Denied on HKLM where full control is restricted
            try:
                key = winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE)
            except FileNotFoundError:
                key = winreg.CreateKeyEx(root, subkey, 0, winreg.KEY_SET_VALUE)

            with key:
                winreg.SetValueEx(key, name, 0, reg_type, value)
            return True
        except PermissionError:
            console.print(f"[error]  ! Access Denied: {path} (Run as Admin)[/error]")
            return False
        except Exception as e:
            console.print(f"[error]Registry Error: {e}[/error]")
            return False

    def step_fonts(self):
        """Installs fonts by leveraging the Shell.Application COM object via PS wrapper."""
        console.rule("[arcane]Step 1: Inscribing Glyphs[/arcane]")
        time.sleep(1)
        console.print("[info]This step will download and install 'Nunito' and 'Fira Code' fonts to your system.[/info]")
        if Confirm.ask("[spell]Ancient Glyphs (Fonts) seem vital. Inscribe them?[/spell]"):
            font_urls = {
                "Nunito": "https://www.1001fonts.com/download/nunito.zip",
                "FiraCode": "https://github.com/tonsky/FiraCode/releases/download/6.2/Fira_Code_v6.2.zip"
            }
            
            # Using a simplified PS script inside Python to handle the heavy lifting of downloading/zipping/shell-copying
            # This is often cleaner than re-implementing zip extraction and COM objects in ctypes
            ps_block = """
            $Fonts = @{
                'Nunito'    = 'https://www.1001fonts.com/download/nunito.zip'
                'Fira Code' = 'https://github.com/tonsky/FiraCode/releases/download/6.2/Fira_Code_v6.2.zip'
            }
            $FontTemp = "$env:TEMP\\CustomFonts"
            New-Item $FontTemp -ItemType Directory -Force | Out-Null
            $Shell = New-Object -ComObject Shell.Application
            $FontsFolder = $Shell.Namespace(0x14)
            
            $RegLM = Get-ItemProperty "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" -ErrorAction SilentlyContinue
            $RegCU = Get-ItemProperty "HKCU:\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts" -ErrorAction SilentlyContinue
            $Installed = @()
            if ($RegLM) { $Installed += $RegLM.PSObject.Properties.Name }
            if ($RegCU) { $Installed += $RegCU.PSObject.Properties.Name }

            foreach ($Key in $Fonts.Keys) {
                # Simple heuristic check to see if font is already registered
                if ($Installed -match $Key) {
                    Write-Host "Glyph $Key is already inscribed." -ForegroundColor DarkGray
                } else {
                    Write-Host "Downloading $Key..."
                    $Zip = "$FontTemp\\$Key.zip"
                    Invoke-WebRequest $Fonts[$Key] -OutFile $Zip
                    Expand-Archive $Zip -Dest "$FontTemp\\$Key" -Force
                    $FontFiles = Get-ChildItem "$FontTemp\\$Key" -Include *.ttf,*.otf -Recurse
                    foreach ($File in $FontFiles) {
                        $FontsFolder.CopyHere($File.FullName, 0x14)
                    }
                }
            }
            """
            
            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), transient=True) as progress:
                progress.add_task("[cyan]Scribing Glyphs...", total=None)
                self.run_ps(ps_block)
            console.print("[success]  + Glyphs inscribed.[/success]")
            time.sleep(1)

    def step_share_and_drive(self):
        """Sets up the Data folder and Maps R:"""
        console.rule("[arcane]Step 2: Setting up the Codex[/arcane]")
        time.sleep(1)

        # Part A: The Codex (Data Folder)
        console.print(f"[info]This step will create the directory '{self.data_path}' and share it as '{self.share_name}' with full access for user '{self.user}'.[/info]")
        if Confirm.ask(f"[spell]The Codex ({self.data_path}). Manage it?[/spell]"):
            if not self.data_path.exists():
                console.print(f"[info]The Codex ({self.data_path}) is hidden. Revealing...[/info]")
                self.data_path.mkdir(exist_ok=True)
                # Share it
                self.run_ps(f"New-SmbShare -Name '{self.share_name}' -Path '{self.data_path}' -FullAccess '{self.user}' -Description 'Data Repository' -ErrorAction SilentlyContinue", "Creating SMB Share")
            else:
                console.print(f"[dim]The Codex ({self.data_path}) stands ready.[/dim]")

        # Part B: The Astral Gateway (Drive Mapping)
        console.print(f"[info]This step will map the drive letter '{self.drive_letter}' to the local share '\\\\localhost\\{self.share_name}' and label it '{self.drive_label}'.[/info]")
        if Confirm.ask(f"[spell]The Astral Gateway ({self.drive_letter}). Bind it?[/spell]"):
            if not Path(self.drive_letter).exists():
                console.print(f"[info]The Astral Gateway ({self.drive_letter}) is closed. Binding...[/info]")
                cmd = f"""
                New-SmbMapping -LocalPath {self.drive_letter} -RemotePath "\\\\localhost\\{self.share_name}" -Persistent $true
                $RegPath = "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MountPoints2\\##localhost#{self.share_name}"
                if (-not (Test-Path $RegPath)) {{ New-Item -Path $RegPath -Force | Out-Null }}
                Set-ItemProperty -Path $RegPath -Name "_LabelFromReg" -Value "{self.drive_label}"
                """
                if self.run_ps(cmd, "Binding Drive"):
                    console.print(f"[success]  + Gateway bound as {self.drive_label}.[/success]")
            else:
                console.print(f"[dim]The Astral Gateway ({self.drive_letter}) is already open.[/dim]")
        time.sleep(1)

    def step_software(self):
        """Installs software via Winget."""
        softwares = {
            "AutoHotkey": "AutoHotkey.AutoHotkey",
            "FFmpeg": "Gyan.FFmpeg",
            "Mullvad VPN": "MullvadVPN.MullvadVPN",
            "Obsidian": "Obsidian.Obsidian",
            "PowerToys": "Microsoft.PowerToys",
            "qBittorrent": "qBittorrent.qBittorrent",
            "VS Code": "Microsoft.VisualStudioCode",
            "Git": "Git.Git",
            "Python": "Python.Python.3.12",
            "Node.js": "OpenJS.NodeJS",
            "Windows Terminal": "Microsoft.WindowsTerminal",
            "Espanso": "Espanso.Espanso"
        }

        console.rule("[arcane]Step 3: Summoning Instruments (Winget)[/arcane]")
        time.sleep(1)
        
        # Dynamic description
        software_list = ", ".join(softwares.keys())
        console.print(f"[info]This step will check for and install the following software using Winget: {software_list}.[/info]")
        
        if Confirm.ask("[spell]Shall we summon these instruments?[/spell]"):
            console.print("[info]Aligning planetary bodies for software download...[/info]")
            time.sleep(2)
            
            # Check if Winget exists
            if shutil.which("winget") is None:
                console.print("[error]Winget is missing from the realm.[/error]")
                return

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                transient=True
            ) as progress:
                # Optimization: Fetch installed list once to avoid calling winget 10+ times
                progress.add_task("Consulting the Archives (Reading installed packages)...", total=None)
                installed_blob = ""
                try:
                    # We fetch the raw list. It's faster to check string presence than to query individually.
                    # Using utf-8 and ignoring errors to handle potential character encoding issues in app names.
                    proc = subprocess.run(["winget", "list"], capture_output=True, text=True, encoding="utf-8", errors="ignore")
                    if proc.returncode == 0:
                        installed_blob = proc.stdout
                except Exception:
                    pass # Fallback to individual checks if this fails

                task_id = progress.add_task("Summoning...", total=len(softwares))
                
                for name, pkg_id in softwares.items():
                    progress.update(task_id, description=f"Seeking presence of {name}...")
                    
                    # Check if the ID is in the bulk list, or fallback to individual check if list failed
                    if pkg_id not in installed_blob:
                        progress.update(task_id, description=f"[yellow]Summoning {name}...[/yellow]")
                        install = subprocess.run(
                            ["winget", "install", "-e", "--id", pkg_id, "--silent", "--accept-source-agreements", "--accept-package-agreements"],
                            capture_output=True
                        )
                        if install.returncode == 0:
                            console.print(f"[success]  + {name} summoned.[/success]")
                        else:
                            console.print(f"[error]  ! Failed to summon {name}.[/error]")
                    else:
                        console.print(f"[dim]  . {name} already present.[/dim]")

    def step_windows_settings(self):
        """Configures Windows UI settings via Registry."""
        console.rule("[arcane]Step 4: Shaping the Apparatus (Settings)[/arcane]")
        time.sleep(1)
        console.print("[info]This step will configure Windows Explorer (hidden files, file extensions), set Dark Mode, and enable mapped drives for elevated processes.[/info]")
        
        if Confirm.ask("[spell]Shall we shape the apparatus?[/spell]"):
            settings = [
                # Explorer
                ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "ShowTaskViewButton", 0),
                ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "Hidden", 1), # Show Hidden
                ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "HideFileExt", 0), # Show Ext
                # Dark Mode
                ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "AppsUseLightTheme", 0),
                ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "SystemUsesLightTheme", 0),
                # Enable Mapped Drives for Elevated Token (Fixes R: drive visibility)
                ("HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System", "EnableLinkedConnections", 1),
            ]

            table = Table(show_header=True, header_style="bold magenta", box=None)
            table.add_column("Configuration Key")
            table.add_column("Value")

            for path, key, val in settings:
                self.set_reg_key(path, key, val)
                table.add_row(key, str(val))
            
            console.print(table)
            console.print("[success]  + Visuals aligned to darkness.[/success]")
            time.sleep(0.5)

            # Wallpaper
            bg_path = self.script_root / "Assets" / "background.png"
            if bg_path.exists():
                console.print(f"[info]Found background artifact at {bg_path}.[/info]")
                dest = Path(os.environ["USERPROFILE"]) / "Pictures" / "background.png"
                shutil.copy(bg_path, dest)
                # Python equivalent of SystemParametersInfo for wallpaper
                SPI_SETDESKWALLPAPER = 20
                ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, str(dest), 3)
                console.print("[success]  + Reality (Wallpaper) rewritten.[/success]")
                time.sleep(1)

    def step_gemini(self):
        """Installs Gemini CLI via NPM."""
        console.rule("[arcane]Step 5: Summoning the Oracle (Gemini CLI)[/arcane]")
        time.sleep(1)
        console.print("[info]This step will install the '@google/gemini-cli' package globally using npm.[/info]")
        
        if Confirm.ask("[spell]Shall we summon the Oracle?[/spell]"):
            if shutil.which("npm"):
                # Simple check if package is installed
                check = subprocess.run("npm list -g @google/gemini-cli", shell=True, capture_output=True, text=True)
                if "@google/gemini-cli" not in check.stdout:
                    with console.status("[bold yellow]Npm is chanting..."):
                        subprocess.run("npm install -g @google/gemini-cli", shell=True)
                    console.print("[success]  + The Oracle is ready.[/success]")
                else:
                    console.print("[dim]  . The Oracle waits.[/dim]")
            else:
                console.print("[warning]  ! npm not found. The Oracle cannot be summoned.[/warning]")

    def step_path(self):
        """Adds paths to PATH."""
        console.rule("[arcane]Step 6: Extending the Ley Lines (PATH)[/arcane]")
        time.sleep(1)
        console.print(r"[info]This step will add 'G:\My Drive\Data\Resonance\Spells' to your user PATH environment variable.[/info]")
        
        if Confirm.ask("[spell]Shall we extend the Ley Lines?[/spell]"):
            targets = [
                r"G:\My Drive\Data\Resonance\Spells"
            ]
            
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                    try:
                        path_val, type_ = winreg.QueryValueEx(key, "Path")
                    except FileNotFoundError:
                        path_val = ""
                        type_ = winreg.REG_EXPAND_SZ
                    
                    parts = [p for p in path_val.split(";") if p]
                    existing_lower = [p.lower() for p in parts]
                    modified = False

                    for target in targets:
                        if target.lower() not in existing_lower:
                            parts.append(target)
                            existing_lower.append(target.lower())
                            modified = True
                            console.print(f"[success]  + The path {target} has been woven into the Ley Lines.[/success]")
                        else:
                            console.print(f"[dim]  . The path {target} is already present.[/dim]")

                    if modified:
                        new_path = ";".join(parts)
                        winreg.SetValueEx(key, "Path", 0, type_, new_path)
            except Exception as e:
                console.print(f"[error]  ! Failed to extend Ley Lines: {e}[/error]")

    def step_desktop_cleanse(self):
        """Hides all desktop icons."""
        console.rule("[arcane]Step 7: Cleansing the Surface (Desktop)[/arcane]")
        time.sleep(1)
        console.print("[info]This step will hide all icons on your Desktop by modifying the registry key 'HideIcons'.[/info]")
        if Confirm.ask("[spell]Shall we banish all icons from the Desktop?[/spell]"):
            # HideIcons = 1
            if self.set_reg_key(r"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideIcons", 1):
                console.print("[success]  + The surface has been silenced.[/success]")
            else:
                console.print("[error]  ! Failed to silence the Desktop.[/error]")

    def step_taskbar_renewal(self):
        """Clears the taskbar and pins Windows Terminal using LayoutModification.xml."""
        console.rule("[arcane]Step 8: Forging the Anchor (Taskbar)[/arcane]")
        time.sleep(1)
        console.print("[info]This step will remove all existing pinned items from the Taskbar and pin ONLY the Windows Terminal.[/info]")
        if Confirm.ask("[spell]Shall we clear the Taskbar and anchor only the Terminal?[/spell]"):
            ps_script = r"""
            $ErrorActionPreference = 'SilentlyContinue'
            
            # Path to LayoutModification.xml
            $LayoutPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml"
            
            # Define the XML content for replacing taskbar pins with Windows Terminal
            $XmlContent = @'
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:UWA AppUserModelID="Microsoft.WindowsTerminal_8wekyb3d8bbwe!App" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
'@
            
            # Backup existing layout if it exists
            if (Test-Path $LayoutPath) {
                Copy-Item $LayoutPath "$LayoutPath.bak" -Force
            }
            
            # Write the new XML
            Set-Content -Path $LayoutPath -Value $XmlContent -Encoding UTF8
            
            # Clear existing pins and registry state to force a reload
            $TaskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
            if (Test-Path $TaskbarPath) {
                Remove-Item "$TaskbarPath\*" -Force
            }
            Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -Name "*"
            
            # Restart Explorer to apply changes
            Stop-Process -Name explorer -Force
            Start-Sleep -Seconds 2
            if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) {
                Start-Process explorer
            }
            """
            
            with console.status("[bold magenta]Reforging the Taskbar (LayoutModification.xml)..."):
                self.run_ps(ps_script)
            console.print("[success]  + Taskbar layout applied. (Explorer restarted)[/success]")

    def finalize(self):
        console.rule("[arcane]~~~ INCANTATION COMPLETE ~~~[/arcane]")
        time.sleep(1)
        console.print("[info]This step will restart Windows Explorer to ensure all registry changes and icon settings take effect immediately.[/info]")
        if Confirm.ask("Restart Explorer to apply all sigils?"):
            subprocess.run(["powershell", "-c", "Stop-Process -Name explorer -Force; Start-Process explorer"])

# --- ENTRY POINT ---
if __name__ == "__main__":
    # Ensure Admin privileges for HKLM writes
    if not ctypes.windll.shell32.IsUserAnAdmin():
        console.print("[warning]Elevation required for system modifications. Summoning UAC...[/warning]")
        script_path = os.path.abspath(__file__)
        # ShellExecuteW returns >32 on success
        ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script_path}"', None, 1)
        if ret > 32:
            sys.exit(0)
        console.print("[error]Elevation failed or cancelled. Proceeding with limited power...[/error]")
        time.sleep(2)

    try:
        Incantation = Incantator()
        Incantation.banner()
        
        Incantation.step_fonts()
        Incantation.step_share_and_drive()
        Incantation.step_software()
        Incantation.step_windows_settings()
        Incantation.step_gemini()
        Incantation.step_path()
        Incantation.step_desktop_cleanse()
        Incantation.step_taskbar_renewal()
        
        Incantation.finalize()
    except Exception as e:
        console.print(Panel(f"[bold red]FATAL ERROR[/bold red]\n\n{e}", border_style="red"))
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
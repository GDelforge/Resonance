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

# --- THEME CONFIGURATION ---
custom_theme = Theme({
    "info": "dim cyan",
    "warning": "yellow",
    "error": "bold red",
    "success": "bold green",
    "arcane": "magenta",
    "spell": "italic violet",
})
console = Console(theme=custom_theme)

class Grimoire:
    def __init__(self):
        self.user = os.environ.get('USERNAME')
        self.script_root = Path(__file__).parent
        self.data_path = Path(r"C:\Data")
        self.share_name = "Data$"
        self.drive_letter = "R:"
        self.drive_label = "Codex"

    def banner(self):
        console.clear()
        console.print(Panel(f"[arcane]~~~ THE GRAND CONJURATION (PYTHON EDITION) ~~~[/arcane]\n[dim]Apprentice: {self.user}[/dim]", border_style="magenta"))

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

            with winreg.CreateKey(root, subkey) as key:
                winreg.SetValueEx(key, name, 0, reg_type, value)
            return True
        except Exception as e:
            console.print(f"[error]Registry Error: {e}[/error]")
            return False

    def step_fonts(self):
        """Installs fonts by leveraging the Shell.Application COM object via PS wrapper."""
        if Confirm.ask("[spell]Ancient Glyphs (Fonts) seem vital. Inscribe them?[/spell]"):
            font_urls = {
                "Nunito": "https://www.1001fonts.com/download/nunito.zip",
                "Raleway": "https://www.1001fonts.com/download/raleway.zip",
                "FiraCode": "https://github.com/tonsky/FiraCode/releases/download/6.2/Fira_Code_v6.2.zip"
            }
            
            # Using a simplified PS script inside Python to handle the heavy lifting of downloading/zipping/shell-copying
            # This is often cleaner than re-implementing zip extraction and COM objects in ctypes
            ps_block = """
            $Fonts = @{
                'Nunito'   = 'https://www.1001fonts.com/download/nunito.zip'
                'Raleway'  = 'https://www.1001fonts.com/download/raleway.zip'
                'FiraCode' = 'https://github.com/tonsky/FiraCode/releases/download/6.2/Fira_Code_v6.2.zip'
            }
            $FontTemp = "$env:TEMP\\CustomFonts"
            New-Item $FontTemp -ItemType Directory -Force | Out-Null
            $Shell = New-Object -ComObject Shell.Application
            $FontsFolder = $Shell.Namespace(0x14)

            foreach ($Key in $Fonts.Keys) {
                Write-Host "Downloading $Key..."
                $Zip = "$FontTemp\\$Key.zip"
                Invoke-WebRequest $Fonts[$Key] -OutFile $Zip
                Expand-Archive $Zip -Dest "$FontTemp\\$Key" -Force
                $FontFiles = Get-ChildItem "$FontTemp\\$Key" -Include *.ttf,*.otf -Recurse
                foreach ($File in $FontFiles) {
                    $FontsFolder.CopyHere($File.FullName, 0x14)
                }
            }
            """
            
            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), transient=True) as progress:
                progress.add_task("[cyan]Scribing Glyphs...", total=None)
                self.run_ps(ps_block)
            console.print("[success]  + Glyphs inscribed.[/success]")

    def step_share_and_drive(self):
        """Sets up the Data folder and Maps R:"""
        if not self.data_path.exists():
            if Confirm.ask(f"[spell]The Codex ({self.data_path}) is hidden. Reveal it?[/spell]"):
                self.data_path.mkdir(exist_ok=True)
                # Share it
                self.run_ps(f"New-SmbShare -Name '{self.share_name}' -Path '{self.data_path}' -FullAccess '{self.user}' -Description 'Data Repository' -ErrorAction SilentlyContinue", "Creating SMB Share")

        if not Path(self.drive_letter).exists():
            if Confirm.ask(f"[spell]The Astral Gateway ({self.drive_letter}) is closed. Bind it?[/spell]"):
                cmd = f"""
                New-SmbMapping -LocalPath {self.drive_letter} -RemotePath "\\\\localhost\\{self.share_name}" -Persistent $true
                $RegPath = "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MountPoints2\\##localhost#{self.share_name}"
                if (-not (Test-Path $RegPath)) {{ New-Item -Path $RegPath -Force | Out-Null }}
                Set-ItemProperty -Path $RegPath -Name "_LabelFromReg" -Value "{self.drive_label}"
                """
                if self.run_ps(cmd, "Binding Drive"):
                    console.print(f"[success]  + Gateway bound as {self.drive_label}.[/success]")

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
            "Node.js": "OpenJS.NodeJS"
        }

        console.print("\n[arcane]Step 3: Summoning Instruments (Winget)[/arcane]")
        
        # Check if Winget exists
        if shutil.which("winget") is None:
            console.print("[error]Winget is missing from the realm.[/error]")
            return

        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            transient=True
        ) as progress:
            task_id = progress.add_task("Checking registry...", total=len(softwares))
            
            for name, pkg_id in softwares.items():
                progress.update(task_id, description=f"Seeking presence of {name}...")
                
                # Check installed
                check = subprocess.run(["winget", "list", "-e", "--id", pkg_id], capture_output=True)
                
                if check.returncode != 0:
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
        console.print("\n[arcane]Step 4: Shaping the Apparatus (Settings)[/arcane]")
        
        settings = [
            # Explorer
            ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "ShowTaskViewButton", 0),
            ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "Hidden", 1), # Show Hidden
            ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced", "HideFileExt", 0), # Show Ext
            # Dark Mode
            ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "AppsUseLightTheme", 0),
            ("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "SystemUsesLightTheme", 0),
        ]

        for path, key, val in settings:
            self.set_reg_key(path, key, val)
        
        console.print("[success]  + Visuals aligned to darkness.[/success]")

        # Wallpaper
        bg_path = self.script_root / "background.png"
        if bg_path.exists():
            dest = Path(os.environ["USERPROFILE"]) / "Pictures" / "background.png"
            shutil.copy(bg_path, dest)
            # Python equivalent of SystemParametersInfo for wallpaper
            SPI_SETDESKWALLPAPER = 20
            ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, str(dest), 3)
            console.print("[success]  + Reality (Wallpaper) rewritten.[/success]")

    def step_gemini(self):
        """Installs Gemini CLI via NPM."""
        console.print("\n[arcane]Step 5: Summoning the Oracle[/arcane]")
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

    def finalize(self):
        console.print("\n[arcane]~~~ CONJURATION COMPLETE ~~~[/arcane]")
        if Confirm.ask("Restart Explorer to apply all sigils?"):
            subprocess.run(["powershell", "-c", "Stop-Process -Name explorer -Force; Start-Process explorer"])

# --- ENTRY POINT ---
if __name__ == "__main__":
    mage = Grimoire()
    mage.banner()
    
    mage.step_fonts()
    mage.step_share_and_drive()
    mage.step_software()
    mage.step_windows_settings()
    mage.step_gemini()
    
    mage.finalize()
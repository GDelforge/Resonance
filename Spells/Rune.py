import os
import sys
from pathlib import Path

# Try to import rich, as it is used in the project
try:
    from rich.console import Console
    from rich.prompt import Prompt
    from rich.panel import Panel
    from rich.theme import Theme
except ImportError:
    print("This spell requires the 'rich' library. Please ensure it is installed.")
    sys.exit(1)

# Theme matching Incantation
custom_theme = Theme({
    "info": "dim white",
    "warning": "yellow",
    "error": "bold red",
    "success": "bold green",
    "arcane": "magenta",
    "spell": "italic violet",
})
console = Console(theme=custom_theme)

def get_espanso_dir():
    """Locates the Espanso match directory."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        console.print("[error]Could not find APPDATA environment variable.[/error]")
        return None
    
    # Standard location
    match_path = Path(appdata) / "espanso" / "match"
    
    if not match_path.exists():
        console.print(f"[warning]Espanso match directory not found at: {match_path}[/warning]")
        return None
        
    return match_path

def create_rune():
    console.clear()
    console.rule("[arcane]~~~ RUNE SCRIBE ~~~[/arcane]")
    console.print("[dim]Inscribing new rune into the fabric of Espanso...[/dim]\n")

    match_dir = get_espanso_dir()
    if not match_dir:
        console.print("[error]Espanso match directory not found. Is Espanso installed?[/error]")
        return

    # Select file - defaulting to base.yml
    target_file = match_dir / "base.yml"
    
    # If base.yml doesn't exist, check for others or create it
    if not target_file.exists():
        yml_files = list(match_dir.glob("*.yml"))
        if yml_files:
            # Just pick the first one if base.yml is missing
            target_file = yml_files[0]
            console.print(f"[info]base.yml not found. Targeting {target_file.name} instead.")
        else:
            # Create base.yml
            console.print("[info]No match files found. Creating base.yml...")
            try:
                match_dir.mkdir(parents=True, exist_ok=True)
                with open(target_file, "w", encoding="utf-8") as f:
                    f.write("matches:\n")
            except Exception as e:
                console.print(f"[error]Failed to create match file: {e}[/error]")
                return

    # Gather input
    console.print("[info]Define your new rune:[/info]")
    trigger = Prompt.ask("[spell]  Trigger (what you type)[/spell]")
    replace = Prompt.ask("[spell]  Replacement (what appears)[/spell]")

    if not trigger or not replace:
        console.print("[warning]The incantation was mumbled (empty input). Aborting.[/warning]")
        return

    # Prepare the block to append
    # We append a simple match entry. 
    # NOTE: We assume 'matches:' parent key exists and we are appending to it.
    # We add a newline before to ensure separation.
    block = f"\n  - trigger: \":{trigger}\"\n    replace: \"{replace}\"\n"
    
    try:
        # check if file needs 'matches:' header
        if target_file.stat().st_size == 0:
             with open(target_file, "w", encoding="utf-8") as f:
                f.write("matches:\n")

        with open(target_file, "a", encoding="utf-8") as f:
            f.write(block)
        
        console.print(Panel(
            f"[bold]Trigger:[/bold] {trigger}\n[bold]Replace:[/bold] {replace}",
            title="[success]Rune Inscribed[/success]",
            border_style="green"
        ))
        console.print("[dim]Espanso should reload automatically.[/dim]")
        
    except Exception as e:
        console.print(f"[error]Failed to inscribe rune: {e}[/error]")

if __name__ == "__main__":
    create_rune()

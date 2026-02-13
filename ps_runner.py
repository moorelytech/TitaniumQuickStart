import subprocess
from typing import Optional, Tuple


def run_ps(ps_command: str) -> Tuple[int, str, str]:
    p = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
        capture_output=True,
        text=True,
    )
    return p.returncode, p.stdout, p.stderr


def run_ps_function(ps1_path: str, function_name: str, prelude: Optional[str] = None) -> Tuple[int, str, str]:
    # Dot-source your script, optionally run prelude (set globals), then call the function.
    prelude = prelude or ""
    cmd = (
        f". '{ps1_path}'; "
        f"$ErrorActionPreference='Stop'; "
        f"{prelude} "
        f"{function_name}"
    )
    return run_ps(cmd)

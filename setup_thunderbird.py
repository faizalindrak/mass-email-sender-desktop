#!/usr/bin/env python3
# Setup script for Thunderbird headless MailExtension bridge and Native Messaging host.
# - Creates platform-specific Native Messaging manifest(s)
# - Registers Windows Registry keys (HKCU) for the native host
# - Builds an XPI package for the MailExtension
# - Creates platform-specific launcher scripts for the native host
#
# Usage:
#   python setup_thunderbird.py --install
#   python setup_thunderbird.py --uninstall   (best-effort cleanup)
#   python setup_thunderbird.py --print-info  (prints computed paths)
#
# After installation:
#   - In Thunderbird: Add-ons and Themes -> Tools (gear) -> Debug Add-ons -> Load Temporary Add-on
#     and select thunderbird_extension/manifest.json (or the generated XPI).
#   - Ensure the extension ID matches the allowed_extensions in the native manifest:
#       email-automation@local
#
# Notes:
#   - Queue directory defaults are aligned with ThunderbirdExtensionSender and native_host.py:
#       Windows: %APPDATA%\\EmailAutomation\\tb_queue
#       macOS:   ~/Library/Application Support/EmailAutomation/tb_queue
#       Linux:   ~/.local/share/email_automation/tb_queue
#   - You can optionally set 'tb_queue_dir' in your profile JSON to a custom path.

import os
import sys
import json
import shutil
import zipfile
import stat
from typing import List, Tuple

HOST_NAME = "com.emailautomation.tbhost"
EXTENSION_ID = "email-automation@local"

def project_root() -> str:
    return os.path.abspath(os.path.dirname(__file__))

def extension_dir() -> str:
    return os.path.join(project_root(), "thunderbird_extension")

def native_host_script_path() -> str:
    return os.path.join(extension_dir(), "native_host.py")

def background_js_path() -> str:
    return os.path.join(extension_dir(), "background.js")

def extension_manifest_path() -> str:
    return os.path.join(extension_dir(), "manifest.json")

def default_queue_dir() -> str:
    try:
        if os.name == "nt":
            appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
            return os.path.join(appdata, "EmailAutomation", "tb_queue")
        elif sys.platform == "darwin":
            return os.path.join(os.path.expanduser("~"), "Library", "Application Support", "EmailAutomation", "tb_queue")
        else:
            return os.path.join(os.path.expanduser("~"), ".local", "share", "email_automation", "tb_queue")
    except Exception:
        return os.path.abspath(os.path.join("tb_queue"))

def ensure_dirs(paths: List[str]):
    for p in paths:
        os.makedirs(p, exist_ok=True)

def build_xpi() -> str:
    """Pack the MailExtension into an XPI file for convenience."""
    xpi_path = os.path.join(extension_dir(), "email-automation-bridge.xpi")
    with zipfile.ZipFile(xpi_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # Only include extension files (exclude native host files)
        zf.write(extension_manifest_path(), arcname="manifest.json")
        zf.write(background_js_path(), arcname="background.js")
    return xpi_path

def create_win_launcher(launcher_path: str, host_py: str, python_path: str = None, queue_dir: str = None):
    """Create a BAT launcher that starts the native host Python script.
    Prefer the absolute python executable path detected during setup to avoid PATH issues.
    Also sets TB_QUEUE_DIR so the native host and app share the same queue.
    """
    lines = [
        "@echo off",
        "setlocal",
        "set PYTHONUNBUFFERED=1",
    ]
    if queue_dir:
        # Quote to keep spaces intact
        lines.append(f'set "TB_QUEUE_DIR={queue_dir}"')
    if python_path and os.path.exists(python_path):
        lines.append(f'"{python_path}" -u "%~dp0native_host.py"')
        lines.append("exit /b %ERRORLEVEL%")
    else:
        # Fallback: try 'py' then 'python'
        lines.extend([
            "where py >nul 2>nul",
            "if %ERRORLEVEL%==0 (",
            '  py -3 -u "%~dp0native_host.py"',
            "  exit /b %ERRORLEVEL%",
            ")",
            "where python >nul 2>nul",
            "if %ERRORLEVEL%==0 (",
            '  python -u "%~dp0native_host.py"',
            "  exit /b %ERRORLEVEL%",
            ")",
            "echo Python not found. Please install Python 3 and ensure 'py' or 'python' is on PATH.",
            "exit /b 1",
        ])
    content = "\r\n".join(lines) + "\r\n"
    with open(launcher_path, "w", encoding="utf-8") as f:
        f.write(content)

def create_posix_launcher(launcher_path: str, host_py: str):
    """Create a POSIX shell script launcher for the native host."""
    content = """#!/usr/bin/env bash
set -euo pipefail
export PYTHONUNBUFFERED=1

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
HOST_PY="${SCRIPT_DIR}/native_host.py"

if command -v python3 >/dev/null 2>&1; then
  exec python3 -u "${HOST_PY}"
elif command -v python >/dev/null 2>&1; then
  exec python -u "${HOST_PY}"
else:
  echo "Python 3 not found. Please install Python 3 and ensure it's on PATH." >&2
  exit 1
fi
"""
    with open(launcher_path, "w", encoding="utf-8") as f:
        f.write(content)
    # Make executable
    st = os.stat(launcher_path)
    os.chmod(launcher_path, st.st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

def native_manifest_json(executable_path: str) -> dict:
    """Return a native messaging manifest JSON dict."""
    return {
        "name": HOST_NAME,
        "description": "Email Automation native host for Thunderbird",
        "path": executable_path,
        "type": "stdio",
        # For Gecko-based apps (Thunderbird), allowed_extensions should contain the extension ID
        "allowed_extensions": [EXTENSION_ID],
    }

def install_windows() -> Tuple[List[str], List[str]]:
    """Install native host on Windows: create manifest(s) and registry keys under HKCU."""
    manifest_paths = []
    messages = []

    appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    mozilla_dir = os.path.join(appdata, "Mozilla", "NativeMessagingHosts")
    thunderbird_dir = os.path.join(appdata, "Thunderbird", "NativeMessagingHosts")
    ensure_dirs([mozilla_dir, thunderbird_dir])

    # Use a project-local queue dir to avoid ambiguity
    qdir = os.path.join(project_root(), "tb_queue")
    ensure_dirs([qdir, os.path.join(qdir, "jobs"), os.path.join(qdir, "results")])
    messages.append(f"Ensured queue dir exists: {qdir}")
    messages.append(f"- jobs: {os.path.join(qdir, 'jobs')}")
    messages.append(f"- results: {os.path.join(qdir, 'results')}")
    
    # Create a BAT launcher next to native_host.py, pinning to the current Python interpreter and queue dir
    launcher_path = os.path.join(extension_dir(), "native_host_launcher.bat")
    create_win_launcher(launcher_path, native_host_script_path(), python_path=sys.executable, queue_dir=qdir)

    manifest_data = native_manifest_json(launcher_path)

    for host_dir in (mozilla_dir, thunderbird_dir):
        manifest_file = os.path.join(host_dir, f"{HOST_NAME}.json")
        with open(manifest_file, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=2, ensure_ascii=False)
        manifest_paths.append(manifest_file)
        messages.append(f"Wrote Windows manifest: {manifest_file}")

    # Write Registry keys (HKCU)
    try:
        import winreg  # type: ignore
        for base in ("Software\\Mozilla\\NativeMessagingHosts", "Software\\Thunderbird\\NativeMessagingHosts"):
            key_path = f"{base}\\{HOST_NAME}"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            # Set the default value to the manifest path in the 'Thunderbird' dir (prefer TB key)
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, os.path.join(thunderbird_dir, f"{HOST_NAME}.json"))
            winreg.CloseKey(key)
            messages.append(f"Registry HKCU\\{key_path} -> {thunderbird_dir}\\{HOST_NAME}.json")
    except Exception as e:
        messages.append(f"Warning: Failed to create registry keys: {e}")

    return manifest_paths, messages

def install_linux() -> Tuple[List[str], List[str]]:
    """Install native host on Linux: write manifests under ~/.mozilla and ~/.thunderbird."""
    manifest_paths = []
    messages = []

    mozilla_dir = os.path.join(os.path.expanduser("~"), ".mozilla", "native-messaging-hosts")
    thunderbird_dir = os.path.join(os.path.expanduser("~"), ".thunderbird", "native-messaging-hosts")
    ensure_dirs([mozilla_dir, thunderbird_dir])

    launcher_path = os.path.join(extension_dir(), "native_host_launcher.sh")
    create_posix_launcher(launcher_path, native_host_script_path())

    manifest_data = native_manifest_json(launcher_path)

    for host_dir in (mozilla_dir, thunderbird_dir):
        manifest_file = os.path.join(host_dir, f"{HOST_NAME}.json")
        with open(manifest_file, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=2, ensure_ascii=False)
        manifest_paths.append(manifest_file)
        messages.append(f"Wrote Linux manifest: {manifest_file}")

    return manifest_paths, messages

def install_macos() -> Tuple[List[str], List[str]]:
    """Install native host on macOS: write manifests under Application Support."""
    manifest_paths = []
    messages = []

    mozilla_dir = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Mozilla", "NativeMessagingHosts")
    thunderbird_dir = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Thunderbird", "NativeMessagingHosts")
    ensure_dirs([mozilla_dir, thunderbird_dir])

    launcher_path = os.path.join(extension_dir(), "native_host_launcher.sh")
    create_posix_launcher(launcher_path, native_host_script_path())

    manifest_data = native_manifest_json(launcher_path)

    for host_dir in (mozilla_dir, thunderbird_dir):
        manifest_file = os.path.join(host_dir, f"{HOST_NAME}.json")
        with open(manifest_file, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=2, ensure_ascii=False)
        manifest_paths.append(manifest_file)
        messages.append(f"Wrote macOS manifest: {manifest_file}")

    return manifest_paths, messages

def uninstall_windows() -> List[str]:
    removed = []
    try:
        import winreg  # type: ignore
        for base in ("Software\\Mozilla\\NativeMessagingHosts", "Software\\Thunderbird\\NativeMessagingHosts"):
            key_path = f"{base}\\{HOST_NAME}"
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
                removed.append(f"Deleted registry key HKCU\\{key_path}")
            except FileNotFoundError:
                pass
    except Exception as e:
        removed.append(f"Registry cleanup warning: {e}")

    # Remove manifest files
    appdata = os.getenv("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    for host_dir in (
        os.path.join(appdata, "Mozilla", "NativeMessagingHosts"),
        os.path.join(appdata, "Thunderbird", "NativeMessagingHosts"),
    ):
        manifest_file = os.path.join(host_dir, f"{HOST_NAME}.json")
        if os.path.exists(manifest_file):
            try:
                os.remove(manifest_file)
                removed.append(f"Removed {manifest_file}")
            except Exception:
                pass
    return removed

def uninstall_posix() -> List[str]:
    removed = []
    for host_dir in (
        os.path.join(os.path.expanduser("~"), ".mozilla", "native-messaging-hosts"),
        os.path.join(os.path.expanduser("~"), ".thunderbird", "native-messaging-hosts"),
        os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Mozilla", "NativeMessagingHosts"),
        os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Thunderbird", "NativeMessagingHosts"),
    ):
        manifest_file = os.path.join(host_dir, f"{HOST_NAME}.json")
        if os.path.exists(manifest_file):
            try:
                os.remove(manifest_file)
                removed.append(f"Removed {manifest_file}")
            except Exception:
                pass
    return removed

def print_info():
    print("Project Root:", project_root())
    print("Extension Dir:", extension_dir())
    print("Native Host Script:", native_host_script_path())
    print("Default Queue Dir:", default_queue_dir())
    print("Extension ID:", EXTENSION_ID)
    print("Host Name:", HOST_NAME)

def main():
    args = sys.argv[1:]
    action = "--install" if not args else args[0]

    if action == "--print-info":
        print_info()
        return

    if action == "--install":
        print_info()
        # Build XPI for convenience
        try:
            xpi_path = build_xpi()
            print("Built XPI:", xpi_path)
        except Exception as e:
            print("Warning: Failed to build XPI:", e)

        if os.name == "nt":
            manifests, messages = install_windows()
        elif sys.platform == "darwin":
            manifests, messages = install_macos()
        else:
            manifests, messages = install_linux()

        for m in messages:
            print(m)
        print("Native messaging installation complete.")
        print("")
        print("Next steps:")
        print("1) Open Thunderbird -> Add-ons and Themes -> Tools (gear) -> Debug Add-ons -> Load Temporary Add-on")
        print(f"   Load: {extension_manifest_path()}  (or the XPI at thunderbird_extension/email-automation-bridge.xpi)")
        print("2) Keep Thunderbird running. The extension will connect to the native host.")
        print("3) In this app, set your profile 'email_client' to 'thunderbird'.")
        print("   Optionally set 'tb_queue_dir' in the profile JSON; otherwise default queue dir is used:")
        print(f"   {default_queue_dir()}")
        print("4) Start monitoring; when a file is processed, the app writes a job JSON into the queue.")
        print("   The native host and extension will send the email and write a result JSON back.")
        return

    if action == "--uninstall":
        removed = []
        if os.name == "nt":
            removed += uninstall_windows()
        else:
            removed += uninstall_posix()
        print("Uninstall cleanup complete. Items removed:")
        for r in removed:
            print(" -", r)
        return

    print("Unknown action. Use --install | --uninstall | --print-info")
    sys.exit(1)

if __name__ == "__main__":
    main()
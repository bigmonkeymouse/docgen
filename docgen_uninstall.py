import os
import sys
import shutil
import subprocess
import tempfile


def _is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def _get_install_dir() -> str:
    if _is_frozen_app():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _get_appdata_dir() -> str:
    roaming = os.environ.get("APPDATA")
    if roaming and roaming.strip():
        return os.path.join(roaming, "DocGen")
    return os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "DocGen")


def _message_box(text: str, title: str) -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x00000040)
    except Exception:
        pass


def _remove_uninstall_registry() -> None:
    try:
        import winreg
        uninstall_root = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, uninstall_root, 0, winreg.KEY_ALL_ACCESS) as root:
            winreg.DeleteKey(root, "DocGen")
    except Exception:
        pass


def _schedule_delete_install_dir(install_dir: str) -> None:
    tmp_dir = tempfile.mkdtemp(prefix="docgen_uninstall_")
    cmd_path = os.path.join(tmp_dir, "remove_docgen.cmd")

    script = "\r\n".join(
        [
            "@echo off",
            "ping 127.0.0.1 -n 3 > nul",
            f'rmdir /s /q "{install_dir}"',
            f'rmdir /s /q "{tmp_dir}"',
        ]
    )
    with open(cmd_path, "w", encoding="utf-8") as f:
        f.write(script)

    subprocess.Popen(
        ["cmd.exe", "/c", "start", "", "/min", cmd_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


def main() -> int:
    install_dir = _get_install_dir()
    appdata_dir = _get_appdata_dir()

    try:
        shutil.rmtree(appdata_dir, ignore_errors=True)
    except Exception:
        pass

    _remove_uninstall_registry()

    if os.path.isdir(install_dir):
        try:
            _schedule_delete_install_dir(install_dir)
        except Exception:
            pass

    _message_box("DocGen 已开始卸载。若文件占用导致残留，请关闭程序后重试。", "DocGen 卸载")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

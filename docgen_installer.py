import argparse
import os
import sys
import shutil
import tempfile
import traceback
import zipfile

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QWidget,
)


APP_NAME = "DocGen"
APP_VERSION = "1.0.0"


def _get_local_programs_dir() -> str:
    local = os.environ.get("LOCALAPPDATA")
    if local and local.strip():
        return os.path.join(local, "Programs", APP_NAME)
    return os.path.join(os.path.expanduser("~"), "AppData", "Local", "Programs", APP_NAME)


def _payload_root() -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "payload")


def _find_payload_file(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    for root in (os.path.join(base, "payload"), base):
        p = os.path.join(root, name)
        if os.path.isfile(p):
            return p
    for root, _, files in os.walk(base):
        if name in files:
            return os.path.join(root, name)
    return ""


def _find_payload_dir(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    for root in (os.path.join(base, "payload"), base):
        p = os.path.join(root, name)
        if os.path.isdir(p):
            return p
    for root, dirs, _ in os.walk(base):
        if name in dirs:
            return os.path.join(root, name)
    return ""


def _copy_tree(src_dir: str, dst_dir: str) -> None:
    os.makedirs(dst_dir, exist_ok=True)
    for root, dirs, files in os.walk(src_dir):
        rel = os.path.relpath(root, src_dir)
        rel = "" if rel == "." else rel
        target_root = os.path.join(dst_dir, rel)
        os.makedirs(target_root, exist_ok=True)
        for name in files:
            src_path = os.path.join(root, name)
            dst_path = os.path.join(target_root, name)
            shutil.copy2(src_path, dst_path)
        for d in dirs:
            os.makedirs(os.path.join(target_root, d), exist_ok=True)


def _write_uninstall_registry(install_dir: str) -> None:
    try:
        import winreg

        uninstall_root = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, uninstall_root + "\\" + APP_NAME) as k:
            winreg.SetValueEx(k, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(k, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(k, "Publisher", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(k, "InstallLocation", 0, winreg.REG_SZ, install_dir)
            winreg.SetValueEx(k, "DisplayIcon", 0, winreg.REG_SZ, os.path.join(install_dir, "DocGen.exe"))
            winreg.SetValueEx(k, "UninstallString", 0, winreg.REG_SZ, os.path.join(install_dir, "DocGenUninstall.exe"))
            winreg.SetValueEx(k, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(k, "NoRepair", 0, winreg.REG_DWORD, 1)
    except Exception:
        pass


def perform_install(install_dir: str) -> None:
    app_exe = _find_payload_file("DocGen.exe")
    uninstaller_exe = _find_payload_file("DocGenUninstall.exe")
    readme = _find_payload_file("安装包使用说明.txt")
    templates_zip = _find_payload_file("word_template.zip")

    for p in (app_exe, uninstaller_exe, readme):
        if not p or not os.path.isfile(p):
            raise FileNotFoundError("安装载荷缺失")

    os.makedirs(install_dir, exist_ok=True)
    shutil.copy2(app_exe, os.path.join(install_dir, "DocGen.exe"))
    shutil.copy2(uninstaller_exe, os.path.join(install_dir, "DocGenUninstall.exe"))
    shutil.copy2(readme, os.path.join(install_dir, "安装包使用说明.txt"))
    if templates_zip and os.path.isfile(templates_zip):
        with zipfile.ZipFile(templates_zip, "r") as zf:
            zf.extractall(install_dir)
    _write_uninstall_registry(install_dir)


class InstallerWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} 安装程序")
        self.setMinimumWidth(560)

        self.path_edit = QLineEdit()
        self.path_edit.setText(_get_local_programs_dir())
        self.browse_btn = QPushButton("选择...")
        self.install_btn = QPushButton("一键安装")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)

        title = QLabel(f"{APP_NAME} {APP_VERSION}")
        title.setStyleSheet("font-size: 18px; font-weight: 700;")
        desc = QLabel("请选择安装目录，然后点击“一键安装”。")
        desc.setStyleSheet("color: #555;")

        layout = QGridLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(10)

        layout.addWidget(title, 0, 0, 1, 3)
        layout.addWidget(desc, 1, 0, 1, 3)

        layout.addWidget(QLabel("安装目录："), 2, 0, alignment=Qt.AlignVCenter)
        layout.addWidget(self.path_edit, 2, 1)
        layout.addWidget(self.browse_btn, 2, 2)

        layout.addWidget(self.progress, 3, 0, 1, 3)
        layout.addWidget(self.install_btn, 4, 0, 1, 3)

        self.browse_btn.clicked.connect(self.on_browse)
        self.install_btn.clicked.connect(self.on_install)

    def on_browse(self) -> None:
        current = self.path_edit.text().strip() or _get_local_programs_dir()
        path = QFileDialog.getExistingDirectory(self, "选择安装目录", current)
        if path:
            self.path_edit.setText(path)

    def on_install(self) -> None:
        install_dir = self.path_edit.text().strip()
        if not install_dir:
            QMessageBox.warning(self, "提示", "请选择安装目录。")
            return

        self.install_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.progress.setValue(5)

        try:
            self.progress.setValue(35)
            self.progress.setValue(55)
            perform_install(install_dir)
            self.progress.setValue(100)

            QMessageBox.information(self, "完成", f"{APP_NAME} 已安装完成。")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "安装失败", str(e))
            self.install_btn.setEnabled(True)
            self.browse_btn.setEnabled(True)
            self.progress.setValue(0)


def main() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--silent", action="store_true")
    parser.add_argument("--dir", default="")
    args, qt_args = parser.parse_known_args(sys.argv[1:])

    if args.silent:
        target = args.dir.strip() or _get_local_programs_dir()
        try:
            perform_install(target)
            return 0
        except Exception:
            try:
                log_path = os.path.join(tempfile.gettempdir(), "docgen_installer_error.log")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except Exception:
                pass
            return 1

    app = QApplication([sys.argv[0]] + qt_args)
    win = InstallerWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())


import os
import sys
import json
import subprocess
import datetime
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QSizePolicy,
    QPushButton,
    QProgressBar,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QSize, QUrl
from PySide6.QtGui import QFont, QColor, QPalette, QDesktopServices

# i18n-ready string resources (English keys, localized values)
STR = {
    "APP_WINDOW_TITLE": "合同自动生成工具",
    "APP_HEADER_TITLE": "合同生成助手",
    "APP_HEADER_SUBTITLE": "自动化批量填充 Word 模板与 Excel 数据",
    "SETTINGS_WINDOW_TITLE": "偏好设置",
    "SETTINGS_HEADER_TITLE": "偏好设置",
    "SETTINGS_TEMPLATE_LABEL": "Word 模板文件夹",
    "SETTINGS_BROWSE": "选择...",
    "SETTINGS_OK": "确认",
    "SETTINGS_CANCEL": "取消",
    "DIALOG_SELECT_TEMPLATE_FOLDER_TITLE": "选择 Word 模板文件夹",
    "DIALOG_SELECT_EXCEL_TITLE": "选择数据源 (Excel)",
    "DIALOG_SELECT_EXCEL_FILTER": "Excel Files (*.xlsx)",
    "DIALOG_TEMPLATE_MISSING_TITLE": "模板缺失",
    "DIALOG_TEMPLATE_MISSING_BODY": "未找到 Word 模板文件夹，请确认 word_template 目录存在并包含 .docx 模板文件。",
    "DIALOG_OPEN_FOLDER": "打开目录",
    "WARN_TITLE": "警告",
    "ERROR_TITLE": "错误",
    "ERROR_CORE_SCRIPT_MISSING": "未找到核心填充脚本：template_filler.py",
    "ERROR_TEMPLATE_DIR_INVALID": "模板文件夹路径无效，请在偏好设置中配置。",
    "ERROR_TEMPLATE_DIR_MISSING": "模板文件夹不存在：\n{path}",
    "ERROR_EXCEL_INVALID": "数据源文件无效：\n{path}",
    "ERROR_CREATE_FOLDER_FAILED": "创建文件夹失败：{error}",
    "STATUS_READY_SELECT_SOURCE": "准备就绪，请选择数据源",
    "STATUS_PREPARING": "正在准备环境...",
    "STATUS_RUNNING": "正在处理：{current} / {total}",
    "STATUS_DONE": "任务已完成",
    "STATUS_FAILED": "任务失败",
    "LOG_CONFIG_UPDATED": "配置已更新",
    "LOG_TASK_STARTED": "任务开始...",
    "LOG_TASK_DONE": "所有合同已生成并打包成功！",
    "LOG_TASK_FAILED": "生成过程中出现错误：{error}",
    "LOG_ZIP_EXPORTED": "已成功导出：{name}",
    "BTN_START": "生成文档",
    "BTN_CHOOSE_TEMPLATE": "选择模板",
    "BTN_OPEN_TEMPLATE_LIBRARY": "模板库",
    "BTN_VIEW_LOG": "查看日志",
    "BTN_OPEN_SETTINGS": "偏好设置",
    "LOG_WINDOW_TITLE": "生成日志",
    "LOG_HEADER_TITLE": "任务日志",
    "LOG_CLEAR": "清空日志记录",
    "MSGBOX_TASK_FAILED": "生成失败：{error}",
}

# Modern Apple-inspired Theme (HIG Compliant)
APPLE_STYLE = """
    QMainWindow {
        background-color: #F5F5F7;
    }
    
    QWidget {
        font-size: 12px;
        color: #1D1D1F;
    }

    /* Card-like container */
    QFrame#MainCard {
        background-color: #FFFFFF;
        border-radius: 16px;
        border: 1px solid rgba(0, 0, 0, 0.05);
    }

    /* Primary Button */
    QPushButton#PrimaryBtn {
        background-color: #0071E3;
        color: white;
        border-radius: 12px;
        font-weight: 600;
        font-size: 15px;
        padding: 0 24px;
        min-height: 44px; /* HIG 44pt touch area */
        border: none;
    }
    QPushButton#PrimaryBtn:hover {
        background-color: #0077ED;
    }
    QPushButton#PrimaryBtn:pressed {
        background-color: #0062C3;
    }
    QPushButton#PrimaryBtn:disabled {
        background-color: #D2D2D7;
        color: #86868B;
    }

    /* Secondary/Action Buttons */
    QPushButton#SecondaryBtn {
        background-color: rgba(0, 0, 0, 0.04);
        color: #0071E3;
        border-radius: 10px;
        font-weight: 500;
        padding: 0 16px;
        min-height: 36px;
        border: 1px solid rgba(0, 0, 0, 0.05);
    }
    QPushButton#SecondaryBtn:hover {
        background-color: rgba(0, 0, 0, 0.08);
    }
    QPushButton#SecondaryBtn:pressed {
        background-color: rgba(0, 0, 0, 0.12);
    }

    /* Progress Bar */
    QProgressBar {
        background-color: rgba(0, 0, 0, 0.05);
        border: none;
        border-radius: 6px;
        text-align: center;
        height: 10px;
        font-size: 10px;
        font-weight: 500;
        color: #1D1D1F;
    }
    QProgressBar::chunk {
        background-color: #0071E3;
        border-radius: 6px;
    }

    /* Settings Dialog Inputs */
    QLineEdit {
        background-color: rgba(0, 0, 0, 0.03);
        border: 1px solid rgba(0, 0, 0, 0.1);
        border-radius: 8px;
        padding: 8px 12px;
        selection-background-color: #0071E3;
    }
    QLineEdit:focus {
        border: 2px solid #0071E3;
        background-color: #FFFFFF;
    }

    /* Log Browser */
    QTextBrowser {
        background-color: #FBFBFD;
        border: 1px solid rgba(0, 0, 0, 0.1);
        border-radius: 12px;
        padding: 12px;
        line-height: 1.5;
    }
    
    QLabel {
        font-weight: 500;
        color: #1D1D1F;
    }
"""

class DraggableLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            self.setText(url.toLocalFile())

class SettingsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle(STR["SETTINGS_WINDOW_TITLE"])
        self.setMinimumWidth(640)
        self.setStyleSheet(APPLE_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        header = QLabel(STR["SETTINGS_HEADER_TITLE"])
        header.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(header)

        self.path_container = QFrame()
        self.path_container.setObjectName("MainCard")
        form_layout = QGridLayout(self.path_container)
        form_layout.setContentsMargins(16, 16, 16, 16)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(12)

        label = QLabel(STR["SETTINGS_TEMPLATE_LABEL"])
        label.setStyleSheet("font-size: 12px; color: #86868B;")
        self.word_template_folder_edit = DraggableLineEdit()
        self.word_template_folder_edit.setReadOnly(True)

        browse_btn = QPushButton(STR["SETTINGS_BROWSE"])
        browse_btn.setObjectName("SecondaryBtn")
        browse_btn.setFixedWidth(90)
        browse_btn.clicked.connect(lambda: self.browse_word_template_path(self.word_template_folder_edit))

        form_layout.addWidget(label, 0, 0, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        form_layout.addWidget(self.word_template_folder_edit, 0, 1)
        form_layout.addWidget(browse_btn, 0, 2)
        form_layout.setColumnStretch(1, 1)

        layout.addWidget(self.path_container)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        button_layout.addStretch()

        self.cancel_btn = QPushButton(STR["SETTINGS_CANCEL"])
        self.cancel_btn.setObjectName("SecondaryBtn")
        self.cancel_btn.clicked.connect(self.reject)

        self.ok_btn = QPushButton(STR["SETTINGS_OK"])
        self.ok_btn.setObjectName("PrimaryBtn")
        self.ok_btn.clicked.connect(self.save_settings)

        button_layout.addWidget(self.cancel_btn)
        button_layout.addWidget(self.ok_btn)
        layout.addLayout(button_layout)

        self.load_settings()

    def browse_word_template_path(self, edit_widget):
        path = QFileDialog.getExistingDirectory(self, STR["DIALOG_SELECT_TEMPLATE_FOLDER_TITLE"])
        if path:
            edit_widget.setText(path)

    def load_settings(self):
        self.word_template_folder_edit.setText(self.config.get("word_template_folder", ""))

    def save_settings(self):
        word_path = self.word_template_folder_edit.text()

        if word_path and not os.path.exists(word_path):
            try:
                os.makedirs(word_path)
            except Exception as e:
                QMessageBox.warning(self, STR["WARN_TITLE"], STR["ERROR_CREATE_FOLDER_FAILED"].format(error=e))
                return

        self.config["word_template_folder"] = word_path
        self.accept()

class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(STR["LOG_WINDOW_TITLE"])
        self.setMinimumSize(700, 500)
        self.setStyleSheet(APPLE_STYLE)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)
        
        header = QLabel(STR["LOG_HEADER_TITLE"])
        header.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(header)
        
        self.log_browser = QTextBrowser()
        self.log_browser.setFont(QFont(["SF Mono", "Courier New"], 10))
        layout.addWidget(self.log_browser)
        
        self.clear_btn = QPushButton(STR["LOG_CLEAR"])
        self.clear_btn.setObjectName("SecondaryBtn")
        self.clear_btn.clicked.connect(self.clear_log)
        layout.addWidget(self.clear_btn, alignment=Qt.AlignRight)

    def append_log(self, msg: str):
        self.log_browser.append(msg)
        self.log_browser.ensureCursorVisible()

    def clear_log(self):
        self.log_browser.clear()

class Worker(QThread):
    log_message = Signal(str)
    progress_updated = Signal(int, int) # current, total
    finished = Signal(int, str)

    def __init__(self, command):
        super().__init__()
        self.command = command

    def run(self):
        try:
            # Use sys.executable to ensure we use the same python environment
            process = subprocess.Popen(
                self.command,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='gbk',  # Windows Chinese systems typically use GBK (CP936)
                errors='replace',  # Handle decoding errors gracefully
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            total_templates = 0
            for line in iter(process.stdout.readline, ''):
                line = line.strip()
                if not line:
                    continue
                
                # Parse markers from template_filler.py
                if line.startswith("TOTAL_TEMPLATES:"):
                    try:
                        total_templates = int(line.split(":")[1])
                        self.progress_updated.emit(0, total_templates)
                    except (ValueError, IndexError):
                        pass
                elif line.startswith("PROGRESS:"):
                    try:
                        current_idx = int(line.split(":")[1].split("/")[0])
                        self.progress_updated.emit(current_idx, total_templates)
                    except (ValueError, IndexError):
                        pass
                
                self.log_message.emit(line)
            
            process.stdout.close()
            return_code = process.wait()
            self.finished.emit(return_code, "")
        except Exception as e:
            self.finished.emit(-1, str(e))

class DocGenGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(STR["APP_WINDOW_TITLE"])
        self.setMinimumSize(560, 380)
        self.resize(640, 420)
        self.setStyleSheet(APPLE_STYLE)

        self.config = self.load_config()
        self.log_dialog = LogDialog(self)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(12, 12, 12, 12)
        self.main_layout.setSpacing(6)

        header_container = QFrame()
        header_container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        header_container.setMaximumWidth(520)
        header_v = QVBoxLayout(header_container)
        header_v.setContentsMargins(0, 0, 0, 0)
        header_v.setSpacing(4)
        title = QLabel(STR["APP_HEADER_TITLE"])
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #1D1D1F;")
        subtitle = QLabel(STR["APP_HEADER_SUBTITLE"])
        subtitle.setStyleSheet("font-size: 12px; color: #86868B;")
        header_v.addWidget(title, alignment=Qt.AlignHCenter)
        header_v.addWidget(subtitle, alignment=Qt.AlignHCenter)
        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.addStretch()
        header_row.addWidget(header_container)
        header_row.addStretch()
        self.main_layout.addLayout(header_row)

        self.status_card = QFrame()
        self.status_card.setObjectName("MainCard")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 4)
        self.status_card.setGraphicsEffect(shadow)
        self.status_card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        
        status_layout = QGridLayout(self.status_card)
        status_layout.setContentsMargins(12, 12, 12, 12)
        status_layout.setHorizontalSpacing(6)
        status_layout.setVerticalSpacing(6)
        
        self.progress_label = QLabel(STR["STATUS_READY_SELECT_SOURCE"])
        self.progress_label.setStyleSheet("font-size: 12px; color: #86868B; font-weight: 600; text-transform: uppercase;")
        status_layout.addWidget(self.progress_label, 0, 0, 1, 3)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        status_layout.addWidget(self.progress_bar, 1, 0, 1, 3)
        
        # Big primary action button on its own row
        self.generate_btn = QPushButton(STR["BTN_START"])
        self.generate_btn.setObjectName("PrimaryBtn")
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.setMinimumHeight(52)
        self.generate_btn.clicked.connect(self.generate_contract)
        status_layout.addWidget(self.generate_btn, 2, 0, 1, 3)

        # Small secondary actions on third row
        self.choose_template_btn = QPushButton(STR["BTN_CHOOSE_TEMPLATE"])
        self.choose_template_btn.setObjectName("SecondaryBtn")
        self.choose_template_btn.setCursor(Qt.PointingHandCursor)
        self.choose_template_btn.setFixedHeight(32)

        self.show_log_btn = QPushButton(STR["BTN_VIEW_LOG"])
        self.show_log_btn.setObjectName("SecondaryBtn")
        self.show_log_btn.setCursor(Qt.PointingHandCursor)
        self.show_log_btn.setFixedHeight(32)

        self.settings_btn = QPushButton(STR["BTN_OPEN_SETTINGS"])
        self.settings_btn.setObjectName("SecondaryBtn")
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setFixedHeight(32)

        status_layout.addWidget(self.choose_template_btn, 3, 0, alignment=Qt.AlignLeft)
        status_layout.addWidget(self.show_log_btn, 3, 1, alignment=Qt.AlignLeft)
        status_layout.addWidget(self.settings_btn, 3, 2, alignment=Qt.AlignRight)
        status_layout.setColumnStretch(0, 0)
        status_layout.setColumnStretch(1, 0)
        status_layout.setColumnStretch(2, 1)

        self.main_layout.addWidget(self.status_card)

        self.settings_btn.clicked.connect(self.open_settings)
        self.choose_template_btn.clicked.connect(self.choose_template_folder)
        self.show_log_btn.clicked.connect(self.log_dialog.exec)
        self.last_zip_path = None
        self.reset_ready_state()
        self.check_templates_ready(show_dialog=False)
        self.setup_tab_order()

    def log(self, message, color="#1D1D1F"):
        if not message.strip():
            return
        
        if "ZIP_PATH:" in message:
            self.last_zip_path = message.split("ZIP_PATH:")[1].strip()
            message = STR["LOG_ZIP_EXPORTED"].format(name=os.path.basename(self.last_zip_path))
            color = "#0071E3"

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_msg = f'<p style="margin:4px 0;"><span style="color:#86868B;">[{timestamp}]</span> <span style="color:{color};">{message.strip()}</span></p>'
        self.log_dialog.append_log(formatted_msg)

    def update_progress(self, current, total):
        if total > 0:
            percent = int((current / total) * 100)
            percent = max(0, min(100, percent))
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(percent)
            self.progress_bar.setFormat(f"{percent}%")
            self.progress_label.setText(STR["STATUS_RUNNING"].format(current=current, total=total))
        else:
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("0%")
            self.progress_label.setText(STR["STATUS_PREPARING"])

    def reset_ready_state(self):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.progress_label.setText(STR["STATUS_READY_SELECT_SOURCE"])
        self.generate_btn.setEnabled(True)

    def load_config(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(script_dir, "config.json")
        default_template_dir = self.compute_default_template_dir(script_dir)
        self.ensure_word_template_folder(script_dir, default_template_dir)

        default_config = {
            "word_template_folder": default_template_dir,
            "last_excel_dir": script_dir,
        }

        if not os.path.exists(self.config_path):
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            return default_config

        with open(self.config_path, "r", encoding="utf-8") as f:
            try:
                config = json.load(f) or {}
            except Exception:
                config = {}

        mutated = False
        for removed_key in ("excel_path", "zip_output_path"):
            if removed_key in config:
                config.pop(removed_key, None)
                mutated = True

        if "word_template_folder" not in config or not str(config.get("word_template_folder") or "").strip():
            config["word_template_folder"] = default_template_dir
            mutated = True

        if "last_excel_dir" not in config or not str(config.get("last_excel_dir") or "").strip():
            config["last_excel_dir"] = script_dir
            mutated = True

        word_template_folder = config.get("word_template_folder")
        if word_template_folder and not os.path.exists(word_template_folder):
            try:
                os.makedirs(word_template_folder, exist_ok=True)
                mutated = True
            except Exception:
                config["word_template_folder"] = default_template_dir
                mutated = True

        if mutated:
            with open(self.config_path, "w", encoding="utf-8") as fw:
                json.dump(config, fw, indent=4, ensure_ascii=False)

        return config

    def save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def choose_template_folder(self):
        current = self.config.get("word_template_folder") or os.path.dirname(os.path.abspath(__file__))
        path = QFileDialog.getExistingDirectory(self, STR["DIALOG_SELECT_TEMPLATE_FOLDER_TITLE"], current)
        if not path:
            return
        self.config["word_template_folder"] = path
        self.save_config()
        self.check_templates_ready(show_dialog=True)

    def open_settings(self):
        dialog = SettingsDialog(self.config, self)
        if dialog.exec():
            self.save_config()
            self.log(STR["LOG_CONFIG_UPDATED"], "#0071E3")

    def generate_contract(self):
        if not self.check_templates_ready(show_dialog=True):
            return
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_initial_dir = self.config.get("last_excel_dir") or script_dir
        excel_path, _ = QFileDialog.getOpenFileName(
            self,
            STR["DIALOG_SELECT_EXCEL_TITLE"],
            excel_initial_dir,
            STR["DIALOG_SELECT_EXCEL_FILTER"],
        )
        if not excel_path:
            self.reset_ready_state()
            return
        if not excel_path.lower().endswith(".xlsx") or not os.path.exists(excel_path):
            self.reset_ready_state()
            QMessageBox.critical(self, STR["ERROR_TITLE"], STR["ERROR_EXCEL_INVALID"].format(path=excel_path))
            return

        self.config["last_excel_dir"] = os.path.dirname(os.path.abspath(excel_path))
        self.save_config()

        template_dir = self.config.get("word_template_folder") or ""
        if not template_dir or not os.path.exists(template_dir):
            self.reset_ready_state()
            QMessageBox.critical(
                self,
                STR["ERROR_TITLE"],
                STR["ERROR_TEMPLATE_DIR_MISSING"].format(path=template_dir),
            )
            return

        zip_output_dir = os.path.dirname(os.path.abspath(excel_path))

        script_path = os.path.join(script_dir, "template_filler.py")
        if not os.path.exists(script_path):
            self.log(STR["ERROR_CORE_SCRIPT_MISSING"], "#FF3B30")
            QMessageBox.critical(self, STR["ERROR_TITLE"], STR["ERROR_CORE_SCRIPT_MISSING"])
            self.reset_ready_state()
            return

        self.log_dialog.clear_log()
        self.log(STR["LOG_TASK_STARTED"])
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.progress_label.setText(STR["STATUS_PREPARING"])

        python_exe = sys.executable
        venv_python_exe = os.path.join(script_dir, ".venv", "Scripts", "python.exe")
        if os.path.exists(venv_python_exe):
            python_exe = venv_python_exe
        
        command = [
            f'"{python_exe}"',
            f'"{script_path}"',
            f'--excel "{excel_path}"',
            f'--template-dir "{template_dir}"',
            f'--zip-output "{zip_output_dir}"'
        ]
        full_command = " ".join(command)

        self.worker = Worker(full_command)
        self.worker.log_message.connect(lambda msg: self.log(msg))
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.start()
        self.generate_btn.setEnabled(False)

    def on_generation_finished(self, return_code, error_message):
        self.generate_btn.setEnabled(True)
        if return_code == 0:
            self.progress_label.setText(STR["STATUS_DONE"])
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.progress_bar.setFormat("100%")
            self.log(STR["LOG_TASK_DONE"], "#34C759")
            
            if self.last_zip_path and os.path.exists(self.last_zip_path):
                try:
                    subprocess.Popen(f'explorer /select,"{os.path.abspath(self.last_zip_path)}"')
                except:
                    pass
        else:
            self.progress_label.setText(STR["STATUS_FAILED"])
            self.log(STR["LOG_TASK_FAILED"].format(error=error_message), "#FF3B30")
            QMessageBox.warning(self, STR["WARN_TITLE"], STR["MSGBOX_TASK_FAILED"].format(error=error_message))
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("0%")
        
        self.last_zip_path = None

    def ensure_word_template_folder(self, script_dir: str, word_template_dir: str):
        created_ok = True
        try:
            os.makedirs(word_template_dir, exist_ok=True)
        except Exception as e:
            print(f"[docgen_gui] 创建默认模板目录失败：{e}")
            created_ok = False
        if not created_ok:
            fallback = os.path.join(os.path.expanduser("~"), "docgen_templates")
            try:
                os.makedirs(fallback, exist_ok=True)
                # update config fallback path
                self.config = getattr(self, "config", {}) or {}
                self.config["word_template_folder"] = fallback
                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(self.config, f, indent=4, ensure_ascii=False)
            except Exception as e:
                print(f"[docgen_gui] 回退目录创建失败：{e}")
                return

        try:
            candidates = []
            for name in os.listdir(script_dir):
                if not name.lower().endswith(".docx"):
                    continue
                if name.startswith("~$"):
                    continue
                candidates.append(os.path.join(script_dir, name))

            for src_path in candidates:
                dest_path = os.path.join(word_template_dir, os.path.basename(src_path))
                if os.path.abspath(src_path) == os.path.abspath(dest_path):
                    continue
                if os.path.exists(dest_path):
                    continue
                try:
                    os.replace(src_path, dest_path)
                except Exception:
                    continue
        except Exception:
            return

    @staticmethod
    def compute_default_template_dir(script_dir: str) -> str:
        return os.path.join(script_dir, "word_template")

    @staticmethod
    def probe_templates_status(template_dir: str) -> bool:
        try:
            if not template_dir or not os.path.exists(template_dir):
                return False
            items = []
            for name in os.listdir(template_dir):
                if name.startswith("~$"):
                    continue
                if name.lower().endswith(".docx"):
                    items.append(name)
            return len(items) > 0
        except PermissionError:
            return False
        except Exception:
            return False

    def check_templates_ready(self, show_dialog: bool = True) -> bool:
        tpl_dir = self.config.get("word_template_folder") or ""
        ok = self.probe_templates_status(tpl_dir)
        self.generate_btn.setEnabled(bool(ok))
        if ok or not show_dialog:
            return ok
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Critical)
        box.setWindowTitle(STR["DIALOG_TEMPLATE_MISSING_TITLE"])
        box.setText(STR["DIALOG_TEMPLATE_MISSING_BODY"])
        open_btn = box.addButton(STR["DIALOG_OPEN_FOLDER"], QMessageBox.ActionRole)
        box.addButton(QMessageBox.Close)
        box.setModal(True)
        box.exec()
        if box.clickedButton() is open_btn:
            parent_dir = os.path.dirname(tpl_dir) if tpl_dir else os.path.dirname(os.path.abspath(__file__))
            try:
                QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(parent_dir)))
            except Exception:
                try:
                    subprocess.Popen(f'explorer "{os.path.abspath(parent_dir)}"')
                except Exception:
                    pass
        return False

    def setup_tab_order(self):
        try:
            self.setTabOrder(self.choose_template_btn, self.generate_btn)
            self.setTabOrder(self.generate_btn, self.show_log_btn)
            self.setTabOrder(self.show_log_btn, self.settings_btn)
        except Exception:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    main_win = DocGenGUI()
    main_win.show()
    sys.exit(app.exec())

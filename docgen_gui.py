
import os
import sys
import json
import subprocess
import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget,
    QTextBrowser, QDialog, QLabel, QLineEdit, QFileDialog, QMessageBox, QProgressBar,
    QGraphicsDropShadowEffect, QFrame
)
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QSize
from PySide6.QtGui import QFont, QColor, QPalette

# Modern Apple-inspired Theme (HIG Compliant)
APPLE_STYLE = """
    QMainWindow {
        background-color: #F5F5F7;
    }
    
    QWidget {
        font-family: "SF Pro Display", "SF Pro Text", "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
        font-size: 13px;
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
        height: 12px;
        font-size: 11px;
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
        self.setWindowTitle("路径配置")
        self.setMinimumWidth(650)
        self.setStyleSheet(APPLE_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        # Header
        header = QLabel("路径配置")
        header.setStyleSheet("font-size: 20px; font-weight: 700; margin-bottom: 8px;")
        layout.addWidget(header)

        # Path settings
        self.path_container = QFrame()
        self.path_container.setObjectName("MainCard")
        path_layout = QVBoxLayout(self.path_container)
        path_layout.setContentsMargins(16, 16, 16, 16)
        path_layout.setSpacing(12)

        path_layout.addLayout(self.create_path_selector("Word 模板文件夹", "word_template_folder", self.browse_word_template_path))
        path_layout.addLayout(self.create_path_selector("Excel 输入文件", "excel_path", self.browse_excel_path))
        path_layout.addLayout(self.create_path_selector("ZIP 输出路径", "zip_output_path", self.browse_zip_output_path))
        
        layout.addWidget(self.path_container)

        # Buttons
        self.save_btn = QPushButton("保存")
        self.save_btn.setObjectName("PrimaryBtn")
        self.save_btn.clicked.connect(self.save_settings)
        
        self.restore_defaults_btn = QPushButton("恢复默认")
        self.restore_defaults_btn.setObjectName("SecondaryBtn")
        self.restore_defaults_btn.clicked.connect(self.restore_defaults)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(12)
        button_layout.addWidget(self.restore_defaults_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.save_btn)
        layout.addLayout(button_layout)

        self.load_settings()

    def create_path_selector(self, label_text, config_key, browse_method):
        layout = QVBoxLayout()
        layout.setSpacing(6)
        
        label = QLabel(label_text)
        label.setStyleSheet("font-size: 12px; color: #86868B;")
        
        row_layout = QHBoxLayout()
        row_layout.setSpacing(8)
        
        edit = DraggableLineEdit()
        edit.setReadOnly(True)
        edit.setPlaceholderText(f"选择 {label_text}...")
        
        browse_btn = QPushButton("选择...")
        browse_btn.setObjectName("SecondaryBtn")
        browse_btn.setFixedWidth(80)
        browse_btn.clicked.connect(lambda: browse_method(edit))
        
        row_layout.addWidget(edit)
        row_layout.addWidget(browse_btn)
        
        layout.addWidget(label)
        layout.addLayout(row_layout)
        
        setattr(self, f"{config_key}_edit", edit)
        return layout

    def browse_word_template_path(self, edit_widget):
        path = QFileDialog.getExistingDirectory(self, "选择Word模板文件夹")
        if path:
            edit_widget.setText(path)

    def browse_excel_path(self, edit_widget):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if path:
            edit_widget.setText(path)

    def browse_zip_output_path(self, edit_widget):
        path = QFileDialog.getExistingDirectory(self, "选择ZIP输出路径")
        if path:
            edit_widget.setText(path)

    def load_settings(self):
        self.word_template_folder_edit.setText(self.config.get("word_template_folder", ""))
        self.excel_path_edit.setText(self.config.get("excel_path", ""))
        self.zip_output_path_edit.setText(self.config.get("zip_output_path", ""))

    def save_settings(self):
        word_path = self.word_template_folder_edit.text()
        excel_path = self.excel_path_edit.text()
        zip_path = self.zip_output_path_edit.text()

        if not excel_path or not os.path.isdir(os.path.dirname(excel_path)):
            QMessageBox.warning(self, "警告", "Excel文件所在目录无效！")
            return

        if not zip_path or not os.access(zip_path, os.W_OK):
            QMessageBox.warning(self, "警告", "ZIP输出路径不可写！")
            return

        if word_path and not os.path.exists(word_path):
            try:
                os.makedirs(word_path)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"创建文件夹失败：{e}")
                return

        self.config["word_template_folder"] = word_path
        self.config["excel_path"] = excel_path
        self.config["zip_output_path"] = zip_path
        self.accept()

    def restore_defaults(self):
        default_dir = os.path.dirname(os.path.abspath(__file__))
        self.word_template_folder_edit.setText(default_dir)
        self.excel_path_edit.setText(os.path.join(default_dir, "input.xlsx"))
        self.zip_output_path_edit.setText(default_dir)

class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("生成日志")
        self.setMinimumSize(700, 500)
        self.setStyleSheet(APPLE_STYLE)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)
        
        header = QLabel("任务日志")
        header.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(header)
        
        self.log_browser = QTextBrowser()
        self.log_browser.setFont(QFont(["SF Mono", "Courier New"], 10))
        layout.addWidget(self.log_browser)
        
        self.clear_btn = QPushButton("清空日志记录")
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
        self.setWindowTitle("合同自动生成工具")
        self.setMinimumSize(500, 350)
        self.setStyleSheet(APPLE_STYLE)

        self.config = self.load_config()
        self.log_dialog = LogDialog(self)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Main layout with 24pt margins
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(32, 32, 32, 32)
        self.main_layout.setSpacing(24)

        # 1. Header Area
        header_layout = QVBoxLayout()
        header_layout.setSpacing(4)
        title = QLabel("合同生成助手")
        title.setStyleSheet("font-size: 28px; font-weight: 700; color: #1D1D1F;")
        subtitle = QLabel("自动化批量填充 Word 模板与 Excel 数据")
        subtitle.setStyleSheet("font-size: 14px; color: #86868B;")
        header_layout.addWidget(title)
        header_layout.addWidget(subtitle)
        self.main_layout.addLayout(header_layout)

        # 2. Status Card (Progress)
        self.status_card = QFrame()
        self.status_card.setObjectName("MainCard")
        # Apply shadow for depth
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 4)
        self.status_card.setGraphicsEffect(shadow)
        
        status_layout = QVBoxLayout(self.status_card)
        status_layout.setContentsMargins(20, 20, 20, 20)
        status_layout.setSpacing(12)
        
        self.progress_label = QLabel("准备就绪")
        self.progress_label.setStyleSheet("font-size: 12px; color: #86868B; font-weight: 600; text-transform: uppercase;")
        status_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("等待启动...")
        status_layout.addWidget(self.progress_bar)
        
        self.main_layout.addWidget(self.status_card)

        # 3. Action Buttons
        self.generate_btn = QPushButton("开始生成合同")
        self.generate_btn.setObjectName("PrimaryBtn")
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate_contract)
        self.main_layout.addWidget(self.generate_btn)

        # 4. Footer Tools
        self.footer_layout = QHBoxLayout()
        self.footer_layout.setSpacing(12)
        
        self.modify_template_btn = QPushButton("模板库")
        self.fill_info_btn = QPushButton("数据源")
        self.show_log_btn = QPushButton("查看日志")
        self.settings_btn = QPushButton("偏好设置")

        for btn in [self.modify_template_btn, self.fill_info_btn, self.show_log_btn, self.settings_btn]:
            btn.setObjectName("SecondaryBtn")
            btn.setCursor(Qt.PointingHandCursor)
            self.footer_layout.addWidget(btn)

        self.main_layout.addStretch()
        self.main_layout.addLayout(self.footer_layout)

        # Connections
        self.settings_btn.clicked.connect(self.open_settings)
        self.modify_template_btn.clicked.connect(self.open_template_folder)
        self.fill_info_btn.clicked.connect(self.open_excel_file)
        self.show_log_btn.clicked.connect(self.log_dialog.exec)
        self.last_zip_path = None

    def log(self, message, color="#1D1D1F"):
        if not message.strip():
            return
        
        # Capture generated ZIP path
        if "ZIP_PATH:" in message:
            self.last_zip_path = message.split("ZIP_PATH:")[1].strip()
            message = f"✅ 已成功导出: {os.path.basename(self.last_zip_path)}"
            color = "#0071E3"

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_msg = f'<p style="margin:4px 0;"><span style="color:#86868B;">[{timestamp}]</span> <span style="color:{color};">{message.strip()}</span></p>'
        self.log_dialog.append_log(formatted_msg)

    def update_progress(self, current, total):
        if total > 0:
            self.progress_bar.setRange(0, total)
            self.progress_bar.setValue(current)
            percent = int(current / total * 100)
            self.progress_bar.setFormat(f"{percent}%")
            self.progress_label.setText(f"正在处理: {current} / {total}")
        else:
            self.progress_bar.setRange(0, 0)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("初始化...")
            self.progress_label.setText("正在解析模板...")

    def load_config(self):
        self.config_path = "config.json"
        if not os.path.exists(self.config_path):
            default_dir = os.path.dirname(os.path.abspath(__file__))
            config = {
                "word_template_folder": default_dir,
                "excel_path": os.path.join(default_dir, "input.xlsx"),
                "zip_output_path": default_dir,
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4)
            return config
        else:
            with open(self.config_path, "r", encoding="utf-8") as f:
                try:
                    return json.load(f)
                except:
                    return {}

    def save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)

    def open_settings(self):
        dialog = SettingsDialog(self.config, self)
        if dialog.exec():
            self.save_config()
            self.log("配置已更新", "#0071E3")

    def open_template_folder(self):
        path = self.config.get("word_template_folder")
        if path and os.path.exists(path):
            subprocess.Popen(f'explorer "{os.path.abspath(path)}"')
        else:
            QMessageBox.warning(self, "警告", "模板文件夹路径无效，请在设置中配置。")

    def open_excel_file(self):
        path = self.config.get("excel_path")
        if path and os.path.exists(path):
            os.startfile(os.path.abspath(path))
        else:
            QMessageBox.warning(self, "警告", "Excel文件路径无效，请在设置中配置。")

    def generate_contract(self):
        if not self.validate_paths():
            return

        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_filler.py")
        if not os.path.exists(script_path):
            self.log("错误: template_filler.py 不存在", "#FF3B30")
            QMessageBox.critical(self, "错误", "未找到核心填充脚本。")
            return

        self.log_dialog.clear_log()
        self.log("任务开始...")
        
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setValue(0)
        self.progress_label.setText("正在准备环境...")
        
        command = [
            f'"{sys.executable}"',
            f'"{script_path}"',
            f'--excel "{self.config["excel_path"]}"',
            f'--template-dir "{self.config["word_template_folder"]}"',
            f'--zip-output "{self.config["zip_output_path"]}"'
        ]
        full_command = " ".join(command)

        self.worker = Worker(full_command)
        self.worker.log_message.connect(lambda msg: self.log(msg))
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.start()
        self.generate_btn.setEnabled(False)

    def validate_paths(self):
        word_path = self.config.get("word_template_folder")
        excel_path = self.config.get("excel_path")
        zip_path = self.config.get("zip_output_path")

        if not all([word_path, excel_path, zip_path]):
            QMessageBox.critical(self, "错误", "路径配置不完整，请先在设置中配置。")
            return False

        if not os.path.exists(word_path):
            QMessageBox.critical(self, "错误", f"模板文件夹不存在:\n{word_path}")
            return False

        if not os.path.exists(excel_path):
            QMessageBox.critical(self, "错误", f"Excel文件不存在:\n{excel_path}")
            return False

        return True

    def on_generation_finished(self, return_code, error_message):
        self.generate_btn.setEnabled(True)
        if return_code == 0:
            self.progress_label.setText("任务已完成")
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.log("✨ 所有合同已生成并打包成功！", "#34C759")
            
            if self.last_zip_path and os.path.exists(self.last_zip_path):
                try:
                    subprocess.Popen(f'explorer /select,"{os.path.abspath(self.last_zip_path)}"')
                except:
                    pass
        else:
            self.progress_label.setText("任务失败")
            self.log(f"❌ 生成过程中出现错误: {error_message}", "#FF3B30")
            QMessageBox.warning(self, "任务失败", f"生成失败: {error_message}")
        
        self.last_zip_path = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    main_win = DocGenGUI()
    main_win.show()
    sys.exit(app.exec())

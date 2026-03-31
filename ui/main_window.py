"""
PyQt6 主窗口模块
提供文件拖拽/选择、样式切换、转换按钮等 UI 交互。
"""

import os
import sys
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QRadioButton, QButtonGroup,
    QFileDialog, QMessageBox, QFrame, QGroupBox,
    QProgressBar, QApplication, QTabWidget, QPlainTextEdit,
)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QIcon

from converter.md_parser import parse_markdown, read_md_file
from converter.docx_writer import DocxWriter


class DropArea(QFrame):
    """文件拖拽区域"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_path = None
        self.parent_window = parent

        self.setMinimumHeight(180)
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 图标文字
        self.icon_label = QLabel("📄")
        self.icon_label.setFont(QFont("Segoe UI Emoji", 36))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label)

        self.label = QLabel("将 Markdown 文件拖到这里\n或点击下方按钮选择文件")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setFont(QFont("Microsoft YaHei", 11))
        self.label.setStyleSheet("color: #666; border: none; background: transparent;")
        layout.addWidget(self.label)

        self.file_label = QLabel("")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setFont(QFont("Microsoft YaHei", 10))
        self.file_label.setStyleSheet("color: #5B8DEF; font-weight: bold; border: none; background: transparent;")
        self.file_label.setWordWrap(True)
        layout.addWidget(self.file_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.toLocalFile().lower().endswith('.md'):
                    event.acceptProposedAction()
                    self.setStyleSheet("""
                        DropArea {
                            border: 2px solid #5B8DEF;
                            border-radius: 12px;
                            background-color: #E8EFFF;
                        }
                    """)
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.md'):
                self.set_file(file_path)
                break

        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #AABBCC;
                border-radius: 12px;
                background-color: #F8FAFC;
            }
            DropArea:hover {
                border-color: #5B8DEF;
                background-color: #EEF2FF;
            }
        """)

    def set_file(self, file_path: str):
        """设置选中的文件。"""
        self.file_path = file_path
        filename = os.path.basename(file_path)
        self.file_label.setText(f"✅ {filename}")
        self.icon_label.setText("📝")
        self.label.setText("已选择文件：")
        if self.parent_window and hasattr(self.parent_window, 'on_file_selected'):
            self.parent_window.on_file_selected()

    def clear_file(self):
        """清除已选文件。"""
        self.file_path = None
        self.file_label.setText("")
        self.icon_label.setText("📄")
        self.label.setText("将 Markdown 文件拖到这里\n或点击下方按钮选择文件")


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("MD → DOCX 转换工具")
        self.setMinimumSize(620, 720)
        self.resize(620, 720)

        # 主容器
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(28, 20, 28, 20)
        main_layout.setSpacing(12)

        # 应用全局样式
        self.setStyleSheet("""
            * {
                font-family: "Microsoft YaHei", "PingFang SC", "Helvetica Neue", sans-serif;
            }
            QMainWindow {
                background-color: #FFFFFF;
            }
            QGroupBox {
                font-family: "Microsoft YaHei", "PingFang SC", "Helvetica Neue", sans-serif;
                font-size: 12px;
                font-weight: bold;
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 16px;
                background-color: #FAFBFC;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #333;
            }
            QRadioButton {
                font-family: "Microsoft YaHei", "PingFang SC", "Helvetica Neue", sans-serif;
                font-size: 12px;
                spacing: 8px;
                padding: 6px 4px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QTabWidget::pane {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background: #FAFBFC;
            }
            QTabBar::tab {
                font-family: "Microsoft YaHei", "PingFang SC", "Helvetica Neue", sans-serif;
                font-size: 11px;
                font-weight: bold;
                padding: 8px 20px;
                margin-right: 2px;
                border: 1px solid #E0E0E0;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                background: #F0F0F0;
                color: #666;
            }
            QTabBar::tab:selected {
                background: #FAFBFC;
                color: #5B8DEF;
                border-bottom: 2px solid #5B8DEF;
            }
            QTabBar::tab:hover:!selected {
                background: #E8EFFF;
            }
        """)

        # ===== 标题 =====
        title_label = QLabel("MD → DOCX 转换工具")
        title_label.setFont(QFont("Microsoft YaHei", 16, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #1A1A2E; margin-bottom: 4px;")
        main_layout.addWidget(title_label)

        subtitle = QLabel("将 Markdown 文件或文本转换为格式化的 Word 文档")
        subtitle.setFont(QFont("Microsoft YaHei", 9))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #888;")
        main_layout.addWidget(subtitle)

        # ===== 标签页 =====
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # ----- Tab 1: 文件模式 -----
        file_tab = QWidget()
        file_tab_layout = QVBoxLayout(file_tab)
        file_tab_layout.setContentsMargins(12, 12, 12, 12)
        file_tab_layout.setSpacing(10)

        self.drop_area = DropArea(self)
        file_tab_layout.addWidget(self.drop_area)

        file_btn_layout = QHBoxLayout()
        
        select_btn = QPushButton("📂  选择 Markdown 文件")
        select_btn.setFont(QFont("Microsoft YaHei", 10))
        select_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        select_btn.setStyleSheet("""
            QPushButton {
                background-color: #F0F4FF;
                color: #5B8DEF;
                border: 1px solid #C8D8F8;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E0EAFF;
                border-color: #5B8DEF;
            }
            QPushButton:pressed {
                background-color: #D0DFFF;
            }
        """)
        select_btn.clicked.connect(self.select_file)
        file_btn_layout.addWidget(select_btn)
        
        clear_btn = QPushButton("❌  清除选择")
        clear_btn.setFont(QFont("Microsoft YaHei", 10))
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFF5F5;
                color: #EF4444;
                border: 1px solid #FDD;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                border-color: #EF4444;
            }
            QPushButton:pressed {
                background-color: #FECACA;
            }
        """)
        clear_btn.clicked.connect(self.clear_selection)
        file_btn_layout.addWidget(clear_btn)
        
        file_tab_layout.addLayout(file_btn_layout)
        self.tab_widget.addTab(file_tab, "📄 文件转换")

        # ----- Tab 2: 文本模式 -----
        text_tab = QWidget()
        text_tab_layout = QVBoxLayout(text_tab)
        text_tab_layout.setContentsMargins(12, 12, 12, 12)
        text_tab_layout.setSpacing(10)

        text_hint = QLabel("在下方粘贴或输入 Markdown 文本，然后点击“开始转换”")
        text_hint.setFont(QFont("Microsoft YaHei", 9))
        text_hint.setStyleSheet("color: #888;")
        text_tab_layout.addWidget(text_hint)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("请在此处粘贴或输入 Markdown 内容……\n\n例如：\n# 标题\n## 二级标题\n正文内容……")
        self.text_edit.setFont(QFont("Consolas", 11))
        self.text_edit.setMinimumHeight(200)
        self.text_edit.setStyleSheet("""
            QPlainTextEdit {
                border: 1px solid #D0D5DD;
                border-radius: 8px;
                padding: 10px;
                background-color: #FFFFFF;
                color: #333;
                selection-background-color: #C8D8F8;
            }
            QPlainTextEdit:focus {
                border-color: #5B8DEF;
            }
        """)
        text_tab_layout.addWidget(self.text_edit)

        clear_text_btn = QPushButton("🗑️  清空文本")
        clear_text_btn.setFont(QFont("Microsoft YaHei", 10))
        clear_text_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_text_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFF5F5;
                color: #EF4444;
                border: 1px solid #FDD;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FEE2E2;
                border-color: #EF4444;
            }
            QPushButton:pressed {
                background-color: #FECACA;
            }
        """)
        clear_text_btn.clicked.connect(lambda: self.text_edit.clear())
        text_tab_layout.addWidget(clear_text_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.tab_widget.addTab(text_tab, "✍️ 文本转换")

        # ===== 样式选择 =====
        style_group = QGroupBox("选择转换样式")
        style_layout = QVBoxLayout(style_group)
        style_layout.setSpacing(4)

        self.style_btn_group = QButtonGroup(self)

        self.radio_a = QRadioButton("样式 A：文档标题作为「标题」（Title）\n    # → Title， ## → Heading 1， ### → Heading 2")
        self.radio_b = QRadioButton("样式 B：文档标题作为「标题 1」（Heading 1）\n    # → Heading 1， ## → Heading 2， ### → Heading 3")
        self.radio_a.setChecked(True)
        self.radio_a.setMinimumHeight(40)
        self.radio_b.setMinimumHeight(40)

        self.style_btn_group.addButton(self.radio_a, 0)
        self.style_btn_group.addButton(self.radio_b, 1)

        style_layout.addWidget(self.radio_a)
        style_layout.addWidget(self.radio_b)
        main_layout.addWidget(style_group)

        # ===== 转换按钮 =====
        self.convert_btn = QPushButton("🚀  开始转换")
        self.convert_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.convert_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_btn.setEnabled(True)
        self.convert_btn.setStyleSheet("""
            QPushButton {
                background-color: #5B8DEF;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                min-height: 20px;
            }
            QPushButton:hover {
                background-color: #4A7CE0;
            }
            QPushButton:pressed {
                background-color: #3A6CD0;
            }
            QPushButton:disabled {
                background-color: #CCC;
                color: #999;
            }
        """)
        self.convert_btn.clicked.connect(self.do_convert)
        self.convert_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        main_layout.addWidget(self.convert_btn)

        # ===== 状态栏 =====
        self.status_label = QLabel("")
        self.status_label.setFont(QFont("Microsoft YaHei", 9))
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #888;")
        self.status_label.setWordWrap(True)
        self.status_label.setMinimumHeight(24)
        main_layout.addWidget(self.status_label)

        main_layout.addStretch()

    def on_file_selected(self):
        """文件选中后启用转换按钮。"""
        self.convert_btn.clearFocus()
        self.status_label.setText("")

    def clear_selection(self):
        """清除已选文件。"""
        self.drop_area.clear_file()
        self.status_label.setText("")
        self.status_label.setStyleSheet("color: #888;")

    def select_file(self):
        """弹出文件选择对话框。"""
        # 手动清空可能聚焦的状态
        for widget in QApplication.topLevelWidgets():
            widget.clearFocus()

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Markdown 文件",
            "",
            "Markdown 文件 (*.md);;所有文件 (*.*)"
        )
        if file_path:
            self.drop_area.set_file(file_path)

    def do_convert(self):
        """执行转换，根据当前 Tab 分发逻辑。"""
        current_tab = self.tab_widget.currentIndex()
        
        if current_tab == 0:
            self._convert_from_file()
        else:
            self._convert_from_text()

    def _convert_from_file(self):
        """文件模式转换。"""
        file_path = self.drop_area.file_path
        
        # 如果没有选择文件，弹出文件选择框
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择 Markdown 文件",
                "",
                "Markdown 文件 (*.md);;所有文件 (*.*)"
            )
            if not file_path:
                return
            self.drop_area.set_file(file_path)

        if not os.path.isfile(file_path):
            QMessageBox.warning(self, "错误", f"文件不存在：\n{file_path}")
            return

        style_mode = 'A' if self.radio_a.isChecked() else 'B'
        
        default_save_path = os.path.splitext(file_path)[0] + '.docx'
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存 Word 文档",
            default_save_path,
            "Word 文档 (*.docx)"
        )
        if not save_path:
            return

        self.status_label.setText("⏳ 正在转换...")
        self.status_label.setStyleSheet("color: #5B8DEF;")
        QApplication.processEvents()

        try:
            md_text = read_md_file(file_path)
            ast_nodes = parse_markdown(md_text)

            writer = DocxWriter(style_mode=style_mode)
            writer.convert(ast_nodes, save_path)

            self.status_label.setText(f"✅ 转换成功！已保存至：{os.path.basename(save_path)}")
            self.status_label.setStyleSheet("color: #22C55E;")
            self._show_success_dialog(save_path)

        except PermissionError as e:
            msg = f"保存失败：文件正在被其他程序（如 Word）占用。\n\n请尝试先关闭已打开的 '{os.path.basename(save_path)}'，然后再重新开始转换！\n\n底层错误信息：{str(e)}"
            self.status_label.setText("❌ 转换失败：文件被占用")
            self.status_label.setStyleSheet("color: #EF4444;")
            QMessageBox.critical(self, "转换失败 - 文件占用", msg)
            
        except Exception as e:
            self.status_label.setText(f"❌ 转换失败：{str(e)}")
            self.status_label.setStyleSheet("color: #EF4444;")
            QMessageBox.critical(self, "转换失败", f"转换过程中出现错误：\n{str(e)}")

    def _convert_from_text(self):
        """文本模式转换。"""
        md_text = self.text_edit.toPlainText().strip()
        
        if not md_text:
            QMessageBox.information(self, "提示", "请先在文本框中输入或粘贴 Markdown 内容。")
            return

        style_mode = 'A' if self.radio_a.isChecked() else 'B'

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存 Word 文档",
            "未命名文档.docx",
            "Word 文档 (*.docx)"
        )
        if not save_path:
            return

        self.status_label.setText("⏳ 正在转换...")
        self.status_label.setStyleSheet("color: #5B8DEF;")
        QApplication.processEvents()

        try:
            ast_nodes = parse_markdown(md_text)

            writer = DocxWriter(style_mode=style_mode)
            writer.convert(ast_nodes, save_path)

            self.status_label.setText(f"✅ 转换成功！已保存至：{os.path.basename(save_path)}")
            self.status_label.setStyleSheet("color: #22C55E;")
            self._show_success_dialog(save_path)

        except PermissionError as e:
            msg = f"保存失败：文件正在被其他程序占用。\n\n请先关闭已打开的 '{os.path.basename(save_path)}'。\n\n底层错误：{str(e)}"
            self.status_label.setText("❌ 转换失败：文件被占用")
            self.status_label.setStyleSheet("color: #EF4444;")
            QMessageBox.critical(self, "转换失败 - 文件占用", msg)
            
        except Exception as e:
            self.status_label.setText(f"❌ 转换失败：{str(e)}")
            self.status_label.setStyleSheet("color: #EF4444;")
            QMessageBox.critical(self, "转换失败", f"转换过程中出现错误：\n{str(e)}")

    def _show_success_dialog(self, save_path: str):
        """显示转换成功对话框。"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("转换成功")
        msg_box.setText(f"文件已保存至：\n{save_path}")
        msg_box.setIcon(QMessageBox.Icon.Information)

        ok_btn = msg_box.addButton("确定", QMessageBox.ButtonRole.AcceptRole)
        open_btn = msg_box.addButton("打开文档", QMessageBox.ButtonRole.ActionRole)

        msg_box.exec()

        if msg_box.clickedButton() == open_btn:
            import platform
            import subprocess
            if platform.system() == 'Windows':
                os.startfile(save_path)
            elif platform.system() == 'Darwin':
                subprocess.call(('open', save_path))
            else:
                subprocess.call(('xdg-open', save_path))

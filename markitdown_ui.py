import sys
from pathlib import Path
import warnings
import re
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QLabel, QLineEdit, QPushButton,
                              QTextEdit, QFileDialog, QMessageBox, QProgressBar,
                              QListWidget, QListWidgetItem, QFrame,
                              QAbstractItemView, QCheckBox)
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtGui import QFont, QDragEnterEvent, QDropEvent

VERSION = "1.1.0.0"

try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# 忽略各种警告
warnings.filterwarnings("ignore", message="Couldn't find ffmpeg or avconv")
warnings.filterwarnings("ignore", message="Unsupported Windows version")
warnings.filterwarnings("ignore", category=UserWarning, module="onnxruntime")

try:
    from markitdown import MarkItDown, UnsupportedFormatException, MissingDependencyException
except ImportError as e:
    print("Error: Cannot import markitdown library")
    print("Please run: pip install markitdown[all]")
    print(f"Details: {e}")
    sys.exit(1)

# 转换工作线程
class ConversionWorker(QThread):
    finished = Signal(str, str)  # markdown_content, source
    error = Signal(str)
    
    def __init__(self, md_instance, source, excel_file=None, selected_sheets=None):
        super().__init__()
        self.md = md_instance
        self.source = source
        self.excel_file = excel_file
        self.selected_sheets = selected_sheets
    
    def run(self):
        try:
            # 检查是否为 Excel 文件且需要特殊处理
            if (self.excel_file and 
                self.excel_file == self.source and 
                EXCEL_SUPPORT and 
                self.selected_sheets):
                
                # 使用自定义的 Excel 转换
                markdown_content = self._convert_excel_sheets(self.source, self.selected_sheets)
                self.finished.emit(markdown_content, self.source)
            else:
                # 使用 MarkItDown 的默认转换
                result = self.md.convert(self.source)
                self.finished.emit(result.markdown, self.source)
                
        except UnsupportedFormatException:
            self.error.emit("不支持的文件格式")
        except MissingDependencyException as e:
            self.error.emit(f"缺少依赖: {e}")
        except Exception as e:
            self.error.emit(f"转换失败: {str(e)}")
    
    def _convert_excel_sheets(self, filename, selected_sheets):
        """转换选中的 Excel sheets"""
        if not selected_sheets:
            raise Exception("请至少选择一个 Sheet")
        
        results = []
        for sheet_name in selected_sheets:
            try:
                workbook = openpyxl.load_workbook(filename, read_only=True)
                worksheet = workbook[sheet_name]
                
                # 将 sheet 数据转换为 markdown 表格
                markdown_content = self._worksheet_to_markdown(worksheet, sheet_name)
                results.append(markdown_content)
                
                workbook.close()
                
            except Exception as e:
                results.append(f"# {sheet_name}\n\n**错误**: 无法转换此 Sheet - {str(e)}\n\n")
        
        return "\n\n---\n\n".join(results)
    
    def _worksheet_to_markdown(self, worksheet, sheet_name):
        """将 Excel worksheet 转换为 Markdown"""
        markdown = f"# {sheet_name}\n\n"
        
        # 获取有数据的区域
        if worksheet.max_row == 1 and worksheet.max_column == 1:
            return markdown + "此 Sheet 为空\n"
        
        # 转换为表格
        rows = []
        for row in worksheet.iter_rows(values_only=True):
            # 跳过完全空的行
            if all(cell is None or str(cell).strip() == '' for cell in row):
                continue
            # 将 None 值转换为空字符串，其他值转换为字符串
            row_data = [str(cell) if cell is not None else '' for cell in row]
            rows.append(row_data)
        
        if not rows:
            return markdown + "此 Sheet 为空\n"
        
        # 确定最大列数
        max_cols = max(len(row) for row in rows) if rows else 0
        
        # 补齐所有行到相同列数
        for row in rows:
            while len(row) < max_cols:
                row.append('')
        
        # 生成 Markdown 表格
        if rows:
            # 表头
            header = "| " + " | ".join(rows[0]) + " |"
            separator = "| " + " | ".join(['---'] * len(rows[0])) + " |"
            markdown += header + "\n" + separator + "\n"
            
            # 数据行
            for row in rows[1:]:
                markdown += "| " + " | ".join(row) + " |\n"
        
        return markdown


# 支持拖拽的文本编辑器
class DragDropTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)
    
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if file_path:
                    # 发送信号给主窗口
                    main_window = self.window()
                    if hasattr(main_window, 'handle_file_drop'):
                        main_window.handle_file_drop(file_path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


# 支持拖拽的输入框
class DragDropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)
    
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if file_path:
                    self.setText(file_path)
                    # 通知主窗口文件已更改
                    main_window = self.window()
                    if hasattr(main_window, 'handle_file_drop'):
                        main_window.handle_file_drop(file_path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class MarkItDownUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"MarkItDown 文件转换器 v{VERSION}")
        self.setGeometry(100, 100, 1000, 850)
        self.setMinimumSize(850, 700)

        # 初始化变量
        self.excel_sheets = []
        self.selected_sheets = []
        self.current_excel_file = None
        self.current_result = ""
        self.current_title = ""
        self.use_default_output = False
        self.pending_save = False
        # 记录最近一次保存/默认输出目录，初始为桌面（若不存在则用户主目录）
        self.last_output_dir = self._get_default_output_dir()

        # 设置现代化样式
        self.setup_style()

        # 初始化MarkItDown
        try:
            self.md = MarkItDown()
            self.setup_ui()
        except Exception as e:
            QMessageBox.critical(self, "初始化错误", f"无法初始化MarkItDown: {e}")
            sys.exit(1)

    def setup_style(self):
        """设置现代化的应用样式 - 基于Material Design原则"""
        self.setStyleSheet("""
            /* ===== 主窗口 ===== */
            QMainWindow {
                background-color: #f8f9fa;
            }

            /* ===== 卡片容器 ===== */
            QWidget#cardContainer {
                background-color: white;
                border-radius: 12px;
                border: 1px solid #e9ecef;
            }

            /* ===== 分组框 ===== */
            QGroupBox {
                background-color: white;
                border: none;
                border-radius: 12px;
                margin-top: 8px;
                padding: 20px 16px 16px 16px;
                font-weight: 600;
                font-size: 14px;
                color: #212529;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 16px;
                top: 8px;
                padding: 0 8px;
                background-color: white;
                color: #495057;
            }

            /* ===== 输入框 ===== */
            QLineEdit {
                padding: 0px 12px;
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: #ffffff;
                font-size: 12px;
                color: #212529;
                selection-background-color: #0d6efd;
                selection-color: white;
                min-height: 20px;
            }

            QLineEdit:hover {
                border: 2px solid #adb5bd;
                background-color: #f8f9fa;
            }

            QLineEdit:focus {
                border: 2px solid #0d6efd;
                background-color: white;
                outline: none;
            }

            QLineEdit::placeholder {
                color: #adb5bd;
            }

            /* ===== 按钮 ===== */
            QPushButton {
                background-color: #0d6efd;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 600;
                min-width: 80px;
                min-height: 36px;
            }

            QPushButton:hover {
                background-color: #0b5ed7;
            }

            QPushButton:pressed {
                background-color: #0a58ca;
                padding: 9px 15px 7px 17px;
            }

            QPushButton:disabled {
                background-color: #e9ecef;
                color: #adb5bd;
            }

            /* 次要按钮 */
            QPushButton#secondaryButton {
                background-color: #6c757d;
                color: white;
            }

            QPushButton#secondaryButton:hover {
                background-color: #5c636a;
            }

            QPushButton#secondaryButton:pressed {
                background-color: #565e64;
            }

            /* 成功按钮 */
            QPushButton#successButton {
                background-color: #198754;
                color: white;
            }

            QPushButton#successButton:hover {
                background-color: #157347;
            }

            QPushButton#successButton:pressed {
                background-color: #146c43;
            }

            /* 危险按钮 */
            QPushButton#dangerButton {
                background-color: #dc3545;
                color: white;
            }

            QPushButton#dangerButton:hover {
                background-color: #bb2d3b;
            }

            QPushButton#dangerButton:pressed {
                background-color: #b02a37;
            }

            /* 浏览按钮 - 轮廓样式 */
            QPushButton#browseButton {
                background-color: transparent;
                color: #0d6efd;
                border: 2px solid #0d6efd;
                border-radius: 6px;
                padding: 0px 16px;
                font-size: 12px;
                font-weight: 600;
                min-width: 70px;
                min-height: 20px;
            }

            QPushButton#browseButton:hover {
                background-color: #0d6efd;
                color: white;
            }

            QPushButton#browseButton:pressed {
                background-color: #0b5ed7;
                border-color: #0b5ed7;
            }

            /* 紧凑按钮 - 用于 Excel 操作等 */
            QPushButton#compactButton {
                background-color: #e9ecef;
                color: #495057;
                border: none;
                border-radius: 4px;
                padding: 4px 12px;
                font-size: 12px;
                font-weight: 500;
                min-width: 50px;
                min-height: 28px;
            }

            QPushButton#compactButton:hover {
                background-color: #dee2e6;
                color: #212529;
            }

            QPushButton#compactButton:pressed {
                background-color: #ced4da;
                padding: 5px 11px 3px 13px;
            }

            /* ===== 文本编辑器 ===== */
            QTextEdit {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                background-color: #ffffff;
                padding: 10px;
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 11px;
                color: #212529;
                selection-background-color: #0d6efd;
                selection-color: white;
            }

            QTextEdit:focus {
                border: 2px solid #0d6efd;
            }

            /* 确保 placeholder 文本完整显示 */
            QTextEdit QAbstractScrollArea {
                padding: 0px;
            }

            /* ===== 列表控件 ===== */
            QListWidget {
                border: 2px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                padding: 6px;
                font-size: 12px;
                color: #212529;
                outline: none;
            }

            QListWidget:focus {
                border: 2px solid #0d6efd;
            }

            QListWidget::item {
                padding: 6px 10px;
                border-radius: 4px;
                margin: 1px 0;
                border: none;
            }

            QListWidget::item:hover {
                background-color: #f8f9fa;
            }

            QListWidget::item:selected {
                background-color: #0d6efd;
                color: white;
            }

            QListWidget::item:selected:hover {
                background-color: #0b5ed7;
            }

            /* ===== 进度条 ===== */
            QProgressBar {
                border: none;
                border-radius: 8px;
                text-align: center;
                background-color: #e9ecef;
                height: 8px;
                font-size: 11px;
                color: #495057;
            }

            QProgressBar::chunk {
                background-color: #0d6efd;
                border-radius: 8px;
            }

            /* ===== 状态标签 ===== */
            QLabel#statusLabel {
                background-color: #e7f1ff;
                border: none;
                border-radius: 8px;
                padding: 12px 16px;
                color: #084298;
                font-size: 13px;
                font-weight: 500;
            }

            QLabel#sectionTitle {
                font-size: 13px;
                font-weight: 600;
                color: #212529;
                padding: 4px 0;
            }

            /* ===== 滚动条 ===== */
            QScrollBar:vertical {
                background-color: #f8f9fa;
                width: 12px;
                border-radius: 6px;
            }

            QScrollBar::handle:vertical {
                background-color: #adb5bd;
                border-radius: 6px;
                min-height: 30px;
            }

            QScrollBar::handle:vertical:hover {
                background-color: #6c757d;
            }

            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar:horizontal {
                background-color: #f8f9fa;
                height: 12px;
                border-radius: 6px;
            }

            QScrollBar::handle:horizontal {
                background-color: #adb5bd;
                border-radius: 6px;
                min-width: 30px;
            }

            QScrollBar::handle:horizontal:hover {
                background-color: #6c757d;
            }

            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
        """)

        
    def setup_ui(self):
        # 创建中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # ===== 顶部：输入区域 =====
        input_container = QWidget()
        input_container.setObjectName("cardContainer")
        input_layout = QVBoxLayout(input_container)
        input_layout.setSpacing(12)
        input_layout.setContentsMargins(16, 16, 16, 16)

        # 统一的输入区域（文件路径或URL）
        input_section = QWidget()
        input_section_layout = QVBoxLayout(input_section)
        input_section_layout.setSpacing(6)
        input_section_layout.setContentsMargins(0, 0, 0, 0)

        input_label = QLabel("文件或URL")
        input_label.setObjectName("sectionTitle")
        input_section_layout.addWidget(input_label)

        input_control_layout = QHBoxLayout()
        input_control_layout.setSpacing(10)

        self.file_entry = DragDropLineEdit()
        self.file_entry.setPlaceholderText("选择文件、拖拽文件到此处，或输入URL...")
        # 总高度 = min-height(20) + padding(0*2) + border(2*2) = 24px
        # 但为了垂直居中文字，使用稍大的高度
        self.file_entry.setFixedHeight(36)
        input_control_layout.addWidget(self.file_entry, stretch=1)

        browse_btn = QPushButton("浏览")
        browse_btn.setObjectName("browseButton")
        browse_btn.setMinimumWidth(80)
        # 设置和输入框完全相同的高度
        browse_btn.setFixedHeight(36)
        browse_btn.clicked.connect(self.browse_file)
        input_control_layout.addWidget(browse_btn)

        input_section_layout.addLayout(input_control_layout)
        input_layout.addWidget(input_section)

        main_layout.addWidget(input_container)

        # ===== 默认输出路径区域（放在输入与转换之间） =====
        output_section = QWidget()
        output_section_layout = QHBoxLayout(output_section)
        output_section_layout.setSpacing(8)
        output_section_layout.setContentsMargins(0, 0, 0, 0)

        self.default_output_checkbox = QCheckBox("使用默认输出路径（可编辑，默认桌面）")
        self.default_output_checkbox.stateChanged.connect(self.toggle_default_output)
        output_section_layout.addWidget(self.default_output_checkbox)

        self.output_dir_edit = QLineEdit(str(self.last_output_dir))
        self.output_dir_edit.setPlaceholderText("选择或输入输出目录")
        self.output_dir_edit.setMinimumWidth(260)
        output_section_layout.addWidget(self.output_dir_edit, stretch=1)

        output_browse_btn = QPushButton("选择路径")
        output_browse_btn.setObjectName("browseButton")
        output_browse_btn.setFixedHeight(30)
        output_browse_btn.clicked.connect(self.select_output_dir)
        output_section_layout.addWidget(output_browse_btn)

        # 初始未勾选时隐藏路径输入和按钮
        self.output_dir_edit.hide()
        output_browse_btn.hide()
        self.output_dir_browse_btn = output_browse_btn  # 保存引用用于显隐控制

        main_layout.addWidget(output_section)

        # ===== Excel Sheet 选择区域（初始隐藏）=====
        self.excel_container = QWidget()
        self.excel_container.setObjectName("cardContainer")
        excel_main_layout = QVBoxLayout(self.excel_container)
        excel_main_layout.setSpacing(8)
        excel_main_layout.setContentsMargins(16, 12, 16, 12)

        # 标题和按钮在同一行
        excel_header_layout = QHBoxLayout()
        excel_header_layout.setSpacing(10)

        excel_title = QLabel("Excel Sheet 选择")
        excel_title.setObjectName("sectionTitle")
        excel_header_layout.addWidget(excel_title)

        excel_header_layout.addStretch()

        # 使用更小的按钮
        select_all_btn = QPushButton("全选")
        select_all_btn.setObjectName("compactButton")
        select_all_btn.clicked.connect(self.select_all_sheets)
        excel_header_layout.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("全不选")
        deselect_all_btn.setObjectName("compactButton")
        deselect_all_btn.clicked.connect(self.deselect_all_sheets)
        excel_header_layout.addWidget(deselect_all_btn)

        invert_btn = QPushButton("反选")
        invert_btn.setObjectName("compactButton")
        invert_btn.clicked.connect(self.invert_sheet_selection)
        excel_header_layout.addWidget(invert_btn)

        excel_main_layout.addLayout(excel_header_layout)

        # Sheet 列表
        self.sheet_listbox = QListWidget()
        self.sheet_listbox.setSelectionMode(QAbstractItemView.MultiSelection)
        self.sheet_listbox.setMinimumHeight(100)
        self.sheet_listbox.setMaximumHeight(150)
        excel_main_layout.addWidget(self.sheet_listbox)

        main_layout.addWidget(self.excel_container)
        self.excel_container.hide()  # 初始隐藏

        # ===== 操作按钮区域 =====
        button_container = QWidget()
        button_main_layout = QVBoxLayout(button_container)
        button_main_layout.setSpacing(8)
        button_main_layout.setContentsMargins(0, 0, 0, 0)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        # 主转换按钮 - 更突出
        convert_btn = QPushButton("转换为Markdown")
        convert_btn.setMinimumHeight(42)
        convert_btn.setMinimumWidth(140)
        convert_btn.clicked.connect(self.convert_file)
        button_layout.addWidget(convert_btn)

        self.copy_btn = QPushButton("复制结果")
        self.copy_btn.setObjectName("secondaryButton")
        self.copy_btn.setMinimumHeight(38)
        self.copy_btn.setMinimumWidth(90)
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(self.copy_result)
        button_layout.addWidget(self.copy_btn)

        # 添加弹性空间
        button_layout.addStretch()

        # 次要按钮 - 更小更紧凑
        save_btn = QPushButton("保存结果")
        save_btn.setObjectName("successButton")
        save_btn.setMinimumHeight(38)
        save_btn.setMinimumWidth(90)
        save_btn.clicked.connect(self.save_result)
        button_layout.addWidget(save_btn)

        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.setObjectName("secondaryButton")
        self.refresh_btn.setMinimumHeight(38)
        self.refresh_btn.setMinimumWidth(70)
        self.refresh_btn.clicked.connect(self.refresh_file)
        button_layout.addWidget(self.refresh_btn)

        self.clear_btn = QPushButton("清空")
        self.clear_btn.setObjectName("secondaryButton")
        self.clear_btn.setMinimumHeight(38)
        self.clear_btn.setMinimumWidth(70)
        self.clear_btn.clicked.connect(self.clear_result)
        button_layout.addWidget(self.clear_btn)

        button_main_layout.addLayout(button_layout)

        # 进度条（初始状态隐藏）
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # 无限进度条
        self.progress.setMinimumHeight(6)
        button_main_layout.addWidget(self.progress)
        self.progress.hide()  # 初始隐藏

        main_layout.addWidget(button_container)

        # ===== 结果显示区域 =====
        result_container = QWidget()
        result_container.setObjectName("cardContainer")
        result_main_layout = QVBoxLayout(result_container)
        result_main_layout.setSpacing(10)
        result_main_layout.setContentsMargins(16, 16, 16, 16)

        result_title = QLabel("转换结果")
        result_title.setObjectName("sectionTitle")
        result_main_layout.addWidget(result_title)

        self.result_text = DragDropTextEdit()
        self.result_text.setPlaceholderText("转换结果将显示在这里...\n\n您也可以直接拖拽文件到此处进行转换。")
        self.result_text.setFont(QFont("Consolas", 10))
        self.result_text.setMinimumHeight(180)
        # 设置文档边距，避免文字被裁剪
        self.result_text.document().setDocumentMargin(5)
        result_main_layout.addWidget(self.result_text)

        main_layout.addWidget(result_container, stretch=1)

        # ===== 状态栏 =====
        self.status_label = QLabel("就绪 - 请选择文件或输入URL")
        self.status_label.setObjectName("statusLabel")
        main_layout.addWidget(self.status_label)
        
    def browse_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "选择要转换的文件",
            "",
            "所有支持的文件 (*.pdf *.docx *.pptx *.xlsx *.csv *.html *.epub *.jpg *.png);;PDF文件 (*.pdf);;Word文档 (*.docx);;PowerPoint (*.pptx);;Excel文件 (*.xlsx *.xls);;图像文件 (*.jpg *.jpeg *.png *.gif *.bmp);;所有文件 (*.*)"
        )
        if filename:
            self.file_entry.setText(filename)
            self._check_excel_file(filename)
    
    def handle_file_drop(self, file_path):
        """处理文件拖拽"""
        self.file_entry.setText(file_path)
        self._check_excel_file(file_path)
            
    def convert_file(self):
        source = self.file_entry.text().strip()

        if not source:
            QMessageBox.warning(self, "错误", "请选择文件或输入URL")
            return

        # 在后台线程中执行转换
        selected_sheets = self._get_selected_sheets() if self.current_excel_file else None
        
        self.worker = ConversionWorker(self.md, source, self.current_excel_file, selected_sheets)
        self.worker.finished.connect(self._conversion_complete)
        self.worker.error.connect(self._conversion_error)
        
        self._start_conversion()
        self.worker.start()
        
    def _start_conversion(self):
        self.progress.show()  # 显示进度条
        self.status_label.setText("正在转换...")
        self.result_text.clear()
        self._set_btn_state(self.copy_btn, False)
        self._set_btn_state(self.refresh_btn, False)
        self._set_btn_state(self.clear_btn, False)
        self._set_btn_state(self.refresh_btn, True)
        self._set_btn_state(self.clear_btn, True)

    def _conversion_complete(self, markdown_content, source):
        self.progress.hide()  # 隐藏进度条
        self.status_label.setText(f"转换完成: {Path(source).name if not source.startswith('http') else source}")
        
        # 显示结果
        self.result_text.setPlainText(markdown_content)
        
        # 存储结果用于保存
        self.current_result = markdown_content
        
        # 根据源文件生成标题
        if source.startswith('http'):
            self.current_title = "web_content"
        else:
            # 使用原文件名（不含扩展名）作为标题
            source_path = Path(source)
            self.current_title = source_path.stem  # 文件名不含扩展名

        # 启用复制按钮
        self._set_btn_state(self.copy_btn, True)
        self._set_btn_state(self.refresh_btn, True)
        self._set_btn_state(self.clear_btn, True)

        # 如果之前触发了“自动转换后保存”，转换完成后自动执行保存
        if self.pending_save:
            self.pending_save = False
            self.save_result()
    
    def _conversion_error(self, error_message):
        self.progress.hide()
        self.status_label.setText(f"转换失败: {error_message}")
        QMessageBox.critical(self, "转换错误", error_message)
        self._set_btn_state(self.copy_btn, False)
        self._set_btn_state(self.refresh_btn, True)
        self._set_btn_state(self.clear_btn, True)
        
    def _sanitize_filename(self, filename):
        """清理文件名中的非法字符"""
        # Windows文件名非法字符
        illegal_chars = r'[<>:"/\\|?*]'
        # 替换非法字符为下划线
        sanitized = re.sub(illegal_chars, '_', filename)
        # 移除多余的空格和点
        sanitized = sanitized.strip('. ')
        # 如果文件名为空，使用默认名称
        if not sanitized:
            sanitized = "converted_document"
        return sanitized

    def save_result(self):
        # 如果还没有结果，先自动转换，再自动保存
        if not self.current_result:
            source = self.file_entry.text().strip()
            if not source:
                QMessageBox.warning(self, "警告", "请选择文件或输入URL")
                return
            self.pending_save = True
            self.status_label.setText("未转换，正在自动转换后保存...")
            self.convert_file()
            return
        # 如果是 Excel 并且有选中的 Sheet，则批量保存为多个文件
        if self.current_excel_file and self._get_selected_sheets():
            selected_sheets = self._get_selected_sheets()
            if self.use_default_output:
                folder = self._get_output_dir()
                folder.mkdir(parents=True, exist_ok=True)
            else:
                folder = QFileDialog.getExistingDirectory(self, "选择保存文件夹", str(self._get_output_dir()))
                if not folder:
                    return

            base_title = self._sanitize_filename(self.current_title or Path(self.current_excel_file).stem)
            success, failed = [], []

            try:
                workbook = openpyxl.load_workbook(self.current_excel_file, read_only=True)
                for sheet_name in selected_sheets:
                    try:
                        if sheet_name not in workbook.sheetnames:
                            failed.append((sheet_name, "Sheet 不存在"))
                            continue
                        worksheet = workbook[sheet_name]
                        markdown_content = self._worksheet_to_markdown(worksheet, sheet_name)

                        sheet_suffix = self._sanitize_filename(sheet_name)
                        file_path = Path(folder) / f"{base_title}_{sheet_suffix}.md"
                        with open(file_path, "w", encoding="utf-8") as f:
                            f.write(markdown_content)
                        success.append(file_path.name)
                    except Exception as e:
                        failed.append((sheet_name, str(e)))
                workbook.close()
            except Exception as e:
                QMessageBox.critical(self, "保存错误", f"处理 Excel 时出错: {e}")
                return

            if success:
                self.status_label.setText(f"已保存 {len(success)} 个文件到: {folder}")
                QMessageBox.information(
                    self,
                    "保存成功",
                    f"成功保存 {len(success)} 个文件。\n位置：{folder}"
                )
                # 记录最近的保存目录
                self.last_output_dir = Path(folder)
                self.output_dir_edit.setText(str(self.last_output_dir))
            if failed:
                fail_msg = "\n".join([f"{name}: {err}" for name, err in failed])
                QMessageBox.warning(self, "部分失败", f"下列 Sheet 保存失败：\n{fail_msg}")
        else:
            # 常规单文件保存
            clean_title = self._sanitize_filename(self.current_title)
            if self.use_default_output:
                folder = self._get_output_dir()
                folder.mkdir(parents=True, exist_ok=True)
                filename = folder / f"{clean_title}.md"
            else:
                filename, _ = QFileDialog.getSaveFileName(
                    self,
                    "保存Markdown文件",
                    str(self._get_output_dir() / f"{clean_title}.md"),
                    "Markdown文件 (*.md);;文本文件 (*.txt);;所有文件 (*.*)"
                )
            
            if filename:
                try:
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(self.current_result)
                    self.status_label.setText(f"已保存: {Path(filename).name}")
                    QMessageBox.information(self, "成功", f"文件已保存到: {filename}")
                    # 记录最近的保存目录
                    self.last_output_dir = Path(filename).parent
                    self.output_dir_edit.setText(str(self.last_output_dir))
                except Exception as e:
                    QMessageBox.critical(self, "保存错误", f"保存文件失败: {str(e)}")
                
    def clear_result(self):
        self.result_text.clear()
        self.file_entry.clear()
        self.status_label.setText(f"就绪 - 请选择文件或输入URL（版本 {VERSION}）")
        self.current_result = ""
        self.pending_save = False
        # 清空后所有操作按钮置为不可用
        self._set_btn_state(self.copy_btn, False)
        self._set_btn_state(self.refresh_btn, False)
        self._set_btn_state(self.clear_btn, False)

        # 隐藏 Excel 选择区域
        self._reset_excel_state()

    def refresh_file(self):
        """重新读取当前文件，刷新 Excel Sheet 列表"""
        source = self.file_entry.text().strip()
        if not source:
            QMessageBox.information(self, "提示", "请先选择文件或输入URL")
            return

        # URL：刷新时仅重置 Excel 状态
        if source.startswith("http://") or source.startswith("https://"):
            self._reset_excel_state()
        else:
            # 本地文件：重新检查是否为 Excel
            self._check_excel_file(source)

        self.status_label.setText(f"已刷新: {Path(source).name if not source.startswith('http') else source}")

        # 刷新后清空预览
        self.result_text.clear()
        self.current_result = ""
        self.pending_save = False
        self._set_btn_state(self.copy_btn, False)
        self._set_btn_state(self.refresh_btn, True)
        self._set_btn_state(self.clear_btn, True)

    def copy_result(self):
        """复制当前 Markdown 结果到剪贴板"""
        if not self.current_result:
            return
        QApplication.clipboard().setText(self.current_result)
        self.status_label.setText("已复制到剪贴板")

    def _reset_excel_state(self):
        """统一重置/隐藏 Excel 相关状态"""
        self.excel_container.hide()
        self.current_excel_file = None
        self.excel_sheets = []
        self.selected_sheets = []

    def _get_default_output_dir(self):
        """获取桌面路径作为默认输出目录，若不存在则退回用户主目录"""
        desktop = Path.home() / "Desktop"
        return desktop if desktop.exists() else Path.home()

    def _get_output_dir(self):
        """从输入框获取输出目录，若为空则使用最近目录"""
        text = self.output_dir_edit.text().strip()
        return Path(text) if text else self.last_output_dir

    def toggle_default_output(self, state):
        self.use_default_output = bool(state)
        # 根据勾选状态显隐输入框与按钮
        if self.use_default_output:
            self.output_dir_edit.setText(str(self.last_output_dir))
            self.output_dir_edit.show()
            self.output_dir_browse_btn.show()
        else:
            self.output_dir_edit.hide()
            self.output_dir_browse_btn.hide()

    def _set_btn_state(self, btn: QPushButton, enabled: bool):
        """统一控制按钮颜色：可用绿色，禁用灰色"""
        btn.setEnabled(enabled)
        if enabled:
            btn.setStyleSheet("background-color: #198754; color: white;")
        else:
            btn.setStyleSheet("background-color: #e9ecef; color: #adb5bd;")

    def select_output_dir(self):
        """选择输出目录并更新最近目录"""
        current_dir = Path(self.output_dir_edit.text().strip() or self.last_output_dir)
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", str(current_dir))
        if folder:
            self.last_output_dir = Path(folder)
            self.output_dir_edit.setText(str(self.last_output_dir))

    def _check_excel_file(self, filename):
        """检查是否为 Excel 文件，如果是则显示 sheet 选择"""
        if not EXCEL_SUPPORT:
            return

        file_ext = Path(filename).suffix.lower()
        if file_ext in ['.xlsx', '.xls']:
            try:
                self.current_excel_file = filename
                self._load_excel_sheets(filename)
                self.excel_container.show()  # 显示 Excel 选择区域
            except Exception as e:
                QMessageBox.critical(self, "Excel 文件错误", f"无法读取 Excel 文件: {str(e)}")
        else:
            self.excel_container.hide()  # 隐藏 Excel 选择区域
            self.current_excel_file = None

    def _load_excel_sheets(self, filename):
        """加载 Excel 文件的所有 sheet"""
        try:
            workbook = openpyxl.load_workbook(filename, read_only=True)
            self.excel_sheets = workbook.sheetnames
            workbook.close()
            
            # 更新 listbox
            self.sheet_listbox.clear()
            for sheet in self.excel_sheets:
                item = QListWidgetItem(sheet)
                self.sheet_listbox.addItem(item)
            
            # 默认选择所有 sheet
            self.select_all_sheets()
            
        except Exception as e:
            raise Exception(f"读取 Excel 文件失败: {str(e)}")

    def select_all_sheets(self):
        """选择所有 sheet"""
        for i in range(self.sheet_listbox.count()):
            self.sheet_listbox.item(i).setSelected(True)

    def deselect_all_sheets(self):
        """取消选择所有 sheet"""
        for i in range(self.sheet_listbox.count()):
            self.sheet_listbox.item(i).setSelected(False)

    def invert_sheet_selection(self):
        """反选 sheet"""
        for i in range(self.sheet_listbox.count()):
            item = self.sheet_listbox.item(i)
            item.setSelected(not item.isSelected())

    def _get_selected_sheets(self):
        """获取选中的 sheet 名称列表"""
        selected_sheets = []
        for i in range(self.sheet_listbox.count()):
            item = self.sheet_listbox.item(i)
            if item.isSelected():
                selected_sheets.append(item.text())
        return selected_sheets

    def _worksheet_to_markdown(self, worksheet, sheet_name):
        """将 Excel worksheet 转换为 Markdown（供直接保存时复用）"""
        markdown = f"# {sheet_name}\n\n"

        # 获取有数据的区域
        if worksheet.max_row == 1 and worksheet.max_column == 1:
            return markdown + "此 Sheet 为空\n"

        # 转换为表格
        rows = []
        for row in worksheet.iter_rows(values_only=True):
            # 跳过完全空的行
            if all(cell is None or str(cell).strip() == '' for cell in row):
                continue
            # 将 None 值转换为空字符串，其他值转换为字符串
            row_data = [str(cell) if cell is not None else '' for cell in row]
            rows.append(row_data)

        if not rows:
            return markdown + "此 Sheet 为空\n"

        # 确定最大列数
        max_cols = max(len(row) for row in rows) if rows else 0

        # 补齐所有行到相同列数
        for row in rows:
            while len(row) < max_cols:
                row.append('')

        # 生成 Markdown 表格
        if rows:
            header = "| " + " | ".join(rows[0]) + " |"
            separator = "| " + " | ".join(['---'] * len(rows[0])) + " |"
            markdown += header + "\n" + separator + "\n"

            for row in rows[1:]:
                markdown += "| " + " | ".join(row) + " |\n"

        return markdown


def main():
    app = QApplication(sys.argv)
    window = MarkItDownUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()




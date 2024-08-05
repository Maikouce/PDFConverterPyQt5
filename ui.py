import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QProgressBar, QFileDialog, QLabel, QComboBox, QTextEdit, QTabWidget
from PyQt5.QtCore import pyqtSlot
from converter import PDFConverterThread

class PDFConverterUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

        # 初始化变量
        self.input_dir_image_to_pdf = None
        self.output_dir_image_to_pdf = None
        self.input_dir_pdf_to_image = None
        self.output_dir_pdf_to_image = None
        self.input_dir_word_to_pdf = None
        self.output_dir_word_to_pdf = None
        self.thread_image_to_pdf = None
        self.thread_pdf_to_image = None
        self.thread_word_to_pdf = None
        # self.mutex = QMutex()

    def initUI(self):
        self.setWindowTitle('MC文件转换器4.0')

        # 样式表
        self.setStyleSheet("""
            QWidget {
                font-family: Arial;
                font-size: 14px;
            }
            QLabel {
                font-weight: bold;
                margin-bottom: 10px;
            }
            QPushButton {
                padding: 10px;
                font-size: 14px;
                border-radius: 5px;
                background-color: #4CAF50;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QProgressBar {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                text-align: center;
            }
            QTextEdit {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                padding: 5px;
            }
            QComboBox {
                border: 1px solid #ccc;
                border-radius: 6px;
                padding: 5px;
                background-color: #ffffff;
                margin: 5px;
                min-width: 200px;
                font-size: 14px;
                background: linear-gradient(to bottom, #ffffff 0%, #f2f2f2 100%);
            }
            QComboBox::drop-down {
                border-left: 1px solid #ccc;
                border-radius: 0;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(icons/down_arrow_gradient.png); /* 更具设计感的下拉箭头图标 */
            }
            QComboBox QAbstractItemView {
                border: 1px solid #ccc;
                border-radius: 6px;
                background-color: #ffffff;
                selection-background-color: #007bff;
                selection-color: white;
            }
            QComboBox QAbstractItemView::item {
                padding: 8px;
            }
        """)

        # 创建选项卡
        self.tabs = QTabWidget()
        self.tabs.addTab(self.createImageToPDFTab(), "图片转PDF")
        self.tabs.addTab(self.createPDFToImageTab(), "PDF转图片")
        self.tabs.addTab(self.createWordToPDFTab(), "Word转PDF")  # 新增的选项卡
        self.tabs.addTab(self.createPDFToWordTab(), "PDF转Word")

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)

        self.setLayout(main_layout)

        # 初始日志
        self.log_image_to_pdf(
            '**软件声明**\n本软件系山东建筑大学测绘工程专业在读本科生孟祥鑫原创开发，主要功能为实现JPG及PNG图像文件的批量转换至PDF格式。该软件旨在为用户提供便捷的文件转换解决方案，并且为非商业用途免费提供使用。\n\n如您对该软件有任何疑问、建议或合作意向，请通过电子邮件地址 mengxiangxin766@gmail.com 与开发者取得联系。\n\n**免责声明**\n本软件为非盈利性质，开发者不对因使用本软件而产生的任何直接或间接后果承担责任。用户在使用过程中若遇到任何问题或损失，开发者不承担法律责任。\n')
        self.log_pdf_to_image(
            '**软件声明**\n本软件系山东建筑大学测绘工程专业在读本科生孟祥鑫原创开发，主要功能是实现PDF文件的批量转换至JPG及PNG格式。该软件旨在为用户提供便捷的文件转换解决方案，并且为非商业用途免费提供使用。\n\n如您对该软件有任何疑问、建议或合作意向，请通过电子邮件地址 mengxiangxin766@gmail.com 与开发者取得联系。\n\n**免责声明**\n本软件为非盈利性质，开发者不对因使用本软件而产生的任何直接或间接后果承担责任。用户在使用过程中若遇到任何问题或损失，开发者不承担法律责任。\n')
        self.log_word_to_pdf(
            '**软件声明**\n本软件系山东建筑大学测绘工程专业在读本科生孟祥鑫原创开发，主要功能为实现Word文件的批量转换至PDF格式。该软件旨在为用户提供便捷的文件转换解决方案，并且为非商业用途免费提供使用。\n\n如您对该软件有任何疑问、建议或合作意向，请通过电子邮件地址 mengxiangxin766@gmail.com 与开发者取得联系。\n\n**免责声明**\n本软件为非盈利性质，开发者不对因使用本软件而产生的任何直接或间接后果承担责任。用户在使用过程中若遇到任何问题或损失，开发者不承担法律责任。\n注：注：仅支持.docx格式\n')
        self.log_pdf_to_word(
            '**软件声明**\n本软件系山东建筑大学测绘工程专业在读本科生孟祥鑫原创开发，主要功能为实现PDF文件的批量转换至Word格式。该软件旨在为用户提供便捷的文件转换解决方案，并且为非商业用途免费提供使用。\n\n如您对该软件有任何疑问、建议或合作意向，请通过电子邮件地址 mengxiangxin766@gmail.com 与开发者取得联系。\n\n**免责声明**\n本软件为非盈利性质，开发者不对因使用本软件而产生的任何直接或间接后果承担责任。用户在使用过程中若遇到任何问题或损失，开发者不承担法律责任。\n注：生成文件docx格式\n')

    def createImageToPDFTab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        # 创建输入和输出目录选择布局
        dir_layout = QHBoxLayout()
        self.btnSelectDir = QPushButton('选择输入文件夹', self)
        self.btnSelectDir.clicked.connect(self.selectInputDirectoryImageToPDF)
        dir_layout.addWidget(self.btnSelectDir)

        self.btnSelectOutputDir = QPushButton('选择输出文件夹', self)
        self.btnSelectOutputDir.clicked.connect(self.selectOutputDirectoryImageToPDF)
        dir_layout.addWidget(self.btnSelectOutputDir)
        layout.addLayout(dir_layout)

        # 显示输入和输出路径
        self.input_dir_label = QLabel('输入文件夹: 未选择', self)
        layout.addWidget(self.input_dir_label)

        self.output_dir_label = QLabel('输出文件夹: 未选择', self)
        layout.addWidget(self.output_dir_label)

        # 标签和下拉框在同一行显示
        options_layout = QVBoxLayout()

        hbox1 = QHBoxLayout()
        hbox1.addWidget(QLabel('选择输出格式选项:', self))
        self.location_option_combo = QComboBox(self)
        self.location_option_combo.addItems(['图片同级输出路径', '图片父级输出路径', '输出路径根路径'])
        hbox1.addWidget(self.location_option_combo)
        options_layout.addLayout(hbox1)

        hbox2 = QHBoxLayout()
        hbox2.addWidget(QLabel('选择文件类型:', self))
        self.file_type_combo = QComboBox(self)
        self.file_type_combo.addItems(['JPG', 'PNG'])
        hbox2.addWidget(self.file_type_combo)
        options_layout.addLayout(hbox2)

        hbox3 = QHBoxLayout()
        hbox3.addWidget(QLabel('选择排序方式:', self))
        self.sort_option_combo = QComboBox(self)
        self.sort_option_combo.addItems(['按名称排序', '按时间排序'])
        hbox3.addWidget(self.sort_option_combo)
        options_layout.addLayout(hbox3)

        layout.addLayout(options_layout)

        # 添加转换按钮和进度条
        self.btnConvert = QPushButton('转换图片为PDF', self)
        self.btnConvert.clicked.connect(self.start_conversion_image_to_pdf)
        layout.addWidget(self.btnConvert)

        self.progress = QProgressBar(self)
        layout.addWidget(self.progress)

        # 添加日志输出窗口
        self.log_output_image_to_pdf = QTextEdit(self)
        self.log_output_image_to_pdf.setReadOnly(True)
        layout.addWidget(self.log_output_image_to_pdf)

        widget.setLayout(layout)
        return widget

    def createPDFToImageTab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        # 创建输入和输出目录选择布局
        dir_layout = QHBoxLayout()
        self.btnSelectDirPdfToImage = QPushButton('选择输入文件夹', self)
        self.btnSelectDirPdfToImage.clicked.connect(self.selectInputDirectoryPDFToImage)
        dir_layout.addWidget(self.btnSelectDirPdfToImage)

        self.btnSelectOutputDirPdfToImage = QPushButton('选择输出文件夹', self)
        self.btnSelectOutputDirPdfToImage.clicked.connect(self.selectOutputDirectoryPDFToImage)
        dir_layout.addWidget(self.btnSelectOutputDirPdfToImage)
        layout.addLayout(dir_layout)

        # 显示输入和输出路径
        self.input_dir_label_pdf_to_image = QLabel('输入文件夹: 未选择', self)
        layout.addWidget(self.input_dir_label_pdf_to_image)

        self.output_dir_label_pdf_to_image = QLabel('输出文件夹: 未选择', self)
        layout.addWidget(self.output_dir_label_pdf_to_image)

        # 标签和下拉框在同一行显示
        options_layout = QVBoxLayout()

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel('选择文件类型:', self))
        self.file_type_combo_pdf_to_image = QComboBox(self)
        self.file_type_combo_pdf_to_image.addItems(['JPG', 'PNG'])
        hbox.addWidget(self.file_type_combo_pdf_to_image)
        options_layout.addLayout(hbox)

        layout.addLayout(options_layout)

        # 添加转换按钮和进度条
        self.btnPdfToImage = QPushButton('转换PDF为图片', self)
        self.btnPdfToImage.clicked.connect(self.start_conversion_pdf_to_image)
        layout.addWidget(self.btnPdfToImage)

        self.progress_pdf_to_image = QProgressBar(self)
        layout.addWidget(self.progress_pdf_to_image)

        # 添加日志输出窗口
        self.log_output_pdf_to_image = QTextEdit(self)
        self.log_output_pdf_to_image.setReadOnly(True)
        layout.addWidget(self.log_output_pdf_to_image)

        widget.setLayout(layout)
        return widget

    def createWordToPDFTab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        # 创建输入和输出目录选择布局
        dir_layout = QHBoxLayout()
        self.btnSelectDirWordToPDF = QPushButton('选择输入文件夹', self)
        self.btnSelectDirWordToPDF.clicked.connect(self.selectInputDirectoryWordToPDF)
        dir_layout.addWidget(self.btnSelectDirWordToPDF)

        self.btnSelectOutputDirWordToPDF = QPushButton('选择输出文件夹', self)
        self.btnSelectOutputDirWordToPDF.clicked.connect(self.selectOutputDirectoryWordToPDF)
        dir_layout.addWidget(self.btnSelectOutputDirWordToPDF)
        layout.addLayout(dir_layout)

        # 显示输入和输出路径
        self.input_dir_label_word_to_pdf = QLabel('输入文件夹: 未选择', self)
        layout.addWidget(self.input_dir_label_word_to_pdf)

        self.output_dir_label_word_to_pdf = QLabel('输出文件夹: 未选择', self)
        layout.addWidget(self.output_dir_label_word_to_pdf)

        # 添加转换按钮和进度条
        self.btnWordToPDF = QPushButton('转换Word为PDF', self)
        self.btnWordToPDF.clicked.connect(self.start_conversion_word_to_pdf)
        layout.addWidget(self.btnWordToPDF)

        self.progress_word_to_pdf = QProgressBar(self)
        layout.addWidget(self.progress_word_to_pdf)

        # 添加日志输出窗口
        self.log_output_word_to_pdf = QTextEdit(self)
        self.log_output_word_to_pdf.setReadOnly(True)
        layout.addWidget(self.log_output_word_to_pdf)

        widget.setLayout(layout)
        return widget
    def createPDFToWordTab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        # 创建输入和输出目录选择布局
        dir_layout = QHBoxLayout()
        self.btnSelectDirPdfToWord = QPushButton('选择输入文件夹', self)
        self.btnSelectDirPdfToWord.clicked.connect(self.selectInputDirectoryPDFToWord)
        dir_layout.addWidget(self.btnSelectDirPdfToWord)

        self.btnSelectOutputDirPdfToWord = QPushButton('选择输出文件夹', self)
        self.btnSelectOutputDirPdfToWord.clicked.connect(self.selectOutputDirectoryPDFToWord)
        dir_layout.addWidget(self.btnSelectOutputDirPdfToWord)
        layout.addLayout(dir_layout)

        # 显示输入和输出路径
        self.input_dir_label_pdf_to_word = QLabel('输入文件夹: 未选择', self)
        layout.addWidget(self.input_dir_label_pdf_to_word)

        self.output_dir_label_pdf_to_word = QLabel('输出文件夹: 未选择', self)
        layout.addWidget(self.output_dir_label_pdf_to_word)

        # 添加转换按钮和进度条
        self.btnPDFToWord = QPushButton('转换PDF为Word', self)
        self.btnPDFToWord.clicked.connect(self.start_conversion_pdf_to_word)
        layout.addWidget(self.btnPDFToWord)

        self.progress_pdf_to_word = QProgressBar(self)
        layout.addWidget(self.progress_pdf_to_word)

        # 添加日志输出窗口
        self.log_output_pdf_to_word = QTextEdit(self)
        self.log_output_pdf_to_word.setReadOnly(True)
        layout.addWidget(self.log_output_pdf_to_word)

        widget.setLayout(layout)
        return widget
    def log_image_to_pdf(self, message):
        self.log_output_image_to_pdf.append(message)

    def log_pdf_to_image(self, message):
        self.log_output_pdf_to_image.append(message)

    def log_word_to_pdf(self, message):
        self.log_output_word_to_pdf.append(message)

    def log_pdf_to_word(self, message):
        self.log_output_pdf_to_word.append(message)
    @pyqtSlot()
    def selectInputDirectoryImageToPDF(self):
        self.input_dir_image_to_pdf = QFileDialog.getExistingDirectory(self, '选择输入文件夹')
        self.input_dir_label.setText(f'输入文件夹: {self.input_dir_image_to_pdf}')

    @pyqtSlot()
    def selectOutputDirectoryImageToPDF(self):
        self.output_dir_image_to_pdf = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        self.output_dir_label.setText(f'输出文件夹: {self.output_dir_image_to_pdf}')

    @pyqtSlot()
    def selectInputDirectoryPDFToImage(self):
        self.input_dir_pdf_to_image = QFileDialog.getExistingDirectory(self, '选择输入文件夹')
        self.input_dir_label_pdf_to_image.setText(f'输入文件夹: {self.input_dir_pdf_to_image}')

    @pyqtSlot()
    def selectOutputDirectoryPDFToImage(self):
        self.output_dir_pdf_to_image = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        self.output_dir_label_pdf_to_image.setText(f'输出文件夹: {self.output_dir_pdf_to_image}')

    @pyqtSlot()
    def selectInputDirectoryWordToPDF(self):
        self.input_dir_word_to_pdf = QFileDialog.getExistingDirectory(self, '选择输入文件夹')
        self.input_dir_label_word_to_pdf.setText(f'输入文件夹: {self.input_dir_word_to_pdf}')

    @pyqtSlot()
    def selectOutputDirectoryWordToPDF(self):
        self.output_dir_word_to_pdf = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        self.output_dir_label_word_to_pdf.setText(f'输出文件夹: {self.output_dir_word_to_pdf}')
    @pyqtSlot()
    def selectInputDirectoryPDFToWord(self):
        self.input_dir_pdf_to_word = QFileDialog.getExistingDirectory(self, '选择输入文件夹')
        self.input_dir_label_pdf_to_word.setText(f'输入文件夹: {self.input_dir_pdf_to_word}')

    @pyqtSlot()
    def selectOutputDirectoryPDFToWord(self):
        self.output_dir_pdf_to_word = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        self.output_dir_label_pdf_to_word.setText(f'输出文件夹: {self.output_dir_pdf_to_word}')
    @pyqtSlot()
    def start_conversion_image_to_pdf(self):
        if not self.input_dir_image_to_pdf or not self.output_dir_image_to_pdf:
            self.log_image_to_pdf('请同时选择输入和输出文件夹。')
            return

        self.progress.setValue(0)
        location_option = self.location_option_combo.currentIndex() + 1
        file_type = self.file_type_combo.currentText().lower()
        sort_option = 'name' if self.sort_option_combo.currentIndex() == 0 else 'time'

        # self.mutex.lock()
        self.thread_image_to_pdf = PDFConverterThread(self.input_dir_image_to_pdf, self.output_dir_image_to_pdf, location_option, file_type, sort_option, 'images_to_pdf')
        self.thread_image_to_pdf.progress_update.connect(self.update_progress_image_to_pdf)
        self.thread_image_to_pdf.log_message.connect(self.log_image_to_pdf)
        # self.thread_image_to_pdf.finished.connect(self.mutex.unlock)
        self.thread_image_to_pdf.start()

    @pyqtSlot()
    def start_conversion_pdf_to_image(self):
        if not self.input_dir_pdf_to_image or not self.output_dir_pdf_to_image:
            self.log_pdf_to_image('请同时选择输入和输出文件夹。')
            return

        file_type = self.file_type_combo_pdf_to_image.currentText().lower()

        self.progress_pdf_to_image.setValue(0)

        # self.mutex.lock()
        self.thread_pdf_to_image = PDFConverterThread(self.input_dir_pdf_to_image, self.output_dir_pdf_to_image, None, file_type, None, 'pdf_to_images')
        self.thread_pdf_to_image.progress_update.connect(self.update_progress_pdf_to_image)
        self.thread_pdf_to_image.log_message.connect(self.log_pdf_to_image)
        # self.thread_pdf_to_image.finished.connect(self.mutex.unlock)
        self.thread_pdf_to_image.start()

    @pyqtSlot()
    def start_conversion_word_to_pdf(self):
        if not self.input_dir_word_to_pdf or not self.output_dir_word_to_pdf:
            self.log_word_to_pdf('请同时选择输入和输出文件夹。')
            return

        self.progress_word_to_pdf.setValue(0)

        # self.mutex.lock()
        self.thread_word_to_pdf = PDFConverterThread(self.input_dir_word_to_pdf, self.output_dir_word_to_pdf, None, None, None, 'word_to_pdf')
        self.thread_word_to_pdf.progress_update.connect(self.update_progress_word_to_pdf)
        self.thread_word_to_pdf.log_message.connect(self.log_word_to_pdf)
        # self.thread_word_to_pdf.finished.connect(self.mutex.unlock)
        self.thread_word_to_pdf.start()
    @pyqtSlot()
    def start_conversion_pdf_to_word(self):
        if not self.input_dir_pdf_to_word or not self.output_dir_pdf_to_word:
            self.log_pdf_to_word('请同时选择输入和输出文件夹。')
            return

        self.progress_pdf_to_word.setValue(0)

        # self.mutex.lock()
        self.thread_pdf_to_word = PDFConverterThread(self.input_dir_pdf_to_word, self.output_dir_pdf_to_word, None, None, None, 'pdf_to_word')
        self.thread_pdf_to_word.progress_update.connect(self.update_progress_pdf_to_word)
        self.thread_pdf_to_word.log_message.connect(self.log_pdf_to_word)
        # self.thread_pdf_to_word.finished.connect(self.mutex.unlock)
        self.thread_pdf_to_word.start()
    @pyqtSlot(int)
    def update_progress_image_to_pdf(self, value):
        self.progress.setValue(value)

    @pyqtSlot(int)
    def update_progress_pdf_to_image(self, value):
        self.progress_pdf_to_image.setValue(value)

    @pyqtSlot(int)
    def update_progress_word_to_pdf(self, value):
        self.progress_word_to_pdf.setValue(value)

    @pyqtSlot(int)
    def update_progress_pdf_to_word(self, value):
        self.progress_pdf_to_word.setValue(value)



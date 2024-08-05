import os
import img2pdf
import fitz  # PyMuPDF
from natsort import natsorted
from PyQt5.QtCore import QThread, pyqtSignal
from pdf2docx import Converter
# from docx2pdf import convert as docx_convert  # 导入 docx2pdf 的 convert 函数
import docx2pdf
class PDFConverterThread(QThread):
    progress_update = pyqtSignal(int)
    log_message = pyqtSignal(str)

    def __init__(self, input_dir, output_dir, location_option, file_type, sort_option, task_type):
        super().__init__()
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.location_option = location_option
        self.file_type = file_type
        self.sort_option = sort_option
        self.task_type = task_type

    def run(self):
        try:
            if self.task_type == 'images_to_pdf':
                self.convert_images_to_pdf()
            elif self.task_type == 'pdf_to_images':
                self.convert_pdfs_to_images()
            elif self.task_type == 'word_to_pdf':
                self.convert_word_to_pdf()
            elif self.task_type == 'pdf_to_word':
                self.convert_pdf_to_word()
        except Exception as e:
            self.log_message.emit(f"错误: {str(e)}")

    def convert_images_to_pdf(self):
        image_files = []

        for root, _, files in os.walk(self.input_dir):
            image_files.extend([os.path.join(root, f) for f in files if f.lower().endswith(f'.{self.file_type}')])

        num = len(image_files)
        self.log_message.emit(f"在 {self.input_dir} 和其子目录中找到 {num} 个 .{self.file_type} 文件。")

        if self.sort_option == 'name':
            image_files = natsorted(image_files)
        elif self.sort_option == 'time':
            image_files.sort(key=lambda x: os.path.getmtime(x))
        else:
            self.log_message.emit("无效的排序选项。")
            return

        i = 0
        for root, _, files in os.walk(self.input_dir):
            image_files_in_dir = [os.path.join(root, f) for f in files if f.lower().endswith(f'.{self.file_type}')]

            image_files_in_dir.sort(key=lambda x: image_files.index(x))

            if image_files_in_dir:
                rel_dir = os.path.relpath(root, self.input_dir)

                if self.location_option == 1:
                    output_pdf_dir = os.path.join(self.output_dir, rel_dir)
                elif self.location_option == 2:
                    output_pdf_dir = os.path.join(self.output_dir, os.path.dirname(rel_dir))
                elif self.location_option == 3:
                    output_pdf_dir = self.output_dir
                else:
                    self.log_message.emit("无效的位置选项。")
                    return

                os.makedirs(output_pdf_dir, exist_ok=True)

                rel_path_parts = rel_dir.split(os.sep)
                pdf_filename = "_".join(rel_path_parts) + ".pdf"
                pdf_path = os.path.join(output_pdf_dir, pdf_filename)

                current_dir_images = [img for img in image_files if os.path.dirname(img) == root]

                if current_dir_images:
                    self.log_message.emit(f'正在转换图片为 {pdf_filename}')
                    try:
                        with open(pdf_path, "wb") as f:
                            f.write(img2pdf.convert(current_dir_images, rotation=img2pdf.Rotation.ifvalid))
                        self.log_message.emit(f'PDF 创建成功: {pdf_path}')
                    except Exception as e:
                        self.log_message.emit(f'创建 PDF 时发生错误: {e}')

                    i += len(current_dir_images)
                    progress_value = int((i / num) * 100)
                    self.progress_update.emit(progress_value)
                    self.log_message.emit(f'当前进度: {progress_value}%')

        self.progress_update.emit(100)
        self.log_message.emit(f"图片转换为PDF完成。")

    def convert_pdfs_to_images(self):
        pdf_files = []

        for root, _, files in os.walk(self.input_dir):
            pdf_files.extend([os.path.join(root, f) for f in files if f.lower().endswith('.pdf')])

        num_pdfs = len(pdf_files)
        self.log_message.emit(f"在 {self.input_dir} 和其子目录中找到 {num_pdfs} 个 PDF 文件。")

        i = 0
        for pdf_file in pdf_files:
            pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
            pdf_output_dir = os.path.join(self.output_dir, pdf_name)
            os.makedirs(pdf_output_dir, exist_ok=True)

            self.log_message.emit(f"开始转换 {pdf_file} 的每一页到图片。")

            doc = fitz.open(pdf_file)
            num_pages = doc.page_count
            image_extension = 'jpg' if self.file_type == 'jpg' else 'png'

            for j in range(num_pages):
                page = doc.load_page(j)
                pix = page.get_pixmap()
                output_image_path = os.path.join(pdf_output_dir, f'page_{j + 1}.{image_extension}')
                pix.save(output_image_path)

                progress = int(((i + j / num_pages) / num_pdfs) * 100)
                self.progress_update.emit(progress)
                self.log_message.emit(f"生成图片文件：{output_image_path}")

            i += 1

        self.progress_update.emit(100)
        self.log_message.emit(f"PDF转换为图片完成。")

    def convert_word_to_pdf(self):
        word_files = []

        for root, _, files in os.walk(self.input_dir):
            word_files.extend([os.path.join(root, f) for f in files if f.lower().endswith('.docx')])

        num_words = len(word_files)
        self.log_message.emit(f"在 {self.input_dir} 和其子目录中找到 {num_words} 个 DOCX 文件。")

        i = 0
        for word_file in word_files:
            output_pdf_path = os.path.join(self.output_dir, f'{os.path.splitext(os.path.basename(word_file))[0]}.pdf')

            try:
                docx2pdf.convert(word_file, output_pdf_path)
                self.log_message.emit(f"生成PDF文件：{output_pdf_path}")
            except Exception as e:
                self.log_message.emit(f"转换错误 {word_file}：{str(e)}")

            i += 1
            progress = int((i / num_words) * 100)
            self.progress_update.emit(progress)

        self.progress_update.emit(100)
        self.log_message.emit(f"Word转换为PDF完成。")

    def convert_pdf_to_word(self):
        pdf_files = []

        for root, _, files in os.walk(self.input_dir):
            pdf_files.extend([os.path.join(root, f) for f in files if f.lower().endswith('.pdf')])

        num_pdfs = len(pdf_files)
        self.log_message.emit(f"在 {self.input_dir} 和其子目录中找到 {num_pdfs} 个 PDF 文件。")

        i = 0
        for pdf_file in pdf_files:
            output_docx_path = os.path.join(self.output_dir, f'{os.path.splitext(os.path.basename(pdf_file))[0]}.docx')

            try:
                self.log_message.emit(f'正在转换 {pdf_file} 为 Word 文件。')
                cv = Converter(pdf_file)
                cv.convert(output_docx_path, start=0, end=None)
                cv.close()
                self.log_message.emit(f'Word 文件创建成功: {output_docx_path}')
            except Exception as e:
                self.log_message.emit(f'转换错误 {pdf_file}：{str(e)}')

            i += 1
            progress = int((i / num_pdfs) * 100)
            self.progress_update.emit(progress)

        self.progress_update.emit(100)
        self.log_message.emit(f"PDF转换为Word完成。")

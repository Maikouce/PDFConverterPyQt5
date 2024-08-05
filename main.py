import sys
from PyQt5.QtWidgets import QApplication
from ui import PDFConverterUI

if __name__ == '__main__':
    app = QApplication(sys.argv)
    converter = PDFConverterUI()
    converter.resize(500, 720)  # 增加窗口高度
    converter.show()
    sys.exit(app.exec_())

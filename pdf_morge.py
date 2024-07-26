import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QPushButton, QVBoxLayout, QWidget,
                             QMessageBox)
from PyPDF2 import PdfReader, PdfWriter


class PDFTool(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # Create widgets
        self.merge_button = QPushButton('Merge PDF', self)
        self.merge_button.clicked.connect(self.merge_pdfs)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.merge_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Window settings
        self.setGeometry(100, 100, 300, 100)
        self.setWindowTitle('PDF Merge Tool')

    def merge_pdfs(self):
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, 'Select PDF Files to Merge', '', 'PDF Files (*.pdf)')

            if file_paths:
                pdf_writer = PdfWriter()

                for file_path in file_paths:
                    pdf_reader = PdfReader(file_path)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)

                output_file, _ = QFileDialog.getSaveFileName(self, 'Save Merged PDF', '', 'PDF Files (*.pdf)')
                if output_file:
                    with open(output_file, 'wb') as output:
                        pdf_writer.write(output)

                    QMessageBox.information(self, 'PDF Merge', 'PDFs merged successfully!')
        except Exception as e:
            print(f"Error during PDF merge: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    pdf_tool = PDFTool()
    pdf_tool.show()
    sys.exit(app.exec_())

import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter


def convert_pdf_to_docx(pdf_path, docx_path):
    obj = Converter(pdf_path)
    obj.convert(docx_path)
    obj.close()











class PDFToWordApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF to Word Converter")
        self.geometry("600x200")

        # Widgets
        tk.Label(self, text="Select PDF file:").grid(row=0, column=0, padx=10, pady=10)
        self.pdf_entry = tk.Entry(self, width=40)
        self.pdf_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self, text="Browse", command=self.browse_pdf).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self, text="Select output DOCX file:").grid(row=1, column=0, padx=10, pady=10)
        self.docx_entry = tk.Entry(self, width=40)
        self.docx_entry.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(self, text="Convert", command=self.convert_pdf).grid(row=2, column=0, columnspan=3, pady=20)

        self.result_label = tk.Label(self, text="")
        self.result_label.grid(row=3, column=0, columnspan=3)

    def browse_output_folder(self):
        folder_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        self.docx_entry.delete(0, tk.END)
        self.docx_entry.insert(0, folder_path)

    def browse_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.pdf_entry.delete(0, tk.END)
        self.pdf_entry.insert(0, file_path)

    def convert_pdf(self):
        pdf_path = self.pdf_entry.get()
        docx_path = self.docx_entry.get()

        if pdf_path and docx_path:
            convert_pdf_to_docx(pdf_path, docx_path)
            self.result_label.config(text="Conversion completed successfully!")
        else:
            self.result_label.config(text="Please select PDF and output DOCX paths.")

def main_pdftoword():
    app = PDFToWordApp()
    app.mainloop()


# Run the GUI
if __name__ == "__main__":
    app = PDFToWordApp()
    app.mainloop()

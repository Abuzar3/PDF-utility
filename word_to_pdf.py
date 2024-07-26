import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert


class DocxToPdfConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DOCX to PDF Converter")

        # Create and set up the GUI elements
        self.label = tk.Label(root, text="Select DOCX file:")
        self.label.pack(pady=10)

        self.file_path_var = tk.StringVar()
        self.entry = tk.Entry(root, textvariable=self.file_path_var, width=50)
        self.entry.pack(pady=10)

        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_button.pack(pady=10)

        self.convert_button = tk.Button(root, text="Convert to PDF", command=self.convert_to_pdf)
        self.convert_button.pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("DOCX Files", "*.docx")])
        if file_path:
            self.file_path_var.set(file_path)

    def convert_to_pdf(self):
        docx_file = self.file_path_var.get()

        if docx_file:
            try:
                # Convert DOCX to PDF
                convert(docx_file)
                messagebox.showinfo("Conversion Complete", "PDF file saved successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
        else:
            messagebox.showwarning("Warning", "Please select a DOCX file.")


def main():
    root = tk.Tk()
    app = DocxToPdfConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToPdfConverterApp(root)
    root.mainloop()

import os
import sys
import tkinter as tk
import tkinter.messagebox as messagebox
import tkinter.simpledialog as simpledialog
from tkinter import filedialog

from PyPDF2 import PdfWriter, PdfReader

class PDFProtectorApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF Protector")
        self.geometry("400x300")

        self.input_file_path = ""
        self.output_file_path = ""
        self.password = ""

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Select PDF file to protect:").pack(pady=10)

        tk.Button(self, text="Browse", command=self.browse_file).pack()

        tk.Label(self, text="Set password for PDF file:").pack(pady=10)

        tk.Button(self, text="Set Password", command=self.set_password).pack()

        tk.Button(self, text="Protect PDF", command=self.protect_pdf).pack(pady=20)

    def browse_file(self):
        self.input_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.input_file_path:
            tk.Label(self, text=f"Selected file: {os.path.basename(self.input_file_path)}").pack()

    def set_password(self):
        self.password = simpledialog.askstring("Password", "Enter password:", show="*")

    def protect_pdf(self):
        if not self.input_file_path:
            messagebox.showerror("Error", "Please select a PDF file to protect.")
            return
        if not self.password:
            messagebox.showerror("Error", "Please set a password for the PDF file.")
            return

        self.output_file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not self.output_file_path:
            return

        try:
            self.protect(self.input_file_path, self.output_file_path, self.password)
            messagebox.showinfo("Success", f"PDF file protected successfully. Saved as: {self.output_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def protect(self, ifname, ofname, password):
        with open(ifname, "rb") as ifile, open(ofname, "wb") as ofile:
            reader = PdfReader(ifile)
            writer = PdfWriter()
            for page_number in range(len(reader.pages)):
                writer.add_page(reader.pages[page_number])
            writer.encrypt(password)
            writer.write(ofile)

if __name__ == "__main__":
    app = PDFProtectorApp()
    app.mainloop()
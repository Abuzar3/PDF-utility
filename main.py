import threading

import customtkinter
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from CTkMenuBar import *
from add_page_number import PDFPageNumberApp
from pdf_morge import PDFTool
from pdftoword import main_pdftoword
from word_to_pdf import main

# Root windows setting
customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

current_version = "v1.0.0"


def pdf_to_word_thread():
    pass


class PDFPageDeleter(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        # configure window
        self.title(f"Digna PDF Utility {current_version} By Karbosh")
        self.iconbitmap('karbosh_logo.ico')
        self.geometry(f"{1100}x{600}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # Create Tkinter menus for the dropdowns
        menu = CTkTitleMenu(master=self)
        menu.add_cascade("Check For Update", command=self.check_for_update)
        menu.add_cascade("About", command=self.about)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)

        # Setting the font
        self.font_large = customtkinter.CTkFont(family="Helvetica", size=16)
        self.font_medium = customtkinter.CTkFont(family="Helvetica", size=14)

        self.file_label = customtkinter.CTkLabel(self.sidebar_frame, text="No file selected",
                                                 font=self.font_medium)
        self.file_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.select_button = customtkinter.CTkButton(self.sidebar_frame, text="Select PDF",
                                                     command=self.select_file, font=self.font_medium,
                                                     fg_color='#4CAF50', text_color='white')
        self.select_button.grid(row=1, column=0, padx=20, pady=10)

        self.page_label = customtkinter.CTkLabel(self.sidebar_frame, text="Pages to Delete (comma separated):",
                                                 font=self.font_medium)
        self.page_label.grid(row=2, column=0, padx=20, pady=(20, 10))

        self.page_entry = customtkinter.CTkEntry(self.sidebar_frame, font=self.font_medium)
        self.page_entry.grid(row=3, column=0, padx=20, pady=10)

        self.delete_button = customtkinter.CTkButton(self.sidebar_frame, text="Delete Pages", width=200,
                                                     command=self.delete_pages, font=self.font_medium,
                                                     fg_color='#f44336', text_color='white')
        self.delete_button.grid(row=4, column=0, padx=20, pady=10)
        self.add_page_numbers = customtkinter.CTkButton(self.sidebar_frame, text="Add Page Numbers", width=200,
                                                        command=self.add_page_numbers, font=self.font_medium,
                                                        fg_color='#2196F3', text_color='white')
        self.add_page_numbers.grid(row=5, column=0, padx=20, pady=10)
        self.pdf_merge = customtkinter.CTkButton(self.sidebar_frame, text="PDF Merge", width=200,
                                                 command=self.pdf_merge, font=self.font_medium,
                                                 fg_color='#2196F3', text_color='white')
        self.pdf_merge.grid(row=6, column=0, padx=20, pady=10)
        self.pdftoword = customtkinter.CTkButton(self.sidebar_frame, text="PDF to Word", width=200,
                                                 command=self.pdftoword, font=self.font_medium,
                                                 fg_color='#2196F3', text_color='white')
        self.pdftoword.grid(row=7, column=0, padx=20, pady=10)
        self.wordtopdf = customtkinter.CTkButton(self.sidebar_frame, text="Word to PDF", width=200,
                                                 command=self.wordtopdf, font=self.font_medium,
                                                 fg_color='#2196F3', text_color='white')
        self.wordtopdf.grid(row=8, column=0, padx=20, pady=10)
        self.protectpdf = customtkinter.CTkButton(self.sidebar_frame, text="Protect PDF", width=200,
                                                  command=self.protectpdf, font=self.font_medium,
                                                  fg_color='#2196F3', text_color='white')
        self.protectpdf.grid(row=9, column=0, padx=20, pady=10)

        self.file_path = None

    def check_for_update(self):
        pass

    def about(self):
        pass

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.file_path:
            self.file_label.configure(text=self.file_path)
        else:
            self.file_label.configure(text="No file selected")

    def delete_pages(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected!")
            return

        try:
            pages_to_delete = list(map(int, self.page_entry.get().split(',')))
            # Convert to 0-indexed
            pages_to_delete = [p - 1 for p in pages_to_delete]
            if any(p < 0 for p in pages_to_delete):
                raise ValueError("Page numbers must be positive")
        except ValueError:
            messagebox.showerror("Error", "Invalid page numbers!")
            return

        try:
            reader = PdfReader(self.file_path)
            num_pages = len(reader.pages)

            if any(p >= num_pages for p in pages_to_delete):
                messagebox.showerror("Error", "One or more page numbers are out of range!")
                return

            writer = PdfWriter()
            for i in range(num_pages):
                if i not in pages_to_delete:
                    writer.add_page(reader.pages[i])

            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if output_path:
                with open(output_path, "wb") as output_file:
                    writer.write(output_file)

                messagebox.showinfo("Success",
                                    f"Pages {', '.join(map(str, [p + 1 for p in pages_to_delete]))} deleted "
                                    f"successfully!")
            else:
                messagebox.showwarning("Warning", "Save operation cancelled!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def add_page_numbers(self):
        if not self.file_path:
            messagebox.showerror("Error", "No file selected!")
            return
        PDFPageNumberApp()
        return

    def pdf_merge(self):
        PDFTool()
        return

    def pdftoword(self):
        #pdftoword_thread = threading.Thread(target=pdf_to_word_thread)
        #pdftoword_thread.start()
        main_pdftoword()

    def wordtopdf(self):
        main()
        return

    def protectpdf(self):
        return


if __name__ == "__main__":
    app = PDFPageDeleter()
    app.mainloop()

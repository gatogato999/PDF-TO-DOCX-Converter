import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from customtkinter import CTk, CTkButton, CTkLabel, CTkProgressBar
from PyPDF2 import PdfReader
from docx import Document


class PDFtoDOCXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to DOCX Converter")
        self.root.geometry("500x300")
        self.pdf_file_path = None
        self.output_folder_path = None

        # GUI Components
        self.create_widgets()

    def create_widgets(self):
        # Labels
        CTkLabel(self.root, text="PDF to DOCX Converter", font=("Arial", 18)).pack(pady=10)
        self.status_label = CTkLabel(self.root, text="Select a PDF file to start.", font=("Arial", 12))
        self.status_label.pack(pady=5)

        # Progress Bar
        self.progress_bar = CTkProgressBar(self.root, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)

        # Buttons
        CTkButton(self.root, text="Select PDF", command=self.select_pdf).pack(pady=5)
        CTkButton(self.root, text="Select Output Folder", command=self.select_output_folder).pack(pady=5)
        CTkButton(self.root, text="Convert", command=self.convert_pdf).pack(pady=5)
        CTkButton(self.root, text="Exit", command=self.root.quit).pack(pady=5)

    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if file_path:
            self.pdf_file_path = file_path
            self.status_label.configure(text=f"Selected PDF: {os.path.basename(file_path)}")

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder_path = folder_path
            self.status_label.configure(text=f"Output Folder: {folder_path}")

    def convert_pdf(self):
        if not self.pdf_file_path:
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        if not self.output_folder_path:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        self.progress_bar.set(0)
        self.status_label.configure(text="Conversion in progress...")
        self.root.update()

        try:
            # Start Conversion
            output_file = os.path.join(self.output_folder_path, "converted.docx")
            self.pdf_to_docx(self.pdf_file_path, output_file)
            self.progress_bar.set(1)
            messagebox.showinfo("Success", f"File converted successfully! Saved at {output_file}")
            self.status_label.configure(text="Conversion completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.configure(text="Conversion failed.")
            self.progress_bar.set(0)

    def pdf_to_docx(self, pdf_path, output_path):
        """Converts a PDF to a DOCX file."""
        reader = PdfReader(pdf_path)
        document = Document()

        for page in reader.pages:
            text = page.extract_text()
            if text:
                document.add_paragraph(text)

        document.save(output_path)


if __name__ == "__main__":
    app = CTk()
    PDFtoDOCXConverter(app)
    app.mainloop()

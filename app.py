import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from customtkinter import CTk, CTkButton, CTkLabel, CTkProgressBar, CTkEntry
from PyPDF2 import PdfReader
from docx import Document


class PDFtoDOCXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to DOCX Converter")
        self.root.geometry("600x400")
        self.pdf_files = []
        self.output_folder_path = None

        # GUI Components
        self.create_widgets()

    def create_widgets(self):
        # Labels
        CTkLabel(self.root, text="PDF to DOCX Converter", font=("Arial", 18)).pack(pady=10)
        self.status_label = CTkLabel(self.root, text="Drag & drop files or select PDF files to start.", font=("Arial", 12))
        self.status_label.pack(pady=5)

        # File List
        self.file_listbox = tk.Listbox(self.root, height=8, width=50, selectmode=tk.EXTENDED)
        self.file_listbox.pack(pady=10)

        # Progress Bar
        self.progress_bar = CTkProgressBar(self.root, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)

        # Customization Options
        self.font_size_entry = CTkEntry(self.root, placeholder_text="Font Size (Default: 12)", width=200)
        self.font_size_entry.pack(pady=5)

        # Buttons
        CTkButton(self.root, text="Select PDFs", command=self.select_pdfs).pack(pady=5)
        CTkButton(self.root, text="Select Output Folder", command=self.select_output_folder).pack(pady=5)
        CTkButton(self.root, text="Convert", command=self.convert_pdfs).pack(pady=5)
        CTkButton(self.root, text="Exit", command=self.root.quit).pack(pady=5)

        # Drag-and-Drop Support
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.drag_and_drop_files)

    def select_pdfs(self):
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if file_paths:
            self.pdf_files.extend(file_paths)
            self.update_file_list()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder_path = folder_path
            self.status_label.configure(text=f"Output Folder: {folder_path}")

    def update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.pdf_files:
            self.file_listbox.insert(tk.END, file)

    def drag_and_drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file in files:
            if file.endswith('.pdf'):
                self.pdf_files.append(file)
        self.update_file_list()

    def convert_pdfs(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "Please select PDF files.")
            return
        if not self.output_folder_path:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        try:
            font_size = int(self.font_size_entry.get()) if self.font_size_entry.get() else 12
        except ValueError:
            messagebox.showerror("Error", "Font size must be an integer.")
            return

        self.progress_bar.set(0)
        self.status_label.configure(text="Conversion in progress...")
        self.root.update()

        try:
            total_files = len(self.pdf_files)
            for index, pdf_file in enumerate(self.pdf_files, start=1):
                output_file = os.path.join(
                    self.output_folder_path,
                    f"{os.path.splitext(os.path.basename(pdf_file))[0]}.docx"
                )
                self.pdf_to_docx(pdf_file, output_file, font_size)
                self.progress_bar.set(index / total_files)
                self.root.update()

            messagebox.showinfo("Success", f"All files converted successfully!")
            self.status_label.configure(text="Conversion completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.status_label.configure(text="Conversion failed.")
            self.progress_bar.set(0)

    def pdf_to_docx(self, pdf_path, output_path, font_size):
        """Converts a PDF to a DOCX file."""
        reader = PdfReader(pdf_path)
        document = Document()

        for page in reader.pages:
            text = page.extract_text()
            if text:
                paragraph = document.add_paragraph(text)
                run = paragraph.runs[0]
                run.font.size = font_size

        document.save(output_path)


if __name__ == "__main__":
    app = TkinterDnD.Tk()
    PDFtoDOCXConverter(app)
    app.mainloop()

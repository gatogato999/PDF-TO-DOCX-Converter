# PDF to DOCX Converter
Version 2.0.

## Description
This is an enhanced Python-based GUI application for converting PDF files to DOCX format. The new version includes advanced features like batch conversion, drag-and-drop functionality, and customizable output options.

## Features
1. Select a PDF File: Choose a single PDF file for conversion.
2. Save Output Location: Specify where the converted DOCX file will be saved.
3. Progress Bar: Displays the progress of the conversion.
4. Success/Failure Messages: Notifies the user when the process is complete.
5. Exit Button: Closes the application.

## New Features
1. Batch Conversion: Convert multiple PDFs at once.
2. Drag-and-Drop: Drag PDF files directly into the application for quick selection.
3. Customization Options: Set the font size for the output DOCX file.
4. Improved UI: Displays a list of selected files for transparency.
5. Custom Output Location: Save DOCX files to a user-specified folder.
6. Custom Icons: Updated branding with a unique icon for the .exe file.
   
## Dependencies
The application uses the following Python libraries:
tkinter (for GUI development)
PyPDF2 (for PDF extraction)
python-docx (for DOCX creation)
customtkinter (for an enhanced GUI experience)
tkinterdnd2 (for drag-and-drop functionality)
Install dependencies with:
`pip install  customtkinter tkinterdnd2 PyPDF2 python-docx`

## How to Use
1. Run the application using: `python app.py`
2. Drag PDF files into the window or use the Select PDFs button.
3. Choose an output folder by clicking Select Output Folder.
4. Optionally, specify a font size for the output DOCX file.
5. Click Convert to begin the batch conversion process.
6. Monitor the progress bar for conversion status.

## Optional Features
-Customize font size.
-Use drag-and-drop to add files quickly.
-Batch process multiple PDFs in a single session.

## Creating an Executable
- To create a standalone .exe file, use the following command:
- pyinstaller --onefile --noconsole --icon=app_icon.ico app.py
   
## Known Limitations
- Complex PDF layouts (e.g., tables, images) may not convert perfectly.
- Processing large files may take time, depending on system resources.

## Planned Improvements
- Add support for preserving advanced PDF formatting (e.g., tables, images).
- Optimize conversion speed for large documents.

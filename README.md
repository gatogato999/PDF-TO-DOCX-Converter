# PDF to DOCX Converter
Version 1.0 (Initial Version).

## Description
This is a Python-based graphical user interface (GUI) application to convert PDF files to DOCX format. 
The application is simple and user-friendly, designed for individual PDF file conversion.

## Features
1. Select a PDF File: Choose a single PDF file for conversion.
2. Save Output Location: Specify where the converted DOCX file will be saved.
3. Progress Bar: Displays the progress of the conversion.
4. Success/Failure Messages: Notifies the user when the process is complete.
5. Exit Button: Closes the application.
   
## Dependencies
The application uses the following Python libraries:
tkinter (for GUI development)
PyPDF2 (for PDF extraction)
python-docx (for DOCX creation)
Install dependencies with:
`pip install PyPDF2 python-docx`

## How to Use
1. Run the application using:
`python app.py`
2. Use the Select PDF button to choose a file.
3. Specify the output location.
4. Click Convert to generate the DOCX file.
5. View the progress on the progress bar.
   
## Known Limitations
- No support for batch file conversion.
- No drag-and-drop feature.
- Formatting may not be preserved perfectly.

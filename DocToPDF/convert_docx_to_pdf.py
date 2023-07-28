'''
first install following

pip install python-docx
pip install fpdf
pip install pywin32

'''

import os
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog

def replace_slashes(input_string):
    # Use the replace method to replace all forward slashes with backslashes
    output_string = input_string.replace("/", "\\")

    return output_string

def replace_extension(input_file, new_extension):
    # Get the base file name without extension
    base_name = os.path.splitext(input_file)[0]

    # Combine the new base name with the new extension
    new_file = f"{base_name}{new_extension}"

    return new_file

def get_word_file_path():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Microsoft Word Documents", "*.doc;*.docx")])

    return file_path


def doc_to_pdf(input_file, output_file):
    # Check if the input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return

    word_app = win32.Dispatch("Word.Application")
    doc = word_app.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=17)  # 17 is the PDF format
    doc.Close()
    word_app.Quit()

if __name__ == "__main__":
    # Provide the file paths with backslashes (\)
    word_file_path = get_word_file_path()

    if word_file_path:
        print("Selected Microsoft Word document path:", word_file_path)
    else:
        print("No file selected.")
    
    modified_path = replace_slashes(word_file_path)

    doc_file =   modified_path   # r"C:\Users\amank\OneDrive - IIT Kanpur\SimpleO\docxpdf\MENIFESTO.docx"

    modified_file_path = replace_extension(modified_path, ".pdf")
    pdf_file = modified_file_path  # r"C:\Users\amank\OneDrive - IIT Kanpur\SimpleO\docxpdf\MENIFESTO.pdf"

    # Call the conversion function
    doc_to_pdf(doc_file, pdf_file)

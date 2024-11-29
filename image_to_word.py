import tkinter as tk
from tkinter import filedialog
from PIL import Image
import pytesseract
from docx import Document
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def select_image():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")])
    if file_path:
        process_image(file_path)

def process_image(image_path):
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image)
    
    doc = Document()
    doc.add_paragraph(text)
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}_output.docx")
    doc.save(output_path)
    print(f"Word document created: {output_path}")

if __name__ == "__main__":
    select_image()

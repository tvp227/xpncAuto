import glob
import openpyxl
import getpass
import subprocess
import os
from fpdf import FPDF
from PIL import Image
from tkinter import Tk, filedialog
from datetime import datetime

username = getpass.getuser()

#PDF convertion.

class PDFWithImageNames(FPDF):
    def header(self):
        pass
    
    def footer(self):
        pass
    
    def chapter_title(self, title):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, title, ln=True, align="C")

def convert_images_to_pdf(directory):
    # Create a PDF object
    pdf = PDFWithImageNames()

    # Get all image files in the directory
    image_files = glob.glob(directory + '/*.jpg') + glob.glob(directory + '/*.jpeg') + glob.glob(directory + '/*.png')

    if not image_files:
        print("No image files found in the directory.")
        return

    # Sort the image files alphabetically
    image_files.sort()

    # Add each image to the PDF
    for image_file in image_files:
        # Add a new page
        pdf.add_page()

        # Get the image name
        image_name = os.path.basename(image_file)

        # Open the image file
        img = Image.open(image_file)

        # Calculate the maximum width and height of the image within the page
        max_width = 210 - 2 * pdf.l_margin
        max_height = 297 - 2 * pdf.t_margin

        # Calculate the aspect ratio of the image
        aspect_ratio = img.width / img.height

        # Calculate the width and height of the image to fit within the page
        if aspect_ratio >= 1:
            width = max_width
            height = max_width / aspect_ratio
        else:
            width = max_height * aspect_ratio
            height = max_height

        # Calculate the position for the image and image name
        x = (210 - width) / 2
        y = 297 - pdf.t_margin - height

        # Add the image to the PDF
        pdf.image(image_file, x=x, y=y, w=width, h=height)

        # Add the image name as a label
        pdf.chapter_title(image_name)

    # Save the PDF file with the current date and time
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    username = getpass.getuser()  # Get the current username
    pdf_file = f"C:/Users/{username}/Documents/XPNCv2/expense_export_{timestamp}.pdf"
    pdf.output(pdf_file)
    print(f"PDF file '{pdf_file}' created successfully.")

# Create a Tkinter root window
root = Tk()
root.withdraw()  # Hide the root window

# Prompt the user to select a directory
directory = filedialog.askdirectory(title="Select a directory")

if not directory:
    print("No directory selected.")
else:
    # Convert images to PDF
    convert_images_to_pdf(directory)
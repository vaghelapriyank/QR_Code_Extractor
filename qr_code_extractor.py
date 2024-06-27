import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
from pyzbar.pyzbar import decode
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
import os
from PIL import Image as PILImage
from io import BytesIO
import matplotlib.pyplot as plt

def extract_images_from_pdf(pdf_path, zoom=2.0):
    """Extract and zoom images from a given PDF file."""
    images = []
    doc = fitz.open(pdf_path)
    for i in range(len(doc)):
        page = doc[i]
        
        # Set the zoom factor for high resolution
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)
        
        # Convert the image to a numpy array
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        # Handle the case where the image has an alpha channel (RGBA)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        
        images.append(img)
        
        # Display the image for verification
        # plt.imshow(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
        # plt.show()
    
    return images

def enhance_image(image):
    """Apply image processing to enhance the visibility of QR codes."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    sharpen_kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
    sharpened = cv2.filter2D(gray, -1, sharpen_kernel)
    enhanced = cv2.adaptiveThreshold(sharpened, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    return enhanced

def detect_qr_code(image):
    """Detect and decode QR code from an image."""
    enhanced_image = enhance_image(image)
    # plt.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    # #plt.title(f'Image from {filename}')
    # plt.show()
    qr_codes = decode(enhanced_image)
    qr_data = []
    for qr_code in qr_codes:
        qr_info = qr_code.data.decode("utf-8")
        qr_data.append({
            'rect': qr_code.rect,
            'data': qr_info
        })
    # print (qr_data)
    return qr_data

def extract_substring(string):
  """Extracts the substring between the second-last "/" and the end of the string.

  Args:
      string: The input string.

  Returns:
      The extracted substring, or an empty string if no such substring exists.
  """

  # Split the string by "/"
  parts = string.split("/")

  # Check if there are at least 2 parts (excluding the empty string at the beginning if there's a leading "/")
  if len(parts) >= 2:
    # Return the element before the last element (which is the second-last)
    return parts[-2]
  else:
    # No substring found, return empty string
    return ""

def process_pdfs_in_folder(folder_path, zoom=1.0):
    """Process all PDFs in the specified folder and store QR code data in a PDF report."""
    qr_code_data = []
    Excel_qr_code_data=[]
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            images = extract_images_from_pdf(pdf_path, zoom=zoom)

            for img in images:
                # plt.imshow(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
                # plt.title(f'Image from {filename}')
                # plt.show()
                
                qr_codes = detect_qr_code(img)
                for qr in qr_codes:
                    # Print QR code data to the console
                    print(f"QR Code found in {filename}: {qr['data']}")
                    
                    # Crop the QR code image
                    rect = qr['rect']
                    qr_img = img[rect.top:rect.top + rect.height, rect.left:rect.left + rect.width]
                    qr_img_pil = PILImage.fromarray(cv2.cvtColor(qr_img, cv2.COLOR_BGR2RGB))
                    img_byte_arr = BytesIO()
                    qr_img_pil.save(img_byte_arr, format='PNG')
                    qr_img_byte = img_byte_arr.getvalue()
                    extracted_substring = extract_substring(qr['data'])
                    qr_code_data.append({
                        'PDF File': filename,
                        'Tag' : extracted_substring,
                        'QR Code': qr['data'],
                        'QR Image': qr_img_byte
                    })
                    Excel_qr_code_data.append({
                        'PDF File': filename,
                        'Tag' : extracted_substring,
                        'QR Code': qr['data']
                    })

    df = pd.DataFrame(Excel_qr_code_data)
    df.to_excel('qr_code_data.xlsx', index=False)
    print("qr_code_data.xlsx Excel Report is Genrated")
    generate_pdf_report(qr_code_data)
    print("qr_code_report.pdf PDF Report is Genrated")

def generate_pdf_report(qr_code_data):
    """Generate a PDF report from the QR code data."""
    report_filename = 'qr_code_report.pdf'
    doc = SimpleDocTemplate(report_filename, pagesize=letter)
    elements = []

    # Create table data
    table_data = [['PDF File', 'Tag' ,'QR Code', 'QR Image']]

    for item in qr_code_data:
        img = PILImage.open(BytesIO(item['QR Image']))
        img_path = f"temp_{item['PDF File']}.png"
        img.save(img_path)
        table_data.append([item['PDF File'], item['Tag'] ,item['QR Code'], Image(img_path, 1 * inch, 1 * inch)])

    # Create table
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    doc.build(elements)

    # Clean up temporary images
    for item in qr_code_data:
        os.remove(f"temp_{item['PDF File']}.png")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python qr_code_extractor.py <path_to_pdf_folder>")
    else:
        folder_path = sys.argv[1]
        process_pdfs_in_folder(folder_path, zoom=1.7)

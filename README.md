# QR Code Extractor from PDFs

## Overview
This project provides a Python script that extracts images from PDF files, enhances them to detect QR codes, and generates a report containing the extracted QR code data. The report is saved as both an Excel file and a PDF file, which includes cropped QR code images.

## Features
- Extracts images from PDF files with optional zoom.
- Enhances images to improve QR code detection.
- Detects and decodes QR codes from images.
- Generates an Excel report with QR code data.
- Generates a PDF report with QR code data and images.

## Prerequisites
Make sure you have the following libraries installed:
- fitz (PyMuPDF)
- cv2 (OpenCV)
- numpy
- pandas
- pyzbar
- reportlab
- PIL (Pillow)
- matplotlib

You can install these dependencies using pip:
bash
pip install pymupdf opencv-python numpy pandas pyzbar reportlab Pillow matplotlib


## Usage
Run the script from the command line by providing the path to the folder containing the PDF files:
bash
python qr_code_extractor.py <path_to_pdf_folder>

For example:
bash
python qr_code_extractor.py /path/to/pdf/folder


## Functions

### extract_images_from_pdf(pdf_path, zoom=2.0)
Extracts and zooms images from a given PDF file.

*Parameters:*
- pdf_path (str): Path to the PDF file.
- zoom (float): Zoom factor for extracting images (default is 2.0).

*Returns:*
- images (list): List of extracted images as numpy arrays.

### enhance_image(image)
Applies image processing to enhance the visibility of QR codes.

*Parameters:*
- image (numpy array): Input image.

*Returns:*
- enhanced (numpy array): Enhanced image.

### detect_qr_code(image)
Detects and decodes QR codes from an image.

*Parameters:*
- image (numpy array): Input image.

*Returns:*
- qr_data (list): List of dictionaries containing QR code data and positions.

### extract_substring(string)
Extracts the substring between the second-last "/" and the end of the string.

*Parameters:*
- string (str): The input string.

*Returns:*
- extracted_substring (str): The extracted substring, or an empty string if no such substring exists.

### process_pdfs_in_folder(folder_path, zoom=1.0)
Processes all PDFs in the specified folder and stores QR code data in an Excel and PDF report.

*Parameters:*
- folder_path (str): Path to the folder containing PDF files.
- zoom (float): Zoom factor for extracting images (default is 1.0).

### generate_pdf_report(qr_code_data)
Generates a PDF report from the QR code data.

*Parameters:*
- qr_code_data (list): List of dictionaries containing QR code data and images.

## Example
Here's an example of how to use the main functions within a script:
python
if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python qr_code_extractor.py <path_to_pdf_folder>")
    else:
        folder_path = sys.argv[1]
        process_pdfs_in_folder(folder_path, zoom=1.7)


In this example, the script expects the folder path containing PDF files as a command-line argument. It processes all PDF files in the folder, extracts and decodes QR codes, and generates Excel and PDF reports.

## Output
- qr_code_data.xlsx: An Excel file containing the QR code data.
- qr_code_report.pdf: A PDF file containing the QR code data and images.

Both files will be generated in the directory where the script is executed.

## License
This project is licensed under the MIT License.

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## Author
Priyank Vaghela
Email: vaghelapriyank142@gmail.com
Contact: +91 9925339131
LinkedIn: https://www.linkedin.com/in/priyank-vaghela-758726173/

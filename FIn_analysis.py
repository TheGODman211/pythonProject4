import pytesseract
from pdf2image import convert_from_path
import cv2
import numpy as np

# Path to the input scanned PDF file
pdf_path = 'path/to/your/scanned.pdf'

# Convert each page of the PDF to images
images = convert_from_path(pdf_path)

# Loop through each image
for i, image in enumerate(images):
    # Convert PIL image to OpenCV format
    img_cv = np.array(image)

    # Preprocess the image for better OCR results (optional)
    img_cv = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    img_cv = cv2.threshold(img_cv, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    img_cv = cv2.medianBlur(img_cv, 3)

    # Perform OCR using Pytesseract
    extracted_text = pytesseract.image_to_string(img_cv)

    # Process the extracted text to identify the table structure
    lines = extracted_text.split('\n')
    table_data = []
    for line in lines:
        cells = line.split('\t')  # Assuming tab-delimited cells
        table_data.append(cells)

    # Print the extracted table data
    print(f"Table data for page {i + 1}:")
    for row in table_data:
        print(row)
    print()

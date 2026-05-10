import cv2
import numpy as np
import pytesseract
import fitz  # PyMuPDF
from PIL import Image
import io

# Point pytesseract to the installed Windows binary
# By default, the silent installer puts it here
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image_for_ocr(pil_img):
    """Applies grayscale, thresholding, and noise removal for better OCR accuracy."""
    # Convert PIL image to OpenCV format (numpy array)
    open_cv_image = np.array(pil_img)
    
    # Convert RGB to BGR
    img = open_cv_image[:, :, ::-1].copy() 

    # 1. Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 2. Rescale image (300 DPI equivalent) for better OCR reading
    gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    # 3. Apply Gaussian Blur to remove high frequency noise
    blur = cv2.GaussianBlur(gray, (5,5), 0)

    # 4. Otsu's thresholding (Binarization)
    # Converts everything strictly to black (text) and white (background)
    _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    return thresh

def extract_text_with_tesseract(pdf_path, page_num):
    """Extracts text from a specific PDF page using local Tesseract OCR."""
    try:
        # Extract page as image using fitz
        doc = fitz.open(pdf_path)
        page = doc[page_num - 1]
        pix = page.get_pixmap(dpi=300) # High DPI for OCR
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()

        # Preprocess the image
        processed_img = preprocess_image_for_ocr(img)

        # Run Tesseract OCR
        # --psm 3: Fully automatic page segmentation, but no OSD. (Default)
        # --psm 6: Assume a single uniform block of text.
        custom_config = r'--oem 3 --psm 3'
        text = pytesseract.image_to_string(processed_img, config=custom_config, lang='eng')
        
        return text.strip()
    except Exception as e:
        print(f"Tesseract Error: {e}")
        return None

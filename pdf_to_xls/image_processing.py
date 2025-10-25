"""Image processing and rotation detection utilities."""

import base64
from io import BytesIO
import fitz  # PyMuPDF
from PIL import Image
import pytesseract


def detect_orientation(img):
    """Detect image orientation using pytesseract OSD.

    Args:
        img: PIL Image object

    Returns:
        tuple: (rotation_angle, confidence)
            - rotation_angle: Degrees to rotate clockwise to correct orientation
            - confidence: Confidence level of the detection
    """
    try:
        # Use pytesseract to detect orientation
        osd = pytesseract.image_to_osd(img)

        # Parse the OSD output to get rotation angle
        rotation = 0
        confidence = 0
        for line in osd.split('\n'):
            if 'Rotate:' in line:
                rotation = int(line.split(':')[1].strip())
            if 'Orientation confidence:' in line:
                confidence = float(line.split(':')[1].strip())

        return rotation, confidence
    except Exception:
        # If OSD fails, return 0 (no rotation)
        return 0, 0


def convert_pdf_page_to_image(pdf_path, page_num):
    """Convert a PDF page to base64-encoded image using PyMuPDF.

    Automatically detects and corrects page orientation.

    Args:
        pdf_path: Path to the PDF file
        page_num: Page number (1-indexed)

    Returns:
        str: Base64-encoded PNG image data, or None if conversion fails
    """
    try:
        # Open PDF with PyMuPDF
        doc = fitz.open(pdf_path)

        # Get the page (0-indexed)
        page = doc[page_num - 1]

        # Render page to image at high resolution
        mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
        pix = page.get_pixmap(matrix=mat)

        # Convert to PIL Image
        img_bytes = pix.tobytes("png")
        img = Image.open(BytesIO(img_bytes))

        # Close document
        doc.close()

        # Detect actual visual orientation using OCR
        detected_rotation, confidence = detect_orientation(img)

        # Apply rotation correction if needed (only if confidence is reasonable)
        if detected_rotation != 0 and confidence > 1.0:
            # Rotation direction conversion:
            # - Tesseract OSD "Rotate" value = degrees to rotate CLOCKWISE to correct orientation
            # - PIL's rotate() method rotates COUNTER-CLOCKWISE by default
            # - Therefore: PIL_angle = -Tesseract_angle to convert conventions
            # - expand=True ensures the canvas expands to fit the rotated image without cropping
            # Example: If text is 90° clockwise (sideways right), Tesseract returns 270,
            #          and rotate(-270) = rotate 90° clockwise = corrects the orientation
            img = img.rotate(-detected_rotation, expand=True)
            print(f"    Detected rotation {detected_rotation}° (confidence: {confidence:.1f}) - correcting")

        # Convert PIL Image to PNG bytes
        output = BytesIO()
        img.save(output, format='PNG')
        final_img_data = output.getvalue()

        # Encode to base64
        return base64.standard_b64encode(final_img_data).decode('utf-8')

    except Exception as e:
        print(f"    Error converting page to image: {e}")
        import traceback
        traceback.print_exc()
        return None

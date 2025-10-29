"""Image processing and rotation detection utilities."""

import base64
from io import BytesIO
from pathlib import Path
import fitz  # PyMuPDF
from PIL import Image
import pytesseract


# Claude API has a 5 MB limit for base64-encoded images
MAX_IMAGE_SIZE_BYTES = 5 * 1024 * 1024  # 5 MB


def resize_image_for_api(img, max_size_bytes=MAX_IMAGE_SIZE_BYTES):
    """Resize image to fit within API size limit.

    Iteratively reduces image dimensions until the base64-encoded PNG
    is under the specified size limit.

    Args:
        img: PIL Image object
        max_size_bytes: Maximum size in bytes for base64-encoded image (default: 5 MB)

    Returns:
        PIL Image object: Resized image that fits within size limit
    """
    # Start with original image
    current_img = img.copy()
    scale_factor = 1.0

    # Get initial size
    output = BytesIO()
    current_img.save(output, format='PNG')
    current_size = len(base64.standard_b64encode(output.getvalue()))

    # If already under limit, return original
    if current_size <= max_size_bytes:
        return current_img

    print(f"    Image size {current_size:,} bytes exceeds {max_size_bytes:,} byte limit, resizing...")

    # Iteratively reduce size until under limit
    # Start with aggressive reduction based on size ratio
    size_ratio = current_size / max_size_bytes
    # Since image size scales roughly with pixel count (width * height),
    # we need to reduce linear dimensions by sqrt(size_ratio)
    scale_factor = 1.0 / (size_ratio ** 0.5)

    # Add some buffer to ensure we get under the limit
    scale_factor *= 0.9

    attempts = 0
    max_attempts = 10

    while current_size > max_size_bytes and attempts < max_attempts:
        attempts += 1

        # Calculate new dimensions
        new_width = int(img.width * scale_factor)
        new_height = int(img.height * scale_factor)

        # Ensure minimum dimensions
        if new_width < 100 or new_height < 100:
            print(f"    Warning: Image dimensions too small ({new_width}x{new_height}), using minimum size")
            new_width = max(new_width, 100)
            new_height = max(new_height, 100)

        # Resize image with high-quality resampling
        current_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # Check new size
        output = BytesIO()
        current_img.save(output, format='PNG')
        current_size = len(base64.standard_b64encode(output.getvalue()))

        if current_size > max_size_bytes:
            # Reduce scale factor further for next attempt
            scale_factor *= 0.85

    print(f"    Resized to {current_img.width}x{current_img.height} ({current_size:,} bytes)")

    return current_img


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


def convert_image_file_to_base64(image_path):
    """Convert an image file to base64-encoded PNG.

    Automatically detects and corrects image orientation.

    Args:
        image_path: Path to the image file (.jpg, .jpeg, .png, .tiff, .tif)

    Returns:
        str: Base64-encoded PNG image data, or None if conversion fails
    """
    try:
        image_path = Path(image_path)

        # Open image file
        img = Image.open(image_path)

        # Convert to RGB if necessary (for TIFF, etc.)
        if img.mode not in ('RGB', 'L', 'RGBA'):
            img = img.convert('RGB')

        # Detect actual visual orientation using OCR
        detected_rotation, confidence = detect_orientation(img)

        # Apply rotation correction if needed (only if confidence is reasonable)
        if detected_rotation != 0 and confidence > 1.0:
            # Rotation direction conversion:
            # - Tesseract OSD "Rotate" value = degrees to rotate CLOCKWISE to correct orientation
            # - PIL's rotate() method rotates COUNTER-CLOCKWISE by default
            # - Therefore: PIL_angle = -Tesseract_angle to convert conventions
            # - expand=True ensures the canvas expands to fit the rotated image without cropping
            img = img.rotate(-detected_rotation, expand=True)
            print(f"    Detected rotation {detected_rotation}° (confidence: {confidence:.1f}) - correcting")

        # Resize image if needed to fit within API size limit
        img = resize_image_for_api(img)

        # Convert PIL Image to PNG bytes
        output = BytesIO()
        img.save(output, format='PNG')
        final_img_data = output.getvalue()

        # Encode to base64
        return base64.standard_b64encode(final_img_data).decode('utf-8')

    except Exception as e:
        print(f"    Error converting image to base64: {e}")
        import traceback
        traceback.print_exc()
        return None


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
        # Use 4x zoom for better OCR accuracy, especially for tables with many columns
        # Higher resolution is critical for wide tables with small text
        mat = fitz.Matrix(4.0, 4.0)  # 4x zoom for better quality
        pix = page.get_pixmap(matrix=mat)

        # Convert to PIL Image
        img_bytes = pix.tobytes("png")
        img = Image.open(BytesIO(img_bytes))

        print(f"    Initial image size: {img.width}x{img.height} pixels")

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

        # Resize image if needed to fit within API size limit
        img = resize_image_for_api(img)
        print(f"    Final image size: {img.width}x{img.height} pixels")

        # Convert PIL Image to PNG bytes
        output = BytesIO()
        img.save(output, format='PNG')
        final_img_data = output.getvalue()
        final_size_mb = len(base64.standard_b64encode(final_img_data)) / (1024 * 1024)
        print(f"    Final encoded size: {final_size_mb:.2f} MB")

        # Encode to base64
        return base64.standard_b64encode(final_img_data).decode('utf-8')

    except Exception as e:
        print(f"    Error converting page to image: {e}")
        import traceback
        traceback.print_exc()
        return None

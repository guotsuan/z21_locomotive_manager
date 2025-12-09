#!/usr/bin/env python3
"""
Test script for PaddleOCR functionality.

This script tests PaddleOCR installation and OCR capabilities on sample images.
It can be used to verify if PaddleOCR works correctly before integrating it into the main application.
"""

import sys
import os
from pathlib import Path
import argparse

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Try to import PaddleOCR
HAS_PADDLEOCR = False
PADDLEOCR_IMPORT_ERROR = None

try:
    # Fix OpenMP library conflict on macOS (must be set before importing PaddleOCR)
    if sys.platform == 'darwin':
        os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
    
    # Allow disabling PaddleOCR via environment variable
    DISABLE_PADDLEOCR = os.environ.get("DISABLE_PADDLEOCR", "").lower() in ("1", "true", "yes")
    
    if DISABLE_PADDLEOCR:
        print("PaddleOCR is disabled via DISABLE_PADDLEOCR environment variable")
        HAS_PADDLEOCR = False
    else:
        from paddleocr import PaddleOCR
        HAS_PADDLEOCR = True
        print("✓ PaddleOCR imported successfully")
except ImportError as e:
    PADDLEOCR_IMPORT_ERROR = str(e)
    print(f"✗ PaddleOCR not installed: {e}")
    print("\nTo install PaddleOCR, run:")
    print("  pip install paddleocr")
except Exception as e:
    PADDLEOCR_IMPORT_ERROR = str(e)
    print(f"✗ Failed to import PaddleOCR: {e}")


def test_paddleocr_initialization():
    """Test PaddleOCR initialization with different configurations."""
    if not HAS_PADDLEOCR:
        print("\n⚠️  Skipping initialization test - PaddleOCR not available")
        return None
    
    print("\n" + "=" * 60)
    print("Testing PaddleOCR Initialization")
    print("=" * 60)
    
    configs = [
        {"name": "Minimal (lang='en')", "params": {"lang": "en"}},
        {"name": "With use_angle_cls", "params": {"use_angle_cls": True, "lang": "en"}},
        {"name": "With use_textline_orientation", "params": {"use_textline_orientation": True, "lang": "en"}},
    ]
    
    ocr = None
    working_config = None
    
    for config in configs:
        print(f"\nTrying: {config['name']}")
        print(f"  Parameters: {config['params']}")
        try:
            ocr = PaddleOCR(**config['params'])
            print(f"  ✓ Success!")
            working_config = config
            break
        except Exception as e:
            print(f"  ✗ Failed: {e}")
            continue
    
    if ocr is None:
        print("\n✗ All initialization attempts failed")
        return None
    
    print(f"\n✓ Using configuration: {working_config['name']}")
    return ocr


def test_ocr_on_image(ocr, image_path: Path):
    """Test OCR on a single image file."""
    if not HAS_PADDLEOCR:
        print("\n⚠️  Skipping OCR test - PaddleOCR not available")
        return None
    
    if not image_path.exists():
        print(f"\n✗ Image file not found: {image_path}")
        return None
    
    print("\n" + "=" * 60)
    print(f"Testing OCR on: {image_path.name}")
    print("=" * 60)
    
    try:
        print(f"\nReading image: {image_path}")
        
        # Test different OCR methods
        print("\n--- Method 1: ocr() ---")
        try:
            result = ocr.ocr(str(image_path), cls=False)
            print("✓ ocr() method succeeded")
            text_lines = []
            if result and result[0]:
                for line in result[0]:
                    if line and len(line) >= 2:
                        text = line[1][0] if isinstance(line[1], (list, tuple)) else str(line[1])
                        confidence = line[1][1] if isinstance(line[1], (list, tuple)) and len(line[1]) > 1 else "N/A"
                        text_lines.append((text, confidence))
                        print(f"  Text: {text} (confidence: {confidence})")
            
            if text_lines:
                full_text = "\n".join([line[0] for line in text_lines])
                print(f"\nExtracted text:\n{full_text}")
                return full_text
            else:
                print("  No text detected")
                return None
        except Exception as e:
            print(f"✗ ocr() method failed: {e}")
            return None
        
    except Exception as e:
        print(f"\n✗ OCR failed with error: {e}")
        import traceback
        traceback.print_exc()
        return None


def test_pdf_ocr(ocr, pdf_path: Path):
    """Test OCR on PDF file."""
    if not HAS_PADDLEOCR:
        print("\n⚠️  Skipping PDF OCR test - PaddleOCR not available")
        return None
    
    if not pdf_path.exists():
        print(f"\n✗ PDF file not found: {pdf_path}")
        return None
    
    print("\n" + "=" * 60)
    print(f"Testing OCR on PDF: {pdf_path.name}")
    print("=" * 60)
    
    try:
        # Try to import pdf2image
        try:
            from pdf2image import convert_from_path
            HAS_PDF2IMAGE = True
        except ImportError:
            print("✗ pdf2image not installed. Install with: pip install pdf2image")
            return None
        
        print(f"\nConverting PDF to images...")
        images = convert_from_path(str(pdf_path))
        print(f"✓ Converted to {len(images)} page(s)")
        
        all_text = []
        for i, image in enumerate(images):
            print(f"\n--- Processing page {i + 1}/{len(images)} ---")
            try:
                # Save to temporary file for PaddleOCR
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    image.save(tmp.name, 'PNG')
                    tmp_path = Path(tmp.name)
                
                try:
                    result = ocr.ocr(str(tmp_path), cls=False)
                    page_text = []
                    if result and result[0]:
                        for line in result[0]:
                            if line and len(line) >= 2:
                                text = line[1][0] if isinstance(line[1], (list, tuple)) else str(line[1])
                                page_text.append(text)
                    
                    if page_text:
                        page_full_text = "\n".join(page_text)
                        print(f"✓ Page {i + 1} text extracted ({len(page_text)} lines)")
                        all_text.append(page_full_text)
                    else:
                        print(f"  No text detected on page {i + 1}")
                finally:
                    tmp_path.unlink()
                    
            except Exception as e:
                print(f"✗ Failed to process page {i + 1}: {e}")
                continue
        
        if all_text:
            full_text = "\n\n--- Page Separator ---\n\n".join(all_text)
            print(f"\n✓ Extracted text from {len(all_text)} page(s)")
            return full_text
        else:
            print("\n✗ No text extracted from PDF")
            return None
            
    except Exception as e:
        print(f"\n✗ PDF OCR failed: {e}")
        import traceback
        traceback.print_exc()
        return None


def compare_with_pytesseract(image_path: Path):
    """Compare PaddleOCR results with pytesseract (if available)."""
    print("\n" + "=" * 60)
    print("Comparison with pytesseract")
    print("=" * 60)
    
    # Test PaddleOCR
    paddleocr_text = None
    if HAS_PADDLEOCR:
        ocr = test_paddleocr_initialization()
        if ocr:
            paddleocr_text = test_ocr_on_image(ocr, image_path)
    
    # Test pytesseract
    pytesseract_text = None
    try:
        import pytesseract
        from PIL import Image
        
        print("\n--- Testing pytesseract ---")
        image = Image.open(image_path)
        pytesseract_text = pytesseract.image_to_string(image)
        if pytesseract_text.strip():
            print(f"✓ pytesseract extracted {len(pytesseract_text)} characters")
        else:
            print("  No text extracted by pytesseract")
    except ImportError:
        print("\n⚠️  pytesseract not installed - skipping comparison")
    except Exception as e:
        print(f"\n✗ pytesseract failed: {e}")
    
    # Compare results
    if paddleocr_text and pytesseract_text:
        print("\n--- Comparison Summary ---")
        print(f"PaddleOCR: {len(paddleocr_text)} characters")
        print(f"pytesseract: {len(pytesseract_text)} characters")
        
        paddleocr_words = set(paddleocr_text.lower().split())
        pytesseract_words = set(pytesseract_text.lower().split())
        common_words = paddleocr_words & pytesseract_words
        
        if common_words:
            print(f"\nCommon words: {len(common_words)}")
            print(f"PaddleOCR unique: {len(paddleocr_words - pytesseract_words)}")
            print(f"pytesseract unique: {len(pytesseract_words - paddleocr_words)}")


def main():
    parser = argparse.ArgumentParser(description="Test PaddleOCR functionality")
    parser.add_argument("file", type=Path, nargs="?", help="Image or PDF file to test OCR on")
    parser.add_argument("--compare", action="store_true", help="Compare with pytesseract")
    parser.add_argument("--pdf", action="store_true", help="Force PDF mode")
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("PaddleOCR Test Script")
    print("=" * 60)
    
    if not HAS_PADDLEOCR:
        print("\n⚠️  PaddleOCR is not available.")
        if PADDLEOCR_IMPORT_ERROR:
            print(f"   Error: {PADDLEOCR_IMPORT_ERROR}")
        print("\nTo install PaddleOCR:")
        print("  pip install paddleocr")
        print("\nNote: PaddleOCR requires additional dependencies and may take a while to install.")
        return 1
    
    # Test initialization
    ocr = test_paddleocr_initialization()
    if ocr is None:
        print("\n✗ Failed to initialize PaddleOCR. Please check your installation.")
        return 1
    
    # Test on file if provided
    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            print(f"\n✗ File not found: {file_path}")
            return 1
        
        if file_path.suffix.lower() == '.pdf' or args.pdf:
            result = test_pdf_ocr(ocr, file_path)
        else:
            result = test_ocr_on_image(ocr, file_path)
            if args.compare:
                compare_with_pytesseract(file_path)
        
        if result:
            print("\n" + "=" * 60)
            print("✓ OCR Test Completed Successfully")
            print("=" * 60)
        else:
            print("\n" + "=" * 60)
            print("⚠️  OCR completed but no text extracted")
            print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("✓ Initialization Test Completed")
        print("=" * 60)
        print("\nTo test OCR on an image, run:")
        print(f"  {sys.argv[0]} <image_path>")
        print("\nTo compare with pytesseract:")
        print(f"  {sys.argv[0]} <image_path> --compare")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())



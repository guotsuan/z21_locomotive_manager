#!/usr/bin/env python3
"""Test script to diagnose PaddleOCR issues."""
import os
import sys

# Set OpenMP fix before importing
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

print("=" * 60)
print("PaddleOCR Diagnostic Test")
print("=" * 60)
print(f"Python version: {sys.version}")
print(f"KMP_DUPLICATE_LIB_OK: {os.environ.get('KMP_DUPLICATE_LIB_OK', 'NOT SET')}")
print()

try:
    print("Step 1: Importing paddleocr...")
    from paddleocr import PaddleOCR
    print("✓ Import successful")
except Exception as e:
    print(f"✗ Import failed: {e}")
    sys.exit(1)

try:
    print("\nStep 2: Initializing PaddleOCR...")
    print("(This is where segmentation fault usually occurs)")
    ocr = PaddleOCR(lang='en')
    print("✓ Initialization successful")
except Exception as e:
    print(f"✗ Initialization failed: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "=" * 60)
print("All tests passed! PaddleOCR should work.")
print("=" * 60)

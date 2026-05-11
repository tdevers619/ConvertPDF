import os
import pdf2docx
from pdf2docx import Converter

print("=== Bulk PDF to DOCX Converter ===\n")

# Prompt user for folders
input_folder = input("Enter the INPUT folder path containing PDFs: ").strip('"')
output_folder = input("Enter the OUTPUT folder path for DOCX files: ").strip('"')

# Validate input folder
if not os.path.exists(input_folder):
    print(f"\nInput folder does not exist:\n{input_folder}")
    exit()

# Create output folder if needed
os.makedirs(output_folder, exist_ok=True)

# Find PDFs
pdf_files = [
    f for f in os.listdir(input_folder)
    if f.lower().endswith(".pdf")
]

if not pdf_files:
    print("\nNo PDF files found in the input folder.")
    exit()

print(f"\nFound {len(pdf_files)} PDF file(s).\n")

# Convert files
for pdf_file in pdf_files:
    try:
        pdf_path = os.path.join(input_folder, pdf_file)

        # Create output DOCX filename
        docx_name = os.path.splitext(pdf_file)[0] + ".docx"
        docx_path = os.path.join(output_folder, docx_name)

        # Skip existing files
        if os.path.exists(docx_path):
            print(f"Skipping (already exists): {docx_name}")
            continue

        print(f"Converting: {pdf_file}")

        # Convert PDF to DOCX
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()

        print(f"Saved: {docx_name}\n")

    except Exception as e:
        print(f"Error converting {pdf_file}: {e}\n")

print("=== Conversion Complete ===")
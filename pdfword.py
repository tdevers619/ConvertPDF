import os
from pdf2docx import Converter
from tqdm import tqdm

print("=== Bulk PDF to DOCX Converter ===\n")

input_folder = input("Enter the INPUT folder path containing PDFs: ").strip('"')
output_folder = input("Enter the OUTPUT folder path for DOCX files: ").strip('"')

if not os.path.exists(input_folder):
    print(f"\nInput folder does not exist:\n{input_folder}")
    exit()

os.makedirs(output_folder, exist_ok=True)

pdf_files = [
    f for f in os.listdir(input_folder)
    if f.lower().endswith(".pdf")
]

if not pdf_files:
    print("\nNo PDF files found in the input folder.")
    exit()

print(f"\nFound {len(pdf_files)} PDF file(s).\n")

for pdf_file in tqdm(pdf_files, desc="Converting PDFs", unit="file"):
    try:
        pdf_path = os.path.join(input_folder, pdf_file)
        docx_name = os.path.splitext(pdf_file)[0] + ".docx"
        docx_path = os.path.join(output_folder, docx_name)

        if os.path.exists(docx_path):
            tqdm.write(f"Skipping existing: {docx_name}")
            continue

        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()

        tqdm.write(f"Converted: {pdf_file}")

    except Exception as e:
        tqdm.write(f"Error converting {pdf_file}: {e}")

print("\n=== Conversion Complete ===")
import os
import win32com.client

input_folder = input("Enter INPUT folder: ").strip('"')
output_folder = input("Enter OUTPUT folder: ").strip('"')

os.makedirs(output_folder, exist_ok=True)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]

for pdf_file in pdf_files:
    pdf_path = os.path.join(input_folder, pdf_file)
    docx_path = os.path.join(
        output_folder,
        os.path.splitext(pdf_file)[0] + ".docx"
    )

    if os.path.exists(docx_path):
        print(f"Skipping existing: {docx_path}")
        continue

    try:
        print(f"Converting: {pdf_file}")

        doc = word.Documents.Open(pdf_path)
        doc.SaveAs2(docx_path, FileFormat=16)  # 16 = DOCX
        doc.Close(False)

        print(f"Saved: {docx_path}")

    except Exception as e:
        print(f"Failed: {pdf_file} - {e}")

word.Quit()
print("Done.")
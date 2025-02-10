import os
import fitz  # PyMuPDF for PDF page count

def count_pdf_pages(pdf_path):
    """
    Counts the number of pages in a PDF using PyMuPDF.
    """
    try:
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        doc.close()
        return page_count
    except Exception as e:
        print(f"❌ Error reading PDF '{pdf_path}': {e}")
        return -1

def check_all_pdfs_in_folder(pdf_folder, expected_page_count=7):
    """
    Checks if all PDF files in the specified folder contain the expected number of pages.
    """
    if not os.path.exists(pdf_folder):
        print(f"⚠️ Folder '{pdf_folder}' does not exist.")
        return

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"⚠️ No PDF files found in '{pdf_folder}'.")
        return

    all_correct = True
    count_file = 0
    for pdf_file in pdf_files:
        count_file += 1
        pdf_path = os.path.join(pdf_folder, pdf_file)
        page_count = count_pdf_pages(pdf_path)

        if page_count == expected_page_count:
            print(f"✅ '{pdf_file}' has the correct {expected_page_count} pages.", count_file)
        else:
            print(f"❌ '{pdf_file}' has {page_count} pages instead of {expected_page_count}.", count_file)
            all_correct = False

    if all_correct:
        print("\n🎉 All PDFs contain the expected number of pages!")
    else:
        print("\n⚠️ Some PDFs have incorrect page counts. Please check the errors above.")

# Example usage
pdf_folder = "pdf"  # Adjust folder path if needed
check_all_pdfs_in_folder(pdf_folder)

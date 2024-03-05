from docx import Document

def bold_name_in_tables(document_path, search_name):
    doc = Document(document_path)
    found = False

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if search_name in cell.text:
                    found = True
                    # Make the text bold
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            if search_name in run.text:
                                run.bold = True

    if found:
        doc.save('modified_document.docx')
        return True
    else:
        return False

# Example usage
document_path = 'takjil jabur fix 1445.docx'
search_name = 'Ibu Sri Wuryani'
if bold_name_in_tables(document_path, search_name):
    print("Name found and bolded.")
else:
    print("Name not found.")

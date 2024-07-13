from docx import Document

def debug_print(message):
    print(message)  # Print to console
    with open('document_debug.txt', 'a', encoding='utf-8') as debug_file:
        debug_file.write(message + '\n')

def parse_document(doc):
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        debug_print(f"Raw content: '{text}'")
        

def main():
    # Clear the debug file at the start of each run
    open('document_debug.txt', 'w').close()
    
    debug_print("Starting document parsing")
    doc = Document('./ignore/84-94A-copy.docx')
    parse_document(doc)
    debug_print("Finished document parsing")

if __name__ == "__main__":
    main()
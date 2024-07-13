from docx import Document

def debug_print(message):
    print(message)
    with open('document_debug.txt', 'a', encoding='utf-8') as debug_file:
        debug_file.write(message + '\n')

def parse_document(doc):
    for i, paragraph in enumerate(doc.paragraphs):
        raw_text = paragraph.text
        debug_print(f"Paragraph {i+1}:")
        debug_print(f"Raw content: {repr(raw_text)}")
        debug_print(f"Length: {len(raw_text)}")
        debug_print("Character by character:")
        for j, char in enumerate(raw_text):
            debug_print(f"  Position {j}: {repr(char)} (ASCII: {ord(char)})")
        debug_print("---")

def main():
    open('document_debug.txt', 'w').close()  # Clear the debug file
    
    debug_print("Starting document parsing")
    doc = Document('./ignore/84-94A.docx')
    parse_document(doc)
    debug_print("Finished document parsing")

if __name__ == "__main__":
    main()
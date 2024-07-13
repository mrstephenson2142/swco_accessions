from docx import Document
import csv
import re

FIELD_NAMES = [
    "Number", "Donor", "Courtesy of", "Street", "City, State, Zip",
    "Donation/Lending Date", "Main Entry", "Quantity", "Restrictions",
    "Priority", "Assigned to Record Group", "Assigned for Processing?",
    "Date assigned", "Processing Completed?", "Date completed", "Processor", "Lender",
    "Provenance", "Temporary Location", "Special Notes", "Returned by", "Returned",  "Date returned",
    "Assigned to", "Scope and Content Note", "Materials Received By", "Permanent Location"
]

def parse_records(doc):
    records = []
    current_record = None
    current_field = None
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        
        # Skip empty lines and "Accession Records" lines
        if not text or text == "Accession Records":
            continue
        
        # Skip lines that look like "[[number]]"
        if re.match(r'\[\[\d+\]\]', text):
            continue
        
        if text.startswith("Number"):
            if current_record:
                records.append(current_record)
            current_record = {}
            current_field = "Number"
            number_value = text.split(None, 1)[1]
            # Always prepend '19' to the number
            number_value = '19' + number_value
            current_record[current_field] = number_value
        elif current_record is not None:
            field_match = False
            for field in FIELD_NAMES:
                if text.startswith(field):
                    current_field = field
                    value = text[len(field):].strip()
                    if value.startswith(":"):
                        value = value[1:].strip()
                    current_record[current_field] = value
                    field_match = True
                    break
            
            if not field_match and current_field:
                # Append to the current field if it's a continuation
                current_record[current_field] += " " + text
    
    if current_record:
        records.append(current_record)
    
    return records

def main():
    doc = Document('./ignore/84-94A.docx')
    records = parse_records(doc)
    
    if records:
        with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=FIELD_NAMES)
            writer.writeheader()
            for record in records:
                writer.writerow(record)
        
        print(f"Processed {len(records)} records and saved to output.csv")
    else:
        print("No records found")

if __name__ == "__main__":
    main()
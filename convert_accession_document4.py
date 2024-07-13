from docx import Document
import csv
import re

FIELD_NAMES = [
    "Number", "Donor", "Courtesy of", "Street", "City, State, Zip",
    "Donation/Lending Date", "Main Entry", "Quantity", "Restrictions",
    "Priority", "Assigned to Record Group", "Assigned for Processing?",
    "Date assigned", "Processing Completed?", "Date completed", "Processor", "Lender",
    "Provenance", "Temporary Location", "Special Notes", "Returned by", "Returned",  "Date returned",
    "Assigned to", "Scope and Content Note", "Materials Received By", "Permanent Location",
    "Biographical/Historical"
]



def parse_records(doc):
    records = []
    current_record = None
    current_field = None
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        
        print(f"Processing line {i+1}: {text}")  # Debug print
        
        # Skip empty lines and "Accession Records" lines
        if not text or text == "Accession Records":
            continue
        
        # Skip lines that look like "[[number]]"
        if re.match(r'\[\[\d+\]\]', text):
            continue
        
        # Check if the line starts with a field name
        field_match = False
        for field in FIELD_NAMES:
            if text.startswith(field):
                if field == "Number":
                    # Check if it matches the specific "Number" pattern
                    if not re.match(r'Number\s+\d{2}-\d+-[a-zA-Z]', text):
                        break
                    if current_record:
                        records.append(current_record)
                    current_record = {}
                current_field = field
                value = text[len(field):].strip()
                if value.startswith(":"):
                    value = value[1:].strip()
                if current_field == "Number":
                    # Extract the number part more safely
                    number_match = re.search(r'\d{2}-\d+-[a-zA-Z]', value)
                    if number_match:
                        value = '19' + number_match.group(0)
                    else:
                        print(f"Warning: Unexpected Number format: {value}")
                        value = '19' + value  # Fallback, might need adjustment
                if current_record is not None:
                    current_record[current_field] = value
                print(f"Found field: {current_field} = {value}")  # Debug print
                field_match = True
                break
        
        if not field_match:
            # If no field match, append to the current field
            if current_record is not None and current_field:
                # Replace newlines and multiple spaces with a single space
                text = re.sub(r'\s+', ' ', text)
                if current_field in current_record:
                    current_record[current_field] += " " + text
                else:
                    current_record[current_field] = text
                print(f"Appended to {current_field}: {text}")  # Debug print
    
    if current_record:
        records.append(current_record)
    
    print(f"Total records found: {len(records)}")  # Debug print
    return records

def main():
    doc = Document('./ignore/84-94A.docx')
    records = parse_records(doc)
    
    if records:
        with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=FIELD_NAMES, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            for record in records:
                writer.writerow(record)
        
        print(f"Processed {len(records)} records and saved to output.csv")
    else:
        print("No records found")

if __name__ == "__main__":
    main()
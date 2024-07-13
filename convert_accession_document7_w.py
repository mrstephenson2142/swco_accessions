from docx import Document
import csv
import re

FIELD_NAMES = [
    "Number", "DonationTypeID", "Donor", "Courtesy of", "Street", "City, State, Zip",
    "Donation/Lending Date", "Main Entry", "Quantity", "Restrictions",
    "Priority", "Assigned to Record Group", "Assigned for Processing?",
    "Date assigned", "Processing Completed?", "Date completed", "Processor", "Lender",
    "Provenance", "Temporary Location", "Special Notes", "Returned by", "Returned", "Date returned",
    "Assigned to", "Scope and Content Note", "Materials Received By", "Permanent Location",
    "Biographical/Historical"
]

DONATION_TYPE_MAP = {'A': '1', 'B': '2', 'C': '3', 'X': '4'}

def debug_print(message):
    print(message)  # Still print to console
    with open('debug_output.txt', 'a', encoding='utf-8') as debug_file:
        debug_file.write(message + '\n')

def parse_records(doc):
    records = []
    current_record = None
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        
        debug_print(f"Processing paragraph {i+1}: {text}")
        
        # Skip empty paragraphs and "Accession Records" paragraphs
        if not text or text == "Accession Records":
            continue
        
        # Skip paragraphs that look like "[[number]]"
        if re.match(r'\[\[\d+\]\]', text):
            continue
        
        # Check if the paragraph starts with a field name
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
                
                # Extract the value (everything after the field name)
                value = text[len(field):].strip()
                
                if field == "Number":
                    # Extract Number and DonationTypeID
                    number_match = re.search(r'(\d{2}-\d+)-([a-zA-Z])', value)
                    if number_match:
                        current_record["Number"] = '19' + number_match.group(1)
                        current_record["DonationTypeID"] = DONATION_TYPE_MAP.get(number_match.group(2), '0')
                    else:
                        debug_print(f"Warning: Unexpected Number format: {value}")
                        current_record["Number"] = '19' + value
                        current_record["DonationTypeID"] = '0'
                    debug_print(f"Found field: Number = {current_record['Number']}")
                    debug_print(f"Found field: DonationTypeID = {current_record['DonationTypeID']}")
                else:
                    current_record[field] = value
                    debug_print(f"Found field: {field} = {value}")
                
                field_match = True
                break
        
        if not field_match and current_record:
            # If no field match and we have a current record, append to the last field
            last_field = list(current_record.keys())[-1]
            current_record[last_field] += " " + text
            debug_print(f"Appended to {last_field}: {text}")
    
    if current_record:
        records.append(current_record)
    
    debug_print(f"Total records found: {len(records)}")
    return records

def main():
    # Clear the debug file at the start of each run
    open('debug_output.txt', 'w').close()
    
    doc = Document('./ignore/84-94A-copy.docx')
    records = parse_records(doc)
    
    if records:
        with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=FIELD_NAMES, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            for record in records:
                writer.writerow(record)
        
        debug_print(f"Processed {len(records)} records and saved to output.csv")
    else:
        debug_print("No records found")

if __name__ == "__main__":
    main()
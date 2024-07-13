from docx import Document
import csv
import re

FIELD_NAMES = [
    "Number", "DonationTypeID", "Donor", "Courtesy of", "Street", "City", "State", "Zip",
    "Donation/Lending Date", "Main Entry", "Quantity", "Restrictions",
    "Priority", "Assigned to Record Group", "Assigned for Processing?",
    "Date assigned", "Processing Completed?", "Date completed", "Processor", "Lender",
    "Provenance", "Temporary Location", "Special Notes", "Returned by", "Returned", "Date returned",
    "Assigned to", "Scope and Content Note", "Materials Received By", "Permanent Location",
    "Biographical/Historical"
]

PROPER_CASE_FIELDS = {
    "Donor", "Courtesy of", "Street", "City", "Processor", "Lender",
    "Returned by", "Assigned to", "Materials Received By"
}

DONATION_TYPE_MAP = {'A': '1', 'B': '2', 'C': '3', 'X': '4'}

def proper_case(text):
    lowercase_words = {'of', 'the', 'in', 'on', 'at', 'to', 'for', 'and', 'by'}
    words = text.split()
    capitalized_words = [word.capitalize() if word.lower() not in lowercase_words else word.lower() for word in words]
    return ' '.join(capitalized_words)

def debug_print(message):
    print(message)
    with open('debug_output.txt', 'a', encoding='utf-8') as debug_file:
        debug_file.write(message + '\n')

def parse_city_state_zip(address):
    address = re.sub(r'\s+Ph\.\s+\d{3}/\d{3}-\d{4}', '', address)
    parts = [p.strip() for p in re.split(r'[,\s]+', address) if p.strip()]
    
    city = state = zip_code = ''
    
    if len(parts) == 1:
        city = parts[0]
    elif len(parts) == 2:
        if parts[1].isnumeric() or (len(parts[1]) == 5 and parts[1][:5].isnumeric()):
            city, zip_code = parts
        else:
            city, state = parts
    elif len(parts) >= 3:
        city = parts[0]
        for part in parts[1:]:
            if part.isnumeric() or (len(part) >= 5 and part[:5].isnumeric()):
                zip_code = part
            elif len(part) == 2 or part.lower() in ['texas', 'new mexico', 'oklahoma', 'arkansas', 'louisiana', 'mississippi', 'alabama', 'georgia', 'florida']:
                state = part
            elif not state:
                state = part
            else:
                city += ' ' + part

    state_map = {
        'tx': 'TX', 'tex': 'TX', 'texas': 'TX',
        'nm': 'NM', 'new mexico': 'NM',
        'ok': 'OK', 'okla': 'OK', 'oklahoma': 'OK',
        'ar': 'AR', 'ark': 'AR', 'arkansas': 'AR',
        'la': 'LA', 'louisiana': 'LA',
        'ms': 'MS', 'miss': 'MS', 'mississippi': 'MS',
        'al': 'AL', 'ala': 'AL', 'alabama': 'AL',
        'ga': 'GA', 'georgia': 'GA',
        'fl': 'FL', 'fla': 'FL', 'florida': 'FL'
    }
    state = state_map.get(state.lower(), state.upper())

    return city, state, zip_code

def parse_records(doc):
    records = []
    current_record = None
    
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        
        debug_print(f"Processing paragraph {i+1}: {text}")
        
        if not text or text == "Accession Records":
            continue
        
        if re.match(r'\[\[\d+\]\]', text):
            continue
        
        field_match = False
        for field in FIELD_NAMES:
            if text.startswith(field):
                if field == "Number":
                    if not re.match(r'Number\s+\d{2}-\d+-[a-zA-Z]', text):
                        break
                    if current_record:
                        records.append(current_record)
                    current_record = {}

                value = text[len(field):].strip()
                
                if field == "Number":
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
                elif field == "City, State, Zip":
                    city, state, zip_code = parse_city_state_zip(value)
                    current_record["City"] = proper_case(city)
                    current_record["State"] = state
                    current_record["Zip"] = zip_code
                    debug_print(f"Found field: City = {current_record['City']}")
                    debug_print(f"Found field: State = {current_record['State']}")
                    debug_print(f"Found field: Zip = {current_record['Zip']}")
                elif field in PROPER_CASE_FIELDS:
                    current_record[field] = proper_case(value)
                else:
                    current_record[field] = value
                    debug_print(f"Found field: {field} = {value}")
                
                field_match = True
                break
        
        if not field_match and current_record:
            last_field = list(current_record.keys())[-1]
            current_record[last_field] += " " + text
            debug_print(f"Appended to {last_field}: {text}")
    
    if current_record:
        records.append(current_record)
    
    debug_print(f"Total records found: {len(records)}")
    return records

def main():
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
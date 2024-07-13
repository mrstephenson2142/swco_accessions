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

PROPER_CASE_FIELDS = {
    "Donor", "Courtesy of", "Street", "City", "Processor", "Lender",
    "Returned by", "Assigned to", "Materials Received By"
}

DONATION_TYPE_MAP = {'A': '1', 'B': '2', 'C': '3', 'X': '4'}

STATE_MAP = {
    'al': 'AL', 'ak': 'AK', 'az': 'AZ', 'ar': 'AR', 'ca': 'CA', 'co': 'CO', 'ct': 'CT', 'de': 'DE', 'fl': 'FL',
    'ga': 'GA', 'hi': 'HI', 'id': 'ID', 'il': 'IL', 'in': 'IN', 'ia': 'IA', 'ks': 'KS', 'ky': 'KY', 'la': 'LA',
    'me': 'ME', 'md': 'MD', 'ma': 'MA', 'mi': 'MI', 'mn': 'MN', 'ms': 'MS', 'mo': 'MO', 'mt': 'MT', 'ne': 'NE',
    'nv': 'NV', 'nh': 'NH', 'nj': 'NJ', 'nm': 'NM', 'ny': 'NY', 'nc': 'NC', 'nd': 'ND', 'oh': 'OH', 'ok': 'OK',
    'or': 'OR', 'pa': 'PA', 'ri': 'RI', 'sc': 'SC', 'sd': 'SD', 'tn': 'TN', 'tx': 'TX', 'ut': 'UT', 'vt': 'VT',
    'va': 'VA', 'wa': 'WA', 'wv': 'WV', 'wi': 'WI', 'wy': 'WY',
    'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR', 'california': 'CA', 'colorado': 'CO',
    'connecticut': 'CT', 'delaware': 'DE', 'florida': 'FL', 'georgia': 'GA', 'hawaii': 'HI', 'idaho': 'ID',
    'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA', 'kansas': 'KS', 'kentucky': 'KY', 'louisiana': 'LA',
    'maine': 'ME', 'maryland': 'MD', 'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN',
    'mississippi': 'MS', 'missouri': 'MO', 'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV',
    'new hampshire': 'NH', 'new jersey': 'NJ', 'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC',
    'north dakota': 'ND', 'ohio': 'OH', 'oklahoma': 'OK', 'oregon': 'OR', 'pennsylvania': 'PA',
    'rhode island': 'RI', 'south carolina': 'SC', 'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX',
    'utah': 'UT', 'vermont': 'VT', 'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV',
    'wisconsin': 'WI', 'wyoming': 'WY'
}

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
    city = state = zip_code = address_other = ''
    
    # Ensure there's a space after each comma
    address = re.sub(r',(\S)', r', \1', address)
    
    # Add comma before state if not present
    for state_name in sorted(STATE_MAP.keys(), key=len, reverse=True):
        pattern = r'(\s)(' + re.escape(state_name) + r')\s'
        if re.search(pattern, address, re.IGNORECASE):
            address = re.sub(pattern, r',\1\2 ', address, flags=re.IGNORECASE)
            break
    
    # Add comma before zip code if not present
    address = re.sub(r'(\s)(\d{5}(-\d{4})?)\b', r',\1\2', address)
    
    # Split the address into parts
    parts = [p.strip() for p in address.split(',')]
    
    if len(parts) >= 2:
        city = parts[0]
        remaining = ' '.join(parts[1:]).strip()
        
        # Try to find state
        state_match = re.search(r'\b(' + '|'.join(STATE_MAP.keys()) + r')\b', remaining.lower())
        if state_match:
            state = STATE_MAP[state_match.group(1).lower()]
            remaining = remaining[:state_match.start()].strip() + ' ' + remaining[state_match.end():].strip()
        
        # Try to find zip code
        zip_match = re.search(r'\b\d{5}(-\d{4})?\b', remaining)
        if zip_match:
            zip_code = zip_match.group()
            remaining = remaining[:zip_match.start()].strip() + ' ' + remaining[zip_match.end():].strip()
        
        address_other = remaining.strip()
    else:
        # If there's only one part, treat it as city
        city = address
    
    return city, state, zip_code, address_other

def parse_records(doc):
    records = []
    current_record = None
    current_field = None
    address_info = {}
    
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
                        records.append((current_record, address_info))
                    current_record = {}
                    address_info = {}
                
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
                    current_record[field] = value  # Keep original value
                    city, state, zip_code, address_other = parse_city_state_zip(value)
                    address_info["City"] = proper_case(city)
                    address_info["State"] = state
                    address_info["Zip"] = zip_code
                    address_info["Address_Other"] = address_other
                    debug_print(f"Found field: City = {address_info['City']}")
                    debug_print(f"Found field: State = {address_info['State']}")
                    debug_print(f"Found field: Zip = {address_info['Zip']}")
                    debug_print(f"Found field: Address_Other = {address_info['Address_Other']}")
                elif field in PROPER_CASE_FIELDS:
                    current_record[field] = proper_case(value)
                else:
                    current_record[field] = value
                debug_print(f"Found field: {field} = {value}")
                
                current_field = field
                field_match = True
                break
        
        if not field_match and current_record and current_field:
            if current_field in current_record:
                current_record[current_field] += " " + text
            else:
                current_record[current_field] = text
            debug_print(f"Appended to {current_field}: {text}")
    
    if current_record:
        records.append((current_record, address_info))
    
    debug_print(f"Total records found: {len(records)}")
    return records

def main():
    open('debug_output.txt', 'w').close()
    
    doc = Document('./ignore/84-94A-copy.docx')
    records = parse_records(doc)
    
    if records:
        with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=FIELD_NAMES + ["City", "State", "Zip", "Address_Other"], quoting=csv.QUOTE_ALL)
            writer.writeheader()
            for record, address_info in records:
                combined_record = {**record, **address_info}
                writer.writerow(combined_record)
        
        debug_print(f"Processed {len(records)} records and saved to output.csv")
    else:
        debug_print("No records found")

if __name__ == "__main__":
    main()
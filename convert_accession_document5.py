from docx import Document
import csv
import re

FIELD_NAMES = [
    "Number", "DonationTypeID", "Donor", "Courtesy of", "Street", "City", "State", "Zip",
    "Donation/Lending Date", "Main Entry", "Quantity", "Restrictions",
    "Priority", "Assigned to Record Group", "Assigned for Processing?",
    "Date assigned", "Processing Completed?", "Date completed", "Processor", "Lender",
    "Provenance", "Temporary Location", "Special Notes",  "Returned by", "Returned", "Date returned",
    "Assigned to", "Scope and Content Note", "Materials Received By", "Permanent Location",
    "Biographical/Historical"
]

US_STATES = [
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA',
    'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT',
    'VA', 'WA', 'WV', 'WI', 'WY'
]

STATE_DICT = {
    'ALABAMA': 'AL', 'ALASKA': 'AK', 'ARIZONA': 'AZ', 'ARKANSAS': 'AR', 'CALIFORNIA': 'CA',
    'COLORADO': 'CO', 'CONNECTICUT': 'CT', 'DELAWARE': 'DE', 'FLORIDA': 'FL', 'GEORGIA': 'GA',
    'HAWAII': 'HI', 'IDAHO': 'ID', 'ILLINOIS': 'IL', 'INDIANA': 'IN', 'IOWA': 'IA',
    'KANSAS': 'KS', 'KENTUCKY': 'KY', 'LOUISIANA': 'LA', 'MAINE': 'ME', 'MARYLAND': 'MD',
    'MASSACHUSETTS': 'MA', 'MICHIGAN': 'MI', 'MINNESOTA': 'MN', 'MISSISSIPPI': 'MS', 'MISSOURI': 'MO',
    'MONTANA': 'MT', 'NEBRASKA': 'NE', 'NEVADA': 'NV', 'NEW HAMPSHIRE': 'NH', 'NEW JERSEY': 'NJ',
    'NEW MEXICO': 'NM', 'NEW YORK': 'NY', 'NORTH CAROLINA': 'NC', 'NORTH DAKOTA': 'ND', 'OHIO': 'OH',
    'OKLAHOMA': 'OK', 'OREGON': 'OR', 'PENNSYLVANIA': 'PA', 'RHODE ISLAND': 'RI', 'SOUTH CAROLINA': 'SC',
    'SOUTH DAKOTA': 'SD', 'TENNESSEE': 'TN', 'TEXAS': 'TX', 'UTAH': 'UT', 'VERMONT': 'VT',
    'VIRGINIA': 'VA', 'WASHINGTON': 'WA', 'WEST VIRGINIA': 'WV', 'WISCONSIN': 'WI', 'WYOMING': 'WY'
}

DONATION_TYPE_MAP = {'A': '1', 'B': '2', 'C': '3', 'X': '4'}

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
                    # Extract Number and DonationTypeID
                    number_match = re.search(r'(\d{2}-\d+)-([a-zA-Z])', text)
                    if number_match:
                        current_record["Number"] = '19' + number_match.group(1)
                        current_record["DonationTypeID"] = DONATION_TYPE_MAP.get(number_match.group(2), '0')
                    else:
                        print(f"Warning: Unexpected Number format: {text}")
                        current_record["Number"] = '19' + text.split(None, 1)[1]
                        current_record["DonationTypeID"] = '0'
                    print(f"Found field: Number = {current_record['Number']}")
                    print(f"Found field: DonationTypeID = {current_record['DonationTypeID']}")
                    field_match = True
                    break
                elif field == "City, State, Zip":
                    current_field = field
                    value = text[len(field):].strip()
                    if value.startswith(":"):
                        value = value[1:].strip()
                    
                    # Handle special case for "Campus"
                    if value.lower() == "campus":
                        current_record["City"] = "Campus"
                        current_record["State"] = ""
                        current_record["Zip"] = ""
                    else:
                        # Split by comma
                        parts = [part.strip() for part in value.split(',') if part.strip()]
                        
                        if len(parts) == 1:
                            # Only one part, assume it's the city
                            current_record["City"] = parts[0]
                            current_record["State"] = ""
                            current_record["Zip"] = ""
                        elif len(parts) == 2:
                            # Two parts, assume city and state+zip
                            current_record["City"] = parts[0]
                            state_zip = parts[1].split()
                            if len(state_zip) > 1 and state_zip[-1].isdigit():
                                state = " ".join(state_zip[:-1])
                                current_record["State"] = STATE_DICT.get(state.upper(), state)
                                current_record["Zip"] = state_zip[-1]
                            else:
                                current_record["State"] = STATE_DICT.get(parts[1].upper(), parts[1])
                                current_record["Zip"] = ""
                        elif len(parts) >= 3:
                            # Three or more parts
                            current_record["City"] = parts[0]
                            state = parts[-2]
                            current_record["State"] = STATE_DICT.get(state.upper(), state)
                            current_record["Zip"] = parts[-1] if parts[-1].isdigit() else ""
                    
                    print(f"Found field: City = {current_record['City']}")
                    print(f"Found field: State = {current_record['State']}")
                    print(f"Found field: Zip = {current_record['Zip']}")
                    field_match = True
                    break
                else:
                    current_field = field
                    value = text[len(field):].strip()
                    if value.startswith(":"):
                        value = value[1:].strip()
                    current_record[current_field] = value
                    print(f"Found field: {current_field} = {value}")
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
                print(f"Appended to {current_field}: {text}")
    
    if current_record:
        records.append(current_record)
    
    print(f"Total records found: {len(records)}")
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
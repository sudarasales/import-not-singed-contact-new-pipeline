import pandas as pd
import numpy as np
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Read the template to get the exact column structure
df_template = pd.read_excel('pipedrive_template_data.xlsx')
template_columns = df_template.columns.tolist()
print(f'Template columns: {template_columns}\n')

def split_name(name_str):
    """Split name into first and last name"""
    if pd.isna(name_str) or name_str == '' or str(name_str) == 'nan':
        return ('', '')
    name_str = str(name_str).strip()
    parts = name_str.split(maxsplit=1)
    if len(parts) >= 2:
        return (parts[0], ' '.join(parts[1:]))
    else:
        return (parts[0], '')

def parse_contact_info(contact_str):
    """Parse contact information to extract phone and email"""
    phone = ''
    email = ''
    
    if pd.isna(contact_str) or contact_str == '' or str(contact_str) == 'nan':
        return phone, email
    
    contact_str = str(contact_str).strip()
    
    # Check if it contains an email
    if '@' in contact_str:
        # Extract email
        import re
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', contact_str)
        if email_match:
            email = email_match.group(0)
            # Remove email from string to extract phone
            contact_str = contact_str.replace(email, '').strip()
    
    # Extract phone numbers (digits, may have + prefix)
    import re
    # Look for phone patterns
    phone_patterns = [
        r'\+?\d{9,15}',  # Standard phone number
        r'\d{9,15}',     # Phone without +
    ]
    
    for pattern in phone_patterns:
        phone_matches = re.findall(pattern, contact_str)
        if phone_matches:
            # Take the first phone number found
            phone = phone_matches[0]
            # Clean up phone (remove spaces, keep + if present)
            phone = phone.replace(' ', '').replace('-', '')
            break
    
    # If no pattern match but it's a number, use it as phone
    if not phone and contact_str.replace(' ', '').replace('-', '').replace('+', '').isdigit():
        phone = contact_str.replace(' ', '').replace('-', '')
    
    # If still no phone and contact_str doesn't look like email, use as phone
    if not phone and '@' not in contact_str and contact_str:
        # Clean and use as phone
        phone = contact_str.replace(' ', '').replace('-', '').replace('Agent Number', '').replace('Client Number', '').replace(',', '').strip()
    
    return phone, email

def format_phone(phone):
    """Format phone number"""
    if pd.isna(phone) or phone == '' or str(phone) == 'nan':
        return ''
    try:
        # If it's a float/int, convert to int then string
        if isinstance(phone, (int, float)):
            return str(int(phone))
        phone_str = str(phone).strip()
        # Remove common prefixes/suffixes
        phone_str = phone_str.replace('Agent Number', '').replace('Client Number', '').replace(',', '').strip()
        return phone_str
    except:
        return str(phone)

# Process each file
joyce_folder = "Joyce Agents"
files_to_process = [
    ("Agent Visit 2023 - Company.xlsx", True),  # Has headers
    ("Agent Visit 2024 - Company.xlsx", False),  # No headers
    ("Agent Visit 2025 - Company.xlsx", False),  # No headers
]

for filename, has_headers in files_to_process:
    filepath = os.path.join(joyce_folder, filename)
    print(f"\n{'='*60}")
    print(f"Processing: {filename}")
    print('='*60)
    
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        continue
    
    # Read the source file
    if has_headers:
        df_source = pd.read_excel(filepath)
    else:
        # Read without headers, treat all rows as data
        df_source = pd.read_excel(filepath, header=None)
        # Assign column names based on structure
        if "2024" in filename:
            df_source.columns = ['Agent name', 'Agency Company', 'Property', 'Unnamed']
        elif "2025" in filename:
            df_source.columns = ['Agent name', 'Phone', 'Email', 'Agency Company']
    
    print(f"Source shape: {df_source.shape}")
    print(f"Source columns: {df_source.columns.tolist()}")
    
    # Create new dataframe with template structure
    df_transformed = pd.DataFrame(index=df_source.index, columns=template_columns)
    
    # Deal - Title and Deal - Value: empty strings
    df_transformed['Deal - Title'] = ''
    df_transformed['Deal - Value'] = ''
    
    # Person - Name *: from 'Agent name' column
    df_transformed['Person - Name *'] = df_source['Agent name'].fillna('').astype(str)
    df_transformed['Person - Name *'] = df_transformed['Person - Name *'].replace('nan', '')
    
    # Person - First name and Last name: split from 'Agent name'
    name_splits = df_source['Agent name'].apply(split_name)
    df_transformed['Person - First name'] = [x[0] for x in name_splits]
    df_transformed['Person - Last name'] = [x[1] for x in name_splits]
    
    # Person - Email and Phone: depends on file structure
    if "2023" in filename:
        # 2023: 'Visitor Contact Information' contains phone/email
        contact_info = df_source['Visitor Contact Information'].fillna('').astype(str)
        emails = []
        phones = []
        for info in contact_info:
            phone, email = parse_contact_info(info)
            phones.append(phone)
            emails.append(email)
        df_transformed['Person - Phone (suggested)'] = phones
        df_transformed['Person - Email (suggested)'] = emails
    elif "2024" in filename:
        # 2024: No separate phone/email columns
        df_transformed['Person - Phone (suggested)'] = ''
        df_transformed['Person - Email (suggested)'] = ''
    elif "2025" in filename:
        # 2025: Has separate 'Phone' and 'Email' columns
        df_transformed['Person - Phone (suggested)'] = df_source['Phone'].apply(format_phone)
        df_transformed['Person - Email (suggested)'] = df_source['Email'].fillna('').astype(str)
        df_transformed['Person - Email (suggested)'] = df_transformed['Person - Email (suggested)'].replace('nan', '')
    
    # Organization - Name *: from 'Agency Company' column
    df_transformed['Organization - Name *'] = df_source['Agency Company'].fillna('').astype(str)
    df_transformed['Organization - Name *'] = df_transformed['Organization - Name *'].replace('nan', '')
    
    # Organization - Address (suggested): empty (not available in source files)
    df_transformed['Organization - Address (suggested)'] = ''
    
    # Clean up any remaining 'nan' strings and NaN values
    for col in df_transformed.columns:
        df_transformed[col] = df_transformed[col].replace('nan', '')
        df_transformed[col] = df_transformed[col].fillna('')
    
    print(f"\nTransformed shape: {df_transformed.shape}")
    print(f"First 3 rows preview:")
    print(df_transformed.head(3))
    
    # Create output filename
    output_filename = filename.replace('.xlsx', '_transformed.xlsx')
    output_path = os.path.join(joyce_folder, output_filename)
    
    # Save transformed data (replace NaN with empty strings)
    df_transformed = df_transformed.fillna('')
    df_transformed.to_excel(output_path, index=False)
    print(f"\n✓ Saved: {output_path}")

print("\n" + "="*60)
print("All files processed successfully!")
print("="*60)


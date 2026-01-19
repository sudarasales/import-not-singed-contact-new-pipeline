import pandas as pd
import numpy as np

# Read the source file (use backup if original was already transformed)
import os
source_file = 'AGENTS ANNA crm_backup.xlsx' if os.path.exists('AGENTS ANNA crm_backup.xlsx') else 'AGENTS ANNA crm.xlsx'
df_source = pd.read_excel(source_file)
print(f'Source file: {df_source.shape[0]} rows, {df_source.shape[1]} columns')
print(f'Source columns: {df_source.columns.tolist()}')

# Read the template to get the exact column structure
df_template = pd.read_excel('pipedrive_template_data.xlsx')
template_columns = df_template.columns.tolist()
print(f'Template columns: {template_columns}')

# Create new dataframe with template structure
df_transformed = pd.DataFrame(index=df_source.index, columns=template_columns)

# Map the data
# Deal - Title and Deal - Value: empty strings as specified
df_transformed['Deal - Title'] = ''
df_transformed['Deal - Value'] = ''

# Person - Name *: from 'name' column
df_transformed['Person - Name *'] = df_source['name'].fillna('').astype(str)

# Person - First name and Last name: split from 'name'
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

name_splits = df_source['name'].apply(split_name)
df_transformed['Person - First name'] = [x[0] for x in name_splits]
df_transformed['Person - Last name'] = [x[1] for x in name_splits]

# Person - Email (suggested): from 'email' column
df_transformed['Person - Email (suggested)'] = df_source['email'].fillna('').astype(str)
# Convert 'nan' strings to empty
df_transformed['Person - Email (suggested)'] = df_transformed['Person - Email (suggested)'].replace('nan', '')

# Person - Phone (suggested): from 'phone' column
df_transformed['Person - Phone (suggested)'] = df_source['phone'].fillna('').astype(str)
# Convert 'nan' strings to empty and format phone numbers
df_transformed['Person - Phone (suggested)'] = df_transformed['Person - Phone (suggested)'].replace('nan', '')
# Convert float phone numbers to int then string (remove decimal points)
def format_phone(phone):
    if pd.isna(phone) or phone == '' or str(phone) == 'nan':
        return ''
    try:
        # If it's a float/int, convert to int then string
        if isinstance(phone, (int, float)):
            return str(int(phone))
        return str(phone)
    except:
        return str(phone)

df_transformed['Person - Phone (suggested)'] = df_transformed['Person - Phone (suggested)'].apply(format_phone)

# Organization - Name *: from 'Company' column (as user specified)
df_transformed['Organization - Name *'] = df_source['Company'].fillna('').astype(str)
df_transformed['Organization - Name *'] = df_transformed['Organization - Name *'].replace('nan', '')

# Organization - Address (suggested): from 'address' column
df_transformed['Organization - Address (suggested)'] = df_source['address'].fillna('').astype(str)
df_transformed['Organization - Address (suggested)'] = df_transformed['Organization - Address (suggested)'].replace('nan', '')

# Replace empty strings with NaN for Deal - Title and Deal - Value to match template format
# But first ensure they are set properly - if user wants empty, we'll use empty strings
# Actually, let's keep them as empty strings initially and see how Excel handles them

print(f'\nTransformed file: {df_transformed.shape[0]} rows, {df_transformed.shape[1]} columns')
print('\nFirst 5 rows preview:')
print(df_transformed.head())

# Save to the original filename (backup first, then overwrite)
import shutil
if not os.path.exists('AGENTS ANNA crm_backup.xlsx'):
    shutil.copy('AGENTS ANNA crm.xlsx', 'AGENTS ANNA crm_backup.xlsx')
    print('\nBackup saved as: AGENTS ANNA crm_backup.xlsx')

# Save transformed data
df_transformed.to_excel('AGENTS ANNA crm.xlsx', index=False)
print('Transformed file saved as: AGENTS ANNA crm.xlsx')
print('\nDone!')


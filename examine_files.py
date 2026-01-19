import pandas as pd
import sys

# Set encoding for output
sys.stdout.reconfigure(encoding='utf-8')

# Read both files
df1 = pd.read_excel('AGENTS ANNA crm.xlsx')
df2 = pd.read_excel('pipedrive_template_data.xlsx')

print('=== AGENTS ANNA crm.xlsx ===')
print('Columns:', df1.columns.tolist())
print('Shape:', df1.shape)
print('\nColumn data types:')
for col in df1.columns:
    print(f'  {col}: {df1[col].dtype}')

print('\n\n=== pipedrive_template_data.xlsx ===')
print('Columns:', df2.columns.tolist())
print('Shape:', df2.shape)
print('\nColumn data types:')
for col in df2.columns:
    print(f'  {col}: {df2[col].dtype}')

# Save sample to file for inspection
with open('output_sample.txt', 'w', encoding='utf-8') as f:
    f.write('=== AGENTS ANNA crm.xlsx ===\n')
    f.write(str(df1.head(5).to_dict()))
    f.write('\n\n=== pipedrive_template_data.xlsx ===\n')
    f.write(str(df2.head(5).to_dict()))

print('\n\nSample data saved to output_sample.txt')


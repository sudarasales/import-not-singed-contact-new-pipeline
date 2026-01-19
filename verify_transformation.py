import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Read the transformed file
df_transformed = pd.read_excel('AGENTS ANNA crm.xlsx')
df_template = pd.read_excel('pipedrive_template_data.xlsx')

print('=== Verification ===')
print(f'Transformed columns: {df_transformed.columns.tolist()}')
print(f'Template columns: {df_template.columns.tolist()}')
print(f'\nColumn match: {list(df_transformed.columns) == list(df_template.columns)}')
print(f'\nShape: {df_transformed.shape}')
print(f'\nFirst 5 rows:')
print(df_transformed.head(5).to_string())

# Check data types
print('\n\nData types:')
for col in df_transformed.columns:
    print(f'{col}: {df_transformed[col].dtype}')
    # Show non-empty sample values
    non_empty = df_transformed[col][df_transformed[col].notna() & (df_transformed[col].astype(str) != '') & (df_transformed[col].astype(str) != 'nan')]
    if len(non_empty) > 0:
        print(f'  Sample values: {non_empty.head(3).tolist()}')


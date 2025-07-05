import pandas as pd

# File paths
input_file = r'C:\Users\dylan\OneDrive\Documents\python_data_cleaner\dirty_data_2.xlsx'
output_file = r'C:\Users\dylan\OneDrive\Documents\python_data_cleaner\cleaned_data.xlsx'

# Step 1. Read raw data with no headers
df_raw = pd.read_excel(input_file, header=None)

# Step 2. Promote appropriate rows as header (rows 0 and 1 in your screenshot)
new_header = df_raw.iloc[0:2]
data = df_raw.iloc[2:].copy()

# Combine multi-level headers into single strings
combined_headers = []
for col in new_header.columns:
    combined_name = ''
    for row in new_header[col]:
        if pd.notnull(row):
            combined_name += str(row).strip() + ' '
    combined_headers.append(combined_name.strip())

# Assign combined headers to data
data.columns = combined_headers

# Step 3. Rename first column explicitly as 'Order Date'
first_col = data.columns[0]
data.rename(columns={first_col: 'Order Date'}, inplace=True)

# Step 4. Melt dataframe to long format
id_vars = ['Order Date']
value_vars = [col for col in data.columns if col != 'Order Date']

df_melted = pd.melt(data, id_vars=id_vars, value_vars=value_vars, var_name='Ship Segment', value_name='Sales')

# Step 5. Drop rows with NaN Sales
df_melted = df_melted.dropna(subset=['Sales'])

# Step 6. Split 'Ship Segment' into 'Ship Mode' and 'Segment'
df_melted[['Ship Mode', 'Segment']] = df_melted['Ship Segment'].str.extract(r'(\w+\s*\w*)\s*(Consumer|Corporate|Home Office)', expand=True)

# Step 7. Standardise date format
df_melted['Order Date'] = pd.to_datetime(df_melted['Order Date'], errors='coerce').dt.strftime('%Y-%m-%d')

# Step 8. Drop rows with invalid dates
df_melted = df_melted.dropna(subset=['Order Date'])

# Step 9. Reorder columns for clarity
df_melted = df_melted[['Order Date', 'Ship Mode', 'Segment', 'Sales']]

# Step 10. Reset index
df_melted.reset_index(drop=True, inplace=True)

# Step 11. Save cleaned data
df_melted.to_excel(output_file, index=False)

print(f"Cleaned data saved to '{output_file}' successfully.")

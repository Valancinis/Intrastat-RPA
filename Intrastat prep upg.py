import pandas as pd
import os
import openpyxl
from openpyxl.styles import PatternFill

# Load the CSV file into a pandas DataFrame, skipping the first two rows
df = pd.read_csv('Intrastat.csv', skiprows=[0,1], dtype={'SKU': str})

# Load the Nav Intrastat data
nav_df = pd.read_excel('Additional Data.xlsx', usecols=['No_', 'Country of Origin Code', 'Tariff No_', 'Weight per purch UoM'], dtype={'Tariff No_': str})

# Load excel x-ref data
df_refs= pd.read_excel('Additional Data.xlsx', sheet_name='IDx-Ref', dtype={'Item number': str, 'Vendor item number' : str})

# Delete rows that have "." as a value in "Declarant Country" column
df = df[df['Declarant Country'] != '.']

# Replace any asterisks in the 'Country Of Original Origin' column with 'FIX'
df['Country Of Original Origin'] = df['Country Of Original Origin'].str.replace('*', 'FIX', regex=False)
df['Country Of Original Origin'] = df['Country Of Original Origin'].str.replace('LT', 'FIX', regex=False)
df['Country Of Original Origin'] = df['Country Of Original Origin'].fillna('FIX')

# Replace any '11111111' or 'BLANK' values in the 'Commodity Code' column with 'FIX'
df['Commodity Code'] = df['Commodity Code'].replace(['11111111', 'BLANK'], 'FIX')
df['Commodity Code'] = df['Commodity Code'].fillna('FIX')

# Fill any blank cells in the 'Commodity Code', 'Net Mass in KG' columns with 'FIX'
df['Commodity Code'] = df['Commodity Code'].fillna('FIX')
df['Net Mass in KG'] = df['Net Mass in KG'].fillna('FIX')

# Add a new column called "Navision check?" that has a value of "Yes" if there is a "FIX" value in the 'Country Of Original Origin', 'Commodity Code', and 'Net Mass in KG' columns of the row
df['Navision check?'] = df.apply(lambda row: 'Yes' if (
    row['Country Of Original Origin'] is not None and not isinstance(row['Country Of Original Origin'], float) and 'FIX' in row['Country Of Original Origin']) or
    (row['Commodity Code'] is not None and not isinstance(row['Commodity Code'], float) and 'FIX' in row['Commodity Code']) or
    (row['Net Mass in KG'] is not None and not isinstance(row['Net Mass in KG'], float) and 'FIX' in row['Net Mass in KG']) else 'No', axis=1)

# Fill any blank cells in the 'Mode of Trn', 'Conditions of Transport', 'QTY Received', and 'Taxable Amount' columns with 'FIX'
df['Mode of Trn'] = df['Mode of Trn'].fillna('FIX')
df['Conditions of Transport'] = df['Conditions of Transport'].fillna('FIX')
df['QTY Received'] = df['QTY Received'].fillna('FIX')
df['Taxable Amount'] = df['Taxable Amount'].fillna('FIX')

# Add a new column called "Intrastat check?" that has a value of "Yes" if there is a "FIX" value in any cell of the row
df['Intrastat check?'] = df.apply(lambda row: 'Yes' if 'FIX' in row.values else 'No', axis=1)

# drop nav_df duplicate values
nav_df = nav_df.drop_duplicates(subset=['No_'], keep='first')

# Save the df to a separate Excel file before merging
df.to_excel('output_original.xlsx', index=False)

# merge dataframes based on "SKU" in "df" and "No_" in "nav_df"
merged_df = pd.merge(df, nav_df, left_on=['SKU'], right_on=['No_'], how='left')

# update columns based on "Needs update?" column
merged_df.loc[(merged_df['Navision check?'] == 'Yes') & (merged_df['Country of Origin Code'].notnull()), 'Country Of Original Origin'] = merged_df['Country of Origin Code']
merged_df.loc[(merged_df['Navision check?'] == 'Yes') & (merged_df['Tariff No_'].notnull()) & (merged_df['Tariff No_'] != '00000000'), 'Commodity Code'] = merged_df['Tariff No_']
merged_df.loc[(merged_df['Navision check?'] == 'Yes') & (merged_df['Weight per purch UoM'].notnull()), 'Net Mass in KG'] = merged_df['Weight per purch UoM'] * merged_df['QTY Received']

# Filter rows based on "Needs update?" column
updated_rows = merged_df.loc[merged_df['Navision check?'] == 'Yes']

# Save the updated rows to a separate Excel file
updated_rows.to_excel('updated_rows_nav.xlsx', index=False)

# Drop unnecessary columns
merged_df.drop(['No_', 'Country of Origin Code', 'Tariff No_', 'Weight per purch UoM', 'Navision check?'], axis=1, inplace=True)

# Update df with the updated values
df = merged_df

# Create a new column called "Needs update?"
df['Needs update?'] = df.apply(lambda row: 'Yes' if (row['Country Of Original Origin'] == 'FIX') or
                                                      (row['Commodity Code'] == 'FIX') or
                                                      (row['Net Mass in KG'] == 'FIX')
                                                      else 'No', axis=1)

# Remove duplicates on df_refs
df_refs = df_refs.drop_duplicates(subset=['Item Number'], keep='first')

# Merge the two dataframes based on the 'Item Number' and 'SKU' columns
df_merged_xref = pd.merge(df, df_refs[['Item Number', 'Vendor Item Number']], how='left', left_on='SKU', right_on='Item Number')

# Drop the 'Item Number' column
df_merged_xref.drop('Item Number', axis=1, inplace=True)

# Setting new df value
df = df_merged_xref

# Save the new DataFrame to a separate Excel file
df.to_excel('output.xlsx', index=False)

# Group the DataFrame by 'Account'
grouped = df.groupby('Account')
# Replace any 'FIX' values in the DataFrame with a blank value
df.replace('FIX', '', inplace=True)

# Create the directory if it doesn't exist
directory = "Requests for suppliers"
if not os.path.exists(directory):
    os.makedirs(directory)

for account, group in grouped:
    # Check if the group has any 'Yes' values in the 'Supplier check?' column
    if 'Yes' in group['Needs update?'].values:
        # Create a new DataFrame with only the rows where 'Supplier check?' is 'Yes'
        group_new = group.loc[
            group['Needs update?'] == 'Yes', ['Account', 'Name', 'SKU', 'Vendor Item Number', 'Country Of Original Origin',
                                                'Commodity Code', 'Net Mass in KG']]

        # Replace blank values in the Vendor Item Number column with the modified SKU value
        group_new.loc[group_new['Vendor Item Number'].isnull(), 'Vendor Item Number'] = group_new['SKU'].str.replace('-VNO', '')

        # Remove duplicate ID values
        group_new = group_new.drop_duplicates(subset=['SKU'])

        # Save the new DataFrame to a separate Excel file named after the account value and name value
        account_value = str(account).replace('.0', '')
        name_value = group.loc[group['Account'] == account, 'Name'].iloc[0].replace('/', '_')
        file_name = f"{account_value} - {name_value}.xlsx"
        file_path = os.path.join(directory, file_name)
        group_new.to_excel(file_path, index=False)

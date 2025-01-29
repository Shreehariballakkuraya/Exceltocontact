import pandas as pd

# Load the Excel file
file_path = "contacts11.xlsx"
df = pd.read_excel(file_path)

# Check the first few rows to see the structure
print(df.head())

# Assuming the Excel sheet has columns 'Name' and 'Phone'
contacts = df[['Name', 'Phone']].values.tolist()

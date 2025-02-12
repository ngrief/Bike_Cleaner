import pandas as pd

# Load the Excel file
file_path = "Resources/merged.xlsx"

try:
    df = pd.read_excel(file_path, engine="openpyxl")  # Use openpyxl for reading .xlsx files
except Exception as e:
    print(f"Error reading file: {e}")
    exit()

# Data Cleaning
df.drop_duplicates(inplace=True)  # Remove duplicate rows
df.dropna(inplace=True)  # Remove rows with missing values
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")  # Standardize column names

# Sampling 10% of the data
sampled_df = df.sample(frac=0.1, random_state=42)  # Use a fixed random state for reproducibility

# Save cleaned & sampled data as a new Excel file
output_file = "Resources/cleaned_sampled_data.xlsx"
sampled_df.to_excel(output_file, index=False, engine="openpyxl")

print(f"Cleaned and sampled data saved to {output_file}")

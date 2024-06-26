import pandas as pd
import glob

# Directory containing CSV files
csv_directory = 'E:/Karan_Bais/python/Ml/Projects/stock/New folder'

# Output Excel file path
excel_file = 'combined_workbook.xlsx'

# Create a new Excel writer object
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    # Find all CSV files in the directory
    csv_files = glob.glob(f"{csv_directory}/*.csv")
    
    # Iterate over CSV files and add each as a new sheet
    for csv_file in csv_files:
        # Extract the file name without extension to use as sheet name
        sheet_name = csv_file.split('\\')[-1].replace('.csv', '')

        # Read CSV file
        # print(sheet_name)
        df = pd.read_csv(csv_file)
        # Write DataFrame to a new sheet in the Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"All CSV files have been combined into {excel_file}")
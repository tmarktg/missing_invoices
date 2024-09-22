"""
This script finds the missing Purchase orders (PO#s) in a set of excel files that are invoices
It compares the PO#s in the excel files against a list of Purchase Order numbers from a directory
If any PO#s are missing from the Excel files, it prints a list of the missing PO#s

Technologies used:
- pandas: for processing the excel files

Inputs:
- A folder in the Desktop directory that contains a series of PO numbers called PO#s
- A folder in the Desktop directory that contains a series of excel files which are invoices

Outputs:
- Prints a message for the excel files that have missing purchase orders that need to be added to the PO#s list
- Ouputs a set of all the missing Purchase order numbers that need to be retrieved from the database

Author: Mark Truong
Date: August 30, 2024
"""
import os
import pandas as pd

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get the paths from environment variables
path_po_numbers = os.getenv('PO_NUMBERS_PATH')
path_invoices = os.getenv('INVOICES_PATH')

# Print paths for verification (optional)
print(f"PO Numbers Path: {path_po_numbers}")
print(f"Invoices Path: {path_invoices}")

# Generate list of available PO#s from the file directory
arr1 = [int(po_number) for po_number in os.listdir(
    path_po_numbers) if po_number.isnumeric()]
arr1_set = set(arr1)  # Convert to set for faster lookup

# Missing invoices set
arr3_set = set()

# Iterate over each excel sheet in a file
for filename in os.listdir(path_invoices):
    # Skip non-Excel files
    if not filename.endswith(('.xlsx', '.xls')):
        continue

    try:
        # Load the invoice sheet
        df = pd.read_excel(os.path.join(
            path_invoices, filename), sheet_name='Sheet1')

        # Extract the PO# column and remove any NaN values
        arr2 = df['PO#'].dropna().tolist()

        # Check for missing invoices
        if all(item in arr1_set for item in arr2):
            print(f"No missing invoices in {filename}")
        else:
            print(f"Missing invoices in {filename}")
            for item in arr2:
                if item not in arr1_set:
                    arr3_set.add(item)

    except Exception as e:
        print(f"Could not read {filename}: {e}")

print("Missing Invoices:", arr3_set)

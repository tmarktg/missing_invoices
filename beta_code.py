"""First Working Version of Code"""

import openpyxl
import os


# find the file list of PO#s
path_po_numbers = os.path.expanduser('~/Desktop/roy_work')
path_invoices = os.path.expanduser('~/Desktop/excel_files')

# See if there are any missing PO#s by seeing if the invoice list is a subset of the PO#s list


def is_subset(arr1, arr2, m, n):
    # Iterate over each element in the second array
    for i in range(n):
        found = False

        # Check if the element exists in the first array
        for j in range(m):
            if arr2[i] == arr1[j]:
                found = True

        # If any element is not found, return false
        if not found:
            return False

    # If all elements are found, return true
    return True


# missing invoices array
arr3 = []

# Iterate over each excel sheet in a file
for i in os.listdir(path_invoices):
    # mock invoice sheet
    wb = openpyxl.load_workbook(os.path.join(path_invoices, i))
    sheet = wb['Sheet1']
    b = sheet['B1']

    arr1 = []
    arr2 = []

    # implement regex later
    # This checks if cell is the PO# so we can grab the list of PO#s from the expense sheet
    if sheet['B1'].value == 'Apples':
        # Generate second array (subset) from invoice
        for po_number in range(1, sheet.max_row+1):
            arr2.append(sheet[f'C{po_number}'].value)
            # print(sheet[f'C{po_number}'].value)

    # grabs only numbers to make a list of PO#s
    for i in os.listdir(path_po_numbers):
        if i.isnumeric():
            arr1.append(int(i))

    if is_subset(arr1, arr2, len(arr1), len(arr2)):
        print("There are no missing invoices")
    else:
        print("There are missing invoices")
        # Print elements in arr1 that aren't in arr2
        for item in arr2:
            if item not in arr1 and item is not None:
                # print(item)
                if item not in arr3:
                    arr3.append(item)

    # print(arr1)
    # print(arr2)
    print(arr3)

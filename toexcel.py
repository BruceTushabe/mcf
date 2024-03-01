from openpyxl import load_workbook
from docx import Document
import pandas as pd

# Function to find data for a specific account number in the Excel sheet
def find_account_data(account_number):
    account_number = int(account_number) # Convert to string for comparison
    workbook = load_workbook('data1.xlsx', data_only=True)
    sheet = workbook.active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[2] == account_number:  # Assuming the account number is in the 6th column (index 4)
            return row  # Return the row data if account number is found
    return None  # Return None if account number is not found

# Load the Word document template
document = Document('template.docx')

# Function to populate the Word document with data for a specific account number
def populate_document_for_account(document, account_data):
    if account_data:
        ACCT_NAME = account_data['ACCT_NAME'].iloc[0]
        Gender = account_data['Gender'].iloc[0]
        Contact = account_data['Contact'].iloc[0]
        ACCNUM = account_data['ACCNUM'].iloc[0]
        AGE_CATEGORY = account_data['AGE_CATEGORY'].iloc[0]
        Address = account_data['Address'].iloc[0]
        DIS_SHDL_DATE = account_data['DIS_SHDL_DATE'].iloc[0]
        DIS_AMT = account_data['DIS_AMT'].iloc[0]
        AMOUNT_CLAIMED = account_data['AMOUNT_CLAIMED'].iloc[0]
        # Add other columns as needed
        
        # Populate the Word document with data
        document.tables[0].cell(0, 1).text = ACCT_NAME
        document.tables[0].cell(1, 1).text = Gender
        document.tables[0].cell(2, 1).text = Contact
        document.tables[0].cell(3, 1).text = str(ACCNUM)
        document.tables[0].cell(4, 1).text = AGE_CATEGORY
        document.tables[0].cell(5, 1).text = Address
        document.tables[0].cell(6, 1).text = str(DIS_SHDL_DATE)
        document.tables[0].cell(7, 1).text = str(DIS_AMT)
        document.tables[0].cell(8, 1).text = str(AMOUNT_CLAIMED)
        # Add other cells as needed
        
        return True  # Data populated successfully
    return False  # Account number not found

# Input the account number
account_number = input("Enter the account number: ")

# Find data for the specified account number in the Excel sheet
account_data = find_account_data(account_number)

# Populate the Word document with data for the specified account number
if populate_document_for_account(document, account_data):
    # Save the populated Word document
    output_word_filename = f"output_word_{account_number}.docx"
    document.save(output_word_filename)
    print(f"Word document generated for account number {account_number}.")
else:
    print(f"Account number {account_number} not found in the Excel sheet.")

# Create a DataFrame for Excel sheet
if account_data:
    excel_data = {
        'ACCT_NAME': [account_data['ACCT_NAME'].iloc[0]],
        'Gender': [account_data['Gender'].iloc[0]],
        'Contact': [account_data['Contact'].iloc[0]],
        'ACCNUM': [account_data['ACCNUM'].iloc[0]],
        'AGE_CATEGORY': [account_data['AGE_CATEGORY'].iloc[0]],
        'Address': [account_data['Address'].iloc[0]],
        'DIS_SHDL_DATE': [account_data['DIS_SHDL_DATE'].iloc[0]],
        'DIS_AMT': [account_data['DIS_AMT'].iloc[0]],
        'AMOUNT_CLAIMED': [account_data['AMOUNT_CLAIMED'].iloc[0]],
        # Add other columns as needed
    }
    df = pd.DataFrame(excel_data)
    
    # Save the DataFrame to Excel sheet
    output_excel_filename = f"output_excel_{account_number}.xlsx"
    df.to_excel(output_excel_filename, index=False)
    print(f"Excel sheet generated for account number {account_number}.")
else:
    print(f"Account number {account_number} not found in the Excel sheet.")


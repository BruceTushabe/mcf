import pandas as pd
from mailmerge import MailMerge

# Load Excel data into a DataFrame
excel_file = 'data1.xlsx'  # Update with your Excel file path
df = pd.read_excel(excel_file)

# Load Word template
template_file = 'template.docx'  # Update with your Word template file path
document = MailMerge(template_file)

# Initialize counter
counter = 0

# Iterate through each row in Excel and fill in the Word template
for index, row in df.iterrows():
    print(row)
    
    # Fill in the placeholders with data from the current row
    document.merge(
        
        name = str(row['ACCT_NAME']),
                
    #    age=int(row['TERM']),              
    #    client_resident=str(row['SCHM_CODE']),  
    #   group=str(row['SOL_ID']),                
    #    pronoun=str(row['SCHM_CODE']),           
    #    occupation=str(row['SCHM_CODE']),        
    #   Branch=str(row['SCHEME NAME']),          
    #   amount_disbursed=float(row['DIS_AMT']),  
    #   cause_of_default=str(row['ACCT_MGR_USER_ID']) 
        
    )
    
    # Save filled document for each individual
    output_file = f'output_{row["ACCT_NAME"]}.docx'  
    document.write(output_file)

    # Increment counter
    counter += 1

    # Check if we've processed 5 individuals
    if counter >= 5:
        break

print("Documents filled and saved successfully for the first 5 individuals.")
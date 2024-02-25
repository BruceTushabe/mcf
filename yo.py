import docx
import pandas as pd


def create_customer_profile(client_data):
    """Defines the structure and content of the customer profile."""
    profile = f"Customer Profile\n\n"
    profile += f"Name: {client_data['ACCT_NAME']}\n"
    profile += f"Account Number: {client_data['ACCTNUM']}\n"
    profile += f"Principal Balance: {client_data['PRINCIPLE_BALANCE']}\n"
    
    # Add more fields as needed
    return profile


def generate_customer_profiles(input_file, output):
    """Reads Excel data, creates customer profiles, and saves them to Word documents."""

    # Read Excel data
    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
        # Create customer profile
        profile_text = create_customer_profile(row)

        # Create a new Word document (append to existing one)
        doc = docx.Document()

        # Add paragraph to the document
        doc.add_paragraph(profile_text)

        # Save the Word document with a unique filename
        client_name = row['ACCT_NAME']
        doc.save(f"{output}/{client_name}.docx")

        # Close the document to release resources
        #doc.close()


# Call the function to generate customer profiles
generate_customer_profiles("data1.xlsx", "output")
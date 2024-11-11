import os
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Set up Google Sheets and Google Docs API credentials
SERVICE_ACCOUNT_FILE = process.env.GCP_CLIENT_SECRET
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive.file'
]
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Configuration
SHEET_ID = '1lZtro8sLdpx8IozYQUbe-dJFhkacxw8W3RvZf45furo'
PACKAGE_META_TAB = 'Package_Meta'
PACKAGE_CONTENTS_TAB = 'Package_Contents'
TEMPLATE_DOC_ID = '19VGSJoRY2n1BmyQcqUBuI0aVEJd_1Uh2wHMDB-0FXi8'
OUTPUT_FOLDER_ID = '1VWF48TKNHBtAfIt1SF9ks1uzmXj02IQj'

def fetch_google_sheet_data(sheet_id, range_name):
    """Fetch data from Google Sheets."""
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    return pd.DataFrame(values[1:], columns=values[0]) if values else pd.DataFrame()

def get_package_data():
    """Fetch data from both tabs and merge them based on Package_Number."""
    # Fetch data from Google Sheets tabs
    package_meta = fetch_google_sheet_data(SHEET_ID, PACKAGE_META_TAB)
    package_contents = fetch_google_sheet_data(SHEET_ID, PACKAGE_CONTENTS_TAB)

    # Convert Package_Number to string for consistent merging
    package_meta['Package_Number'] = package_meta['Package_Number'].astype(str)
    package_contents['Package_Number'] = package_contents['Package_Number'].astype(str)

    # Merge data based on Package_Number
    merged_data = pd.merge(package_meta, package_contents, on='Package_Number', how='left')
    return merged_data

def copy_template(template_id, output_name, folder_id):
    """Create a copy of the Google Docs template for each invoice."""
    drive_service = build('drive', 'v3', credentials=creds)
    copied_file = drive_service.files().copy(
        fileId=template_id,
        body={'name': output_name, 'parents': [folder_id]}
    ).execute()
    return copied_file['id']

def replace_placeholders(document_id, replacements):
    """Replace placeholders in a Google Docs document."""
    docs_service = build('docs', 'v1', credentials=creds)
    requests = [
        {'replaceAllText': {
            'containsText': {'text': '{{' + key + '}}', 'matchCase': True},
            'replaceText': value
        }} for key, value in replacements.items()
    ]
    docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

def generate_invoices():
    """Main function to generate invoices based on merged data."""
    merged_data = get_package_data()

    # Group by Package_Number to create individual invoices
    for package_number, group in merged_data.groupby('Package_Number'):
        package_info = group.iloc[0]  # Get customer info for this package

        # Format items list with details from Package_Contents
        items = "\n".join([
            f"Card: {row['Card']}, Set: {row['Set']}, Cond.: {row['Cond.']}, Finish: {row['Finish']}, "
            f"Lang.: {row['Lang.']}, Amount: {row['Amount']}, Status: {row['Status']}, Date: {row['Date']}"
            for _, row in group.iterrows()
        ])

        # Prepare replacements dictionary
        replacements = {
            'Customer_Name': package_info['Customer_Name'],
            'Customer_Email': package_info['Customer_Email'],
            'Package_Number': package_number,
            'Order_Date': package_info['Order_Date'],
            'Shipping_Address': package_info['Shipping_Address'],
            'Items': items,
            'Total_Amount': str(group['Amount'].astype(float).sum())
        }

        # Copy the template and replace placeholders
        document_id = copy_template(TEMPLATE_DOC_ID, f"Invoice_{package_number}", OUTPUT_FOLDER_ID)
        replace_placeholders(document_id, replacements)
        print(f"Invoice generated for Package {package_number}.")

if __name__ == '__main__':
    generate_invoices()

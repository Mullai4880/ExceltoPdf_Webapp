import os
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename
import msal
import requests
import xlwings as xw

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.abspath('uploads')
app.config['SECRET_KEY'] = os.urandom(24)  # Replace with your actual secret key

# Ensure the upload directory exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# MSAL Configuration
CLIENT_ID = 'your_client_id'
CLIENT_SECRET = 'your_client_secret'
AUTHORITY = 'https://login.microsoftonline.com/your_tenant_id'
REDIRECT_PATH = '/get_token'
SCOPES = ['https://graph.microsoft.com/.default']

# Global variable to temporarily store the PDF path
pdf_path_global = None

@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and convert to PDF."""
    global pdf_path_global  # Use the global variable

    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            print(f"Attempting to save file to {file_path}")
            file.save(file_path)
            print(f"File saved successfully to {file_path}")
            pdf_path = export_to_pdf(file_path)
            if pdf_path:
                pdf_path_global = pdf_path  # Store the PDF path in the global variable
                return redirect(url_for('login'))
            else:
                return 'Error generating PDF.'
        except Exception as e:
            print(f"An error occurred while saving the file: {e}")
            return f"An error occurred: {e}"

def export_to_pdf(excel_path):
    """Convert the uploaded Excel file to a PDF."""
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], "Combined_Report.pdf")
    try:
        xl_app = xw.App(visible=False)
        book = xl_app.books.open(excel_path)
        book.api.ExportAsFixedFormat(0, pdf_path)
        book.close()
        xl_app.quit()
        print(f"Exported entire workbook to {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"Error exporting workbook to PDF: {e}")
        return None

@app.route('/login')
def login():
    """Redirect to Microsoft login for authentication."""
    auth_url = f'{AUTHORITY}/oauth2/v2.0/authorize?client_id={CLIENT_ID}&response_type=code&redirect_uri=http://localhost:5000/get_token&response_mode=query&scope={SCOPES[0]}&state=12345'
    return redirect(auth_url)

@app.route('/get_token')
def get_token():
    """Handle token retrieval after login and upload the PDF to OneDrive and SharePoint."""
    global pdf_path_global  # Use the global variable

    code = request.args.get('code')
    print(f"Retrieved PDF path from global variable: {pdf_path_global}")

    if not code:
        return 'No authorization code found.'
    
    if not pdf_path_global:
        print("PDF path not found in global variable.")
        return 'PDF path not found.'

    # Create a ConfidentialClientApplication for acquiring tokens
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri='http://localhost:5000/get_token')
    
    if "access_token" in result:
        access_token = result['access_token']
        upload_to_onedrive(pdf_path_global, access_token)
        upload_to_sharepoint(pdf_path_global, access_token)  # Upload to SharePoint
        pdf_path_global = None  # Clear the global variable
        return 'File uploaded successfully to OneDrive and SharePoint'
    else:
        print(f"Failed to obtain access token: {result.get('error_description')}")
        return 'Failed to obtain access token.'

def upload_to_onedrive(file_path, access_token):
    """Upload the PDF to OneDrive."""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/pdf'
    }
    file_name = os.path.basename(file_path)
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/Uploads/{file_name}:/content"
    
    with open(file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
        
    if response.status_code == 201:
        print(f"File {file_name} uploaded successfully to OneDrive")
    else:
        print(f"Failed to upload file {file_name}: {response.status_code}, {response.text}")

def upload_to_sharepoint(file_path, access_token):
    """Upload the PDF to SharePoint."""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/pdf'
    }
    file_name = os.path.basename(file_path)
    site_id = 'your_sharepoint_site_id'  # Replace with your SharePoint site ID
    document_library_id = 'your_document_library_id'  # Replace with your document library ID
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{document_library_id}/root:/Uploads/{file_name}:/content"
    
    with open(file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file)
        
    if response.status_code == 201:
        print(f"File {file_name} uploaded successfully to SharePoint")
    else:
        print(f"Failed to upload file {file_name} to SharePoint: {response.status_code}, {response.text}")

if __name__ == '__main__':
    app.run(debug=True)

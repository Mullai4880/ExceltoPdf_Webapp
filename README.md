# Excel to PDF Converter

## Overview
This web application allows users to upload Excel files and convert them to PDF format. The generated PDFs are then uploaded to both OneDrive and SharePoint. This project aims to replace the current Excel-based report generation process with a streamlined web application to enhance efficiency and accessibility.

## Features
- **Upload Excel files**: Users can upload Excel files through a simple web interface.
- **Convert Excel to PDF**: The application converts the uploaded Excel file into a PDF document.
- **Upload PDFs to OneDrive and SharePoint**: The generated PDFs are uploaded to both OneDrive and SharePoint for easy access and sharing.
- **User Authentication with MSAL**: The application uses Microsoft Authentication Library (MSAL) for secure user authentication.

## Setup Instructions

### Prerequisites
- Python 3.x
- Flask
- xlwings
- MSAL
- Requests

### Installation
1. **Clone the repository**:
    ```sh
    git clone https://github.com/Mullai4880/ExceltoPdf_Webapp.git
    cd ExceltoPdf_Webapp
    ```

2. **Install dependencies**:
    ```sh
    pip install -r requirements.txt
    ```

3. **Configure the application**:
    - Update `app.py` with your OneDrive and SharePoint details including `CLIENT_ID`, `CLIENT_SECRET`, `AUTHORITY`, `site_id`, and `document_library_id`.

4. **Run the application**:
    ```sh
    flask run
    ```

## Usage
1. **Open your web browser** and go to `http://localhost:5000`.
2. **Upload an Excel file** by clicking on the "Choose File" button and selecting an Excel file from your computer.
3. **Click "Convert to PDF"** to start the conversion process.
4. **Authenticate with your Microsoft account** when prompted to allow the application to upload the PDF to OneDrive and SharePoint.

## Screenshots
Here are some screenshots to give you an idea of the application interface:

### Home Page
![Home Page](https://i.imgur.com/wAKGi1T.png)
*Screenshot of the home page where users can upload their Excel files.*

### File Upload Form
![File Upload Form](https://i.imgur.com/XmpmaYC.png)
*Screenshot of the file upload form.*

### Conversion Result
![Conversion Result](https://i.imgur.com/fXrzXaJ.png)
*Screenshot of the confirmation message after converting the Excel file to PDF.*

### OneDrive and SharePoint Upload Confirmation
![Upload Confirmation](https://i.imgur.com/w1GxFVn.png)
*Screenshot showing the confirmation that the PDF has been uploaded to OneDrive and SharePoint.*

## Demo Video
Watch the demo video to see the application in action:

[Demo Video](https://i.imgur.com/dmptiwr.mp4)

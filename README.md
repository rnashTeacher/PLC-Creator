# Student Report PDF Generator

A browser-based tool to generate personalized PDF reports for students directly from an Excel spreadsheet. This application allows educators to quickly create detailed feedback documents, including topic-by-topic analysis and QR-code-linked resources, without needing any server-side processing.

![Screenshot of the application UI](screenshot.png) <!-- You should create a screenshot of your app and name it screenshot.png -->

## Features

- **Excel File Parsing**: Directly loads and parses `.xls` and `.xlsx` files in the browser.
- **Interactive Configuration**: No hardcoded values. Users can interactively select all important data points directly from the displayed spreadsheet table:
  - Sheet selection for multi-sheet workbooks.
  - Topic header row.
  - First student row.
  - Row containing maximum scores for each topic.
  - Row containing guidance text/links.
  - Columns for student grades and the start/end of topics.
- **Dynamic PDF Generation**: Creates a unique, multi-page PDF for each student.
- **Customizable PDF Header**: Add a custom logo to the header of each generated PDF.
- **Topic Analysis**: Automatically categorizes topics into "Very Insecure," "Insecure," and "Secure" based on student scores relative to the maximum possible score for each topic.
- **Targeted Guidance**: For topics identified as "Very Insecure," the tool generates a second table that includes:
  - The name of the topic.
  - A scannable QR code that links to a corresponding URL for additional resources or guidance.
- **Smart Layout**:
  - The guidance table is automatically split into chunks of 5 rows to prevent awkward page breaks.
  - Column widths are adjusted to ensure readability.
  - QR codes are rendered at a fixed, scannable size.

## How to Use

1.  **Open the Application**: Open the `sheets-viewer.html` file in a modern web browser (like Chrome, Firefox, or Edge).

2.  **Load Files**:
    - **Step 1**: Click "Choose File" to select the Excel spreadsheet containing your student data.
    - **Step 2 (Optional)**: Click "Choose File" to select a logo image (`.png` or `.jpg`) to be placed in the header of each PDF.

3.  **Select a Sheet**:
    - **Step 3**: If your workbook has multiple sheets, a dropdown will appear. Select the sheet that contains the student data you wish to process. The data will be displayed in a table on the page.

4.  **Configure PDF Settings**:
    - **Step 4**: A "PDF Generation Settings" panel will appear. For each setting:
        - Click the corresponding "Select" button.
        - An alert will confirm you are in selection mode.
        - Click the appropriate cell, row, or column in the data table below as prompted.
    - You must configure all settings:
        - **Topic Header Row**: The row with the full topic names (e.g., "Q1 - 1.1...").
        - **First Student Row**: The first row containing student data.
        - **Guidance Row**: The row containing the URLs that will be converted into QR codes.
        - **Max Score Row**: The row containing the maximum possible score for each topic.
        - **Grade Column**: The column containing the students' overall grades.
        - **First/Last Topic Column**: The columns that mark the beginning and end of your topic data.

5.  **Generate PDFs**:
    - Once all settings are configured, click the **"Generate PDFs for Students"** button.
    - The application will process the data and your browser will begin downloading a separate PDF file for each student. You may need to grant your browser permission to download multiple files.

## Technical Stack

- **HTML/CSS/JavaScript**: The core front-end technologies.
- **SheetJS (xlsx.js)**: For parsing Excel files directly in the browser.
- **jsPDF**: For creating PDF documents on the client-side.
- **jsPDF-AutoTable**: A plugin for jsPDF to easily generate tables.
- **qrcode-generator**: For generating QR codes from URLs.

## Setup for Local Development

No special setup is required. Simply open the `sheets-viewer.html` file in a web browser. For debugging purposes, it can be helpful to run a local web server to avoid potential browser security restrictions with local files.

---
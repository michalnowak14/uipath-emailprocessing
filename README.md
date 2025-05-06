🧾 Automated Invoice Processing & Email Reporting – UiPath Project
This UiPath project automates the process of retrieving invoice PDFs from Outlook emails, extracting relevant data, compiling it into a structured format, and emailing a summary report. It is designed to reduce manual effort in invoice handling and ensure accuracy in reporting.

📌 Features
Fetches emails with PDF attachments from Outlook

Saves PDF attachments locally

Reads and parses invoice data from PDFs

Builds a structured DataTable of invoice information

Writes the data to an Excel file (optional)

Generates and sends a summary email with total invoice count and total amount

🗂️ Project Structure
graphql
Kopiuj
Edytuj
Main.xaml
│
├── Sequences
│   ├── ProcessEmailAttachments.xaml         # Downloads PDF attachments from emails
│   ├── ExtractPDFDataToDataTable.xaml       # Parses PDF content and stores data in a table
│   ├── WriteInvoicesToExcel.xaml            # Writes DataTable to Excel (optional)
│   └── PrepareAndSendInvoiceSummaryEmail.xaml # Prepares and sends a summary report via Outlook
🔄 Workflow Overview
Retrieve Outlook Emails

Filters emails in a specified folder.

Extracts only those with PDF attachments.

Save PDF Attachments

Downloads attachments to a local folder for processing.

Read PDF Content

Uses Read PDF Text to extract invoice data.

Parses key information (e.g., invoice number, total amount).

Build Data Table

Adds rows for each PDF into a structured table.

Can be exported to Excel if needed.

Prepare Email Summary

Calculates:

Total number of invoices

Combined total amount

Composes an email body with this data.

Send Summary Email

Sends a formatted email via Outlook to a predefined recipient.

🛠️ Technologies Used
UiPath Studio

Outlook Desktop Integration

PDF Activities

Excel Activities (optional)

VB.NET for expressions and logic

✅ Requirements
UiPath Studio (tested on 2023.X or later)

Microsoft Outlook (desktop client)

PDF files must contain readable text (not image-based PDFs unless OCR is added)

Excel (for optional export)

📨 Example Email Output
Subject: Invoice Summary Report
Body:

Total Invoices Processed: 5

Total Amount: $8,230.50

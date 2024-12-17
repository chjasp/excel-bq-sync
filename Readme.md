# Excel BigQuery Integration

A VBA-based solution for bidirectional data synchronization between Microsoft Excel and Google BigQuery, supported by a Python Cloud Function for authentication.

## Overview

This project enables users to:
1. Query data from BigQuery and populate Excel sheets
2. Upload Excel data back to BigQuery tables
3. Securely authenticate with Google Cloud using service accounts

## Components

### 1. Excel VBA Module
The main VBA module (`bq-sync.vba`) provides functionality to:
- Execute BigQuery queries and populate results in a "Data" sheet
- Upload data from Excel to BigQuery using the insertAll API
- Handle JSON parsing and HTTP requests
- Manage authentication tokens

### 2. Cloud Function
A Python-based Google Cloud Function that:
- Handles service account authentication
- Generates JWT tokens for BigQuery API access
- Provides a secure endpoint for token generation

## Setup

1. Create a Google Cloud project and enable the BigQuery API
2. Create a service account and download the credentials
3. Deploy the cloud function from the `02-cloud-function` directory
4. Import the VBA module into your Excel workbook
5. Configure the "Controls" sheet with:
   - Service account credentials (Cell B1)
   - Project ID (Cell B2)
   - Dataset name (Cell B3)
   - Table name (Cell B4)
   - SQL Query (Cell B5) - for data retrieval

## Usage

### Querying Data
1. Enter your BigQuery SQL query in the Controls sheet (Cell B5)
2. Run the `ExecuteBigQueryAndPopulateSheet` subroutine
3. Results will appear in the "Data" sheet

### Uploading Data
1. Prepare your data in the "Data" sheet with appropriate column headers
2. Run the `UploadDataToBigQuery` subroutine
3. Data will be uploaded to the specified BigQuery table

## Dependencies

### Cloud Function
- Flask
- Google Auth
- Werkzeug
- Google Auth OAuth Library

### Excel
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) - A JSON parser for VBA
- MacOS. Shell commands are tailored for MacOS. Windows users will have to use a different method to send the HTTP request.

## Security Notes

- Service account credentials should be handled securely
- The cloud function should be deployed with appropriate IAM permissions
- Access tokens are generated per session and handled securely in memory

## Error Handling

The solution includes comprehensive error handling for:
- Network requests
- JSON parsing
- File operations
- BigQuery API interactions

## Limitations

- Currently designed for MacOS (uses MacScript for HTTP requests)
- Requires manual configuration of cloud function URL
- Maximum data transfer limits based on BigQuery quotas
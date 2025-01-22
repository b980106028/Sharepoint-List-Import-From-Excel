# Excel to SharePoint Import Tool

A PowerShell-based tool for importing data from Excel files to SharePoint lists with advanced features like progress tracking, error handling, and import status management.

## Features

### Core Functionality
- Imports data from Excel files to SharePoint lists
- Maps Excel column headers to SharePoint internal column names
- Tracks import progress with detailed logging
- Prevents duplicate imports using ImportStatus tracking
- Checks for open Excel files before processing

### Safety Features
- Validates Excel file existence and accessibility
- Confirms column mapping with user before processing
- Handles errors gracefully with detailed error messages
- Creates backup of original Excel file before modifications
- Uses web login for secure SharePoint authentication

### Progress Tracking
- Maintains a detailed progress log in `Progress.MD`
- Records timestamps for all operations
- Tracks successful and failed imports
- Shows detailed error messages when issues occur

## Prerequisites

### Required PowerShell Modules
- PnP.PowerShell: For SharePoint operations
- ImportExcel: For Excel file handling

### Required Files
- `.env`: Configuration file for environment variables
- Source Excel file with data to import
- SharePoint list with matching columns

## Installation

1. Clone this repository
2. Copy `.env.example` to `.env`
3. Configure the following variables in `.env`:
   ```
   SHAREPOINT_URL=https://your-sharepoint-site-url
   LIST_NAME=Your_SharePoint_List_Name
   EXCEL_FILE=source_excel.xlsx
   EXCEL_FILE_UPDATED=updated_excel.xlsx
   ```
4. Ensure required PowerShell modules are installed (script will handle this automatically)

## Usage

1. Close any open Excel files that will be processed
2. Run the script:
   ```powershell
   .\Import-ExcelToSharePoint.ps1
   ```
3. Follow the interactive prompts:
   - Confirm SharePoint connection
   - Review column mapping
   - Monitor import progress

## Excel File Requirements

### Source Excel File
- Must have headers in the first row
- Headers should match SharePoint column names (will be mapped automatically)
- File should not be open in Excel during import

### Updated Excel File
- Created automatically by the script
- Contains original data plus ImportStatus column
- ImportStatus values:
  - 0: Row not yet imported
  - 1: Row successfully imported

## Error Handling

The script handles various error scenarios:
- Missing Excel files
- Open Excel files (waits for user to close)
- SharePoint connection issues
- Column mapping mismatches
- Data validation errors

## Logging

All operations are logged in `Progress.MD` with timestamps:
- Script start/end
- SharePoint connection status
- Excel file operations
- Row-by-row import status
- Error messages

## Security

- Uses interactive web login for SharePoint authentication
- No stored credentials
- Original Excel file is preserved
- Sensitive configuration in `.env` file (gitignored)

## Files

### Core Files
- `Import-ExcelToSharePoint.ps1`: Main script
- `.env`: Configuration file
- `Progress.MD`: Progress log

### Generated Files
- `tobeimported_excel_last.xlsx`: Updated Excel file with ImportStatus

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

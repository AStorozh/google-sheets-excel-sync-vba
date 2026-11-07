# Google Sheets Excel Sync VBA

VBA automation for syncing data from Google Sheets into an Excel report.

The macro downloads a Google Sheets document as an `.xlsx` file, opens it in Excel, copies selected sheets into target Excel sheets, keeps formatting, enables filters and fixes barcode columns displayed in scientific notation.

## Features

- Sync data from Google Sheets to Excel
- Copy only selected sheets
- Map Google Sheets tabs to specific Excel tabs
- Preserve cell values
- Preserve cell colors and formatting
- Preserve column widths
- Preserve row heights
- Enable filters on imported tables
- Fix barcode columns displayed as scientific notation
- Works on Windows and macOS

## Use Case

This project is useful when a business report is stored in Excel, but the source data is updated in Google Sheets.

Instead of manually copying data between files, the user can press one button in Excel and refresh all required sheets automatically.

Example workflow:

```text
Google Sheets updates data
↓
User opens Excel report
↓
User clicks Refresh button
↓
Excel downloads fresh data
↓
Required sheets are updated automatically
```

## Project Structure

```text
google-sheets-excel-sync-vba/
├── README.md
└── google_sheets_excel_sync.bas
```

## How It Works

The macro:

1. Downloads a Google Sheets file as `.xlsx`
2. Opens the downloaded file as a temporary workbook
3. Copies selected sheets from the source workbook
4. Pastes data into selected sheets in the target Excel workbook
5. Copies formatting, column widths and row heights
6. Enables filters
7. Formats barcode columns correctly
8. Closes the temporary workbook

## Setup

### 1. Export Google Sheets as XLSX

Use this URL format:

```text
https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/export?format=xlsx
```

Replace:

```text
YOUR_SPREADSHEET_ID
```

with your real Google Sheets document ID.

Example:

```vba
googleXlsxUrl = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/export?format=xlsx"
```

Important: the Google Sheets file must be accessible by link or published with proper permissions.

## Sheet Mapping

Update the sheet mapping section in the VBA code:

```vba
CopySheetData sourceWb, "Source Sheet 1", ThisWorkbook, "Target Sheet 1"
CopySheetData sourceWb, "Source Sheet 2", ThisWorkbook, "Target Sheet 2"
CopySheetData sourceWb, "Source Sheet 3", ThisWorkbook, "Target Sheet 3"
CopySheetData sourceWb, "Source Sheet 4", ThisWorkbook, "Target Sheet 4"
CopySheetData sourceWb, "Source Sheet 5", ThisWorkbook, "Target Sheet 5"
```

Where:

```text
Source Sheet
```

is the sheet name in Google Sheets.

```text
Target Sheet
```

is the sheet name in the Excel report.

## How to Add the Macro to Excel

1. Open your Excel file.
2. Save it as `.xlsm`.
3. Open the VBA editor:
   - Windows: `Alt + F11`
   - macOS: `Option + F11`
4. Insert a new module:
   - `Insert → Module`
5. Paste the code from `google_sheets_excel_sync.bas`.
6. Save the file.

## How to Add a Refresh Button

You can add a button directly inside the Excel sheet:

1. Open the `Developer` tab.
2. Click `Insert`.
3. Choose `Button`.
4. Draw the button on the sheet.
5. Assign the macro:

```text
Refresh_Data_From_Google_Sheets
```

6. Rename the button to:

```text
Refresh
```

Now the report can be updated by clicking one button.

## Windows Notes

On Windows, the macro uses:

```vba
URLDownloadToFile
```

to download the Google Sheets `.xlsx` file.

If macros are blocked:

1. Close the Excel file.
2. Right-click the file.
3. Open `Properties`.
4. Click or check `Unblock`.
5. Open the file again.
6. Enable macros.

## macOS Notes

On macOS, the macro uses:

```text
curl
```

through `MacScript` to download the Google Sheets `.xlsx` file.

## Barcode Formatting

Excel often displays long barcode values like this:

```text
2.04347E+12
```

The macro detects columns named:

```text
barcode
bar code
```

(the Russian spelling of "barcode" is matched too)

and applies number formatting to display them correctly.

## Important Security Notes

Do not publish:

- Private Google Sheets links
- Real company data
- Client data
- Financial reports
- Product analytics
- Internal sheet names if they reveal business logic

This repository contains a reusable template only.

## Technologies

- VBA
- Microsoft Excel
- Google Sheets export
- Windows API `URLDownloadToFile`
- macOS `curl`

## Example Result

Before automation:

```text
Manual copy-paste from Google Sheets to Excel
Formatting breaks
Filters need to be recreated
Barcode values display incorrectly
```

After automation:

```text
One-click Excel update
Formatting preserved
Filters enabled
Barcode values displayed correctly
```

## License

MIT License

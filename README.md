# Excel Import Service

A C# ASP.NET Framework 4.8 Web API application that extracts data from Excel files and returns JSON. This project is designed for IIS hosting (System.Web / Web API 5.x).

## Project Structure

```
ImportExcelApp/
├── Models/
│   ├── ExcelImportRequest.cs      # Input model for API requests
│   └── ExcelImportResponse.cs     # Output model for API responses
├── Services/
│   └── ExcelService.cs            # Business logic for Excel parsing
├── Properties/
│   └── launchSettings.json        # Launch configuration
├── Program.cs                      # Application entry point and API endpoints
├── ImportExcelApp.csproj          # Project configuration
├── ExcelImport.http               # Sample HTTP requests for testing
└── appsettings.json               # Application settings
```

## Features

- ✅ Extract data from Excel files (.xlsx and .xls)
- ✅ Support for custom sheet selection via the request body
- ✅ Only columns matching model field names (template or line items) are retained
- ✅ Automatic header detection
- ✅ Returns data as JSON with descriptive response
- ✅ Comprehensive error handling
- ✅ Logging support

## Dependencies

- **EPPlus 7.4.1** - Excel file reading library
- **.NET Framework 4.8** - Target framework
- **ASP.NET Web API 5.x** - Web API framework for IIS hosting

## API Endpoint

### POST /api/excel/import

**Description:** Import Excel file data and return as JSON via built-in HTTP listener. You can optionally supply a `sheetName` in the request to choose which sheet to import; otherwise the first sheet is used.

**Request Headers:**
```
Content-Type: application/json
```

**Request Body:**
```json
{
  "filePath": "C:\\path\\to\\file.xlsx",
  "hasHeaders": true
}
```

**Request Parameters:**
- `filePath` (required, string): Local file path to the Excel file
- `sheetName` (optional, string): Name of the sheet to import. If not specified, the first sheet is used
- `hasHeaders` (optional, boolean): Whether to treat the first row as headers. Default: true

The service scans the *entire sheet* for cells whose values exactly match one of the template field names (Role, Requestor, TypeofPurchase, PurchasingGroup, PurchaseCategory, SpeedDivision, Vendor, VendorCurrency, ParentIC, BusinessUnit, DirectIndirect, Region, LevelCode, CustomerCode).
When a matching cell is found, the service attempts to read a value in one of two places:
1. the text immediately following the label within the same cell (e.g. "Purchasing Group Z13"), or
2. the adjacent cell to the right on the same row.
All discovered values are merged into a **single object**; the response always contains at most one element in the `TemplateRecords` array.

If the sheet contains a tabular section whose header row includes **"Item Number"**, the rows below that header are treated as line items.  Each column name is matched (ignoring spaces/punctuation) to a property on `CreateICLineItem`; any filled cells are copied into the corresponding object.  The `LineItems` array in the response contains one object per non‑empty row under that section.
**Response (Success - 200 OK):**
```json
{
  "success": true,
  "message": "Successfully imported template data from 'Sheet1' sheet",
  "sheetName": "Sheet1",
  "TemplateRecords": [
    {
      "Role": "John",
      "Requestor": "john@example.com",
      "Vendor": "Contoso",
      "Region": "West"
    }
  ],
  "LineItems": [
    {
      "ItemNumber": "1001",
      "MaterialCode": "ABC",
      "ShortText": "sample",
      "Quantity": "10",
      "Price": "100"
    }
  ]
}
```

**Response (Error - 400 Bad Request):**
```json
{
  "success": false,
  "message": "File not found: C:\\invalid\\path.xlsx",
  "sheetName": null,
  "TemplateRecords": []
}
```

## Getting Started

### Prerequisites
- .NET Framework 4.8 developer targeting pack (can build with Visual Studio 2019/2022)
- Visual Studio (for IIS Express or publish deployment)


### Installation

1. **Build the project** (using Visual Studio or `msbuild` or `dotnet build` targeting `net48`).
   ```bash
   dotnet build -f net48
   ```

2. **Publish for IIS deployment:**
   ```bash
   dotnet publish -c Release -f net48 -o publish
   ```

3. **Deploy to IIS** (recommended):
   - create website/app in IIS Manager
   - physical path -> `...\publish`
   - app pool .NET CLR v4.0, Integrated pipeline
   - verify `web.config` is present in published output

4. **Run with IIS Express (local dev)**:
   - use Visual Studio F5 with the Web API profile

> `dotnet run -f net48` is not applicable for IIS-hosted `System.Web` Web API projects.

### Usage

#### Example with sheet selection

```json
{
  "filePath": "C:\\Users\\user\\Documents\\data.xlsx",
  "sheetName": "Customers",
  "hasHeaders": true
}
```


#### Option 1: Using REST Client Extension (VS Code)

1. Open `ExcelImport.http` file
2. Update the file path in the request body
3. Click "Send Request" above the request

#### Option 2: Using cURL

```bash
curl -X POST https://localhost:7180/api/excel/import \
  -H "Content-Type: application/json" \
  -d "{\"filePath\":\"C:\\path\\to\\file.xlsx\",\"hasHeaders\":true}"
```

#### Option 3: Using Postman

1. Create a new POST request
2. URL: `https://localhost:7180/api/excel/import`
3. Headers: `Content-Type: application/json`
4. Body (raw JSON):
```json
{
  "filePath": "C:\\path\\to\\file.xlsx",
  "hasHeaders": true
}
```


## Error Handling

The service handles various error scenarios:

| Error | Status Code | Message |
|-------|------------|---------|
| Missing file path | 400 | "File path is required" |
| File not found | 400 | "File not found: [path]" |
| Invalid file format | 400 | "Only .xlsx and .xls files are supported" |
| Sheet not found | 400 | "Sheet '[name]' not found in the Excel file" |
| General error | 400 | "Error processing Excel file: [details]" |

## Code Structure

### Models

**ExcelImportRequest:**
- `FilePath`: Path to the Excel file
- `SheetName`: Optional sheet name to import; defaults to first sheet if omitted
- `HasHeaders`: Whether first row contains headers

**ExcelImportResponse:**
- `Success`: Operation success status
- `Message`: Status message
- `SheetName`: Name of the sheet that was imported
- `RowCount`: Number of template records found
- `Headers`: List of column headers
- `TemplateRecords`: Array containing a single merged template field object
- `LineItems`: (optional) list of line-item records extracted from the sheet

### Services

**ExcelService:**
- Implements `IExcelService` interface
- Uses EPPlus to read Excel files
- Handles data extraction and validation
- Includes comprehensive logging

## API Documentation

Use standard API tools to test this endpoint:
- `POST /api/excel/import`

## Features & Enhancements

### Current Features
- ✅ Local file path input
- ✅ Single sheet import
- ✅ Header detection
- ✅ JSON response

### Potential Enhancements
- [ ] Multiple sheet import
- [ ] File upload via multipart/form-data
- [ ] Data validation and transformation
- [ ] Export to other formats (CSV, JSON file)
- [ ] Batch processing
- [ ] Data filtering and pagination
- [ ] Cell formatting preservation
- [ ] Formula evaluation

## Logging

The service logs important events to help with debugging:
- File access attempts
- Data import completion
- Errors and exceptions

Check the application console output for logs during execution.

## License

This project is provided as-is for development purposes.

## Support

For issues or questions, ensure that:
1. The Excel file path is correct and accessible
2. The file has a valid .xlsx or .xls extension
3. If specifying a sheet name, it exists in the workbook
4. The application has read permissions for the file

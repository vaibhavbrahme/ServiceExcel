using ImportExcelApp.Models;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;
using static System.Console;
using System.IO;
using System;
using System.Linq;


namespace ImportExcelApp.Services
{
    /// <summary>
    /// Service for handling Excel file operations
    /// </summary>
    public interface IExcelService
    {
        /// <summary>
        /// Imports data from an Excel file
        /// </summary>
        Task<ExcelImportResponse> ImportExcelAsync(ExcelImportRequest request);
    }

    /// <summary>
    /// Implementation of Excel service using OleDB
    /// </summary>
    public class ExcelService : IExcelService
    {
        // simple console logger instead of ILogger
        public ExcelService()
        {
        }

        /// <summary>
        /// Helper method to safely get cell value from DataRow by column name
        /// </summary>
        private string? GetCellValue(DataRow row, string columnName)
        {
            try
            {
                if (row.Table.Columns.Contains(columnName))
                {
                    var value = row[columnName];
                    return value == DBNull.Value ? null : value?.ToString();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Builds the OleDB connection string for Excel files
        /// </summary>
        private string BuildConnectionString(string filePath, bool hasHeaders)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            string connectionString;

            if (extension == ".xlsx")
            {
                // For Excel 2007+ (.xlsx)
                connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR={(hasHeaders ? "YES" : "NO")};'";
            }
            else if (extension == ".xls")
            {
                // For Excel 97-2003 (.xls)
                connectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR={(hasHeaders ? "YES" : "NO")};'";
            }
            else
            {
                throw new InvalidOperationException("Unsupported file format");
            }

            return connectionString;
        }

        /// <summary>
        /// Gets all sheet names from the Excel file
        /// </summary>
        private List<string> GetSheetNames(string connectionString)
        {
            var sheetNames = new List<string>();

            try
            {
                using (var connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    var schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (schemaTable != null)
                    {
                        foreach (DataRow row in schemaTable.Rows)
                        {
                            var tableName = row["TABLE_NAME"].ToString();
                            // only include worksheet entries (ending with $ or $')
                            if (tableName.EndsWith("$") || tableName.EndsWith("$'"))
                            {
                                // trim any surrounding whitespace, dollar sign, or quote characters
                                var cleanName = tableName.Trim(' ', '$', '\'');
                                sheetNames.Add(cleanName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting sheet names from Excel file: {ex.Message}");
            }

            return sheetNames;
        }

        /// <summary>
        /// Imports data from an Excel file and returns it as JSON
        /// </summary>
        public async Task<ExcelImportResponse> ImportExcelAsync(ExcelImportRequest request)
        {
            var response = new ExcelImportResponse();

            try
            {
                // Validate file path
                if (string.IsNullOrWhiteSpace(request.FilePath))
                {
                    response.Success = false;
                    response.Message = "File path is required";
                    return response;
                }

                // Check if file exists
                if (!File.Exists(request.FilePath))
                {
                    response.Success = false;
                    response.Message = $"File not found: {request.FilePath}";
                    Console.WriteLine($"Warning: Excel file not found: {request.FilePath}");
                    return response;
                }

                // Validate file extension
                var extension = Path.GetExtension(request.FilePath).ToLower();
                if (extension != ".xlsx" && extension != ".xls")
                {
                    response.Success = false;
                    response.Message = "Only .xlsx and .xls files are supported";
                    return response;
                }

                // Build connection string
                string connectionString = BuildConnectionString(request.FilePath, request.HasHeaders);

                // Get all sheet names
                var sheetNames = GetSheetNames(connectionString);

                if (sheetNames.Count == 0)
                {
                    response.Success = false;
                    response.Message = "Excel file contains no accessible sheets";
                    return response;
                }

                // if multiple sheets exist and caller didn't specify one, we can't pick arbitrarily
                if (sheetNames.Count > 1 && string.IsNullOrWhiteSpace(request.SheetName))
                {
                    response.Success = false;
                    response.Message = "Workbook contains multiple sheets; please specify the sheet name in the request.";
                    return response;
                }

                // Determine which sheet to read (only from request body)
                string sheetToRead;
                if (string.IsNullOrWhiteSpace(request.SheetName))
                {
                    sheetToRead = sheetNames[0];
                }
                else
                {
                    // normalize input by trimming whitespace and quotes
                    sheetToRead = request.SheetName!.Trim(' ', '\'');
                    if (!sheetNames.Exists(s => string.Equals(s, sheetToRead, StringComparison.OrdinalIgnoreCase)))
                    {
                        response.Success = false;
                        response.Message = $"Sheet '{sheetToRead}' not found in the Excel file. Available sheets: {string.Join(", ", sheetNames)}";
                        return response;
                    }
                }
                response.SheetName = sheetToRead;
                // Use OleDbDataAdapter to fill DataSet
                using (var connection = new OleDbConnection(connectionString))
                {
                    // Query the sheet - add $ to sheet name for OleDB
                    string query = $"SELECT * FROM [{sheetToRead}$]";
                    var adapter = new OleDbDataAdapter(query, connection);
                    var dataSet = new DataSet();

                    adapter.Fill(dataSet);

                    if (dataSet.Tables.Count == 0)
                    {
                        response.Success = false;
                        response.Message = "Unable to read data from the sheet";
                        return response;
                    }

                    var dataTable = dataSet.Tables[0];

                    // Handle empty sheet
                    if (dataTable.Rows.Count == 0)
                    {
                        response.Success = true;
                        response.Message = "Excel sheet is empty";
                        // removed ColumnCount
                        
    
                        return response;
                    }


                    // build mapping of normalized property name -> PropertyInfo
                    var propMap = new Dictionary<string, System.Reflection.PropertyInfo>(StringComparer.OrdinalIgnoreCase);
                    var templateProps = typeof(CreateICTemplateFields).GetProperties();
                    foreach (var prop in templateProps)
                    {
                        // normalize by removing spaces
                        var norm = prop.Name.Replace(" ", string.Empty).ToLowerInvariant();
                        propMap[norm] = prop;
                    }

                    // scan entire sheet for template field names; allow label=value in same cell or adjacent cell
                    var mergedRecord = new CreateICTemplateFields();
                    bool anyFound = false;
                    for (int r = 0; r < dataTable.Rows.Count; r++)
                    {
                        var row = dataTable.Rows[r];
                        for (int c = 0; c < dataTable.Columns.Count; c++)
                        {
                            var cellRaw = row[c]?.ToString();
                            var cell = cellRaw?.Trim();
                            if (string.IsNullOrEmpty(cell))
                                continue;

                            // normalize cell text (remove spaces and punctuation)
                            var normCell = new string(cell.Where(ch => !char.IsWhiteSpace(ch) && !char.IsPunctuation(ch)).ToArray()).ToLowerInvariant();

                            foreach (var kv in propMap)
                            {
                                var normProp = kv.Key;
                                if (normCell.Contains(normProp))
                                {
                                    // determine value: text after the prop label in the cell, if any
                                    string extracted = null;
                                    int idx = cell.ToLowerInvariant().IndexOf(kv.Value.Name.ToLowerInvariant());
                                    if (idx >= 0)
                                    {
                                        extracted = cell.Substring(idx + kv.Value.Name.Length).Trim(new char[] {':','-',' '});
                                    }
                                    if (string.IsNullOrEmpty(extracted) && c + 1 < dataTable.Columns.Count)
                                    {
                                        extracted = row[c + 1]?.ToString();
                                    }
                                    if (!string.IsNullOrEmpty(extracted))
                                    {
                                        kv.Value.SetValue(mergedRecord, extracted);
                                        anyFound = true;
                                    }
                                }
                            }
                        }
                    }
                    if (anyFound)
                    {
                        response.TemplateRecords.Add(mergedRecord);
                    }

                    // now attempt to locate a header row for line items and populate them
                    int headerRowIndex = -1;
                    // find a row containing something that looks like "Item Number" after normalization
                    for (int r = 0; r < dataTable.Rows.Count; r++)
                    {
                        for (int c = 0; c < dataTable.Columns.Count; c++)
                        {
                            var text = dataTable.Rows[r][c]?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(text)) continue;
                            var norm = new string(text.Where(ch => !char.IsWhiteSpace(ch) && !char.IsPunctuation(ch)).ToArray()).ToLowerInvariant();
                            if (norm.Contains("itemnumber"))
                            {
                                headerRowIndex = r;
                                break;
                            }
                        }
                        if (headerRowIndex >= 0) break;
                    }
                    if (headerRowIndex >= 0)
                    {
                        // build header->colIndex map
                        var liPropMap = new Dictionary<string, System.Reflection.PropertyInfo>(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in typeof(CreateICLineItem).GetProperties())
                        {
                            var norm = prop.Name.Replace(" ", string.Empty).ToLowerInvariant();
                            liPropMap[norm] = prop;
                        }

                        var headers = new List<string>();
                        for (int c = 0; c < dataTable.Columns.Count; c++)
                            headers.Add(dataTable.Rows[headerRowIndex][c]?.ToString()?.Trim() ?? string.Empty);

                        for (int r = headerRowIndex + 1; r < dataTable.Rows.Count; r++)
                        {
                            // stop when first cell empty
                            if (string.IsNullOrWhiteSpace(dataTable.Rows[r][0]?.ToString()))
                                break;
                            var lineItem = new CreateICLineItem();
                            bool liFound = false;
                            for (int c = 0; c < headers.Count; c++)
                        {
                            var hdr = headers[c];
                            if (string.IsNullOrEmpty(hdr)) continue;
                            var normHdr = new string(hdr.Where(ch => !char.IsWhiteSpace(ch) && !char.IsPunctuation(ch)).ToArray()).ToLowerInvariant();
                            System.Reflection.PropertyInfo prop = null;
                            if (!liPropMap.TryGetValue(normHdr, out prop))
                            {
                                // fuzzy search: either key contains hdr or hdr contains key
                                prop = liPropMap.FirstOrDefault(kvp => kvp.Key.Contains(normHdr) || normHdr.Contains(kvp.Key)).Value;
                            }
                            if (prop != null)
                            {
                                var value = dataTable.Rows[r][c]?.ToString()?.Trim();
                                if (!string.IsNullOrEmpty(value))
                                {
                                    prop.SetValue(lineItem, value);
                                    liFound = true;
                                }
                            }
                        }
                            if (liFound)
                                response.LineItems.Add(lineItem);
                        }
                    }

                    // template records already populated by scanning earlier
                    response.Success = true;
                    response.Message = $"Successfully imported template data from '{sheetToRead}' sheet";

                    Console.WriteLine($"Successfully imported template data from sheet '{sheetToRead}' in {request.FilePath}");
                }
            }
            catch (OleDbException ex)
            {
                response.Success = false;
                response.Message = $"Database error: {ex.Message}. This might occur if the required database provider (Microsoft.ACE.OLEDB or Jet.OLEDB) is not installed.";
                Console.WriteLine($"OleDB error processing Excel file: {request.FilePath} - {ex}");
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.Message = $"Error processing Excel file: {ex.Message}";
                Console.WriteLine($"Error processing Excel file: {request.FilePath} - {ex}");
            }

            return response;
        }
    }
}



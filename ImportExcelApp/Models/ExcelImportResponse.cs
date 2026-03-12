using System.Collections.Generic;

namespace ImportExcelApp.Models
{
    /// <summary>
    /// Response model for Excel data extraction
    /// Contains imported IC template and line items data
    /// </summary>
    public class ExcelImportResponse
    {
        /// <summary>
        /// Indicates if the operation was successful
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// Success or error message
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// Name of the imported sheet
        /// </summary>
        public string? SheetName { get; set; }


        /// <summary>
        /// IC Template Fields discovered anywhere in the sheet
        /// Each entry corresponds to a row that contained one or more
        /// template field names; values are taken from the adjacent cell.
        /// </summary>
        public List<CreateICTemplateFields> TemplateRecords { get; set; } = new();

        /// <summary>
        /// List of IC Line Items - detailed row-level information
        /// Each item represents a line in the purchase order
        /// </summary>
        public List<CreateICLineItem> LineItems { get; set; } = new();

    }
}


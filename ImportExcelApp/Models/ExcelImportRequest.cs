using System.Collections.Generic;

namespace ImportExcelApp.Models
{
    /// <summary>
    /// Request model for Excel file import
    /// </summary>
    public class ExcelImportRequest
    {
        /// <summary>
        /// Local file path to the Excel file
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Optional: Sheet name to import. If not specified, first sheet is used
        /// </summary>
        public string? SheetName { get; set; }

        /// <summary>
        /// Optional: Whether to treat first row as headers
        /// </summary>
        public bool HasHeaders { get; set; } = true;
    }
}

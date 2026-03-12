namespace ImportExcelApp.Models
{
    /// <summary>
    /// Model for IC (Intercompany) Template Fields
    /// Contains header/template level information
    /// </summary>
    public class CreateICTemplateFields
    {
        /// <summary>
        /// Role identifier
        /// </summary>
        public string? Role { get; set; }

        /// <summary>
        /// Type of Purchase (e.g., Service, Material, etc.)
        /// </summary>
        public string? TypeofPurchase { get; set; }

        /// <summary>
        /// Requestor name or ID
        /// </summary>
        public string? Requestor { get; set; }


        /// <summary>
        /// Purchasing Group code
        /// </summary>
        public string? PurchasingGroup { get; set; }

        /// <summary>
        /// Purchase Category
        /// </summary>
        public string? PurchaseCategory { get; set; }

        /// <summary>
        /// Speed Division identifier
        /// </summary>
        public string? SpeedDivision { get; set; }

        /// <summary>
        /// Vendor name or code
        /// </summary>
        public string? Vendor { get; set; }

        /// <summary>
        /// Vendor Currency code (e.g., USD, EUR, etc.)
        /// </summary>
        public string? VendorCurrency { get; set; }

        /// <summary>
        /// Parent IC reference
        /// </summary>
        public string? ParentIC { get; set; }

        /// <summary>
        /// Business Unit identifier
        /// </summary>
        public string? BusinessUnit { get; set; }

        /// <summary>
        /// Direct/Indirect designation
        /// </summary>
        public string? DirectIndirect { get; set; }

        /// <summary>
        /// Region identifier
        /// </summary>
        public string? Region { get; set; }

        /// <summary>
        /// Level Code
        /// </summary>
        public string? LevelCode { get; set; }

        /// <summary>
        /// Customer Code
        /// </summary>
        public string? CustomerCode { get; set; }
    }
}

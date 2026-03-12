namespace ImportExcelApp.Models
{
    /// <summary>
    /// Model for IC (Intercompany) Line Items
    /// Contains detailed information for each line item in a purchase order
    /// </summary>
    public class CreateICLineItem
    {
        /// <summary>
        /// Data Item Index - unique identifier for the line item
        /// </summary>
        public string? DataItemIndex { get; set; }

        /// <summary>
        /// Material Code - SKU or material identifier
        /// </summary>
        public string? MaterialCode { get; set; }

        /// <summary>
        /// Short Text - brief description of the material
        /// </summary>
        public string? ShortText { get; set; }

        /// <summary>
        /// Cost Center code
        /// </summary>
        public string? CostCenter { get; set; }

        /// <summary>
        /// Quantity ordered
        /// </summary>
        public string? Quantity { get; set; }

        /// <summary>
        /// Unit Price
        /// </summary>
        public string? Price { get; set; }

        /// <summary>
        /// Net Price (total for the line item)
        /// </summary>
        public string? NetPrice { get; set; }

        // additional fields from sample:
        public string? ItemNumber { get; set; }
        public string? Plant { get; set; }
        public string? DeliveryDate { get; set; }
        public string? InternalOrder { get; set; }
        public string? HSNSACCode { get; set; }
        public string? SNSACDescription { get; set; }
        public string? POQuantity { get; set; }
        public string? OrderingUnit { get; set; }
        public string? MaterialGroup { get; set; }
        public string? TaxCode { get; set; }
        public string? Budget { get; set; }
        public string? ReferencePrice { get; set; }
        public string? BusinessArea { get; set; }
        public string? ControllingArea { get; set; }
        public string? ProfitCenter { get; set; }
        public string? Signature { get; set; }
        public string? BrandCode { get; set; }
        public string? SubBrand { get; set; }
        public string? Reference { get; set; }
        public string? GLCode { get; set; }
        public string? GLDescription { get; set; }
    }
}

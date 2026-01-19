namespace Core.Entities
{
    public class PurchaseOrderItem
    {
        public string PoNumber { get; set; } = "PO";

        public string VendorName {  get; set; } = string.Empty;

        public string InvoiceNumber {  get; set; } = string.Empty;

        public string Incoterm {  get; set; } = string.Empty;

        public int LineNumber { get; set; }

        public string PartNumber {  get; set; } = string.Empty;

        public string Description {  get; set; } = string.Empty;

        public decimal Quantity { get; set; }

        public string Unit {  get; set; } = string.Empty;

        public decimal UnitPrice { get; set; }

        public decimal TotalPrice { get; set; }

        public string SourceFileName { get; set; } = string.Empty;        
    }
}

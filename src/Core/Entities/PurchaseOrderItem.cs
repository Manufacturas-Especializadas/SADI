namespace Core.Entities
{
    public class PurchaseOrderItem
    {
        public int PoNumber { get; set; }
        
        public string InvoiceNumber { get; set; } = string.Empty;

        public string VendorName { get; set; } = string.Empty;

        public string PartNumber {  get; set; } = string.Empty;

        public string SourceFileName {  get; set; } = string.Empty;

        public decimal QtyPoPz { get; set; }

        public decimal QtyPoKg { get; set; }

        public decimal QtyInvPz { get; set; }

        public decimal QtyInvKg { get; set; }

        public int LineNumber { get; set; }

        public decimal TotalPrice { get; set; }

        public decimal UnitPrice { get; set; }
    }
}

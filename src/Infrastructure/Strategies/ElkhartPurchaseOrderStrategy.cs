using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Infrastructure.Strategies
{
    public class ElkhartPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "02 - Elkhart";

        public string DocumentTypeSubFolder => "02 - Factura";

        private class PoLineData
        {
            public int LineNumber { get; set; }

            public string Part {  get; set; }

            public decimal Qty { get; set; }

            public string Unit { get; set; }

            public decimal UnitPrice { get; set; }

            public decimal Amount { get; set; }
        }

        public List<PurchaseOrderItem> Extract(string invoiceFilePath)
        {
            var items = new List<PurchaseOrderItem>();
            string fileName = Path.GetFileName(invoiceFilePath);

            int poNumber = 0;
            string invoiceNum = "UNKNOWN";
            string vendorName = "ETI, LLC";

            using (var pdfInv = PdfDocument.Open(invoiceFilePath))
            {
                var page1 = pdfInv.GetPage(1);

                invoiceNum = ExtractInvoiceNumber(page1);
                poNumber = ExtractPoNumberFromInvoice(page1);

                foreach (var page in pdfInv.GetPages())
                {
                    var invoiceLines = ExtractInvoiceLines(page);

                    foreach (var line in invoiceLines)
                    {
                        var newItem = new PurchaseOrderItem
                        {
                            PoNumber = poNumber,
                            VendorName = vendorName,
                            InvoiceNumber = invoiceNum,
                            PartNumber = line.Part,
                            SourceFileName = fileName,
                            QtyInvPz = line.Qty,
                            QtyPoPz = line.Qty,
                            UnitPrice = line.UnitPrice,
                            TotalPrice = line.Amount
                        };

                        items.Add(newItem);
                    }
                }
            }

            return items;
        }

        private List<PoLineData> ExtractInvoiceLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();

            var custPartHeader = words.FirstOrDefault(w => w.Text.Contains("CUSTOMER") &&
                                                           words.Any(n => n.Text.Contains("NUMBER") && Math.Abs(n.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            var qtyHeader = words.FirstOrDefault(w => w.Text.Contains("QUANTITY") || w.Text.Contains("SHIPPED"));
            var unitPriceHeader = words.FirstOrDefault(w => w.Text.Contains("UNIT") && words.Any(p => p.Text.Contains("PRICE")));
            var extPriceHeader = words.FirstOrDefault(w => w.Text.Contains("EXTENDED"));

            if (custPartHeader == null || unitPriceHeader == null) return list;

            double tableTopY = custPartHeader.BoundingBox.Bottom;

            var partCandidates = words.Where(w =>
                w.BoundingBox.Top < tableTopY &&
                w.Text.Length > 3 &&
                Regex.IsMatch(w.Text, @"^[A-Za-z]") &&
                Regex.IsMatch(w.Text, @"\d") &&
                w.BoundingBox.Left >= custPartHeader.BoundingBox.Left - 50 &&
                w.BoundingBox.Right <= custPartHeader.BoundingBox.Right + 50
            ).ToList();

            foreach (var partWord in partCandidates)
            {
                var line = new PoLineData { Part = partWord.Text };
                double rowY = partWord.BoundingBox.Centroid.Y;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 10).ToList();

                if (qtyHeader != null)
                {
                    var qtyWord = rowWords.FirstOrDefault(w =>
                        w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 40 &&
                        w.BoundingBox.Right <= qtyHeader.BoundingBox.Right + 40 &&
                        Regex.IsMatch(w.Text, @"\d"));

                    if (qtyWord != null)
                    {
                        if (decimal.TryParse(qtyWord.Text.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q)) line.Qty = q;
                    }
                }

                var priceWord = rowWords.FirstOrDefault(w =>
                    w.BoundingBox.Left >= unitPriceHeader.BoundingBox.Left - 30 &&
                    w.BoundingBox.Right <= unitPriceHeader.BoundingBox.Right + 30 &&
                    Regex.IsMatch(w.Text, @"\d"));

                if (priceWord != null)
                {
                    if (decimal.TryParse(priceWord.Text.Replace("$", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal p)) line.UnitPrice = p;
                }

                if (extPriceHeader != null)
                {
                    var amtWord = rowWords.FirstOrDefault(w =>
                        w.BoundingBox.Left >= extPriceHeader.BoundingBox.Left - 30 &&
                        Regex.IsMatch(w.Text, @"\d"));

                    if (amtWord != null)
                    {
                        if (decimal.TryParse(amtWord.Text.Replace("$", "").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal a)) line.Amount = a;
                    }
                }

                list.Add(line);
            }

            return list;
        }

        private string ExtractInvoiceNumber(Page page)
        {
            var header = page.GetWords().FirstOrDefault(w => w.Text.Contains("INVOICE") &&
                         page.GetWords().Any(x => x.Text.Contains("NO") && Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            if (header != null)
            {
                var number = page.GetWords().FirstOrDefault(w =>
                    w.BoundingBox.Top < header.BoundingBox.Bottom &&
                    Math.Abs(w.BoundingBox.Left - header.BoundingBox.Left) < 50 &&
                    (Regex.IsMatch(w.Text, @"\d") || w.Text.EndsWith("RI")));

                if (number != null) return number.Text;
            }
            return "UNKNOWN";
        }

        private int ExtractPoNumberFromInvoice(Page page)
        {
            var match = Regex.Match(page.Text, @"P\.?O\.?\s*NO\.?[:\s]*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int val)) return val;
            return 0;
        }
    }
}
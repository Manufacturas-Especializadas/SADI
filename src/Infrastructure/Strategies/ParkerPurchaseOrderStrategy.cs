using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Infrastructure.Strategies
{
    public class ParkerPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "01 - Parker";
        public string DocumentTypeSubFolder => "02 - Factura";

        private class PoLineData
        {
            public string Part { get; set; } = "";
            public decimal Qty { get; set; }
            public string Unit { get; set; } = "";
            public decimal Amount { get; set; }
            public decimal UnitPrice { get; set; }
        }

        public List<PurchaseOrderItem> Extract(string invoiceFilePath)
        {
            var items = new List<PurchaseOrderItem>();
            string fileName = Path.GetFileName(invoiceFilePath);

            Console.WriteLine($"--- PROCESANDO PARKER: {fileName} ---");

            int poNumber = 0;
            string invoiceNum = "UNKNOWN";
            var poLinesReference = new List<PoLineData>();

            using (var pdfInv = PdfDocument.Open(invoiceFilePath))
            {
                var page1 = pdfInv.GetPage(1);

                invoiceNum = ExtractInvoiceNumberRegex(page1);
                Console.WriteLine($"Invoice: {invoiceNum}");

                poNumber = ExtractCustomerPoFromInvoice(page1);
                Console.WriteLine($"PO Number: {poNumber}");

                if (poNumber != 0)
                {
                    string poPath = FindPoFile(invoiceFilePath, poNumber);
                    if (poPath != "NOT FOUND")
                    {
                        using (var pdfPo = PdfDocument.Open(poPath))
                        {
                            foreach (var page in pdfPo.GetPages()) poLinesReference.AddRange(ExtractPoLines(page));
                        }
                    }
                }

                foreach (var page in pdfInv.GetPages())
                {
                    var invoiceLines = ExtractInvoiceLines(page);

                    foreach (var invLine in invoiceLines)
                    {
                        var newItem = new PurchaseOrderItem
                        {
                            PoNumber = poNumber,
                            VendorName = "PARKER HANNIFIN CORPORATION",
                            InvoiceNumber = invoiceNum,
                            PartNumber = invLine.Part,
                            TotalPrice = invLine.Amount,
                            UnitPrice = invLine.UnitPrice,
                            SourceFileName = fileName,
                            QtyInvPz = invLine.Qty
                        };

                        var matchPo = poLinesReference.FirstOrDefault(p =>
                                p.Part.Contains(invLine.Part) || invLine.Part.Contains(p.Part));

                        if (matchPo != null) newItem.QtyPoPz = matchPo.Qty;
                        else newItem.QtyPoPz = invLine.Qty;

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

            var unitPriceHeader = words.FirstOrDefault(w => w.Text.Contains("UNIT") &&
                words.Any(x => x.Text.Contains("PRICE") &&
                Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 20));

            var netHeader = words.FirstOrDefault(w =>
                w.Text.Contains("NET") &&
                words.Any(x => x.Text.Contains("AMOUNT") &&
                Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 20));

            var qtyHeader = words.FirstOrDefault(w =>
                Regex.IsMatch(w.Text, @"^QTY\.?$", RegexOptions.IgnoreCase));

            var boHeader = words.FirstOrDefault(w => w.Text.Contains("B/O"));

            if (unitPriceHeader == null || qtyHeader == null || netHeader == null)
                return list;

            double tableTopY = unitPriceHeader.BoundingBox.Bottom;
            double pageRight = page.Width;

            double startDescX = (boHeader != null)
                ? boHeader.BoundingBox.Right + 10
                : qtyHeader.BoundingBox.Right + 50;

            double endDescX = unitPriceHeader.BoundingBox.Left - 10;

            var lineCodes = words
                .Where(w =>
                    w.BoundingBox.Top < tableTopY &&
                    w.BoundingBox.Left > startDescX &&
                    w.BoundingBox.Right < endDescX &&
                    Regex.IsMatch(w.Text, @"^\d{4,}-\d+$"))
                .ToList();

            foreach (var codeWord in lineCodes)
            {
                var line = new PoLineData();

                double rowY = codeWord.BoundingBox.Centroid.Y;
                double rowTop = rowY + 22;
                double rowBottom = rowY - 22;

                var partCandidate = words
                    .Where(w =>
                        w.BoundingBox.Left > startDescX &&
                        w.BoundingBox.Right < endDescX &&
                        w.BoundingBox.Centroid.Y <= rowTop &&
                        w.BoundingBox.Centroid.Y >= rowBottom &&
                        Regex.IsMatch(w.Text, @"^(?=.*[A-Z])(?=.*\d)[A-Z0-9\-]{6,}$") &&
                        !Regex.IsMatch(w.Text, @"^\d{4,}-\d+$") &&
                        !w.Text.Contains("LOC") &&
                        !w.Text.Contains("PO#") &&
                        !w.Text.Contains("PRO") &&
                        w.Text != "7T"
                    )
                    .OrderBy(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY))
                    .FirstOrDefault();

                if (partCandidate == null) continue;
                line.Part = partCandidate.Text;

                var qtyWord = words
                    .Where(w =>
                        w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 25 &&
                        w.BoundingBox.Right <= qtyHeader.BoundingBox.Right + 25 &&
                        w.BoundingBox.Centroid.Y <= rowTop &&
                        w.BoundingBox.Centroid.Y >= rowBottom &&
                        Regex.IsMatch(w.Text, @"^\d+(\.\d+)?$"))
                    .FirstOrDefault();

                if (qtyWord != null &&
                    decimal.TryParse(qtyWord.Text.Replace(",", ""),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q))
                    line.Qty = q;

                var priceWord = words
                    .Where(w =>
                        w.BoundingBox.Left >= unitPriceHeader.BoundingBox.Left - 30 &&
                        w.BoundingBox.Right <= unitPriceHeader.BoundingBox.Right + 40 &&
                        w.BoundingBox.Centroid.Y <= rowTop &&
                        w.BoundingBox.Centroid.Y >= rowBottom &&
                        Regex.IsMatch(w.Text, @"\d+\.\d+"))
                    .FirstOrDefault();

                if (priceWord != null &&
                    decimal.TryParse(priceWord.Text.Replace("$", ""),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out decimal p))
                    line.UnitPrice = p;

                var netWord = words
                    .Where(w =>
                        w.BoundingBox.Left > pageRight * 0.72 &&
                        w.BoundingBox.Centroid.Y <= rowTop &&
                        w.BoundingBox.Centroid.Y >= rowBottom &&
                        Regex.IsMatch(w.Text.Replace("$", "").Replace(",", ""), @"^\d+(\.\d+)?$")
                    )
                    .OrderBy(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY))
                    .FirstOrDefault();

                if (netWord != null &&
                    decimal.TryParse(netWord.Text.Replace("$", ""),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out decimal net))
                    line.Amount = net;

                list.Add(line);

                Console.WriteLine($"✅ Item: {line.Qty} | {line.Part} | {line.UnitPrice} | {line.Amount}");
            }

            return list;
        }


        private List<PoLineData> ExtractPoLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();
            var partHeader = words.FirstOrDefault(w => w.Text.Contains("Part") || w.Text.Contains("Item"));
            var qtyHeader = words.FirstOrDefault(w => w.Text.Contains("Qty") || w.Text.Contains("Quantity"));

            if (partHeader == null || qtyHeader == null) return list;

            double tableTopY = partHeader.BoundingBox.Bottom;
            var partCandidates = words.Where(w => w.BoundingBox.Top < tableTopY && w.BoundingBox.Left >= partHeader.BoundingBox.Left - 20 && w.Text.Length > 3).ToList();

            foreach (var partWord in partCandidates)
            {
                double rowY = partWord.BoundingBox.Centroid.Y;
                var qtyWord = words.FirstOrDefault(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 15 && w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 30 && Regex.IsMatch(w.Text, @"\d"));
                if (qtyWord != null)
                {
                    var poLine = new PoLineData { Part = partWord.Text };
                    string cleanQty = Regex.Match(qtyWord.Text, @"[\d.,]+").Value.Replace(",", "");
                    if (decimal.TryParse(cleanQty, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q)) poLine.Qty = q;
                    list.Add(poLine);
                }
            }
            return list;
        }

        private string ExtractInvoiceNumberRegex(Page page)
        {
            var words = page.GetWords().ToList();

            var anchor = words.FirstOrDefault(w => w.Text.Contains("PLEASE") &&
                         words.Any(x => x.Text.Contains("INVOICE") && Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 20));

            if (anchor != null)
            {
                var number = words.FirstOrDefault(w =>
                    w.BoundingBox.Top < anchor.BoundingBox.Bottom &&
                    w.BoundingBox.Top > anchor.BoundingBox.Bottom - 100 &&
                    w.BoundingBox.Left >= anchor.BoundingBox.Left && 
                    (Regex.IsMatch(w.Text, @"\d") || w.Text.Contains("XF")));

                if (number != null) return number.Text;
            }

            var header = words.FirstOrDefault(w => w.Text == "INVOICE" && words.Any(n => n.Text.Contains("NO.") && n.BoundingBox.Left > w.BoundingBox.Left));
            if (header != null)
            {
                var num = words.FirstOrDefault(w => w.BoundingBox.Top < header.BoundingBox.Bottom && w.BoundingBox.Left > header.BoundingBox.Left - 50 && w.Text.Length > 5);
                if (num != null) return num.Text;
            }

            return "UNKNOWN";
        }

        private int ExtractCustomerPoFromInvoice(Page page)
        {
            var words = page.GetWords().ToList();
            var header = words.FirstOrDefault(w => w.Text.Contains("CUSTOMER") &&
                                             words.Any(x => x.Text.Contains("P.O.") &&
                                             Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 15));

            if (header != null)
            {
                var number = words.FirstOrDefault(w =>
                    w.BoundingBox.Top < header.BoundingBox.Bottom &&
                    w.BoundingBox.Top > header.BoundingBox.Bottom - 60 &&
                    Math.Abs(w.BoundingBox.Left - header.BoundingBox.Left) < 100 &&
                    Regex.IsMatch(w.Text, @"^\d+$"));

                if (number != null && int.TryParse(number.Text, out int val)) return val;
            }
            return 0;
        }

        private string FindPoFile(string invoiceFilePath, int poNumber)
        {
            try
            {
                var currentDir = Directory.GetParent(invoiceFilePath);
                if (currentDir?.Parent == null) return "NOT FOUND";
                string poFolder = Path.Combine(currentDir.Parent.FullName, "01 - Orden de compra");
                if (!Directory.Exists(poFolder)) return "NOT FOUND";
                var files = Directory.GetFiles(poFolder, $"*{poNumber}*.pdf");
                if (files.Any()) return files.First();
            }
            catch { }
            return "NOT FOUND";
        }
    }
}
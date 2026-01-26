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
            public string Part { get; set; } = "";
            public decimal Qty { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal Amount { get; set; }
            public int PoLineNumber { get; set; }
            public decimal PoQty { get; set; }
        }

        public List<PurchaseOrderItem> Extract(string invoiceFilePath)
        {
            var items = new List<PurchaseOrderItem>();
            string fileName = Path.GetFileName(invoiceFilePath);

            int searchPoNumber = 0;
            int officialPoNumber = 0;
            string invoiceNum = "UNKNOWN";
            string vendorName = "ETI, LLC";

            var poLinesReference = new List<PoLineData>();

            using (var pdfInv = PdfDocument.Open(invoiceFilePath))
            {
                var page1 = pdfInv.GetPage(1);

                invoiceNum = ExtractInvoiceNumber(page1);

                searchPoNumber = ExtractPoNumberFromInvoice(page1);

                if (searchPoNumber != 0)
                {
                    string poPath = FindPoFile(invoiceFilePath, searchPoNumber);
                    if (poPath != "NOT FOUND")
                    {
                        using (var pdfPo = PdfDocument.Open(poPath))
                        {
                            var poPage1 = pdfPo.GetPage(1);

                            officialPoNumber = ExtractOfficialPoNumber(poPage1);

                            if (officialPoNumber == 0) officialPoNumber = searchPoNumber;

                            foreach (var page in pdfPo.GetPages())
                            {
                                poLinesReference.AddRange(ExtractPoLines(page));
                            }
                        }
                    }
                    else
                    {
                        officialPoNumber = searchPoNumber;
                    }
                }

                foreach (var page in pdfInv.GetPages())
                {
                    var invoiceLines = ExtractInvoiceLines(page);

                    foreach (var line in invoiceLines)
                    {
                        var newItem = new PurchaseOrderItem
                        {
                            PoNumber = officialPoNumber,
                            VendorName = vendorName,
                            InvoiceNumber = invoiceNum,
                            PartNumber = line.Part,
                            SourceFileName = fileName,
                            QtyInvPz = line.Qty,
                            UnitPrice = line.UnitPrice,
                            TotalPrice = line.Amount,
                            QtyPoPz = line.Qty
                        };

                        var matchPo = poLinesReference.FirstOrDefault(p => p.Part.Contains(line.Part) || line.Part.Contains(p.Part));
                        if (matchPo != null)
                        {
                            newItem.QtyPoPz = matchPo.Qty;
                            newItem.LineNumber = matchPo.PoLineNumber;
                        }

                        items.Add(newItem);
                    }
                }
            }

            return items;
        }

        private int ExtractPoNumberFromInvoice(Page page)
        {
            var words = page.GetWords().ToList();

            var header = words.FirstOrDefault(w => w.Text.Contains("CUSTOMER") &&
                         words.Any(n => n.Text.Contains("ORDER") &&
                                        n.BoundingBox.Left > w.BoundingBox.Right &&
                                        Math.Abs(n.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            if (header != null)
            {
                var number = words.FirstOrDefault(w =>
                    w.BoundingBox.Top < header.BoundingBox.Bottom &&
                    Math.Abs(w.BoundingBox.Left - header.BoundingBox.Left) < 30 &&
                    Regex.IsMatch(w.Text, @"^\d+$"));

                if (number != null && int.TryParse(number.Text, out int result)) return result;
            }

            var match = Regex.Match(page.Text, @"CUSTOMER\s*ORDER\s*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int val)) return val;

            return 0;
        }

        private List<PoLineData> ExtractInvoiceLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();

            var custPartHeader = words.FirstOrDefault(w => w.Text.Contains("CUSTOMER") &&
                                                           words.Any(n => n.Text.Contains("NUMBER") && Math.Abs(n.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            var qtyHeader = words.FirstOrDefault(w => w.Text.Contains("SHIPPED"));
            if (qtyHeader == null) qtyHeader = words.FirstOrDefault(w => w.Text.Contains("QUANTITY"));

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

        private int ExtractOfficialPoNumber(Page page)
        {
            var match = Regex.Match(page.Text, @"(?:Nro\.?|No\.?|Num)\.?\s*(?:de)?\s*OC[:\s]*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int val)) return val;

            match = Regex.Match(page.Text, @"PO\s*Number[:\s]+(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int val2)) return val2;

            return 0;
        }

        private List<PoLineData> ExtractPoLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();

            var partHeader = words.FirstOrDefault(w => w.Text.Contains("Parte", StringComparison.OrdinalIgnoreCase) ||
                                                       w.Text.Contains("Part", StringComparison.OrdinalIgnoreCase));
            var qtyHeader = words.FirstOrDefault(w => w.Text.StartsWith("Cant", StringComparison.OrdinalIgnoreCase) ||
                                                      w.Text.StartsWith("Qty", StringComparison.OrdinalIgnoreCase));

            if (partHeader == null || qtyHeader == null) return list;

            double tableTopY = partHeader.BoundingBox.Bottom;

            var partCandidates = words.Where(w =>
                w.BoundingBox.Top < tableTopY &&
                w.Text.Length > 3 &&
                !w.Text.Contains("/") &&
                !w.Text.Contains("POForm", StringComparison.OrdinalIgnoreCase) &&
                !w.Text.Contains("Page", StringComparison.OrdinalIgnoreCase) &&
                Regex.IsMatch(w.Text, @"^[A-Za-z]") &&
                Regex.IsMatch(w.Text, @"\d") &&
                w.BoundingBox.Left >= partHeader.BoundingBox.Left - 80
            ).ToList();

            foreach (var partWord in partCandidates)
            {
                var lineData = new PoLineData { Part = partWord.Text };
                double rowY = partWord.BoundingBox.Centroid.Y;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 10).ToList();

                var lineNumWord = rowWords.FirstOrDefault(w => w.BoundingBox.Right < partWord.BoundingBox.Left && Regex.IsMatch(w.Text, @"^\d+$"));
                if (lineNumWord != null && int.TryParse(lineNumWord.Text, out int ln)) lineData.PoLineNumber = ln;

                var qtyWord = rowWords.FirstOrDefault(w =>
                    w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 30 &&
                    w.BoundingBox.Right <= qtyHeader.BoundingBox.Right + 30 &&
                    Regex.IsMatch(w.Text, @"\d"));

                if (qtyWord != null)
                {
                    string cleanNum = qtyWord.Text.Replace(",", "");
                    var match = Regex.Match(cleanNum, @"([\d.]+)");
                    if (match.Success && decimal.TryParse(match.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q))
                    {
                        lineData.Qty = q;
                    }
                }
                list.Add(lineData);
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
            var match = Regex.Match(page.Text, @"INVOICE\s*NO\.?\s*(\S+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;
            return "UNKNOWN";
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
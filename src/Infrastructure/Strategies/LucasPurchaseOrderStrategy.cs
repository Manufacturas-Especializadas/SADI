using Core.Entities;
using Core.Interfaces;
using System.Globalization;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Infrastructure.Strategies
{
    public class LucasPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "06 - LUCAS";
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

            int poNumber = 0;
            string invoiceNum = "UNKNOWN";
            var poLinesReference = new List<PoLineData>();

            using (var pdfInv = PdfDocument.Open(invoiceFilePath))
            {
                var page1 = pdfInv.GetPage(1);

                invoiceNum = ExtractInvoiceNumberRegex(page1.Text);
                poNumber = ExtractCustomerPoFromInvoice(page1.Text);

                if (poNumber != 0)
                {
                    string poPath = FindPoFile(invoiceFilePath, poNumber);
                    if (poPath != "NOT FOUND")
                    {
                        using (var pdfPo = PdfDocument.Open(poPath))
                        {
                            foreach (var page in pdfPo.GetPages())
                            {
                                poLinesReference.AddRange(ExtractPoLines(page));
                            }
                        }

                        items.Add(newItem);
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
                            VendorName = "LUCAS MILHAUPT INC",
                            InvoiceNumber = invoiceNum,
                            PartNumber = invLine.Part,
                            TotalPrice = invLine.Amount,
                            UnitPrice = invLine.UnitPrice,
                            SourceFileName = fileName
                        };

                        AssignQuantity(newItem, invLine.Qty, invLine.Unit, isInvoice: true);

                        var matchPo = poLinesReference.FirstOrDefault(p =>
                            p.Part.Contains(invLine.Part) || invLine.Part.Contains(p.Part));

                        if (matchPo != null)
                        {
                            AssignQuantity(newItem, matchPo.Qty, matchPo.Unit, isInvoice: false);
                        }

                        items.Add(newItem);
                    }
                }
            }

            return items;
        }

        private void AssignQuantity(PurchaseOrderItem item, decimal qty, string unit, bool isInvoice)
        {
            unit = unit.ToUpper().Trim();
            bool isKg = unit.Contains("KG") || unit.Contains("LB");

            if (isInvoice)
            {
                if (isKg) item.QtyInvKg = qty;
                else item.QtyInvPz = qty;
            }
            else
            {
                if (isKg) item.QtyPoKg = qty;
                else item.QtyPoPz = qty;
            }
        }

        private List<PoLineData> ExtractInvoiceLines(Page page)
        {
            var lines = new List<PoLineData>();
            var words = page.GetWords().ToList();

            var customerHeaders = words.Where(w => w.Text.Contains("Customer"));

            Word? correctCustomerHeader = null;
            Word? descHeader = null;
            Word? qtyHeader = null;
            Word? priceHeader = null;
            Word? amountHeader = null; 

            foreach (var candidate in customerHeaders)
            {
                double y = candidate.BoundingBox.Centroid.Y;

                var descCandidate = words.FirstOrDefault(w => w.Text.Contains("Description") && Math.Abs(w.BoundingBox.Centroid.Y - y) < 10);
                var qtyCandidate = words.FirstOrDefault(w => w.Text.Contains("Quantity") && Math.Abs(w.BoundingBox.Centroid.Y - y) < 10);

                var priceCandidate = words.FirstOrDefault(w => w.Text.Contains("Price") && Math.Abs(w.BoundingBox.Centroid.Y - y) < 10);
                var amtCandidate = words.FirstOrDefault(w => w.Text.Contains("Amount") && Math.Abs(w.BoundingBox.Centroid.Y - y) < 10);

                if (descCandidate != null && qtyCandidate != null && amtCandidate != null)
                {
                    correctCustomerHeader = candidate;
                    descHeader = descCandidate;
                    qtyHeader = qtyCandidate;
                    priceHeader = priceCandidate;
                    amountHeader = amtCandidate; 
                    break;
                }
            }

            if (correctCustomerHeader == null || descHeader == null || qtyHeader == null || amountHeader == null) return lines;

            double tableTopY = correctCustomerHeader.BoundingBox.Bottom;
            double minX = correctCustomerHeader.BoundingBox.Left - 10;
            double maxX = descHeader.BoundingBox.Left - 5;

            var partCandidates = words.Where(w =>
                w.BoundingBox.Top < tableTopY &&
                w.BoundingBox.Left >= minX &&
                w.BoundingBox.Right <= maxX &&
                w.Text.Length > 3
            ).ToList();

            foreach (var partWord in partCandidates)
            {
                if (partWord.Text.Contains("Customer") || partWord.Text == "PN") continue;

                double rowY = partWord.BoundingBox.Centroid.Y;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 10).ToList();

                var qtyWord = rowWords.FirstOrDefault(w =>
                    w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 50 &&
                    w.BoundingBox.Right <= qtyHeader.BoundingBox.Right + 50 &&
                    Regex.IsMatch(w.Text, @"\d")
                );

                var priceWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= priceHeader!.BoundingBox.Left - 40 
                            && w.BoundingBox.Left < amountHeader.BoundingBox.Left 
                            && Regex.IsMatch(w.Text, @"\d"));

                var amountWord = rowWords.FirstOrDefault(w =>
                            w.BoundingBox.Left >= amountHeader.BoundingBox.Left - 50 &&
                            Regex.IsMatch(w.Text, @"\d"));

                if (qtyWord != null)
                {
                    var line = new PoLineData { Part = partWord.Text };

                    var match = Regex.Match(qtyWord.Text, @"([\d,.]+)\s*([A-Za-z]*)");
                    if (match.Success)
                    {
                        string val = match.Groups[1].Value;
                        if (decimal.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q)) line.Qty = q;

                        string u = match.Groups[2].Value;
                        if (string.IsNullOrEmpty(u))
                        {
                            var unitWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left > qtyWord.BoundingBox.Right &&
                                                                        w.BoundingBox.Left < qtyWord.BoundingBox.Right + 80);
                            if (unitWord != null) u = unitWord.Text;
                        }
                        line.Unit = u;
                    }

                    if(priceWord != null)
                    {
                        string rawPrice = priceWord.Text.Replace("$", "").Replace(",", "").Trim();
                        if(decimal.TryParse(rawPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal pr))
                        {
                            line.UnitPrice = pr;
                        }
                    }

                    if (amountWord != null)
                    {
                        string rawAmount = amountWord.Text.Replace("$", "").Replace(",", "").Trim();

                        if (decimal.TryParse(rawAmount, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amt))
                        {
                            line.Amount = amt;
                        }
                    }

                    lines.Add(line);
                }
            }
            return lines;
        }

        private List<PoLineData> ExtractPoLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();
            var partHeader = words.FirstOrDefault(w => w.Text.Contains("Part"));
            var qtyHeader = words.FirstOrDefault(w => w.Text.Contains("Order") && words.Any(q => q.Text.Contains("Qty")));

            if (partHeader != null && qtyHeader != null)
            {
                double tableTopY = partHeader.BoundingBox.Bottom;
                var partCandidates = words.Where(w =>
                    w.BoundingBox.Top < tableTopY &&
                    w.BoundingBox.Left >= partHeader.BoundingBox.Left - 20 &&
                    w.BoundingBox.Right <= partHeader.BoundingBox.Right + 150 &&
                    w.Text.Length > 3 &&
                    !w.Text.Contains("Handy") && !w.Text.Contains("Ring")
                ).ToList();

                foreach (var partWord in partCandidates)
                {
                    double rowY = partWord.BoundingBox.Centroid.Y;
                    var qtyWord = words.FirstOrDefault(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 15 &&
                                                            w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 30 &&
                                                            Regex.IsMatch(w.Text, @"\d"));

                    if (qtyWord != null)
                    {
                        var poLine = new PoLineData { Part = partWord.Text };
                        var match = Regex.Match(qtyWord.Text, @"([\d,.]+)\s*([A-Za-z]*)");
                        if (match.Success)
                        {
                            string val = match.Groups[1].Value.Replace(",", "");
                            if (decimal.TryParse(val, out decimal q)) poLine.Qty = q;

                            string u = match.Groups[2].Value;
                            if (string.IsNullOrEmpty(u))
                            {
                                var unitWord = words.FirstOrDefault(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 15 &&
                                                                         w.BoundingBox.Left > qtyWord.BoundingBox.Right);
                                if (unitWord != null) u = unitWord.Text;
                            }
                            poLine.Unit = u;
                        }
                        list.Add(poLine);
                    }
                }
            }
            return list;
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

        private int ExtractCustomerPoFromInvoice(string text)
        {
            var match = Regex.Match(text, @"Customer\s*PO\s*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int result)) return result;

            match = Regex.Match(text, @"PO\s*#?\s*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int result2)) return result2;

            return 0;
        }

        private string ExtractInvoiceNumberRegex(string text)
        {
            var match = Regex.Match(text, @"Invoice\s+Number\s+([A-Z0-9]+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string val = match.Groups[1].Value;
                if (val.Contains("Please")) val = val.Substring(0, val.IndexOf("Please"));
                return val;
            }
            return "UNKNOWN";
        }
    }
}
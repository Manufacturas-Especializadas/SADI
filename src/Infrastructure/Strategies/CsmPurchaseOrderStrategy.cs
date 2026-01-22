using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Infrastructure.Strategies
{
    public class CsmPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "05 - CSM";
        public string DocumentTypeSubFolder => "02 - Factura";

        private class PoLineData
        {
            public string Part { get; set; } = "";
            public decimal Qty { get; set; }
            public string Unit { get; set; } = "";
            public decimal Amount { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal WeightKg { get; set; }
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

                invoiceNum = ExtractInvoiceNumber(page1);
                poNumber = ExtractPoNumber(page1);

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
                            VendorName = "CSM CORPORATION",
                            InvoiceNumber = invoiceNum,
                            PartNumber = invLine.Part,
                            SourceFileName = fileName,
                            TotalPrice = invLine.Amount,
                            UnitPrice = invLine.UnitPrice,
                            QtyInvPz = invLine.Qty
                        };

                        var matchPo = poLinesReference.FirstOrDefault(p =>
                                    p.Part.Contains(invLine.Part) || invLine.Part.Contains(p.Part));

                        if (matchPo != null)
                        {
                            newItem.QtyPoPz = matchPo.Qty;
                            newItem.QtyPoKg = matchPo.WeightKg;
                            newItem.QtyInvKg = matchPo.WeightKg;
                        }

                        items.Add(newItem);
                    }
                }
            }

            return items;
        }

        private List<PoLineData> ExtractInvoiceLines(Page page)
        {
            var lines = new List<PoLineData>();
            var words = page.GetWords().ToList();

            // 1. ENCONTRAR ENCABEZADOS
            var qtyHeader = words.FirstOrDefault(w => w.Text.Contains("Quantity"));
            var descHeader = words.FirstOrDefault(w => w.Text.Contains("Description") && Math.Abs(w.BoundingBox.Centroid.Y - (qtyHeader?.BoundingBox.Centroid.Y ?? 0)) < 20);
            var itemCodeHeader = words.FirstOrDefault(w => w.Text.Contains("Item") &&
                                                           words.Any(n => n.Text.Contains("Code") && Math.Abs(n.BoundingBox.Centroid.Y - w.BoundingBox.Centroid.Y) < 10));
            var priceHeader = words.FirstOrDefault(w => w.Text.Contains("Price") && words.Any(x => x.Text.Contains("Each")));
            var amountHeader = words.FirstOrDefault(w => w.Text.Contains("Amount"));

            if (qtyHeader == null || itemCodeHeader == null || priceHeader == null) return lines;

            double tableTopY = itemCodeHeader.BoundingBox.Bottom;
            double minX = (descHeader != null) ? descHeader.BoundingBox.Right : itemCodeHeader.BoundingBox.Left - 10;
            double maxX = priceHeader.BoundingBox.Left + 10;

            // --- DETECCIÓN DE PISO (FRENO DE MANO) ---
            var footerKeywords = new[] { "Total", "Phone", "COUNTRY", "THESE", "DIVERSION", "ORIGIN" };
            var footerWord = words
                .Where(w => w.BoundingBox.Top < tableTopY &&
                            footerKeywords.Any(k => w.Text.Contains(k, StringComparison.OrdinalIgnoreCase)))
                .OrderByDescending(w => w.BoundingBox.Top)
                .FirstOrDefault();

            double tableBottomY = footerWord != null ? footerWord.BoundingBox.Top : 0;

            // 2. FILTRAR CANDIDATOS
            var partCandidates = words.Where(w =>
                w.BoundingBox.Top < tableTopY &&
                w.BoundingBox.Bottom > tableBottomY &&
                w.BoundingBox.Left >= minX &&
                w.BoundingBox.Right <= maxX &&
                w.Text.Length > 3 &&
                !w.Text.ToUpper().Contains("FREIGHT") &&
                !w.Text.ToUpper().Contains("DESCRIPTION")
            ).ToList();

            foreach (var partWord in partCandidates)
            {
                if (partWord.Text == "U.S." || partWord.Text == "DIVERSION") continue;

                double rowY = partWord.BoundingBox.Centroid.Y;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 10).ToList();

                var line = new PoLineData { Part = partWord.Text };

                // Quantity
                var qtyWord = rowWords.FirstOrDefault(w =>
                    w.BoundingBox.Left >= qtyHeader.BoundingBox.Left - 40 &&
                    w.BoundingBox.Right <= qtyHeader.BoundingBox.Right + 30 &&
                    Regex.IsMatch(w.Text, @"\d"));

                if (qtyWord != null && decimal.TryParse(qtyWord.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal q))
                {
                    line.Qty = q;
                }

                // --- CORRECCIÓN EN PRICE EACH ---
                // Definimos el límite derecho usando AMOUNT. Si no hay Amount, damos un margen amplio (150).
                double priceMaxX = (amountHeader != null) ? amountHeader.BoundingBox.Left - 10 : priceHeader.BoundingBox.Right + 150;

                var priceWord = rowWords.FirstOrDefault(w =>
                    w.BoundingBox.Left >= priceHeader.BoundingBox.Left - 20 && // Desde donde empieza "Price"
                    w.BoundingBox.Right <= priceMaxX && // Hasta donde empieza "Amount"
                    Regex.IsMatch(w.Text, @"\d"));

                if (priceWord != null)
                {
                    string rawPrice = priceWord.Text.Replace("$", "").Replace(",", "").Trim();
                    if (decimal.TryParse(rawPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal p))
                    {
                        line.UnitPrice = p;
                    }
                }

                if (amountHeader != null)
                {
                    var amtWord = rowWords.FirstOrDefault(w =>
                        w.BoundingBox.Left >= amountHeader.BoundingBox.Left - 40 &&
                        Regex.IsMatch(w.Text, @"\d"));

                    if (amtWord != null)
                    {
                        string rawAmt = amtWord.Text.Replace("$", "").Replace(",", "").Trim();
                        if (decimal.TryParse(rawAmt, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal a))
                        {
                            line.Amount = a;
                        }
                    }
                }

                lines.Add(line);
            }

            return lines;
        }

        private List<PoLineData> ExtractPoLines(Page page)
        {
            var list = new List<PoLineData>();
            var words = page.GetWords().ToList();
            var partHeader = words.FirstOrDefault(w => w.Text.Contains("Part") || w.Text.Contains("Item"));

            if (partHeader != null)
            {
                var candidates = words.Where(w => w.BoundingBox.Left >= partHeader.BoundingBox.Left - 20
                                              && w.BoundingBox.Top < partHeader.BoundingBox.Bottom &&
                                              w.Text.Length > 3).ToList();

                foreach (var part in candidates)
                {
                    var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - part.BoundingBox.Centroid.Y) < 15).ToList();

                    var weightWord = rowWords.FirstOrDefault(w => w.Text.ToUpper().Contains("KG") || w.Text.ToUpper().Contains("LB"));

                    if (weightWord != null)
                    {
                        var match = Regex.Match(weightWord.Text, @"([\d,.]+)\s*(KG|LB|LBS)", RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            decimal weight = 0;
                            decimal.TryParse(match.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out weight);
                            if (match.Groups[2].Value.ToUpper().Contains("LB")) weight = weight * 0.453592m;

                            list.Add(new PoLineData { Part = part.Text, WeightKg = weight, Qty = weight });
                        }
                    }
                }
            }
            return list;
        }

        private string ExtractInvoiceNumber(Page page)
        {
            var header = page.GetWords().FirstOrDefault(w => w.Text.Contains("Invoice") &&
                         page.GetWords().Any(x => x.Text.Contains("#") && Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            if (header != null)
            {
                var number = page.GetWords().FirstOrDefault(w =>
                    w.BoundingBox.Top < header.BoundingBox.Bottom &&
                    Math.Abs(w.BoundingBox.Left - header.BoundingBox.Left) < 50 &&
                    Regex.IsMatch(w.Text, @"^\d+$"));

                if (number != null) return number.Text;
            }

            var match = Regex.Match(page.Text, @"Invoice\s*#\s*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;
            return "UNKNOWN";
        }

        private int ExtractPoNumber(Page page)
        {
            var words = page.GetWords().ToList();
            var header = words.FirstOrDefault(w => w.Text.Contains("P.O.") &&
                         words.Any(x => x.Text.Contains("Number") && Math.Abs(x.BoundingBox.Bottom - w.BoundingBox.Bottom) < 10));

            if (header != null)
            {
                var number = words.FirstOrDefault(w =>
                    w.BoundingBox.Top < header.BoundingBox.Bottom &&
                    Math.Abs(w.BoundingBox.Left - header.BoundingBox.Left) < 50 &&
                    Regex.IsMatch(w.Text, @"^\d+$"));

                if (number != null && int.TryParse(number.Text, out int res)) return res;
            }

            var match = Regex.Match(page.Text, @"P\.?O\.?\s*Number[:\s]*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int resRegex)) return resRegex;

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
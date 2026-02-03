using ClosedXML.Excel;
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

        private class ExtraLogisticsInfo
        {
            public string Reference { get; set; } = "";
            public string Weight { get; set; } = "";
            public string Guia { get; set; } = "";
        }

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

            var allInvoiceLines = new List<PoLineData>();
            var poLinesReference = new List<PoLineData>();
            var logisticsInfo = new ExtraLogisticsInfo();

            using (var pdfInv = PdfDocument.Open(invoiceFilePath))
            {
                var page1 = pdfInv.GetPage(1);
                invoiceNum = ExtractInvoiceNumberRegex(page1.Text);
                poNumber = ExtractCustomerPoFromInvoice(page1.Text);

                foreach (var page in pdfInv.GetPages())
                {
                    allInvoiceLines.AddRange(ExtractInvoiceLines(page));
                }
            }

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


                string foundPackingListPath =
                            FindPackingListByCustItemNumbers(invoiceFilePath);



                if (foundPackingListPath != "NOT FOUND")
                {
                    string packingListName = Path.GetFileName(foundPackingListPath);

                    logisticsInfo = GetLogisticsFromExcelByFilename(invoiceFilePath, packingListName);
                }
            }

            foreach (var invLine in allInvoiceLines)
            {
                var newItem = new PurchaseOrderItem
                {
                    PoNumber = poNumber,
                    VendorName = "LUCAS MILHAUPT INC",
                    InvoiceNumber = invoiceNum,
                    PartNumber = invLine.Part,
                    TotalPrice = invLine.Amount,
                    UnitPrice = invLine.UnitPrice,
                    SourceFileName = fileName,
                    Reference = logisticsInfo.Reference,
                    Weight = logisticsInfo.Weight,
                    Guia = logisticsInfo.Guia
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

            return items;
        }

        private string FindPackingListByCustItemNumbers(
            string invoiceFilePath
            )
        {
            try
            {
                var currentDir = Directory.GetParent(invoiceFilePath);
                if (currentDir?.Parent == null) return "NOT FOUND";

                string parentPath = currentDir.Parent.FullName;
                string packingFolder = Path.Combine(parentPath, "03 - Packing List");

                if (!Directory.Exists(packingFolder))
                    return "NOT FOUND";

                var files = Directory.GetFiles(packingFolder, "*.pdf");

                foreach (var file in files)
                {
                    var custItems = ExtractCustItemNumbersFromPackingList(file);

                    if (custItems.Any())
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"📦 Packing List válido: {Path.GetFileName(file)}");
                        Console.ResetColor();

                        foreach (var ci in custItems.Distinct())
                        {
                            Console.WriteLine($"   ✔ CustItem: {ci}");
                        }

                        return file;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] {ex.Message}");
            }

            return "NOT FOUND";
        }

        private List<string> ExtractCustItemNumbersFromPackingList(string filePath)
        {
            var result = new HashSet<string>();

            using (var pdf = PdfDocument.Open(filePath))
            {
                foreach (var page in pdf.GetPages())
                {
                    var words = page.GetWords().ToList();

                    var custWords = words
                        .Where(w => w.Text.ToUpper().Contains("CUST"))
                        .ToList();

                    foreach (var cust in custWords)
                    {
                        double baseY = cust.BoundingBox.Centroid.Y;

                        var belowWords = words
                            .Where(w =>
                                w.BoundingBox.Centroid.Y < baseY - 3 &&
                                w.BoundingBox.Centroid.Y > baseY - 40 &&
                                w.BoundingBox.Left >= cust.BoundingBox.Left - 5)
                            .OrderBy(w => w.BoundingBox.Centroid.Y)
                            .ThenBy(w => w.BoundingBox.Left)
                            .Select(w => w.Text)
                            .ToList();

                        if (!belowWords.Any())
                            continue;

                        string combined = Regex.Replace(
                            string.Concat(belowWords).ToUpper(),
                            @"[^A-Z0-9]",
                            ""
                        );

                        if (combined.Length >= 6 && Regex.IsMatch(combined, @"[A-Z]+\d+"))
                        {
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine($"📦 CustItem encontrado en {Path.GetFileName(filePath)}: {combined}");
                            Console.ResetColor();

                            result.Add(combined);
                        }
                    }
                }
            }

            return result.ToList();
        }


        private ExtraLogisticsInfo GetLogisticsFromExcelByFilename(string invoiceFilePath, string targetFilename)
        {
            var info = new ExtraLogisticsInfo();
            try
            {
                var rootDir = Directory.GetParent(invoiceFilePath)?.Parent?.Parent;
                if (rootDir == null) return info;

                string excelPath = Path.Combine(rootDir.FullName, "Email Info.xlsx");

                if (File.Exists(excelPath))
                {
                    using (var workbook = new XLWorkbook(excelPath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var rows = worksheet.RangeUsed()?.RowsUsed();
                        if (rows == null) return info;

                        var headerRow = rows.First();
                        int colPdf = -1, colRef = -1, colWeight = -1, colGuia = -1;

                        foreach (var cell in headerRow.Cells())
                        {
                            string val = cell.GetString().Trim().ToUpper();
                            if (val == "PDF") colPdf = cell.Address.ColumnNumber;
                            if (val == "REFERENCE") colRef = cell.Address.ColumnNumber;
                            if (val == "WEIGHT") colWeight = cell.Address.ColumnNumber;
                            if (val == "GUIA") colGuia = cell.Address.ColumnNumber;
                        }

                        if (colPdf != -1)
                        {
                            foreach (var row in rows.Skip(1))
                            {
                                string cellPdfName = row.Cell(colPdf).GetString().Trim();

                                if (string.Equals(cellPdfName, targetFilename, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (colRef != -1) info.Reference = row.Cell(colRef).GetString();
                                    if (colWeight != -1) info.Weight = row.Cell(colWeight).GetString();
                                    if (colGuia != -1) info.Guia = row.Cell(colGuia).GetString();
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return info;
        }

        private void AssignQuantity(PurchaseOrderItem item, decimal qty, string unit, bool isInvoice)
        {
            unit = unit.ToUpper().Trim();
            bool isKg = unit.Contains("KG") || unit.Contains("LB");

            if (isInvoice) { if (isKg) item.QtyInvKg = qty; else item.QtyInvPz = qty; }
            else { if (isKg) item.QtyPoKg = qty; else item.QtyPoPz = qty; }
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

                var sameRowWords = words
                    .Where(w =>
                        Math.Abs(w.BoundingBox.Centroid.Y - rowY) < 8 &&
                        w.BoundingBox.Left >= partWord.BoundingBox.Left)
                    .OrderBy(w => w.BoundingBox.Left)
                    .Select(w => w.Text)
                    .ToList();

                string combinedPart = string.Concat(sameRowWords);

                var poLine = new PoLineData
                {
                    Part = combinedPart
                };

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

                    if (priceWord != null)
                    {
                        string rawPrice = priceWord.Text.Replace("$", "").Replace(",", "").Trim();
                        if (decimal.TryParse(rawPrice, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal pr)) line.UnitPrice = pr;
                    }

                    if (amountWord != null)
                    {
                        string rawAmount = amountWord.Text.Replace("$", "").Replace(",", "").Trim();
                        if (decimal.TryParse(rawAmount, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amt)) line.Amount = amt;
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
    }
}
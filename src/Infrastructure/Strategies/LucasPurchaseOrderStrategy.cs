using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;

namespace Infrastructure.Strategies
{
    public class LucasPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "06- LUCAS";
        public string DocumentTypeSubFolder => "orden de compra";

        public List<PurchaseOrderItem> Extract(string filePath)
        {
            var items = new List<PurchaseOrderItem>();
            string fileName = Path.GetFileName(filePath);

            using (var pdf = PdfDocument.Open(filePath))
            {
                var page1 = pdf.GetPage(1);
                string poNumber = GetPoNumberRegex(page1);
                string vendorName = GetVendorName(page1);

                if (poNumber == "UNKNOWN") Console.WriteLine($"[ERROR] No PO found in {fileName}");

                var invoiceData = GetInvoiceDetails(filePath, poNumber);

                string incoterm = invoiceData.Incoterm;
                string invoiceNum = invoiceData.InvoiceNum;

                foreach (var page in pdf.GetPages())
                {
                    var pageItems = ProcessTableOnPage(page, poNumber, vendorName, incoterm, invoiceNum, fileName);
                    items.AddRange(pageItems);
                }
            }
            return items;
        }

        private (string Incoterm, string InvoiceNum) GetInvoiceDetails(string poFilePath, string targetPoNumber)
        {
            if (targetPoNumber == "UNKNOWN") return ("PO MISSING", "PO MISSING");

            try
            {
                var poDirInfo = new DirectoryInfo(Path.GetDirectoryName(poFilePath)!);
                var clientRoot = poDirInfo.Parent;
                if (clientRoot == null) return ("PATH ERROR", "PATH ERROR");

                string invoiceFolder = Path.Combine(clientRoot.FullName, "factura");

                if (!Directory.Exists(invoiceFolder)) return ("NO INVOICE FOLDER", "NO INVOICE FOLDER");

                var invoiceFiles = Directory.GetFiles(invoiceFolder, "*.pdf");

                foreach (var invoicePath in invoiceFiles)
                {
                    using (var pdf = PdfDocument.Open(invoicePath))
                    {
                        var page = pdf.GetPage(1);
                        var text = page.Text;

                        if (text.Contains(targetPoNumber))
                        {
                            string foundIncoterm = ExtractIncotermFromPage(page, targetPoNumber);

                            string foundInvoice = ExtractInvoiceNumberRegex(text);

                            return (foundIncoterm, foundInvoice);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading invoice: {ex.Message}");
            }

            return ("NOT FOUND", "NOT FOUND");
        }

        private string ExtractInvoiceNumberRegex(string fullText)
        {
            var match = Regex.Match(fullText, @"Invoice\s+Number\s+([A-Z0-9]+)", RegexOptions.IgnoreCase);

            if (match.Success)
            {
                string rawValue = match.Groups[1].Value;

                if (rawValue.Contains("Please", StringComparison.OrdinalIgnoreCase))
                {
                    int index = rawValue.IndexOf("Please", StringComparison.OrdinalIgnoreCase);
                    return rawValue.Substring(0, index);
                }

                return rawValue;
            }
            return "UNKNOWN";
        }

        private string ExtractIncotermFromPage(Page page, string poNumber)
        {            
            try
            {
                var annotations = page.ExperimentalAccess.GetAnnotations().ToList();
                foreach (var annotation in annotations)
                {
                    string content = annotation.Content;
                    if (!string.IsNullOrEmpty(content))
                    {
                        string cleanContent = content.Replace("\r", " ").Replace("\n", " ").Trim();
                        if (cleanContent.Contains("Incoterm", StringComparison.OrdinalIgnoreCase))
                        {
                            var match = Regex.Match(cleanContent, @"Incoterm\s*[:\-\s]*([A-Za-z\s]+)", RegexOptions.IgnoreCase);
                            if (match.Success) return match.Groups[1].Value.Trim();
                            return cleanContent;
                        }
                        if (cleanContent.Contains("Ex Works", StringComparison.OrdinalIgnoreCase) ||
                           cleanContent.Contains("EXW", StringComparison.OrdinalIgnoreCase))
                        {
                            return cleanContent;
                        }
                    }
                }
            }
            catch { }

            var words = page.GetWords().ToList();
            var anchorTop = words.FirstOrDefault(w => w.Text.Contains(poNumber));
            if (anchorTop != null)
            {
                double searchCeiling = anchorTop.BoundingBox.Bottom;
                double searchLeft = anchorTop.BoundingBox.Left - 50;
                double searchRight = anchorTop.BoundingBox.Right + 150;

                var candidateWords = words.Where(w =>
                    w.BoundingBox.Top < searchCeiling &&
                    w.BoundingBox.Top > searchCeiling - 100 &&
                    w.BoundingBox.Right > searchLeft &&
                    w.BoundingBox.Left < searchRight
                ).OrderByDescending(w => w.BoundingBox.Top).ThenBy(w => w.BoundingBox.Left).ToList();

                var result = new List<string>();
                foreach (var word in candidateWords)
                {
                    if (word.Text.Contains("Incoterm", StringComparison.OrdinalIgnoreCase)) continue;
                    if (word.Text.Contains("Customer", StringComparison.OrdinalIgnoreCase)) continue;
                    if (word.Text.Contains("Mode", StringComparison.OrdinalIgnoreCase)) break;
                    if (word.Text.Contains("Buyer", StringComparison.OrdinalIgnoreCase)) break;
                    if (word.Text.Contains("Phone", StringComparison.OrdinalIgnoreCase)) break;
                    result.Add(word.Text);
                }
                if (result.Any()) return string.Join(" ", result);
            }
            return "UNKNOWN";
        }
        private string GetPoNumberRegex(Page page)
        {
            var text = page.Text;
            var match = Regex.Match(text, @"PO\s*Number[:\s]+(\d+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;
            return "UNKNOWN";
        }

        private string GetVendorName(Page page)
        {
            var words = page.GetWords();
            var lucasWord = words.FirstOrDefault(w => w.Text.ToUpper().Contains("LUCAS") && w.Text.ToUpper().Contains("MILHAUPT"));
            if (lucasWord != null) return "LUCAS MILHAUPT INC";
            return "UNKNOWN";
        }

        private List<PurchaseOrderItem> ProcessTableOnPage(Page page, string po, string vendor, string incoterm, string invoiceNum, string file)
        {
            var items = new List<PurchaseOrderItem>();
            var words = page.GetWords().ToList();

            var lineHeader = words.FirstOrDefault(w => w.Text == "Line");
            var partHeader = words.FirstOrDefault(w => w.Text == "Part");
            var qtyHeader = words.FirstOrDefault(w => w.Text == "Order" || w.Text == "Qty");
            var priceHeader = words.FirstOrDefault(w => w.Text == "Unit" && words.Any(n => n.Text == "Price" && n.BoundingBox.Left > w.BoundingBox.Left));
            var totalHeader = words.FirstOrDefault(w => w.Text == "Ext" && words.Any(n => n.Text == "Price" && n.BoundingBox.Left > w.BoundingBox.Left));

            if (lineHeader == null || partHeader == null) return items;

            double tableTopY = lineHeader.BoundingBox.Bottom;

            var lineCandidates = words.Where(w =>
               w.BoundingBox.Top < tableTopY &&
               w.BoundingBox.Centroid.Y < lineHeader.BoundingBox.Bottom &&
               w.BoundingBox.Left >= lineHeader.BoundingBox.Left - 20 &&
               w.BoundingBox.Right <= lineHeader.BoundingBox.Right + 20 &&
               int.TryParse(w.Text, out _)
           ).ToList();

            foreach (var lineWord in lineCandidates)
            {
                int lineNumber = int.Parse(lineWord.Text);
                double rowY = lineWord.BoundingBox.Centroid.Y;
                double rowTolerance = 10;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < rowTolerance && w != lineWord).ToList();

                var item = new PurchaseOrderItem
                {
                    PoNumber = po,
                    VendorName = vendor,
                    Incoterm = incoterm,
                    InvoiceNumber = invoiceNum, 
                    LineNumber = lineNumber,
                    SourceFileName = file
                };

                var partNumWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= partHeader.BoundingBox.Left - 10);
                if (partNumWord != null) item.PartNumber = partNumWord.Text;

                double searchQtyX = qtyHeader != null ? qtyHeader.BoundingBox.Left : partHeader.BoundingBox.Right + 50;
                var qtyWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= searchQtyX - 20 && Regex.IsMatch(w.Text, @"\d"));

                if (qtyWord != null)
                {
                    var match = Regex.Match(qtyWord.Text, @"([\d,.]+)([A-Za-z]*)");
                    if (match.Success)
                    {
                        string numberPart = match.Groups[1].Value.Replace(",", "");
                        string unitPart = match.Groups[2].Value;
                        if (decimal.TryParse(numberPart, out decimal q)) item.Quantity = q;
                        if (!string.IsNullOrEmpty(unitPart)) item.Unit = unitPart;
                        else
                        {
                            var nextWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left > qtyWord.BoundingBox.Right && w.BoundingBox.Left < qtyWord.BoundingBox.Right + 20);
                            if (nextWord != null && Regex.IsMatch(nextWord.Text, @"^[A-Za-z]+$")) item.Unit = nextWord.Text;
                        }
                    }
                }

                if (priceHeader != null)
                {
                    var priceWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= priceHeader.BoundingBox.Left - 20 && w.Text.Any(char.IsDigit));
                    if (priceWord != null)
                    {
                        string cleanPrice = priceWord.Text.Split('/')[0].Replace(",", "");
                        if (decimal.TryParse(cleanPrice, out decimal p)) item.UnitPrice = p;
                    }
                }

                if (totalHeader != null)
                {
                    var totalWord = rowWords.LastOrDefault(w => w.BoundingBox.Left >= totalHeader.BoundingBox.Left - 20);
                    if (totalWord != null)
                    {
                        string cleanTotal = Regex.Match(totalWord.Text, @"[\d.,]+").Value.Replace(",", "");
                        if (decimal.TryParse(cleanTotal, out decimal t)) item.TotalPrice = t;
                    }
                }

                items.Add(item);
            }
            return items;
        }
    }
}
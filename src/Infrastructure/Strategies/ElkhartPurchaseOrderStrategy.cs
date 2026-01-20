using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;

namespace Infrastructure.Strategies
{
    public class ElkhartPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "02 - Elkhart";
        public string DocumentTypeSubFolder => "01 - Orden de compra";

        public List<PurchaseOrderItem> Extract(string filePath)
        {
            var items = new List<PurchaseOrderItem>();
            string fileName = Path.GetFileName(filePath);

            using (var pdf = PdfDocument.Open(filePath))
            {
                var page1 = pdf.GetPage(1);
                string poNumber = GetPoNumberRegex(page1);
                string vendorName = GetVendorName(page1);

                if (poNumber == "UNKNOWN") Console.WriteLine($"[ADVERTENCIA] PO no encontrada en {fileName}");

                string invoiceNum = GetInvoiceNumberFromFolder(filePath, poNumber);

                foreach (var page in pdf.GetPages())
                {
                    var pageItems = ProcessTableOnPage(page, poNumber, vendorName, "N/A", invoiceNum, fileName);
                    items.AddRange(pageItems);
                }
            }

            return items;
        }

        private List<PurchaseOrderItem> ProcessTableOnPage(Page page, string po, string vendor, string incoterm, string invoiceNum, string file)
        {
            var items = new List<PurchaseOrderItem>();
            var words = page.GetWords().ToList();

            var partHeader = words.FirstOrDefault(w => w.Text.Contains("Parte", StringComparison.OrdinalIgnoreCase) ||
                                                       w.Text.Contains("Descrip", StringComparison.OrdinalIgnoreCase));

            var qtyHeader = words.FirstOrDefault(w => w.Text.StartsWith("Cant", StringComparison.OrdinalIgnoreCase) ||
                                                      w.Text.StartsWith("Qty", StringComparison.OrdinalIgnoreCase));

            if (partHeader == null) return items;

            double tableTopY = partHeader.BoundingBox.Bottom;

            var partCandidates = words.Where(w =>
                w.BoundingBox.Top < tableTopY &&
                w.Text.Length > 3 &&

                !w.Text.Contains("POForm", StringComparison.OrdinalIgnoreCase) &&
                !w.Text.Contains("Page", StringComparison.OrdinalIgnoreCase) &&
                !w.Text.Contains("Página", StringComparison.OrdinalIgnoreCase) &&

                Regex.IsMatch(w.Text, @"[A-Z]") &&
                Regex.IsMatch(w.Text, @"\d") && 

                w.BoundingBox.Left >= partHeader.BoundingBox.Left - 80 &&
                w.BoundingBox.Right <= partHeader.BoundingBox.Right + 150
            ).ToList();

            foreach (var partWord in partCandidates)
            {
                double rowY = partWord.BoundingBox.Centroid.Y;
                double rowTolerance = 15;

                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY) < rowTolerance).ToList();

                var item = new PurchaseOrderItem
                {
                    PoNumber = po,
                    VendorName = vendor,
                    Incoterm = incoterm,
                    InvoiceNumber = invoiceNum,
                    PartNumber = partWord.Text,
                    SourceFileName = file
                };

                var lineWord = rowWords.FirstOrDefault(w => w.BoundingBox.Right < partWord.BoundingBox.Left && Regex.IsMatch(w.Text, @"^\d+$"));
                if (lineWord != null && int.TryParse(lineWord.Text, out int ln))
                {
                    item.LineNumber = ln;
                }
                else
                {
                    item.LineNumber = 0;
                }

                double searchQtyX = (qtyHeader != null) ? qtyHeader.BoundingBox.Left - 20 : partWord.BoundingBox.Right + 20;
                var qtyWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= searchQtyX && Regex.IsMatch(w.Text, @"\d"));

                if (qtyWord != null)
                {
                    var match = Regex.Match(qtyWord.Text, @"([\d,.]+)");
                    if (match.Success)
                    {
                        string cleanNum = match.Groups[1].Value.Replace(",", "");
                        if (decimal.TryParse(cleanNum, out decimal q)) item.Quantity = q;
                    }
                }

                items.Add(item);
            }

            return items;
        }

        private string GetInvoiceNumberFromFolder(string poFilePath, string targetPoNumber)
        {
            if (targetPoNumber == "UNKNOWN") return "PO MISSING";

            try
            {
                var poDirInfo = new DirectoryInfo(Path.GetDirectoryName(poFilePath)!);
                var clientRoot = poDirInfo.Parent;
                if (clientRoot == null) return "PATH ERROR";

                string invoiceFolder = Path.Combine(clientRoot.FullName, "02 - Factura");

                if (!Directory.Exists(invoiceFolder)) return "NO INVOICE FOLDER";

                var invoiceFiles = Directory.GetFiles(invoiceFolder, "*.pdf");

                foreach (var invoicePath in invoiceFiles)
                {
                    using (var pdf = PdfDocument.Open(invoicePath))
                    {
                        var page = pdf.GetPage(1);
                        var text = page.Text;

                        if (text.Contains(targetPoNumber))
                        {
                            return ExtractInvoiceNumberElkhart(page);
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine($"Error Invoice: {ex.Message}"); }

            return "NOT FOUND";
        }

        private string ExtractInvoiceNumberElkhart(Page page)
        {
            var words = page.GetWords().ToList();

            var anchor = words.FirstOrDefault(w => w.Text.ToUpper().Contains("INVOICE") &&
                                                words.Any(next => (next.Text.ToUpper().Contains("NO.") || next.Text.ToUpper() == "NO") &&
                                                next.BoundingBox.Left > w.BoundingBox.Left &&
                                                Math.Abs(next.BoundingBox.Bottom - w.BoundingBox.Bottom) < 5));

            if (anchor != null)
            {
                double searchTop = anchor.BoundingBox.Bottom - 2;
                double searchBottom = searchTop - 25;

                double searchLeft = anchor.BoundingBox.Left - 5;
                double searchRight = anchor.BoundingBox.Right + 60;

                var candidateWords = words.Where(w =>
                        w.BoundingBox.Top < searchTop &&
                        w.BoundingBox.Bottom > searchBottom &&
                        w.BoundingBox.Left > searchLeft &&
                        w.BoundingBox.Right < searchRight &&                       
                        (Regex.IsMatch(w.Text, @"\d") || w.Text.ToUpper() == "RI")

                        ).OrderBy(w => w.BoundingBox.Left).ToList();

                if (candidateWords.Any())
                {
                    return string.Join(" ", candidateWords.Select(w => w.Text));
                }
            }

            var match = Regex.Match(page.Text, @"INVOICE\s*NO\.?\s*(\d+\s*[A-Z]*)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;

            return "UNKNOWN";
        }

        private string GetPoNumberRegex(Page page)
        {
            var match = Regex.Match(page.Text, @"(?:Nro\.?|No\.?|Num)\.?\s*(?:de)?\s*OC[:\s]*(\d+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;

            match = Regex.Match(page.Text, @"PO\s*Number[:\s]+(\d+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;

            return "UNKNOWN";
        }

        private string GetVendorName(Page page)
        {
            var words = page.GetWords();
            if (words.Any(w => w.Text.ToUpper().Contains("ETI"))) return "ETI, LLC";
            return "ETI";
        }
    }
}
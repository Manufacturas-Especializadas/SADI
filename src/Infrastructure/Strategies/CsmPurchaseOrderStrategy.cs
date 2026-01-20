using Core.Entities;
using Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using System.Text.RegularExpressions;

namespace Infrastructure.Strategies
{
    public class CsmPurchaseOrderStrategy : IExtractionStrategy
    {
        public string ClientFolderIdentifier => "05 - CSM";

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

                if(poNumber == "UNKNOWN") Console.WriteLine($"[ERROR No PO found in {fileName}");

                string invoiceNum = GetInvoiceNumberFromFolder(filePath, poNumber);

                foreach (var page in pdf.GetPages())
                {
                    var pageItems = ProcessTableOnPage(page, poNumber, vendorName, "N/A", invoiceNum, fileName);
                    items.AddRange(pageItems);
                }
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

                var invoiceFiles = Directory.GetFiles(invoiceFolder, "*pdf");

                foreach (var invoicePath in invoiceFiles)
                {
                    using (var pdf = PdfDocument.Open(invoicePath))
                    {
                        var page = pdf.GetPage(1);
                        var text = page.Text;

                        if (text.Contains(targetPoNumber))
                        {
                            return ExtractInvoiceNumberCsm(page);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error CSM Invoice: {ex.Message}");
            }

            return "NOT FOUND";
        }

        private string ExtractInvoiceNumberCsm(Page page)
        {
            var words = page.GetWords().ToList();

            var headerAnchor = words.FirstOrDefault(w => w.Text.Contains("Invoice") &&
                                                    words.Any(hash => hash.Text.Contains("#") &&
                                                    Math.Abs(hash.BoundingBox.Bottom - w.BoundingBox.Bottom) < 5));

            if(headerAnchor != null)
            {
                double searchTop = headerAnchor.BoundingBox.Bottom;
                double searchBottom = searchTop - 50;
                double searchLeft = headerAnchor.BoundingBox.Left - 20;
                double searchRight = headerAnchor.BoundingBox.Right + 50;

                var candidate = words.FirstOrDefault(w => 
                                    w.BoundingBox.Top < searchTop &&
                                    w.BoundingBox.Bottom > searchBottom &&
                                    w.BoundingBox.Left > searchLeft &&
                                    w.BoundingBox.Right < searchRight && 
                                    Regex.IsMatch(w.Text, @"^\d+$"));

                if(candidate != null) return candidate.Text;
            }

            var match = Regex.Match(page.Text, @"Invoice\s*#\s*(\d+)", RegexOptions.IgnoreCase);

            if (match.Success) return match.Groups[1].Value;

            return "UNKNOWN";
        }

        private string GetPoNumberRegex(Page page)
        {
            var match = Regex.Match(page.Text, @"PO\s*Number[:\s]+(\d+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;

            return "UNKNOWN";
        }

        private string GetVendorName(Page page)
        {
            var words = page.GetWords();
            if (words.Any(w => w.Text.ToUpper().Contains("CSM"))) return "CSM CORPORATION";

            return "CSM";
        }

        private List<PurchaseOrderItem> ProcessTableOnPage(Page page, string po, string vendor, string incoterm, string invoiceNum, string file)
        {
            var items = new List<PurchaseOrderItem>();
            var words = page.GetWords().ToList();

            var lineHeader = words.FirstOrDefault(w => w.Text == "Line");
            var partHeader = words.FirstOrDefault(w => w.Text == "Part");
            var qtyHeader = words.FirstOrDefault(w => w.Text == "Order");

            if (lineHeader == null || partHeader == null) return items;

            double tableTopY = lineHeader.BoundingBox.Bottom;

            var lineCandidates = words.Where(w =>
                    w.BoundingBox.Top < tableTopY &&
                    w.BoundingBox.Centroid.Y < lineHeader.BoundingBox.Bottom &&
                    w.BoundingBox.Left >= lineHeader.BoundingBox.Left - 20 &&
                    w.BoundingBox.Right <= lineHeader.BoundingBox.Right + 20 &&
                    int.TryParse(w.Text, out _)).ToList();


            foreach(var lineWord in lineCandidates)
            {
                int lineNumber = int.Parse(lineWord.Text);
                double rowY = lineWord.BoundingBox.Centroid.Y;
                double rowTolerence = 10;
                var rowWords = words.Where(w => Math.Abs(w.BoundingBox.Centroid.Y - rowY)
                    < rowTolerence && w != lineWord).ToList();

                var item = new PurchaseOrderItem
                {
                    PoNumber = po,
                    VendorName = vendor,
                    Incoterm = incoterm,
                    InvoiceNumber = invoiceNum,
                    LineNumber = lineNumber,
                    SourceFileName = file
                };

                var partNumWord = rowWords
                                .FirstOrDefault(w => w.BoundingBox.Left >= partHeader.BoundingBox.Left - 10);

                if( partNumWord != null) item.PartNumber = partNumWord.Text;

                double searchQtyX = qtyHeader != null ? qtyHeader.BoundingBox.Left : partHeader.BoundingBox.Right + 50;
                var qtyWord = rowWords.FirstOrDefault(w => w.BoundingBox.Left >= searchQtyX - 20 && Regex.IsMatch(w.Text, @"\d"));

                if(qtyWord != null)
                {
                    var match = Regex.Match(qtyWord.Text, @"([\d,.]+)([A-Za-z]*)");

                    if( match.Success )
                    {
                        string numberPart = match.Groups[1].Value.Replace(",", "");
                        if (decimal.TryParse(numberPart, out decimal q)) item.Quantity = q;
                    }
                }

                items.Add(item);
            }

            return items;
        }

    }
}
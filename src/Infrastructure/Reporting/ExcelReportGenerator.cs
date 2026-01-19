using ClosedXML.Excel;
using Core.Entities;
using Core.Interfaces;

namespace Infrastructure.Reporting
{
    public class ExcelReportGenerator : IReportGenerator
    {
        public void AppendToMasterLog(List<PurchaseOrderItem> data, string masterFilePath)
        {
            using (var workbook = new XLWorkbook(masterFilePath))
            {
                var worksheet = workbook.Worksheet(1);
              
                int row = 2;
               
                while (!string.IsNullOrWhiteSpace(worksheet.Cell(row, "B").GetString()))
                {
                    row++;
                    if (row > 10000) break;
                }

                Console.WriteLine($"   [Excel] Escribiendo {data.Count} registros a partir de la fila: {row}");

                foreach (var item in data)
                {
                    worksheet.Cell(row, "B").Value = item.VendorName;
                    worksheet.Cell(row, "C").Value = item.PartNumber;
                    worksheet.Cell(row, "D").Value = item.Quantity;
                    worksheet.Cell(row, "G").Value = item.InvoiceNumber;
                    worksheet.Cell(row, "J").Value = item.PoNumber;
                    worksheet.Cell(row, "R").Value = item.SourceFileName;
                   
                    if (row > 2)
                    {
                        var rowRange = worksheet.Range(row, 1, row, 17);
                        rowRange.Style = worksheet.Row(row - 1).Style;
                    }                    

                    row++;
                }

                Console.WriteLine("   [Excel] Guardando archivo (esto puede tardar unos segundos)...");
                workbook.Save();
            }
        }

        public HashSet<string> GetProcessFileNames(string masterFilePath)
        {
            var processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(masterFilePath)) return processedFiles;

            try
            {
                using (var workbook = new XLWorkbook(masterFilePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    var lastRow = worksheet.LastRowUsed();
                    if (lastRow == null) return processedFiles;

                    var rows = worksheet.Range($"R2:R{lastRow.RowNumber()}").CellsUsed();

                    foreach(var cell in rows)
                    {
                        string fileName = cell.GetString().Trim();

                        if (!string.IsNullOrEmpty(fileName))
                        {
                            processedFiles.Add(fileName);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine($"[ADVERTENCIA] No se pudo leer el historial de archivos: {ex.Message}");
            }

            return processedFiles;
        }
    }
}
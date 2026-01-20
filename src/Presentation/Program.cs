using Core.Entities;
using Core.Interfaces;
using Infrastructure.Reporting;
using Infrastructure.Strategies;

string rootPath = @"\\192.168.25.54\Cadena de suministro\04 - Import - Export";
string masterFileName = "Relación de Importación 2026.xlsx";
string masterFilePath = Path.Combine(rootPath, masterFileName);

if (!File.Exists(masterFilePath))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"[ERROR] No existe el archivo maestro: {masterFilePath}");
    return;
}

IReportGenerator reportGenerator = new ExcelReportGenerator();

Console.Write("Leyendo historial de archivos procesados... ");
var processedFiles = reportGenerator.GetProcessFileNames(masterFilePath);
Console.WriteLine($"Encontrados {processedFiles.Count} archivos previos.");

var strategies = new List<IExtractionStrategy>
{
    new LucasPurchaseOrderStrategy(),
    new CsmPurchaseOrderStrategy(),
    new ElkhartPurchaseOrderStrategy()
};

var newItemsToSave = new List<PurchaseOrderItem>();

foreach (var strategy in strategies)
{
    string strategyPath = Path.Combine(rootPath, strategy.ClientFolderIdentifier, strategy.DocumentTypeSubFolder);
    Console.WriteLine($"\nEscaneando carpeta: {strategy.ClientFolderIdentifier}...");

    if (!Directory.Exists(strategyPath)) continue;

    var directoryInfo = new DirectoryInfo(strategyPath);
    var allFiles = directoryInfo.GetFiles("*.pdf")
                                .OrderByDescending(f => f.LastWriteTime)
                                .ToList();

    int countNew = 0;
    int countSkipped = 0;

    foreach (var fileInfo in allFiles)
    {
        if (processedFiles.Contains(fileInfo.Name))
        {
            countSkipped++;
            continue;
        }

        Console.Write($"   [NUEVO] Procesando {fileInfo.Name} ({fileInfo.LastWriteTime:dd/MM HH:mm})... ");

        try
        {
            var extracted = strategy.Extract(fileInfo.FullName);

            if (extracted.Count > 0)
            {
                newItemsToSave.AddRange(extracted);
                Console.WriteLine("OK");
                countNew++;
            }
            else
            {
                Console.WriteLine("Sin datos legibles.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
        }
    }

    Console.WriteLine($"   Resumen carpeta: {countNew} nuevos procesados, {countSkipped} ignorados por duplicado.");
}

if (newItemsToSave.Count > 0)
{
    Console.WriteLine("\n------------------------------------------------");
    Console.WriteLine($"Agregando {newItemsToSave.Count} registros nuevos al Excel...");

    try
    {
        reportGenerator.AppendToMasterLog(newItemsToSave, masterFilePath);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("¡ACTUALIZACIÓN COMPLETADA CON ÉXITO!");
    }
    catch (IOException)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("ERROR: El archivo Excel está ABIERTO. Ciérralo e intenta de nuevo.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ERROR CRÍTICO: {ex.Message}");
    }
}
else
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("\nNo hay archivos nuevos. Todo está actualizado.");
}

Console.ResetColor();
Console.WriteLine("\nPresiona cualquier tecla para salir...");
Console.ReadKey();
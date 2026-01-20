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
    Console.WriteLine($"[ERROR FATAL] No existe el archivo maestro: {masterFilePath}");
    Console.ResetColor();
    Console.ReadKey();
    return;
}

IReportGenerator reportGenerator = new ExcelReportGenerator();
bool continuarPrograma = true;

while (continuarPrograma)
{
    Console.Clear();
    Console.WriteLine("================================================");
    Console.WriteLine("   ACTUALIZADOR DE RELACIÓN DE IMPORTACIÓN");
    Console.WriteLine("================================================");
    Console.WriteLine("\nSeleccione el cliente a procesar:");
    Console.WriteLine("1. Lucas");
    Console.WriteLine("2. CSM");
    Console.WriteLine("3. Elkhart");
    Console.WriteLine("4. TODOS (Procesar todos los clientes)");
    Console.WriteLine("5. Salir");
    Console.Write("\nIngrese su opción (1-5): ");

    string? option = Console.ReadLine();

    var strategies = new List<IExtractionStrategy>();

    switch (option)
    {
        case "1":
            Console.WriteLine("\n-> Seleccionado: Lucas");
            strategies.Add(new LucasPurchaseOrderStrategy());
            break;
        case "2":
            Console.WriteLine("\n-> Seleccionado: CSM");
            strategies.Add(new CsmPurchaseOrderStrategy());
            break;
        case "3":
            Console.WriteLine("\n-> Seleccionado: Elkhart");
            strategies.Add(new ElkhartPurchaseOrderStrategy());
            break;
        case "4":
            Console.WriteLine("\n-> Seleccionado: TODOS");
            strategies.Add(new LucasPurchaseOrderStrategy());
            strategies.Add(new CsmPurchaseOrderStrategy());
            strategies.Add(new ElkhartPurchaseOrderStrategy());
            break;
        case "5":
            Console.WriteLine("\nSaliendo del programa...");
            continuarPrograma = false;
            continue;
        default:
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n[!] Opción no válida. Por favor intente de nuevo.");
            Console.ResetColor();
            Console.WriteLine("Presione cualquier tecla para continuar...");
            Console.ReadKey();
            continue;
    }

    if (strategies.Count > 0)
    {
        Console.Write("Leyendo historial actualizado del Excel... ");
        var processedFiles = reportGenerator.GetProcessFileNames(masterFilePath);
        Console.WriteLine($"Historial cargado ({processedFiles.Count} archivos previos).");

        var newItemsToSave = new List<PurchaseOrderItem>();

        foreach (var strategy in strategies)
        {
            string strategyPath = Path.Combine(rootPath, strategy.ClientFolderIdentifier, strategy.DocumentTypeSubFolder);
            Console.WriteLine($"\n------------------------------------------------");
            Console.WriteLine($"Escaneando carpeta: {strategy.ClientFolderIdentifier}...");

            if (!Directory.Exists(strategyPath))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"[ADVERTENCIA] No existe la carpeta: {strategyPath}");
                Console.ResetColor();
                continue;
            }

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

            Console.WriteLine($"   Resumen: {countNew} nuevos, {countSkipped} ignorados.");
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
            Console.WriteLine("\nNo se encontraron archivos nuevos para procesar.");
        }

        Console.ResetColor();
        Console.WriteLine("\nPresione cualquier tecla para volver al menú...");
        Console.ReadKey();
    }
}
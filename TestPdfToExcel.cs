using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace PdfToExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("==========================================");
            Console.WriteLine("  Test: Ejecutar pdf_to_excel.py desde C#");
            Console.WriteLine("==========================================");
            Console.WriteLine();

            // TODO: Get PDF path from database
            string pdfPath = GetPdfPathFromDatabase();
            

            // Verificar que el PDF existe
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"❌ Error: El archivo PDF no existe: {pdfPath}");
                Console.WriteLine();
                Console.WriteLine("Presiona cualquier tecla para salir...");
                try
                {
                    Console.ReadKey();
                }
                catch (InvalidOperationException)
                {
                    Console.WriteLine("Ejecución completada.");
                }
                return;
            }

            Console.WriteLine($"📄 PDF a procesar: {pdfPath}");
            Console.WriteLine();

            // Execute Python script
            var startInfo = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"pdf_to_excel.py\" \"{pdfPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            

            Console.WriteLine("🔄 Ejecutando pdf_to_excel.py...");
            Console.WriteLine();

            using (var process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    Console.WriteLine("❌ Error: No se pudo iniciar el proceso Python.");
                    Console.WriteLine("   Verifica que Python esté instalado y en PATH.");
                    return;
                }


                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                Console.WriteLine("==========================================");
                Console.WriteLine("  RESULTADOS");
                Console.WriteLine("==========================================");
                Console.WriteLine();

                if (process.ExitCode == 0)
                {
                    // Extract Excel path from output: "✅ Excel file created successfully -> {path}"
                    var match = Regex.Match(output, @"Excel file created successfully -> (.+)");
                    string excelPath = match.Success ? match.Groups[1].Value.Trim() : Path.ChangeExtension(pdfPath, ".xlsx");

                    Console.WriteLine("✅ Estado: ÉXITO");
                    if (File.Exists(excelPath))
                    {
                        var fileInfo = new FileInfo(excelPath);
                        Console.WriteLine($"📊 Excel creado: {excelPath} ({fileInfo.Length:N0} bytes)");
                    }
                    else
                    {
                        Console.WriteLine($"📊 Excel creado: {excelPath}");
                    }
                    Console.WriteLine();

                    // TODO: Update database with excelPath and status 'Done'
                    UpdateDatabaseWithExcelPath(pdfPath, excelPath, "Done", string.Empty);
                }
                else
                {
                    Console.WriteLine("❌ Estado: ERROR");
                    Console.WriteLine($"Error: {error}");
                    Console.WriteLine();

                    // TODO: Update database with status 'Failed' and error message
                    UpdateDatabaseWithExcelPath(pdfPath, string.Empty, "Failed", error);
                }

                Console.WriteLine("==========================================");
                Console.WriteLine();
                Console.WriteLine("  OUTPUT DEL SCRIPT PYTHON");
                Console.WriteLine();
                Console.WriteLine("==========================================");
                Console.WriteLine();
                Console.WriteLine(output);
                Console.WriteLine();
            }

            Console.WriteLine("==========================================");
            Console.WriteLine();
            Console.WriteLine("Presiona cualquier tecla para salir...");
            try
            {
                Console.ReadKey();
            }
            catch (InvalidOperationException)
            {
                Console.WriteLine("Ejecución completada.");
            }
        }

        // Dummy method: Get PDF path from database
        static string GetPdfPathFromDatabase()
        {
            // TODO: Replace with actual database query
            // Example SQL: SELECT PDFTOOLPATH FROM BANK_STTEMENT_SEARCH_HISTORY_FORWINDOWSERVICE WHERE Status = 'Pending'
            return @"Test\Bank Statement\Final Test\SCOTIABANK TEST JULY 2022.pdf";
        }

        // Dummy method: Update database with Excel path and status
        static void UpdateDatabaseWithExcelPath(string pdfPath, string excelPath, string status, string errorMessage)
        {
            // TODO: Replace with actual database update
            // Example SQL: UPDATE BANK_STTEMENT_SEARCH_HISTORY_FORWINDOWSERVICE 
            // SET FILEPATH = @excelPath, Status = @status, ErrorMessage = @errorMessage
            // WHERE PDFTOOLPATH = @pdfPath
            Console.WriteLine($"📝 Se actualizaría la BD: PDF={pdfPath}, Excel={excelPath ?? "null"}, Status={status}");
        }
    }
}

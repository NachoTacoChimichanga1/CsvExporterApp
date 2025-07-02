using System;
using System.Collections.Generic;
using System.Data;
using CsvExporterLibrary;
using ExcelDataReader;
using System.Text;
using Microsoft.Win32;
using System.IO;
using System.Linq;

namespace MyApp

{
    internal class Program
    {

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            while (true)
            {
                Console.WriteLine("\nИзбери опция:");
                Console.WriteLine("1. Избери .xlsx файл и го конвертирай в .csv");
                Console.WriteLine("2. Създай произволна таблица и експортирай");
                Console.WriteLine("0. Изход");
                Console.Write("Избор: ");

                var choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ConvertExcelToCsv();
                        break;
                    case "2":
                        CustomTableEntry();
                        break;
                    case "0":
                        return;
                    default:
                        Console.WriteLine("Невалиден избор.");
                        break;
                }
            }
        }

        static void ConvertExcelToCsv()
        {
            Console.Write("Въведи път до Excel (.xlsx) файл: ");
            var filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Файлът не съществува.");
                return;
            }

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var sb = new StringBuilder();

            //popravka
            while (reader.Read())
            {
                var values = new List<string>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var value = reader.GetValue(i)?.ToString() ?? "";

                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                    {
                        value = value.Replace("\"", "\"\"");
                        value = $"\"{value}\"";
                    }

                    values.Add(value);
                }

                sb.AppendLine(string.Join(",", values));
            }
            //gotovo

            Console.Write("Въведи път за запис на .csv файл: ");
            var outputPath = Console.ReadLine();
            outputPath = EnsureCsvExtension(outputPath);

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

            Console.WriteLine("Успешно конвертирано.");
        }

        static void CustomTableEntry()
        {
            Console.Write("Въведи имена на колоните, разделени със запетая (NA LATINICA!!!): ");
            var columnInput = Console.ReadLine();
            var columns = columnInput.Split(',').Select(c => c.Trim()).ToList();

            var rows = new List<List<string>>();

            while (true)
            {
                var currentRow = new List<string>();
                Console.WriteLine("Въведи стойности за реда:");

                foreach (var column in columns)
                {
                    Console.Write($"{column}: ");
                    currentRow.Add(Console.ReadLine());
                }

                rows.Add(currentRow);

                Console.Write("Добави още един ред? (y/n): ");
                if (Console.ReadLine().ToLower() != "y") break;
            }

            Console.Write("Въведи път за запис на .csv файл: ");
            var filePath = Console.ReadLine();
            filePath = EnsureCsvExtension(filePath);

            CsvExporter.ExportCustomListToCsv(columns, rows, filePath);
            Console.WriteLine("Файлът е запазен успешно.");
        }

        static string EnsureCsvExtension(string path)
        {
            if (!path.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                path += ".csv";
            }
            return path;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;

namespace CsvExporterLibrary
{
    public static class CsvExporter
    {
        public static void ExportListToCsv<T>(List<T> list, string filePath)
        {
            var properties = typeof(T).GetProperties();
            var sb = new StringBuilder();

            sb.AppendLine(string.Join(";", properties.Select(p => p.Name)));

            foreach (var item in list)
            {
                var values = properties.Select(p => p.GetValue(item)?.ToString()?.Replace(";", ",") ?? "");
                sb.AppendLine(string.Join(";", values));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static void ExportCustomListToCsv(List<string> columns, List<List<string>> rows, string filePath)
        {
            var sb = new StringBuilder();

            sb.AppendLine(string.Join(";", columns));

            foreach (var row in rows)
            {
                var safeRow = row.Select(value => $"\"{value?.Replace("\"", "\"\"")}\"");
                sb.AppendLine(string.Join(";", safeRow));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
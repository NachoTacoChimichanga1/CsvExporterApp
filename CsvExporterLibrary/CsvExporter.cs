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
        public static void ExportListToCsv<Entity>(List<Entity> list, string filePath)
        {
            var properties = typeof(Entity).GetProperties();
            var sb = new StringBuilder();

            sb.AppendLine(string.Join(",", properties.Select(p => p.Name)));

            foreach (var item in list)
            {
                var values = properties.Select(p =>
                {
                    var value = p.GetValue(item)?.ToString() ?? "";
                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                    {
                        value = value.Replace("\"", "\"\"");
                        value = $"\"{value}\"";
                    }
                    return value;
                });

                sb.AppendLine(string.Join(",", values));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static void ExportCustomListToCsv(List<string> columns, List<List<string>> rows, string filePath)
        {
            var sb = new StringBuilder();

            sb.AppendLine(string.Join(",", columns));

            foreach (var row in rows)
            {
                var safeRow = row.Select(value =>
                {
                    var val = value?.Replace("\"", "\"\"") ?? "";
                    if (val.Contains(",") || val.Contains("\"") || val.Contains("\n"))
                        val = $"\"{val}\"";
                    return val;
                });

                sb.AppendLine(string.Join(",", safeRow));
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
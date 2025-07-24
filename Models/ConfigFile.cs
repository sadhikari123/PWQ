using System;
using System.Collections.Generic;
using System.IO;
using CsvHelper;
using System.Globalization;
using System.Linq;

namespace PWQ.Models
{
    public class ConfigFile
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public List<Dictionary<string, string>> Rows { get; set; } = new List<Dictionary<string, string>>();
        public List<string> Columns { get; set; } = new List<string>();

        public void LoadFromFile()
        {
            if (!File.Exists(Path))
            {
                throw new FileNotFoundException($"File not found: {Path}");
            }

            Rows.Clear();
            Columns.Clear();

            using (var reader = new StreamReader(Path))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            using (var dr = new CsvHelper.CsvDataReader(csv))
            {
                var dt = new System.Data.DataTable();
                dt.Load(dr);
                Columns = dt.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToList();
                foreach (System.Data.DataRow drow in dt.Rows)
                {
                    var dict = new Dictionary<string, string>();
                    foreach (var col in Columns)
                        dict[col] = drow[col]?.ToString() ?? "";
                    Rows.Add(dict);
                }
            }
        }

        public void SaveToFile()
        {
            using (var writer = new StreamWriter(Path))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                foreach (var col in Columns)
                    csv.WriteField(col);
                csv.NextRecord();
                foreach (var row in Rows)
                {
                    foreach (var col in Columns)
                        csv.WriteField(row.ContainsKey(col) ? row[col] : "");
                    csv.NextRecord();
                }
            }
        }

        public void CreateEmptyStructure()
        {
            Rows.Clear();
            Columns.Clear();

            // Create default columns based on file type
            switch (Name)
            {
                case "LHD_CONFIG":
                    Columns = new List<string> { "ID", "Name", "Value", "Description", "Category" };
                    break;
                case "LHD_Run_Plan":
                    Columns = new List<string> { "ID", "Step", "Action", "Parameters", "Expected_Result" };
                    break;
                case "LHD_Settings":
                    Columns = new List<string> { "Setting_Name", "Setting_Value", "Default_Value", "Type", "Description" };
                    break;
                default:
                    Columns = new List<string> { "ID", "Name", "Value" };
                    break;
            }
        }

        public void AddRow(Dictionary<string, object> rowData)
        {
            var dict = rowData.ToDictionary(kvp => kvp.Key, kvp => kvp.Value?.ToString() ?? "");
            Rows.Add(dict);
        }

        public void UpdateRow(int index, Dictionary<string, object> rowData)
        {
            if (index >= 0 && index < Rows.Count)
            {
                var dict = rowData.ToDictionary(kvp => kvp.Key, kvp => kvp.Value?.ToString() ?? "");
                Rows[index] = dict;
            }
        }

        public void DeleteRow(int index)
        {
            if (index >= 0 && index < Rows.Count)
            {
                Rows.RemoveAt(index);
            }
        }
    }
}
